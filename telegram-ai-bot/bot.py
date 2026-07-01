import os
import re
import json
import base64
import logging
import urllib.parse
from datetime import datetime
import httpx
import gspread
from google.oauth2.service_account import Credentials
import google.generativeai as genai
from telegram import Update, InputMediaPhoto
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN     = os.environ["TELEGRAM_TOKEN"]
GEMINI_API_KEY     = os.environ["GEMINI_API_KEY"]
GOOGLE_CREDENTIALS = os.environ["GOOGLE_CREDENTIALS"]
SPREADSHEET_ID     = os.environ["SPREADSHEET_ID"]
ALLOWED_USER_ID    = int(os.environ.get("ALLOWED_USER_ID", "0"))

genai.configure(api_key=GEMINI_API_KEY)
gemini = genai.GenerativeModel("gemini-2.5-flash")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

sessions: dict[int, genai.ChatSession] = {}


def is_allowed(user_id: int) -> bool:
    return ALLOWED_USER_ID == 0 or user_id == ALLOWED_USER_ID


def get_gc():
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDENTIALS), scopes=SCOPES
    )
    return gspread.authorize(creds)


def format_rupiah(amount: int) -> str:
    return f"Rp {amount:,.0f}".replace(",", ".")


def setup_spreadsheet():
    gc = get_gc()
    ss = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws_t = ss.worksheet("Transaksi")
    except gspread.WorksheetNotFound:
        ws_t = ss.add_worksheet("Transaksi", 1000, 7)

    ws_t.update("A1:G1", [["No", "Tanggal", "Waktu", "Tipe", "Jumlah (Rp)", "Kategori", "Keterangan"]])
    ws_t.format("A1:G1", {
        "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 11},
        "horizontalAlignment": "CENTER",
    })
    ss.batch_update({"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": ws_t.id, "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount",
        }
    }]})

    try:
        ws_l = ss.worksheet("Laporan")
    except gspread.WorksheetNotFound:
        ws_l = ss.add_worksheet("Laporan", 100, 5)

    _build_laporan(ws_l, ss)


def _build_laporan(ws, ss):
    bulan = datetime.now().strftime("%B %Y")
    ws.clear()

    ss.batch_update({"requests": [{
        "mergeCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "mergeType": "MERGE_ALL",
        }
    }]})
    ws.update("A1", [[f"LAPORAN KEUANGAN — {bulan}"]])
    ws.format("A1", {
        "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 14},
        "horizontalAlignment": "CENTER",
    })

    ws.update("A3", [["RINGKASAN"]])
    ws.format("A3:E3", {
        "backgroundColor": {"red": 0.20, "green": 0.40, "blue": 0.60},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })
    ws.update("A4:B6", [
        ["Total Pemasukan",   '=SUMIF(Transaksi!D:D;"MASUK";Transaksi!E:E)'],
        ["Total Pengeluaran", '=SUMIF(Transaksi!D:D;"KELUAR";Transaksi!E:E)'],
        ["Saldo Bersih",      "=B4-B5"],
    ], value_input_option="USER_ENTERED")
    ws.format("A4:A6", {"textFormat": {"bold": True}})
    ws.format("B4", {"backgroundColor": {"red": 0.85, "green": 0.93, "blue": 0.83}})
    ws.format("B5", {"backgroundColor": {"red": 0.99, "green": 0.91, "blue": 0.91}})
    ws.format("B6", {
        "backgroundColor": {"red": 0.93, "green": 0.95, "blue": 1.0},
        "textFormat": {"bold": True},
    })
    ws.format("B4:B6", {"numberFormat": {"type": "NUMBER", "pattern": '"Rp "#,##0'}})

    ws.update("A8", [["PENGELUARAN PER KATEGORI"]])
    ws.format("A8:E8", {
        "backgroundColor": {"red": 0.80, "green": 0.20, "blue": 0.20},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })
    kat_keluar = ["Makan", "Transport", "Belanja", "Tagihan", "Hiburan", "Kesehatan", "Lainnya"]
    for i, k in enumerate(kat_keluar):
        r = 9 + i
        ws.update(f"A{r}:B{r}", [[k, f'=SUMIFS(Transaksi!E:E;Transaksi!D:D;"KELUAR";Transaksi!F:F;"{k}")']], value_input_option="USER_ENTERED")
        ws.format(f"B{r}", {"numberFormat": {"type": "NUMBER", "pattern": '"Rp "#,##0'}})
        if i % 2 == 0:
            ws.format(f"A{r}:B{r}", {"backgroundColor": {"red": 0.99, "green": 0.95, "blue": 0.95}})

    sr = 9 + len(kat_keluar) + 2
    ws.update(f"A{sr}", [["PEMASUKAN PER KATEGORI"]])
    ws.format(f"A{sr}:E{sr}", {
        "backgroundColor": {"red": 0.20, "green": 0.60, "blue": 0.30},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })
    kat_masuk = ["Gaji", "Jualan", "Freelance", "Investasi", "Bonus", "Lainnya"]
    for i, k in enumerate(kat_masuk):
        r = sr + 1 + i
        ws.update(f"A{r}:B{r}", [[k, f'=SUMIFS(Transaksi!E:E;Transaksi!D:D;"MASUK";Transaksi!F:F;"{k}")']], value_input_option="USER_ENTERED")
        ws.format(f"B{r}", {"numberFormat": {"type": "NUMBER", "pattern": '"Rp "#,##0'}})
        if i % 2 == 0:
            ws.format(f"A{r}:B{r}", {"backgroundColor": {"red": 0.90, "green": 0.98, "blue": 0.90}})


def catat_ke_sheets(tipe: str, jumlah: int, kategori: str, keterangan: str):
    gc = get_gc()
    ss = gc.open_by_key(SPREADSHEET_ID)
    ws = ss.worksheet("Transaksi")

    now      = datetime.now()
    all_rows = ws.get_all_values()
    next_row = len(all_rows) + 1
    no       = next_row - 1

    ws.append_row(
        [no, now.strftime("%d/%m/%Y"), now.strftime("%H:%M"),
         tipe.upper(), jumlah, kategori, keterangan],
        value_input_option="USER_ENTERED",
    )

    bg = ({"red": 0.85, "green": 0.93, "blue": 0.83}
          if tipe.upper() == "MASUK"
          else {"red": 0.99, "green": 0.91, "blue": 0.91})
    ws.format(f"A{next_row}:G{next_row}", {"backgroundColor": bg})
    ws.format(f"E{next_row}", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}})


async def parse_transaksi(text: str) -> dict | None:
    prompt = f"""Analisa pesan ini. Jika pencatatan transaksi keuangan, ekstrak detailnya.

Pesan: "{text}"

Jika transaksi, return JSON:
{{"is_transaction": true, "tipe": "masuk" atau "keluar", "jumlah": angka_rupiah, "kategori": "nama", "keterangan": "singkat"}}

Kategori pengeluaran: Makan, Transport, Belanja, Tagihan, Hiburan, Kesehatan, Lainnya
Kategori pemasukan: Gaji, Jualan, Freelance, Investasi, Bonus, Lainnya

Contoh:
- "beli makan 25rb" → {{"is_transaction":true,"tipe":"keluar","jumlah":25000,"kategori":"Makan","keterangan":"makan"}}
- "masuk 500rb jualan" → {{"is_transaction":true,"tipe":"masuk","jumlah":500000,"kategori":"Jualan","keterangan":"jualan"}}
- "halo apa kabar" → {{"is_transaction":false}}

Return JSON saja."""

    try:
        resp = gemini.generate_content(prompt)
        raw  = re.sub(r"```json\n?|\n?```", "", resp.text.strip()).strip()
        return json.loads(raw)
    except Exception as e:
        logger.error(f"Parse error: {e}")
        return None


async def get_saldo_text() -> str:
    try:
        gc      = get_gc()
        records = gc.open_by_key(SPREADSHEET_ID).worksheet("Transaksi").get_all_records()
        masuk   = sum(r["Jumlah (Rp)"] for r in records if r["Tipe"] == "MASUK")
        keluar  = sum(r["Jumlah (Rp)"] for r in records if r["Tipe"] == "KELUAR")
        saldo   = masuk - keluar
        return (
            f"*Saldo Keuangan*\n\n"
            f"Pemasukan   : {format_rupiah(masuk)}\n"
            f"Pengeluaran : {format_rupiah(keluar)}\n"
            f"────────────────────\n"
            f"*Saldo Bersih: {format_rupiah(saldo)}*"
        )
    except Exception as e:
        logger.error(f"Saldo error: {e}")
        return "Gagal mengambil data saldo."


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Halo! Bot AI Pribadi.\n\n"
        "KEUANGAN — cukup chat biasa:\n"
        "  beli makan 25rb\n"
        "  masuk gaji 3 juta\n"
        "  bayar listrik 150000\n\n"
        "PERINTAH AI:\n"
        "/image <prompt>  — Generate gambar\n"
        "/editimage       — Edit foto (kirim foto + caption)\n"
        "/story <tema>    — Buat script + gambar per scene\n"
        "/ugc <produk>    — Generate foto UGC style\n\n"
        "KEUANGAN:\n"
        "/saldo  — Cek saldo\n"
        "/setup  — Setup Google Sheets\n"
        "/clear  — Reset chat AI\n"
        "/help   — Bantuan lengkap"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "AI IMAGE:\n"
        "/image kucing astronot       — Text to image\n"
        "/editimage                   — Lihat cara edit foto\n"
        "  → Kirim foto + caption instruksi edit\n"
        "  → Contoh: foto selfie + caption 'ganti background pantai'\n"
        "/story promosi kopi kekinian — Script + gambar 2 scene\n"
        "/ugc skincare vitamin C      — Foto UGC style otomatis\n\n"
        "KEUANGAN (chat biasa):\n"
        "  beli makan siang 25rb\n"
        "  masuk gaji 3 juta\n"
        "  bayar listrik 150000\n\n"
        "PERINTAH KEUANGAN:\n"
        "/saldo  — Cek total saldo\n"
        "/setup  — Setup ulang Google Sheets\n"
        "/clear  — Reset percakapan AI"
    )


async def setup_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg = await update.message.reply_text("Menyiapkan Google Sheets...")
    try:
        setup_spreadsheet()
        await msg.edit_text("Google Sheets siap! Cek spreadsheet kamu.")
    except Exception as e:
        logger.error(f"Setup error: {e}")
        await msg.edit_text(f"Error: {e}")


async def saldo_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    msg  = await update.message.reply_text("Mengambil data...")
    text = await get_saldo_text()
    await msg.edit_text(text, parse_mode="Markdown")


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return
    sessions[user_id] = gemini.start_chat(history=[])
    await update.message.reply_text("Riwayat percakapan dihapus!")


async def image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    if not context.args:
        await update.message.reply_text("Contoh: /image kucing astronot di bulan")
        return
    prompt = " ".join(context.args)
    msg    = await update.message.reply_text("Generating gambar...")
    try:
        encoded = urllib.parse.quote(prompt)
        seed    = abs(hash(prompt)) % 99999
        url     = (f"https://image.pollinations.ai/prompt/{encoded}"
                   f"?width=1024&height=1024&nologo=true&enhance=true&seed={seed}")
        async with httpx.AsyncClient(timeout=90) as client:
            resp = await client.get(url)
            resp.raise_for_status()
        await msg.delete()
        await update.message.reply_photo(photo=resp.content, caption=prompt)
    except Exception as e:
        logger.error(f"Image error: {e}")
        await msg.edit_text("Gagal generate gambar. Coba lagi.")


async def chat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan menggunakan bot ini.")
        return

    text = update.message.text
    await update.message.chat.send_action("typing")

    parsed = await parse_transaksi(text)
    if parsed and parsed.get("is_transaction"):
        tipe       = parsed["tipe"]
        jumlah     = int(parsed["jumlah"])
        kategori   = parsed.get("kategori", "Lainnya")
        keterangan = parsed.get("keterangan", text)
        try:
            catat_ke_sheets(tipe, jumlah, kategori, keterangan)
            tanda = "+" if tipe == "masuk" else "-"
            await update.message.reply_text(
                f"Tercatat!\n\n"
                f"Tipe      : {'Pemasukan' if tipe == 'masuk' else 'Pengeluaran'}\n"
                f"Jumlah    : {tanda}{format_rupiah(jumlah)}\n"
                f"Kategori  : {kategori}\n"
                f"Keterangan: {keterangan}\n\n"
                f"Tersimpan di Google Sheets."
            )
        except Exception as e:
            logger.error(f"Sheets error: {e}")
            await update.message.reply_text("Gagal menyimpan ke Sheets. Coba lagi.")
        return

    if user_id not in sessions:
        sessions[user_id] = gemini.start_chat(history=[])
    try:
        reply = sessions[user_id].send_message(text).text
        if len(reply) > 4096:
            for i in range(0, len(reply), 4096):
                await update.message.reply_text(reply[i:i + 4096])
        else:
            await update.message.reply_text(reply)
    except Exception as e:
        logger.error(f"Gemini error: {e}")
        await update.message.reply_text("Terjadi kesalahan. Coba lagi.")


async def editimage_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    await update.message.reply_text(
        "Cara pakai Edit Image:\n\n"
        "1. Kirim foto ke bot ini\n"
        "2. Isi caption dengan instruksi editnya\n\n"
        "Contoh:\n"
        "  Foto selfie + caption: ganti background jadi pantai sunset\n"
        "  Foto produk + caption: buat jadi foto studio profesional\n"
        "  Foto + caption: ubah jadi gaya anime\n\n"
        "Kirim foto TANPA caption → bot mendeskripsikan gambarnya."
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return

    photo   = update.message.photo[-1]
    caption = (update.message.caption or "").strip()
    msg     = await update.message.reply_text("Memproses foto...")

    try:
        file        = await photo.get_file()
        image_bytes = bytes(await file.download_as_bytearray())
        image_part  = {"mime_type": "image/jpeg", "data": image_bytes}

        if not caption:
            response = gemini.generate_content([
                "Deskripsikan gambar ini dengan detail dalam bahasa Indonesia.",
                image_part,
            ])
            await msg.edit_text(response.text[:4096])
            return

        await msg.edit_text("Mengedit gambar... (30-60 detik)")

        vision_prompt = (
            f'Kamu AI image editor. Analisa foto ini dan buat prompt bahasa Inggris '
            f'yang detail untuk menghasilkan gambar baru berdasarkan instruksi: "{caption}"\n\n'
            f"Gabungkan elemen visual dari foto asli dengan perubahan yang diminta.\n"
            f"Return HANYA prompt bahasa Inggris (max 80 kata) untuk AI image generator."
        )
        vision_resp = gemini.generate_content([vision_prompt, image_part])
        new_prompt  = vision_resp.text.strip()

        encoded = urllib.parse.quote(new_prompt)
        seed    = abs(hash(new_prompt + caption)) % 99999
        url     = (f"https://image.pollinations.ai/prompt/{encoded}"
                   f"?width=1024&height=1024&nologo=true&enhance=true&seed={seed}")

        async with httpx.AsyncClient(timeout=120) as client:
            img_resp = await client.get(url)
            img_resp.raise_for_status()

        await msg.delete()
        await update.message.reply_photo(
            photo=img_resp.content,
            caption=f"Edit: {caption}",
        )
    except Exception as e:
        logger.error(f"Photo error: {e}")
        await msg.edit_text("Gagal memproses foto. Coba lagi.")


async def story_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    if not context.args:
        await update.message.reply_text(
            "Contoh:\n"
            "/story promosi skincare untuk remaja\n"
            "/story petualangan di hutan hujan\n"
            "/story iklan kafe aesthetic"
        )
        return

    tema = " ".join(context.args)
    msg  = await update.message.reply_text(f"Membuat story: {tema}...\n(bisa 1-2 menit)")

    try:
        story_prompt = f"""Buat script/story menarik dengan tema: "{tema}"

Return JSON ini SAJA:
{{
  "judul": "judul menarik",
  "scenes": [
    {{
      "narasi": "narasi scene 1 bahasa Indonesia (2-3 kalimat)",
      "image_prompt": "detailed English prompt for AI image generator, vivid, max 60 words"
    }},
    {{
      "narasi": "narasi scene 2 bahasa Indonesia (2-3 kalimat)",
      "image_prompt": "detailed English prompt for AI image generator, vivid, max 60 words"
    }}
  ],
  "penutup": "kalimat penutup atau CTA bahasa Indonesia"
}}"""

        resp = gemini.generate_content(story_prompt)
        raw  = re.sub(r"```json\n?|\n?```", "", resp.text.strip()).strip()
        data = json.loads(raw)

        for i, scene in enumerate(data["scenes"], 1):
            await msg.edit_text(f"Generating gambar scene {i} dari {len(data['scenes'])}...")

            encoded = urllib.parse.quote(scene["image_prompt"])
            seed    = abs(hash(scene["image_prompt"])) % 99999
            url     = (f"https://image.pollinations.ai/prompt/{encoded}"
                       f"?width=1024&height=576&nologo=true&enhance=true&seed={seed}")

            async with httpx.AsyncClient(timeout=120) as client:
                img_resp = await client.get(url)
                img_resp.raise_for_status()

            cap = f"*{data['judul']}*\n\n*Scene {i}*\n{scene['narasi']}"
            if i == len(data["scenes"]):
                cap += f"\n\n_{data['penutup']}_"

            await update.message.reply_photo(
                photo=img_resp.content,
                caption=cap,
                parse_mode="Markdown",
            )

        await msg.delete()

    except Exception as e:
        logger.error(f"Story error: {e}")
        await msg.edit_text("Gagal membuat story. Coba lagi.")


async def ugc_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update.effective_user.id):
        return
    if not context.args:
        await update.message.reply_text(
            "Contoh:\n"
            "/ugc skincare vitamin C serum\n"
            "/ugc kopi susu kekinian\n"
            "/ugc sepatu sneakers putih"
        )
        return

    produk = " ".join(context.args)
    msg    = await update.message.reply_text(f"Generating UGC image: {produk}...")

    try:
        ugc_req = f"""Buat prompt foto UGC (User Generated Content) style untuk: "{produk}"

UGC style: casual authentic lifestyle photography, seperti influencer/pengguna nyata, natural lighting, candid feel, real environment.

Return JSON ini SAJA:
{{
  "image_prompt": "detailed UGC-style image prompt in English, max 80 words, mention lighting/setting/mood/aesthetic",
  "caption": "caption media sosial bahasa Indonesia yang engaging (1-2 kalimat + 3 hashtag relevan)"
}}"""

        resp = gemini.generate_content(ugc_req)
        raw  = re.sub(r"```json\n?|\n?```", "", resp.text.strip()).strip()
        data = json.loads(raw)

        image_prompt = data["image_prompt"]
        caption      = data["caption"]

        encoded = urllib.parse.quote(image_prompt)
        seed    = abs(hash(image_prompt + produk)) % 99999
        url     = (f"https://image.pollinations.ai/prompt/{encoded}"
                   f"?width=1024&height=1024&nologo=true&enhance=true&seed={seed}")

        async with httpx.AsyncClient(timeout=120) as client:
            img_resp = await client.get(url)
            img_resp.raise_for_status()

        await msg.delete()
        await update.message.reply_photo(
            photo=img_resp.content,
            caption=f"{caption}",
        )

    except Exception as e:
        logger.error(f"UGC error: {e}")
        await msg.edit_text("Gagal generate UGC image. Coba lagi.")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start",      start))
    app.add_handler(CommandHandler("help",       help_command))
    app.add_handler(CommandHandler("clear",      clear))
    app.add_handler(CommandHandler("setup",      setup_command))
    app.add_handler(CommandHandler("saldo",      saldo_command))
    app.add_handler(CommandHandler("image",      image_command))
    app.add_handler(CommandHandler("editimage",  editimage_command))
    app.add_handler(CommandHandler("story",      story_command))
    app.add_handler(CommandHandler("ugc",        ugc_command))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat))
    logger.info("Bot berjalan...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
