import os
import re
import json
import logging
import urllib.parse
from datetime import datetime
import httpx
import gspread
from google.oauth2.service_account import Credentials
import google.generativeai as genai
from telegram import Update
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
        ["Total Pemasukan",   '=SUMIF(Transaksi!D:D,"MASUK",Transaksi!E:E)'],
        ["Total Pengeluaran", '=SUMIF(Transaksi!D:D,"KELUAR",Transaksi!E:E)'],
        ["Saldo Bersih",      "=B4-B5"],
    ])
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
        ws.update(f"A{r}:B{r}", [[k, f'=SUMIFS(Transaksi!E:E,Transaksi!D:D,"KELUAR",Transaksi!F:F,"{k}")']])
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
        ws.update(f"A{r}:B{r}", [[k, f'=SUMIFS(Transaksi!E:E,Transaksi!D:D,"MASUK",Transaksi!F:F,"{k}")']])
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
        "Halo! Bot AI + Pencatat Keuangan.\n\n"
        "Catat transaksi cukup chat biasa:\n"
        "  beli makan 25rb\n"
        "  masuk gaji 3 juta\n"
        "  bayar listrik 150000\n\n"
        "Perintah:\n"
        "/saldo  — Cek saldo\n"
        "/setup  — Setup Google Sheets\n"
        "/image  — Generate gambar\n"
        "/clear  — Reset chat AI\n"
        "/help   — Bantuan"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "CATAT TRANSAKSI (chat biasa):\n"
        "  beli makan siang 25rb\n"
        "  masuk gaji 3 juta\n"
        "  bensin 50ribu\n"
        "  terima bayaran freelance 500rb\n\n"
        "PERINTAH:\n"
        "/saldo         — Cek total saldo\n"
        "/setup         — Setup ulang Google Sheets\n"
        "/image <prompt>— Generate gambar\n"
        "/clear         — Reset percakapan AI"
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


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("help",   help_command))
    app.add_handler(CommandHandler("clear",  clear))
    app.add_handler(CommandHandler("setup",  setup_command))
    app.add_handler(CommandHandler("saldo",  saldo_command))
    app.add_handler(CommandHandler("image",  image_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat))
    logger.info("Bot berjalan...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
