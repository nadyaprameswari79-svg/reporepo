import os
import logging
import urllib.parse
import httpx
import google.generativeai as genai
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
GEMINI_API_KEY  = os.environ["GEMINI_API_KEY"]
ALLOWED_USER_ID = int(os.environ.get("ALLOWED_USER_ID", "0"))

genai.configure(api_key=GEMINI_API_KEY)
gemini = genai.GenerativeModel("gemini-2.0-flash")

# Simpan sesi chat per user (biar Gemini ingat konteks)
sessions: dict[int, genai.ChatSession] = {}


def is_allowed(user_id: int) -> bool:
    return ALLOWED_USER_ID == 0 or user_id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Halo! Saya bot AI gratis.\n\n"
        "Perintah:\n"
        "/image <prompt> — Generate gambar dari teks\n"
        "/clear         — Hapus riwayat percakapan\n"
        "/help          — Bantuan\n\n"
        "Kirim pesan apapun untuk mulai chat!"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Cara pakai:\n"
        "- Pesan biasa         : chat dengan Gemini AI\n"
        "- /image <deskripsi>  : generate gambar\n"
        "- /clear              : reset percakapan\n\n"
        "Contoh generate gambar:\n"
        "/image kucing astronot di bulan, digital art\n"
        "/image pantai sunset minimalist watercolor"
    )


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return
    sessions[user_id] = gemini.start_chat(history=[])
    await update.message.reply_text("Riwayat percakapan dihapus. Mulai dari awal!")


async def image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    if not context.args:
        await update.message.reply_text(
            "Tulis deskripsi gambarnya.\n"
            "Contoh: /image kucing astronot di bulan"
        )
        return

    prompt = " ".join(context.args)
    msg = await update.message.reply_text("Generating gambar, tunggu sebentar...")

    try:
        encoded = urllib.parse.quote(prompt)
        seed    = abs(hash(prompt)) % 99999
        url     = (
            f"https://image.pollinations.ai/prompt/{encoded}"
            f"?width=1024&height=1024&nologo=true&enhance=true&seed={seed}"
        )

        async with httpx.AsyncClient(timeout=90) as client:
            resp = await client.get(url)
            resp.raise_for_status()

        await msg.delete()
        await update.message.reply_photo(photo=resp.content, caption=prompt)

    except Exception as e:
        logger.error(f"Image error: {e}")
        await msg.edit_text("Gagal generate gambar. Coba lagi atau ubah prompt-nya.")


async def chat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan menggunakan bot ini.")
        return

    if user_id not in sessions:
        sessions[user_id] = gemini.start_chat(history=[])

    await update.message.chat.send_action("typing")

    try:
        response = sessions[user_id].send_message(update.message.text)
        reply    = response.text

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

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help",  help_command))
    app.add_handler(CommandHandler("clear", clear))
    app.add_handler(CommandHandler("image", image_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat))

    logger.info("Bot berjalan...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
