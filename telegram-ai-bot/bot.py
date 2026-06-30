import os
import logging
from openai import OpenAI
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
OPENAI_API_KEY  = os.environ["OPENAI_API_KEY"]
ALLOWED_USER_ID = int(os.environ.get("ALLOWED_USER_ID", "0"))

client = OpenAI(api_key=OPENAI_API_KEY)

# Simpan riwayat percakapan per user
conversations: dict[int, list] = {}

SYSTEM_PROMPT = (
    "Kamu adalah asisten AI yang helpful, ramah, dan cerdas. "
    "Jawab dalam bahasa yang sama dengan pertanyaan pengguna."
)


def is_allowed(user_id: int) -> bool:
    return ALLOWED_USER_ID == 0 or user_id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Halo! Saya bot AI yang terhubung ke ChatGPT (GPT-4o).\n\n"
        "Kirim pesan apapun untuk mulai ngobrol.\n\n"
        "Perintah:\n"
        "/clear  — Hapus riwayat percakapan\n"
        "/model  — Ganti model AI\n"
        "/help   — Bantuan"
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Cara pakai:\n"
        "• Kirim pesan biasa → bot reply pakai AI\n"
        "• /clear  — Reset percakapan (mulai dari awal)\n"
        "• /model  — Lihat/ganti model (gpt-4o, gpt-4o-mini, dll)\n"
        "• /start  — Pesan selamat datang\n\n"
        "Bot ini menyimpan riwayat percakapan supaya AI ingat konteks sebelumnya."
    )


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return
    conversations[user_id] = []
    await update.message.reply_text("Riwayat dihapus. Percakapan baru dimulai!")


async def model_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_allowed(user_id):
        return

    if context.args:
        new_model = context.args[0]
        context.user_data["model"] = new_model
        await update.message.reply_text(f"Model diganti ke: {new_model}")
    else:
        current = context.user_data.get("model", "gpt-4o")
        await update.message.reply_text(
            f"Model sekarang: {current}\n\n"
            "Ganti dengan: /model <nama-model>\n"
            "Contoh:\n"
            "  /model gpt-4o\n"
            "  /model gpt-4o-mini\n"
            "  /model gpt-4.1"
        )


async def chat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if not is_allowed(user_id):
        await update.message.reply_text("Maaf, kamu tidak diizinkan menggunakan bot ini.")
        return

    user_message = update.message.text
    model = context.user_data.get("model", "gpt-4o")

    if user_id not in conversations:
        conversations[user_id] = []

    conversations[user_id].append({"role": "user", "content": user_message})

    # Batasi 20 pesan terakhir supaya tidak over token limit
    if len(conversations[user_id]) > 20:
        conversations[user_id] = conversations[user_id][-20:]

    await update.message.chat.send_action("typing")

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": SYSTEM_PROMPT}] + conversations[user_id],
        )
        reply = response.choices[0].message.content
        conversations[user_id].append({"role": "assistant", "content": reply})

        # Telegram max 4096 karakter per pesan
        if len(reply) > 4096:
            for i in range(0, len(reply), 4096):
                await update.message.reply_text(reply[i:i + 4096])
        else:
            await update.message.reply_text(reply)

    except Exception as e:
        logger.error(f"OpenAI error: {e}")
        await update.message.reply_text(
            "Terjadi kesalahan saat menghubungi AI. Coba lagi sebentar."
        )


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("clear", clear))
    app.add_handler(CommandHandler("model", model_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat))

    logger.info("Bot berjalan...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
