import asyncio
import csv
import html
import logging
import os
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from telegram import (
    BotCommand,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Update,
)
from telegram.constants import ParseMode
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# ============================================================
# KONFIGURASI
# ============================================================

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
DB_PATH = os.getenv("DB_PATH", "data/surat_keluar.db").strip()
ALLOWED_USER_IDS_RAW = os.getenv("ALLOWED_USER_IDS", "").strip()
ADMIN_IDS_RAW = os.getenv("ADMIN_IDS", "").strip()
SEQUENCE_MODE = os.getenv("SEQUENCE_MODE", "global").strip().lower()
START_NUMBER = int(os.getenv("START_NUMBER", "1"))

if SEQUENCE_MODE not in {"global", "per_classification"}:
    raise ValueError("SEQUENCE_MODE harus 'global' atau 'per_classification'.")

def parse_id_list(raw: str) -> set[int]:
    result: set[int] = set()
    if not raw:
        return result
    for item in raw.split(","):
        item = item.strip()
        if item:
            result.add(int(item))
    return result

ALLOWED_USER_IDS = parse_id_list(ALLOWED_USER_IDS_RAW)
ADMIN_IDS = parse_id_list(ADMIN_IDS_RAW)

# ============================================================
# KLASIFIKASI SURAT
# Disalin dari tabel klasifikasi yang diberikan pengguna.
# ============================================================

CLASSIFICATIONS = {
    "500.13": "PARIWISATA DAN EKONOMI KREATIF",
    "500.13.1": "Kebijakan di bidang Pariwisata dan Ekonomi Kreatif yang dilakukan oleh Pemerintah Daerah",

    "500.13.2": "Pengembangan Destinasi Wisata",
    "500.13.2.1": "Perancangan Destinasi dan Investasi Pariwisata",
    "500.13.2.2": "Pengembangan Daya Tarik Wisata",
    "500.13.2.3": "Industri Pariwisata",
    "500.13.2.4": "Pemberdayaan Masyarakat Destinasi Pariwisata",
    "500.13.2.5": "Pengembangan Wisata Minat Khusus, Konvensi, Insentif, dan Event",

    "500.13.3": "Pemasaran Pariwisata",
    "500.13.3.1": "Pengembangan Pasar dan Informasi Pariwisata",
    "500.13.3.2": "Promosi Pariwisata Luar Negeri",
    "500.13.3.3": "Promosi Pariwisata Dalam Negeri",
    "500.13.3.4": "Pencitraan Indonesia",

    "500.13.4": "Ekonomi Kreatif Berbasis Seni dan Budaya",
    "500.13.4.1": "Pengembangan Industri Perfilman",
    "500.13.4.2": "Pengembangan Seni Pertunjukan dan Industri Musik",
    "500.13.4.3": "Pengembangan Seni Rupa",

    "500.13.5": "Ekonomi Kreatif Berbasis Media, Desain, dan IPTEK",
    "500.13.5.1": "Pengembangan Ekonomi Kreatif Berbasis Media",
    "500.13.5.2": "Desain dan Arsitektur",
    "500.13.5.3": "Kerjasama dan Fasilitasi",

    "500.13.6": "Pengembangan Sumber Daya Pariwisata dan Ekonomi Kreatif",
    "500.13.6.1": "Penelitian dan Pengembangan Kebijakan Kepariwisataan",
    "500.13.6.2": "Penelitian dan Pengembangan Kebijakan Ekonomi Kreatif",
    "500.13.6.3": "Pengembangan SDM Kepariwisataan dan Ekonomi Kreatif",
    "500.13.6.4": "Kompetensi Kepariwisataan dan Ekonomi Kreatif",
}

CATEGORY_CODES = [
    "500.13",
    "500.13.1",
    "500.13.2",
    "500.13.3",
    "500.13.4",
    "500.13.5",
    "500.13.6",
]

# Conversation states
WAIT_SUBJECT, WAIT_DESTINATION, CONFIRM = range(3)

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("bot-surat-keluar")


# ============================================================
# DATABASE
# ============================================================

def ensure_db_parent() -> None:
    parent = Path(DB_PATH).expanduser().resolve().parent
    parent.mkdir(parents=True, exist_ok=True)

def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    return conn

def init_db() -> None:
    ensure_db_parent()
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS counters (
                scope TEXT PRIMARY KEY,
                current_number INTEGER NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS letters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sequence_number INTEGER NOT NULL,
                counter_scope TEXT NOT NULL,
                classification_code TEXT NOT NULL,
                classification_name TEXT NOT NULL,
                letter_number TEXT NOT NULL UNIQUE,
                subject TEXT NOT NULL,
                destination TEXT,
                created_by INTEGER NOT NULL,
                created_by_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'ACTIVE'
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_letters_created_at
            ON letters(created_at DESC)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_letters_classification
            ON letters(classification_code)
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                detail TEXT NOT NULL,
                user_id INTEGER NOT NULL,
                user_name TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.commit()

def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")

def get_counter_scope(classification_code: str) -> str:
    if SEQUENCE_MODE == "per_classification":
        return classification_code
    return "GLOBAL"

def allocate_letter_number(
    classification_code: str,
    subject: str,
    destination: str,
    user_id: int,
    user_name: str,
) -> tuple[int, str]:
    """
    Alokasi nomor dilakukan dalam transaksi BEGIN IMMEDIATE agar dua pengguna
    tidak memperoleh nomor urut yang sama pada saat bersamaan.
    """
    scope = get_counter_scope(classification_code)
    conn = db_connect()
    try:
        conn.execute("BEGIN IMMEDIATE")

        row = conn.execute(
            "SELECT current_number FROM counters WHERE scope = ?",
            (scope,),
        ).fetchone()

        if row is None:
            next_number = START_NUMBER
            conn.execute(
                "INSERT INTO counters(scope, current_number) VALUES (?, ?)",
                (scope, next_number),
            )
        else:
            next_number = int(row["current_number"]) + 1
            conn.execute(
                "UPDATE counters SET current_number = ? WHERE scope = ?",
                (next_number, scope),
            )

        if next_number > 99999:
            raise ValueError("Nomor urut sudah melebihi batas 5 digit (99999).")

        letter_number = f"{classification_code}/{next_number:05d}"
        classification_name = CLASSIFICATIONS[classification_code]

        conn.execute("""
            INSERT INTO letters (
                sequence_number,
                counter_scope,
                classification_code,
                classification_name,
                letter_number,
                subject,
                destination,
                created_by,
                created_by_name,
                created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            next_number,
            scope,
            classification_code,
            classification_name,
            letter_number,
            subject,
            destination,
            user_id,
            user_name,
            now_iso(),
        ))

        conn.execute("""
            INSERT INTO audit_log(action, detail, user_id, user_name, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (
            "CREATE_NUMBER",
            letter_number,
            user_id,
            user_name,
            now_iso(),
        ))

        conn.commit()
        return next_number, letter_number
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def get_last_letters(limit: int = 10) -> list[sqlite3.Row]:
    with db_connect() as conn:
        return conn.execute("""
            SELECT * FROM letters
            ORDER BY id DESC
            LIMIT ?
        """, (limit,)).fetchall()

def get_last_number() -> Optional[sqlite3.Row]:
    with db_connect() as conn:
        return conn.execute("""
            SELECT * FROM letters
            ORDER BY id DESC
            LIMIT 1
        """).fetchone()

def search_letters(term: str, limit: int = 15) -> list[sqlite3.Row]:
    pattern = f"%{term}%"
    with db_connect() as conn:
        return conn.execute("""
            SELECT * FROM letters
            WHERE letter_number LIKE ?
               OR subject LIKE ?
               OR destination LIKE ?
               OR classification_code LIKE ?
               OR classification_name LIKE ?
            ORDER BY id DESC
            LIMIT ?
        """, (pattern, pattern, pattern, pattern, pattern, limit)).fetchall()

def get_max_sequence_global() -> int:
    with db_connect() as conn:
        row = conn.execute("""
            SELECT COALESCE(MAX(sequence_number), 0) AS max_seq
            FROM letters
            WHERE counter_scope = 'GLOBAL'
        """).fetchone()
        return int(row["max_seq"])

def set_next_global_number(next_number: int, user_id: int, user_name: str) -> None:
    if SEQUENCE_MODE != "global":
        raise ValueError("Perintah ini hanya tersedia saat SEQUENCE_MODE=global.")
    if not 1 <= next_number <= 99999:
        raise ValueError("Nomor berikutnya harus antara 1 dan 99999.")

    max_used = get_max_sequence_global()
    if next_number <= max_used:
        raise ValueError(
            f"Nomor berikutnya harus lebih besar dari nomor yang sudah pernah dipakai ({max_used:05d})."
        )

    current_number = next_number - 1
    with db_connect() as conn:
        conn.execute("""
            INSERT INTO counters(scope, current_number)
            VALUES ('GLOBAL', ?)
            ON CONFLICT(scope) DO UPDATE SET current_number = excluded.current_number
        """, (current_number,))
        conn.execute("""
            INSERT INTO audit_log(action, detail, user_id, user_name, created_at)
            VALUES (?, ?, ?, ?, ?)
        """, (
            "SET_NEXT_NUMBER",
            f"Next global number = {next_number:05d}",
            user_id,
            user_name,
            now_iso(),
        ))
        conn.commit()

def export_csv_file() -> str:
    fd, path = tempfile.mkstemp(prefix="surat_keluar_", suffix=".csv")
    os.close(fd)

    with db_connect() as conn:
        rows = conn.execute("""
            SELECT
                letter_number,
                classification_code,
                classification_name,
                subject,
                destination,
                created_by_name,
                created_by,
                created_at,
                status
            FROM letters
            ORDER BY id ASC
        """).fetchall()

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Nomor Surat",
            "Kode Klasifikasi",
            "Nama Klasifikasi",
            "Perihal",
            "Tujuan",
            "Dibuat Oleh",
            "Telegram ID",
            "Tanggal",
            "Status",
        ])
        for row in rows:
            writer.writerow([
                row["letter_number"],
                row["classification_code"],
                row["classification_name"],
                row["subject"],
                row["destination"] or "",
                row["created_by_name"],
                row["created_by"],
                row["created_at"],
                row["status"],
            ])
    return path


# ============================================================
# UTILITAS TELEGRAM
# ============================================================

def user_display_name(update: Update) -> str:
    user = update.effective_user
    if not user:
        return "Unknown"
    return user.full_name or user.username or str(user.id)

async def access_allowed(update: Update) -> bool:
    user = update.effective_user
    if not user:
        return False

    # Jika ALLOWED_USER_IDS kosong, bot terbuka untuk pengujian.
    # Untuk produksi, isi ALLOWED_USER_IDS di environment variable.
    if ALLOWED_USER_IDS and user.id not in ALLOWED_USER_IDS and user.id not in ADMIN_IDS:
        if update.callback_query:
            await update.callback_query.answer("Akses ditolak.", show_alert=True)
        elif update.effective_message:
            await update.effective_message.reply_text(
                "⛔ Anda belum terdaftar sebagai pengguna bot.\n\n"
                f"Telegram ID Anda: <code>{user.id}</code>\n"
                "Hubungi admin agar ID ini ditambahkan.",
                parse_mode=ParseMode.HTML,
            )
        return False
    return True

def is_admin(update: Update) -> bool:
    user = update.effective_user
    return bool(user and user.id in ADMIN_IDS)

def esc(text: Optional[str]) -> str:
    return html.escape(text or "-")

def short_label(text: str, max_len: int = 46) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"

def main_menu() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("➕ Buat Nomor Surat", callback_data="menu:new")],
        [
            InlineKeyboardButton("🕘 Riwayat Terakhir", callback_data="menu:history"),
            InlineKeyboardButton("🔢 Nomor Terakhir", callback_data="menu:last"),
        ],
        [InlineKeyboardButton("🗂 Daftar Klasifikasi", callback_data="menu:classes")],
        [InlineKeyboardButton("🆔 ID Telegram Saya", callback_data="menu:myid")],
    ])

def category_keyboard() -> InlineKeyboardMarkup:
    rows = []
    for code in CATEGORY_CODES:
        rows.append([
            InlineKeyboardButton(
                f"{code} • {short_label(CLASSIFICATIONS[code], 34)}",
                callback_data=f"cat:{code}",
            )
        ])
    rows.append([InlineKeyboardButton("⬅️ Kembali", callback_data="menu:home")])
    return InlineKeyboardMarkup(rows)

def classification_keyboard(category_code: str) -> InlineKeyboardMarkup:
    rows = []

    # Selalu izinkan pemilihan kode induk itu sendiri.
    rows.append([
        InlineKeyboardButton(
            f"✅ Gunakan {category_code}",
            callback_data=f"class:{category_code}",
        )
    ])

    prefix = category_code + "."
    direct_children = []
    parent_depth = category_code.count(".")
    for code, name in CLASSIFICATIONS.items():
        if code.startswith(prefix) and code.count(".") == parent_depth + 1:
            direct_children.append((code, name))

    for code, name in direct_children:
        rows.append([
            InlineKeyboardButton(
                f"{code} • {short_label(name, 32)}",
                callback_data=f"class:{code}",
            )
        ])

    rows.append([InlineKeyboardButton("⬅️ Pilih Kelompok Lain", callback_data="menu:new")])
    return InlineKeyboardMarkup(rows)

def confirm_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Terbitkan Nomor", callback_data="confirm:yes"),
            InlineKeyboardButton("❌ Batalkan", callback_data="confirm:no"),
        ]
    ])


# ============================================================
# HANDLERS
# ============================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return

    text = (
        "📨 <b>BOT PENOMORAN SURAT KELUAR BIDANG</b>\n\n"
        "Bot ini membuat nomor surat dengan format:\n"
        "<code>KODE_KLASIFIKASI/00001</code>\n\n"
        "Contoh:\n"
        "<code>500.13.3.1/00001</code>\n\n"
        "Nomor 5 digit dibuat otomatis dan disimpan ke database."
    )
    await update.effective_message.reply_text(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=main_menu(),
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return

    text = (
        "ℹ️ <b>Panduan Singkat</b>\n\n"
        "• /start — buka menu utama\n"
        "• /baru — buat nomor surat baru\n"
        "• /riwayat — lihat 10 nomor terakhir\n"
        "• /terakhir — lihat nomor terakhir\n"
        "• /cari kata — cari nomor/perihal/tujuan\n"
        "• /id — lihat Telegram ID\n"
        "• /batal — batalkan proses input\n\n"
        "<b>Admin:</b>\n"
        "• /setnomor 123 — set nomor berikutnya menjadi 00123\n"
        "• /export — ekspor seluruh register ke CSV"
    )
    await update.effective_message.reply_text(text, parse_mode=ParseMode.HTML)

async def my_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    if not user:
        return
    await update.effective_message.reply_text(
        f"🆔 Telegram ID Anda: <code>{user.id}</code>",
        parse_mode=ParseMode.HTML,
    )

async def menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> Optional[int]:
    if not await access_allowed(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "menu:home":
        await query.edit_message_text(
            "📨 <b>BOT PENOMORAN SURAT KELUAR BIDANG</b>\n\nSilakan pilih menu:",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu(),
        )
        return ConversationHandler.END

    if data == "menu:new":
        context.user_data.clear()
        await query.edit_message_text(
            "🗂 <b>Pilih kelompok klasifikasi surat:</b>",
            parse_mode=ParseMode.HTML,
            reply_markup=category_keyboard(),
        )
        return ConversationHandler.END

    if data == "menu:history":
        await show_history_message(query.edit_message_text)
        return ConversationHandler.END

    if data == "menu:last":
        await show_last_message(query.edit_message_text)
        return ConversationHandler.END

    if data == "menu:classes":
        await query.edit_message_text(
            build_classification_list(),
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ Kembali", callback_data="menu:home")]
            ]),
        )
        return ConversationHandler.END

    if data == "menu:myid":
        user = update.effective_user
        await query.edit_message_text(
            f"🆔 Telegram ID Anda: <code>{user.id}</code>",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ Kembali", callback_data="menu:home")]
            ]),
        )
        return ConversationHandler.END

    return ConversationHandler.END

async def category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return

    query = update.callback_query
    await query.answer()
    code = query.data.split(":", 1)[1]

    await query.edit_message_text(
        f"🗂 <b>{esc(code)} — {esc(CLASSIFICATIONS[code])}</b>\n\n"
        "Pilih kode yang akan digunakan:",
        parse_mode=ParseMode.HTML,
        reply_markup=classification_keyboard(code),
    )

async def new_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    context.user_data.clear()
    await update.effective_message.reply_text(
        "🗂 <b>Pilih kelompok klasifikasi surat:</b>",
        parse_mode=ParseMode.HTML,
        reply_markup=category_keyboard(),
    )
    return ConversationHandler.END

async def class_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()

    code = query.data.split(":", 1)[1]
    context.user_data["classification_code"] = code
    context.user_data["classification_name"] = CLASSIFICATIONS[code]

    await query.edit_message_text(
        f"✅ <b>Klasifikasi dipilih</b>\n"
        f"<code>{esc(code)}</code>\n"
        f"{esc(CLASSIFICATIONS[code])}\n\n"
        "📝 Ketik <b>perihal surat</b>:",
        parse_mode=ParseMode.HTML,
    )
    return WAIT_SUBJECT

async def receive_subject(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    subject = (update.effective_message.text or "").strip()
    if len(subject) < 3:
        await update.effective_message.reply_text(
            "Perihal terlalu pendek. Silakan ketik perihal surat yang lebih jelas."
        )
        return WAIT_SUBJECT

    context.user_data["subject"] = subject

    await update.effective_message.reply_text(
        "🎯 Ketik <b>tujuan/penerima surat</b>.\n"
        "Atau tekan <b>Lewati</b> jika tidak ingin dicatat.",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⏭ Lewati", callback_data="destination:skip")]
        ]),
    )
    return WAIT_DESTINATION

async def skip_destination(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()
    context.user_data["destination"] = "-"
    await show_confirmation(query.edit_message_text, context)
    return CONFIRM

async def receive_destination(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    destination = (update.effective_message.text or "").strip()
    context.user_data["destination"] = destination or "-"
    await show_confirmation(update.effective_message.reply_text, context)
    return CONFIRM

async def show_confirmation(send_func, context: ContextTypes.DEFAULT_TYPE) -> None:
    code = context.user_data["classification_code"]
    name = context.user_data["classification_name"]
    subject = context.user_data["subject"]
    destination = context.user_data.get("destination", "-")

    text = (
        "🔎 <b>Konfirmasi Data Surat</b>\n\n"
        f"<b>Klasifikasi:</b>\n<code>{esc(code)}</code>\n{esc(name)}\n\n"
        f"<b>Perihal:</b>\n{esc(subject)}\n\n"
        f"<b>Tujuan:</b>\n{esc(destination)}\n\n"
        "Nomor 5 digit akan diterbitkan saat tombol konfirmasi ditekan."
    )
    await send_func(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=confirm_keyboard(),
    )

async def confirm_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await access_allowed(update):
        return ConversationHandler.END

    query = update.callback_query
    await query.answer()
    choice = query.data

    if choice == "confirm:no":
        context.user_data.clear()
        await query.edit_message_text(
            "❌ Pembuatan nomor dibatalkan.",
            reply_markup=main_menu(),
        )
        return ConversationHandler.END

    required = {"classification_code", "classification_name", "subject"}
    if not required.issubset(context.user_data):
        context.user_data.clear()
        await query.edit_message_text(
            "⚠️ Data proses sudah tidak lengkap. Silakan mulai lagi.",
            reply_markup=main_menu(),
        )
        return ConversationHandler.END

    code = context.user_data["classification_code"]
    subject = context.user_data["subject"]
    destination = context.user_data.get("destination", "-")
    user = update.effective_user

    try:
        sequence, letter_number = await asyncio.to_thread(
            allocate_letter_number,
            code,
            subject,
            destination,
            user.id,
            user_display_name(update),
        )
    except sqlite3.IntegrityError:
        logger.exception("Duplicate letter number detected")
        await query.edit_message_text(
            "⚠️ Terjadi benturan nomor. Silakan coba lagi.",
            reply_markup=main_menu(),
        )
        return ConversationHandler.END
    except Exception as exc:
        logger.exception("Failed to allocate number")
        await query.edit_message_text(
            f"⚠️ Gagal membuat nomor: {esc(str(exc))}",
            parse_mode=ParseMode.HTML,
            reply_markup=main_menu(),
        )
        return ConversationHandler.END

    text = (
        "✅ <b>NOMOR SURAT BERHASIL DITERBITKAN</b>\n\n"
        f"📌 <b>Nomor:</b>\n"
        f"<code>{esc(letter_number)}</code>\n\n"
        f"🗂 <b>Klasifikasi:</b>\n"
        f"{esc(CLASSIFICATIONS[code])}\n\n"
        f"📝 <b>Perihal:</b>\n"
        f"{esc(subject)}\n\n"
        f"🎯 <b>Tujuan:</b>\n"
        f"{esc(destination)}\n\n"
        f"👤 <b>Dibuat oleh:</b> {esc(user_display_name(update))}\n"
        f"🕒 <b>Waktu:</b> {esc(datetime.now().astimezone().strftime('%d-%m-%Y %H:%M:%S'))}"
    )

    context.user_data.clear()
    await query.edit_message_text(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=main_menu(),
    )
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.effective_message.reply_text(
        "❌ Proses dibatalkan.",
        reply_markup=main_menu(),
    )
    return ConversationHandler.END

def format_letter_row(row: sqlite3.Row) -> str:
    created = row["created_at"]
    try:
        created_dt = datetime.fromisoformat(created)
        created_text = created_dt.strftime("%d-%m-%Y %H:%M")
    except ValueError:
        created_text = created

    return (
        f"📌 <code>{esc(row['letter_number'])}</code>\n"
        f"📝 {esc(row['subject'])}\n"
        f"🎯 {esc(row['destination'] or '-')}\n"
        f"👤 {esc(row['created_by_name'])} • {esc(created_text)}"
    )

async def show_history_message(send_func) -> None:
    rows = await asyncio.to_thread(get_last_letters, 10)
    if not rows:
        text = "📭 Belum ada nomor surat yang tersimpan."
    else:
        parts = ["🕘 <b>10 NOMOR SURAT TERAKHIR</b>"]
        for i, row in enumerate(rows, start=1):
            parts.append(f"\n<b>{i}.</b> {format_letter_row(row)}")
        text = "\n".join(parts)

    await send_func(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Kembali", callback_data="menu:home")]
        ]),
    )

async def history_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return
    await show_history_message(update.effective_message.reply_text)

async def show_last_message(send_func) -> None:
    row = await asyncio.to_thread(get_last_number)
    if not row:
        text = "📭 Belum ada nomor surat yang tersimpan."
    else:
        text = (
            "🔢 <b>NOMOR SURAT TERAKHIR</b>\n\n"
            + format_letter_row(row)
        )

    await send_func(
        text,
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ Kembali", callback_data="menu:home")]
        ]),
    )

async def last_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return
    await show_last_message(update.effective_message.reply_text)

async def search_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return

    term = " ".join(context.args).strip()
    if not term:
        await update.effective_message.reply_text(
            "Gunakan format:\n<code>/cari kata_kunci</code>\n\n"
            "Contoh:\n<code>/cari promosi</code>\n"
            "<code>/cari 500.13.3.1</code>\n"
            "<code>/cari 00025</code>",
            parse_mode=ParseMode.HTML,
        )
        return

    rows = await asyncio.to_thread(search_letters, term, 15)
    if not rows:
        await update.effective_message.reply_text(
            f"🔎 Tidak ada hasil untuk: <b>{esc(term)}</b>",
            parse_mode=ParseMode.HTML,
        )
        return

    parts = [f"🔎 <b>HASIL PENCARIAN: {esc(term)}</b>"]
    for i, row in enumerate(rows, start=1):
        parts.append(f"\n<b>{i}.</b> {format_letter_row(row)}")

    await update.effective_message.reply_text(
        "\n".join(parts),
        parse_mode=ParseMode.HTML,
    )

def build_classification_list() -> str:
    lines = ["🗂 <b>DAFTAR KLASIFIKASI</b>\n"]
    for code, name in CLASSIFICATIONS.items():
        lines.append(f"<code>{esc(code)}</code> — {esc(name)}")
    return "\n".join(lines)

async def classes_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return
    await update.effective_message.reply_text(
        build_classification_list(),
        parse_mode=ParseMode.HTML,
    )

async def set_number_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return
    if not is_admin(update):
        await update.effective_message.reply_text("⛔ Perintah ini hanya untuk admin.")
        return

    if len(context.args) != 1 or not context.args[0].isdigit():
        await update.effective_message.reply_text(
            "Format:\n<code>/setnomor 123</code>\n\n"
            "Artinya nomor berikutnya yang diterbitkan adalah <code>00123</code>.",
            parse_mode=ParseMode.HTML,
        )
        return

    next_number = int(context.args[0])
    user = update.effective_user

    try:
        await asyncio.to_thread(
            set_next_global_number,
            next_number,
            user.id,
            user_display_name(update),
        )
    except Exception as exc:
        await update.effective_message.reply_text(
            f"⚠️ Gagal: {esc(str(exc))}",
            parse_mode=ParseMode.HTML,
        )
        return

    await update.effective_message.reply_text(
        f"✅ Nomor berikutnya diset menjadi <code>{next_number:05d}</code>.",
        parse_mode=ParseMode.HTML,
    )

async def export_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not await access_allowed(update):
        return
    if not is_admin(update):
        await update.effective_message.reply_text("⛔ Perintah ini hanya untuk admin.")
        return

    path = await asyncio.to_thread(export_csv_file)
    try:
        with open(path, "rb") as f:
            await update.effective_message.reply_document(
                document=f,
                filename=f"register_surat_keluar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                caption="📊 Export register nomor surat keluar.",
            )
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

async def unknown_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        "Perintah tidak dikenali. Ketik /help untuk melihat panduan."
    )

async def post_init(application: Application) -> None:
    commands = [
        BotCommand("start", "Buka menu utama"),
        BotCommand("baru", "Buat nomor surat baru"),
        BotCommand("riwayat", "Lihat 10 nomor terakhir"),
        BotCommand("terakhir", "Lihat nomor terakhir"),
        BotCommand("cari", "Cari register surat"),
        BotCommand("klasifikasi", "Lihat daftar klasifikasi"),
        BotCommand("id", "Lihat Telegram ID"),
        BotCommand("help", "Panduan penggunaan"),
        BotCommand("batal", "Batalkan proses input"),
    ]
    await application.bot.set_my_commands(commands)

def build_application() -> Application:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN belum diisi.")

    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    conversation = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(class_callback, pattern=r"^class:"),
        ],
        states={
            WAIT_SUBJECT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_subject),
            ],
            WAIT_DESTINATION: [
                CallbackQueryHandler(skip_destination, pattern=r"^destination:skip$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_destination),
            ],
            CONFIRM: [
                CallbackQueryHandler(confirm_callback, pattern=r"^confirm:"),
            ],
        },
        fallbacks=[
            CommandHandler("batal", cancel),
            CommandHandler("start", start),
        ],
        allow_reentry=True,
    )

    # Command handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("id", my_id))
    app.add_handler(CommandHandler("baru", new_command))
    app.add_handler(CommandHandler("riwayat", history_command))
    app.add_handler(CommandHandler("terakhir", last_command))
    app.add_handler(CommandHandler("cari", search_command))
    app.add_handler(CommandHandler("klasifikasi", classes_command))
    app.add_handler(CommandHandler("setnomor", set_number_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CommandHandler("batal", cancel))

    # Conversation harus dipasang sebelum generic callback handlers.
    app.add_handler(conversation)

    # Callback handlers
    app.add_handler(CallbackQueryHandler(menu_callback, pattern=r"^menu:"))
    app.add_handler(CallbackQueryHandler(category_callback, pattern=r"^cat:"))

    # Unknown commands
    app.add_handler(MessageHandler(filters.COMMAND, unknown_command))

    return app

def main() -> None:
    init_db()
    app = build_application()
    logger.info("Bot started. DB=%s | SEQUENCE_MODE=%s", DB_PATH, SEQUENCE_MODE)
    app.run_polling(
        allowed_updates=Update.ALL_TYPES,
        drop_pending_updates=False,
    )

if __name__ == "__main__":
    main()
