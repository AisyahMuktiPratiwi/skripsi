import json
import os
import re
import hashlib
import sqlite3
import time
from datetime import timedelta, datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from docx import Document
from flask import (
    Flask, flash, redirect, render_template_string, request, send_from_directory, session, url_for,
)
from PyPDF2 import PdfReader
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

# ==================== KONFIGURASI ====================
BASE_DIR = Path(__file__).resolve().parent

# Muat variabel dari file .env (kalau ada) supaya tidak perlu set env var manual tiap buka terminal.
# PENTING: load_dotenv() TANPA path mencari .env berdasarkan current working directory proses
# yang menjalankan app ini — bukan lokasi file app.py. Di local biasanya cwd = folder project
# (karena dijalankan lewat `python app.py` dari situ), jadi ketemu. Tapi di PythonAnywhere,
# proses WSGI yang menjalankan app cwd-nya BEDA (bukan folder project), jadi .env gak pernah
# kebaca -> semua os.environ.get(...) balik ke default kosong (makanya FONNTE_API_KEY kosong
# hanya di PythonAnywhere, padahal di local jalan normal). Kasih path eksplisit ke BASE_DIR/.env
# supaya selalu ketemu, gak peduli proses ini dijalankan dari direktori mana pun.
try:
    from dotenv import load_dotenv
    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

UPLOAD_FOLDER = BASE_DIR / "uploads"
DB_PATH = BASE_DIR / "similarity.db"
UPLOAD_FOLDER.mkdir(exist_ok=True)

# Folder untuk aset statis (logo, dll). Taruh file logo di sini, contoh: static/logo.png
STATIC_FOLDER = BASE_DIR / "static"
STATIC_FOLDER.mkdir(exist_ok=True)
LOGO_FILENAME = "api.png"

ALLOWED_EXTENSIONS = {"txt", "pdf", "docx", "html", "htm"}
K_GRAM = 5
WINDOW_SIZE = 4

# Berapa hari tanpa upload sebelum streak dianggap padam
DAYS_BEFORE_WARNING = 3   # hari ke-3 tanpa upload -> peringatan
DAYS_BEFORE_DEATH = 4     # hari ke-4 tanpa upload -> streak reset & api padam
KEEPALIVE_START_DAYS = 5      # reminder susulan pertama dikirim di hari ke-5 sejak api padam
KEEPALIVE_INTERVAL_DAYS = 3   # setelah itu, reminder diulang tiap 3 hari
REMINDER_STOP_AFTER_DAYS = 90  # ~3 bulan sejak padam, reminder dihentikan total

# Konfigurasi Fonnte WhatsApp API
FONNTE_API_KEY = os.environ.get("FONNTE_API_KEY", "")
FONNTE_BASE_URL = "https://api.fonnte.com"
ADMIN_CRON_TOKEN = os.environ.get("ADMIN_CRON_TOKEN", "ganti-token-ini")

REMINDER_HOUR = int(os.environ.get("REMINDER_HOUR", "8"))
REMINDER_MINUTE = int(os.environ.get("REMINDER_MINUTE", "0"))

app = Flask(__name__)
app.secret_key = os.environ.get(
    "APP_SECRET_KEY",
    "8ea93c21ee2b0060a2fb9dbd5f61c783ce940f653db8beb87809dcf6bac901c9",
)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

INDO_MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
INDO_MONTHS_FULL = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
INDO_DAYS_FULL = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
INDO_DAYS_SHORT = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]

# Warna kurva berdasarkan jenis/kategori progres (selaras dengan get_progress_label)
PROGRESS_LABEL_COLORS = {
    "Perubahan sangat kecil": "#ff4d6d",
    "Progres ringan": "#ff8a00",
    "Progres sedang": "#ffb347",
    "Progres besar": "#34d399",
}
NO_UPLOAD_COLOR = "rgba(255,255,255,0.18)"
NO_UPLOAD_LABEL = "Belum upload"


# ==================== FUNGSI UTILITY ====================
def now_dt():
    return datetime.now()


def today_str():
    return now_dt().strftime("%Y-%m-%d")


def parse_date(date_str: str):
    return datetime.strptime(date_str, "%Y-%m-%d").date()


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def normalize_phone_number(phone_number: str) -> str:
    """Normalisasi nomor telepon ke format Fonnte (62xxxxxxxxxx)."""
    digits = re.sub(r"[^0-9]", "", phone_number or "")
    if not digits:
        return ""
    if digits.startswith("0"):
        digits = "62" + digits[1:]
    elif not digits.startswith("62"):
        digits = "62" + digits
    return digits


# ==================== INISIALISASI DATABASE ====================
def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            phone_number TEXT,
            created_at TEXT NOT NULL,
            welcome_sent INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            original_name TEXT NOT NULL,
            saved_name TEXT NOT NULL,
            file_ext TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            upload_date TEXT NOT NULL,
            raw_text TEXT,
            clean_text TEXT,
            kgrams_json TEXT,
            hashes_json TEXT,
            windows_json TEXT,
            fingerprints_json TEXT,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS comparisons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            doc1_id INTEGER NOT NULL,
            doc2_id INTEGER NOT NULL,
            similarity_percent REAL NOT NULL,
            difference_percent REAL NOT NULL,
            process_time REAL NOT NULL,
            progress_label TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(doc1_id, doc2_id),
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        """
    )
    # streak_state: status "api"/streak per user
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS streak_state (
            user_id INTEGER PRIMARY KEY,
            streak_count INTEGER NOT NULL DEFAULT 0,
            last_valid_upload_date TEXT,
            flame_status TEXT NOT NULL DEFAULT 'off',
            flame_off_date TEXT,
            last_reminder_type TEXT,
            last_reminder_date TEXT,
            thesis_finished INTEGER NOT NULL DEFAULT 0,
            first_upload_notified INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS reminder_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            reminder_type TEXT NOT NULL,
            sent_at TEXT NOT NULL,
            phone_number TEXT NOT NULL,
            status TEXT NOT NULL,
            response TEXT,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        """
    )
    conn.commit()

    # ---- migrasi kolom baru (untuk database lama hasil versi sebelumnya) ----
    existing_streak_cols = [row[1] for row in cur.execute("PRAGMA table_info(streak_state)").fetchall()]
    streak_migrations = {
        "flame_off_date": "ALTER TABLE streak_state ADD COLUMN flame_off_date TEXT",
        "last_reminder_type": "ALTER TABLE streak_state ADD COLUMN last_reminder_type TEXT",
        "last_reminder_date": "ALTER TABLE streak_state ADD COLUMN last_reminder_date TEXT",
        "thesis_finished": "ALTER TABLE streak_state ADD COLUMN thesis_finished INTEGER NOT NULL DEFAULT 0",
        "first_upload_notified": "ALTER TABLE streak_state ADD COLUMN first_upload_notified INTEGER NOT NULL DEFAULT 0",
    }
    for col, ddl in streak_migrations.items():
        if col not in existing_streak_cols:
            try:
                cur.execute(ddl)
            except sqlite3.OperationalError:
                pass

    existing_user_cols = [row[1] for row in cur.execute("PRAGMA table_info(users)").fetchall()]
    user_migrations = {
        "welcome_sent": "ALTER TABLE users ADD COLUMN welcome_sent INTEGER NOT NULL DEFAULT 0",
    }
    for col, ddl in user_migrations.items():
        if col not in existing_user_cols:
            try:
                cur.execute(ddl)
            except sqlite3.OperationalError:
                pass

    conn.commit()
    conn.close()


# ==================== FUNGSI AUTENTIKASI ====================
def get_user_by_username(username):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return row


def get_user_by_id(user_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    return row


def create_user(username, email, password, phone_number=None):
    conn = get_conn()
    cur = conn.cursor()
    password_hash = generate_password_hash(password)
    created_at = now_dt().strftime("%Y-%m-%d %H:%M:%S")
    try:
        cur.execute(
            """
            INSERT INTO users (username, email, password_hash, phone_number, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (username, email, password_hash, phone_number, created_at),
        )
        user_id = cur.lastrowid
        # Requirement: "pertama daftar, pertama streak" -> inisialisasi streak_state
        cur.execute(
            """
            INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
            VALUES (?, 0, NULL, 'off')
            """,
            (user_id,),
        )
        conn.commit()
        conn.close()
        return user_id
    except sqlite3.IntegrityError:
        conn.close()
        raise


# ==================== FUNGSI WHATSAPP (FONNTE) ====================
def send_whatsapp_reminder(phone_number, message):
    """Mengirim pesan WhatsApp menggunakan Fonnte API."""
    if not FONNTE_API_KEY or FONNTE_API_KEY == "YOUR_FONNTE_API_KEY":
        return {"status": "error", "message": "API Key Fonnte belum dikonfigurasi"}
    url = f"{FONNTE_BASE_URL}/send"
    headers = {"Authorization": FONNTE_API_KEY}
    data = {
        "target": phone_number,
        "message": message,
        "countryCode": "62",
    }
    try:
        response = requests.post(url, headers=headers, data=data, timeout=15)
        return {
            "status": "success" if response.status_code == 200 else "error",
            "response": response.json() if response.status_code == 200 else response.text,
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}


def log_reminder(user_id, reminder_type, phone_number, status, response=None):
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO reminder_log (user_id, reminder_type, sent_at, phone_number, status, response)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (user_id, reminder_type, now_dt().strftime("%Y-%m-%d %H:%M:%S"), phone_number, status, response),
    )
    conn.commit()
    conn.close()


def send_direct_message(user_id, phone_number, reminder_type, message):
    """Kirim pesan WA di luar alur sync_streak_and_notify (welcome, first upload, thesis finished, dll)."""
    if not phone_number:
        log_reminder(user_id, reminder_type, "-", "skipped_no_phone", None)
        return
    result = send_whatsapp_reminder(phone_number, message)
    status = result.get("status", "error")
    log_reminder(user_id, reminder_type, phone_number, status, json.dumps(result))


# ==================== PESAN-PESAN WHATSAPP (HUMANIS & MEMOTIVASI) ====================
def msg_welcome(username):
    return (
        f"Halo {username}! \U0001F44B\U0001F525\n\n"
        "Selamat bergabung di Skripsiku! Aku bakal nemenin & ingetin kamu biar konsisten "
        "ngerjain skripsi, pelan-pelan nggak apa, yang penting jalan terus.\n\n"
        "Yuk upload progres pertamamu kapan pun kamu siap, dan mulai nyalain semangatnya! \U0001F525"
    )


def msg_first_upload(username):
    return (
        f"Mantap, {username}! \U0001F389\U0001F525\n\n"
        "Ini upload progres pertamamu, langkah paling berat udah kamu lewati. Aku bakal ingetin "
        "kamu lewat WA biar tetap konsisten.\n\n"
        "Info: kalau nanti skripsimu sudah selesai, tinggal klik tombol 'Skripsi Selesai' di "
        "dashboard buat berhenti reminder. Semangat terus! \U0001F4AA"
    )


def msg_warn_day4(username, gap):
    return (
        f"Halo {username} \U0001F525\n\n"
        f"Sudah {gap} hari belum upload progres. Besok hari terakhir sebelum streak-mu reset "
        "dan api padam.\n\n"
        "Yuk luangkan waktu sebentar hari ini, walau kecil tetap progres. Semangat! \U0001F4AA"
    )


def msg_dead(username, gap):
    return (
        f"Halo {username}\n\n"
        f"Streak-mu terhenti setelah {gap} hari tanpa progres, apinya padam sementara. Ini "
        "bukan akhir, cuma jeda.\n\n"
        "Upload kapan pun kamu siap, kita nyalakan lagi apinya bareng-bareng. \U0001F525 Kalau "
        "skripsimu sudah selesai, klik 'Skripsi Selesai' di dashboard ya."
    )


def msg_low_progress(username, avg_diff):
    return (
        f"Halo {username} \U0001F525\n\n"
        f"Progres 1-3 hari terakhir masih kecil (rata-rata {avg_diff}%), tapi tenang, streak-mu "
        "masih aman.\n\n"
        "Coba tambah sedikit lagi hari ini, progres kecil tetap progres. Semangat terus! \U0001F4AA"
    )


def msg_reactivated(username):
    return (
        f"Selamat datang kembali, {username}! \U0001F525\U0001F525\n\n"
        "Api semangatmu menyala lagi, streak mulai dari hari ke-1. Nggak masalah sempat vakum, "
        "yang penting kamu lanjut lagi.\n\n"
        "Ayo selesaikan skripsinya sampai tuntas! \U0001F4AA"
    )


def msg_keepalive(username, days_off):
    return (
        f"Halo {username}\n\n"
        f"Sudah {days_off} hari sejak progres terakhirmu. Nggak perlu langsung banyak, mulai "
        "dari yang kecil aja hari ini.\n\n"
        "Kalau sudah selesai, klik 'Skripsi Selesai' di dashboard ya. Kalau belum, semangat "
        "terus, aku masih nemenin kamu! \U0001F525"
    )


def msg_thesis_finished(username):
    return (
        f"Selamat, {username}! \U0001F393\U0001F389\n\n"
        "Skripsimu resmi ditandai SELESAI, kerja kerasmu selama ini luar biasa. Reminder "
        "progres aku hentikan mulai sekarang.\n\n"
        "Semoga sidang lancar, selamat menuju wisuda! \U0001F973"
    )


def msg_thesis_resumed(username):
    return (
        f"Halo lagi, {username}! \U0001F525\n\n"
        "Kamu lanjutkan lagi progres skripsimu, reminder aku aktifkan lagi biar tetap ada yang "
        "ingetin.\n\n"
        "Apapun alasannya, yang penting tetap melangkah. Semangat! \U0001F4AA"
    )


# ==================== FUNGSI STREAK & REMINDER ====================
def _get_or_create_streak_row(cur, user_id):
    row = cur.execute("SELECT * FROM streak_state WHERE user_id = ?", (user_id,)).fetchone()
    if not row:
        cur.execute(
            """
            INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
            VALUES (?, 0, NULL, 'off')
            """,
            (user_id,),
        )
        row = cur.execute("SELECT * FROM streak_state WHERE user_id = ?", (user_id,)).fetchone()
    return row


def sync_streak_and_notify(user_id):
    """Cek status streak user, perbarui bila perlu, dan kirim WA notifikasi sesuai aturan."""
    conn = get_conn()
    cur = conn.cursor()
    row = _get_or_create_streak_row(cur, user_id)
    conn.commit()

    # Kalau skripsi sudah ditandai selesai, jangan kirim reminder apa pun lagi.
    if row["thesis_finished"]:
        conn.close()
        return

    user = cur.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    phone = user["phone_number"] if user else None
    today = today_str()
    today_date = parse_date(today)

    flame_status = row["flame_status"]
    last_valid = row["last_valid_upload_date"]
    flame_off_date = row["flame_off_date"]
    last_reminder_type = row["last_reminder_type"]
    last_reminder_date = row["last_reminder_date"]

    state = {"last_reminder_type": last_reminder_type, "last_reminder_date": last_reminder_date}

    def already_sent_today(rtype):
        return state["last_reminder_type"] == rtype and state["last_reminder_date"] == today

    def send_and_log(rtype, message):
        if not phone:
            log_reminder(user_id, rtype, "-", "skipped_no_phone", None)
            return
        result = send_whatsapp_reminder(phone, message)
        status = result.get("status", "error")
        log_reminder(user_id, rtype, phone, status, json.dumps(result))
        cur.execute(
            "UPDATE streak_state SET last_reminder_type = ?, last_reminder_date = ? WHERE user_id = ?",
            (rtype, today, user_id),
        )
        conn.commit()
        state["last_reminder_type"] = rtype
        state["last_reminder_date"] = today

    username = user["username"] if user else "Sobat Skripsi"

    if flame_status == "on" and last_valid:
        gap = (today_date - parse_date(last_valid)).days
        if gap >= DAYS_BEFORE_DEATH:
            # ---- streak resmi padam ----
            cur.execute(
                """
                UPDATE streak_state
                SET streak_count = 0, flame_status = 'off', flame_off_date = ?
                WHERE user_id = ?
                """,
                (today, user_id),
            )
            conn.commit()
            flame_status = "off"
            flame_off_date = today
            if not already_sent_today("dead"):
                send_and_log("dead", msg_dead(username, gap))
        elif gap == DAYS_BEFORE_WARNING:
            # ---- peringatan H-1 sebelum padam ----
            if not already_sent_today("warn_day4"):
                send_and_log("warn_day4", msg_warn_day4(username, gap))

        # ---- progress kecil dalam 1-3 hari terakhir, streak tetap aman ----
        if flame_status == "on":
            highlight = get_recent_progress_highlight_by_user(user_id)
            if highlight and highlight["label"] == "Perubahan sangat kecil" and not already_sent_today("low_progress"):
                send_and_log("low_progress", msg_low_progress(username, highlight["avg_difference"]))

    elif flame_status == "off" and flame_off_date:
        # ---- reminder berkala setelah padam: mulai hari ke-5, lalu tiap 3 hari, stop ~3 bulan ----
        days_off = (today_date - parse_date(flame_off_date)).days
        if days_off >= REMINDER_STOP_AFTER_DAYS:
            pass  # reminder dihentikan permanen untuk user ini
        elif days_off >= KEEPALIVE_START_DAYS and (days_off - KEEPALIVE_START_DAYS) % KEEPALIVE_INTERVAL_DAYS == 0:
            if not already_sent_today("keepalive"):
                send_and_log("keepalive", msg_keepalive(username, days_off))

    conn.close()


def get_or_reset_streak_state(user_id):
    """Dipanggil dari halaman dashboard. Selalu sinkron sebelum mengembalikan data."""
    sync_streak_and_notify(user_id)
    conn = get_conn()
    row = conn.execute("SELECT * FROM streak_state WHERE user_id = ?", (user_id,)).fetchone()
    conn.close()
    return row


def update_streak_after_valid_upload(user_id):
    """Dipanggil setelah upload dokumen yang menghasilkan perbandingan valid (similarity < 100%)."""
    conn = get_conn()
    cur = conn.cursor()
    row = _get_or_create_streak_row(cur, user_id)
    conn.commit()

    was_off = row["flame_status"] == "off"
    was_finished = bool(row["thesis_finished"])
    streak_count = row["streak_count"]
    last_valid_upload_date = row["last_valid_upload_date"]
    is_first_ever_upload = last_valid_upload_date is None and not row["first_upload_notified"]
    today = today_str()

    if last_valid_upload_date == today:
        conn.close()
        return  # sudah upload valid hari ini, tidak perlu diproses ulang

    if last_valid_upload_date is None:
        streak_count = 1
    else:
        gap_days = (parse_date(today) - parse_date(last_valid_upload_date)).days
        if gap_days == 1:
            streak_count += 1
        else:
            streak_count = 1

    cur.execute(
        """
        UPDATE streak_state
        SET streak_count = ?, last_valid_upload_date = ?, flame_status = 'on',
            flame_off_date = NULL, last_reminder_type = NULL, last_reminder_date = NULL,
            thesis_finished = 0, first_upload_notified = 1
        WHERE user_id = ?
        """,
        (streak_count, today, user_id),
    )
    conn.commit()
    conn.close()

    user = get_user_by_id(user_id)
    phone = user["phone_number"] if user else None
    username = user["username"] if user else "Sobat Skripsi"

    if is_first_ever_upload:
        # ---- upload valid pertama sepanjang masa untuk user ini ----
        send_direct_message(user_id, phone, "first_upload", msg_first_upload(username))
    elif was_finished:
        # ---- user melanjutkan lagi setelah sebelumnya menandai skripsi selesai ----
        send_direct_message(user_id, phone, "thesis_resumed", msg_thesis_resumed(username))
    elif was_off:
        # ---- api kembali menyala setelah sebelumnya padam ----
        send_direct_message(user_id, phone, "reactivated", msg_reactivated(username))


def mark_thesis_finished(user_id):
    conn = get_conn()
    conn.execute(
        "UPDATE streak_state SET thesis_finished = 1, last_reminder_type = NULL, last_reminder_date = NULL WHERE user_id = ?",
        (user_id,),
    )
    conn.commit()
    conn.close()
    user = get_user_by_id(user_id)
    if user:
        send_direct_message(user_id, user["phone_number"], "thesis_finished", msg_thesis_finished(user["username"]))


def mark_thesis_resumed(user_id):
    conn = get_conn()
    conn.execute("UPDATE streak_state SET thesis_finished = 0 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    user = get_user_by_id(user_id)
    if user:
        send_direct_message(user_id, user["phone_number"], "thesis_resumed", msg_thesis_resumed(user["username"]))


def run_reminder_scheduler():
    """Dipanggil oleh cron job / scheduler harian untuk semua user."""
    conn = get_conn()
    user_ids = [r["user_id"] for r in conn.execute("SELECT user_id FROM streak_state").fetchall()]
    conn.close()
    print(f"[{now_dt()}] Memulai pengecekan reminder untuk {len(user_ids)} user...")
    for uid in user_ids:
        try:
            sync_streak_and_notify(uid)
        except Exception as e:
            print(f"[{now_dt()}] Gagal cek reminder user {uid}: {e}")
    print(f"[{now_dt()}] Pengecekan reminder selesai.")


# ==================== FUNGSI FILE PROCESSING (WINNOWING) ====================
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_from_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def extract_text_from_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    texts = [page.extract_text() or "" for page in reader.pages]
    return "\n".join(texts)


def extract_text_from_docx(path: Path) -> str:
    doc = Document(str(path))
    return "\n".join(paragraph.text for paragraph in doc.paragraphs)


def extract_text_from_html(path: Path) -> str:
    html = path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text(separator=" ")


def extract_text(path: Path, ext: str) -> str:
    ext = ext.lower()
    if ext == "txt":
        return extract_text_from_txt(path)
    if ext == "pdf":
        return extract_text_from_pdf(path)
    if ext == "docx":
        return extract_text_from_docx(path)
    if ext in {"html", "htm"}:
        return extract_text_from_html(path)
    raise ValueError(f"Format file tidak didukung: {ext}")


def preprocess_text(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace(" ", "")
    return text


def make_kgrams(text: str, k: int):
    if len(text) < k:
        return []
    return [text[i:i + k] for i in range(len(text) - k + 1)]


def stable_hash(kgram: str) -> int:
    return int(hashlib.md5(kgram.encode("utf-8")).hexdigest()[:8], 16)


def hash_kgrams(kgrams):
    return [stable_hash(kgram) for kgram in kgrams]


def make_windows(hashes, w: int):
    if len(hashes) < w:
        return []
    return [hashes[i:i + w] for i in range(len(hashes) - w + 1)]


def winnowing_fingerprints(hashes, w: int):
    windows = make_windows(hashes, w)
    fingerprints = []
    chosen = set()
    for i, window in enumerate(windows):
        min_hash = min(window)
        rightmost_index = max(idx for idx, value in enumerate(window) if value == min_hash)
        absolute_index = i + rightmost_index
        key = (min_hash, absolute_index)
        if key not in chosen:
            chosen.add(key)
            fingerprints.append({"hash": min_hash, "index": absolute_index})
    return windows, fingerprints


def calculate_similarity(fp1, fp2) -> float:
    set1 = {item["hash"] for item in fp1}
    set2 = {item["hash"] for item in fp2}
    if not set1 and not set2:
        return 100.0
    if not set1 or not set2:
        return 0.0
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    return round((intersection / union) * 100, 2)


def calculate_difference(similarity_percent: float) -> float:
    return round(100.0 - similarity_percent, 2)


def get_progress_label(difference_percent: float) -> str:
    # Kategori progres berdasarkan persentase difference (100% - similarity):
    #   0%   - 10%  -> "Perubahan sangat kecil"
    #   >10% - 30%  -> "Progres ringan"
    #   >30% - 60%  -> "Progres sedang"
    #   >60% - 100% -> "Progres besar"
    if difference_percent <= 10:
        return "Perubahan sangat kecil"
    if difference_percent <= 30:
        return "Progres ringan"
    if difference_percent <= 60:
        return "Progres sedang"
    return "Progres besar"


def process_document_text(raw_text: str, k: int = K_GRAM, w: int = WINDOW_SIZE):
    clean_text = preprocess_text(raw_text)
    kgrams = make_kgrams(clean_text, k)
    hashes = hash_kgrams(kgrams)
    windows, fingerprints = winnowing_fingerprints(hashes, w)
    return {
        "raw_text": raw_text,
        "clean_text": clean_text,
        "kgrams": kgrams,
        "hashes": hashes,
        "windows": windows,
        "fingerprints": fingerprints,
    }


# ==================== FUNGSI DATABASE: DOCUMENTS & COMPARISONS ====================
def save_document_record(user_id, original_name, saved_name, ext, processed):
    conn = get_conn()
    cur = conn.cursor()
    uploaded_at = now_dt().strftime("%Y-%m-%d %H:%M:%S")
    upload_date = today_str()
    cur.execute(
        """
        INSERT INTO documents (
            user_id, original_name, saved_name, file_ext, uploaded_at, upload_date,
            raw_text, clean_text, kgrams_json, hashes_json, windows_json, fingerprints_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id, original_name, saved_name, ext, uploaded_at, upload_date,
            processed["raw_text"], processed["clean_text"],
            json.dumps(processed["kgrams"]), json.dumps(processed["hashes"]),
            json.dumps(processed["windows"]), json.dumps(processed["fingerprints"]),
        ),
    )
    doc_id = cur.lastrowid
    conn.commit()
    conn.close()
    return doc_id


def get_documents_by_user(user_id):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM documents WHERE user_id = ? ORDER BY id DESC", (user_id,)).fetchall()
    conn.close()
    return rows


def get_documents_by_user_asc(user_id):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM documents WHERE user_id = ? ORDER BY id ASC", (user_id,)).fetchall()
    conn.close()
    return rows


def get_document_by_id_and_user(doc_id, user_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM documents WHERE id = ? AND user_id = ?", (doc_id, user_id)).fetchone()
    conn.close()
    return row


def get_fingerprints_from_row(row):
    return json.loads(row["fingerprints_json"]) if row["fingerprints_json"] else []


def upsert_comparison(user_id, doc1_id, doc2_id, similarity_percent, difference_percent, process_time, progress_label):
    a, b = sorted([doc1_id, doc2_id])
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO comparisons (
            user_id, doc1_id, doc2_id, similarity_percent, difference_percent,
            process_time, progress_label, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(doc1_id, doc2_id) DO UPDATE SET
            similarity_percent = excluded.similarity_percent,
            difference_percent = excluded.difference_percent,
            process_time = excluded.process_time,
            progress_label = excluded.progress_label,
            created_at = excluded.created_at
        """,
        (user_id, a, b, similarity_percent, difference_percent, process_time, progress_label,
         now_dt().strftime("%Y-%m-%d %H:%M:%S")),
    )
    conn.commit()
    conn.close()


def compare_new_document_against_all(user_id, new_doc_id):
    new_doc = get_document_by_id_and_user(new_doc_id, user_id)
    if not new_doc:
        return
    new_fp = get_fingerprints_from_row(new_doc)
    for old_doc in get_documents_by_user_asc(user_id):
        if old_doc["id"] == new_doc_id:
            continue
        start = time.perf_counter()
        old_fp = get_fingerprints_from_row(old_doc)
        similarity = calculate_similarity(new_fp, old_fp)
        difference = calculate_difference(similarity)
        label = get_progress_label(difference)
        elapsed = round(time.perf_counter() - start, 6)
        upsert_comparison(user_id, new_doc_id, old_doc["id"], similarity, difference, elapsed, label)


def get_comparisons_by_user(user_id):
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT c.*, d1.original_name AS doc1_name, d2.original_name AS doc2_name
        FROM comparisons c
        JOIN documents d1 ON c.doc1_id = d1.id
        JOIN documents d2 ON c.doc2_id = d2.id
        WHERE c.user_id = ?
        ORDER BY c.difference_percent DESC, c.id DESC
        """, (user_id,),
    ).fetchall()
    conn.close()
    return rows


def get_best_highlight_for_date(user_id, date_str):
    # PENTING: kriteria "terbaik" di sini HARUS sama dengan get_recent_progress_highlight_by_user
    # dan get_comparisons_by_user (urut berdasarkan difference_percent DESC = progres paling
    # besar dianggap terbaik). Sebelumnya fungsi ini urut berdasarkan similarity_percent DESC
    # (progres paling KECIL yang dianggap "terbaik"), sehingga kurva mingguan, "Progress Terbaik
    # Hari Ini", dan "Progress Hari Ini" menampilkan hasil yang berbeda/tidak sinkron dengan
    # "Highlight Progress 1-3 Hari Terakhir" dan tabel Progress Berurutan / Semua Comparison.
    # doc1_name/doc2_name juga di-select di sini supaya info "compare file vs" ikut tersedia.
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT c.*, d1.original_name AS doc1_name, d2.original_name AS doc2_name
        FROM comparisons c
        JOIN documents d1 ON c.doc1_id = d1.id
        JOIN documents d2 ON c.doc2_id = d2.id
        WHERE c.user_id = ? AND (d1.upload_date = ? OR d2.upload_date = ?)
        ORDER BY c.difference_percent DESC, c.id DESC
        """,
        (user_id, date_str, date_str),
    ).fetchall()
    conn.close()
    if not rows:
        return None
    for row in rows:
        if float(row["similarity_percent"]) < 100.0:
            return row
    return rows[0]


def get_today_best_highlight(user_id):
    return get_best_highlight_for_date(user_id, today_str())


def get_upload_days_by_user(user_id):
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT upload_date, COUNT(*) AS total
        FROM documents WHERE user_id = ?
        GROUP BY upload_date ORDER BY upload_date DESC
        """,
        (user_id,),
    ).fetchall()
    conn.close()
    return rows


def get_sequential_progress_by_user(user_id):
    docs = get_documents_by_user_asc(user_id)
    if len(docs) < 2:
        return []
    conn = get_conn()
    cur = conn.cursor()
    results = []
    for i in range(len(docs) - 1):
        d1, d2 = docs[i], docs[i + 1]
        a, b = sorted([d1["id"], d2["id"]])
        row = cur.execute(
            "SELECT * FROM comparisons WHERE user_id = ? AND doc1_id = ? AND doc2_id = ?",
            (user_id, a, b),
        ).fetchone()
        if row:
            results.append(
                {
                    "from_name": d1["original_name"],
                    "to_name": d2["original_name"],
                    "difference_percent": row["difference_percent"],
                    "progress_label": row["progress_label"],
                    "similarity_percent": row["similarity_percent"],
                }
            )
    conn.close()
    return results


def get_recent_progress_highlight_by_user(user_id):
    """Ambil rata-rata progress (difference) dari 1-3 hari upload terakhir."""
    conn = get_conn()
    cur = conn.cursor()
    dates = cur.execute(
        "SELECT DISTINCT upload_date FROM documents WHERE user_id = ? ORDER BY upload_date DESC",
        (user_id,),
    ).fetchall()
    results = []
    for d in dates:
        upload_date = d["upload_date"]
        rows = cur.execute(
            """
            SELECT c.* FROM comparisons c
            JOIN documents d1 ON c.doc1_id = d1.id
            JOIN documents d2 ON c.doc2_id = d2.id
            WHERE c.user_id = ? AND (d1.upload_date = ? OR d2.upload_date = ?)
            ORDER BY c.difference_percent DESC, c.id DESC
            """,
            (user_id, upload_date, upload_date),
        ).fetchall()
        if rows:
            best = rows[0]
            results.append({
                "upload_date": upload_date,
                "best_difference": float(best["difference_percent"]),
                "best_similarity": float(best["similarity_percent"]),
                "progress_label": best["progress_label"],
            })
    conn.close()
    if not results:
        return None
    recent = results[:3]
    values = [item["best_difference"] for item in recent if item["best_difference"] is not None]
    if not values:
        return None
    avg_diff = round(sum(values) / len(values), 2)
    if avg_diff <= 10:
        label, message, alert = (
            "Perubahan sangat kecil",
            "Dalam 1-3 hari terakhir, progress kamu masih sangat kecil. Coba tambahkan perubahan yang lebih signifikan.",
            "warning",
        )
    elif avg_diff <= 30:
        label, message, alert = (
            "Progres ringan",
            "Dalam 1-3 hari terakhir, progress kamu masih ringan. Ayo tingkatkan progress pekerjaanmu.",
            "warning",
        )
    elif avg_diff <= 60:
        label, message, alert = (
            "Progres sedang",
            "Dalam 1-3 hari terakhir, progress kamu sudah cukup terlihat.",
            "info",
        )
    else:
        label, message, alert = (
            "Progres besar",
            "Dalam 1-3 hari terakhir, progress kamu sangat baik dan terlihat signifikan.",
            "success",
        )
    return {
        "avg_difference": avg_diff,
        "label": label,
        "message": message,
        "alert": alert,
        "days_count": len(values),
    }


def get_weekly_progress_chart_data(user_id):
    """Kurva progres Senin-Minggu pada minggu berjalan: persentase difference + jenis progresnya."""
    today = now_dt().date()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    data = []
    for i in range(7):
        d = monday + timedelta(days=i)
        d_str = d.strftime("%Y-%m-%d")
        value = 0.0
        label = NO_UPLOAD_LABEL
        color = NO_UPLOAD_COLOR
        has_upload = False
        if d <= today:
            highlight = get_best_highlight_for_date(user_id, d_str)
            if highlight:
                value = float(highlight["difference_percent"])
                label = highlight["progress_label"]
                color = PROGRESS_LABEL_COLORS.get(label, NO_UPLOAD_COLOR)
                has_upload = True
        data.append({
            "day_name": INDO_DAYS_SHORT[i],
            "date_label": f"{d.day} {INDO_MONTHS_SHORT[d.month - 1]}",
            "full_date": d_str,
            "value": value,
            "label": label,
            "color": color,
            "has_upload": has_upload,
            "is_today": d_str == today_str(),
            "is_future": d > today,
        })
    if monday.month == sunday.month:
        month_label = f"{INDO_MONTHS_FULL[monday.month - 1]} {monday.year}"
    else:
        month_label = f"{INDO_MONTHS_FULL[monday.month - 1]}-{INDO_MONTHS_FULL[sunday.month - 1]} {sunday.year}"
    return {"days": data, "month_label": month_label}


def get_weekly_upload_progress(user_id):
    """Konsistensi upload Senin-Minggu pada minggu berjalan (7 hari)."""
    today = now_dt().date()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT upload_date FROM documents WHERE user_id = ? AND upload_date BETWEEN ? AND ?",
        (user_id, monday.strftime("%Y-%m-%d"), sunday.strftime("%Y-%m-%d")),
    ).fetchall()
    conn.close()
    days_uploaded = len(rows)
    percent = round((days_uploaded / 7) * 100)
    return {"days_uploaded": days_uploaded, "percent": percent, "target": 7}


# SVG api berlapis (outer/mid/inner) dengan gradasi merah -> oranye membara.
FLAME_SVG = """
<svg class="flame-svg {css_class}" viewBox="0 0 100 130" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="flameOuterGrad" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#7a0d02"/>
      <stop offset="55%" stop-color="#ff2d55"/>
      <stop offset="100%" stop-color="#ff8a00"/>
    </linearGradient>
    <linearGradient id="flameMidGrad" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#ff8a00"/>
      <stop offset="100%" stop-color="#ffb347"/>
    </linearGradient>
    <linearGradient id="flameInnerGrad" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#ffcf7a"/>
      <stop offset="100%" stop-color="#fff6e6"/>
    </linearGradient>
  </defs>
  <path class="flame-layer flame-layer-outer" fill="url(#flameOuterGrad)"
    d="M50 125 C18 108 10 68 32 38 C36 55 47 56 45 33 C58 52 74 40 63 12 C96 40 92 92 66 118 C71 98 54 92 50 125 Z"/>
  <path class="flame-layer flame-layer-mid" fill="url(#flameMidGrad)"
    d="M50 122 C28 108 24 78 38 55 C41 66 49 66 47 50 C56 63 66 54 60 34 C80 55 78 90 60 112 C64 96 53 94 50 122 Z"/>
  <path class="flame-layer flame-layer-inner" fill="url(#flameInnerGrad)"
    d="M50 118 C36 108 34 88 43 70 C45 78 50 78 49 66 C55 76 62 69 58 54 C71 70 70 94 58 108 C61 96 53 95 50 118 Z"/>
</svg>
"""


def flame_meta(flame_status: str):
    css_class = "flame-svg-on" if flame_status == "on" else "flame-svg-off"
    svg_markup = FLAME_SVG.format(css_class=css_class)
    if flame_status == "on":
        return {"icon": "\U0001F525", "css_class": "flame-on", "badge_class": "badge-flame-on",
                "text": "Api menyala", "flame_status": "on", "svg": svg_markup}
    return {"icon": "\U0001F4A8", "css_class": "flame-off", "badge_class": "badge-flame-off",
            "text": "Api padam", "flame_status": "off", "svg": svg_markup}


# ==================== SHARED THEME (merah - oranye elektrik - hitam) ====================
THEME_STYLE = """
:root {
  --bg-1: #0a0000;
  --bg-2: #210603;
  --bg-3: #360a05;
  --purple: #ff2d55;
  --purple-light: #ff6b4a;
  --blue: #ff8a00;
  --blue-glow: #ffb347;
  --text: #fff6f2;
  --text-dim: rgba(255, 246, 242, 0.62);
  --glass: rgba(255, 255, 255, 0.05);
  --glass-border: rgba(255, 255, 255, 0.1);
  --danger: #ff4d4d;
  --success: #34d399;
  --warning: #fbbf24;
  --info: #ff8a00;
}
* { font-family: 'Inter', sans-serif; }
body {
  background: radial-gradient(circle at 15% 10%, #4a0f06 0%, #200604 45%, #0a0000 100%);
  color: var(--text);
  min-height: 100vh;
}
a { color: var(--blue-glow); }
.navbar-custom {
  background: linear-gradient(135deg, #1a0503 0%, #3a0b05 100%) !important;
  border-bottom: 1px solid var(--glass-border);
  box-shadow: 0 4px 30px rgba(255, 45, 85, 0.25);
}
.navbar-brand { letter-spacing: 0.3px; display: flex; align-items: center; }
.navbar-brand .brand-logo { height: 32px; width: auto; margin-right: 8px; border-radius: 6px; object-fit: contain; }
.btn-settings {
  background: rgba(255,255,255,0.08);
  border: 1px solid var(--glass-border);
  color: var(--text);
}
.btn-settings:hover { background: rgba(255,255,255,0.16); color: #fff; }
.card-shadow, .soft-card, .auth-card {
  background: var(--glass);
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  border: 1px solid var(--glass-border);
  border-radius: 20px;
  box-shadow: 0 20px 50px rgba(0,0,0,0.45);
  color: var(--text);
}
.soft-card { padding: 18px; }
.text-muted { color: var(--text-dim) !important; }
h1, h2, h3, h4, h5, h6 { color: var(--text); }
.form-control {
  background: rgba(255,255,255,0.06);
  border: 2px solid var(--glass-border);
  border-radius: 12px;
  color: var(--text);
  padding: 12px 16px;
}
.form-control:focus {
  background: rgba(255,255,255,0.09);
  border-color: var(--purple-light);
  box-shadow: 0 0 0 4px rgba(255, 45, 85, 0.18);
  color: var(--text);
}
.form-control::placeholder { color: rgba(255,246,242,0.35); }
.form-label { color: var(--text-dim); font-weight: 500; }
.btn-dark, .btn-primary {
  background: linear-gradient(135deg, var(--purple), var(--blue));
  border: none;
  border-radius: 12px;
  font-weight: 600;
  color: #1a0000;
  transition: all .25s ease;
}
.btn-dark:hover, .btn-primary:hover {
  transform: translateY(-2px);
  box-shadow: 0 10px 25px rgba(255, 138, 0, 0.35);
  color: #1a0000;
}
.btn-outline-secondary {
  border-radius: 12px;
  border: 2px solid var(--glass-border);
  color: var(--text);
}
.btn-outline-secondary:hover { background: rgba(255,255,255,0.08); color: #fff; }
.btn-outline-success {
  border-radius: 12px;
  border: 2px solid var(--success);
  color: var(--success);
}
.btn-outline-success:hover { background: rgba(52,211,153,0.15); color: #fff; }
.table { color: var(--text); }
.table-striped > tbody > tr:nth-of-type(odd) > * { background: rgba(255,255,255,0.03); color: var(--text); }
.table thead { color: var(--blue-glow); border-bottom: 1px solid var(--glass-border); }
.mono-box {
  background: #0a0000;
  color: #ffe0c2;
  border-radius: 12px;
  padding: 16px;
  white-space: pre-wrap;
  word-break: break-word;
  max-height: 320px;
  overflow-y: auto;
  font-family: Consolas, monospace;
  font-size: 13px;
  border: 1px solid var(--glass-border);
}
.small-scroll {
  max-height: 240px;
  overflow-y: auto;
  font-family: Consolas, monospace;
  font-size: 13px;
  background: rgba(255,255,255,0.03);
  padding: 12px;
  border: 1px solid var(--glass-border);
  border-radius: 12px;
  white-space: pre-wrap;
  word-break: break-word;
  color: var(--text);
}
.hero-highlight {
  background: linear-gradient(135deg, rgba(255,45,85,0.25), rgba(255,138,0,0.15));
  border: 1px solid var(--glass-border);
}
.badge-flame-on {
  background: linear-gradient(135deg, var(--purple), var(--blue));
  color: #1a0000;
  font-weight: 600;
}
.badge-flame-off { background: rgba(255,255,255,0.12); color: var(--text-dim); }
.badge-finished { background: linear-gradient(135deg, #34d399, #10b981); color: #04241a; font-weight: 700; }
.step-card {
  border-left: 3px solid var(--purple-light);
  padding-left: 16px;
  margin-bottom: 22px;
}
.step-number {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 28px; height: 28px;
  border-radius: 50%;
  background: linear-gradient(135deg, var(--purple), var(--blue));
  color: #1a0000;
  font-weight: 700;
  font-size: 13px;
  margin-right: 8px;
}
.stat-pill {
  display: inline-block;
  padding: 4px 12px;
  border-radius: 999px;
  background: rgba(255,255,255,0.08);
  border: 1px solid var(--glass-border);
  font-size: 12.5px;
  color: var(--blue-glow);
  margin-right: 6px;
}
/* ---- Streak / api animasi SVG ---- */
.flame-stage { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 10px 0 4px; }
.flame-orb {
  position: relative;
  width: 150px; height: 170px;
  border-radius: 24px;
  display: flex; align-items: center; justify-content: center;
  background: radial-gradient(circle at 50% 65%, rgba(255,45,85,0.16), rgba(10,0,0,0.4) 72%);
  border: 1px solid var(--glass-border);
}
.flame-orb-on { animation: orbPulse 2.2s ease-in-out infinite; }
.flame-orb-on::before {
  content: "";
  position: absolute;
  inset: -14px;
  border-radius: 24px;
  background: radial-gradient(circle, rgba(255,138,0,0.4) 0%, rgba(255,45,85,0.32) 45%, transparent 72%);
  filter: blur(18px);
  z-index: 0;
  animation: glowPulse 1.3s ease-in-out infinite;
}
.flame-orb-off {
  background: radial-gradient(circle at 50% 65%, rgba(255,255,255,0.05), rgba(10,0,0,0.5) 72%);
  filter: grayscale(1);
  opacity: .7;
}
.flame-svg { position: relative; z-index: 1; width: 92px; height: 118px; overflow: visible; }
.flame-layer { transform-box: fill-box; }
.flame-svg-on .flame-layer-outer { transform-origin: 50% 100%; animation: flameRiseOuter 1.5s ease-in-out infinite; }
.flame-svg-on .flame-layer-mid { transform-origin: 50% 100%; animation: flameRiseMid 1.1s ease-in-out infinite; }
.flame-svg-on .flame-layer-inner { transform-origin: 50% 100%; animation: flameRiseInner 0.85s ease-in-out infinite; }
.flame-svg-off .flame-layer { filter: grayscale(1) brightness(.55); opacity: .55; }
@keyframes flameRiseOuter {
  0%, 100% { transform: scaleY(1) scaleX(1) translateY(0) rotate(0deg); }
  25% { transform: scaleY(1.14) scaleX(0.95) translateY(-5px) rotate(-2deg); }
  50% { transform: scaleY(0.9) scaleX(1.06) translateY(3px) rotate(1deg); }
  75% { transform: scaleY(1.08) scaleX(0.97) translateY(-3px) rotate(2deg); }
}
@keyframes flameRiseMid {
  0%, 100% { transform: scaleY(1) translateY(0) rotate(0deg); }
  30% { transform: scaleY(1.2) translateY(-7px) rotate(2deg); }
  60% { transform: scaleY(0.86) translateY(4px) rotate(-2deg); }
}
@keyframes flameRiseInner {
  0%, 100% { transform: scaleY(1) translateY(0); }
  35% { transform: scaleY(1.28) translateY(-9px); }
  65% { transform: scaleY(0.82) translateY(5px); }
}
@keyframes glowPulse {
  0%, 100% { opacity: .5; transform: scale(1); }
  50% { opacity: 1; transform: scale(1.15); }
}
@keyframes orbPulse {
  0%, 100% { box-shadow: 0 0 0 rgba(255,45,85,0); }
  50% { box-shadow: 0 0 40px rgba(255,138,0,0.25); }
}
/* ---- Progress bar custom ---- */
.progress-block { margin-bottom: 18px; }
.progress-block .progress-label-row { display: flex; justify-content: space-between; font-size: 13px; color: var(--text-dim); margin-bottom: 6px; }
.progress-track {
  width: 100%; height: 14px; border-radius: 999px;
  background: rgba(255,255,255,0.08);
  border: 1px solid var(--glass-border);
  overflow: hidden;
}
.progress-fill {
  height: 100%; border-radius: 999px;
  background: linear-gradient(90deg, var(--purple), var(--blue));
  transition: width .6s ease;
}
/* ---- Alert restyle ---- */
.alert { border-radius: 14px; border: 1px solid var(--glass-border); }
.alert-success { background: rgba(52,211,153,0.12); color: #6ee7b7; }
.alert-danger { background: rgba(255,77,109,0.12); color: #ff9fb3; }
.alert-warning { background: rgba(251,191,36,0.12); color: #fde68a; }
.alert-info { background: rgba(255,138,0,0.12); color: #ffc98a; }
.btn-close { filter: invert(1) grayscale(1) brightness(1.8); opacity: .8; }
.btn-close:hover { opacity: 1; }
.form-select {
  background-color: rgba(255,255,255,0.06);
  border: 2px solid var(--glass-border);
  border-radius: 12px;
  color: var(--text);
}
.form-select:focus {
  background-color: rgba(255,255,255,0.09);
  border-color: var(--purple-light);
  box-shadow: 0 0 0 4px rgba(255,45,85,0.18);
  color: var(--text);
}
.form-select option { background: #1c0704; color: var(--text); }
input[type="file"] { display: none; }
.upload-dropzone {
  border: 2px dashed rgba(255,45,85,0.45);
  border-radius: 18px;
  background: rgba(255,255,255,0.03);
  padding: 34px 20px;
  text-align: center;
  cursor: pointer;
  transition: all .25s ease;
}
.upload-dropzone:hover, .upload-dropzone.dragover {
  border-color: var(--blue-glow);
  background: rgba(255,138,0,0.06);
  box-shadow: 0 0 30px rgba(255,138,0,0.15);
}
.upload-dropzone .bi { font-size: 40px; background: linear-gradient(135deg, var(--purple), var(--blue)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
.upload-dropzone p { margin: 10px 0 2px; font-weight: 600; color: var(--text); }
.upload-dropzone small { color: var(--text-dim); }
.upload-dropzone.has-file { border-style: solid; border-color: var(--success); background: rgba(52,211,153,0.06); }
.scenario-card {
  background: rgba(255,255,255,0.04);
  border: 1px solid var(--glass-border);
  border-radius: 16px;
  padding: 16px 18px; height: 100%;
}
.scenario-card h6 { color: #fff; font-weight: 700; margin-bottom: 6px; }
.scenario-card p { color: var(--text-dim); font-size: 13px; margin-bottom: 12px; }
.state-pill { font-family: Consolas, monospace; font-size: 12.5px; }
.auth-logo { text-align: center; margin-bottom: 28px; }
.auth-logo .icon {
  font-size: 46px;
  background: linear-gradient(135deg, var(--purple), var(--blue));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}
.auth-title { font-weight: 700; font-size: 26px; color: #fff; }
.auth-subtitle { color: var(--text-dim); font-size: 14px; }
.input-icon { position: relative; }
.input-icon .form-control { padding-left: 46px; }
.input-icon .bi { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: rgba(255,246,242,0.35); font-size: 18px; }
.chart-card canvas { max-height: 280px; }
"""

AUTH_PAGE_STYLE = """
* { font-family: 'Inter', sans-serif; }
body {
  background: radial-gradient(circle at 15% 10%, #4a0f06 0%, #200604 45%, #0a0000 100%);
  min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 20px;
}
.auth-wrap {
  background: rgba(255,255,255,0.05);
  backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
  border-radius: 32px; padding: 50px 45px; max-width: 460px; width: 100%;
  border: 1px solid rgba(255,255,255,0.08);
  box-shadow: 0 40px 80px rgba(0,0,0,0.6);
  animation: fadeInUp 0.7s ease-out;
}
@keyframes fadeInUp { from { opacity:0; transform: translateY(40px);} to { opacity:1; transform: translateY(0);} }
.auth-wrap .logo { text-align: center; margin-bottom: 30px; }
.auth-wrap .logo .brand-logo-img { height: 64px; width: auto; object-fit: contain; margin-bottom: 6px; border-radius: 10px; }
.auth-wrap .logo .icon {
  font-size: 48px;
  background: linear-gradient(135deg, #ff2d55, #ff8a00);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  display: inline-block;
}
.auth-wrap .logo h1 { color: #fff; font-size: 27px; font-weight: 700; margin-top: 8px; letter-spacing: -0.5px; }
.auth-wrap .logo p { color: rgba(255,255,255,0.55); font-size: 14px; margin-top: 4px; }
.auth-wrap .form-group { margin-bottom: 18px; }
.auth-wrap .form-group label { color: rgba(255,255,255,0.7); font-size: 13px; font-weight: 500; margin-bottom: 6px; display: block; }
.auth-wrap .input-wrapper { position: relative; }
.auth-wrap .input-wrapper .bi { position: absolute; left: 16px; top: 50%; transform: translateY(-50%); color: rgba(255,255,255,0.3); font-size: 18px; transition: color .3s; }
.auth-wrap .input-wrapper input {
  width: 100%; padding: 14px 16px 14px 48px;
  background: rgba(255,255,255,0.06); border: 2px solid rgba(255,255,255,0.09);
  border-radius: 14px; color: #fff; font-size: 15px; outline: none; transition: all .3s;
}
.auth-wrap .input-wrapper input::placeholder { color: rgba(255,255,255,0.25); }
.auth-wrap .input-wrapper input:focus {
  border-color: rgba(255, 138, 0, 0.55);
  background: rgba(255,255,255,0.09);
  box-shadow: 0 0 0 4px rgba(255, 138, 0, 0.12);
}
.auth-wrap .input-wrapper input:focus ~ .bi { color: #ff8a00; }
.auth-wrap .btn-auth {
  width: 100%; padding: 14px; border: none; border-radius: 14px;
  background: linear-gradient(135deg, #ff2d55, #ff8a00);
  color: #1a0000; font-size: 16px; font-weight: 700; cursor: pointer; margin-top: 8px;
  transition: all .3s ease;
}
.auth-wrap .btn-auth:hover { transform: translateY(-2px); box-shadow: 0 12px 30px rgba(255,138,0,0.35); }
.auth-wrap .switch-link { text-align: center; margin-top: 20px; color: rgba(255,255,255,0.45); font-size: 14px; }
.auth-wrap .switch-link a { color: #ff8a00; text-decoration: none; font-weight: 600; }
.auth-wrap .switch-link a:hover { color: #ff6b4a; }
.auth-wrap .hint { color: rgba(255,255,255,0.32); font-size: 12px; margin-top: 4px; display: block; }
.alert-custom {
  background: rgba(255,255,255,0.05); backdrop-filter: blur(10px);
  border: 1px solid rgba(255,255,255,0.08); border-radius: 14px; padding: 12px 16px;
  color: #fff; font-size: 14px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px;
}
.alert-custom .bi { font-size: 18px; }
.alert-custom.danger { border-color: rgba(255,77,109,0.35); background: rgba(255,77,109,0.12); }
.alert-custom.danger .bi { color: #ff4d6d; }
.alert-custom.success { border-color: rgba(52,211,153,0.35); background: rgba(52,211,153,0.12); }
.alert-custom.success .bi { color: #34d399; }
.alert-custom.warning { border-color: rgba(251,191,36,0.35); background: rgba(251,191,36,0.12); }
.alert-custom.warning .bi { color: #fbbf24; }
@media (max-width: 480px) { .auth-wrap { padding: 32px 24px; } }
"""

# ==================== TEMPLATE DASAR (NAVBAR + LAYOUT) ====================
BASE_TEMPLATE = """
<!doctype html><html lang="id" data-bs-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }} - Skripsiku</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>""" + THEME_STYLE + """</style>
</head>
<body>
<nav class="navbar navbar-expand-lg navbar-custom mb-4">
  <div class="container">
    <a class="navbar-brand fw-bold text-white" href="{{ url_for('index') }}">
      <img src="{{ url_for('static', filename='api.png') }}" alt="Skripsiku" class="brand-logo"
           onerror="this.style.display='none'; document.getElementById('brand-fallback-icon').style.display='inline-block';">
      <i class="bi bi-fire me-2" id="brand-fallback-icon" style="display:none;"></i>
      Skripsiku
    </a>
    <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
      <span class="navbar-toggler-icon"></span>
    </button>
    <div class="collapse navbar-collapse" id="navbarNav">
      <ul class="navbar-nav ms-auto">
        {% if session.user_id %}
        <li class="nav-item"><span class="nav-link text-white-50"><i class="bi bi-person-circle me-1"></i>{{ session.username }}</span></li>
        <li class="nav-item"><a class="nav-link text-white" href="{{ url_for('index') }}"><i class="bi bi-house me-1"></i>Home</a></li>
        <li class="nav-item"><a class="nav-link text-white" href="{{ url_for('documents') }}"><i class="bi bi-file-earmark me-1"></i>Dokumen</a></li>
        <li class="nav-item"><a class="nav-link text-white" href="{{ url_for('comparisons') }}"><i class="bi bi-bar-chart me-1"></i>Progress</a></li>
        <li class="nav-item"><a class="nav-link text-white" href="{{ url_for('settings') }}"><i class="bi bi-gear me-1"></i>Settings</a></li>
        <li class="nav-item"><a class="nav-link text-danger" href="{{ url_for('logout') }}"><i class="bi bi-box-arrow-right me-1"></i>Logout</a></li>
        {% else %}
        <li class="nav-item"><a class="nav-link text-white" href="{{ url_for('login') }}"><i class="bi bi-box-arrow-in-right me-1"></i>Login</a></li>
        <li class="nav-item"><a class="nav-link btn btn-outline-light btn-sm px-3" href="{{ url_for('register') }}"><i class="bi bi-person-plus me-1"></i>Register</a></li>
        {% endif %}
      </ul>
    </div>
  </div>
</nav>
<div class="container">
{% with messages = get_flashed_messages(with_categories=true) %}
  {% if messages %}
    {% for category, msg in messages %}
    <div class="alert alert-{{ category or 'info' }} alert-dismissible fade show shadow-sm" role="alert">
      <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' if category == 'warning' else 'info-circle' }} me-2"></i>
      {{ msg }}
      <button type="button" class="btn-close btn-close-white" data-bs-dismiss="alert"></button>
    </div>
    {% endfor %}
  {% endif %}
{% endwith %}
</div>
<div class="container pb-5">
{{ content | safe }}
</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""


# ==================== ROUTE: AUTH ====================
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        phone_number = request.form.get("phone_number", "").strip()

        if not username or not email or not password:
            flash("Semua field wajib diisi!", "danger")
            return redirect(url_for("register"))
        if password != confirm_password:
            flash("Password tidak cocok!", "danger")
            return redirect(url_for("register"))
        if len(password) < 6:
            flash("Password minimal 6 karakter!", "danger")
            return redirect(url_for("register"))

        phone_number = normalize_phone_number(phone_number) if phone_number else None

        try:
            user_id = create_user(username, email, password, phone_number)
            # ---- Welcome message: dikirim begitu user pertama kali daftar ----
            if phone_number:
                send_direct_message(user_id, phone_number, "welcome", msg_welcome(username))
                conn = get_conn()
                conn.execute("UPDATE users SET welcome_sent = 1 WHERE id = ?", (user_id,))
                conn.commit()
                conn.close()
            flash("Registrasi berhasil! Silakan login.", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("Username atau email sudah terdaftar!", "danger")
            return redirect(url_for("register"))

    return render_template_string("""
<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Register - Skripsiku</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
<style>""" + AUTH_PAGE_STYLE + """</style>
</head>
<body>
<div class="auth-wrap">
  <div class="logo">
    <img src="{{ url_for('static', filename='api.png') }}" alt="Skripsiku" class="brand-logo-img"
         onerror="this.style.display='none'; document.getElementById('register-fallback-icon').style.display='inline-block';">
    <span class="icon" id="register-fallback-icon" style="display:none;"><i class="bi bi-lightning-charge-fill"></i></span>
    <h1>Mulai Sekarang</h1>
    <p>Daftar dan pantau progres skripsimu</p>
  </div>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}{% for category, msg in messages %}
    <div class="alert-custom {{ category }}">
      <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' }}"></i>{{ msg }}
    </div>
    {% endfor %}{% endif %}
  {% endwith %}
  <form method="post">
    <div class="form-group">
      <label>Username</label>
      <div class="input-wrapper"><i class="bi bi-person"></i>
        <input type="text" name="username" placeholder="Masukkan username" required>
      </div>
    </div>
    <div class="form-group">
      <label>Email</label>
      <div class="input-wrapper"><i class="bi bi-envelope"></i>
        <input type="email" name="email" placeholder="Masukkan email" required>
      </div>
    </div>
    <div class="form-group">
      <label>Password</label>
      <div class="input-wrapper"><i class="bi bi-lock"></i>
        <input type="password" name="password" placeholder="Minimal 6 karakter" required>
      </div>
    </div>
    <div class="form-group">
      <label>Konfirmasi Password</label>
      <div class="input-wrapper"><i class="bi bi-lock-fill"></i>
        <input type="password" name="confirm_password" placeholder="Ulangi password" required>
      </div>
    </div>
    <div class="form-group">
      <label>Nomor WhatsApp</label>
      <div class="input-wrapper"><i class="bi bi-whatsapp"></i>
        <input type="tel" name="phone_number" placeholder="Contoh: 08123456789">
      </div>
      <span class="hint">Isi nomor WA-mu supaya langsung dapat pesan selamat datang & reminder progres (via Fonnte)</span>
    </div>
    <button type="submit" class="btn-auth"><i class="bi bi-person-plus me-2"></i>Daftar Sekarang</button>
  </form>
  <div class="switch-link">Sudah punya akun? <a href="{{ url_for('login') }}">Login</a></div>
</div>
</body>
</html>
""", title="Registrasi")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if not username or not password:
            flash("Username dan password wajib diisi!", "danger")
            return redirect(url_for("login"))
        user = get_user_by_username(username)
        if not user or not check_password_hash(user["password_hash"], password):
            flash("Username atau password salah!", "danger")
            return redirect(url_for("login"))
        session.permanent = True
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        session["email"] = user["email"]
        session["phone_number"] = user["phone_number"]
        flash(f"Selamat datang, {user['username']}!", "success")
        return redirect(url_for("index"))

    return render_template_string("""
<!doctype html><html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login - Skripsiku</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
<style>""" + AUTH_PAGE_STYLE + """</style>
</head>
<body>
<div class="auth-wrap">
  <div class="logo">
    <img src="{{ url_for('static', filename='api.png') }}" alt="Skripsiku" class="brand-logo-img"
         onerror="this.style.display='none'; document.getElementById('login-fallback-icon').style.display='inline-block';">
    <span class="icon" id="login-fallback-icon" style="display:none;"><i class="bi bi-lightning-charge-fill"></i></span>
    <h1>Selamat Datang</h1>
    <p>Login untuk pantau progres skripsimu</p>
  </div>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}{% for category, msg in messages %}
    <div class="alert-custom {{ category }}">
      <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' }}"></i>{{ msg }}
    </div>
    {% endfor %}{% endif %}
  {% endwith %}
  <form method="post">
    <div class="form-group">
      <label>Username</label>
      <div class="input-wrapper"><i class="bi bi-person"></i>
        <input type="text" name="username" placeholder="Masukkan username" required>
      </div>
    </div>
    <div class="form-group">
      <label>Password</label>
      <div class="input-wrapper"><i class="bi bi-lock"></i>
        <input type="password" name="password" placeholder="Masukkan password" required>
      </div>
    </div>
    <button type="submit" class="btn-auth"><i class="bi bi-box-arrow-in-right me-2"></i>Login</button>
  </form>
  <div class="switch-link">Belum punya akun? <a href="{{ url_for('register') }}">Daftar Sekarang</a></div>
</div>
</body>
</html>
""", title="Login")


@app.route("/logout")
def logout():
    session.clear()
    flash("Anda telah logout.", "info")
    return redirect(url_for("login"))


# ==================== ROUTE: DASHBOARD (UPLOAD ADA DI SINI SAJA) ====================
@app.route("/", methods=["GET", "POST"])
def index():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]

    if request.method == "POST":
        if "file" not in request.files:
            flash("File tidak ditemukan.", "danger")
            return redirect(url_for("index"))
        file = request.files["file"]
        if file.filename == "":
            flash("Pilih file terlebih dahulu.", "warning")
            return redirect(url_for("index"))
        if not allowed_file(file.filename):
            flash("Format file tidak didukung. Gunakan txt, pdf, docx, html, atau htm.", "danger")
            return redirect(url_for("index"))

        original_name = file.filename
        ext = original_name.rsplit(".", 1)[1].lower()
        timestamp = now_dt().strftime("%Y%m%d%H%M%S%f")
        saved_name = f"{timestamp}_{secure_filename(original_name)}"
        save_path = UPLOAD_FOLDER / saved_name

        # Dicek SEBELUM insert: apakah ini benar-benar dokumen pertama user ini sepanjang masa.
        # Upload pertama tidak punya dokumen lain untuk dibandingkan, jadi tidak boleh
        # menunggu hasil compare untuk dianggap "valid" — kalau tidak, notifikasi
        # "upload pertama kali" baru akan terkirim setelah upload dokumen KEDUA.
        is_very_first_upload = len(get_documents_by_user(user_id)) == 0

        try:
            file.save(save_path)
            raw_text = extract_text(save_path, ext)
            processed = process_document_text(raw_text)
            new_doc_id = save_document_record(user_id, original_name, saved_name, ext, processed)
            compare_new_document_against_all(user_id, new_doc_id)
            best = get_today_best_highlight(user_id)
            is_valid_progress = is_very_first_upload or (best and float(best["similarity_percent"]) < 100.0)
            if is_valid_progress:
                update_streak_after_valid_upload(user_id)
            flash(f"File '{original_name}' berhasil diupload dan progress diperbarui.", "success")
        except Exception as e:
            if save_path.exists():
                save_path.unlink()
            flash(f"Gagal memproses file: {e}", "danger")
        return redirect(url_for("index"))

    # ---- GET: tampilkan dashboard ----
    streak = get_or_reset_streak_state(user_id)
    flame = flame_meta(streak["flame_status"])
    docs = get_documents_by_user(user_id)
    docs_count = len(docs)
    comparison_count = len(get_comparisons_by_user(user_id))
    sequential_progress = get_sequential_progress_by_user(user_id)
    best = get_today_best_highlight(user_id)
    upload_days = get_upload_days_by_user(user_id)
    recent_highlight = get_recent_progress_highlight_by_user(user_id)
    weekly_chart = get_weekly_progress_chart_data(user_id)
    weekly_progress = get_weekly_upload_progress(user_id)
    thesis_finished = bool(streak["thesis_finished"])

    content = render_template_string(
        """
{% if thesis_finished %}
<div class="alert alert-success d-flex justify-content-between align-items-center flex-wrap gap-2 mb-4">
  <div><i class="bi bi-mortarboard-fill me-2"></i><b>Skripsi kamu sudah ditandai SELESAI.</b> Reminder progres sedang non-aktif. Selamat! \U0001F393</div>
  <form method="post" action="{{ url_for('thesis_resume') }}" class="m-0">
    <button class="btn btn-outline-secondary btn-sm" type="submit"><i class="bi bi-arrow-repeat me-1"></i>Lanjutkan Progres Lagi</button>
  </form>
</div>
{% endif %}
<div class="row g-4 mb-4">
  <div class="col-lg-7">
    <div class="card card-shadow">
      <div class="card-body p-4">
        <h3 class="fw-bold mb-3"><i class="bi bi-cloud-upload me-2"></i>Upload Versi Dokumen</h3>
        <p class="text-muted">
          Upload minimal 1 file per hari agar streak tetap jalan.
          Streak hanya bertambah jika hari itu ada hasil compare terbaik yang bukan 100% similarity.
        </p>
        <form method="post" enctype="multipart/form-data" id="uploadForm">
          <div class="upload-dropzone" id="dropzone">
            <input type="file" id="fileInput" name="file" required>
            <div>
              <i class="bi bi-cloud-arrow-up"></i>
              <p id="fileLabel">Klik atau seret file ke sini</p>
              <small>.txt, .pdf, .docx, .html — maksimal 1 file</small>
            </div>
          </div>
          <button class="btn btn-dark w-100 mt-3" type="submit"><i class="bi bi-cloud-upload me-2"></i>Upload & Hitung Progress</button>
        </form>
        <script>
          (function () {
            const dz = document.getElementById('dropzone');
            const fi = document.getElementById('fileInput');
            const label = document.getElementById('fileLabel');
            dz.addEventListener('click', function () { fi.click(); });
            fi.addEventListener('change', function () {
              if (fi.files.length) { label.textContent = fi.files[0].name; dz.classList.add('has-file'); }
            });
            ['dragover', 'dragenter'].forEach(function (evt) {
              dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.add('dragover'); });
            });
            ['dragleave', 'dragend'].forEach(function (evt) {
              dz.addEventListener(evt, function () { dz.classList.remove('dragover'); });
            });
            dz.addEventListener('drop', function (e) {
              e.preventDefault();
              dz.classList.remove('dragover');
              if (e.dataTransfer.files.length) {
                fi.files = e.dataTransfer.files;
                label.textContent = e.dataTransfer.files[0].name;
                dz.classList.add('has-file');
              }
            });
          })();
        </script>
      </div>
    </div>
    {% if not thesis_finished %}
    <div class="card card-shadow mt-4">
      <div class="card-body p-4 text-center">
        <h5 class="fw-bold mb-2"><i class="bi bi-mortarboard me-2"></i>Sudah Selesai Skripsi?</h5>
        <p class="text-muted">Kalau skripsimu sudah kelar dan tidak butuh reminder progres lagi, tandai di sini ya.</p>
        <form method="post" action="{{ url_for('thesis_finish') }}" onsubmit="return confirm('Yakin skripsi sudah selesai? Reminder progres akan dihentikan.');">
          <button class="btn btn-outline-success" type="submit"><i class="bi bi-check2-circle me-2"></i>Skripsi Selesai</button>
        </form>
      </div>
    </div>
    {% endif %}
  </div>
  <div class="col-lg-5">
    <div class="card card-shadow">
      <div class="card-body p-4 text-center">
        <div class="flame-stage">
          <div class="flame-orb {{ 'flame-orb-on' if flame.flame_status == 'on' else 'flame-orb-off' }}">
            {{ flame.svg | safe }}
          </div>
        </div>
        <div class="big-number mt-2" style="font-size:40px;font-weight:700;">{{ streak["streak_count"] }}</div>
        <div class="fw-semibold">Streak Hari Aktif</div>
        <div class="mt-2"><span class="badge {{ flame.badge_class }}">{{ flame.icon }} {{ flame.text }}</span></div>
        <p class="text-muted mt-3 mb-0">
          Jika {{ days_warning }} hari tidak upload, kamu akan diingatkan.
          Jika sampai hari ke-{{ days_death }} masih belum upload, api padam dan streak reset ke 0.
        </p>
      </div>
    </div>
  </div>
</div>

<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h4 class="fw-bold mb-3"><i class="bi bi-speedometer2 me-2"></i>Progress & Konsistensi</h4>
    <div class="progress-block">
      <div class="progress-label-row">
        <span>Konsistensi Minggu Ini (Senin-Minggu)</span>
        <span>{{ weekly_progress.days_uploaded }} / {{ weekly_progress.target }} hari</span>
      </div>
      <div class="progress-track"><div class="progress-fill" style="width: {{ weekly_progress.percent }}%;"></div></div>
    </div>
    {% if best %}
    <div class="progress-block mb-0">
      <div class="progress-label-row">
        <span>Progress Terbaik Hari Ini &mdash; {{ best["progress_label"] }}</span>
        <span>{{ best["difference_percent"] }}%</span>
      </div>
      <div class="progress-track"><div class="progress-fill" style="width: {{ best["difference_percent"] }}%;"></div></div>
    </div>
    {% endif %}
  </div>
</div>

<div class="card card-shadow chart-card mb-4">
  <div class="card-body p-4">
    <h4 class="fw-bold mb-1"><i class="bi bi-graph-up-arrow me-2"></i>Kurva Progres Mingguan</h4>
    <p class="text-muted mb-2">{{ weekly_chart.month_label }} &middot; Senin - Minggu, persentase &amp; jenis progres per hari, update real-time</p>
    <div class="mb-3">
      <span class="stat-pill" style="color:#ff4d6d;">&#9679; Perubahan sangat kecil</span>
      <span class="stat-pill" style="color:#ff8a00;">&#9679; Progres ringan</span>
      <span class="stat-pill" style="color:#ffb347;">&#9679; Progres sedang</span>
      <span class="stat-pill" style="color:#34d399;">&#9679; Progres besar</span>
      <span class="stat-pill" style="color:rgba(255,255,255,0.55);">&#9679; Belum upload</span>
    </div>
    <canvas id="weeklyChart" height="110"></canvas>
  </div>
</div>
<script>
(function () {
  const ctx = document.getElementById('weeklyChart');
  if (!ctx || !window.Chart) { return; }
  const chartData = {{ weekly_chart.days | tojson }};
  new Chart(ctx, {
    type: 'line',
    data: {
      labels: chartData.map(function (d) { return d.day_name + '\\n' + d.date_label; }),
      datasets: [{
        label: 'Progress (%)',
        data: chartData.map(function (d) { return d.value; }),
        borderColor: '#ff8a00',
        backgroundColor: 'rgba(255,45,85,0.18)',
        pointBackgroundColor: chartData.map(function (d) { return d.color; }),
        pointBorderColor: chartData.map(function (d) { return d.is_today ? '#ffffff' : d.color; }),
        pointBorderWidth: chartData.map(function (d) { return d.is_today ? 2 : 0; }),
        pointRadius: chartData.map(function (d) { return d.is_today ? 7 : 5; }),
        borderWidth: 3,
        tension: 0.35,
        fill: true,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function (item) {
              const d = chartData[item.dataIndex];
              return d.has_upload ? (d.label + ': ' + d.value + '%') : 'Belum upload';
            }
          }
        }
      },
      scales: {
        x: { ticks: { color: '#fff6f2' }, grid: { color: 'rgba(255,255,255,0.06)' } },
        y: { beginAtZero: true, max: 100, ticks: { color: '#fff6f2' }, grid: { color: 'rgba(255,255,255,0.06)' } }
      }
    }
  });
})();
</script>

{% if recent_highlight %}
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h4 class="fw-bold mb-3"><i class="bi bi-graph-up me-2"></i>Highlight Progress 1-3 Hari Terakhir</h4>
    <div class="alert alert-{{ recent_highlight.alert }} mb-0">
      <h5 class="mb-2">{{ recent_highlight.label }}</h5>
      <p class="mb-1">{{ recent_highlight.message }}</p>
      <small>Rata-rata difference: <b>{{ recent_highlight.avg_difference }}%</b> dari {{ recent_highlight.days_count }} hari upload terakhir.</small>
    </div>
  </div>
</div>
{% endif %}

<div class="card card-shadow hero-highlight mb-4">
  <div class="card-body p-4">
    <h4 class="fw-bold mb-3"><i class="bi bi-calendar-check me-2"></i>Progress Hari Ini</h4>
    {% if best %}
    <div class="row g-3">
      <div class="col-md-6">
        <div class="soft-card h-100">
          <div class="small text-muted mb-1">Compare utama hari ini</div>
          <div class="fw-semibold">{{ best["doc1_name"] }}</div>
          <div class="text-muted">vs</div>
          <div class="fw-semibold">{{ best["doc2_name"] }}</div>
        </div>
      </div>
      <div class="col-md-2">
        <div class="soft-card text-center h-100">
          <div class="small text-muted">Similarity</div>
          <div class="fw-bold fs-4">{{ best["similarity_percent"] }}%</div>
        </div>
      </div>
      <div class="col-md-2">
        <div class="soft-card text-center h-100">
          <div class="small text-muted">Difference</div>
          <div class="fw-bold fs-4">{{ best["difference_percent"] }}%</div>
        </div>
      </div>
      <div class="col-md-2">
        <div class="soft-card text-center h-100">
          <div class="small text-muted">Progress</div>
          <div class="fw-bold">{{ best["progress_label"] }}</div>
        </div>
      </div>
    </div>
    {% else %}
    <p class="mb-0">Belum ada progress hari ini.</p>
    {% endif %}
  </div>
</div>

<div class="row g-4 mb-4">
  <div class="col-lg-6">
    <div class="card card-shadow">
      <div class="card-body p-4">
        <h5 class="fw-bold"><i class="bi bi-info-circle me-2"></i>Ringkasan Sistem</h5>
        <ul class="mb-0">
          <li>K-Gram: <b>{{ k_gram }}</b></li>
          <li>Window: <b>{{ window_size }}</b></li>
          <li>Jumlah dokumen: <b>{{ docs_count }}</b></li>
          <li>Jumlah perbandingan: <b>{{ comparison_count }}</b></li>
          <li>Status API: <span class="badge {{ flame.badge_class }}">{{ flame.icon }} {{ flame.text }}</span></li>
        </ul>
      </div>
    </div>
  </div>
  <div class="col-lg-6">
    <div class="card card-shadow">
      <div class="card-body p-4">
        <h5 class="fw-bold"><i class="bi bi-calendar me-2"></i>Hari Upload Tercatat</h5>
        <div class="table-responsive">
          <table class="table table-sm">
            <thead><tr><th>Tanggal</th><th>Total File</th></tr></thead>
            <tbody>
              {% if upload_days_recent %}
                {% for day in upload_days_recent %}
                <tr><td>{{ day["upload_date"] }}</td><td>{{ day["total"] }}</td></tr>
                {% endfor %}
              {% else %}
                <tr><td colspan="2" class="text-center">Belum ada upload.</td></tr>
              {% endif %}
            </tbody>
          </table>
        </div>
        <p class="text-muted mb-0 mt-2 small">Menampilkan 3 tanggal terbaru. Riwayat lengkap ada di halaman <a href="{{ url_for('documents') }}">Dokumen</a>.</p>
      </div>
    </div>
  </div>
</div>

<div class="card card-shadow">
  <div class="card-body p-4">
    <h4 class="fw-bold mb-3"><i class="bi bi-arrow-right-circle me-2"></i>Progress Berurutan</h4>
    <div class="table-responsive">
      <table class="table table-striped">
        <thead><tr><th>No</th><th>Dari</th><th>Ke</th><th>Similarity</th><th>Difference</th><th>Status</th></tr></thead>
        <tbody>
          {% if sequential_progress_recent %}
            {% for item in sequential_progress_recent %}
            <tr>
              <td>{{ loop.index }}</td>
              <td>{{ item.from_name }}</td>
              <td>{{ item.to_name }}</td>
              <td>{{ item.similarity_percent }}%</td>
              <td><span class="badge badge-flame-on">{{ item.difference_percent }}%</span></td>
              <td>{{ item.progress_label }}</td>
            </tr>
            {% endfor %}
          {% else %}
            <tr><td colspan="6" class="text-center">Belum cukup dokumen untuk menghitung progres berurutan.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
    <p class="text-muted mb-0 mt-2">Menampilkan 3 pasangan terbaru. Detail lengkap semua pasangan file ada di halaman <a href="{{ url_for('comparisons') }}">Progress</a>.</p>
  </div>
</div>
""",
        k_gram=K_GRAM, window_size=WINDOW_SIZE, docs_count=docs_count, comparison_count=comparison_count,
        sequential_progress_recent=sequential_progress[-3:][::-1], best=best, streak=streak, flame=flame,
        upload_days_recent=upload_days[:3], recent_highlight=recent_highlight,
        days_warning=DAYS_BEFORE_WARNING, days_death=DAYS_BEFORE_DEATH,
        weekly_chart=weekly_chart, weekly_progress=weekly_progress, thesis_finished=thesis_finished,
    )
    return render_template_string(BASE_TEMPLATE, title="Dashboard", content=content)


# ==================== ROUTE: SKRIPSI SELESAI ====================
@app.route("/thesis/finish", methods=["POST"])
def thesis_finish():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    mark_thesis_finished(session["user_id"])
    flash("Selamat! Skripsimu ditandai selesai. Reminder progres sudah dihentikan.", "success")
    return redirect(url_for("index"))


@app.route("/thesis/resume", methods=["POST"])
def thesis_resume():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    mark_thesis_resumed(session["user_id"])
    flash("Reminder progres diaktifkan kembali. Semangat lanjutkan skripsinya!", "info")
    return redirect(url_for("index"))


# ==================== ROUTE: DOKUMEN (HANYA LIST, TANPA UPLOAD) ====================
@app.route("/documents")
def documents():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    docs = get_documents_by_user(session["user_id"])
    content = render_template_string(
        """
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h3 class="fw-bold mb-1"><i class="bi bi-file-earmark-text me-2"></i>Daftar Dokumen</h3>
    <p class="text-muted">Upload dokumen baru dilakukan lewat <a href="{{ url_for('index') }}">Dashboard</a>.</p>
    <div class="table-responsive">
      <table class="table table-striped">
        <thead><tr><th>No</th><th>Nama File</th><th>Ext</th><th>Uploaded At</th><th>Preview</th><th>Aksi</th></tr></thead>
        <tbody>
          {% if docs %}
            {% for doc in docs %}
            <tr>
              <td>{{ loop.index }}</td>
              <td>{{ doc["original_name"] }}</td>
              <td><span class="stat-pill">{{ doc["file_ext"] }}</span></td>
              <td>{{ doc["uploaded_at"] }}</td>
              <td>
                <button type="button" class="btn btn-sm btn-outline-light btn-preview-file"
                        data-src="{{ url_for('document_preview', doc_id=doc['id']) }}"
                        data-name="{{ doc['original_name'] }}">
                  <i class="bi bi-eye me-1"></i>Preview
                </button>
              </td>
              <td><a class="btn btn-sm btn-dark" href="{{ url_for('document_detail', doc_id=doc['id']) }}"><i class="bi bi-diagram-3 me-1"></i>Detail</a></td>
            </tr>
            {% endfor %}
          {% else %}
            <tr><td colspan="6" class="text-center">Belum ada dokumen.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</div>

<div class="modal fade" id="filePreviewModal" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-lg modal-dialog-scrollable">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="filePreviewLabel">Preview File</h5>
        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
      </div>
      <div class="modal-body p-0" style="height:75vh;">
        <iframe id="filePreviewFrame" src="" style="width:100%;height:100%;border:0;background:#fff;" title="Preview File"></iframe>
      </div>
    </div>
  </div>
</div>
<script>
(function () {
  var modalEl = document.getElementById('filePreviewModal');
  var frame = document.getElementById('filePreviewFrame');
  var label = document.getElementById('filePreviewLabel');
  if (!modalEl) { return; }

  // Tombol dipasang listener-nya duluan, cek ketersediaan bootstrap baru dilakukan
  // saat tombol DIKLIK (bukan saat script ini jalan) — karena <script> bootstrap.bundle.min.js
  // baru dimuat di paling bawah <body>, setelah konten halaman ini.
  document.querySelectorAll('.btn-preview-file').forEach(function (btn) {
    btn.addEventListener('click', function () {
      frame.src = btn.getAttribute('data-src');
      label.textContent = 'Preview: ' + btn.getAttribute('data-name');
      if (window.bootstrap) {
        bootstrap.Modal.getOrCreateInstance(modalEl).show();
      }
    });
  });
  modalEl.addEventListener('hidden.bs.modal', function () {
    frame.src = '';
  });
})();
</script>
""",
        docs=docs,
    )
    return render_template_string(BASE_TEMPLATE, title="Dokumen", content=content)


@app.route("/documents/<int:doc_id>/file")
def document_file(doc_id):
    """Serve file ASLI apa adanya (dipakai untuk PDF, karena browser bisa render PDF native)."""
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    doc = get_document_by_id_and_user(doc_id, session["user_id"])
    if not doc:
        flash("Dokumen tidak ditemukan.", "danger")
        return redirect(url_for("documents"))
    file_path = UPLOAD_FOLDER / doc["saved_name"]
    if not file_path.exists():
        flash("File fisik dokumen ini tidak ditemukan di server.", "danger")
        return redirect(url_for("documents"))
    return send_from_directory(
        str(UPLOAD_FOLDER), doc["saved_name"], as_attachment=False, download_name=doc["original_name"]
    )


@app.route("/documents/<int:doc_id>/preview")
def document_preview(doc_id):
    """
    Konten yang dimuat ke iframe modal preview di halaman Daftar Dokumen.
    - PDF: kirim file aslinya, browser sudah bisa render PDF secara native lewat iframe.
    - txt/docx/html/htm: browser TIDAK bisa merender docx secara native, jadi supaya preview
      konsisten untuk semua jenis file yang boleh diupload, tampilkan teks hasil ekstraksi
      (raw_text, yang sudah disimpan sejak upload) di dalam halaman sederhana.
    """
    if "user_id" not in session:
        return "Silakan login terlebih dahulu.", 401
    doc = get_document_by_id_and_user(doc_id, session["user_id"])
    if not doc:
        return "Dokumen tidak ditemukan.", 404

    if doc["file_ext"] == "pdf":
        file_path = UPLOAD_FOLDER / doc["saved_name"]
        if not file_path.exists():
            return "<p style='font-family:sans-serif;color:#555;padding:16px;'>File fisik tidak ditemukan di server.</p>"
        return send_from_directory(
            str(UPLOAD_FOLDER), doc["saved_name"], as_attachment=False, download_name=doc["original_name"]
        )

    text = doc["raw_text"] or "(Tidak ada teks yang bisa diekstrak dari file ini.)"
    return render_template_string(
        """
<!doctype html><html><head><meta charset="utf-8">
<style>
  body { background:#1c1c1c; color:#f4f1ee; font-family: 'Inter', Arial, sans-serif; margin:0; padding:20px; }
  pre { white-space: pre-wrap; word-break: break-word; font-family: inherit; font-size: 14px; line-height:1.7; margin:0; }
</style></head>
<body><pre>{{ text }}</pre></body></html>
""",
        text=text,
    )



@app.route("/documents/<int:doc_id>")
def document_detail(doc_id):
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    doc = get_document_by_id_and_user(doc_id, session["user_id"])
    if not doc:
        flash("Dokumen tidak ditemukan.", "danger")
        return redirect(url_for("documents"))

    kgrams = json.loads(doc["kgrams_json"]) if doc["kgrams_json"] else []
    hashes = json.loads(doc["hashes_json"]) if doc["hashes_json"] else []
    windows = json.loads(doc["windows_json"]) if doc["windows_json"] else []
    fingerprints = json.loads(doc["fingerprints_json"]) if doc["fingerprints_json"] else []

    kgrams_preview = ", ".join(kgrams[:150])
    hashes_preview = ", ".join(str(h) for h in hashes[:150])
    windows_preview = [", ".join(str(h) for h in w) for w in windows[:40]]
    fingerprints_preview = fingerprints[:200]

    content = render_template_string(
        """
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h3 class="fw-bold mb-1"><i class="bi bi-file-earmark me-2"></i>Detail Dokumen &amp; Tahapan Winnowing</h3>
    <p class="text-muted mb-1">{{ doc["original_name"] }}</p>
    <p class="text-muted mb-0">Diupload: {{ doc["uploaded_at"] }}</p>
    <div class="mt-2">
      <span class="stat-pill">K-Gram = {{ k_gram }}</span>
      <span class="stat-pill">Window = {{ window_size }}</span>
      <span class="stat-pill">{{ kgrams_count }} k-gram</span>
      <span class="stat-pill">{{ hashes_count }} hash</span>
      <span class="stat-pill">{{ windows_count }} window</span>
      <span class="stat-pill">{{ fp_count }} fingerprint</span>
    </div>
  </div>
</div>
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <div class="step-card">
      <h5><span class="step-number">1</span>Ekstraksi &amp; Raw Text</h5>
      <p class="text-muted small">Teks mentah hasil ekstraksi dari file yang diupload.</p>
      <div class="mono-box">{{ doc["raw_text"] or "" }}</div>
    </div>
    <div class="step-card">
      <h5><span class="step-number">2</span>Preprocessing (Cleaning + Normalisasi)</h5>
      <p class="text-muted small">Lowercase, hapus karakter non alfanumerik, hapus spasi berlebih, lalu semua spasi dihilangkan.</p>
      <div class="mono-box">{{ doc["clean_text"] or "" }}</div>
    </div>
    <div class="step-card">
      <h5><span class="step-number">3</span>K-Gram Tokenization (k = {{ k_gram }})</h5>
      <p class="text-muted small">Teks bersih dipecah menjadi potongan sepanjang {{ k_gram }} karakter (sliding window per karakter). Total: {{ kgrams_count }} k-gram.</p>
      <div class="small-scroll">{{ kgrams_preview }}{% if kgrams_count > 150 %} ...(+{{ kgrams_count - 150 }} lagi){% endif %}</div>
    </div>
    <div class="step-card">
      <h5><span class="step-number">4</span>Hashing (MD5, 8 digit hex &rarr; integer)</h5>
      <p class="text-muted small">Setiap k-gram di-hash agar dapat dibandingkan secara efisien. Total: {{ hashes_count }} hash.</p>
      <div class="small-scroll">{{ hashes_preview }}{% if hashes_count > 150 %} ...(+{{ hashes_count - 150 }} lagi){% endif %}</div>
    </div>
    <div class="step-card">
      <h5><span class="step-number">5</span>Windowing (w = {{ window_size }})</h5>
      <p class="text-muted small">Deretan hash dikelompokkan ke dalam window berukuran {{ window_size }} untuk proses winnowing. Total: {{ windows_count }} window.</p>
      <div class="small-scroll">
        {% for w in windows_preview %}
        [{{ loop.index0 }}] {{ w }}<br>
        {% endfor %}
        {% if windows_count > 40 %} ...(+{{ windows_count - 40 }} window lagi){% endif %}
      </div>
    </div>
    <div class="step-card mb-0">
      <h5><span class="step-number">6</span>Winnowing Fingerprint</h5>
      <p class="text-muted small">Dari setiap window diambil hash minimum (posisi paling kanan jika ada duplikat) sebagai fingerprint dokumen. Fingerprint inilah yang dipakai untuk menghitung similarity antar dokumen. Total: {{ fp_count }} fingerprint.</p>
      <div class="table-responsive small-scroll" style="max-height:320px;">
        <table class="table table-sm table-striped mb-0">
          <thead><tr><th>#</th><th>Hash</th><th>Index</th></tr></thead>
          <tbody>
            {% for fp in fingerprints_preview %}
            <tr><td>{{ loop.index }}</td><td>{{ fp.hash }}</td><td>{{ fp.index }}</td></tr>
            {% endfor %}
          </tbody>
        </table>
        {% if fp_count > 200 %}<div class="text-muted small mt-2">...(+{{ fp_count - 200 }} fingerprint lagi)</div>{% endif %}
      </div>
    </div>
  </div>
</div>
""",
        doc=doc, k_gram=K_GRAM, window_size=WINDOW_SIZE,
        kgrams_count=len(kgrams), hashes_count=len(hashes), windows_count=len(windows), fp_count=len(fingerprints),
        kgrams_preview=kgrams_preview, hashes_preview=hashes_preview,
        windows_preview=windows_preview, fingerprints_preview=fingerprints_preview,
    )
    return render_template_string(BASE_TEMPLATE, title="Detail Dokumen", content=content)


# ==================== ROUTE: PROGRESS / COMPARISONS ====================
@app.route("/progress")
@app.route("/comparisons")
def comparisons():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    comparisons_data = get_comparisons_by_user(session["user_id"])
    content = render_template_string(
        """
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h3 class="fw-bold mb-3"><i class="bi bi-bar-chart me-2"></i>Semua Hasil Progress Comparison</h3>
    <p class="text-muted">Detail lengkap semua pasangan file ditampilkan di sini.</p>
    <div class="table-responsive">
      <table class="table table-striped align-middle">
        <thead>
          <tr>
            <th>No</th><th>Nama File 1</th><th>Nama File 2</th>
            <th>Similarity</th><th>Difference</th><th>Status Progress</th><th>Waktu</th>
          </tr>
        </thead>
        <tbody>
          {% if comparisons_data %}
            {% for row in comparisons_data %}
            <tr>
              <td>{{ loop.index }}</td>
              <td>{{ row["doc1_name"] }}</td>
              <td>{{ row["doc2_name"] }}</td>
              <td>{{ row["similarity_percent"] }}%</td>
              <td><span class="badge badge-flame-on">{{ row["difference_percent"] }}%</span></td>
              <td>{{ row["progress_label"] }}</td>
              <td>{{ row["process_time"] }} detik</td>
            </tr>
            {% endfor %}
          {% else %}
            <tr><td colspan="7" class="text-center">Belum ada hasil perbandingan.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</div>
""",
        comparisons_data=comparisons_data,
    )
    return render_template_string(BASE_TEMPLATE, title="Progress", content=content)


# ==================== ROUTE: SETTINGS ====================
@app.route("/settings", methods=["GET", "POST"])
def settings():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user = get_user_by_id(session["user_id"])
    if request.method == "POST":
        phone_number = request.form.get("phone_number", "").strip()
        phone_number = normalize_phone_number(phone_number) if phone_number else None

        should_send_welcome = bool(phone_number) and not user["phone_number"] and not user["welcome_sent"]

        conn = get_conn()
        conn.execute("UPDATE users SET phone_number = ? WHERE id = ?", (phone_number, session["user_id"]))
        if should_send_welcome:
            conn.execute("UPDATE users SET welcome_sent = 1 WHERE id = ?", (session["user_id"],))
        conn.commit()
        conn.close()
        session["phone_number"] = phone_number

        if should_send_welcome:
            send_direct_message(session["user_id"], phone_number, "welcome", msg_welcome(user["username"]))

        flash("Nomor telepon berhasil diperbarui!", "success")
        return redirect(url_for("settings"))

    content = render_template_string(
        """
<div class="row justify-content-center">
  <div class="col-md-6">
    <div class="card auth-card">
      <div class="card-body p-4">
        <h4 class="fw-bold mb-4"><i class="bi bi-gear me-2"></i>Pengaturan Akun</h4>
        <form method="post">
          <div class="mb-3">
            <label class="form-label fw-semibold">Username</label>
            <input type="text" class="form-control" value="{{ user['username'] }}" disabled>
          </div>
          <div class="mb-3">
            <label class="form-label fw-semibold">Email</label>
            <input type="email" class="form-control" value="{{ user['email'] }}" disabled>
          </div>
          <div class="mb-3 input-icon">
            <i class="bi bi-whatsapp"></i>
            <input type="tel" class="form-control" name="phone_number" value="{{ user['phone_number'] or '' }}" placeholder="08123456789">
            <small class="text-muted d-block mt-1">Masukkan nomor WhatsApp aktif untuk menerima notifikasi reminder streak (via Fonnte).</small>
          </div>
          <button type="submit" class="btn btn-primary w-100"><i class="bi bi-save me-2"></i>Simpan Perubahan</button>
          <a href="{{ url_for('index') }}" class="btn btn-outline-secondary w-100 mt-2"><i class="bi bi-arrow-left me-2"></i>Kembali</a>
        </form>
      </div>
    </div>
  </div>
</div>
""",
        user=user,
    )
    return render_template_string(BASE_TEMPLATE, title="Pengaturan", content=content)


# ==================== ROUTE: TRIGGER REMINDER (untuk CRON) ====================
@app.route("/admin/run-reminders", methods=["POST", "GET"])
def admin_run_reminders():
    token = request.args.get("token", "")
    if token != ADMIN_CRON_TOKEN:
        return {"status": "error", "message": "unauthorized"}, 401
    run_reminder_scheduler()
    return {"status": "ok", "message": "reminder scheduler dijalankan"}


# ==================== TESTING / DUMMY PANEL ====================
TESTING_SCENARIOS = {
    "fresh": "1. Baru daftar — streak 0, api padam, belum pernah upload",
    "on_today": "2. Baru upload hari ini — streak aktif, gap 0 hari (baseline, tidak ada reminder)",
    "warn_day4": "3. Sudah 3 hari tidak upload — peringatan H-1 sebelum padam",
    "dead": "4. Sudah 4+ hari tidak upload — streak resmi padam",
    "low_progress": "5. Progress kecil 1-3 hari terakhir — streak tetap aman, cuma diingatkan",
    "reactivated": "6. Upload lagi setelah api padam — notifikasi api menyala kembali",
    "keepalive_5": "7a. Api padam sejak 5 hari lalu — reminder lanjutkan skripsi",
    "keepalive_8": "7b. Api padam sejak 8 hari lalu — reminder lanjutkan skripsi",
    "stopped": "8. Api padam sejak 95 hari lalu (>3 bulan) — reminder dihentikan",
}


def create_dummy_comparison_today(user_id, difference_percent):
    similarity_percent = round(100 - difference_percent, 2)
    label = get_progress_label(difference_percent)
    today = today_str()
    now = now_dt().strftime("%Y-%m-%d %H:%M:%S")
    suffix = str(int(time.time() * 1000))
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO documents (user_id, original_name, saved_name, file_ext, uploaded_at, upload_date,
            raw_text, clean_text, kgrams_json, hashes_json, windows_json, fingerprints_json)
        VALUES (?, ?, ?, 'txt', ?, ?, '', '', '[]', '[]', '[]', '[]')
        """,
        (user_id, f"dummy_a_{suffix}.txt", f"dummy_a_{suffix}.txt", now, today),
    )
    doc1_id = cur.lastrowid
    cur.execute(
        """
        INSERT INTO documents (user_id, original_name, saved_name, file_ext, uploaded_at, upload_date,
            raw_text, clean_text, kgrams_json, hashes_json, windows_json, fingerprints_json)
        VALUES (?, ?, ?, 'txt', ?, ?, '', '', '[]', '[]', '[]', '[]')
        """,
        (user_id, f"dummy_b_{suffix}.txt", f"dummy_b_{suffix}.txt", now, today),
    )
    doc2_id = cur.lastrowid
    conn.commit()
    conn.close()
    upsert_comparison(user_id, doc1_id, doc2_id, similarity_percent, difference_percent, 0.0, label)


@app.route("/testing")
def testing_panel():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]
    user = get_user_by_id(user_id)
    conn = get_conn()
    cur = conn.cursor()
    _get_or_create_streak_row(cur, user_id)
    conn.commit()
    state = cur.execute("SELECT * FROM streak_state WHERE user_id = ?", (user_id,)).fetchone()
    logs = cur.execute(
        "SELECT * FROM reminder_log WHERE user_id = ? ORDER BY id DESC LIMIT 20", (user_id,)
    ).fetchall()
    dummy_count = cur.execute(
        "SELECT COUNT(*) AS c FROM documents WHERE user_id = ? AND original_name LIKE 'dummy_%'", (user_id,)
    ).fetchone()["c"]
    conn.close()
    flame = flame_meta(state["flame_status"])

    content = render_template_string(
        """
<div class="card card-shadow mb-4">
  <div class="card-body p-4">
    <h3 class="fw-bold mb-1"><i class="bi bi-flask me-2"></i>Testing Panel — Simulasi Semua Kondisi Reminder</h3>
    <p class="text-muted mb-3">
      Klik salah satu skenario di bawah untuk langsung mengubah status streak akunmu
      ke kondisi tertentu, lalu sistem otomatis mengecek &amp; mengirim reminder WA
      sesuai kondisi itu (kalau nomor WhatsApp di Settings sudah diisi).
    </p>
    {% if not user["phone_number"] %}
    <div class="alert alert-warning">
      <i class="bi bi-exclamation-triangle me-2"></i>
      Nomor WhatsApp kamu belum diisi di <a href="{{ url_for('settings') }}">Settings</a>.
      Skenario tetap bisa dites, tapi pesan tidak akan benar-benar terkirim (statusnya <code>skipped_no_phone</code>).
    </div>
    {% endif %}
    <div class="alert alert-info d-flex justify-content-between align-items-center flex-wrap gap-2">
      <div>
        <i class="bi bi-send me-2"></i>
        <b>Bingung kenapa WA nggak masuk?</b> Coba dulu kirim test langsung
        (bypass semua logic streak) buat mastiin Fonnte-nya sendiri sudah benar.
      </div>
      <form method="post" action="{{ url_for('testing_send_test') }}" class="m-0">
        <button class="btn btn-dark btn-sm" type="submit"><i class="bi bi-send-fill me-1"></i>Test Kirim WA Sekarang</button>
      </form>
    </div>
    <div class="row g-2 mb-3">
      <div class="col-md-3"><div class="soft-card text-center"><div class="small text-muted">Streak</div><div class="fw-bold fs-4">{{ state["streak_count"] }}</div></div></div>
      <div class="col-md-3"><div class="soft-card text-center"><div class="small text-muted">Status Api</div><div class="fw-bold">{{ flame.icon }} {{ flame.text }}</div></div></div>
      <div class="col-md-3"><div class="soft-card text-center"><div class="small text-muted">Upload Valid Terakhir</div><div class="fw-bold state-pill">{{ state["last_valid_upload_date"] or "-" }}</div></div></div>
      <div class="col-md-3"><div class="soft-card text-center"><div class="small text-muted">Api Padam Sejak</div><div class="fw-bold state-pill">{{ state["flame_off_date"] or "-" }}</div></div></div>
    </div>
    <form method="post" action="{{ url_for('testing_reset') }}" class="d-inline">
      <button class="btn btn-outline-secondary btn-sm" type="submit"><i class="bi bi-arrow-counterclockwise me-1"></i>Reset Status ke Awal</button>
    </form>
    <form method="post" action="{{ url_for('testing_cleanup') }}" class="d-inline">
      <button class="btn btn-outline-secondary btn-sm" type="submit"><i class="bi bi-trash me-1"></i>Hapus {{ dummy_count }} Dokumen Dummy</button>
    </form>
  </div>
</div>
<div class="row g-3 mb-4">
  {% for key, label in scenarios.items() %}
  <div class="col-md-6 col-lg-4">
    <div class="scenario-card d-flex flex-column">
      <h6>{{ label }}</h6>
      <p class="flex-grow-1"></p>
      <form method="post" action="{{ url_for('testing_simulate', scenario=key) }}">
        <button class="btn btn-dark btn-sm w-100" type="submit"><i class="bi bi-play-fill me-1"></i>Jalankan Skenario</button>
      </form>
    </div>
  </div>
  {% endfor %}
</div>
<div class="card card-shadow">
  <div class="card-body p-4">
    <h5 class="fw-bold mb-3"><i class="bi bi-clock-history me-2"></i>20 Riwayat Reminder Terakhir</h5>
    <div class="table-responsive">
      <table class="table table-striped">
        <thead><tr><th>Waktu</th><th>Tipe</th><th>Status</th><th>Nomor</th></tr></thead>
        <tbody>
          {% if logs %}
            {% for log in logs %}
            <tr>
              <td>{{ log["sent_at"] }}</td>
              <td><span class="badge badge-flame-on">{{ log["reminder_type"] }}</span></td>
              <td>{{ log["status"] }}</td>
              <td>{{ log["phone_number"] }}</td>
            </tr>
            {% endfor %}
          {% else %}
            <tr><td colspan="4" class="text-center">Belum ada riwayat reminder.</td></tr>
          {% endif %}
        </tbody>
      </table>
    </div>
  </div>
</div>
""",
        state=state, flame=flame, user=user, scenarios=TESTING_SCENARIOS, logs=logs, dummy_count=dummy_count,
    )
    return render_template_string(BASE_TEMPLATE, title="Testing", content=content)


@app.route("/testing/simulate/<scenario>", methods=["POST"])
def testing_simulate(scenario):
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]
    if scenario not in TESTING_SCENARIOS:
        flash("Skenario tidak dikenal.", "danger")
        return redirect(url_for("testing_panel"))

    today = today_str()

    def date_offset(days):
        return (now_dt() - timedelta(days=days)).strftime("%Y-%m-%d")

    conn = get_conn()
    cur = conn.cursor()
    _get_or_create_streak_row(cur, user_id)
    conn.commit()

    def set_state(streak_count, flame_status, last_valid, flame_off):
        cur.execute(
            """
            UPDATE streak_state
            SET streak_count = ?, flame_status = ?, last_valid_upload_date = ?, flame_off_date = ?,
                last_reminder_type = NULL, last_reminder_date = NULL, thesis_finished = 0
            WHERE user_id = ?
            """,
            (streak_count, flame_status, last_valid, flame_off, user_id),
        )
        conn.commit()

    if scenario == "fresh":
        set_state(0, "off", None, None)
    elif scenario == "on_today":
        set_state(3, "on", today, None)
    elif scenario == "warn_day4":
        set_state(3, "on", date_offset(DAYS_BEFORE_WARNING), None)
    elif scenario == "dead":
        set_state(5, "on", date_offset(DAYS_BEFORE_DEATH), None)
    elif scenario == "low_progress":
        set_state(2, "on", today, None)
    elif scenario == "reactivated":
        set_state(0, "off", date_offset(10), date_offset(5))
    elif scenario == "keepalive_5":
        set_state(0, "off", date_offset(12), date_offset(KEEPALIVE_START_DAYS))
    elif scenario == "keepalive_8":
        set_state(0, "off", date_offset(15), date_offset(KEEPALIVE_START_DAYS + KEEPALIVE_INTERVAL_DAYS))
    elif scenario == "stopped":
        set_state(0, "off", date_offset(100), date_offset(95))

    conn.close()

    if scenario == "low_progress":
        create_dummy_comparison_today(user_id, difference_percent=5)

    if scenario == "reactivated":
        update_streak_after_valid_upload(user_id)
    else:
        sync_streak_and_notify(user_id)

    conn = get_conn()
    last_log = conn.execute(
        "SELECT * FROM reminder_log WHERE user_id = ? ORDER BY id DESC LIMIT 1", (user_id,)
    ).fetchone()
    conn.close()

    label = TESTING_SCENARIOS[scenario]
    if last_log:
        detail = ""
        if last_log["response"]:
            try:
                resp_obj = json.loads(last_log["response"])
                detail = f" — Detail: {json.dumps(resp_obj)[:250]}"
            except Exception:
                detail = f" — Detail: {str(last_log['response'])[:250]}"
        category = "success" if last_log["status"] == "success" else ("warning" if last_log["status"] == "skipped_no_phone" else "danger")
        flash(f"Skenario '{label}' dijalankan. Reminder: [{last_log['reminder_type']}] status={last_log['status']}{detail}", category)
    else:
        flash(f"Skenario '{label}' dijalankan. Tidak ada reminder yang perlu dikirim untuk kondisi ini.", "info")
    return redirect(url_for("testing_panel"))


@app.route("/testing/send-test", methods=["POST"])
def testing_send_test():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]
    user = get_user_by_id(user_id)
    if not user["phone_number"]:
        flash("Isi dulu nomor WhatsApp di Settings sebelum coba kirim test.", "warning")
        return redirect(url_for("testing_panel"))
    if not FONNTE_API_KEY:
        flash("FONNTE_API_KEY masih kosong di .env. Isi dulu, lalu restart aplikasi.", "danger")
        return redirect(url_for("testing_panel"))

    msg = (
        "Ini pesan TEST dari Skripsiku.\n\n"
        "Kalau kamu terima pesan ini, artinya konfigurasi Fonnte kamu sudah benar "
        "dan reminder streak akan berfungsi normal."
    )
    result = send_whatsapp_reminder(user["phone_number"], msg)
    status = result.get("status", "error")
    log_reminder(user_id, "manual_test", user["phone_number"], status, json.dumps(result))
    if status == "success":
        flash(f"Request kirim ke Fonnte berhasil (target: {user['phone_number']}). Cek WhatsApp kamu dalam beberapa detik.", "success")
    else:
        flash(f"Fonnte menolak/gagal mengirim. Detail: {json.dumps(result)[:300]}", "danger")
    return redirect(url_for("testing_panel"))


@app.route("/testing/cleanup", methods=["POST"])
def testing_cleanup():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]
    conn = get_conn()
    cur = conn.cursor()
    dummy_ids = [
        r["id"] for r in cur.execute(
            "SELECT id FROM documents WHERE user_id = ? AND original_name LIKE 'dummy_%'", (user_id,)
        ).fetchall()
    ]
    for did in dummy_ids:
        cur.execute("DELETE FROM comparisons WHERE user_id = ? AND (doc1_id = ? OR doc2_id = ?)", (user_id, did, did))
        cur.execute("DELETE FROM documents WHERE id = ?", (did,))
    conn.commit()
    conn.close()
    flash(f"{len(dummy_ids)} dokumen dummy berhasil dihapus.", "success")
    return redirect(url_for("testing_panel"))


@app.route("/testing/reset", methods=["POST"])
def testing_reset():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))
    user_id = session["user_id"]
    conn = get_conn()
    conn.execute(
        """
        UPDATE streak_state
        SET streak_count = 0, flame_status = 'off', last_valid_upload_date = NULL,
            flame_off_date = NULL, last_reminder_type = NULL, last_reminder_date = NULL,
            thesis_finished = 0, first_upload_notified = 0
        WHERE user_id = ?
        """,
        (user_id,),
    )
    conn.execute("DELETE FROM reminder_log WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    flash("Status streak & riwayat reminder direset ke kondisi awal.", "success")
    return redirect(url_for("testing_panel"))


# ==================== RUN APP ====================
if __name__ == "__main__":
    init_db()
    if not FONNTE_API_KEY:
        print("=" * 70)
        print(" FONNTE_API_KEY KOSONG! WhatsApp reminder TIDAK AKAN TERKIRIM.")
        print(" Isi FONNTE_API_KEY di file .env, lalu restart aplikasi ini.")
        print("=" * 70)
    else:
        _masked = FONNTE_API_KEY[:4] + "..." + FONNTE_API_KEY[-4:] if len(FONNTE_API_KEY) > 8 else "(pendek)"
        print(f" FONNTE_API_KEY terbaca dari .env: {_masked}")

    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        scheduler = BackgroundScheduler(daemon=True)
        scheduler.add_job(run_reminder_scheduler, "cron", hour=REMINDER_HOUR, minute=REMINDER_MINUTE)
        scheduler.start()
        print(f"APScheduler aktif: reminder akan dicek otomatis tiap jam {REMINDER_HOUR:02d}:{REMINDER_MINUTE:02d}.")
    except ImportError:
        print("APScheduler tidak terinstall — reminder tidak berjalan otomatis.")
        print("Install dengan: pip install apscheduler")
        print("Atau panggil endpoint /admin/run-reminders?token=... lewat cron job harian.")

    app.run(debug=True, host="0.0.0.0", port=5000)
