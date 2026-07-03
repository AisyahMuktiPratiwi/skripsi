import json
import hashlib
import re
import sqlite3
import time
from datetime import datetime, timedelta
from itertools import combinations
from pathlib import Path
import requests

from bs4 import BeautifulSoup
from docx import Document
from flask import Flask, flash, redirect, render_template_string, request, url_for, session
from PyPDF2 import PdfReader
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

# ==================== KONFIGURASI ====================
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
DB_PATH = BASE_DIR / "similarity.db"

UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"txt", "pdf", "docx", "html", "htm"}

K_GRAM = 5
WINDOW_SIZE = 4

# Konfigurasi Fonnte WhatsApp API
FONNTE_API_KEY = "K73cbVVpqTVUgc3kZk4n"  # Ganti dengan API Key Fonnte Anda
FONNTE_BASE_URL = "https://api.fonnte.com"

app = Flask(__name__)
app.secret_key = "8ea93c21ee2b0060a2fb9dbd5f61c783ce940f653db8beb87809dcf6bac901c9"  # Ganti dengan secret key yang aman
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

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

# ==================== INISIALISASI DATABASE ====================
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    # Tabel users untuk login/registrasi
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            phone_number TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    # Tabel documents
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

    # Tabel comparisons
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

    # Tabel streak_state - DIPERBAIKI: tanpa CHECK (id = 1)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS streak_state (
            user_id INTEGER PRIMARY KEY,
            streak_count INTEGER NOT NULL DEFAULT 0,
            last_valid_upload_date TEXT,
            flame_status TEXT NOT NULL DEFAULT 'off',
            last_reminder_sent TEXT,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        """
    )

    # Tabel reminder_log untuk tracking pengiriman reminder
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
    conn.close()

# ==================== FUNGSI AUTHENTIKASI ====================
def get_user_by_username(username):
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return row

def get_user_by_id(user_id):
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
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
            (username, email, password_hash, phone_number, created_at)
        )
        user_id = cur.lastrowid
        
        # Inisialisasi streak_state untuk user baru
        cur.execute(
            """
            INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
            VALUES (?, 0, NULL, 'off')
            """,
            (user_id,)
        )
        
        conn.commit()
        conn.close()
        return user_id
    except sqlite3.IntegrityError as e:
        conn.close()
        raise e

# ==================== FUNGSI REMINDER WA (FONNTE) ====================
def send_whatsapp_reminder(phone_number, message):
    """
    Mengirim pesan WhatsApp menggunakan Fonnte API
    """
    if not FONNTE_API_KEY or FONNTE_API_KEY == "YOUR_FONNTE_API_KEY":
        return {"status": "error", "message": "API Key belum dikonfigurasi"}
    
    url = f"{FONNTE_BASE_URL}/send"
    
    headers = {
        "Authorization": FONNTE_API_KEY
    }
    
    data = {
        "target": phone_number,
        "message": message,
        "countryCode": "62",  # Kode negara Indonesia
    }
    
    try:
        response = requests.post(url, headers=headers, data=data)
        return {
            "status": "success" if response.status_code == 200 else "error",
            "response": response.json() if response.status_code == 200 else response.text
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

def log_reminder(user_id, reminder_type, phone_number, status, response=None):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO reminder_log (user_id, reminder_type, sent_at, phone_number, status, response)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (user_id, reminder_type, now_dt().strftime("%Y-%m-%d %H:%M:%S"), phone_number, status, response)
    )
    conn.commit()
    conn.close()
def get_users_needing_reminder():
    """
    Mendapatkan user yang perlu dikirim reminder:
    - Sudah 3 hari tidak upload (hari ke-4 streak akan reset)
    - Memiliki nomor telepon terdaftar
    - Belum dikirim reminder hari ini
    """
    conn = get_conn()
    cur = conn.cursor()
    
    today = today_str()
    three_days_ago = (parse_date(today) - timedelta(days=3)).strftime("%Y-%m-%d")
    
    # PERBAIKI: Tambahkan u.id AS user_id
    rows = cur.execute(
        """
        SELECT 
            u.id AS user_id,  -- <-- INI PERUBAHANNYA!
            u.username, 
            u.phone_number,
            s.streak_count, 
            s.last_valid_upload_date, 
            s.flame_status,
            s.last_reminder_sent
        FROM streak_state s
        JOIN users u ON s.user_id = u.id
        WHERE s.flame_status = 'on'
        AND s.last_valid_upload_date <= ?
        AND u.phone_number IS NOT NULL
        AND u.phone_number != ''
        AND (s.last_reminder_sent IS NULL OR s.last_reminder_sent != ?)
        """,
        (three_days_ago, today)
    ).fetchall()
    
    conn.close()
    return rows
def check_and_send_reminders():
    """
    Cek user yang membutuhkan reminder dan kirim notifikasi
    """
    users = get_users_needing_reminder()
    
    for user in users:
        last_date = parse_date(user["last_valid_upload_date"])
        today = parse_date(today_str())
        days_gap = (today - last_date).days
        
        if days_gap >= 3:  # Sudah 3 hari tidak upload
            message = (
                f"Halo {user['username']}! 👋\n\n"
                f"Kami perhatikan Anda sudah {days_gap} hari tidak mengupload progres skripsi.\n\n"
                f"⚠️ PERINGATAN: Jika besok (hari ke-4) Anda masih belum upload, maka STREAK akan RESET ke 0 dan API akan PADAM! 🔥\n\n"
                f"Segera upload dokumen terbaru Anda di: http://your-domain.com\n\n"
                f"Tetap semangat menyelesaikan skripsinya! 💪"
            )
        else:
            continue
        
        # Kirim reminder
        result = send_whatsapp_reminder(user["phone_number"], message)
        
        # Log pengiriman
        status = result.get("status", "error")
        response = json.dumps(result) if result else None
        
        log_reminder(
            user["user_id"],
            "daily_reminder",
            user["phone_number"],
            status,
            response
        )
        
        # Update last_reminder_sent jika berhasil
        if status == "success":
            conn = get_conn()
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE streak_state
                SET last_reminder_sent = ?
                WHERE user_id = ?
                """,
                (today_str(), user["user_id"])
            )
            conn.commit()
            conn.close()

# ==================== FUNGSI STREAK (DENGAN USER ID) ====================
def get_or_reset_streak_state(user_id):
    conn = get_conn()
    cur = conn.cursor()

    row = cur.execute(
        "SELECT * FROM streak_state WHERE user_id = ?",
        (user_id,)
    ).fetchone()

    if not row:
        # Jika belum ada streak_state untuk user ini, buat baru
        cur.execute(
            """
            INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
            VALUES (?, 0, NULL, 'off')
            """,
            (user_id,)
        )
        conn.commit()
        row = cur.execute(
            "SELECT * FROM streak_state WHERE user_id = ?",
            (user_id,)
        ).fetchone()

    streak_count = row["streak_count"]
    last_valid_upload_date = row["last_valid_upload_date"]
    flame_status = row["flame_status"]

    # Cek jika sudah 4 hari tidak upload -> reset streak
    if last_valid_upload_date:
        last_date = parse_date(last_valid_upload_date)
        today = parse_date(today_str())
        gap_days = (today - last_date).days

        if gap_days >= 4:
            streak_count = 0
            flame_status = "off"

            cur.execute(
                """
                UPDATE streak_state
                SET streak_count = ?, flame_status = ?
                WHERE user_id = ?
                """,
                (streak_count, flame_status, user_id),
            )
            conn.commit()

    # Ambil ulang data terbaru
    row = cur.execute(
        "SELECT * FROM streak_state WHERE user_id = ?",
        (user_id,)
    ).fetchone()
    
    conn.close()
    return row

def update_streak_after_valid_upload(user_id):
    conn = get_conn()
    cur = conn.cursor()
    
    row = cur.execute(
        "SELECT * FROM streak_state WHERE user_id = ?",
        (user_id,)
    ).fetchone()
    
    if not row:
        cur.execute(
            """
            INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
            VALUES (?, 0, NULL, 'off')
            """,
            (user_id,)
        )
        conn.commit()
        row = cur.execute(
            "SELECT * FROM streak_state WHERE user_id = ?",
            (user_id,)
        ).fetchone()

    streak_count = row["streak_count"]
    last_valid_upload_date = row["last_valid_upload_date"]
    today = today_str()

    # Cek apakah hari ini sudah upload
    if last_valid_upload_date == today:
        conn.close()
        return

    # Cek apakah upload ini valid (ada perubahan > 0%)
    # Validasi dilakukan di fungsi panggil dengan cek best similarity < 100%
    
    if last_valid_upload_date is None:
        streak_count = 1
    else:
        gap_days = (parse_date(today) - parse_date(last_valid_upload_date)).days

        if gap_days == 1:
            streak_count += 1
        elif gap_days >= 4:
            streak_count = 1
        else:
            streak_count = 1

    flame_status = "on"

    cur.execute(
        """
        UPDATE streak_state
        SET streak_count = ?, last_valid_upload_date = ?, flame_status = ?, last_reminder_sent = NULL
        WHERE user_id = ?
        """,
        (streak_count, today, flame_status, user_id),
    )

    conn.commit()
    conn.close()

# ==================== FUNGSI FILE PROCESSING (DENGAN USER ID) ====================
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")

def extract_text_from_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    texts = []
    for page in reader.pages:
        texts.append(page.extract_text() or "")
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

    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))
    return round((intersection / union) * 100, 2)

def calculate_difference(similarity_percent: float) -> float:
    return round(100.0 - similarity_percent, 2)

def get_progress_label(difference_percent: float) -> str:
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

# ==================== FUNGSI DATABASE OPERATIONS (DENGAN USER ID) ====================
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
            user_id,
            original_name,
            saved_name,
            ext,
            uploaded_at,
            upload_date,
            processed["raw_text"],
            processed["clean_text"],
            json.dumps(processed["kgrams"]),
            json.dumps(processed["hashes"]),
            json.dumps(processed["windows"]),
            json.dumps(processed["fingerprints"]),
        ),
    )

    doc_id = cur.lastrowid
    conn.commit()
    conn.close()
    return doc_id

def get_documents_by_user(user_id):
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT * FROM documents WHERE user_id = ? ORDER BY id DESC",
        (user_id,)
    ).fetchall()
    conn.close()
    return rows

def get_documents_by_user_asc(user_id):
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        "SELECT * FROM documents WHERE user_id = ? ORDER BY id ASC",
        (user_id,)
    ).fetchall()
    conn.close()
    return rows

def get_document_by_id_and_user(doc_id, user_id):
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute(
        "SELECT * FROM documents WHERE id = ? AND user_id = ?",
        (doc_id, user_id)
    ).fetchone()
    conn.close()
    return row

def get_fingerprints_from_row(row):
    return json.loads(row["fingerprints_json"]) if row["fingerprints_json"] else []

def upsert_comparison(user_id, doc1_id, doc2_id, similarity_percent, difference_percent, process_time, progress_label):
    a, b = sorted([doc1_id, doc2_id])

    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO comparisons (
            user_id, doc1_id, doc2_id, similarity_percent, difference_percent,
            process_time, progress_label, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(doc1_id, doc2_id)
        DO UPDATE SET
            similarity_percent = excluded.similarity_percent,
            difference_percent = excluded.difference_percent,
            process_time = excluded.process_time,
            progress_label = excluded.progress_label,
            created_at = excluded.created_at
        """,
        (
            user_id,
            a,
            b,
            similarity_percent,
            difference_percent,
            process_time,
            progress_label,
            now_dt().strftime("%Y-%m-%d %H:%M:%S"),
        ),
    )

    conn.commit()
    conn.close()

def compare_new_document_against_all(user_id, new_doc_id):
    new_doc = get_document_by_id_and_user(new_doc_id, user_id)
    if not new_doc:
        return

    new_fp = get_fingerprints_from_row(new_doc)
    all_docs = get_documents_by_user_asc(user_id)

    for old_doc in all_docs:
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
    cur = conn.cursor()

    rows = cur.execute(
        """
        SELECT
            c.*,
            d1.original_name AS doc1_name,
            d2.original_name AS doc2_name
        FROM comparisons c
        JOIN documents d1 ON c.doc1_id = d1.id
        JOIN documents d2 ON c.doc2_id = d2.id
        WHERE c.user_id = ?
        ORDER BY c.difference_percent DESC, c.id DESC
        """,
        (user_id,)
    ).fetchall()

    conn.close()
    return rows

def get_today_best_highlight(user_id):
    conn = get_conn()
    cur = conn.cursor()

    rows = cur.execute(
        """
        SELECT
            c.*,
            d1.original_name AS doc1_name,
            d2.original_name AS doc2_name
        FROM comparisons c
        JOIN documents d1 ON c.doc1_id = d1.id
        JOIN documents d2 ON c.doc2_id = d2.id
        WHERE c.user_id = ?
        AND (d1.upload_date = ? OR d2.upload_date = ?)
        ORDER BY c.similarity_percent DESC, c.id DESC
        """,
        (user_id, today_str(), today_str()),
    ).fetchall()

    conn.close()

    if not rows:
        return None

    for row in rows:
        if float(row["similarity_percent"]) < 100.0:
            return row

    return rows[0] if rows else None

def get_upload_days_by_user(user_id):
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute(
        """
        SELECT upload_date, COUNT(*) AS total
        FROM documents
        WHERE user_id = ?
        GROUP BY upload_date
        ORDER BY upload_date DESC
        """,
        (user_id,)
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
        d1 = docs[i]
        d2 = docs[i + 1]
        a, b = sorted([d1["id"], d2["id"]])

        row = cur.execute(
            """
            SELECT * FROM comparisons 
            WHERE user_id = ? AND doc1_id = ? AND doc2_id = ?
            """,
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
    """
    Ambil best progress per hari berdasarkan difference tertinggi
    """
    conn = get_conn()
    cur = conn.cursor()

    dates = cur.execute(
        """
        SELECT DISTINCT upload_date
        FROM documents
        WHERE user_id = ?
        ORDER BY upload_date DESC
        """,
        (user_id,)
    ).fetchall()

    results = []

    for d in dates:
        upload_date = d["upload_date"]

        rows = cur.execute(
            """
            SELECT
                c.*,
                d1.upload_date AS doc1_date,
                d2.upload_date AS doc2_date
            FROM comparisons c
            JOIN documents d1 ON c.doc1_id = d1.id
            JOIN documents d2 ON c.doc2_id = d2.id
            WHERE c.user_id = ?
            AND (d1.upload_date = ? OR d2.upload_date = ?)
            ORDER BY c.difference_percent DESC, c.id DESC
            """,
            (user_id, upload_date, upload_date),
        ).fetchall()

        if rows:
            best = rows[0]
            results.append(
                {
                    "upload_date": upload_date,
                    "best_difference": float(best["difference_percent"]),
                    "best_similarity": float(best["similarity_percent"]),
                    "progress_label": best["progress_label"],
                }
            )

    conn.close()
    
    if not results:
        return None

    recent = results[:3]
    values = [item["best_difference"] for item in recent if item["best_difference"] is not None]

    if not values:
        return None

    avg_diff = round(sum(values) / len(values), 2)

    if avg_diff <= 10:
        label = "Perubahan sangat kecil"
        message = "Dalam 1–3 hari terakhir, progress kamu masih sangat kecil. Coba tambahkan perubahan yang lebih signifikan."
        alert = "warning"
    elif avg_diff <= 30:
        label = "Progres ringan"
        message = "Dalam 1–3 hari terakhir, progress kamu masih ringan. Ayo tingkatkan progress pekerjaanmu."
        alert = "warning"
    elif avg_diff <= 60:
        label = "Progres sedang"
        message = "Dalam 1–3 hari terakhir, progress kamu sudah cukup terlihat."
        alert = "info"
    else:
        label = "Progres besar"
        message = "Dalam 1–3 hari terakhir, progress kamu sangat baik dan terlihat signifikan."
        alert = "success"

    return {
        "avg_difference": avg_diff,
        "label": label,
        "message": message,
        "alert": alert,
        "days_count": len(values),
    }

def flame_meta(flame_status: str):
    if flame_status == "on":
        return {
            "icon": "🔥",
            "badge_class": "bg-danger",
            "text": "Api menyala",
        }
    return {
        "icon": "🪨",
        "badge_class": "bg-secondary",
        "text": "Api padam",
    }

# ==================== TEMPLATE BASE ====================
BASE_TEMPLATE = """
<!doctype html>
<html lang="id">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ title }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <style>
        body {
            background: linear-gradient(135deg, #f6f7fb 0%, #e9edf5 100%);
            min-height: 100vh;
        }
        .card-shadow {
            box-shadow: 0 10px 24px rgba(0,0,0,0.08);
            border: none;
            border-radius: 18px;
        }
        .mono-box {
            background: #111827;
            color: #f9fafb;
            border-radius: 12px;
            padding: 16px;
            white-space: pre-wrap;
            word-break: break-word;
            max-height: 320px;
            overflow-y: auto;
            font-family: Consolas, monospace;
            font-size: 13px;
        }
        .small-scroll {
            max-height: 240px;
            overflow-y: auto;
            font-family: Consolas, monospace;
            font-size: 13px;
            background: #fff;
            padding: 12px;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .hero-highlight {
            background: linear-gradient(135deg, #111827, #1f2937);
            color: white;
            border-radius: 20px;
        }
        .streak-icon {
            font-size: 44px;
            line-height: 1;
        }
        .big-number {
            font-size: 40px;
            font-weight: 700;
        }
        .soft-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 18px;
            border: 1px solid #eceff4;
        }
        .auth-card {
            border-radius: 24px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.12);
            overflow: hidden;
        }
        .auth-card .card-header {
            background: linear-gradient(135deg, #1a1a2e, #16213e);
            color: white;
            padding: 30px 30px 20px;
            border: none;
        }
        .auth-card .card-body {
            padding: 35px 30px;
        }
        .auth-card .form-control {
            border-radius: 12px;
            padding: 12px 16px;
            border: 2px solid #e5e7eb;
            transition: all 0.3s ease;
        }
        .auth-card .form-control:focus {
            border-color: #1a1a2e;
            box-shadow: 0 0 0 4px rgba(26, 26, 46, 0.1);
        }
        .auth-card .btn-primary {
            background: linear-gradient(135deg, #1a1a2e, #2d2d44);
            border: none;
            border-radius: 12px;
            padding: 12px;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        .auth-card .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(26, 26, 46, 0.3);
        }
        .auth-card .btn-outline-secondary {
            border-radius: 12px;
            padding: 12px;
            font-weight: 600;
        }
        .auth-logo {
            font-size: 48px;
            margin-bottom: 10px;
        }
        .auth-title {
            font-weight: 700;
            font-size: 28px;
        }
        .auth-subtitle {
            opacity: 0.8;
            font-size: 14px;
        }
        .input-icon {
            position: relative;
        }
        .input-icon .form-control {
            padding-left: 45px;
        }
        .input-icon .bi {
            position: absolute;
            left: 15px;
            top: 50%;
            transform: translateY(-50%);
            color: #9ca3af;
            font-size: 18px;
        }
        .navbar-custom {
            background: linear-gradient(135deg, #1a1a2e, #16213e) !important;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
        .btn-settings {
            background: rgba(255,255,255,0.1);
            border: 1px solid rgba(255,255,255,0.2);
            color: white;
        }
        .btn-settings:hover {
            background: rgba(255,255,255,0.2);
            color: white;
        }
    </style>
</head>
<body>

<!-- Navbar -->
<nav class="navbar navbar-expand-lg navbar-custom mb-4">
    <div class="container">
        <a class="navbar-brand fw-bold text-white" href="{{ url_for('index') }}">
            <i class="bi bi-graph-up-arrow me-2"></i>Winnowing Progress
        </a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNav">
            <ul class="navbar-nav ms-auto">
                {% if session.user_id %}
                    <li class="nav-item">
                        <span class="nav-link text-white-50">
                            <i class="bi bi-person-circle me-1"></i> {{ session.username }}
                        </span>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link text-white" href="{{ url_for('index') }}">
                            <i class="bi bi-house me-1"></i>Home
                        </a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link text-white" href="{{ url_for('documents') }}">
                            <i class="bi bi-file-earmark me-1"></i>Dokumen
                        </a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link text-white" href="{{ url_for('comparisons') }}">
                            <i class="bi bi-bar-chart me-1"></i>Progress
                        </a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link text-white" href="{{ url_for('settings') }}">
                            <i class="bi bi-gear me-1"></i>Settings
                        </a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link text-danger" href="{{ url_for('logout') }}">
                            <i class="bi bi-box-arrow-right me-1"></i>Logout
                        </a>
                    </li>
                {% else %}
                    <li class="nav-item">
                        <a class="nav-link text-white" href="{{ url_for('login') }}">
                            <i class="bi bi-box-arrow-in-right me-1"></i>Login
                        </a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link btn btn-outline-light btn-sm px-3" href="{{ url_for('register') }}">
                            <i class="bi bi-person-plus me-1"></i>Register
                        </a>
                    </li>
                {% endif %}
            </ul>
        </div>
    </div>
</nav>

<!-- Flash Messages -->
<div class="container">
    {% with messages = get_flashed_messages(with_categories=true) %}
        {% if messages %}
            {% for category, msg in messages %}
                <div class="alert alert-{{ category or 'info' }} alert-dismissible fade show shadow-sm" role="alert">
                    <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' if category == 'warning' else 'info-circle' }} me-2"></i>
                    {{ msg }}
                    <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                </div>
            {% endfor %}
        {% endif %}
    {% endwith %}
</div>

<!-- Content -->
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

        # Format nomor telepon (untuk Fonnte)
        if phone_number:
            phone_number = re.sub(r"[^0-9]", "", phone_number)
            if phone_number.startswith("0"):
                phone_number = "62" + phone_number[1:]
            elif not phone_number.startswith("62"):
                phone_number = "62" + phone_number

        try:
            user_id = create_user(username, email, password, phone_number)
            flash("🎉 Registrasi berhasil! Silakan login.", "success")
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
            <title>Register - Winnowing Progress</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
            <style>
                * {
                    font-family: 'Inter', sans-serif;
                }
                body {
                    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
                    min-height: 100vh;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    padding: 20px;
                }
                .register-card {
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(20px);
                    -webkit-backdrop-filter: blur(20px);
                    border-radius: 32px;
                    padding: 50px 45px;
                    max-width: 460px;
                    width: 100%;
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    box-shadow: 0 40px 80px rgba(0, 0, 0, 0.6);
                    animation: fadeInUp 0.8s ease-out;
                }
                @keyframes fadeInUp {
                    from {
                        opacity: 0;
                        transform: translateY(40px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }
                .register-card .logo {
                    text-align: center;
                    margin-bottom: 32px;
                }
                .register-card .logo .icon {
                    font-size: 48px;
                    background: linear-gradient(135deg, #f093fb, #f5576c);
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                    display: inline-block;
                }
                .register-card .logo h1 {
                    color: #fff;
                    font-size: 28px;
                    font-weight: 700;
                    margin-top: 8px;
                    letter-spacing: -0.5px;
                }
                .register-card .logo p {
                    color: rgba(255, 255, 255, 0.5);
                    font-size: 14px;
                    font-weight: 400;
                    margin-top: 4px;
                }
                .register-card .form-group {
                    margin-bottom: 18px;
                }
                .register-card .form-group label {
                    color: rgba(255, 255, 255, 0.7);
                    font-size: 13px;
                    font-weight: 500;
                    margin-bottom: 6px;
                    display: block;
                }
                .register-card .form-group .input-wrapper {
                    position: relative;
                }
                .register-card .form-group .input-wrapper .bi {
                    position: absolute;
                    left: 16px;
                    top: 50%;
                    transform: translateY(-50%);
                    color: rgba(255, 255, 255, 0.3);
                    font-size: 18px;
                    transition: color 0.3s ease;
                }
                .register-card .form-group .input-wrapper input {
                    width: 100%;
                    padding: 14px 16px 14px 48px;
                    background: rgba(255, 255, 255, 0.06);
                    border: 2px solid rgba(255, 255, 255, 0.08);
                    border-radius: 14px;
                    color: #fff;
                    font-size: 15px;
                    font-weight: 400;
                    transition: all 0.3s ease;
                    outline: none;
                }
                .register-card .form-group .input-wrapper input::placeholder {
                    color: rgba(255, 255, 255, 0.25);
                }
                .register-card .form-group .input-wrapper input:focus {
                    border-color: rgba(245, 87, 108, 0.5);
                    background: rgba(255, 255, 255, 0.08);
                    box-shadow: 0 0 0 4px rgba(245, 87, 108, 0.1);
                }
                .register-card .form-group .input-wrapper input:focus + .bi,
                .register-card .form-group .input-wrapper input:focus ~ .bi {
                    color: #f5576c;
                }
                .register-card .form-group .input-wrapper input:focus {
                    border-color: rgba(245, 87, 108, 0.5);
                    background: rgba(255, 255, 255, 0.08);
                    box-shadow: 0 0 0 4px rgba(245, 87, 108, 0.1);
                }
                .register-card .form-group .input-wrapper input:focus + .bi,
                .register-card .form-group .input-wrapper input:focus ~ .bi {
                    color: #f5576c;
                }
                .register-card .btn-register {
                    width: 100%;
                    padding: 14px;
                    background: linear-gradient(135deg, #f093fb, #f5576c);
                    border: none;
                    border-radius: 14px;
                    color: #fff;
                    font-size: 16px;
                    font-weight: 600;
                    transition: all 0.3s ease;
                    cursor: pointer;
                    margin-top: 8px;
                }
                .register-card .btn-register:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 12px 30px rgba(245, 87, 108, 0.35);
                }
                .register-card .btn-register:active {
                    transform: translateY(0);
                }
                .register-card .login-link {
                    text-align: center;
                    margin-top: 20px;
                    color: rgba(255, 255, 255, 0.4);
                    font-size: 14px;
                }
                .register-card .login-link a {
                    color: #f5576c;
                    text-decoration: none;
                    font-weight: 600;
                    transition: color 0.3s ease;
                }
                .register-card .login-link a:hover {
                    color: #f093fb;
                }
                .register-card .hint {
                    color: rgba(255, 255, 255, 0.3);
                    font-size: 12px;
                    margin-top: 4px;
                    display: block;
                }
                .alert-custom {
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(10px);
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    border-radius: 14px;
                    padding: 12px 16px;
                    color: #fff;
                    font-size: 14px;
                    margin-bottom: 20px;
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }
                .alert-custom .bi {
                    font-size: 18px;
                }
                .alert-custom.danger {
                    border-color: rgba(245, 87, 108, 0.3);
                    background: rgba(245, 87, 108, 0.1);
                }
                .alert-custom.danger .bi {
                    color: #f5576c;
                }
                .alert-custom.success {
                    border-color: rgba(52, 211, 153, 0.3);
                    background: rgba(52, 211, 153, 0.1);
                }
                .alert-custom.success .bi {
                    color: #34d399;
                }
                @media (max-width: 480px) {
                    .register-card {
                        padding: 32px 24px;
                    }
                }
            </style>
        </head>
        <body>
            <div class="register-card">
                <div class="logo">
                    <span class="icon">🚀</span>
                    <h1>Mulai Sekarang</h1>
                    <p>Daftar dan pantau progres skripsimu</p>
                </div>

                {% with messages = get_flashed_messages(with_categories=true) %}
                    {% if messages %}
                        {% for category, msg in messages %}
                            <div class="alert-custom {{ category }}">
                                <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' }}"></i>
                                {{ msg }}
                            </div>
                        {% endfor %}
                    {% endif %}
                {% endwith %}

                <form method="post">
                    <div class="form-group">
                        <label>Username</label>
                        <div class="input-wrapper">
                            <i class="bi bi-person"></i>
                            <input type="text" name="username" placeholder="Masukkan username" required>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Email</label>
                        <div class="input-wrapper">
                            <i class="bi bi-envelope"></i>
                            <input type="email" name="email" placeholder="Masukkan email" required>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Password</label>
                        <div class="input-wrapper">
                            <i class="bi bi-lock"></i>
                            <input type="password" name="password" placeholder="Minimal 6 karakter" required>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Konfirmasi Password</label>
                        <div class="input-wrapper">
                            <i class="bi bi-lock-fill"></i>
                            <input type="password" name="confirm_password" placeholder="Ulangi password" required>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Nomor WhatsApp</label>
                        <div class="input-wrapper">
                            <i class="bi bi-whatsapp"></i>
                            <input type="tel" name="phone_number" placeholder="Contoh: 08123456789">
                        </div>
                        <span class="hint">Opsional — untuk menerima reminder skripsi</span>
                    </div>

                    <button type="submit" class="btn-register">
                        <i class="bi bi-person-plus me-2"></i>Daftar Sekarang
                    </button>
                </form>

                <div class="login-link">
                    Sudah punya akun? <a href="{{ url_for('login') }}">Login</a>
                </div>
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

        flash(f"👋 Selamat datang, {user['username']}!", "success")
        return redirect(url_for("index"))

    return render_template_string("""
        <!doctype html>
        <html lang="id">
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>Login - Winnowing Progress</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
            <link href="https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,100..900;1,100..900&display=swap" rel="stylesheet">
            <style>
                * {
                    font-family: 'Inter', sans-serif;
                }
                body {
                    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
                    min-height: 100vh;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    padding: 20px;
                }
                .login-card {
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(20px);
                    -webkit-backdrop-filter: blur(20px);
                    border-radius: 32px;
                    padding: 50px 45px;
                    max-width: 420px;
                    width: 100%;
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    box-shadow: 0 40px 80px rgba(0, 0, 0, 0.6);
                    animation: fadeInUp 0.8s ease-out;
                }
                @keyframes fadeInUp {
                    from {
                        opacity: 0;
                        transform: translateY(40px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }
                .login-card .logo {
                    text-align: center;
                    margin-bottom: 32px;
                }
                .login-card .logo .icon {
                    font-size: 48px;
                    background: linear-gradient(135deg, #f093fb, #f5576c);
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                    display: inline-block;
                }
                .login-card .logo h1 {
                    color: #fff;
                    font-size: 28px;
                    font-weight: 700;
                    margin-top: 8px;
                    letter-spacing: -0.5px;
                }
                .login-card .logo p {
                    color: rgba(255, 255, 255, 0.5);
                    font-size: 14px;
                    font-weight: 400;
                    margin-top: 4px;
                }
                .login-card .form-group {
                    margin-bottom: 20px;
                }
                .login-card .form-group label {
                    color: rgba(255, 255, 255, 0.7);
                    font-size: 13px;
                    font-weight: 500;
                    margin-bottom: 6px;
                    display: block;
                }
                .login-card .form-group .input-wrapper {
                    position: relative;
                }
                .login-card .form-group .input-wrapper .bi {
                    position: absolute;
                    left: 16px;
                    top: 50%;
                    transform: translateY(-50%);
                    color: rgba(255, 255, 255, 0.3);
                    font-size: 18px;
                    transition: color 0.3s ease;
                }
                .login-card .form-group .input-wrapper input {
                    width: 100%;
                    padding: 14px 16px 14px 48px;
                    background: rgba(255, 255, 255, 0.06);
                    border: 2px solid rgba(255, 255, 255, 0.08);
                    border-radius: 14px;
                    color: #fff;
                    font-size: 15px;
                    font-weight: 400;
                    transition: all 0.3s ease;
                    outline: none;
                }
                .login-card .form-group .input-wrapper input::placeholder {
                    color: rgba(255, 255, 255, 0.25);
                }
                .login-card .form-group .input-wrapper input:focus {
                    border-color: rgba(245, 87, 108, 0.5);
                    background: rgba(255, 255, 255, 0.08);
                    box-shadow: 0 0 0 4px rgba(245, 87, 108, 0.1);
                }
                .login-card .form-group .input-wrapper input:focus + .bi,
                .login-card .form-group .input-wrapper input:focus ~ .bi {
                    color: #f5576c;
                }
                .login-card .btn-login {
                    width: 100%;
                    padding: 14px;
                    background: linear-gradient(135deg, #f093fb, #f5576c);
                    border: none;
                    border-radius: 14px;
                    color: #fff;
                    font-size: 16px;
                    font-weight: 600;
                    transition: all 0.3s ease;
                    cursor: pointer;
                    margin-top: 8px;
                }
                .login-card .btn-login:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 12px 30px rgba(245, 87, 108, 0.35);
                }
                .login-card .btn-login:active {
                    transform: translateY(0);
                }
                .login-card .register-link {
                    text-align: center;
                    margin-top: 20px;
                    color: rgba(255, 255, 255, 0.4);
                    font-size: 14px;
                }
                .login-card .register-link a {
                    color: #f5576c;
                    text-decoration: none;
                    font-weight: 600;
                    transition: color 0.3s ease;
                }
                .login-card .register-link a:hover {
                    color: #f093fb;
                }
                .alert-custom {
                    background: rgba(255, 255, 255, 0.05);
                    backdrop-filter: blur(10px);
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    border-radius: 14px;
                    padding: 12px 16px;
                    color: #fff;
                    font-size: 14px;
                    margin-bottom: 20px;
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }
                .alert-custom .bi {
                    font-size: 18px;
                }
                .alert-custom.danger {
                    border-color: rgba(245, 87, 108, 0.3);
                    background: rgba(245, 87, 108, 0.1);
                }
                .alert-custom.danger .bi {
                    color: #f5576c;
                }
                .alert-custom.success {
                    border-color: rgba(52, 211, 153, 0.3);
                    background: rgba(52, 211, 153, 0.1);
                }
                .alert-custom.success .bi {
                    color: #34d399;
                }
                @media (max-width: 480px) {
                    .login-card {
                        padding: 32px 24px;
                    }
                }
            </style>
        </head>
        <body>
            <div class="login-card">
                <div class="logo">
                    <span class="icon">🚀</span>
                    <h1>Selamat Datang</h1>
                    <p>Login untuk pantau progres skripsimu</p>
                </div>

                {% with messages = get_flashed_messages(with_categories=true) %}
                    {% if messages %}
                        {% for category, msg in messages %}
                            <div class="alert-custom {{ category }}">
                                <i class="bi bi-{{ 'check-circle' if category == 'success' else 'exclamation-triangle' }}"></i>
                                {{ msg }}
                            </div>
                        {% endfor %}
                    {% endif %}
                {% endwith %}

                <form method="post">
                    <div class="form-group">
                        <label>Username</label>
                        <div class="input-wrapper">
                            <i class="bi bi-person"></i>
                            <input type="text" name="username" placeholder="Masukkan username" required>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Password</label>
                        <div class="input-wrapper">
                            <i class="bi bi-lock"></i>
                            <input type="password" name="password" placeholder="Masukkan password" required>
                        </div>
                    </div>

                    <button type="submit" class="btn-login">
                        <i class="bi bi-box-arrow-in-right me-2"></i>Login
                    </button>
                </form>

                <div class="register-link">
                    Belum punya akun? <a href="{{ url_for('register') }}">Daftar Sekarang</a>
                </div>
            </div>
        </body>
        </html>
    """, title="Login")

@app.route("/logout")
def logout():
    session.clear()
    flash("Anda telah logout.", "info")
    return redirect(url_for("login"))

# ==================== ROUTE: MAIN (DENGAN AUTH) ====================
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

        try:
            file.save(save_path)
            raw_text = extract_text(save_path, ext)
            processed = process_document_text(raw_text)
            new_doc_id = save_document_record(user_id, original_name, saved_name, ext, processed)
            compare_new_document_against_all(user_id, new_doc_id)
            
            # Cek apakah upload ini valid (ada perubahan)
            best = get_today_best_highlight(user_id)
            if best and float(best["similarity_percent"]) < 100.0:
                update_streak_after_valid_upload(user_id)
            
            flash(f"✅ File '{original_name}' berhasil diupload dan progress diperbarui.", "success")
        except Exception as e:
            if save_path.exists():
                save_path.unlink()
            flash(f"Gagal memproses file: {e}", "danger")

        return redirect(url_for("index"))

    # GET request - tampilkan dashboard
    streak = get_or_reset_streak_state(user_id)
    flame = flame_meta(streak["flame_status"])
    docs = get_documents_by_user(user_id)
    docs_count = len(docs)
    comparison_count = len(get_comparisons_by_user(user_id))
    sequential_progress = get_sequential_progress_by_user(user_id)
    best = get_today_best_highlight(user_id)
    upload_days = get_upload_days_by_user(user_id)
    recent_highlight = get_recent_progress_highlight_by_user(user_id)

    content = render_template_string(
        """
        <div class="row g-4 mb-4">
            <div class="col-lg-7">
                <div class="card card-shadow">
                    <div class="card-body p-4">
                        <h3 class="fw-bold mb-3"><i class="bi bi-cloud-upload me-2"></i>Upload Versi Dokumen</h3>
                        <p class="text-muted">
                            Upload minimal 1 file per hari agar streak tetap jalan.
                            Streak hanya bertambah jika hari itu ada hasil compare terbaik yang bukan 100% similarity.
                        </p>

                        <form method="post" enctype="multipart/form-data">
                            <div class="mb-3">
                                <label class="form-label">Pilih file</label>
                                <input type="file" class="form-control" name="file" required>
                            </div>
                            <button class="btn btn-dark" type="submit">
                                <i class="bi bi-cloud-upload me-2"></i>Upload & Hitung Progress
                            </button>
                        </form>
                    </div>
                </div>
            </div>

            <div class="col-lg-5">
                <div class="card card-shadow">
                    <div class="card-body p-4 text-center">
                        <div class="streak-icon mb-2">{{ flame.icon }}</div>
                        <div class="big-number">{{ streak["streak_count"] }}</div>
                        <div class="fw-semibold">Streak Hari Aktif</div>
                        <div class="mt-2">
                            <span class="badge {{ flame.badge_class }}">{{ flame.text }}</span>
                        </div>
                        <p class="text-muted mt-3 mb-0">
                            Jika 3 hari tidak upload, maka hari ke-4 api padam dan streak kembali 0.
                        </p>
                    </div>
                </div>
            </div>
        </div>

        {% if recent_highlight %}
        <div class="card card-shadow mb-4">
            <div class="card-body p-4">
                <h4 class="fw-bold mb-3"><i class="bi bi-graph-up me-2"></i>Highlight Progress 1–3 Hari Terakhir</h4>
                <div class="alert alert-{{ recent_highlight.alert }} mb-0">
                    <h5 class="mb-2">{{ recent_highlight.label }}</h5>
                    <p class="mb-1">{{ recent_highlight.message }}</p>
                    <small>
                        Rata-rata difference: <b>{{ recent_highlight.avg_difference }}%</b>
                        dari {{ recent_highlight.days_count }} hari upload terakhir.
                    </small>
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
                            <div class="soft-card text-dark h-100">
                                <div class="small text-muted mb-1">Compare utama hari ini</div>
                                <div class="fw-semibold">{{ best["doc1_name"] }}</div>
                                <div class="text-muted">vs</div>
                                <div class="fw-semibold">{{ best["doc2_name"] }}</div>
                            </div>
                        </div>

                        <div class="col-md-2">
                            <div class="soft-card text-dark text-center h-100">
                                <div class="small text-muted">Similarity</div>
                                <div class="fw-bold fs-4">{{ best["similarity_percent"] }}%</div>
                            </div>
                        </div>

                        <div class="col-md-2">
                            <div class="soft-card text-dark text-center h-100">
                                <div class="small text-muted">Difference</div>
                                <div class="fw-bold fs-4">{{ best["difference_percent"] }}%</div>
                            </div>
                        </div>

                        <div class="col-md-2">
                            <div class="soft-card text-dark text-center h-100">
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
                            <li>Status API: <span class="badge {{ 'bg-success' if flame.flame_status == 'on' else 'bg-secondary' }}">
                                {{ "🔥 Menyala" if flame.flame_status == 'on' else "🪨 Padam" }}
                            </span></li>
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
                                <thead>
                                    <tr>
                                        <th>Tanggal</th>
                                        <th>Total File</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {% if upload_days %}
                                        {% for day in upload_days %}
                                        <tr>
                                            <td>{{ day["upload_date"] }}</td>
                                            <td>{{ day["total"] }}</td>
                                        </tr>
                                        {% endfor %}
                                    {% else %}
                                        <tr><td colspan="2" class="text-center">Belum ada upload.</td></tr>
                                    {% endif %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="card card-shadow">
            <div class="card-body p-4">
                <h4 class="fw-bold mb-3"><i class="bi bi-arrow-right-circle me-2"></i>Progress Berurutan</h4>
                <div class="table-responsive">
                    <table class="table table-striped">
                        <thead>
                            <tr>
                                <th>Dari</th>
                                <th>Ke</th>
                                <th>Similarity</th>
                                <th>Difference</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% if sequential_progress %}
                                {% for item in sequential_progress %}
                                <tr>
                                    <td>{{ item.from_name }}</td>
                                    <td>{{ item.to_name }}</td>
                                    <td>{{ item.similarity_percent }}%</td>
                                    <td><span class="badge bg-dark">{{ item.difference_percent }}%</span></td>
                                    <td>{{ item.progress_label }}</td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr>
                                    <td colspan="5" class="text-center">Belum cukup dokumen untuk menghitung progres berurutan.</td>
                                </tr>
                            {% endif %}
                        </tbody>
                    </table>
                </div>
                <p class="text-muted mb-0 mt-2">Detail lengkap semua pasangan file ada di halaman Progress.</p>
            </div>
        </div>
        """,
        k_gram=K_GRAM,
        window_size=WINDOW_SIZE,
        docs_count=docs_count,
        comparison_count=comparison_count,
        sequential_progress=sequential_progress,
        best=best,
        streak=streak,
        flame=flame,
        upload_days=upload_days,
        recent_highlight=recent_highlight,
    )

    return render_template_string(BASE_TEMPLATE, title="Home", content=content)

@app.route("/documents")
def documents():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))

    docs = get_documents_by_user(session["user_id"])

    content = render_template_string(
        """
        <div class="card card-shadow">
            <div class="card-body p-4">
                <h3 class="fw-bold mb-3"><i class="bi bi-file-earmark-text me-2"></i>Daftar Dokumen</h3>

                <div class="table-responsive">
                    <table class="table table-striped">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Nama File</th>
                                <th>Ext</th>
                                <th>Uploaded At</th>
                                <th>Aksi</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% if docs %}
                                {% for doc in docs %}
                                <tr>
                                    <td>{{ doc["id"] }}</td>
                                    <td>{{ doc["original_name"] }}</td>
                                    <td>{{ doc["file_ext"] }}</td>
                                    <td>{{ doc["uploaded_at"] }}</td>
                                    <td>
                                        <a class="btn btn-sm btn-dark" href="{{ url_for('document_detail', doc_id=doc['id']) }}">
                                            <i class="bi bi-eye me-1"></i>Detail
                                        </a>
                                    </td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr>
                                    <td colspan="5" class="text-center">Belum ada dokumen.</td>
                                </tr>
                            {% endif %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        """,
        docs=docs,
    )

    return render_template_string(BASE_TEMPLATE, title="Dokumen", content=content)

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

    content = render_template_string(
        """
        <div class="card card-shadow mb-4">
            <div class="card-body p-4">
                <h3 class="fw-bold mb-1"><i class="bi bi-file-earmark me-2"></i>Detail Dokumen</h3>
                <p class="text-muted mb-1">{{ doc["original_name"] }}</p>
                <p class="text-muted mb-0">Diupload: {{ doc["uploaded_at"] }}</p>
            </div>
        </div>

        <div class="row g-4">
            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>Raw Text</h5>
                        <div class="mono-box">{{ doc["raw_text"] or "" }}</div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>Preprocessing</h5>
                        <div class="mono-box">{{ doc["clean_text"] or "" }}</div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>K-Gram (sample 200)</h5>
                        <div class="small-scroll">{{ kgrams[:200] }}</div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>Hashes (sample 200)</h5>
                        <div class="small-scroll">{{ hashes[:200] }}</div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>Windows (sample 100)</h5>
                        <div class="small-scroll">{{ windows[:100] }}</div>
                    </div>
                </div>
            </div>

            <div class="col-lg-6">
                <div class="card card-shadow">
                    <div class="card-body">
                        <h5>Fingerprints</h5>
                        <div class="small-scroll">{{ fingerprints }}</div>
                    </div>
                </div>
            </div>
        </div>
        """,
        doc=doc,
        kgrams=kgrams,
        hashes=hashes,
        windows=windows,
        fingerprints=fingerprints,
    )

    return render_template_string(BASE_TEMPLATE, title="Detail Dokumen", content=content)

@app.route("/progress")
@app.route("/comparisons")
def comparisons():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))

    comparisons_data = get_comparisons_by_user(session["user_id"])

    content = render_template_string(
        """
        <div class="card card-shadow">
            <div class="card-body p-4">
                <h3 class="fw-bold mb-3"><i class="bi bi-bar-chart me-2"></i>Semua Hasil Progress Comparison</h3>

                <p class="text-muted">
                    Detail lengkap semua pasangan file ditampilkan di sini.
                </p>

                <div class="table-responsive">
                    <table class="table table-striped align-middle">
                        <thead>
                            <tr>
                                <th>ID Doc 1</th>
                                <th>Nama File 1</th>
                                <th>ID Doc 2</th>
                                <th>Nama File 2</th>
                                <th>Similarity</th>
                                <th>Difference</th>
                                <th>Status Progress</th>
                                <th>Waktu</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% if comparisons_data %}
                                {% for row in comparisons_data %}
                                <tr>
                                    <td>{{ row["doc1_id"] }}</td>
                                    <td>{{ row["doc1_name"] }}</td>
                                    <td>{{ row["doc2_id"] }}</td>
                                    <td>{{ row["doc2_name"] }}</td>
                                    <td>{{ row["similarity_percent"] }}%</td>
                                    <td><span class="badge bg-dark">{{ row["difference_percent"] }}%</span></td>
                                    <td>{{ row["progress_label"] }}</td>
                                    <td>{{ row["process_time"] }} detik</td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr>
                                    <td colspan="8" class="text-center">Belum ada hasil perbandingan.</td>
                                </tr>
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

# ==================== ROUTE: SETTINGS (UPDATE PHONE NUMBER) ====================
@app.route("/settings", methods=["GET", "POST"])
def settings():
    if "user_id" not in session:
        flash("Silakan login terlebih dahulu.", "warning")
        return redirect(url_for("login"))

    user = get_user_by_id(session["user_id"])

    if request.method == "POST":
        phone_number = request.form.get("phone_number", "").strip()
        
        # Format nomor telepon
        if phone_number:
            phone_number = re.sub(r"[^0-9]", "", phone_number)
            if phone_number.startswith("0"):
                phone_number = "62" + phone_number[1:]
            elif not phone_number.startswith("62"):
                phone_number = "62" + phone_number

        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "UPDATE users SET phone_number = ? WHERE id = ?",
            (phone_number or None, session["user_id"])
        )
        conn.commit()
        conn.close()

        session["phone_number"] = phone_number
        flash("Nomor telepon berhasil diperbarui!", "success")
        return redirect(url_for("settings"))

    content = render_template_string("""
        <div class="row justify-content-center">
            <div class="col-md-6">
                <div class="card auth-card">
                    <div class="card-header text-center">
                        <h4 class="mb-0"><i class="bi bi-gear me-2"></i>Pengaturan</h4>
                    </div>
                    <div class="card-body">
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
                                <input type="tel" class="form-control" name="phone_number" 
                                       value="{{ user['phone_number'] or '' }}" 
                                       placeholder="08123456789">
                                <small class="text-muted d-block mt-1">
                                    Masukkan nomor WhatsApp aktif untuk menerima notifikasi reminder.
                                </small>
                            </div>
                            <button type="submit" class="btn btn-primary w-100">
                                <i class="bi bi-save me-2"></i>Simpan Perubahan
                            </button>
                            <a href="{{ url_for('index') }}" class="btn btn-outline-secondary w-100 mt-2">
                                <i class="bi bi-arrow-left me-2"></i>Kembali
                            </a>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    """, title="Pengaturan", user=user)

    return render_template_string(BASE_TEMPLATE, title="Pengaturan", content=content)

# ==================== FUNGSI REMINDER SCHEDULER ====================
def run_reminder_scheduler():
    """
    Fungsi ini dipanggil setiap hari untuk mengirim reminder
    """
    print(f"[{now_dt()}] Memulai pengecekan reminder...")
    check_and_send_reminders()
    print(f"[{now_dt()}] Pengecekan reminder selesai.")

# ==================== RUN APP ====================
if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000) 