# seed_dummy.py - Buat dummy data untuk testing (REALTIME)
import sqlite3
import hashlib
from datetime import datetime, timedelta
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "similarity.db"

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def create_dummy_user(days_ago=3):
    """
    Buat user dummy dengan nomor WA
    days_ago: berapa hari yang lalu last upload (default 3 hari)
    """
    conn = get_conn()
    cur = conn.cursor()
    
    # Hapus data dummy sebelumnya
    cur.execute("DELETE FROM users WHERE username = 'testuser'")
    cur.execute("DELETE FROM streak_state WHERE user_id IN (SELECT id FROM users WHERE username = 'testuser')")
    
    # Buat user dummy
    password_hash = hashlib.sha256("password123".encode()).hexdigest()
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # GANTI dengan nomor WA asli kamu!
    phone_number = "081319894484"  # <-- GANTI DENGAN NOMOR WA ASLI KAMU!
    
    cur.execute("""
        INSERT INTO users (username, email, password_hash, phone_number, created_at)
        VALUES (?, ?, ?, ?, ?)
    """, ("testuser", "test@email.com", password_hash, phone_number, created_at))
    
    user_id = cur.lastrowid
    
    # Buat streak_state dengan flame_status = 'on' dan last_valid_upload_date = days_ago hari lalu
    days_ago_date = (datetime.now() - timedelta(days=days_ago)).strftime("%Y-%m-%d")
    
    cur.execute("""
        INSERT INTO streak_state (user_id, streak_count, last_valid_upload_date, flame_status)
        VALUES (?, 1, ?, 'on')
    """, (user_id, days_ago_date))
    
    conn.commit()
    conn.close()
    
    print(f"✅ User dummy berhasil dibuat:")
    print(f"   Username: testuser")
    print(f"   Password: password123")
    print(f"   No WA: {phone_number}")
    print(f"   Terakhir upload: {days_ago_date} ({days_ago} hari lalu)")
    print()
    return user_id

def create_dummy_document(user_id, days_ago=3):
    """Buat dokumen dummy"""
    conn = get_conn()
    cur = conn.cursor()
    
    uploaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    upload_date = (datetime.now() - timedelta(days=days_ago)).strftime("%Y-%m-%d")
    
    cur.execute("""
        INSERT INTO documents (
            user_id, original_name, saved_name, file_ext, 
            uploaded_at, upload_date, raw_text, clean_text,
            kgrams_json, hashes_json, windows_json, fingerprints_json
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        user_id,
        "test_document.txt",
        "test_document.txt",
        "txt",
        uploaded_at,
        upload_date,
        "Ini adalah teks dummy untuk testing reminder WhatsApp",
        "inidummy",
        '["dummy"]',
        '["dummy"]',
        '["dummy"]',
        '["dummy"]'
    ))
    
    conn.commit()
    conn.close()
    print(f"✅ Dokumen dummy berhasil dibuat untuk user testuser")
    print(f"   Upload date: {upload_date} ({days_ago} hari lalu)")
    print()

def create_dummy_comparison(user_id):
    """Buat comparison dummy"""
    conn = get_conn()
    cur = conn.cursor()
    
    # Ambil document id
    doc = cur.execute("SELECT id FROM documents WHERE user_id = ? ORDER BY id DESC LIMIT 1", (user_id,)).fetchone()
    if not doc:
        print("❌ Tidak ada dokumen untuk dibuat comparison")
        return
    
    doc_id = doc["id"]
    
    cur.execute("""
        INSERT INTO comparisons (
            user_id, doc1_id, doc2_id, similarity_percent, difference_percent,
            process_time, progress_label, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        user_id,
        doc_id,
        doc_id,
        85.5,
        14.5,
        0.5,
        "Progres ringan",
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))
    
    conn.commit()
    conn.close()
    print(f"✅ Comparison dummy berhasil dibuat")
    print()

def reset_streak_to_on(user_id, days_ago=3):
    """Reset streak state ke ON untuk testing"""
    conn = get_conn()
    cur = conn.cursor()
    
    days_ago_date = (datetime.now() - timedelta(days=days_ago)).strftime("%Y-%m-%d")
    
    cur.execute("""
        UPDATE streak_state 
        SET flame_status = 'on', 
            streak_count = 1, 
            last_valid_upload_date = ?
        WHERE user_id = ?
    """, (days_ago_date, user_id))
    
    conn.commit()
    conn.close()
    print(f"✅ Streak state di-reset ke ON ({days_ago} hari lalu)")
    print()

def show_data():
    """Tampilkan data yang ada di database"""
    conn = get_conn()
    cur = conn.cursor()
    
    print("=" * 50)
    print("📊 DATA DI DATABASE")
    print("=" * 50)
    
    # Cek users
    users = cur.execute("SELECT id, username, phone_number FROM users").fetchall()
    print("\n👤 USERS:")
    if users:
        for u in users:
            print(f"   ID: {u['id']}, Username: {u['username']}, No WA: {u['phone_number']}")
    else:
        print("   ❌ Tidak ada user")
    
    # Cek streak_state
    streaks = cur.execute("SELECT user_id, streak_count, last_valid_upload_date, flame_status FROM streak_state").fetchall()
    print("\n🔥 STREAK STATE:")
    if streaks:
        for s in streaks:
            print(f"   User ID: {s['user_id']}, Streak: {s['streak_count']}, Last Upload: {s['last_valid_upload_date']}, Status: {s['flame_status']}")
            
            # Hitung selisih hari
            if s['last_valid_upload_date']:
                last_date = datetime.strptime(s['last_valid_upload_date'], "%Y-%m-%d").date()
                today = datetime.now().date()
                diff = (today - last_date).days
                print(f"      ↳ Selisih: {diff} hari (harus >= 3 untuk kirim reminder)")
    else:
        print("   ❌ Tidak ada streak_state")
    
    conn.close()
    print("=" * 50)

def reset_all(days_ago=3):
    """
    Reset semua data dan buat ulang dummy
    days_ago: berapa hari yang lalu last upload (default 3)
    """
    print("🔄 Menghapus data dummy lama...")
    
    conn = get_conn()
    cur = conn.cursor()
    
    # Hapus semua data dummy
    cur.execute("DELETE FROM comparisons WHERE user_id IN (SELECT id FROM users WHERE username = 'testuser')")
    cur.execute("DELETE FROM documents WHERE user_id IN (SELECT id FROM users WHERE username = 'testuser')")
    cur.execute("DELETE FROM streak_state WHERE user_id IN (SELECT id FROM users WHERE username = 'testuser')")
    cur.execute("DELETE FROM users WHERE username = 'testuser'")
    
    conn.commit()
    conn.close()
    
    print("✅ Data dummy berhasil dihapus")
    print()
    
    # Buat data baru
    user_id = create_dummy_user(days_ago)
    create_dummy_document(user_id, days_ago)
    create_dummy_comparison(user_id)
    reset_streak_to_on(user_id, days_ago)
    show_data()
    
    print("\n" + "=" * 50)
    print("📱 TESTING REMINDER")
    print("=" * 50)
    print("\n🚀 Jalankan perintah berikut untuk test kirim WA:")
    print("   python -c \"from app import check_and_send_reminders; check_and_send_reminders()\"")
    print("\n📝 Atau jalankan:")
    print("   python reminder.py")
    print("=" * 50)

if __name__ == "__main__":
    import sys
    
    # Bisa pilih berapa hari yang lalu
    # Contoh: python seed_dummy.py 3 (default)
    #         python seed_dummy.py 4
    #         python seed_dummy.py 5
    
    days = 3
    if len(sys.argv) > 1:
        try:
            days = int(sys.argv[1])
        except:
            pass
    
    print(f"\n📅 Setting last upload = {days} hari yang lalu")
    print(f"   Hari ini: {datetime.now().strftime('%Y-%m-%d')}")
    print(f"   Tanggal upload: {(datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')}")
    print()
    
    reset_all(days)