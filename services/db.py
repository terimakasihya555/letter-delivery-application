import sqlite3
import sys
from pathlib import Path

from services.backup import backup_database


def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = get_base_dir()
DB_PATH = BASE_DIR / "database.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS surat_rekap (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bulan TEXT,
            tahun TEXT,
            nomor_urut INTEGER,
            nomor_referensi TEXT,
            kode_surat TEXT,
            klasifikasi TEXT,
            alamat TEXT
        )
    """)

    conn.commit()
    conn.close()


def get_all_surat(q="", bulan="", tahun=""):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT *
        FROM surat_rekap
        WHERE 1=1
    """
    params = []

    if q:
        query += """
            AND (
                nomor_referensi LIKE ?
                OR kode_surat LIKE ?
                OR klasifikasi LIKE ?
                OR alamat LIKE ?
            )
        """
        keyword = f"%{q}%"
        params.extend([keyword, keyword, keyword, keyword])

    if bulan:
        query += " AND bulan = ?"
        params.append(bulan)

    if tahun:
        query += " AND tahun = ?"
        params.append(tahun)

    query += " ORDER BY tahun DESC, bulan DESC, nomor_urut ASC, id ASC"

    cursor.execute(query, params)
    rows = cursor.fetchall()

    conn.close()
    return rows


def get_surat_by_id(surat_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT * FROM surat_rekap WHERE id = ?",
        (surat_id,)
    )

    row = cursor.fetchone()

    conn.close()
    return row


def insert_surat(data):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO surat_rekap (
            bulan,
            tahun,
            nomor_urut,
            nomor_referensi,
            kode_surat,
            klasifikasi,
            alamat
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        data.get("bulan", ""),
        data.get("tahun", ""),
        data.get("nomor_urut", ""),
        data.get("nomor_referensi", ""),
        data.get("kode_surat", ""),
        data.get("klasifikasi", ""),
        data.get("alamat", "")
    ))

    conn.commit()
    conn.close()

    backup_database()


def insert_many_surat(rows):
    """
    Insert banyak data sekaligus dari hasil import Excel.

    rows harus berupa list of dict dengan key:
    bulan, tahun, nomor_urut, nomor_referensi, kode_surat, klasifikasi, alamat
    """
    if not rows:
        return 0

    conn = get_connection()
    cursor = conn.cursor()

    cursor.executemany("""
        INSERT INTO surat_rekap (
            bulan,
            tahun,
            nomor_urut,
            nomor_referensi,
            kode_surat,
            klasifikasi,
            alamat
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, [
        (
            row.get("bulan", ""),
            row.get("tahun", ""),
            row.get("nomor_urut", ""),
            row.get("nomor_referensi", ""),
            row.get("kode_surat", ""),
            row.get("klasifikasi", ""),
            row.get("alamat", "")
        )
        for row in rows
    ])

    conn.commit()
    inserted_count = cursor.rowcount
    conn.close()

    backup_database()

    return inserted_count


def update_surat(surat_id, data):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE surat_rekap
        SET
            bulan = ?,
            tahun = ?,
            nomor_urut = ?,
            nomor_referensi = ?,
            kode_surat = ?,
            klasifikasi = ?,
            alamat = ?
        WHERE id = ?
    """, (
        data.get("bulan", ""),
        data.get("tahun", ""),
        data.get("nomor_urut", ""),
        data.get("nomor_referensi", ""),
        data.get("kode_surat", ""),
        data.get("klasifikasi", ""),
        data.get("alamat", ""),
        surat_id
    ))

    conn.commit()
    conn.close()

    backup_database()


def delete_surat(surat_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM surat_rekap WHERE id = ?",
        (surat_id,)
    )

    conn.commit()
    conn.close()

    backup_database()