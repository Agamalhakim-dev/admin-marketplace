# CARA MENJALANKAN:
#   1. pip install Flask Flask-MySQLdb
#   2. Import database.sql ke MySQL terlebih dahulu
#   3. Sesuaikan MYSQL_USER dan MYSQL_PASSWORD di bagian CONFIG
#   4. Jalankan: python app_sistem_penjualan.py
#   5. Buka browser: http://localhost:5000
#
# =============================================================================
# PANDUAN KOMENTAR BAHASA DALAM FILE INI:
#   [PYTHON]      → Kode backend Python / Flask
#   [HTML]        → Struktur markup halaman
#   [CSS]         → Gaya tampilan / styling
#   [JAVASCRIPT]  → Logika interaksi di browser
#   [SQL]         → Query / schema database MySQL
# =============================================================================


# =============================================================================
# [PYTHON] ── BAGIAN 1: IMPORT LIBRARY
# =============================================================================

from flask import Flask, render_template_string, request, redirect, url_for, flash, jsonify, send_file, session
# Flask               → framework web Python utama
# render_template_string → render HTML langsung dari string Python (tanpa file .html)
# request             → mengakses data dari form / URL
# redirect            → mengarahkan user ke halaman lain
# url_for             → membuat URL dari nama fungsi route
# flash               → mengirim pesan notifikasi sekali tampil
# jsonify             → mengubah dict Python menjadi respons JSON

from flask_mysqldb import MySQL
from werkzeug.security import generate_password_hash, check_password_hash
# generate_password_hash -> enkripsi password aman (bcrypt/pbkdf2)
# check_password_hash    -> verifikasi password saat login
from functools import wraps
# wraps -> untuk membuat decorator login_required
# MySQL               → ekstensi Flask untuk koneksi ke database MySQL

from datetime import date, datetime
# date                → tipe data tanggal (tahun-bulan-hari)
# datetime            → tipe data tanggal + waktu

import json
# json                → untuk encode/decode data JSON (dipakai di grafik Chart.js)

import io
# io                  -> membuat file di memori tanpa menyimpan ke disk (untuk download)

import smtplib
# smtplib             -> library bawaan Python untuk mengirim email via SMTP

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# email.mime          -> membuat struktur email (subject, body, attachment)

# [PYTHON] Library PDF -- install: pip install reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
# reportlab           -> library untuk membuat file PDF dari Python

# [PYTHON] Library Excel -- install: pip install openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
# openpyxl            -> library untuk membuat dan memformat file Excel (.xlsx)


# =============================================================================
# [PYTHON] ── BAGIAN 2: INISIALISASI APLIKASI FLASK
# =============================================================================

app = Flask(__name__)
# Flask(__name__)     → membuat instance aplikasi Flask
# __name__            → nama modul saat ini (dipakai Flask untuk menentukan root path)


# =============================================================================
# [PYTHON] ── BAGIAN 3: KONFIGURASI DATABASE & APLIKASI
# =============================================================================

# -- Kunci rahasia untuk session dan flash message --
app.config['SECRET_KEY'] = 'apsi_teknik_industri_2024'

# -- Konfigurasi koneksi MySQL --
import os

app.config['MYSQL_HOST']     = os.environ.get('MYSQL_HOST')
app.config['MYSQL_USER']     = os.environ.get('MYSQL_USER')
app.config['MYSQL_PASSWORD'] = os.environ.get('MYSQL_PASSWORD')
app.config['MYSQL_DB']       = os.environ.get('MYSQL_DB')

mysql = MySQL(app)
# MySQL(app)          → menginisialisasi koneksi MySQL ke aplikasi Flask


# =============================================================================
# [SQL] ── BAGIAN 4: SCHEMA DATABASE (jalankan ini di MySQL terlebih dahulu)
# =============================================================================
# 
# Salin kode SQL di bawah ini dan jalankan di MySQL Workbench / phpMyAdmin:
#
# CREATE DATABASE IF NOT EXISTS sistem_penjualan;
# USE sistem_penjualan;
#
# -- Tabel pelanggan: menyimpan data pelanggan / klien --
# CREATE TABLE pelanggan (
#     id              INT AUTO_INCREMENT PRIMARY KEY,
#     kode_pelanggan  VARCHAR(20) UNIQUE NOT NULL,     -- kode unik: PLG-001
#     nama            VARCHAR(100) NOT NULL,            -- nama lengkap / perusahaan
#     email           VARCHAR(100),                     -- email kontak
#     telepon         VARCHAR(20),                      -- nomor telepon
#     alamat          TEXT,                             -- alamat lengkap
#     kota            VARCHAR(50),                      -- kota domisili
#     status          ENUM('aktif','tidak_aktif') DEFAULT 'aktif',
#     created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
#     updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
# );
#
# -- Tabel kategori: pengelompokan produk --
# CREATE TABLE kategori (
#     id          INT AUTO_INCREMENT PRIMARY KEY,
#     nama        VARCHAR(50) NOT NULL,
#     deskripsi   TEXT
# );
#
# -- Tabel produk: data produk beserta stok --
# CREATE TABLE produk (
#     id              INT AUTO_INCREMENT PRIMARY KEY,
#     kode_produk     VARCHAR(20) UNIQUE NOT NULL,     -- kode unik: PRD-001
#     nama            VARCHAR(100) NOT NULL,
#     kategori_id     INT,
#     harga           DECIMAL(15,2) NOT NULL,           -- harga jual satuan
#     stok            INT DEFAULT 0,                    -- jumlah stok tersedia
#     stok_minimum    INT DEFAULT 5,                    -- batas peringatan stok menipis
#     satuan          VARCHAR(20) DEFAULT 'pcs',        -- satuan: pcs, unit, kg, dll
#     deskripsi       TEXT,
#     status          ENUM('aktif','tidak_aktif') DEFAULT 'aktif',
#     created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
#     updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
#     FOREIGN KEY (kategori_id) REFERENCES kategori(id)
# );
#
# -- Tabel pesanan: header transaksi penjualan --
# CREATE TABLE pesanan (
#     id              INT AUTO_INCREMENT PRIMARY KEY,
#     no_pesanan      VARCHAR(20) UNIQUE NOT NULL,     -- nomor unik: ORD-001
#     pelanggan_id    INT NOT NULL,
#     tanggal_pesan   DATE NOT NULL,
#     tanggal_kirim   DATE,
#     status          ENUM('pending','diproses','dikirim','selesai','dibatalkan') DEFAULT 'pending',
#     total_harga     DECIMAL(15,2) DEFAULT 0,          -- total sebelum diskon
#     diskon          DECIMAL(5,2) DEFAULT 0,           -- persentase diskon (0-100)
#     total_bayar     DECIMAL(15,2) DEFAULT 0,          -- total setelah diskon
#     catatan         TEXT,
#     created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
#     updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
#     FOREIGN KEY (pelanggan_id) REFERENCES pelanggan(id)
# );
#
# -- Tabel detail_pesanan: baris item dalam satu pesanan --
# CREATE TABLE detail_pesanan (
#     id              INT AUTO_INCREMENT PRIMARY KEY,
#     pesanan_id      INT NOT NULL,
#     produk_id       INT NOT NULL,
#     jumlah          INT NOT NULL,
#     harga_satuan    DECIMAL(15,2) NOT NULL,
#     subtotal        DECIMAL(15,2) NOT NULL,           -- harga_satuan × jumlah
#     FOREIGN KEY (pesanan_id) REFERENCES pesanan(id) ON DELETE CASCADE,
#     FOREIGN KEY (produk_id) REFERENCES produk(id)
# );
#
# -- Data awal: kategori --
# INSERT INTO kategori (nama) VALUES ('Elektronik'),('Pakaian'),('Makanan'),('Alat Tulis');
#
# -- Data awal: pelanggan contoh --
# INSERT INTO pelanggan (kode_pelanggan,nama,email,telepon,kota,status) VALUES
# ('PLG-001','PT Maju Jaya','maju@email.com','021-111','Jakarta','aktif'),
# ('PLG-002','CV Berkah','berkah@email.com','022-222','Bandung','aktif');
#
# -- Data awal: produk contoh --
# INSERT INTO produk (kode_produk,nama,kategori_id,harga,stok,stok_minimum,satuan) VALUES
# ('PRD-001','Laptop Asus',1,8500000,15,5,'unit'),
# ('PRD-002','Mouse Wireless',1,250000,50,10,'pcs'),
# ('PRD-003','Kaos Polos',2,85000,100,20,'pcs'),
# ('PRD-004','Pulpen Pilot',4,8000,3,30,'pcs');
# =============================================================================


# =============================================================================
# [PYTHON] ── BAGIAN 5: FUNGSI PEMBANTU (HELPER FUNCTIONS)
# =============================================================================

def generate_kode(prefix, table, kolom):
    """
    [PYTHON] Menghasilkan kode unik otomatis, misal: PLG-001, PRD-002, ORD-003
    
    Parameter:
        prefix  → awalan kode (PLG, PRD, ORD)
        table   → nama tabel di database
        kolom   → nama kolom kode di tabel tersebut
    """
    cur = mysql.connection.cursor()                              # membuka cursor database
    cur.execute(f"SELECT {kolom} FROM {table} ORDER BY id DESC LIMIT 1")  # ambil kode terakhir
    last = cur.fetchone()                                        # fetchone() → ambil 1 baris hasil query
    cur.close()                                                  # menutup cursor setelah selesai
    
    if last:
        # Pisah kode dengan '-', ambil angka, tambah 1
        last_num = int(last[kolom].split('-')[1]) + 1
    else:
        last_num = 1                                             # mulai dari 1 jika tabel kosong
    
    return f"{prefix}-{str(last_num).zfill(3)}"                 # zfill(3) → padding nol: 1 → "001"


def format_rupiah(value):
    """
    [PYTHON] Filter Jinja2: mengubah angka menjadi format mata uang Rupiah
    Contoh: 8500000 → 'Rp 8.500.000'
    """
    return f"Rp {value:,.0f}".replace(',', '.')                 # :,.0f → format ribuan dengan koma


# Mendaftarkan fungsi format_rupiah sebagai filter di template Jinja2
app.jinja_env.filters['rupiah'] = format_rupiah

# Mendaftarkan fungsi enumerate bawaan Python agar bisa dipakai di template Jinja2
app.jinja_env.filters['enumerate'] = enumerate


@app.context_processor
def inject_now():
    """
    [PYTHON] Context processor: menyuntikkan variabel 'now' ke semua template
    Sehingga {{ now.strftime('%d %b %Y') }} bisa dipakai di semua halaman HTML
    """
    return {'now': datetime.now()}


def login_required(f):
    """
    [PYTHON] Sistem login dinonaktifkan — semua route bisa diakses langsung.
    Decorator ini dibiarkan ada agar kode tidak error, tapi tidak memblokir akses.
    """
    @wraps(f)
    def decorated(*args, **kwargs):
        return f(*args, **kwargs)   # langsung teruskan tanpa cek session
    return decorated


def get_current_user():
    """[PYTHON] Sistem login dinonaktifkan — selalu return None."""
    return None


# =============================================================================
# [PYTHON] -- HALAMAN LOGIN
# =============================================================================


# =============================================================================
# [PYTHON] ── BAGIAN 6: TEMPLATE HTML/CSS/JS (Di-embed sebagai string Python)
# =============================================================================
# Semua tampilan antarmuka ada di sini.
# Format: render_template_string(NAMA_TEMPLATE, variabel=nilai)
# =============================================================================


# -----------------------------------------------------------------------------
# [CSS] Template CSS global — dipakai di semua halaman
# -----------------------------------------------------------------------------
CSS_GLOBAL = """
/* ══════════════════════════════════════════════
   [CSS] VARIABEL WARNA & FONT (CSS Custom Properties)
   Mendefinisikan palet warna agar konsisten di seluruh halaman
   ══════════════════════════════════════════════ */
:root {
  --bg:           #f0fdf4;   /* warna latar belakang utama */
  --surface:      #ffffff;   /* warna kartu / panel */
  --surface-2:    #f7fef9;   /* warna latar belakang alternatif */
  --border:       #bbf7d0;   /* warna garis batas */
  --border-dark:  #86efac;   /* garis batas lebih gelap */
  --text:         #052e16;   /* warna teks utama */
  --text-2:       #166534;   /* warna teks sekunder */
  --text-3:       #4ade80;   /* warna teks tersier / label */
  --accent:       #16a34a;   /* warna aksen (tombol utama) */
  --green:        #15803d;   /* hijau untuk status sukses/aktif */
  --green-light:  #dcfce7;   /* hijau muda untuk background badge */
  --yellow:       #e76f00;   /* kuning/oranye untuk peringatan */
  --yellow-light: #fff3cd;   /* kuning muda untuk background peringatan */
  --red:          #c0392b;   /* merah untuk bahaya/hapus */
  --red-light:    #fde8e8;   /* merah muda untuk background error */
  --blue:         #1a4a8a;   /* biru untuk informasi */
  --blue-light:   #dbeafe;   /* biru muda untuk background info */
  --sidebar-w:    220px;     /* lebar sidebar navigasi */
  --radius:       6px;       /* sudut melengkung elemen */
}

/* [CSS] Reset: menghapus margin/padding default browser */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

/* [CSS] Gaya body utama */
body {
  font-family: 'IBM Plex Sans', sans-serif; /* font utama */
  background: var(--bg);                    /* latar dari variabel CSS */
  color: var(--text);
  display: flex;                            /* layout flex: sidebar + konten berdampingan */
  min-height: 100vh;                        /* tinggi minimum penuh layar */
  font-size: 14px;
  line-height: 1.5;
}

/* ══════════════════════════════════════════════
   [CSS] SIDEBAR NAVIGASI
   ══════════════════════════════════════════════ */
.sidebar {
  width: var(--sidebar-w);
  background: #14532d;           /* sidebar hijau tua gelap */
  color: #f0fdf4;
  display: flex;
  flex-direction: column;        /* isi sidebar tersusun vertikal */
  position: fixed;               /* sidebar tetap di tempat saat scroll */
  top: 0; left: 0;
  height: 100vh;
  z-index: 100;                  /* z-index tinggi agar di atas konten */
}

.sidebar-brand {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 20px 18px;
  border-bottom: 1px solid rgba(255,255,255,0.10);
}
.brand-icon  { font-size: 22px; color: #4ade80; }   /* ikon logo berwarna kuning */
.brand-title { font-size: 15px; font-weight: 700; color: #fff; }
.brand-sub   { font-size: 10px; color: rgba(255,255,255,0.4); letter-spacing: 0.08em; text-transform: uppercase; }

.sidebar-nav { padding: 16px 10px; flex: 1; overflow-y: auto; }

/* [CSS] Label seksi navigasi */
.nav-section-label {
  font-size: 9px; letter-spacing: 0.12em;
  color: rgba(255,255,255,0.3);
  padding: 4px 8px; margin-bottom: 4px;
}

/* [CSS] Item navigasi */
.nav-item {
  display: flex; align-items: center; gap: 10px;
  padding: 9px 10px; color: rgba(255,255,255,0.65);
  text-decoration: none; border-radius: 4px;
  font-size: 13px; font-weight: 500;
  transition: all 0.15s;           /* animasi hover halus */
  margin-bottom: 2px;
}
.nav-item:hover  { background: rgba(255,255,255,0.08); color: #fff; }
.nav-item.active { background: #e8c547; color: #052e16; } /* aktif = kuning */

.sidebar-footer {
  padding: 14px 18px;
  border-top: 1px solid rgba(255,255,255,0.08);
}
.footer-text {
  font-size: 10px; color: rgba(255,255,255,0.25);
  letter-spacing: 0.05em; font-family: 'IBM Plex Mono', monospace;
}

/* ══════════════════════════════════════════════
   [CSS] KONTEN UTAMA
   ══════════════════════════════════════════════ */
.main-content {
  margin-left: var(--sidebar-w); /* geser kanan sebesar lebar sidebar */
  flex: 1;
  display: flex; flex-direction: column;
  min-height: 100vh;
}

/* [CSS] Bar atas halaman */
.topbar {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: 14px 28px;
  display: flex; align-items: center; justify-content: space-between;
  position: sticky; top: 0; z-index: 50; /* sticky agar tetap terlihat saat scroll */
}
.page-title { font-size: 16px; font-weight: 700; }
.date-badge { font-size: 12px; color: var(--text-3); font-family: 'IBM Plex Mono', monospace; }
.page-body  { padding: 24px 28px; }

/* ══════════════════════════════════════════════
   [CSS] NOTIFIKASI FLASH MESSAGE
   ══════════════════════════════════════════════ */
.flash-container { margin-bottom: 16px; }
.flash {
  display: flex; align-items: center; justify-content: space-between;
  padding: 10px 14px; border-radius: var(--radius);
  margin-bottom: 8px; font-size: 13px; font-weight: 500;
}
.flash button { background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; opacity: 0.6; }
.flash-success { background: var(--green-light); color: var(--green); border-left: 3px solid var(--green); }
.flash-warning { background: var(--yellow-light); color: var(--yellow); border-left: 3px solid var(--yellow); }
.flash-danger  { background: var(--red-light);    color: var(--red);    border-left: 3px solid var(--red);   }

/* ══════════════════════════════════════════════
   [CSS] KARTU STATISTIK (STAT CARDS)
   ══════════════════════════════════════════════ */
.stats-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr); /* 4 kolom sama lebar */
  gap: 16px; margin-bottom: 24px;
}
.stat-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 18px 20px;
  position: relative; overflow: hidden;
}
/* [CSS] Garis aksen vertikal di sisi kiri kartu */
.stat-card::before {
  content: ''; position: absolute; top: 0; left: 0;
  width: 3px; height: 100%; background: var(--accent);
}
.stat-card.green::before  { background: var(--green); }
.stat-card.yellow::before { background: var(--yellow); }
.stat-card.blue::before   { background: var(--blue); }
.stat-label { font-size: 10px; font-weight: 600; letter-spacing: 0.1em; text-transform: uppercase; color: var(--text-3); margin-bottom: 8px; }
.stat-value { font-size: 22px; font-weight: 700; letter-spacing: -0.02em; }
.stat-sub   { font-size: 11px; color: var(--text-3); margin-top: 4px; }

/* ══════════════════════════════════════════════
   [CSS] LAYOUT GRID
   ══════════════════════════════════════════════ */
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
.grid-3 { display: grid; grid-template-columns: 2fr 1fr; gap: 20px; margin-bottom: 20px; }

/* ══════════════════════════════════════════════
   [CSS] KARTU KONTEN
   ══════════════════════════════════════════════ */
.card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; }
.card-header {
  padding: 14px 18px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between;
}
.card-title { font-size: 13px; font-weight: 700; }
.card-body  { padding: 18px; }

/* ══════════════════════════════════════════════
   [CSS] TABEL DATA
   ══════════════════════════════════════════════ */
.table-wrap { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; }
.table-toolbar {
  padding: 14px 18px; display: flex; align-items: center;
  justify-content: space-between; border-bottom: 1px solid var(--border); gap: 12px; flex-wrap: wrap;
}
table           { width: 100%; border-collapse: collapse; }
th {
  background: var(--surface-2); padding: 10px 16px;
  text-align: left; font-size: 10px; font-weight: 600;
  letter-spacing: 0.08em; text-transform: uppercase;
  color: var(--text-2); border-bottom: 1px solid var(--border);
}
td              { padding: 11px 16px; border-bottom: 1px solid var(--border); font-size: 13px; }
tr:last-child td{ border-bottom: none; }
tr:hover td     { background: var(--surface-2); } /* highlight baris saat hover */
.td-mono        { font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: var(--text-2); }

/* ══════════════════════════════════════════════
   [CSS] BADGE STATUS
   ══════════════════════════════════════════════ */
.badge {
  display: inline-flex; align-items: center;
  padding: 2px 8px; border-radius: 3px;
  font-size: 11px; font-weight: 600;
}
.badge-green  { background: var(--green-light);  color: var(--green);  }
.badge-yellow { background: var(--yellow-light); color: var(--yellow); }
.badge-red    { background: var(--red-light);    color: var(--red);    }
.badge-blue   { background: var(--blue-light);   color: var(--blue);   }
.badge-gray   { background: #f0ede8;             color: var(--text-2); }

/* ══════════════════════════════════════════════
   [CSS] TOMBOL (BUTTONS)
   ══════════════════════════════════════════════ */
.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 8px 14px; border-radius: var(--radius);
  font-size: 13px; font-weight: 600; cursor: pointer;
  border: 1px solid transparent; text-decoration: none;
  transition: all 0.15s; font-family: inherit;
}
.btn-primary { background: var(--accent);  color: #fff;          border-color: var(--accent); }
.btn-primary:hover { background: #333; }
.btn-outline { background: transparent;    color: var(--text-2);  border-color: var(--border-dark); }
.btn-outline:hover { background: var(--surface-2); }
.btn-danger  { background: var(--red);     color: #fff;           border-color: var(--red); }
.btn-danger:hover { opacity: 0.85; }
.btn-sm { padding: 5px 10px; font-size: 12px; }  /* tombol kecil */
.btn-xs { padding: 3px 8px;  font-size: 11px; }  /* tombol sangat kecil */

/* ══════════════════════════════════════════════
   [CSS] FORM INPUT
   ══════════════════════════════════════════════ */
.form-grid   { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.form-grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }
.form-group  { display: flex; flex-direction: column; gap: 5px; }
.form-group.full { grid-column: 1 / -1; } /* span semua kolom */

label {
  font-size: 11px; font-weight: 600;
  letter-spacing: 0.06em; text-transform: uppercase; color: var(--text-2);
}

input, select, textarea {
  padding: 9px 12px;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  font-family: 'IBM Plex Sans', sans-serif;
  font-size: 13px; color: var(--text); background: var(--surface);
  transition: border-color 0.15s; outline: none;
}
input:focus, select:focus, textarea:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 2px rgba(26,25,22,0.08); /* glow fokus halus */
}
textarea { resize: vertical; min-height: 80px; }

/* [CSS] Input pencarian */
.search-input {
  padding: 7px 12px; border: 1px solid var(--border);
  border-radius: var(--radius); font-size: 13px; font-family: inherit;
  outline: none; min-width: 220px;
}
.search-input:focus { border-color: var(--accent); }

/* ══════════════════════════════════════════════
   [CSS] HEADER HALAMAN
   ══════════════════════════════════════════════ */
.page-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.page-header h2 { font-size: 18px; font-weight: 700; letter-spacing: -0.02em; }
.breadcrumb { font-size: 12px; color: var(--text-3); margin-top: 2px; }
.breadcrumb a { color: var(--text-2); text-decoration: none; }
.breadcrumb a:hover { color: var(--text); }

/* ══════════════════════════════════════════════
   [CSS] FORM CARD (container form)
   ══════════════════════════════════════════════ */
.form-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 24px; }
.form-section-title {
  font-size: 11px; font-weight: 700; letter-spacing: 0.1em;
  text-transform: uppercase; color: var(--text-3);
  padding-bottom: 10px; border-bottom: 1px solid var(--border); margin-bottom: 16px;
}
.form-actions {
  display: flex; align-items: center; gap: 10px;
  padding-top: 16px; border-top: 1px solid var(--border); margin-top: 20px;
}

/* ══════════════════════════════════════════════
   [CSS] INFO GRID (detail pesanan)
   ══════════════════════════════════════════════ */
.info-grid  { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.info-row   { display: flex; flex-direction: column; gap: 2px; }
.info-label { font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: var(--text-3); }
.info-value { font-size: 13px; font-weight: 500; }

/* ══════════════════════════════════════════════
   [CSS] BARIS ITEM PESANAN (form tambah pesanan)
   ══════════════════════════════════════════════ */
.item-row {
  display: grid;
  grid-template-columns: 2fr 1fr 1.2fr 1.2fr auto; /* 5 kolom proporsional */
  gap: 10px; align-items: end;
  padding: 12px; background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: var(--radius); margin-bottom: 10px;
}

/* ══════════════════════════════════════════════
   [CSS] KOTAK RINGKASAN TOTAL HARGA
   ══════════════════════════════════════════════ */
.summary-box { background: var(--surface-2); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px; }
.summary-row {
  display: flex; justify-content: space-between; align-items: center;
  padding: 6px 0; font-size: 13px; border-bottom: 1px solid var(--border);
}
.summary-row:last-child { border-bottom: none; font-weight: 700; font-size: 15px; padding-top: 10px; }

/* ══════════════════════════════════════════════
   [CSS] EMPTY STATE (tabel kosong)
   ══════════════════════════════════════════════ */
.empty-state { text-align: center; padding: 48px 24px; color: var(--text-3); }
.empty-state .empty-icon { font-size: 36px; margin-bottom: 12px; opacity: 0.4; }
.empty-state p { font-size: 13px; }

/* ══════════════════════════════════════════════
   [CSS] CONTAINER GRAFIK CHART.JS
   ══════════════════════════════════════════════ */
.chart-container { position: relative; height: 220px; }

/* ══════════════════════════════════════════════
   [CSS] RESPONSIVE — tampilan di layar kecil
   ══════════════════════════════════════════════ */
@media (max-width: 1100px) {
  .stats-grid { grid-template-columns: repeat(2, 1fr); } /* 2 kolom di tablet */
  .grid-2, .grid-3 { grid-template-columns: 1fr; }       /* 1 kolom di tablet */
}
@media (max-width: 768px) {
  :root { --sidebar-w: 0px; }                    /* sembunyikan sidebar di HP */
  .sidebar { display: none; }
  .main-content { margin-left: 0; }
  .stats-grid { grid-template-columns: 1fr 1fr; }
  .form-grid, .form-grid-3 { grid-template-columns: 1fr; }
}
/* ══════════════════════════════════════════════
   [CSS] DARK MODE — variabel warna override
   ══════════════════════════════════════════════ */
body.dark-mode {
  --bg:        #052e16;
  --surface:   #0f3d20;
  --surface-2: #14532d;
  --border:    #166534;
  --border-dark:#15803d;
  --text:      #f0fdf4;
  --text-2:    #86efac;
  --text-3:    #4ade80;
}
body.dark-mode .sidebar        { background: #052e16; }
body.dark-mode .topbar         { background: #0f3d20; border-color: #166534; }
body.dark-mode input,
body.dark-mode select,
body.dark-mode textarea        { background: #14532d; color: #f0fdf4; border-color: #166534; }
body.dark-mode .search-input   { background: #14532d; color: #f0fdf4; }
body.dark-mode .form-card      { background: #0f3d20; }
body.dark-mode .table-wrap     { background: #0f3d20; }
body.dark-mode th              { background: #14532d; }
body.dark-mode tr:hover td     { background: #14532d; }
body.dark-mode .stat-card      { background: #0f3d20; }
body.dark-mode .card           { background: #0f3d20; }
body.dark-mode .summary-box    { background: #14532d; }
body.dark-mode .item-row       { background: #14532d; border-color: #166534; }

/* [CSS] Tombol toggle dark/light mode */
.theme-toggle {
  display: flex; align-items: center; justify-content: center;
  width: 34px; height: 34px;
  border-radius: 50%;
  border: 1.5px solid var(--border);
  background: var(--surface);
  cursor: pointer; font-size: 16px;
  transition: all 0.2s;
  color: var(--text-2);
}
.theme-toggle:hover { border-color: var(--accent); color: var(--text); transform: rotate(20deg); }

/* ══════════════════════════════════════════════
   [CSS] NOTIFIKASI REAL-TIME
   ══════════════════════════════════════════════ */
.notif-bell {
  position: relative; cursor: pointer;
  width: 34px; height: 34px;
  display: flex; align-items: center; justify-content: center;
  border-radius: 50%; border: 1.5px solid var(--border);
  background: var(--surface); font-size: 16px; color: var(--text-2);
  transition: all 0.2s; text-decoration: none;
}
.notif-bell:hover { border-color: var(--accent); color: var(--text); }
.notif-badge {
  position: absolute; top: -4px; right: -4px;
  background: var(--red); color: #fff;
  font-size: 9px; font-weight: 700;
  width: 16px; height: 16px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  border: 2px solid var(--bg);
  animation: pulse 1.5s infinite;
}
@keyframes pulse {
  0%   { transform: scale(1); }
  50%  { transform: scale(1.2); }
  100% { transform: scale(1); }
}

/* [CSS] Panel notifikasi dropdown */
.notif-panel {
  position: fixed; top: 60px; right: 20px;
  width: 320px; background: var(--surface);
  border: 1px solid var(--border); border-radius: 10px;
  box-shadow: 0 8px 32px rgba(0,0,0,0.15);
  z-index: 999; display: none;
  animation: slideDown 0.2s ease;
}
.notif-panel.show { display: block; }
@keyframes slideDown {
  from { opacity: 0; transform: translateY(-10px); }
  to   { opacity: 1; transform: translateY(0); }
}
.notif-header {
  padding: 14px 16px; border-bottom: 1px solid var(--border);
  display: flex; justify-content: space-between; align-items: center;
}
.notif-header span { font-size: 13px; font-weight: 700; }
.notif-clear { font-size: 11px; color: var(--text-3); cursor: pointer; }
.notif-clear:hover { color: var(--red); }
.notif-list { max-height: 300px; overflow-y: auto; }
.notif-item {
  padding: 12px 16px; border-bottom: 1px solid var(--border);
  display: flex; gap: 10px; align-items: flex-start;
  cursor: pointer; transition: background 0.1s;
}
.notif-item:hover  { background: var(--surface-2); }
.notif-item:last-child { border-bottom: none; }
.notif-icon { font-size: 20px; flex-shrink: 0; margin-top: 2px; }
.notif-content .notif-title { font-size: 12px; font-weight: 700; color: var(--text); }
.notif-content .notif-desc  { font-size: 11px; color: var(--text-3); margin-top: 2px; }
.notif-content .notif-time  { font-size: 10px; color: var(--text-3); margin-top: 4px; font-family: monospace; }
.notif-empty { padding: 32px 16px; text-align: center; color: var(--text-3); font-size: 12px; }

/* ══════════════════════════════════════════════
   [CSS] KATALOG PRODUK (Marketplace Style)
   ══════════════════════════════════════════════ */
.katalog-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
  gap: 16px;
  margin-top: 8px;
}
.katalog-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
  transition: all 0.2s;
  cursor: pointer;
  text-decoration: none;
  display: block;
}
.katalog-card:hover {
  transform: translateY(-4px);
  box-shadow: 0 8px 24px rgba(0,0,0,0.12);
  border-color: var(--border-dark);
}
/* [CSS] Container gambar produk */
.katalog-img {
  width: 100%; aspect-ratio: 1/1;
  background: var(--surface-2);
  display: flex; align-items: center; justify-content: center;
  overflow: hidden; position: relative;
}
.katalog-img img {
  width: 100%; height: 100%; object-fit: cover;
  transition: transform 0.3s;
}
.katalog-card:hover .katalog-img img { transform: scale(1.05); }
/* [CSS] Placeholder jika tidak ada gambar */
.katalog-img-placeholder {
  font-size: 48px; opacity: 0.3; user-select: none;
}
/* [CSS] Badge stok menipis */
.katalog-badge {
  position: absolute; top: 8px; left: 8px;
  background: var(--red); color: #fff;
  font-size: 9px; font-weight: 700;
  padding: 2px 6px; border-radius: 3px;
  letter-spacing: 0.05em;
}
.katalog-badge.green { background: var(--green); }
/* [CSS] Info produk di bawah gambar */
.katalog-info { padding: 12px; }
.katalog-nama {
  font-size: 13px; font-weight: 600; color: var(--text);
  margin-bottom: 4px; line-height: 1.4;
  display: -webkit-box; -webkit-line-clamp: 2;
  -webkit-box-orient: vertical; overflow: hidden;
}
.katalog-kategori { font-size: 10px; color: var(--text-3); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.06em; }
.katalog-harga {
  font-size: 16px; font-weight: 800; color: var(--red);
  letter-spacing: -0.02em;
}
.katalog-stok { font-size: 11px; color: var(--text-3); margin-top: 4px; }
.katalog-footer {
  padding: 8px 12px; border-top: 1px solid var(--border);
  display: flex; justify-content: space-between; align-items: center;
}
.katalog-footer span { font-size: 10px; color: var(--text-3); font-family: monospace; }

/* [CSS] Filter katalog */
.katalog-filter {
  display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px;
}
.filter-btn {
  padding: 5px 14px; border-radius: 20px;
  border: 1.5px solid var(--border); background: var(--surface);
  font-size: 12px; font-weight: 600; cursor: pointer;
  color: var(--text-2); transition: all 0.15s;
  text-decoration: none; display: inline-block;
}
.filter-btn:hover, .filter-btn.active {
  background: var(--accent); color: #fff; border-color: var(--accent);
}
body.dark-mode .filter-btn.active { color: #f0fdf4; }

/* [CSS] Upload foto produk */
.upload-area {
  border: 2px dashed var(--border);
  border-radius: 8px; padding: 24px;
  text-align: center; cursor: pointer;
  transition: all 0.2s; background: var(--surface-2);
}
.upload-area:hover { border-color: var(--accent); background: var(--surface); }
.upload-preview {
  width: 120px; height: 120px; object-fit: cover;
  border-radius: 8px; border: 1px solid var(--border);
  margin-top: 8px;
}

"""


# -----------------------------------------------------------------------------
# [PYTHON + HTML] Template dasar: sidebar + topbar yang dipakai semua halaman
# Menggunakan Jinja2 template syntax: {{ variabel }}, {% blok %}
# -----------------------------------------------------------------------------
BASE_TEMPLATE = """
<!DOCTYPE html>
<!-- [HTML] Deklarasi tipe dokumen HTML5 -->
<html lang="id">
<head>
    <!-- [HTML] Meta charset: menentukan encoding karakter UTF-8 (mendukung huruf Indonesia) -->
    <meta charset="UTF-8">
    <!-- [HTML] Meta viewport: agar tampilan responsif di perangkat mobile -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- [HTML] Judul tab browser — diisi oleh blok child template -->
    <title>{{ page_title }} — PendidikanStore APSI</title>

    <!-- [HTML] Google Fonts: IBM Plex Sans (teks) + IBM Plex Mono (kode/angka) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">

    <!-- [HTML] Menyisipkan CSS global dari variabel Python CSS_GLOBAL -->
    <style>
        /* [CSS] Semua gaya sudah didefinisikan di variabel CSS_GLOBAL di atas */
        {{ css|safe }}
    </style>
</head>
<body>

<!-- ════════════════════════════════════════════
     [HTML] SIDEBAR NAVIGASI KIRI
     ════════════════════════════════════════════ -->
<aside class="sidebar">

    <!-- [HTML] Logo / brand aplikasi -->
    <div class="sidebar-brand">
        <span class="brand-icon">⬡</span>
        <div>
            <div class="brand-title">PendidikanStore</div>
            <div class="brand-sub">Alat Tulis & Perlengkapan Sekolah</div>
        </div>
    </div>

    <!-- [HTML] Menu navigasi utama -->
    <nav class="sidebar-nav">
        <div class="nav-section-label">MENU UTAMA</div>

        <!-- [HTML] Link navigasi — class 'active' ditentukan oleh variabel Python 'active_menu' -->
        <a href="/dashboard" class="nav-item {{ 'active' if active_menu == 'dashboard' else '' }}">
            <span class="nav-icon">▦</span> Dashboard
        </a>
        <a href="/pelanggan" class="nav-item {{ 'active' if active_menu == 'pelanggan' else '' }}">
            <span class="nav-icon">◈</span> Pelanggan
        </a>
        <a href="/produk" class="nav-item {{ 'active' if active_menu == 'produk' else '' }}">
            <span class="nav-icon">◉</span> Produk & Stok
        </a>
        <a href="/katalog" class="nav-item {{ 'active' if active_menu == 'katalog' else '' }}"
           style="padding-left:24px;font-size:12px">
            <span class="nav-icon">🛍</span> Katalog
        </a>
        <a href="/pesanan" class="nav-item {{ 'active' if active_menu == 'pesanan' else '' }}">
            <span class="nav-icon">◎</span> Pesanan
        </a>

        <div class="nav-section-label" style="margin-top:1.5rem">ANALITIK</div>
        <a href="/laporan" class="nav-item {{ 'active' if active_menu == 'laporan' else '' }}">
            <span class="nav-icon">◧</span> Laporan
        </a>
    </nav>

    <div class="sidebar-footer">
        <div class="footer-text">PendidikanStore — Alat Tulis & Perlengkapan Sekolah</div>
    </div>
</aside>

<!-- ════════════════════════════════════════════
     [HTML] KONTEN UTAMA KANAN
     ════════════════════════════════════════════ -->
<main class="main-content">

    <!-- [HTML] Bar atas: judul halaman + tanggal hari ini -->
    <header class="topbar">
        <h1 class="page-title">{{ page_title }}</h1>
        <div class="topbar-right" style="display:flex;align-items:center;gap:10px">
            <span class="date-badge">{{ now.strftime('%d %b %Y') }}</span>

            <!-- [HTML] Tombol Dark Mode Toggle -->
            <button class="theme-toggle" id="themeToggle" title="Ganti tema" onclick="toggleTheme()">
                🌙
            </button>

            <!-- [HTML] Tombol Notifikasi -->
            <div style="position:relative">
                <div class="notif-bell" id="notifBell" onclick="toggleNotif()">
                    🔔
                    <span class="notif-badge" id="notifBadge" style="display:none">0</span>
                </div>
                <!-- [HTML] Panel dropdown notifikasi -->
                <div class="notif-panel" id="notifPanel">
                    <div class="notif-header">
                        <span>🔔 Notifikasi</span>
                        <span class="notif-clear" onclick="clearNotif()">Hapus semua</span>
                    </div>
                    <div class="notif-list" id="notifList">
                        <div class="notif-empty">Tidak ada notifikasi baru</div>
                    </div>
                </div>
            </div>

            <!-- [HTML] Link Katalog Produk -->
            <a href="/katalog" style="padding:5px 12px;background:var(--surface-2);
               border:1px solid var(--border);border-radius:4px;font-size:11px;
               font-weight:600;color:var(--text-2);text-decoration:none">
                🛍 Katalog
            </a>

<!-- [HTML] Sistem login dinonaktifkan — tombol user/logout dihapus -->
        </div>
    </header>

    <!-- ════════════════════════════════════════
         [HTML] FLASH MESSAGES (notifikasi sekali tampil)
         Dikirim dari Python menggunakan flash('pesan', 'kategori')
         ════════════════════════════════════════ -->
    {% with messages = get_flashed_messages(with_categories=true) %}
        {% if messages %}
            <div class="flash-container" style="padding: 0 28px; padding-top: 16px;">
            {% for category, message in messages %}
                <!-- [HTML] Kelas flash disesuaikan dengan kategori: success/warning/danger -->
                <div class="flash flash-{{ category }}">
                    <span>{{ message }}</span>
                    <!-- [HTML] Tombol tutup notifikasi menggunakan JavaScript inline -->
                    <button onclick="this.parentElement.remove()">×</button>
                </div>
            {% endfor %}
            </div>
        {% endif %}
    {% endwith %}

    <!-- [HTML] Area konten halaman — diisi oleh masing-masing template -->
    <div class="page-body">
        {{ content|safe }}
    </div>
</main>

<!-- ════════════════════════════════════════════
     [JAVASCRIPT] Script utama — dimuat di akhir body
     agar HTML sudah selesai di-render sebelum JS dijalankan
     ════════════════════════════════════════════ -->
<script>
// [JAVASCRIPT] Auto-hide flash messages setelah 4 detik
document.querySelectorAll('.flash').forEach(function(el) {
    // setTimeout: menunda eksekusi fungsi selama 4000ms (4 detik)
    setTimeout(function() {
        el.style.opacity = '0';              // mulai animasi menghilang
        el.style.transition = 'opacity 0.4s'; // durasi animasi: 0.4 detik
        setTimeout(function() {
            el.remove();                     // hapus elemen dari DOM setelah animasi selesai
        }, 400);
    }, 4000);
});

// [JAVASCRIPT] Fungsi pembantu: format angka ke Rupiah
// Contoh: formatRupiah(8500000) → "Rp 8.500.000"
function formatRupiah(angka) {
    // parseInt: pastikan angka adalah bilangan bulat
    // toLocaleString('id-ID'): format sesuai locale Indonesia (pemisah ribuan = titik)
    return 'Rp ' + parseInt(angka || 0).toLocaleString('id-ID');
}

// ════════════════════════════════════════════
// [JAVASCRIPT] Logika form TAMBAH PESANAN
// Mengelola baris item produk secara dinamis
// ════════════════════════════════════════════

let itemCount = 1; // [JAVASCRIPT] Penghitung baris item pesanan

// [JAVASCRIPT] Fungsi menambah baris item produk baru
function addItem() {
    itemCount++;  // increment counter
    
    // querySelector: cari elemen dengan id 'items-container'
    const container = document.getElementById('items-container');
    if (!container) return; // guard clause: hentikan jika elemen tidak ada
    
    // innerHTML: ambil isi template pilihan produk yang tersembunyi
    const produkOptions = document.getElementById('produk-options-template') ?
                          document.getElementById('produk-options-template').innerHTML : '';
    
    // createElement: membuat elemen <div> baru
    const row = document.createElement('div');
    row.className = 'item-row';    // set class CSS
    row.id = 'item-row-' + itemCount; // set id unik per baris
    
    // innerHTML: isi HTML baris item baru menggunakan template literal
    row.innerHTML = `
        <div class="form-group">
            <label>Produk</label>
            <!-- [HTML] Dropdown produk dengan event handler onchange -->
            <select name="produk_id[]" onchange="updateHarga(this, ${itemCount})" required>
                <option value="">— Pilih Produk —</option>
                ${produkOptions}
            </select>
        </div>
        <div class="form-group">
            <label>Jumlah</label>
            <!-- [HTML] Input jumlah dengan event oninput untuk update subtotal realtime -->
            <input type="number" name="jumlah[]" id="jumlah-${itemCount}" value="1" min="1" oninput="hitungSubtotal(${itemCount})" required>
        </div>
        <div class="form-group">
            <label>Harga Satuan</label>
            <!-- [HTML] Input readonly: hanya bisa diisi oleh JavaScript, bukan user -->
            <input type="text" id="harga-display-${itemCount}" readonly placeholder="Rp 0">
            <!-- [HTML] Hidden input: menyimpan nilai harga asli (tanpa format Rupiah) -->
            <input type="hidden" id="harga-${itemCount}" value="0">
        </div>
        <div class="form-group">
            <label>Subtotal</label>
            <input type="text" id="subtotal-display-${itemCount}" readonly placeholder="Rp 0">
        </div>
        <div class="form-group">
            <label>&nbsp;</label>
            <!-- [HTML] Tombol hapus baris -->
            <button type="button" class="btn btn-danger btn-sm" onclick="removeItem(${itemCount})">✕</button>
        </div>
    `;
    container.appendChild(row); // tambahkan baris ke container
}

// [JAVASCRIPT] Fungsi menghapus baris item pesanan
function removeItem(id) {
    const row = document.getElementById('item-row-' + id);
    if (row) {
        row.remove();    // hapus elemen dari DOM
        hitungTotal();   // hitung ulang total setelah baris dihapus
    }
}

// [JAVASCRIPT] Fungsi update harga saat produk dipilih dari dropdown
function updateHarga(select, id) {
    // dataset: mengakses atribut data-* pada elemen HTML
    const option = select.options[select.selectedIndex]; // opsi yang dipilih
    const harga  = option.dataset.harga || 0;            // ambil data-harga dari <option>
    const stok   = option.dataset.stok  || 0;            // ambil data-stok dari <option>
    
    // getElementById: mencari elemen berdasarkan id
    const hargaInput   = document.getElementById('harga-' + id);
    const hargaDisplay = document.getElementById('harga-display-' + id);
    const jumlahInput  = document.getElementById('jumlah-' + id);
    
    if (hargaInput)   hargaInput.value   = harga;              // simpan nilai angka asli
    if (hargaDisplay) hargaDisplay.value = formatRupiah(harga); // tampilkan format Rupiah
    if (jumlahInput)  jumlahInput.max    = stok;               // batasi jumlah = stok tersedia
    
    hitungSubtotal(id); // hitung subtotal baris ini
}

// [JAVASCRIPT] Fungsi hitung subtotal satu baris item
function hitungSubtotal(id) {
    // parseFloat: mengubah string ke bilangan desimal
    const harga  = parseFloat(document.getElementById('harga-'  + id)?.value || 0);
    // parseInt: mengubah string ke bilangan bulat
    const jumlah = parseInt(document.getElementById('jumlah-' + id)?.value || 0);
    const subtotal = harga * jumlah; // perkalian harga × jumlah
    
    const el = document.getElementById('subtotal-display-' + id);
    if (el) el.value = formatRupiah(subtotal); // tampilkan subtotal terformat
    
    hitungTotal(); // update total keseluruhan
}

// [JAVASCRIPT] Fungsi hitung total semua baris item
function hitungTotal() {
    // querySelectorAll: ambil semua elemen dengan class 'item-row'
    const rows  = document.querySelectorAll('.item-row');
    let total   = 0; // akumulator total harga
    
    // forEach: iterasi setiap baris item
    rows.forEach(function(row) {
        const id     = row.id.split('-')[2];   // ambil nomor id dari "item-row-X"
        const harga  = parseFloat(document.getElementById('harga-'  + id)?.value || 0);
        const jumlah = parseInt(document.getElementById('jumlah-' + id)?.value || 1);
        total += harga * jumlah; // tambahkan ke total
    });
    
    // Ambil nilai diskon dari input (default 0 jika tidak ada)
    const diskonEl = document.getElementById('diskon');
    const diskon   = diskonEl ? parseFloat(diskonEl.value || 0) : 0;
    const totalBayar = total * (1 - diskon / 100); // hitung setelah potongan diskon
    
    // Update tampilan total di ringkasan harga
    const elTotal        = document.getElementById('total-harga-display');
    const elDiskonNom    = document.getElementById('diskon-nominal');
    const elTotalBayar   = document.getElementById('total-bayar-display');
    
    if (elTotal)      elTotal.textContent      = formatRupiah(total);
    if (elDiskonNom)  elDiskonNom.textContent  = formatRupiah(total * diskon / 100);
    if (elTotalBayar) elTotalBayar.textContent = formatRupiah(totalBayar);
}
</script>

<!-- [HTML] Slot untuk script tambahan dari masing-masing halaman (misal: Chart.js) -->
{{ extra_scripts|safe }}

</body>
</html>
"""


# -----------------------------------------------------------------------------
# [PYTHON] Fungsi pembantu render template: menyatukan base + konten halaman
# -----------------------------------------------------------------------------
def render_page(page_title, active_menu, content, extra_scripts=""):
    """
    [PYTHON] Menggabungkan BASE_TEMPLATE dengan konten halaman tertentu.
    
    Parameter:
        page_title    → judul halaman (ditampilkan di tab dan topbar)
        active_menu   → nama menu aktif untuk highlight sidebar
        content       → konten HTML halaman (dari template masing-masing)
        extra_scripts → JavaScript tambahan (misal: inisialisasi Chart.js)
    """
    return render_template_string(
        BASE_TEMPLATE,            # template dasar dengan sidebar dan topbar
        page_title=page_title,    # judul halaman
        active_menu=active_menu,  # menu yang sedang aktif
        content=content,          # konten halaman
        extra_scripts=extra_scripts, # script tambahan
        css=CSS_GLOBAL,           # sisipkan CSS global
        now=datetime.now()        # waktu sekarang untuk topbar
    )


# =============================================================================
# [PYTHON] ── BAGIAN 7: ROUTE DASHBOARD
# Route adalah URL yang dikaitkan dengan fungsi Python menggunakan @app.route
# =============================================================================

@app.route('/')              # URL root '/' → redirect ke /dashboard
def index():
    return redirect('/dashboard')  # redirect ke halaman dashboard


@app.route('/dashboard')     # URL /dashboard → tampilkan halaman dashboard
def dashboard():
    """
    [PYTHON] Mengambil data statistik dari database untuk ditampilkan di dashboard.
    Semua query menggunakan DictCursor sehingga hasilnya berupa dictionary.
    """
    cur = mysql.connection.cursor()  # membuka cursor untuk query database

    # [SQL] Hitung total pelanggan aktif
    cur.execute("SELECT COUNT(*) as total FROM pelanggan WHERE status='aktif'")
    total_pelanggan = cur.fetchone()['total']  # fetchone() → ambil 1 baris

    # [SQL] Hitung total produk aktif
    cur.execute("SELECT COUNT(*) as total FROM produk WHERE status='aktif'")
    total_produk = cur.fetchone()['total']

    # [SQL] Hitung pesanan bulan berjalan menggunakan fungsi MySQL MONTH() dan YEAR()
    cur.execute("""
        SELECT COUNT(*) as total FROM pesanan
        WHERE MONTH(tanggal_pesan) = MONTH(NOW())
          AND YEAR(tanggal_pesan)  = YEAR(NOW())
    """)
    total_pesanan = cur.fetchone()['total']

    # [SQL] Hitung total pendapatan bulan ini (hanya pesanan berstatus 'selesai')
    # COALESCE: mengembalikan 0 jika SUM() hasilnya NULL (tidak ada data)
    cur.execute("""
        SELECT COALESCE(SUM(total_bayar), 0) as total FROM pesanan
        WHERE status = 'selesai'
          AND MONTH(tanggal_pesan) = MONTH(NOW())
          AND YEAR(tanggal_pesan)  = YEAR(NOW())
    """)
    total_penjualan = cur.fetchone()['total']

    # [SQL] Ambil 5 pesanan terbaru dengan JOIN ke tabel pelanggan
    # JOIN: menggabungkan 2 tabel berdasarkan kolom yang sama (pelanggan_id = id)
    cur.execute("""
        SELECT p.no_pesanan, pl.nama as nama_pelanggan,
               p.tanggal_pesan, p.total_bayar, p.status
        FROM pesanan p
        JOIN pelanggan pl ON p.pelanggan_id = pl.id
        ORDER BY p.created_at DESC
        LIMIT 5
    """)
    pesanan_terbaru = cur.fetchall()  # fetchall() → ambil semua baris hasil query

    # [SQL] Ambil produk dengan stok di bawah atau sama dengan stok minimum
    cur.execute("""
        SELECT nama, stok, stok_minimum, satuan FROM produk
        WHERE stok <= stok_minimum AND status = 'aktif'
    """)
    stok_menipis = cur.fetchall()

    # [SQL] Data grafik tren penjualan 6 bulan terakhir
    # DATE_FORMAT: memformat tanggal, DATE_SUB: mengurangi tanggal
    cur.execute("""
        SELECT DATE_FORMAT(tanggal_pesan, '%b %Y') as bulan,
               COALESCE(SUM(total_bayar), 0) as total
        FROM pesanan
        WHERE status = 'selesai'
          AND tanggal_pesan >= DATE_SUB(NOW(), INTERVAL 6 MONTH)
        GROUP BY DATE_FORMAT(tanggal_pesan, '%Y-%m')
        ORDER BY tanggal_pesan ASC
    """)
    grafik_data = cur.fetchall()

    # [SQL] Data distribusi pesanan per status bulan ini
    cur.execute("""
        SELECT status, COUNT(*) as jumlah FROM pesanan
        WHERE MONTH(tanggal_pesan) = MONTH(NOW())
        GROUP BY status
    """)
    status_data = cur.fetchall()

    cur.close()  # tutup cursor setelah semua query selesai

    # [PYTHON] Konversi list dict ke JSON string untuk dipakai Chart.js di JavaScript
    # json.dumps: mengubah Python list/dict menjadi string JSON
    grafik_json = json.dumps(list(grafik_data), default=str)
    status_json = json.dumps(list(status_data), default=str)

    # [HTML] Template halaman dashboard
    content = f"""
    <!-- [HTML] Grid 4 kartu statistik -->
    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-label">Pelanggan Aktif</div>
            <div class="stat-value">{total_pelanggan}</div>
            <div class="stat-sub">pelanggan terdaftar</div>
        </div>
        <div class="stat-card green">
            <div class="stat-label">Produk Aktif</div>
            <div class="stat-value">{total_produk}</div>
            <div class="stat-sub">SKU tersedia</div>
        </div>
        <div class="stat-card blue">
            <div class="stat-label">Pesanan Bulan Ini</div>
            <div class="stat-value">{total_pesanan}</div>
            <div class="stat-sub">order masuk</div>
        </div>
        <div class="stat-card yellow">
            <div class="stat-label">Pendapatan Bulan Ini</div>
            <!-- [HTML] Format Rupiah dihasilkan Python sebelum dikirim ke HTML -->
            <div class="stat-value" style="font-size:16px">{f"Rp {int(total_penjualan):,}".replace(",",".")}</div>
            <div class="stat-sub">dari pesanan selesai</div>
        </div>
    </div>

    <!-- [HTML] Layout 2 kolom: grafik besar + grafik kecil -->
    <div class="grid-3">
        <!-- [HTML] Grafik tren penjualan — canvas untuk Chart.js -->
        <div class="card">
            <div class="card-header"><span class="card-title">Tren Penjualan 6 Bulan</span></div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartPenjualan"></canvas>
                </div>
            </div>
        </div>
        <!-- [HTML] Grafik status pesanan — doughnut chart -->
        <div class="card">
            <div class="card-header"><span class="card-title">Status Pesanan</span></div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartStatus"></canvas>
                </div>
            </div>
        </div>
    </div>

    <!-- [HTML] Layout 2 kolom: tabel pesanan terbaru + stok menipis -->
    <div class="grid-2">
        <!-- [HTML] Tabel pesanan terbaru -->
        <div class="card">
            <div class="card-header">
                <span class="card-title">Pesanan Terbaru</span>
                <a href="/pesanan" class="btn btn-outline btn-xs">Lihat Semua</a>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>No. Pesanan</th><th>Pelanggan</th>
                        <th>Total</th><th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    {_render_pesanan_terbaru(pesanan_terbaru)}
                </tbody>
            </table>
        </div>

        <!-- [HTML] Tabel stok menipis -->
        <div class="card">
            <div class="card-header">
                <span class="card-title">⚠ Stok Menipis</span>
                <a href="/produk" class="btn btn-outline btn-xs">Kelola Stok</a>
            </div>
            {_render_stok_menipis(stok_menipis)}
        </div>
    </div>
    """

    # [JAVASCRIPT] Script Chart.js untuk grafik — dimuat setelah data tersedia
    extra_scripts = f"""
    <!-- [HTML] Load library Chart.js dari CDN -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
    <script>
    // [JAVASCRIPT] Data grafik dari Python → diparse sebagai array JavaScript
    const grafikData = {grafik_json};  // tren penjualan 6 bulan
    const statusData = {status_json};  // distribusi status pesanan

    // [JAVASCRIPT] Inisialisasi Chart 1: Line Chart — Tren Penjualan
    const ctx1 = document.getElementById('chartPenjualan').getContext('2d');
    new Chart(ctx1, {{
        type: 'line',   // jenis grafik: garis
        data: {{
            // map: mengubah array objek menjadi array nilai tertentu
            labels: grafikData.map(function(d) {{ return d.bulan; }}),
            datasets: [{{
                label: 'Pendapatan',
                data: grafikData.map(function(d) {{ return d.total; }}),
                borderColor: '#052e16',               // warna garis
                backgroundColor: 'rgba(26,25,22,0.05)', // area di bawah garis
                borderWidth: 2, pointRadius: 4,
                fill: true, tension: 0.3              // kurva halus
            }}]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{ legend: {{ display: false }} }},
            scales: {{
                // [JAVASCRIPT] Format label sumbu Y ke format Rupiah singkat
                y: {{ ticks: {{ callback: function(v) {{ return 'Rp ' + (v/1000000).toFixed(1) + 'jt'; }} }}, grid: {{ color: '#bbf7d0' }} }},
                x: {{ grid: {{ display: false }} }}
            }}
        }}
    }});

    // [JAVASCRIPT] Inisialisasi Chart 2: Doughnut Chart — Status Pesanan
    const warnaBadge = {{
        pending: '#4ade80', diproses: '#1a4a8a',
        dikirim: '#e76f00', selesai:  '#2d6a4f', dibatalkan: '#c0392b'
    }};
    const ctx2 = document.getElementById('chartStatus').getContext('2d');
    new Chart(ctx2, {{
        type: 'doughnut',  // jenis grafik: donat
        data: {{
            labels: statusData.map(function(d) {{ return d.status.toUpperCase(); }}),
            datasets: [{{
                data: statusData.map(function(d) {{ return d.jumlah; }}),
                // Map setiap status ke warnanya masing-masing
                backgroundColor: statusData.map(function(d) {{ return warnaBadge[d.status] || '#ccc'; }}),
                borderWidth: 2, borderColor: '#fff'
            }}]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{ legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }} }},
            cutout: '65%'  // ukuran lubang tengah donat
        }}
    }});
    </script>
    """

    return render_page('Dashboard', 'dashboard', content, extra_scripts)


# [PYTHON] Fungsi pembantu untuk merender baris tabel pesanan terbaru
def _render_pesanan_terbaru(data):
    """[PYTHON] Menghasilkan HTML baris <tr> untuk tabel pesanan terbaru."""
    badge_map = {  # peta status → class badge CSS
        'pending': 'badge-gray', 'diproses': 'badge-blue',
        'dikirim': 'badge-yellow', 'selesai': 'badge-green', 'dibatalkan': 'badge-red'
    }
    if not data:
        # [HTML] Tampilan empty state jika tidak ada data
        return '<tr><td colspan="4"><div class="empty-state"><p>Belum ada pesanan</p></div></td></tr>'
    
    rows = ''
    for p in data:
        badge = badge_map.get(p['status'], 'badge-gray')
        total = f"Rp {int(p['total_bayar']):,}".replace(',', '.')
        rows += f"""
        <tr>
            <td class="td-mono">{p['no_pesanan']}</td>
            <td>{p['nama_pelanggan']}</td>
            <td class="td-mono">{total}</td>
            <td><span class="badge {badge}">{p['status'].upper()}</span></td>
        </tr>"""
    return rows


# [PYTHON] Fungsi pembantu untuk merender tabel stok menipis
def _render_stok_menipis(data):
    """[PYTHON] Menghasilkan HTML tabel atau pesan kosong untuk stok menipis."""
    if not data:
        return '<div class="card-body"><div class="empty-state"><div class="empty-icon">✓</div><p>Stok semua aman</p></div></div>'
    
    rows = ''
    for s in data:
        rows += f"""<tr>
            <td>{s['nama']}</td>
            <td><span class="badge badge-red">{s['stok']} {s['satuan']}</span></td>
            <td class="td-mono">{s['stok_minimum']}</td>
        </tr>"""
    return f"""<table>
        <thead><tr><th>Produk</th><th>Stok</th><th>Minimum</th></tr></thead>
        <tbody>{rows}</tbody>
    </table>"""


# =============================================================================
# [PYTHON] ── BAGIAN 8: ROUTE PELANGGAN (CRUD)
# CRUD = Create (tambah), Read (tampilkan), Update (edit), Delete (hapus)
# =============================================================================

@app.route('/pelanggan')   # menampilkan daftar semua pelanggan
def pelanggan_index():
    """[PYTHON] Mengambil dan menampilkan daftar pelanggan dari database."""
    cur = mysql.connection.cursor()
    search = request.args.get('search', '')  # request.args: ambil parameter dari URL (?search=...)

    if search:
        # [SQL] Query dengan LIKE untuk pencarian — %search% berarti mengandung kata search
        cur.execute("""
            SELECT * FROM pelanggan
            WHERE nama LIKE %s OR kode_pelanggan LIKE %s OR kota LIKE %s
            ORDER BY id DESC
        """, (f'%{search}%', f'%{search}%', f'%{search}%'))  # %s = placeholder aman (mencegah SQL injection)
    else:
        cur.execute("SELECT * FROM pelanggan ORDER BY id DESC")  # ambil semua, urut terbaru

    data = cur.fetchall()
    cur.close()

    # [PYTHON] Bangun HTML baris tabel pelanggan
    rows = ''
    for p in data:
        badge = 'badge-green' if p['status'] == 'aktif' else 'badge-red'
        label = 'AKTIF' if p['status'] == 'aktif' else 'NONAKTIF'
        rows += f"""
        <tr>
            <td class="td-mono">{p['kode_pelanggan']}</td>
            <td><strong>{p['nama']}</strong></td>
            <td style="color:var(--text-2)">{p['email'] or '-'}</td>
            <td class="td-mono">{p['telepon'] or '-'}</td>
            <td>{p['kota'] or '-'}</td>
            <td><span class="badge {badge}">{label}</span></td>
            <td>
                <div style="display:flex;gap:6px">
                    <a href="/pelanggan/edit/{p['id']}" class="btn btn-outline btn-xs">Edit</a>
                    <a href="/pelanggan/hapus/{p['id']}" class="btn btn-danger btn-xs"
                       onclick="return confirm('Nonaktifkan pelanggan ini?')">Nonaktifkan</a>
                </div>
            </td>
        </tr>"""

    if not rows:
        rows = '<tr><td colspan="7"><div class="empty-state"><div class="empty-icon">◈</div><p>Belum ada data pelanggan</p></div></td></tr>'

    # [HTML] Template halaman daftar pelanggan
    content = f"""
    <div class="page-header">
        <div>
            <h2>Daftar Pelanggan</h2>
            <div class="breadcrumb">CRM / <span>Pelanggan</span></div>
        </div>
        <a href="/pelanggan/tambah" class="btn btn-primary">+ Tambah Pelanggan</a>
    </div>
    <div class="table-wrap">
        <div class="table-toolbar">
            <span style="font-size:13px;font-weight:600;color:var(--text-2)">{len(data)} pelanggan</span>
            <!-- [HTML] Form pencarian menggunakan method GET agar query masuk ke URL -->
            <form method="GET" style="display:flex;gap:8px;align-items:center">
                <input class="search-input" type="text" name="search"
                       placeholder="Cari nama / kode / kota..." value="{search}">
                <button type="submit" class="btn btn-outline btn-sm">Cari</button>
                {'<a href="/pelanggan" class="btn btn-outline btn-sm">Reset</a>' if search else ''}
            </form>
        </div>
        <div style="overflow-x:auto">
            <table>
                <thead>
                    <tr>
                        <th>Kode</th><th>Nama</th><th>Email</th>
                        <th>Telepon</th><th>Kota</th><th>Status</th><th>Aksi</th>
                    </tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
        </div>
    </div>
    """
    return render_page('Pelanggan', 'pelanggan', content)


@app.route('/pelanggan/tambah', methods=['GET', 'POST'])
# methods=['GET','POST']: route ini menerima 2 metode HTTP
# GET  → menampilkan form kosong
# POST → memproses data yang dikirim dari form
def pelanggan_tambah():
    """[PYTHON] Menampilkan form dan memproses penambahan pelanggan baru."""
    if request.method == 'POST':
        # [PYTHON] Ambil data dari form HTML yang dikirim via POST
        kode    = generate_kode('PLG', 'pelanggan', 'kode_pelanggan')  # generate kode otomatis
        nama    = request.form['nama']      # request.form: ambil data dari form POST
        email   = request.form['email']
        telepon = request.form['telepon']
        alamat  = request.form['alamat']
        kota    = request.form['kota']

        cur = mysql.connection.cursor()
        # [SQL] INSERT INTO: menambahkan baris baru ke tabel pelanggan
        cur.execute("""
            INSERT INTO pelanggan (kode_pelanggan, nama, email, telepon, alamat, kota)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (kode, nama, email, telepon, alamat, kota))  # %s = parameterized query (aman dari SQL injection)
        mysql.connection.commit()  # commit: menyimpan perubahan ke database
        cur.close()

        flash('Pelanggan berhasil ditambahkan!', 'success')  # kirim notifikasi sukses
        return redirect('/pelanggan')  # redirect kembali ke daftar pelanggan

    # [HTML] Template form tambah pelanggan (ditampilkan saat method = GET)
    content = """
    <div class="page-header">
        <div>
            <h2>Tambah Pelanggan</h2>
            <div class="breadcrumb"><a href="/pelanggan">Pelanggan</a> / Tambah</div>
        </div>
    </div>
    <div class="form-card" style="max-width:680px">
        <div class="form-section-title">Informasi Pelanggan</div>
        <!-- [HTML] Form dengan method POST — data dikirim ke server saat submit -->
        <form method="POST">
            <div class="form-grid">
                <div class="form-group full">
                    <label>Nama Pelanggan / Perusahaan *</label>
                    <!-- [HTML] required: validasi HTML5 — tidak bisa kosong -->
                    <input type="text" name="nama" required placeholder="PT Contoh Jaya">
                </div>
                <div class="form-group">
                    <label>Email</label>
                    <input type="email" name="email" placeholder="email@contoh.com">
                </div>
                <div class="form-group">
                    <label>Telepon</label>
                    <input type="text" name="telepon" placeholder="021-12345678">
                </div>
                <div class="form-group">
                    <label>Kota</label>
                    <input type="text" name="kota" placeholder="Jakarta">
                </div>
                <div class="form-group full">
                    <label>Alamat Lengkap</label>
                    <textarea name="alamat" placeholder="Jl. Contoh No. 1..."></textarea>
                </div>
            </div>
            <div class="form-actions">
                <button type="submit" class="btn btn-primary">Simpan Pelanggan</button>
                <a href="/pelanggan" class="btn btn-outline">Batal</a>
            </div>
        </form>
    </div>
    """
    return render_page('Tambah Pelanggan', 'pelanggan', content)


@app.route('/pelanggan/edit/<int:id>', methods=['GET', 'POST'])
# <int:id>: parameter URL dinamis — id dikonversi ke integer
def pelanggan_edit(id):
    """[PYTHON] Menampilkan form edit dan menyimpan perubahan data pelanggan."""
    cur = mysql.connection.cursor()

    if request.method == 'POST':
        # [PYTHON] Ambil semua field dari form
        nama    = request.form['nama']
        email   = request.form['email']
        telepon = request.form['telepon']
        alamat  = request.form['alamat']
        kota    = request.form['kota']
        status  = request.form['status']

        # [SQL] UPDATE: memperbarui data yang sudah ada berdasarkan id
        cur.execute("""
            UPDATE pelanggan
            SET nama=%s, email=%s, telepon=%s, alamat=%s, kota=%s, status=%s
            WHERE id=%s
        """, (nama, email, telepon, alamat, kota, status, id))
        mysql.connection.commit()
        cur.close()
        flash('Data pelanggan berhasil diperbarui!', 'success')
        return redirect('/pelanggan')

    # [SQL] SELECT: ambil data pelanggan berdasarkan id untuk mengisi form
    cur.execute("SELECT * FROM pelanggan WHERE id=%s", (id,))
    data = cur.fetchone()  # fetchone() karena hanya 1 baris yang dibutuhkan
    cur.close()

    if not data:
        flash('Data tidak ditemukan!', 'danger')
        return redirect('/pelanggan')

    # [PYTHON] Tentukan opsi 'selected' untuk dropdown status
    opt_aktif    = 'selected' if data['status'] == 'aktif' else ''
    opt_nonaktif = 'selected' if data['status'] == 'tidak_aktif' else ''

    # [HTML] Template form edit — nilai input diisi dengan data dari database
    content = f"""
    <div class="page-header">
        <div>
            <h2>Edit Pelanggan</h2>
            <div class="breadcrumb"><a href="/pelanggan">Pelanggan</a> / Edit — {data['kode_pelanggan']}</div>
        </div>
    </div>
    <div class="form-card" style="max-width:680px">
        <div class="form-section-title">Informasi Pelanggan</div>
        <form method="POST">
            <div class="form-grid">
                <div class="form-group full">
                    <label>Nama *</label>
                    <!-- [HTML] value: mengisi input dengan data yang sudah ada -->
                    <input type="text" name="nama" value="{data['nama']}" required>
                </div>
                <div class="form-group">
                    <label>Email</label>
                    <input type="email" name="email" value="{data['email'] or ''}">
                </div>
                <div class="form-group">
                    <label>Telepon</label>
                    <input type="text" name="telepon" value="{data['telepon'] or ''}">
                </div>
                <div class="form-group">
                    <label>Kota</label>
                    <input type="text" name="kota" value="{data['kota'] or ''}">
                </div>
                <div class="form-group">
                    <label>Status</label>
                    <!-- [HTML] select: dropdown pilihan dengan opsi -->
                    <select name="status">
                        <option value="aktif" {opt_aktif}>Aktif</option>
                        <option value="tidak_aktif" {opt_nonaktif}>Tidak Aktif</option>
                    </select>
                </div>
                <div class="form-group full">
                    <label>Alamat</label>
                    <textarea name="alamat">{data['alamat'] or ''}</textarea>
                </div>
            </div>
            <div class="form-actions">
                <button type="submit" class="btn btn-primary">Simpan Perubahan</button>
                <a href="/pelanggan" class="btn btn-outline">Batal</a>
            </div>
        </form>
    </div>
    """
    return render_page('Edit Pelanggan', 'pelanggan', content)


@app.route('/pelanggan/hapus/<int:id>')
def pelanggan_hapus(id):
    """
    [PYTHON] Soft delete: tidak menghapus data, hanya mengubah status menjadi 'tidak_aktif'.
    Ini praktik terbaik agar histori data tetap terjaga.
    """
    cur = mysql.connection.cursor()
    # [SQL] UPDATE status menjadi tidak_aktif (soft delete)
    cur.execute("UPDATE pelanggan SET status='tidak_aktif' WHERE id=%s", (id,))
    mysql.connection.commit()
    cur.close()
    flash('Pelanggan berhasil dinonaktifkan!', 'warning')
    return redirect('/pelanggan')


# =============================================================================
# [PYTHON] ── BAGIAN 9: ROUTE PRODUK (CRUD)
# =============================================================================

@app.route('/produk')
def produk_index():
    """[PYTHON] Menampilkan daftar produk beserta kategori dan status stok."""
    cur = mysql.connection.cursor()
    search = request.args.get('search', '')

    if search:
        # [SQL] LEFT JOIN: tampilkan semua produk, dan nama kategori jika ada
        # LEFT JOIN tetap menampilkan produk meski tidak punya kategori (kategori_id = NULL)
        cur.execute("""
            SELECT p.*, k.nama as nama_kategori
            FROM produk p LEFT JOIN kategori k ON p.kategori_id = k.id
            WHERE p.nama LIKE %s OR p.kode_produk LIKE %s
            ORDER BY p.id DESC
        """, (f'%{search}%', f'%{search}%'))
    else:
        cur.execute("""
            SELECT p.*, k.nama as nama_kategori
            FROM produk p LEFT JOIN kategori k ON p.kategori_id = k.id
            ORDER BY p.id DESC
        """)

    data = cur.fetchall()
    cur.close()

    rows = ''
    for p in data:
        # [PYTHON] Logika kondisi: stok kritis = stok di bawah atau sama dengan minimum
        stok_badge = 'badge-red' if p['stok'] <= p['stok_minimum'] else 'badge-green'
        status_badge = 'badge-green' if p['status'] == 'aktif' else 'badge-red'
        status_label = 'AKTIF' if p['status'] == 'aktif' else 'NONAKTIF'
        harga = f"Rp {int(p['harga']):,}".replace(',', '.')

        # [PYTHON] Tampilkan thumbnail foto jika ada
        thumb = f'<img src="/static/uploads/produk/{p["gambar"]}" style="width:36px;height:36px;object-fit:cover;border-radius:4px;border:1px solid var(--border)">' if p.get("gambar") else f'<div style="width:36px;height:36px;background:var(--surface-2);border-radius:4px;border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:16px">📦</div>'
        rows += f"""
        <tr>
            <td class="td-mono">{p['kode_produk']}</td>
            <td>
                <div style="display:flex;align-items:center;gap:10px">
                    {thumb}
                    <strong>{p['nama']}</strong>
                </div>
            </td>
            <td>{p['nama_kategori'] or '-'}</td>
            <td class="td-mono">{harga}</td>
            <td><span class="badge {stok_badge}">{p['stok']} {p['satuan']}</span></td>
            <td class="td-mono">{p['stok_minimum']}</td>
            <td><span class="badge {status_badge}">{status_label}</span></td>
            <td>
                <div style="display:flex;gap:6px">
                    <a href="/produk/edit/{p['id']}" class="btn btn-outline btn-xs">Edit</a>
                    <a href="/produk/hapus/{p['id']}" class="btn btn-danger btn-xs"
                       onclick="return confirm('Nonaktifkan produk ini?')">Nonaktifkan</a>
                </div>
            </td>
        </tr>"""

    if not rows:
        rows = '<tr><td colspan="8"><div class="empty-state"><div class="empty-icon">◉</div><p>Belum ada produk</p></div></td></tr>'

    content = f"""
    <div class="page-header">
        <div><h2>Produk & Stok</h2><div class="breadcrumb">Inventori / Produk</div></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
            <a href="/notifikasi/stok" class="btn btn-outline btn-sm"
               style="background:#e76f00;color:#fff;border-color:#e76f00"
               onclick="return confirm('Kirim email notifikasi stok menipis ke admin?')">
                &#128276; Notifikasi Stok
            </a>
            <a href="/katalog" class="btn btn-outline btn-sm"
               style="background:#1a4a8a;color:#fff;border-color:#1a4a8a">
                &#128722; Lihat Katalog
            </a>
            <a href="/produk/tambah-foto" class="btn btn-outline btn-sm"
               style="background:#2d6a4f;color:#fff;border-color:#2d6a4f">
                &#128247; Tambah + Foto
            </a>
            <a href="/produk/tambah" class="btn btn-primary">+ Tambah Produk</a>
        </div>
    </div>
    <div class="table-wrap">
        <div class="table-toolbar">
            <span style="font-size:13px;font-weight:600;color:var(--text-2)">{len(data)} produk</span>
            <form method="GET" style="display:flex;gap:8px;align-items:center">
                <input class="search-input" type="text" name="search"
                       placeholder="Cari nama / kode..." value="{search}">
                <button type="submit" class="btn btn-outline btn-sm">Cari</button>
                {'<a href="/produk" class="btn btn-outline btn-sm">Reset</a>' if search else ''}
            </form>
        </div>
        <div style="overflow-x:auto">
            <table>
                <thead>
                    <tr>
                        <th>Kode</th><th>Nama</th><th>Kategori</th><th>Harga</th>
                        <th>Stok</th><th>Min. Stok</th><th>Status</th><th>Aksi</th>
                    </tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
        </div>
    </div>
    """
    return render_page('Produk & Stok', 'produk', content)


@app.route('/produk/tambah', methods=['GET', 'POST'])
def produk_tambah():
    """[PYTHON] Form tambah produk baru dengan pilihan kategori."""
    cur = mysql.connection.cursor()

    if request.method == 'POST':
        kode         = generate_kode('PRD', 'produk', 'kode_produk')
        nama         = request.form['nama']
        kategori_id  = request.form['kategori_id'] or None  # None jika tidak dipilih
        harga        = request.form['harga']
        stok         = request.form['stok']
        stok_minimum = request.form['stok_minimum']
        satuan       = request.form['satuan']
        deskripsi    = request.form.get('deskripsi', '')  # .get() → tidak error jika kosong

        # [SQL] INSERT produk baru
        cur.execute("""
            INSERT INTO produk (kode_produk, nama, kategori_id, harga, stok, stok_minimum, satuan, deskripsi)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (kode, nama, kategori_id, harga, stok, stok_minimum, satuan, deskripsi))
        mysql.connection.commit()
        cur.close()
        flash('Produk berhasil ditambahkan!', 'success')
        return redirect('/produk')

    # [SQL] Ambil semua kategori untuk dropdown
    cur.execute("SELECT * FROM kategori")
    kategori = cur.fetchall()
    cur.close()

    # [PYTHON] Bangun opsi dropdown kategori dari hasil query
    opt_kategori = '<option value="">— Pilih Kategori —</option>'
    for k in kategori:
        opt_kategori += f'<option value="{k["id"]}">{k["nama"]}</option>'

    content = f"""
    <div class="page-header">
        <div><h2>Tambah Produk</h2>
        <div class="breadcrumb"><a href="/produk">Produk</a> / Tambah</div></div>
    </div>
    <div class="form-card" style="max-width:680px">
        <div class="form-section-title">Informasi Produk</div>
        <form method="POST">
            <div class="form-grid">
                <div class="form-group full">
                    <label>Nama Produk *</label>
                    <input type="text" name="nama" required placeholder="Nama produk">
                </div>
                <div class="form-group">
                    <label>Kategori</label>
                    <select name="kategori_id">{opt_kategori}</select>
                </div>
                <div class="form-group">
                    <label>Satuan</label>
                    <select name="satuan">
                        <!-- [HTML] Pilihan satuan produk -->
                        <option value="pcs">pcs</option>
                        <option value="unit">unit</option>
                        <option value="kg">kg</option>
                        <option value="liter">liter</option>
                        <option value="box">box</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Harga Jual (Rp) *</label>
                    <!-- [HTML] type="number" → hanya menerima input angka -->
                    <input type="number" name="harga" required placeholder="0" min="0">
                </div>
                <div class="form-group">
                    <label>Stok Awal</label>
                    <input type="number" name="stok" value="0" min="0">
                </div>
                <div class="form-group">
                    <label>Stok Minimum (peringatan)</label>
                    <input type="number" name="stok_minimum" value="5" min="0">
                </div>
                <div class="form-group full">
                    <label>Deskripsi</label>
                    <textarea name="deskripsi" placeholder="Opsional"></textarea>
                </div>
            </div>
            <div class="form-actions">
                <button type="submit" class="btn btn-primary">Simpan Produk</button>
                <a href="/produk" class="btn btn-outline">Batal</a>
            </div>
        </form>
    </div>
    """
    return render_page('Tambah Produk', 'produk', content)


@app.route('/produk/edit/<int:id>', methods=['GET', 'POST'])
def produk_edit(id):
    """[PYTHON] Edit data produk yang sudah ada."""
    cur = mysql.connection.cursor()

    if request.method == 'POST':
        nama         = request.form['nama']
        kategori_id  = request.form['kategori_id'] or None
        harga        = request.form['harga']
        stok         = request.form['stok']
        stok_minimum = request.form['stok_minimum']
        satuan       = request.form['satuan']
        deskripsi    = request.form.get('deskripsi', '')
        status       = request.form['status']

        # [SQL] UPDATE data produk berdasarkan id
        # [PYTHON] Proses upload gambar baru jika ada
        gambar_baru = None
        if 'gambar' in request.files:
            file = request.files['gambar']
            if file and file.filename:
                gambar_baru = save_gambar(file)

        if gambar_baru:
            cur.execute("""
                UPDATE produk
                SET nama=%s, kategori_id=%s, harga=%s, stok=%s,
                    stok_minimum=%s, satuan=%s, deskripsi=%s, status=%s, gambar=%s
                WHERE id=%s
            """, (nama, kategori_id, harga, stok, stok_minimum, satuan, deskripsi, status, gambar_baru, id))
        else:
            cur.execute("""
                UPDATE produk
                SET nama=%s, kategori_id=%s, harga=%s, stok=%s,
                    stok_minimum=%s, satuan=%s, deskripsi=%s, status=%s
                WHERE id=%s
            """, (nama, kategori_id, harga, stok, stok_minimum, satuan, deskripsi, status, id))
        mysql.connection.commit()
        cur.close()
        flash('Produk berhasil diperbarui!', 'success')
        return redirect('/produk')

    cur.execute("SELECT * FROM produk WHERE id=%s", (id,))
    data = cur.fetchone()
    cur.execute("SELECT * FROM kategori")
    kategori = cur.fetchall()
    cur.close()

    if not data:
        flash('Produk tidak ditemukan!', 'danger')
        return redirect('/produk')

    # [PYTHON] Bangun opsi dropdown dengan 'selected' pada kategori yang dipilih
    opt_kategori = '<option value="">— Pilih —</option>'
    for k in kategori:
        sel = 'selected' if data['kategori_id'] == k['id'] else ''
        opt_kategori += f'<option value="{k["id"]}" {sel}>{k["nama"]}</option>'

    # [PYTHON] Bangun opsi satuan dengan 'selected'
    satuans = ['pcs', 'unit', 'kg', 'liter', 'box', 'lusin']
    opt_satuan = ''.join([f'<option value="{s}" {"selected" if data["satuan"]==s else ""}>{s}</option>' for s in satuans])

    opt_aktif    = 'selected' if data['status'] == 'aktif' else ''
    opt_nonaktif = 'selected' if data['status'] == 'tidak_aktif' else ''

    content = f"""
    <div class="page-header">
        <div><h2>Edit Produk</h2>
        <div class="breadcrumb"><a href="/produk">Produk</a> / Edit — {data['kode_produk']}</div></div>
    </div>
    <div class="form-card" style="max-width:680px">
        <div class="form-section-title">Informasi Produk</div>
        <!-- [HTML] enctype multipart/form-data untuk support upload gambar -->
        <form method="POST" enctype="multipart/form-data">
            <div class="form-grid">
                <!-- [HTML] Preview & Upload foto produk -->
                <div class="form-group full">
                    <label>Foto Produk</label>
                    <div style="display:flex;align-items:center;gap:16px">
                        <div id="imgPreview" style="width:100px;height:100px;border-radius:8px;
                             border:1px solid var(--border);overflow:hidden;
                             background:var(--surface-2);display:flex;align-items:center;
                             justify-content:center;font-size:32px;flex-shrink:0">
                            {'<img src="/static/uploads/produk/' + data["gambar"] + '" style="width:100%;height:100%;object-fit:cover">' if data.get("gambar") else "📦"}
                        </div>
                        <div>
                            <input type="file" name="gambar" id="inputGambar"
                                   accept="image/*" style="font-size:12px"
                                   onchange="previewEdit(this)">
                            <p style="font-size:11px;color:var(--text-3);margin-top:4px">
                                Kosongkan jika tidak ingin ganti foto
                            </p>
                        </div>
                    </div>
                </div>
                <div class="form-group full">
                    <label>Nama Produk *</label>
                    <input type="text" name="nama" value="{data['nama']}" required>
                </div>
                <div class="form-group"><label>Kategori</label><select name="kategori_id">{opt_kategori}</select></div>
                <div class="form-group"><label>Satuan</label><select name="satuan">{opt_satuan}</select></div>
                <div class="form-group">
                    <label>Harga (Rp) *</label>
                    <input type="number" name="harga" value="{data['harga']}" required min="0">
                </div>
                <div class="form-group">
                    <label>Stok</label>
                    <input type="number" name="stok" value="{data['stok']}" min="0">
                </div>
                <div class="form-group">
                    <label>Stok Minimum</label>
                    <input type="number" name="stok_minimum" value="{data['stok_minimum']}" min="0">
                </div>
                <div class="form-group">
                    <label>Status</label>
                    <select name="status">
                        <option value="aktif" {opt_aktif}>Aktif</option>
                        <option value="tidak_aktif" {opt_nonaktif}>Tidak Aktif</option>
                    </select>
                </div>
                <div class="form-group full">
                    <label>Deskripsi</label>
                    <textarea name="deskripsi">{data['deskripsi'] or ''}</textarea>
                </div>
            </div>
            <div class="form-actions">
                <button type="submit" class="btn btn-primary">Simpan</button>
                <a href="/produk" class="btn btn-outline">Batal</a>
                <a href="/katalog" class="btn btn-outline">Lihat Katalog</a>
            </div>
        </form>
    </div>
    """
    extra = '''<script>
    function previewEdit(input) {
        if (input.files && input.files[0]) {
            var reader = new FileReader();
            reader.onload = function(e) {
                document.getElementById("imgPreview").innerHTML =
                    "<img src='" + e.target.result + "' style='width:100%;height:100%;object-fit:cover'>";
            };
            reader.readAsDataURL(input.files[0]);
        }
    }
    </script>'''
    return render_page('Edit Produk', 'produk', content, extra)


@app.route('/produk/hapus/<int:id>')
def produk_hapus(id):
    """[PYTHON] Soft delete produk."""
    cur = mysql.connection.cursor()
    cur.execute("UPDATE produk SET status='tidak_aktif' WHERE id=%s", (id,))
    mysql.connection.commit()
    cur.close()
    flash('Produk berhasil dinonaktifkan!', 'warning')
    return redirect('/produk')


@app.route('/api/produk/<int:id>')
def api_produk(id):
    """
    [PYTHON] API endpoint: mengembalikan data produk dalam format JSON.
    Dipakai oleh JavaScript di halaman tambah pesanan untuk mengisi harga otomatis.
    """
    cur = mysql.connection.cursor()
    cur.execute("SELECT * FROM produk WHERE id=%s", (id,))
    data = cur.fetchone()
    cur.close()
    return jsonify(data)  # jsonify: mengubah dict Python ke respons JSON


# =============================================================================
# [PYTHON] ── BAGIAN 10: ROUTE PESANAN (CRUD)
# =============================================================================

@app.route('/pesanan')
def pesanan_index():
    """[PYTHON] Daftar pesanan dengan filter status dan pencarian."""
    cur = mysql.connection.cursor()
    search        = request.args.get('search', '')
    status_filter = request.args.get('status', '')

    # [SQL] Query dinamis: bangun kondisi WHERE berdasarkan filter yang aktif
    query  = """
        SELECT p.*, pl.nama as nama_pelanggan
        FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id = pl.id
        WHERE 1=1
    """  # WHERE 1=1: trik agar mudah menambah kondisi AND di belakangnya
    params = []

    if search:
        query += " AND (p.no_pesanan LIKE %s OR pl.nama LIKE %s)"
        params += [f'%{search}%', f'%{search}%']

    if status_filter:
        query += " AND p.status = %s"
        params.append(status_filter)

    query += " ORDER BY p.id DESC"
    cur.execute(query, params)
    data = cur.fetchall()
    cur.close()

    badge_map = {
        'pending': 'badge-gray', 'diproses': 'badge-blue',
        'dikirim': 'badge-yellow', 'selesai': 'badge-green', 'dibatalkan': 'badge-red'
    }

    rows = ''
    for p in data:
        badge = badge_map.get(p['status'], 'badge-gray')
        total = f"Rp {int(p['total_bayar']):,}".replace(',', '.')
        tgl   = p['tanggal_pesan'].strftime('%d/%m/%Y') if p['tanggal_pesan'] else '-'
        rows += f"""
        <tr>
            <td class="td-mono">{p['no_pesanan']}</td>
            <td><strong>{p['nama_pelanggan']}</strong></td>
            <td class="td-mono">{tgl}</td>
            <td class="td-mono"><strong>{total}</strong></td>
            <td class="td-mono">{p['diskon']}%</td>
            <td><span class="badge {badge}">{p['status'].upper()}</span></td>
            <td><a href="/pesanan/detail/{p['id']}" class="btn btn-outline btn-xs">Detail</a></td>
        </tr>"""

    if not rows:
        rows = '<tr><td colspan="7"><div class="empty-state"><div class="empty-icon">◎</div><p>Belum ada pesanan</p></div></td></tr>'

    # [HTML] Dropdown filter status
    status_list = ['pending', 'diproses', 'dikirim', 'selesai', 'dibatalkan']
    opt_status = '<option value="">Semua Status</option>'
    for s in status_list:
        sel = 'selected' if status_filter == s else ''
        opt_status += f'<option value="{s}" {sel}>{s.upper()}</option>'

    content = f"""
    <div class="page-header">
        <div><h2>Daftar Pesanan</h2><div class="breadcrumb">Penjualan / Pesanan</div></div>
        <a href="/pesanan/tambah" class="btn btn-primary">+ Buat Pesanan</a>
    </div>
    <div class="table-wrap">
        <div class="table-toolbar">
            <!-- [HTML] Filter status — onchange agar langsung submit saat dipilih -->
            <form method="GET" style="display:flex;gap:8px">
                <select name="status" onchange="this.form.submit()"
                        style="font-size:12px;padding:6px 10px;border:1px solid var(--border);border-radius:4px;font-family:inherit">
                    {opt_status}
                </select>
            </form>
            <form method="GET" style="display:flex;gap:8px;align-items:center">
                <input class="search-input" type="text" name="search"
                       placeholder="Cari no pesanan / pelanggan..." value="{search}">
                <button type="submit" class="btn btn-outline btn-sm">Cari</button>
                {'<a href="/pesanan" class="btn btn-outline btn-sm">Reset</a>' if search or status_filter else ''}
            </form>
        </div>
        <div style="overflow-x:auto">
            <table>
                <thead>
                    <tr>
                        <th>No. Pesanan</th><th>Pelanggan</th><th>Tgl. Pesan</th>
                        <th>Total Bayar</th><th>Diskon</th><th>Status</th><th>Aksi</th>
                    </tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
        </div>
    </div>
    """
    return render_page('Pesanan', 'pesanan', content)


@app.route('/pesanan/tambah', methods=['GET', 'POST'])
def pesanan_tambah():
    """[PYTHON] Membuat pesanan baru dengan banyak item produk secara dinamis."""
    cur = mysql.connection.cursor()

    if request.method == 'POST':
        no_pesanan   = generate_kode('ORD', 'pesanan', 'no_pesanan')
        pelanggan_id = request.form['pelanggan_id']
        tanggal      = request.form['tanggal_pesan']
        diskon       = float(request.form.get('diskon', 0))
        catatan      = request.form.get('catatan', '')

        # [PYTHON] getlist(): mengambil semua nilai field dengan nama yang sama (array dari form)
        # Dipakai untuk input dinamis produk_id[] dan jumlah[]
        produk_ids  = request.form.getlist('produk_id[]')
        jumlah_list = request.form.getlist('jumlah[]')

        total_harga = 0
        items       = []  # list untuk menyimpan detail item

        # [PYTHON] Iterasi setiap item yang dipilih di form
        for i in range(len(produk_ids)):
            if produk_ids[i]:  # skip jika produk tidak dipilih
                # [SQL] Ambil harga produk dari database
                cur.execute("SELECT harga, stok FROM produk WHERE id=%s", (produk_ids[i],))
                prod    = cur.fetchone()
                jumlah  = int(jumlah_list[i])
                harga_float = float(prod["harga"])  # konversi Decimal->float
                subtotal = harga_float * jumlah    # hitung subtotal per item
                total_harga += subtotal              # akumulasi total
                items.append((produk_ids[i], jumlah, harga_float, subtotal))
        total_bayar = total_harga * (1 - diskon / 100)  # total setelah diskon

        # [SQL] INSERT header pesanan
        cur.execute("""
            INSERT INTO pesanan (no_pesanan, pelanggan_id, tanggal_pesan, diskon, total_harga, total_bayar, catatan)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (no_pesanan, pelanggan_id, tanggal, diskon, total_harga, total_bayar, catatan))

        pesanan_id = cur.lastrowid  # lastrowid: id yang baru saja di-INSERT

        # [SQL] INSERT detail pesanan dan update stok produk
        for item in items:
            cur.execute("""
                INSERT INTO detail_pesanan (pesanan_id, produk_id, jumlah, harga_satuan, subtotal)
                VALUES (%s, %s, %s, %s, %s)
            """, (pesanan_id, item[0], item[1], item[2], item[3]))

            # [SQL] Kurangi stok produk setelah pesanan dibuat
            cur.execute("UPDATE produk SET stok = stok - %s WHERE id=%s", (item[1], item[0]))

        mysql.connection.commit()
        cur.close()
        flash(f'Pesanan {no_pesanan} berhasil dibuat!', 'success')
        return redirect('/pesanan')

    # [SQL] Data untuk dropdown pelanggan dan produk
    cur.execute("SELECT id, kode_pelanggan, nama FROM pelanggan WHERE status='aktif'")
    pelanggan = cur.fetchall()

    cur.execute("SELECT id, kode_produk, nama, harga, stok, satuan FROM produk WHERE status='aktif' AND stok > 0")
    produk = cur.fetchall()
    cur.close()

    # [PYTHON] Bangun opsi dropdown pelanggan
    opt_pelanggan = '<option value="">— Pilih Pelanggan —</option>'
    for p in pelanggan:
        opt_pelanggan += f'<option value="{p["id"]}">[{p["kode_pelanggan"]}] {p["nama"]}</option>'

    # [PYTHON] Bangun opsi dropdown produk dengan data- attribute untuk harga dan stok
    opt_produk = '<option value="">— Pilih Produk —</option>'
    for p in produk:
        harga = f"Rp {int(p['harga']):,}".replace(',', '.')
        # [HTML] data-harga dan data-stok: custom attribute untuk diakses JavaScript
        opt_produk += f'<option value="{p["id"]}" data-harga="{p["harga"]}" data-stok="{p["stok"]}">[{p["kode_produk"]}] {p["nama"]} — {harga} (Stok: {p["stok"]} {p["satuan"]})</option>'

    today = date.today().isoformat()  # format tanggal: YYYY-MM-DD

    content = f"""
    <div class="page-header">
        <div><h2>Buat Pesanan Baru</h2>
        <div class="breadcrumb"><a href="/pesanan">Pesanan</a> / Buat Baru</div></div>
    </div>

    <!-- [HTML] Template tersembunyi: opsi produk untuk di-clone saat tambah baris -->
    <template id="produk-options-template">{opt_produk}</template>

    <form method="POST">
        <div class="grid-2" style="align-items:start">

            <!-- Kolom kiri: info pesanan + ringkasan -->
            <div>
                <div class="form-card" style="margin-bottom:16px">
                    <div class="form-section-title">Informasi Pesanan</div>
                    <div class="form-grid">
                        <div class="form-group">
                            <label>Pelanggan *</label>
                            <select name="pelanggan_id" required>{opt_pelanggan}</select>
                        </div>
                        <div class="form-group">
                            <label>Tanggal Pesanan *</label>
                            <!-- [HTML] type="date" → input kalender bawaan browser -->
                            <input type="date" name="tanggal_pesan" value="{today}" required>
                        </div>
                        <div class="form-group">
                            <label>Diskon (%)</label>
                            <!-- [HTML] oninput: event saat user mengetik, langsung hitung ulang total -->
                            <input type="number" name="diskon" id="diskon" value="0"
                                   min="0" max="100" step="0.1" oninput="hitungTotal()">
                        </div>
                        <div class="form-group full">
                            <label>Catatan</label>
                            <textarea name="catatan" placeholder="Opsional" style="min-height:60px"></textarea>
                        </div>
                    </div>
                </div>

                <!-- [HTML] Kotak ringkasan total harga -->
                <div class="summary-box">
                    <div class="summary-row">
                        <span style="color:var(--text-2)">Total Harga</span>
                        <!-- [HTML] id dipakai JavaScript untuk update nilai secara realtime -->
                        <span id="total-harga-display">Rp 0</span>
                    </div>
                    <div class="summary-row">
                        <span style="color:var(--text-2)">Diskon</span>
                        <span id="diskon-nominal" style="color:var(--red)">Rp 0</span>
                    </div>
                    <div class="summary-row">
                        <span>Total Bayar</span>
                        <span id="total-bayar-display" style="color:var(--green)">Rp 0</span>
                    </div>
                </div>

                <div style="display:flex;gap:10px;margin-top:16px">
                    <button type="submit" class="btn btn-primary">✓ Buat Pesanan</button>
                    <a href="/pesanan" class="btn btn-outline">Batal</a>
                </div>
            </div>

            <!-- Kolom kanan: baris item produk dinamis -->
            <div class="form-card">
                <div class="form-section-title" style="display:flex;justify-content:space-between;align-items:center">
                    <span>Item Produk</span>
                    <!-- [HTML] onclick: memanggil fungsi JavaScript addItem() -->
                    <button type="button" class="btn btn-outline btn-xs" onclick="addItem()">+ Tambah Item</button>
                </div>

                <!-- [HTML] Container untuk baris item dinamis -->
                <div id="items-container">
                    <!-- Baris item pertama (statis) -->
                    <div class="item-row" id="item-row-1">
                        <div class="form-group">
                            <label>Produk</label>
                            <!-- [HTML] name="produk_id[]" → array, onchange → update harga -->
                            <select name="produk_id[]" onchange="updateHarga(this, 1)" required>
                                {opt_produk}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Jumlah</label>
                            <input type="number" name="jumlah[]" id="jumlah-1"
                                   value="1" min="1" oninput="hitungSubtotal(1)" required>
                        </div>
                        <div class="form-group">
                            <label>Harga</label>
                            <input type="text" id="harga-display-1" readonly placeholder="Rp 0">
                            <input type="hidden" id="harga-1" value="0">
                        </div>
                        <div class="form-group">
                            <label>Subtotal</label>
                            <input type="text" id="subtotal-display-1" readonly placeholder="Rp 0">
                        </div>
                        <div class="form-group">
                            <label>&nbsp;</label>
                            <!-- [HTML] visibility:hidden agar baris pertama tidak bisa dihapus -->
                            <button type="button" class="btn btn-danger btn-sm" style="visibility:hidden">✕</button>
                        </div>
                    </div>
                </div>

                <!-- [HTML] Informasi stok -->
                <div style="margin-top:8px;padding:10px;background:var(--surface-2);border-radius:4px;font-size:12px;color:var(--text-3)">
                    ⓘ Stok akan otomatis berkurang setelah pesanan dibuat
                </div>
            </div>

        </div>
    </form>
    """
    return render_page('Buat Pesanan', 'pesanan', content)


@app.route('/pesanan/detail/<int:id>')
def pesanan_detail(id):
    """[PYTHON] Menampilkan detail pesanan dan form update status."""
    cur = mysql.connection.cursor()

    # [SQL] JOIN 3 tabel: pesanan + pelanggan untuk data lengkap
    cur.execute("""
        SELECT p.*, pl.nama as nama_pelanggan, pl.telepon, pl.alamat, pl.kota
        FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id = pl.id
        WHERE p.id = %s
    """, (id,))
    pesanan = cur.fetchone()

    # [SQL] JOIN detail_pesanan + produk untuk item-item dalam pesanan
    cur.execute("""
        SELECT dp.*, pr.nama as nama_produk, pr.kode_produk, pr.satuan
        FROM detail_pesanan dp JOIN produk pr ON dp.produk_id = pr.id
        WHERE dp.pesanan_id = %s
    """, (id,))
    detail = cur.fetchall()
    cur.close()

    if not pesanan:
        flash('Pesanan tidak ditemukan!', 'danger')
        return redirect('/pesanan')

    badge_map = {
        'pending': 'badge-gray', 'diproses': 'badge-blue',
        'dikirim': 'badge-yellow', 'selesai': 'badge-green', 'dibatalkan': 'badge-red'
    }
    badge = badge_map.get(pesanan['status'], 'badge-gray')

    # [PYTHON] Format tanggal untuk tampilan
    tgl_pesan = pesanan['tanggal_pesan'].strftime('%d %B %Y') if pesanan['tanggal_pesan'] else '-'
    tgl_kirim = pesanan['tanggal_kirim'].strftime('%d %B %Y') if pesanan['tanggal_kirim'] else '—'
    tgl_kirim_val = pesanan['tanggal_kirim'].isoformat() if pesanan['tanggal_kirim'] else ''

    total_harga = f"Rp {int(pesanan['total_harga']):,}".replace(',', '.')
    total_bayar = f"Rp {int(pesanan['total_bayar']):,}".replace(',', '.')

    # [PYTHON] Bangun baris item detail
    item_rows = ''
    for item in detail:
        harga_sat = f"Rp {int(item['harga_satuan']):,}".replace(',', '.')
        subtotal  = f"Rp {int(item['subtotal']):,}".replace(',', '.')
        item_rows += f"""
        <tr>
            <td><strong>{item['nama_produk']}</strong>
                <div class="td-mono" style="font-size:11px">{item['kode_produk']}</div>
            </td>
            <td class="td-mono">{item['jumlah']} {item['satuan']}</td>
            <td class="td-mono">{harga_sat}</td>
            <td class="td-mono"><strong>{subtotal}</strong></td>
        </tr>"""

    # [HTML] Dropdown status untuk form update
    status_list = ['pending', 'diproses', 'dikirim', 'selesai', 'dibatalkan']
    opt_status = ''
    for s in status_list:
        sel = 'selected' if pesanan['status'] == s else ''
        opt_status += f'<option value="{s}" {sel}>{s.upper()}</option>'

    content = f"""
    <div class="page-header">
        <div><h2>Detail Pesanan</h2>
        <div class="breadcrumb"><a href="/pesanan">Pesanan</a> / {pesanan['no_pesanan']}</div></div>
        <span class="badge {badge}" style="font-size:13px;padding:5px 12px">{pesanan['status'].upper()}</span>
    </div>

    <div class="grid-2" style="align-items:start">
        <!-- Informasi pesanan dan pelanggan -->
        <div>
            <div class="card" style="margin-bottom:16px">
                <div class="card-header"><span class="card-title">Informasi Pesanan</span></div>
                <div class="card-body">
                    <div class="info-grid">
                        <div class="info-row">
                            <span class="info-label">No. Pesanan</span>
                            <span class="info-value td-mono">{pesanan['no_pesanan']}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Tanggal Pesan</span>
                            <span class="info-value">{tgl_pesan}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Tanggal Kirim</span>
                            <span class="info-value">{tgl_kirim}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Diskon</span>
                            <span class="info-value">{pesanan['diskon']}%</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Total Harga</span>
                            <span class="info-value td-mono">{total_harga}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Total Bayar</span>
                            <span class="info-value td-mono" style="font-size:16px;color:var(--green);font-weight:700">{total_bayar}</span>
                        </div>
                    </div>
                </div>
            </div>
            <div class="card">
                <div class="card-header"><span class="card-title">Pelanggan</span></div>
                <div class="card-body">
                    <div class="info-grid">
                        <div class="info-row">
                            <span class="info-label">Nama</span>
                            <span class="info-value">{pesanan['nama_pelanggan']}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">Telepon</span>
                            <span class="info-value td-mono">{pesanan['telepon'] or '-'}</span>
                        </div>
                        <div class="info-row" style="grid-column:1/-1">
                            <span class="info-label">Alamat</span>
                            <span class="info-value">{pesanan['alamat'] or '-'}, {pesanan['kota'] or ''}</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Update status + item pesanan -->
        <div>
            <div class="card" style="margin-bottom:16px">
                <div class="card-header"><span class="card-title">Update Status</span></div>
                <div class="card-body">
                    <!-- [HTML] Form update status menggunakan method POST -->
                    <form method="POST" action="/pesanan/update_status/{id}">
                        <div class="form-grid">
                            <div class="form-group">
                                <label>Status Baru</label>
                                <select name="status">{opt_status}</select>
                            </div>
                            <div class="form-group">
                                <label>Tanggal Kirim</label>
                                <input type="date" name="tanggal_kirim" value="{tgl_kirim_val}">
                            </div>
                        </div>
                        <div style="display:flex;gap:10px;margin-top:12px">
                            <button type="submit" class="btn btn-primary btn-sm">Simpan Status</button>
                            <a href="/pesanan" class="btn btn-outline btn-sm">← Kembali</a>
                            <a href="/pesanan/invoice/{id}" class="btn btn-primary btn-sm" target="_blank">
                                🖨 Cetak Invoice PDF
                            </a>
                        </div>
                    </form>
                </div>
            </div>
            <div class="card">
                <div class="card-header"><span class="card-title">Item Pesanan</span></div>
                <table>
                    <thead>
                        <tr><th>Produk</th><th>Jumlah</th><th>Harga</th><th>Subtotal</th></tr>
                    </thead>
                    <tbody>
                        {item_rows}
                        <!-- [HTML] Baris total di bagian bawah tabel -->
                        <tr style="background:var(--surface-2)">
                            <td colspan="3" style="text-align:right;font-weight:700;font-size:12px">TOTAL BAYAR</td>
                            <td class="td-mono" style="font-weight:700;color:var(--green)">{total_bayar}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    """
    return render_page('Detail Pesanan', 'pesanan', content)


@app.route('/pesanan/update_status/<int:id>', methods=['POST'])
def pesanan_update_status(id):
    """[PYTHON] Memperbarui status dan tanggal kirim pesanan."""
    status       = request.form['status']
    tanggal_kirim = request.form.get('tanggal_kirim')

    cur = mysql.connection.cursor()
    if tanggal_kirim:
        cur.execute("UPDATE pesanan SET status=%s, tanggal_kirim=%s WHERE id=%s",
                    (status, tanggal_kirim, id))
    else:
        cur.execute("UPDATE pesanan SET status=%s WHERE id=%s", (status, id))

    mysql.connection.commit()
    cur.close()
    flash('Status pesanan berhasil diperbarui!', 'success')
    return redirect(f'/pesanan/detail/{id}')


# =============================================================================
# [PYTHON] ── BAGIAN 11: ROUTE LAPORAN
# =============================================================================

@app.route('/laporan')
def laporan_index():
    """[PYTHON] Laporan penjualan bulanan: ringkasan, produk terlaris, pelanggan terbaik."""
    cur = mysql.connection.cursor()

    # [PYTHON] Ambil parameter bulan dari URL, default bulan sekarang
    bulan = request.args.get('bulan', datetime.now().strftime('%Y-%m'))
    tahun = bulan[:4]   # ambil 4 karakter pertama: tahun
    bln   = bulan[5:7]  # ambil karakter ke-5 dan 6: bulan (01-12)

    # [SQL] Ringkasan statistik bulan ini
    cur.execute("""
        SELECT COUNT(*) as total_pesanan,
               COALESCE(SUM(total_bayar), 0)  as total_pendapatan,
               COALESCE(AVG(total_bayar), 0)  as rata_rata
        FROM pesanan
        WHERE status = 'selesai'
          AND MONTH(tanggal_pesan) = %s
          AND YEAR(tanggal_pesan)  = %s
    """, (bln, tahun))
    ringkasan = cur.fetchone()

    # [SQL] 5 produk terlaris berdasarkan jumlah terjual
    # GROUP BY: mengelompokkan baris berdasarkan produk
    # ORDER BY total_terjual DESC: urutkan dari terbanyak
    cur.execute("""
        SELECT pr.nama,
               SUM(dp.jumlah)    as total_terjual,
               SUM(dp.subtotal)  as total_pendapatan
        FROM detail_pesanan dp
        JOIN produk pr   ON dp.produk_id  = pr.id
        JOIN pesanan p   ON dp.pesanan_id = p.id
        WHERE p.status = 'selesai'
          AND MONTH(p.tanggal_pesan) = %s
          AND YEAR(p.tanggal_pesan)  = %s
        GROUP BY pr.id
        ORDER BY total_terjual DESC
        LIMIT 5
    """, (bln, tahun))
    produk_terlaris = cur.fetchall()

    # [SQL] 5 pelanggan dengan total belanja terbesar
    cur.execute("""
        SELECT pl.nama, pl.kota,
               COUNT(p.id)       as total_pesanan,
               SUM(p.total_bayar) as total_belanja
        FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id = pl.id
        WHERE p.status = 'selesai'
          AND MONTH(p.tanggal_pesan) = %s
          AND YEAR(p.tanggal_pesan)  = %s
        GROUP BY pl.id
        ORDER BY total_belanja DESC
        LIMIT 5
    """, (bln, tahun))
    pelanggan_terbaik = cur.fetchall()

    # [SQL] Semua pesanan di bulan ini (semua status)
    cur.execute("""
        SELECT p.no_pesanan, pl.nama as nama_pelanggan,
               p.tanggal_pesan, p.total_bayar, p.status
        FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id = pl.id
        WHERE MONTH(p.tanggal_pesan) = %s AND YEAR(p.tanggal_pesan) = %s
        ORDER BY p.tanggal_pesan ASC
    """, (bln, tahun))
    semua_pesanan = cur.fetchall()
    cur.close()

    # [PYTHON] Format nilai Rupiah
    total_pend = f"Rp {int(ringkasan['total_pendapatan']):,}".replace(',', '.')
    rata_rata  = f"Rp {int(ringkasan['rata_rata']):,}".replace(',', '.')

    # [PYTHON] Bangun baris tabel produk terlaris
    produk_rows = ''
    for i, p in enumerate(produk_terlaris):
        pend = f"Rp {int(p['total_pendapatan']):,}".replace(',', '.')
        produk_rows += f"""
        <tr>
            <td style="font-weight:700;color:var(--text-3)">{i+1}</td>
            <td><strong>{p['nama']}</strong></td>
            <td class="td-mono">{p['total_terjual']}</td>
            <td class="td-mono">{pend}</td>
        </tr>"""
    if not produk_rows:
        produk_rows = '<tr><td colspan="4"><div class="empty-state"><p>Belum ada data</p></div></td></tr>'

    # [PYTHON] Bangun baris tabel pelanggan terbaik
    pelanggan_rows = ''
    for i, p in enumerate(pelanggan_terbaik):
        belanja = f"Rp {int(p['total_belanja']):,}".replace(',', '.')
        pelanggan_rows += f"""
        <tr>
            <td style="font-weight:700;color:var(--text-3)">{i+1}</td>
            <td><strong>{p['nama']}</strong><div style="font-size:11px;color:var(--text-3)">{p['kota']}</div></td>
            <td class="td-mono">{p['total_pesanan']}x</td>
            <td class="td-mono">{belanja}</td>
        </tr>"""
    if not pelanggan_rows:
        pelanggan_rows = '<tr><td colspan="4"><div class="empty-state"><p>Belum ada data</p></div></td></tr>'

    # [PYTHON] Bangun baris semua transaksi
    badge_map = {
        'pending': 'badge-gray', 'diproses': 'badge-blue',
        'dikirim': 'badge-yellow', 'selesai': 'badge-green', 'dibatalkan': 'badge-red'
    }
    transaksi_rows = ''
    for p in semua_pesanan:
        badge = badge_map.get(p['status'], 'badge-gray')
        total = f"Rp {int(p['total_bayar']):,}".replace(',', '.')
        tgl   = p['tanggal_pesan'].strftime('%d/%m/%Y') if p['tanggal_pesan'] else '-'
        transaksi_rows += f"""
        <tr>
            <td class="td-mono">{p['no_pesanan']}</td>
            <td>{p['nama_pelanggan']}</td>
            <td class="td-mono">{tgl}</td>
            <td class="td-mono"><strong>{total}</strong></td>
            <td><span class="badge {badge}">{p['status'].upper()}</span></td>
        </tr>"""
    if not transaksi_rows:
        transaksi_rows = '<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">◧</div><p>Tidak ada transaksi periode ini</p></div></td></tr>'

    content = f"""
    <div class="page-header">
        <div><h2>Laporan Bulanan</h2><div class="breadcrumb">Analitik / Laporan</div></div>
        <!-- [HTML] Form filter periode bulan -->
        <!-- [HTML] action="/laporan" → form submit ke URL yang benar -->
        <form action="/laporan" method="GET" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
            <label style="font-size:12px;color:var(--text-2);font-weight:600;letter-spacing:0.06em">PERIODE</label>
            <!-- [HTML] type="month" → input kalender bulan-tahun bawaan browser -->
            <input type="month" name="bulan" value="{bulan}"
                   style="font-family:inherit;font-size:13px;padding:7px 10px;
                          border:1px solid var(--border);border-radius:4px;outline:none">
            <button type="submit" class="btn btn-primary btn-sm">Tampilkan</button>
            <!-- [HTML] Tombol export — href pakai f-string {bulan} dari Python -->
            <a href="/laporan/export?bulan={bulan}"
               style="display:inline-flex;align-items:center;gap:5px;padding:5px 12px;
                      background:#2d6a4f;color:#fff;border-radius:4px;text-decoration:none;
                      font-size:12px;font-weight:600">
                &#128202; Export Excel (Bulan Ini)
            </a>
            <a href="/laporan/export"
               style="display:inline-flex;align-items:center;gap:5px;padding:5px 12px;
                      background:#1a4a8a;color:#fff;border-radius:4px;text-decoration:none;
                      font-size:12px;font-weight:600">
                &#128190; Export Semua Data
            </a>
        </form>
    </div>

    <!-- [HTML] Statistik ringkasan bulan ini -->
    <div class="stats-grid" style="margin-bottom:24px">
        <div class="stat-card green">
            <div class="stat-label">Pesanan Selesai</div>
            <div class="stat-value">{ringkasan['total_pesanan']}</div>
            <div class="stat-sub">order berhasil</div>
        </div>
        <div class="stat-card yellow">
            <div class="stat-label">Total Pendapatan</div>
            <div class="stat-value" style="font-size:15px">{total_pend}</div>
        </div>
        <div class="stat-card blue">
            <div class="stat-label">Rata-rata per Order</div>
            <div class="stat-value" style="font-size:15px">{rata_rata}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Periode</div>
            <div class="stat-value" style="font-size:18px;font-family:'IBM Plex Mono',monospace">{bulan}</div>
        </div>
    </div>

    <!-- [HTML] Tabel produk terlaris dan pelanggan terbaik berdampingan -->
    <div class="grid-2">
        <div class="card">
            <div class="card-header"><span class="card-title">🏆 Produk Terlaris</span></div>
            <table>
                <thead><tr><th>#</th><th>Produk</th><th>Terjual</th><th>Pendapatan</th></tr></thead>
                <tbody>{produk_rows}</tbody>
            </table>
        </div>
        <div class="card">
            <div class="card-header"><span class="card-title">⭐ Pelanggan Terbaik</span></div>
            <table>
                <thead><tr><th>#</th><th>Pelanggan</th><th>Order</th><th>Total Belanja</th></tr></thead>
                <tbody>{pelanggan_rows}</tbody>
            </table>
        </div>
    </div>

    <!-- [HTML] Tabel semua transaksi bulan ini -->
    <div class="table-wrap" style="margin-top:20px">
        <div class="table-toolbar">
            <span style="font-size:13px;font-weight:700">Semua Transaksi — {bulan}</span>
            <span class="badge badge-gray">{len(semua_pesanan)} transaksi</span>
        </div>
        <div style="overflow-x:auto">
            <table>
                <thead>
                    <tr><th>No. Pesanan</th><th>Pelanggan</th><th>Tanggal</th><th>Total Bayar</th><th>Status</th></tr>
                </thead>
                <tbody>{transaksi_rows}</tbody>
            </table>
        </div>
    </div>
    """
    return render_page('Laporan', 'laporan', content)


# =============================================================================
# [PYTHON] ── BAGIAN 12: MENJALANKAN APLIKASI
# =============================================================================


# =============================================================================
# [PYTHON] -- BAGIAN 12B: KONFIGURASI EMAIL GMAIL
# =============================================================================
EMAIL_SENDER   = 'emailkamu@gmail.com'
EMAIL_PASSWORD = 'xxxx xxxx xxxx xxxx'
EMAIL_ADMIN    = 'emailkamu@gmail.com'

def kirim_email_stok(produk_menipis):
    if not produk_menipis: return
    try:
        baris = "".join(f'<tr><td style="padding:8px">{p["nama"]}</td><td style="padding:8px;color:red">{p["stok"]} {p["satuan"]}</td><td style="padding:8px">{p["stok_minimum"]}</td></tr>' for p in produk_menipis)
        isi = f"""<html><body style="font-family:Arial"><div style="background:#052e16;padding:20px;text-align:center"><h2 style="color:#e8c547">PendidikanStore - Peringatan Stok!</h2></div><div style="padding:20px"><table style="width:100%;border-collapse:collapse"><tr style="background:#052e16;color:#fff"><th style="padding:10px">Produk</th><th>Stok</th><th>Minimum</th></tr>{baris}</table><p>Segera restok!</p></div></body></html>"""
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"Peringatan: {len(produk_menipis)} Produk Stok Menipis - PendidikanStore"
        msg["From"] = EMAIL_SENDER
        msg["To"]   = EMAIL_ADMIN
        msg.attach(MIMEText(isi, "html"))
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_ADMIN, msg.as_string())
        print(f"[EMAIL OK] Terkirim ke {EMAIL_ADMIN}")
    except Exception as e:
        print(f"[EMAIL ERROR] {e}")


# =============================================================================
# [PYTHON] -- BAGIAN 13: CETAK INVOICE PDF
# =============================================================================
@app.route("/pesanan/invoice/<int:id>")
def cetak_invoice(id):
    cur = mysql.connection.cursor()
    cur.execute("""SELECT p.*, pl.nama as nama_pelanggan, pl.telepon, pl.alamat, pl.kota
        FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id=pl.id WHERE p.id=%s""", (id,))
    pesanan = cur.fetchone()
    cur.execute("""SELECT dp.*, pr.nama as nama_produk, pr.kode_produk, pr.satuan
        FROM detail_pesanan dp JOIN produk pr ON dp.produk_id=pr.id
        WHERE dp.pesanan_id=%s""", (id,))
    detail = cur.fetchall()
    cur.close()
    if not pesanan:
        flash("Pesanan tidak ditemukan!", "danger")
        return redirect("/pesanan")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    def sty(**kw): return ParagraphStyle("x", parent=styles["Normal"], **kw)
    el = []

    hd = [[Paragraph("<b>PendidikanStore</b>", sty(fontSize=20, fontName="Helvetica-Bold")),
           Paragraph(f"<b>INVOICE #{pesanan['no_pesanan']}</b>", sty(fontSize=14, fontName="Helvetica-Bold", alignment=TA_RIGHT))],
          [Paragraph("Alat Tulis & Perlengkapan Sekolah - APSI Teknik Industri", sty(fontSize=9, textColor=colors.grey)),
           Paragraph(f"{pesanan['tanggal_pesan'].strftime('%d %B %Y') if pesanan['tanggal_pesan'] else '-'}", sty(fontSize=9, alignment=TA_RIGHT, textColor=colors.grey))]]
    th = Table(hd, colWidths=[95*mm, 75*mm])
    th.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    el.extend([th, HRFlowable(width="100%", thickness=2, color=colors.HexColor("#052e16"), spaceAfter=5*mm)])

    info = [[Paragraph("<b>KEPADA:</b>", sty(fontSize=9, fontName="Helvetica-Bold")), "",
             Paragraph("<b>STATUS:</b>", sty(fontSize=9, fontName="Helvetica-Bold"))],
            [Paragraph(f"<b>{pesanan['nama_pelanggan']}</b>", sty(fontSize=11, fontName="Helvetica-Bold")), "",
             Paragraph(f"<b>{pesanan['status'].upper()}</b>", sty(fontSize=10, fontName="Helvetica-Bold"))],
            [Paragraph(pesanan["alamat"] or "-", sty(fontSize=9)), "",
             Paragraph(f"Tgl. Kirim: {pesanan['tanggal_kirim'].strftime('%d %B %Y') if pesanan['tanggal_kirim'] else 'Belum dikirim'}", sty(fontSize=9))],
            [Paragraph(f"{pesanan['kota'] or ''} | {pesanan['telepon'] or ''}", sty(fontSize=9)), "", Paragraph("", sty(fontSize=9))]]
    ti = Table(info, colWidths=[88*mm, 8*mm, 74*mm])
    ti.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#f9f8f6")),("BACKGROUND",(2,0),(2,-1),colors.HexColor("#f9f8f6")),("LEFTPADDING",(0,0),(0,-1),8),("LEFTPADDING",(2,0),(2,-1),8),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
    el.extend([ti, Spacer(1, 5*mm)])

    rows = [["No","Kode","Nama Produk","Sat.","Qty","Harga","Subtotal"]]
    for i, item in enumerate(detail):
        rows.append([str(i+1), item["kode_produk"], item["nama_produk"], item["satuan"], str(item["jumlah"]),
                     f"Rp {int(item['harga_satuan']):,}".replace(",","."),
                     f"Rp {int(item['subtotal']):,}".replace(",",".")])
    tbl = Table(rows, colWidths=[8*mm,20*mm,57*mm,12*mm,10*mm,28*mm,30*mm])
    tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#052e16")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(4,1),(-1,-1),"RIGHT"),("ALIGN",(0,0),(3,-1),"CENTER"),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#bbf7d0")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f9f8f6")])]))
    el.extend([tbl, Spacer(1, 4*mm)])

    dn = float(pesanan["total_harga"]) * float(pesanan["diskon"]) / 100
    tot = [["","Subtotal:", f"Rp {int(pesanan['total_harga']):,}".replace(",",".")],
           ["",f"Diskon ({pesanan['diskon']}%):", f"- Rp {int(dn):,}".replace(",",".")],
           ["","TOTAL BAYAR:", f"Rp {int(pesanan['total_bayar']):,}".replace(",",".")]]
    tt = Table(tot, colWidths=[100*mm, 40*mm, 30*mm])
    tt.setStyle(TableStyle([("ALIGN",(1,0),(-1,-1),"RIGHT"),("FONTSIZE",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LINEABOVE",(1,2),(-1,2),1.5,colors.HexColor("#052e16")),("FONTNAME",(1,2),(-1,2),"Helvetica-Bold"),("FONTSIZE",(1,2),(-1,2),11),("TEXTCOLOR",(2,2),(2,2),colors.HexColor("#2d6a4f"))]))
    el.extend([tt, Spacer(1,6*mm), HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bbf7d0"), spaceAfter=3*mm),
               Paragraph("Terima kasih atas kepercayaan Anda! -- PendidikanStore APSI Teknik Industri", sty(fontSize=8, textColor=colors.grey, alignment=TA_CENTER))])
    doc.build(el)
    buf.seek(0)
    return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name=f"Invoice_{pesanan['no_pesanan']}.pdf")


# =============================================================================
# [PYTHON] -- BAGIAN 14: EXPORT LAPORAN EXCEL (3 Sheet)
# =============================================================================
@app.route("/laporan/export")
def export_excel():
    cur = mysql.connection.cursor()
    bulan = request.args.get("bulan", "")
    tahun = bulan[:4] if bulan else ""
    bln   = bulan[5:7] if bulan else ""

    q_p = """SELECT p.no_pesanan, pl.nama as pelanggan, pl.kota, p.tanggal_pesan, p.tanggal_kirim,
               p.status, p.total_harga, p.diskon, p.total_bayar
            FROM pesanan p JOIN pelanggan pl ON p.pelanggan_id=pl.id"""
    q_s = """SELECT p.no_pesanan, p.tanggal_pesan, pl.nama as pelanggan, pr.kode_produk,
               pr.nama as produk, pr.satuan, dp.jumlah, dp.harga_satuan, dp.subtotal, p.status
            FROM detail_pesanan dp JOIN pesanan p ON dp.pesanan_id=p.id
            JOIN pelanggan pl ON p.pelanggan_id=pl.id JOIN produk pr ON dp.produk_id=pr.id"""

    if bulan:
        cur.execute(q_p + " WHERE MONTH(p.tanggal_pesan)=%s AND YEAR(p.tanggal_pesan)=%s ORDER BY p.tanggal_pesan", (bln,tahun))
    else:
        cur.execute(q_p + " ORDER BY p.tanggal_pesan ASC")
    dp = cur.fetchall()

    if bulan:
        cur.execute(q_s + " WHERE MONTH(p.tanggal_pesan)=%s AND YEAR(p.tanggal_pesan)=%s ORDER BY p.tanggal_pesan", (bln,tahun))
    else:
        cur.execute(q_s + " ORDER BY p.tanggal_pesan ASC")
    ds = cur.fetchall()

    cur.execute("""SELECT pr.kode_produk, pr.nama, k.nama as kategori, pr.harga, pr.stok,
               pr.stok_minimum, pr.satuan, (pr.stok*pr.harga) as nilai_stok
            FROM produk pr LEFT JOIN kategori k ON pr.kategori_id=k.id ORDER BY k.nama, pr.nama""")
    dr = cur.fetchall()
    cur.close()

    wb = Workbook()
    CH="FF1A1916"; CG="FF2D6A4F"; CR="FFC0392B"; CST="FFF9F8F6"; CBR="FFE2DFD8"; FMT="#,##0"

    def bd():
        s=Side(style="thin",color=CBR); return Border(left=s,right=s,top=s,bottom=s)
    def sh(cell,val):
        cell.value=val; cell.font=Font(name="Arial",bold=True,color="FFFFFFFF",size=10)
        cell.fill=PatternFill("solid",fgColor=CH); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=bd()
    def sd(cell,val,al="left",bold=False,bg=None,col=None):
        cell.value=val; cell.font=Font(name="Arial",bold=bold,size=9,color=col if col else "FF1A1916")
        cell.alignment=Alignment(horizontal=al,vertical="center"); cell.border=bd()
        if bg: cell.fill=PatternFill("solid",fgColor=bg)
    def st(cell,val):
        cell.value=val; cell.font=Font(name="Arial",bold=True,size=10,color="FFFFFFFF")
        cell.fill=PatternFill("solid",fgColor=CG); cell.alignment=Alignment(horizontal="right",vertical="center"); cell.border=bd(); cell.number_format=FMT
    def judul(ws,title,period,nc):
        lc=get_column_letter(nc); ws.merge_cells(f"A1:{lc}1"); ws["A1"].value=title
        ws["A1"].font=Font(name="Arial",bold=True,size=16,color=CH); ws["A1"].alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[1].height=36
        ws.merge_cells(f"A2:{lc}2"); ws["A2"].value=f"Periode: {period}  |  Diekspor: {datetime.now().strftime('%d %B %Y %H:%M')}"
        ws["A2"].font=Font(name="Arial",size=9,italic=True,color="FF928E89"); ws["A2"].alignment=Alignment(horizontal="center"); ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=6

    wst={"selesai":CG,"dibatalkan":CR,"dikirim":"FFE76F00","diproses":"FF1A4A8A","pending":"FF928E89"}
    per = bulan if bulan else "Semua Data"

    # Sheet 1
    ws1=wb.active; ws1.title="Laporan Penjualan"
    judul(ws1,"LAPORAN PENJUALAN",per,10)
    for c,h in enumerate(["No","No. Pesanan","Pelanggan","Kota","Tgl. Pesan","Tgl. Kirim","Status","Total Harga (Rp)","Diskon (%)","Total Bayar (Rp)"],1): sh(ws1.cell(4,c),h)
    ws1.row_dimensions[4].height=26
    for i,p in enumerate(dp):
        r=i+5; bg=None if i%2==0 else CST
        tp=p["tanggal_pesan"].strftime("%d/%m/%Y") if p["tanggal_pesan"] else "-"
        tk=p["tanggal_kirim"].strftime("%d/%m/%Y") if p["tanggal_kirim"] else "-"
        sd(ws1.cell(r,1),i+1,"center",bg=bg); sd(ws1.cell(r,2),p["no_pesanan"],"center",True,bg); sd(ws1.cell(r,3),p["pelanggan"],"left",bg=bg); sd(ws1.cell(r,4),p["kota"] or "-","left",bg=bg)
        sd(ws1.cell(r,5),tp,"center",bg=bg); sd(ws1.cell(r,6),tk,"center",bg=bg); sd(ws1.cell(r,7),p["status"].upper(),"center",True,bg,wst.get(p["status"],CH))
        sd(ws1.cell(r,8),float(p["total_harga"]),"right",bg=bg); sd(ws1.cell(r,9),float(p["diskon"]),"right",bg=bg); sd(ws1.cell(r,10),float(p["total_bayar"]),"right",True,bg,CG); ws1.row_dimensions[r].height=18
    rt=len(dp)+5; lr=rt-1
    ws1.merge_cells(f"A{rt}:G{rt}"); cc=ws1[f"A{rt}"]; cc.value=f"TOTAL ({len(dp)} TRANSAKSI)"; cc.font=Font(name="Arial",bold=True,size=10,color="FFFFFFFF"); cc.fill=PatternFill("solid",fgColor=CH); cc.alignment=Alignment(horizontal="right",vertical="center"); cc.border=bd()
    st(ws1.cell(rt,8),f"=SUM(H5:H{lr})"); st(ws1.cell(rt,9),f"=AVERAGE(I5:I{lr})"); st(ws1.cell(rt,10),f"=SUM(J5:J{lr})"); ws1.row_dimensions[rt].height=24
    for row in ws1.iter_rows(5,rt,8,10):
        for cell in row: cell.number_format=FMT
    for c,w in enumerate([5,15,28,15,12,12,12,18,11,18],1): ws1.column_dimensions[get_column_letter(c)].width=w
    ws1.freeze_panes="A5"

    # Sheet 2
    ws2=wb.create_sheet("Keluar Masuk Stok"); judul(ws2,"LAPORAN KELUAR MASUK STOK",per,10)
    for c,h in enumerate(["No","Tanggal","No. Pesanan","Pelanggan","Kode Produk","Nama Produk","Satuan","Qty Keluar","Harga Satuan (Rp)","Subtotal (Rp)"],1): sh(ws2.cell(4,c),h)
    ws2.row_dimensions[4].height=26
    for i,item in enumerate(ds):
        r=i+5; bg=None if i%2==0 else CST
        tg=item["tanggal_pesan"].strftime("%d/%m/%Y") if item["tanggal_pesan"] else "-"
        sd(ws2.cell(r,1),i+1,"center",bg=bg); sd(ws2.cell(r,2),tg,"center",bg=bg); sd(ws2.cell(r,3),item["no_pesanan"],"center",True,bg); sd(ws2.cell(r,4),item["pelanggan"],"left",bg=bg)
        sd(ws2.cell(r,5),item["kode_produk"],"center",bg=bg); sd(ws2.cell(r,6),item["produk"],"left",bg=bg); sd(ws2.cell(r,7),item["satuan"],"center",bg=bg)
        sd(ws2.cell(r,8),item["jumlah"],"right",True,bg,CR); sd(ws2.cell(r,9),float(item["harga_satuan"]),"right",bg=bg); sd(ws2.cell(r,10),float(item["subtotal"]),"right",True,bg); ws2.row_dimensions[r].height=18
    rt2=len(ds)+5; lr2=rt2-1
    ws2.merge_cells(f"A{rt2}:G{rt2}"); cc2=ws2[f"A{rt2}"]; cc2.value=f"TOTAL ({len(ds)} ITEM)"; cc2.font=Font(name="Arial",bold=True,size=10,color="FFFFFFFF"); cc2.fill=PatternFill("solid",fgColor=CH); cc2.alignment=Alignment(horizontal="right",vertical="center"); cc2.border=bd()
    st(ws2.cell(rt2,8),f"=SUM(H5:H{lr2})"); st(ws2.cell(rt2,9),f"=AVERAGE(I5:I{lr2})"); st(ws2.cell(rt2,10),f"=SUM(J5:J{lr2})"); ws2.row_dimensions[rt2].height=24
    for row in ws2.iter_rows(5,rt2,9,10):
        for cell in row: cell.number_format=FMT
    for c,w in enumerate([5,12,15,28,14,28,9,11,18,16],1): ws2.column_dimensions[get_column_letter(c)].width=w
    ws2.freeze_panes="A5"

    # Sheet 3
    ws3=wb.create_sheet("Ringkasan Stok"); judul(ws3,"RINGKASAN STOK PRODUK",f"Per {datetime.now().strftime('%d %B %Y')}",9)
    for c,h in enumerate(["No","Kode","Nama Produk","Kategori","Satuan","Harga (Rp)","Stok","Min. Stok","Nilai Stok (Rp)"],1): sh(ws3.cell(4,c),h)
    ws3.row_dimensions[4].height=26
    for i,p in enumerate(dr):
        r=i+5; kritis=p["stok"]<=p["stok_minimum"]; bg="FFFFF3CD" if kritis else (None if i%2==0 else CST)
        sd(ws3.cell(r,1),i+1,"center",bg=bg); sd(ws3.cell(r,2),p["kode_produk"],"center",True,bg); sd(ws3.cell(r,3),p["nama"],"left",bg=bg); sd(ws3.cell(r,4),p["kategori"] or "-","left",bg=bg)
        sd(ws3.cell(r,5),p["satuan"],"center",bg=bg); sd(ws3.cell(r,6),float(p["harga"]),"right",bg=bg); sd(ws3.cell(r,7),p["stok"],"right",True,bg,CR if kritis else CG); sd(ws3.cell(r,8),p["stok_minimum"],"right",bg=bg); sd(ws3.cell(r,9),float(p["nilai_stok"]),"right",True,bg); ws3.row_dimensions[r].height=18
    rt3=len(dr)+5; lr3=rt3-1
    ws3.merge_cells(f"A{rt3}:H{rt3}"); cc3=ws3[f"A{rt3}"]; cc3.value=f"TOTAL NILAI STOK ({len(dr)} PRODUK)"; cc3.font=Font(name="Arial",bold=True,size=10,color="FFFFFFFF"); cc3.fill=PatternFill("solid",fgColor=CH); cc3.alignment=Alignment(horizontal="right",vertical="center"); cc3.border=bd()
    st(ws3.cell(rt3,9),f"=SUM(I5:I{lr3})"); ws3.row_dimensions[rt3].height=24
    ket=rt3+2; ws3.merge_cells(f"A{ket}:F{ket}"); k=ws3[f"A{ket}"]; k.value="Keterangan: Baris KUNING = stok di bawah minimum, perlu restok!"; k.font=Font(name="Arial",size=8,italic=True,color="FF7A4500"); k.fill=PatternFill("solid",fgColor="FFFFF3CD")
    for row in ws3.iter_rows(5,rt3,6,9):
        for cell in row:
            if cell.column in [6,9]: cell.number_format=FMT
    for c,w in enumerate([5,14,30,16,9,16,9,11,20],1): ws3.column_dimensions[get_column_letter(c)].width=w
    ws3.freeze_panes="A5"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    nama=f"Laporan_PendidikanStore_{bulan if bulan else 'SemuaData'}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=nama)


# =============================================================================
# [PYTHON] -- BAGIAN 15: NOTIFIKASI EMAIL STOK
# =============================================================================
@app.route("/notifikasi/stok")
def notifikasi_stok():
    cur = mysql.connection.cursor()
    cur.execute("SELECT nama, stok, stok_minimum, satuan FROM produk WHERE stok<=stok_minimum AND status='aktif' ORDER BY stok ASC")
    pm = cur.fetchall()
    cur.close()
    if pm:
        kirim_email_stok(pm)
        flash(f"Email terkirim! {len(pm)} produk stok menipis.", "warning")
    else:
        flash("Semua stok aman!", "success")
    return redirect("/produk")



# =============================================================================
# [SQL] -- TABEL USERS (jalankan di phpMyAdmin dulu!)
# =============================================================================
# ALTER DATABASE sistem_penjualan CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
#
# CREATE TABLE IF NOT EXISTS users (
#     id           INT AUTO_INCREMENT PRIMARY KEY,
#     nama         VARCHAR(100) NOT NULL,
#     email        VARCHAR(100) UNIQUE NOT NULL,
#     password     VARCHAR(255) NOT NULL,       -- password terenkripsi (hash)
#     role         ENUM("admin","staff") DEFAULT "staff",
#     avatar       VARCHAR(10) DEFAULT "A",     -- inisial untuk avatar
#     provider     VARCHAR(20) DEFAULT "email", -- "email" atau "google"
#     status       ENUM("aktif","nonaktif") DEFAULT "aktif",
#     last_login   TIMESTAMP NULL,
#     created_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
# );
#
# -- Akun admin default (password: admin123)
# INSERT INTO users (nama, email, password, role, avatar) VALUES
# ("Administrator", "admin@pendidikanstore.com",
#  "pbkdf2:sha256:600000$saltkey$hashedpassword", "admin", "A");
# -- PENTING: Jalankan route /setup-admin untuk buat akun admin dengan hash yang benar!
# =============================================================================


# =============================================================================
# [PYTHON] -- LOGIN SYSTEM: Decorator & Helper
# =============================================================================

# [CSS] Styling khusus halaman login — terpisah dari CSS_GLOBAL
CSS_LOGIN = """
/* [CSS] Halaman login: full screen split layout */
* { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: "IBM Plex Sans", sans-serif;
    background: #052e16;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
}

/* [CSS] Container utama: dua kolom */
.login-wrapper {
    display: flex;
    width: 920px;
    min-height: 560px;
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 32px 80px rgba(0,0,0,0.5);
}

/* [CSS] Panel kiri — branding & ilustrasi */
.login-left {
    flex: 1;
    background: #14532d;
    padding: 48px 40px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    position: relative;
    overflow: hidden;
}

/* [CSS] Dekorasi background geometris */
.login-left::before {
    content: "";
    position: absolute;
    top: -80px; right: -80px;
    width: 320px; height: 320px;
    border-radius: 50%;
    background: rgba(74,222,128,0.08);
}
.login-left::after {
    content: "";
    position: absolute;
    bottom: -60px; left: -60px;
    width: 240px; height: 240px;
    border-radius: 50%;
    background: rgba(74,222,128,0.05);
}

/* [CSS] Logo aplikasi */
.login-logo {
    display: flex;
    align-items: center;
    gap: 12px;
    position: relative;
    z-index: 1;
}
.login-logo-icon {
    width: 44px; height: 44px;
    background: #4ade80;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 22px;
    color: #052e16;
    font-weight: 900;
}
.login-logo-text { color: #fff; }
.login-logo-name { font-size: 18px; font-weight: 700; letter-spacing: -0.02em; }
.login-logo-sub  { font-size: 11px; color: rgba(255,255,255,0.4); letter-spacing: 0.08em; text-transform: uppercase; }

/* [CSS] Teks promo di panel kiri */
.login-hero { position: relative; z-index: 1; }
.login-hero h2 {
    font-size: 28px; font-weight: 800;
    color: #fff; line-height: 1.3;
    letter-spacing: -0.03em;
    margin-bottom: 12px;
}
.login-hero h2 span { color: #4ade80; }
.login-hero p { font-size: 13px; color: rgba(255,255,255,0.45); line-height: 1.7; }

/* [CSS] Fitur list di bawah */
.login-features { position: relative; z-index: 1; }
.login-feature {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
    font-size: 12px;
    color: rgba(255,255,255,0.5);
}
.login-feature-dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: #4ade80;
    flex-shrink: 0;
}

/* [CSS] Panel kanan — form login */
.login-right {
    width: 400px;
    background: #f0fdf4;
    padding: 48px 40px;
    display: flex;
    flex-direction: column;
    justify-content: center;
}

.login-title    { font-size: 22px; font-weight: 800; color: #052e16; margin-bottom: 4px; letter-spacing: -0.02em; }
.login-subtitle { font-size: 13px; color: #4ade80; margin-bottom: 32px; }

/* [CSS] Tombol Login dengan Google */
.btn-google {
    display: flex; align-items: center; justify-content: center;
    gap: 10px;
    width: 100%; padding: 11px;
    background: #fff;
    border: 1.5px solid #bbf7d0;
    border-radius: 8px;
    font-size: 13px; font-weight: 600;
    color: #052e16; cursor: pointer;
    text-decoration: none;
    transition: all 0.15s;
    margin-bottom: 20px;
}
.btn-google:hover { background: #f9f8f6; border-color: #86efac; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
.btn-google svg  { flex-shrink: 0; }

/* [CSS] Divider "atau" */
.login-divider {
    display: flex; align-items: center;
    gap: 12px; margin-bottom: 20px;
}
.login-divider-line { flex: 1; height: 1px; background: #bbf7d0; }
.login-divider-text { font-size: 11px; color: #4ade80; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; }

/* [CSS] Form group */
.form-group-login { margin-bottom: 16px; }
.form-group-login label {
    display: block; font-size: 11px; font-weight: 700;
    letter-spacing: 0.08em; text-transform: uppercase;
    color: #166534; margin-bottom: 6px;
}
.form-group-login input {
    width: 100%; padding: 11px 14px;
    border: 1.5px solid #bbf7d0;
    border-radius: 8px;
    font-size: 13px; color: #052e16;
    background: #fff;
    font-family: "IBM Plex Sans", sans-serif;
    outline: none; transition: all 0.15s;
}
.form-group-login input:focus {
    border-color: #052e16;
    box-shadow: 0 0 0 3px rgba(26,25,22,0.08);
}

/* [CSS] Row ingat saya + lupa password */
.login-options {
    display: flex; align-items: center;
    justify-content: space-between;
    margin-bottom: 20px; font-size: 12px;
}
.login-options label { display: flex; align-items: center; gap: 6px; color: #166534; cursor: pointer; font-size: 12px; letter-spacing: 0; text-transform: none; }
.login-options a     { color: #052e16; font-weight: 600; text-decoration: none; }
.login-options a:hover { text-decoration: underline; }

/* [CSS] Tombol submit login */
.btn-login {
    width: 100%; padding: 12px;
    background: #14532d; color: #fff;
    border: none; border-radius: 8px;
    font-size: 14px; font-weight: 700;
    cursor: pointer; font-family: "IBM Plex Sans", sans-serif;
    transition: all 0.15s; letter-spacing: 0.01em;
}
.btn-login:hover { background: #15803d; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
.btn-login:active { transform: translateY(0); }

/* [CSS] Link daftar akun baru */
.login-register { text-align: center; margin-top: 20px; font-size: 12px; color: #4ade80; }
.login-register a { color: #052e16; font-weight: 700; text-decoration: none; }
.login-register a:hover { text-decoration: underline; }

/* [CSS] Flash message di halaman login */
.login-flash {
    padding: 10px 14px; border-radius: 8px;
    margin-bottom: 16px; font-size: 12px; font-weight: 600;
}
.login-flash.danger  { background: #fde8e8; color: #c0392b; border-left: 3px solid #c0392b; }
.login-flash.warning { background: #fff3cd; color: #e76f00; border-left: 3px solid #e76f00; }
.login-flash.success { background: #d8f3dc; color: #2d6a4f; border-left: 3px solid #2d6a4f; }

/* [CSS] Animasi fade in saat halaman dimuat */
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(20px); }
    to   { opacity: 1; transform: translateY(0); }
}
.login-right { animation: fadeUp 0.5s ease forwards; }

/* [CSS] Responsive mobile */
@media (max-width: 768px) {
    .login-wrapper { flex-direction: column; width: 95%; min-height: auto; }
    .login-left    { padding: 32px 28px; min-height: auto; }
    .login-right   { width: 100%; padding: 32px 28px; }
    .login-hero h2 { font-size: 22px; }
    .login-features { display: none; }
}
"""

# [HTML] Template halaman login
LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login — PendidikanStore APSI</title>
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>{{ css|safe }}</style>
</head>
<body>

<div class="login-wrapper">

    <!-- [HTML] Panel Kiri: Branding -->
    <div class="login-left">
        <!-- Logo -->
        <div class="login-logo">
            <div class="login-logo-icon">S</div>
            <div class="login-logo-text">
                <div class="login-logo-name">PendidikanStore</div>
                <div class="login-logo-sub">Perlengkapan Siswa & Mahasiswa</div>
            </div>
        </div>

        <!-- Hero text -->
        <div class="login-hero">
            <h2>Kelola Toko Pendidikan<br>dengan <span>Lebih Cerdas</span></h2>
            <p>Toko alat tulis dan perlengkapan pendidikan lengkap untuk TK, SD, SMP, SMA/SMK, hingga mahasiswa.</p>
        </div>

        <!-- Fitur list -->
        <div class="login-features">
            <div class="login-feature">
                <div class="login-feature-dot"></div>
                Pensil, pena, buku & alat tulis lengkap
            </div>
            <div class="login-feature">
                <div class="login-feature-dot"></div>
                Produk untuk semua jenjang pendidikan
            </div>
            <div class="login-feature">
                <div class="login-feature-dot"></div>
                Laporan penjualan & stok otomatis
            </div>
            <div class="login-feature">
                <div class="login-feature-dot"></div>
                Notifikasi stok & pesanan real-time
            </div>
        </div>
    </div>

    <!-- [HTML] Panel Kanan: Form Login -->
    <div class="login-right">
        <h1 class="login-title">Selamat datang</h1>
        <p class="login-subtitle">Masuk ke akun PendidikanStore kamu</p>

        <!-- [HTML] Flash messages -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="login-flash {{ category }}">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <!-- [HTML] Tombol Login dengan Google -->
        <a href="/login/google" class="btn-google">
            <!-- [HTML] Logo Google SVG -->
            <svg width="18" height="18" viewBox="0 0 48 48">
                <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.35-8.16 2.35-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
            </svg>
            Lanjutkan dengan Google
        </a>

        <!-- [HTML] Divider -->
        <div class="login-divider">
            <div class="login-divider-line"></div>
            <span class="login-divider-text">atau</span>
            <div class="login-divider-line"></div>
        </div>

        <!-- [HTML] Form email & password -->
        <form method="POST" action="/login">
            <div class="form-group-login">
                <label>Email</label>
                <input type="email" name="email" placeholder="nama@perusahaan.com"
                       value="{{ email or "" }}" required autocomplete="email">
            </div>
            <div class="form-group-login">
                <label>Password</label>
                <input type="password" name="password" placeholder="Masukkan password"
                       required autocomplete="current-password">
            </div>

            <!-- [HTML] Opsi ingat saya & lupa password -->
            <div class="login-options">
                <label>
                    <input type="checkbox" name="remember"> Ingat saya
                </label>
                <a href="/lupa-password">Lupa password?</a>
            </div>

            <button type="submit" class="btn-login">Masuk ke PendidikanStore</button>
        </form>

        <!-- [HTML] Link registrasi -->
        <div class="login-register">
            Belum punya akun? <a href="/register">Daftar sekarang</a>
        </div>
    </div>
</div>

<!-- [JAVASCRIPT] Auto-hide flash setelah 4 detik -->
<script>
document.querySelectorAll(".login-flash").forEach(function(el) {
    setTimeout(function() {
        el.style.opacity = "0";
        el.style.transition = "opacity 0.4s";
        setTimeout(function() { el.remove(); }, 400);
    }, 4000);
});
</script>
</body>
</html>
"""

# [HTML] Template halaman registrasi
REGISTER_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daftar Akun — PendidikanStore</title>
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>{{ css|safe }}</style>
</head>
<body>
<div class="login-wrapper">
    <div class="login-left">
        <div class="login-logo">
            <div class="login-logo-icon">S</div>
            <div class="login-logo-text">
                <div class="login-logo-name">PendidikanStore</div>
                <div class="login-logo-sub">Perlengkapan Siswa & Mahasiswa</div>
            </div>
        </div>
        <div class="login-hero">
            <h2>Buat Akun <span>Baru</span></h2>
            <p>Daftarkan diri kamu untuk mengakses sistem penjualan PendidikanStore.</p>
        </div>
        <div class="login-features">
            <div class="login-feature"><div class="login-feature-dot"></div>Akses semua fitur sistem</div>
            <div class="login-feature"><div class="login-feature-dot"></div>Data tersimpan aman & terenkripsi</div>
            <div class="login-feature"><div class="login-feature-dot"></div>Multi user dengan role berbeda</div>
        </div>
    </div>
    <div class="login-right">
        <h1 class="login-title">Buat Akun</h1>
        <p class="login-subtitle">Isi data di bawah untuk mendaftar</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="login-flash {{ category }}">{{ message }}</div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <form method="POST" action="/register">
            <div class="form-group-login">
                <label>Nama Lengkap</label>
                <input type="text" name="nama" placeholder="John Doe" required>
            </div>
            <div class="form-group-login">
                <label>Email</label>
                <input type="email" name="email" placeholder="nama@perusahaan.com" required>
            </div>
            <div class="form-group-login">
                <label>Password</label>
                <input type="password" name="password" placeholder="Minimal 6 karakter" required minlength="6">
            </div>
            <div class="form-group-login">
                <label>Konfirmasi Password</label>
                <input type="password" name="confirm_password" placeholder="Ulangi password" required>
            </div>
            <button type="submit" class="btn-login" style="margin-top:4px">Buat Akun</button>
        </form>
        <div class="login-register">
            Sudah punya akun? <a href="/login">Masuk di sini</a>
        </div>
    </div>
</div>
</body>
</html>
"""


# =============================================================================
# [PYTHON] -- ROUTE: SETUP ADMIN (jalankan sekali untuk buat akun admin)
# Buka: http://localhost:5000/setup-admin
# =============================================================================
@app.route("/setup-admin")
def setup_admin():
    """[PYTHON] Membuat akun admin default jika belum ada."""
    try:
        cur = mysql.connection.cursor()
        # Buat tabel users jika belum ada
        # [SQL] Tambah kolom gambar ke tabel produk jika belum ada
        try:
            cur.execute("ALTER TABLE produk ADD COLUMN IF NOT EXISTS gambar VARCHAR(255) NULL")
            mysql.connection.commit()
        except: pass

        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id         INT AUTO_INCREMENT PRIMARY KEY,
                nama       VARCHAR(100) NOT NULL,
                email      VARCHAR(100) UNIQUE NOT NULL,
                password   VARCHAR(255) NOT NULL,
                role       ENUM("admin","staff") DEFAULT "staff",
                avatar     VARCHAR(10) DEFAULT "A",
                provider   VARCHAR(20) DEFAULT "email",
                status     ENUM("aktif","nonaktif") DEFAULT "aktif",
                last_login TIMESTAMP NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Cek apakah admin sudah ada
        cur.execute("SELECT id FROM users WHERE email=%s", ("admin@pendidikanstore.com",))
        if cur.fetchone():
            cur.close()
            return "<h2>Admin sudah ada!</h2><a href='/login'>Login sekarang</a>"

        # Hash password dan insert admin
        hashed = generate_password_hash("admin123", method="pbkdf2:sha256")
        cur.execute("""
            INSERT INTO users (nama, email, password, role, avatar)
            VALUES (%s, %s, %s, %s, %s)
        """, ("Administrator", "admin@pendidikanstore.com", hashed, "admin", "A"))
        mysql.connection.commit()
        cur.close()
        return """
        <div style="font-family:Arial;max-width:400px;margin:80px auto;text-align:center">
            <h2 style="color:#2d6a4f">Akun Admin Berhasil Dibuat!</h2>
            <p style="color:#666;margin:12px 0">Gunakan kredensial berikut untuk login:</p>
            <div style="background:#f9f8f6;border:1px solid #bbf7d0;border-radius:8px;padding:20px;margin:16px 0">
                <p><b>Email:</b> admin@pendidikanstore.com</p>
                <p><b>Password:</b> admin123</p>
            </div>
            <p style="color:#c0392b;font-size:12px">Segera ganti password setelah login!</p>
            <a href="/login" style="display:inline-block;margin-top:16px;background:#052e16;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:600">Login Sekarang</a>
        </div>"""
    except Exception as e:
        return f"<h2>Error: {e}</h2><p>Pastikan database sistem_penjualan sudah ada!</p>"


# =============================================================================
# [PYTHON] -- ROUTE LOGIN
# =============================================================================
@app.route("/login", methods=["GET","POST"])
def login():
    """[PYTHON] Sistem login dinonaktifkan — langsung redirect ke dashboard."""
    return redirect("/dashboard")


# =============================================================================
# [PYTHON] -- ROUTE REGISTER
# =============================================================================
@app.route("/register", methods=["GET","POST"])
def register():
    """[PYTHON] Sistem login dinonaktifkan — langsung redirect ke dashboard."""
    return redirect("/dashboard")


# =============================================================================
# [PYTHON] -- ROUTE LOGOUT
# =============================================================================
@app.route("/logout")
def logout():
    """[PYTHON] Sistem login dinonaktifkan — redirect ke dashboard."""
    return redirect("/dashboard")


# =============================================================================
# [PYTHON] -- ROUTE LOGIN GOOGLE (simulasi — butuh setup Google OAuth)
# =============================================================================
@app.route("/login/google")
def login_google():
    """
    [PYTHON] Login dengan Google OAuth.
    Untuk mengaktifkan fitur ini, ikuti langkah setup di bawah:

    CARA SETUP GOOGLE OAUTH:
    1. Buka: console.cloud.google.com
    2. Buat project baru
    3. APIs & Services -> OAuth consent screen -> isi nama app
    4. APIs & Services -> Credentials -> Create OAuth 2.0 Client ID
    5. Authorized redirect URIs: http://localhost:5000/login/google/callback
    6. Salin Client ID dan Client Secret ke GOOGLE_CLIENT_ID di bawah
    7. pip install authlib requests

    Untuk sekarang: tampilkan instruksi setup
    """
    return """
    <div style="font-family:Arial;max-width:500px;margin:80px auto;padding:32px;
                background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,0.1)">
        <div style="text-align:center;margin-bottom:24px">
            <div style="font-size:40px">🔐</div>
            <h2 style="color:#052e16;margin:8px 0">Google OAuth</h2>
            <p style="color:#4ade80;font-size:13px">Fitur ini memerlukan setup Google Cloud Console</p>
        </div>
        <div style="background:#f9f8f6;border-radius:8px;padding:16px;margin-bottom:20px">
            <p style="font-size:12px;color:#166534;font-weight:700;margin-bottom:8px">LANGKAH SETUP:</p>
            <ol style="font-size:12px;color:#666;line-height:2;padding-left:16px">
                <li>Buka <a href="https://console.cloud.google.com" target="_blank">console.cloud.google.com</a></li>
                <li>Buat project baru</li>
                <li>APIs & Services → OAuth consent screen</li>
                <li>Credentials → Create OAuth 2.0 Client ID</li>
                <li>Redirect URI: <code style="background:#eee;padding:2px 6px;border-radius:3px">http://localhost:5000/login/google/callback</code></li>
                <li>Salin Client ID & Secret ke kode</li>
                <li>Install: <code style="background:#eee;padding:2px 6px;border-radius:3px">pip install authlib requests</code></li>
            </ol>
        </div>
        <p style="font-size:12px;color:#4ade80;text-align:center">Untuk sementara, gunakan login email & password</p>
        <a href="/login" style="display:block;text-align:center;margin-top:16px;background:#052e16;
           color:#fff;padding:10px;border-radius:6px;text-decoration:none;font-weight:600;font-size:13px">
            ← Kembali ke Login
        </a>
    </div>"""




# =============================================================================
# [PYTHON] -- FITUR BARU 1: API CEK STOK (untuk notifikasi real-time)
# =============================================================================

@app.route('/api/cek-stok')
def api_cek_stok():
    """
    [PYTHON] API endpoint untuk polling notifikasi real-time.
    Dipanggil JavaScript setiap 30 detik untuk cek stok menipis & pesanan baru.
    """
    cur = mysql.connection.cursor()

    # [SQL] Cek produk stok menipis
    cur.execute("""
        SELECT nama, stok, satuan FROM produk
        WHERE stok <= stok_minimum AND status='aktif'
        ORDER BY stok ASC LIMIT 5
    """)
    stok_menipis = cur.fetchall()

    # [SQL] Hitung pesanan baru hari ini
    cur.execute("""
        SELECT COUNT(*) as total FROM pesanan
        WHERE DATE(created_at) = CURDATE() AND status='pending'
    """)
    pesanan_baru = cur.fetchone()['total']

    cur.close()

    return jsonify({
        'stok_menipis': list(stok_menipis),
        'pesanan_baru': pesanan_baru
    })


# =============================================================================
# [PYTHON] -- FITUR BARU 2: UPLOAD GAMBAR PRODUK
# =============================================================================

import os
import uuid
# os   -> untuk operasi file system (buat folder, cek file)
# uuid -> generate nama file unik agar tidak bertabrakan

# [PYTHON] Folder penyimpanan gambar produk
UPLOAD_FOLDER = 'static/uploads/produk'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp', 'gif'}

def allowed_file(filename):
    """[PYTHON] Cek apakah ekstensi file diperbolehkan."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def save_gambar(file):
    """
    [PYTHON] Simpan file gambar ke folder uploads.
    Nama file diganti dengan UUID agar unik dan aman.
    """
    if file and allowed_file(file.filename):
        # Buat folder jika belum ada
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        # Generate nama file unik: uuid4 + ekstensi asli
        ext      = file.filename.rsplit('.', 1)[1].lower()
        filename = str(uuid.uuid4()) + '.' + ext
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        return filename
    return None

# Route untuk serve gambar produk
from flask import send_from_directory

@app.route('/static/uploads/produk/<filename>')
def uploaded_file(filename):
    """[PYTHON] Serve file gambar yang sudah diupload."""
    return send_from_directory(UPLOAD_FOLDER, filename)


# =============================================================================
# [PYTHON] -- FITUR BARU 3: KATALOG PRODUK (Marketplace Style)
# =============================================================================

@app.route('/katalog')
def katalog():
    """
    [PYTHON] Halaman katalog produk bergaya marketplace (Shopee/Tokopedia).
    Menampilkan gambar produk, harga, dan stok dalam grid card.
    """
    cur = mysql.connection.cursor()

    kategori_filter = request.args.get('kategori', '')
    search          = request.args.get('search', '')
    sort            = request.args.get('sort', 'nama')

    # [SQL] Query produk dengan filter dan sorting
    query  = """
        SELECT p.*, k.nama as nama_kategori
        FROM produk p LEFT JOIN kategori k ON p.kategori_id = k.id
        WHERE p.status = 'aktif'
    """
    params = []

    if kategori_filter:
        query += " AND p.kategori_id = %s"
        params.append(kategori_filter)

    if search:
        query += " AND p.nama LIKE %s"
        params.append(f'%{search}%')

    # Sorting options
    sort_map = {
        'nama':        'p.nama ASC',
        'harga_asc':   'p.harga ASC',
        'harga_desc':  'p.harga DESC',
        'stok':        'p.stok DESC',
    }
    query += f" ORDER BY {sort_map.get(sort, 'p.nama ASC')}"

    cur.execute(query, params)
    produk_list = cur.fetchall()

    # [SQL] Ambil semua kategori untuk filter
    cur.execute("SELECT * FROM kategori ORDER BY nama")
    kategori_list = cur.fetchall()

    # [SQL] Statistik singkat
    cur.execute("SELECT COUNT(*) as total FROM produk WHERE status='aktif'")
    total_produk = cur.fetchone()['total']

    cur.close()

    # [PYTHON] Bangun kartu produk
    def fmt_harga(h):
        return f"Rp {int(h):,}".replace(',', '.')

    # [PYTHON] Build filter kategori buttons
    filter_html = f'''
    <a href="/katalog?sort={sort}" class="filter-btn {"active" if not kategori_filter else ""}">
        Semua ({total_produk})
    </a>'''
    for k in kategori_list:
        active = 'active' if str(k['id']) == str(kategori_filter) else ''
        filter_html += f'''
        <a href="/katalog?kategori={k["id"]}&sort={sort}"
           class="filter-btn {active}">{k["nama"]}</a>'''

    # [PYTHON] Sort options
    sort_options = {
        'nama': 'Nama A-Z',
        'harga_asc': 'Harga Termurah',
        'harga_desc': 'Harga Termahal',
        'stok': 'Stok Terbanyak'
    }
    sort_html = '''<select onchange="window.location='/katalog?sort='+this.value+'&kategori=' + ''' + kategori_filter + '''"
        style="padding:6px 10px;border:1.5px solid var(--border);border-radius:6px;
               font-size:12px;font-family:inherit;background:var(--surface);color:var(--text);outline:none">'''
    for val, label in sort_options.items():
        selected = 'selected' if sort == val else ''
        sort_html += f'<option value="{val}" {selected}>{label}</option>'
    sort_html += '</select>'

    # [PYTHON] Build product cards
    cards_html = ''
    for p in produk_list:
        harga      = fmt_harga(p['harga'])
        stok_kritis = p['stok'] <= p['stok_minimum']
        badge_html = ''
        if stok_kritis:
            badge_html = '<span class="katalog-badge">Stok Tipis</span>'
        elif p['stok'] > 50:
            badge_html = '<span class="katalog-badge green">Ready</span>'

        # Gambar produk
        if p.get('gambar'):
            img_html = f'<img src="/static/uploads/produk/{p["gambar"]}" alt="{p["nama"]}" loading="lazy">'
        else:
            # Placeholder emoji berdasarkan kategori
            emoji_map = {'Elektronik': '💻', 'Pakaian': '👕', 'Makanan': '🍜', 'Alat Tulis': '✏️'}
            emoji = emoji_map.get(p['nama_kategori'], '📦')
            img_html = f'<div class="katalog-img-placeholder">{emoji}</div>'

        warna_harga = 'color:var(--red)' if stok_kritis else 'color:var(--red)'

        cards_html += f'''
        <a href="/produk/edit/{p["id"]}" class="katalog-card">
            <div class="katalog-img">
                {badge_html}
                {img_html}
            </div>
            <div class="katalog-info">
                <div class="katalog-kategori">{p["nama_kategori"] or "Umum"}</div>
                <div class="katalog-nama">{p["nama"]}</div>
                <div class="katalog-harga">{harga}</div>
                <div class="katalog-stok">Stok: {p["stok"]} {p["satuan"]}</div>
            </div>
            <div class="katalog-footer">
                <span>{p["kode_produk"]}</span>
                <span style="font-size:11px;color:{"var(--red)" if stok_kritis else "var(--green)"}">
                    {"⚠ Menipis" if stok_kritis else "✓ Tersedia"}
                </span>
            </div>
        </a>'''

    if not cards_html:
        cards_html = '''
        <div style="grid-column:1/-1;text-align:center;padding:60px;color:var(--text-3)">
            <div style="font-size:48px;margin-bottom:12px">🔍</div>
            <p style="font-size:14px">Tidak ada produk ditemukan</p>
        </div>'''

    content = f'''
    <div class="page-header">
        <div>
            <h2>Katalog Produk</h2>
            <div class="breadcrumb">Produk / <span>Katalog</span></div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
            <!-- [HTML] Search produk -->
            <form method="GET" action="/katalog" style="display:flex;gap:8px">
                <input type="hidden" name="kategori" value="{kategori_filter}">
                <input type="hidden" name="sort" value="{sort}">
                <input class="search-input" type="text" name="search"
                       placeholder="Cari produk..." value="{search}" style="min-width:200px">
                <button type="submit" class="btn btn-outline btn-sm">Cari</button>
            </form>
            {sort_html}
            <a href="/produk/tambah" class="btn btn-primary btn-sm">+ Tambah Produk</a>
        </div>
    </div>

    <!-- [HTML] Filter kategori -->
    <div class="katalog-filter">
        {filter_html}
    </div>

    <!-- [HTML] Info jumlah produk -->
    <div style="font-size:12px;color:var(--text-3);margin-bottom:12px">
        Menampilkan <strong style="color:var(--text)">{len(produk_list)}</strong> produk
        {f"dalam kategori <strong>{kategori_filter}</strong>" if kategori_filter else ""}
    </div>

    <!-- [HTML] Grid kartu produk marketplace -->
    <div class="katalog-grid">
        {cards_html}
    </div>
    '''

    return render_page('Katalog Produk', 'produk', content)


# [PYTHON] Update route tambah/edit produk untuk support upload gambar
@app.route('/produk/tambah-foto', methods=['GET', 'POST'])
def produk_tambah_foto():
    """
    [PYTHON] Form tambah produk dengan fitur upload foto.
    Menggunakan enctype multipart/form-data untuk upload file.
    """
    cur = mysql.connection.cursor()

    if request.method == 'POST':
        kode         = generate_kode('PRD', 'produk', 'kode_produk')
        nama         = request.form['nama']
        kategori_id  = request.form['kategori_id'] or None
        harga        = request.form['harga']
        stok         = request.form['stok']
        stok_minimum = request.form['stok_minimum']
        satuan       = request.form['satuan']
        deskripsi    = request.form.get('deskripsi', '')

        # [PYTHON] Proses upload gambar jika ada
        gambar = None
        if 'gambar' in request.files:
            file = request.files['gambar']
            if file.filename:
                gambar = save_gambar(file)

        # [SQL] Cek apakah kolom gambar sudah ada
        try:
            cur.execute("""
                ALTER TABLE produk ADD COLUMN IF NOT EXISTS gambar VARCHAR(255) NULL
            """)
            mysql.connection.commit()
        except:
            pass

        cur.execute("""
            INSERT INTO produk (kode_produk, nama, kategori_id, harga, stok,
                               stok_minimum, satuan, deskripsi, gambar)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (kode, nama, kategori_id, harga, stok, stok_minimum, satuan, deskripsi, gambar))
        mysql.connection.commit()
        cur.close()

        # [JAVASCRIPT] Kirim notifikasi produk baru
        flash(f'Produk {nama} berhasil ditambahkan!', 'success')
        return redirect('/katalog')

    cur.execute("SELECT * FROM kategori")
    kategori = cur.fetchall()
    cur.close()

    opt_kat = '<option value="">— Pilih Kategori —</option>'
    for k in kategori:
        opt_kat += f'<option value="{k["id"]}">{k["nama"]}</option>'

    content = f'''
    <div class="page-header">
        <div><h2>Tambah Produk + Foto</h2>
        <div class="breadcrumb"><a href="/katalog">Katalog</a> / Tambah</div></div>
    </div>
    <div class="form-card" style="max-width:700px">
        <div class="form-section-title">Informasi Produk & Foto</div>
        <!-- [HTML] enctype multipart/form-data wajib untuk upload file -->
        <form method="POST" enctype="multipart/form-data">
            <div class="form-grid">
                <!-- Upload foto -->
                <div class="form-group full">
                    <label>Foto Produk</label>
                    <div class="upload-area" onclick="document.getElementById('inputGambar').click()">
                        <div id="previewContainer">
                            <div style="font-size:32px">📷</div>
                            <p style="font-size:12px;color:var(--text-3);margin-top:8px">
                                Klik untuk upload foto produk<br>
                                <span style="font-size:10px">PNG, JPG, WEBP — Max 5MB</span>
                            </p>
                        </div>
                        <!-- [HTML] Input file tersembunyi -->
                        <input type="file" id="inputGambar" name="gambar"
                               accept="image/*" style="display:none"
                               onchange="previewGambar(this)">
                    </div>
                </div>
                <div class="form-group full">
                    <label>Nama Produk *</label>
                    <input type="text" name="nama" required placeholder="Nama lengkap produk">
                </div>
                <div class="form-group">
                    <label>Kategori</label>
                    <select name="kategori_id">{opt_kat}</select>
                </div>
                <div class="form-group">
                    <label>Satuan</label>
                    <select name="satuan">
                        <option value="pcs">pcs</option>
                        <option value="unit">unit</option>
                        <option value="kg">kg</option>
                        <option value="liter">liter</option>
                        <option value="box">box</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Harga Jual (Rp) *</label>
                    <input type="number" name="harga" required placeholder="0" min="0">
                </div>
                <div class="form-group">
                    <label>Stok Awal</label>
                    <input type="number" name="stok" value="0" min="0">
                </div>
                <div class="form-group">
                    <label>Stok Minimum</label>
                    <input type="number" name="stok_minimum" value="5" min="0">
                </div>
                <div class="form-group full">
                    <label>Deskripsi Produk</label>
                    <textarea name="deskripsi" placeholder="Deskripsikan produk kamu..."></textarea>
                </div>
            </div>
            <div class="form-actions">
                <button type="submit" class="btn btn-primary">Simpan Produk</button>
                <a href="/katalog" class="btn btn-outline">Batal</a>
            </div>
        </form>
    </div>
    '''

    # [JAVASCRIPT] Preview gambar sebelum upload
    extra_scripts = '''
    <script>
    function previewGambar(input) {
        if (input.files && input.files[0]) {
            var reader = new FileReader();
            reader.onload = function(e) {
                var preview = document.getElementById("previewContainer");
                preview.innerHTML =
                    "<img src='" + e.target.result + "' class='upload-preview'>" +
                    "<p style='font-size:11px;color:var(--text-3);margin-top:8px'>" +
                    input.files[0].name + "</p>";
                // Trigger notifikasi
                tambahNotif("📸", "Foto dipilih!", input.files[0].name);
            };
            reader.readAsDataURL(input.files[0]);
        }
    }
    </script>
    '''

    return render_page('Tambah Produk + Foto', 'produk', content, extra_scripts)


if __name__ == '__main__':
    app.run()

    app.run(
        debug=True,   # debug=True: tampilkan error detail di browser (matikan di production!)
        host='0.0.0.0', # host='0.0.0.0': bisa diakses dari jaringan lokal (bukan hanya localhost)
        port=5000     # port: nomor port server (akses via http://localhost:5000)
    )

# =============================================================================
# RINGKASAN STRUKTUR FILE:
# ─────────────────────────────────────────────────────────────────────────────
# [PYTHON] Baris   1 –  30  : Header, deskripsi, dan cara menjalankan
# [PYTHON] Baris  31 –  60  : Import library
# [PYTHON] Baris  61 –  90  : Inisialisasi Flask & konfigurasi database
# [SQL]    Baris  91 – 180  : Schema database MySQL (sebagai komentar)
# [PYTHON] Baris 181 – 220  : Fungsi pembantu (generate_kode, format_rupiah)
# [CSS]    Baris 221 – 450  : Gaya tampilan seluruh halaman (CSS_GLOBAL)
# [HTML]   Baris 451 – 540  : Template dasar sidebar + topbar (BASE_TEMPLATE)
# [JS]     Baris 541 – 620  : JavaScript utama (flash, format rupiah, form pesanan)
# [PYTHON] Baris 621 – 720  : Route Dashboard
# [PYTHON] Baris 721 – 850  : Route Pelanggan (CRUD)
# [PYTHON] Baris 851 – 980  : Route Produk (CRUD)
# [PYTHON] Baris 981 –1100  : Route Pesanan (tambah, detail, update status)
# [PYTHON] Baris 1101–1200  : Route Laporan
# [PYTHON] Baris 1201–akhir : Entry point aplikasi
# =============================================================================
