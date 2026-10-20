#!/usr/bin/env python3
"""
Hotel Voucher Recap - Collector
================================
Membaca email voucher OTA (Agoda, Traveloka, Tiket.com, Bookandlink: Booking.com/Ctrip/dll)
dari mailbox reservasi via IMAP (READ-ONLY, email tidak ditandai dibaca / tidak dipindah),
membaca data booking, membuat PDF voucher, lalu mengirim ke Google Apps Script
yang menyimpan PDF ke Google Drive (Unit / Bulan / Tanggal Check-in) dan mencatat ke Sheet rekap.

Mode:
  python collector.py                 -> jalan normal (dipakai GitHub Actions)
  python collector.py --dry-run       -> baca IMAP + parse, TIDAK kirim ke Drive (cek log saja)
  python collector.py --test a.eml .. -> uji parser dari file .eml lokal (PDF disimpan ke ./test_output)

Environment (GitHub Secrets):
  IMAP_HOST, IMAP_PORT (default 993), IMAP_USER, IMAP_PASS, IMAP_FOLDER (default INBOX)
  UNIT_NAME   (mis. "Verse Lite Gajah Mada")
  GAS_URL     (URL Web App Apps Script), GAS_TOKEN (token rahasia, sama dengan di Apps Script)
  LOOKBACK_DAYS (default 3)
"""
import base64
import datetime as dt
import email
import email.policy
import imaplib
import io
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.request
from email.utils import parsedate_to_datetime

from bs4 import BeautifulSoup

# ----------------------------------------------------------------------------
# Konfigurasi
# ----------------------------------------------------------------------------
SENDERS = {
    "agoda.com": "AGODA",
    "traveloka.com": "TRAVELOKA",
    "tiket.com": "TIKET",
    "bookandlink.com": "BOOKANDLINK",
    "trip.com": "TRIPCOM",
}
# OTA yang sudah punya email langsung -> email Bookandlink untuk OTA ini dilewati (hindari dobel)
DIRECT_OTAS = {"AGODA", "TRAVELOKA", "CTRIP", "TRIP.COM"}   # Tiket.com tidak: Oak Tree hanya menerima Tiket lewat Bookandlink (dobel dicegah di server)
# Jaring pengaman (16 Okt 2026): notifikasi Bookandlink untuk OTA di atas biasanya dobel dengan email langsung OTA.
# Tapi kalau email langsung OTA TIDAK sampai (kasus Traveloka #...6326 Oak Tree 15 Okt), booking hilang dari laporan.
# Maka: email Bookandlink yang sudah berumur >= BNL_TUNDA_JAM tetap dikirim ke Apps Script -> di sana dibuang bila
# booking ID yang sama sudah tercatat, atau disimpan sebagai cadangan bila belum ada. Yang lebih muda ditunda (tidak
# dicatat) supaya dicek lagi run berikutnya, saat email langsung OTA mestinya sudah masuk.
BNL_TUNDA_JAM = float(os.environ.get("BNL_TUNDA_JAM") or "3")
BNL_TANDA = "via Bookandlink"   # penanda di Catatan (Apps Script BNL_TANDA_) -> server tahu ini cadangan
# Batas waktu per unit (16 Okt 2026): run 16.05 WIB 15 Okt macet 55 menit di unit pertama -> semua unit tidak terbaca.
EMAIL_MAKS_MENIT = float(os.environ.get("EMAIL_MAKS_MENIT") or "9")    # lewat ini: email sisa dilanjut run berikutnya
UNIT_MAKS_MENIT = float(os.environ.get("UNIT_MAKS_MENIT") or "14")     # lewat ini: unit dihentikan paksa, unit berikutnya jalan
# Laporan VHP (PDF upload FO) SEMUA unit dibaca paling awal pada step pertama job (17 Okt 2026): unit terakhir dulu baru
# terbaca ~08.40, sesudah laporan 08.30 -> "Laporan VHP belum ada" padahal FO sudah upload.
SEMUA_UNIT = [u.strip() for u in (os.environ.get("SEMUA_UNIT") or
              "Verse Lite Gajah Mada|Oak Tree Mahakam Blok M|Verse Hotel Cirebon|Verse Luxe Hotel Wahid Hasyim").split("|") if u.strip()]

MONTHS = {
    "jan": 1, "january": 1, "januari": 1,
    "feb": 2, "february": 2, "februari": 2, "peb": 2,
    "mar": 3, "march": 3, "maret": 3,
    "apr": 4, "april": 4,
    "may": 5, "mei": 5,
    "jun": 6, "june": 6, "juni": 6,
    "jul": 7, "july": 7, "juli": 7,
    "aug": 8, "august": 8, "agu": 8, "agt": 8, "agustus": 8, "ags": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "okt": 10, "oktober": 10,
    "nov": 11, "november": 11, "nop": 11, "nopember": 11,
    "dec": 12, "december": 12, "des": 12, "desember": 12,
}
MON_RE = "|".join(sorted(MONTHS.keys(), key=len, reverse=True))


# ----------------------------------------------------------------------------
# Helper umum
# ----------------------------------------------------------------------------
def log(*a):
    print(*a, flush=True)


# MODE LOG PUBLIK (repo public): log TIDAK boleh memuat data tamu (nama, harga, nama file, isi error).
# Aktif bila env LOG_PUBLIK=1 (diset di workflow). Detail lengkap tetap tersimpan di Sheet/Drive.
LOG_PUBLIK = os.environ.get("LOG_PUBLIK") == "1"


def _e(e):
    """Pesan error yang aman ditampilkan: di mode publik hanya jenis error (+ kode HTTP / timeout)."""
    if not LOG_PUBLIK:
        return str(e)
    t = str(e)
    kode = re.search(r"HTTP Error \d{3}|timed out|timeout|Connection (?:reset|refused)|EOF", t, re.I)
    return type(e).__name__ + (f" ({kode.group(0)})" if kode else "")


def _samar_id(x):
    """Nomor booking disamarkan: hanya 4 digit terakhir."""
    x = str(x or "")
    return ("…" + x[-4:]) if (LOG_PUBLIK and len(x) > 4) else x


if LOG_PUBLIK:   # error tak terduga: jangan cetak traceback (bisa memuat data tamu)
    def _hook(t, v, tb):
        print("FATAL:", _e(v), flush=True)
    sys.excepthook = _hook


def to_number(s):
    """'IDR 655,159.00' / 'Rp 427,608' / '333.534,80' / '-147,374.00' -> float"""
    if s is None:
        return None
    s = str(s)
    m = re.search(r"-?\d[\d.,]*", s)
    if not m:
        return None
    t = m.group(0)
    neg = t.startswith("-")
    t = t.lstrip("-")
    last_dot, last_com = t.rfind("."), t.rfind(",")
    last = max(last_dot, last_com)
    if last >= 0 and len(t) - last - 1 in (1, 2):
        # separator terakhir = desimal
        dec_sep = t[last]
        intpart = re.sub(r"[.,]", "", t[:last])
        t = intpart + "." + t[last + 1:]
    else:
        t = re.sub(r"[.,]", "", t)
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def parse_date(s):
    """Mengembalikan datetime.date dari berbagai format tanggal OTA."""
    if not s:
        return None
    s = s.strip()
    # (29-09-2026) / 29-09-2026 / 29/09/2026
    m = re.search(r"\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b", s)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return dt.date(y, mo, d)
        except ValueError:
            pass
    # 2026-09-29 / 2026/09/24
    m = re.search(r"\b(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\b", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return dt.date(y, mo, d)
        except ValueError:
            pass
    # 29-Sep-2026 / 29 Sep 2026 / 01 Okt 2026
    m = re.search(r"\b(\d{1,2})[\s\-]+(" + MON_RE + r")[a-z]*\.?[\s\-,]+(\d{4})\b", s, re.I)
    if m:
        try:
            return dt.date(int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1)))
        except (ValueError, KeyError):
            pass
    # September 29, 2026 / Sep 29, 2026
    m = re.search(r"\b(" + MON_RE + r")[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b", s, re.I)
    if m:
        try:
            return dt.date(int(m.group(3)), MONTHS[m.group(1).lower()], int(m.group(2)))
        except (ValueError, KeyError):
            pass
    return None


def html_to_lines(html):
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style", "head", "title"]):
        t.decompose()
    text = soup.get_text("\n")
    lines = []
    for ln in text.splitlines():
        ln = re.sub(r"[\s ​]+", " ", ln).strip()
        if ln:
            lines.append(ln)
    return lines


def text_to_lines(text):
    out = []
    for ln in (text or "").splitlines():
        ln = re.sub(r"[\s ​]+", " ", ln).strip()
        if ln:
            out.append(ln)
    return out


class Lines:
    """Pencarian label -> nilai pada teks email yang sudah dipecah per baris."""

    def __init__(self, lines):
        self.lines = lines
        self.text = "\n".join(lines)

    def find(self, pattern, flags=re.I):
        m = re.search(pattern, self.text, flags)
        return m.group(1).strip() if m else None

    def after(self, label_re, n=1, skip_re=None, max_look=6):
        """Nilai pada baris SETELAH baris yang cocok label_re (atau sisa baris yang sama)."""
        for i, ln in enumerate(self.lines):
            m = re.search(label_re, ln, re.I)
            if not m:
                continue
            rest = ln[m.end():].strip(" :\t-")
            if rest and not (skip_re and re.fullmatch(skip_re, rest, re.I)):
                return rest
            got = []
            for nxt in self.lines[i + 1:i + 1 + max_look]:
                nxt = nxt.strip(" :")
                if not nxt:
                    continue
                if skip_re and re.fullmatch(skip_re, nxt, re.I):
                    continue
                got.append(nxt)
                if len(got) >= n:
                    return " ".join(got)
        return None


NIGHT_RE = re.compile(r"((?:" + MON_RE + r")[a-z]*\.?\s+\d{1,2},\s*\d{4}|\d{1,2}[- ](?:" + MON_RE + r")[a-z]*[- ]\d{4})"
                      r"\s*\n?\s*(?:IDR|Rp\.?)\s*\n?\s*([\d.,]+)", re.I)


def nightly_rates(text, ci=None, co=None):
    """Rincian harga per malam di voucher (Agoda 'October 1, 2026 IDR 309,750.00', Traveloka 'October 01, 2026 / IDR 321285 ...').
    Return daftar harga urut tanggal; hanya tanggal di dalam masa inap."""
    seen = {}
    for m in NIGHT_RE.finditer(text or ""):
        d, amt = parse_date(m.group(1)), to_number(m.group(2))
        if not d or not amt or amt < 1000:
            continue
        if ci and co and not (ci <= d < co):
            continue
        seen.setdefault(d, amt)
    return [seen[d] for d in sorted(seen)]


def set_nightly(r, text):
    rates = nightly_rates(text, r.get("checkin"), r.get("checkout"))
    if rates:
        r["rate_per_malam"] = "|".join(str(int(round(x))) for x in rates)
        if not r.get("rate_malam1"):
            r["rate_malam1"] = rates[0]


def empty_record():
    return {
        "ota": "", "jenis": "BARU", "booking_id": "", "tamu": "",
        "checkin": "", "checkout": "", "malam": "", "kamar": "", "tipe_kamar": "",
        "pembayaran": "", "mata_uang": "IDR",
        "harga_jual": None, "komisi": None, "pajak_komisi": None, "net": None,
        "rate_malam1": None,  # harga malam pertama (jika voucher mencantumkan rincian per malam)
        "rate_per_malam": None,  # daftar harga tiap malam "309750|309750|340362" (harga bisa beda tiap malam)
        "catatan": [], "status": "OK",
    }


def detect_jenis(subject, body_text):
    s = (subject or "").lower()
    b = (body_text or "")[:1500].lower()
    if re.search(r"cancel|batal|dibatalkan|pembatalan", s):
        return "CANCEL"
    if re.search(r"modif|amend|change|ubah|diubah|perubahan|revis", s):
        return "UBAH"
    if re.search(r"\bcancell?ed\b|pesanan dibatalkan|reservation cancel", b):
        return "CANCEL"
    return "BARU"


# ----------------------------------------------------------------------------
# Parser per OTA
# ----------------------------------------------------------------------------
def parse_agoda(L, subject):
    r = empty_record()
    r["ota"] = "Agoda"
    r["booking_id"] = L.find(r"ID Pesanan Agoda\s*(\d{6,})") or L.find(r"Booking ID[^\d]{0,40}(\d{6,})") or \
        (re.search(r"(\d{8,})", subject or "") or [None, None])[1] or ""
    first = L.after(r"^Customer First Name", skip_re=r"(Nama Awal( Pelanggan)?|Pelanggan)")
    last = L.after(r"^Customer Last Name", skip_re=r"(Nama Akhir( Pelanggan)?|Pelanggan)")
    first = re.sub(r"^Nama Awal\s*(Pelanggan)?\s*", "", first or "", flags=re.I)
    last = re.sub(r"^Nama Akhir\s*(Pelanggan)?\s*", "", last or "", flags=re.I)
    r["tamu"] = " ".join(x for x in [first, last] if x).strip()
    _d = r"(\d{1,2}-[A-Za-z]{3}-\d{4}(?:\s*\(\d{2}-\d{2}-\d{4}\))?|[A-Za-z]{3,9}\.? \d{1,2}, \d{4})"
    ci = L.find(r"Check-in(?:\s*Check-in)?\s*:?\s*\n?\s*" + _d)
    co = L.find(r"Check-out(?:\s*Check-out)?\s*:?\s*\n?\s*" + _d)
    r["checkin"], r["checkout"] = parse_date(ci), parse_date(co)
    # Email cancel Agoda berbahasa Indonesia: "Nama Awal Pelanggan: X Nama Akhir Pelanggan: Y ... Ketibaan: October 25, 2026 Keberangkatan: ..."
    if not r["tamu"]:
        f2 = L.find(r"(?:Nama Awal Pelanggan|Customer First Name)\s*:\s*([^\n:]+?)\s*(?=Nama Akhir Pelanggan|Customer Last Name|Kota\s*:|City\s*:|\n|$)")
        l2 = L.find(r"(?:Nama Akhir Pelanggan|Customer Last Name)\s*:\s*([^\n:]+?)\s*(?=Kota\s*:|City\s*:|Negara\s*:|Country\s*:|\n|$)")
        r["tamu"] = " ".join(x for x in [f2, l2] if x).strip()
    if not r["checkin"]:
        r["checkin"] = parse_date(L.find(r"(?:Ketibaan|Arrival(?: date)?)\s*:\s*\n?\s*" + _d))
    if not r["checkout"]:
        r["checkout"] = parse_date(L.find(r"(?:Keberangkatan|Departure(?: date)?)\s*:\s*\n?\s*" + _d))
    rt = L.after(r"^Tipe Kamar$|^Room Type$", n=1, skip_re=r"(Tipe Kamar|No\. of Rooms|Jumlah kamar|Occupancy|Tingkat hunian|No\. of Extra Bed|Jumlah Extra Bed:?|"
                 r"Children\S?s age|Child(ren)? age|Usia anak|Umur anak)", max_look=12)
    r["tipe_kamar"] = rt or ""
    # Jumlah kamar: baris tabel "Superior Room King Bed | 2 | 3 Adults | 0" (harga di voucher = total semua kamar)
    m = re.search(r"\b(\d{1,2})\s+\d{1,2}\s+(?:Adults?|Dewasa|Orang Dewasa)\b", L.text, re.I)
    if m and 0 < int(m.group(1)) <= 30:
        r["kamar"] = m.group(1)
    up = L.text.upper()
    if "PRA BAYAR" in up or "PREPAID" in up or "PRE-PAID" in up:
        r["pembayaran"] = "Prepaid (Agoda)"
    elif "BAYAR DI HOTEL" in up or "PAY AT HOTEL" in up or "PAY AT THE HOTEL" in up:
        r["pembayaran"] = "Bayar di Hotel"
    # Uang: gunakan regex multi-baris (label bisa terpisah dari angka)
    t = L.text
    m = re.search(r"Referensi harga jual.*?(?:IDR|USD|SGD)\s*\n?\s*(-?[\d.,]+)", t, re.S | re.I) or \
        re.search(r"Reference sell rate.*?(?:IDR|USD|SGD)\s*\n?\s*(-?[\d.,]+)", t, re.S | re.I)
    if m and len(m.group(0)) < 300:
        r["harga_jual"] = abs(to_number(m.group(1)))
    m = re.search(r"Komisi\s*\n?\s*(?:IDR\s*)?\n?\s*(-?[\d.,]+)", t, re.I) or \
        re.search(r"Commission\s*(?:Komisi)?\s*\n?\s*(?:IDR\s*)?\n?\s*(-?[\d.,]+)", t, re.I)
    if m:
        r["komisi"] = abs(to_number(m.group(1)))
    m = re.search(r"Tax on Commission\s*\n?\s*(?:IDR\s*)?\n?\s*(?:Tax on Commission\s*)?\n?\s*(-?[\d.,]+)", t, re.I)
    if m:
        r["pajak_komisi"] = abs(to_number(m.group(1)))
    m = re.search(r"Tarif bersih.*?(?:IDR|USD|SGD)\s*\n?\s*([\d.,]+)", t, re.S | re.I) or \
        re.search(r"Net rate \(incl.*?(?:IDR|USD|SGD)\s*\n?\s*([\d.,]+)", t, re.S | re.I)
    if m and len(m.group(0)) < 300:
        r["net"] = to_number(m.group(1))
    set_nightly(r, t)
    # Validasi silang: harga jual - komisi - pajak komisi harus = net.
    # Jika tidak cocok (urutan teks tabel berantakan), cari kombinasi angka negatif yang cocok.
    hj, net = r["harga_jual"], r["net"]
    if hj and net:
        target = round(hj - net, 2)
        cur = (r["komisi"] or 0) + (r["pajak_komisi"] or 0)
        if abs(cur - target) > 1:
            negs = sorted({abs(to_number(x)) for x in re.findall(r"-\s?[\d][\d.,]*\d", t) if to_number(x)}, reverse=True)
            found = None
            for i, k in enumerate(negs):
                if abs(k - target) <= 1:
                    found = (k, None)
                    break
                for tx in negs[i + 1:]:
                    if abs(k + tx - target) <= 1:
                        found = (k, tx)
                        break
                if found:
                    break
            if found:
                r["komisi"], r["pajak_komisi"] = found
            else:
                r["catatan"].append("Cek manual: harga jual - komisi - pajak tidak sama dengan net")
    m = re.search(r"\b(IDR|USD|SGD)\b", t)
    if m:
        r["mata_uang"] = m.group(1)
    return r


def parse_traveloka(L, subject):
    r = empty_record()
    r["ota"] = "Traveloka"
    r["booking_id"] = L.find(r"Itinerary ID\s*:?\s*\n?\s*(\d{8,})") or \
        (re.search(r"(\d{10,})", subject or "") or [None, None])[1] or ""
    first = L.after(r"^Customer First Name", skip_re=r"Customer Last Name")
    last = L.after(r"^Customer Last Name", skip_re=r"(Customer First Name)")
    # Tata letak tabel: "Customer First Name / Customer Last Name / Yusep Abdul / Kodir"
    idx = next((i for i, ln in enumerate(L.lines) if re.match(r"Customer First Name", ln, re.I)), None)
    if idx is not None and idx + 3 < len(L.lines) and re.match(r"Customer Last Name", L.lines[idx + 1], re.I):
        first, last = L.lines[idx + 2], L.lines[idx + 3]
    elif idx is not None:
        m = re.match(r"Customer First Name\s+Customer Last Name$", L.lines[idx], re.I)
        if m and idx + 1 < len(L.lines):
            first, last = L.lines[idx + 1], ""
    if not (first or last):  # versi Bahasa Indonesia: "Nama Depan Tamu / Nama Belakang Tamu"
        first = L.after(r"^Nama Depan Tamu", skip_re=r"Nama Belakang Tamu")
        last = L.after(r"^Nama Belakang Tamu", skip_re=r"(Nama Depan Tamu|Email Tamu.*)")
        idx = next((i for i, ln in enumerate(L.lines) if re.match(r"Nama Depan Tamu$", ln, re.I)), None)
        if idx is not None and idx + 3 < len(L.lines) and re.match(r"Nama Belakang Tamu$", L.lines[idx + 1], re.I):
            first, last = L.lines[idx + 2], L.lines[idx + 3]
        if last and re.match(r"Email Tamu", last, re.I):
            last = ""
    r["tamu"] = " ".join(x for x in [first, last] if x and not re.search(r"Customer|Nama (Depan|Belakang) Tamu", x)).strip()
    r["checkin"] = parse_date(L.after(r"^Check-in$|^Check-in\b"))
    r["checkout"] = parse_date(L.after(r"^Check-out$|^Check-out\b"))
    m = re.search(r"\(\s*(\d+)\s*[×x]\s*\)\s*(.+)", L.text)
    if m:
        r["kamar"] = m.group(1)
        rt = m.group(2).strip()
        nxt = L.text[m.end():m.end() + 60].split("\n")
        if len(nxt) > 1 and re.fullmatch(r"(Only|Room Only|With Breakfast|Breakfast)", nxt[1].strip(), re.I):
            rt += " " + nxt[1].strip()
        r["tipe_kamar"] = rt
    up = L.text.upper()
    if "PREPAID" in up or "PRABAYAR" in up:
        r["pembayaran"] = "Prepaid (Traveloka)"
    elif "PAY AT HOTEL" in up:
        r["pembayaran"] = "Bayar di Hotel"
    m = re.search(r"(?:Total you will receive|Nominal diterima)\s*:?\s*\n?\s*(IDR|USD)?\s*\n?\s*([\d.,]+)", L.text, re.I)
    if m:
        r["net"] = to_number(m.group(2))
        if m.group(1):
            r["mata_uang"] = m.group(1)
    set_nightly(r, L.text)
    r["catatan"].append("Traveloka hanya mencantumkan net (harga jual & komisi tidak tercantum)")
    return r


def parse_tiket(L, subject):
    """Tiket.com: isi email (English) dibaca dulu, teks PDF e-voucher sebagai pelengkap (harga)."""
    r = empty_record()
    r["ota"] = "Tiket.com"
    r["booking_id"] = L.find(r"Itinerary\s*ID\s*:?\s*\n?\s*:?\s*#?\s*(\d{6,})") or \
        (re.search(r"(\d{7,})", subject or "") or [None, None])[1] or ""
    tamu = L.after(r"^Guest Name$", skip_re=r":") or L.after(r"^Guest Details", skip_re=r":") or ""
    if not tamu:
        m = re.search(r"Kamar\s*\d+\s*:\s*([^\n(]+)", L.text)
        tamu = m.group(1) if m else ""
    r["tamu"] = re.sub(r"^\d+\.\s*", "", tamu).replace("(Dewasa)", "").replace("(Adult)", "").strip()
    ci = L.after(r"^Check-in$")
    co = L.after(r"^Check-out$")
    if ci and co and parse_date(ci) and parse_date(co):
        r["checkin"], r["checkout"] = parse_date(ci), parse_date(co)
    else:
        rng = L.after(r"^Check-in\s*&\s*Check-out", skip_re=r":") or L.after(r"Tanggal Check-in\s*&\s*Check-out") or ""
        parts = re.split(r"\s+-\s+", rng)
        r["checkin"] = parse_date(parts[0]) if parts else None
        r["checkout"] = parse_date(parts[1]) if len(parts) > 1 else None
    r["tipe_kamar"] = L.after(r"^Room (and|&) Rate Plan", skip_re=r":") or L.after(r"^Kamar dan Rate Plan") or ""
    nr = L.after(r"^Number of Rooms?$", skip_re=r":") or ""
    m = re.search(r"(\d+)", nr)
    r["kamar"] = m.group(1) if m else ""
    m = re.search(r"Total Harga\s*\n?\s*(IDR|Rp)?\s*([\d.,]+)", L.text, re.I)
    if m:
        r["net"] = to_number(m.group(2))
    if re.search(r"sudah membayar|already paid|Payment Plan\s*\n?\s*Checkout", L.text, re.I):
        r["pembayaran"] = "Prepaid (Tiket.com)"
    elif re.search(r"bayar di hotel|pay at hotel", L.text, re.I):
        r["pembayaran"] = "Bayar di Hotel"
    return r


def parse_tripcom(L, subject):
    r = empty_record()
    r["ota"] = "Ctrip / Trip.com"
    r["booking_id"] = L.after(r"^Reservation no\.?$") or (re.search(r"(\d{12,})", subject or "") or [None, ""])[1]
    g = L.after(r"^Guest Name:?$") or ""
    if "/" in g:  # RUWANSYAH/AYUB -> AYUB RUWANSYAH
        last, first = g.split("/", 1)
        g = f"{first.strip()} {last.strip()}"
    r["tamu"] = g
    sp = L.after(r"^Staying period:?$") or ""
    parts = re.split(r"\s+-\s+", sp.split("|")[0])
    r["checkin"] = parse_date(parts[0]) if parts else None
    r["checkout"] = parse_date(parts[1]) if len(parts) > 1 else None
    m = re.search(r"(\d+)\s*night", sp)
    if m:
        r["malam"] = m.group(1)
    rt = L.after(r"^Room Type:?$") or ""
    m = re.search(r"\|\s*(\d+)\s*room", rt)
    r["kamar"] = m.group(1) if m else ""
    r["tipe_kamar"] = rt.split("|")[0].strip()
    m = re.search(r"Your payout\s*\n\s*(IDR|USD)?\s*\n?\s*([\d.,]+)", L.text)
    if m:
        r["net"] = to_number(m.group(2))
    m = re.search(r"Original room rate \(incl\. taxes and fees\)\s*\n\s*([\d.,]+)\s*\n", L.text)
    if m:
        r["harga_jual"] = to_number(m.group(1))
    m = re.search(r"\nVAT\s*\n\s*\+?([\d.,]+)\s*\n", L.text)
    if m:
        r["pajak_komisi"] = None
        r["catatan"].append(f"VAT {int(to_number(m.group(1))):,}".replace(",", "."))
    if re.search(r"prepaid", L.text, re.I):
        r["pembayaran"] = "Prepaid (Trip.com)"
    r["catatan"].append("Trip.com: net = 'Your payout'; komisi tidak tercantum")
    return r


def parse_bookandlink_web(L, subject):
    """Booking website hotel via Bookandlink: subjek 'Confirmation Booking #52765704' (invoice ke tamu)."""
    r = empty_record()
    r["ota"] = "Website (Bookandlink)"
    r["_bookandlink_ota"] = "WEBSITE"
    r["booking_id"] = (re.search(r"#\s*(\d{5,})", subject or "") or [None, ""])[1] or \
        L.find(r"Reservation\s*number\s*\n?\s*#?\s*(\d{5,})") or ""
    name = L.after(r"^Guest details$|^Guest details\b") or ""
    # Booking >1 kamar: "Guest details / Room 1: / SAHRU MASKUR / Room 2: ..." -> nama tamu, bukan "Room 1" (18 Okt 2026)
    m = re.search(r"Guest details\b[\s:]*Room\s*\d+\s*:\s*([^\n]+?)\s*(?=Room\s*\d+\s*:|\n|$)", L.text, re.I)
    if m and m.group(1).strip():
        name = m.group(1)
    elif re.fullmatch(r"Room\s*\d+\s*:?", name.strip(), re.I):
        name = L.after(r"^Room\s*1\s*:?$") or ""
    r["tamu"] = re.sub(r"\s+", " ", name).strip()
    lab_in = r"^(Check[- ]?in(\s*date)?|Arrival(\s*date)?|Tanggal check[- ]?in|Kedatangan)\b"
    lab_out = r"^(Check[- ]?out(\s*date)?|Departure(\s*date)?|Tanggal check[- ]?out|Keberangkatan)\b"
    # "Check-in date (02:00pm - 06:00pm) Tuesday, September 29, 2026" (jam di antara label & tanggal)
    _dt = r"([A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{4}-\d{2}-\d{2})"
    m = re.search(r"Check[- ]?in(?:\s*date)?.{0,80}?" + _dt, L.text, re.S | re.I)
    r["checkin"] = parse_date(m.group(1)) if m else parse_date(L.after(lab_in, skip_re=r"(Check[- ]?out.*|Departure.*)"))
    m = re.search(r"Check[- ]?out(?:\s*date)?.{0,80}?" + _dt, L.text, re.S | re.I)
    r["checkout"] = parse_date(m.group(1)) if m else parse_date(L.after(lab_out))
    if not r["checkin"]:   # format rentang: "Oct 01, 2026 - Oct 03, 2026" / "01 Oct 2026 - 03 Oct 2026"
        m = re.search(r"((?:\d{1,2}\s+[A-Za-z]{3,9}|[A-Za-z]{3,9}\s+\d{1,2},?)\s+\d{4})\s*(?:-|–|to|s/d|until)\s*"
                      r"((?:\d{1,2}\s+[A-Za-z]{3,9}|[A-Za-z]{3,9}\s+\d{1,2},?)\s+\d{4})", L.text)
        if m:
            r["checkin"], r["checkout"] = parse_date(m.group(1)), parse_date(m.group(2))
    m = re.search(r"Room\s*Type\s*\(\s*(\d+)\s*Rooms?\s*\)\s*([^\n]+)", L.text, re.I)
    if m:
        r["kamar"], r["tipe_kamar"] = m.group(1), m.group(2).strip()
    else:
        r["tipe_kamar"] = L.after(r"^Room\s*type\b|^Room\s*name\b|^Tipe kamar\b") or ""
    n_kamar = len(set(re.findall(r"Room\s*(\d+)\s*-\s*CONFIRM", L.text, re.I)))
    if n_kamar > 1:   # beberapa kamar dalam 1 reservasi -> jumlah kamar & semua tipe kamarnya
        r["kamar"] = str(n_kamar)
        tipe = list(dict.fromkeys(t.strip() for t in re.findall(r"Room\s*Type\s*\(\s*\d+\s*Rooms?\s*\)\s*([^\n]+)", L.text, re.I)))
        if tipe:
            r["tipe_kamar"] = " + ".join(tipe)
        # harga TIAP kamar ("Total reservation IDR 395,214" per kamar) -> dicocokkan per kamar dengan rate VHP
        tot = [to_number(x) for x in re.findall(r"Total reservation\s*(?:IDR|USD|SGD)?\s*([\d.,]+)", L.text, re.I)]
        tot = [x for x in tot if x]
        try:
            n_mlm = max(1, (dt.date.fromisoformat(str(r["checkout"])[:10]) - dt.date.fromisoformat(str(r["checkin"])[:10])).days)
        except (ValueError, TypeError):
            n_mlm = 1
        if len(tot) >= n_kamar:
            r["rate_per_malam"] = "|".join(str(int(round(x / n_mlm))) for x in tot)
    m = re.search(r"Grand total\s*\n?\s*:?\s*(IDR|USD|SGD)?\s*([\d.,]+)", L.text, re.I) or \
        re.search(r"Invoice total\s*\n?\s*(IDR|USD|SGD)?\s*([\d.,]+)", L.text, re.I)
    if m:
        r["harga_jual"] = to_number(m.group(2))
        if m.group(1):
            r["mata_uang"] = m.group(1)
    up = L.text.upper()
    if re.search(r"\bPAID\b|LUNAS|PAYMENT SUCCESS|ONLINE PAYMENT|DUE PAYMENT IDR 0\b", up):
        r["pembayaran"] = "Prepaid (Website)"
    r["catatan"].append("Booking website hotel (Bookandlink)")
    return r


def parse_bookandlink(L, subject):
    r = empty_record()
    # Subject: "BOOKING.COM Reservation # 5198953888" / "CTRIP Reservation # 13678..."
    m = re.match(r"\s*([A-Za-z0-9 .\-]+?)\s+(Reservation|Cancellation|Modification|Cancel|Modify|Amend\w*)\b", subject or "", re.I)
    ota = (m.group(1).strip() if m else "") or (L.find(r"^(.+?) Reservation$", re.I | re.M) or "")
    ota_up = ota.upper()
    nice = {"BOOKING.COM": "Booking.com", "CTRIP": "Ctrip / Trip.com", "TRIP.COM": "Ctrip / Trip.com",
            "EXPEDIA": "Expedia", "AGODA": "Agoda", "TRAVELOKA": "Traveloka", "TIKET.COM": "Tiket.com", "TIKET": "Tiket.com"}
    r["ota"] = nice.get(ota_up, ota.title() or "Bookandlink")
    r["_bookandlink_ota"] = ota_up
    r["booking_id"] = L.find(r"Reservation\s*#\s*(\w+)") or (re.search(r"#\s*(\w+)", subject or "") or [None, ""])[1]
    r["tamu"] = L.after(r"^Guest\s*:") or L.after(r"^Booked by$|^Booked by\b") or ""
    r["checkin"] = parse_date(L.after(r"^Arrival date"))
    r["checkout"] = parse_date(L.after(r"^Departure date"))
    rn = L.after(r"^Room name") or ""
    r["tipe_kamar"] = re.sub(r"\(\d+[:\d]*\)", "", rn).strip()
    # "System room : Family Loft Quads:Room Only" = nama kamar baku hotel (Room name bisa bahasa lain, mis. "Kamar Keluarga Loteng")
    sysr = (L.after(r"^System room") or "").split(":")[0].strip()
    if sysr and not re.fullmatch(r"N/?A|-", sysr, re.I):
        r["tipe_kamar"] = sysr
    ng = L.after(r"^Number of guest")
    rooms = len(re.findall(r"Room Reservation\s*#", L.text))
    r["kamar"] = str(rooms) if rooms else ""
    m = re.search(r"Total price\s*\n?\s*:?\s*([\d.,]+)\s*(IDR|USD|SGD)?", L.text, re.I)
    if m:
        r["harga_jual"] = to_number(m.group(1))
        if m.group(2):
            r["mata_uang"] = m.group(2)
    m = re.search(r"Total commission\s*\n?\s*:?\s*([\d.,]+)", L.text, re.I) or \
        re.search(r"Commission\s*\n?\s*:?\s*([\d.,]+)", L.text, re.I)
    if m:
        r["komisi"] = to_number(m.group(1))
    # Rincian per malam: "Room rates ... 2026-09-30 204559301:8836884 366,521 IDR"
    # Harga tiap malam bisa berbeda (mis. 425,088 lalu 347,599) -> semua malam dicatat, bukan hanya malam pertama
    i = L.text.lower().find("room rates")
    if i >= 0:
        seen, rates = set(), []
        for d, amt in re.findall(r"\b(\d{4}-\d{2}-\d{2})\s+\S+\s+([\d.,]+)\s*(?:IDR|USD|SGD)", L.text[i:], re.I):
            a = to_number(amt)
            if a and a >= 1000 and (d, a) not in seen:
                seen.add((d, a))
                rates.append(a)
        if rates:
            r["rate_malam1"] = rates[0]
            r["rate_per_malam"] = "|".join(str(int(round(x))) for x in rates)
    remarks = L.text.lower()
    if "prepaid" in remarks or "pre-paid" in remarks or "prepay" in remarks:
        r["pembayaran"] = "Prepaid"
    elif "pay at hotel" in remarks or "pay at the property" in remarks:
        r["pembayaran"] = "Bayar di Hotel"
    if ota_up.startswith("CTRIP") or ota_up.startswith("TRIP"):
        # Ctrip: angka Bookandlink = net rate
        if r["harga_jual"] is not None:
            r["net"], r["harga_jual"] = r["harga_jual"], None
        r["catatan"].append("Ctrip: 'Total price' Bookandlink dicatat sebagai net (harga jual & komisi tidak tercantum)")
    elif r["komisi"] is None:
        r["catatan"].append("Komisi tidak tercantum di email Bookandlink")
    return r


PARSERS = {"AGODA": parse_agoda, "TRAVELOKA": parse_traveloka, "TIKET": parse_tiket,
           "BOOKANDLINK": parse_bookandlink, "TRIPCOM": parse_tripcom}


# ----------------------------------------------------------------------------
# Email -> record + PDF
# ----------------------------------------------------------------------------
def sender_key(from_hdr):
    f = (from_hdr or "").lower()
    for dom, key in SENDERS.items():
        if dom in f:
            return key
    return None


def get_parts(msg):
    html, text, pdfs, cids = None, None, [], {}
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = (part.get("Content-Disposition") or "").lower()
        fname = part.get_filename() or ""
        if part.is_multipart():
            continue
        if ctype == "application/pdf" or fname.lower().endswith(".pdf"):
            pdfs.append((fname or "voucher.pdf", part.get_payload(decode=True)))
            continue
        cid = part.get("Content-ID")
        if cid and ctype.startswith("image/"):
            cids[cid.strip("<>")] = (ctype, part.get_payload(decode=True))
            continue
        if "attachment" in disp:
            continue
        if ctype == "text/html" and html is None:
            html = part.get_content()
        elif ctype == "text/plain" and text is None:
            text = part.get_content()
    return html, text, pdfs, cids


def pdf_text(data):
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception as e:
        log("   ! gagal baca teks PDF:", _e(e))
        return ""


def build_header_html(msg):
    esc = lambda s: (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (
        '<div style="font-family:Arial,sans-serif;font-size:11px;border-bottom:1px solid #999;'
        'padding:6px 0 8px;margin-bottom:10px;color:#222">'
        f"<b>Subject:</b> {esc(str(msg.get('Subject', '')))}<br>"
        f"<b>From:</b> {esc(str(msg.get('From', '')))}<br>"
        f"<b>Date:</b> {esc(str(msg.get('Date', '')))}<br>"
        f"<b>To:</b> {esc(str(msg.get('To', '')))}"
        "</div>"
    )


_browser = None


_CHROMIUM_SIAP = False
def _pastikan_chromium():
    """HEMAT KUOTA: Chromium baru dipasang saat benar-benar ada voucher yang perlu dicetak ke PDF
    (bukan di setiap jalan). Aktif bila env PW_LAZY=1 (diset di workflow)."""
    global _CHROMIUM_SIAP
    if _CHROMIUM_SIAP or os.environ.get("PW_LAZY") != "1":
        return
    import subprocess
    log("   … memasang Chromium (hanya karena ada voucher baru untuk dicetak)")
    subprocess.run([sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"], check=False,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    _CHROMIUM_SIAP = True


def html_to_pdf(html, msg, cids):
    """Render isi email ke PDF (seperti print manual)."""
    global _browser
    for cid, (ctype, data) in cids.items():
        uri = f"data:{ctype};base64,{base64.b64encode(data).decode()}"
        html = html.replace(f"cid:{cid}", uri)
    header = build_header_html(msg)
    if re.search(r"<body[^>]*>", html, re.I):
        html = re.sub(r"(<body[^>]*>)", r"\1" + header.replace("\\", "\\\\"), html, count=1, flags=re.I)
    else:
        html = header + html
    from playwright.sync_api import sync_playwright
    if _browser is None:
        _pastikan_chromium()
        pw = sync_playwright().start()
        _browser = pw.chromium.launch()
    page = _browser.new_page()
    try:
        page.set_content(html, wait_until="domcontentloaded", timeout=15000)
        page.wait_for_load_state("load", timeout=8000)
    except Exception:
        pass  # gambar eksternal lambat -> tetap cetak
    pdf = page.pdf(format="A4", print_background=True,
                   margin={"top": "10mm", "bottom": "10mm", "left": "8mm", "right": "8mm"})
    page.close()
    return pdf


def safe_name(s, maxlen=40):
    s = re.sub(r"[^\w\s\-.]", "", s or "", flags=re.U).strip()
    s = re.sub(r"\s+", " ", s)
    return s[:maxlen].strip()


def finalize(r):
    """Hitung nilai turunan & tandai status."""
    ci, co = r.get("checkin"), r.get("checkout")
    if ci and co and not r.get("malam"):
        r["malam"] = str((co - ci).days)
    if r["net"] is None and r["harga_jual"] is not None and r["komisi"] is not None:
        r["net"] = round(r["harga_jual"] - r["komisi"] - (r["pajak_komisi"] or 0), 2)
        r["catatan"].append("Net dihitung = harga jual - komisi" + (" - pajak komisi" if r["pajak_komisi"] else ""))
    if r["ota"] == "Booking.com" and r["harga_jual"] is not None and r["komisi"] is not None and r["net"] is None:
        r["net"] = round(r["harga_jual"] - r["komisi"], 2)
    missing = [k for k in ("booking_id", "checkin") if not r.get(k)]
    if missing:
        r["status"] = "PERLU DICEK"
        r["catatan"].append("Tidak terbaca: " + ", ".join(missing))
    r["checkin"] = ci.isoformat() if ci else ""
    r["checkout"] = co.isoformat() if co else ""
    r["catatan"] = "; ".join(dict.fromkeys(r["catatan"]))
    return r


def _umur_jam(msg):
    """Umur email (jam) dari header Date; None bila tidak terbaca."""
    try:
        d = parsedate_to_datetime(str(msg.get("Date", "")))
        if d.tzinfo is None:
            d = d.replace(tzinfo=dt.timezone.utc)
        return (dt.datetime.now(dt.timezone.utc) - d).total_seconds() / 3600
    except Exception:
        return None


_ADA_OFF = []   # server Apps Script versi lama (belum kenal 'ada_booking') -> tidak ditanya lagi di jalan ini
_BIDS = {}      # unit -> set Booking ID yang sudah tercatat (dikirim server sekali per run lewat 'check', 17 Okt 2026)


def _ada_booking(unit, bid):
    """Booking ID ini sudah tercatat? Pakai daftar dari server (sekali per run); server lama -> tanya per email.
    Gagal -> False: simpan, server tetap cegah dobel."""
    if not bid:
        return False
    if unit in _BIDS:
        return str(bid) in _BIDS[unit]
    if _ADA_OFF:
        return False
    try:
        return bool(gas_post({"action": "ada_booking", "unit": unit, "booking_id": str(bid)}, retries=1).get("ada"))
    except Exception as e:
        _ADA_OFF.append(1)
        log(f"   ~ cek booking tidak tersedia (lanjut simpan, server tetap cegah dobel): {_e(e)}")
        return False


def process_message(raw, unit, render_pdf=True):
    """raw bytes -> (record, pdf_bytes, filename) atau (None, None, alasan) jika dilewati."""
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    subject = str(msg.get("Subject", "") or "")
    frm = str(msg.get("From", "") or "")
    key = sender_key(frm)
    if not key:
        return None, None, "pengirim bukan OTA"
    html, text, pdfs, cids = get_parts(msg)
    lines = html_to_lines(html) if html else text_to_lines(text)
    # Tiket: gabungkan teks PDF lampiran (jika PDF berbasis teks)
    if key == "TIKET" and pdfs:
        pt = pdf_text(pdfs[0][1])
        if pt.strip():
            lines = lines + text_to_lines(pt)
    L = Lines(lines)
    # Filter: hanya email voucher/booking
    s_low = subject.lower()
    if key == "AGODA" and not re.search(r"id pesanan|booking id|pemesanan|reservation|booking confirmation|dibatalkan|cancel", s_low):
        return None, None, "Agoda bukan voucher"
    # Email tim Customer Experience Agoda ("... ACCOM"): MINTA IZIN hotel membatalkan / mengubah booking -> bukan voucher & belum cancel resmi
    if key == "AGODA" and (re.search(r"\bACCOM\b", subject) or re.search(
            r"permintaan dari pelanggan untuk (membatalkan|mengubah)|request from the customer to (cancel|amend|change)", L.text, re.I)):
        return None, None, "Agoda permintaan pembatalan/perubahan dari tamu (bukan voucher, belum cancel resmi)"
    if key == "TRAVELOKA" and not re.search(r"itinerary", s_low):
        return None, None, "Traveloka bukan voucher"
    # Email tim hotelops Traveloka: MINTA PERSETUJUAN hotel untuk membatalkan -> belum cancel resmi
    if key == "TRAVELOKA" and (re.search(r"pengajuan pembatalan|cancellation request", s_low) or re.search(
            r"mohon konfirmasi persetujuan hotel|konfirmasi persetujuan hotel via email", L.text, re.I)):
        return None, None, "Traveloka pengajuan pembatalan (menunggu persetujuan hotel, belum cancel resmi)"
    if key == "TIKET" and not (pdfs or re.search(r"reservation|reservasi|itinerary|pesanan|cancel|batal", s_low)):
        return None, None, "Tiket bukan voucher"
    if key == "BOOKANDLINK" and re.search(r"payment status|status pembayaran", s_low):
        return None, None, "Bookandlink notifikasi pembayaran (bukan voucher)"
    if key == "BOOKANDLINK" and not re.search(r"reservation|cancel|modif|amend|confirmation booking|booking confirmation", s_low):
        return None, None, "Bookandlink bukan voucher (v2)"
    if key == "TIKET" and re.search(r"request to change|permintaan (perubahan|ubah)", s_low):
        return None, None, "Tiket 'Request to change' = notifikasi (voucher/cancel final dikirim email terpisah)"
    if key == "TRIPCOM" and not re.search(r"pemesanan|booking no|reservation|dibatalkan|cancel", s_low):
        return None, None, "Trip.com bukan voucher"
    if key == "TRIPCOM" and "Reservation no" not in L.text and "No. pemesanan" not in L.text:
        return None, None, "Trip.com tanpa detail reservasi"

    if key == "BOOKANDLINK" and re.search(r"confirmation booking|booking confirmation", s_low):
        r = parse_bookandlink_web(L, subject)   # booking langsung dari website hotel (booking engine Bookandlink)
    else:
        r = PARSERS[key](L, subject)
    r["jenis"] = detect_jenis(subject, L.text)
    if key == "BOOKANDLINK" and r.get("_bookandlink_ota") in DIRECT_OTAS and r["jenis"] != "CANCEL":
        umur = _umur_jam(msg)
        if umur is not None and umur < BNL_TUNDA_JAM:
            return None, None, f"TUNDA: Bookandlink {r['_bookandlink_ota']} (cek lagi setelah {BNL_TUNDA_JAM:g} jam)"
        r["catatan"].append(BNL_TANDA + " (cadangan bila email langsung OTA tidak ada)")
    r["via_bookandlink"] = (key == "BOOKANDLINK")
    r.pop("_bookandlink_ota", None)
    r = finalize(r)
    r["unit"] = unit
    r["subjek"] = subject[:250]
    try:
        r["tgl_email"] = parsedate_to_datetime(msg.get("Date")).astimezone(dt.timezone(dt.timedelta(hours=7))).strftime("%Y-%m-%d %H:%M")
    except Exception:
        r["tgl_email"] = str(msg.get("Date", ""))
    r["message_id"] = (str(msg.get("Message-ID", "")) or "").strip() or f"{r['ota']}-{r['booking_id']}-{r['tgl_email']}"

    pdf_bytes = None
    if render_pdf:
        if key == "TIKET" and pdfs and len(pdfs[0][1] or b"") > 1000:
            pdf_bytes = pdfs[0][1]
        elif html or text:  # lampiran kosong/rusak -> cetak isi email
            body = html or "<pre style='font-family:Arial;white-space:pre-wrap'>" + (text or "").replace("<", "&lt;") + "</pre>"
            pdf_bytes = html_to_pdf(body, msg, cids)
    jenis_tag = {"BARU": "", "UBAH": "UBAH_", "CANCEL": "CANCEL BOOKING_"}.get(r["jenis"], "")
    fname = f"{jenis_tag}{safe_name(r['ota'].split(' ')[0], 12)}_{r['booking_id'] or 'NA'}_{safe_name(r['tamu']) or 'Tamu'}.pdf"
    fname = fname.replace(" ", "_")
    return r, pdf_bytes, fname


# ----------------------------------------------------------------------------
# Apps Script API
# ----------------------------------------------------------------------------
class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def _post_follow(url, data, hops=5):
    """POST ke Web App Apps Script dan ikuti redirect dengan benar:
    - redirect ke googleusercontent (hasil)  -> GET
    - redirect ke /macros/ (exec lain)       -> POST ulang (bukan GET, supaya tidak jatuh ke doGet)"""
    method_post = True
    for n in range(hops):
        req = urllib.request.Request(url, data=data if method_post else None,
                                     headers={"Content-Type": "application/json"} if method_post else {})
        t0 = time.time()
        try:
            with _opener.open(req, timeout=120) as resp:
                return resp.read().decode()
        except urllib.error.HTTPError as e:
            if e.code in (301, 302, 303, 307, 308) and e.headers.get("Location"):
                url = e.headers["Location"]
                method_post = "/macros/" in url and "googleusercontent" not in url
                continue
            # Diagnosa: di langkah mana (POST ke /exec atau GET hasil) Google menolak, berapa lama, dan apa pesannya
            raise RuntimeError(_ringkas_galat_http(e, n + 1, method_post, url, time.time() - t0)) from e
    raise RuntimeError("terlalu banyak redirect")


def _ringkas_galat_http(e, langkah, method_post, url, detik):
    isi = ""
    try:
        mentah = e.read(6000).decode("utf-8", "replace")
        mentah = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", mentah)
        isi = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", mentah)).strip()[:160]
    except Exception:
        pass
    try:
        host = url.split("/")[2]
    except Exception:
        host = "?"
    return (f"HTTP Error {e.code}: {e.reason} [langkah {langkah} {'POST' if method_post else 'GET'} {host}, "
            f"{detik:.0f} dtk{' | ' + isi if isi else ''}]")


_WAKTU_GAS = {}   # lama tiap jenis panggilan Apps Script (detik) -> diringkas di akhir, untuk mencari bagian yang lambat


def _catat_waktu_gas(action, detik):
    a = _WAKTU_GAS.setdefault(str(action), [0, 0.0])
    a[0] += 1
    a[1] += detik
    if detik >= 15:
        log(f"   ⏱ Apps Script '{action}' lambat: {detik:.0f} dtk")


def gas_post(payload, retries=5):
    url, token = os.environ["GAS_URL"], os.environ["GAS_TOKEN"]
    payload = dict(payload, token=token)
    data = json.dumps(payload).encode()
    last = None
    for i in range(retries):
        t0 = time.time()
        try:
            out = json.loads(_post_follow(url, data))
            if not out.get("ok"):
                raise RuntimeError(out.get("error", "unknown error"))
            _catat_waktu_gas(payload.get("action"), time.time() - t0)
            return out
        except Exception as e:
            last = e
            g = _WAKTU_GAS.setdefault("GAGAL " + str(payload.get("action")), [0, 0.0])   # waktu yang terbuang untuk percobaan gagal
            g[0] += 1
            g[1] += time.time() - t0
            log(f"   ! GAS gagal (percobaan {i + 1}, {payload.get('action')}):", _e(e))
            time.sleep(min(20, 4 * (i + 1)))
    raise RuntimeError(f"Apps Script gagal: {last}")


# ----------------------------------------------------------------------------
# Arrival List VHP (PDF di-upload FO ke folder tanggal)
# ----------------------------------------------------------------------------
ARR_ROW = re.compile(r"^\s*(\d{4,})\s+(.+?)\s+(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(\d+)\s+(\S+)\s+(\S+)\s+(.*)$")


# Departed Guest List: No ResNo VoucherNo Reservation-Name Guest-Name Email Mobile ... Nat Seg Arrival Depart Qty RmTyp RmNo ...
DEP_ROW = re.compile(r"^\s*\d{1,4}\s+(\d{5,})\s+(.*?)\s+(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(\d{1,3})\s+(\S+)\s+(\S+)\s*(.*)$")
# Arrival Guest List format Oak Tree: No [RmNo] ResNo [*] GuestName Segment Argt RmCat Ratecode RoomRate A/C ... Company/TA, CreateDate Arrival Departure ...
EA_ROW = re.compile(r"^\s*\d{1,3}\s+(?:(\S{1,5})\s+)?(\d{5,})\s+(?:\*\s+)?(.+?)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+([\d,]{3,}(?:\.\d{2})?)\s+\d+/\d+\s+"
                    r"(.*?)\s+(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\b(.*)$")
EA_ARGT = {"RO", "RB", "BB", "HB", "FB", "AI", "RBB", "ROB"}   # kolom Arrangement (room only / room breakfast ...)
EA_SEG = {"OTA", "FIT", "TA-FIT", "TA", "COR", "CORP", "GOV", "COMP", "HU", "GRP", "WALK", "WIG", "REG"}
VHP_SOURCES = re.compile(r"^(\[[^\]]+\],?|TRAVELOKA\s+TERA|TRAVELOKA|AGODA|BOOKING\.COM|TIKET\.COM|TIKET|C-TRIP|CTRIP|TRIP\.COM|"
                         r"EXPEDIA|BOOKNLINK|BOOKANDLINK)\s*", re.I)
VHP_REPORT = re.compile(r"Guest\s+Arrival\s+List|Resident\s+Guest\s+List|Departed\s+Guest\s+List|Guest\s+List|"
                        r"In\s*House|Expected\s+Arrival|\bResNo\b", re.I)


def _iso(s):
    return "20%s-%s-%s" % (s[6:8], s[3:5], s[0:2])


def _parse_departed_mid(mid):
    """'20261124509847 TRAVELOKA TERA ABDUL KODIR, YUSEP N/A ABW 8' -> (voucher, sumber, tamu)"""
    mid = re.sub(r"\s+[A-Z]{3}\s+\d+\s*$", "", mid.strip())                 # Nat + Segment
    voucher = ""
    m = re.match(r"(\d{6,})\s+(.*)$", mid)
    if m:
        voucher, mid = m.group(1), m.group(2)
    toks = [t for t in mid.split() if "@" not in t and t.upper() != "N/A" and not re.fullmatch(r"\+?[\d\-]{6,}", t)]
    mid = " ".join(toks)
    m = VHP_SOURCES.match(mid)
    if m:
        source, name = m.group(1).strip().strip(","), mid[m.end():]
    else:
        source, name = (mid.split(" ", 1) + [""])[:2]
    return voucher, source.strip(), name.strip().strip(",").strip()


def parse_arrival_pdf(data):
    """Baca laporan VHP (Arrival List / Guest In House / Departed Guest List) dari ISI file.
    Return (jenis_laporan, rows). jenis_laporan '' = bukan laporan VHP (mis. voucher TA)."""
    import pdfplumber
    best = ("", [])
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        # Font VHP tiap unit beda kerapatan: coba ekstraksi standar & rapat (x_tolerance=1), pakai yang barisnya terbanyak
        for tol in (None, 1):
            kw = {"layout": True} if tol is None else {"layout": True, "x_tolerance": tol}
            text = "".join((p.extract_text(**kw) or "") + "\n" for p in pdf.pages)
            res = _parse_vhp_text(text)
            if len(res[1]) > len(best[1]) or (not best[0] and res[0]):
                best = res
    return best


def _parse_vhp_text(text):
    rows = []
    head = text[:3000]
    if not VHP_REPORT.search(head):
        return "", []
    kind = ("departed" if re.search(r"Departed|Check\s*-?\s*Out|VoucherNo", head, re.I) else
            "arrival" if re.search(r"Arrival\s+Guest\s+List", head, re.I) else
            "inhouse" if re.search(r"Resident|In\s*House", head, re.I) else "arrival")
    for ln in text.splitlines():
        m = DEP_ROW.match(ln) if kind == "departed" else None
        if m:
            resno, mid, arr, dep, qty, rmtyp, rmno, rest = m.groups()
            voucher, source, name = _parse_departed_mid(mid)
            rows.append({"resno": resno, "tamu": name, "arrival": _iso(arr), "depart": _iso(dep), "tipe": rmtyp,
                         "kamar": rmno, "sumber": source, "catatan": voucher, "rate": None})
            continue
        m = ARR_ROW.match(ln)
        if not m:
            e = EA_ROW.match(ln)
            if e:   # format Arrival Guest List (Oak Tree)
                rmno, resno, name, seg, argt, rmcat, rcode, rate, mid, _cr, arr, dep, _tail = e.groups()
                # Ratecode bisa kosong (tamu walk-in/individual) atau 2 kata ("BAR 3") -> kolom bergeser.
                # Patokan: kolom Arrangement (RO/RB/...) -> sesudahnya RmCat, sebelumnya Segment, sisanya nama tamu.
                tok = " ".join([name, seg, argt, rmcat, rcode]).split()
                ai = [i for i, x in enumerate(tok) if i >= 2 and i + 1 < len(tok) and x in EA_ARGT]
                if ai:
                    seg_ok = [i for i in ai if tok[i - 1] in EA_SEG]
                    i = seg_ok[-1] if seg_ok else ai[-1]
                    name, seg, argt, rmcat = " ".join(tok[:i - 1]), tok[i - 1], tok[i], tok[i + 1]
                src = re.search(r"(\[[^\]]+\]|[^\s,]+),?\s*$", mid)
                if "[" in mid and "]" not in mid[mid.rfind("["):]:   # "[INDIVIDUAL RE" terpotong di kolom PDF
                    src = re.search(r"(\[.*)$", mid)
                rows.append({"resno": resno, "tamu": name.strip().strip(",").strip(), "arrival": _iso(arr), "depart": _iso(dep),
                             "tipe": rmcat, "kamar": rmno or "", "sumber": (src.group(1).strip() + ("]" if src.group(1).count("[") > src.group(1).count("]") else "")) if src else "",
                             "catatan": "", "rate": to_number(rate)})
            continue
        resno, name, arr, dep, qty, rmtyp, rmno, rest = m.groups()
        # Kolom sesudah RmTyp: [RmNo] A C C CO CC Stay Argt ... -> kamar yang BELUM diberi nomor hanya punya 6 angka
        # (kalau 7 angka, angka pertama = nomor kamar). Tanpa cek ini jumlah tamu "2" terbaca sebagai nomor kamar.
        nums = re.match(r"((?:\d+\s+)+)", rmno + " " + rest)
        if re.fullmatch(r"\d{1,2}", rmno) and nums and len(nums.group(1).split()) <= 6:
            rmno = ""
        m2 = re.search(r"([\d,]+\.\d{2})\s+(.*)$", rest)
        rate = to_number(m2.group(1)) if m2 else None
        after = m2.group(2) if m2 else rest
        # "<Company/TA> [Flight] ETA Kcard [Created-date] [By] Comment" — Created-date & By kadang kosong di laporan pagi
        m3 = re.match(r"(.*?)\s*\d{1,2}:\d{2}\s+\d+(?:\s+\d{2}/\d{2}/\d{2})?(?:\s+(.*))?$", after)
        source = (m3.group(1) if m3 else after).strip().strip(",").strip()
        comment = re.sub(r"^\*\*\s*", "", m3.group(2) or "").strip() if m3 else ""
        rows.append({"resno": resno, "tamu": name.strip(), "arrival": _iso(arr), "depart": _iso(dep),
                     "tipe": rmtyp, "kamar": rmno, "sumber": source, "catatan": comment, "rate": rate})
    return kind, rows


def arrival_sync(unit, dry_run=False):
    """Baca PDF yang di-upload FO di folder tanggal. Laporan VHP dikenali dari ISI file (nama file bebas)."""
    try:
        files = gas_post({"action": "arrival_pending", "unit": unit}).get("files", [])
    except Exception as e:
        log("   ~ cek laporan VHP gagal (dicoba lagi nanti):", _e(e))
        return
    saved = 0
    for f in files:
        try:
            kind, rows = parse_arrival_pdf(base64.b64decode(f["pdf_base64"]))
            log(f"   # {f['date']} {'(file)' if LOG_PUBLIK else f['name']}: " + (f"laporan VHP ({kind}), {len(rows)} tamu" if kind else "bukan laporan VHP"))
            if not dry_run:
                gas_post({"action": "arrival_save", "unit": unit, "date": f["date"], "file_id": f["file_id"],
                          "stamp": f["stamp"], "rows": rows, "is_report": bool(kind and rows), "kind": kind})
                saved += 1 if (kind and rows) else 0
        except Exception as e:
            log(f"   ! gagal baca {'(file)' if LOG_PUBLIK else f.get('name')}: {_e(e)}")
    if saved and not dry_run:
        try:
            out = gas_post({"action": "arrival_done"}, retries=2)
            if out.get("revisi"):
                log("   ✉ email REVISI dikirim:", (f"{len(out['revisi'])} unit/tanggal" if LOG_PUBLIK else ", ".join(out["revisi"])))
        except Exception as e:
            log("   ~ kirim revisi gagal:", _e(e))


# ----------------------------------------------------------------------------
# IMAP
# ----------------------------------------------------------------------------
def imap_uids(M, since, before=None):
    uids = set()
    rng = f'SINCE "{since}"' + (f' BEFORE "{before}"' if before else "")
    for dom in SENDERS:
        typ, data = M.uid("search", None, f'({rng} FROM "{dom}")')
        if typ == "OK" and data and data[0]:
            uids.update(data[0].split())
    return sorted(uids, key=int)


def fetch_header_ids(M, uids):
    out = {}
    for i in range(0, len(uids), 50):
        chunk = b",".join(uids[i:i + 50])
        typ, data = M.uid("fetch", chunk, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
        for item in data:
            if not isinstance(item, tuple):
                continue
            m = re.search(rb"UID (\d+)", item[0])
            hdr = email.message_from_bytes(item[1])
            if m:
                out[m.group(1)] = (hdr.get("Message-ID") or "").strip()
    return out


BF_STEP_DAYS = 10   # lebar jendela tarik-ulang otomatis email lama per jalan (kecil supaya tidak kena batas waktu job)
BF_MAX_MENIT = 12   # tarik-ulang otomatis dilewati bila unit ini sudah memakai waktu lebih dari ini pada jalan tsb
BF_STOP_MENIT = 15  # tarik-ulang yang sedang berjalan dijeda bila unit ini sudah memakai waktu lebih dari ini (job dibagi 4 unit)
_T0 = time.time()


def _tgl(d):
    return d.strftime("%d-%b-%Y")


def _proses_email(M, unit, since, before, backfill, dry_run, stats, min_ci):
    """Satu putaran baca email pada jendela [since, before). Return: jumlah error pada putaran ini.
    backfill=True (TARIK ULANG email lama): hanya voucher yang check-in-nya kemarin/hari ini/ke depan yang disimpan;
    email lain dilewati tanpa dicatat."""
    err0 = stats["error"]
    uids = imap_uids(M, _tgl(since), _tgl(before) if before else None)
    log(f"   {len(uids)} email dari OTA ditemukan")
    ids = fetch_header_ids(M, uids)

    seen, need_price = set(), set()
    if not dry_run and ids:
        out = gas_post({"action": "check", "unit": unit, "message_ids": [v for v in ids.values() if v]})
        seen = set(out.get("seen", []))
        if isinstance(out.get("bids"), list):
            _BIDS[unit] = set(str(b) for b in out["bids"])
        need_price = set(out.get("need_price", []))
    todo = [u for u in uids if not ids.get(u) or ids[u] not in seen]
    log(f"   {len(todo)} email baru untuk diproses")

    for k, uid in enumerate(todo):
        if not backfill and time.time() - _T0 > EMAIL_MAKS_MENIT * 60:
            log(f"   ⏸ batas {EMAIL_MAKS_MENIT:g} menit/unit: sisa {len(todo) - k} email dilanjut run berikutnya")
            stats["sisa"] = len(todo) - k
            return stats["error"] - err0
        if backfill and time.time() - _T0 > BF_STOP_MENIT * 60:   # login ulang otomatis -> tarik ulang bisa sangat lama: dibatasi
            log(f"   ⏸ tarik ulang dijeda (batas {BF_STOP_MENIT} menit/unit), sisa {len(todo) - k} email dilanjut run berikutnya")
            return -1   # belum selesai -> jendela ini diulang run berikutnya (email yang sudah tersimpan dilewati)
        try:
            typ, data = M.uid("fetch", uid, "(BODY.PEEK[])")
            raw = next(x[1] for x in data if isinstance(x, tuple))
            if backfill:   # baca dulu tanpa cetak PDF; yang check-in-nya sudah lewat / bukan voucher tidak disentuh
                rec, pdf, info = process_message(raw, unit, render_pdf=False)
                if rec is None or not rec.get("checkin") or str(rec["checkin"]) < min_ci:
                    stats["lama"] = stats.get("lama", 0) + 1
                    continue
                if not dry_run:
                    rec, pdf, info = process_message(raw, unit, render_pdf=True)
            else:
                rec, pdf, info = process_message(raw, unit, render_pdf=False)
                if rec is not None and not dry_run:
                    if BNL_TANDA in str(rec.get("catatan") or "") and _ada_booking(unit, rec.get("booking_id")):
                        rec, pdf, info = None, None, "Duplikat Bookandlink: booking ini sudah tercatat"
                    else:
                        rec, pdf, info = process_message(raw, unit, render_pdf=True)
            mid = ids.get(uid) or ""
            if rec is None and str(info).startswith("TUNDA"):   # belum dicatat -> dicek lagi run berikutnya
                stats["tunda"] = stats.get("tunda", 0) + 1
                continue
            if rec is None:
                log(f"   - lewati UID {uid.decode()}" + ("" if LOG_PUBLIK else f": {info}"))
                stats["skipped"] += 1
                if not dry_run and mid:
                    try:
                        gas_post({"action": "skip", "message_id": mid, "reason": info}, retries=2)
                    except Exception as e:  # tidak fatal: akan dicoba lagi di run berikutnya
                        log(f"   ~ gagal catat skip (dicoba lagi nanti): {_e(e)}")
                continue
            if LOG_PUBLIK:
                log(f"   + {rec['ota']} {rec['jenis']} #{_samar_id(rec['booking_id'])} | CI {rec['checkin']} CO {rec['checkout']}"
                    f" | harga {'ada' if rec.get('harga_jual') else 'kosong'} | {rec['status']}")
            else:
                log(f"   + {rec['ota']} {rec['jenis']} #{rec['booking_id']} {rec['tamu']} | CI {rec['checkin']} CO {rec['checkout']}"
                    f" | jual {rec['harga_jual']} kom {rec['komisi']} net {rec['net']} | {rec['status']}")
            if dry_run:
                continue
            gas_post({"action": "save", "record": rec, "filename": info,
                      "pdf_base64": base64.b64encode(pdf).decode() if pdf else ""})
            stats["saved"] += 1
            if unit in _BIDS and rec.get("booking_id") and rec.get("jenis") != "CANCEL":
                _BIDS[unit].add(str(rec["booking_id"]))   # salinan Bookandlink berikutnya di run ini langsung dikenali
        except Exception as e:
            stats["error"] += 1
            log(f"   ! ERROR UID {uid!r}: {_e(e)}")
    # Voucher lama yang belum punya catatan harga -> baca ulang email, kirim harganya saja
    for uid in [u for u in uids if ids.get(u) in need_price]:
        if not backfill and time.time() - _T0 > EMAIL_MAKS_MENIT * 60:
            break   # isi harga voucher lama tidak mendesak -> lanjut run berikutnya
        try:
            typ, data = M.uid("fetch", uid, "(BODY.PEEK[])")
            raw = next(x[1] for x in data if isinstance(x, tuple))
            rec, _, _ = process_message(raw, unit, render_pdf=False)
            if rec is None:
                continue
            gas_post({"action": "price", "message_id": ids[uid], "harga_jual": rec["harga_jual"], "komisi": rec["komisi"],
                      "net": rec["net"], "rate_malam1": rec.get("rate_malam1"),
                      "rate_per_malam": rec.get("rate_per_malam") or "", "kamar": rec.get("kamar") or ""}, retries=2)
            stats["price"] = stats.get("price", 0) + 1
        except Exception as e:
            log(f"   ~ gagal isi harga UID {uid!r}: {_e(e)}")
    return stats["error"] - err0


class _Imap:
    """Koneksi IMAP baca-saja yang LOGIN ULANG sendiri bila server memutus koneksi
    (server email unit memutus sesi panjang setelah +-15 menit -> 'socket error: EOF')."""

    def __init__(self, host, port, user, pw, folder):
        self._a = (host, port, user, pw, folder)
        self.M, self.ulang = None, 0
        self._buka()

    def _buka(self):
        host, port, user, pw, folder = self._a
        if self.M is not None:
            try:
                self.M.shutdown()
            except Exception:
                pass
        self.M = imaplib.IMAP4_SSL(host, port, timeout=120)
        self.M.login(user, pw)
        self.M.select(f'"{folder}"', readonly=True)  # READ-ONLY: email tidak berubah status

    def uid(self, *args):
        for coba in range(3):
            try:
                return self.M.uid(*args)
            except (imaplib.IMAP4.abort, OSError) as e:
                if coba == 2:
                    raise
                self.ulang += 1
                log(f"   ↻ koneksi email terputus ({_e(e)}) - login ulang & lanjut")
                time.sleep(3 * (coba + 1))
                self._buka()

    def logout(self):
        try:
            self.M.logout()
        except Exception:
            pass


def _pasang_batas_unit(unit, menit):
    """Unit macet (server email / Apps Script / Chromium tidak menjawab) -> dihentikan paksa setelah UNIT_MAKS_MENIT,
    supaya unit berikutnya tetap jalan. Email yang belum tersimpan otomatis diproses run berikutnya."""
    def _stop():
        log(f"   ⛔ {unit}: melewati batas {menit:g} menit -> dihentikan, dilanjut run berikutnya")
        sys.stdout.flush()
        os._exit(2)
    t = threading.Timer(max(menit * 60 - (time.time() - _T0), 1), _stop)
    t.daemon = True
    t.start()
    return t


def _vhp_semua_unit(dry_run):
    """Sekali per job (penanda file di runner): baca laporan VHP baru semua unit sebelum email diproses."""
    tanda = os.path.join(os.environ.get("RUNNER_TEMP") or "/tmp", f"hvr_vhp_{os.environ.get('GITHUB_RUN_ID') or 'lokal'}")
    if dry_run or os.path.exists(tanda):
        return
    try:
        open(tanda, "w").close()
    except OSError:
        pass
    for u in SEMUA_UNIT:
        log(f"== Laporan VHP lebih dulu | {u}")
        try:
            arrival_sync(u, dry_run)
        except Exception as e:
            log("   ~ laporan VHP gagal (dicoba lagi di step unitnya):", _e(e))


def run(dry_run=False):
    unit = os.environ.get("UNIT_NAME", "Verse Lite Gajah Mada")
    host = os.environ["IMAP_HOST"]
    port = int(os.environ.get("IMAP_PORT") or "993")
    folder = os.environ.get("IMAP_FOLDER") or "INBOX"
    today = dt.date.today()
    days = int(os.environ.get("LOOKBACK_DAYS") or "3")
    bdays = int(os.environ.get("BEFORE_DAYS") or "0")   # batas atas jendela (0 = sampai sekarang), untuk memecah per jendela
    manual = days > 7 or bdays > 0                      # tarik ulang yang diminta manual lewat input workflow

    # Status unit dari Apps Script: kapan terakhir berhasil jalan + sampai mana email lama sudah ditarik ulang
    state = {}
    if not dry_run:
        try:
            state = gas_post({"action": "state", "unit": unit}, retries=2).get("state") or {}
        except Exception as e:
            log("   ~ status unit tidak terbaca (pakai bawaan):", _e(e))
    # OTOMATIS 1 - sistem sempat berhenti: baca mundur sampai hari terakhir berhasil jalan (bukan hanya 3 hari)
    auto_ext = False
    if not manual and state.get("last_ok"):
        try:
            gap = (today - dt.date.fromisoformat(str(state["last_ok"])[:10])).days + 2
            if gap > days:
                days, auto_ext = min(gap, 120), True
        except ValueError:
            pass
    # TARIK ULANG EMAIL LAMA (lookback > 7 hari, atau > 14 hari bila diperpanjang otomatis): hanya voucher yang
    # check-in-nya maks. 15 hari lalu/ke depan yang disimpan; email lain dilewati tanpa dicatat.
    backfill = days > (14 if auto_ext else 7)
    # tarik ulang (manual / unit baru / sistem sempat berhenti) memang lama -> batasnya lebih longgar
    bf_aktif = str(state.get("bf_date") or "") > str(state.get("bf_target") or "") and bool(state.get("bf_target"))
    if not dry_run:
        _pasang_batas_unit(unit, BF_STOP_MENIT + 8 if (backfill or manual or bf_aktif) else UNIT_MAKS_MENIT)
        _vhp_semua_unit(dry_run)
    since = today - dt.timedelta(days=days)
    before = (today - dt.timedelta(days=bdays)) if bdays else None
    min_ci = (today - dt.timedelta(days=15)).isoformat()   # = REV_BACK_DAYS (14) + 1: voucher telat masih disalin ke folder Rev Updates

    log(f"== Hotel Voucher Recap | {unit} | {host} | sejak {_tgl(since)}" + (f" s/d {_tgl(before)}" if before else "") +
        f" | dry_run={dry_run}" + (f" | TARIK ULANG (check-in >= {min_ci})" if backfill else "") +
        (f" | diperpanjang otomatis: terakhir berhasil {state.get('last_ok')}" if auto_ext else ""))
    M = _Imap(host, port, os.environ["IMAP_USER"], os.environ["IMAP_PASS"], folder)   # login ulang otomatis bila terputus

    stats = {"saved": 0, "skipped": 0, "error": 0}
    _proses_email(M, unit, since, before, backfill, dry_run, stats, min_ci)

    # OTOMATIS 2 - unit baru: email lama (pesanan yang dibuat sebelum sistem jalan) ditarik ulang sendiri,
    # satu jendela per jalan, sampai tanggal target yang ditetapkan Apps Script.
    bf_new = ""
    bf_date, bf_target = str(state.get("bf_date") or ""), str(state.get("bf_target") or "")
    if (not dry_run and not manual and bf_date and bf_target and bf_date > bf_target
            and time.time() - _T0 < BF_MAX_MENIT * 60):
        try:
            d1 = dt.date.fromisoformat(bf_date)
            d0 = max(d1 - dt.timedelta(days=BF_STEP_DAYS), dt.date.fromisoformat(bf_target))
            log(f"   ↺ tarik ulang otomatis email lama {d0} s/d {d1} (berlanjut tiap jalan sampai {bf_target})")
            if _proses_email(M, unit, d0, d1 + dt.timedelta(days=1), True, dry_run, stats, min_ci) == 0:
                bf_new = d0.isoformat()
        except Exception as e:
            log("   ~ tarik ulang otomatis gagal (dicoba lagi run berikutnya):", _e(e))
    M.logout()
    if M.ulang:
        log(f"   ↻ login ulang email: {M.ulang}x")
    arrival_sync(unit, dry_run)
    if not dry_run:   # salin voucher tamu in-house ke folder Rev Updates (tim owner), nama file = nomor kamar
        try:          # sekaligus lapor: jalan ini berhasil (ok) + sampai mana email lama sudah ditarik (bf_date)
            rev = gas_post({"action": "rev_sync", "unit": unit, "ok": stats["error"] == 0 and bdays == 0,
                            "bf_date": bf_new}, retries=1).get("rev") or {}
            log(f"   > Rev Updates{' (simulasi)' if rev.get('simulasi') else ''}: {rev.get('disalin', 0)} disalin, "
                f"{rev.get('sudahAdaFO', 0)} sudah ada dari FO, {rev.get('gantiKamar', 0)} ganti kamar")
            for ln in ([] if LOG_PUBLIK else rev.get("baris", [])):
                log("     ·", ln)
        except Exception as e:
            log("   ~ Rev Updates gagal (dicoba lagi run berikutnya):", _e(e))
    if not dry_run and stats["saved"]:
        gas_post({"action": "finish"})
    if _WAKTU_GAS:
        log("   ⏱ waktu Apps Script: " + ", ".join(f"{k} {v[0]}x {v[1]:.0f} dtk" for k, v in sorted(_WAKTU_GAS.items(), key=lambda x: -x[1][1])))
    log(f"== Selesai: {stats}")
    if stats["error"]:
        sys.exit(1)


def test_files(paths):
    os.makedirs("test_output", exist_ok=True)
    for p in paths:
        with open(p, "rb") as f:
            raw = f.read()
        rec, pdf, info = process_message(raw, os.environ.get("UNIT_NAME", "Verse Lite Gajah Mada"))
        print("=" * 80, "\n", p)
        if rec is None:
            print("  DILEWATI:", info)
            continue
        print(json.dumps(rec, indent=2, ensure_ascii=False))
        print("  file:", info)
        if pdf:
            with open(os.path.join("test_output", info), "wb") as f:
                f.write(pdf)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        test_files(sys.argv[2:])
    else:
        run(dry_run="--dry-run" in sys.argv or os.environ.get("DRY_RUN") == "1")
