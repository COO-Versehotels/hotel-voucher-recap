#!/usr/bin/env python3
"""Cari voucher Agoda yang memakai promo tertentu (default "Mega Sale") di semua folder email voucher unit
(INBOX + Archive + subfolder; Spam/Trash/Sent dilewati). READ-ONLY: tidak menyimpan apa pun ke Drive/Sheet.

Hasil (nama tamu, harga) TIDAK ditulis ke log publik: disimpan sebagai JSON terenkripsi (kunci publik
laporan_pub.pem, hanya bisa dibuka pemegang kunci privat) lalu diunggah sebagai artifact 1 hari.
Log publik hanya berisi jumlah."""
import email, email.policy, email.utils, imaplib, json, os, re, sys

import collector as C

PROMO = os.environ.get("PROMO", "Mega Sale")
SINCE = os.environ.get("SINCE", "01-Nov-2026")
SKIP = re.compile(r"(^|[./])(spam|junk|trash|sampah|sent|drafts?|terkirim)([./]|$)", re.I)
PROMO_RE = re.compile(r"([A-Za-z][A-Za-z0-9 &'+./]{1,40}?)\s*-\s*(\d{1,2} [A-Za-z]{3} \d{4})\s*([\d.]+%)")


def folders(M):
    out = []
    for ln in (M.list()[1] or []):
        s = ln.decode("utf-8", "replace") if isinstance(ln, bytes) else str(ln)
        m = re.match(r'\((.*?)\)\s+(?:"[^"]*"|NIL)\s+(.+)$', s)
        if not m or "\\Noselect" in m.group(1):
            continue
        nama = m.group(2).strip().strip('"')
        if SKIP.search(nama):
            continue
        out.append(nama)
    return out


def promos(text):
    hasil = []
    for m in PROMO_RE.finditer(text):
        nama = re.sub(r"^.*?(?:Last-Minute -|IDR [\d.,-]+)\s*", "", m.group(1)).strip()
        nama = re.sub(r"^[\d.]+%\s*-?\s*\d{0,2} ?[A-Za-z]{0,3} ?\d{0,4}\s*", "", nama).strip()
        hasil.append({"nama": nama, "tgl": m.group(2), "persen": m.group(3)})
    return hasil


def main():
    unit = os.environ.get("UNIT_NAME", "")
    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"), timeout=60)
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
    rows, total, kena = [], 0, 0
    seen = set()
    for fo in folders(M):
        try:
            if M.select(f'"{fo}"', readonly=True)[0] != "OK":
                continue
        except Exception:
            continue
        typ, d = M.uid("search", None, f'(SINCE "{SINCE}" FROM "agoda")')
        uids = d[0].split() if typ == "OK" and d and d[0] else []
        n_f = 0
        for u in uids:
            typ, data = M.uid("fetch", u, "(BODY.PEEK[])")
            raw = next((x[1] for x in data if isinstance(x, tuple)), None)
            if not raw:
                continue
            msg = email.message_from_bytes(raw, policy=email.policy.default)
            mid = (msg.get("Message-ID") or "").strip()
            if mid and mid in seen:
                continue
            seen.add(mid)
            try:
                rec, _, alasan = C.process_message(raw, unit, render_pdf=False)
            except Exception as e:
                rec, alasan = None, type(e).__name__
            if not rec:
                continue
            total += 1
            html, text, _, _ = C.get_parts(msg)
            lines = C.html_to_lines(html) if html else C.text_to_lines(text)
            t = "\n".join(lines)
            pr = promos(t)
            ada = PROMO.lower() in t.lower()
            if not ada:
                continue
            kena += 1
            n_f += 1
            try:
                tgl = email.utils.parsedate_to_datetime(msg.get("Date")).astimezone(
                    __import__("datetime").timezone(__import__("datetime").timedelta(hours=7))).strftime("%Y-%m-%d %H:%M")
            except Exception:
                tgl = ""
            rows.append({"unit": unit, "folder": fo, "email_wib": tgl, "subject": str(msg.get("Subject") or ""),
                         "jenis": rec.get("jenis"), "booking_id": rec.get("booking_id"), "tamu": rec.get("tamu"),
                         "checkin": rec.get("checkin"), "checkout": rec.get("checkout"), "tipe_kamar": rec.get("tipe_kamar"),
                         "kamar": rec.get("kamar"), "pembayaran": rec.get("pembayaran"), "harga_jual": rec.get("harga_jual"),
                         "komisi": rec.get("komisi"), "pajak_komisi": rec.get("pajak_komisi"), "net": rec.get("net"),
                         "rate_per_malam": rec.get("rate_per_malam"), "promos": pr,
                         "promo_baris": [ln for ln in lines if PROMO.lower() in ln.lower()][:3]})
        print(f"   folder #{len(seen)}: {len(uids)} email Agoda, {n_f} memuat promo", flush=True)
    M.logout()
    print(f"== {unit}: {total} voucher Agoda sejak {SINCE}, {kena} memuat '{PROMO}'", flush=True)
    os.makedirs("hasil", exist_ok=True)
    with open(f"hasil/{re.sub(r'[^A-Za-z]', '_', unit)}.json", "w") as f:
        json.dump(rows, f, ensure_ascii=False)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, str(e)[:120].replace(os.environ.get("IMAP_USER", "~"), "***"), flush=True)
        sys.exit(1)
