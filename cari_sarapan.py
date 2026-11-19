#!/usr/bin/env python3
"""Cari voucher OTA satu unit yang MENGINAP pada rentang tanggal tertentu (default 7 - 31 Des 2026) dan catat
keterangan paket makannya (Breakfast / Room Only), untuk menghitung kamar RB/RBF vs RO per malam.

READ-ONLY: email tidak ditandai dibaca / dipindah / dihapus; tidak menyentuh Drive / Sheet.
Semua folder email dibaca (INBOX + Archive + subfolder; Spam/Trash/Sent/Draft dilewati), tanpa batas tanggal booking
(env EMAIL_SEJAK, default 01-Jan-2025). Yang menentukan = tanggal menginap.

Hasil (nama tamu) TIDAK ditulis ke log publik: JSON terenkripsi (laporan_pub.pem) diunggah sebagai artifact 1 hari.
Log publik hanya berisi jumlah.

Env: IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASS, UNIT_NAME, INAP_DARI (2026-12-07), INAP_SAMPAI (2026-12-31, malam
terakhir yang dihitung), EMAIL_SEJAK (01-Jan-2025)"""
import datetime as dt
import email
import email.policy
import imaplib
import json
import os
import re
import sys
import time

os.environ.setdefault("LOG_PUBLIK", "1")
import collector as C  # noqa: E402  (parser voucher yang sama dengan sistem harian)

SKIP = re.compile(r"(^|[./])(spam|junk|trash|sampah|sent|drafts?|terkirim|deleted|outbox)([./]|$)", re.I)
MAKAN_RE = re.compile(r"breakfast|sarapan|room\s*only|tanpa\s*sarapan|\bmeal|makan|\bRB\b|\bRBF\b|\bRO\b|\bBB\b|"
                      r"benefit|fasilitas|termasuk|includ|rate\s*plan|rate\s*name|paket|package|board", re.I)
T0 = time.time()


def log(*a):
    print(f"[{(time.time() - T0) / 60:5.1f}m]", *a, flush=True)


def folders(M):
    out = []
    for ln in (M.list()[1] or []):
        s = ln.decode("utf-8", "replace") if isinstance(ln, bytes) else str(ln)
        m = re.match(r'\((.*?)\)\s+(?:"[^"]*"|NIL)\s+(.+)$', s)
        if not m or "\\Noselect" in m.group(1):
            continue
        nama = m.group(2).strip()
        if nama.startswith('"') and nama.endswith('"'):
            nama = nama[1:-1].replace('\\"', '"')
        if SKIP.search(nama):
            continue
        out.append(nama)
    return out


def baris_makan(lines):
    info = []
    for i, ln in enumerate(lines):
        if MAKAN_RE.search(ln):
            s = " | ".join(x.strip() for x in lines[max(0, i - 1):i + 2])[:300]
            if s not in info:
                info.append(s)
        if len(info) >= 25:
            break
    return info


def main():
    unit = os.environ.get("UNIT_NAME", "Verse Lite Gajah Mada")
    dari = dt.date.fromisoformat(os.environ.get("INAP_DARI") or "2026-12-07")
    sampai = dt.date.fromisoformat(os.environ.get("INAP_SAMPAI") or "2026-12-31")
    sejak = os.environ.get("EMAIL_SEJAK") or "01-Jan-2025"
    host, port = os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993")

    def sambung():
        M = imaplib.IMAP4_SSL(host, port, timeout=90)
        M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
        return M

    M = sambung()
    fl = folders(M)
    log(f"== {unit} | menginap {dari} s/d {sampai} | email sejak {sejak} | {len(fl)} folder")
    hasil, mids = [], set()
    st = {"email_ota": 0, "voucher": 0, "dalam_rentang": 0, "ubah_cancel_lain": 0, "error": 0}
    for no, fo in enumerate(fl, 1):
        try:
            if M.select(f'"{fo}"', readonly=True)[0] != "OK":
                continue
        except Exception:
            continue
        uids = set()
        for dom in C.SENDERS:
            typ, d = M.uid("search", None, f'(SINCE "{sejak}" FROM "{dom}")')
            if typ == "OK" and d and d[0]:
                uids.update(d[0].split())
        uids = sorted(uids, key=int)
        log(f"   folder #{no}: {len(uids)} email OTA")
        for k, u in enumerate(uids):
            try:
                typ, data = M.uid("fetch", u, "(BODY.PEEK[])")
                raw = next((x[1] for x in data if isinstance(x, tuple)), None)
                if not raw:
                    continue
                st["email_ota"] += 1
                rec, _, alasan = C.process_message(raw, unit, render_pdf=False)
                if not rec:
                    continue
                mid = rec.get("message_id") or ""
                if mid and mid in mids:
                    continue
                mids.add(mid)
                st["voucher"] += 1
                ci = rec.get("checkin") or ""
                co = rec.get("checkout") or ""
                try:
                    dalam = dt.date.fromisoformat(ci) <= sampai and dt.date.fromisoformat(co) > dari
                except ValueError:
                    dalam = False
                if not dalam and rec.get("jenis") not in ("CANCEL", "UBAH"):
                    continue
                msg = email.message_from_bytes(raw, policy=email.policy.default)
                key = C.sender_key(str(msg.get("From", "") or ""))
                html, text, pdfs, _ = C.get_parts(msg)
                lines = C.html_to_lines(html) if html else C.text_to_lines(text)
                if key == "TIKET" and pdfs:
                    pt = C.pdf_text(pdfs[0][1])
                    if pt.strip():
                        lines = lines + C.text_to_lines(pt)
                for kk in ("harga_jual", "komisi", "pajak_komisi", "net", "rate_malam1"):
                    if isinstance(rec.get(kk), float):
                        rec[kk] = round(rec[kk], 2)
                rec["_key"] = key
                rec["_dalam"] = dalam
                rec["_makan"] = baris_makan(lines)
                rec["_folder"] = no
                hasil.append(rec)
                st["dalam_rentang" if dalam else "ubah_cancel_lain"] += 1
            except Exception as e:
                st["error"] += 1
                log("   ! error:", type(e).__name__)
                try:
                    M.noop()
                except Exception:
                    M = sambung()
                    M.select(f'"{fo}"', readonly=True)
            if k and k % 300 == 0:
                log(f"   ... {k}/{len(uids)}")
    try:
        M.logout()
    except Exception:
        pass
    # ubah/cancel di luar rentang: simpan hanya bila booking ID-nya ada di rentang (bisa jadi tanggal digeser keluar)
    bid = {str(r.get("booking_id")) for r in hasil if r["_dalam"] and r.get("booking_id")}
    hasil = [r for r in hasil if r["_dalam"] or str(r.get("booking_id") or "") in bid]
    log(f"== {unit}: {len(bid)} booking menginap di rentang | {st}")
    os.makedirs("hasil", exist_ok=True)
    with open(f"hasil/{re.sub(r'[^A-Za-z]', '_', unit)}.json", "w") as f:
        json.dump(hasil, f, ensure_ascii=False, default=str)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, flush=True)
        sys.exit(1)
