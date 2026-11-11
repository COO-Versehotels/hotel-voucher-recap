#!/usr/bin/env python3
"""
Cari voucher PAKET MAKAN (All Inclusive / Full Board / Half Board / Lunch / Dinner) di mailbox FO satu unit.
Dipakai sekali-sekali lewat workflow "Cari Voucher Paket Makan" (manual). READ-ONLY: email tidak ditandai dibaca,
tidak dipindah, tidak dihapus. Tidak menyentuh Drive / Sheet rekap voucher harian.

Repo public -> log bisa dilihat siapa saja. Karena itu hasil (data tamu) DIENKRIPSI dengan kunci publik PAKET_PUBKEY
dan dicetak sebagai teks acak; hanya pemegang kunci privat yang bisa membacanya. Log biasa hanya berisi angka.

Env: IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASS, UNIT_NAME, PAKET_PUBKEY,
     EMAIL_DARI (default 2026-01-01), EMAIL_SAMPAI (default 2026-11-01, eksklusif),
     CEK_DARI / CEK_SAMPAI (jendela uji: semua email dibaca penuh untuk membandingkan dgn pencarian server)
"""
import base64
import datetime as dt
import email
import email.policy
import email.utils
import imaplib
import json
import os
import re
import sys
import time
import zlib

os.environ.setdefault("LOG_PUBLIK", "1")
import collector as C  # noqa: E402  (parser voucher yang sama dengan sistem harian)

MEAL_RE = re.compile(r"all[\s-]*inclusive(?!\s*(?:of|tax|pajak))|full[\s-]*board|half[\s-]*board|\blunch(?:es)?\b|"
                     r"\bdinner\b|makan\s*siang|makan\s*malam|3\s*x\s*makan", re.I)
INFO_RE = re.compile(r"struktur harga|rate plan|rate name|meal|benefit|termasuk|includ|breakfast|sarapan|paket|package", re.I)
UBAH_RE = re.compile(r"cancel|batal|pembatalan|modif|amend|change|ubah|perubahan|revis", re.I)
LEWATI_FOLDER = re.compile(r"trash|junk|spam|sent|draft|deleted|terkirim|sampah|outbox", re.I)
BODY_KATA = ["inclusive", "lunch", "dinner", "full board", "half board", "makan siang", "makan malam"]
T0 = time.time()


def log(*a):
    print(f"[{(time.time() - T0) / 60:5.1f}m]", *a, flush=True)


def tgl(s):
    return dt.date.fromisoformat(s).strftime("%d-%b-%Y")


def teks_email(msg, key):
    html, text, pdfs, _ = C.get_parts(msg)
    lines = C.html_to_lines(html) if html else C.text_to_lines(text)
    if key == "TIKET" and pdfs:
        pt = C.pdf_text(pdfs[0][1])
        if pt.strip():
            lines = lines + C.text_to_lines(pt)
    return lines


def baca(raw, unit):
    """raw -> (rec | None, ada_paket_makan, potongan baris info)"""
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    key = C.sender_key(str(msg.get("From", "") or ""))
    if not key:
        return None, False, []
    try:
        lines = teks_email(msg, key)
    except Exception:
        return None, False, []
    txt = "\n".join(lines)
    ada = bool(MEAL_RE.search(txt))
    info = []
    for i, ln in enumerate(lines):
        if MEAL_RE.search(ln) or INFO_RE.search(ln):
            s = ln if len(ln) > 3 else ln + " " + (lines[i + 1] if i + 1 < len(lines) else "")
            s = s[:220]
            if s not in info:
                info.append(s)
        if len(info) >= 12:
            break
    try:
        rec, _, alasan = C.process_message(raw, unit, render_pdf=False)
    except Exception as e:
        rec, alasan = None, "parse error " + type(e).__name__
    if rec is None:
        rec = {"_lewati": str(alasan)[:120], "subjek": str(msg.get("Subject", ""))[:200], "ota": key}
        try:
            rec["tgl_email"] = email.utils.parsedate_to_datetime(msg.get("Date")).strftime("%Y-%m-%d %H:%M")
        except Exception:
            rec["tgl_email"] = ""
    for k in ("harga_jual", "komisi", "pajak_komisi", "net", "rate_malam1"):
        if isinstance(rec.get(k), float):
            rec[k] = round(rec[k], 2)
    rec["_key"] = key
    return rec, ada, info


def cari(M, kriteria):
    typ, data = M.uid("search", None, kriteria)
    return data[0].split() if typ == "OK" and data and data[0] else []


def ambil(M, uid):
    typ, data = M.uid("fetch", uid, "(BODY.PEEK[])")
    return next(x[1] for x in data if isinstance(x, tuple))


def main():
    unit = os.environ.get("UNIT_NAME", "Verse Luxe Hotel Wahid Hasyim")
    dari, sampai = os.environ.get("EMAIL_DARI") or "2026-01-01", os.environ.get("EMAIL_SAMPAI") or "2026-11-01"
    cek_dari, cek_sampai = os.environ.get("CEK_DARI") or "", os.environ.get("CEK_SAMPAI") or ""
    pub = os.environ["PAKET_PUBKEY"]
    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"))
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
    folders = []
    typ, data = M.list()
    for ln in data or []:
        s = ln.decode(errors="replace")
        m = re.search(r'\) "?[^"]*"? (.+)$', s)
        nama = (m.group(1) if m else s).strip()
        if "\\Noselect" in s or LEWATI_FOLDER.search(nama):
            continue
        folders.append(nama)
    log(f"== {unit} | email {dari} s/d {sampai} | {len(folders)} folder dibaca")

    hasil, mids = [], set()
    stat = {"paket": 0, "ubah_cancel": 0, "uji_email": 0, "uji_paket": 0, "uji_terlewat": 0, "error": 0}
    rng = f'SINCE "{tgl(dari)}" BEFORE "{tgl(sampai)}"'
    for fo in folders:
        try:
            typ, _ = M.select(fo if fo.startswith('"') else f'"{fo}"', readonly=True)
            if typ != "OK":
                continue
        except Exception:
            continue
        # 1) pencarian di server: email OTA yang memuat kata paket makan
        kata = " ".join(["OR"] * (len(BODY_KATA) - 1)) + " " + " ".join(f'BODY "{k}"' for k in BODY_KATA)
        cocok = set()
        for dom in C.SENDERS:
            cocok.update(cari(M, f'({rng} FROM "{dom}" {kata})'))
        # 2) email cancel / ubah (subjek) -> versi terakhir booking paket makan
        ubah = set()
        for dom in C.SENDERS:
            for w in ("cancel", "batal", "modif", "amend", "ubah", "perubahan", "change"):
                ubah.update(cari(M, f'({rng} FROM "{dom}" SUBJECT "{w}")'))
        # 3) jendela uji: SEMUA email OTA dibaca penuh -> cek apakah pencarian server ada yang terlewat
        uji = set()
        if cek_dari and cek_sampai:
            for dom in C.SENDERS:
                uji.update(cari(M, f'(SINCE "{tgl(cek_dari)}" BEFORE "{tgl(cek_sampai)}" FROM "{dom}")'))
        log(f"   folder #{folders.index(fo) + 1}: {len(cocok)} kandidat paket, {len(ubah)} ubah/cancel, {len(uji)} email uji")
        semua = sorted(cocok | ubah | uji, key=int)
        for k, uid in enumerate(semua):
            try:
                raw = ambil(M, uid)
                rec, ada, info = baca(raw, unit)
                if rec is None:
                    continue
                mid = rec.get("message_id") or ""
                if uid in uji:
                    stat["uji_email"] += 1
                    if ada:
                        stat["uji_paket"] += 1
                        if uid not in cocok:
                            stat["uji_terlewat"] += 1
                if mid and mid in mids:
                    continue
                if ada or uid in ubah:
                    if mid:
                        mids.add(mid)
                    rec["_paket"] = ada
                    rec["_info"] = info if ada else []
                    rec["_folder"] = folders.index(fo) + 1
                    hasil.append(rec)
                    stat["paket" if ada else "ubah_cancel"] += 1
            except Exception as e:
                stat["error"] += 1
                log("   ! error:", type(e).__name__)
                try:
                    M.noop()
                except Exception:
                    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"))
                    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
                    M.select(fo if fo.startswith('"') else f'"{fo}"', readonly=True)
            if k and k % 200 == 0:
                log(f"   ... {k}/{len(semua)} email dibaca")
    try:
        M.logout()
    except Exception:
        pass

    # Hanya ubah/cancel yang booking ID-nya termasuk booking paket makan yang disimpan
    bid_paket = {str(r.get("booking_id")) for r in hasil if r.get("_paket") and r.get("booking_id")}
    hasil = [r for r in hasil if r.get("_paket") or str(r.get("booking_id") or "") in bid_paket]
    log(f"== Ringkasan: {len(bid_paket)} booking paket makan | {stat}")

    from nacl.public import PublicKey, SealedBox
    blob = SealedBox(PublicKey(base64.b64decode(pub))).encrypt(
        zlib.compress(json.dumps(hasil, ensure_ascii=False, default=str).encode(), 9))
    b64 = base64.b64encode(blob).decode()
    print("=====HASIL-TERENKRIPSI-MULAI=====")
    for i in range(0, len(b64), 3000):
        print(b64[i:i + 3000])
    print("=====HASIL-TERENKRIPSI-SELESAI=====", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, flush=True)
        sys.exit(1)
