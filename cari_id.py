#!/usr/bin/env python3
"""Cari satu booking ID di SEMUA folder mailbox satu unit (READ-ONLY). Log hanya: folder, waktu email (WIB),
domain pengirim, subjek. Tidak ada isi email / nama tamu.
Pencarian: (1) SUBJECT di server, (2) BODY/TEXT di server, (3) baca penuh semua email OTA di rentang tanggal
CEK_SEJAK..CEK_SAMPAI lalu cari ID di isi (termasuk teks PDF) - karena pencarian isi di server tidak selalu andal.
Env: IMAP_*, BOOKING_ID, CEK_SEJAK (01-Oct-2026), CEK_SAMPAI (20-Nov-2026)"""
import datetime as dt, email, email.policy, email.utils, imaplib, os, re, sys

os.environ.setdefault("LOG_PUBLIK", "1")
import collector as C  # noqa: E402

BID = os.environ["BOOKING_ID"].strip()
SEJAK = os.environ.get("CEK_SEJAK") or "01-Oct-2026"
SAMPAI = os.environ.get("CEK_SAMPAI") or "20-Nov-2026"
WIB = dt.timezone(dt.timedelta(hours=7))


def info(msg):
    try:
        t = email.utils.parsedate_to_datetime(msg.get("Date")).astimezone(WIB).strftime("%d %b %Y %H:%M WIB")
    except Exception:
        t = str(msg.get("Date"))
    frm = str(msg.get("From") or "")
    dom = (re.search(r"@([\w.-]+)", frm) or [None, "?"])[1]
    return t, dom, str(msg.get("Subject") or "")[:120]


def main():
    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"), timeout=90)
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
    fl = []
    for ln in (M.list()[1] or []):
        s = ln.decode("utf-8", "replace") if isinstance(ln, bytes) else str(ln)
        m = re.match(r'\((.*?)\)\s+(?:"[^"]*"|NIL)\s+(.+)$', s)
        if not m or "\\Noselect" in m.group(1):
            continue
        n = m.group(2).strip()
        if n.startswith('"') and n.endswith('"'):
            n = n[1:-1]
        fl.append(n)
    print(f"== cari #{BID} di {len(fl)} folder: {fl}", flush=True)
    ketemu = {}
    for fo in fl:
        try:
            if M.select(f'"{fo}"', readonly=True)[0] != "OK":
                continue
        except Exception:
            continue
        uids = set()
        for kr in (f'(SUBJECT "{BID}")', f'(BODY "{BID}")', f'(TEXT "{BID}")'):
            try:
                typ, d = M.uid("search", None, kr)
                if typ == "OK" and d and d[0]:
                    uids.update(d[0].split())
            except Exception as e:
                print("   search gagal", fo, kr, type(e).__name__)
        cand = set()
        for dom in C.SENDERS:
            typ, d = M.uid("search", None, f'(SINCE "{SEJAK}" BEFORE "{SAMPAI}" FROM "{dom}")')
            if typ == "OK" and d and d[0]:
                cand.update(d[0].split())
        n_scan = 0
        for u in sorted(cand | uids, key=int):
            typ, data = M.uid("fetch", u, "(BODY.PEEK[])")
            raw = next((x[1] for x in data if isinstance(x, tuple)), None)
            if not raw:
                continue
            n_scan += 1
            msg = email.message_from_bytes(raw, policy=email.policy.default)
            html, text, pdfs, _ = C.get_parts(msg)
            isi = str(msg.get("Subject") or "") + "\n" + (html or "") + "\n" + (text or "")
            if BID not in isi:
                for _, b in pdfs:
                    try:
                        isi += C.pdf_text(b)
                    except Exception:
                        pass
            if BID in isi or u in uids:
                ketemu[(fo, u)] = info(msg) + (("isi" if BID in isi else "server-saja"),)
        print(f"   folder {fo}: {n_scan} email dibaca", flush=True)
    M.logout()
    print(f"== HASIL #{BID}: {len(ketemu)} email", flush=True)
    for (fo, u), (t, dom, subj, cara) in sorted(ketemu.items(), key=lambda x: x[1][0]):
        print(f"   [{fo}] {t} | dari {dom} | {subj} | cocok:{cara}", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, flush=True)
        sys.exit(1)
