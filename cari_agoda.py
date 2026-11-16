#!/usr/bin/env python3
"""Cari email konfirmasi Agoda untuk nomor booking tertentu di SEMUA folder mailbox (termasuk Spam/Junk).
READ-ONLY. Log publik: nomor booking disamarkan (4 digit terakhir), tanpa nama tamu."""
import email, email.utils, imaplib, os, re, sys

def main():
    ids = [x.strip() for x in (os.environ.get("AGODA_IDS") or "").split(",") if x.strip()]
    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"))
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
    folders = []
    for ln in (M.list()[1] or []):
        s = ln.decode(errors="replace")
        if "\\Noselect" in s:
            continue
        m = re.search(r'\) "?[^"]*"? (.+)$', s)
        folders.append((m.group(1) if m else s).strip())
    print(f"== {os.environ.get('UNIT_NAME')} | {len(ids)} booking | folder: {', '.join(f.strip(chr(34)) for f in folders)}", flush=True)
    hasil = {b: [] for b in ids}
    for fo in folders:
        q = fo if fo.startswith('"') else f'"{fo}"'
        try:
            if M.select(q, readonly=True)[0] != "OK":
                continue
        except Exception:
            continue
        typ, d = M.uid("search", None, '(SINCE "01-Oct-2026" FROM "agoda.com")')
        n_ag = len(d[0].split()) if typ == "OK" and d and d[0] else 0
        ketemu = 0
        for b in ids:
            typ, d = M.uid("search", None, f'(SUBJECT "{b}")')
            uids = d[0].split() if typ == "OK" and d and d[0] else []
            for u in uids:
                typ, h = M.uid("fetch", u, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
                raw = next((x[1] for x in h if isinstance(x, tuple)), b"")
                msg = email.message_from_bytes(raw)
                frm = (msg.get("From") or "").lower()
                src = "agoda" if "agoda" in frm else ("bookandlink" if "bookandlink" in frm else "lain")
                sub = (msg.get("Subject") or "").lower()
                jenis = "batal" if re.search(r"batal|cancel", sub) else ("ubah" if re.search(r"ubah|amend|modif", sub) else "konfirmasi")
                try:
                    tgl = email.utils.parsedate_to_datetime(msg.get("Date")).strftime("%d/%m %H:%M")
                except Exception:
                    tgl = "?"
                hasil[b].append(f"{fo.strip(chr(34))}:{src}:{jenis}:{tgl}")
                ketemu += 1
        print(f"   folder {fo.strip(chr(34))}: {n_ag} email Agoda sejak 1 Okt, {ketemu} email cocok nomor booking", flush=True)
    M.logout()
    tidak = 0
    for b in ids:
        ag = [x for x in hasil[b] if ":agoda:" in x]
        if not ag:
            tidak += 1
        print(f"   #…{b[-4:]}: " + ("; ".join(hasil[b]) if hasil[b] else "tidak ada email sama sekali"), flush=True)
    print(f"== Ringkas: {len(ids) - tidak} booking ada email Agoda-nya, {tidak} tidak ada email dari Agoda", flush=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, flush=True)
        sys.exit(1)
