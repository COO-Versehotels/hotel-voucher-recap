#!/usr/bin/env python3
"""
Cetak ulang voucher tertentu (berdasarkan Message-ID) menjadi PDF dan simpan ke folder audit di Drive COO
lewat Apps Script (action 'audit_pdf'). TIDAK masuk Log Voucher / laporan harian. Mailbox dibaca READ-ONLY.
Daftar email: file JSON (AUDIT_IDS) berisi [{"mid": "<message-id>", "sub": "Juni 2026"}, ...].
Log publik hanya berisi angka.
"""
import base64, imaplib, json, os, sys, time
os.environ.setdefault("LOG_PUBLIK", "1")
os.environ.setdefault("PW_LAZY", "1")
import collector as C  # noqa: E402


def main():
    unit = os.environ.get("UNIT_NAME", "Verse Luxe Hotel Wahid Hasyim")
    ids = json.load(open(os.environ.get("AUDIT_IDS") or "audit_luxe_ids.json"))
    folder = os.environ.get("AUDIT_FOLDER") or "Audit All Inclusive Luxe"
    M = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"], int(os.environ.get("IMAP_PORT") or "993"))
    M.login(os.environ["IMAP_USER"], os.environ["IMAP_PASS"])
    M.select("INBOX", readonly=True)
    st = {"total": len(ids), "disimpan": 0, "tidak_ketemu": 0, "gagal": 0}
    t0 = time.time()
    for k, it in enumerate(ids):
        try:
            typ, data = M.uid("search", None, f'HEADER Message-ID "{it["mid"]}"')
            uids = data[0].split() if typ == "OK" and data and data[0] else []
            if not uids:
                st["tidak_ketemu"] += 1
                continue
            typ, data = M.uid("fetch", uids[-1], "(BODY.PEEK[])")
            raw = next(x[1] for x in data if isinstance(x, tuple))
            rec, pdf, fname = C.process_message(raw, unit, render_pdf=True)
            if rec is None or not pdf:
                st["gagal"] += 1
                continue
            C.gas_post({"action": "audit_pdf", "folder": folder, "sub": it.get("sub") or "",
                        "nama": fname, "pdf_base64": base64.b64encode(pdf).decode()}, retries=3)
            st["disimpan"] += 1
        except Exception as e:
            st["gagal"] += 1
            print("   ! gagal:", type(e).__name__, flush=True)
        if (k + 1) % 10 == 0:
            print(f"   ... {k + 1}/{len(ids)} ({(time.time() - t0) / 60:.1f} menit)", flush=True)
    try:
        M.logout()
    except Exception:
        pass
    print("== Selesai:", st, flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL:", type(e).__name__, flush=True)
        sys.exit(1)
