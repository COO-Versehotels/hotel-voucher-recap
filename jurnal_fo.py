"""Pembaca Jurnal FO VHP (F/O Journal Transaction). Dikenali dari ISI file, nama file bebas."""
import io, re, json, sys
import pdfplumber

JURNAL_RE = re.compile(r"F\s*/\s*O\s+Journal\s+Transaction", re.I)
HDR = ["Date", "RmNo", "Bill-No", "ArtNo", "Description", "Voucher", "Department", "Qty", "Amount", "Guest",
       "Bill", "SegmentCode", "Check", "Time"]
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
NUM_RE = re.compile(r"^-?[\d,]+\.\d{2}$")


def _iso(s):
    return "20%s-%s-%s" % (s[6:8], s[3:5], s[0:2])


def _num(s):
    return float(s.replace(",", ""))


def _lines(words):
    rows = {}
    for w in words:
        k = round(w["top"] / 2.0)
        rows.setdefault(k, []).append(w)
    out, last = [], None
    for k in sorted(rows):
        ws = sorted(rows[k], key=lambda w: w["x0"])
        if last is not None and k - last <= 1:
            out[-1] = sorted(out[-1] + ws, key=lambda w: w["x0"])
        else:
            out.append(ws)
        last = k
    return out


SRC_RE = re.compile(r"(\[[^\]]*\]?|TRAVELOKA\s+TERA|TRAVELOKA|AGODA|BOOKING\.COM|TIKET\.COM|C[\s-]?TRIP(?:,\s*T&T)?|TRIP\.COM|"
                    r"EXPEDIA|BOOKNLINK|BOOKANDLINK|MG\s+HOLIDAY|MGHOLIDAY|OTA)\s*,?\s*$", re.I)


def _row(ln, col):
    tk = [(w["text"], w["x0"]) for w in ln]
    i, rm = 1, []
    while i < len(tk) and not re.fullmatch(r"\d{5,8}", tk[i][0]):
        rm.append(tk[i][0]); i += 1
    if i >= len(tk):
        return None
    bill = tk[i][0]; i += 1
    art = ""
    if i < len(tk) and re.fullmatch(r"\d{4}", tk[i][0]):
        art = tk[i][0]; i += 1
    j = next((k for k in range(i, len(tk)) if NUM_RE.match(tk[k][0])), None)
    if j is None:
        return None
    amount = _num(tk[j][0])
    k = j
    qty = 1
    if k - 1 >= i and re.fullmatch(r"-?\d+", tk[k - 1][0]):
        qty = int(tk[k - 1][0]); k -= 1
    mid = tk[i:k]
    dx, vx = col.get("Department", 9e9) - 12, col.get("Voucher", 9e9) - 10
    dept = [t for t, x in mid if x >= dx]
    if dept and dept[0] != "Front" and "Front" in [t for t, x in mid]:
        pass
    dv = [(t, x) for t, x in mid if x < dx]
    if dv and dv[-1][0] == "Front":   # 'Front' kadang masuk kolom voucher
        dept = ["Front"] + dept; dv = dv[:-1]
    desc = [t for t, x in dv if x < vx]
    vou = [t for t, x in dv if x >= vx]
    rest = tk[j + 1:]
    d = next((k for k in range(len(rest)) if DATE_RE.match(rest[k][0])), len(rest))
    nm, tail = rest[:d], rest[d:]
    sx = col.get("SegmentCode", 9e9) - 12
    seg = [t for t, x in nm if x >= sx]
    nm = [(t, x) for t, x in nm if x < sx]
    txt = " ".join(t for t, x in nm)
    t2 = re.sub(r"[\s,]+OTA\s*$", "", txt)
    m = SRC_RE.search(t2) if t2 != txt else None
    if m and m.start() > 0:
        txt = t2
    else:
        m = SRC_RE.search(txt)
    if m and m.start() > 0:
        tamu, recv = txt[:m.start()], m.group(0)
    else:
        bx = col.get("Bill", 9e9) + 5
        tamu = " ".join(t for t, x in nm if x < bx); recv = " ".join(t for t, x in nm if x >= bx)
    dts = [t for t, x in tail if DATE_RE.match(t)]
    oth = [t for t, x in tail if not DATE_RE.match(t)]
    tm = next((t for t in oth if re.fullmatch(r"\d{2}:\d{2}:\d{2}", t)), "")
    aft = oth[oth.index(tm) + 1:] if tm else oth
    return {"tgl": _iso(tk[0][0]), "kamar": " ".join(rm), "bill": bill, "art": art,
            "desc": " ".join(desc), "voucher": " ".join(vou), "dept": " ".join(dept), "qty": qty, "amount": amount,
            "tamu": tamu.strip().strip(",").strip(), "penerima": recv.strip().strip(",").strip(),
            "segmen": " ".join(seg).strip(", "),
            "ci": _iso(dts[0]) if dts else "", "co": _iso(dts[1]) if len(dts) > 1 else "",
            "jam": tm, "user": aft[0] if aft else "", "remark": " ".join(aft[1:])}


def parse_jurnal_pdf(data):
    """Jurnal FO VHP (F/O Journal Transaction). Dikenali dari ISI file. Return None bila bukan jurnal."""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        first = pdf.pages[0].extract_text() or ""
        if not JURNAL_RE.search(first[:1500]):
            return None
        hotel = first.splitlines()[0]
        hotel = re.sub(r"\s+Page\s*:.*$", "", hotel).strip()
        mp = re.search(r"Period\s*:\s*(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})", first)
        dari, sampai = (_iso(mp.group(1)), _iso(mp.group(2))) if mp else ("", "")
        rows, grand = [], None
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
            lines = _lines(words)
            col = None
            for ln in lines:
                t = [w["text"] for w in ln]
                if col is None:
                    if len(t) > 10 and t[0] == "Date" and "ArtNo" in t:
                        col = {}
                        for w in ln:
                            if w["text"] in HDR and w["text"] not in col:
                                col[w["text"]] = w["x0"]
                    continue
                if t and t[0] == "Grand" and "TOTAL" in t:
                    nums = [x for x in t if NUM_RE.match(x)]
                    if nums: grand = _num(nums[-1])
                    continue
                if not t or not DATE_RE.match(t[0]) or ln[0]["x0"] > col["RmNo"] - 5:
                    continue
                r = _row(ln, col)
                if r: rows.append(r)
        return {"hotel": hotel, "dari": dari, "sampai": sampai, "rows": rows, "grand": grand}


if __name__ == "__main__":   # uji lokal: python jurnal_fo.py file.pdf
    for f in sys.argv[1:]:
        r = parse_jurnal_pdf(open(f, "rb").read())
        print(f, "->", "bukan jurnal" if r is None else "%s %s..%s %d baris, total %s (Grand TOTAL %s)" % (
            r["hotel"], r["dari"], r["sampai"], len(r["rows"]), round(sum(x["amount"] for x in r["rows"]), 2), r["grand"]))
