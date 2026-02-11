import os
import fnmatch
import datetime as dt

EXCEL_EXTS = ('.xlsx', '.xls', '.xlsm')
WORD_EXTS = ('.doc', '.docx')
PDF_EXTS = ('.pdf',)


def parse_date(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return dt.datetime.strptime(value, "%Y-%m-%d").date()
    except Exception:
        return None


def parse_pagespec(spec: str):
    pages = set()
    spec = (spec or "1").replace(" ", "")
    if not spec:
        return [1], [(1, 1)]
    for part in spec.split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            if a.isdigit() and b.isdigit():
                a = int(a)
                b = int(b)
                if a <= b:
                    for x in range(a, b + 1):
                        pages.add(x)
        else:
            if part.isdigit():
                pages.add(int(part))
    if not pages:
        pages = {1}
    pages = sorted(pages)
    ranges = []
    start = prev = pages[0]
    for x in pages[1:]:
        if x == prev + 1:
            prev = x
        else:
            ranges.append((start, prev))
            start = prev = x
    ranges.append((start, prev))
    return pages, ranges


def within_date(path, date_from, date_to):
    if not (date_from or date_to):
        return True
    try:
        m = dt.date.fromtimestamp(os.path.getmtime(path))
    except Exception:
        return False
    if date_from and m < date_from:
        return False
    if date_to and m > date_to:
        return False
    return True


def gather_files(folders, recursive, patterns, use_excel, use_word, use_pdf, dfrom, dto):
    exts = []
    if use_excel:
        exts += list(EXCEL_EXTS)
    if use_word:
        exts += list(WORD_EXTS)
    if use_pdf:
        exts += list(PDF_EXTS)
    if not exts:
        return []

    pats = ["*"] if not patterns else [p.strip() for p in patterns.split(";") if p.strip()]

    def match(name):
        low = name.lower()
        if low.startswith("~$") or low.endswith(".tmp"):
            return False
        if not any(low.endswith(e) for e in exts):
            return False
        return any(fnmatch.fnmatch(name, pat) for pat in pats)

    files = []
    for base in folders:
        if not os.path.isdir(base):
            continue
        it = os.walk(base) if recursive else [(base, [], os.listdir(base))]
        for root, _, names in it:
            for name in names:
                if match(name):
                    path = os.path.join(root, name)
                    if within_date(path, dfrom, dto):
                        files.append(path)
    return files


def make_pdf_name(out_dir, in_path, suffix):
    base = os.path.splitext(os.path.basename(in_path))[0]
    safe = "".join(ch for ch in base if ch not in '\\/:*?"<>|')
    return os.path.join(out_dir, f"{safe}_{suffix}.pdf")
