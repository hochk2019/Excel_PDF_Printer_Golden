import os
import tempfile

from app_utils import make_pdf_name


# ---------- Excel handlers (fixed for ALL sheets) ----------

def excel_print_first_sheet_ranges(path, ranges, printer, log):
    import win32com.client
    ok = True
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        if printer:
            try:
                excel.ActivePrinter = printer
            except Exception as e:
                log(f"[WARN] Excel không đặt được máy in '{printer}': {e}")
        wb = excel.Workbooks.Open(path, ReadOnly=True)
        ws = wb.Sheets(1)
        for a, b in ranges:
            ws.PrintOut(From=a, To=b)
        log(f"[OK] Excel in (sheet 1): {path} {ranges}")
        wb.Close(SaveChanges=False)
    except Exception as e:
        ok = False
        log(f"[ERR] Excel lỗi: {path} - {e}")
    finally:
        if excel:
            excel.Quit()
    return ok


def excel_print_all_sheets_ranges(path, ranges, printer, log):
    import win32com.client
    ok = True
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        if printer:
            try:
                excel.ActivePrinter = printer
            except Exception as e:
                log(f"[WARN] Excel không đặt được máy in '{printer}': {e}")
        wb = excel.Workbooks.Open(path, ReadOnly=True)
        count = wb.Sheets.Count
        for i in range(1, count + 1):
            ws = wb.Sheets(i)
            for a, b in ranges:
                ws.PrintOut(From=a, To=b)
        log(f"[OK] Excel in (ALL sheets, ranges {ranges}): {path} ({count} sheets)")
        wb.Close(SaveChanges=False)
    except Exception as e:
        ok = False
        log(f"[ERR] Excel lỗi: {path} - {e}")
    finally:
        if excel:
            excel.Quit()
    return ok


def excel_export_first_sheet(path, pages, ranges, printer, out_dir, merge, log):
    import win32com.client
    from pypdf import PdfWriter, PdfReader
    ok = True
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        if printer:
            try:
                excel.ActivePrinter = printer
            except Exception as e:
                log(f"[ERR] không thể chọn máy in PDF '{printer}': {e}")
                return False
        os.makedirs(out_dir, exist_ok=True)
        if merge:
            temps = []
            wb = excel.Workbooks.Open(path, ReadOnly=True)
            ws = wb.Sheets(1)
            for p in pages:
                tmp = make_pdf_name(out_dir, path, f"p{p}_tmp")
                ws.PrintOut(From=p, To=p, PrintToFile=True, PrToFileName=tmp)
                temps.append(tmp)
            wb.Close(SaveChanges=False)
            writer = PdfWriter()
            for t in temps:
                r = PdfReader(t)
                writer.add_page(r.pages[0])
            final = make_pdf_name(out_dir, path, "pages")
            with open(final, "wb") as f:
                writer.write(f)
            for t in temps:
                try:
                    os.remove(t)
                except Exception:
                    pass
            log(f"[OK] Excel xuất PDF (sheet1, gộp): {final}")
        else:
            wb = excel.Workbooks.Open(path, ReadOnly=True)
            ws = wb.Sheets(1)
            for a, b in ranges:
                for p in range(a, b + 1):
                    outp = make_pdf_name(out_dir, path, f"p{p}")
                    ws.PrintOut(From=p, To=p, PrintToFile=True, PrToFileName=outp)
            wb.Close(SaveChanges=False)
            log(f"[OK] Excel xuất PDF (sheet1, tách): {path}")
    except Exception as e:
        ok = False
        log(f"[ERR] Excel xuất PDF lỗi: {path} - {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
    return ok


def excel_export_all_sheets_pages(path, pages, ranges, printer, out_dir, merge, log):
    from pypdf import PdfWriter, PdfReader
    import win32com.client
    ok = True
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        if printer:
            try:
                excel.ActivePrinter = printer
            except Exception as e:
                log(f"[ERR] không thể chọn máy in PDF '{printer}': {e}")
                return False
        os.makedirs(out_dir, exist_ok=True)
        wb = excel.Workbooks.Open(path, ReadOnly=True)
        count = wb.Sheets.Count
        if merge:
            temps = []
            for i in range(1, count + 1):
                ws = wb.Sheets(i)
                safe = "".join(ch for ch in ws.Name if ch not in '\\/:*?"<>|') or f"S{i}"
                for p in pages:
                    tmp = make_pdf_name(out_dir, path, f"{safe}_p{p}_tmp")
                    ws.PrintOut(From=p, To=p, PrintToFile=True, PrToFileName=tmp)
                    temps.append(tmp)
            wb.Close(SaveChanges=False)
            writer = PdfWriter()
            for t in temps:
                r = PdfReader(t)
                for pg in r.pages:
                    writer.add_page(pg)
            final = make_pdf_name(out_dir, path, "allsheets_pages")
            with open(final, "wb") as f:
                writer.write(f)
            for t in temps:
                try:
                    os.remove(t)
                except Exception:
                    pass
            log(f"[OK] Excel xuất PDF (ALL sheets, gộp, pages={pages}): {final}")
        else:
            for i in range(1, count + 1):
                ws = wb.Sheets(i)
                safe = "".join(ch for ch in ws.Name if ch not in '\\/:*?"<>|') or f"S{i}"
                for p in pages:
                    outp = make_pdf_name(out_dir, path, f"{safe}_p{p}")
                    ws.PrintOut(From=p, To=p, PrintToFile=True, PrToFileName=outp)
            wb.Close(SaveChanges=False)
            log(f"[OK] Excel xuất PDF (ALL sheets, tách, pages={pages}): {path}")
    except Exception as e:
        ok = False
        log(f"[ERR] Excel xuất PDF lỗi: {path} - {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
    return ok


# ---------- Word/PDF handlers ----------

def word_print_ranges(path, ranges, printer, log):
    import win32com.client
    wdPrintFromTo = 4
    word = None
    ok = True
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        if printer:
            try:
                word.ActivePrinter = printer
            except Exception as e:
                log(f"[WARN] Word không đặt được máy in '{printer}': {e}")
        doc = word.Documents.Open(path, ReadOnly=True)
        for a, b in ranges:
            doc.PrintOut(Range=wdPrintFromTo, From=str(a), To=str(b))
        log(f"[OK] Word in: {path} {ranges}")
        doc.Close(SaveChanges=False)
    except Exception as e:
        ok = False
        log(f"[ERR] Word lỗi: {path} - {e}")
    finally:
        if word:
            word.Quit()
    return ok


def pdf_print_ranges(path, ranges, printer, log):
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception:
        log("[ERR] Thiếu pypdf: pip install pypdf")
        return False
    import win32print, win32api
    ok = True
    default_before = None
    if printer:
        try:
            default_before = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer)
        except Exception as e:
            log(f"[WARN] không đặt được máy in mặc định '{printer}': {e}")
    try:
        reader = PdfReader(path)
        total = len(reader.pages)
        for a, b in ranges:
            if a < 1 or b > total:
                log(f"[WARN] PDF {path}: range {(a, b)} vượt {total}")
                ok = False
                continue
            writer = PdfWriter()
            for i in range(a, b + 1):
                writer.add_page(reader.pages[i - 1])
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                writer.write(tmp)
                temp_path = tmp.name
            win32api.ShellExecute(0, "print", temp_path, None, ".", 0)
        log(f"[OK] PDF in: {path} {ranges}")
    except Exception as e:
        ok = False
        log(f"[ERR] PDF lỗi: {path} - {e}")
    finally:
        if default_before:
            try:
                win32print.SetDefaultPrinter(default_before)
            except Exception:
                pass
    return ok


def word_export_pdf(path, pages, ranges, printer, out_dir, merge, log):
    import win32com.client
    from pypdf import PdfWriter, PdfReader
    wdPrintFromTo = 4
    ok = True
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        if printer:
            try:
                word.ActivePrinter = printer
            except Exception as e:
                log(f"[ERR] không thể chọn máy in PDF '{printer}': {e}")
                return False
        os.makedirs(out_dir, exist_ok=True)
        if merge:
            temps = []
            doc = word.Documents.Open(path, ReadOnly=True)
            for p in pages:
                tmp = make_pdf_name(out_dir, path, f"p{p}_tmp")
                doc.PrintOut(Range=wdPrintFromTo, From=str(p), To=str(p), PrintToFile=True, OutputFileName=tmp)
                temps.append(tmp)
            doc.Close(SaveChanges=False)
            writer = PdfWriter()
            for t in temps:
                r = PdfReader(t)
                writer.add_page(r.pages[0])
            final = make_pdf_name(out_dir, path, "pages")
            with open(final, "wb") as f:
                writer.write(f)
            for t in temps:
                try:
                    os.remove(t)
                except Exception:
                    pass
            log(f"[OK] Word xuất PDF (gộp): {final}")
        else:
            doc = word.Documents.Open(path, ReadOnly=True)
            for a, b in ranges:
                for p in range(a, b + 1):
                    outp = make_pdf_name(out_dir, path, f"p{p}")
                    doc.PrintOut(Range=wdPrintFromTo, From=str(p), To=str(p), PrintToFile=True, OutputFileName=outp)
            doc.Close(SaveChanges=False)
            log(f"[OK] Word xuất PDF (tách): {path}")
    except Exception as e:
        ok = False
        log(f"[ERR] Word xuất PDF lỗi: {path} - {e}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return ok


def pdf_export(path, pages, ranges, out_dir, merge, log):
    from pypdf import PdfReader, PdfWriter
    ok = True
    try:
        r = PdfReader(path)
        os.makedirs(out_dir, exist_ok=True)
        if merge:
            w = PdfWriter()
            for p in pages:
                if p < 1 or p > len(r.pages):
                    ok = False
                    log(f"[WARN] PDF {path}: trang {p} vượt {len(r.pages)}")
                    continue
                w.add_page(r.pages[p - 1])
            final = make_pdf_name(out_dir, path, "pages")
            with open(final, "wb") as f:
                w.write(f)
            log(f"[OK] PDF trích (gộp): {final}")
        else:
            for p in pages:
                if p < 1 or p > len(r.pages):
                    ok = False
                    log(f"[WARN] PDF {path}: trang {p} vượt {len(r.pages)}")
                    continue
                w = PdfWriter()
                w.add_page(r.pages[p - 1])
                outp = make_pdf_name(out_dir, path, f"p{p}")
                with open(outp, "wb") as f:
                    w.write(f)
            log(f"[OK] PDF trích (tách): {path}")
    except Exception as e:
        ok = False
        log(f"[ERR] PDF xuất lỗi: {path} - {e}")
    return ok





