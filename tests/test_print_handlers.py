import unittest
import print_handlers as ph


class TestPrintHandlers(unittest.TestCase):
    def test_handlers_available(self):
        names = [
            "excel_print_first_sheet_ranges",
            "excel_print_all_sheets_ranges",
            "excel_export_first_sheet",
            "excel_export_all_sheets_pages",
            "word_print_ranges",
            "word_export_pdf",
            "pdf_print_ranges",
            "pdf_export",
        ]
        for name in names:
            self.assertTrue(hasattr(ph, name), name)
            self.assertTrue(callable(getattr(ph, name)), name)


if __name__ == "__main__":
    unittest.main()
