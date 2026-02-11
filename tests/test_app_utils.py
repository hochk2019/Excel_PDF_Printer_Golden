import os
import tempfile
import unittest
import datetime as dt

from app_utils import parse_date, parse_pagespec, gather_files, make_pdf_name


class TestAppUtils(unittest.TestCase):
    def test_parse_pagespec_ranges(self):
        pages, ranges = parse_pagespec("1,3,5-7")
        self.assertEqual(pages, [1, 3, 5, 6, 7])
        self.assertEqual(ranges, [(1, 1), (3, 3), (5, 7)])

    def test_parse_pagespec_default(self):
        pages, ranges = parse_pagespec("")
        self.assertEqual(pages, [1])
        self.assertEqual(ranges, [(1, 1)])

    def test_parse_date(self):
        self.assertEqual(parse_date("2026-02-10"), dt.date(2026, 2, 10))
        self.assertIsNone(parse_date("2026/02/10"))
        self.assertIsNone(parse_date(""))

    def test_gather_files_filters(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = tmp
            sub = os.path.join(tmp, "sub")
            os.makedirs(sub, exist_ok=True)
            paths = [
                os.path.join(root, "a.xlsx"),
                os.path.join(root, "b.docx"),
                os.path.join(root, "c.pdf"),
                os.path.join(root, "ignore.txt"),
                os.path.join(root, "~$temp.xlsx"),
                os.path.join(sub, "d.xlsm"),
            ]
            for p in paths:
                with open(p, "w", encoding="utf-8") as f:
                    f.write("x")

            files = gather_files(
                [root],
                recursive=True,
                patterns="*",
                use_excel=True,
                use_word=False,
                use_pdf=False,
                dfrom=None,
                dto=None,
            )
            self.assertIn(os.path.join(root, "a.xlsx"), files)
            self.assertIn(os.path.join(sub, "d.xlsm"), files)
            self.assertNotIn(os.path.join(root, "b.docx"), files)
            self.assertNotIn(os.path.join(root, "c.pdf"), files)
            self.assertNotIn(os.path.join(root, "~$temp.xlsx"), files)

    def test_make_pdf_name(self):
        out = make_pdf_name("C:/out", "C:/in/fi:le?.xlsx", "pages")
        self.assertTrue(out.endswith("file_pages.pdf"))


if __name__ == "__main__":
    unittest.main()
