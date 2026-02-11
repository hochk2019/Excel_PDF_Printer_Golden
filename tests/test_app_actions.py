# -*- coding: utf-8 -*-
import unittest
from unittest.mock import patch

import app_actions as aa


class SimpleVar:
    def __init__(self, value=None):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class ListboxStub:
    def __init__(self):
        self.items = []

    def delete(self, _start, _end=None):
        self.items = []

    def insert(self, _idx, value):
        self.items.append(value)


class LabelStub:
    def __init__(self):
        self.text = None

    def config(self, text=None):
        self.text = text


class DummyFilter:
    def __init__(self):
        self.current_files = ["C:/Xx/b.txt", "C:/Xx/a.txt", "C:/Xx/c.txt"]
        self.filtered_files = []
        self.search_var = SimpleVar("a.txt")
        self.sort_var = SimpleVar("Tên A→Z")
        self.file_listbox = ListboxStub()
        self.lbl_count = LabelStub()
        self.summary_called = False

    def update_summary(self):
        self.summary_called = True


class DummyState:
    def __init__(self):
        self.folders = []
        self.recursive = SimpleVar(False)
        self.pattern = SimpleVar("")
        self.date_from = SimpleVar("")
        self.date_to = SimpleVar("")
        self.use_excel = SimpleVar(False)
        self.use_word = SimpleVar(False)
        self.use_pdf = SimpleVar(False)
        self.page_spec = SimpleVar("")
        self.mode_pdf_export = SimpleVar(False)
        self.pdf_merge = SimpleVar(False)
        self.sheet_scope = SimpleVar("first")
        self.paper_size = SimpleVar("A4")
        self.out_dir = SimpleVar("")
        self.enable_logging = SimpleVar(False)
        self.log_dir = SimpleVar("")
        self.selected_printer = SimpleVar("Printer A")
        self.search_var = SimpleVar("")
        self.sort_options = ["Tên A→Z", "Tên Z→A"]
        self.sort_var = SimpleVar(self.sort_options[0])
        self.listbox = ListboxStub()
        self.scan_called = False
        self.toggle_pdf_called = False
        self.toggle_log_called = False

    def scan_files(self):
        self.scan_called = True

    def toggle_pdf_dir(self):
        self.toggle_pdf_called = True

    def toggle_logging_fields(self):
        self.toggle_log_called = True


class DummyPreset:
    def __init__(self):
        self.presets = {}
        self.preset_name = SimpleVar("")
        self.preset_input = SimpleVar("Cấu hình A")
        self.logged = []
        self.applied = None
        self.refreshed = False

    def collect_state(self):
        return {"k": 1}

    def apply_state(self, state):
        self.applied = state

    def refresh_preset_list(self):
        self.refreshed = True

    def log(self, msg):
        self.logged.append(msg)


class TestAppActions(unittest.TestCase):
    def test_apply_file_filter(self):
        dummy = DummyFilter()
        aa.apply_file_filter(dummy)
        self.assertEqual(dummy.filtered_files, ["C:/Xx/a.txt"])
        self.assertEqual(dummy.file_listbox.items, ["C:/Xx/a.txt"])
        self.assertEqual(dummy.lbl_count.text, "Tìm thấy: 1 file")
        self.assertTrue(dummy.summary_called)

    def test_collect_apply_state(self):
        dummy = DummyState()
        dummy.folders = ["C:/in"]
        dummy.recursive.set(True)
        dummy.pattern.set("*")
        dummy.use_excel.set(True)
        dummy.use_pdf.set(True)
        dummy.page_spec.set("1,3")
        dummy.mode_pdf_export.set(True)
        dummy.out_dir.set("C:/out")
        dummy.enable_logging.set(True)
        dummy.log_dir.set("C:/log")
        dummy.search_var.set("abc")
        dummy.sort_var.set("Tên Z→A")

        state = aa.collect_state(dummy)
        fresh = DummyState()
        aa.apply_state(fresh, state)

        self.assertEqual(fresh.folders, ["C:/in"])
        self.assertTrue(fresh.recursive.get())
        self.assertEqual(fresh.page_spec.get(), "1,3")
        self.assertTrue(fresh.mode_pdf_export.get())
        self.assertEqual(fresh.out_dir.get(), "C:/out")
        self.assertTrue(fresh.enable_logging.get())
        self.assertEqual(fresh.search_var.get(), "abc")
        self.assertEqual(fresh.sort_var.get(), "Tên Z→A")
        self.assertEqual(fresh.listbox.items, ["C:/in"])
        self.assertTrue(fresh.scan_called)
        self.assertTrue(fresh.toggle_pdf_called)
        self.assertTrue(fresh.toggle_log_called)

    @patch("app_actions.messagebox.askyesno", return_value=True)
    @patch("app_actions.messagebox.showinfo")
    def test_save_delete_apply_preset(self, _showinfo, _askyesno):
        dummy = DummyPreset()
        aa.save_preset(dummy)
        self.assertIn("Cấu hình A", dummy.presets)
        self.assertTrue(dummy.refreshed)

        dummy.preset_name.set("Cấu hình A")
        aa.apply_selected_preset(dummy)
        self.assertEqual(dummy.applied, {"k": 1})

        aa.delete_preset(dummy)
        self.assertNotIn("Cấu hình A", dummy.presets)

    @patch("app_actions.messagebox.showinfo")
    @patch("app_actions.os.startfile")
    @patch("app_actions.os.path.isdir", return_value=True)
    def test_open_output_folder(self, _isdir, startfile, _showinfo):
        class Dummy:
            def __init__(self):
                self.mode_pdf_export = SimpleVar(True)
                self.out_dir = SimpleVar("C:/out")

        dummy = Dummy()
        aa.open_output_folder(dummy)
        startfile.assert_called_once_with("C:/out")

    @patch("app_actions.messagebox.askyesno", return_value=True)
    @patch("app_actions.messagebox.showinfo")
    def test_clear_folders(self, _showinfo, _askyesno):
        dummy = DummyState()
        dummy.folders = ["C:/in", "C:/in2"]
        aa.clear_folders(dummy)
        self.assertEqual(dummy.folders, [])
        self.assertEqual(dummy.listbox.items, [])
        self.assertTrue(dummy.scan_called)


if __name__ == "__main__":
    unittest.main()
