import unittest
import ui_builder


class TestUiBuilder(unittest.TestCase):
    def test_exports(self):
        self.assertTrue(callable(ui_builder.build_source_panel))
        self.assertTrue(callable(ui_builder.build_settings_panel))


if __name__ == "__main__":
    unittest.main()
