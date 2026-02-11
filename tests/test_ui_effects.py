# -*- coding: utf-8 -*-
import unittest
import ui_effects


class TestUiEffects(unittest.TestCase):
    def test_exports(self):
        self.assertTrue(callable(ui_effects.apply_button_effects))
        self.assertTrue(callable(ui_effects.bind_tooltip))
        self.assertTrue(callable(ui_effects.apply_focus_ring))
        self.assertTrue(hasattr(ui_effects, "Tooltip"))


if __name__ == "__main__":
    unittest.main()
