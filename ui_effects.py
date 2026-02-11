import tkinter as tk


class Tooltip:
    def __init__(self, widget, text, colors, delay=400, wrap=280):
        self.widget = widget
        self.text = text
        self.colors = colors
        self.delay = delay
        self.wrap = wrap
        self._after_id = None
        self._tip = None

        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress-1>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self._tip or not self.text:
            return
        x = self.widget.winfo_pointerx() + 12
        y = self.widget.winfo_pointery() + 16
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        try:
            self._tip.attributes("-topmost", True)
        except Exception:
            pass
        self._tip.geometry(f"+{x}+{y}")
        label = tk.Label(
            self._tip,
            text=self.text,
            bg=self.colors.get("text", "#0F172A"),
            fg="white",
            padx=8,
            pady=4,
            justify="left",
            wraplength=self.wrap,
            font=("Segoe UI", 9),
        )
        label.pack()

    def _hide(self, _event=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


def bind_tooltip(widget, text, colors, delay=400):
    widget._tooltip = Tooltip(widget, text, colors, delay=delay)
    return widget._tooltip


def apply_focus_ring(widget, colors, base=2, focus=3, base_color=None, focus_color=None):
    try:
        current_base = widget.cget("highlightthickness")
        base_thickness = int(current_base) if str(current_base).isdigit() else base
    except Exception:
        base_thickness = base

    if base_color is None:
        try:
            base_color = widget.cget("highlightbackground")
        except Exception:
            base_color = colors.get("border", "#E2E8F0")
    if focus_color is None:
        try:
            focus_color = widget.cget("highlightcolor")
        except Exception:
            focus_color = colors.get("primary", "#3B82F6")

    try:
        widget.configure(
            highlightthickness=base_thickness,
            highlightbackground=base_color,
            highlightcolor=focus_color,
        )
    except Exception:
        return

    def on_focus_in(_event=None):
        try:
            widget.configure(highlightthickness=focus, highlightcolor=focus_color)
        except Exception:
            pass

    def on_focus_out(_event=None):
        try:
            widget.configure(highlightthickness=base_thickness, highlightbackground=base_color)
        except Exception:
            pass

    widget.bind("<FocusIn>", on_focus_in, add="+")
    widget.bind("<FocusOut>", on_focus_out, add="+")


def apply_button_effects(button, colors, style="secondary"):
    palette = {
        "primary": {
            "bg": colors.get("primary", "#3B82F6"),
            "fg": "white",
            "hover": colors.get("primary_dark", "#1D4ED8"),
            "pressed": "#1E40AF",
        },
        "secondary": {
            "bg": colors.get("border", "#E2E8F0"),
            "fg": colors.get("text", "#1E293B"),
            "hover": "#D5DEE9",
            "pressed": "#C7D2E0",
        },
        "danger": {
            "bg": colors.get("danger", "#EF4444"),
            "fg": "white",
            "hover": "#DC2626",
            "pressed": "#B91C1C",
        },
    }
    use = palette.get(style, palette["secondary"])

    normal_bg = use["bg"]
    normal_fg = use["fg"]
    hover_bg = use["hover"]
    pressed_bg = use["pressed"]
    hover_fg = use.get("hover_fg", normal_fg)
    pressed_fg = use.get("pressed_fg", normal_fg)

    try:
        button.configure(
            cursor="hand2",
            activebackground=pressed_bg,
            activeforeground=pressed_fg,
        )
    except Exception:
        pass

    original_relief = button.cget("relief")
    original_bd = button.cget("bd")

    def is_disabled():
        try:
            return str(button.cget("state")) == "disabled"
        except Exception:
            return False

    def on_enter(_event=None):
        if is_disabled():
            return
        button.configure(bg=hover_bg, fg=hover_fg)

    def on_leave(_event=None):
        if is_disabled():
            return
        button.configure(bg=normal_bg, fg=normal_fg, relief=original_relief, bd=original_bd)

    def on_press(_event=None):
        if is_disabled():
            return
        button.configure(bg=pressed_bg, fg=pressed_fg, relief="sunken")

    def on_release(_event=None):
        if is_disabled():
            return
        inside = button == button.winfo_containing(
            button.winfo_pointerx(), button.winfo_pointery()
        )
        button.configure(
            bg=hover_bg if inside else normal_bg,
            fg=hover_fg if inside else normal_fg,
            relief=original_relief,
            bd=original_bd,
        )

    button.bind("<Enter>", on_enter, add="+")
    button.bind("<Leave>", on_leave, add="+")
    button.bind("<ButtonPress-1>", on_press, add="+")
    button.bind("<ButtonRelease-1>", on_release, add="+")
    apply_focus_ring(button, colors)
