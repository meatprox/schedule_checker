"""スケジュール把握のPython版。ブラウザも通信も使わない。"""

from __future__ import annotations

import math
import sys
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk

from pdf_export import pdf_file_name, pdf_overview_file_name, write_overview_pdf, write_schedule_pdf
from model import (
    SLOT_COUNT,
    STYLES,
    Point,
    Segment,
    Task,
    assign_slots,
    clone_tasks,
    slot_count_for,
    effort_spans,
    export_file_name,
    parse_workbook,
    same_tasks,
    LINE_COLORS,
    sample_workbook,
    save_tasks,
    task_span,
)

AXIS_H = 56
EFFORT_H = 32
PAD_L = 28
PAD_R = 40
GUTTER_W = 72
MIN_PX_PER_HOUR = 32
MAX_PX_PER_HOUR = 3600
COLORS = LINE_COLORS
WEEKDAYS = "日月火水木金土"
STYLE_LABELS = (
    ("solid", "実線"),
    ("dashed", "破線"),
    ("double", "二重線"),
    ("triple", "三重線"),
    ("quad", "四重線"),
    ("quint", "五重線"),
    ("none", "線無し"),
)


def enable_dpi() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        return


def clamp(low, value, high):
    return min(high, max(low, value))


def start_of_day(value: datetime) -> datetime:
    return datetime(value.year, value.month, value.day)


def add_days(value: datetime, days: int) -> datetime:
    shifted = start_of_day(value) + timedelta(days=days)
    return datetime(shifted.year, shifted.month, shifted.day)


def format_hm(value: datetime) -> str:
    return f"{value.hour:02d}:{value.minute:02d}"


def format_day(value: datetime) -> str:
    weekday = WEEKDAYS[(value.weekday() + 1) % 7]
    return f"{value.month}/{value.day}（{weekday}）"


def snap_minute(value: datetime) -> datetime:
    base = datetime(value.year, value.month, value.day)
    minutes = int(round((value - base).total_seconds() / 60))
    return base + timedelta(minutes=minutes)


def mix(hex_color: str, hex_background: str, alpha: float) -> str:
    def channels(text: str):
        return tuple(int(text[index : index + 2], 16) for index in (1, 3, 5))

    fg = channels(hex_color)
    bg = channels(hex_background)
    mixed = tuple(round(back + (front - back) * alpha) for front, back in zip(fg, bg))
    return "#{:02x}{:02x}{:02x}".format(*mixed)


def effort_color(total: float, alert_at: float | None = None) -> tuple[str, str]:
    amount = 0.14 + total * 0.18
    alpha = min(0.9, amount)
    if alert_at is not None and round(total * 10) >= round(alert_at * 10):
        alert_alpha = min(0.92, 0.38 + total * 0.1)
        fill = mix("#c81e1e", "#f8ecec", alert_alpha)
        text = "#fff5f5" if alert_alpha > 0.55 else "#9f1239"
        return fill, text
    fill = mix("#0e6b67", "#e7f3f2", alpha)
    text = "#f7fffe" if amount > 0.55 else "#0e6b67"
    return fill, text


def format_effort(value: float) -> str:
    rounded = round(value * 10) / 10
    if float(rounded).is_integer():
        return str(int(rounded))
    return f"{rounded:.1f}"


class ScheduleApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("スケジュール把握（Python版）")
        self.root.geometry("1320x858")
        self.root.minsize(960, 640)
        self.root.configure(bg="#e7eef0")

        self.tasks = []
        self.snapshot = []
        self.file_name = ""
        self.sheet_name = ""
        self.skipped = 0
        self.axis_start = None
        self.axis_end = None
        self.px_per_hour = 80.0
        self.slot_h = 64
        self.slot_count = SLOT_COUNT
        self.slot_names = [""] * SLOT_COUNT
        self.saved_slot_names = [""] * SLOT_COUNT
        self.gutter_w = GUTTER_W
        self.slot_entry = None
        self.slot_edit_index = None
        self.tool = None
        self.pen_anchor = None
        self.pen_draft = None
        self.name_box = None
        self.name_entry = None
        self.alert_var = tk.StringVar()
        self.drag = None
        self.selected = None
        self.color_task_id = None
        self.point_hits = []
        self.label_hits = []
        self.seg_hits = []
        self.day_labels = []
        self.timeline_width = 1
        self.view_height = 1
        self._suspend_scroll = False

        self._fonts()
        self._build()
        self.root.bind("<Escape>", self.on_escape)
        self.root.bind_all("<Left>", self.on_arrow)
        self.root.bind_all("<Right>", self.on_arrow)
        self.root.bind_all("<ButtonRelease-1>", self.on_release, add="+")

    def _fonts(self) -> None:
        family = "Yu Gothic UI"
        self.font_ui = (family, 10)
        self.font_title = (family, 16, "bold")
        self.font_small = (family, 9)
        self.font_axis = (family, 9)
        self.font_day = (family, 10, "bold")
        self.font_label = (family, 9, "bold")
        self.font_menu = (family, 8)
        self.measure = tkfont.Font(root=self.root, family=family, size=9, weight="bold")
        self.slot_measure = tkfont.Font(root=self.root, family=family, size=9)

    def _build(self) -> None:
        top = tk.Frame(self.root, bg="#e7eef0")
        top.pack(fill="x", padx=16, pady=(12, 6))
        titles = tk.Frame(top, bg="#e7eef0")
        titles.pack(side="left", fill="x", expand=True)
        tk.Label(titles, text="スケジュール把握", font=self.font_title, bg="#e7eef0", fg="#1b2428").pack(anchor="w")
        self.status = tk.Label(
            titles,
            text="Excelの1〜4列目を読み、1日幅の時間軸にタスクを並べます。",
            font=self.font_ui,
            bg="#e7eef0",
            fg="#5d6c74",
            anchor="w",
        )
        self.status.pack(anchor="w")
        actions = tk.Frame(top, bg="#e7eef0")
        actions.pack(side="right")
        tk.Label(actions, text="この工数以上を赤く", font=self.font_ui, bg="#e7eef0", fg="#5d6c74").pack(side="left", padx=(0, 4))
        self.alert_entry = ttk.Entry(actions, textvariable=self.alert_var, width=6)
        self.alert_entry.pack(side="left", padx=(0, 10))
        self.alert_var.trace_add("write", self.on_alert_changed)
        self.reset_button = ttk.Button(actions, text="変更を元に戻す", command=self.restore)
        self.save_button = ttk.Menubutton(actions, text="保存")
        self.save_menu = tk.Menu(self.save_button, tearoff=0)
        self.save_menu.add_command(label="Excelに保存", command=self.save_dialog)
        self.save_menu.add_command(label="PDFに保存", command=self.save_pdf_dialog)
        self.save_menu.add_command(label="PDF（全体）", command=self.save_overview_dialog)
        self.save_button["menu"] = self.save_menu
        self.open_button = ttk.Button(actions, text="Excelを開く", command=self.open_dialog)
        for button in (self.reset_button, self.save_button, self.open_button):
            button.pack(side="left", padx=3)
        self.reset_button.state(["disabled"])
        self.set_save_enabled(False)

        self.error = tk.Label(self.root, text="", font=self.font_ui, bg="#fff1f2", fg="#9f1239", anchor="w")
        self.hint = tk.Label(
            self.root,
            text="点を左右へドラッグして時刻を変更    線かタスク名をドラッグして時刻とスロットを移動    線をクリックして線種を変更    タスク名をクリックして色を変更    点にマウスを乗せて説明を表示    Ctrl＋ホイールで時間軸を拡大・縮小",
            font=self.font_small,
            bg="#e7eef0",
            fg="#5d6c74",
            anchor="w",
        )

        self.empty = tk.Frame(self.root, bg="#ffffff", highlightbackground="#d5dee3", highlightthickness=1)
        self.empty.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        body = tk.Frame(self.empty, bg="#ffffff")
        body.pack(anchor="w", padx=22, pady=18)
        tk.Label(body, text="Excelの並び", font=(self.font_title[0], 13, "bold"), bg="#ffffff").pack(anchor="w")
        tk.Label(
            body,
            text="1列目がタスク名、2列目が日付、3列目が時間、4列目がその時間の説明です。5列目に番号を書くと、そのスロットに配置します。10より大きい番号があると、その数まで段を増やします。空欄なら自動で配置します。同じタスク名が続いている行を結びます。空行や別のタスク名を挟むと、別の線になります。",
            font=self.font_ui,
            bg="#ffffff",
            wraplength=760,
            justify="left",
        ).pack(anchor="w", pady=(4, 8))
        tk.Label(
            body,
            text="この画面は Python だけで動きます。ブラウザも通信も使いません。",
            font=self.font_ui,
            bg="#ffffff",
            fg="#0e6b67",
        ).pack(anchor="w", pady=(0, 8))
        for line in (
            "点を左右へドラッグすると、その点の時刻だけが変わります。スロットは移動しません。",
            "スロットの番号をダブルクリックすると名前を変えられます。空にすると番号に戻ります。",
            "線かタスク名をドラッグすると、タスク全体の時刻がずれ、上下でスロットへ移ります。",
            "点と点のあいだの線をクリックすると、実線・破線・二重線・三重線・四重線・五重線・線無しを選べます。",
            "タスク名をクリックすると、そのライン全体の色を10色から選べます。ドラッグすると、これまでどおり移動します。",
            "工数は線無しが 0、破線が 0.5、実線が 1、二重線が 2、三重線が 3、四重線が 4、五重線が 5 です。時間軸の下に合計を表示します。",
            "PDFに保存すると、1日1枚の横長ページに予定表を書き出します。PDF（全体）は全期間を1枚にまとめます。縮小すると全体が見え、拡大すると文字が読めます。",
            "「この工数以上を赤く」に数を入れると、その数以上の合計が赤くなります。空欄なら色は緑のままです。",
            "Ctrlを押しながらホイールを回すと、時間軸が拡大・縮小します。上で拡大、下で縮小です。",
            "下のペンで、同じスロット上を2か所クリックすると線を引けます。消しゴムを選んでいるあいだは、線をダブルクリックするとその1本だけ消えます。",
        ):
            tk.Label(body, text="・ " + line, font=self.font_ui, bg="#ffffff", anchor="w").pack(anchor="w")

        self.tool_bar = tk.Frame(self.root, bg="#e7eef0")
        tk.Label(self.tool_bar, text="ツール", font=self.font_small, bg="#e7eef0", fg="#5d6c74").pack(side="left", padx=(0, 8))
        self.pen_button = tk.Button(
            self.tool_bar,
            text="ペン",
            font=self.font_ui,
            relief="raised",
            bg="#ffffff",
            fg="#1b2428",
            padx=14,
            pady=2,
            command=lambda: self.toggle_tool("pen"),
        )
        self.eraser_button = tk.Button(
            self.tool_bar,
            text="消しゴム",
            font=self.font_ui,
            relief="raised",
            bg="#ffffff",
            fg="#1b2428",
            padx=14,
            pady=2,
            command=lambda: self.toggle_tool("eraser"),
        )
        self.pen_button.pack(side="left", padx=3)
        self.eraser_button.pack(side="left", padx=3)
        self.pen_button.configure(state="disabled")
        self.eraser_button.configure(state="disabled")

        self.board = tk.Frame(self.root, bg="#ffffff", highlightbackground="#d5dee3", highlightthickness=1)
        self.gutter = tk.Canvas(self.board, width=GUTTER_W, bg="#f7fbfb", highlightthickness=0, bd=0)
        self.canvas = tk.Canvas(self.board, bg="#ffffff", highlightthickness=0, bd=0, xscrollincrement=1)
        self.hbar = ttk.Scrollbar(self.board, orient="horizontal", command=self.canvas.xview)
        self.vbar = ttk.Scrollbar(self.board, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self._on_xview, yscrollcommand=self._on_yview)
        self.gutter.grid(row=0, column=0, sticky="ns")
        self.canvas.grid(row=0, column=1, sticky="nsew")
        self.hbar.grid(row=1, column=1, sticky="ew")
        self.gutter.bind("<MouseWheel>", self.on_vertical_wheel)
        self.gutter.bind("<Double-Button-1>", self.on_gutter_double_click)
        self.board.rowconfigure(0, weight=1)
        self.board.columnconfigure(1, weight=1)
        self.board.bind("<Configure>", self.on_board_configure)
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<Double-Button-1>", self.on_canvas_double)
        self.root.bind_all("<B1-Motion>", self.on_drag_motion)
        self.canvas.bind("<Motion>", self.on_hover)
        self.canvas.bind("<Leave>", lambda _event: self.hide_tip())
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Control-MouseWheel>", self.on_zoom)

        self.tip = tk.Toplevel(self.root)
        self.tip.overrideredirect(True)
        self.tip.withdraw()
        self.tip.configure(bg="#1b2428")
        self.tip_time = tk.Label(self.tip, font=self.font_label, bg="#1b2428", fg="#ffffff")
        self.tip_task = tk.Label(self.tip, font=self.font_ui, bg="#1b2428", fg="#d5dee3")
        self.tip_desc = tk.Label(self.tip, font=self.font_ui, bg="#1b2428", fg="#ffffff", wraplength=280, justify="left")
        for label in (self.tip_time, self.tip_task, self.tip_desc):
            label.pack(anchor="w", padx=8, pady=1)

        self.menu = tk.Toplevel(self.root)
        self.menu.overrideredirect(True)
        self.menu.withdraw()
        self.menu.configure(bg="#ffffff", highlightbackground="#d5dee3", highlightthickness=1)
        self.menu_buttons = {}
        for style, label in STYLE_LABELS:
            button = tk.Button(
                self.menu,
                text=label,
                font=self.font_menu,
                relief="flat",
                bg="#ffffff",
                anchor="w",
                padx=8,
                pady=0,
                command=lambda chosen=style: self.choose_style(chosen),
            )
            button.pack(fill="x")
            self.menu_buttons[style] = button

        self.color_menu = tk.Toplevel(self.root)
        self.color_menu.overrideredirect(True)
        self.color_menu.withdraw()
        self.color_menu.configure(bg="#ffffff", highlightbackground="#d5dee3", highlightthickness=1)
        self.color_swatches = {}
        for index, color in enumerate(COLORS):
            swatch = tk.Frame(
                self.color_menu,
                width=22,
                height=22,
                bg=color,
                cursor="hand2",
                highlightthickness=2,
                highlightbackground="#ffffff",
            )
            swatch.grid(row=index // 5, column=index % 5, padx=4, pady=4)
            swatch.grid_propagate(False)
            swatch.bind("<Button-1>", lambda _event, chosen=color: self.choose_color(chosen))
            self.color_swatches[color] = swatch

    def show_error(self, message: str) -> None:
        self.error.configure(text=message)
        if not self.error.winfo_ismapped():
            self.error.pack(fill="x", padx=16, pady=(0, 6), before=self.empty if self.empty.winfo_ismapped() else self.board)

    def hide_error(self) -> None:
        self.error.pack_forget()

    def show_board(self) -> None:
        self.empty.pack_forget()
        if not self.hint.winfo_ismapped():
            self.hint.pack(fill="x", padx=16, pady=(0, 6))
        if not self.tool_bar.winfo_ismapped():
            self.tool_bar.pack(side="bottom", fill="x", padx=16, pady=(0, 12))
        if not self.board.winfo_ismapped():
            self.board.pack(fill="both", expand=True, padx=16, pady=(0, 8))

    def load_sample(self) -> None:
        self.apply_workbook(sample_workbook(), "sample.xlsx")

    def open_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="Excelを開く",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.load_path(path)

    def load_path(self, path) -> None:
        try:
            parsed_name = Path(path).name
            workbook = parse_workbook(path)
        except Exception as exc:
            self.show_error(f"このExcelは読み込めませんでした。{exc}")
            return
        self.apply_result(workbook, parsed_name)

    def apply_workbook(self, workbook, file_name: str) -> None:
        try:
            result = parse_workbook(workbook)
        except Exception as exc:
            self.show_error(f"このExcelは読み込めませんでした。{exc}")
            return
        self.apply_result(result, file_name)

    def apply_result(self, result, file_name: str) -> None:
        if not result.tasks:
            self.show_error(
                "時刻を読み取れませんでした。1列目にタスク名、2列目に日付、3列目に時間、4列目に説明があるか確認してください。"
            )
            return
        loaded_names = [str(name or "").strip() for name in result.slot_names]
        self.slot_count = max(slot_count_for(result.tasks), len(loaded_names))
        assign_slots(result.tasks, self.slot_count)
        self.slot_names = [loaded_names[index] if index < len(loaded_names) else "" for index in range(self.slot_count)]
        self.saved_slot_names = list(self.slot_names)
        self.cancel_slot_edit()
        self.clear_pen_draft()
        self.tool = None
        self.refresh_tool_buttons()
        self.file_name = file_name
        self.sheet_name = result.sheet_name
        self.tasks = result.tasks
        self.snapshot = clone_tasks(result.tasks)
        self.skipped = result.skipped
        self.selected = None
        self.drag = None
        self.close_menu()
        self.hide_tip()
        self.hide_error()
        self.fit_axis()
        self.show_board()
        self.root.update_idletasks()
        first = min((point.time for task in self.tasks for point in task.points), default=None)
        restore = 0 if first is None else max(0, self.time_to_x(first) - 96)
        self.redraw(restore, 0)

    def save_dialog(self) -> None:
        if not self.tasks:
            return
        initial = export_file_name(self.file_name)
        path = filedialog.asksaveasfilename(
            title="Excelに保存",
            defaultextension=".xlsx",
            initialfile=initial,
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.save_to(path)

    def save_to(self, path) -> None:
        try:
            save_tasks(self.tasks, self.sheet_name or "予定", path, self.slot_names)
        except OSError as exc:
            self.show_error(f"保存できませんでした。{exc}")
            return
        self.hide_error()
        self.status.configure(text=f"保存しました: {path}")

    def save_pdf_dialog(self) -> None:
        if not self.tasks or self.axis_start is None:
            return
        initial = pdf_file_name(self.file_name)
        path = filedialog.asksaveasfilename(
            title="PDFに保存",
            defaultextension=".pdf",
            initialfile=initial,
            filetypes=[("PDF", "*.pdf")],
        )
        if path:
            self.save_pdf_to(path)

    def save_pdf_to(self, path) -> None:
        try:
            write_schedule_pdf(
                path,
                self.tasks,
                self.axis_start,
                self.axis_end,
                self.alert_level(),
                self.file_name,
                self.sheet_name,
                self.slot_count,
                self.slot_names,
            )
        except OSError as exc:
            self.show_error(f"PDFを保存できませんでした。{exc}")
            return
        except Exception as exc:
            self.show_error(str(exc) or "PDFを保存できませんでした。")
            return
        self.hide_error()
        self.status.configure(text=f"PDFを保存しました: {path}")

    def save_overview_dialog(self) -> None:
        if not self.tasks or self.axis_start is None:
            return
        initial = pdf_overview_file_name(self.file_name)
        path = filedialog.asksaveasfilename(
            title="全体を1枚のPDFに保存",
            defaultextension=".pdf",
            initialfile=initial,
            filetypes=[("PDF", "*.pdf")],
        )
        if path:
            self.save_overview_to(path)

    def save_overview_to(self, path) -> None:
        try:
            write_overview_pdf(
                path,
                self.tasks,
                self.axis_start,
                self.axis_end,
                self.alert_level(),
                self.file_name,
                self.sheet_name,
                self.slot_count,
                self.slot_names,
            )
        except OSError as exc:
            self.show_error(f"PDFを保存できませんでした。{exc}")
            return
        except Exception as exc:
            self.show_error(str(exc) or "PDFを保存できませんでした。")
            return
        self.hide_error()
        self.status.configure(text=f"全体のPDFを保存しました: {path}")

    def restore(self) -> None:
        if not self.snapshot:
            return
        self.tasks = clone_tasks(self.snapshot)
        self.slot_names = list(self.saved_slot_names)
        self.slot_count = max(self.slot_count, len(self.slot_names))
        self.cancel_slot_edit()
        self.clear_pen_draft()
        self.selected = None
        self.cancel_drag()
        self.close_menu()
        self.fit_axis()
        self.redraw()

    def fit_axis(self) -> None:
        times = [point.time for task in self.tasks for point in task.points]
        self.axis_start = start_of_day(min(times))
        self.axis_end = add_days(start_of_day(max(times)), 1)

    def on_alert_changed(self, *_args) -> None:
        if self.tasks and self.axis_start and self.drag is None:
            self.redraw()

    def alert_level(self):
        text = self.alert_var.get().strip().replace("，", ".").replace("、", ".")
        if not text:
            return None
        try:
            value = float(text)
        except ValueError:
            return None
        if value < 0:
            return None
        return value

    def on_board_configure(self, event) -> None:
        available = event.height
        if available < 80:
            return
        slot_h = int(clamp(36, (available - AXIS_H - EFFORT_H - 16) / self.slot_count, 120))
        if slot_h != self.slot_h and self.tasks:
            self.slot_h = slot_h
            self.redraw()
        else:
            self.slot_h = slot_h

    def content_top(self) -> int:
        return AXIS_H + EFFORT_H

    def time_to_x(self, value: datetime) -> float:
        hours = (value - self.axis_start).total_seconds() / 3600
        return PAD_L + hours * self.px_per_hour

    def x_to_time(self, x: float) -> datetime:
        hours = (x - PAD_L) / self.px_per_hour
        return self.axis_start + timedelta(seconds=hours * 3600)

    def timeline_size(self) -> tuple[float, int]:
        hours = (self.axis_end - self.axis_start).total_seconds() / 3600
        width = PAD_L + hours * self.px_per_hour + PAD_R
        height = self.content_top() + self.slot_count * self.slot_h
        return width, int(height)

    def redraw(self, restore_x=None, restore_y=None) -> None:
        if not self.tasks or self.axis_start is None:
            return
        if restore_x is None:
            restore_x = self.canvas.canvasx(0)
        if restore_y is None:
            restore_y = self.canvas.canvasy(0)
        width, height = self.timeline_size()
        self.timeline_width = width
        self.view_height = height
        self._suspend_scroll = True
        self.canvas.delete("all")
        self.gutter.delete("all")
        self.point_hits.clear()
        self.label_hits.clear()
        self.seg_hits.clear()
        self.day_labels.clear()
        self._fit_gutter_width()
        self._draw_sheet(width, height)
        self._draw_gutter(height)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        self.gutter.configure(scrollregion=(0, 0, self.gutter_w, height))
        fraction = 0 if width <= 0 else clamp(0, restore_x / width, 1)
        self.canvas.xview_moveto(fraction)
        y_fraction = 0 if height <= 0 else clamp(0, restore_y / height, 1)
        self.canvas.yview_moveto(y_fraction)
        self.gutter.yview_moveto(y_fraction)
        self._update_vertical_scroll(height)
        self._suspend_scroll = False
        self._place_day_labels()
        self.update_status()

    def slot_label(self, slot: int) -> str:
        if 0 <= slot < len(self.slot_names) and self.slot_names[slot].strip():
            return self.slot_names[slot].strip()
        return str(slot + 1)

    def _fit_gutter_width(self) -> None:
        widest = 0
        for slot in range(self.slot_count):
            widest = max(widest, self.slot_measure.measure(self.slot_label(slot)))
        self.gutter_w = int(clamp(GUTTER_W, widest + 18, 200))
        self.gutter.configure(width=self.gutter_w)

    def _draw_gutter(self, height: int) -> None:
        self.gutter.create_rectangle(0, 0, self.gutter_w, AXIS_H, fill="#f6fafb", outline="")
        self.gutter.create_text(self.gutter_w / 2, AXIS_H / 2, text="スロット", font=self.font_small, fill="#5d6c74")
        self.gutter.create_rectangle(0, AXIS_H, self.gutter_w, self.content_top(), fill="#e7f3f2", outline="")
        self.gutter.create_text(self.gutter_w / 2, AXIS_H + EFFORT_H / 2, text="工数", font=self.font_label, fill="#0e6b67")
        active = self.drag.slot if self.drag and self.drag.kind == "task" else -1
        for slot in range(self.slot_count):
            y = self.content_top() + slot * self.slot_h
            fill = "#d9efee" if slot == active else "#ffffff" if slot % 2 == 0 else "#f7fbfb"
            self.gutter.create_rectangle(0, y, self.gutter_w, y + self.slot_h, fill=fill, outline="")
            self.gutter.create_text(self.gutter_w / 2, y + self.slot_h / 2, text=self.slot_label(slot), font=self.font_small, fill="#1b2428")
            self.gutter.create_line(0, y, self.gutter_w, y, fill="#e1e8ec")

    def _draw_sheet(self, width: float, height: int) -> None:
        self.canvas.create_rectangle(0, 0, width, height, fill="#ffffff", outline="")
        self.canvas.create_rectangle(0, 0, width, AXIS_H, fill="#f6fafb", outline="")
        self.canvas.create_rectangle(0, AXIS_H, width, self.content_top(), fill="#e7f3f2", outline="")
        active = self.drag.slot if self.drag and self.drag.kind == "task" else -1
        for slot in range(self.slot_count):
            y = self.content_top() + slot * self.slot_h
            fill = "#e7f4f3" if slot == active else "#ffffff" if slot % 2 == 0 else "#f7fbfb"
            self.canvas.create_rectangle(0, y, width, y + self.slot_h, fill=fill, outline="")
            self.canvas.create_line(0, y, width, y, fill="#e1e8ec")

        hours = int(round((self.axis_end - self.axis_start).total_seconds() / 3600))
        for index in range(hours + 1):
            moment = self.axis_start + timedelta(hours=index)
            x = self.time_to_x(moment)
            midnight = moment.hour == 0
            self.canvas.create_line(x, 0 if midnight else 28, x, height, fill="#c5ced6" if midnight else "#e7eef2")
            if index == hours:
                continue
            if self.px_per_hour < 480:
                half = x + self.px_per_hour / 2
                self.canvas.create_line(half, AXIS_H - 8, half, height, fill="#f1f5f7")
            if midnight:
                self.canvas.create_text(x + 8, 18, text=format_day(moment), anchor="w", font=self.font_day, fill="#1b2428")
            self.canvas.create_text(x + 6, 42, text=format_hm(moment), anchor="w", font=self.font_axis, fill="#5d6c74")
        self._draw_minute_marks(height)

        day_count = int(round(hours / 24))
        for day in range(day_count):
            day_start = add_days(self.axis_start, day)
            x0 = self.time_to_x(day_start)
            x1 = self.time_to_x(add_days(day_start, 1))
            item = self.canvas.create_text(x0 + 8, 18, text=format_day(day_start), anchor="w", font=self.font_day, fill="#1b2428")
            self.day_labels.append((item, x0, x1))

        self._draw_effort(width)
        ordered = self._draw_order()
        for task in ordered:
            self._draw_lines(task)
        for task in ordered:
            self._draw_label(task)
        for task in ordered:
            self._draw_points(task)
        if self.drag and self.drag.kind == "point":
            self._draw_point_guide(self.drag)
        if self.drag and self.drag.kind == "task" and self.drag.moved:
            self._draw_task_guide(self.drag)
        self._draw_pen_anchor()
        self._draw_pen_draft()

    def _draw_minute_marks(self, height: int) -> None:
        px_per_minute = self.px_per_hour / 60
        if px_per_minute < 8:
            return
        step = 1 if px_per_minute >= 16 else 5
        label_every = 1 if px_per_minute >= 36 else 5
        view_left = self.canvas.canvasx(0) - 40
        view_right = view_left + max(self.canvas.winfo_width(), 480) + 80
        start = self.x_to_time(max(PAD_L, view_left))
        end = min(self.axis_end, self.x_to_time(view_right))
        if end <= self.axis_start:
            return
        cursor = start.replace(second=0, microsecond=0)
        cursor = cursor.replace(minute=(cursor.minute // step) * step)
        if cursor < start:
            cursor += timedelta(minutes=step)
        while cursor < end:
            if cursor >= self.axis_start and cursor.minute != 0:
                x = self.time_to_x(cursor)
                on_five = cursor.minute % 5 == 0
                self.canvas.create_line(
                    x,
                    36 if on_five else AXIS_H - 6,
                    x,
                    height if on_five else AXIS_H,
                    fill="#e7eef2" if on_five else "#dbe3e8",
                )
                if cursor.minute % label_every == 0:
                    self.canvas.create_text(x + 3, 42, text=format_hm(cursor), anchor="w", font=self.font_axis, fill="#8aa0a8")
            cursor += timedelta(minutes=step)

    def _draw_effort(self, width: float) -> None:
        boundaries = set()
        for span in effort_spans(self.tasks):
            x1 = self.time_to_x(span.start)
            x2 = self.time_to_x(span.end)
            if x2 < x1:
                x1, x2 = x2, x1
            fill, text_fill = effort_color(span.total, self.alert_level())
            self.canvas.create_rectangle(x1, AXIS_H, x2, self.content_top(), fill=fill, outline="")
            label = format_effort(span.total)
            if x2 - x1 >= self.measure.measure(label) + 6:
                self.canvas.create_text((x1 + x2) / 2, AXIS_H + 21, text=label, font=self.font_label, fill=text_fill)
            boundaries.add(span.start)
            boundaries.add(span.end)
        for moment in boundaries:
            x = self.time_to_x(moment)
            self.canvas.create_line(x, AXIS_H - 10, x, self.content_top(), fill="#ffffff", width=3)
            self.canvas.create_line(x, AXIS_H - 10, x, self.content_top(), fill="#0e6b67", width=2)
        self.canvas.create_line(0, self.content_top(), width, self.content_top(), fill="#d5dee3")

    def _draw_order(self):
        dragging = self.drag.task_id if self.drag and self.drag.kind == "task" else None
        return sorted(
            self.tasks,
            key=lambda task: (task.id == dragging, task.slot, task.points[0].time),
        )

    def task_color(self, task) -> str:
        if task.color in COLORS:
            return task.color
        index = next((position for position, item in enumerate(self.tasks) if item.id == task.id), 0)
        return COLORS[index % len(COLORS)]

    def lane_cluster(self, task):
        start, end = task_span(task)
        peers = [item for item in self.tasks if item.slot == task.slot]
        peers.sort(key=lambda item: self._numeric_id(item.id))
        return [item for item in peers if ranges_touch(task_span(item), (start, end))]

    def lane_y(self, task) -> float:
        peers = self.lane_cluster(task)
        index = next((position for position, item in enumerate(peers) if item.id == task.id), 0)
        center = self.content_top() + task.slot * self.slot_h + self.slot_h / 2
        if len(peers) <= 1:
            return center
        spread = min(28, max(0, (self.slot_h - 36) / 2))
        step = (spread * 2) / (len(peers) - 1)
        return center - spread + index * step

    def _draw_lines(self, task) -> None:
        color = self.task_color(task)
        y = self.lane_y(task)
        for index, point in enumerate(task.points[:-1]):
            nxt = task.points[index + 1]
            x1 = self.time_to_x(point.time)
            x2 = self.time_to_x(nxt.time)
            style = task.segments[index].style if index < len(task.segments) else "solid"
            if style not in STYLES:
                style = "solid"
            if style not in {"dashed", "none"}:
                width = {"double": 12, "triple": 18, "quad": 24, "quint": 30}.get(style, 6)
                self.canvas.create_line(x1, y, x2, y, fill="#ffffff", width=width, capstyle="round")
            if self.selected and self.selected[0] == task.id and self.selected[1] == index:
                self.canvas.create_line(x1, y, x2, y, fill=mix(color, "#ffffff", 0.18), width=16, capstyle="round")
            if style == "none":
                self.canvas.create_line(x1, y, x2, y, fill="#c5ced6", width=1, dash=(2, 5))
            for offset in style_offsets(style):
                options = {
                    "fill": color,
                    "width": 3,
                    "capstyle": "butt" if style == "dashed" else "round",
                }
                if style == "dashed":
                    options["dash"] = (8, 6)
                self.canvas.create_line(x1, y + offset, x2, y + offset, **options)
            hit_x1 = x1 - 6 if x1 == x2 else x1
            hit_x2 = x2 + 6 if x1 == x2 else x2
            self.seg_hits.append({"task_id": task.id, "index": index, "x1": hit_x1, "x2": hit_x2, "y": y})

    def _draw_label(self, task) -> None:
        color = self.task_color(task)
        text = task.name
        width = self.measure.measure(text) + 16
        height = 22
        x1 = self.time_to_x(task.points[0].time)
        x2 = self.time_to_x(task.points[-1].time)
        x = self.label_x(x1, x2, width)
        y_line = self.lane_y(task)
        band = self.label_band(task)
        y = y_line - height - 8
        if y < band[0] + 2:
            y = min(band[1] - height - 2, y_line + 10)
        self.canvas.create_rectangle(x, y, x + width, y + height, fill="#ffffff", outline=color, width=1.5)
        self.canvas.create_text(x + 8, y + 11, text=text, anchor="w", font=self.font_label, fill="#1b2428")
        self.label_hits.append({"task_id": task.id, "x1": x, "y1": y, "x2": x + width, "y2": y + height})

    def _draw_points(self, task) -> None:
        color = self.task_color(task)
        y = self.lane_y(task)
        for point in task.points:
            x = self.time_to_x(point.time)
            self.canvas.create_oval(x - 6.5, y - 6.5, x + 6.5, y + 6.5, fill=color, outline="#ffffff", width=2)
            self.point_hits.append({"task_id": task.id, "point_id": point.id, "x": x, "y": y})

    def _draw_point_guide(self, drag) -> None:
        task = self.find_task(drag.task_id)
        point = next((item for item in task.points if item.id == drag.point_id), None) if task else None
        if point is None:
            return
        x = self.time_to_x(point.time)
        y = self.lane_y(task)
        self._guide(x, y, format_hm(point.time))

    def _draw_task_guide(self, drag) -> None:
        task = self.find_task(drag.task_id)
        if not task or not task.points:
            return
        first = task.points[0].time
        last = task.points[-1].time
        text = format_hm(first) if first == last else f"{format_hm(first)}–{format_hm(last)}"
        self._guide(self.time_to_x(first), self.lane_y(task), text)

    def _guide(self, x: float, y: float, text: str) -> None:
        bottom = self.content_top() + self.slot_count * self.slot_h
        self.canvas.create_line(x, self.content_top(), x, bottom, fill="#0e6b67", dash=(3, 4))
        width = self.measure.measure(text) + 16
        self.canvas.create_rectangle(x - width / 2, y - 36, x + width / 2, y - 16, fill="#1b2428", outline="")
        self.canvas.create_text(x, y - 26, text=text, font=self.font_label, fill="#ffffff")

    def label_x(self, x1: float, x2: float, width: float) -> float:
        left, right = min(x1, x2), max(x1, x2)
        limit = self.timeline_width - width - 8
        if right - left < width + 16:
            return clamp(8, left + (right - left - width) / 2, limit)
        span_right = max(left, right - width)
        view_left = self.canvas.canvasx(0) + 8
        view_width = max(160, self.canvas.winfo_width())
        view_right = self.canvas.canvasx(0) + view_width - width - 8
        if view_right < left or view_left > span_right:
            return clamp(8, left, limit)
        return clamp(left, view_left, span_right)

    def label_band(self, task):
        peers = self.lane_cluster(task)
        index = next((position for position, item in enumerate(peers) if item.id == task.id), 0)
        top = self.content_top() + task.slot * self.slot_h
        band = self.slot_h / max(1, len(peers))
        return top + band * index, top + band * (index + 1)

    def _place_day_labels(self) -> None:
        if self._suspend_scroll:
            return
        view = self.canvas.canvasx(8)
        for item, x0, x1 in self.day_labels:
            if view < x0 + 108 or view > x1 - 120:
                self.canvas.itemconfigure(item, state="hidden")
            else:
                self.canvas.itemconfigure(item, state="normal")
                self.canvas.coords(item, view, 18)

    def _on_xview(self, first, last) -> None:
        self.hbar.set(first, last)
        self._place_day_labels()

    def _on_yview(self, first, last) -> None:
        self.vbar.set(first, last)
        if not self._suspend_scroll:
            self.gutter.yview_moveto(float(first))

    def _update_vertical_scroll(self, height: int) -> None:
        if height > self.canvas.winfo_height() + 1:
            self.vbar.grid(row=0, column=2, sticky="ns")
        else:
            self.vbar.grid_remove()

    def update_status(self) -> None:
        if not self.tasks or self.axis_start is None:
            return
        points = sum(len(task.points) for task in self.tasks)
        where = f" / {self.sheet_name}" if self.sheet_name else ""
        skipped = f" ・ 読み飛ばし {self.skipped}行" if self.skipped else ""
        start = format_day(self.axis_start)
        end = format_day(add_days(self.axis_end, -1))
        span = start if start == end else f"{start} 〜 {end}"
        self.status.configure(text=f"{self.file_name}{where} ・ {span} ・ タスク {len(self.tasks)} ・ 点 {points}{skipped}")
        names_changed = list(self.slot_names) != list(self.saved_slot_names)
        if same_tasks(self.tasks, self.snapshot) and not names_changed:
            self.reset_button.state(["disabled"])
        else:
            self.reset_button.state(["!disabled"])
        self.set_save_enabled(True)
        self.pen_button.configure(state="normal")
        self.eraser_button.configure(state="normal")
        self.refresh_tool_buttons()

    def set_save_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.save_button.state(["!disabled"] if enabled else ["disabled"])
        last = self.save_menu.index("end")
        if last is None:
            return
        for index in range(last + 1):
            self.save_menu.entryconfigure(index, state=state)

    def find_task(self, task_id):
        return next((task for task in self.tasks if task.id == task_id), None)

    def find_point(self, task, point_id):
        return next((point for point in task.points if point.id == point_id), None)

    def hit_test(self, x: float, y: float):
        for hit in reversed(self.point_hits):
            if math.hypot(x - hit["x"], y - hit["y"]) <= 14:
                return ("point", hit)
        for hit in reversed(self.label_hits):
            if hit["x1"] <= x <= hit["x2"] and hit["y1"] <= y <= hit["y2"]:
                return ("label", hit)
        for hit in reversed(self.seg_hits):
            left, right = min(hit["x1"], hit["x2"]), max(hit["x1"], hit["x2"])
            if left - 2 <= x <= right + 2 and abs(y - hit["y"]) <= 11:
                return ("seg", hit)
        return None

    def on_press(self, event) -> None:
        self.canvas.focus_set()
        self.close_menu()
        self.hide_tip()
        if self.tool == "pen":
            self.on_pen_press(event)
            return
        if self.tool == "eraser":
            return
        _canvas_x, _canvas_y, _widget_x, _widget_y = self.pointer_canvas(event)
        found = self.hit_test(_canvas_x, _canvas_y)
        if found is None:
            self.drag = Drag("pan", start_root_x=event.x_root, origin_scroll=self.canvas.canvasx(0))
            self.canvas.configure(cursor="fleur")
            return
        kind, hit = found
        task = self.find_task(hit["task_id"])
        if task is None:
            return
        if kind == "point":
            point = self.find_point(task, hit["point_id"])
            if point is None:
                return
            self.drag = Drag(
                "point",
                task_id=task.id,
                point_id=point.id,
                origin_time=point.time,
                slot=task.slot,
            )
            self.canvas.configure(cursor="sb_h_double_arrow")
            return
        self.drag = Drag(
            "task",
            task_id=task.id,
            origin_slot=task.slot,
            origin_times=[point.time for point in task.points],
            slot=task.slot,
            start_root_x=event.x_root,
            start_root_y=event.y_root,
            start_canvas_x=_canvas_x,
            seg_index=hit["index"] if kind == "seg" else None,
        )
        self.canvas.configure(cursor="fleur")

    def on_drag_motion(self, event) -> None:
        drag = self.drag
        if drag is None:
            return
        if drag.kind == "pan":
            width = self.timeline_width or 1
            view = self.canvas.winfo_width()
            left = drag.origin_scroll - (event.x_root - drag.start_root_x)
            left = clamp(0, left, max(0, width - view))
            self.canvas.xview_moveto(left / width)
            drag.moved = True
            return
        if drag.kind == "point":
            self.move_point(drag, event)
            return
        if drag.kind == "task":
            distance = math.hypot(event.x_root - drag.start_root_x, event.y_root - drag.start_root_y)
            if not drag.moved and distance < 6:
                return
            self.move_task(drag, event)

    def move_task(self, drag, event) -> None:
        task = self.find_task(drag.task_id)
        if task is None:
            return
        _canvas_x, _canvas_y, widget_x, _widget_y = self.pointer_canvas(event)
        self.auto_scroll(widget_x)
        canvas_x, canvas_y, _widget_x, _widget_y = self.pointer_canvas(event)
        task.slot = self.slot_from_y(canvas_y)
        drag.slot = task.slot
        delta_px = canvas_x - drag.start_canvas_x
        delta_min = int(round((delta_px / self.px_per_hour) * 60))
        earliest = min(drag.origin_times)
        latest = max(drag.origin_times)
        min_delta = int(round((self.axis_start - earliest).total_seconds() / 60))
        max_delta = int(round(((self.axis_end - timedelta(minutes=1)) - latest).total_seconds() / 60))
        delta_min = int(clamp(min_delta, delta_min, max(min_delta, max_delta)))
        for point, origin in zip(task.points, drag.origin_times):
            point.time = origin + timedelta(minutes=delta_min)
        drag.moved = True
        self.redraw()

    def move_point(self, drag, event) -> None:
        task = self.find_task(drag.task_id)
        if task is None:
            return
        index = next((position for position, point in enumerate(task.points) if point.id == drag.point_id), -1)
        if index < 0:
            return
        task.slot = drag.slot
        _canvas_x, _canvas_y, widget_x, _widget_y = self.pointer_canvas(event)
        self.auto_scroll(widget_x)
        canvas_x, _canvas_y, _widget_x, _widget_y = self.pointer_canvas(event)
        point = task.points[index]
        previous = task.points[index - 1].time + timedelta(minutes=1) if index > 0 else self.axis_start
        nxt = (
            task.points[index + 1].time - timedelta(minutes=1)
            if index < len(task.points) - 1
            else self.axis_end - timedelta(minutes=1)
        )
        low, high = previous, nxt
        if low > high:
            low = high = point.time
        snapped = snap_minute(self.x_to_time(canvas_x))
        point.time = clamp_time(low, snapped, high)
        drag.moved = True
        self.redraw()

    def on_release(self, event) -> None:
        drag = self.drag
        if drag is None:
            return
        self.drag = None
        self.apply_tool_cursor()
        if drag.kind == "task" and not drag.moved and drag.seg_index is not None:
            self.open_menu(drag.task_id, drag.seg_index, event.x_root, event.y_root)
        elif drag.kind == "task" and not drag.moved and drag.seg_index is None:
            self.open_color_menu(drag.task_id, event.x_root, event.y_root)
        else:
            self.redraw()

    def on_hover(self, event) -> None:
        if self.drag:
            self.hide_tip()
            return
        if self.tool:
            self.apply_tool_cursor()
            found = self.hit_test(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
            if self.tool == "pen" and found is not None and found[0] == "point":
                self.show_tip(found[1], event)
            else:
                self.hide_tip()
            self._update_pen_preview(self.canvas.canvasx(event.x))
            return
        found = self.hit_test(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if found is None:
            self.canvas.configure(cursor="arrow")
            self.hide_tip()
            return
        kind, hit = found
        if kind == "point":
            self.canvas.configure(cursor="sb_h_double_arrow")
            self.show_tip(hit, event)
            return
        self.canvas.configure(cursor="fleur")
        self.hide_tip()

    def show_tip(self, hit, event) -> None:
        task = self.find_task(hit["task_id"])
        point = self.find_point(task, hit["point_id"]) if task else None
        if task is None or point is None:
            self.hide_tip()
            return
        self.tip_time.configure(text=f"{format_day(point.time)} {format_hm(point.time)}")
        self.tip_task.configure(text=task.name)
        self.tip_desc.configure(text=point.description or "説明はありません")
        self.tip.deiconify()
        self.tip.geometry(f"+{event.x_root + 14}+{event.y_root + 16}")

    def hide_tip(self) -> None:
        self.tip.withdraw()

    def open_menu(self, task_id, index, x_root, y_root) -> None:
        self.close_color_menu()
        self.selected = (task_id, index)
        task = self.find_task(task_id)
        current = task.segments[index].style if task and index < len(task.segments) else "solid"
        for style, button in self.menu_buttons.items():
            button.configure(relief="sunken" if style == current else "flat", bg="#e7f3f2" if style == current else "#ffffff")
        self.menu.deiconify()
        self._place_popup(self.menu, x_root, y_root)
        self.redraw()

    def open_color_menu(self, task_id, x_root, y_root) -> None:
        self.menu.withdraw()
        self.selected = None
        task = self.find_task(task_id)
        if task is None:
            return
        self.color_task_id = task_id
        current = task.color if task.color in COLORS else ""
        for color, swatch in self.color_swatches.items():
            swatch.configure(highlightbackground="#1b2428" if color == current else "#ffffff")
        self.color_menu.deiconify()
        self._place_popup(self.color_menu, x_root, y_root)

    def _place_popup(self, popup, x_root, y_root) -> None:
        popup.update_idletasks()
        width = popup.winfo_reqwidth()
        height = popup.winfo_reqheight()
        left = self.root.winfo_rootx()
        top = self.root.winfo_rooty()
        right = min(self.root.winfo_screenwidth(), left + self.root.winfo_width())
        bottom = min(self.root.winfo_screenheight(), top + self.root.winfo_height())
        x = x_root + 8
        y = y_root + 8
        if x + width > right - 4:
            x = max(left + 4, right - width - 4)
        if y + height > bottom - 4:
            y = y_root - height - 4
        if y < top + 4:
            y = top + 4
        if y + height > bottom - 4:
            y = max(top + 4, bottom - height - 4)
        popup.geometry(f"+{int(x)}+{int(y)}")
        popup.lift()

    def close_color_menu(self) -> bool:
        visible = self.color_menu.winfo_viewable()
        self.color_menu.withdraw()
        had = self.color_task_id is not None
        self.color_task_id = None
        return visible and had

    def close_menu(self) -> None:
        closed_color = self.close_color_menu()
        visible = self.menu.winfo_viewable()
        self.menu.withdraw()
        had = self.selected is not None
        self.selected = None
        if (visible and had or closed_color) and self.tasks and self.drag is None:
            self.redraw()

    def choose_color(self, color: str) -> None:
        task = self.find_task(self.color_task_id) if self.color_task_id else None
        self.close_color_menu()
        if task is None or color not in COLORS:
            return
        task.color = color
        self.redraw()

    def choose_style(self, style: str) -> None:
        if self.selected is None or style not in STYLES:
            return
        task = self.find_task(self.selected[0])
        index = self.selected[1]
        if task is None or index >= len(task.segments):
            self.close_menu()
            return
        task.segments[index].style = style
        self.selected = None
        self.menu.withdraw()
        self.redraw()

    def on_escape(self, _event=None) -> None:
        if self.slot_entry is not None:
            self.cancel_slot_edit()
            return "break"
        if self.name_box is not None:
            self.cancel_pen_line()
            return "break"
        if self.pen_anchor is not None:
            self.pen_anchor = None
            self.redraw()
            return "break"
        self.cancel_drag()
        self.close_menu()

    def cancel_drag(self) -> None:
        drag = self.drag
        if drag is None:
            return
        if drag.kind == "point":
            task = self.find_task(drag.task_id)
            point = self.find_point(task, drag.point_id) if task else None
            if point is not None:
                point.time = drag.origin_time
            if task is not None:
                task.slot = drag.slot
        elif drag.kind == "task":
            task = self.find_task(drag.task_id)
            if task is not None:
                task.slot = drag.origin_slot
                for point, origin in zip(task.points, drag.origin_times):
                    point.time = origin
        self.drag = None
        self.apply_tool_cursor()
        self.redraw()

    def on_arrow(self, event) -> None:
        step = self.px_per_hour if event.keysym == "Right" else -self.px_per_hour
        width = self.timeline_width or 1
        view = max(1, self.canvas.winfo_width())
        left = clamp(0, self.canvas.canvasx(0) + step, max(0, width - view))
        self.canvas.xview_moveto(left / width)
        return "break"

    def on_vertical_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def on_wheel(self, event) -> None:
        if event.state & 0x4:
            return
        if event.state & 0x1:
            return self.on_vertical_wheel(event)
        width = self.timeline_width
        view = self.canvas.winfo_width()
        if width <= view:
            return
        left = clamp(0, self.canvas.canvasx(0) - event.delta, max(0, width - view))
        self.canvas.xview_moveto(left / width)
        return "break"

    def on_zoom(self, event):
        self.zoom_by(event.x, event.delta)
        return "break"

    def zoom_by(self, widget_x: float, delta: float) -> None:
        if self.axis_start is None or not delta:
            return
        anchor = self.x_to_time(self.canvas.canvasx(widget_x))
        factor = math.exp(delta / 280)
        self.px_per_hour = clamp(MIN_PX_PER_HOUR, self.px_per_hour * factor, MAX_PX_PER_HOUR)
        restore = self.time_to_x(anchor) - widget_x
        self.redraw(restore)

    def pointer_canvas(self, event):
        widget_x = event.x_root - self.canvas.winfo_rootx()
        widget_y = event.y_root - self.canvas.winfo_rooty()
        return self.canvas.canvasx(widget_x), self.canvas.canvasy(widget_y), widget_x, widget_y

    def auto_scroll(self, widget_x: float) -> None:
        width = self.timeline_width or 1
        view = self.canvas.winfo_width()
        left = self.canvas.canvasx(0)
        if widget_x > view - 48:
            left += 18
        elif widget_x < 48:
            left -= 18
        else:
            return
        left = clamp(0, left, max(0, width - view))
        self.canvas.xview_moveto(left / width)

    def apply_tool_cursor(self) -> None:
        cursor = { "pen": "crosshair", "eraser": "X_cursor" }.get(self.tool, "arrow")
        try:
            self.canvas.configure(cursor=cursor)
        except tk.TclError:
            fallback = "crosshair" if self.tool == "pen" else "circle" if self.tool == "eraser" else "arrow"
            try:
                self.canvas.configure(cursor=fallback)
            except tk.TclError:
                self.canvas.configure(cursor="arrow")

    def refresh_tool_buttons(self) -> None:
        self.pen_button.configure(
            bg="#0e6b67" if self.tool == "pen" else "#ffffff",
            fg="#ffffff" if self.tool == "pen" else "#1b2428",
            relief="sunken" if self.tool == "pen" else "raised",
        )
        self.eraser_button.configure(
            bg="#9f1239" if self.tool == "eraser" else "#ffffff",
            fg="#ffffff" if self.tool == "eraser" else "#1b2428",
            relief="sunken" if self.tool == "eraser" else "raised",
        )
        self.apply_tool_cursor()

    def toggle_tool(self, name: str) -> None:
        self.clear_pen_draft()
        self.tool = None if self.tool == name else name
        self.refresh_tool_buttons()
        self.hide_tip()
        if self.tasks and self.axis_start is not None:
            self.redraw()

    def clear_pen_draft(self) -> None:
        self.pen_anchor = None
        self.pen_draft = None
        self._close_name_box()

    def on_pen_press(self, event) -> None:
        if self.name_box is not None or self.axis_start is None:
            return
        canvas_x, canvas_y, _widget_x, _widget_y = self.pointer_canvas(event)
        place = self._pen_place(canvas_x, canvas_y)
        if place is None:
            return
        slot, when = place
        if self.pen_anchor is None:
            self.pen_anchor = {"slot": slot, "time": when}
            self.redraw()
            return
        if slot != self.pen_anchor["slot"]:
            self.pen_anchor = {"slot": slot, "time": when}
            self.redraw()
            return
        target = self._pen_target_time(canvas_x, self.pen_anchor["time"])
        if target is None:
            return
        self.pen_draft = {
            "slot": slot,
            "start": self.pen_anchor["time"],
            "end": target,
            "after_id": None,
        }
        self.redraw()
        self.open_name_box(event.x_root, event.y_root)

    def _pen_place(self, canvas_x: float, canvas_y: float):
        top = self.content_top()
        bottom = top + self.slot_count * self.slot_h
        if canvas_y < top or canvas_y >= bottom:
            return None
        when = snap_minute(self.x_to_time(canvas_x))
        when = clamp_time(self.axis_start, when, self.axis_end - timedelta(minutes=1))
        return self.slot_from_y(canvas_y), when

    def _pen_target_time(self, canvas_x: float, anchor_time: datetime):
        anchor_x = self.time_to_x(anchor_time)
        if abs(canvas_x - anchor_x) < 2:
            return None
        direction = 1 if canvas_x > anchor_x else -1
        snapped = snap_minute(self.x_to_time(canvas_x))
        snapped = clamp_time(self.axis_start, snapped, self.axis_end - timedelta(minutes=1))
        if (snapped - anchor_time).total_seconds() * direction <= 0:
            snapped = anchor_time + timedelta(minutes=direction)
            snapped = clamp_time(self.axis_start, snapped, self.axis_end - timedelta(minutes=1))
        if snapped == anchor_time:
            return None
        return snapped

    def _update_pen_preview(self, canvas_x: float) -> None:
        self.canvas.delete("penpreview")
        anchor = self.pen_anchor
        if self.tool != "pen" or not isinstance(anchor, dict) or self.pen_draft is not None:
            return
        target = self._pen_target_time(canvas_x, anchor["time"])
        if target is None:
            return
        y = self.content_top() + anchor["slot"] * self.slot_h + self.slot_h / 2
        self.canvas.create_line(
            self.time_to_x(anchor["time"]),
            y,
            self.time_to_x(target),
            y,
            fill="#0e6b67",
            width=2,
            dash=(4, 3),
            tags="penpreview",
        )

    def _draw_pen_anchor(self) -> None:
        anchor = self.pen_anchor
        if not isinstance(anchor, dict) or self.pen_draft is not None:
            return
        x = self.time_to_x(anchor["time"])
        y = self.content_top() + anchor["slot"] * self.slot_h + self.slot_h / 2
        self.canvas.create_oval(x - 11, y - 11, x + 11, y + 11, outline="#0e6b67", width=3)

    def _draw_pen_draft(self) -> None:
        draft = self.pen_draft
        if not draft:
            return
        y = self.content_top() + draft["slot"] * self.slot_h + self.slot_h / 2
        x1 = self.time_to_x(draft["start"])
        x2 = self.time_to_x(draft["end"])
        self.canvas.create_line(x1, y, x2, y, fill="#0e6b67", width=3, capstyle="round")
        for x in (x1, x2):
            self.canvas.create_oval(x - 6, y - 6, x + 6, y + 6, fill="#0e6b67", outline="#ffffff", width=2)

    def open_name_box(self, x_root: int, y_root: int) -> None:
        self._close_name_box()
        box = tk.Toplevel(self.root)
        box.overrideredirect(True)
        box.configure(bg="#ffffff", highlightbackground="#0e6b67", highlightthickness=1)
        tk.Label(box, text="線の名前", font=self.font_label, bg="#ffffff", fg="#1b2428").pack(anchor="w", padx=10, pady=(8, 2))
        entry = tk.Entry(box, font=self.font_ui, width=22)
        entry.pack(padx=10, pady=(0, 6))
        actions = tk.Frame(box, bg="#ffffff")
        actions.pack(padx=10, pady=(0, 8))
        tk.Button(actions, text="決定", font=self.font_ui, command=self.confirm_pen_line, bg="#0e6b67", fg="#ffffff", relief="flat", padx=10).pack(side="left", padx=(0, 6))
        tk.Button(actions, text="キャンセル", font=self.font_ui, command=self.cancel_pen_line, relief="flat", padx=8).pack(side="left")
        entry.bind("<Return>", lambda _event: self.confirm_pen_line())
        box.bind("<Escape>", lambda _event: self.cancel_pen_line())
        self.name_box = box
        self.name_entry = entry
        box.update_idletasks()
        width = box.winfo_reqwidth()
        height = box.winfo_reqheight()
        left = self.root.winfo_rootx()
        top = self.root.winfo_rooty()
        right = min(self.root.winfo_screenwidth(), left + self.root.winfo_width())
        bottom = min(self.root.winfo_screenheight(), top + self.root.winfo_height())
        x = x_root + 8
        y = y_root + 8
        if x + width > right - 4:
            x = max(left + 4, right - width - 4)
        if y + height > bottom - 4:
            y = max(top + 4, y_root - height - 8)
        box.geometry(f"+{int(x)}+{int(y)}")
        box.lift()
        entry.focus_set()
        try:
            box.grab_set()
        except tk.TclError:
            pass

    def confirm_pen_line(self) -> None:
        if self.pen_draft is None or self.name_entry is None:
            return
        name = self.name_entry.get().strip()
        if not name:
            return
        draft = self.pen_draft
        self._append_pen_task(draft, name)
        self.pen_draft = None
        self.pen_anchor = None
        self._close_name_box()
        self.tool = None
        self.refresh_tool_buttons()
        self.redraw()

    def cancel_pen_line(self) -> None:
        self.pen_draft = None
        self.pen_anchor = None
        self._close_name_box()
        if self.tasks and self.axis_start is not None:
            self.redraw()

    def _close_name_box(self) -> None:
        box = self.name_box
        self.name_box = None
        self.name_entry = None
        if box is None:
            return
        try:
            box.grab_release()
        except tk.TclError:
            pass
        box.destroy()

    def _max_suffix(self, prefix: str) -> int:
        best = 0
        values = []
        for task in self.tasks:
            values.append(task.id)
            values.extend(point.id for point in task.points)
        for value in values:
            if not str(value).startswith(prefix):
                continue
            tail = str(value)[len(prefix) :]
            if tail.isdigit():
                best = max(best, int(tail))
        return best

    def _append_pen_task(self, draft, name: str) -> None:
        start, end = sorted((draft["start"], draft["end"]))
        task_id = f"t{self._max_suffix('t') + 1}"
        point_base = self._max_suffix("p")
        task = Task(
            id=task_id,
            name=name,
            slot=draft["slot"],
            slot_assigned=True,
            points=[
                Point(f"p{point_base + 1}", start, "", 0),
                Point(f"p{point_base + 2}", end, "", 0),
            ],
            segments=[Segment("solid")],
        )
        position = next((index for index, item in enumerate(self.tasks) if item.id == draft.get("after_id")), len(self.tasks) - 1)
        self.tasks.insert(position + 1, task)

    def on_canvas_double(self, event) -> None:
        if self.tool != "eraser" or self.axis_start is None:
            return
        canvas_x, canvas_y, _widget_x, _widget_y = self.pointer_canvas(event)
        found = self.hit_test(canvas_x, canvas_y)
        if found is None or found[0] != "seg":
            return
        hit = found[1]
        self.erase_segment(hit["task_id"], hit["index"])

    def erase_segment(self, task_id: str, index: int) -> None:
        task = self.find_task(task_id)
        if task is None or index < 0 or index >= len(task.segments):
            return
        left_points = list(task.points[: index + 1])
        right_points = list(task.points[index + 1 :])
        left_segments = list(task.segments[:index])
        right_segments = list(task.segments[index + 1 :])
        if len(left_points) <= 1:
            left_points, left_segments = [], []
        if len(right_points) <= 1:
            right_points, right_segments = [], []
        pieces = [(left_points, left_segments), (right_points, right_segments)]
        pieces = [(points, segments) for points, segments in pieces if points]
        if not pieces:
            self.tasks = [item for item in self.tasks if item.id != task.id]
        else:
            task.points, task.segments = pieces[0]
            if len(pieces) == 2:
                newbie = Task(
                    id=f"t{self._max_suffix('t') + 1}",
                    name=task.name,
                    slot=task.slot,
                    slot_assigned=True,
                    points=pieces[1][0],
                    segments=pieces[1][1],
                    color=task.color,
                )
                position = next(item_index for item_index, item in enumerate(self.tasks) if item.id == task.id)
                self.tasks.insert(position + 1, newbie)
        self.selected = None
        self.redraw()

    def on_gutter_double_click(self, event) -> None:
        if not self.tasks:
            return
        y = self.gutter.canvasy(event.y)
        if y < self.content_top():
            return
        self.start_slot_edit(self.slot_from_y(y))

    def start_slot_edit(self, slot: int) -> None:
        if slot < 0 or slot >= self.slot_count:
            return
        self.finish_slot_edit()
        top = self.content_top() + slot * self.slot_h - self.gutter.canvasy(0)
        entry = tk.Entry(self.gutter, font=self.font_small, justify="center", relief="solid", bd=1)
        entry.insert(0, self.slot_label(slot))
        entry.selection_range(0, "end")
        entry.place(x=2, y=top + 2, width=max(20, self.gutter_w - 4), height=max(16, self.slot_h - 4))
        entry.focus_set()
        entry.bind("<Return>", self.finish_slot_edit)
        entry.bind("<Escape>", lambda _event: self.cancel_slot_edit())
        entry.bind("<FocusOut>", self.finish_slot_edit)
        self.slot_entry = entry
        self.slot_edit_index = slot

    def cancel_slot_edit(self) -> None:
        entry = self.slot_entry
        self.slot_entry = None
        self.slot_edit_index = None
        if entry is not None:
            entry.destroy()

    def finish_slot_edit(self, _event=None) -> None:
        entry = self.slot_entry
        index = self.slot_edit_index
        if entry is None or index is None:
            return
        self.slot_entry = None
        self.slot_edit_index = None
        text = entry.get().strip()
        entry.destroy()
        if text == str(index + 1):
            text = ""
        if index >= len(self.slot_names):
            self.slot_names.extend([""] * (index + 1 - len(self.slot_names)))
        if self.slot_names[index] == text:
            return
        self.slot_names[index] = text
        self.redraw()

    def rename_slot(self, slot: int, name: str) -> None:
        if slot < 0 or slot >= self.slot_count:
            return
        if slot >= len(self.slot_names):
            self.slot_names.extend([""] * (slot + 1 - len(self.slot_names)))
        text = str(name or "").strip()
        if text == str(slot + 1):
            text = ""
        self.slot_names[slot] = text
        self.redraw()

    def slot_from_y(self, y: float) -> int:
        return int(clamp(0, math.floor((y - self.content_top()) / self.slot_h), self.slot_count - 1))

    def _numeric_id(self, value: str):
        parts = []
        chunk = ""
        for char in value:
            if char.isdigit():
                chunk += char
            else:
                if chunk:
                    parts.append(int(chunk))
                    chunk = ""
                parts.append(char)
        if chunk:
            parts.append(int(chunk))
        return parts


class Drag:
    def __init__(self, kind, **kwargs):
        self.kind = kind
        self.moved = False
        self.seg_index = None
        self.slot = 0
        for key, value in kwargs.items():
            setattr(self, key, value)


def ranges_touch(left, right) -> bool:
    return left[0] < right[1] and left[1] > right[0]


def style_offsets(style: str):
    if style == "none":
        return ()
    if style == "double":
        return (-4, 4)
    if style == "triple":
        return (-7, 0, 7)
    if style == "quad":
        return (-9, -3, 3, 9)
    if style == "quint":
        return (-12, -6, 0, 6, 12)
    return (0,)


def clamp_time(low: datetime, value: datetime, high: datetime) -> datetime:
    if value < low:
        return low
    if value > high:
        return high
    return value


def main() -> None:
    enable_dpi()
    root = tk.Tk()
    ScheduleApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
