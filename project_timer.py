#!/usr/bin/env python3
"""
Project Time Tracker — Standalone desktop app for Obsidian vault time tracking.

Reads and writes YAML frontmatter in Obsidian markdown notes (.md) to provide
a timer UI, weekly views, adjusted bookings, and a duration calculator.

Compatible with existing Obsidian DataviewJS queries — only modifies fields
that the JS already reads.
"""

import base64
import calendar
import io
import json
import math
import os
import struct
import sys
import tkinter as tk
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
import frontmatter
from dateutil import parser as dtparser


def _generate_clock_ico() -> bytes:
    """Generate a minimal 16x16 clock icon as .ico bytes.
    Draws a circle with hour/minute hands on a transparent background."""
    size = 16
    pixels = bytearray(size * size * 4)  # BGRA

    cx, cy, r = 7.5, 7.5, 6.5

    for y in range(size):
        for x in range(size):
            off = (y * size + x) * 4
            dx, dy = x - cx, y - cy
            dist = math.sqrt(dx * dx + dy * dy)

            if abs(dist - r) < 1.2:
                # Circle outline — white
                alpha = max(0, min(255, int(255 * (1.2 - abs(dist - r)))))
                pixels[off:off + 4] = bytes([255, 255, 255, alpha])  # BGRA
            elif dist < r - 0.5:
                # Inside — dark fill
                pixels[off:off + 4] = bytes([60, 50, 40, 220])  # BGRA
            else:
                pixels[off:off + 4] = bytes([0, 0, 0, 0])

    # Draw clock hands (minute hand pointing to 12, hour hand to 2)
    import math as m
    hands = [
        (cx, cy, cx, cy - 5, 255),       # minute hand — straight up
        (cx, cy, cx + 3, cy - 3, 200),   # hour hand — ~2 o'clock
    ]
    for x0, y0, x1, y1, brightness in hands:
        steps = 20
        for t in range(steps + 1):
            px = x0 + (x1 - x0) * t / steps
            py = y0 + (y1 - y0) * t / steps
            ix, iy = int(round(px)), int(round(py))
            if 0 <= ix < size and 0 <= iy < size:
                off = (iy * size + ix) * 4
                pixels[off:off + 4] = bytes([brightness, brightness, brightness, 255])

    # Build .ico file
    bmp_data = bytes(pixels)
    # ICO header
    ico = bytearray()
    ico += struct.pack('<HHH', 0, 1, 1)  # reserved, type=icon, count=1
    # Directory entry
    bmp_size = 40 + len(bmp_data)
    ico += struct.pack('<BBBBHHII', size, size, 0, 0, 1, 32, bmp_size, 22)
    # BITMAPINFOHEADER
    ico += struct.pack('<IiiHHIIiiII', 40, size, size * 2, 1, 32, 0, len(bmp_data), 0, 0, 0, 0)
    # Pixel data — ICO BMPs are stored bottom-up
    for row in range(size - 1, -1, -1):
        ico += bmp_data[row * size * 4:(row + 1) * size * 4]

    return bytes(ico)

# ─── Constants ────────────────────────────────────────────────────────────────

CONFIG_FILE = "timer_config.json"
DEFAULT_CONFIG = {
    "vault_path": r"C:\Users\jacob.hand\OneDrive - Stockport Metropolitan Borough Council\Documents\Jacob Hand SMBC PKM\Slip Box",
    "window_width": 600,
    "window_height": 400,
    "window_x": 100,
    "window_y": 100,
}

EFFORT_ORDER = [
    ("1 - Owned", "Owned"),
    ("2 - High", "High"),
    ("3 - Medium", "Medium"),
    ("4 - Low", "Low"),
]
EFFORT_LABELS = {k: v for k, v in EFFORT_ORDER}

COLOR_IDLE = "#28a745"
COLOR_RUNNING = "#dc3545"
DAYS_OF_WEEK = ["Mon", "Tue", "Wed", "Thu", "Fri"]

# Table colors
TBL_BG = "#1e1e2e"
TBL_ROW_EVEN = "#252535"
TBL_ROW_ODD = "#2b2b3b"
TBL_HEADER_BG = "#333348"
TBL_HEADER_FG = "#c0c0d0"
TBL_TEXT = "#e0e0e0"
TBL_TOTALS_BG = "#3a3a50"
TBL_TOTALS_FG = "#ffffff"

# Calendar popup colors
CAL_BG = "#1e1e2e"
CAL_HEADER_FG = "#e0e0e0"
CAL_DOW_FG = "#888899"
CAL_DAY_FG = "#e0e0e0"
CAL_DAY_DISABLED = "#444455"
CAL_DAY_HOVER = "#333348"
CAL_DAY_SELECTED_BG = "#3b82f6"
CAL_DAY_TODAY_FG = "#3b82f6"
CAL_BORDER = "#444458"

# ─── Utility Functions ────────────────────────────────────────────────────────


def config_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / CONFIG_FILE


def load_config() -> dict:
    cp = config_path()
    if cp.exists():
        try:
            with open(cp, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"Warning: could not save config: {e}")


def parse_iso(s) -> datetime | None:
    if not s:
        return None
    try:
        dt = dtparser.parse(str(s))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        print(f"Warning: malformed ISO timestamp: {s}")
        return None


def format_hm(minutes: int | float) -> str:
    m = int(round(minutes))
    if m <= 0:
        return "-"
    h, rem = divmod(m, 60)
    if h and rem:
        return f"{h}h {rem}m"
    if h:
        return f"{h}h"
    return f"{rem}m"


def format_hmm(minutes: int | float) -> str:
    m = int(round(minutes))
    h, rem = divmod(m, 60)
    return f"{h}:{rem:02d}"


def format_hhmmss(total_seconds: float) -> str:
    ts = max(0, int(total_seconds))
    h, rem = divmod(ts, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}"


def week_bounds_local(ref_date: datetime) -> tuple[datetime, datetime]:
    if ref_date.tzinfo:
        ref_date = ref_date.astimezone(tz=None).replace(tzinfo=None)
    monday = ref_date - timedelta(days=ref_date.weekday())
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    sunday = monday + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)
    return monday, sunday


def session_day_local(session: dict) -> datetime | None:
    et = parse_iso(session.get("end_time"))
    if et:
        return et.astimezone(tz=None).replace(tzinfo=None)
    st = parse_iso(session.get("start_time"))
    if st:
        return st.astimezone(tz=None).replace(tzinfo=None)
    return None


def safe_write_frontmatter(filepath: Path, post: frontmatter.Post):
    tmp = filepath.with_suffix(".tmp")
    content = frontmatter.dumps(post)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(content)
    tmp.replace(filepath)


# ─── Calendar Date Picker ─────────────────────────────────────────────────────
# Uses plain tk.Toplevel instead of CTkToplevel to avoid the known Windows bug
# where CTkToplevel with overrideredirect(True) creates unkillable ghost windows.


class CalendarPicker(ctk.CTkFrame):
    """Date picker with entry field and dropdown calendar popup.
    Only allows selecting Mondays when mondays_only=True."""

    def __init__(self, master, mondays_only=True, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.mondays_only = mondays_only
        self._popup = None

        today = datetime.now()
        mon = today - timedelta(days=today.weekday())
        self._selected_date = mon if mondays_only else today
        self._view_year = self._selected_date.year
        self._view_month = self._selected_date.month

        self.date_entry = ctk.CTkEntry(self, width=120, justify="center")
        self.date_entry.pack(side="left", padx=(0, 4))
        self.date_entry.insert(0, self._selected_date.strftime("%d/%m/%Y"))

        self.cal_btn = ctk.CTkButton(
            self, text="📅", width=32, height=28,
            fg_color="#444455", hover_color="#555566",
            command=self._toggle_popup,
        )
        self.cal_btn.pack(side="left")

    def get_date(self) -> datetime | None:
        text = self.date_entry.get().strip()
        try:
            dt = datetime.strptime(text, "%d/%m/%Y")
        except ValueError:
            return None
        if self.mondays_only and dt.weekday() != 0:
            return None
        return dt

    def set_date(self, dt: datetime):
        self._selected_date = dt
        self._view_year = dt.year
        self._view_month = dt.month
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, dt.strftime("%d/%m/%Y"))

    def _toggle_popup(self):
        if self._popup is not None:
            try:
                if self._popup.winfo_exists():
                    self._popup.destroy()
            except Exception:
                pass
            self._popup = None
            return
        self._show_popup()

    def _show_popup(self):
        # Sync view month to entry text
        text = self.date_entry.get().strip()
        try:
            dt = datetime.strptime(text, "%d/%m/%Y")
            self._view_year = dt.year
            self._view_month = dt.month
        except ValueError:
            pass

        # Use plain tk.Toplevel — avoids CTkToplevel ghost window bug on Windows
        self._popup = tk.Toplevel(self)
        self._popup.overrideredirect(True)
        self._popup.configure(bg=CAL_BG)
        self._popup.attributes("-topmost", True)

        # Position below the entry
        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        self._popup.geometry(f"+{x}+{y}")

        self._build_calendar()

        # Close when clicking anywhere outside the popup
        self._popup.bind("<FocusOut>", self._schedule_close_check)
        self._popup.focus_set()

    def _schedule_close_check(self, event=None):
        if self._popup:
            self._popup.after(200, self._close_if_lost_focus)

    def _close_if_lost_focus(self):
        try:
            if not self._popup or not self._popup.winfo_exists():
                return
            # Check if focus is still within the popup
            focused = self._popup.focus_get()
            if focused is None:
                self._close_popup()
                return
            # Walk up the widget hierarchy to see if focus is in popup
            w = focused
            while w is not None:
                if w == self._popup:
                    return  # Focus still inside popup
                w = w.master
            self._close_popup()
        except Exception:
            self._close_popup()

    def _close_popup(self):
        if self._popup:
            try:
                self._popup.destroy()
            except Exception:
                pass
            self._popup = None

    def _build_calendar(self):
        if not self._popup:
            return

        # Clear previous content
        for w in self._popup.winfo_children():
            w.destroy()

        outer = tk.Frame(self._popup, bg=CAL_BORDER, padx=1, pady=1)
        outer.pack()
        frame = tk.Frame(outer, bg=CAL_BG, padx=8, pady=8)
        frame.pack()

        # ── Header row: ‹  Month Year  › ──
        header = tk.Frame(frame, bg=CAL_BG)
        header.pack(fill="x", pady=(0, 6))

        prev_btn = tk.Button(
            header, text="‹", font=("Segoe UI", 14), width=2,
            bg=CAL_BG, fg=CAL_HEADER_FG, activebackground=CAL_DAY_HOVER,
            activeforeground=CAL_HEADER_FG, bd=0, relief="flat",
            command=self._prev_month,
        )
        prev_btn.pack(side="left")

        month_name = calendar.month_name[self._view_month]
        tk.Label(
            header, text=f"{month_name} {self._view_year}",
            font=("Segoe UI", 11, "bold"), bg=CAL_BG, fg=CAL_HEADER_FG,
        ).pack(side="left", expand=True)

        next_btn = tk.Button(
            header, text="›", font=("Segoe UI", 14), width=2,
            bg=CAL_BG, fg=CAL_HEADER_FG, activebackground=CAL_DAY_HOVER,
            activeforeground=CAL_HEADER_FG, bd=0, relief="flat",
            command=self._next_month,
        )
        next_btn.pack(side="right")

        # ── Day-of-week labels ──
        dow_frame = tk.Frame(frame, bg=CAL_BG)
        dow_frame.pack(fill="x")
        for d in ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]:
            tk.Label(
                dow_frame, text=d, width=4, font=("Segoe UI", 9),
                bg=CAL_BG, fg=CAL_DOW_FG,
            ).pack(side="left", padx=1)

        # ── Day grid ──
        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdayscalendar(self._view_year, self._view_month)
        today = datetime.now().date()

        for week in weeks:
            row = tk.Frame(frame, bg=CAL_BG)
            row.pack(fill="x")

            for day_num in week:
                if day_num == 0:
                    tk.Label(row, text="", width=4, bg=CAL_BG).pack(side="left", padx=1, pady=1)
                    continue

                dt = datetime(self._view_year, self._view_month, day_num)
                is_monday = dt.weekday() == 0
                is_today = dt.date() == today
                is_selected = self._selected_date and dt.date() == self._selected_date.date()
                selectable = (not self.mondays_only) or is_monday

                if is_selected:
                    bg = CAL_DAY_SELECTED_BG
                    fg = "#ffffff"
                elif is_today:
                    bg = CAL_BG
                    fg = CAL_DAY_TODAY_FG
                elif selectable:
                    bg = CAL_BG
                    fg = CAL_DAY_FG
                else:
                    bg = CAL_BG
                    fg = CAL_DAY_DISABLED

                if selectable:
                    btn = tk.Button(
                        row, text=str(day_num), width=4,
                        font=("Segoe UI", 9, "bold" if is_selected else "normal"),
                        bg=bg, fg=fg, activebackground=CAL_DAY_HOVER,
                        activeforeground="#ffffff", bd=0, relief="flat",
                        command=lambda d=dt: self._pick_date(d),
                    )
                    # Hover effect
                    if not is_selected:
                        btn.bind("<Enter>", lambda e, b=btn: b.config(bg=CAL_DAY_HOVER))
                        btn.bind("<Leave>", lambda e, b=btn, c=bg: b.config(bg=c))
                    btn.pack(side="left", padx=1, pady=1)
                else:
                    lbl = tk.Label(
                        row, text=str(day_num), width=4,
                        font=("Segoe UI", 9), bg=bg, fg=fg,
                    )
                    lbl.pack(side="left", padx=1, pady=1)

        # ── Bottom: Clear / Today ──
        bottom = tk.Frame(frame, bg=CAL_BG)
        bottom.pack(fill="x", pady=(6, 0))

        tk.Button(
            bottom, text="Clear", font=("Segoe UI", 9),
            bg=CAL_BG, fg=CAL_DOW_FG, activebackground=CAL_DAY_HOVER,
            activeforeground="#ffffff", bd=0, relief="flat",
            command=self._clear_date,
        ).pack(side="left")

        tk.Button(
            bottom, text="Today", font=("Segoe UI", 9),
            bg=CAL_BG, fg=CAL_DAY_TODAY_FG, activebackground=CAL_DAY_HOVER,
            activeforeground="#ffffff", bd=0, relief="flat",
            command=self._go_today,
        ).pack(side="right")

    def _prev_month(self):
        if self._view_month == 1:
            self._view_month = 12
            self._view_year -= 1
        else:
            self._view_month -= 1
        self._build_calendar()

    def _next_month(self):
        if self._view_month == 12:
            self._view_month = 1
            self._view_year += 1
        else:
            self._view_month += 1
        self._build_calendar()

    def _pick_date(self, dt: datetime):
        self._selected_date = dt
        self.date_entry.delete(0, "end")
        self.date_entry.insert(0, dt.strftime("%d/%m/%Y"))
        self._close_popup()

    def _clear_date(self):
        self.date_entry.delete(0, "end")
        self._close_popup()

    def _go_today(self):
        today = datetime.now()
        if self.mondays_only:
            today = today - timedelta(days=today.weekday())
        self._pick_date(today)


# ─── Dark Table Widget ────────────────────────────────────────────────────────


class DarkTable(ctk.CTkFrame):
    """Dark-themed table with both vertical and horizontal scrolling.
    Built with tk.Canvas + CTk labels. No ttk Treeview."""

    def __init__(self, master, headers: list[str], col_widths: list[int] | None = None,
                 col_anchors: list[str] | None = None, **kwargs):
        super().__init__(master, fg_color=TBL_BG, **kwargs)
        self.headers = headers
        self.col_count = len(headers)
        self.col_widths = col_widths or [120] * self.col_count
        self.col_anchors = col_anchors or ["w"] + ["center"] * (self.col_count - 1)
        self._row_count = 0
        self._total_width = sum(self.col_widths) + self.col_count * 8 + 20

        # Canvas + scrollbars
        self._canvas = tk.Canvas(self, bg=TBL_BG, highlightthickness=0)
        self._v_scroll = ctk.CTkScrollbar(self, command=self._canvas.yview)
        self._h_scroll = ctk.CTkScrollbar(self, orientation="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=self._v_scroll.set, xscrollcommand=self._h_scroll.set)

        self._h_scroll.pack(side="bottom", fill="x")
        self._v_scroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Inner frame inside canvas
        self._inner = ctk.CTkFrame(self._canvas, fg_color=TBL_BG)
        self._canvas_window = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # Mouse wheel scrolling
        self._canvas.bind("<Enter>", self._bind_mousewheel)
        self._canvas.bind("<Leave>", self._unbind_mousewheel)

        # Header row
        hdr_frame = ctk.CTkFrame(self._inner, fg_color=TBL_HEADER_BG, corner_radius=0)
        hdr_frame.pack(fill="x", padx=0, pady=(0, 1))
        for i, h in enumerate(headers):
            ctk.CTkLabel(
                hdr_frame, text=h, width=self.col_widths[i],
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color=TBL_HEADER_FG, anchor=self.col_anchors[i],
            ).pack(side="left", padx=4, pady=4)

    def _on_inner_configure(self, event=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        # Ensure inner frame is at least as wide as content
        canvas_width = event.width if event else self._canvas.winfo_width()
        min_width = max(canvas_width, self._total_width)
        self._canvas.itemconfig(self._canvas_window, width=min_width)

    def _bind_mousewheel(self, event=None):
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _unbind_mousewheel(self, event=None):
        self._canvas.unbind_all("<MouseWheel>")
        self._canvas.unbind_all("<Shift-MouseWheel>")

    def _on_mousewheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self._canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_row(self, values: list[str], is_totals: bool = False):
        bg = TBL_TOTALS_BG if is_totals else (TBL_ROW_EVEN if self._row_count % 2 == 0 else TBL_ROW_ODD)
        fg = TBL_TOTALS_FG if is_totals else TBL_TEXT
        weight = "bold" if is_totals else "normal"

        row_frame = ctk.CTkFrame(self._inner, fg_color=bg, corner_radius=0)
        row_frame.pack(fill="x", padx=0, pady=0)

        for i, val in enumerate(values):
            if i == 0:
                # Project name column — use tk.Label with wraplength for long names
                lbl = tk.Label(
                    row_frame, text=val, width=0,
                    font=("Segoe UI", 10 if not is_totals else 10, weight),
                    bg=bg, fg=fg, anchor="w", justify="left",
                    wraplength=self.col_widths[i],
                )
                lbl.pack(side="left", padx=4, pady=3, ipadx=0)
                lbl.configure(width=self.col_widths[i] // 7)  # approx char width
            else:
                ctk.CTkLabel(
                    row_frame, text=val, width=self.col_widths[i],
                    font=ctk.CTkFont(size=11, weight=weight),
                    text_color=fg, anchor=self.col_anchors[i],
                ).pack(side="left", padx=4, pady=3)

        self._row_count += 1


# ─── Project Data ─────────────────────────────────────────────────────────────


class ProjectNote:
    def __init__(self, filepath: Path):
        self.filepath = filepath
        self.name = filepath.stem
        self.post: frontmatter.Post | None = None
        self.load()

    def load(self):
        with open(self.filepath, "r", encoding="utf-8") as f:
            self.post = frontmatter.load(f)

    @property
    def meta(self) -> dict:
        return self.post.metadata

    @property
    def cls(self) -> str:
        return str(self.meta.get("Class", ""))

    @property
    def status(self) -> str:
        return str(self.meta.get("Status", ""))

    @property
    def effort(self) -> str:
        return str(self.meta.get("Effort", ""))

    @property
    def effort_group(self) -> str:
        e = self.effort.strip().strip('"').strip("'")
        return EFFORT_LABELS.get(e, "Unassigned")

    @property
    def timer_running(self) -> bool:
        return bool(self.meta.get("timer_running", False))

    @property
    def session_start(self) -> datetime | None:
        return parse_iso(self.meta.get("session_start"))

    @property
    def sessions(self) -> list[dict]:
        s = self.meta.get("time_sessions")
        if isinstance(s, list):
            return s
        return []

    def minutes_in_range(self, start: datetime, end: datetime) -> float:
        total = 0.0
        for sess in self.sessions:
            dt = session_day_local(sess)
            if dt is None:
                continue
            if start <= dt <= end:
                dur = sess.get("duration")
                if dur is not None:
                    total += float(dur)
        return total

    def minutes_per_day(self, monday: datetime) -> list[float]:
        days = [0.0] * 5
        for sess in self.sessions:
            dt = session_day_local(sess)
            if dt is None:
                continue
            idx = (dt.date() - monday.date()).days
            if 0 <= idx <= 4:
                dur = sess.get("duration")
                if dur is not None:
                    days[idx] += float(dur)
        return days

    def start_timer(self):
        now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.") + \
                  f"{datetime.now(timezone.utc).microsecond // 1000:03d}Z"
        self.meta["timer_running"] = True
        self.meta["session_start"] = now_iso
        safe_write_frontmatter(self.filepath, self.post)

    def stop_timer(self):
        start = self.session_start
        if start is None:
            self.meta["timer_running"] = False
            self.meta["session_start"] = None
            safe_write_frontmatter(self.filepath, self.post)
            return 0

        now = datetime.now(timezone.utc)
        duration = max(1, round((now - start).total_seconds() / 60))

        now_iso = now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}Z"
        start_iso = self.meta["session_start"]

        session = {
            "start_time": start_iso,
            "end_time": now_iso,
            "duration": duration,
        }

        if not isinstance(self.meta.get("time_sessions"), list):
            self.meta["time_sessions"] = []
        self.meta["time_sessions"].append(session)

        prev_total = self.meta.get("total_time_minutes", 0) or 0
        self.meta["total_time_minutes"] = prev_total + duration
        self.meta["timer_running"] = False
        self.meta["session_start"] = None

        safe_write_frontmatter(self.filepath, self.post)
        return duration


def scan_projects(vault_path: str) -> list[ProjectNote]:
    projects = []
    vp = Path(vault_path)
    if not vp.exists():
        return projects
    for md in sorted(vp.rglob("*.md")):
        try:
            p = ProjectNote(md)
            if p.cls == "Project":
                projects.append(p)
        except Exception as e:
            print(f"Warning: could not load {md}: {e}")
    return projects


# ─── Adjusted Bookings Algorithm ─────────────────────────────────────────────


def calculate_adjusted_bookings(
    projects: list[ProjectNote],
    monday: datetime,
    target_minutes: int = 2250,
) -> dict:
    """
    Compute adjusted bookings for a given week.

    Args:
        projects: list of ProjectNote objects
        monday: the Monday starting the week
        target_minutes: total weekly target in minutes (default 2250 = 37.5h)

    Returns dict with keys:
        rows, excluded, actual_grand, multiplier, target_minutes
    """
    target_daily = target_minutes // 5

    # Gather raw data
    raw = []
    for p in projects:
        days = p.minutes_per_day(monday)
        total = sum(days)
        if total > 0:
            raw.append({"name": p.name, "days": days, "total": total})

    # ── Step 1: Filter — Remove projects with week total < 15 minutes ──
    excluded = [r["name"] for r in raw if r["total"] < 15]
    filtered = [r for r in raw if r["total"] >= 15]

    if not filtered:
        return {
            "rows": [],
            "excluded": excluded,
            "actual_grand": sum(r["total"] for r in raw),
            "multiplier": 0,
            "target_minutes": target_minutes,
        }

    actual_grand = sum(r["total"] for r in filtered)

    # ── Step 2: Scale — multiplier = target / totalActualMinutes ──
    multiplier = target_minutes / actual_grand if actual_grand > 0 else 1.0
    for r in filtered:
        r["adj_days"] = [v * multiplier for v in r["days"]]
        r["adj_total"] = r["total"] * multiplier

    # ── Step 3: Round to 15 mins ──
    for r in filtered:
        r["final_days"] = [round(v / 15) * 15 for v in r["adj_days"]]
        r["final_total"] = round(r["adj_total"] / 15) * 15

        day_sum = sum(r["final_days"])
        diff = r["final_total"] - day_sum
        if diff != 0:
            max_idx = max(range(5), key=lambda i: r["adj_days"][i])
            r["final_days"][max_idx] += diff

    # ── Step 4: Enforce target total ──
    grand = sum(r["final_total"] for r in filtered)
    if grand != target_minutes:
        diff = target_minutes - grand
        top = max(filtered, key=lambda r: r["final_total"])
        worked_indices = [i for i in range(5) if top["final_days"][i] > 0]
        if worked_indices:
            increments = round(diff / 15)
            per_day = increments // len(worked_indices)
            remainder = increments % len(worked_indices)
            if per_day != 0 or remainder != 0:
                if abs(per_day) > 0:
                    for i in worked_indices:
                        top["final_days"][i] += per_day * 15
                if remainder != 0:
                    best_idx = max(worked_indices, key=lambda i: top["final_days"][i])
                    top["final_days"][best_idx] += remainder * 15
                top["final_total"] = sum(top["final_days"])

    # ── Step 5: Balance days to target_daily each ──
    for _iteration in range(20):
        balanced = True
        for day_idx in range(5):
            day_sum = sum(r["final_days"][day_idx] for r in filtered)
            if day_sum != target_daily:
                balanced = False
                diff = target_daily - day_sum
                step = 15 if diff > 0 else -15
                target_row = max(filtered, key=lambda r: r["final_days"][day_idx])
                target_row["final_days"][day_idx] += step
                target_row["final_total"] = sum(target_row["final_days"])
        if balanced:
            break

    rows = sorted(filtered, key=lambda r: r["final_total"], reverse=True)
    result_rows = []
    for r in rows:
        result_rows.append({
            "name": r["name"],
            "actual_days": r["days"],
            "actual_total": r["total"],
            "final_days": r["final_days"],
            "final_total": r["final_total"],
        })

    return {
        "rows": result_rows,
        "excluded": excluded,
        "actual_grand": actual_grand,
        "multiplier": multiplier,
        "target_minutes": target_minutes,
    }


# ─── GUI Application ─────────────────────────────────────────────────────────


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.config = load_config()
        self.title("Project Time Tracker")
        self.geometry(
            f"{self.config['window_width']}x{self.config['window_height']}"
            f"+{self.config['window_x']}+{self.config['window_y']}"
        )
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Set clock icon
        try:
            ico_data = _generate_clock_ico()
            import tempfile
            ico_path = Path(tempfile.gettempdir()) / "ptt_clock.ico"
            with open(ico_path, "wb") as f:
                f.write(ico_data)
            self.iconbitmap(str(ico_path))
        except Exception:
            pass  # Fall back to default icon

        self.projects: list[ProjectNote] = []
        self.running_project: ProjectNote | None = None
        self._timer_after_id = None
        self._refresh_after_id = None

        self._build_ui()
        self._initial_load()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI Construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # Top row: tabs on left, settings+refresh on right
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=4, pady=(2, 0))

        ctk.CTkButton(
            top_bar, text="⚙", width=24, height=20,
            font=ctk.CTkFont(size=9),
            command=self._open_settings,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="right", padx=(2, 0))

        ctk.CTkButton(
            top_bar, text="🔄", width=24, height=20,
            font=ctk.CTkFont(size=9),
            command=self._refresh_projects,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="right")

        # Tab view
        self.tabview = ctk.CTkTabview(self, height=28)
        self.tabview.pack(fill="both", expand=True, padx=4, pady=(0, 0))
        self.tabview._segmented_button.configure(font=ctk.CTkFont(size=11))

        self.tab_timer = self.tabview.add("Timer")
        self.tab_weekly = self.tabview.add("Weekly View")
        self.tab_adjusted = self.tabview.add("Adjusted Bookings")
        self.tab_calc = self.tabview.add("Duration Calculator")

        self._build_timer_tab()
        self._build_weekly_tab()
        self._build_adjusted_tab()
        self._build_calc_tab()

        # Compact status bar
        self.status_frame = ctk.CTkFrame(self, height=24, corner_radius=0, fg_color="#1a1a2e")
        self.status_frame.pack(fill="x", side="bottom")
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame, text="No timer running",
            font=ctk.CTkFont(size=11), anchor="w"
        )
        self.status_label.pack(fill="x", padx=8, pady=2)
        self.status_frame.bind("<Button-1>", self._status_bar_click)
        self.status_label.bind("<Button-1>", self._status_bar_click)

    def _build_timer_tab(self):
        # Dual-scroll frame for timer buttons
        timer_outer = ctk.CTkFrame(self.tab_timer, fg_color="transparent")
        timer_outer.pack(fill="both", expand=True, padx=2, pady=2)

        self._timer_canvas = tk.Canvas(timer_outer, bg="#2b2b2b", highlightthickness=0)
        timer_v = ctk.CTkScrollbar(timer_outer, command=self._timer_canvas.yview)
        timer_h = ctk.CTkScrollbar(timer_outer, orientation="horizontal", command=self._timer_canvas.xview)
        self._timer_canvas.configure(yscrollcommand=timer_v.set, xscrollcommand=timer_h.set)

        timer_h.pack(side="bottom", fill="x")
        timer_v.pack(side="right", fill="y")
        self._timer_canvas.pack(side="left", fill="both", expand=True)

        self.timer_scroll = ctk.CTkFrame(self._timer_canvas, fg_color="transparent")
        self._timer_canvas_window = self._timer_canvas.create_window((0, 0), window=self.timer_scroll, anchor="nw")

        self.timer_scroll.bind("<Configure>", lambda e: self._timer_canvas.configure(
            scrollregion=self._timer_canvas.bbox("all")
        ))
        self._timer_canvas.bind("<Configure>", self._on_timer_canvas_configure)
        self._timer_canvas.bind("<Enter>", lambda e: (
            self._timer_canvas.bind_all("<MouseWheel>", lambda ev: self._timer_canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")),
            self._timer_canvas.bind_all("<Shift-MouseWheel>", lambda ev: self._timer_canvas.xview_scroll(int(-1 * (ev.delta / 120)), "units")),
        ))
        self._timer_canvas.bind("<Leave>", lambda e: (
            self._timer_canvas.unbind_all("<MouseWheel>"),
            self._timer_canvas.unbind_all("<Shift-MouseWheel>"),
        ))

        self.timer_buttons: dict[str, ctk.CTkButton] = {}

    def _build_weekly_tab(self):
        ctrl = ctk.CTkFrame(self.tab_weekly, fg_color="transparent")
        ctrl.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(ctrl, text="Week start:").pack(side="left", padx=(0, 6))
        self.weekly_cal = CalendarPicker(ctrl, mondays_only=True)
        self.weekly_cal.pack(side="left", padx=(0, 8))
        ctk.CTkButton(ctrl, text="Load Week", width=100, command=self._load_weekly).pack(side="left")

        self.weekly_table_frame = ctk.CTkFrame(self.tab_weekly, fg_color="transparent")
        self.weekly_table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _build_adjusted_tab(self):
        ctrl = ctk.CTkFrame(self.tab_adjusted, fg_color="transparent")
        ctrl.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(ctrl, text="Week start:").pack(side="left", padx=(0, 6))
        self.adj_cal = CalendarPicker(ctrl, mondays_only=True)
        self.adj_cal.pack(side="left", padx=(0, 12))

        ctk.CTkLabel(ctrl, text="Target hours:").pack(side="left", padx=(0, 6))
        self.target_entry = ctk.CTkEntry(ctrl, width=60, justify="center")
        self.target_entry.pack(side="left", padx=(0, 12))
        self.target_entry.insert(0, "37.5")

        ctk.CTkButton(ctrl, text="Calculate", width=100, command=self._load_adjusted).pack(side="left")

        self.adj_output_frame = ctk.CTkFrame(self.tab_adjusted, fg_color="transparent")
        self.adj_output_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _build_calc_tab(self):
        frame = ctk.CTkFrame(self.tab_calc, fg_color="transparent")
        frame.pack(padx=20, pady=20, anchor="nw")

        ctk.CTkLabel(frame, text="Start Time (ISO 8601):").grid(row=0, column=0, sticky="w", pady=4)
        self.calc_start = ctk.CTkEntry(frame, width=360)
        self.calc_start.grid(row=0, column=1, padx=8, pady=4)

        ctk.CTkLabel(frame, text="End Time (ISO 8601):").grid(row=1, column=0, sticky="w", pady=4)
        self.calc_end = ctk.CTkEntry(frame, width=360)
        self.calc_end.grid(row=1, column=1, padx=8, pady=4)

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=2, pady=8)

        ctk.CTkButton(btn_frame, text="Calculate", width=100, command=self._calc_duration).pack(side="left", padx=4)
        ctk.CTkButton(
            btn_frame, text="Clear", width=80, command=self._calc_clear,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="left", padx=4)

        self.calc_result = ctk.CTkLabel(frame, text="", font=ctk.CTkFont(size=16, weight="bold"), anchor="w")
        self.calc_result.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 2))

        self.calc_detail = ctk.CTkLabel(frame, text="", font=ctk.CTkFont(size=12), text_color="#aaaaaa", anchor="w")
        self.calc_detail.grid(row=4, column=0, columnspan=2, sticky="w")

        self.calc_start.bind("<Return>", lambda e: self._calc_duration())
        self.calc_end.bind("<Return>", lambda e: self._calc_duration())

    # ── Data Loading ──────────────────────────────────────────────────────

    def _initial_load(self):
        vault = self.config["vault_path"]
        if not Path(vault).exists():
            messagebox.showerror(
                "Vault Not Found",
                f"Projects folder not found:\n{vault}\n\nUse Settings to set the correct path."
            )
        self._refresh_projects()

    def _refresh_projects(self):
        vault = self.config["vault_path"]
        self.projects = scan_projects(vault)
        self._detect_running_timer()
        self._populate_timer_buttons()
        self._schedule_auto_refresh()

    def _detect_running_timer(self):
        running = [p for p in self.projects if p.timer_running]

        if len(running) > 1:
            names = ", ".join(p.name for p in running)
            for p in running:
                try:
                    p.meta["timer_running"] = False
                    p.meta["session_start"] = None
                    safe_write_frontmatter(p.filepath, p.post)
                except Exception as e:
                    print(f"Error resetting {p.name}: {e}")
            self.running_project = None
            messagebox.showwarning(
                "Multiple Timers Detected",
                f"Found multiple running timers ({names}).\n"
                "All have been stopped to fix the corrupted state."
            )
            return

        if len(running) == 1:
            p = running[0]
            if p.session_start is not None:
                self.running_project = p
            else:
                p.meta["timer_running"] = False
                p.meta["session_start"] = None
                try:
                    safe_write_frontmatter(p.filepath, p.post)
                except Exception:
                    pass
                self.running_project = None
        else:
            self.running_project = None

        self._update_status_bar()

    def _schedule_auto_refresh(self):
        if self._refresh_after_id:
            self.after_cancel(self._refresh_after_id)
        self._refresh_after_id = self.after(60000, self._auto_refresh)

    def _auto_refresh(self):
        running_name = self.running_project.name if self.running_project else None
        vault = self.config["vault_path"]
        self.projects = scan_projects(vault)

        if running_name:
            match = [p for p in self.projects if p.name == running_name and p.timer_running]
            self.running_project = match[0] if match else None

        self._populate_timer_buttons()
        self._schedule_auto_refresh()

    # ── Timer Tab ─────────────────────────────────────────────────────────

    def _on_timer_canvas_configure(self, event=None):
        canvas_width = event.width if event else self._timer_canvas.winfo_width()
        content_width = self.timer_scroll.winfo_reqwidth()
        self._timer_canvas.itemconfig(self._timer_canvas_window, width=max(canvas_width, content_width))

    def _populate_timer_buttons(self):
        for w in self.timer_scroll.winfo_children():
            w.destroy()
        self.timer_buttons.clear()

        now = datetime.now()
        mon_start, sun_end = week_bounds_local(now)
        active = [p for p in self.projects if p.status == "Active"]

        groups: dict[str, list[ProjectNote]] = {}
        for label in [v for _, v in EFFORT_ORDER] + ["Unassigned"]:
            groups[label] = []

        for p in active:
            g = p.effort_group
            if g not in groups:
                groups["Unassigned"].append(p)
            else:
                groups[g].append(p)

        for group_label, projs in groups.items():
            if not projs:
                continue

            ctk.CTkLabel(
                self.timer_scroll, text=group_label,
                font=ctk.CTkFont(size=10, weight="bold"), anchor="w"
            ).pack(anchor="w", padx=2, pady=(4, 1))

            container = ctk.CTkFrame(self.timer_scroll, fg_color="transparent")
            container.pack(anchor="w", padx=2, pady=0)

            for p in sorted(projs, key=lambda x: x.name):
                week_mins = p.minutes_in_range(mon_start, sun_end)

                if self.running_project and p.name == self.running_project.name:
                    ss = self.running_project.session_start
                    if ss:
                        elapsed = (datetime.now(timezone.utc) - ss).total_seconds() / 60
                        week_mins += elapsed

                is_running = (self.running_project and p.name == self.running_project.name)

                time_str = format_hm(week_mins)
                btn = ctk.CTkButton(
                    container,
                    text=f"{p.name} ({time_str})",
                    height=22,
                    font=ctk.CTkFont(size=9),
                    fg_color=COLOR_RUNNING if is_running else COLOR_IDLE,
                    hover_color="#c0392b" if is_running else "#218838",
                    border_width=1 if is_running else 0,
                    border_color="#ff6b6b" if is_running else None,
                    corner_radius=4,
                    command=lambda proj=p: self._toggle_timer(proj),
                )
                btn.pack(side="left", padx=1, pady=1)
                self.timer_buttons[p.name] = btn

    def _toggle_timer(self, project: ProjectNote):
        try:
            if self.running_project and self.running_project.name == project.name:
                self.running_project.stop_timer()
                self.running_project = None
            else:
                if self.running_project:
                    self.running_project.stop_timer()
                    self.running_project = None
                project.load()
                project.start_timer()
                self.running_project = project
        except Exception as e:
            messagebox.showerror("File Write Error", f"Failed to write {project.filepath.name}:\n{e}")
            return

        self._populate_timer_buttons()
        self._update_status_bar()

    def _update_status_bar(self):
        if self._timer_after_id:
            self.after_cancel(self._timer_after_id)
            self._timer_after_id = None

        if self.running_project and self.running_project.session_start:
            elapsed = (datetime.now(timezone.utc) - self.running_project.session_start).total_seconds()
            self.status_label.configure(
                text=f"⏱ {self.running_project.name} — {format_hhmmss(elapsed)}",
                text_color="#ff6b6b"
            )
            self._timer_after_id = self.after(1000, self._tick)
        else:
            self.status_label.configure(text="No timer running", text_color="#aaaaaa")

    def _tick(self):
        self._update_status_bar()
        if self.running_project and self.running_project.name in self.timer_buttons:
            now = datetime.now()
            mon_start, sun_end = week_bounds_local(now)
            week_mins = self.running_project.minutes_in_range(mon_start, sun_end)
            ss = self.running_project.session_start
            if ss:
                week_mins += (datetime.now(timezone.utc) - ss).total_seconds() / 60
            self.timer_buttons[self.running_project.name].configure(
                text=f"{self.running_project.name} ({format_hm(week_mins)})"
            )

    def _status_bar_click(self, event=None):
        if self.running_project:
            self._toggle_timer(self.running_project)

    # ── Weekly View Tab ───────────────────────────────────────────────────

    def _load_weekly(self):
        monday = self.weekly_cal.get_date()
        if monday is None:
            messagebox.showerror("Invalid Date", "Please select a valid Monday using the calendar.")
            return

        for w in self.weekly_table_frame.winfo_children():
            w.destroy()

        day_headers = []
        for i in range(5):
            d = monday + timedelta(days=i)
            day_headers.append(f"{DAYS_OF_WEEK[i]} {d.strftime('%d/%m')}")

        table = DarkTable(
            self.weekly_table_frame,
            headers=["Project"] + day_headers + ["Week Total"],
            col_widths=[200] + [110] * 5 + [110],
            col_anchors=["w"] + ["center"] * 6,
        )
        table.pack(fill="both", expand=True)

        rows = []
        for p in self.projects:
            days = p.minutes_per_day(monday)
            total = sum(days)
            if total > 0:
                rows.append((p.name, days, total))

        rows.sort(key=lambda r: r[2], reverse=True)

        daily_totals = [0.0] * 5
        grand = 0.0

        for name, days, total in rows:
            vals = [name]
            for i in range(5):
                vals.append(format_hm(days[i]))
                daily_totals[i] += days[i]
            vals.append(format_hm(total))
            grand += total
            table.add_row(vals)

        tot_vals = ["Daily Totals"]
        for i in range(5):
            tot_vals.append(format_hm(daily_totals[i]))
        tot_vals.append(format_hm(grand))
        table.add_row(tot_vals, is_totals=True)

    # ── Adjusted Bookings Tab ─────────────────────────────────────────────

    def _load_adjusted(self):
        monday = self.adj_cal.get_date()
        if monday is None:
            messagebox.showerror("Invalid Date", "Please select a valid Monday using the calendar.")
            return

        try:
            target_hours = float(self.target_entry.get().strip())
            if target_hours <= 0:
                raise ValueError
            target_minutes = int(round(target_hours * 60))
        except ValueError:
            messagebox.showerror("Invalid Target", "Please enter a valid number of hours (e.g. 37.5).")
            return

        for w in self.adj_output_frame.winfo_children():
            w.destroy()

        result = calculate_adjusted_bookings(self.projects, monday, target_minutes)

        if not result["rows"]:
            ctk.CTkLabel(
                self.adj_output_frame, text="No project data for this week.",
                font=ctk.CTkFont(size=13)
            ).pack(pady=20)
            return

        actual_hmm = format_hmm(result["actual_grand"])
        target_hmm = format_hmm(target_minutes)
        summary = f"Actual: {actual_hmm} → Target: {target_hmm} (Multiplier: {result['multiplier']:.6f})"
        ctk.CTkLabel(
            self.adj_output_frame, text=summary,
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w"
        ).pack(fill="x", padx=4, pady=(4, 2))

        if result["excluded"]:
            exc_text = f"Excluded (<15 min): {', '.join(result['excluded'])}"
            ctk.CTkLabel(
                self.adj_output_frame, text=exc_text,
                font=ctk.CTkFont(size=11), text_color="#ff9800", anchor="w"
            ).pack(fill="x", padx=4, pady=(0, 4))

        day_headers = []
        for i in range(5):
            d = monday + timedelta(days=i)
            day_headers.append(f"{DAYS_OF_WEEK[i]} {d.strftime('%d/%m')}")

        table = DarkTable(
            self.adj_output_frame,
            headers=["Project"] + day_headers + ["Week Total"],
            col_widths=[200] + [110] * 5 + [110],
            col_anchors=["w"] + ["center"] * 6,
        )
        table.pack(fill="both", expand=True)

        daily_totals = [0] * 5
        grand = 0

        for row in result["rows"]:
            vals = [row["name"]]
            for i in range(5):
                v = row["final_days"][i]
                vals.append(format_hmm(v) if v > 0 else "-")
                daily_totals[i] += v
            vals.append(format_hmm(row["final_total"]))
            grand += row["final_total"]
            table.add_row(vals)

        tot_vals = ["Daily Totals"]
        for i in range(5):
            tot_vals.append(format_hmm(daily_totals[i]))
        tot_vals.append(format_hmm(grand))
        table.add_row(tot_vals, is_totals=True)

    # ── Duration Calculator Tab ───────────────────────────────────────────

    def _calc_duration(self):
        start_text = self.calc_start.get().strip()
        end_text = self.calc_end.get().strip()

        if not start_text or not end_text:
            self.calc_result.configure(text="Please enter both start and end times.", text_color="#ff6b6b")
            self.calc_detail.configure(text="")
            return

        start = parse_iso(start_text)
        end = parse_iso(end_text)

        if start is None:
            self.calc_result.configure(text="Invalid start time format.", text_color="#ff6b6b")
            self.calc_detail.configure(text="")
            return
        if end is None:
            self.calc_result.configure(text="Invalid end time format.", text_color="#ff6b6b")
            self.calc_detail.configure(text="")
            return
        if end <= start:
            self.calc_result.configure(text="End time must be after start time.", text_color="#ff6b6b")
            self.calc_detail.configure(text="")
            return

        diff = (end - start).total_seconds()
        exact_minutes = diff / 60
        ceil_minutes = math.ceil(exact_minutes)

        self.calc_result.configure(
            text=f"{ceil_minutes} minutes — {format_hm(ceil_minutes)}",
            text_color="white"
        )
        self.calc_detail.configure(text=f"Exact: {exact_minutes:.4f} minutes")

    def _calc_clear(self):
        self.calc_start.delete(0, "end")
        self.calc_end.delete(0, "end")
        self.calc_result.configure(text="")
        self.calc_detail.configure(text="")

    # ── Settings ──────────────────────────────────────────────────────────

    def _open_settings(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Settings")
        dialog.geometry("500x180")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="Vault Path:").pack(padx=16, pady=(16, 4), anchor="w")

        path_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        path_frame.pack(fill="x", padx=16, pady=(0, 12))

        path_entry = ctk.CTkEntry(path_frame, width=380)
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        path_entry.insert(0, self.config["vault_path"])

        def browse():
            d = filedialog.askdirectory(initialdir=self.config["vault_path"])
            if d:
                path_entry.delete(0, "end")
                path_entry.insert(0, d)

        ctk.CTkButton(path_frame, text="Browse", width=80, command=browse).pack(side="right")

        def save():
            new_path = path_entry.get().strip()
            if new_path and Path(new_path).exists():
                self.config["vault_path"] = new_path
                save_config(self.config)
                self._refresh_projects()
                dialog.destroy()
            else:
                messagebox.showerror("Invalid Path", "The specified path does not exist.", parent=dialog)

        ctk.CTkButton(dialog, text="Save", width=100, command=save).pack(pady=8)

    # ── Window Lifecycle ──────────────────────────────────────────────────

    def _on_close(self):
        try:
            geo = self.geometry()
            parts = geo.replace("+", "x").split("x")
            if len(parts) >= 4:
                self.config["window_width"] = int(parts[0])
                self.config["window_height"] = int(parts[1])
                self.config["window_x"] = int(parts[2])
                self.config["window_y"] = int(parts[3])
        except Exception:
            pass
        save_config(self.config)
        self.destroy()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
