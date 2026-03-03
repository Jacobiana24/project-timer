#!/usr/bin/env python3
"""
Project Time Tracker — Standalone desktop app for Obsidian vault time tracking.

Reads and writes YAML frontmatter in Obsidian markdown notes (.md) to provide
a timer UI, weekly views, adjusted bookings, and a duration calculator.

Compatible with existing Obsidian DataviewJS queries — only modifies fields
that the JS already reads.
"""

import json
import math
import os
import sys
import textwrap
import tkinter as tk
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import frontmatter
from dateutil import parser as dtparser

# ─── Constants ────────────────────────────────────────────────────────────────

CONFIG_FILE = "timer_config.json"
DEFAULT_CONFIG = {
    "vault_path": r"C:\Users\jacob.hand\Documents\My Brain\Work\Projects",
    "window_width": 1100,
    "window_height": 700,
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
TARGET_MINUTES = 2250
TARGET_DAILY = 450
DAYS_OF_WEEK = ["Mon", "Tue", "Wed", "Thu", "Fri"]

# ─── Utility Functions ────────────────────────────────────────────────────────


def config_path() -> Path:
    """Return path to config file next to the script/exe."""
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
    """Parse ISO 8601 string (including Z suffix) to aware UTC datetime."""
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
    """Format minutes as Xh Ym."""
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
    """Format minutes as H:MM."""
    m = int(round(minutes))
    h, rem = divmod(m, 60)
    return f"{h}:{rem:02d}"


def format_hhmmss(total_seconds: float) -> str:
    """Format seconds as H:MM:SS."""
    ts = max(0, int(total_seconds))
    h, rem = divmod(ts, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}"


def week_bounds_local(ref_date: datetime) -> tuple[datetime, datetime]:
    """Return (monday 00:00, sunday 23:59:59.999999) in local time for the
    week containing ref_date."""
    if ref_date.tzinfo:
        ref_date = ref_date.astimezone(tz=None).replace(tzinfo=None)
    monday = ref_date - timedelta(days=ref_date.weekday())
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    sunday = monday + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)
    return monday, sunday


def session_day_local(session: dict) -> datetime | None:
    """Return the local-time date a session should be attributed to.
    Uses end_time, falls back to start_time."""
    et = parse_iso(session.get("end_time"))
    if et:
        return et.astimezone(tz=None).replace(tzinfo=None)
    st = parse_iso(session.get("start_time"))
    if st:
        return st.astimezone(tz=None).replace(tzinfo=None)
    return None


def safe_write_frontmatter(filepath: Path, post: frontmatter.Post):
    """Write frontmatter safely via tmp file + rename."""
    tmp = filepath.with_suffix(".tmp")
    content = frontmatter.dumps(post)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(content)
    tmp.replace(filepath)


# ─── Project Data ─────────────────────────────────────────────────────────────


class ProjectNote:
    """Represents a single Obsidian project note."""

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
        """Sum session minutes within a local-time range."""
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
        """Return [Mon, Tue, Wed, Thu, Fri] minute totals for the week starting monday."""
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
    """Scan vault path for all Project class markdown notes."""
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


def calculate_adjusted_bookings(projects: list[ProjectNote], monday: datetime) -> dict:
    """
    Compute adjusted bookings for a given week.

    Returns dict with keys:
        rows: list of dicts {name, actual_days, actual_total, adj_days, adj_total, final_days, final_total}
        excluded: list of project names excluded (<15 min)
        actual_grand: total actual minutes
        multiplier: scaling multiplier
    """
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
        }

    actual_grand = sum(r["total"] for r in filtered)

    # ── Step 2: Scale — multiplier = 2250 / totalActualMinutes ──
    multiplier = TARGET_MINUTES / actual_grand if actual_grand > 0 else 1.0
    for r in filtered:
        r["adj_days"] = [v * multiplier for v in r["days"]]
        r["adj_total"] = r["total"] * multiplier

    # ── Step 3: Round to 15 mins ──
    # Round each adjusted day value: round(v / 15) * 15
    # Round adjusted weekly total independently
    # If sum of rounded days ≠ rounded total, add difference to highest-value day
    for r in filtered:
        r["final_days"] = [round(v / 15) * 15 for v in r["adj_days"]]
        r["final_total"] = round(r["adj_total"] / 15) * 15

        day_sum = sum(r["final_days"])
        diff = r["final_total"] - day_sum
        if diff != 0:
            # Find the day with the highest adjusted value (before rounding)
            max_idx = max(range(5), key=lambda i: r["adj_days"][i])
            r["final_days"][max_idx] += diff

    # ── Step 4: Enforce 2250 total ──
    # Sum all finalTotal values. If ≠ 2250, find project with highest total
    # and adjust its worked days evenly in 15-min increments.
    grand = sum(r["final_total"] for r in filtered)
    if grand != TARGET_MINUTES:
        diff = TARGET_MINUTES - grand
        # Find project with highest final_total
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
                # Apply remainder to single day (highest value worked day)
                if remainder != 0:
                    best_idx = max(worked_indices, key=lambda i: top["final_days"][i])
                    top["final_days"][best_idx] += remainder * 15
                top["final_total"] = sum(top["final_days"])

    # ── Step 5: Balance days to 450 mins each ──
    # For each Mon–Fri, if daily sum ≠ 450, add/subtract 15 from largest project.
    # Repeat up to 20 iterations.
    for _iteration in range(20):
        balanced = True
        for day_idx in range(5):
            day_sum = sum(r["final_days"][day_idx] for r in filtered)
            if day_sum != TARGET_DAILY:
                balanced = False
                diff = TARGET_DAILY - day_sum
                step = 15 if diff > 0 else -15
                # Find project with largest value on this day
                target_row = max(filtered, key=lambda r: r["final_days"][day_idx])
                target_row["final_days"][day_idx] += step
                target_row["final_total"] = sum(target_row["final_days"])
        if balanced:
            break

    # Build output rows sorted by final_total descending
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

        self.projects: list[ProjectNote] = []
        self.running_project: ProjectNote | None = None
        self._timer_after_id = None
        self._refresh_after_id = None

        self._build_ui()
        self._initial_load()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI Construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # Top bar with settings
        top = ctk.CTkFrame(self, height=36, fg_color="transparent")
        top.pack(fill="x", padx=8, pady=(4, 0))

        settings_btn = ctk.CTkButton(
            top, text="⚙ Settings", width=100, height=28,
            command=self._open_settings,
            fg_color="#555555", hover_color="#666666"
        )
        settings_btn.pack(side="right")

        # Tab view
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=8, pady=4)

        self.tab_timer = self.tabview.add("Timer")
        self.tab_weekly = self.tabview.add("Weekly View")
        self.tab_adjusted = self.tabview.add("Adjusted Bookings")
        self.tab_calc = self.tabview.add("Duration Calculator")

        self._build_timer_tab()
        self._build_weekly_tab()
        self._build_adjusted_tab()
        self._build_calc_tab()

        # Status bar
        self.status_frame = ctk.CTkFrame(self, height=32, corner_radius=0, fg_color="#1a1a2e")
        self.status_frame.pack(fill="x", side="bottom")
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame, text="No timer running",
            font=ctk.CTkFont(size=13), anchor="w"
        )
        self.status_label.pack(fill="x", padx=12, pady=4)
        self.status_frame.bind("<Button-1>", self._status_bar_click)
        self.status_label.bind("<Button-1>", self._status_bar_click)

    def _build_timer_tab(self):
        # Refresh button
        ctrl = ctk.CTkFrame(self.tab_timer, fg_color="transparent")
        ctrl.pack(fill="x", padx=4, pady=4)

        ctk.CTkButton(
            ctrl, text="🔄 Refresh", width=100, height=28,
            command=self._refresh_projects,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="left")

        # Scrollable area
        self.timer_scroll = ctk.CTkScrollableFrame(self.tab_timer)
        self.timer_scroll.pack(fill="both", expand=True, padx=4, pady=4)

        self.timer_buttons: dict[str, ctk.CTkButton] = {}

    def _build_weekly_tab(self):
        ctrl = ctk.CTkFrame(self.tab_weekly, fg_color="transparent")
        ctrl.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(ctrl, text="Week start (DD/MM/YYYY):").pack(side="left", padx=(0, 6))
        self.weekly_date_entry = ctk.CTkEntry(ctrl, width=140)
        self.weekly_date_entry.pack(side="left", padx=(0, 6))
        # Pre-fill with current week's Monday
        today = datetime.now()
        mon = today - timedelta(days=today.weekday())
        self.weekly_date_entry.insert(0, mon.strftime("%d/%m/%Y"))

        ctk.CTkButton(ctrl, text="Load Week", width=100, command=self._load_weekly).pack(side="left")

        # Treeview container
        tree_frame = ctk.CTkFrame(self.tab_weekly, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.weekly_tree = None
        self.weekly_tree_frame = tree_frame

    def _build_adjusted_tab(self):
        ctrl = ctk.CTkFrame(self.tab_adjusted, fg_color="transparent")
        ctrl.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(ctrl, text="Week start (DD/MM/YYYY):").pack(side="left", padx=(0, 6))
        self.adj_date_entry = ctk.CTkEntry(ctrl, width=140)
        self.adj_date_entry.pack(side="left", padx=(0, 6))
        today = datetime.now()
        mon = today - timedelta(days=today.weekday())
        self.adj_date_entry.insert(0, mon.strftime("%d/%m/%Y"))

        ctk.CTkButton(ctrl, text="Calculate", width=100, command=self._load_adjusted).pack(side="left")

        # Output area
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

        # Bind Enter key
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
        """On startup/refresh, find any running timers. Handle corruption."""
        running = [p for p in self.projects if p.timer_running]

        if len(running) > 1:
            # Corrupted state — stop all
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
                # timer_running=true but no session_start — reset silently
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
        """Re-scan vault without disrupting running timer."""
        running_name = self.running_project.name if self.running_project else None
        vault = self.config["vault_path"]
        self.projects = scan_projects(vault)

        # Re-link running project
        if running_name:
            match = [p for p in self.projects if p.name == running_name and p.timer_running]
            self.running_project = match[0] if match else None

        self._populate_timer_buttons()
        self._schedule_auto_refresh()

    # ── Timer Tab ─────────────────────────────────────────────────────────

    def _populate_timer_buttons(self):
        # Clear old widgets
        for w in self.timer_scroll.winfo_children():
            w.destroy()
        self.timer_buttons.clear()

        # Get current week bounds
        now = datetime.now()
        mon_start, sun_end = week_bounds_local(now)

        # Filter active projects for timer tab
        active = [p for p in self.projects if p.status == "Active"]

        # Group by effort
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

            # Group heading
            ctk.CTkLabel(
                self.timer_scroll, text=group_label,
                font=ctk.CTkFont(size=14, weight="bold"),
                anchor="w"
            ).pack(fill="x", padx=4, pady=(10, 4))

            # Button container for wrapping
            container = ctk.CTkFrame(self.timer_scroll, fg_color="transparent")
            container.pack(fill="x", padx=4, pady=2)

            for p in sorted(projs, key=lambda x: x.name):
                week_mins = p.minutes_in_range(mon_start, sun_end)

                # Add live elapsed if this project is running
                if self.running_project and p.name == self.running_project.name:
                    ss = self.running_project.session_start
                    if ss:
                        elapsed = (datetime.now(timezone.utc) - ss).total_seconds() / 60
                        week_mins += elapsed

                is_running = (self.running_project and p.name == self.running_project.name)

                btn_text = f"{p.name}\n{format_hm(week_mins)} this week"
                btn = ctk.CTkButton(
                    container,
                    text=btn_text,
                    width=170,
                    height=52,
                    font=ctk.CTkFont(size=12),
                    fg_color=COLOR_RUNNING if is_running else COLOR_IDLE,
                    hover_color="#c0392b" if is_running else "#218838",
                    border_width=2 if is_running else 0,
                    border_color="#ff6b6b" if is_running else None,
                    command=lambda proj=p: self._toggle_timer(proj),
                )
                btn.pack(side="left", padx=3, pady=3)
                self.timer_buttons[p.name] = btn

    def _toggle_timer(self, project: ProjectNote):
        """Start or stop a timer for the given project."""
        try:
            if self.running_project and self.running_project.name == project.name:
                # Stop current
                self.running_project.stop_timer()
                self.running_project = None
            else:
                # Stop existing if any
                if self.running_project:
                    self.running_project.stop_timer()
                    self.running_project = None

                # Reload the project fresh before starting
                project.load()
                project.start_timer()
                self.running_project = project

        except Exception as e:
            messagebox.showerror("File Write Error", f"Failed to write {project.filepath.name}:\n{e}")
            return

        self._populate_timer_buttons()
        self._update_status_bar()

    def _update_status_bar(self):
        """Update the status bar text. Schedules itself every second if running."""
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
        """Called every second to update status bar and button text."""
        self._update_status_bar()
        # Update the running button's week time
        if self.running_project and self.running_project.name in self.timer_buttons:
            now = datetime.now()
            mon_start, sun_end = week_bounds_local(now)
            week_mins = self.running_project.minutes_in_range(mon_start, sun_end)
            ss = self.running_project.session_start
            if ss:
                elapsed = (datetime.now(timezone.utc) - ss).total_seconds() / 60
                week_mins += elapsed
            btn = self.timer_buttons[self.running_project.name]
            btn.configure(text=f"{self.running_project.name}\n{format_hm(week_mins)} this week")

    def _status_bar_click(self, event=None):
        """Clicking the status bar while a timer is running stops it."""
        if self.running_project:
            self._toggle_timer(self.running_project)

    # ── Weekly View Tab ───────────────────────────────────────────────────

    def _parse_week_entry(self, entry_widget) -> datetime | None:
        """Parse DD/MM/YYYY from an entry widget, validate it's a Monday."""
        text = entry_widget.get().strip()
        try:
            dt = datetime.strptime(text, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter a date in DD/MM/YYYY format.")
            return None
        if dt.weekday() != 0:
            messagebox.showerror("Not a Monday", "The date must be a Monday.")
            return None
        return dt

    def _load_weekly(self):
        monday = self._parse_week_entry(self.weekly_date_entry)
        if monday is None:
            return

        # Clear previous
        for w in self.weekly_tree_frame.winfo_children():
            w.destroy()

        mon_start = monday.replace(hour=0, minute=0, second=0, microsecond=0)
        sun_end = mon_start + timedelta(days=6, hours=23, minutes=59, seconds=59)

        # Columns
        day_headers = []
        for i in range(5):
            d = monday + timedelta(days=i)
            day_headers.append(f"{DAYS_OF_WEEK[i]} {d.strftime('%d/%m')}")

        cols = ("project",) + tuple(f"d{i}" for i in range(5)) + ("total",)
        headings = ["Project"] + day_headers + ["Week Total"]

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Weekly.Treeview", background="#2b2b2b", foreground="white",
                        fieldbackground="#2b2b2b", rowheight=26, font=("Segoe UI", 10))
        style.configure("Weekly.Treeview.Heading", background="#3b3b3b", foreground="white",
                        font=("Segoe UI", 10, "bold"))

        tree = ttk.Treeview(self.weekly_tree_frame, columns=cols, show="headings",
                            style="Weekly.Treeview")

        tree.column("project", width=200, anchor="w")
        tree.heading("project", text="Project")
        for i in range(5):
            tree.column(f"d{i}", width=100, anchor="center")
            tree.heading(f"d{i}", text=day_headers[i])
        tree.column("total", width=100, anchor="center")
        tree.heading("total", text="Week Total")

        # Gather data
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
            tree.insert("", "end", values=vals)

        # Totals row
        tot_vals = ["Daily Totals"]
        for i in range(5):
            tot_vals.append(format_hm(daily_totals[i]))
        tot_vals.append(format_hm(grand))
        tree.insert("", "end", values=tot_vals, tags=("totals",))
        tree.tag_configure("totals", font=("Segoe UI", 10, "bold"))

        scrollbar = ttk.Scrollbar(self.weekly_tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(fill="both", expand=True, side="left")
        scrollbar.pack(fill="y", side="right")

    # ── Adjusted Bookings Tab ─────────────────────────────────────────────

    def _load_adjusted(self):
        monday = self._parse_week_entry(self.adj_date_entry)
        if monday is None:
            return

        # Clear previous
        for w in self.adj_output_frame.winfo_children():
            w.destroy()

        result = calculate_adjusted_bookings(self.projects, monday)

        if not result["rows"]:
            ctk.CTkLabel(
                self.adj_output_frame, text="No project data for this week.",
                font=ctk.CTkFont(size=13)
            ).pack(pady=20)
            return

        # Summary line
        actual_hmm = format_hmm(result["actual_grand"])
        summary = f"Actual: {actual_hmm} → Target: 37:30 (Multiplier: {result['multiplier']:.6f})"
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

        # Build table
        day_headers = []
        for i in range(5):
            d = monday + timedelta(days=i)
            day_headers.append(f"{DAYS_OF_WEEK[i]} {d.strftime('%d/%m')}")

        cols = ("project",) + tuple(f"d{i}" for i in range(5)) + ("total",)

        style = ttk.Style()
        style.configure("Adj.Treeview", background="#2b2b2b", foreground="white",
                        fieldbackground="#2b2b2b", rowheight=26, font=("Segoe UI", 10))
        style.configure("Adj.Treeview.Heading", background="#3b3b3b", foreground="white",
                        font=("Segoe UI", 10, "bold"))

        tree = ttk.Treeview(self.adj_output_frame, columns=cols, show="headings",
                            style="Adj.Treeview")

        tree.column("project", width=200, anchor="w")
        tree.heading("project", text="Project")
        for i in range(5):
            tree.column(f"d{i}", width=100, anchor="center")
            tree.heading(f"d{i}", text=day_headers[i])
        tree.column("total", width=100, anchor="center")
        tree.heading("total", text="Week Total")

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
            tree.insert("", "end", values=vals)

        # Totals row
        tot_vals = ["Daily Totals"]
        for i in range(5):
            tot_vals.append(format_hmm(daily_totals[i]))
        tot_vals.append(format_hmm(grand))
        tree.insert("", "end", values=tot_vals, tags=("totals",))
        tree.tag_configure("totals", font=("Segoe UI", 10, "bold"))

        scrollbar = ttk.Scrollbar(self.adj_output_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(fill="both", expand=True, side="left")
        scrollbar.pack(fill="y", side="right")

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
        # Save window geometry
        try:
            geo = self.geometry()
            # format: WxH+X+Y
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
