import os
import sys
import threading
import json
import ctypes
import calendar
from datetime import date, datetime, timedelta
import pandas as pd
import customtkinter as ctk
from pathlib import Path

APP_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
SYSTEM_DIR = APP_DIR / "System Files"
UPLOAD_DIR = APP_DIR / "Upload files"
SYSTEM_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Locate embedded icon inside PyInstaller _MEIPASS temp runtime
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    BUNDLE_DIR = Path(__file__).resolve().parent

ICO_FILE = BUNDLE_DIR / "icon.ico"

# 1. Set explicit Windows AppUserModelID so Taskbar uses the custom icon
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("mindyourlogic.bytebrew.automation.app")
except Exception:
    pass

# 2. Fix DPI Scaling on Windows for sharp UI rendering
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    except Exception:
        pass

# Import your existing ByteBrew downloader processing engine
try:
    import bytebrew_downloader as downloader
except ImportError:
    downloader = None

# Set UI Theme & Palette
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

SHEET_ID = "1GVaA0ajVKbrPbqdlY5vbXa6ONlXEjGnotS3-3pYI5Ww"
SHEET_NAME = "Dp1"

CREDENTIALS_FILE = SYSTEM_DIR / "credentials.json"


class DatePickerPopup(ctk.CTkToplevel):
    """Modal popup calendar overlay that avoids main window layout shifts."""
    def __init__(self, master, current_date, on_select_callback):
        super().__init__(master)
        self.master_app = master
        self.current_date = current_date
        self.view_year = current_date.year
        self.view_month = current_date.month
        self.on_select_callback = on_select_callback

        self.title("Select Date")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        if ICO_FILE.exists():
            try:
                self.iconbitmap(str(ICO_FILE))
            except Exception:
                pass

        self._build_ui()
        self._position_popup()

    def _position_popup(self):
        self.update_idletasks()
        parent_x = self.master_app.winfo_rootx()
        parent_y = self.master_app.winfo_rooty()
        parent_w = self.master_app.winfo_width()
        parent_h = self.master_app.winfo_height()

        w = 340
        h = 320
        x = parent_x + max(0, (parent_w - w) // 2)
        y = parent_y + max(0, (parent_h - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(pady=12, padx=14, fill="both", expand=True)

        # Month Switcher Header
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 8))

        ctk.CTkButton(
            header, text="◄", width=32, height=28, fg_color="#F1F5F9", text_color="#0F172A",
            hover_color="#E2E8F0", font=("Arial", 11, "bold"), command=self._prev_month
        ).pack(side="left")

        self.lbl_month = ctk.CTkLabel(
            header, text=f"{calendar.month_name[self.view_month]} {self.view_year}",
            font=("Arial", 13, "bold"), text_color="#0F172A"
        )
        self.lbl_month.pack(side="left", expand=True, padx=10)

        ctk.CTkButton(
            header, text="►", width=32, height=28, fg_color="#F1F5F9", text_color="#0F172A",
            hover_color="#E2E8F0", font=("Arial", 11, "bold"), command=self._next_month
        ).pack(side="right")

        # Calendar Day Grid
        self.grid_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.grid_frame.pack(pady=4)

        self._render_days()

        # Footer Actions
        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.pack(fill="x", pady=(10, 0))

        yest = date.today() - timedelta(days=1)
        ctk.CTkButton(
            footer, text="↺ Reset to Yesterday", fg_color="#F8FAFC", hover_color="#EDF2F7",
            text_color="#2A789B", font=("Arial", 11, "bold"), border_color="#CBD5E1", border_width=1,
            height=28, command=lambda: self._select_day(yest.year, yest.month, yest.day)
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            footer, text="✕ Close", fg_color="#F8FAFC", hover_color="#FEE2E2",
            text_color="#DC2626", font=("Arial", 11, "bold"), border_color="#CBD5E1", border_width=1,
            width=65, height=28, command=self.destroy
        ).pack(side="right")

    def _render_days(self):
        for widget in self.grid_frame.winfo_children():
            widget.destroy()

        for col, d in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
            ctk.CTkLabel(
                self.grid_frame, text=d, font=("Arial", 10, "bold"), text_color="#64748B", width=36
            ).grid(row=0, column=col, pady=(0, 2))

        cal_matrix = calendar.monthcalendar(self.view_year, self.view_month)
        for r_idx, week in enumerate(cal_matrix, start=1):
            for c_idx, day_num in enumerate(week):
                if day_num == 0:
                    continue

                is_selected = (
                    self.view_year == self.current_date.year and
                    self.view_month == self.current_date.month and
                    day_num == self.current_date.day
                )

                btn = ctk.CTkButton(
                    self.grid_frame,
                    text=str(day_num),
                    width=36,
                    height=28,
                    corner_radius=6,
                    font=("Arial", 11, "bold" if is_selected else "normal"),
                    fg_color="#2A789B" if is_selected else "transparent",
                    text_color="#FFFFFF" if is_selected else "#0F172A",
                    hover_color="#1F5A7B" if is_selected else "#EDF2F7",
                    command=lambda d=day_num: self._select_day(self.view_year, self.view_month, d)
                )
                btn.grid(row=r_idx, column=c_idx, padx=1, pady=1)

    def _prev_month(self):
        if self.view_month == 1:
            self.view_month = 12
            self.view_year -= 1
        else:
            self.view_month -= 1
        self.lbl_month.configure(text=f"{calendar.month_name[self.view_month]} {self.view_year}")
        self._render_days()

    def _next_month(self):
        if self.view_month == 12:
            self.view_month = 1
            self.view_year += 1
        else:
            self.view_month += 1
        self.lbl_month.configure(text=f"{calendar.month_name[self.view_month]} {self.view_year}")
        self._render_days()

    def _select_day(self, year, month, day):
        self.current_date = date(year, month, day)
        self.on_select_callback(self.current_date)
        self.destroy()


class CredentialsDialog(ctk.CTkToplevel):
    """Modal popup for configuring ByteBrew Login credentials styled identically to DatePickerPopup."""
    def __init__(self, master, on_save_callback=None):
        super().__init__(master)
        self.master_app = master
        self.on_save_callback = on_save_callback

        self.title("ByteBrew Credentials")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        if ICO_FILE.exists():
            try:
                self.iconbitmap(str(ICO_FILE))
            except Exception:
                pass

        self._build_ui()
        self._position_popup()

    def _position_popup(self):
        self.update_idletasks()
        parent_x = self.master_app.winfo_rootx()
        parent_y = self.master_app.winfo_rooty()
        parent_w = self.master_app.winfo_width()
        parent_h = self.master_app.winfo_height()

        w = 340
        h = 280
        x = parent_x + max(0, (parent_w - w) // 2)
        y = parent_y + max(0, (parent_h - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def load_saved(self):
        if CREDENTIALS_FILE.exists():
            try:
                with open(CREDENTIALS_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _build_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(pady=16, padx=16, fill="both", expand=True)

        ctk.CTkLabel(
            container, text="ByteBrew Login Setup", font=("Arial", 15, "bold"), text_color="#0F172A"
        ).pack(pady=(0, 14))

        creds = self.load_saved()

        self.email_entry = ctk.CTkEntry(container, placeholder_text="ByteBrew Email", width=300, height=36, font=("Arial", 12))
        self.email_entry.pack(pady=6)
        if creds.get("email"):
            self.email_entry.insert(0, creds["email"])

        self.pass_entry = ctk.CTkEntry(container, placeholder_text="ByteBrew Password", show="•", width=300, height=36, font=("Arial", 12))
        self.pass_entry.pack(pady=6)
        if creds.get("password"):
            self.pass_entry.insert(0, creds["password"])

        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.pack(fill="x", pady=(14, 0))

        ctk.CTkButton(
            footer, text="Save Credentials", fg_color="#10B981", hover_color="#059669",
            text_color="#FFFFFF", font=("Arial", 12, "bold"),
            height=34, command=self.save_and_close
        ).pack(fill="x")

    def save_and_close(self):
        data = {
            "email": self.email_entry.get().strip(),
            "password": self.pass_entry.get().strip()
        }
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump(data, f)
        self.destroy()
        if self.on_save_callback:
            self.on_save_callback()


class TaskCard(ctk.CTkFrame):
    """Custom UI component for each game task row matching the card design."""
    def __init__(self, master, index, task_data, start_single_callback):
        super().__init__(master, fg_color="#FFFFFF", border_color="#E2E8F0", border_width=1, corner_radius=12)
        self.task_data = task_data
        self.index = index
        self.start_single_callback = start_single_callback

        self.grid_columnconfigure(0, weight=1)

        task_title = str(task_data.get("Game Name") or task_data.get("Task Name") or "Unnamed Task").strip()

        self.title_label = ctk.CTkLabel(
            self, text=task_title, font=("Arial", 13, "bold"), text_color="#0F172A", anchor="w"
        )
        self.title_label.grid(row=0, column=0, sticky="w", padx=(16, 10), pady=16)

        # Single Action Button
        self.action_btn = ctk.CTkButton(
            self, text="Run", width=85, height=32, corner_radius=16,
            fg_color="#2A789B", hover_color="#1F5A7B", text_color="#FFFFFF", font=("Arial", 12, "bold"),
            command=lambda: self.start_single_callback(self)
        )
        self.action_btn.grid(row=0, column=1, padx=(10, 16), sticky="e")

    def update_status(self, status):
        """Updates action button state cleanly."""
        if status == "Running":
            self.action_btn.configure(state="disabled", text="Running...", fg_color="#94A3B8")
        elif status == "Completed":
            self.action_btn.configure(state="normal", text="Completed", fg_color="#10B981", hover_color="#059669")
        elif status == "Failed":
            self.action_btn.configure(state="normal", text="Retry", fg_color="#EF4444", hover_color="#DC2626")
        else:
            self.action_btn.configure(state="normal", text="Run", fg_color="#2A789B", hover_color="#1F5A7B")


class GroupCard(ctk.CTkFrame):
    """Main section container card with header toggle and nested task cards."""
    def __init__(self, master, title_text, run_all_callback):
        super().__init__(master, fg_color="#F8FAFC", border_color="#D9E2EC", border_width=1, corner_radius=14)
        self.is_expanded = True
        self.run_all_callback = run_all_callback

        # Section Header
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=16, pady=12)

        # Toggle Button + Section Label
        self.title_btn = ctk.CTkButton(
            self.header_frame, text=f"▼  {title_text}", font=("Arial", 16, "bold"),
            text_color="#0F172A", fg_color="transparent", hover_color="#EDF2F7",
            anchor="w", command=self.toggle_collapse
        )
        self.title_btn.pack(side="left")

        # Run All Button
        self.btn_run_all = ctk.CTkButton(
            self.header_frame, text="▶  Run All", font=("Arial", 12, "bold"),
            fg_color="#287A88", hover_color="#1E616C", text_color="#FFFFFF",
            width=90, height=32, corner_radius=14,
            command=self.run_all_callback
        )
        self.btn_run_all.pack(side="right")

        # Tasks Container inside section card
        self.tasks_container = ctk.CTkFrame(self, fg_color="transparent")
        self.tasks_container.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def toggle_collapse(self):
        """Collapses or expands the inner task rows."""
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.title_btn.configure(text=self.title_btn.cget("text").replace("►", "▼"))
            self.tasks_container.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        else:
            self.title_btn.configure(text=self.title_btn.cget("text").replace("▼", "►"))
            self.tasks_container.pack_forget()


class DP1AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ByteBrew Automation MYL")
        self.geometry("820x640")
        self.minsize(720, 520)
        self.configure(fg_color="#ECEFF4")

        if ICO_FILE.exists():
            try:
                self.iconbitmap(str(ICO_FILE))
            except Exception:
                pass

        self.tasks_data = []
        self.task_cards = []
        self.is_running_all = False
        self.selected_date = date.today() - timedelta(days=1)

        self._build_ui()
        self.load_tasks_from_sheet()

    def _build_ui(self):
        # 1. Top Centered Control Bar
        header_bar = ctk.CTkFrame(self, fg_color="transparent")
        header_bar.pack(fill="x", padx=20, pady=(18, 10))

        center_buttons_frame = ctk.CTkFrame(header_bar, fg_color="transparent")
        center_buttons_frame.pack(anchor="center")

        btn_style = {
            "font": ("Arial", 12),
            "fg_color": "#F1F5F9",
            "hover_color": "#E2E8F0",
            "text_color": "#1E293B",
            "border_color": "#CBD5E1",
            "border_width": 1,
            "height": 35,
            "corner_radius": 10
        }

        self.btn_date = ctk.CTkButton(
            center_buttons_frame,
            text=f"📅  {self.selected_date.isoformat()}",
            command=self.open_date_picker,
            **btn_style
        )
        self.btn_date.pack(side="left", padx=6)

        self.btn_refresh = ctk.CTkButton(
            center_buttons_frame, text="🔄 Refresh Pull", command=self.load_tasks_from_sheet, **btn_style
        )
        self.btn_refresh.pack(side="left", padx=6)

        self.btn_open_folder = ctk.CTkButton(
            center_buttons_frame, text="📁 Open Output Folder", command=self.open_output_folder, **btn_style
        )
        self.btn_open_folder.pack(side="left", padx=6)

        self.btn_creds = ctk.CTkButton(
            center_buttons_frame, text="🔑 Login Info", command=lambda: CredentialsDialog(self), **btn_style
        )
        self.btn_creds.pack(side="left", padx=6)

        # 2. Main Scrollable View Area
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # 3. Main Section Card Wrapper
        self.group_card = GroupCard(self.scroll_frame, "Daily Game DP1", self.run_all_tasks)
        self.group_card.pack(fill="x", pady=5)

        # 4. Bottom Status Footer Bar
        self.footer = ctk.CTkFrame(self, height=32, fg_color="#E2E8F0")
        self.footer.pack(fill="x", side="bottom")

        self.lbl_summary = ctk.CTkLabel(
            self.footer, text="System Ready.", font=("Arial", 11), text_color="#475569"
        )
        self.lbl_summary.pack(side="left", padx=15, pady=4)

    def load_tasks_from_sheet(self):
        """Fetches live Google Sheet config and renders UI rows."""
        self.lbl_summary.configure(text="Fetching rules from Google Sheet...")
        
        # Clear existing card widgets
        for card in self.task_cards:
            card.destroy()
        self.task_cards.clear()

        def _fetch():
            url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"
            try:
                df = pd.read_csv(url)
                # Clean column headers and drop empty Unnamed columns
                df.columns = df.columns.str.strip()
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                
                # Filter active tasks cleanly
                active_df = df[df['Active'].astype(str).str.strip().str.lower().isin(['yes', 'y', '1', 'true', 'active'])].copy()
                active_df['Order'] = pd.to_numeric(active_df['Order'], errors='coerce').fillna(9999)
                active_df = active_df.sort_values('Order')
                
                self.tasks_data = active_df.to_dict('records')
                self.after(0, self._render_tasks)
            except Exception as e:
                self.after(0, lambda: self.lbl_summary.configure(text=f"Error loading sheet: {e}"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _render_tasks(self):
        if not self.tasks_data:
            self.lbl_summary.configure(text="No active tasks found in Google Sheet.")
            return

        for idx, task in enumerate(self.tasks_data, start=1):
            card = TaskCard(self.group_card.tasks_container, idx, task, self.run_single_task)
            card.pack(fill="x", pady=5)
            self.task_cards.append(card)

        self.lbl_summary.configure(text=f"Loaded {len(self.tasks_data)} active task(s).")

    def open_date_picker(self):
        """Opens a modal calendar popup directly over the application."""
        DatePickerPopup(self, self.selected_date, self._on_date_picked)

    def _on_date_picked(self, chosen_date):
        self.selected_date = chosen_date
        self.btn_date.configure(text=f"📅  {self.selected_date.isoformat()}")
        self.lbl_summary.configure(text=f"Target cohort date set to: {self.selected_date.isoformat()}")

    def get_selected_date(self):
        return self.selected_date.isoformat()

    def open_output_folder(self):
        """Opens selected date's upload folder or upload root in Windows Explorer."""
        target_date = self.get_selected_date()
        folder_path = UPLOAD_DIR / target_date
        
        if not folder_path.exists():
            folder_path = UPLOAD_DIR
            
        os.startfile(folder_path)

    def run_single_task(self, card):
        """Executes a single game task in a background thread."""
        def _worker():
            card.update_status("Running")
            self.lbl_summary.configure(text=f"Running: {card.task_data.get('Game Name')}...")
            
            res = self._execute_bytebrew_job(card.task_data, on_login_needed=lambda: self.run_single_task(card))
            if res == "LOGIN_PROMPT":
                card.update_status("Run")
                return

            status = "Completed" if res is True else "Failed"
            card.update_status(status)
            self.lbl_summary.configure(text=f"Finished {card.task_data.get('Game Name')}: {status}")

        threading.Thread(target=_worker, daemon=True).start()

    def run_all_tasks(self):
        """Runs all active tasks sequentially (one by one)."""
        if self.is_running_all:
            return

        def _worker_all():
            self.is_running_all = True
            self.group_card.btn_run_all.configure(state="disabled", text="Running...", fg_color="#94A3B8")

            for card in self.task_cards:
                card.update_status("Running")
                self.lbl_summary.configure(text=f"Processing queue: {card.task_data.get('Game Name')}...")
                
                res = self._execute_bytebrew_job(card.task_data, on_login_needed=lambda: self.run_all_tasks())
                if res == "LOGIN_PROMPT":
                    self.is_running_all = False
                    card.update_status("Run")
                    self.group_card.btn_run_all.configure(state="normal", text="▶  Run All", fg_color="#287A88")
                    return

                status = "Completed" if res is True else "Failed"
                card.update_status(status)

            self.is_running_all = False
            self.group_card.btn_run_all.configure(state="normal", text="▶  Run All", fg_color="#287A88")
            self.lbl_summary.configure(text="Queue completed!")

        threading.Thread(target=_worker_all, daemon=True).start()

    def _execute_bytebrew_job(self, task_dict, on_login_needed=None):
        """Integration bridge to your bytebrew_downloader script."""
        if not downloader:
            import time
            time.sleep(2)
            return True

        # Upfront Credential Check
        creds = {}
        if CREDENTIALS_FILE.exists():
            try:
                with open(CREDENTIALS_FILE, "r") as f:
                    creds = json.load(f)
            except Exception:
                pass

        email = (creds.get("email") or "").strip()
        password = (creds.get("password") or "").strip()

        if not email or not password:
            self.after(0, lambda: CredentialsDialog(self, on_save_callback=on_login_needed))
            self.after(0, lambda: self.lbl_summary.configure(text="Please enter your ByteBrew login credentials."))
            return "LOGIN_PROMPT"

        # Construct standard game dictionary expected by bytebrew_downloader
        game_config = {
            "order": task_dict.get("Order", 9999),
            "game_name": str(task_dict.get("Game Name") or task_dict.get("Game ID")),
            "game_id": str(task_dict.get("Game ID")),
            "page": str(task_dict.get("Page") or "funnelexplorer"),
            "saved_funnels": [
                f for f in [
                    str(task_dict.get("Saved Funnel 1", "")).strip(),
                    str(task_dict.get("Saved Funnel 2", "")).strip(),
                ] if f and f.lower() != "nan"
            ],
            "date_range": self.get_selected_date(),
            "build_version": "" if pd.isna(task_dict.get("Build Version")) else str(task_dict.get("Build Version")),
            "country": "" if pd.isna(task_dict.get("Country")) else str(task_dict.get("Country")),
            "output_name": str(task_dict.get("Output Name") or task_dict.get("Game Name")),
        }

        try:
            from playwright.sync_api import sync_playwright
            profile_dir = SYSTEM_DIR / "bytebrew_profile"
            profile_dir.mkdir(parents=True, exist_ok=True)

            with sync_playwright() as p:
                context = None
                for channel in ["chrome", "msedge", None]:
                    try:
                        kwargs = {
                            "user_data_dir": str(profile_dir),
                            "headless": True,
                            "accept_downloads": True,
                            "viewport": {"width": 1600, "height": 950},
                        }
                        if channel:
                            kwargs["channel"] = channel
                        context = p.chromium.launch_persistent_context(**kwargs)
                        break
                    except Exception:
                        continue

                if not context:
                    return False
                page = context.pages[0] if context.pages else context.new_page()
                page.set_default_timeout(15000)

                # Hook log callback to update footer in real-time step-by-step
                downloader.LOG_CALLBACK = lambda msg: self.after(0, lambda: self.lbl_summary.configure(text=msg))

                # Process single game sequential job
                downloader.process_game(page, game_config, email=email, password=password)
                context.close()
                return True
        except Exception as e:
            err_msg = str(e)
            # Route fatal crash messages to the new execution_log.txt file as well as the UI
            downloader.log(f"[Error executing {game_config['game_name']}]: {err_msg}")
            
            try:
                if 'page' in locals():
                    # Take a final crash screenshot right into the daily folder before aborting
                    downloader.screenshot(page, game_config["game_name"], "fatal_error")
            except Exception:
                pass
                
            if "LOGIN_" in err_msg:
                self.after(0, lambda: CredentialsDialog(self, on_save_callback=on_login_needed))
                self.after(0, lambda: self.lbl_summary.configure(text="Login failed. Please check credentials."))
                return "LOGIN_PROMPT"
            return False


if __name__ == "__main__":
    app = DP1AutomationApp()
    app.mainloop()