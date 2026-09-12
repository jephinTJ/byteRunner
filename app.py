import os
import sys
from datetime import date, timedelta
from pathlib import Path
import webview
import threading
import json

# --- SMART PATHING LOGIC ---
if getattr(sys, 'frozen', False):
    # Compiled Mode: PyInstaller's --onedir hides bundled assets (like UI) in the _internal folder mapped to sys._MEIPASS
    BUNDLE_DIR = Path(sys._MEIPASS)
    # The actual .exe location for generating local user folders
    ROOT_DIR = Path(sys.executable).parent 
else:
    # Dev Mode
    BUNDLE_DIR = Path(__file__).resolve().parent
    ROOT_DIR = BUNDLE_DIR

SYSTEM_DIR = ROOT_DIR / "System Files"
UPLOAD_DIR = ROOT_DIR / "Upload files"
SYSTEM_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

SHEET_ID = "1GVaA0ajVKbrPbqdlY5vbXa6ONlXEjGnotS3-3pYI5Ww"
SHEET_NAME_DP1 = "Dp1"
SHEET_NAME_ALL_GEO = "allGeo"
SHEET_NAME_DROP_SHEET = "DropSheet"
CREDENTIALS_FILE = SYSTEM_DIR / "credentials.json"
SETTINGS_FILE = SYSTEM_DIR / "settings.json"

class DesktopAPI:
    """JS Bridge: Connects Tailwind frontend directly to Playwright scraper engine."""
    def __init__(self):
        self._window = None  # Private attribute prevents PyWebView infinite recursion crash
        self.tasks_dp1 = []
        self.tasks_all_geo = []
        self.is_running_all_dp1 = False
        self.is_running_all_all_geo = False
        self.tasks_dropsheet = []
        self.is_running_all_dropsheet = False
        saved_folder = str(Path.home() / "Downloads")
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    saved_folder = json.load(f).get("dropsheet_folder", saved_folder)
            except Exception:
                pass
        self.dropsheet_folder = saved_folder
        self.abort_signal = False

    def emit_log(self, msg, queue_type=None, idx=None):
        """Pushes backend log events to the frontend UI terminal."""
        if self._window:
            safe_msg = json.dumps(msg)
            q_val = json.dumps(queue_type) if queue_type else "null"
            idx_val = idx if idx is not None else "null"
            try:
                self._window.evaluate_js(f"appendLog({safe_msg}, {q_val}, {idx_val})")
            except Exception:
                pass

    def kill_task(self):
        """Instantly aborts Playwright browser context to kill running tasks."""
        self.is_running_all_dp1 = False
        self.is_running_all_all_geo = False
        self.abort_signal = True
        
        # 1. Instant UI telemetry
        self.emit_log("[System] KILL SIGNAL RECEIVED. Force severing Playwright connection...", None)
        
        # 2. Brutally kill the active Playwright page from the inside to instantly snap the blocked thread
        if hasattr(self, 'active_page') and self.active_page:
            try:
                self.active_page.evaluate("window.stop();", timeout=100)
            except Exception:
                pass
            try:
                self.active_page.close()
            except Exception:
                pass
        if hasattr(self, 'active_context') and self.active_context:
            try:
                self.active_context.close()
            except Exception:
                pass

    def select_dropsheet_folder(self):
        if self._window:
            try:
                # Correct newer pywebview syntax
                dialog_type = webview.FileDialog.FOLDER
            except AttributeError:
                # Fallback for older versions
                dialog_type = webview.FOLDER_DIALOG
            
            result = self._window.create_file_dialog(dialog_type)
            if result and len(result) > 0:
                self.dropsheet_folder = result[0]
                data = {}
                if SETTINGS_FILE.exists():
                    try:
                        with open(SETTINGS_FILE, "r") as f:
                            data = json.load(f)
                    except Exception:
                        pass
                data["dropsheet_folder"] = self.dropsheet_folder
                with open(SETTINGS_FILE, "w") as f:
                    json.dump(data, f)
                self.emit_log(f"[System] Drop Sheet folder updated to: {self.dropsheet_folder}")
        return self.dropsheet_folder
                
    def get_initial_state(self):
        delete_csvs = False
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    delete_csvs = json.load(f).get("delete_csvs", False)
            except Exception:
                pass
        return {
            "selected_date": (date.today() - timedelta(days=1)).isoformat(),
            "status": "MindYourLogic Studios",
            "delete_csvs": delete_csvs
        }

    def get_dropsheet_settings(self):
        data = {}
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
            except Exception:
                pass
        return {
            "folder": data.get("dropsheet_folder", self.dropsheet_folder),
            "delete_csvs": data.get("delete_csvs", False),
            "formatting": data.get("dropsheet_formatting", {})
        }

    def save_dropsheet_settings(self, payload):
        data = {}
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
            except Exception:
                pass
        
        self.dropsheet_folder = payload.get("folder", self.dropsheet_folder)
        data["dropsheet_folder"] = self.dropsheet_folder
        data["delete_csvs"] = bool(payload.get("delete_csvs", False))
        data["dropsheet_formatting"] = payload.get("formatting", {})
        
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f)
        
        self.emit_log("[System] Drop Sheet configurations saved.")
        return True

    def select_folder_dialog_only(self):
        if self._window:
            try:
                dialog_type = webview.FileDialog.FOLDER
            except AttributeError:
                dialog_type = webview.FOLDER_DIALOG
            result = self._window.create_file_dialog(dialog_type)
            if result and len(result) > 0:
                return result[0]
        return None
    
    def set_delete_csvs(self, enabled):
        data = {}
        if SETTINGS_FILE.exists():
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
            except Exception:
                pass
        data["delete_csvs"] = bool(enabled)
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f)
        return True

    def open_output_folder(self, target_date):
        folder_path = UPLOAD_DIR / target_date
        if not folder_path.exists():
            folder_path = UPLOAD_DIR
        os.startfile(folder_path)

    def get_credentials(self):
        if CREDENTIALS_FILE.exists():
            try:
                with open(CREDENTIALS_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"email": "", "password": ""}

    def save_credentials(self, email, password):
        data = {"email": email, "password": password}
        with open(CREDENTIALS_FILE, "w") as f:
            json.dump(data, f)
        return True

    def load_tasks_from_sheet(self):
        import pandas as pd  # Lazy-load massive library to prevent UI lag on startup
        
        url_dp1 = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME_DP1}"
        url_all_geo = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME_ALL_GEO}"
        url_dropsheet = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME_DROP_SHEET}"
        meta_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=metaData"
        try:
            icon_map = {}
            try:
                meta_df = pd.read_csv(meta_url)
                meta_df.columns = meta_df.columns.str.strip()
                if 'Game Name' in meta_df.columns and 'Playstore Icon' in meta_df.columns:
                    meta_df = meta_df.dropna(subset=['Game Name', 'Playstore Icon'])
                    icon_map = dict(zip(meta_df['Game Name'].astype(str).str.strip(), meta_df['Playstore Icon'].astype(str).str.strip()))
            except Exception as e:
                print(f"[DesktopAPI] Warning: Could not load metaData sheet: {e}", flush=True)

            def parse_sheet_df(url, is_all_geo=False):
                try:
                    df = pd.read_csv(url)
                    df.columns = df.columns.str.strip()
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    active_df = df[df['Active'].astype(str).str.strip().str.lower().isin(['yes', 'y', '1', 'true', 'active'])].copy()
                    active_df['Order'] = pd.to_numeric(active_df['Order'], errors='coerce').fillna(9999)
                    active_df = active_df.sort_values('Order').fillna("")
                    tasks = active_df.to_dict('records')
                    for task in tasks:
                        g_name = str(task.get('Game Name', '')).strip()
                        task['Playstore Icon'] = icon_map.get(g_name, "")
                        task['is_all_geo'] = is_all_geo
                    return tasks
                except Exception as ex:
                    print(f"[DesktopAPI] Error parsing sheet from {url}: {ex}", flush=True)
                    return []

            self.tasks_dp1 = parse_sheet_df(url_dp1, is_all_geo=False)
            self.tasks_all_geo = parse_sheet_df(url_all_geo, is_all_geo=True)
            self.tasks_dropsheet = parse_sheet_df(url_dropsheet, is_all_geo=False)

            return {
                "dp1": self.tasks_dp1,
                "all_geo": self.tasks_all_geo,
                "dropsheet": self.tasks_dropsheet
            }
        except Exception as e:
            print(f"[DesktopAPI] Error loading sheets: {e}", flush=True)
            return {"dp1": [], "all_geo": [], "dropsheet": []}

    def _execute_dropsheet_job(self, task_dict, idx, delete_csvs=False):
        import requests
        self.abort_signal = False
        script_url = str(task_dict.get("Script URL", "")).strip()
        prefix = str(task_dict.get("Output Prefix", "")).strip()
        g_name = str(task_dict.get("Game Name", "Unnamed")).strip()
        
        if not script_url:
            self.emit_log(f"[Error] No script URL found for {g_name}", "dropsheet", idx)
            return False

        cache_dir = SYSTEM_DIR / "cached_modules"
        cache_dir.mkdir(parents=True, exist_ok=True)
        local_cache = cache_dir / f"drop_{idx}.py"
        code_text = None

        try:
            res = requests.get(script_url, headers={"Cache-Control": "no-cache"}, timeout=5)
            if res.status_code == 200:
                code_text = res.text
                with open(local_cache, "w", encoding="utf-8") as f:
                    f.write(code_text)
        except Exception as e:
            self.emit_log(f"[System] Network fetch failed: {e}. Checking cache...", "dropsheet", idx)

        if not code_text and local_cache.exists():
            with open(local_cache, "r", encoding="utf-8") as f:
                code_text = f.read()

        if not code_text:
            self.emit_log(f"[Error] Fatal: Cannot load script for {g_name} offline.", "dropsheet", idx)
            return False

        try:
            self.emit_log(f"[{g_name}] Processing in: {self.dropsheet_folder}", "dropsheet", idx)
            module_scope = {}
            exec(code_text, module_scope)
            
            # Fetch latest configurations dynamically
            settings = self.get_dropsheet_settings()
            use_delete = settings.get("delete_csvs", False)

            import time
            start_time = time.time()

            # Block the online script from opening the file immediately
            res = module_scope['run'](
                folder_path=self.dropsheet_folder,
                custom_filename=prefix,
                delete_csvs=use_delete,
                auto_open=False
            )
            
            # Locate target Excel file (from return value or auto-detect latest file)
            target_file = None
            if isinstance(res, str) and res.endswith('.xlsx') and os.path.exists(res):
                target_file = res
            else:
                folder = Path(self.dropsheet_folder)
                candidates = []
                if prefix:
                    candidates = [f for f in folder.glob(f"*{prefix}*.xlsx") if f.is_file()]
                if not candidates:
                    candidates = [f for f in folder.glob("*.xlsx") if f.is_file()]
                
                # Find files created or modified around this execution run
                recent = [f for f in candidates if f.stat().st_mtime >= (start_time - 5)]
                pool = recent if recent else candidates
                if pool:
                    target_file = str(max(pool, key=lambda p: p.stat().st_mtime))

            if target_file and os.path.exists(target_file):
                self._apply_excel_formatting(target_file, idx)
                if os.name == 'nt':
                    os.startfile(target_file)
            else:
                self.emit_log(f"[{g_name}] Notice: Output Excel file could not be detected.", "dropsheet", idx)

            self.emit_log(f"[{g_name}] Completed successfully.", "dropsheet", idx)
            return True
        except Exception as e:
            self.emit_log(f"[Error executing {g_name}]: {e}", "dropsheet", idx)
            return False
        
    def _apply_excel_formatting(self, file_path, idx):
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font
            from openpyxl.formatting.rule import CellIsRule
            from openpyxl.utils import get_column_letter
        except ImportError:
            self.emit_log("[System] openpyxl missing. Skipping conditional styling.", "dropsheet", idx)
            return

        try:
            settings = self.get_dropsheet_settings()
            fmt = settings.get("formatting", {})
            
            def get_val(key, default):
                try:
                    return round(float(fmt.get(key, default)), 1) / 100.0
                except (ValueError, TypeError):
                    return round(float(default), 1) / 100.0

            # Map the exact frontend logic IDs
            rules_config = {
                "Level Drop%": {
                    "flag": get_val("ld_flag", 3.0),
                    "danger": get_val("ld_danger", 6.0)
                },
                "Interruption Drop%": {
                    "flag": get_val("id_flag", 4.0),
                    "danger": get_val("id_danger", 8.0)
                },
                "Total Drop%": {
                    "flag": get_val("td_flag", 5.0),
                    "danger": get_val("td_danger", 9.9)
                }
            }

            wb = openpyxl.load_workbook(file_path)
            ws = wb.active

            # Preset styling arrays based on requirement
            fill_flag = PatternFill(start_color="FF5050", end_color="FF5050", fill_type="solid")
            font_flag = Font(color="000000", bold=False)
            
            fill_danger = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
            font_danger = Font(color="000000", bold=False)

            header_row = 1
            col_map = {}
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=header_row, column=col).value
                if cell_val in rules_config:
                    col_map[cell_val] = get_column_letter(col)

            for col_name, col_letter in col_map.items():
                rng = f"{col_letter}2:{col_letter}{ws.max_row}"
                cfg = rules_config[col_name]
                
                # Rule 1: Danger threshold (Dark Red, White text)
                rule_danger = CellIsRule(operator='greaterThanOrEqual', formula=[str(cfg["danger"])], stopIfTrue=True, fill=fill_danger, font=font_danger)
                # Rule 2: Flagged threshold (Light Red, Black text)
                rule_flag = CellIsRule(operator='greaterThanOrEqual', formula=[str(cfg["flag"])], stopIfTrue=True, fill=fill_flag, font=font_flag)

                ws.conditional_formatting.add(rng, rule_danger)
                ws.conditional_formatting.add(rng, rule_flag)

            wb.save(file_path)
            self.emit_log("[System] Dynamic custom cell formatting applied.", "dropsheet", idx)
        except Exception as e:
            self.emit_log(f"[Error] Excel formatting failed: {str(e)}", "dropsheet", idx)

    def _execute_job(self, task_dict, queue_type, idx, target_date):
        self.abort_signal = False
        
        try:
            import bytebrew_downloader as downloader  # Lazy-load Playwright engine only when task runs
        except ImportError:
            downloader = None

        if not downloader:
            import time
            time.sleep(1.5)
            return True

        creds = {}
        if CREDENTIALS_FILE.exists():
            try:
                with open(CREDENTIALS_FILE, "r") as f:
                    creds = json.load(f)
            except Exception:
                pass

        email = (creds.get("email") or "").strip()
        password = (creds.get("password") or "").strip()

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
            "date_range": target_date,
            "build_version": str(task_dict.get("Build Version", "")).strip() if task_dict.get("Build Version") else "",
            "country": str(task_dict.get("Country", "")).strip() if task_dict.get("Country") else "",
            "output_name": str(task_dict.get("Output Name") or task_dict.get("Game Name")),
            "is_all_geo": bool(task_dict.get("is_all_geo", False)),
        }

        try:
            def custom_log(msg):
                self.emit_log(msg, queue_type, idx)
                
            if downloader:
                downloader.LOG_CALLBACK = custom_log

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
                        self.active_context = context
                        break
                    except Exception:
                        continue

                if not context:
                    return False

                page = context.pages[0] if context.pages else context.new_page()
                self.active_page = page
                page.set_default_timeout(15000)

                downloader.process_game(page, game_config, email=email, password=password)
                
                self.active_page = None
                self.active_context = None
                context.close()
                return True
        except Exception as e:
            err_str = str(e)
            if self.abort_signal:
                downloader.log(f"Task for {game_config['game_name']} was manually stopped.")
                return "aborted"
            elif "LOGIN_REQUIRED" in err_str or "LOGIN_FAILED" in err_str:
                downloader.log(f"Login missing or invalid. Pausing queue to request credentials.")
                return "login_error"
            else:
                downloader.log(f"[Error executing {game_config['game_name']}]: {e}")
            return False

    def run_single_task(self, queue_type, idx, target_date, delete_csvs=False):
        if queue_type == "dp1": tasks = self.tasks_dp1
        elif queue_type == "all_geo": tasks = self.tasks_all_geo
        else: tasks = self.tasks_dropsheet
        
        if idx >= len(tasks): return

        def _worker():
            self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Running')")
            if queue_type == "dropsheet":
                res = self._execute_dropsheet_job(tasks[idx], idx, delete_csvs)
            else:
                res = self._execute_job(tasks[idx], queue_type, idx, target_date)
            
            if res == "login_error":
                self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Failed')")
                self._window.evaluate_js(f"triggerLoginRecovery('{queue_type}', {idx}, false)")
                return
                
            status = 'Completed' if res == True else ('Aborted' if res == 'aborted' else 'Failed')
            self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, '{status}')")

        threading.Thread(target=_worker, daemon=True).start()

    def run_all_tasks(self, queue_type, target_date, start_idx=0):
        if queue_type == "dp1": tasks = self.tasks_dp1
        elif queue_type == "all_geo": tasks = self.tasks_all_geo
        else: tasks = self.tasks_dropsheet
        if not tasks: return

        def _worker_all():
            if queue_type == "dp1": self.is_running_all_dp1 = True
            elif queue_type == "all_geo": self.is_running_all_all_geo = True
            else: self.is_running_all_dropsheet = True

            total = len(tasks)
            self._window.evaluate_js(f"setRunAllState('{queue_type}', true, {start_idx}, {total})")

            success_count = start_idx
            for idx, task in enumerate(tasks):
                if idx < start_idx:
                    continue
                    
                if queue_type == "dp1": is_running = self.is_running_all_dp1
                elif queue_type == "all_geo": is_running = self.is_running_all_all_geo
                else: is_running = self.is_running_all_dropsheet

                if not is_running: break
                
                self._window.evaluate_js(f"setRunAllState('{queue_type}', true, {idx+1}, {total})")
                self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Running')")
                
                if queue_type == "dropsheet":
                    res = self._execute_dropsheet_job(task, idx)
                else:
                    res = self._execute_job(task, queue_type, idx, target_date)
                
                if res == "login_error":
                    self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Failed')")
                    if queue_type == "dp1": self.is_running_all_dp1 = False
                    elif queue_type == "all_geo": self.is_running_all_all_geo = False
                    else: self.is_running_all_dropsheet = False
                    self._window.evaluate_js(f"setRunAllCompleted('{queue_type}', {success_count}, {total})")
                    self._window.evaluate_js(f"triggerLoginRecovery('{queue_type}', {idx}, true)")
                    return

                if res == True:
                    success_count += 1
                status = 'Completed' if res == True else ('Aborted' if res == 'aborted' else 'Failed')
                self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, '{status}')")

            # Reset internal flags
            if queue_type == "dp1":
                self.is_running_all_dp1 = False
            else:
                self.is_running_all_all_geo = False

            # Trigger JS to lock the progress bar and show "X/Y Done" acknowledgement
            self._window.evaluate_js(f"setRunAllCompleted('{queue_type}', {success_count}, {total})")

        threading.Thread(target=_worker_all, daemon=True).start()


def main():
    api = DesktopAPI()
    html_path = BUNDLE_DIR / "ui" / "index.html"

    window = webview.create_window(
        title="ByteRunner",
        url=str(html_path),
        js_api=api,
        width=880,
        height=650,
        resizable=False,
        background_color="#09090b"
    )
    api._window = window
    webview.start(debug=False)


if __name__ == "__main__":
    main()