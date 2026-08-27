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
CREDENTIALS_FILE = SYSTEM_DIR / "credentials.json"

class DesktopAPI:
    """JS Bridge: Connects Tailwind frontend directly to Playwright scraper engine."""
    def __init__(self):
        self._window = None  # Private attribute prevents PyWebView infinite recursion crash
        self.tasks_dp1 = []
        self.tasks_all_geo = []
        self.is_running_all_dp1 = False
        self.is_running_all_all_geo = False
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

    def get_initial_state(self):
        return {
            "selected_date": (date.today() - timedelta(days=1)).isoformat(),
            "status": "MindYourLogic Studios"
        }

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

            return {
                "dp1": self.tasks_dp1,
                "all_geo": self.tasks_all_geo
            }
        except Exception as e:
            print(f"[DesktopAPI] Error loading sheets: {e}", flush=True)
            return {"dp1": [], "all_geo": []}

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

    def run_single_task(self, queue_type, idx, target_date):
        tasks = self.tasks_dp1 if queue_type == "dp1" else self.tasks_all_geo
        if idx >= len(tasks):
            return

        def _worker():
            self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Running')")
            res = self._execute_job(tasks[idx], queue_type, idx, target_date)
            
            if res == "login_error":
                self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Failed')")
                self._window.evaluate_js(f"triggerLoginRecovery('{queue_type}', {idx}, false)")
                return
                
            status = 'Completed' if res == True else ('Aborted' if res == 'aborted' else 'Failed')
            self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, '{status}')")

        threading.Thread(target=_worker, daemon=True).start()

    def run_all_tasks(self, queue_type, target_date, start_idx=0):
        tasks = self.tasks_dp1 if queue_type == "dp1" else self.tasks_all_geo
        if not tasks:
            return

        def _worker_all():
            if queue_type == "dp1":
                self.is_running_all_dp1 = True
            else:
                self.is_running_all_all_geo = True

            total = len(tasks)
            self._window.evaluate_js(f"setRunAllState('{queue_type}', true, {start_idx}, {total})")

            success_count = start_idx
            for idx, task in enumerate(tasks):
                if idx < start_idx:
                    continue
                    
                is_running = self.is_running_all_dp1 if queue_type == "dp1" else self.is_running_all_all_geo
                if not is_running:
                    break
                self._window.evaluate_js(f"setRunAllState('{queue_type}', true, {idx+1}, {total})")
                self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Running')")
                res = self._execute_job(task, queue_type, idx, target_date)
                
                if res == "login_error":
                    self._window.evaluate_js(f"updateTaskStatus('{queue_type}', {idx}, 'Failed')")
                    if queue_type == "dp1":
                        self.is_running_all_dp1 = False
                    else:
                        self.is_running_all_all_geo = False
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