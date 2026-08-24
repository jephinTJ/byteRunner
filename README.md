# ByteRunner: ByteBrew DP1 Automation Engine

ByteRunner is a desktop automation and analytics tool built with Python, CustomTkinter, and Playwright. It automates cohort configuration, multi-part funnel extraction, and DP1 KPI aggregation for mobile games on the ByteBrew dashboard.

---

## Features

* **Clicker–Checker State Machine:** Implements action-assertion verification across DOM elements before executing steps to avoid race conditions and network lag.
* **Automated Funnel Extraction:** Navigates Funnel and Mechanic explorers, sets date ranges, configures cohorts, and downloads multi-part datasets.
* **Automated DP1 KPI Reports:** Merges Part A and Part B exports into styled `.xlsx` spreadsheets calculating level progressions, ad views, and average ads per user.
* **Self-Healing API Recovery:** Detects dropped network packets or missing Geo filters in real time and triggers automatic reload loops.
* **Desktop GUI:** Built with CustomTkinter, featuring a calendar date picker, Google Sheets task sync, and local credential storage.

---

## Project Structure

```text
├── app.py                   # Main GUI application
├── bytebrew_downloader.py   # Playwright automation & data processing engine
├── icon.ico                 # Application window and taskbar icon
├── requirements.txt         # Python dependencies
└── README.md                # Project documentation
```

---

## Installation & Setup

### 1. Clone the Repository
```bash
git clone [https://github.com/](https://github.com/)<jephinTJ>/byteRunner.git
cd byterunner
```

### 2. Set Up Virtual Environment & Dependencies
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Install Playwright Chromium Browser
```bash
playwright install chromium
```

### 4. Run the Application
```bash
python app.py
```

---

## Building Standalone Executable (.exe)

Compile the application into a single standalone binary using PyInstaller:

```bash
pyinstaller --noconfirm --onefile --windowed --icon="icon.ico" --add-data="icon.ico;." --collect-all customtkinter app.py
```

The compiled `.exe` will be generated inside the `dist/` directory.

---

## Output Architecture

* **Excel Reports:** `Upload files/YYYY-MM-DD/<Output_Name>.xlsx`
* **Execution Logs:** `Upload files/YYYY-MM-DD/execution_log.txt`
* **Local Credentials:** `System Files/credentials.json` (created on first save)
