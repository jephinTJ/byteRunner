# ByteRunner: Automated ByteBrew DP1 Metrics Engine

An enterprise-grade desktop automation tool and analytics aggregator built with **Python**, **CustomTkinter**, and **Playwright**. ByteRunner automates cohort selection, multi-part funnel extraction, and dynamic DP1 metrics reporting across multiple mobile titles on the ByteBrew dashboard.

---

## ⚡ Core Features

+ **Clicker–Checker State Machine:** Employs an Action-Assertion architecture that validates DOM updates before triggering subsequent events to eliminate API race conditions and network lag.
+ **Automated Funnel Extraction:** Navigates ByteBrew mechanics and funnel explorers, sets cohorts, and exports multi-part CSVs automatically.
+ **Automated Excel Report Generation:** Parses and cleans raw event tables, computes user conversion and ad frequency metrics, and exports formatted `.xlsx` reports with KPI styling.
+ **Self-Healing DOM Handler:** Detects server packet drops or missing preset configurations and executes automatic reload cycles to prevent corrupted data exports.
+ **Modern Desktop GUI:** Built with CustomTkinter featuring modal calendar pickers, credential managers, and localized task controls.

---

## 🛠 Tech Stack

+ **GUI Framework:** CustomTkinter
+ **Browser Engine:** Playwright (Chromium)
+ **Data Processing:** Pandas, OpenPyXL
+ **Compiler:** PyInstaller

---

## 🚀 Setup & Installation

### 1. Clone the Repository
```bash
git clone [https://github.com/](https://github.com/)<your-username>/byterunner.git
cd byterunner
