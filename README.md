# 📡 BTS Alarm Log Analysis Engine

A Streamlit-based web application to clean, map, and analyze legacy network alarm logs cross-referenced with a PKEY Master Mapping file.

---

## 🚀 Features

- Upload **1–31 daily `.xls` alarm log files** for any month
- Cross-reference with **PKEY Master Mapping** to link incharge/JTO details
- **8 analysis tabs**: Overview, BTS Type & Vendor, Fault Analysis, Duration Deep-Dive, Incharge Analytics, Daily Trend, Summary Report, Raw Data
- Automatic `.xls` → `.xlsx` conversion via LibreOffice
- Dark-themed interactive Plotly charts
- Export processed data as a multi-sheet Excel report

---

## 📁 Expected File Formats

### PKEY Master Mapping (`.xlsx` / `.csv`)
| Column | Description |
|---|---|
| `PKEY` | Primary key |
| `BTSIPID` | BTS IP ID (join key) |
| `SSACODE`, `SSANAME` | SSA details |
| `SDCANAME`, `SDCA` | SDCA details |
| `SITENAME`, `LOCATION` | Site info |
| `incharge`, `JTO INCHARGE` | Responsible personnel |

### Alarm Log Files (`.xls` / `.xlsx`)
| Column | Description |
|---|---|
| `bts_id`, `bts_name` | BTS identifiers |
| `bts_ip_id` | IP ID (for PKEY join) |
| `bts_type` | 2G / 3G / 4G |
| `vendor` | Equipment vendor |
| `bts_down_dt`, `bts_up_dt` | Downtime timestamps |
| `downPeriod` | Human-readable duration (e.g. `0 days 5 hours 48 minutes`) |
| `fault_type` | Fault category |
| `sdca_name`, `ssa_name` | Location hierarchy |

---

## 🛠️ Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/<your-username>/bts-alarm-engine.git
cd bts-alarm-engine

# 2. Create virtual environment
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate

# 3. Install Python dependencies
pip install -r requirements.txt

# 4. Install LibreOffice (for legacy .xls conversion)
# Ubuntu/Debian:
sudo apt-get install -y libreoffice
# macOS:
brew install --cask libreoffice

# 5. Run the app
streamlit run app.py
```

Open your browser at `http://localhost:8501`

---

## ☁️ Deploy to Streamlit Community Cloud

1. Push this repo to GitHub (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **"New app"** → connect your GitHub repo
4. Set:
   - **Repository:** `<your-username>/bts-alarm-engine`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Click **Deploy**

> **Note:** `packages.txt` tells Streamlit Cloud to install `libreoffice` automatically before launching the app.

---

## 📊 Analysis Tabs

| Tab | Contents |
|---|---|
| 📊 Overview | Fleet-level event & downtime distribution |
| 📡 BTS Type & Vendor | 2G/3G/4G breakdown, vendor share, hourly heatmap |
| ⚠️ Fault Analysis | Top fault types, Pareto chart, fault × vendor matrix |
| ⏱️ Duration Deep-Dive | Duration band distribution, long-outage BTS list |
| 👤 Incharge Analytics | Per-incharge downtime & event counts |
| 📅 Daily Trend | Day-by-day event volume and fault type trends |
| 📋 Summary Report | Executive summary with key metrics & export |
| 🗃️ Raw Data | Filterable full dataset view |

---

## 🏢 About

Built for **BSNL TN Circle – Network Operations**  
Version 2.0 | BTS Alarm Analysis Engine
