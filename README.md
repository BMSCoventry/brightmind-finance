# 💼 Brightmind Solutions — Finance Tracker

A free, browser-based income and expense tracker for Brightmind Solutions Limited.  
Hosted on **GitHub Pages** — no server, no database, no subscription needed.

## 🌐 Live App

👉 **https://BMSCoventry.github.io/brightmind-finance/**

---

## ✨ Features

- **Income tracking** — record payments by student, cash or bank transfer, advance / end-of-lesson / deferred
- **Expense tracking** — log by category (Rent, Utilities, Materials, Marketing, etc.), cash or card
- **Dashboard** — monthly summary cards and 6-month bar chart
- **Reports** — full financial year (April–March) breakdown with HMRC SA103 reference figures
- **Excel export** — monthly or full-year multi-sheet workbook
- **JSON export/import** — back up and restore all data
- **Coloured Excel generator** — Python script for beautifully formatted reports with charts

## 💾 Your Data

All data is stored **locally in your browser** using `localStorage`.  
Nothing is sent to any server. Only you can see your data.

> ⚠️ **Important:** Data is tied to the browser you use. To back up or move to another device, use **Export JSON**.

---

## 🚀 How to Host on GitHub Pages

### Step 1 — Create a GitHub account
Go to [github.com](https://github.com) and sign up (free).

### Step 2 — Create a new repository
1. Click the **+** button → **New repository**
2. Name it `brightmind-finance` (or any name you like)
3. Set it to **Public**
4. Click **Create repository**

### Step 3 — Upload the files
1. In your new repository, click **Add file → Upload files**
2. Upload both files:
   - `index.html`
   - `generate_excel_report.py`
3. Click **Commit changes**

### Step 4 — Enable GitHub Pages
1. Go to your repository **Settings**
2. Scroll down to **Pages** (left sidebar)
3. Under **Source**, select **Deploy from a branch**
4. Choose branch: `main`, folder: `/ (root)`
5. Click **Save**

### Step 5 — Visit your app
After 1–2 minutes, your app will be live at:
```
https://YOUR-USERNAME.github.io/REPOSITORY-NAME/
```

---

## 📊 Coloured Excel Reports (Python Script)

The `generate_excel_report.py` script creates a professionally formatted Excel workbook with:

- ✅ 12 monthly sheets (income in green, expenses in red)
- ✅ FY Summary sheet with running totals and bar chart
- ✅ Income breakdown by student (cash vs bank)
- ✅ Expense breakdown by category
- ✅ HMRC SA103 reference figures

### Usage

```bash
# Install (one-time)
pip install openpyxl

# Export JSON from the app first, then run:
python generate_excel_report.py BMS_finance_data.json

# Optional: specify output filename
python generate_excel_report.py BMS_finance_data.json MyReport.xlsx
```

---

## 📁 Files

| File | Purpose |
|------|---------|
| `index.html` | The main web app (open in any browser) |
| `generate_excel_report.py` | Python script for coloured Excel reports |
| `README.md` | This file |

---

## 🔒 Privacy

- No accounts required
- No data leaves your device
- The app works fully offline once loaded
- GitHub Pages only hosts the HTML/JS files — your financial data is never on GitHub

---

*Brightmind Solutions Limited — Finance Tracker*
