# 📊 Weekly Report Automation

This project automates the full process of **generating, updating, and distributing a weekly Excel performance report**.  
It integrates data from multiple sources, refreshes Power Query connections, captures a dashboard image, and sends an automated Outlook email — all in under 5 minutes.  

Originally developed for internal operations, all files, names, and paths have been **fully anonymized** in this public version.

## ⚙️ Overview

**Automated workflow:**
1. Detects the current week and creates a new Excel file from a predefined template.  
2. Imports and consolidates data from the latest source report.  
3. Filters and appends complementary data from a master dataset.  
4. Refreshes Power Query connections to update all tables.  
5. Hides auxiliary sheets and saves the updated workbook.  
6. Captures a dashboard range and sends an Outlook email with an embedded image.  

**Frequency:** Weekly  
**Execution time:** ~5 minutes  
**Manual time saved:** ≈33 hours per year (~4 workdays)

---

## 🧩 Technologies Used

| Component | Technology | Purpose |
|------------|-------------|----------|
| **Automation** | `Python 3.10+` | Core automation logic |
| **Excel Integration** | `xlwings`, `win32com.client` | Workbook manipulation, Power Query refresh |
| **Email Automation** | `win32com.client` | Outlook email creation and sending |
| **Image Capture** | `PIL (ImageGrab)` | Captures the dashboard as an embedded PNG |
| **Utilities** | `os`, `datetime`, `shutil`, `urllib.parse` | File operations and path management |

---

## 💡 Key Features

- 🧮 **Fully automated report generation** from an Excel template  
- 🔁 **Multi-source integration** — imports from weekly and master datasets  
- 📊 **Automatic refresh** of Power Query connections  
- 📧 **Outlook email delivery** with an embedded dashboard image  
- 🧱 **Modular design** — functions separated by logical workflow steps  
- ⚠️ **Error-safe execution** — handles missing files and connection issues gracefully  

---

## 🔄 Detailed Workflow

### 1️⃣ Template Duplication
- Copies the  `/Template /` file to the `/output/` folder.  
- Renames the file dynamically (based on the week).  
- Automatically updates week references inside the Excel sheet (`Service Levels`).

### 2️⃣ Data Import & Integration
- Opens the source report and copies the latest weekly data into the Auxiliary sheet.
- Opens the master file and filters rows matching the current week (e.g., WEEK-43).
- Copies only the relevant columns into the Next Week sheet.

### 3️⃣ Power Query Refresh & File Preparation
- Triggers all Power Query refreshes via COM API.
- Waits briefly to allow connections to complete.
- Hides auxiliary worksheets and saves the workbook.

### 4️⃣ Outlook Email Automation
- Opens the completed Excel file in background mode.
- Copies a dashboard range as an image.
- Embeds the image in an HTML email body and sends it to recipients.

---

## ⭐ Acknowledgements
This project was developed to streamline reporting workflows, ensuring higher data consistency and freeing up time for analysis instead of repetitive manual tasks.
It demonstrates how Python + Excel + Outlook automation can transform routine processes into efficient, reliable pipelines.
