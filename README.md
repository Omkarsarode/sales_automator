# 🚀 1-Click Excel Automation: Raw Data to Executive Report

**Stop wasting hours on manual copy-pasting, pivot tables, and Excel formatting.** 

This project demonstrates a fully automated data pipeline that takes messy, raw transactional data (CSV/Excel) and instantly converts it into a formatted, boardroom-ready Excel dashboard with zero manual effort.

## 🎯 What This Project Does
Many businesses lose up to 10-15 hours a week manually cleaning data and updating weekly reports. This automation solves that bottleneck. 

* **The Problem:** Raw sales exports are messy, unformatted, and too large to easily analyze.
* **The Solution:** A custom standalone application that cleans the data, calculates key metrics (Revenue, Profit by Region/Category), builds charts, applies corporate styling, and saves the final file automatically.

## ⚙️ How It Works (The Workflow)
This automation runs entirely locally on your machine for maximum data security. It uses **Python** under the hood, but is packaged as a simple desktop app.

1. **User Input:** A simple pop-up window asks the user to select the raw data file (no coding required).
2. **Data Crunching (Pandas):** The script ingests thousands of rows instantly, groups the data, and aggregates sales and profit metrics.
3. **Report Generation (OpenPyXL):** It generates a brand new `.xlsx` file, injecting the summarized data.
4. **Automated Styling:** It applies custom header colors, currency formatting, auto-adjusts column widths, and injects a dynamic Bar Chart.
5. **Output:** A pristine `Automated_Sales_Report.xlsx` is saved directly to your folder in under 3 seconds.

---

## 🛠️ Try It Yourself (Zero Code Required)

Want to see the magic happen on your own machine? You don't need to install Python or know how to code. Just follow these 3 simple steps:

### Step 1: Download the Application and Sample Data
1. Download the executable application file: `sales_automator.exe` (Link to your hosted .exe file / GitHub Releases).
2. Download the sample raw data: `Sample_Superstore.csv` (Link to the raw Kaggle CSV in your repo).

### Step 2: Run the Automation
1. Double-click `sales_automator.exe`.
2. A file selection window will pop up. Select the `Sample_Superstore.csv` file you just downloaded.
3. Wait 2-3 seconds. A pop-up will notify you that the process is complete.

### Step 3: View the Result
Check the folder where your CSV file is located. You will see a brand new file named `Automated_Sales_Report.xlsx`. Open it to see the beautifully formatted tables and charts generated instantly.

---

## 💻 For Developers (Running from Source)
If you prefer to review the source code and run the Python script directly:

1. Clone this repository.
2. Install the required dependencies: 
   `pip install pandas openpyxl`
3. Run the script:
   `python sales_automator.py`

## 🤝 Let's Automate Your Business
This is just a baseline demonstration. This script can be fully customized for your specific business needs to include:
* Connecting directly to a SQL database or API (no CSV downloading required).
* Automatically emailing the final report to your team via Outlook/Gmail.
* Applying complex Machine Learning models to forecast next month's sales.

**Like what you see? Let's chat about automating your most time-consuming workflows.** 
connect me at sarodeomkar1@gmail.com
