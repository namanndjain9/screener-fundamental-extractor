# Screener Fundamental Analyzer

An automated **equity fundamental analysis** tool that extracts company financials and valuation ratios from **Screener.in** and generates a **fully formatted Excel analysis dashboard**.

Designed for **investment research, equity screening, and portfolio analysis**, this project automates repetitive data collection and standardizes company comparisons at scale.

---

## 🚀 Key Features

- 🔍 **Automated Fundamental Data Extraction**
  - Revenue, EBITDA, PAT, Net Worth, Debt
  - 5-Year Revenue & PAT CAGR
  - Market price and balance sheet metrics

- 📊 **Valuation & Quality Ratios**
  - Market Cap, EPS, EV/EBITDA, PEG
  - Piotroski Score
  - Debt-to-Equity & Debt-to-EBITDA
  - P/E, P/S, P/BV (Excel-calculated)

- 📈 **Multi-Company Comparison**
  - Analyze multiple companies in a single run
  - Side-by-side structured output

- 📁 **Excel Output (Investment-Ready)**
  - Auto-formatted Excel workbook
  - Percentage & numeric formatting
  - Frozen headers and clean layout

- 🔐 **Secure Credential Handling**
  - Credentials managed via environment variables
  - `.gitignore` prevents sensitive data leaks

---

## 🛠️ Tech Stack

- **Python**
- **Selenium** – Web automation
- **Pandas** – Data processing
- **OpenPyXL** – Excel formatting
- **Regex** – Data cleaning

---

## 📁 Project Structure

screener_fundamental_extractor/
│
├── screener_scraper.py # Main automation & extraction script
├── requirements.txt
├── README.md
└── .gitignore
