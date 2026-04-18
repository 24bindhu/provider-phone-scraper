# provider-phone-scraper

# Provider Phone Scraper Automation

This project automates extraction of phone numbers from the Virginia DBHDS Provider Search website and updates an Excel file.

---

## 🚀 What it does

- Opens DBHDS Provider Search website
- Selects Active providers
- Clicks Submit to load results
- Iterates through provider list
- Opens each provider detail page
- Extracts phone number
- Writes it into Excel file in the correct row
- Skips already processed rows (resume support)

---

## 📁 Files

- `scraper.py` → Main automation script
- `providers.xlsx` → Input file with provider list
- `providers_with_phones.xlsx` → Output file with phone numbers
- `requirements.txt` → Dependencies

---

## ⚙️ Installation

```bash
pip install -r requirements.txt
python -m playwright install



##How to Run
python scraper.py








1. Clone the repo
git clone <your-repo-link>
cd provider-phone-scraper

2. Install dependencies
pip install -r requirements.txt

3.Install Playwright browsers
python -m playwright install


4.Run the script
python scraper.py
