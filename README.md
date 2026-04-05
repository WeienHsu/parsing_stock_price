# 📈 TW & US Stock Price Scraper with Google Drive Sync

A one-click tool that fetches real-time closing prices for Taiwan (TWSE/TPEX) and US (S&P 500) stocks, merges them into a formatted Excel report, and automatically uploads it to Google Drive.

## ✨ Features

- **Taiwan Stock Scraping** — Collects listed and OTC stock quotes from HiStock, TWSE, and TPEX simultaneously
- **US Stock Scraping** — Supports all S&P 500 constituents plus a customizable watchlist via yfinance
- **Automated Report Generation** — Deduplicates and merges TW/US data side-by-side into a formatted `.xlsx` file
- **Google Drive Sync** — Uploads Excel as Google Sheets automatically; creates new or updates existing files
- **Standalone Executable** — Supports packaging via PyInstaller for Windows `.exe` or macOS app

## 📁 Project Structure

```
parsing_stock_price/
├── parsing_and_upload_to_drive.py   # Main script: scrape → merge → export → upload
├── parsing_me.py                    # Legacy version (TW stocks only, from wespai.com)
├── stock_list.csv                   # S&P 500 constituent list (~504 symbols)
├── cacert.pem                       # SSL certificate bundle
├── bagofmoney_5108.ico              # Application icon
├── requirements.txt                 # Python dependencies
├── .env                             # Environment variables (not tracked in git)
├── token.json                       # Google OAuth token (not tracked in git)
└── client_secret_*.json             # Google OAuth credentials (not tracked in git)
```

## 🚀 Getting Started

### 1. Prerequisites

- Python 3.9+
- pip

### 2. Install Dependencies

```bash
# Using a virtual environment is recommended
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
# .venv\Scripts\activate    # Windows

pip install -r requirements.txt
```

### 3. Configure Environment Variables

Create a `.env` file in the project root:

```bash
# [Required] Google OAuth credential filename (without .json extension)
CRED_NAME=client_secret_xxxxxxxx

# [Required] Target Google Drive folder ID
FOLDER_ID=your_google_drive_folder_id

# [Optional] Specific Google Drive file ID to update (auto-matches by name if empty)
FILE_ID=

# [Optional] Local path to copy the generated Excel file to
CP_EXCEL_TO=

# [Optional] Additional US stock symbols to track, comma-separated
CUSTOM_US_STOCKS=GBTC,GLD,ITA,SOXX,TSM,VOO,QQQ
```

> **⚠️ Security Notice:** `.env`, `token.json`, and `client_secret_*.json` are all listed in `.gitignore`. Never commit these files to version control.

### 4. Google Drive API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create a project with the **Google Drive API** enabled
2. Create an **OAuth 2.0 Client ID** (application type: Desktop)
3. Download the credential JSON file and place it in the project root
4. Set the filename (without `.json`) in `CRED_NAME` within `.env`
5. On first run, a browser window will open for OAuth authorization; `token.json` is generated automatically upon completion

### 5. Run

```bash
python parsing_and_upload_to_drive.py
```

## 📊 Data Sources

| Market | Source | Method | Description |
|--------|--------|--------|-------------|
| 🇹🇼 TW Listed | [HiStock](https://histock.tw/stock/rank.aspx) | HTML scraping (BeautifulSoup) | Full listed stock rankings |
| 🇹🇼 TW Listed Avg | [TWSE](https://www.twse.com.tw/) | CSV download | Daily closing average prices |
| 🇹🇼 TW OTC | [TPEX](https://www.tpex.org.tw/) | CSV download | OTC fixed-price trading data |
| 🇺🇸 US Stocks | S&P 500 + custom list | yfinance API | Batch download (10 per batch, single-threaded) |

### US Stock List

- By default, reads S&P 500 constituents from the local `stock_list.csv`
- If the CSV does not exist, it automatically scrapes the latest list from [Wikipedia](https://en.wikipedia.org/wiki/List_of_S%26P_500_companies) and caches it locally
- Additional symbols can be added via `CUSTOM_US_STOCKS` in `.env`

## 🔄 Execution Flow

```
1. Scrape Taiwan stocks
   ├── HiStock (listed stocks)
   ├── TWSE CSV (listed average prices)
   └── TPEX CSV (OTC fixed prices)
2. Merge & deduplicate TW data
3. Scrape US stocks
   ├── Load stock_list.csv (S&P 500)
   ├── Append custom symbols from .env
   └── Batch download via yfinance (10 per batch)
4. Merge & deduplicate US data
5. Horizontally merge TW + US → save as "股價N.xlsx"
6. Upload to Google Drive (converted to Google Sheets)
7. [Optional] Copy Excel to a specified local path
```

## 📦 Output Format

The generated `股價N.xlsx` contains a single sheet named "股價N" with the following layout:

| 代號 (Symbol) | 名稱 (Name) | 價格 (Price) | 成交量 (Volume) | 美股代號 (US Symbol) | 美股名稱 (US Name) | 美股價格 (US Price) | 美股成交量 (US Volume) |
|---------------|-------------|-------------|----------------|---------------------|-------------------|--------------------|-----------------------|
| 2330 | TSMC | 590.00 | 25,381 | AAPL | Apple Inc. | 178.72 | 48,235,678 |

- Row 1: Execution timestamp
- Row 2: Column headers
- Price columns are formatted as `#,##0.00`

## 🛠️ Building Standalone Executables

### Windows

```bash
pyinstaller -F parsing_and_upload_to_drive.py
```

> Must be built in a Windows environment. The `.exe` output is located in the `dist/` directory.
> Place `cacert.pem`, `.env`, and `client_secret_*.json` in the same folder as the executable.

### macOS

```bash
pyinstaller --onefile --windowed --icon=./bagofmoney_5108.ico parsing_and_upload_to_drive.py
```

## ❓ FAQ

### Q: US stock download fails with `TypeError: 'NoneType' object is not subscriptable`
This occurs when Yahoo Finance API blocks or returns malformed data. The script already includes batch processing (10 per batch + single-threaded) to mitigate this. Usually, simply re-running resolves the issue.

### Q: `OperationalError: unable to open database file`
The internal SQLite cache used by yfinance gets locked due to concurrent writes. The script uses `threads=False` to prevent this. If it still occurs, try deleting the yfinance cache directory and re-running.

### Q: Taiwan stock data returns empty
TW stock data sources only update after market close on trading days. Data may be unavailable during non-trading hours or holidays.

### Q: Google Drive upload fails
1. Verify `CRED_NAME` and `FOLDER_ID` are correctly set in `.env`
2. Check if `token.json` has expired — delete it and re-authorize if needed
3. Confirm that the Google Drive API is enabled in Google Cloud Console

## 🧰 Tech Stack

- **Python 3.9+**
- [pandas](https://pandas.pydata.org/) — Data processing and DataFrame operations
- [yfinance](https://github.com/ranaroussi/yfinance) — Yahoo Finance US stock quote API
- [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/) — HTML web scraping
- [xlsxwriter](https://xlsxwriter.readthedocs.io/) — Excel file generation
- [google-api-python-client](https://github.com/googleapis/google-api-python-client) — Google Drive API
- [requests](https://docs.python-requests.org/) — HTTP requests
- [PyInstaller](https://pyinstaller.org/) — Standalone executable packaging

## 📄 License

This project is licensed under the [MIT License](LICENSE).
