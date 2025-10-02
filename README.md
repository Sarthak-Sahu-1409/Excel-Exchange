## Excel Exchange Setup Guide

### 🖥️ Windows Instructions

```
git clone https://github.com/Sarthak-Sahu-1409/Excel-Exchange.git
cd Excel-Exchange

python -m venv venv
.\venv\Scripts\activate

pip install requests xlwings tk
python app.py
```

### 🍏 macOS Instructions

```
git clone https://github.com/Sarthak-Sahu-1409/Excel-Exchange.git
cd Excel-Exchange

python3 -m venv venv
source venv/bin/activate

pip3 install requests xlwings tk
python3 app.py
```

## App Preview

![App Preview](App%20Preview.png "App Preview Screenshot")

## Key Features

- **Live Exchange Rates:** Fetches up-to-date currency rates directly from a reliable exchange API for accurate conversions.
- **Excel Integration:** Easily reads from and writes to your active Excel workbook using the `xlwings` library, enabling seamless currency management.
- **Currency Conversion:** Instantly converts values in selected cells or ranges between different currencies, streamlining your workflow.
- **User-Friendly GUI:** Intuitive graphical interface powered by Tkinter for hassle-free interactions and effortless use.
- **Smart Caching:** Locally caches exchange rates for 30 minutes to boost performance and minimize unnecessary API requests.
- **Comprehensive Logging:** Maintains detailed logs of all operations to simplify troubleshooting and support smooth debugging.
