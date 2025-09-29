# Excel Exchange

### 🖥️ Windows Setup

git clone https://github.com/Sarthak-Sahu-1409/Excel-Exchange.git
cd Excel-Exchange

python -m venv venv
.\venv\Scripts\activate

pip install requests xlwings tk

python app.py

### 🍏 macOS Setup

git clone https://github.com/Sarthak-Sahu-1409/Excel-Exchange.git
cd Excel-Exchange

python3 -m venv venv
source venv/bin/activate

pip3 install requests xlwings tk

python3 app.py

## Features

*   **Live Exchange Rates:** Fetches the latest currency exchange rates from a reliable API.
*   **Excel Integration:** Reads and writes data to your active Excel spreadsheet using `xlwings`.
*   **Currency Conversion:** Converts a selected cell or range of cells from one currency to another.
*   **User-Friendly GUI:** A simple and intuitive graphical user interface built with Tkinter.
*   **Caching:** Caches exchange rates locally to improve performance and reduce API calls. The time to live for the entries is 30 mins.
*   **Logging:** Keeps a log of all operations for easy debugging.

