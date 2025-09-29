#!/usr/bin/env python3
"""
================================================================================
Excel Currency Converter - Professional Desktop Application
================================================================================
Version: 2.0.0
Author: Gemini
License: MIT

DESCRIPTION:
    Professional desktop currency converter with live Excel integration, GUI,
    caching, and multi-output support. This version uses Excel's native
    InputBox for foolproof range selection.

DEPENDENCIES:
    pip install xlwings requests

USAGE:
    1. Run directly: python app.py
    2. Build executable: pyinstaller --onefile --windowed app.py
================================================================================
"""

# ================================ IMPORTS ====================================
import os
import sys
import json
import time
import logging
import threading
import traceback
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Any
from dataclasses import dataclass, asdict
from pathlib import Path
from enum import Enum

# Third-party imports
try:
    import requests
except ImportError:
    print("ERROR: requests library not installed. Run: pip install requests")
    sys.exit(1)

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings library not installed. Run: pip install xlwings")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog

# ================================ CONSTANTS ==================================
API_BASE_URL = "https://api.frankfurter.app"
CACHE_FILE = Path(__file__).parent / "rates_cache.json"
CACHE_TTL_HOURS = 2
LOG_FILE = Path(__file__).parent / "currency_converter.log"

CURRENCIES = sorted([
    "USD", "EUR", "JPY", "GBP", "AUD", "CAD", "CHF", "CNY", "HKD", "NZD",
    "SEK", "KRW", "SGD", "NOK", "MXN", "INR", "RUB", "ZAR", "TRY", "BRL",
    "PLN", "PHP", "THB", "IDR", "HUF", "CZK", "ILS", "DKK", "MYR", "RON"
])

COLORS = {
    'bg_primary': '#f0f4f8',
    'bg_secondary': '#ffffff',
    'fg_primary': '#2c3e50',
    'accent': '#3498db',
    'success': '#27ae60',
    'warning': '#f39c12',
    'error': '#e74c3c',
    'border': '#dce4ec'
}

# ================================ LOGGING SETUP ==============================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ================================ DATA CLASSES ===============================
@dataclass
class CacheEntry:
    base_currency: str
    rates: Dict[str, float]
    timestamp: float
    source: str = "api"
    def is_expired(self) -> bool:
        return (time.time() - self.timestamp) / 3600 > CACHE_TTL_HOURS

@dataclass
class ConversionRequest:
    from_currency: str
    to_currency: str
    precision: int = 2
    add_suffix: bool = False

class OutputMode(Enum):
    OVERWRITE = "Overwrite selected cells"
    ADJACENT_COLUMN = "Write to adjacent column"
    NEW_SHEET = "Create new sheet"

# ================================ EXCEPTIONS =================================
class CurrencyConverterError(Exception): pass
class ExcelConnectionError(CurrencyConverterError): pass
class APIError(CurrencyConverterError): pass

# ================================ API FUNCTIONS ==============================
class ExchangeRateProvider:
    # This class remains unchanged from the previous version
    def __init__(self, cache_file: Path = CACHE_FILE):
        self.cache_file = cache_file
        self._cache: Dict[str, CacheEntry] = {}
        self._load_cache()

    def _load_cache(self) -> None:
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r') as f:
                    data = json.load(f)
                    for key, entry_data in data.items():
                        self._cache[key] = CacheEntry(**entry_data)
                logger.info(f"Loaded {len(self._cache)} cache entries")
        except Exception as e:
            logger.warning(f"Failed to load cache: {e}")

    def _save_cache(self) -> None:
        try:
            with open(self.cache_file, 'w') as f:
                json.dump({k: asdict(v) for k, v in self._cache.items()}, f, indent=2)
            logger.info("Cache saved.")
        except Exception as e:
            logger.error(f"Failed to save cache: {e}")

    def get_rate(self, from_currency: str, to_currency: str) -> Tuple[float, str]:
        if from_currency == to_currency:
            return 1.0, "same"
        cache_key = from_currency
        if cache_key in self._cache:
            entry = self._cache[cache_key]
            if not entry.is_expired() and to_currency in entry.rates:
                return entry.rates[to_currency], "cache"
        try:
            rates = self._fetch_from_api(from_currency)
            self._cache[cache_key] = CacheEntry(base_currency=from_currency, rates=rates, timestamp=time.time())
            self._save_cache()
            if to_currency in rates:
                return rates[to_currency], "api"
            raise APIError(f"Currency '{to_currency}' not in API response.")
        except Exception as api_exc:
            logger.warning(f"API fetch failed: {api_exc}")
            if cache_key in self._cache and to_currency in self._cache[cache_key].rates:
                return self._cache[cache_key].rates[to_currency], "offline"
            raise APIError(f"No rate for {from_currency}->{to_currency}.")

    def _fetch_from_api(self, base_currency: str) -> Dict[str, float]:
        url = f"{API_BASE_URL}/latest"
        params = {"from": base_currency, "to": ",".join(c for c in CURRENCIES if c != base_currency)}
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        rates = data.get("rates")
        if not rates:
            raise APIError("API returned no rates.")
        return rates

    def refresh_all_rates(self, progress_callback=None) -> Dict[str, bool]:
        results = {}
        for i, currency in enumerate(CURRENCIES):
            if progress_callback:
                progress_callback(i, len(CURRENCIES), f"Fetching {currency}...")
            try:
                rates = self._fetch_from_api(currency)
                self._cache[currency] = CacheEntry(base_currency=currency, rates=rates, timestamp=time.time())
                results[currency] = True
            except Exception as e:
                logger.error(f"Failed to refresh {currency}: {e}")
                results[currency] = False
        self._save_cache()
        return results

# ================================ EXCEL INTERFACE ============================
class ExcelInterface(ABC):
    @abstractmethod
    def connect(self) -> bool: pass
    @abstractmethod
    def get_selection_from_inputbox(self) -> Optional[Any]: pass
    @abstractmethod
    def read_values(self, selection: Any) -> List[List[Any]]: pass
    @abstractmethod
    def write_values(self, selection: Any, values: List[List[Any]], mode: OutputMode) -> None: pass
    @abstractmethod
    def is_connected(self) -> bool: pass

class XLWingsExcelInterface(ExcelInterface):
    def __init__(self):
        self.app: Optional[xw.App] = None
        self.book: Optional[xw.Book] = None
        self._last_connection_state = False
        self._last_active_book = None

    def connect(self) -> bool:
        """Tries to connect to an existing Excel instance and set the active book."""
        try:
            # First, try to get the active app if it exists
            try:
                self.app = xw.apps.active
            except:
                self.app = None
            
            # If no active app found, look for any running Excel instances
            if not self.app and xw.apps:
                for app in xw.apps:
                    try:
                        # Check if the app is responsive
                        _ = app.pid
                        self.app = app
                        break
                    except:
                        continue

            # If we found a valid app, try to get its active book
            current_connection_state = False
            current_book_name = None
            
            if self.app:
                try:
                    if self.app.books:
                        self.book = self.app.books.active
                        if not self.book:  # If no active book, take the first one
                            self.book = self.app.books[0]
                    else:
                        self.book = None
                    
                    current_connection_state = True
                    current_book_name = self.book.name if self.book else None
                    
                    # Only log if there's been a change in state
                    if (current_connection_state != self._last_connection_state or 
                        current_book_name != self._last_active_book):
                        logger.info(f"Connected to Excel. Active book: {current_book_name or 'None'}")
                    
                except Exception as e:
                    if self._last_connection_state:  # Only log if this is a change in state
                        logger.warning(f"Found Excel but couldn't get active book: {e}")
                    self.book = None
                    current_connection_state = True  # Still connected, just no active book
            else:
                if self._last_connection_state:  # Only log if this is a change in state
                    logger.warning("No running Excel instance found")
                self.app = self.book = None
                current_connection_state = False
            
            # Update state tracking
            self._last_connection_state = current_connection_state
            self._last_active_book = current_book_name
            
            return current_connection_state
            
        except Exception as e:
            if self._last_connection_state:  # Only log if this is a change in state
                logger.error(f"Error connecting to Excel: {e}")
            self.app = self.book = None
            self._last_connection_state = False
            self._last_active_book = None
            return False

    def is_connected(self) -> bool:
        try:
            # More thorough connection check
            if not self.app:
                return False
            
            # Try to access essential properties to verify the connection
            try:
                _ = self.app.pid  # This will fail if Excel is not responding
                _ = self.app.books  # This will fail if Excel is not accessible
                return True
            except:
                return False
                
        except Exception:
            return False

    def get_selection_from_inputbox(self) -> Optional[xw.Range]:
        """
        Uses Excel's native InputBox to get a range selection from the user.
        """
        if not self.is_connected() or not self.app or not self.book:
            raise ExcelConnectionError("Excel is not connected to a workbook.")
        try:
            prompt = "Please select the range of cells to convert."
            title = "Select Range for Conversion"
            address = self.app.api.InputBox(prompt, title, Type=8)
            
            if not address:
                logger.warning("User cancelled the InputBox or selection was invalid.")
                return None
            
            return self.book.sheets.active.range(address)
        except Exception as e:
            if "Cancel" in str(e) or "OLE error 0x800a03ec" in str(e):
                logger.info("User cancelled the range selection InputBox.")
                return None
            logger.error(f"Error getting selection from InputBox: {e}", exc_info=True)
            raise ExcelConnectionError(f"Failed to get selection from Excel: {e}")

    def read_values(self, selection: xw.Range) -> List[List[Any]]:
        return selection.options(ndim=2).value

    def list_open_workbooks(self) -> List[str]:
        if not self.is_connected() or not self.app:
            return []
        return [book.name for book in self.app.books]

    def set_active_workbook(self, name: str) -> bool:
        if not self.is_connected() or not self.app:
            return False
        try:
            self.book = self.app.books[name]
            self.book.activate()
            return True
        except Exception as e:
            logger.error(f"Failed to set active workbook to '{name}': {e}")
            return False

    def open_workbook(self, path: str) -> Optional[xw.Book]:
        if not self.is_connected() or not self.app:
            return None
        try:
            book = self.app.books.open(path)
            self.book = book
            self.book.activate()
            return book
        except Exception as e:
            logger.error(f"Failed to open workbook at '{path}': {e}")
            return None

    def write_values(self, selection: xw.Range, values: List[List[Any]], mode: OutputMode) -> None:
        target_range = None
        if mode == OutputMode.OVERWRITE:
            target_range = selection
        elif mode == OutputMode.ADJACENT_COLUMN:
            target_range = selection.offset(column_offset=selection.shape[1])
        elif mode == OutputMode.NEW_SHEET:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_sheet = self.book.sheets.add(f"Converted_{timestamp}", after=self.book.active_sheet)
            target_range = new_sheet.range('A1')
        if target_range:
            target_range.options(expand='table').value = values

# ================================ CONVERSION LOGIC ===========================
class CurrencyConverter:
    # This class remains unchanged
    def __init__(self):
        self.rate_provider = ExchangeRateProvider()
        self.excel = XLWingsExcelInterface()

    def convert_value(self, value: Any, request: ConversionRequest) -> Tuple[Optional[Any], str]:
        if value is None or value == "": return value, "skipped (empty)"
        try:
            rate, source = self.rate_provider.get_rate(request.from_currency, request.to_currency)
            converted = float(value) * rate
            if request.add_suffix:
                return f"{converted:,.{request.precision}f} {request.to_currency}", f"converted ({source})"
            return round(converted, request.precision), f"converted ({source})"
        except (ValueError, TypeError):
            return value, "skipped (non-numeric)"
        except Exception as e:
            return value, f"error: {e}"

    def convert_range(self, values: List[List[Any]], request: ConversionRequest, progress_callback=None) -> Tuple[List[List[Any]], Dict[str, int]]:
        converted_values, stats = [], {"total": 0, "converted": 0, "skipped": 0, "errors": 0}
        total_cells = sum(len(row) for row in values)
        for r_idx, row in enumerate(values):
            converted_row = []
            for c_idx, value in enumerate(row):
                if progress_callback:
                    progress_callback(r_idx * len(row) + c_idx, total_cells, "Processing...")
                converted, status = self.convert_value(value, request)
                converted_row.append(converted)
                stats["total"] += 1
                if "converted" in status: stats["converted"] += 1
                elif "skipped" in status: stats["skipped"] += 1
                else: stats["errors"] += 1
            converted_values.append(converted_row)
        return converted_values, stats

# ================================ GUI APPLICATION ============================

class WorkbookSelectionDialog(simpledialog.Dialog):
    """A dialog to allow the user to select from open workbooks or open a new one."""
    def __init__(self, parent, title, open_workbooks: List[str]):
        self.open_workbooks = open_workbooks
        self.result = None
        self.selection = tk.StringVar()
        if self.open_workbooks:
            self.selection.set(self.open_workbooks[0])
        super().__init__(parent, title)

    def body(self, master):
        master.pack(padx=10, pady=10)
        ttk.Label(master, text="Choose an open workbook or open a new file:").pack(pady=(0,10))
        
        list_frame = ttk.Frame(master)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, exportselection=False)
        for book_name in self.open_workbooks:
            self.listbox.insert(tk.END, book_name)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        if self.open_workbooks:
            self.listbox.selection_set(0)

        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        return self.listbox

    def on_select(self, evt):
        w = evt.widget
        if w.curselection():
            index = int(w.curselection()[0])
            self.selection.set(w.get(index))

    def buttonbox(self):
        box = ttk.Frame(self)
        ttk.Button(box, text="Use Selected", width=15, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Open File...", width=15, command=self.open_file).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def ok(self, event=None):
        if not self.selection.get():
            messagebox.showwarning("No Selection", "Please select a workbook from the list.", parent=self)
            return
        self.result = ("select", self.selection.get())
        super().ok()

    def open_file(self):
        self.result = ("open", None)
        super().ok()

class CurrencyConverterGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Currency Converter Pro")
        self.root.geometry("700x600")
        self.root.minsize(600, 500)
        self.root.configure(bg=COLORS['bg_primary'])
        self.converter = CurrencyConverter()
        self.current_selection: Optional[xw.Range] = None
        self.excel_values: Optional[List[List[Any]]] = None
        self._setup_styles()
        self._build_gui()
        self._center_window()
        self.root.after(100, self._periodic_check)

    def _setup_styles(self):
        # This method is unchanged
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure('TFrame', background=COLORS['bg_primary'])
        style.configure('TLabelframe', background=COLORS['bg_primary'], bordercolor=COLORS['border'])
        style.configure('TLabelframe.Label', background=COLORS['bg_primary'], foreground=COLORS['fg_primary'])
        style.configure('TLabel', background=COLORS['bg_primary'], foreground=COLORS['fg_primary'])
        style.configure('TButton', background=COLORS['bg_secondary'], foreground=COLORS['fg_primary'], bordercolor=COLORS['border'])
        style.map('TButton', background=[('active', COLORS['accent']), ('disabled', COLORS['bg_primary'])])
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), background=COLORS['accent'], foreground=COLORS['bg_secondary'])
        style.map('Primary.TButton', background=[('active', '#2980b9'), ('disabled', COLORS['border'])])

    def _build_gui(self):
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        self._build_status_bar(main_frame)
        self._build_currency_section(main_frame)
        self._build_input_section(main_frame)
        self._build_options_section(main_frame)
        self._build_action_section(main_frame)
        self._build_progress_section(main_frame)
        self._build_log_section(main_frame)

    def _build_status_bar(self, parent):
        status_frame = ttk.Frame(parent, padding=5)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        self.excel_status_label = ttk.Label(status_frame, text="● Excel: Checking...")
        self.excel_status_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.change_workbook_button = ttk.Button(status_frame, text="Change...", command=self._change_workbook, width=10)
        self.change_workbook_button.pack(side=tk.LEFT, padx=(0, 20))
        self.change_workbook_button.config(state='disabled')

        self.api_status_label = ttk.Label(status_frame, text="● API: Checking...")
        self.api_status_label.pack(side=tk.LEFT, padx=(0, 20))
        
    def _change_workbook(self):
        """
        ### NEW METHOD IN v2.0.0 ###
        Opens a dialog to let the user select from open workbooks or open a new file.
        """
        # Force a fresh connection attempt
        if not self.converter.excel.connect():
            messagebox.showerror("Excel Error", "Excel is not connected. Please open Excel and try again.")
            return

        open_workbooks = self.converter.excel.list_open_workbooks()
        if not open_workbooks:
            if messagebox.askyesno("No Workbooks", "No open workbooks found. Would you like to open a new workbook?"):
                self._open_new_workbook()
            return

        dialog = WorkbookSelectionDialog(self.root, "Select Workbook", open_workbooks)
        if not dialog.result:
            return  # User cancelled

        action, selection = dialog.result
        success = False
        
        if action == "select":
            # Ensure Excel is still connected
            if not self.converter.excel.connect():
                messagebox.showerror("Excel Error", "Lost connection to Excel. Please try again.")
                return

            success = self.converter.excel.set_active_workbook(selection)
            if success:
                self._log(f"Switched to workbook: {selection}", "success")
                # Clear any existing selection
                self.current_selection = None
                self.excel_values = None
                self.selection_info_var.set("No selection yet.")
            else:
                self._log(f"Failed to switch to workbook: {selection}", "error")
        elif action == "open":
            self._open_new_workbook()
        
        # Force a connection check to update the status
        self._periodic_check()

    def _open_new_workbook(self):
        """Helper method to handle opening a new workbook."""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Open Excel Workbook",
            filetypes=[
                ("Excel files", "*.xlsx;*.xlsm;*.xls"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            # Ensure Excel is still connected
            if not self.converter.excel.connect():
                messagebox.showerror("Excel Error", "Lost connection to Excel. Please try again.")
                return

            book = self.converter.excel.open_workbook(file_path)
            if book:
                self._log(f"Opened workbook: {book.name}", "success")
                # Clear any existing selection
                self.current_selection = None
                self.excel_values = None
                self.selection_info_var.set("No selection yet.")
            else:
                self._log(f"Failed to open workbook: {file_path}", "error")

    def _build_currency_section(self, parent): # Unchanged
        frame = ttk.LabelFrame(parent, text="Currency Settings", padding=10)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame, text="From:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.from_currency_var = tk.StringVar(value="USD")
        ttk.Combobox(frame, textvariable=self.from_currency_var, values=CURRENCIES, state='readonly', width=10).grid(row=0, column=1)
        ttk.Label(frame, text="To:").grid(row=0, column=2, padx=(20, 5), pady=5, sticky='w')
        self.to_currency_var = tk.StringVar(value="EUR")
        ttk.Combobox(frame, textvariable=self.to_currency_var, values=CURRENCIES, state='readonly', width=10).grid(row=0, column=3)
        ttk.Label(frame, text="Decimals:").grid(row=0, column=4, padx=(20, 5), pady=5, sticky='w')
        self.precision_var = tk.IntVar(value=2)
        ttk.Spinbox(frame, from_=0, to=10, textvariable=self.precision_var, width=5).grid(row=0, column=5)
        self.add_suffix_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Add currency suffix (e.g., '123.45 EUR')", variable=self.add_suffix_var).grid(row=1, columnspan=6, sticky='w', pady=5)

    def _build_input_section(self, parent):
        """
        ### MODIFIED ###
        This section now includes a read-only entry to display the selected range.
        """
        frame = ttk.LabelFrame(parent, text="Input Source", padding=10)
        frame.pack(fill=tk.X, pady=5)
        frame.columnconfigure(1, weight=1)

        self.select_button = ttk.Button(frame, text="Select Range in Excel...", command=self._select_from_excel_interactive)
        self.select_button.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        self.selection_info_var = tk.StringVar(value="No selection yet.")
        self.selection_info_entry = ttk.Entry(frame, textvariable=self.selection_info_var, state='readonly', font=('Segoe UI', 9, 'italic'))
        self.selection_info_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')


    def _build_options_section(self, parent): # Unchanged
        frame = ttk.LabelFrame(parent, text="Excel Output Mode", padding=10)
        frame.pack(fill=tk.X, pady=5)
        self.output_mode_var = tk.StringVar(value=OutputMode.OVERWRITE.value)
        for mode in OutputMode:
            ttk.Radiobutton(frame, text=mode.value, variable=self.output_mode_var, value=mode.value).pack(anchor='w')

    def _build_action_section(self, parent): # Unchanged
        frame = ttk.Frame(parent, padding=(0, 10))
        frame.pack(fill=tk.X, pady=5)
        self.convert_button = ttk.Button(frame, text="Convert", command=self._convert, style='Primary.TButton')
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        self.refresh_button = ttk.Button(frame, text="Refresh Rates", command=self._refresh_rates)
        self.refresh_button.pack(side=tk.LEFT, padx=(0, 10))
        self.log_button = ttk.Button(frame, text="Open Log File", command=self._open_log_file)
        self.log_button.pack(side=tk.RIGHT)
        self.clear_log_button = ttk.Button(frame, text="Clear Log", command=self._clear_log)
        self.clear_log_button.pack(side=tk.RIGHT, padx=(0, 10))

    def _build_progress_section(self, parent): # Unchanged
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        self.progress_bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, expand=True, pady=(0, 5))
        self.progress_label = ttk.Label(frame, text="Ready", anchor='center')
        self.progress_label.pack(fill=tk.X, expand=True)

    def _build_log_section(self, parent): # Unchanged
        frame = ttk.LabelFrame(parent, text="Activity Log", padding=5)
        frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = scrolledtext.ScrolledText(frame, height=8, wrap='word', font=('Consolas', 9), relief='flat')
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config('success', foreground=COLORS['success'])
        self.log_text.tag_config('warning', foreground=COLORS['warning'])
        self.log_text.tag_config('error', foreground=COLORS['error'])
        self._log("Application started. Ready for conversion.")
        
    def _center_window(self): # Unchanged
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_width()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_height()) // 2
        self.root.geometry(f'+{x}+{y}')

    # --- NEW AND MODIFIED METHODS ---

    def _select_from_excel_interactive(self):
        """
        ### NEW WORKFLOW in v2.0.0 ###
        Uses Excel's native InputBox to get a range from the user.
        """
        # Force a connection check before proceeding
        if not self.converter.excel.connect():
            messagebox.showerror("Excel Error", "Excel is not connected. Please open Excel and try again.")
            return

        try:
            # Ensure we have an active workbook
            if not self.converter.excel.book:
                messagebox.showwarning("No Workbook", "Please open or select a workbook first.")
                return

            # This call is blocking and will wait for user input in Excel
            selection = self.converter.excel.get_selection_from_inputbox()
            
            if selection:
                self._process_selection(selection)
                # Activate Excel window after selection
                try:
                    if self.converter.excel.app:
                        self.converter.excel.app.activate()
                except:
                    pass
            else:
                self._log("Range selection was cancelled by the user in Excel.", "warning")
        except Exception as e:
            messagebox.showerror("Selection Error", f"Could not get the selection from Excel.\n\nError: {e}")
            self._log(f"Error during interactive selection: {e}", "error")

    def _process_selection(self, selection: xw.Range):
        """
        ### NEW HELPER METHOD ###
        A single, refactored method to handle a valid selection object,
        whether it comes from the mouse or manual input.
        """
        try:
            # Verify the selection is still valid
            if not selection or not selection.sheet or not selection.address:
                raise ValueError("Invalid selection")

            self.current_selection = selection
            self.excel_values = self.converter.excel.read_values(selection)
            
            # Get detailed selection info
            sheet_name = selection.sheet.name
            address = selection.address
            rows, cols = selection.shape
            cell_count = rows * cols
            
            info_text = f"Selected: {sheet_name}!{address} ({rows}×{cols}, {cell_count} cells)"
            self.selection_info_var.set(info_text)
            self._log(info_text, "success")
            
            # Enable the convert button now that we have a valid selection
            self.convert_button.config(state='normal')
            
        except Exception as e:
            self.current_selection = None
            self.excel_values = None
            self.selection_info_var.set("No selection yet.")
            self.convert_button.config(state='disabled')
            raise ValueError(f"Failed to process selection: {str(e)}")
    
    # --- Other methods are mostly unchanged ---
    
    def _log(self, message: str, tag: str = 'info'):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        self.log_text.see(tk.END)

    def _update_progress(self, current: int, total: int, message: str):
        if total > 0:
            self.progress_bar['value'] = (current / total) * 100
            self.progress_label['text'] = message
            self.root.update_idletasks()

    def _set_ui_state(self, enabled: bool):
        state = 'normal' if enabled else 'disabled'
        # Convert button is only enabled if a selection has been made
        self.convert_button.config(state='normal' if self.excel_values else 'disabled')
        
        # Other buttons are enabled based on the general state
        for widget in [self.refresh_button, self.select_button]:
            widget.config(state=state)

    def _periodic_check(self):
        """
        Periodically checks the connection to Excel and the API.
        This runs in a background thread to avoid freezing the GUI.
        """
        def task():
            # Check Excel and get detailed status
            excel_ok = self.converter.excel.connect()
            book_name = None
            
            if excel_ok:
                try:
                    # Get current workbook name
                    if self.converter.excel.book:
                        book_name = self.converter.excel.book.name
                    
                    # Only update button state if needed
                    def update_button():
                        current_state = self.change_workbook_button['state']
                        new_state = 'normal' if excel_ok else 'disabled'
                        if current_state != new_state:
                            self.change_workbook_button.config(state=new_state)
                    
                    self.root.after(0, update_button)
                    
                except Exception as e:
                    if book_name != "Error getting workbook info":  # Only log state changes
                        logger.warning(f"Error getting Excel details: {e}")
                    book_name = "Error getting workbook info"
                    excel_ok = False
            else:
                book_name = "Not Connected"

            # Check API (less frequently)
            api_ok = False
            if not hasattr(self, '_last_api_check') or time.time() - self._last_api_check > 10:
                try:
                    self.converter.rate_provider.get_rate("USD", "EUR")
                    api_ok = True
                except Exception:
                    api_ok = False
                self._last_api_check = time.time()
            else:
                # Use the last known API state
                api_ok = hasattr(self, '_last_api_state') and self._last_api_state
            
            self._last_api_state = api_ok
            
            # Schedule GUI update on the main thread
            self.root.after(0, self._update_connection_status, excel_ok, api_ok, book_name)
            
            # Schedule the next check (every 5 seconds instead of 2)
            self.root.after(5000, self._periodic_check)

        # Run the check in a thread to not block the GUI
        if not hasattr(self, '_check_thread') or not self._check_thread.is_alive():
            self._check_thread = threading.Thread(target=task, daemon=True)
            self._check_thread.start()

    def _update_connection_status(self, excel_ok: bool, api_ok: bool, book_name: Optional[str]):
        if excel_ok and book_name != "No active workbook":
            self.excel_status_label.config(text=f"● Excel: Connected ({book_name})", foreground=COLORS['success'])
            self._set_ui_state(True)
        elif excel_ok:
            self.excel_status_label.config(text=f"● Excel: Connected (No active workbook)", foreground=COLORS['warning'])
            self._set_ui_state(False)
        else:
            self.excel_status_label.config(text="● Excel: Not Connected", foreground=COLORS['error'])
            self._set_ui_state(False)

        if api_ok:
            self.api_status_label.config(text="● API: Online", foreground=COLORS['success'])
        else:
            self.api_status_label.config(text="● API: Offline/Error", foreground=COLORS['warning'])

    def _refresh_rates(self):
        def task():
            self._set_ui_state(False)
            results = self.converter.rate_provider.refresh_all_rates(self._update_progress)
            self.root.after(0, self._on_refresh_complete, sum(1 for v in results.values() if v), len(results))
        threading.Thread(target=task, daemon=True).start()

    def _on_refresh_complete(self, success_count: int, total: int):
        self._set_ui_state(True)
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Ready"
        self._log(f"Refreshed {success_count}/{total} currencies.", 'success' if success_count == total else 'warning')

    def _convert(self):
        if not self.excel_values:
            messagebox.showwarning("No Input", "Please select a range from Excel first using one of the methods above.")
            return
        
        try:
            request = ConversionRequest(
                from_currency=self.from_currency_var.get(),
                to_currency=self.to_currency_var.get(),
                precision=self.precision_var.get(),
                add_suffix=self.add_suffix_var.get()
            )
            self._convert_excel(request)
        except tk.TclError:
            messagebox.showerror("Input Error", "Invalid precision value.")

    def _convert_excel(self, request: ConversionRequest):
        def task():
            self._set_ui_state(False)
            try:
                converted_data, stats = self.converter.convert_range(self.excel_values, request, self._update_progress)
                write_success = False
                if stats['converted'] > 0:
                    output_mode = OutputMode(self.output_mode_var.get())
                    self.converter.excel.write_values(self.current_selection, converted_data, output_mode)
                    write_success = True
                self.root.after(0, self._on_convert_complete, stats, write_success, None)
            except Exception as e:
                self.root.after(0, self._on_convert_complete, {}, False, e)
        threading.Thread(target=task, daemon=True).start()

    def _on_convert_complete(self, stats: Dict, write_success: bool, error: Optional[Exception]):
        self._set_ui_state(True)
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Ready"
        if error:
            summary = f"An unexpected error occurred: {error}"
            messagebox.showerror("Conversion Failed", summary)
            return
        summary = (f"Conversion finished. Total: {stats.get('total', 0)}, "
                   f"Converted: {stats.get('converted', 0)}, Skipped: {stats.get('skipped', 0)}, "
                   f"Errors: {stats.get('errors', 0)}.")
        if stats.get('errors', 0) > 0 or not write_success and stats.get('converted', 0) > 0:
            messagebox.showerror("Conversion Finished with Errors", summary)
        else:
            messagebox.showinfo("Conversion Successful", summary)

    def _clear_log(self):
        self.log_text.delete('1.0', tk.END)
        self._log("Log cleared.")

    def _open_log_file(self):
        if not LOG_FILE.exists():
            messagebox.showwarning("Log File", "Log file does not exist yet.")
            return
        try:
            if sys.platform == "win32": os.startfile(LOG_FILE)
            elif sys.platform == "darwin": os.system(f'open "{LOG_FILE}"')
            else: os.system(f'xdg-open "{LOG_FILE}"')
        except Exception as e:
            self._log(f"Failed to open log file: {e}", 'error')

# ================================ MAIN =======================================
def main():
    root = tk.Tk()
    try:
        app = CurrencyConverterGUI(root)
        root.mainloop()
    except Exception as e:
        logger.critical(f"A fatal error occurred: {e}\n{traceback.format_exc()}")
        messagebox.showerror("Fatal Error", f"A critical error occurred: {e}\nSee {LOG_FILE} for details.")

if __name__ == "__main__":
    main()