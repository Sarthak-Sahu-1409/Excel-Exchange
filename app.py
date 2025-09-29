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

class OutputMode(Enum):
    OVERWRITE = "Overwrite selected cells"

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
        total_currencies = len(CURRENCIES)
        
        try:
            if progress_callback:
                progress_callback(0, total_currencies, "Fetching base rates...")
            
            # Fetch USD rates as base (USD is most commonly used)
            usd_rates = self._fetch_from_api('USD')
            if not usd_rates:
                raise APIError("Failed to fetch USD rates")
                
            # Store USD rates
            self._cache['USD'] = CacheEntry(
                base_currency='USD',
                rates=usd_rates,
                timestamp=time.time(),
                source="api"
            )
            results['USD'] = True
            
            # Calculate all other rates using USD as base
            for i, currency in enumerate(c for c in CURRENCIES if c != 'USD'):
                if progress_callback:
                    progress_callback(i + 1, total_currencies, f"Processing {currency}...")
                
                try:
                    # Get rate to convert from USD to current currency
                    usd_to_curr = usd_rates.get(currency)
                    if usd_to_curr is None:
                        raise ValueError(f"No rate found for {currency}")
                    
                    # Calculate rates for this currency to all others
                    rates = {}
                    for target in CURRENCIES:
                        if target == currency:
                            rates[target] = 1.0  # Same currency
                        elif target == 'USD':
                            rates[target] = 1 / usd_to_curr  # Inverse of USD rate
                        else:
                            # Cross rate: first convert to USD, then to target
                            usd_to_target = usd_rates.get(target)
                            if usd_to_target is None:
                                raise ValueError(f"No rate found for {target}")
                            rates[target] = (1 / usd_to_curr) * usd_to_target
                    
                    # Store the calculated rates
                    self._cache[currency] = CacheEntry(
                        base_currency=currency,
                        rates=rates,
                        timestamp=time.time(),
                        source="calculated"
                    )
                    results[currency] = True
                    
                except Exception as e:
                    logger.error(f"Failed to calculate rates for {currency}: {str(e)}")
                    results[currency] = False
            
            # Save all rates to cache
            self._save_cache()
            logger.info("Successfully refreshed all currency rates")
            
        except Exception as e:
            logger.error(f"Failed to refresh rates: {str(e)}")
            return {curr: False for curr in CURRENCIES}
            
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

import pythoncom

class XLWingsExcelInterface(ExcelInterface):
    def __init__(self):
        self.app: Optional[xw.App] = None
        self.book: Optional[xw.Book] = None
        self._last_connection_state = False
        self._last_active_book = None
        self._thread_initialized = False

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
            # Ensure Excel is in the foreground
            self.app.activate()
            
            # Get the active sheet before showing InputBox
            active_sheet = self.book.sheets.active
            if not active_sheet:
                raise ExcelConnectionError("No active sheet in workbook.")

            # Show the InputBox and get range directly (instead of address)
            try:
                # Use Excel's InputBox to get range directly
                xl_range = self.app.selection
                prompt = "Please select the range of cells to convert."
                title = "Select Range for Conversion"
                xl_range = self.app.api.Application.InputBox(
                    Prompt=prompt,
                    Title=title,
                    Type=8  # xlRange constant
                )
                
                # Handle cancellation
                if not xl_range:
                    logger.info("User cancelled the range selection.")
                    return None
                
                # If we got a valid range from InputBox, get its address and create a new Range object
                if xl_range:
                    # Get the address and sheet name from the COM range object
                    try:
                        address = xl_range.Address
                        sheet = xl_range.Worksheet.Name
                        # Create a new Range object using the public API
                        range_obj = self.book.sheets[sheet].range(address)
                        if range_obj and range_obj.address:
                            return range_obj
                    except Exception as e:
                        logger.warning(f"Error converting Excel range: {e}")
                        return None
                else:
                    logger.warning("Invalid range selection.")
                    return None

            except Exception as e:
                if any(err in str(e).lower() for err in ['cancel', 'user-interrupted', '0x800a03ec']):
                    logger.info("User cancelled the range selection.")
                    return None
                raise

        except Exception as e:
            logger.error(f"Error during range selection: {e}", exc_info=True)
            if 'com_error' in str(type(e)).lower():
                raise ExcelConnectionError("Excel communication error. Please try again.")
            raise ExcelConnectionError(f"Failed to get selection: {str(e)}")

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
        """Write values to Excel. This should be called from the main thread."""
        try:
            # In simplified version, we only overwrite the selected range
            selection.options(expand='table').value = values
            
            # Log success
            logger.info(f"Successfully wrote values to {selection.address} in {selection.sheet.name}")
            
        except Exception as e:
            logger.error(f"Error writing values to Excel: {e}")
            raise

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

    def _build_currency_section(self, parent):
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

    def _build_input_section(self, parent):
        """
        ### MODIFIED ###
        This section now includes manual range coordinate inputs.
        """
        frame = ttk.LabelFrame(parent, text="Range Selection", padding=10)
        frame.pack(fill=tk.X, pady=5)
        
        # Top section for inputs
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Left Top coordinate
        left_frame = ttk.LabelFrame(input_frame, text="Start Cell", padding=5)
        left_frame.pack(side=tk.LEFT, padx=5)
        self.start_cell_var = tk.StringVar(value="A1")
        ttk.Entry(left_frame, textvariable=self.start_cell_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # Right Bottom coordinate
        right_frame = ttk.LabelFrame(input_frame, text="End Cell", padding=5)
        right_frame.pack(side=tk.LEFT, padx=5)
        self.end_cell_var = tk.StringVar(value="A1")
        ttk.Entry(right_frame, textvariable=self.end_cell_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # Apply button
        self.apply_range_button = ttk.Button(input_frame, text="Apply Range", command=self._apply_range)
        self.apply_range_button.pack(side=tk.LEFT, padx=5)
        
        # Range display
        display_frame = ttk.Frame(frame)
        display_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(display_frame, text="Current Range:").pack(side=tk.LEFT, padx=(5, 0))
        self.selection_info_var = tk.StringVar(value="No selection yet.")
        self.selection_info_entry = ttk.Entry(display_frame, textvariable=self.selection_info_var, 
                                            state='readonly', font=('Segoe UI', 9, 'italic'))
        self.selection_info_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)


    def _build_options_section(self, parent):
        # Options section removed as we only have one output mode
        self.output_mode_var = tk.StringVar(value=OutputMode.OVERWRITE.value)

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

    def _apply_range(self):
        """
        Applies the manually entered range coordinates.
        Ensures Excel operations run on the main thread.
        """
        def validate_and_get_range():
            # Get the entered coordinates
            start_cell = self.start_cell_var.get().strip().upper()
            end_cell = self.end_cell_var.get().strip().upper()

            # Basic validation
            if not start_cell or not end_cell:
                raise ValueError("Please enter both start and end cell references")

            # Validate cell references format (e.g., A1, B2)
            if not all(c.isalnum() for c in start_cell + end_cell):
                raise ValueError("Invalid cell reference format")

            return start_cell, end_cell

        def apply_range_in_main_thread():
            try:
                # Connect to Excel in main thread
                if not self.converter.excel.connect():
                    messagebox.showerror("Excel Error", "Excel is not connected.")
                    return

                if not self.converter.excel.book:
                    messagebox.showwarning("No Workbook", "Please select a workbook first.")
                    return

                # Get and validate range
                try:
                    start_cell, end_cell = validate_and_get_range()
                except ValueError as e:
                    messagebox.showerror("Input Error", str(e))
                    self._log(str(e), "error")
                    return

                # Create the range address
                range_address = f"{start_cell}:{end_cell}" if start_cell != end_cell else start_cell

                # Get the range object without activating Excel
                try:
                    active_sheet = self.converter.excel.book.sheets.active
                    selection = active_sheet.range(range_address)
                    self._process_selection(selection)
                    self._log(f"Range {range_address} applied successfully.", "success")
                except Exception as e:
                    messagebox.showerror("Range Error", 
                        f"Could not apply the range {range_address}.\n\nError: {str(e)}")
                    self._log(f"Error applying range: {e}", "error")

            except Exception as e:
                messagebox.showerror("Error", 
                    f"An unexpected error occurred.\n\nError: {str(e)}")
                self._log(f"Unexpected error: {e}", "error")

        # Schedule the Excel operations to run on the main thread
        self.root.after(0, apply_range_in_main_thread)

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
        self.refresh_button.config(state=state)
        self.apply_range_button.config(state=state)

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
                    
                    # Workbook is connected and available
                    pass
                    
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
            self.progress_label['text'] = "Starting rates refresh..."
            self.progress_bar['value'] = 0
            self.root.update_idletasks()
            
            try:
                results = self.converter.rate_provider.refresh_all_rates(self._update_progress)
                total_success = sum(1 for v in results.values() if v)
                self.root.after(0, self._on_refresh_complete, total_success, len(results))
            except Exception as e:
                self.root.after(0, lambda: self._log(f"Error refreshing rates: {e}", "error"))
                self.root.after(0, lambda: self._on_refresh_complete(0, len(CURRENCIES)))
            
        # Show immediate feedback
        self.refresh_button.config(state='disabled')
        self.progress_label['text'] = "Preparing to refresh rates..."
        self.root.update_idletasks()
        
        # Start the refresh in a background thread
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
                precision=self.precision_var.get()
            )
            self._convert_excel(request)
        except tk.TclError:
            messagebox.showerror("Input Error", "Invalid precision value.")

    def _convert_excel(self, request: ConversionRequest):
        def conversion_task(request):
            """This runs in a background thread to do the conversion calculations"""
            try:
                converted_data, stats = self.converter.convert_range(self.excel_values, request, self._update_progress)
                if stats['converted'] > 0:
                    # Schedule the Excel write operation on the main thread
                    self.root.after(0, lambda: self._write_to_excel(converted_data, stats))
                else:
                    self.root.after(0, lambda: self._on_convert_complete(stats, True, None))
            except Exception as e:
                self.root.after(0, lambda: self._on_convert_complete({}, False, e))

        def start_conversion():
            self._set_ui_state(False)
            # Start the conversion in a background thread
            thread = threading.Thread(target=lambda: conversion_task(request), daemon=True)
            thread.start()

        # Verify Excel connection on main thread before starting
        if not self.converter.excel.connect():
            messagebox.showerror("Excel Error", "Lost connection to Excel. Please try again.")
            return

        # Start the conversion process
        start_conversion()

    def _write_to_excel(self, converted_data: List[List[Any]], stats: Dict):
        """Handles writing to Excel on the main thread"""
        try:
            # Verify connection again before writing
            if not self.converter.excel.connect():
                raise ExcelConnectionError("Lost connection to Excel")

            # Get output mode
            output_mode = OutputMode(self.output_mode_var.get())
            
            # Try to activate Excel
            try:
                if self.converter.excel.app:
                    self.converter.excel.app.activate()
            except:
                pass

            # Write the values
            self.converter.excel.write_values(self.current_selection, converted_data, output_mode)
            self._on_convert_complete(stats, True, None)

        except Exception as e:
            logger.error(f"Error writing to Excel: {e}")
            self._on_convert_complete({}, False, 
                ExcelConnectionError(f"Failed to write to Excel: {str(e)}"))

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