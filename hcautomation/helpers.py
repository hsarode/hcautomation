from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

import rsa
import pandas as pd
import numpy as np
from glob import glob
import re
import os
import time
from pathlib import Path
import shutil
from datetime import datetime, date, timedelta
from typing import Mapping, Sequence, Optional, Union
from pandas.api.types import is_datetime64_ns_dtype
from dataclasses import dataclass
import win32com.client
import pythoncom

from importlib import resources
from hcautomation.login_generator import generate_login_file

class ERDownloaderError(Exception):
    pass
class LoginTimeoutError(ERDownloaderError):
    pass
class BookmarkNotFoundError(ERDownloaderError):
    pass
class FilterNotFoundError(ERDownloaderError):
    pass
class DownloadTimeoutError(ERDownloaderError):
    pass
class ConfirmationTimeoutError(ERDownloaderError):
    pass


@dataclass
class FilterSpec:
    filter_name: str
    start_date: str | None = None
    end_date: str | None = None
    single_value_filter: str | None = None

    def __post_init__(self):
        if not self.filter_name:
            raise ValueError("filter_name is required")

        if self.filter_name == "Date":
            if not self.start_date or not self.end_date:
                raise ValueError("Date filter requires start_date and end_date")

        # elif self.filter_name in ("Territory", "Location Code"):
        else:
            if not self.single_value_filter:
                raise ValueError(
                    f"{self.filter_name} filter requires single_value_filter"
                )

class ERDownloader:
    def __init__(self, download_dir=None):
        check_kill_switch()
        with resources.files("hcautomation").joinpath("priv_key.PEM").open("rb") as f:
            private_key = rsa.PrivateKey.load_pkcs1(f.read())

        self.er_login = Path.home() / '.hcautomation' / 'er_login.txt'
        if not self.er_login.exists():
            generate_login_file()

        with self.er_login.open("rb") as f:
            er = f.read()
            self.username, self.password = rsa.decrypt(er, private_key).decode().split('###')

        self.download_dir = Path(download_dir) if download_dir else Path.home()/'Downloads'
        self.driver = None
        self.action = None

    def init_driver(self, url):
        edge_options = Options()
        edge_options.use_chromium = True
        edge_options.add_argument(
            "user-data-dir=C:\\Users\\{}\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default".format(os.getlogin())
        )
        edge_options.add_argument("profile-directory=Default")
        edge_options.add_experimental_option("prefs", {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })

        self.driver = webdriver.Edge(options=edge_options)
        self.driver.get(url)
        self.action = ActionChains(self.driver)
        return self.driver, self.action

    def _login(self):
        wait = WebDriverWait(self.driver, 10)
        try:
            username_box = wait.until(EC.presence_of_element_located((By.NAME, "username")))
            password_box = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        except TimeoutException:
            raise LoginTimeoutError("Timed out trying to wait for Username & Password fields") from None

        username_box.send_keys(self.username)
        password_box.send_keys(self.password)
        password_box.send_keys(Keys.RETURN)

    def _wait_after_login(self):
        wait = WebDriverWait(self.driver, 20)
        try:
            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "//span[contains(text(),'Enterprise Data Team')]")
                )
            )
        except TimeoutException:
            raise LoginTimeoutError("Timed out trying to wait ER to load post login") from None

    def _find_existing_bookmarks(self, timeout=20):
        deadline = time.time() + timeout
        while time.time() < deadline:
            elements = self.driver.find_elements(By.CLASS_NAME, "CatalogActionLink")
            element_text = [elem.text for elem in elements]
            if elements:
                return elements, element_text
            time.sleep(1)
        raise BookmarkNotFoundError("Timed out waiting for bookmark/action links to load")

    def _click_more_on_bookmark(self, bookmark, concept, user, direct_export=True, timeout=20):
        elements, element_text = self._find_existing_bookmarks()
        bookmark_present_status = bookmark in element_text

        if bookmark_present_status:
            deadline = time.time() + timeout
            while time.time() < deadline:
                try:
                    elements[element_text.index(bookmark) + 3].click()
                    break
                except Exception:
                    elements = self.driver.find_elements(By.CLASS_NAME, "CatalogActionLink")
                    element_text = [elem.text for elem in elements]
                    time.sleep(1)
            else:
                raise BookmarkNotFoundError(f"Unable to click More for bookmark: {bookmark}")
        else:
            wait = WebDriverWait(self.driver, 20)

            wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//span[contains(@class,'HeaderMenuBarText') and contains(@class,'HeaderMenuNavBarText') and normalize-space()='Catalog']"
                    )
                )
            ).click()

            wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//span[@class='treeNodeText' and normalize-space()='Shared Folders']")
                )
            ).click()

            self._wait_and_click_action("Concepts", "Expand")
            self._wait_and_click_action(concept, "Expand")
            self._wait_and_click_action(user, "Expand")

            if direct_export:
                self._wait_and_click_action(bookmark, "More")

    def _wait_for_confirmation_dialog(self, timeout=60):
        locator = (By.CSS_SELECTOR, "span.dialogTitle")

        def confirmation_loaded(drv):
            try:
                title = drv.find_element(*locator).text.strip()
                if title == "Confirmation":
                    return True
                if title == "Processing":
                    return False
                return False
            except (NoSuchElementException, StaleElementReferenceException):
                return False

        try:
            WebDriverWait(self.driver, timeout, poll_frequency=0.25).until(confirmation_loaded)
            return True
        except TimeoutException:
            raise ConfirmationTimeoutError(
                f"ER failed to show Confirmation despite waiting {timeout}s"
            ) from None

    def _wait_and_click_action(self, item_title, action_text, timeout=20):
        wait = WebDriverWait(self.driver, timeout)

        def _click(drv):
            try:
                items = drv.find_elements(
                    By.CSS_SELECTOR,
                    ".ListItem.masterAccordionBottomContentAreaPanel.CatalogListVerboseCell"
                )

                for item in items:
                    title = item.find_element(
                        By.CSS_SELECTOR,
                        ".masterHeader.CatalogObjectListItemTitle"
                    ).text.strip()

                    if title == item_title:
                        for action in item.find_elements(By.CSS_SELECTOR, ".CatalogActionLink"):
                            if action.text.strip() == action_text:
                                drv.execute_script("arguments[0].scrollIntoView({block: 'center'});", action)
                                action.click()
                                return True
                return False
            except StaleElementReferenceException:
                return False

        wait.until(_click)

    def _edit_filter(self, filter_spec: FilterSpec, filters=10):
        filter_name = filter_spec.filter_name
        index = None

        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.ID, "cell1_hightlightrow0_hightlighttable0"))
        )

        for i in range(filters):
            text = self.driver.find_element(
                By.ID, f"cell1_hightlightrow{i}_hightlighttable0"
            ).text
            if filter_name in text:
                index = str(i)
                break

        if index is None:
            raise FilterNotFoundError(f"Filter not found: {filter_name}")

        selected_filter = self.driver.find_element(By.ID, f"floatcell_hightlightrow{index}_hightlighttable0")
        selected_filter.find_element(By.XPATH, ".//img[@title='Edit Filter']").click()

        if filter_name == "Date":
            if not filter_spec.start_date or not filter_spec.end_date:
                raise ValueError("Date filter requires start_date and end_date")

            # final_expected_text = f"Date  is between  {filter_spec.start_date} and {filter_spec.end_date}"

            date_inputs = self.driver.find_elements(By.ID, "datePicker_D")
            date_inputs[0].clear()
            date_inputs[0].send_keys(filter_spec.start_date)
            date_inputs[1].clear()
            date_inputs[1].send_keys(filter_spec.end_date)
            self.driver.find_element(By.NAME, "OK").click()

            # actual_text = self.driver.find_element(
            #     By.ID, f"cell1_hightlightrow{index}_hightlighttable0"
            # ).text

            # if actual_text != final_expected_text:
            #     raise ERDownloaderError("Date filter was not applied as expected")

        elif filter_name in ("Territory", "Location Code"):
            if not filter_spec.single_value_filter:
                raise ValueError(f"{filter_name} filter requires single_value_filter")

            dropdown = self.driver.find_elements(By.ID, "dropdownid")[0]
            dropdown.clear()
            dropdown.send_keys(filter_spec.single_value_filter)
            self.driver.find_element(By.NAME, "OK").click()

        else:
            raise ValueError(f"Unsupported filter name: {filter_name}")

    def _perform_export(
        self,
        direct_export,
        export_format,
        filter_spec=None,
        num_filters=10,
    ):
        
        if direct_export:
            for element in self.driver.find_elements(By.CLASS_NAME, "contextMenuOptionText"):
                if element.text == "Export":
                    element.click()
                    break

            for element in self.driver.find_elements(By.CLASS_NAME, "contextMenuOptionText"):
                if element.text == "Data":
                    element.click()
                    break

            if export_format == "csv":
                self.driver.find_element(By.ID, "menuoptionCell_CSV").click()
            elif export_format == "xlsx":
                self.driver.find_element(By.ID, "menuoptionCell_Excel").click()
            else:
                raise ValueError(f"Unsupported export format: {export_format}")

        else:
            for elem in self.driver.find_elements(By.CLASS_NAME, "CatalogActionLink"):
                if elem.text == "Edit":
                    elem.click()
                    break

            element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "criteriaTab_tab"))
            )
            element.click()

            if filter_spec is not None:
                self._edit_filter(filter_spec, num_filters)

            element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "resultsTab_tab"))
            )
            element.click()

            element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "idAnswersCompoundViewToolbar_export_image"))
            )
            element.click()

            element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.NAME, "exportData"))
            )
            element.click()

            export_name = export_format if export_format == "csv" else "excel_data"
            element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.NAME, export_name))
            )
            element.click()

    def _wait_for_download(self, timeout=180, stable_time=2.0):
        folder = self.download_dir
        folder.mkdir(parents=True, exist_ok=True)
        before = {p.name for p in folder.iterdir() if p.is_file()}

        deadline = time.time() + timeout
        candidate = None
        last_size = None
        last_change_time = None

        while time.time() < deadline:
            files = [p for p in folder.iterdir() if p.is_file()]
            new_files = [p for p in files if p.name not in before]

            completed = [
                p for p in new_files
                if not p.name.endswith(".crdownload") and not p.name.endswith(".tmp")
            ]

            if completed:
                newest = max(completed, key=lambda p: p.stat().st_mtime)
                size = newest.stat().st_size

                if candidate != newest or size != last_size:
                    candidate = newest
                    last_size = size
                    last_change_time = time.time()
                else:
                    if time.time() - last_change_time >= stable_time:
                        return candidate

            time.sleep(0.5)

        raise DownloadTimeoutError("Download did not complete in time")

    def _move_dloaded_file(self, dload_path, save_path):
        save_path = Path(save_path)
        save_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(dload_path), str(save_path))

    def er(
        self,
        bookmark_name,
        save_path,
        direct_export=True,
        filter_spec=None,
        num_filters=10,
        concept="Home Center",
        user="Omkar",
    ):
        save_path = Path(save_path)

        if save_path.suffix.lower() not in {".csv", ".xlsx"}:
            raise ValueError(f"save_path must end with .csv or .xlsx. Got: {save_path}")

        if not direct_export and filter_spec is None:
            raise ValueError("filter_spec is required when direct_export=False")

        url = "https://lmeraz.landmarkgroup.com/"

        try:
            self.init_driver(url)
            self._login()
            self._wait_after_login()
            self._click_more_on_bookmark(
                bookmark_name, concept, user, direct_export=direct_export
            )

            self._perform_export(
                direct_export,
                save_path.suffix.replace(".", "").lower(),
                filter_spec=filter_spec,
                num_filters=num_filters,
            )

            self._wait_for_confirmation_dialog()
            file_path = self._wait_for_download()
            self._move_dloaded_file(file_path, save_path)

        finally:
            if self.driver is not None:
                self.driver.quit()
                self.driver = None
                self.action = None

class Helpers:
    def __init__(self):
        check_kill_switch()

    def strip_chars_v3(self, df, column_names):
        df_old = df.copy()
        # Compile the regex pattern to remove non-numeric characters except '-' and '.'
        pattern = re.compile(r'[^\d.-]+')

        try:
            for column in column_names:
                df[column] = (
                    df[column]
                    .astype(str)
                    .str.replace(pattern, '', regex=True)
                    .astype(float)
                )
        except:
            print('{} column failing, check manually. Returning old df'.format(column))
            return df_old
        else:
            return df
        
    def call_macro(self, excel_file_path, macro_name):
        xlapp = win32com.client.DispatchEx("Excel.Application")
        filename = os.path.basename(excel_file_path)
        try:
            wb = xlapp.Workbooks.Open(excel_file_path)
        except:
            xlapp.Application.Quit()
            return 0
        xlapp.DisplayAlerts = False
        for macro in macro_name:
            xlapp.Run(macro)
        wb.Save()
        wb.Close()
        xlapp.Application.Quit()

    def refresh_excel(self, excel_file_path):
        pythoncom.CoInitialize()  # <-- init COM for this thread
        xlapp = win32com.client.DispatchEx("Excel.Application")
        filename = os.path.basename(excel_file_path)
        try:
            wb = xlapp.Workbooks.Open(excel_file_path)
        except:
            xlapp.Application.Quit()
            return 0
        xlapp.DisplayAlerts = False
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        xlapp.Quit()
        xlapp = None
        del xlapp

    def refresh_excel_safe(self, excel_file_path):
        pythoncom.CoInitialize()  # <-- init COM for this thread
        try:
            xlapp = win32com.client.DispatchEx("Excel.Application")
            try:
                wb = xlapp.Workbooks.Open(excel_file_path)
                wb.RefreshAll()
                xlapp.CalculateUntilAsyncQueriesDone()
                wb.Save()
                wb.Close(SaveChanges=True)
            finally:
                xlapp.Quit()
        finally:
            pythoncom.CoUninitialize()  # <-- clean up COM

    def send_mail(self, to_list, cc_list, subject, attachments=[], html_body='', body='', send_flag=False):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_list
        mail.CC = cc_list
        mail.Subject = subject
        if body == '' and html_body != '':
            mail.HTMLBody = html_body
        elif(body != '' and html_body == ''):
            mail.Body = body
        if len(attachments) > 0:
            for attachment in attachments:
                mail.Attachments.Add(attachment)
        if send_flag:
            mail.Send()
        else:
            mail.Display()

    def process_semantic_dumps(path, col_rename_map=None, sheet_name=None, skiprows=2, date_cols=(), numeric_cols=(), errors='raise') -> pd.DataFrame:
        try:
            excel_obj = path if isinstance(path, pd.ExcelFile) else path
            df = pd.read_excel(excel_obj, sheet_name=sheet_name, skiprows=skiprows)
        except PermissionError:
            return None

        cols_to_remove = [col for col in df.columns if 'FormatString' not in col]
        df = df[cols_to_remove]
        df.columns = [col[:-1].split('[')[-1] for col in df.columns]

        if col_rename_map is None:
            raise ValueError(f"[ERROR] Column rename dictionary empty. PFB columns in the df: {df.columns}")
        
        final_cols = set(col_rename_map.values())
        requested_cols = set(date_cols) | set(numeric_cols)

        missing = requested_cols - final_cols
        if missing:
            raise KeyError(
                f"[ERROR] Conversion columns not found in rename map values: {sorted(missing)}. "
                f"[WARN] Allowed columns are: {sorted(final_cols)}"
            )
        overlap = set(date_cols) & set(numeric_cols)
        if overlap:
            raise ValueError(f"[ERROR] Columns cannot be both date and numeric: {sorted(overlap)}")
        
        df = df.rename(columns=col_rename_map)
        for c in date_cols:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors=errors)
                df = df.dropna(subset=[c])

        for c in numeric_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.replace(",", "")
                df[c] = pd.to_numeric(df[c], errors=errors)
                df = df.dropna(subset=[c])
        return df

    def clean_exit(message: str = "\nPress Ctrl+C to exit...") -> None:
        print(message)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nExiting.")

    def fetch_pl_files(terr:str, 
                    omni_letter:str='O', 
                    pl_columns:Sequence[str]=('skuCode','concept'), 
                    col_rename_map: Mapping[str, str] | None = None,
                    marketplace:bool=False, 
                    dtype_dict:Mapping[str,str] | None = None,
                    fetch_last_month:bool=False
                    ) -> pd.DataFrame:
        """
        Fetch latest Product List (PL) file for a territory.
        Falls back to previous month if current month not available.
        """

        terr_map = {'QAT': 'QAT', 'UAE': 'UAE', 'KWT': 'KUW', 'KSA': 'KSA', 'OMN': 'OMA', 'BAH': 'BAH', 'EGP': 'EGP'}
        rev_terr_map = {'QAT': 'QAT', 'UAE': 'UAE', 'KUW': 'KWT', 'KSA': 'KSA', 'OMA': 'OMN', 'BAH': 'BAH', 'EGP': 'EGP'}

        if terr not in terr_map:
            raise ValueError(f"Invalid territory code: {terr}") from None

        terr = terr_map[terr]
        terr_modified = f"{terr}-Marketplace" if marketplace else terr

        today_str = date.today().strftime('%Y%m%d')
        year_month = today_str[:6]
        last_month = (datetime.strptime(year_month + "01", "%Y%m%d") - timedelta(days=1)).strftime('%Y%m')

        if fetch_last_month:
            pl_path = f"{omni_letter}:\\06. Raw Data\\02. Product List Daily\\{terr_modified}\\{last_month}"
        else:
            pl_path = f"{omni_letter}:\\06. Raw Data\\02. Product List Daily\\{terr_modified}\\{year_month}"
        pl_files = glob(os.path.join(pl_path, "*.csv"))

        # Roll back to previous month if needed
        if not pl_files and not fetch_last_month:
            print("[WARN] Current month's PL is not available, fetching last month's")
            pl_path = f"{omni_letter}:\\06. Raw Data\\02. Product List Daily\\{terr_modified}\\{last_month}"
            pl_files = glob(os.path.join(pl_path, "*.csv"))
        elif not pl_files:
            raise ValueError("fetch_last_month arguement was set True while calling the function. Last month's PL not available") from None

        if not pl_files:
            raise FileNotFoundError(f"[ERROR] No PL files found for territory for current & last month: {terr_modified}") from None

        mandatory_cols = ["skuCode", "concept"]
        user_defined_pl_columns = pl_columns + ['Order Location', 'isMP']
        pl_columns = set(pl_columns) | set(mandatory_cols)

        latest_pl_file = max(pl_files, key=os.path.getmtime)
        print(f"[INFO] Found PL file: {os.path.basename(latest_pl_file)} | ", end="")
        df = pd.read_csv(latest_pl_file, dtype=dtype_dict, usecols=pl_columns)
        loaded_rows = len(df)
        
        if "createdTime" in df.columns:
            df["createdTime"] = pd.to_datetime(df["createdTime"], errors="coerce")
        df = df[pd.to_numeric(df['skuCode'], errors='coerce').notna()].copy()
        df['skuCode'] = pd.to_numeric(df['skuCode']).astype('Int64')
        print(f"{loaded_rows - len(df)} Rows removed due to invalid SKUs")
        df['Order Location'] = rev_terr_map[terr]
        df['isMP'] = 'MP' if marketplace else 'HC'
        df = df.drop_duplicates(subset=['skuCode'])
        if terr == 'UAE':
            df = df[df['concept'] != 'OTATH']

        df = df[user_defined_pl_columns]

        if col_rename_map is None:
            print(f"\t[WARN] Column rename dictionary empty. PFB columns in the df: {df.columns}")
        else:
            df = df.rename(columns=col_rename_map)
        return df, latest_pl_file

    def is_file_updated_today(
        file_path: str,
        raise_on_fail: bool = True
    ) -> bool:
        """
        Returns
        -------
        bool
            True if file was modified today, False otherwise
            (only when raise_on_fail=False)
        """
        if not os.path.isfile(file_path):
            if raise_on_fail:
                raise FileNotFoundError(f"[ERROR] File not found: {file_path}") from None
            return False

        last_modified = date.fromtimestamp(os.path.getmtime(file_path))

        if last_modified != date.today():
            if raise_on_fail:
                raise FileNotFoundError(
                    f"[ERROR] File has not been updated today. "
                    f"Last modified on: {last_modified}"
                ) from None
            return False

        return True

    def define_cust_type(df, cust_id='Customer ID'):
        cols = {'Date', 'First Transcation Date', 'Customer ID'}
        if not cols.issubset(df.columns):
            raise KeyError('Date, First Transcation Date', 'Customer ID columns not in df')
        
        cols = ['Date', 'First Transcation Date']
        assert all(is_datetime64_ns_dtype(df[c]) for c in cols), "One or more columns are not datetime64[ns]"

        temp = df[[cust_id, 'Date', 'First Transcation Date']].drop_duplicates(subset=[cust_id]).copy()
        temp['new_repeat'] = np.where(temp['Date'] == temp['First Transcation Date'], 'New', 'Repeat')
        df = pd.merge(df, temp[['Customer ID', 'new_repeat']], on=['Customer ID'], how='left', validate='m:1')
        return df

    def get_latest_file(path, typ='c'):
        path = str(path)
        if typ == 'c':
            return max(glob(path), key=os.path.getctime)
        elif typ == 'm':
            return max(glob(path), key=os.path.getmtime)
        else:
            raise ValueError(f'[ERROR] {typ} value provided as type. Allowed values are c=createdTime, m=modifiedTime')
        
class InternalVerificationFailed(Exception):
    """Raised when global kill switch is enabled"""
    pass

KILL_SWITCH_URL = "https://docs.google.com/spreadsheets/d/1_JdOBe7SxX9oYdQr5UZ_YMElwp6cCN9ZXttq6v3ilC4/export?format=csv"

def is_kill_switch_on() -> bool:
    ATTEMPT = 1
    MAX_ATTEMPTS = 5
    while ATTEMPT <= MAX_ATTEMPTS:
        try:
            df = pd.read_csv(KILL_SWITCH_URL)
            value = df.loc[df['Item'] == 'Kill Switch', 'Value'].iloc[0]
            ATTEMPT += 1
            return str(value).strip().upper() == "TRUE"
        except Exception as e:
            ATTEMPT += 1
            continue

def check_kill_switch():
    if is_kill_switch_on():
        raise InternalVerificationFailed("Execution blocked by kernel")