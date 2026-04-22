# hcautomation

A Windows-focused automation toolkit for Enterprise Reporting downloads, Excel refresh/macro workflows, Outlook email automation, and DataFrame/file utilities.

## Features

- Automates Enterprise Reporting exports with Selenium + Edge
- Supports direct export and filter-based export flows
- Includes typed filter configuration via `FilterSpec`
- Refreshes Excel workbooks and runs VBA macros through COM
- Sends Outlook emails with attachments
- Cleans and transforms pandas DataFrames
- Loads and validates Product List (PL) files
- Loads UDA and DCS lite extracts
- Provides small file/date utility helpers

## Installation

```bash
pip install hcautomation
```

## Requirements

This package depends on both Python packages and local Windows applications.

### Python dependencies

- `selenium`
- `rsa`
- `pandas`
- `numpy`
- `pywin32`

### System requirements

- Windows
- Microsoft Edge
- Compatible Edge WebDriver / Selenium setup
- Microsoft Excel
- Microsoft Outlook

## Quick Start

### Download a report directly

```python
from hcautomation import ERDownloader

downloader = ERDownloader(download_dir=r"C:\Temp")
downloader.er(
    bookmark_name="Daily Sales",
    save_path=r"C:\Reports\daily_sales.csv"
)
```

### Download a report with filters

```python
from hcautomation import ERDownloader, FilterSpec

downloader = ERDownloader()

date_filter = FilterSpec(
    filter_name="Date",
    start_date="01/03/2026",
    end_date="31/03/2026"
)

downloader.er(
    bookmark_name="Daily Sales",
    save_path=r"C:\Reports\daily_sales.xlsx",
    filter_spec=date_filter
)
```

### Use helper utilities

```python
from hcautomation import Helpers

helpers = Helpers()
```

---

## Exceptions

The package defines the following custom exceptions for Enterprise Reporting and runtime checks:

- `ERDownloaderError`
- `LoginTimeoutError`
- `BookmarkNotFoundError`
- `FileNotUpdatedError`
- `DownloadTimeoutError`
- `ConfirmationTimeoutError`
- `InternalVerificationFailed`

---

# API Reference

## `FilterSpec`

Dataclass used to define filter inputs before applying them in Enterprise Reporting.

### Signature

```python
FilterSpec(
    filter_name: str,
    start_date: str | None = None,
    end_date: str | None = None,
    single_value_filter: str | None = None
)
```

### Purpose

- Use `filter_name="Date"` with both `start_date` and `end_date`
- Use `single_value_filter` for supported single-value filters such as `Territory` or `Location Code`

### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `filter_name` | Yes | `str` | — | Name of the filter |
| `start_date` | Conditional | `str \| None` | `None` | Required when `filter_name == "Date"` |
| `end_date` | Conditional | `str \| None` | `None` | Required when `filter_name == "Date"` |
| `single_value_filter` | Conditional | `str \| None` | `None` | Required for non-date single-value filters |

### Example

```python
date_filter = FilterSpec(
    filter_name="Date",
    start_date="01/03/2026",
    end_date="31/03/2026"
)

territory_filter = FilterSpec(
    filter_name="Territory",
    single_value_filter="UAE"
)
```

---

## `ERDownloader`

Automates Enterprise Reporting login, bookmark navigation, filtering, export, download detection, and final file movement.

### `ERDownloader(download_dir=None)`

Initializes the downloader and loads encrypted credentials.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `download_dir` | No | `str \| Path \| None` | User Downloads folder | Temporary download directory |

#### Example

```python
downloader = ERDownloader(download_dir=r"C:\Reports\Temp")
```

---

### `er(bookmark_name, save_path, filter_spec=None, num_filters=10, user="Omkar", timeout=60)`

Main public method that runs the full Enterprise Reporting export workflow.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `bookmark_name` | Yes | `str` | — | Name of the bookmark/report |
| `save_path` | Yes | `str \| Path` | — | Final output path ending in `.csv` or `.xlsx` |
| `filter_spec` | No | `FilterSpec \| None` | `None` | Filter configuration object |
| `num_filters` | No | `int` | `10` | Number of filter rows to scan |
| `user` | No | `str` | `"Omkar"` | User folder name |
| `timeout` | No | `int` | `60` | Timeout in seconds for UI waits and download confirmation |

#### Example

```python
from hcautomation import ERDownloader, FilterSpec

downloader = ERDownloader(download_dir=r"C:\Temp")

date_filter = FilterSpec(
    filter_name="Date",
    start_date="01/03/2026",
    end_date="31/03/2026"
)

downloader.er(
    bookmark_name="Daily Sales Report",
    save_path=r"C:\Reports\daily_sales.csv",
    filter_spec=date_filter,
    num_filters=10,
    user="Omkar",
    timeout=60
)
```

---

## `Helpers`

Collection of utility methods for dataframe cleanup, Excel COM automation, Outlook mail generation, PL file loading, UDA/DCS extraction, and file/date checks.

### `Helpers()`

Creates a helper instance.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| None | — | — | — | No constructor parameters |

---

### `strip_chars_v3(df, column_names)`

Removes non-numeric characters except `-` and `.` from selected columns, then converts them to `float`.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `df` | Yes | `pd.DataFrame` | — | DataFrame to modify |
| `column_names` | Yes | iterable | — | Columns to clean |

#### Returns

- Cleaned `pd.DataFrame`
- Original dataframe copy if conversion fails

#### Example

```python
helpers = Helpers()
df = helpers.strip_chars_v3(df, ["Sales", "Margin"])
```

---

### `call_macro(excel_file_path, macro_names, save=False)`

Opens an Excel workbook, runs one or more VBA macros, optionally saves the workbook, and closes Excel cleanly.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `excel_file_path` | Yes | `str \| Path` | — | Workbook path |
| `macro_names` | Yes | `str \| Sequence[str]` | — | Macro name or names to execute |
| `save` | No | `bool` | `False` | Save workbook after running macros |

#### Returns

- `dict[str, Any]` mapping each macro name to its return value

#### Example

```python
results = helpers.call_macro(
    excel_file_path=r"C:\Reports\automation.xlsm",
    macro_names=["RefreshData", "FormatOutput"],
    save=True,
)
```

---

### `get_uda_dcs(uda_rename_map=None, dcs_rename_map=None)`

Generates or refreshes lite UDA and DCS files when needed, then loads both into pandas DataFrames.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `uda_rename_map` | No | `Mapping[str, str] \| None` | `None` | Optional rename mapping for UDA columns |
| `dcs_rename_map` | No | `Mapping[str, str] \| None` | `None` | Optional rename mapping for DCS columns |

#### Returns

```python
(uda_df, dcs_df)
```

#### Notes

- Uses `lite_file_generator.xlsm` and runs macro `CompressUDADCS_v1` when fresh files are not already present.
- Checks whether `uda_lite.xlsx` and `dcs_lite.xlsx` were updated recently before regenerating them.

---

### `refresh_excel_safe(excel_file_path)`

Safer Excel refresh helper with explicit COM initialization and cleanup.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `excel_file_path` | Yes | `str` | — | Workbook path |

---

### `send_mail(to_list, cc_list, subject, attachments=[], html_body='', body='', send_flag=False)`

Creates an Outlook email, optionally attaches files, and either displays or sends it.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `to_list` | Yes | `str` | — | To recipients |
| `cc_list` | Yes | `str` | — | CC recipients |
| `subject` | Yes | `str` | — | Email subject |
| `attachments` | No | `list` | `[]` | Attachment paths |
| `html_body` | No | `str` | `''` | HTML body |
| `body` | No | `str` | `''` | Plain text body |
| `send_flag` | No | `bool` | `False` | Send immediately if `True`, otherwise display draft |

#### Example

```python
helpers.send_mail(
    to_list="user@example.com",
    cc_list="manager@example.com",
    subject="Daily Report",
    attachments=[r"C:\Reports\daily_sales.xlsx"],
    body="Please find the report attached.",
    send_flag=True
)
```

---

### `process_semantic_dumps(path, col_rename_map=None, sheet_name=None, skiprows=2, date_cols=(), numeric_cols=(), errors='raise')`

Reads a semantic dump Excel file, normalizes columns, renames them, and applies optional date/numeric conversion.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `path` | Yes | `str \| pd.ExcelFile` | — | Excel file path or `pd.ExcelFile` object |
| `col_rename_map` | Yes | `Mapping[str, str] \| None` | `None` | Column rename mapping |
| `sheet_name` | No | `str \| int \| None` | `None` | Sheet name or index |
| `skiprows` | No | `int` | `2` | Number of rows to skip |
| `date_cols` | No | sequence | `()` | Columns to convert to datetime |
| `numeric_cols` | No | sequence | `()` | Columns to convert to numeric |
| `errors` | No | `str` | `'raise'` | Conversion error handling behavior |

#### Returns

- Processed `pd.DataFrame`
- `None` if opening fails with `PermissionError`

#### Raises

- `ValueError`
- `KeyError`

---

### `clean_exit(message="\nPress Ctrl+C to exit...")`

Prints a message and blocks until the user stops the process with `Ctrl+C`.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `message` | No | `str` | `"\nPress Ctrl+C to exit..."` | Message shown before waiting |

---

### `fetch_pl_files(terr, omni_letter='O', pl_columns=('skuCode','concept'), col_rename_map=None, marketplace=False, dtype_dict=None, fetch_last_month=False)`

Fetches the latest Product List file for a territory, with fallback to the previous month when needed.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `terr` | Yes | `str` | — | Territory code: `QAT`, `UAE`, `KWT`, `KSA`, `OMN`, `BAH`, or `EGP` |
| `omni_letter` | No | `str` | `'O'` | Drive letter prefix |
| `pl_columns` | No | `Sequence[str]` | `('skuCode', 'concept')` | Columns to read from the CSV |
| `col_rename_map` | No | `Mapping[str, str] \| None` | `None` | Optional rename mapping |
| `marketplace` | No | `bool` | `False` | Use marketplace PL path if `True` |
| `dtype_dict` | No | `Mapping[str, str] \| None` | `None` | Optional dtype mapping for `read_csv` |
| `fetch_last_month` | No | `bool` | `False` | Force previous-month lookup |

#### Returns

```python
(df, latest_pl_file)
```

#### Raises

- `ValueError`
- `FileNotFoundError`

#### Example

```python
df, file_path = helpers.fetch_pl_files(
    terr="UAE",
    marketplace=False,
    pl_columns=("skuCode", "concept", "createdTime")
)
```

---

### `is_file_updated(file_path, days=0, raise_on_fail=True)`

Checks whether a file exists and was modified within the last `days` days.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `file_path` | Yes | `Path` | — | File path to check |
| `days` | No | `int` | `0` | Allowed age in days; `0` means today |
| `raise_on_fail` | No | `bool` | `True` | Raise instead of returning `False` on failure |

#### Returns

- `True` if updated within the requested window
- `False` if not updated and `raise_on_fail=False`

#### Raises

- `FileNotFoundError`
- `FileNotUpdatedError`

---

### `define_cust_type(df, cust_id='Customer ID')`

Classifies each customer as `New` or `Repeat` by comparing transaction date and first transaction date.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `df` | Yes | `pd.DataFrame` | — | Input dataframe |
| `cust_id` | No | `str` | `'Customer ID'` | Customer ID column name |

#### Returns

- `pd.DataFrame` with a `new_repeat` column

#### Raises

- `KeyError`
- `AssertionError`

---

### `get_latest_file(path, typ='c')`

Returns the latest file matching a glob pattern based on creation or modification time.

#### Arguments

| Argument | Required | Type | Default | Description |
|---|---|---|---|---|
| `path` | Yes | `str \| Path` | — | Glob pattern |
| `typ` | No | `str` | `'c'` | `'c'` for creation time, `'m'` for modification time |

#### Returns

- Latest matching file path

#### Raises

- `ValueError`

#### Example

```python
latest = helpers.get_latest_file(r"C:\Reports\*.xlsx", typ="m")
```

---

## Notes

- This package is Windows-specific for COM-based Excel and Outlook automation.
- Enterprise Reporting automation depends on the current ER UI structure and labels.
- `get_uda_dcs()` currently contains what looks like a typo in the rename step for DCS: it uses `df_dcs` instead of `dcs`.
- `define_cust_type()` currently merges back using `'Customer ID'` explicitly, even though `cust_id` is configurable.
- `send_mail()` uses a mutable default argument for `attachments`; consider changing it to `None` in code.

## License

Add your project license here, for example:

```text
MIT
```
