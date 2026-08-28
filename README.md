# Trend Micro Billing Comparator

A small desktop program for comparing two monthly Trend Micro MSP customer-summary workbooks.

The comparator focuses on:

- new and dropped customers;
- service plan changes;
- provisioned-seat changes; and
- used-seat changes.

It supports both the original complete Trend Micro download and carved-down workbooks, provided the file still contains the **Customer**, **Service Plan**, **Provisioned**, and **Used** columns. The program locates the header automatically, ignores total rows, and uses the Customer ID when both files include it. Otherwise, it matches customers by a normalized customer name.

## Download the Windows app

Download the latest version from the [GitHub Releases page](https://github.com/sitescomputer/StatementComparator-TrendMicro/releases/latest), or download [TrendMicroComparator.exe](https://github.com/sitescomputer/StatementComparator-TrendMicro/releases/latest/download/TrendMicroComparator.exe) directly.

No Python installation is required. Save the EXE anywhere on your computer and double-click it to run. Because the application is currently unsigned, Windows may display a SmartScreen warning. Only continue if you downloaded the file from this repository's official release page.

## Run from source

Install Python 3.11 or newer, then run:

```powershell
py -m pip install -r requirements.txt
py TrendMicroComparator.py
```

In the window, select the older month first and the newer month second. The comparison CSV is saved in the current user's Downloads folder and can be opened directly in Excel.

The CSV includes the old and new values, numeric seat differences, and a `Status` column. Unchanged rows use `-`, while changed rows can contain one or more of these labels:

- `NEW CUSTOMER`
- `DROPPED CUSTOMER`
- `SERVICE PLAN CHANGE`
- `PROVISIONED CHANGE`
- `USED SEATS CHANGE`
- `CUSTOMER NAME CHANGE`
- `NEW SERVICE PLAN`
- `DROPPED SERVICE PLAN`

You can also run it without the GUI:

```powershell
py TrendMicroComparator.py --old .\June.xlsx --new .\July.xlsx --output .\comparison.csv
```

## Build the Windows EXE

Run the included build script from PowerShell:

```powershell
.\build_exe.ps1
```

The script installs the build dependencies, runs the tests, and creates:

```text
dist\TrendMicroComparator.exe
```

The EXE is a single-file, windowed application suitable for attaching to a GitHub release. Windows may show a SmartScreen warning for unsigned internal applications; code-signing the final EXE avoids that warning.

The included GitHub Actions workflow also builds the EXE on demand. Pushing a tag such as `v1.0.0` builds the program and attaches it to that tag's GitHub release.
