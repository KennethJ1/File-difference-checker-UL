Quick CLI runner for Difference Checker

This document explains how to run a lightweight CLI version of the comparison without using PowerShell activation scripts (useful when scripts are disabled).

Steps (from project root)

1) Create a virtual environment (one-time):

```
python -m venv .venv
```

2) Install requirements using the venv Python (do NOT rely on activating the PowerShell script if execution policy blocks it):

```
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

3) Run the CLI (example):

```
# compare two Excel files (explicit type)
.venv\Scripts\python.exe compare_cli.py "C:\path\A.xlsx" "C:\path\B.xlsx" excel

# or compare two PDFs
.venv\Scripts\python.exe compare_cli.py "C:\path\A.pdf" "C:\path\B.pdf" pdf
```

What the CLI does
- Calls the existing `run_compare` function in `src.core.runner`.
- Prints progress and writes `comparison_result.txt` in the current folder containing a short repr() of the returned result.

If you prefer a one-click runner, use `run_cli.bat` (provided). If you need the GUI instead, run:

```
.venv\Scripts\python.exe -m src.gui.main
```

Notes for non-admin environments
- Many corporate machines disable PowerShell script execution. Using the venv python executable directly avoids this.
- If a coworker reports the exe is blocked by antimalware, ask them to run the CLI approach above and send back `comparison_result.txt` and the output of `Get-FileHash -Algorithm SHA256 .\difference_checker.exe` from the extracted folder so IT can whitelist.
