# Difference Checker

A small desktop tool to visually and semantically compare PDFs and Excel files.

Quick start

1. Create and activate a virtual environment (Windows PowerShell):

   ```powershell
   python -m venv .venv
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

2. Run the GUI

   ```powershell
   python -m src.gui.main
   ```

Packaging

- Use the included PowerShell helper to build a Windows single-folder app using PyInstaller:

   ```powershell
   # from project root
   .\tools\build_windows.ps1
   # or to force a clean build
   .\tools\build_windows.ps1 -Clean
   ```

- Under the hood the script creates a temporary virtualenv, installs requirements, runs `pyinstaller` with
   `difference_checker.spec` and copies the `dist\difference_checker` folder to `artifacts\` with a timestamp.

- Troubleshooting PySide6 plugin errors: if the packaged app fails to start due to missing Qt plugins, verify
   the `PySide6/plugins/platforms` folder is present inside the `dist\difference_checker` folder. If missing,
   add an explicit `datas` entry to `difference_checker.spec` pointing to your local site-packages PySide6 plugins.

Installer

- After running the build script, create an installer with Inno Setup (recommended) or a ZIP fallback:

   ```powershell
   # produce an installer (pass the dist\difference_checker path from artifacts)
   .\tools\build_installer.ps1 -Source "artifacts\difference_checker_20250821_151923\difference_checker"
   ```

- The script looks for Inno Setup at the default `C:\Program Files (x86)\Inno Setup 6\ISCC.exe`. If not found,
   it will create a ZIP of the dist folder in `tools\installer_output\` instead.

Fixed ZIP for portable distribution

- A portable ZIP with a clear top-level folder is available at:

   `tools/installer_output/difference_checker_setup_fixed.zip`

- Coworker instructions (recommended):

   1. Download and extract the ZIP to a local folder (do not run directly from OneDrive or a network share).
       - Example extraction location: `%USERPROFILE%\Desktop\difference_checker_run`

   2. Run the application by double-clicking `difference_checker.exe` inside the extracted `difference_checker` folder.

   3. (Optional) Create Start Menu and Desktop shortcuts using the included helper script:

       ```powershell
       # from the project root (or the folder where you put the extracted files)
       .\tools\create_shortcut.ps1 -ExePath "C:\Users\you\Desktop\difference_checker_run\difference_checker\difference_checker.exe"
       ```

       The script creates a Start Menu shortcut for the current user and an optional Desktop shortcut.

Logs

- The app writes logs to a `logs/` folder in the project root when run from source. Packaged apps will write logs next to the running executable.

License

- Add your license here.

git

To create a git repo and push to GitHub (example):

```powershell
git init
git add .
git commit -m "Initial commit - Difference Checker"
# create remote repo on GitHub and replace <your-remote-url>
git remote add origin <your-remote-url>
git branch -M main
git push -u origin main
```
