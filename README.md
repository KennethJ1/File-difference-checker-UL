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

- Use `pyinstaller difference_checker.spec` to build an installer-ready folder (see `difference_checker.spec`).

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
