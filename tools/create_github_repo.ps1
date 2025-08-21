Param(
    [string]$RepoName = "difference-checker",
    [string]$RemoteUrl = ""
)

# Simple helper: initializes git, creates initial commit and optionally sets remote
# Requires git to be installed and available in PATH

if (-not (Test-Path -Path .git)) {
    git init
    git add .
    git commit -m "Initial commit - Difference Checker"
    if ($RemoteUrl) {
        git remote add origin $RemoteUrl
        git branch -M main
        git push -u origin main
    }
    Write-Host "Git repository initialized."
} else {
    Write-Host ".git already exists. Skipping init."
}
