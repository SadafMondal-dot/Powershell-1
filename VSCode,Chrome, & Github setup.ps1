# GitHub & Software Provisioning Toolkit
# Author: Don with copilot
# Purpose: Automate installs, guide setup, log progress, and email results

$logPath = "$env:USERPROFILE\GitHubSetupLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

function Log-Step {
    param ([string]$Message, [string]$Status = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp [$Status] $Message"
    Add-Content -Path $logPath -Value $entry
    Write-Host $entry -ForegroundColor Gray
}

function Pause-Step {
    param ([string]$Message)
    Write-Host "`n--- $Message ---" -ForegroundColor Yellow
    Log-Step $Message
    Read-Host "Press Enter to continue"
}

function Install-Chocolatey {
    if (!(Get-Command choco -ErrorAction SilentlyContinue)) {
        Log-Step "Installing Chocolatey..."
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            Log-Step "Chocolatey installed successfully." "SUCCESS"
        } catch {
            Log-Step "Chocolatey installation failed: $_" "ERROR"
        }
    } else {
        Log-Step "Chocolatey already installed." "SUCCESS"
    }
}

function Configure-GitHub {
    Pause-Step "Navigate to GitHub and create an account using your @outlook.com email"
    Start-Process "https://github.com"
    Pause-Step "Create a new repository named 'PowerShell'"
    Pause-Step "Click 'Get started by creating a new file', name it 'readme.md', and type your name"
    Pause-Step "Scroll down and click 'Commit new file'"
}

function Install-Software {
    Pause-Step "Creating GitHubRepositories folder"
    try {
        New-Item -Path "$env:USERPROFILE\GitHubRepositories" -ItemType Directory -Force
        Log-Step "Created GitHubRepositories folder." "SUCCESS"
    } catch {
        Log-Step "Failed to create folder: $_" "ERROR"
    }

    Log-Step "Installing software using Chocolatey..."
    try {
        choco install googlechrome -y
        choco install git -y
        choco install vscode -y
        Log-Step "Software installed successfully." "SUCCESS"
    } catch {
        Log-Step "Software installation failed: $_" "ERROR"
    }

    Pause-Step "Launch Git Bash and run:"
    Write-Host 'git config --global user.name "Your Name"' -ForegroundColor Cyan
    Write-Host 'git config --global user.email "your_email@outlook.com"' -ForegroundColor Cyan

    Pause-Step "Launch Visual Studio Code and install extensions:"
    Write-Host "1. PowerShell extension by Microsoft" -ForegroundColor Cyan
    Write-Host "2. GitHub Pull Requests and Issues" -ForegroundColor Cyan
    Write-Host "3. Turn on Settings Sync and sign in with GitHub" -ForegroundColor Cyan
}

function Clone-Repository {
    Pause-Step "In VS Code, click Explorer (top left), then 'Clone Repository'"
    Pause-Step "Choose 'Clone from GitHub' and allow sign-in if prompted"
    Pause-Step "Select the 'GitHubRepositories' folder you created earlier"
    Pause-Step "Click 'Open' when prompted to open the cloned repository"
    Log-Step "Repository cloned successfully." "SUCCESS"
}

function Show-SummaryReport {
    Write-Host "`n=== Setup Summary ===" -ForegroundColor Cyan
    $logContent = Get-Content $logPath
    $successes = $logContent | Where-Object { $_ -match "\[SUCCESS\]" }
    $errors    = $logContent | Where-Object { $_ -match "\[ERROR\]" }
    $info      = $logContent | Where-Object { $_ -match "\[INFO\]" }

    Write-Host "`n✅ Successful Steps:" -ForegroundColor Green
    $successes | ForEach-Object { Write-Host $_ -ForegroundColor Green }

    Write-Host "`n⚠️ Errors Encountered:" -ForegroundColor Red
    if ($errors.Count -eq 0) {
        Write-Host "None 🎉" -ForegroundColor Green
    } else {
        $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    }

    Write-Host "`n📘 Informational Notes:" -ForegroundColor Yellow
    $info | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }

    Write-Host "`nLog file saved to: $logPath" -ForegroundColor Cyan
}

function Send-LogEmail {
    $send = Read-Host "Do you want to email this log file? (y/n)"
    if ($send -eq "y") {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0)
            $Mail.Subject = "GitHub Setup Log - $(Get-Date -Format 'yyyy-MM-dd')"
            $Mail.Body = "Hi,`n`nAttached is the setup log file from today's GitHub and software configuration.`n`nRegards,`nDon"
            $Mail.To = "recipient@example.com"  # Replace or prompt for email
            $Mail.Attachments.Add($logPath)
            $Mail.Display()  # Use .Send() to auto-send
            Log-Step "Log file opened in Outlook for review." "INFO"
        } catch {
            Log-Step "Failed to launch Outlook or create email: $_" "ERROR"
        }
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "=== GitHub & Software Setup ===" -ForegroundColor Cyan
    Write-Host "1. Install Chocolatey"
    Write-Host "2. Configure GitHub Account"
    Write-Host "3. Install Required Software"
    Write-Host "4. Clone GitHub Repository"
    Write-Host "5. Exit and Show Summary"
    return Read-Host "Select an option (1–5)"
}

# Main Loop
do {
    $choice = Show-Menu
    switch ($choice) {
        "1" { Install-Chocolatey }
        "2" { Configure-GitHub }
        "3" { Install-Software }
        "4" { Clone-Repository }
        "5" {
            Log-Step "User exited setup." "INFO"
            Show-SummaryReport
            Send-LogEmail
            Write-Host "`nExiting setup. Have a great day!" -ForegroundColor Magenta
        }
        default { Write-Host "`nInvalid selection. Please choose 1–5." -ForegroundColor Red }
    }
    if ($choice -ne "5") {
        Pause-Step "Return to main menu"
    }
} while ($choice -ne "5")