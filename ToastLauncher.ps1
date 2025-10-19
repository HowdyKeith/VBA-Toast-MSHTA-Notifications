<#
.SYNOPSIS
ToastWatcherLauncher.ps1
Launches and monitors ToastWatcherRT.ps1 for changes, restarting it when updated.
Version: 5.3
#>

# --- CONFIGURATION ---
$ScriptPath = "C:\Users\howdy\OneDrive\Documents\2025\Powershell\ToastWatcherRT.ps1"
$LogDir    = Split-Path $ScriptPath -Parent
$LogFile   = Join-Path $LogDir "toast_launcher_log_$((Get-Date).ToString('yyyy-MM-dd')).log"

# --- LOGGING HELPER ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    try { Add-Content -Path $LogFile -Value $line -ErrorAction Stop } catch {}
    Write-Host $line
}

# --- GET SCRIPT LAST WRITE TIME ---
function Get-ScriptLastWriteTime {
    if (Test-Path $ScriptPath) { return (Get-Item $ScriptPath).LastWriteTime }
    return $null
}

# --- CHECK IF PROCESS IS RUNNING ---
function Get-ToastWatcherProcess {
    $proc = Get-Process -Name "powershell" -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -and (Resolve-Path $_.Path).Path -eq (Get-Item $ScriptPath).FullName }
    return $proc
}

# --- START LISTENER ---
function Start-ToastWatcher {
    Write-Log "Starting ToastWatcherRT.ps1..."
    $proc = Start-Process -FilePath "powershell.exe" `
                          -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" `
                          -PassThru
    Write-Log "ToastWatcher started (PID $($proc.Id))"
    return $proc
}

# --- STOP LISTENER ---
function Stop-ToastWatcher($proc) {
    if ($proc -and -not $proc.HasExited) {
        Write-Log "Stopping ToastWatcher (PID $($proc.Id))"
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
}

# --- INITIALIZE ---
Write-Log "ToastWatcherLauncher started. Monitoring $ScriptPath"
$lastWriteTime = Get-ScriptLastWriteTime

# Launch listener if not already running
$currentProcess = Get-ToastWatcherProcess
if (-not $currentProcess) {
    $currentProcess = Start-ToastWatcher
} else {
    Write-Log "ToastWatcher already running (PID $($currentProcess.Id))"
}

# --- FILESYSTEM WATCHER ---
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = Split-Path $ScriptPath -Parent
$watcher.Filter = Split-Path $ScriptPath -Leaf
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent $watcher "Changed" -Action {
    $data = $Event.MessageData
    $newWriteTime = (Get-Item $data.ScriptPath).LastWriteTime
    if ($newWriteTime -gt $data.LastWriteTime) {
        Write-Log "Detected change in $($data.ScriptPath)"
        Stop-ToastWatcher $data.CurrentProcess
        $newProcess = Start-ToastWatcher
        $data.CurrentProcess = $newProcess
        $data.LastWriteTime = $newWriteTime
    }
} -MessageData @{
    ScriptPath     = $ScriptPath
    CurrentProcess = $currentProcess
    LastWriteTime  = $lastWriteTime
}

# --- KEEP RUNNING ---
while ($true) { Start-Sleep -Seconds 1 }
