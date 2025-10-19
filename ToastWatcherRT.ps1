<#
.SYNOPSIS
ToastWatcherRT.ps1 v5.4
Real-time PowerShell listener for VBA/UI notifications via Named Pipe and Folder Watcher.
Supports WinRT toasts, fallback balloon tips, and JSON message parsing.

Author: Keith Swerling + ChatGPT (GPT-5)
#>

# ==============================
# CONFIGURATION
# ==============================
$Version = "5.4"
$PipeName = "ExcelToastPipe"
$TempFolder = "$env:TEMP\ToastQueue"
$LogFile = "$env:TEMP\ToastWatcherRT.log"
$UseWinRT = $true         # Try WinRT first
$EnableTempWatch = $true  # Fallback file-based watcher
$VerboseLog = $true

# ==============================
# LOGGING
# ==============================
Function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    if ($VerboseLog) { Write-Host $line }
    Add-Content -Path $LogFile -Value $line
}

# ==============================
# TOAST CORE
# ==============================
Function Show-WinRTToast {
    param([string]$Title, [string]$Message, [string]$Level, [int]$Progress = 0)

    try {
        Add-Type -AssemblyName Windows.Data, Windows.UI
        Add-Type -AssemblyName Windows.UI.Notifications
        Add-Type -AssemblyName Windows.Data.Xml.Dom

        $xml = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>$Title</text>
      <text>$Message</text>
    </binding>
  </visual>
</toast>
"@
        $xmlDoc = [Windows.Data.Xml.Dom.XmlDocument]::new()
        $xmlDoc.LoadXml($xml)

        $toast = [Windows.UI.Notifications.ToastNotification]::new($xmlDoc)
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Excel.ToastWatcherRT")
        $notifier.Show($toast)
        return $true
    }
    catch {
        Write-Log "[ERROR] WinRT Toast failed: $_" "ERROR"
        return $false
    }
}

Function Show-BalloonToast {
    param([string]$Title, [string]$Message)
    try {
        [reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.Visible = $true
        $notify.ShowBalloonTip(4000, $Title, $Message, [System.Windows.Forms.ToolTipIcon]::Info)
        Start-Sleep -Milliseconds 5000
        $notify.Dispose()
    }
    catch {
        Write-Log "[ERROR] Balloon toast failed: $_" "ERROR"
    }
}

# ==============================
# MESSAGE HANDLER
# ==============================
Function Handle-Toast {
    param([string]$Json)

    try {
        $data = $null
        $data = $Json | ConvertFrom-Json -ErrorAction Stop
        $Title = $data.Title
        $Message = $data.Message
        $Level = $data.Level
        $Progress = $data.Progress
        $Attr = $data.Attribution

        Write-Log "Toast: $Title | $Message | $Level | $Progress%"

        # Try WinRT first
        if ($UseWinRT -and (Show-WinRTToast $Title $Message $Level $Progress)) { return }

        # Fallback
        Show-BalloonToast $Title $Message
    }
    catch {
        Write-Log "[ERROR] Handle-Toast failed: $_" "ERROR"
    }
}

# ==============================
# PIPE SERVER
# ==============================
Function Start-PipeListener {
    Write-Log "Listening on named pipe: $PipeName ..."
    while ($true) {
        try {
            $pipe = New-Object System.IO.Pipes.NamedPipeServerStream($PipeName, [IO.Pipes.PipeDirection]::In)
            $pipe.WaitForConnection()
            $reader = New-Object System.IO.StreamReader($pipe)
            $msg = $reader.ReadToEnd()
            if ($msg -ne "") {
                Write-Log "Received raw: $msg"
                Handle-Toast $msg
            }
            $reader.Close()
            $pipe.Dispose()
        } catch {
            Write-Log "[ERROR] Pipe listener: $_" "ERROR"
            Start-Sleep -Seconds 1
        }
    }
}

# ==============================
# FOLDER WATCHER (FALLBACK)
# ==============================
Function Start-FolderWatcher {
    if (-not (Test-Path $TempFolder)) {
        New-Item -ItemType Directory -Force -Path $TempFolder | Out-Null
    }
    Write-Log "Watching folder: $TempFolder"

    $watcher = New-Object IO.FileSystemWatcher $TempFolder, "*.json"
    $watcher.EnableRaisingEvents = $true

    Register-ObjectEvent $watcher "Created" -Action {
        $path = $Event.SourceEventArgs.FullPath
        Start-Sleep -Milliseconds 200
        try {
            $json = Get-Content $path -Raw
            Write-Log "Detected new file: $path"
            Handle-Toast $json
            Remove-Item $path -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Log "[ERROR] FolderWatcher read failed: $_" "ERROR"
        }
    }
}

# ==============================
# MAIN
# ==============================
Write-Log "=== ToastWatcherRT.ps1 v$Version started ==="
Write-Log "Pipe: \\.\pipe\$PipeName"
Write-Log "Log file: $LogFile"

if ($EnableTempWatch) { Start-FolderWatcher }
Start-PipeListener
