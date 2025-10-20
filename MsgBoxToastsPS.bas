Attribute VB_Name = "MsgBoxToastsPS"
' MsgBoxToastsPS.bas v5.9.1
' Purpose: PowerShell Toast Bridge with Single-Use and VBS Fallback for Progress Updates
' Key Fix: Removed duplicate progressFile declaration in ShowToastPowerShell
' Dependencies: ToastWatcher.ps1 (for listener mode)

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Function tempPath() As String
    tempPath = Environ$("TEMP")
End Function

Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

Public Function ShowToastPowerShell( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal DurationSec As Long, _
    ByVal ToastType As String, _
    Optional ByVal LinkUrl As String = "", _
    Optional ByVal Icon As String = "", _
    Optional ByVal Sound As String = "", _
    Optional ByVal ImagePath As String = "", _
    Optional ByVal ImageSize As String = "Small", _
    Optional ByVal CallbackMacro As String = "", _
    Optional ByVal NoDismiss As Boolean = False, _
    Optional ByVal Position As String = "BR", _
    Optional ByVal Progress As Long = 0, _
    Optional ByVal SingleUse As Boolean = False, _
    Optional ByVal ProgressFile As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim tmp As String: tmp = tempPath()
    Dim requestFile As String: requestFile = tmp & "\ToastRequest.json"
    Dim lockFile As String: lockFile = tmp & "\ToastRequest.lock"
    Dim responseFile As String: responseFile = tmp & "\ToastResponse.txt"
    ' Use ProgressFile parameter or default to %TEMP%\ProgressRequest.json
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim ts As Object
    
    ' Clear old files
    On Error Resume Next
    If fso.FileExists(lockFile) Then fso.DeleteFile lockFile
    If fso.FileExists(requestFile) Then fso.DeleteFile requestFile
    If fso.FileExists(responseFile) Then fso.DeleteFile responseFile
    If Len(ProgressFile) > 0 And fso.FileExists(ProgressFile) Then fso.DeleteFile ProgressFile
    On Error GoTo ErrorHandler
    
    Sleep 50 ' Ensure file deletion
    
    ' Build JSON
    Dim tJson As String
    tJson = "{"
    tJson = tJson & """Title"":""" & EscapeJson(Title) & ""","
    tJson = tJson & """Message"":""" & EscapeJson(Message) & ""","
    tJson = tJson & """DurationSec"":" & DurationSec & ","
    tJson = tJson & """ToastType"":""" & UCase$(ToastType) & ""","
    tJson = tJson & """LinkUrl"":""" & EscapeJson(LinkUrl) & ""","
    tJson = tJson & """Icon"":""" & EscapeJson(Icon) & ""","
    tJson = tJson & """Sound"":""" & Sound & ""","
    tJson = tJson & """ImagePath"":""" & EscapeJson(ImagePath) & ""","
    tJson = tJson & """ImageSize"":""" & ImageSize & ""","
    tJson = tJson & """CallbackMacro"":""" & CallbackMacro & ""","
    tJson = tJson & """NoDismiss"":" & IIf(NoDismiss, "true", "false") & ","
    tJson = tJson & """Position"":""" & Position & ""","
    tJson = tJson & """Progress"":" & Progress & ","
    tJson = tJson & """ProgressFile"":""" & EscapeJson(IIf(Len(ProgressFile) > 0, ProgressFile, tmp & "\ProgressRequest.json")) & """"
    tJson = tJson & "}"
    
    ' Play beep
    If UCase$(Sound) = "BEEP" Then
        On Error Resume Next
        Beep 800, 200
        On Error GoTo ErrorHandler
    End If
    
    ' Single-use or no listener
    If SingleUse Or Not PowershellListenerRunning() Then
        If SingleUse Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Single-use mode enabled"
        Else
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Listener not running - falling back"
        End If
        
        If Progress > 0 Then
            ' VBS fallback for progress
            Dim vbsFile As String: vbsFile = tmp & "\ProgressToast.vbs"
            Dim htaFile As String: htaFile = tmp & "\ProgressToast.hta"
            Set ts = fso.CreateTextFile(vbsFile, True, True)
            ts.Write GetVBSCode(IIf(Len(ProgressFile) > 0, ProgressFile, tmp & "\ProgressRequest.json"), htaFile)
            ts.Close
            
            ' Initialize progress JSON
            Dim json As String
            json = "{""Progress"":" & Progress & ",""Message"":""" & EscapeJson(Message) & """,""Running"":true}"
            Set ts = fso.CreateTextFile(IIf(Len(ProgressFile) > 0, ProgressFile, tmp & "\ProgressRequest.json"), True, True)
            ts.Write json
            ts.Close
            
            On Error Resume Next
            shell.Run "wscript """ & vbsFile & """", 0, False
            If Err.Number <> 0 Then
                Debug.Print "[" & Format(Now, "hh:nn:ss") & "] ERROR: Failed to launch VBS - " & Err.Description
                ShowToastPowerShell = False
                Exit Function
            End If
            On Error GoTo ErrorHandler
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Launched VBS progress toast"
            ShowToastPowerShell = True
            Exit Function
        Else
            ' Single-use PowerShell
            Dim singleUsePs1 As String: singleUsePs1 = tmp & "\SingleUseToast.ps1"
            Dim psCode As String
            psCode = BuildSingleUsePs1Code(Title, Message, DurationSec, ToastType, LinkUrl, Icon, Sound, ImagePath, ImageSize, CallbackMacro, NoDismiss, Position, Progress)
            
            Set ts = fso.CreateTextFile(singleUsePs1, True, True)
            ts.Write psCode
            ts.Close
            
            Dim cmd As String
            cmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & singleUsePs1 & """"
            On Error Resume Next
            shell.Run cmd, 0, False
            If Err.Number <> 0 Then
                Debug.Print "[" & Format(Now, "hh:nn:ss") & "] ERROR: Failed to launch PowerShell - " & Err.Description
                ShowToastPowerShell = False
                Exit Function
            End If
            On Error GoTo ErrorHandler
            
            Sleep 1000
            If fso.FileExists(singleUsePs1) Then fso.DeleteFile singleUsePs1
            
            ShowToastPowerShell = True
            Exit Function
        End If
    End If
    
    ' Listener mode
    Set ts = fso.CreateTextFile(requestFile, True, True)
    ts.Write tJson
    ts.Close
    
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Wrote toast request: " & Title
    
    ' Wait for PowerShell to process
    Dim maxWait As Long: maxWait = 30 ' 3 seconds
    Dim i As Long
    For i = 1 To maxWait
        If Not fso.FileExists(requestFile) Then
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Toast processed by PowerShell"
            ShowToastPowerShell = True
            Exit Function
        End If
        Sleep 100
        DoEvents
    Next i
    
    ' Timeout
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] WARNING: Toast request timeout (listener not running?)"
    If fso.FileExists(responseFile) Then
        Set ts = fso.OpenTextFile(responseFile, 1)
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Listener error: " & ts.ReadAll
        ts.Close
    End If
    ShowToastPowerShell = False
    
    Exit Function

ErrorHandler:
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] ERROR: " & Err.Description
    ShowToastPowerShell = False
End Function

' Build single-use PowerShell script code for toast
Private Function BuildSingleUsePs1Code( _
    Title As String, Message As String, DurationSec As Long, ToastType As String, _
    LinkUrl As String, Icon As String, Sound As String, ImagePath As String, _
    ImageSize As String, CallbackMacro As String, NoDismiss As Boolean, _
    Position As String, Progress As Long) As String
    
    Dim psCode As String
    psCode = "<#" & vbCrLf
    psCode = psCode & ".SYNOPSIS" & vbCrLf
    psCode = psCode & "Single-use PowerShell script to display a toast notification." & vbCrLf
    psCode = psCode & "#>" & vbCrLf
    psCode = psCode & "$ErrorActionPreference = 'Stop'" & vbCrLf
    psCode = psCode & "$appId = 'Microsoft.Office.Excel'" & vbCrLf
    psCode = psCode & "try {" & vbCrLf
    psCode = psCode & "    if (Get-Module -ListAvailable -Name BurntToast) {" & vbCrLf
    psCode = psCode & "        Import-Module BurntToast" & vbCrLf
    psCode = psCode & "        $params = @{" & vbCrLf
    psCode = psCode & "            Text = @('" & Replace(Title, "'", "''") & "', '" & Replace(Message, "'", "''") & "')" & vbCrLf
    psCode = psCode & "            AppId = $appId" & vbCrLf
    If Len(ImagePath) > 0 Then
        psCode = psCode & "        if (Test-Path '" & Replace(ImagePath, "'", "''") & "') {" & vbCrLf
        psCode = psCode & "            if ('" & ImageSize & "' -eq 'Large') { $params['HeroImage'] = '" & Replace(ImagePath, "'", "''") & "' } else { $params['AppLogo'] = '" & Replace(ImagePath, "'", "''") & "' }" & vbCrLf
        psCode = psCode & "        }" & vbCrLf
    End If
    If Len(Icon) > 0 Then
        psCode = psCode & "        if (Test-Path '" & Replace(Icon, "'", "''") & "') { $params['AppLogo'] = '" & Replace(Icon, "'", "''") & "' }" & vbCrLf
    End If
    If Len(Sound) > 0 And Sound <> "BEEP" Then
        psCode = psCode & "        if ('" & Sound & "' -eq 'SystemAsterisk') { $params['Sound'] = 'IM' } else { $params['Sound'] = 'Default' }" & vbCrLf
    End If
    If Len(LinkUrl) > 0 Then
        psCode = psCode & "        $buttons = @()" & vbCrLf
        psCode = psCode & "        $buttons += New-BTButton -Content 'Open Link' -Arguments '" & Replace(LinkUrl, "'", "''") & "'" & vbCrLf
        psCode = psCode & "        $params['Button'] = $buttons" & vbCrLf
    End If
    If Len(CallbackMacro) > 0 Then
        psCode = psCode & "        if (-not $buttons) { $buttons = @() }" & vbCrLf
        psCode = psCode & "        $buttons += New-BTButton -Content 'Callback' -Arguments 'protocolHandler:" & Replace(CallbackMacro, "'", "''") & "'" & vbCrLf
        psCode = psCode & "        $params['Button'] = $buttons" & vbCrLf
    End If
    If Not NoDismiss Then
        psCode = psCode & "        if (-not $buttons) { $buttons = @() }" & vbCrLf
        psCode = psCode & "        $buttons += New-BTButton -Dismiss" & vbCrLf
        psCode = psCode & "        $params['Button'] = $buttons" & vbCrLf
    End If
    If DurationSec > 0 Then
        psCode = psCode & "        $params['ExpirationTime'] = (Get-Date).AddSeconds(" & DurationSec & ")" & vbCrLf
    End If
    psCode = psCode & "        New-BurntToastNotification @params" & vbCrLf
    psCode = psCode & "    } else {" & vbCrLf
    psCode = psCode & "        $bgColor = switch ('" & UCase$(ToastType) & "') { 'WARN' { 'linear-gradient(135deg, #ffeb3b, #ffa000)' } 'ERROR' { 'linear-gradient(135deg, #ff6b6b, #d32f2f)' } default { 'linear-gradient(135deg, #4caf50, #2e7d32)' } }" & vbCrLf
    psCode = psCode & "        $textColor = if ('" & UCase$(ToastType) & "' -eq 'WARN') { '#000000' } else { '#ffffff' }" & vbCrLf
    psCode = psCode & "        $iconChar = if ('" & Icon & "') { '" & Replace(Icon, "'", "''") & "' } else { switch ('" & UCase$(ToastType) & "') { 'WARN' { '?' } 'ERROR' { '?' } default { '?' } } }" & vbCrLf
    psCode = psCode & "        $screenW = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width)" & vbCrLf
    psCode = psCode & "        $screenH = ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height)" & vbCrLf
    psCode = psCode & "        $toastW = 350" & vbCrLf
    psCode = psCode & "        $toastH = $(if (" & Progress & " -gt 0) { 180 } else { 150 })" & vbCrLf
    psCode = psCode & "        $margin = 20" & vbCrLf
    psCode = psCode & "        switch ('" & UCase$(Position) & "') {" & vbCrLf
    psCode = psCode & "            'TL' { $x = $margin; $y = $margin }" & vbCrLf
    psCode = psCode & "            'TR' { $x = $screenW - $toastW - $margin; $y = $margin }" & vbCrLf
    psCode = psCode & "            'BL' { $x = $margin; $y = $screenH - $toastH - $margin }" & vbCrLf
    psCode = psCode & "            'BR' { $x = $screenW - $toastW - $margin; $y = $screenH - $toastH - $margin }" & vbCrLf
    psCode = psCode & "            'CR' { $x = $screenW - $toastW - $margin; $y = ($screenH - $toastH) / 2 }" & vbCrLf
    psCode = psCode & "            'C'  { $x = ($screenW - $toastW) / 2; $y = ($screenH - $toastH) / 2 }" & vbCrLf
    psCode = psCode & "            default { $x = $screenW - $toastW - $margin; $y = $screenH - $toastH - $margin }" & vbCrLf
    psCode = psCode & "        }" & vbCrLf
    If Progress > 0 Then
        psCode = psCode & "        $progressWidth = [math]::Max(0, [math]::Min(100, " & Progress & "))" & vbCrLf
        psCode = psCode & "        $progressHtml = '<div class=''progress-container''><div class=''progress-bar'' style=''width:' + $progressWidth + '%''></div></div>'" & vbCrLf
    Else
        psCode = psCode & "        $progressHtml = ''" & vbCrLf
    End If
    psCode = psCode & "        $hta = @\"" & vbCrLf"
    psCode = psCode & "<html><head><meta charset='UTF-8'><title>$($title -replace '\" ', '&quot;')</title>" & vbCrLf
    psCode = psCode & "<HTA:APPLICATION id='htaToast' border='none' showintaskbar='no' sysmenu='no' scroll='no' singleinstance='no'>" & vbCrLf
    psCode = psCode & "<style>" & vbCrLf
    psCode = psCode & "body {margin:0;padding:10px;font-family:'Segoe UI',Arial,sans-serif;background:$bgColor;color:$textColor;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);animation:slideIn 0.5s ease-in;}" & vbCrLf
    psCode = psCode & "@keyframes slideIn {from {transform:translateX(100%);} to {transform:translateX(0);}}" & vbCrLf
    psCode = psCode & "@keyframes slideOut {from {transform:translateX(0);} to {transform:translateX(100%);}}" & vbCrLf
    psCode = psCode & "h3 {margin:0 0 8px;font-size:18px;font-weight:600;}" & vbCrLf
    psCode = psCode & ".icon {font-size:24px;margin-right:8px;}" & vbCrLf
    psCode = psCode & "p {margin:0 0 10px;font-size:14px;}" & vbCrLf
    psCode = psCode & ".progress-container {width:100%;background:#ccc;border-radius:4px;}" & vbCrLf
    psCode = psCode & ".progress-bar {height:20px;background:#2196f3;border-radius:4px;}" & vbCrLf
    psCode = psCode & "button {position:absolute;top:5px;right:5px;padding:4px 8px;border:none;border-radius:4px;background:rgba(0,0,0,0.3);color:$textColor;cursor:pointer;}" & vbCrLf
    psCode = psCode & "</style>" & vbCrLf
    psCode = psCode & "<script language='VBScript'>" & vbCrLf
    psCode = psCode & "window.resizeTo " & (350 + 20) & "," & IIf(Progress > 0, 180 + 20, 150 + 20) & vbCrLf
    psCode = psCode & "window.moveTo $x,$y" & vbCrLf
    If Len(LinkUrl) > 0 Then
        psCode = psCode & "Sub btnLink_OnClick" & vbCrLf
        psCode = psCode & "  Dim sh" & vbCrLf
        psCode = psCode & "  Set sh = CreateObject(""WScript.Shell"")" & vbCrLf
        psCode = psCode & "  sh.Run """ & Replace(LinkUrl, """", """""") & """, 1, False" & vbCrLf
        psCode = psCode & "End Sub" & vbCrLf
    End If
    If Len(CallbackMacro) > 0 Then
        psCode = psCode & "Sub btnCallback_OnClick" & vbCrLf
        psCode = psCode & "  MsgBox ""Callback triggered: " & Replace(CallbackMacro, """", """""") & """, 64, ""Callback""" & vbCrLf
        psCode = psCode & "End Sub" & vbCrLf
    End If
    If DurationSec > 0 Then
        psCode = psCode & "setTimeout ""document.body.style.animation='slideOut 0.5s ease-out':setTimeout """"window.close"""",500"", " & (DurationSec * 1000) & vbCrLf
    End If
    psCode = psCode & "</script></head>" & vbCrLf
    psCode = psCode & "<body>" & vbCrLf
    psCode = psCode & "<h3><span class='icon'>$iconChar</span>$($title -replace '\" ', '&quot;')</h3>" & vbCrLf
    psCode = psCode & "<p>$($message -replace '\" ', '&quot;')</p>" & vbCrLf
    psCode = psCode & "$progressHtml" & vbCrLf
    If Len(LinkUrl) > 0 Then
        psCode = psCode & "<button onclick='btnLink_OnClick'>Open Link</button>" & vbCrLf
    End If
    If Len(CallbackMacro) > 0 Then
        psCode = psCode & "<button onclick='btnCallback_OnClick'>Callback</button>" & vbCrLf
    End If
    If Not NoDismiss Or Len(LinkUrl) > 0 Or Len(CallbackMacro) > 0 Then
        psCode = psCode & "<button style='top:5px;right:5px' onclick='document.body.style.animation=""slideOut 0.5s ease-out"";setTimeout ""window.close"",500'>×</button>" & vbCrLf
    End If
    psCode = psCode & "</body></html>" & vbCrLf
    psCode = psCode & "@""" & vbCrLf
    psCode = psCode & "        $htaPath = [System.IO.Path]::Combine($env:TEMP, 'Toast_$([System.Guid]::NewGuid()).hta')" & vbCrLf
    psCode = psCode & "        Set-Content -Path $htaPath -Value $hta -Encoding UTF8" & vbCrLf
psCode = psCode & "        Start-Process 'mshta.exe' -ArgumentList ""'"" + $htaPath + ""'"" -WindowStyle Hidden" & vbCrLf

    psCode = psCode & "        Start-Sleep -Milliseconds 500" & vbCrLf
    psCode = psCode & "        Remove-Item -Path $htaPath -Force" & vbCrLf
    psCode = psCode & "    }" & vbCrLf
    psCode = psCode & "} catch {" & vbCrLf
    psCode = psCode & "    $responsePath = [System.IO.Path]::Combine($env:TEMP, 'ToastResponse.txt')" & vbCrLf
    psCode = psCode & "    Set-Content -Path $responsePath -Value $_.Exception.Message" & vbCrLf
    psCode = psCode & "}" & vbCrLf
    
    BuildSingleUsePs1Code = psCode
End Function

' Generate VBS code for progress toasts
Private Function GetVBSCode(ProgressFile As String, htaFile As String) As String
    Dim vbsCode As String
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim fso, shell" & vbCrLf
    vbsCode = vbsCode & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf
    vbsCode = vbsCode & "RunDynamicProgressToast ""Progress Demo"", ""Starting process..."", ""INFO"", """", """", """", 0, ""C"", 0, """ & Replace(ProgressFile, """", """""") & """, """ & Replace(htaFile, """", """""") & """" & vbCrLf
    vbsCode = vbsCode & GetRunDynamicProgressToastCode()
    GetVBSCode = vbsCode
End Function

Private Function GetRunDynamicProgressToastCode() As String
    Dim code As String
    code = "Sub RunDynamicProgressToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, DurationSec, Position, Progress, ProgressFile, HTAFile)" & vbCrLf
    code = code & "  Dim fso" & vbCrLf
    code = code & "  Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "  Dim bgColor, textColor, iconChar" & vbCrLf
    code = code & "  Select Case UCase(ToastType)" & vbCrLf
    code = code & "    Case ""WARN"": bgColor = ""linear-gradient(135deg, #ffeb3b, #ffa000)"": textColor = ""#000000"": iconChar = ""?""" & vbCrLf
    code = code & "    Case ""ERROR"": bgColor = ""linear-gradient(135deg, #ff6b6b, #d32f2f)"": textColor = ""#ffffff"": iconChar = ""?""" & vbCrLf
    code = code & "    Case Else: bgColor = ""linear-gradient(135deg, #4caf50, #2e7d32)"": textColor = ""#ffffff"": iconChar = ""?""" & vbCrLf
    code = code & "  End Select" & vbCrLf
    code = code & "  Dim screenW, screenH, toastW, toastH, margin, x, y" & vbCrLf
    code = code & "  screenW = CreateObject(""Shell.Application"").Namespace(0).Self.InvokeVerb(""Display"").Width" & vbCrLf
    code = code & "  screenH = CreateObject(""Shell.Application"").Namespace(0).Self.InvokeVerb(""Display"").Height" & vbCrLf
    code = code & "  toastW = 350: toastH = 180: margin = 20" & vbCrLf
    code = code & "  Select Case UCase(Position)" & vbCrLf
    code = code & "    Case ""TL"": x = margin: y = margin" & vbCrLf
    code = code & "    Case ""TR"": x = screenW - toastW - margin: y = margin" & vbCrLf
    code = code & "    Case ""BL"": x = margin: y = screenH - toastH - margin" & vbCrLf
    code = code & "    Case ""BR"": x = screenW - toastW - margin: y = screenH - toastH - margin" & vbCrLf
    code = code & "    Case ""CR"": x = screenW - toastW - margin: y = (screenH - toastH) / 2" & vbCrLf
    code = code & "    Case ""C"": x = (screenW - toastW) / 2: y = (screenH - toastH) / 2" & vbCrLf
    code = code & "    Case Else: x = screenW - toastW - margin: y = screenH - toastH - margin" & vbCrLf
    code = code & "  End Select" & vbCrLf
    code = code & "  Dim html" & vbCrLf
    code = code & "  html = ""<html><head><meta charset='UTF-8'><title>"" & Replace(Title, """""", ""&quot;"") & ""</title>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<HTA:APPLICATION id='htaToast' border='none' showintaskbar='no' sysmenu='no' scroll='no' singleinstance='no'>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<style>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""body {margin:0;padding:10px;font-family:'Segoe UI',Arial,sans-serif;background:"" & bgColor & "";color:"" & textColor & "";border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);animation:slideIn 0.5s ease-in;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""@keyframes slideIn {from {transform:translateX(100%);} to {transform:translateX(0);}}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""@keyframes slideOut {from {transform:translateX(0);} to {transform:translateX(100%);}}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""h3 {margin:0 0 8px;font-size:18px;font-weight:600;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & "".icon {font-size:24px;margin-right:8px;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""p {margin:0 0 10px;font-size:14px;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & "".progress-container {width:100%;background:#ccc;border-radius:4px;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & "".progress-bar {height:20px;background:#2196f3;border-radius:4px;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""button {position:absolute;top:5px;right:5px;padding:4px 8px;border:none;border-radius:4px;background:rgba(0,0,0,0.3);color:"" & textColor & "";cursor:pointer;}"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""</style>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<script language='VBScript'>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""window.resizeTo "" & (toastW + 20) & "","" & (toastH + 20) & vbCrLf" & vbCrLf
    code = code & "  html = html & ""window.moveTo "" & x & "","" & y & vbCrLf" & vbCrLf
    code = code & "  html = html & ""Sub UpdateProgress"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""  On Error Resume Next"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""  Dim fso, ts, json, progress, message, running"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""  Set fso = CreateObject(""""Scripting.FileSystemObject"""")"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""  If fso.FileExists("""" & ProgressFile & """") Then"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    Set ts = fso.OpenTextFile("""" & ProgressFile & """", 1)"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    json = ts.ReadAll"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    ts.Close"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    progress = CInt(Mid(json, InStr(json, """"Progress"""":) + 10, InStr(json, "",""""Message"""") - InStr(json, """"Progress"""":) - 10))"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    message = Mid(json, InStr(json, """"Message"""":""""""") + 10, InStr(json, """""Running"""") - InStr(json, """"Message"""":""""""") - 11)"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    running = (InStr(json, """"Running"""":true) > 0)"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    If running Then"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""      document.getElementById(""""progressBar"""").style.width = progress & """"%"""""" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""      document.getElementById(""""msg"""").innerText = message"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""      setTimeout """"UpdateProgress"""", 500"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    Else"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""      document.body.style.animation = """"slideOut 0.5s ease-out"""""" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""      setTimeout """"window.close"""", 500"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    End If"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""  Else"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    document.body.style.animation = """"slideOut 0.5s ease-out"""""" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""    setTimeout """"window.close"""", 500"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""End Sub"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""setTimeout """"UpdateProgress"""", 500"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""</script></head>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<body>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<h3><span class='icon'>"" & iconChar & ""</span>"" & Replace(Title, """""", ""&quot;"") & ""</h3>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<p id='msg'>"" & Replace(Message, """""", ""&quot;"") & ""</p>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<div class='progress-container'><div id='progressBar' class='progress-bar' style='width:"" & Progress & ""%'></div></div>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""<button onclick='document.body.style.animation=""""slideOut 0.5s ease-out"""":setTimeout """"window.close"""",500'>×</button>"" & vbCrLf" & vbCrLf
    code = code & "  html = html & ""</body></html>"" & vbCrLf" & vbCrLf
    code = code & "  Dim ts" & vbCrLf
    code = code & "  Set ts = fso.CreateTextFile(HTAFile, True)" & vbCrLf
    code = code & "  ts.Write html" & vbCrLf
    code = code & "  ts.Close" & vbCrLf
    code = code & "  CreateObject(""WScript.Shell"").Run ""mshta.exe """" & HTAFile & """""", 0, False" & vbCrLf
    code = code & "End Sub" & vbCrLf
    GetRunDynamicProgressToastCode = code
End Function

' Demo for dynamic progress updates
Public Sub ProgressToastDemo()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim tmp As String: tmp = tempPath()
    Dim ProgressFile As String: ProgressFile = tmp & "\ProgressRequest.json"
    Dim vbsFile As String: vbsFile = tmp & "\ProgressToast.vbs"
    Dim htaFile As String: htaFile = tmp & "\ProgressToast.hta"
    Dim singleUsePs1 As String: singleUsePs1 = tmp & "\SingleUseToast.ps1"
    
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Starting progress demo"
    
    ' Try listener mode
    If PowershellListenerRunning() Then
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Using PowerShell listener"
        Dim i As Long
        For i = 0 To 100 Step 10
            If Not ShowToastPowerShell("Progress Demo", "Processing: " & i & "%", 0, "INFO", , , , , , , , "C", i, False, ProgressFile) Then
                Debug.Print "[" & Format(Now, "hh:nn:ss") & "] ERROR: Listener failed to process toast"
                Exit Sub
            End If
            Dim json As String
            json = "{""Progress"":" & i & ",""Message"":""" & EscapeJson("Processing: " & i & "%") & """,""Running"":true}"
            Set ts = fso.CreateTextFile(ProgressFile, True, True)
            ts.Write json
            ts.Close
            Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Updated progress to " & i & "%"
            Application.Wait Now + TimeValue("00:00:01")
        Next i
        json = "{""Progress"":100,""Message"":""" & EscapeJson("Processing complete!") & """,""Running"":false}"
        Set ts = fso.CreateTextFile(ProgressFile, True, True)
        ts.Write json
        ts.Close
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Progress complete (listener mode)"
        Application.Wait Now + TimeValue("00:00:02")
        If fso.FileExists(ProgressFile) Then fso.DeleteFile ProgressFile
        Exit Sub
    End If
    
    ' Fallback to VBS
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Listener not running - using VBS"
    Set ts = fso.CreateTextFile(vbsFile, True, True)
    ts.Write GetVBSCode(ProgressFile, htaFile)
    ts.Close
    
    On Error Resume Next
    shell.Run "wscript """ & vbsFile & """", 0, False
    If Err.Number <> 0 Then
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] ERROR: Failed to launch VBS - " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Launched VBS progress toast"
    
    ' Simulate progress
    Dim j As Long
    For j = 0 To 100 Step 10
        Dim json2 As String
        json2 = "{""Progress"":" & j & ",""Message"":""" & EscapeJson("Processing: " & j & "%") & """,""Running"":true}"
        Set ts = fso.CreateTextFile(ProgressFile, True, True)
        ts.Write json2
        ts.Close
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Updated progress to " & j & "%"
        Application.Wait Now + TimeValue("00:00:01")
    Next j
    
    ' Signal completion
    json2 = "{""Progress"":100,""Message"":""" & EscapeJson("Processing complete!") & """,""Running"":false}"
    Set ts = fso.CreateTextFile(ProgressFile, True, True)
    ts.Write json2
    ts.Close
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Progress complete (VBS mode)"
    
    ' Cleanup
    Application.Wait Now + TimeValue("00:00:02")
    If fso.FileExists(vbsFile) Then fso.DeleteFile vbsFile
    If fso.FileExists(ProgressFile) Then fso.DeleteFile ProgressFile
    If fso.FileExists(htaFile) Then fso.DeleteFile htaFile
End Sub

' Check if PowerShell listener is running
Private Function PowershellListenerRunning() As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String: tmp = tempPath()
    Dim sentinel As String: sentinel = tmp & "\ToastLastProcessed.txt"
    
    If fso.FileExists(sentinel) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(sentinel, 1)
        Dim lastTime As String: lastTime = ts.ReadLine
        ts.Close
        Dim diff As Double
        diff = DateDiff("s", CDate(lastTime), Now)
        PowershellListenerRunning = (diff < 10)
    Else
        PowershellListenerRunning = False
    End If
    On Error GoTo 0
End Function

' Test menu for all toast features
Public Sub TestNotificationsPS()
    Dim choice As String
    Dim Prompt As String
    
    Prompt = "===== PowerShell Toast Test Menu =====" & vbCrLf & vbCrLf
    Prompt = Prompt & "BASIC TESTS:" & vbCrLf
    Prompt = Prompt & " 1 - Simple Info Toast (BR)" & vbCrLf
    Prompt = Prompt & " 2 - Warning Toast (TR)" & vbCrLf
    Prompt = Prompt & " 3 - Error Toast (TL)" & vbCrLf & vbCrLf
    Prompt = Prompt & "POSITION TESTS:" & vbCrLf
    Prompt = Prompt & " 4 - Bottom-Left (BL)" & vbCrLf
    Prompt = Prompt & " 5 - Center-Right (CR)" & vbCrLf
    Prompt = Prompt & " 6 - Center (C)" & vbCrLf
    Prompt = Prompt & " 7 - Test All 6 Positions" & vbCrLf & vbCrLf
    Prompt = Prompt & "ADVANCED FEATURES:" & vbCrLf
    Prompt = Prompt & " 8 - Toast with Link" & vbCrLf
    Prompt = Prompt & " 9 - Toast with Callback" & vbCrLf
    Prompt = Prompt & "10 - Toast with Image (Small)" & vbCrLf
    Prompt = Prompt & "11 - Toast with Image (Large)" & vbCrLf
    Prompt = Prompt & "12 - No Dismiss Button" & vbCrLf
    Prompt = Prompt & "13 - Long Duration (10s)" & vbCrLf
    Prompt = Prompt & "14 - Persistent (No Auto-Close)" & vbCrLf
    Prompt = Prompt & "15 - Progress Toast (50%)" & vbCrLf
    Prompt = Prompt & "16 - Dynamic Progress Demo" & vbCrLf
    Prompt = Prompt & "17 - Single-Use Launch (No Listener)" & vbCrLf & vbCrLf
    Prompt = Prompt & "SYSTEM:" & vbCrLf
    Prompt = Prompt & "18 - Check Listener Status" & vbCrLf
    Prompt = Prompt & "19 - Multiple Toasts Demo" & vbCrLf
    Prompt = Prompt & "20 - Cleanup Temp Files" & vbCrLf & vbCrLf
    Prompt = Prompt & " 0 - Exit"
    
    Do
        choice = InputBox(Prompt, "PowerShell Toast Tests", "1")
        
        If choice = "" Or choice = "0" Then Exit Sub
        
        Select Case val(choice)
            Case 1
                If ShowToastPowerShell("Info Toast", "This is a simple info toast.", 3, "INFO", , , "BEEP", , , , , "BR") Then
                    Debug.Print "? Info toast sent"
                Else
                    MsgBox "Failed - is listener running?", vbExclamation
                End If
                
            Case 2
                ShowToastPowerShell "Warning", "This is a warning toast with sound!", 5, "WARN", , , "BEEP", , , , , "TR"
                
            Case 3
                ShowToastPowerShell "Error", "This is an error toast!", 5, "ERROR", , , "BEEP", , , , , "TL"
                
            Case 4
                ShowToastPowerShell "Bottom-Left", "Toast at bottom-left corner.", 3, "INFO", , , , , , , , "BL"
                
            Case 5
                ShowToastPowerShell "Center-Right", "Toast at center-right.", 3, "INFO", , , , , , , , "CR"
                
            Case 6
                ShowToastPowerShell "Center", "Toast at screen center.", 3, "INFO", , , , , , , , "C"
                
            Case 7
                TestAllPositions
                
            Case 8
                ShowToastPowerShell "Link Toast", "Click the link to open Microsoft.", 0, "INFO", "https://www.microsoft.com", , "BEEP", , , , , "BR"
                
            Case 9
                ShowToastPowerShell "Callback Toast", "Click to trigger callback macro.", 0, "INFO", , , "BEEP", , , "OnToastClicked", , "BR"
                
            Case 10
                Dim imgPath As String
                imgPath = InputBox("Enter image path (or leave blank to skip):", "Image Path", "C:\Windows\Web\Wallpaper\Windows\img0.jpg")
                If Len(imgPath) > 0 Then
                    ShowToastPowerShell "Small Image", "Toast with small image.", 5, "INFO", , , , imgPath, "Small", , , "BR"
                End If
                
            Case 11
                imgPath = InputBox("Enter image path (or leave blank to skip):", "Image Path", "C:\Windows\Web\Wallpaper\Windows\img0.jpg")
                If Len(imgPath) > 0 Then
                    ShowToastPowerShell "Large Image", "Toast with large hero image.", 5, "INFO", , , , imgPath, "Large", , , "BR"
                End If
                
            Case 12
                ShowToastPowerShell "No Dismiss", "This toast has no dismiss button.", 5, "INFO", , , "BEEP", , , , True, "BR"
                
            Case 13
                ShowToastPowerShell "Long Duration", "This toast will stay for 10 seconds.", 10, "INFO", , , , , , , , "BR"
                
            Case 14
                ShowToastPowerShell "Persistent", "This toast won't auto-close. Click X to dismiss.", 0, "INFO", , , , , , , , "BR"
                
            Case 15
                ShowToastPowerShell "Progress Toast", "Processing task (50%)...", 8, "INFO", , , "BEEP", , , , , "BR", 50
                
            Case 16
                ProgressToastDemo
                
            Case 17
                ShowToastPowerShell "Single-Use", "This toast uses single-use launch.", 5, "INFO", , , "BEEP", , , , , "BR", 0, True
                
            Case 18
                CheckListener
                
            Case 19
                MultipleToastsDemo
                
            Case 20
                CleanupTempFiles
                
            Case Else
                MsgBox "Invalid choice (0-20)", vbExclamation
        End Select
    Loop
End Sub

Private Sub MultipleToastsDemo()
    Dim i As Long
    Dim positions As Variant
    positions = Array("TL", "TR", "BL", "BR")
    
    MsgBox "Will display 4 toasts in different corners.", vbInformation, "Multiple Toasts"
    
    For i = 0 To 3
        ShowToastPowerShell "Toast #" & (i + 1), "This is toast " & (i + 1) & " at " & positions(i), 3, "INFO", , , , , , , , CStr(positions(i))
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    MsgBox "Demo complete!", vbInformation
End Sub

Private Sub CleanupTempFiles()
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String: tmp = tempPath()
    Dim files As Variant
    files = Array("ToastRequest.json", "ToastRequest.lock", "ToastResponse.txt", "ToastLastProcessed.txt", "ProgressRequest.json", "ProgressToast.vbs", "ProgressToast.hta", "SingleUseToast.ps1")
    
    Dim f As Variant
    Dim count As Long
    For Each f In files
        If fso.FileExists(tmp & "\" & f) Then
            fso.DeleteFile tmp & "\" & f
            count = count + 1
        End If
    Next f
    
    MsgBox "Cleaned up " & count & " temp files.", vbInformation, "Cleanup Complete"
End Sub

Private Sub TestAllPositions()
    Dim positions As Variant
    Dim i As Long
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    
    For i = LBound(positions) To UBound(positions)
        If ShowToastPowerShell("Position Test", "Testing: " & positions(i), 2, "INFO", , , , , , , , CStr(positions(i))) Then
            Debug.Print "Position " & positions(i) & " - OK"
        Else
            Debug.Print "Position " & positions(i) & " - FAILED"
        End If
        Application.Wait Now + TimeValue("00:00:02")
    Next i
    
    MsgBox "Position test complete!", vbInformation
End Sub

Private Sub CheckListener()
    Dim running As Boolean
    running = PowershellListenerRunning()
    
    If running Then
        MsgBox "? PowerShell listener is RUNNING" & vbCrLf & vbCrLf & _
               "Sentinel file updated < 10 seconds ago", vbInformation, "Listener Status"
    Else
        MsgBox "? PowerShell listener is NOT running" & vbCrLf & vbCrLf & _
               "Start it with:" & vbCrLf & _
               "powershell -File ToastWatcher.ps1", vbExclamation, "Listener Status"
    End If
End Sub

Public Sub OnToastClicked()
    MsgBox "Callback executed!", vbInformation, "Toast Callback"
End Sub

Private Sub Beep(Frequency As Long, Duration As Long)
    On Error Resume Next
    CreateObject("WScript.Shell").Run "powershell -Command [Console]::Beep(" & Frequency & "," & Duration & ")", 0, True
End Sub


