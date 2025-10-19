'========================================================
' MsgBoxToastTester.vbs v2.0 (FIXED)
' Universal Toast Tester (PS + MSHTA + WScript)
' 
' FIXES:
' - Added position support (TL, TR, BL, BR, CR, C)
' - Fixed PowerShell native toast (no BurntToast dependency)
' - Fixed MSHTA toast with proper animations
' - Improved callback handling
' - Better error handling
'========================================================

Option Explicit

'================= GLOBALS =================
Dim UsePowerShellToasts, DefaultPosition
UsePowerShellToasts = False ' Default to MSHTA (more reliable)
DefaultPosition = "BR"

'================= MAIN MENU =================
Sub MainMenu()
    Dim choice
    Do
        choice = InputBox( _
            "===== Universal Toast Tester v2.0 =====" & vbCrLf & vbCrLf & _
            "QUICK TESTS:" & vbCrLf & _
            "1 - Info Toast (BR)" & vbCrLf & _
            "2 - Warning Toast (TR)" & vbCrLf & _
            "3 - Error Toast (TL)" & vbCrLf & vbCrLf & _
            "ADVANCED:" & vbCrLf & _
            "4 - Toast with Link" & vbCrLf & _
            "5 - Persistent Toast (No Auto-Close)" & vbCrLf & _
            "6 - Queue Multiple Toasts" & vbCrLf & vbCrLf & _
            "TESTS:" & vbCrLf & _
            "7 - Test All Positions" & vbCrLf & _
            "8 - Test All Types (Info/Warn/Error)" & vbCrLf & vbCrLf & _
            "SETTINGS:" & vbCrLf & _
            "9 - Toggle PS/MSHTA Mode (Current: " & GetMode() & ")" & vbCrLf & _
            "10 - System Check" & vbCrLf & vbCrLf & _
            "0 - Exit", _
            "Universal Toast Tester", "1")
        
        If choice = "" Or choice = "0" Then Exit Sub
        
        Select Case CInt(choice)
            Case 1
                ShowToast "Info", "This is an informational toast.", "INFO", "", "", "", "", 5, "BR"
            Case 2
                ShowToast "Warning", "This is a warning toast with sound!", "WARN", "", "", "", "BEEP", 7, "TR"
            Case 3
                ShowToast "Error", "This is an error toast!", "ERROR", "", "", "", "BEEP", 8, "TL"
            Case 4
                ShowToast "Link Toast", "Click the link to open Microsoft.", "INFO", "https://www.microsoft.com", "", "", "", 0, "CR"
            Case 5
                ShowToast "Persistent", "This toast won't auto-close. Click X to dismiss.", "INFO", "", "", "", "", 0, "BR"
            Case 6
                QueueToasts
            Case 7
                PositionTest
            Case 8
                TypeTest
            Case 9
                ToggleMode
            Case 10
                SystemCheck
            Case Else
                MsgBox "Invalid choice (0-10).", vbExclamation, "Error"
        End Select
    Loop
End Sub

'================= SHOW TOAST =================
Sub ShowToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position)
    ' Validate parameters
    If IsEmpty(Duration) Or Duration = "" Then Duration = 5
    If IsEmpty(Position) Or Position = "" Then Position = DefaultPosition
    If IsEmpty(ToastType) Or ToastType = "" Then ToastType = "INFO"
    
    ' Choose display method
    If UsePowerShellToasts And PowerShellAvailable() Then
        RunPSToast Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position
    ElseIf MSHTAAvailable() Then
        RunMSHTAToast Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position
    Else
        ' Fallback to classic MsgBox
        MsgBox Message, vbOKOnly, Title
    End If
End Sub

'================= CHECK PS AVAILABILITY =================
Function PowerShellAvailable()
    On Error Resume Next
    Dim shell, result
    Set shell = CreateObject("WScript.Shell")
    result = shell.Run("powershell -Command ""exit 0""", 0, True)
    PowerShellAvailable = (Err.Number = 0 And result = 0)
    Err.Clear
End Function

'================= CHECK MSHTA AVAILABILITY =================
Function MSHTAAvailable()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    MSHTAAvailable = fso.FileExists(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%WINDIR%") & "\System32\mshta.exe")
End Function

'================= RUN PS TOAST =================
Sub RunPSToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position)
    Dim tempPS, fso, shell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    tempPS = fso.GetSpecialFolder(2) & "\TempToastVBS_" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".ps1"
    
    ' Build PowerShell script
    Dim psScript
    psScript = BuildPSScript(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position)
    
    ' Write script
    Dim ts
    Set ts = fso.CreateTextFile(tempPS, True)
    ts.Write psScript
    ts.Close
    
    ' Execute
    shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & tempPS & """", 0, False
End Sub

'================= BUILD PS SCRIPT =================
Function BuildPSScript(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position)
    Dim ps
    ps = "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf
    ps = ps & "Add-Type -AssemblyName System.Drawing" & vbCrLf & vbCrLf
    
    ' Colors
    Dim bgColor, fgColor, icon
    Select Case UCase(ToastType)
        Case "WARN"
            bgColor = "255,235,59"
            fgColor = "0,0,0"
            icon = "⚠"
        Case "ERROR"
            bgColor = "244,67,54"
            fgColor = "255,255,255"
            icon = "❌"
        Case Else
            bgColor = "76,175,80"
            fgColor = "255,255,255"
            icon = "ℹ"
    End Select
    
    ' Create form
    ps = ps & "$form = New-Object System.Windows.Forms.Form" & vbCrLf
    ps = ps & "$form.Text = '" & PSEscape(Title) & "'" & vbCrLf
    ps = ps & "$form.Width = 350" & vbCrLf
    ps = ps & "$form.Height = 150" & vbCrLf
    ps = ps & "$form.FormBorderStyle = 'None'" & vbCrLf
    ps = ps & "$form.StartPosition = 'Manual'" & vbCrLf
    ps = ps & "$form.TopMost = $true" & vbCrLf
    ps = ps & "$form.BackColor = [System.Drawing.Color]::FromArgb(" & bgColor & ")" & vbCrLf
    ps = ps & "$form.Opacity = 0" & vbCrLf & vbCrLf
    
    ' Position
    ps = ps & CalculatePositionPS(Position)
    
    ' Title
    ps = ps & "$lblTitle = New-Object System.Windows.Forms.Label" & vbCrLf
    ps = ps & "$lblTitle.Text = '" & PSEscape(icon & " " & Title) & "'" & vbCrLf
    ps = ps & "$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)" & vbCrLf
    ps = ps & "$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(" & fgColor & ")" & vbCrLf
    ps = ps & "$lblTitle.AutoSize = $true" & vbCrLf
    ps = ps & "$lblTitle.Location = New-Object System.Drawing.Point(10,10)" & vbCrLf
    ps = ps & "$form.Controls.Add($lblTitle)" & vbCrLf & vbCrLf
    
    ' Message
    ps = ps & "$lblMsg = New-Object System.Windows.Forms.Label" & vbCrLf
    ps = ps & "$lblMsg.Text = '" & PSEscape(Message) & "'" & vbCrLf
    ps = ps & "$lblMsg.Font = New-Object System.Drawing.Font('Segoe UI',10)" & vbCrLf
    ps = ps & "$lblMsg.ForeColor = [System.Drawing.Color]::FromArgb(" & fgColor & ")" & vbCrLf
    ps = ps & "$lblMsg.AutoSize = $false" & vbCrLf
    ps = ps & "$lblMsg.Width = 320" & vbCrLf
    ps = ps & "$lblMsg.Height = 60" & vbCrLf
    ps = ps & "$lblMsg.Location = New-Object System.Drawing.Point(10,40)" & vbCrLf
    ps = ps & "$form.Controls.Add($lblMsg)" & vbCrLf & vbCrLf
    
    ' Close button
    ps = ps & "$btnClose = New-Object System.Windows.Forms.Button" & vbCrLf
    ps = ps & "$btnClose.Text = 'X'" & vbCrLf
    ps = ps & "$btnClose.Width = 25" & vbCrLf
    ps = ps & "$btnClose.Height = 25" & vbCrLf
    ps = ps & "$btnClose.Location = New-Object System.Drawing.Point(315,5)" & vbCrLf
    ps = ps & "$btnClose.FlatStyle = 'Flat'" & vbCrLf
    ps = ps & "$btnClose.Add_Click({ $form.Close() })" & vbCrLf
    ps = ps & "$form.Controls.Add($btnClose)" & vbCrLf & vbCrLf
    
    ' Link
    If LinkUrl <> "" Then
        ps = ps & "$lnk = New-Object System.Windows.Forms.LinkLabel" & vbCrLf
        ps = ps & "$lnk.Text = '" & PSEscape(LinkUrl) & "'" & vbCrLf
        ps = ps & "$lnk.AutoSize = $true" & vbCrLf
        ps = ps & "$lnk.Location = New-Object System.Drawing.Point(10,110)" & vbCrLf
        ps = ps & "$lnk.LinkColor = [System.Drawing.Color]::FromArgb(" & fgColor & ")" & vbCrLf
        ps = ps & "$lnk.Add_LinkClicked({ Start-Process '" & PSEscape(LinkUrl) & "'; $form.Close() })" & vbCrLf
        ps = ps & "$form.Controls.Add($lnk)" & vbCrLf & vbCrLf
    End If
    
    ' Fade-in
    ps = ps & "$timer = New-Object System.Windows.Forms.Timer" & vbCrLf
    ps = ps & "$timer.Interval = 30" & vbCrLf
    ps = ps & "$timer.Add_Tick({ if ($form.Opacity -lt 1) { $form.Opacity += 0.1 } else { $timer.Stop() } })" & vbCrLf
    ps = ps & "$timer.Start()" & vbCrLf & vbCrLf
    
    ' Auto-close
    If Duration > 0 Then
        ps = ps & "$closeTimer = New-Object System.Windows.Forms.Timer" & vbCrLf
        ps = ps & "$closeTimer.Interval = " & (Duration * 1000) & vbCrLf
        ps = ps & "$closeTimer.Add_Tick({ $form.Close() })" & vbCrLf
        ps = ps & "$closeTimer.Start()" & vbCrLf & vbCrLf
    End If
    
    ' Sound
    If Sound <> "" And UCase(Sound) <> "NONE" Then
        ps = ps & "[System.Media.SystemSounds]::Beep.Play()" & vbCrLf & vbCrLf
    End If
    
    ' Show
    ps = ps & "$form.Add_Shown({ $form.Activate() })" & vbCrLf
    ps = ps & "[void]$form.ShowDialog()" & vbCrLf
    ps = ps & "$form.Dispose()" & vbCrLf
    
    BuildPSScript = ps
End Function

'================= CALCULATE PS POSITION =================
Function CalculatePositionPS(Position)
    Dim ps
    ps = "$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea" & vbCrLf
    
    Select Case UCase(Position)
        Case "TL"
            ps = ps & "$form.Left = 20" & vbCrLf
            ps = ps & "$form.Top = 20" & vbCrLf
        Case "TR"
            ps = ps & "$form.Left = $screen.Right - $form.Width - 20" & vbCrLf
            ps = ps & "$form.Top = 20" & vbCrLf
        Case "BL"
            ps = ps & "$form.Left = 20" & vbCrLf
            ps = ps & "$form.Top = $screen.Bottom - $form.Height - 20" & vbCrLf
        Case "CR"
            ps = ps & "$form.Left = $screen.Right - $form.Width - 20" & vbCrLf
            ps = ps & "$form.Top = ($screen.Height - $form.Height) / 2" & vbCrLf
        Case "C"
            ps = ps & "$form.Left = ($screen.Width - $form.Width) / 2" & vbCrLf
            ps = ps & "$form.Top = ($screen.Height - $form.Height) / 2" & vbCrLf
        Case Else ' BR
            ps = ps & "$form.Left = $screen.Right - $form.Width - 20" & vbCrLf
            ps = ps & "$form.Top = $screen.Bottom - $form.Height - 20" & vbCrLf
    End Select
    
    CalculatePositionPS = ps & vbCrLf
End Function

'================= RUN MSHTA TOAST =================
Sub RunMSHTAToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position)
    Dim fso, tempHTA, shell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    
    tempHTA = fso.GetSpecialFolder(2) & "\TempToastVBS_" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".hta"
    
    ' Build HTML
    Dim html
    html = BuildHTAToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position, tempHTA)
    
    ' Write file
    Dim ts
    Set ts = fso.CreateTextFile(tempHTA, True, True)
    ts.Write html
    ts.Close
    
    ' Launch
    shell.Run "mshta.exe """ & tempHTA & """", 0, False
End Sub

'================= BUILD HTA TOAST =================
Function BuildHTAToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, Sound, Duration, Position, tempFile)
    Dim html, bgColor, textColor, borderColor, icon
    
    ' Determine colors
    Select Case UCase(ToastType)
        Case "WARN"
            bgColor = "linear-gradient(135deg, #ffeb3b, #ffa000)"
            textColor = "#000000"
            borderColor = "#e65100"
            icon = "⚠"
        Case "ERROR"
            bgColor = "linear-gradient(135deg, #ff6b6b, #d32f2f)"
            textColor = "#ffffff"
            borderColor = "#b71c1c"
            icon = "❌"
        Case Else
            bgColor = "linear-gradient(135deg, #4caf50, #2e7d32)"
            textColor = "#ffffff"
            borderColor = "#1b5e20"
            icon = "ℹ"
    End Select
    
    ' Calculate position
    Dim posX, posY
    CalculatePositionMSHTA Position, posX, posY
    
    ' Build HTML
    html = "<html><head><meta charset='UTF-8'><title>" & HTMLEscape(Title) & "</title>" & vbCrLf
    html = html & "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no' SYSMENU='no' SCROLL='no'>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body{margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;background:transparent;overflow:hidden;}" & vbCrLf
    html = html & "#toast{position:fixed;padding:15px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);" & vbCrLf
    html = html & "width:350px;min-height:150px;background:" & bgColor & ";" & vbCrLf
    html = html & "border:2px solid " & borderColor & ";color:" & textColor & ";transform:translateX(120%);opacity:0;" & vbCrLf
    html = html & "transition:transform 500ms cubic-bezier(0.68,-0.55,0.265,1.55),opacity 500ms ease;}" & vbCrLf
    html = html & "h3{margin:0 0 8px;font-size:18px;font-weight:600;}" & vbCrLf
    html = html & "p{margin:0 0 10px;font-size:14px;line-height:1.5;}" & vbCrLf
    html = html & "a{color:" & textColor & ";text-decoration:underline;cursor:pointer;font-size:13px;}" & vbCrLf
    html = html & "a:hover{opacity:0.8;}" & vbCrLf
    html = html & "#dismissBtn{position:absolute;top:8px;right:8px;padding:4px 8px;font-size:16px;font-weight:bold;" & vbCrLf
    html = html & "background:rgba(0,0,0,0.3);border:none;color:" & textColor & ";border-radius:4px;cursor:pointer;}" & vbCrLf
    html = html & "#dismissBtn:hover{background:rgba(0,0,0,0.5);}" & vbCrLf
    html = html & "</style>" & vbCrLf
    html = html & "<script>" & vbCrLf
    
    ' Position window
    html = html & "window.resizeTo(370,170);" & vbCrLf
    html = html & "window.moveTo(" & posX & "," & posY & ");" & vbCrLf
    
    ' Dismiss function
    html = html & "function dismiss(){" & vbCrLf
    html = html & "var t=document.getElementById('toast');" & vbCrLf
    html = html & "t.style.transform='translateX(120%)';" & vbCrLf
    html = html & "t.style.opacity='0';" & vbCrLf
    html = html & "setTimeout(function(){" & vbCrLf
    html = html & "window.close();" & vbCrLf
    html = html & "try{new ActiveXObject('Scripting.FileSystemObject').DeleteFile('" & Replace(tempFile, "\", "\\") & "');}catch(e){}" & vbCrLf
    html = html & "},600);" & vbCrLf
    html = html & "}" & vbCrLf
    
    ' Slide-in
    html = html & "setTimeout(function(){" & vbCrLf
    html = html & "var t=document.getElementById('toast');" & vbCrLf
    html = html & "t.style.transform='translateX(0)';" & vbCrLf
    html = html & "t.style.opacity='1';" & vbCrLf
    html = html & "},100);" & vbCrLf
    
    ' Auto-close
    If Duration > 0 Then
        html = html & "setTimeout(function(){dismiss();}," & (Duration * 1000) & ");" & vbCrLf
    End If
    
    ' Link handler
    If LinkUrl <> "" Then
        html = html & "function openLink(){" & vbCrLf
        html = html & "try{new ActiveXObject('WScript.Shell').Run('cmd /c start \"\" \"" & JSEscape(LinkUrl) & "\"',0,false);}catch(e){}" & vbCrLf
        html = html & "dismiss();" & vbCrLf
        html = html & "}" & vbCrLf
    End If
    
    html = html & "</script></head><body><div id='toast'>" & vbCrLf
    
    ' Dismiss button
    html = html & "<button id='dismissBtn' onclick='dismiss()' title='Close'>×</button>" & vbCrLf
    
    ' Content
    html = html & "<h3>" & icon & " " & HTMLEscape(Title) & "</h3>" & vbCrLf
    html = html & "<p>" & Replace(HTMLEscape(Message), vbCrLf, "<br>") & "</p>" & vbCrLf
    
    ' Link
    If LinkUrl <> "" Then
        html = html & "<a onclick='openLink();'>" & HTMLEscape(LinkUrl) & "</a>" & vbCrLf
    End If
    
    html = html & "</div></body></html>"
    
    BuildHTAToast = html
End Function

'================= CALCULATE MSHTA POSITION =================
Sub CalculatePositionMSHTA(Position, ByRef outX, ByRef outY)
    Dim screenW, screenH, margin
    screenW = 1920 ' Default, can't easily get actual screen size in VBS
    screenH = 1080
    margin = 20
    
    Select Case UCase(Position)
        Case "TL"
            outX = margin
            outY = margin
        Case "TR"
            outX = screenW - 370 - margin
            outY = margin
        Case "BL"
            outX = margin
            outY = screenH - 170 - margin
        Case "CR"
            outX = screenW - 370 - margin
            outY = (screenH - 170) / 2
        Case "C"
            outX = (screenW - 370) / 2
            outY = (screenH - 170) / 2
        Case Else ' BR
            outX = screenW - 370 - margin
            outY = screenH - 170 - margin
    End Select
End Sub

'================= ESCAPE FUNCTIONS =================
Function HTMLEscape(txt)
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, """", "&quot;")
    txt = Replace(txt, "'", "&#39;")
    HTMLEscape = txt
End Function

Function JSEscape(txt)
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, "'", "\'")
    JSEscape = txt
End Function

Function PSEscape(txt)
    txt = Replace(txt, "'", "''")
    txt = Replace(txt, "`", "``")
    txt = Replace(txt, "$", "`$")
    PSEscape = txt
End Function

'================= QUEUE TOASTS =================
Sub QueueToasts()
    Dim i
    For i = 1 To 5
        ShowToast "Queued Toast #" & i, "This is toast number " & i & " of 5.", "INFO", "", "", "", "", 3, "BR"
        WScript.Sleep 1000
    Next
    MsgBox "All 5 toasts queued!", vbInformation, "Queue Complete"
End Sub

'================= POSITION TEST =================
Sub PositionTest()
    Dim positions, posNames, i
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    posNames = Array("Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center-Right", "Center")
    
    MsgBox "Will display 6 toasts in different positions." & vbCrLf & _
           "Each will appear for 3 seconds.", vbInformation, "Position Test"
    
    For i = 0 To UBound(positions)
        ShowToast "Position Test", "Testing: " & posNames(i), "INFO", "", "", "", "", 3, positions(i)
        WScript.Sleep 3500
    Next
    
    MsgBox "Position test complete!", vbInformation, "Test Complete"
End Sub

'================= TYPE TEST =================
Sub TypeTest()
    MsgBox "Will display 3 toasts: Info, Warning, Error", vbInformation, "Type Test"
    
    ShowToast "Info Toast", "This is an informational toast.", "INFO", "", "", "", "", 3, "BR"
    WScript.Sleep 3500
    
    ShowToast "Warning Toast", "This is a warning toast!", "WARN", "", "", "", "BEEP", 3, "TR"
    WScript.Sleep 3500
    
    ShowToast "Error Toast", "This is an error toast!", "ERROR", "", "", "", "BEEP", 3, "TL"
    WScript.Sleep 3500
    
    MsgBox "Type test complete!", vbInformation, "Test Complete"
End Sub

'================= TOGGLE MODE =================
Sub ToggleMode()
    UsePowerShellToasts = Not UsePowerShellToasts
    MsgBox "Toast Mode: " & GetMode() & vbCrLf & vbCrLf & _
           "PowerShell: " & IIf(UsePowerShellToasts, "ENABLED", "DISABLED") & vbCrLf & _
           "MSHTA: " & IIf(Not UsePowerShellToasts, "ENABLED", "DISABLED"), _
           vbInformation, "Mode Changed"
End Sub

'================= GET MODE =================
Function GetMode()
    If UsePowerShellToasts Then
        GetMode = "PowerShell"
    Else
        GetMode = "MSHTA"
    End If
End Function

'================= SYSTEM CHECK =================
Sub SystemCheck()
    Dim report
    report = "===== SYSTEM CHECK =====" & vbCrLf & vbCrLf
    report = report & "PowerShell Available: " & IIf(PowerShellAvailable(), "YES", "NO") & vbCrLf
    report = report & "MSHTA Available: " & IIf(MSHTAAvailable(), "YES", "NO") & vbCrLf & vbCrLf
    report = report & "Current Mode: " & GetMode() & vbCrLf
    report = report & "Default Position: " & DefaultPosition & vbCrLf & vbCrLf
    
    If PowerShellAvailable() Or MSHTAAvailable() Then
        report = report & "Status: OPERATIONAL ✓"
    Else
        report = report & "Status: LIMITED (Fallback to MsgBox)"
    End If
    
    MsgBox report, vbInformation, "System Check"
End Sub

'================= IIF FUNCTION (VBS doesn't have native IIf) =================
Function IIf(condition, trueVal, falseVal)
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

'================= START =================
MainMenu