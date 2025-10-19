Attribute VB_Name = "MsgBoxUniversal"
' =====================================================================
'  Module: MsgBoxUniversal
'  Version: 5.2
'  Purpose: Unified Notification System for VBA (PowerShell, MSHTA, Classic)
'  Author: Keith Swerling + ChatGPT (GPT-5)
' =====================================================================

Option Explicit

' === GLOBAL SETTINGS =================================================
Public UsePowerShellToasts As Boolean
Public Const MSGBOX_VERSION As String = "5.2"
Public Const MSGBOX_UNIVERSAL_TITLE As String = "MsgBoxUniversal v5.2"

' Class instances for temp files & callbacks
Private FileIO As clsFileIO
Private Callbacks As clsCallbacks
' === PLACEHOLDER UTILITIES ==========================================
Public Sub StartToastListener()
    Dim psPath As String
    psPath = Environ$("USERPROFILE") & "\OneDrive\Documents\2025\Powershell\ToastWatcher.ps1"
    
    If Dir(psPath) <> "" Then
        ' Launch ToastWatcher v5.3 in hidden window
        shell "powershell -NoProfile -ExecutionPolicy Bypass -File """ & psPath & """", vbHide
        MsgBox "ToastWatcher v5.3 launched.", vbInformation, "Listener"
    Else
        MsgBox "ToastWatcher.ps1 not found: " & psPath, vbExclamation, "Listener"
    End If
End Sub

Public Sub StopToastListener()
    ' Attempt to kill all PowerShell processes running ToastWatcher.ps1
    Dim cmd As String
    cmd = "Get-Process powershell | Where-Object {$_.Path -like '*ToastWatcher.ps1*'} | Stop-Process -Force"
    shell "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & cmd & """", vbHide
    MsgBox "ToastWatcher v5.3 stopped.", vbInformation, "Listener"
End Sub

' === UNIVERSAL TEST SUITE (Updated) ==================================
Public Sub MsgBoxUnifiedTestMenu()
    UI.Notify "Notification", "Simple notification (auto mode)"
    UI.Progress "File Upload", "Uploading to server...", 45
    
    ' --- New: Test sending to ToastWatcher v5.3 directly via temp JSON ---
    Dim jsonFile As String
    jsonFile = Environ$("TEMP") & "\ExcelToastTest.json"
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(jsonFile, True)
    ts.Write "{""Title"":""Excel Test"",""Message"":""This is a live test from VBA."",""Level"":""INFO"",""Progress"":25,""Duration"":5}"
    ts.Close
    
    MsgBox "Test toast JSON created: " & jsonFile, vbInformation, "Temp File Test"
End Sub

Private Sub Init()
    If FileIO Is Nothing Then Set FileIO = New clsFileIO
    If Callbacks Is Nothing Then Set Callbacks = New clsCallbacks
End Sub

' === MAIN UNIFIED NOTIFICATION ENTRY =================================
Public Function ShowMsgBoxUnified( _
    ByVal Message As String, _
    Optional ByVal Title As String = "Notification", _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal Mode As String = "auto", _
    Optional ByVal TimeoutSeconds As Long = 5, _
    Optional ByVal Level As String = "INFO", _
    Optional ByVal CallbackName As String = "" _
) As VbMsgBoxResult

    On Error GoTo SafeExit
    Init

    Dim modeResolved As String
    modeResolved = LCase$(Mode)

    ' Auto-select mode
    If modeResolved = "auto" Then
        If UsePowerShellToasts Then
            modeResolved = "ps"
        Else
            modeResolved = "classic"
        End If
    End If

    Select Case modeResolved
        Case "ps", "powershell"
            Call ShowPowerShellToast(Message, Title, TimeoutSeconds, Level, CallbackName)
        Case "mshta"
            Call ShowMSHTAToast(Message, Title, TimeoutSeconds, Level, CallbackName)
        Case "wscript"
            Call ShowWScriptPopup(Message, Title, TimeoutSeconds)
        Case Else
            ShowMsgBoxUnified = MsgBox(Message, Buttons, Title)
    End Select

SafeExit:
    Exit Function
End Function

' === SUBSYSTEM: POWERSHELL ===========================================
Private Sub ShowPowerShellToast(msg As String, Title As String, Timeout As Long, Level As String, Optional CallbackName As String = "")
    On Error Resume Next
    Init

    Dim psPath As String
    psPath = Environ$("USERPROFILE") & "\OneDrive\Documents\2025\Powershell\ToastNotify.ps1"

    If Dir(psPath) <> "" Then
        Dim tmpFile As String
        tmpFile = FileIO.TempFile(".txt")
        ' Write callback reference
        If Len(CallbackName) > 0 Then
            FileIO.WriteTextFile tmpFile, CallbackName
        End If

        shell "powershell -ExecutionPolicy Bypass -File """ & psPath & _
              """ -Message """ & Replace(msg, """", "'") & _
              """ -Title """ & Replace(Title, """", "'") & _
              """ -Level " & Level & " -Timeout " & Timeout & _
              " -CallbackFile """ & tmpFile & """", vbHide
    Else
        MsgBox msg, vbInformation, Title & " (PowerShell toast missing)"
    End If
End Sub

' === SUBSYSTEM: MSHTA ================================================
Private Sub ShowMSHTAToast(msg As String, Title As String, Timeout As Long, Level As String, Optional CallbackName As String = "")
    On Error Resume Next
    Init

    Dim html As String
    html = "<html><head><title>" & Title & "</title>" & _
           "<script>" & _
           "function closeMe(){window.close();}" & _
           "setTimeout(closeMe, " & (Timeout * 1000) & ");" & _
           "function doCallback(){"
    If Len(CallbackName) > 0 Then
        html = html & "alert('Callback triggered: " & CallbackName & "');"
    End If
    html = html & "}" & _
           "</script>" & _
           "<style>body{font-family:Segoe UI; background:#222; color:white; padding:10px;}</style>" & _
           "</head><body onclick='doCallback()'>" & _
           "<h4>" & Title & "</h4><p>" & msg & "</p></body></html>"

    Dim tmpPath As String
    tmpPath = FileIO.TempFile(".html")
    FileIO.WriteTextFile tmpPath, html

    shell "mshta """ & tmpPath & """", vbHide
End Sub

' === SUBSYSTEM: WScript ==============================================
Private Sub ShowWScriptPopup(msg As String, Title As String, Timeout As Long)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Popup msg, Timeout, Title, 64
End Sub

' === UNIVERSAL TEST MENU ============================================
Public Sub MsgBoxUniversalMainMenu()
    Dim choice As String
    choice = InputBox( _
        "Unified Notification System (v5.2)" & vbCrLf & _
        "------------------------------------------" & vbCrLf & _
        "1. Universal Test Menu (mixed modes)" & vbCrLf & _
        "2. PowerShell Toast Tests" & vbCrLf & _
        "3. MSHTA / HTML5 Popup Tests" & vbCrLf & _
        "4. Classic / WScript Mode Tests" & vbCrLf & _
        "5. Start ToastWatcher.ps1 Listener" & vbCrLf & _
        "6. Stop Listener" & vbCrLf & _
        "7. Toggle PowerShell Toast Mode" & vbCrLf & _
        "8. Cleanup Temp Toast Files" & vbCrLf & _
        "9. Exit", _
        MSGBOX_UNIVERSAL_TITLE, "1")

    If choice = "" Or choice = "9" Then Exit Sub

    Select Case Val(choice)
        Case 1
            MsgBoxUnifiedTestMenu
        Case 2
            Call UI.Notify("PowerShell Toast Test", "Test via MsgBoxToastsPS", 5, "INFO", "TestCallback")
        Case 3
            Call UI.Notify("MSHTA Popup Test", "Test via HTML/MSHTA", 5, "INFO", "TestCallback", "mshta")
        Case 4
            MsgBoxClassicAndPopupTests
        Case 5
            StartToastListener
        Case 6
            StopToastListener
        Case 7
            UsePowerShellToasts = Not UsePowerShellToasts
            MsgBox "PowerShell Toast Mode: " & IIf(UsePowerShellToasts, "ENABLED", "DISABLED"), vbInformation
        Case 8
            CleanupTempPS
        Case Else
            MsgBox "Invalid selection.", vbExclamation
    End Select

    MsgBoxUniversalMainMenu
End Sub

' === CLASSIC/WScript TESTS ==========================================
Private Sub MsgBoxClassicAndPopupTests()
    Dim choice As String
    choice = InputBox("Classic/WScript Tests:" & vbCrLf & _
                      "1. Classic MsgBox" & vbCrLf & _
                      "2. WScript Popup" & vbCrLf & _
                      "0. Back", "Classic/WScript", "1")

    Select Case Val(choice)
        Case 1: MsgBox "Classic MsgBox test.", vbInformation, "Classic"
        Case 2: ShowMsgBoxUnified "Popup via WScript", "WScript", vbOKOnly, "wscript", 3
        Case Else: Exit Sub
    End Select
    MsgBoxClassicAndPopupTests
End Sub


Public Sub CleanupTempPS()
    On Error Resume Next
    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\toast_*.html"
    Kill tmpPath
    MsgBox "Temporary files cleaned.", vbInformation
End Sub


' =====================================================================
'  UI.NOTIFICATION WRAPPER LAYER (PUBLIC INTERFACE)
' =====================================================================

    Public Sub Notify(Optional ByVal Title As String = "Notice", _
                      Optional ByVal Message As String = "", _
                      Optional ByVal Timeout As Long = 5, _
                      Optional ByVal Level As String = "INFO", _
                      Optional ByVal CallbackName As String = "", _
                      Optional ByVal Mode As String = "auto")
        Call ShowMsgBoxUnified(Message, Title, vbOKOnly, Mode, Timeout, Level, CallbackName)
    End Sub

    Public Sub Progress(ByVal Title As String, _
                        ByVal Message As String, _
                        ByVal Percent As Double)
        Dim displayMsg As String
        displayMsg = Message & vbCrLf & "[" & Format(Percent, "0.0") & "% Complete]"
        Call ShowMsgBoxUnified(displayMsg, Title, vbOKOnly, "ps", 4, "PROGRESS")
    End Sub

    Public Function Inputu(ByVal Title As String, _
                          ByVal Prompt As String, _
                          Optional ByVal DefaultValue As String = "") As String
        Inputu = InputBox(Prompt, Title, DefaultValue)
    End Function


