Attribute VB_Name = "MsgBoxUI"
Option Explicit

' =====================================================================
' Module: MsgBoxUI.bas
' Version: 5.4
' Purpose: Unified VBA Notification Wrappers & ToastWatcher Integration
' Author: Keith Swerling + ChatGPT (GPT-5)
' Updated: 2025-10-18
' =====================================================================

Public Const MSGBOXUI_VERSION As String = "5.4"
Public Const MSGBOXUI_TITLE As String = "MsgBoxUI v5.4"
Public UsePowerShellToasts As Boolean
Public UseTempJsonFallback As Boolean
Public ToastPipeName As String

' =====================================================================
' INITIALIZATION
' =====================================================================
Public Sub MsgBoxUI_Init(Optional ByVal EnableToasts As Boolean = True, _
                         Optional ByVal EnableFallback As Boolean = True, _
                         Optional ByVal PipeName As String = "\\.\pipe\ExcelToastPipe")

    UsePowerShellToasts = EnableToasts
    UseTempJsonFallback = EnableFallback
    ToastPipeName = PipeName
End Sub

' =====================================================================
' SIMPLE WRAPPERS
' =====================================================================
Public Sub Notify(Optional ByVal Title As String = "Notice", _
                  Optional ByVal Message As String = "", _
                  Optional ByVal Timeout As Long = 5, _
                  Optional ByVal Level As String = "INFO", _
                  Optional ByVal Mode As String = "auto")
    Call ShowMsgBoxUnified(Message, Title, vbOKOnly, Mode, Timeout, Level)
End Sub

Public Sub Progress(ByVal Title As String, _
                    ByVal Message As String, _
                    ByVal Percent As Double)
    Dim displayMsg As String
    displayMsg = Message & vbCrLf & "[" & Format(Percent, "0.0") & "% Complete]"
    Call ShowMsgBoxUnified(displayMsg, Title, vbOKOnly, "ps", 4, "PROGRESS")
End Sub

Public Function InputText(ByVal Title As String, _
                          ByVal Prompt As String, _
                          Optional ByVal DefaultValue As String = "") As String
    InputText = InputBox(Prompt, Title, DefaultValue)
End Function


' =====================================================================
' LIVE TOAST OUTPUT
' =====================================================================
Public Sub LiveToast(ByVal Title As String, _
                     ByVal Message As String, _
                     Optional ByVal Level As String = "INFO", _
                     Optional ByVal Progress As Long = 0)
    On Error Resume Next
    Dim PipeName As String
    PipeName = IIf(ToastPipeName = "", "\\.\pipe\ExcelToastPipe", ToastPipeName)
    
    ' First try pipe-based push
    If SendToastViaPipe(PipeName, Title, Message, Level, Progress) Then Exit Sub
    
    ' Fallback to JSON temp file if enabled
    If UseTempJsonFallback Then
        Dim fso As Object, TempFolder As String, jsonFile As String, jsonContent As String
        TempFolder = Environ$("TEMP") & "\ExcelToasts"
        If Dir(TempFolder, vbDirectory) = "" Then MkDir TempFolder
        jsonFile = TempFolder & "\ToastRequest.json"
        jsonContent = "{""Title"":""" & Replace(Title, """", "'") & """," & _
                      """Message"":""" & Replace(Message, """", "'") & """," & _
                      """Level"":""" & Level & """," & _
                      """Progress"":" & Progress & "," & _
                      """Attribution"":""Excel VBA""}"
        Set fso = CreateObject("Scripting.FileSystemObject")
        With fso.CreateTextFile(jsonFile, True)
            .Write jsonContent
            .Close
        End With
    End If
End Sub

Private Function SendToastViaPipe(ByVal PipeName As String, _
                                  ByVal Title As String, _
                                  ByVal Message As String, _
                                  ByVal Level As String, _
                                  ByVal Progress As Long) As Boolean
    On Error Resume Next
    Dim json As String
    json = "{""Title"":""" & Title & """," & _
           """Message"":""" & Message & """," & _
           """Level"":""" & Level & """," & _
           """Progress"":" & Progress & "," & _
           """Attribution"":""Excel VBA""}"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "utf-8"
        .Open
        .WriteText json
        .Position = 0
        Dim bytes() As Byte
        bytes = .Read(-1)
        .Close
    End With

    Dim pipe As Object
    Set pipe = CreateObject("Scripting.FileSystemObject").OpenTextFile(PipeName, 2, False)
    pipe.Write json
    pipe.Close
    SendToastViaPipe = True
End Function


' =====================================================================
' MAIN UNIFIED MSGBOX ROUTER
' =====================================================================
Public Function ShowMsgBoxUnified(ByVal Message As String, _
                                  Optional ByVal Title As String = "Notification", _
                                  Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal Mode As String = "auto", _
                                  Optional ByVal TimeoutSeconds As Long = 5, _
                                  Optional ByVal Level As String = "INFO") As VbMsgBoxResult

    On Error GoTo SafeExit
    Dim resolvedMode As String
    resolvedMode = LCase$(Mode)

    If resolvedMode = "auto" Then
        If UsePowerShellToasts Then
            resolvedMode = "ps"
        Else
            resolvedMode = "classic"
        End If
    End If

    Select Case resolvedMode
        Case "ps", "powershell"
            Call LiveToast(Title, Message, Level, 0)
        Case "mshta"
            Call ShowMSHTAToast(Message, Title, TimeoutSeconds, Level)
        Case "wscript"
            Call ShowWScriptPopup(Message, Title, TimeoutSeconds)
        Case Else
            ShowMsgBoxUnified = MsgBox(Message, Buttons, Title)
    End Select

SafeExit:
    Exit Function
End Function


' =====================================================================
' MSHTA TOAST
' =====================================================================
Private Sub ShowMSHTAToast(msg As String, Title As String, Timeout As Long, Level As String)
    On Error Resume Next
    Dim html As String, tmpPath As String, f As Object
    html = "<html><head><title>" & Title & "</title><script>" & _
           "setTimeout('window.close()', " & (Timeout * 1000) & ");</script>" & _
           "<style>body{font-family:Segoe UI; background:#202020; color:white; padding:10px;}</style></head>" & _
           "<body><h4>" & Title & "</h4><p>" & msg & "</p></body></html>"

    tmpPath = Environ$("TEMP") & "\toast.html"
    Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(tmpPath, True)
    f.Write html
    f.Close
    shell "mshta """ & tmpPath & """", vbHide
End Sub

' =====================================================================
' WScript Popup
' =====================================================================
Private Sub ShowWScriptPopup(msg As String, Title As String, Timeout As Long)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Popup msg, Timeout, Title, 64
End Sub


' =====================================================================
' TOAST WATCHER CONTROL
' =====================================================================
Public Sub StartToastListener()
    Dim psPath As String
    psPath = Environ$("USERPROFILE") & "\OneDrive\Documents\2025\Powershell\ToastWatcherRT.ps1"

    If Dir(psPath) <> "" Then
        shell "powershell -ExecutionPolicy Bypass -File """ & psPath & """ -UseWinRT -VerboseLog", vbHide
        MsgBox "ToastWatcherRT listener launched.", vbInformation, "MsgBoxUI"
    Else
        MsgBox "ToastWatcherRT.ps1 not found." & vbCrLf & psPath, vbExclamation, "MsgBoxUI"
    End If
End Sub

Public Sub StopToastListener()
    On Error Resume Next
    Dim exitFlag As String
    exitFlag = Environ$("TEMP") & "\ExcelToasts\ToastWatcherExit.flag"
    CreateObject("Scripting.FileSystemObject").CreateTextFile(exitFlag, True).Close
    MsgBox "Stop signal sent to ToastWatcherRT listener.", vbInformation, "MsgBoxUI"
End Sub


' =====================================================================
' DEMO MENU
' =====================================================================
Public Sub MsgBoxUI_MainMenu()
    Dim choice As String
    choice = InputBox( _
        MSGBOXUI_TITLE & " Demo Menu" & vbCrLf & _
        "------------------------------------" & vbCrLf & _
        "1. Simple Notification" & vbCrLf & _
        "2. Progress Demo" & vbCrLf & _
        "3. Input Prompt Demo" & vbCrLf & _
        "4. Start ToastWatcherRT.ps1 Listener" & vbCrLf & _
        "5. Stop ToastWatcherRT.ps1 Listener" & vbCrLf & _
        "6. Send Live Toast (manual entry)" & vbCrLf & _
        "7. Exit", MSGBOXUI_TITLE, "1")

    If choice = "" Or choice = "7" Then Exit Sub

    Select Case val(choice)
        Case 1
            Notify "Demo Notification", "This is a simple notification!"
        Case 2
            Progress "Demo Progress", "Uploading data...", 42
        Case 3
            Dim result As String
            result = InputText("Demo Input", "Enter your name:", "Guest")
            MsgBox "You entered: " & result, vbInformation, "Input Result"
        Case 4
            StartToastListener
        Case 5
            StopToastListener
        Case 6
            SendLiveToastDemo
        Case Else
            MsgBox "Invalid selection.", vbExclamation
    End Select

    MsgBoxUI_MainMenu
End Sub


' =====================================================================
' LIVE TOAST TEST (INTERACTIVE)
' =====================================================================
Public Sub SendLiveToastDemo()
    Dim Title As String, msg As String, lvl As String, prog As Long
    Title = InputBox("Enter toast title:", "Live Toast Test", "Excel Test")
    msg = InputBox("Enter toast message:", "Live Toast Test", "This is a live test from VBA.")
    lvl = InputBox("Enter level (INFO/WARN/ERROR/PROGRESS):", "Live Toast Test", "INFO")
    prog = CLng(InputBox("Enter progress % (0–100):", "Live Toast Test", "25"))
    If prog < 0 Then prog = 0
    If prog > 100 Then prog = 100
    LiveToast Title, msg, lvl, prog
    MsgBox "Toast sent to listener (" & lvl & ")", vbInformation, "MsgBoxUI"
End Sub


