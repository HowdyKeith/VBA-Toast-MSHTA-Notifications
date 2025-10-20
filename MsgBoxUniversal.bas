Attribute VB_Name = "MsgBoxUniversal"
'***************************************************************
' Module: MsgBoxUniversal
' Purpose: Unified HTML/JSON-based Toast Engine for MSHTA and WinRT
' Features:
'   - Native Sound Alerts (custom WAV or system sound)
'   - Optional Application.OnTime cleanup scheduler
'   - Callback Macro auto-invoker
'   - Progress & WinRT-compatible JSON payload
' Dependencies:
'   clsToastNotification.cls, clsToastProgress.cls, clsCallbacks.cls, clsFileIO.cls
'***************************************************************
' Module: MsgBoxUniversal
' Version: 7.0
' Purpose: Unified notification interface
'   - Wraps clsToastNotification v7.1
'   - Supports HTA, WinRT, ToastAPI delivery
'   - Queued & stacked toasts, progress updates
'   - Native sound alerts + legacy beep fallback
'   - Callback macro auto-invoker
'   - Optional OnTime timer cleanup
'***************************************************************

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#End If

Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_ALIAS As Long = &H10000


' --- PUBLIC INTERFACE ---

Public Function ShowToast(Optional ByVal Title As String = "", _
                          Optional ByVal Message As String = "", _
                          Optional ByVal Level As String = "INFO", _
                          Optional ByVal Duration As Long = 5, _
                          Optional ByVal Position As String = "BR", _
                          Optional ByVal Icon As String = "", _
                          Optional ByVal CallbackMacro As String = "", _
                          Optional ByVal SoundName As String = "", _
                          Optional ByVal UseOnTime As Boolean = False, _
                          Optional ByVal Mode As String = "auto") As Boolean
    On Error GoTo ErrHandler
    
    Dim toast As clsToastNotification
    Set toast = New clsToastNotification
    
    toast.Title = Title
    toast.Message = Message
    toast.Level = Level
    toast.Duration = Duration
    toast.Position = Position
    toast.Icon = Icon
    toast.CallbackMacro = CallbackMacro
    toast.SoundName = SoundName
    toast.UseOnTime = UseOnTime
    
    ShowToast = toast.Show(Mode)
    Exit Function
ErrHandler:
    Debug.Print "[MsgBoxUniversal] ShowToast error: " & Err.Description
    ShowToast = False
End Function

Public Function ShowQueuedToast(Optional ByVal Title As String = "", _
                                Optional ByVal Message As String = "", _
                                Optional ByVal Level As String = "INFO", _
                                Optional ByVal Duration As Long = 5, _
                                Optional ByVal Position As String = "BR", _
                                Optional ByVal Icon As String = "", _
                                Optional ByVal CallbackMacro As String = "", _
                                Optional ByVal SoundName As String = "", _
                                Optional ByVal UseOnTime As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    
    Dim toast As clsToastNotification
    Set toast = New clsToastNotification
    
    toast.Title = Title
    toast.Message = Message
    toast.Level = Level
    toast.Duration = Duration
    toast.Position = Position
    toast.Icon = Icon
    toast.CallbackMacro = CallbackMacro
    toast.SoundName = SoundName
    toast.UseOnTime = UseOnTime
    
    toast.ShowQueued
    ShowQueuedToast = True
    Exit Function
ErrHandler:
    Debug.Print "[MsgBoxUniversal] ShowQueuedToast error: " & Err.Description
    ShowQueuedToast = False
End Function

Public Sub UpdateProgressToast(ByVal toast As clsToastNotification, ByVal Percent As Long, Optional ByVal NewMessage As String = "")
    On Error Resume Next
    If Not toast Is Nothing Then toast.UpdateProgress Percent, NewMessage
End Sub

Public Sub CloseToast(ByVal toast As clsToastNotification)
    On Error Resume Next
    If Not toast Is Nothing Then toast.Close
End Sub

' --- UTILITY HELPERS ---
Public Function IsToastListenerRunning() As Boolean
    On Error Resume Next
    Dim temp As String
    temp = Environ$("TEMP") & "\ExcelToasts\ListenerHeartbeat.txt"
    If Dir(temp) = "" Then
        IsToastListenerRunning = False
    Else
        Dim fso As Object, lastMod As Date
        Set fso = CreateObject("Scripting.FileSystemObject")
        lastMod = fso.GetFile(temp).DateLastModified
        IsToastListenerRunning = (DateDiff("s", lastMod, Now) < 10)
    End If
End Function











'***************************************************************
' PUBLIC ENTRY POINT
'***************************************************************
Public Function ShowHTMLToast(ByVal toast As clsToastNotification) As Boolean
    On Error GoTo FailSafe
    
    Dim htmlPath As String
    htmlPath = BuildHTMLFile(toast)
    
    If Len(htmlPath) = 0 Then Exit Function
    
    ' === Sound Support ===
    If Len(toast.SoundName) > 0 Then
        PlayToastSound toast.SoundName
    End If
    
    ' === Launch MSHTA ===
    Dim cmd As String
    cmd = "mshta.exe """ & htmlPath & """"
    shell cmd, vbHide
    
    ' === Callback Support ===
    If Len(toast.CallbackMacro) > 0 Then
        ScheduleCallback toast
    End If
    
    ' === Auto Cleanup ===
    If g_UseOnTime Then
        ScheduleCleanup toast
    End If
    
    ShowHTMLToast = True
    Exit Function
    
FailSafe:
    Debug.Print "[MsgBoxToastUnified] Error: " & Err.Description
    ShowHTMLToast = False
End Function

'***************************************************************
' SOUND ENGINE
'***************************************************************
Private Sub PlayToastSound(ByVal SoundName As String)
    On Error Resume Next
    
    Select Case UCase$(SoundName)
        Case "INFO": PlaySound "Windows Notify System Generic", 0, SND_ASYNC Or SND_ALIAS
        Case "SUCCESS": PlaySound "Windows Notify System Mail", 0, SND_ASYNC Or SND_ALIAS
        Case "WARNING": PlaySound "Windows Exclamation", 0, SND_ASYNC Or SND_ALIAS
        Case "ERROR": PlaySound "Windows Critical Stop", 0, SND_ASYNC Or SND_ALIAS
        Case "CRITICAL": PlaySound "SystemHand", 0, SND_ASYNC Or SND_ALIAS
        Case "PROGRESS": PlaySound "Windows Balloon", 0, SND_ASYNC Or SND_ALIAS
        Case Else
            ' Try file path
            If InStr(SoundName, "\") > 0 Then
                PlaySound SoundName, 0, SND_ASYNC Or SND_FILENAME
            End If
    End Select
End Sub

'***************************************************************
' CALLBACK SUPPORT
'***************************************************************
Private Sub ScheduleCallback(ByVal toast As clsToastNotification)
    On Error Resume Next
    
    If Not g_Callbacks Is Nothing Then
        g_Callbacks.RegisterCallback toast.CallbackMacro, toast.Title, toast.Level
    Else
        Application.OnTime Now + TimeSerial(0, 0, toast.Duration), "'" & toast.CallbackMacro & "'"
    End If
End Sub

'***************************************************************
' CLEANUP SCHEDULING
'***************************************************************
Private Sub ScheduleCleanup(ByVal toast As clsToastNotification)
    On Error Resume Next
    If toast.Duration <= 0 Then Exit Sub
    Dim runTime As Date
    runTime = Now + TimeSerial(0, 0, toast.Duration + 1)
    Application.OnTime runTime, "MsgBoxToastAPI.CleanupToasts"
End Sub

'***************************************************************
' HTML BUILDER
'***************************************************************
Private Function BuildHTMLFile(ByVal toast As clsToastNotification) As String
    On Error GoTo FailSafe
    
    Dim html As String
    html = BuildHTML(toast)
    
    Dim tempPath As String
    tempPath = g_FileIO.GetTempFilePath("Toast_" & Format(Now, "hhmmss") & ".hta")
    
    g_FileIO.WriteText tempPath, html
    BuildHTMLFile = tempPath
    Exit Function
    
FailSafe:
    Debug.Print "[MsgBoxToastUnified] Failed to build HTML: " & Err.Description
    BuildHTMLFile = vbNullString
End Function

'***************************************************************
' HTML RENDERER
'***************************************************************
Private Function BuildHTML(ByVal toast As clsToastNotification) As String
    Dim html As String
    html = ""
    
    html = html & "<html><head><meta http-equiv='X-UA-Compatible' content='IE=11'>" & vbCrLf
    html = html & "<title>" & EscapeHtml(toast.Title) & "</title>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body {font-family:Segoe UI; background:#202020; color:#fff; margin:0; padding:0;}" & vbCrLf
    html = html & ".toast {border-radius:8px; padding:14px; margin:8px; box-shadow:0 4px 12px rgba(0,0,0,0.4);}" & vbCrLf
    html = html & ".title {font-weight:bold; font-size:14pt;}" & vbCrLf
    html = html & ".message {margin-top:4px; font-size:11pt;}" & vbCrLf
    html = html & ".progress {width:100%; height:6px; border-radius:3px; background:#333; margin-top:8px;}" & vbCrLf
    html = html & ".bar {height:6px; border-radius:3px; background:#00ccff;}" & vbCrLf
    html = html & "</style></head><body>" & vbCrLf
    
    html = html & "<div class='toast'>" & vbCrLf
    html = html & "<div class='title'>" & EscapeHtml(toast.Title) & "</div>" & vbCrLf
    html = html & "<div class='message'>" & EscapeHtml(toast.Message) & "</div>" & vbCrLf
    
    If toast.Level = "PROGRESS" Then
        html = html & "<div class='progress'><div class='bar' style='width:" & CInt(toast.Progress * 100) & "%'></div></div>" & vbCrLf
    End If
    
    html = html & "</div>" & vbCrLf
    html = html & "<script>" & vbCrLf
    html = html & "setTimeout(function(){window.close();}, " & (toast.Duration * 1000) & ");" & vbCrLf
    html = html & "</script>" & vbCrLf
    html = html & "</body></html>"
    
    BuildHTML = html
End Function

'***************************************************************
' UTILITY
'***************************************************************
Private Function EscapeHtml(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")
    EscapeHtml = text
End Function


Public Sub DemoMsgBoxToasts()
    ' Initialize the toast system
    InitMsgBoxToast useWinRT:=True, UseOnTime:=False
    
    ' === Simple notifications ===
    NotifyInfo "Info Toast", "This is an informational toast.", 3, "CallbackMacroExample", "Asterisk"
    NotifySuccess "Success Toast", "Operation completed successfully.", 3, "CallbackMacroExample", "Asterisk"
    NotifyWarning "Warning Toast", "Be careful with this action.", 3, "CallbackMacroExample", "Exclamation"
    NotifyError "Error Toast", "An error has occurred!", 5, "CallbackMacroExample", "Exclamation"
    NotifyCritical "Critical Toast", "Critical failure!", 7, "CallbackMacroExample", "Exclamation"
    
    ' === Progress toast demonstration ===
    Dim i As Long
    For i = 0 To 100 Step 10
        NotifyProgress "Progress Toast", "Processing item " & i & " of 100...", i, 0, "", "Asterisk"
        DoEvents
        Application.Wait Now + TimeValue("0:00:0.3")
    Next i
    
    ' === Queued notifications ===
    Dim j As Long
    For j = 1 To 3
        ShowToast "Queued Toast " & j, "This toast is queued.", "INFO", 2
    Next j
    
    ' === Demonstrate callback macro ===
    ShowToast "Callback Example", "This toast triggers a macro.", "SUCCESS", 3, "CallbackMacroExample", "Asterisk"
    
    ' === Show stats ===
    Debug.Print GetToastStats()
    
    MsgBox "Demo complete. Check the Immediate Window for queue stats."
End Sub

' Example callback macro
Public Sub CallbackMacroExample()
    Debug.Print "[CallbackMacroExample] Toast callback triggered at " & Now
End Sub

