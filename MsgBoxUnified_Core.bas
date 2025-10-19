Attribute VB_Name = "MsgBoxUnified_Core"
' =====================================================================
'  Module: MsgBoxUnified_Core
'  Version: 6.0
'  Purpose: Unified Notification System - Core Module
'  Author: Keith Swerling + ChatGPT
'  Changes: Consolidated all versions, added queue system, improved error handling
' =====================================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
' === VERSION & CONFIGURATION =========================================
Public Const MSGBOX_VERSION As String = "6.0"
Private Const DEFAULT_TIMEOUT As Long = 5
Private Const MAX_RETRIES As Long = 3
Private Const PIPE_TIMEOUT As Long = 3000 ' milliseconds

' === GLOBAL SETTINGS =================================================
Public UsePowerShellToasts As Boolean       ' Use PS listener if available
Public UseAutoFallback As Boolean           ' Auto-fallback to MSHTA/VBS
Public VerboseLogging As Boolean            ' Debug output
Public ToastPipeName As String              ' Named pipe path

' === PATH RESOLUTION =================================================
Private m_PSScriptPath As String
Private m_TempFolder As String

Public Property Get PSScriptPath() As String
    If m_PSScriptPath = "" Then
        ' Try multiple locations
        Dim paths As Variant
        paths = Array( _
            Environ$("USERPROFILE") & "\Documents\PowerShell\ToastWatcherRT.ps1", _
            Environ$("USERPROFILE") & "\OneDrive\Documents\PowerShell\ToastWatcherRT.ps1", _
            Environ$("APPDATA") & "\PowerShell\ToastWatcherRT.ps1" _
        )
        
        Dim path As Variant
        For Each path In paths
            If Dir(CStr(path)) <> "" Then
                m_PSScriptPath = CStr(path)
                Exit For
            End If
        Next
    End If
    PSScriptPath = m_PSScriptPath
End Property

Public Property Let PSScriptPath(ByVal value As String)
    m_PSScriptPath = value
End Property

Public Property Get TempFolder() As String
    If m_TempFolder = "" Then
        m_TempFolder = Environ$("TEMP") & "\ExcelToasts"
        If Dir(m_TempFolder, vbDirectory) = "" Then MkDir m_TempFolder
    End If
    TempFolder = m_TempFolder
End Property

' === INITIALIZATION ==================================================
Public Sub Initialize(Optional ByVal EnablePS As Boolean = True, _
                     Optional ByVal EnableFallback As Boolean = True, _
                     Optional ByVal PipeName As String = "")
    UsePowerShellToasts = EnablePS
    UseAutoFallback = EnableFallback
    If PipeName = "" Then
        ToastPipeName = "\\.\pipe\ExcelToastPipe"
    Else
        ToastPipeName = PipeName
    End If
    LogDebug "Initialized MsgBoxUnified v" & MSGBOX_VERSION
End Sub

' === MAIN API ========================================================
Public Function Notify(ByVal Title As String, _
                      ByVal Message As String, _
                      Optional ByVal Level As String = "INFO", _
                      Optional ByVal Timeout As Long = DEFAULT_TIMEOUT, _
                      Optional ByVal Position As String = "BR", _
                      Optional ByVal LinkUrl As String = "", _
                      Optional ByVal CallbackMacro As String = "", _
                      Optional ByVal Icon As String = "", _
                      Optional ByVal ImagePath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Build toast data
    Dim toast As ToastData
    toast.Title = Title
    toast.Message = Message
    toast.Level = UCase$(Level)
    toast.Timeout = Timeout
    toast.Position = UCase$(Position)
    toast.LinkUrl = LinkUrl
    toast.CallbackMacro = CallbackMacro
    toast.Icon = Icon
    toast.ImagePath = ImagePath
    toast.Progress = -1  ' No progress
    
    ' Try delivery with fallback chain
    If UsePowerShellToasts And IsListenerRunning() Then
        If DeliverViaPipe(toast) Then
            Notify = True
            Exit Function
        End If
    End If
    
    ' Fallback chain
    If UseAutoFallback Then
        If DeliverViaMSHTA(toast) Then
            Notify = True
            Exit Function
        End If
        
        ' Last resort: Classic MsgBox
        MsgBox Message, vbInformation, Title
        Notify = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "Notify failed: " & Err.description
    Notify = False
End Function

Public Function Progress(ByVal Title As String, _
                        ByVal Message As String, _
                        ByVal Percent As Long, _
                        Optional ByVal Position As String = "BR") As String
    
    On Error GoTo ErrorHandler
    
    ' Validate percent
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    
    Dim toast As ToastData
    toast.Title = Title
    toast.Message = Message & " [" & Percent & "%]"
    toast.Level = "PROGRESS"
    toast.Timeout = 0  ' Persistent
    toast.Position = UCase$(Position)
    toast.Progress = Percent
    
    ' Create progress file for tracking
    Dim ProgressFile As String
    ProgressFile = TempFolder & "\Progress_" & Format(Now, "yyyymmddhhnnss") & ".json"
    
    ' Write progress JSON
    Dim json As String
    json = ToastToJson(toast)
    WriteTextFile ProgressFile, json
    
    ' Deliver toast
    If UsePowerShellToasts And IsListenerRunning() Then
        If DeliverViaPipe(toast, ProgressFile) Then
            Progress = ProgressFile
            Exit Function
        End If
    End If
    
    ' Fallback to MSHTA with dynamic update
    If DeliverViaMSHTAProgress(toast, ProgressFile) Then
        Progress = ProgressFile
        Exit Function
    End If
    
    Progress = ""
    Exit Function
    
ErrorHandler:
    LogError "Progress failed: " & Err.description
    Progress = ""
End Function

Public Sub UpdateProgress(ByVal ProgressFile As String, _
                         ByVal Percent As Long, _
                         Optional ByVal Message As String = "")
    On Error Resume Next
    
    If Dir(ProgressFile) = "" Then Exit Sub
    
    ' Update JSON
    Dim json As String
    json = "{""Progress"":" & Percent & ","
    If Message <> "" Then
        json = json & """Message"":""" & EscapeJson(Message) & ""","
    End If
    json = json & """Running"":" & IIf(Percent < 100, "true", "false") & "}"
    
    WriteTextFile ProgressFile, json
End Sub

' === LISTENER CONTROL ================================================
Public Function StartListener() As Boolean
    On Error GoTo ErrorHandler
    
    If IsListenerRunning() Then
        LogDebug "Listener already running"
        StartListener = True
        Exit Function
    End If
    
    Dim psPath As String
    psPath = PSScriptPath
    
    If psPath = "" Then
        LogError "ToastWatcherRT.ps1 not found"
        StartListener = False
        Exit Function
    End If
    
    ' Launch listener
    Dim cmd As String
    cmd = "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psPath & """"
    shell cmd, vbHide
    
    ' Wait for startup (max 5 seconds)
    Dim i As Long
    For i = 1 To 50
        If IsListenerRunning() Then
            LogDebug "Listener started successfully"
            StartListener = True
            Exit Function
        End If
        Sleep 100
    Next
    
    LogError "Listener failed to start"
    StartListener = False
    Exit Function
    
ErrorHandler:
    LogError "StartListener error: " & Err.description
    StartListener = False
End Function

Public Function StopListener() As Boolean
    On Error Resume Next
    
    ' Send exit signal via flag file
    Dim exitFlag As String
    exitFlag = TempFolder & "\ExitListener.flag"
    WriteTextFile exitFlag, "EXIT"
    
    ' Wait for shutdown
    Dim i As Long
    For i = 1 To 30
        If Not IsListenerRunning() Then
            StopListener = True
            Exit Function
        End If
        Sleep 100
    Next
    
    ' Force kill if still running
    shell "taskkill /F /FI ""WINDOWTITLE eq *ToastWatcherRT*""", vbHide
    StopListener = True
End Function

Public Function IsListenerRunning() As Boolean
    On Error Resume Next
    
    ' Check sentinel file timestamp
    Dim sentinelFile As String
    sentinelFile = TempFolder & "\ListenerHeartbeat.txt"
    
    If Dir(sentinelFile) = "" Then
        IsListenerRunning = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim lastMod As Date
    lastMod = fso.GetFile(sentinelFile).DateLastModified
    
    ' Listener should update every 5 seconds
    IsListenerRunning = (DateDiff("s", lastMod, Now) < 10)
End Function

' === DELIVERY METHODS ================================================
Private Function DeliverViaPipe(ByRef toast As ToastData, _
                               Optional ByVal ProgressFile As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = ToastToJson(toast)
    If ProgressFile <> "" Then
        json = Left$(json, Len(json) - 1) & ",""ProgressFile"":""" & EscapeJson(ProgressFile) & """}"
    End If
    
    ' Try to write to pipe (with retry)
    Dim retries As Long
    For retries = 1 To MAX_RETRIES
        If WritePipe(ToastPipeName, json) Then
            LogDebug "Delivered via pipe (attempt " & retries & ")"
            DeliverViaPipe = True
            Exit Function
        End If
        Sleep 500
    Next
    
    LogDebug "Pipe delivery failed after " & MAX_RETRIES & " attempts"
    DeliverViaPipe = False
    Exit Function
    
ErrorHandler:
    LogError "DeliverViaPipe error: " & Err.description
    DeliverViaPipe = False
End Function

Private Function DeliverViaMSHTA(ByRef toast As ToastData) As Boolean
    On Error GoTo ErrorHandler
    
    ' Generate HTA file
    Dim htaPath As String
    htaPath = TempFolder & "\Toast_" & Format(Now, "yyyymmddhhnnss") & ".hta"
    
    Dim html As String
    html = BuildToastHTA(toast)
    
    WriteTextFile htaPath, html
    
    ' Launch MSHTA
    shell "mshta.exe """ & htaPath & """", vbHide
    
    ' Schedule cleanup
    ScheduleCleanup htaPath, toast.Timeout + 2
    
    LogDebug "Delivered via MSHTA: " & htaPath
    DeliverViaMSHTA = True
    Exit Function
    
ErrorHandler:
    LogError "DeliverViaMSHTA error: " & Err.description
    DeliverViaMSHTA = False
End Function

Private Function DeliverViaMSHTAProgress(ByRef toast As ToastData, _
                                        ByVal ProgressFile As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Generate dynamic progress HTA
    Dim htaPath As String
    htaPath = TempFolder & "\ProgressToast_" & Format(Now, "yyyymmddhhnnss") & ".hta"
    
    Dim html As String
    html = BuildProgressHTA(toast, ProgressFile)
    
    WriteTextFile htaPath, html
    
    shell "mshta.exe """ & htaPath & """", vbHide
    
    LogDebug "Delivered progress via MSHTA: " & htaPath
    DeliverViaMSHTAProgress = True
    Exit Function
    
ErrorHandler:
    LogError "DeliverViaMSHTAProgress error: " & Err.description
    DeliverViaMSHTAProgress = False
End Function

' === HTML BUILDERS ===================================================
Private Function BuildToastHTA(ByRef toast As ToastData) As String
    Dim html As String
    Dim bgColor As String, textColor As String, iconChar As String
    
    ' Theme colors
    Select Case toast.Level
        Case "WARN", "WARNING"
            bgColor = "linear-gradient(135deg, #ffeb3b, #ffa000)"
            textColor = "#000000"
            iconChar = "?"
        Case "ERROR"
            bgColor = "linear-gradient(135deg, #ff6b6b, #d32f2f)"
            textColor = "#ffffff"
            iconChar = "?"
        Case Else
            bgColor = "linear-gradient(135deg, #4caf50, #2e7d32)"
            textColor = "#ffffff"
            iconChar = "?"
    End Select
    
    If toast.Icon <> "" Then iconChar = toast.Icon
    
    ' Calculate position
    Dim posX As Long, posY As Long
    CalculatePosition toast.Position, posX, posY
    
    ' Build HTML
    html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" & vbCrLf
    html = html & "<title>" & EscapeHtml(toast.Title) & "</title>" & vbCrLf
    html = html & "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no' " & _
                  "SYSMENU='no' SCROLL='no' SINGLEINSTANCE='no'>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body{margin:0;padding:12px;font-family:'Segoe UI',Arial;background:" & _
                  bgColor & ";color:" & textColor & ";border-radius:8px;" & _
                  "box-shadow:0 4px 16px rgba(0,0,0,0.4);animation:slideIn 0.4s ease-out;}" & vbCrLf
    html = html & "@keyframes slideIn{from{transform:translateX(100%);opacity:0;}" & _
                  "to{transform:translateX(0);opacity:1;}}" & vbCrLf
    html = html & "@keyframes slideOut{from{transform:translateX(0);opacity:1;}" & _
                  "to{transform:translateX(100%);opacity:0;}}" & vbCrLf
    html = html & "h3{margin:0 0 8px;font-size:18px;font-weight:600;}" & vbCrLf
    html = html & ".icon{font-size:24px;margin-right:8px;}" & vbCrLf
    html = html & "p{margin:0;font-size:14px;line-height:1.4;}" & vbCrLf
    html = html & "button{margin-top:10px;padding:6px 12px;border:none;border-radius:4px;" & _
                  "background:rgba(255,255,255,0.2);color:" & textColor & ";" & _
                  "cursor:pointer;font-size:12px;}" & vbCrLf
    html = html & "button:hover{background:rgba(255,255,255,0.3);}" & vbCrLf
    html = html & "</style>" & vbCrLf
    html = html & "<script>" & vbCrLf
    html = html & "window.resizeTo(370,170);" & vbCrLf
    html = html & "window.moveTo(" & posX & "," & posY & ");" & vbCrLf
    If toast.Timeout > 0 Then
        html = html & "setTimeout(function(){" & _
                      "document.body.style.animation='slideOut 0.3s ease-in';" & _
                      "setTimeout(function(){window.close();},300);" & _
                      "}," & (toast.Timeout * 1000) & ");" & vbCrLf
    End If
    If toast.LinkUrl <> "" Then
        html = html & "function openLink(){window.open('" & EscapeJs(toast.LinkUrl) & "');}" & vbCrLf
    End If
    html = html & "function dismiss(){" & _
                  "document.body.style.animation='slideOut 0.3s ease-in';" & _
                  "setTimeout(function(){window.close();},300);}" & vbCrLf
    html = html & "</script></head><body>" & vbCrLf
    html = html & "<h3><span class='icon'>" & iconChar & "</span>" & _
                  EscapeHtml(toast.Title) & "</h3>" & vbCrLf
    html = html & "<p>" & EscapeHtml(toast.Message) & "</p>" & vbCrLf
    If toast.LinkUrl <> "" Then
        html = html & "<button onclick='openLink()'>Open Link</button>" & vbCrLf
    End If
    html = html & "<button onclick='dismiss()'>Dismiss</button>" & vbCrLf
    html = html & "</body></html>"
    
    BuildToastHTA = html
End Function

Private Function BuildProgressHTA(ByRef toast As ToastData, _
                                 ByVal ProgressFile As String) As String
    ' TODO: Implement dynamic progress HTA with JavaScript polling
    BuildProgressHTA = BuildToastHTA(toast) ' Placeholder
End Function

' === UTILITY FUNCTIONS ===============================================
Private Type ToastData
    Title As String
    Message As String
    Level As String
    Timeout As Long
    Position As String
    LinkUrl As String
    CallbackMacro As String
    Icon As String
    ImagePath As String
    Progress As Long
End Type

Private Function ToastToJson(ByRef toast As ToastData) As String
    Dim json As String
    json = "{"
    json = json & """Title"":""" & EscapeJson(toast.Title) & ""","
    json = json & """Message"":""" & EscapeJson(toast.Message) & ""","
    json = json & """Level"":""" & toast.Level & ""","
    json = json & """Timeout"":" & toast.Timeout & ","
    json = json & """Position"":""" & toast.Position & ""","
    json = json & """LinkUrl"":""" & EscapeJson(toast.LinkUrl) & ""","
    json = json & """CallbackMacro"":""" & EscapeJson(toast.CallbackMacro) & ""","
    json = json & """Icon"":""" & EscapeJson(toast.Icon) & ""","
    json = json & """ImagePath"":""" & EscapeJson(toast.ImagePath) & """"
    If toast.Progress >= 0 Then
        json = json & ",""Progress"":" & toast.Progress
    End If
    json = json & "}"
    ToastToJson = json
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

Private Function EscapeHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    EscapeHtml = s
End Function

Private Function EscapeJs(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    EscapeJs = s
End Function

Private Sub CalculatePosition(ByVal pos As String, ByRef x As Long, ByRef y As Long)
    Const MARGIN As Long = 20
    Const TOAST_W As Long = 350
    Const TOAST_H As Long = 150
    
    ' Get screen dimensions (simplified - could use API)
    Dim screenW As Long: screenW = 1920
    Dim screenH As Long: screenH = 1080
    
    Select Case pos
        Case "TL": x = MARGIN: y = MARGIN
        Case "TR": x = screenW - TOAST_W - MARGIN: y = MARGIN
        Case "BL": x = MARGIN: y = screenH - TOAST_H - MARGIN
        Case "BR": x = screenW - TOAST_W - MARGIN: y = screenH - TOAST_H - MARGIN
        Case "CR": x = screenW - TOAST_W - MARGIN: y = (screenH - TOAST_H) / 2
        Case "C": x = (screenW - TOAST_W) / 2: y = (screenH - TOAST_H) / 2
        Case Else: x = screenW - TOAST_W - MARGIN: y = screenH - TOAST_H - MARGIN
    End Select
End Sub

Private Function WritePipe(ByVal PipeName As String, ByVal Data As String) As Boolean
    ' TODO: Implement Win32 pipe writing
    ' For now, use fallback to temp file
    Dim requestFile As String
    requestFile = TempFolder & "\ToastRequest.json"
    WriteTextFile requestFile, Data
    
    ' Wait for processing
    Dim i As Long
    For i = 1 To 30
        If Dir(requestFile) = "" Then
            WritePipe = True
            Exit Function
        End If
        Sleep 100
    Next
    
    WritePipe = False
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal Content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write Content
    ts.Close
End Sub

Private Sub ScheduleCleanup(ByVal filePath As String, ByVal DelaySeconds As Long)
    ' TODO: Implement async cleanup
    ' For now, files will be cleaned up by explicit CleanupTempFiles call
End Sub

Private Sub LogDebug(ByVal msg As String)
    If VerboseLogging Then Debug.Print "[MsgBoxUnified] " & msg
End Sub

Private Sub LogError(ByVal msg As String)
    Debug.Print "[MsgBoxUnified ERROR] " & msg
End Sub


' === PUBLIC UTILITIES ================================================
Public Sub CleanupTempFiles()
    On Error Resume Next
    Dim fso As Object, folder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(TempFolder)
    
    For Each file In folder.files
        If DateDiff("h", file.DateLastModified, Now) > 24 Then
            file.Delete True
        End If
    Next
End Sub

Public Function GetStatus() As String
    Dim status As String
    status = "MsgBoxUnified v" & MSGBOX_VERSION & vbCrLf
    status = status & "Listener Running: " & IsListenerRunning() & vbCrLf
    status = status & "PS Mode Enabled: " & UsePowerShellToasts & vbCrLf
    status = status & "Auto Fallback: " & UseAutoFallback & vbCrLf
    status = status & "PS Script: " & PSScriptPath & vbCrLf
    status = status & "Temp Folder: " & TempFolder & vbCrLf
    status = status & "Pipe Name: " & ToastPipeName
    GetStatus = status
End Function

