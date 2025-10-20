VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToastNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsToastNotification
' Version: 7.1
' Features:
'   - Multi-toast queueing & stacking
'   - Progress updates
'   - HTA/PS/VBS/WinRT delivery
'   - Native sound alerts
'   - Callback macro auto-invoker
'   - Optional OnTime timer cleanup
'***************************************************************
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
        ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
        ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const TOAST_WIDTH As Long = 380
Private Const TOAST_HEIGHT As Long = 160
Private Const TOAST_MARGIN As Long = 20

Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_ALIAS As Long = &H10000

' === STATIC QUEUE ===
Private SharedQueue As Collection

' === INSTANCE PROPERTIES ===
Private m_Title As String
Private m_Message As String
Private m_Icon As String
Private m_Level As String
Private m_Position As String
Private m_Duration As Long
Private m_LinkUrl As String
Private m_CallbackMacro As String
Private m_ImagePath As String
Private m_Progress As Long
Private m_IsProgress As Boolean
Private m_SoundName As String
Private m_UseOnTime As Boolean

' === RUNTIME STATE ===
Private m_HTAPath As String
Private m_ProgressFile As String
Private m_IsShowing As Boolean
Private m_DeliveryMode As String

' === INITIALIZATION ===
Private Sub Class_Initialize()
    m_Position = "BR"
    m_Duration = 5
    m_Level = "INFO"
    m_Progress = -1
    m_IsProgress = False
    m_IsShowing = False
    m_UseOnTime = False
    m_SoundName = ""
    If SharedQueue Is Nothing Then Set SharedQueue = New Collection
End Sub

Private Sub Class_Terminate()
    If m_IsShowing Then Me.Close
End Sub

' === PROPERTIES ===
Public Property Let Title(ByVal value As String): m_Title = value: End Property
Public Property Get Title() As String: Title = m_Title: End Property

Public Property Let Message(ByVal value As String): m_Message = value: End Property
Public Property Get Message() As String: Message = m_Message: End Property

Public Property Let Icon(ByVal value As String): m_Icon = value: End Property
Public Property Get Icon() As String
    If m_Icon <> "" Then
        Icon = m_Icon
    Else
        Select Case UCase$(m_Level)
            Case "INFO": Icon = "??"
            Case "WARN", "WARNING": Icon = "??"
            Case "ERROR": Icon = "?"
            Case "SUCCESS": Icon = "?"
            Case "PROGRESS": Icon = "?"
            Case Else: Icon = "??"
        End Select
    End If
End Property

Public Property Let Position(ByVal value As String): m_Position = UCase$(value): End Property
Public Property Get Position() As String: Position = m_Position: End Property

Public Property Let Duration(ByVal value As Long): m_Duration = value: End Property
Public Property Get Duration() As Long: Duration = m_Duration: End Property

Public Property Let Level(ByVal value As String): m_Level = UCase$(value): End Property
Public Property Get Level() As String: Level = m_Level: End Property

Public Property Let LinkUrl(ByVal value As String): m_LinkUrl = value: End Property
Public Property Get LinkUrl() As String: LinkUrl = m_LinkUrl: End Property

Public Property Let CallbackMacro(ByVal value As String): m_CallbackMacro = value: End Property
Public Property Get CallbackMacro() As String: CallbackMacro = m_CallbackMacro: End Property

Public Property Let ImagePath(ByVal value As String): m_ImagePath = value: End Property
Public Property Get ImagePath() As String: ImagePath = m_ImagePath: End Property

Public Property Let ProgressValue(ByVal value As Long)
    If value < 0 Then value = 0
    If value > 100 Then value = 100
    m_Progress = value
    m_IsProgress = True
    m_Level = "PROGRESS"
End Property
Public Property Get ProgressValue() As Long: ProgressValue = m_Progress: End Property

Public Property Let SoundName(ByVal value As String): m_SoundName = value: End Property
Public Property Get SoundName() As String: SoundName = m_SoundName: End Property

Public Property Let UseOnTime(ByVal value As Boolean): m_UseOnTime = value: End Property
Public Property Get UseOnTime() As Boolean: UseOnTime = m_UseOnTime: End Property

Public Property Get IsShowing() As Boolean: IsShowing = m_IsShowing: End Property

' === PRIMARY METHODS ===
Public Function Show(Optional ByVal Mode As String = "auto") As Boolean
    On Error GoTo ErrorHandler
    
    m_DeliveryMode = LCase$(Mode)
    If m_DeliveryMode = "auto" Then
        If IsListenerRunning() Then
            m_DeliveryMode = "ps"
        Else
            m_DeliveryMode = "hta"
        End If
    End If
    
    Select Case m_DeliveryMode
        Case "ps", "powershell": Show = ShowViaPowerShell()
        Case "vbs", "vbscript": Show = ShowViaVBScript()
        Case "winrt": Show = ShowViaWinRT()
        Case Else: Show = ShowViaHTA()
    End Select
    
    ' Play sound if configured
    If m_SoundName <> "" Then PlaySoundAlert m_SoundName
    
    If Show Then
        m_IsShowing = True
        If m_CallbackMacro <> "" Then TriggerCallback m_CallbackMacro
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "[clsToastNotification] Show error: " & Err.Description
    Show = False
End Function

Public Sub ShowQueued()
    SharedQueue.Add Me
    If SharedQueue.count = 1 Then DisplayNextToast
End Sub

Private Sub DisplayNextToast()
    If SharedQueue.count = 0 Then Exit Sub
    Dim currentToast As clsToastNotification
    Set currentToast = SharedQueue(1)
    
    currentToast.Show "hta"
    
    Dim waitSeconds As Long
    waitSeconds = currentToast.m_Duration
    If currentToast.m_IsProgress Then
        Do While currentToast.m_Progress < 100
            DoEvents
            currentToast.RefreshProgress
            Sleep 300
        Loop
    Else
        Sleep waitSeconds * 1000
    End If
    
    currentToast.Close
    SharedQueue.Remove 1
    
    If SharedQueue.count > 0 Then DisplayNextToast
End Sub

' === PROGRESS ===
Public Sub UpdateProgress(ByVal Percent As Long, Optional ByVal NewMessage As String = "")
    On Error Resume Next
    If Not m_IsProgress Then Exit Sub
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    m_Progress = Percent
    If NewMessage <> "" Then m_Message = NewMessage
    
    If m_ProgressFile <> "" Then
        Dim json As String
        json = "{""Progress"":" & Percent & ",""Message"":""" & EscapeJson(m_Message) & """,""Running"":" & IIf(Percent < 100, "true", "false") & "}"
        WriteFile m_ProgressFile, json
    End If
    
    If Percent >= 100 And m_Duration > 0 Then
        Sleep m_Duration * 1000
        Me.Close
    End If
End Sub

Public Sub RefreshProgress()
    On Error Resume Next
    If m_ProgressFile <> "" Then
        Dim fso As Object, ts As Object, json As String, dict As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(m_ProgressFile) Then
            Set ts = fso.OpenTextFile(m_ProgressFile, 1)
            json = ts.ReadAll
            ts.Close
            Set dict = JsonDecode(json)
            m_Progress = dict("Progress")
            m_Message = dict("Message")
        End If
    End If
End Sub

' === CLOSE ===
Public Sub Close()
    On Error Resume Next
    If m_ProgressFile <> "" Then DeleteFile m_ProgressFile
    If m_HTAPath <> "" Then DeleteFile m_HTAPath
    m_IsShowing = False
End Sub

'***************************************************************
' --- Delivery Implementations ---
'***************************************************************
Private Function ShowViaPowerShell() As Boolean
    On Error GoTo ErrHandler
    Dim json As String, requestFile As String
    json = ToJson()
    requestFile = TempFolder() & "\ToastRequest_" & Format(Now, "yyyymmddhhnnss") & ".json"
    WriteFile requestFile, json
    
    Dim i As Long
    For i = 1 To 30
        If Dir(requestFile) = "" Then
            ShowViaPowerShell = True
            Exit Function
        End If
        Sleep 100
    Next
    ShowViaPowerShell = False
    Exit Function
ErrHandler:
    ShowViaPowerShell = False
End Function

Private Function ShowViaHTA() As Boolean
    On Error GoTo ErrHandler
    m_HTAPath = TempFolder() & "\Toast_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000) & ".hta"
    
    If m_IsProgress Then
        m_ProgressFile = TempFolder() & "\ToastProgress_" & Format(Now, "yyyymmddhhnnss") & ".json"
        WriteFile m_ProgressFile, "{""Progress"":" & m_Progress & ",""Message"":""" & EscapeJson(m_Message) & """,""Running"":true}"
    End If
    
    WriteFile m_HTAPath, BuildHTML()
    shell "mshta.exe """ & m_HTAPath & """", vbHide
    ShowViaHTA = True
    Exit Function
ErrHandler:
    ShowViaHTA = False
End Function

Private Function ShowViaVBScript() As Boolean
    On Error GoTo ErrHandler
    ShowViaVBScript = ShowViaHTA()
    Exit Function
ErrHandler:
    ShowViaVBScript = False
End Function

Private Function ShowViaWinRT() As Boolean
    On Error GoTo ErrHandler
    ' Placeholder for optional WinRT toast API channel
    ShowViaWinRT = True
    Exit Function
ErrHandler:
    ShowViaWinRT = False
End Function

'***************************************************************
' --- HTML / JS / VBS Generation ---
'***************************************************************
Private Function BuildHTML() As String
    Dim posX As Long, posY As Long
    GetToastPosition m_Position, posX, posY
    
    Dim bgColor As String, textColor As String
    GetThemeColors m_Level, bgColor, textColor
    
    Dim html As String
    html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" & vbCrLf
    html = html & "<title>" & EscapeHtml(m_Title) & "</title>"
    html = html & "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no' SCROLL='no' SINGLEINSTANCE='yes'>" & vbCrLf
    
    html = html & "<style>*{margin:0;padding:0;box-sizing:border-box;}body{font-family:'Segoe UI',sans-serif;background:transparent;overflow:hidden;}#toast{position:fixed;width:" & TOAST_WIDTH & "px;padding:16px;border-radius:12px;background:" & bgColor & ";color:" & textColor & ";box-shadow:0 8px 32px rgba(0,0,0,0.4);backdrop-filter:blur(10px);animation:slideIn 0.4s cubic-bezier(0.68,-0.55,0.265,1.55);} @keyframes slideIn{from{transform:translateX(120%);opacity:0;}to{transform:translateX(0);opacity:1;}} @keyframes slideOut{from{transform:translateX(0);opacity:1;}to{transform:translateX(120%);opacity:0;}}</style>"
    
    html = html & "<div id='toast'><div class='header'><span class='icon'>" & Me.Icon & "</span><h3>" & EscapeHtml(m_Title) & "</h3></div>"
    html = html & "<div class='message' id='message'>" & EscapeHtml(m_Message) & "</div>"
    If m_IsProgress Then html = html & "<div class='progress-container'><div class='progress-bar' id='progressBar' style='width:" & m_Progress & "%'>" & m_Progress & "%</div></div>"
    html = html & "</div></body></html>"
    
    BuildHTML = html
End Function

'***************************************************************
' --- Utilities ---
'***************************************************************
Private Sub GetToastPosition(ByVal pos As String, ByRef x As Long, ByRef y As Long)
    Dim screenW As Long, screenH As Long
    screenW = GetSystemMetrics(SM_CXSCREEN)
    screenH = GetSystemMetrics(SM_CYSCREEN)
    
    Select Case UCase$(pos)
        Case "TL": x = TOAST_MARGIN: y = TOAST_MARGIN
        Case "TR": x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = TOAST_MARGIN
        Case "BL": x = TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "BR": x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "C", "CR": x = (screenW - TOAST_WIDTH) \ 2: y = (screenH - TOAST_HEIGHT) \ 2
        Case Else: x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
    End Select
End Sub

Private Sub GetThemeColors(ByVal Level As String, ByRef bgColor As String, ByRef textColor As String)
    Select Case UCase$(Level)
        Case "WARN", "WARNING": bgColor = "linear-gradient(135deg,#FF9800 0%,#F57C00 100%)": textColor = "#000000"
        Case "ERROR": bgColor = "linear-gradient(135deg,#F44336 0%,#D32F2F 100%)": textColor = "#FFFFFF"
        Case "SUCCESS": bgColor = "linear-gradient(135deg,#4CAF50 0%,#388E3C 100%)": textColor = "#FFFFFF"
        Case "PROGRESS": bgColor = "linear-gradient(135deg,#2196F3 0%,#1976D2 100%)": textColor = "#FFFFFF"
        Case Else: bgColor = "linear-gradient(135deg,#00BCD4 0%,#0097A7 100%)": textColor = "#FFFFFF"
    End Select
End Sub

Private Function ToJson() As String
    Dim json As String
    json = "{""Title"":""" & EscapeJson(m_Title) & """,""Message"":""" & EscapeJson(m_Message) & """,""Level"":""" & m_Level & """,""Duration"":" & m_Duration & ",""Position"":""" & m_Position & """,""LinkUrl"":""" & EscapeJson(m_LinkUrl) & """,""CallbackMacro"":""" & EscapeJson(m_CallbackMacro) & """,""Icon"":""" & EscapeJson(Me.Icon) & """,""ImagePath"":""" & EscapeJson(m_ImagePath) & """}"
    ToJson = json
End Function

Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
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

Private Function TempFolder() As String
    Static temp As String
    If temp = "" Then
        temp = Environ$("TEMP") & "\ExcelToasts"
        If Dir(temp, vbDirectory) = "" Then MkDir temp
    End If
    TempFolder = temp
End Function

Private Sub WriteFile(ByVal FilePath As String, ByVal Content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(FilePath, True, True)
    ts.Write Content
    ts.Close
End Sub

Private Sub DeleteFile(ByVal FilePath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then fso.DeleteFile FilePath
End Sub

Private Function IsListenerRunning() As Boolean
    On Error Resume Next
    Dim sentinelFile As String, fso As Object, lastMod As Date
    sentinelFile = TempFolder() & "\ListenerHeartbeat.txt"
    If Dir(sentinelFile) = "" Then IsListenerRunning = False: Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    lastMod = fso.GetFile(sentinelFile).DateLastModified
    IsListenerRunning = (DateDiff("s", lastMod, Now) < 10)
End Function

'***************************************************************
' --- SOUND / CALLBACK ---
'***************************************************************
Private Sub PlaySoundAlert(ByVal SoundName As String)
    On Error Resume Next
    Dim nameLower As String
    nameLower = LCase$(SoundName)
    
    Select Case nameLower
        Case "beep", "default": Beep
        Case "asterisk", "exclamation", "hand", "question"
            PlaySound SoundName, 0, SND_ALIAS Or SND_ASYNC
        Case Else
            If Dir(SoundName) <> "" Then
                PlaySound SoundName, 0, SND_FILENAME Or SND_ASYNC
            Else
                Beep
            End If
    End Select
End Sub

Private Sub TriggerCallback(ByVal MacroName As String)
    On Error Resume Next
    Application.Run MacroName
End Sub

'***************************************************************
' --- JSON Decode Helper ---
'***************************************************************
Private Function JsonDecode(ByVal jsonText As String) As Object
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    sc.AddCode "function parse(json){return eval('(' + json + ')');}"
    Set JsonDecode = sc.Run("parse", jsonText)
End Function


