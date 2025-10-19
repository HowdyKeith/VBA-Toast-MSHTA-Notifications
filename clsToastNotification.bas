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
' Purpose: Complete toast notification system with progress support
' Version: 6.0
' Features: Real-time updates, multiple delivery modes, smart positioning
'***************************************************************
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79

Private Const TOAST_WIDTH As Long = 380
Private Const TOAST_HEIGHT As Long = 160
Private Const TOAST_MARGIN As Long = 20

' === PROPERTIES ===
Private m_Title As String
Private m_Message As String
Private m_Icon As String
Private m_Position As String
Private m_Duration As Long
Private m_Level As String
Private m_LinkUrl As String
Private m_CallbackMacro As String
Private m_ImagePath As String
Private m_Progress As Long
Private m_IsProgress As Boolean

' === RUNTIME STATE ===
Private m_HTAPath As String
Private m_ProgressFile As String
Private m_IsShowing As Boolean
Private m_DeliveryMode As String  ' "HTA", "PS", "VBS"

' === INITIALIZATION ===
Private Sub Class_Initialize()
    m_Position = "BR"
    m_Duration = 5
    m_Level = "INFO"
    m_Progress = -1
    m_IsProgress = False
    m_IsShowing = False
End Sub

Private Sub Class_Terminate()
    If m_IsShowing Then Me.Close
End Sub

' === PUBLIC CONFIGURATION ===
Public Property Let Title(ByVal value As String): m_Title = value: End Property
Public Property Get Title() As String: Title = m_Title: End Property

Public Property Let Message(ByVal value As String): m_Message = value: End Property
Public Property Get Message() As String: Message = m_Message: End Property

Public Property Let Icon(ByVal value As String): m_Icon = value: End Property
Public Property Get Icon() As String
    If m_Icon <> "" Then
        Icon = m_Icon
    Else
        ' Auto-select icon
        Select Case UCase$(m_Level)
            Case "INFO": Icon = "?"
            Case "WARN", "WARNING": Icon = "?"
            Case "ERROR": Icon = "?"
            Case "SUCCESS": Icon = "?"
            Case "PROGRESS": Icon = "?"
            Case Else: Icon = "?"
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

Public Property Get IsShowing() As Boolean: IsShowing = m_IsShowing: End Property

' === PRIMARY METHODS ===
Public Function Show(Optional ByVal Mode As String = "auto") As Boolean
    On Error GoTo ErrorHandler
    
    ' Resolve delivery mode
    m_DeliveryMode = LCase$(Mode)
    If m_DeliveryMode = "auto" Then
        If IsListenerRunning() Then
            m_DeliveryMode = "ps"
        Else
            m_DeliveryMode = "hta"
        End If
    End If
    
    ' Deliver toast
    Select Case m_DeliveryMode
        Case "ps", "powershell"
            Show = ShowViaPowerShell()
        Case "vbs", "vbscript"
            Show = ShowViaVBScript()
        Case Else
            Show = ShowViaHTA()
    End Select
    
    If Show Then m_IsShowing = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[clsToastNotification] Show error: " & Err.description
    Show = False
End Function

Public Sub UpdateProgress(ByVal Percent As Long, Optional ByVal NewMessage As String = "")
    On Error Resume Next
    
    If Not m_IsProgress Then Exit Sub
    
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    m_Progress = Percent
    
    If NewMessage <> "" Then m_Message = NewMessage
    
    ' Update progress file
    If m_ProgressFile <> "" Then
        Dim json As String
        json = "{""Progress"":" & Percent & ","
        json = json & """Message"":""" & EscapeJson(m_Message) & ""","
        json = json & """Running"":" & IIf(Percent < 100, "true", "false") & "}"
        
        WriteFile m_ProgressFile, json
    End If
    
    ' Auto-close at 100%
    If Percent >= 100 And m_Duration > 0 Then
        Sleep m_Duration * 1000
        Me.Close
    End If
End Sub

Public Sub Close()
    On Error Resume Next
    
    ' Signal close
    If m_ProgressFile <> "" Then
        Dim json As String
        json = "{""Progress"":100,""Message"":""Complete"",""Running"":false}"
        WriteFile m_ProgressFile, json
        Sleep 500
        DeleteFile m_ProgressFile
    End If
    
    ' Delete HTA file
    If m_HTAPath <> "" Then
        DeleteFile m_HTAPath
    End If
    
    m_IsShowing = False
End Sub

' === DELIVERY IMPLEMENTATIONS ===
Private Function ShowViaPowerShell() As Boolean
    On Error GoTo ErrorHandler
    
    ' Create JSON request
    Dim json As String
    json = ToJson()
    
    ' Write to temp file for listener
    Dim requestFile As String
    requestFile = TempFolder() & "\ToastRequest.json"
    WriteFile requestFile, json
    
    ' Wait for processing (max 3 seconds)
    Dim i As Long
    For i = 1 To 30
        If Dir(requestFile) = "" Then
            ShowViaPowerShell = True
            Exit Function
        End If
        Sleep 100
    Next
    
    ' Timeout - listener might not be running
    ShowViaPowerShell = False
    Exit Function
    
ErrorHandler:
    Debug.Print "[ShowViaPowerShell] Error: " & Err.description
    ShowViaPowerShell = False
End Function

Private Function ShowViaHTA() As Boolean
    On Error GoTo ErrorHandler
    
    ' Generate HTA path
    m_HTAPath = TempFolder() & "\Toast_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 10000) & ".hta"
    
    ' Setup progress tracking if needed
    If m_IsProgress Then
        m_ProgressFile = TempFolder() & "\ToastProgress_" & Format(Now, "yyyymmddhhnnss") & ".json"
        Dim initJson As String
        initJson = "{""Progress"":" & m_Progress & ",""Message"":""" & EscapeJson(m_Message) & """,""Running"":true}"
        WriteFile m_ProgressFile, initJson
    End If
    
    ' Build HTML
    Dim html As String
    html = BuildHTML()
    
    ' Write HTA file
    WriteFile m_HTAPath, html
    
    ' Launch
    shell "mshta.exe """ & m_HTAPath & """", vbHide
    
    ShowViaHTA = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[ShowViaHTA] Error: " & Err.description
    ShowViaHTA = False
End Function

Private Function ShowViaVBScript() As Boolean
    On Error GoTo ErrorHandler
    
    ' Generate VBS wrapper (for complex progress scenarios)
    Dim vbsPath As String
    vbsPath = TempFolder() & "\Toast_" & Format(Now, "yyyymmddhhnnss") & ".vbs"
    
    m_HTAPath = Replace(vbsPath, ".vbs", ".hta")
    
    If m_IsProgress Then
        m_ProgressFile = TempFolder() & "\ToastProgress_" & Format(Now, "yyyymmddhhnnss") & ".json"
    End If
    
    ' Build VBS that monitors progress file and updates HTA
    Dim vbsCode As String
    vbsCode = BuildVBSWrapper()
    
    WriteFile vbsPath, vbsCode
    
    ' Launch VBS
    shell "wscript.exe """ & vbsPath & """", vbHide
    
    ShowViaVBScript = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[ShowViaVBScript] Error: " & Err.description
    ShowViaVBScript = False
End Function

' === HTML GENERATION ===
Private Function BuildHTML() As String
    Dim posX As Long, posY As Long
    GetToastPosition m_Position, posX, posY
    
    Dim bgColor As String, textColor As String
    GetThemeColors m_Level, bgColor, textColor
    
    Dim html As String
    html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" & vbCrLf
    html = html & "<title>" & EscapeHtml(m_Title) & "</title>" & vbCrLf
    html = html & "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no' " & _
                  "SYSMENU='no' SCROLL='no' SINGLEINSTANCE='yes'>" & vbCrLf
    
    ' === STYLES ===
    html = html & "<style>" & vbCrLf
    html = html & "* { margin:0; padding:0; box-sizing:border-box; }" & vbCrLf
    html = html & "body { font-family:'Segoe UI',system-ui,-apple-system,sans-serif; " & _
                  "background:transparent; overflow:hidden; }" & vbCrLf
    html = html & "#toast { position:fixed; width:" & TOAST_WIDTH & "px; " & _
                  "padding:16px; border-radius:12px; " & _
                  "background:" & bgColor & "; color:" & textColor & "; " & _
                  "box-shadow:0 8px 32px rgba(0,0,0,0.4); " & _
                  "backdrop-filter:blur(10px); " & _
                  "animation:slideIn 0.4s cubic-bezier(0.68,-0.55,0.265,1.55); }" & vbCrLf
    
    html = html & "@keyframes slideIn { from { transform:translateX(120%); opacity:0; } " & _
                  "to { transform:translateX(0); opacity:1; } }" & vbCrLf
    html = html & "@keyframes slideOut { from { transform:translateX(0); opacity:1; } " & _
                  "to { transform:translateX(120%); opacity:0; } }" & vbCrLf
    html = html & "@keyframes pulse { 0%,100% { transform:scale(1); } " & _
                  "50% { transform:scale(1.05); } }" & vbCrLf
    
    html = html & ".header { display:flex; align-items:center; margin-bottom:12px; }" & vbCrLf
    html = html & ".icon { font-size:28px; margin-right:12px; animation:pulse 2s infinite; }" & vbCrLf
    html = html & "h3 { font-size:18px; font-weight:600; flex:1; }" & vbCrLf
    html = html & ".close-btn { width:28px; height:28px; border:none; border-radius:6px; " & _
                  "background:rgba(0,0,0,0.2); color:" & textColor & "; " & _
                  "font-size:18px; font-weight:bold; cursor:pointer; " & _
                  "transition:all 0.2s; }" & vbCrLf
    html = html & ".close-btn:hover { background:rgba(0,0,0,0.4); transform:rotate(90deg); }" & vbCrLf
    html = html & ".message { font-size:14px; line-height:1.5; margin-bottom:12px; " & _
                  "opacity:0.95; }" & vbCrLf
    
    ' Progress bar styles
    If m_IsProgress Then
        html = html & ".progress-container { width:100%; height:24px; " & _
                      "background:rgba(0,0,0,0.2); border-radius:12px; " & _
                      "overflow:hidden; margin-top:12px; }" & vbCrLf
        html = html & ".progress-bar { height:100%; background:rgba(255,255,255,0.9); " & _
                      "border-radius:12px; transition:width 0.3s ease; " & _
                      "display:flex; align-items:center; justify-content:center; " & _
                      "font-size:12px; font-weight:600; color:#000; }" & vbCrLf
    End If
    
    ' Button styles
    If m_LinkUrl <> "" Or m_CallbackMacro <> "" Then
        html = html & ".btn { margin-top:12px; padding:8px 16px; border:none; " & _
                      "border-radius:6px; background:rgba(255,255,255,0.2); " & _
                      "color:" & textColor & "; font-size:13px; cursor:pointer; " & _
                      "transition:all 0.2s; margin-right:8px; }" & vbCrLf
        html = html & ".btn:hover { background:rgba(255,255,255,0.35); transform:translateY(-2px); }" & vbCrLf
    End If
    
    html = html & "</style>" & vbCrLf
    
    ' === BODY ===
    html = html & "</head><body>" & vbCrLf
    html = html & "<div id='toast'>" & vbCrLf
    html = html & "<div class='header'>" & vbCrLf
    html = html & "<span class='icon'>" & Me.Icon & "</span>" & vbCrLf
    html = html & "<h3>" & EscapeHtml(m_Title) & "</h3>" & vbCrLf
    html = html & "<button class='close-btn' onclick='closeToast()'>×</button>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<div class='message' id='message'>" & EscapeHtml(m_Message) & "</div>" & vbCrLf
    
    ' Progress bar
    If m_IsProgress Then
        html = html & "<div class='progress-container'>" & vbCrLf
        html = html & "<div class='progress-bar' id='progressBar' style='width:" & m_Progress & "%'>" & _
                      m_Progress & "%</div>" & vbCrLf
        html = html & "</div>" & vbCrLf
    End If
    
    ' Action buttons
    If m_LinkUrl <> "" Then
        html = html & "<button class='btn' onclick='openLink()'>Open Link</button>" & vbCrLf
    End If
    If m_CallbackMacro <> "" Then
        html = html & "<button class='btn' onclick='triggerCallback()'>Action</button>" & vbCrLf
    End If
    
    html = html & "</div>" & vbCrLf
    
    ' === JAVASCRIPT ===
    html = html & "<script>" & vbCrLf
    html = html & "window.resizeTo(" & (TOAST_WIDTH + 20) & "," & (TOAST_HEIGHT + 40) & ");" & vbCrLf
    html = html & "window.moveTo(" & posX & "," & posY & ");" & vbCrLf
    
    ' Auto-close timer
    If m_Duration > 0 And Not m_IsProgress Then
        html = html & "setTimeout(function(){ closeToast(); }," & (m_Duration * 1000) & ");" & vbCrLf
    End If
    
    ' Progress monitoring
    If m_IsProgress And m_ProgressFile <> "" Then
        html = html & "var progressFile = '" & Replace(m_ProgressFile, "\", "\\") & "';" & vbCrLf
        html = html & "function updateProgress() {" & vbCrLf
        html = html & "  try {" & vbCrLf
        html = html & "    var fso = new ActiveXObject('Scripting.FileSystemObject');" & vbCrLf
        html = html & "    if (fso.FileExists(progressFile)) {" & vbCrLf
        html = html & "      var file = fso.OpenTextFile(progressFile, 1);" & vbCrLf
        html = html & "      var json = file.ReadAll();" & vbCrLf
        html = html & "      file.Close();" & vbCrLf
        html = html & "      var data = JSON.parse(json);" & vbCrLf
        html = html & "      document.getElementById('progressBar').style.width = data.Progress + '%';" & vbCrLf
        html = html & "      document.getElementById('progressBar').innerText = data.Progress + '%';" & vbCrLf
        html = html & "      document.getElementById('message').innerText = data.Message;" & vbCrLf
        html = html & "      if (data.Running && data.Progress < 100) {" & vbCrLf
        html = html & "        setTimeout(updateProgress, 300);" & vbCrLf
        html = html & "      } else {" & vbCrLf
        html = html & "        setTimeout(function(){ closeToast(); }, 2000);" & vbCrLf
        html = html & "      }" & vbCrLf
        html = html & "    }" & vbCrLf
        html = html & "  } catch(e) {}" & vbCrLf
        html = html & "}" & vbCrLf
        html = html & "setTimeout(updateProgress, 500);" & vbCrLf
    End If
    
    ' Close function
    html = html & "function closeToast() {" & vbCrLf
    html = html & "  document.getElementById('toast').style.animation = 'slideOut 0.3s ease-in';" & vbCrLf
    html = html & "  setTimeout(function(){ window.close(); }, 300);" & vbCrLf
    html = html & "}" & vbCrLf
    
    ' Link handler
    If m_LinkUrl <> "" Then
        html = html & "function openLink() {" & vbCrLf
        html = html & "  var shell = new ActiveXObject('WScript.Shell');" & vbCrLf
        html = html & "  shell.Run('" & Replace(m_LinkUrl, "'", "\'") & "');" & vbCrLf
        html = html & "  closeToast();" & vbCrLf
        html = html & "}" & vbCrLf
    End If
    
    ' Callback handler
    If m_CallbackMacro <> "" Then
        html = html & "function triggerCallback() {" & vbCrLf
        html = html & "  var fso = new ActiveXObject('Scripting.FileSystemObject');" & vbCrLf
        html = html & "  var file = fso.CreateTextFile('" & TempFolder() & "\\ToastCallback.txt', true);" & vbCrLf
        html = html & "  file.Write('" & m_CallbackMacro & "');" & vbCrLf
        html = html & "  file.Close();" & vbCrLf
        html = html & "  closeToast();" & vbCrLf
        html = html & "}" & vbCrLf
    End If
    
    html = html & "</script>" & vbCrLf
    html = html & "</body></html>"
    
    BuildHTML = html
End Function

Private Function BuildVBSWrapper() As String
    ' VBS code for complex progress monitoring
    Dim vbs As String
    vbs = "' Auto-generated VBS wrapper for toast progress" & vbCrLf
    vbs = vbs & "' Generated: " & Now & vbCrLf
    ' TODO: Implement VBS monitoring code
    BuildVBSWrapper = vbs
End Function

' === UTILITIES ===
Private Sub GetToastPosition(ByVal pos As String, ByRef x As Long, ByRef y As Long)
    Dim screenW As Long, screenH As Long
    screenW = GetSystemMetrics(SM_CXSCREEN)
    screenH = GetSystemMetrics(SM_CYSCREEN)
    
    If screenW = 0 Then screenW = 1920
    If screenH = 0 Then screenH = 1080
    
    Select Case UCase$(pos)
        Case "TL": x = TOAST_MARGIN: y = TOAST_MARGIN
        Case "TR": x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = TOAST_MARGIN
        Case "BL": x = TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "BR": x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "CR": x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = (screenH - TOAST_HEIGHT) \ 2
        Case "C": x = (screenW - TOAST_WIDTH) \ 2: y = (screenH - TOAST_HEIGHT) \ 2
        Case Else: x = screenW - TOAST_WIDTH - TOAST_MARGIN: y = screenH - TOAST_HEIGHT - TOAST_MARGIN
    End Select
End Sub

Private Sub GetThemeColors(ByVal Level As String, ByRef bgColor As String, ByRef textColor As String)
    Select Case UCase$(Level)
        Case "WARN", "WARNING"
            bgColor = "linear-gradient(135deg, #FF9800 0%, #F57C00 100%)"
            textColor = "#000000"
        Case "ERROR"
            bgColor = "linear-gradient(135deg, #F44336 0%, #D32F2F 100%)"
            textColor = "#FFFFFF"
        Case "SUCCESS"
            bgColor = "linear-gradient(135deg, #4CAF50 0%, #388E3C 100%)"
            textColor = "#FFFFFF"
        Case "PROGRESS"
            bgColor = "linear-gradient(135deg, #2196F3 0%, #1976D2 100%)"
            textColor = "#FFFFFF"
        Case Else ' INFO
            bgColor = "linear-gradient(135deg, #00BCD4 0%, #0097A7 100%)"
            textColor = "#FFFFFF"
    End Select
End Sub

Private Function ToJson() As String
    Dim json As String
    json = "{"
    json = json & """Title"":""" & EscapeJson(m_Title) & ""","
    json = json & """Message"":""" & EscapeJson(m_Message) & ""","
    json = json & """Level"":""" & m_Level & ""","
    json = json & """Duration"":" & m_Duration & ","
    json = json & """Position"":""" & m_Position & ""","
    json = json & """LinkUrl"":""" & EscapeJson(m_LinkUrl) & ""","
    json = json & """CallbackMacro"":""" & EscapeJson(m_CallbackMacro) & ""","
    json = json & """Icon"":""" & EscapeJson(Me.Icon) & ""","
    json = json & """ImagePath"":""" & EscapeJson(m_ImagePath) & """"
    If m_IsProgress Then
        json = json & ",""Progress"":" & m_Progress
        If m_ProgressFile <> "" Then
            json = json & ",""ProgressFile"":""" & EscapeJson(m_ProgressFile) & """"
        End If
    End If
    json = json & "}"
    ToJson = json
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

Private Function TempFolder() As String
    Static temp As String
    If temp = "" Then
        temp = Environ$("TEMP") & "\ExcelToasts"
        If Dir(temp, vbDirectory) = "" Then MkDir temp
    End If
    TempFolder = temp
End Function

Private Sub WriteFile(ByVal filePath As String, ByVal Content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write Content
    ts.Close
End Sub

Private Sub DeleteFile(ByVal filePath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then fso.DeleteFile filePath
End Sub

Private Function IsListenerRunning() As Boolean
    On Error Resume Next
    Dim sentinelFile As String
    sentinelFile = TempFolder() & "\ListenerHeartbeat.txt"
    
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

'***************************************************************
' USAGE EXAMPLE (in a standard module):
'
' Sub TestToast()
'     Dim toast As New clsToastNotification
'     toast.Title = "File Upload"
'     toast.Message = "Uploading data..."
'     toast.Level = "PROGRESS"
'     toast.Position = "BR"
'     toast.ProgressValue = 0
'
'     If toast.Show() Then
'         Dim i As Long
'         For i = 0 To 100 Step 10
'             toast.UpdateProgress i, "Uploading: " & i & "%"
'             Application.Wait Now + TimeValue("00:00:01")
'         Next
'     End If
' End Sub
'***************************************************************

