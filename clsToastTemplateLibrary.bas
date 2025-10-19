VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToastTemplateLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsToastTemplateLibrary
' Purpose: Pre-configured toast templates for common scenarios
' Version: 6.0
' Features: Built-in templates, custom templates, template persistence
'***************************************************************
Option Explicit

Private m_Templates As Object  ' Dictionary of templates
Private m_ConfigFile As String

' === INITIALIZATION ===
Private Sub Class_Initialize()
    Set m_Templates = CreateObject("Scripting.Dictionary")
    m_ConfigFile = Environ$("TEMP") & "\ExcelToasts\Templates.json"
    
    ' Load built-in templates
    LoadBuiltInTemplates
    
    ' Load custom templates from file
    LoadCustomTemplates
End Sub

' === BUILT-IN TEMPLATES ===
Private Sub LoadBuiltInTemplates()
    ' File Operations
    RegisterTemplate "FileUpload", "INFO", "?", 0, "File Upload", "Uploading {filename}...", "C"
    RegisterTemplate "FileDownload", "INFO", "?", 0, "File Download", "Downloading {filename}...", "C"
    RegisterTemplate "FileSaved", "SUCCESS", "?", 3, "File Saved", "{filename} saved successfully", "BR"
    RegisterTemplate "FileError", "ERROR", "?", 5, "File Error", "Failed to save {filename}", "TR"
    
    ' Data Operations
    RegisterTemplate "DataProcessing", "PROGRESS", "?", 0, "Processing Data", "Processing {count} records...", "BR"
    RegisterTemplate "DataComplete", "SUCCESS", "?", 4, "Processing Complete", "Successfully processed {count} records", "BR"
    RegisterTemplate "DataError", "ERROR", "?", 6, "Processing Error", "Error processing data: {error}", "TR"
    
    ' Network Operations
    RegisterTemplate "Connecting", "INFO", "??", 0, "Connecting", "Connecting to {server}...", "BR"
    RegisterTemplate "Connected", "SUCCESS", "?", 3, "Connected", "Connected to {server}", "BR"
    RegisterTemplate "Disconnected", "WARNING", "?", 4, "Disconnected", "Connection to {server} lost", "TR"
    RegisterTemplate "NetworkError", "ERROR", "?", 5, "Network Error", "Failed to connect: {error}", "TR"
    
    ' User Actions
    RegisterTemplate "UserLogin", "SUCCESS", "??", 3, "Welcome", "Welcome back, {username}!", "TR"
    RegisterTemplate "UserLogout", "INFO", "??", 2, "Goodbye", "Goodbye, {username}!", "TR"
    RegisterTemplate "PermissionDenied", "ERROR", "??", 5, "Access Denied", "You don't have permission to {action}", "C"
    
    ' System Status
    RegisterTemplate "LowMemory", "WARNING", "?", 0, "Low Memory", "Memory usage: {percent}%", "TR"
    RegisterTemplate "LowDiskSpace", "WARNING", "??", 0, "Low Disk Space", "Only {space} remaining on {drive}", "TR"
    RegisterTemplate "SystemError", "ERROR", "?", 0, "System Error", "An error occurred: {error}", "C"
    
    ' Validation
    RegisterTemplate "ValidationSuccess", "SUCCESS", "?", 3, "Validation Passed", "All {count} records are valid", "BR"
    RegisterTemplate "ValidationWarning", "WARNING", "?", 5, "Validation Warnings", "Found {count} warnings", "TR"
    RegisterTemplate "ValidationError", "ERROR", "?", 0, "Validation Failed", "Found {count} errors", "TR"
    
    ' API/External Services
    RegisterTemplate "APIRequest", "INFO", "??", 0, "API Request", "Sending request to {endpoint}...", "BR"
    RegisterTemplate "APISuccess", "SUCCESS", "?", 3, "API Success", "Request completed successfully", "BR"
    RegisterTemplate "APIError", "ERROR", "?", 5, "API Error", "{endpoint} returned error: {error}", "TR"
    RegisterTemplate "APIRateLimited", "WARNING", "?", 5, "Rate Limited", "Too many requests. Try again in {seconds}s", "C"
    
    ' Database Operations
    RegisterTemplate "DBConnecting", "INFO", "??", 0, "Database", "Connecting to database...", "BR"
    RegisterTemplate "DBQuery", "INFO", "??", 0, "Database Query", "Executing query...", "BR"
    RegisterTemplate "DBSuccess", "SUCCESS", "?", 3, "Query Complete", "Retrieved {count} records", "BR"
    RegisterTemplate "DBError", "ERROR", "?", 5, "Database Error", "Query failed: {error}", "TR"
    
    ' Email Operations
    RegisterTemplate "SendingEmail", "INFO", "??", 0, "Sending Email", "Sending email to {recipient}...", "BR"
    RegisterTemplate "EmailSent", "SUCCESS", "?", 3, "Email Sent", "Email sent to {recipient}", "BR"
    RegisterTemplate "EmailError", "ERROR", "?", 5, "Email Failed", "Failed to send email: {error}", "TR"
    
    ' Report Generation
    RegisterTemplate "GeneratingReport", "PROGRESS", "??", 0, "Generating Report", "Creating {reportname}...", "C"
    RegisterTemplate "ReportComplete", "SUCCESS", "?", 4, "Report Ready", "{reportname} is ready", "BR"
    RegisterTemplate "ReportError", "ERROR", "?", 5, "Report Failed", "Error generating report: {error}", "TR"
    
    ' Updates/Patches
    RegisterTemplate "UpdateAvailable", "INFO", "??", 0, "Update Available", "Version {version} is available", "TR"
    RegisterTemplate "Updating", "PROGRESS", "?", 0, "Updating", "Installing update...", "C"
    RegisterTemplate "UpdateComplete", "SUCCESS", "?", 4, "Update Complete", "Updated to version {version}", "BR"
    RegisterTemplate "UpdateError", "ERROR", "?", 5, "Update Failed", "Update failed: {error}", "TR"
    
    ' Task Management
    RegisterTemplate "TaskStarted", "INFO", "?", 0, "Task Started", "Started: {taskname}", "BR"
    RegisterTemplate "TaskProgress", "PROGRESS", "?", 0, "Task Progress", "{taskname}: {percent}%", "BR"
    RegisterTemplate "TaskComplete", "SUCCESS", "?", 3, "Task Complete", "{taskname} finished", "BR"
    RegisterTemplate "TaskCanceled", "WARNING", "?", 3, "Task Canceled", "{taskname} was canceled", "BR"
    RegisterTemplate "TaskError", "ERROR", "?", 5, "Task Failed", "{taskname} failed: {error}", "TR"
End Sub

' === TEMPLATE REGISTRATION ===
Private Sub RegisterTemplate(ByVal name As String, ByVal Level As String, _
                            ByVal Icon As String, ByVal Duration As Long, _
                            ByVal Title As String, ByVal MessageTemplate As String, _
                            ByVal Position As String)
    Dim template As Object
    Set template = CreateObject("Scripting.Dictionary")
    template("Level") = Level
    template("Icon") = Icon
    template("Duration") = Duration
    template("Title") = Title
    template("MessageTemplate") = MessageTemplate
    template("Position") = Position
    
    m_Templates(name) = template
End Sub

' === PUBLIC METHODS ===

' Get list of available templates
Public Function GetTemplateNames() As String()
    Dim names() As String
    ReDim names(0 To m_Templates.count - 1)
    
    Dim i As Long
    Dim key As Variant
    For Each key In m_Templates.Keys
        names(i) = CStr(key)
        i = i + 1
    Next
    
    GetTemplateNames = names
End Function

' Check if template exists
Public Function TemplateExists(ByVal name As String) As Boolean
    TemplateExists = m_Templates.Exists(name)
End Function

' Create toast from template
Public Function CreateFromTemplate(ByVal TemplateName As String, _
                                   ParamArray Parameters() As Variant) As clsToastNotification
    On Error GoTo ErrorHandler
    
    If Not m_Templates.Exists(TemplateName) Then
        Err.Raise 9999, "ToastTemplateLibrary", "Template '" & TemplateName & "' not found"
    End If
    
    Dim template As Object
    Set template = m_Templates(TemplateName)
    
    ' Create toast
    Dim toast As New clsToastNotification
    toast.Level = template("Level")
    toast.Icon = template("Icon")
    toast.Duration = template("Duration")
    toast.Title = template("Title")
    toast.Position = template("Position")
    
    ' Process message template with parameters
    Dim msg As String
    msg = template("MessageTemplate")
    
    ' Replace placeholders
    If UBound(Parameters) >= LBound(Parameters) Then
        msg = ProcessPlaceholders(msg, Parameters)
    End If
    
    toast.Message = msg
    
    Set CreateFromTemplate = toast
    Exit Function
    
ErrorHandler:
    Debug.Print "[ToastTemplateLibrary] Error: " & Err.description
    Set CreateFromTemplate = Nothing
End Function

' Quick show methods
Public Sub ShowTemplate(ByVal TemplateName As String, ParamArray Parameters() As Variant)
    On Error Resume Next
    
    Dim toast As clsToastNotification
    If UBound(Parameters) >= LBound(Parameters) Then
        Set toast = CreateFromTemplate(TemplateName, Parameters)
    Else
        Set toast = CreateFromTemplate(TemplateName)
    End If
    
    If Not toast Is Nothing Then
        toast.Show
    End If
End Sub

' Save custom template
Public Sub SaveCustomTemplate(ByVal name As String, _
                             ByVal Level As String, _
                             ByVal Icon As String, _
                             ByVal Duration As Long, _
                             ByVal Title As String, _
                             ByVal MessageTemplate As String, _
                             ByVal Position As String)
    RegisterTemplate name, Level, Icon, Duration, Title, MessageTemplate, Position
    SaveCustomTemplates
End Sub

' Delete custom template
Public Sub DeleteCustomTemplate(ByVal name As String)
    If m_Templates.Exists(name) Then
        m_Templates.Remove name
        SaveCustomTemplates
    End If
End Sub

' Get template info
Public Function GetTemplateInfo(ByVal name As String) As String
    If Not m_Templates.Exists(name) Then
        GetTemplateInfo = "Template not found"
        Exit Function
    End If
    
    Dim template As Object
    Set template = m_Templates(name)
    
    Dim info As String
    info = "Template: " & name & vbCrLf
    info = info & "Level: " & template("Level") & vbCrLf
    info = info & "Icon: " & template("Icon") & vbCrLf
    info = info & "Duration: " & template("Duration") & "s" & vbCrLf
    info = info & "Title: " & template("Title") & vbCrLf
    info = info & "Message: " & template("MessageTemplate") & vbCrLf
    info = info & "Position: " & template("Position")
    
    GetTemplateInfo = info
End Function

' === PRIVATE HELPERS ===

Private Function ProcessPlaceholders(ByVal msg As String, ByRef params() As Variant) As String
    ' Replace {param1}, {param2}, etc. with actual values
    ' Also support named placeholders if params are passed as "name=value"
    
    Dim i As Long
    For i = LBound(params) To UBound(params)
        If VarType(params(i)) = vbString Then
            Dim parts() As String
            If InStr(params(i), "=") > 0 Then
                ' Named parameter: "filename=test.xlsx"
                parts = Split(params(i), "=")
                If UBound(parts) = 1 Then
                    msg = Replace(msg, "{" & parts(0) & "}", parts(1))
                End If
            Else
                ' Positional parameter: replace {0}, {1}, etc.
                msg = Replace(msg, "{" & (i - LBound(params)) & "}", CStr(params(i)))
            End If
        Else
            ' Non-string parameter
            msg = Replace(msg, "{" & (i - LBound(params)) & "}", CStr(params(i)))
        End If
    Next
    
    ProcessPlaceholders = msg
End Function

Private Sub SaveCustomTemplates()
    ' Save custom templates to JSON file
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure folder exists
    Dim folder As String
    folder = fso.GetParentFolderName(m_ConfigFile)
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder folder
    End If
    
    ' Build JSON (simplified - in production use proper JSON library)
    Dim json As String
    json = "{""templates"": ["
    
    Dim key As Variant
    Dim first As Boolean: first = True
    For Each key In m_Templates.Keys
        Dim template As Object
        Set template = m_Templates(key)
        
        If Not first Then json = json & ","
        json = json & "{"
        json = json & """name"":""" & key & ""","
        json = json & """level"":""" & template("Level") & ""","
        json = json & """icon"":""" & template("Icon") & ""","
        json = json & """duration"":" & template("Duration") & ","
        json = json & """title"":""" & template("Title") & ""","
        json = json & """message"":""" & template("MessageTemplate") & ""","
        json = json & """position"":""" & template("Position") & """"
        json = json & "}"
        first = False
    Next
    
    json = json & "]}"
    
    ' Write file
    Dim ts As Object
    Set ts = fso.CreateTextFile(m_ConfigFile, True, True)
    ts.Write json
    ts.Close
End Sub

Private Sub LoadCustomTemplates()
    ' Load custom templates from JSON file
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(m_ConfigFile) Then Exit Sub
    
    ' Read file
    Dim ts As Object
    Set ts = fso.OpenTextFile(m_ConfigFile, 1)
    Dim json As String
    json = ts.ReadAll
    ts.Close
    
    ' Parse JSON (simplified - in production use proper JSON library)
    ' For now, just skip loading custom templates
    ' A full implementation would use a JSON parser
End Sub

'***************************************************************
' USAGE EXAMPLES:
'
' Sub Example1_BasicUsage()
'     Dim lib As New clsToastTemplateLibrary
'
'     ' Simple file upload notification
'     lib.ShowTemplate "FileUpload", "filename=report.xlsx"
'
'     ' Data processing with count
'     lib.ShowTemplate "DataProcessing", "count=1500"
'
'     ' Network connection
'     lib.ShowTemplate "Connected", "server=api.example.com"
' End Sub
'
' Sub Example2_WithProgress()
'     Dim lib As New clsToastTemplateLibrary
'     Dim toast As clsToastNotification
'
'     ' Start a task
'     Set toast = lib.CreateFromTemplate("TaskStarted", "taskname=Data Export")
'     toast.Show
'
'     ' Show progress
'     Set toast = lib.CreateFromTemplate("TaskProgress", "taskname=Data Export", "percent=50")
'     toast.ProgressValue = 50
'     toast.Show
'
'     ' Complete
'     lib.ShowTemplate "TaskComplete", "taskname=Data Export"
' End Sub
'
' Sub Example3_CustomTemplate()
'     Dim lib As New clsToastTemplateLibrary
'
'     ' Create custom template
'     lib.SaveCustomTemplate "MyCustom", "INFO", "??", 5, _
'         "Custom Alert", "Custom message: {details}", "C"
'
'     ' Use custom template
'     lib.ShowTemplate "MyCustom", "details=Something important!"
' End Sub
'
' Sub Example4_ListTemplates()
'     Dim lib As New clsToastTemplateLibrary
'     Dim names() As String
'     names = lib.GetTemplateNames()
'
'     Dim i As Long
'     For i = LBound(names) To UBound(names)
'         Debug.Print names(i) & ": " & lib.GetTemplateInfo(names(i))
'     Next
' End Sub
'***************************************************************

