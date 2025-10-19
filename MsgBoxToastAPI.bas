Attribute VB_Name = "MsgBoxToastAPI"
'***************************************************************
' Module: MsgBoxToastAPI
' Purpose: Unified API for Toast Notification System v6.0
' Author: Keith Swerling + ChatGPT
' Description: Complete notification system with multiple delivery modes,
'              queue management, templates, analytics, and visual builder
'***************************************************************
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'***************************************************************
' USAGE EXAMPLES:
'
' Sub Example1_QuickStart()
'     ' Simple notifications
'     NotifyInfo "Hello", "Welcome to Toast System!"
'     NotifySuccess "Saved", "File saved successfully"
'     NotifyWarning "Warning", "Disk space low"
'     NotifyError "Error", "Network connection failed"
' End Sub
'
' Sub Example2_Progress()
'     Dim progress As clsToastNotification
'     Set progress = ShowProgress("Upload", "Uploading file...", 0)
'
'     Dim i As Long
'     For i = 0 To 100 Step 10
'         UpdateProgress progress, i, "Uploading: " & i & "%"
'         Application.Wait Now + TimeValue("00:00:01")
'     Next
' End Sub
'
' Sub Example3_Templates()
'     ' Use built-in templates
'     ShowFromTemplate "FileUpload", "filename=Report.xlsx"
'     ShowFromTemplate "DataProcessing", "count=5000"
'     ShowFromTemplate "Connected", "server=api.company.com"
' End Sub
'
' Sub Example4_FullDemo()
'     ShowMainMenu  ' Interactive menu
' End Sub
'***************************************************************


' === GLOBAL INSTANCES ===
Private m_QueueManager As clsToastQueueManager
Private m_TemplateLibrary As clsToastTemplateLibrary
Private m_Analytics As clsToastAnalytics
Private m_Initialized As Boolean

' === VERSION INFO ===
Public Const TOAST_API_VERSION As String = "6.0"
Public Const TOAST_API_DATE As String = "2025-10-19"

' === INITIALIZATION ===
Public Sub InitializeToastSystem(Optional ByVal EnableQueue As Boolean = True, _
                                Optional ByVal EnableAnalytics As Boolean = True, _
                                Optional ByVal EnableDND As Boolean = False)
    
    If m_Initialized Then Exit Sub
    
    ' Initialize queue manager
    If EnableQueue Then
        Set m_QueueManager = New clsToastQueueManager
        m_QueueManager.MaxVisibleToasts = 5
        m_QueueManager.DefaultPosition = "BR"
        m_QueueManager.AutoStack = True
        m_QueueManager.DoNotDisturb = EnableDND
    End If
    
    ' Initialize template library
    Set m_TemplateLibrary = New clsToastTemplateLibrary
    
    ' Initialize analytics
    If EnableAnalytics Then
        Set m_Analytics = New clsToastAnalytics
        m_Analytics.EnableAnalytics = True
        m_Analytics.LogToFile = False
        m_Analytics.MaxHistorySize = 1000
    End If
    
    m_Initialized = True
    
    Debug.Print "[ToastAPI] System initialized v" & TOAST_API_VERSION
End Sub

Public Sub ShutdownToastSystem()
    If Not m_Initialized Then Exit Sub
    
    ' Cleanup
    If Not m_QueueManager Is Nothing Then
        m_QueueManager.ClearAll
        Set m_QueueManager = Nothing
    End If
    
    Set m_TemplateLibrary = Nothing
    Set m_Analytics = Nothing
    
    m_Initialized = False
    
    Debug.Print "[ToastAPI] System shutdown"
End Sub

' === QUICK NOTIFICATION METHODS ===

' Show a simple info notification
Public Sub NotifyInfo(ByVal Title As String, ByVal Message As String, _
                     Optional ByVal Duration As Long = 5, _
                     Optional ByVal Position As String = "BR")
    EnsureInitialized
    
    If m_QueueManager Is Nothing Then
        ' Direct show
        ShowQuickToast Title, Message, "INFO", Duration, Position
    Else
        ' Use queue
        m_QueueManager.NotifyInfo Title, Message, PriorityNormal
    End If
End Sub

' Show a success notification
Public Sub NotifySuccess(ByVal Title As String, ByVal Message As String, _
                        Optional ByVal Duration As Long = 4, _
                        Optional ByVal Position As String = "BR")
    EnsureInitialized
    
    If m_QueueManager Is Nothing Then
        ShowQuickToast Title, Message, "SUCCESS", Duration, Position
    Else
        m_QueueManager.NotifySuccess Title, Message, PriorityNormal
    End If
End Sub

' Show a warning notification
Public Sub NotifyWarning(ByVal Title As String, ByVal Message As String, _
                        Optional ByVal Duration As Long = 6, _
                        Optional ByVal Position As String = "TR")
    EnsureInitialized
    
    If m_QueueManager Is Nothing Then
        ShowQuickToast Title, Message, "WARNING", Duration, Position
    Else
        m_QueueManager.NotifyWarning Title, Message, PriorityHigh
    End If
End Sub

' Show an error notification
Public Sub NotifyError(ByVal Title As String, ByVal Message As String, _
                      Optional ByVal Duration As Long = 0, _
                      Optional ByVal Position As String = "C")
    EnsureInitialized
    
    If m_QueueManager Is Nothing Then
        ShowQuickToast Title, Message, "ERROR", Duration, Position
    Else
        m_QueueManager.NotifyError Title, Message, PriorityCritical
    End If
End Sub

' === PROGRESS NOTIFICATIONS ===

' Show progress toast and return handle for updates
Public Function ShowProgress(ByVal Title As String, _
                            ByVal Message As String, _
                            Optional ByVal InitialPercent As Long = 0, _
                            Optional ByVal Position As String = "BR") As clsToastNotification
    EnsureInitialized
    
    Dim toast As New clsToastNotification
    toast.Title = Title
    toast.Message = Message
    toast.Level = "PROGRESS"
    toast.Position = Position
    toast.Duration = 0  ' Persistent
    toast.ProgressValue = InitialPercent
    
    ' Track in analytics
    Dim startTime As Double
    startTime = Timer
    
    If toast.Show() Then
        If Not m_Analytics Is Nothing Then
            Dim deliveryTime As Long
            deliveryTime = CLng((Timer - startTime) * 1000)
            m_Analytics.LogToast Title, Message, "PROGRESS", Position, 0, "HTA", deliveryTime, True
        End If
        
        Set ShowProgress = toast
    Else
        Set ShowProgress = Nothing
    End If
End Function

' Update existing progress toast
Public Sub UpdateProgress(ByRef toast As clsToastNotification, _
                         ByVal Percent As Long, _
                         Optional ByVal NewMessage As String = "")
    If toast Is Nothing Then Exit Sub
    toast.UpdateProgress Percent, NewMessage
End Sub

' === TEMPLATE-BASED NOTIFICATIONS ===

' Show notification from template
Public Sub ShowFromTemplate(ByVal TemplateName As String, ParamArray Parameters() As Variant)
    EnsureInitialized
    
    If m_TemplateLibrary Is Nothing Then
        MsgBox "Template library not initialized", vbExclamation
        Exit Sub
    End If
    
    Dim toast As clsToastNotification
    If UBound(Parameters) >= LBound(Parameters) Then
        Set toast = m_TemplateLibrary.CreateFromTemplate(TemplateName, Parameters)
    Else
        Set toast = m_TemplateLibrary.CreateFromTemplate(TemplateName)
    End If
    
    If Not toast Is Nothing Then
        Dim startTime As Double
        startTime = Timer
        
        Dim success As Boolean
        success = toast.Show()
        
        If Not m_Analytics Is Nothing Then
            Dim deliveryTime As Long
            deliveryTime = CLng((Timer - startTime) * 1000)
            m_Analytics.LogToast toast.Title, toast.Message, toast.Level, _
                                 toast.Position, toast.Duration, "HTA", _
                                 deliveryTime, success
        End If
    End If
End Sub

' === ADVANCED FEATURES ===

' Open visual toast builder
Public Sub ShowToastBuilder()
    On Error Resume Next
    frmToastBuilder.Show
    If Err.Number <> 0 Then
        MsgBox "Toast Builder form not found. Please create frmToastBuilder UserForm.", _
               vbExclamation, "Error"
    End If
End Sub

' Show analytics dashboard
Public Sub ShowAnalyticsDashboard()
    EnsureInitialized
    
    If m_Analytics Is Nothing Then
        MsgBox "Analytics not enabled", vbExclamation
        Exit Sub
    End If
    
    Dim report As String
    report = "+--------------------------------------------+" & vbCrLf
    report = report & "¦   TOAST NOTIFICATION ANALYTICS DASHBOARD   ¦" & vbCrLf
    report = report & "+--------------------------------------------+" & vbCrLf & vbCrLf
    
    report = report & m_Analytics.GetStatisticsReport() & vbCrLf & vbCrLf
    report = report & String(50, "-") & vbCrLf & vbCrLf
    report = report & m_Analytics.GetPerformanceMetrics() & vbCrLf & vbCrLf
    report = report & String(50, "-") & vbCrLf & vbCrLf
    report = report & m_Analytics.GetRecentHistory(5)
    
    ' Show in message box or custom form
    MsgBox report, vbInformation, "Analytics Dashboard"
End Sub

' Export analytics to CSV
Public Sub ExportAnalytics(Optional ByVal filePath As String = "")
    EnsureInitialized
    
    If m_Analytics Is Nothing Then
        MsgBox "Analytics not enabled", vbExclamation
        Exit Sub
    End If
    
    If filePath = "" Then
        filePath = Environ$("USERPROFILE") & "\Desktop\ToastAnalytics_" & _
                   Format(Now, "yyyymmdd_hhnnss") & ".csv"
    End If
    
    If m_Analytics.ExportToCSV(filePath) Then
        MsgBox "Analytics exported to:" & vbCrLf & filePath, vbInformation, "Export Complete"
    Else
        MsgBox "Failed to export analytics", vbExclamation, "Export Failed"
    End If
End Sub

' List available templates
Public Sub ListTemplates()
    EnsureInitialized
    
    If m_TemplateLibrary Is Nothing Then
        MsgBox "Template library not initialized", vbExclamation
        Exit Sub
    End If
    
    Dim names() As String
    names = m_TemplateLibrary.GetTemplateNames()
    
    Dim list As String
    list = "--- Available Templates ---" & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = LBound(names) To UBound(names)
        list = list & "• " & names(i) & vbCrLf
    Next
    
    list = list & vbCrLf & "Total: " & (UBound(names) - LBound(names) + 1) & " templates"
    
    MsgBox list, vbInformation, "Template Library"
End Sub

' Queue management
Public Sub EnableDoNotDisturb(Optional ByVal DurationMinutes As Long = 60)
    EnsureInitialized
    
    If Not m_QueueManager Is Nothing Then
        m_QueueManager.EnableDND DurationMinutes
        MsgBox "Do Not Disturb enabled for " & DurationMinutes & " minutes", _
               vbInformation, "DND Mode"
    End If
End Sub

Public Sub DisableDoNotDisturb()
    EnsureInitialized
    
    If Not m_QueueManager Is Nothing Then
        m_QueueManager.DisableDND
        MsgBox "Do Not Disturb disabled", vbInformation, "DND Mode"
    End If
End Sub

Public Sub ShowQueueStatus()
    EnsureInitialized
    
    If m_QueueManager Is Nothing Then
        MsgBox "Queue manager not enabled", vbExclamation
        Exit Sub
    End If
    
    MsgBox m_QueueManager.GetStats(), vbInformation, "Queue Status"
End Sub

' === SYSTEM INFO ===
Public Function GetSystemInfo() As String
    Dim info As String
    info = "+--------------------------------------------+" & vbCrLf
    info = info & "¦   TOAST NOTIFICATION SYSTEM v" & TOAST_API_VERSION & "      ¦" & vbCrLf
    info = info & "+--------------------------------------------+" & vbCrLf & vbCrLf
    
    info = info & "Status: " & IIf(m_Initialized, "? Initialized", "? Not Initialized") & vbCrLf
    info = info & "Version: " & TOAST_API_VERSION & vbCrLf
    info = info & "Release Date: " & TOAST_API_DATE & vbCrLf & vbCrLf
    
    info = info & "Components:" & vbCrLf
    info = info & "  Queue Manager: " & IIf(Not m_QueueManager Is Nothing, "? Enabled", "? Disabled") & vbCrLf
    info = info & "  Template Library: " & IIf(Not m_TemplateLibrary Is Nothing, "? Enabled", "? Disabled") & vbCrLf
    info = info & "  Analytics: " & IIf(Not m_Analytics Is Nothing, "? Enabled", "? Disabled") & vbCrLf
    
    If Not m_QueueManager Is Nothing Then
        info = info & vbCrLf & "Queue Status:" & vbCrLf
        info = info & "  Queued: " & m_QueueManager.QueueLength & vbCrLf
        info = info & "  Visible: " & m_QueueManager.visibleCount & vbCrLf
        info = info & "  DND Mode: " & IIf(m_QueueManager.DoNotDisturb, "ON", "OFF") & vbCrLf
    End If
    
    If Not m_Analytics Is Nothing Then
        info = info & vbCrLf & "Analytics:" & vbCrLf
        info = info & "  History Entries: " & m_Analytics.HistoryCount & vbCrLf
    End If
    
    GetSystemInfo = info
End Function

' === TEST SUITE ===
Public Sub RunTestSuite()
    MsgBox "Starting Toast Notification Test Suite...", vbInformation, "Test Suite"
    
    InitializeToastSystem EnableQueue:=True, EnableAnalytics:=True
    
    ' Test 1: Basic notifications
    NotifyInfo "Test 1", "Basic info notification"
    Sleep 1000
    
    NotifySuccess "Test 2", "Success notification"
    Sleep 1000
    
    NotifyWarning "Test 3", "Warning notification"
    Sleep 1000
    
    NotifyError "Test 4", "Error notification (5 sec auto-close)", Duration:=5
    Sleep 1000
    
    ' Test 2: Progress notification
    Dim progressToast As clsToastNotification
    Set progressToast = ShowProgress("Test 5", "Progress test", 0, "C")
    
    If Not progressToast Is Nothing Then
        Dim i As Long
        For i = 0 To 100 Step 20
            UpdateProgress progressToast, i, "Processing: " & i & "%"
            Sleep 500
        Next
        progressToast.Close
    End If
    
    ' Test 3: Template-based
    ShowFromTemplate "FileUpload", "filename=TestFile.xlsx"
    Sleep 1000
    
    ShowFromTemplate "TaskComplete", "taskname=Test Suite"
    Sleep 1000
    
    ' Test 4: Multiple positions
    Dim positions As Variant
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    
    For i = LBound(positions) To UBound(positions)
        NotifyInfo "Position Test", "Testing position: " & positions(i), 2, CStr(positions(i))
        Sleep 500
    Next
    
    ' Show results
    Sleep 2000
    ShowAnalyticsDashboard
    
    MsgBox "Test suite complete!", vbInformation, "Test Suite"
End Sub

' === DEMO MENU ===
Public Sub ShowMainMenu()
    Dim choice As String
    Dim menu As String
    
    Do
        menu = "+--------------------------------------------+" & vbCrLf
        menu = menu & "¦   TOAST NOTIFICATION SYSTEM v" & TOAST_API_VERSION & "      ¦" & vbCrLf
        menu = menu & "+--------------------------------------------+" & vbCrLf & vbCrLf
        menu = menu & "QUICK ACTIONS:" & vbCrLf
        menu = menu & " 1 - Show Info Notification" & vbCrLf
        menu = menu & " 2 - Show Success Notification" & vbCrLf
        menu = menu & " 3 - Show Warning Notification" & vbCrLf
        menu = menu & " 4 - Show Error Notification" & vbCrLf
        menu = menu & " 5 - Show Progress Demo" & vbCrLf & vbCrLf
        menu = menu & "TOOLS:" & vbCrLf
        menu = menu & " 6 - Open Toast Builder (Visual Designer)" & vbCrLf
        menu = menu & " 7 - List Available Templates" & vbCrLf
        menu = menu & " 8 - Show Analytics Dashboard" & vbCrLf
        menu = menu & " 9 - Export Analytics to CSV" & vbCrLf
        menu = menu & "10 - Show Queue Status" & vbCrLf & vbCrLf
        menu = menu & "SETTINGS:" & vbCrLf
        menu = menu & "11 - Enable Do Not Disturb" & vbCrLf
        menu = menu & "12 - Disable Do Not Disturb" & vbCrLf
        menu = menu & "13 - System Information" & vbCrLf & vbCrLf
        menu = menu & "TESTING:" & vbCrLf
        menu = menu & "14 - Run Test Suite" & vbCrLf & vbCrLf
        menu = menu & " 0 - Exit" & vbCrLf
        
        choice = InputBox(menu, "Toast Notification System", "1")
        
        If choice = "" Or choice = "0" Then Exit Sub
        
        Select Case Val(choice)
            Case 1
                NotifyInfo "Info", "This is an information notification"
            Case 2
                NotifySuccess "Success", "Operation completed successfully!"
            Case 3
                NotifyWarning "Warning", "Please review this warning"
            Case 4
                NotifyError "Error", "An error has occurred", Duration:=5
            Case 5
                RunProgressDemo
            Case 6
                ShowToastBuilder
            Case 7
                ListTemplates
            Case 8
                ShowAnalyticsDashboard
            Case 9
                ExportAnalytics
            Case 10
                ShowQueueStatus
            Case 11
                EnableDoNotDisturb 30
            Case 12
                DisableDoNotDisturb
            Case 13
                MsgBox GetSystemInfo(), vbInformation, "System Information"
            Case 14
                RunTestSuite
            Case Else
                MsgBox "Invalid choice", vbExclamation
        End Select
        
        DoEvents
    Loop
End Sub

' === PRIVATE HELPERS ===
Private Sub EnsureInitialized()
    If Not m_Initialized Then
        InitializeToastSystem
    End If
End Sub

Private Sub ShowQuickToast(ByVal Title As String, ByVal Message As String, _
                          ByVal Level As String, ByVal Duration As Long, _
                          ByVal Position As String)
    Dim toast As New clsToastNotification
    toast.Title = Title
    toast.Message = Message
    toast.Level = Level
    toast.Duration = Duration
    toast.Position = Position
    
    Dim startTime As Double
    startTime = Timer
    
    Dim success As Boolean
    success = toast.Show()
    
    If Not m_Analytics Is Nothing Then
        Dim deliveryTime As Long
        deliveryTime = CLng((Timer - startTime) * 1000)
        m_Analytics.LogToast Title, Message, Level, Position, Duration, "HTA", deliveryTime, success
    End If
End Sub

Private Sub RunProgressDemo()
    Dim toast As clsToastNotification
    Set toast = ShowProgress("File Processing", "Initializing...", 0, "BR")
    
    If Not toast Is Nothing Then
        Dim i As Long
        For i = 0 To 100 Step 10
            UpdateProgress toast, i, "Processing file: " & i & "% complete"
            Sleep 300
        Next
        
        toast.Close
        NotifySuccess "Complete", "File processing finished!"
    End If
End Sub
