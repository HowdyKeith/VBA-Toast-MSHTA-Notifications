VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToastAnalytics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsToastAnalytics
' Purpose: Track toast notification history and generate analytics
' Version: 6.0
' Features: History logging, performance metrics, error tracking, reporting
'***************************************************************
Option Explicit

' === HISTORY ENTRY ===
Private Type ToastHistoryEntry
    Timestamp As Double
    Title As String
    Message As String
    Level As String
    Position As String
    Duration As Long
    DeliveryMode As String
    deliveryTime As Long  ' milliseconds
    WasShown As Boolean
    ErrorMessage As String
    UserInteraction As String  ' "Dismissed", "LinkClicked", "CallbackExecuted", "AutoClosed"
End Type

' === STATISTICS ===
Private Type ToastStatistics
    TotalSent As Long
    TotalShown As Long
    TotalFailed As Long
    TotalErrors As Long
    TotalWarnings As Long
    TotalInfo As Long
    TotalSuccess As Long
    AverageDeliveryTime As Double
    MostUsedLevel As String
    MostUsedPosition As String
    SessionStartTime As Double
End Type

' === STATE ===
Private m_History() As ToastHistoryEntry
Private m_HistoryCount As Long
Private m_MaxHistorySize As Long
Private m_Stats As ToastStatistics
Private m_LogToFile As Boolean
Private m_LogFilePath As String
Private m_EnableAnalytics As Boolean

' === INITIALIZATION ===
Private Sub Class_Initialize()
    m_MaxHistorySize = 1000
    ReDim m_History(0 To m_MaxHistorySize - 1)
    m_HistoryCount = 0
    m_LogToFile = False
    m_EnableAnalytics = True
    m_LogFilePath = Environ$("TEMP") & "\ExcelToasts\ToastHistory.log"
    
    ' Initialize stats
    m_Stats.SessionStartTime = Timer
End Sub

' === PROPERTIES ===
Public Property Let MaxHistorySize(ByVal value As Long)
    If value > 0 Then
        m_MaxHistorySize = value
        ReDim Preserve m_History(0 To m_MaxHistorySize - 1)
    End If
End Property
Public Property Get MaxHistorySize() As Long
    MaxHistorySize = m_MaxHistorySize
End Property

Public Property Let LogToFile(ByVal value As Boolean)
    m_LogToFile = value
End Property
Public Property Get LogToFile() As Boolean
    LogToFile = m_LogToFile
End Property

Public Property Let LogFilePath(ByVal value As String)
    m_LogFilePath = value
End Property
Public Property Get LogFilePath() As String
    LogFilePath = m_LogFilePath
End Property

Public Property Let EnableAnalytics(ByVal value As Boolean)
    m_EnableAnalytics = value
End Property
Public Property Get EnableAnalytics() As Boolean
    EnableAnalytics = m_EnableAnalytics
End Property

Public Property Get HistoryCount() As Long
    HistoryCount = m_HistoryCount
End Property

' === PUBLIC METHODS ===

' Log a toast notification
Public Sub LogToast(ByVal Title As String, _
                   ByVal Message As String, _
                   ByVal Level As String, _
                   ByVal Position As String, _
                   ByVal Duration As Long, _
                   ByVal DeliveryMode As String, _
                   ByVal DeliveryTimeMs As Long, _
                   ByVal WasShown As Boolean, _
                   Optional ByVal ErrorMsg As String = "")
    
    On Error Resume Next
    
    If Not m_EnableAnalytics Then Exit Sub
    
    ' Rotate history if at max
    If m_HistoryCount >= m_MaxHistorySize Then
        RotateHistory
    End If
    
    ' Add entry
    With m_History(m_HistoryCount)
        .Timestamp = Timer
        .Title = Title
        .Message = Message
        .Level = UCase$(Level)
        .Position = UCase$(Position)
        .Duration = Duration
        .DeliveryMode = DeliveryMode
        .deliveryTime = DeliveryTimeMs
        .WasShown = WasShown
        .ErrorMessage = ErrorMsg
        .UserInteraction = ""
    End With
    
    m_HistoryCount = m_HistoryCount + 1
    
    ' Update statistics
    UpdateStatistics Level, WasShown, Position, DeliveryTimeMs
    
    ' Write to log file if enabled
    If m_LogToFile Then
        WriteToLogFile m_HistoryCount - 1
    End If
End Sub

' Log user interaction
Public Sub LogInteraction(ByVal ToastIndex As Long, ByVal InteractionType As String)
    On Error Resume Next
    
    If ToastIndex >= 0 And ToastIndex < m_HistoryCount Then
        m_History(ToastIndex).UserInteraction = InteractionType
        
        If m_LogToFile Then
            AppendToLogFile "  Interaction: " & InteractionType
        End If
    End If
End Sub

' Get recent history
Public Function GetRecentHistory(Optional ByVal count As Long = 10) As String
    Dim result As String
    result = "=== Recent Toast History ===" & vbCrLf & vbCrLf
    
    Dim startIdx As Long
    startIdx = IIf(m_HistoryCount > count, m_HistoryCount - count, 0)
    
    Dim i As Long
    For i = m_HistoryCount - 1 To startIdx Step -1
        result = result & FormatHistoryEntry(i) & vbCrLf & vbCrLf
    Next
    
    GetRecentHistory = result
End Function

' Get statistics report
Public Function GetStatisticsReport() As String
    Dim report As String
    report = "=== Toast Statistics Report ===" & vbCrLf & vbCrLf
    
    ' Session info
    Dim sessionDuration As Long
    sessionDuration = CLng((Timer - m_Stats.SessionStartTime) / 60)
    
    report = report & "Session Duration: " & sessionDuration & " minutes" & vbCrLf
    report = report & "Total Toasts Sent: " & m_Stats.TotalSent & vbCrLf
    report = report & "Successfully Shown: " & m_Stats.TotalShown & vbCrLf
    report = report & "Failed: " & m_Stats.TotalFailed & vbCrLf & vbCrLf
    
    ' Success rate
    Dim successRate As Double
    If m_Stats.TotalSent > 0 Then
        successRate = (m_Stats.TotalShown / m_Stats.TotalSent) * 100
        report = report & "Success Rate: " & Format(successRate, "0.00") & "%" & vbCrLf & vbCrLf
    End If
    
    ' Breakdown by level
    report = report & "Breakdown by Level:" & vbCrLf
    report = report & "  INFO: " & m_Stats.TotalInfo & vbCrLf
    report = report & "  SUCCESS: " & m_Stats.TotalSuccess & vbCrLf
    report = report & "  WARNING: " & m_Stats.TotalWarnings & vbCrLf
    report = report & "  ERROR: " & m_Stats.TotalErrors & vbCrLf & vbCrLf
    
    ' Performance
    report = report & "Performance:" & vbCrLf
    report = report & "  Average Delivery Time: " & Format(m_Stats.AverageDeliveryTime, "0.00") & " ms" & vbCrLf
    report = report & "  Most Used Level: " & m_Stats.MostUsedLevel & vbCrLf
    report = report & "  Most Used Position: " & m_Stats.MostUsedPosition & vbCrLf
    
    GetStatisticsReport = report
End Function

' Search history
Public Function SearchHistory(ByVal SearchTerm As String, _
                              Optional ByVal SearchField As String = "ALL") As String
    Dim results As String
    results = "=== Search Results for '" & SearchTerm & "' ===" & vbCrLf & vbCrLf
    
    Dim found As Long
    Dim i As Long
    
    SearchField = UCase$(SearchField)
    SearchTerm = LCase$(SearchTerm)
    
    For i = 0 To m_HistoryCount - 1
        Dim match As Boolean
        match = False
        
        Select Case SearchField
            Case "TITLE"
                match = (InStr(1, LCase$(m_History(i).Title), SearchTerm) > 0)
            Case "MESSAGE"
                match = (InStr(1, LCase$(m_History(i).Message), SearchTerm) > 0)
            Case "LEVEL"
                match = (InStr(1, LCase$(m_History(i).Level), SearchTerm) > 0)
            Case "ERROR"
                match = (InStr(1, LCase$(m_History(i).ErrorMessage), SearchTerm) > 0)
            Case Else ' ALL
                match = (InStr(1, LCase$(m_History(i).Title), SearchTerm) > 0) Or _
                       (InStr(1, LCase$(m_History(i).Message), SearchTerm) > 0) Or _
                       (InStr(1, LCase$(m_History(i).Level), SearchTerm) > 0) Or _
                       (InStr(1, LCase$(m_History(i).ErrorMessage), SearchTerm) > 0)
        End Select
        
        If match Then
            results = results & FormatHistoryEntry(i) & vbCrLf & vbCrLf
            found = found + 1
        End If
    Next
    
    results = results & vbCrLf & "Found " & found & " matches."
    SearchHistory = results
End Function

' Get error summary
Public Function GetErrorSummary() As String
    Dim summary As String
    summary = "=== Error Summary ===" & vbCrLf & vbCrLf
    
    Dim errorCount As Long
    Dim i As Long
    
    For i = 0 To m_HistoryCount - 1
        If Not m_History(i).WasShown Or m_History(i).ErrorMessage <> "" Then
            summary = summary & FormatHistoryEntry(i) & vbCrLf & vbCrLf
            errorCount = errorCount + 1
        End If
    Next
    
    summary = summary & "Total Errors: " & errorCount
    GetErrorSummary = summary
End Function

' Export history to CSV
Public Function ExportToCSV(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    
    ' Write header
    ts.WriteLine "Timestamp,Title,Message,Level,Position,Duration,DeliveryMode,DeliveryTime,WasShown,ErrorMessage,UserInteraction"
    
    ' Write entries
    Dim i As Long
    For i = 0 To m_HistoryCount - 1
        With m_History(i)
            ts.WriteLine FormatCSVLine( _
                FormatTimestamp(.Timestamp), _
                .Title, _
                .Message, _
                .Level, _
                .Position, _
                CStr(.Duration), _
                .DeliveryMode, _
                CStr(.deliveryTime), _
                IIf(.WasShown, "TRUE", "FALSE"), _
                .ErrorMessage, _
                .UserInteraction _
            )
        End With
    Next
    
    ts.Close
    ExportToCSV = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[ToastAnalytics] Export error: " & Err.description
    ExportToCSV = False
End Function

' Clear history
Public Sub ClearHistory()
    m_HistoryCount = 0
    ReDim m_History(0 To m_MaxHistorySize - 1)
    ResetStatistics
End Sub

' Get performance metrics
Public Function GetPerformanceMetrics() As String
    Dim metrics As String
    metrics = "=== Performance Metrics ===" & vbCrLf & vbCrLf
    
    If m_HistoryCount = 0 Then
        metrics = metrics & "No data available"
        GetPerformanceMetrics = metrics
        Exit Function
    End If
    
    ' Calculate metrics
    Dim totalTime As Long, minTime As Long, maxTime As Long
    Dim i As Long
    
    minTime = 999999
    maxTime = 0
    
    For i = 0 To m_HistoryCount - 1
        If m_History(i).WasShown Then
            totalTime = totalTime + m_History(i).deliveryTime
            If m_History(i).deliveryTime < minTime Then minTime = m_History(i).deliveryTime
            If m_History(i).deliveryTime > maxTime Then maxTime = m_History(i).deliveryTime
        End If
    Next
    
    Dim avgTime As Double
    If m_Stats.TotalShown > 0 Then
        avgTime = totalTime / m_Stats.TotalShown
    End If
    
    metrics = metrics & "Delivery Time Statistics:" & vbCrLf
    metrics = metrics & "  Average: " & Format(avgTime, "0.00") & " ms" & vbCrLf
    metrics = metrics & "  Minimum: " & minTime & " ms" & vbCrLf
    metrics = metrics & "  Maximum: " & maxTime & " ms" & vbCrLf & vbCrLf
    
    ' Distribution by position
    metrics = metrics & "Distribution by Position:" & vbCrLf
    Dim posCount As Object
    Set posCount = CreateObject("Scripting.Dictionary")
    
    For i = 0 To m_HistoryCount - 1
        Dim pos As String
        pos = m_History(i).Position
        If posCount.Exists(pos) Then
            posCount(pos) = posCount(pos) + 1
        Else
            posCount(pos) = 1
        End If
    Next
    
    Dim key As Variant
    For Each key In posCount.Keys
        Dim pct As Double
        pct = (posCount(key) / m_HistoryCount) * 100
        metrics = metrics & "  " & key & ": " & posCount(key) & " (" & Format(pct, "0.0") & "%)" & vbCrLf
    Next
    
    GetPerformanceMetrics = metrics
End Function

' === PRIVATE METHODS ===

Private Sub UpdateStatistics(ByVal Level As String, _
                            ByVal WasShown As Boolean, _
                            ByVal Position As String, _
                            ByVal DeliveryTimeMs As Long)
    m_Stats.TotalSent = m_Stats.TotalSent + 1
    
    If WasShown Then
        m_Stats.TotalShown = m_Stats.TotalShown + 1
        
        ' Update average delivery time
        Dim newAvg As Double
        newAvg = ((m_Stats.AverageDeliveryTime * (m_Stats.TotalShown - 1)) + DeliveryTimeMs) / m_Stats.TotalShown
        m_Stats.AverageDeliveryTime = newAvg
    Else
        m_Stats.TotalFailed = m_Stats.TotalFailed + 1
    End If
    
    ' Count by level
    Select Case UCase$(Level)
        Case "INFO": m_Stats.TotalInfo = m_Stats.TotalInfo + 1
        Case "SUCCESS": m_Stats.TotalSuccess = m_Stats.TotalSuccess + 1
        Case "WARNING", "WARN": m_Stats.TotalWarnings = m_Stats.TotalWarnings + 1
        Case "ERROR": m_Stats.TotalErrors = m_Stats.TotalErrors + 1
    End Select
    
    ' Track most used
    ' (Simplified - in production, maintain proper counters)
    m_Stats.MostUsedLevel = Level
    m_Stats.MostUsedPosition = Position
End Sub

Private Sub RotateHistory()
    ' Remove oldest half of entries
    Dim keepCount As Long
    keepCount = m_MaxHistorySize \ 2
    
    Dim i As Long
    For i = 0 To keepCount - 1
        m_History(i) = m_History(m_HistoryCount - keepCount + i)
    Next
    
    m_HistoryCount = keepCount
End Sub

Private Function FormatHistoryEntry(ByVal Index As Long) As String
    If Index < 0 Or Index >= m_HistoryCount Then
        FormatHistoryEntry = "Invalid index"
        Exit Function
    End If
    
    Dim entry As String
    With m_History(Index)
        entry = "[" & FormatTimestamp(.Timestamp) & "] "
        entry = entry & .Level & " - " & .Title & vbCrLf
        entry = entry & "  Message: " & .Message & vbCrLf
        entry = entry & "  Position: " & .Position & " | Duration: " & .Duration & "s" & vbCrLf
        entry = entry & "  Delivery: " & .DeliveryMode & " (" & .deliveryTime & "ms)" & vbCrLf
        entry = entry & "  Status: " & IIf(.WasShown, "Shown", "Failed")
        
        If .ErrorMessage <> "" Then
            entry = entry & " - Error: " & .ErrorMessage
        End If
        
        If .UserInteraction <> "" Then
            entry = entry & vbCrLf & "  User Action: " & .UserInteraction
        End If
    End With
    
    FormatHistoryEntry = entry
End Function

Private Function FormatTimestamp(ByVal TimerValue As Double) As String
    ' Convert Timer value to readable timestamp
    Dim baseDate As Date
    baseDate = Date
    
    Dim seconds As Long
    seconds = CLng(TimerValue)
    
    Dim hours As Long, mins As Long, secs As Long
    hours = seconds \ 3600
    mins = (seconds Mod 3600) \ 60
    secs = seconds Mod 60
    
    FormatTimestamp = Format(hours, "00") & ":" & Format(mins, "00") & ":" & Format(secs, "00")
End Function

Private Function FormatCSVLine(ParamArray values() As Variant) As String
    Dim line As String
    Dim i As Long
    
    For i = LBound(values) To UBound(values)
        Dim value As String
        value = CStr(values(i))
        
        ' Escape quotes and wrap in quotes if needed
        If InStr(value, ",") > 0 Or InStr(value, """") > 0 Or InStr(value, vbCrLf) > 0 Then
            value = """" & Replace(value, """", """""") & """"
        End If
        
        line = line & value
        If i < UBound(values) Then line = line & ","
    Next
    
    FormatCSVLine = line
End Function

Private Sub WriteToLogFile(ByVal Index As Long)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure folder exists
    Dim folder As String
    folder = fso.GetParentFolderName(m_LogFilePath)
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder folder
    End If
    
    ' Append to log
    Dim ts As Object
    Set ts = fso.OpenTextFile(m_LogFilePath, 8, True, True)  ' 8 = ForAppending
    ts.WriteLine FormatHistoryEntry(Index)
    ts.WriteLine String(60, "-")
    ts.Close
End Sub

Private Sub AppendToLogFile(ByVal Text As String)
    On Error Resume Next
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(m_LogFilePath, 8, True, True)
    ts.WriteLine Text
    ts.Close
End Sub

Private Sub ResetStatistics()
    With m_Stats
        .TotalSent = 0
        .TotalShown = 0
        .TotalFailed = 0
        .TotalErrors = 0
        .TotalWarnings = 0
        .TotalInfo = 0
        .TotalSuccess = 0
        .AverageDeliveryTime = 0
        .MostUsedLevel = ""
        .MostUsedPosition = ""
        .SessionStartTime = Timer
    End With
End Sub

'***************************************************************
' USAGE EXAMPLES:
'
' Sub Example1_BasicLogging()
'     Dim analytics As New clsToastAnalytics
'     analytics.EnableAnalytics = True
'     analytics.LogToFile = True
'
'     ' Log some toasts
'     analytics.LogToast "File Upload", "Uploading file...", "INFO", "BR", 5, "HTA", 120, True
'     analytics.LogToast "Error", "Upload failed", "ERROR", "TR", 5, "HTA", 95, True, "Network timeout"
'
'     ' View recent history
'     Debug.Print analytics.GetRecentHistory(10)
'
'     ' View statistics
'     Debug.Print analytics.GetStatisticsReport()
' End Sub
'
' Sub Example2_SearchAndExport()
'     Dim analytics As New clsToastAnalytics
'
'     ' Search for errors
'     Debug.Print analytics.SearchHistory("error", "LEVEL")
'
'     ' Export to CSV
'     analytics.ExportToCSV "C:\Temp\ToastHistory.csv"
' End Sub
'
' Sub Example3_PerformanceAnalysis()
'     Dim analytics As New clsToastAnalytics
'
'     ' Get performance metrics
'     Debug.Print analytics.GetPerformanceMetrics()
'
'     ' Get error summary
'     Debug.Print analytics.GetErrorSummary()
' End Sub
'***************************************************************

