Attribute VB_Name = "Logs"

Option Explicit

' ============================================================
' Module: Logs
' Purpose: Writes timestamped debug, status, and leveled messages to a daily log file
'          in C:\SmartTraffic\Logs\server_log_YYYY-MM-DD.log and optionally
'          displays messages in Excel's status bar. Creates the directory if it
'          doesn't exist. Thread-safe for non-blocking server operations.
' Parameters:
'   - Log:
'       - msg (String): The message to log.
'       - level (String): Log level (e.g., "INFO", "WARN", "ERROR").
'       - showInStatusBar (Boolean, Optional): If True, displays the message in Excel's status bar.
'       - logFolder (String, Optional): Custom log folder path (default: C:\SmartTraffic\Logs\).
'   - DebugLog:
'       - msg (String): The message to log.
'       - showInStatusBar (Boolean, Optional): If True, displays the message in Excel's status bar.
'       - logFolder (String, Optional): Custom log folder path.
'   - LogStatus:
'       - msg (String): The status message to log and display in the status bar.
'       - logFolder (String, Optional): Custom log folder path.
' Dependencies:
'   - None (uses standard VBA file operations and Excel Application object).
' Usage:
'   - Log: Called by modules needing leveled logging (e.g., INFO, WARN, ERROR).
'   - DebugLog: Called by MQTT modules for debug and error logging.
'   - LogStatus: Called for status updates (e.g., connection status, server state).
' Notes:
'   - Logs are appended to avoid overwriting.
'   - Directory is created if missing.
'   - Status bar messages are cleared after 5 seconds.
'   - Errors are logged to the Immediate Window for debugging.
'   - Added Log method to support leveled logging for compatibility with Govee.bas.
' ============================================================

Private m_nextStatusClear As Date
' --- Compatibility Wrapper ---
Public Sub LogEvent(ByVal msg As String, Optional ByVal Level As String = "INFO")
    ' Redirects legacy LogEvent calls to new Log() method
    On Error Resume Next
    Call Log(msg, Level)
End Sub
' --- Leveled Logging ---
Public Sub Log(ByVal msg As String, ByVal Level As String, Optional ByVal showInStatusBar As Boolean = False, Optional ByVal logFolder As String = "C:\SmartTraffic\Logs\")
    On Error GoTo ErrorHandler
10  Dim f As Integer
20  Dim path As String
30  Dim Timestamp As String
    
40  ' Format timestamp as YYYY-MM-DD HH:MM:SS
50  Timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & UCase(Level) & "] "
    
60  ' Set log file path
70  If Right(logFolder, 1) <> "\" Then logFolder = logFolder & "\"
80  path = logFolder & "server_log_" & Format(Date, "yyyy-mm-dd") & ".log"
    
90  ' Create log directory if it doesn't exist
100 If Dir(logFolder, vbDirectory) = "" Then
110     MkDir logFolder
120 End If
    
130 ' Open file in append mode
140 f = FreeFile
150 Open path For Append As #f
160 Print #f, Timestamp & msg
170 Close #f
    
180 ' Print to Immediate Window
190 Debug.Print Timestamp & msg
    
200 ' Update status bar if requested
210 If showInStatusBar Then
220     Application.StatusBar = "[" & UCase(Level) & "] " & msg
230     ScheduleStatusClear
240 End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "[Logs] Error in Log: " & Err.Description & " (Line: " & Erl & ")"
    On Error Resume Next
    Close #f
End Sub

' --- Debug Logging ---
Public Sub DebugLog(ByVal msg As String, Optional ByVal showInStatusBar As Boolean = False, Optional ByVal logFolder As String = "C:\SmartTraffic\Logs\")
    On Error GoTo ErrorHandler
10  Dim f As Integer
20  Dim path As String
30  Dim Timestamp As String
    
40  ' Format timestamp as YYYY-MM-DD HH:MM:SS
50  Timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [DEBUG] "
    
60  ' Set log file path
70  If Right(logFolder, 1) <> "\" Then logFolder = logFolder & "\"
80  path = logFolder & "server_log_" & Format(Date, "yyyy-mm-dd") & ".log"
    
90  ' Create log directory if it doesn't exist
100 If Dir(logFolder, vbDirectory) = "" Then
110     MkDir logFolder
120 End If
    
130 ' Open file in append mode
140 f = FreeFile
150 Open path For Append As #f
160 Print #f, Timestamp & msg
170 Close #f
    
180 ' Print to Immediate Window
190 Debug.Print Timestamp & msg
    
200 ' Update status bar if requested
210 If showInStatusBar Then
220     Application.StatusBar = msg
230     ScheduleStatusClear
240 End If
    
    Exit Sub

ErrorHandler:
    Debug.Print "[Logs] Error in DebugLog: " & Err.Description & " (Line: " & Erl & ")"
    On Error Resume Next
    Close #f
End Sub

' --- Status Logging ---
Public Sub LogStatus(ByVal msg As String, Optional ByVal logFolder As String = "C:\SmartTraffic\Logs\")
    On Error GoTo ErrorHandler
10  Dim f As Integer
20  Dim path As String
30  Dim Timestamp As String
    
40  ' Format timestamp as YYYY-MM-DD HH:MM:SS
50  Timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [STATUS] "
    
60  ' Set log file path
70  If Right(logFolder, 1) <> "\" Then logFolder = logFolder & "\"
80  path = logFolder & "server_log_" & Format(Date, "yyyy-mm-dd") & ".log"
    
90  ' Create log directory if it doesn't exist
100 If Dir(logFolder, vbDirectory) = "" Then
110     MkDir logFolder
120 End If
    
130 ' Open file in append mode
140 f = FreeFile
150 Open path For Append As #f
160 Print #f, Timestamp & msg
170 Close #f
    
180 ' Print to Immediate Window
190 Debug.Print Timestamp & msg
    
200 ' Update status bar
210 Application.StatusBar = msg
220 ScheduleStatusClear
    
    Exit Sub

ErrorHandler:
    Debug.Print "[Logs] Error in LogStatus: " & Err.Description & " (Line: " & Erl & ")"
    On Error Resume Next
    Close #f
End Sub

' --- Clear Status Bar ---
Private Sub ClearStatusBar()
    On Error GoTo ErrorHandler
10  Application.StatusBar = False
    Exit Sub
ErrorHandler:
    Debug.Print "[Logs] Error in ClearStatusBar: " & Err.Description & " (Line: " & Erl & ")"
End Sub

' --- Schedule Status Bar Clear ---
Private Sub ScheduleStatusClear()
    On Error GoTo ErrorHandler
10  m_nextStatusClear = Now + TimeSerial(0, 0, 5)
20  Application.OnTime m_nextStatusClear, "Logs.ClearStatusBar"
    Exit Sub
ErrorHandler:
    Debug.Print "[Logs] Error in ScheduleStatusClear: " & Err.Description & " (Line: " & Erl & ")"
End Sub




