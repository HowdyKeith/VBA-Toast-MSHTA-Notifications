Attribute VB_Name = "MsgBoxUI_Declares"
Option Explicit

' =====================================================================
' Module: MsgBoxUI_Declares.bas
' Version: 5.4
' Purpose: API Declarations and Utility Support for MsgBoxUI / ToastWatcher
' Author: Keith Swerling + ChatGPT (GPT-5)
' Updated: 2025-10-18
' =====================================================================

' ============================================================
' Win32 API DECLARATIONS
' ============================================================

' Handle utilities
Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

' Pipe + file I/O
Public Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileW" ( _
    ByVal lpFileName As LongPtr, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Public Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long

' Process control
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As LongPtr, _
    ByVal uExitCode As Long) As Long

Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' ============================================================
' CONSTANTS
' ============================================================

Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const OPEN_EXISTING As Long = 3
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const INVALID_HANDLE_VALUE As LongPtr = -1

Public Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Public Const PROCESS_TERMINATE As Long = &H1

' ============================================================
' UTILITIES
' ============================================================

' === Pipe existence check ===
Public Function PipeExists(ByVal PipeName As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    PipeExists = fso.FileExists(PipeName)
End Function

' === JSON fallback existence ===
Public Function ToastJsonPending() As Boolean
    On Error Resume Next
    Dim tmpFile As String
    tmpFile = Environ$("TEMP") & "\ExcelToasts\ToastRequest.json"
    ToastJsonPending = (Dir(tmpFile) <> "")
End Function

' === Detect if ToastWatcherRT.ps1 is active ===
Public Function IsToastWatcherRunning() As Boolean
    On Error Resume Next
    Dim wmi As Object, procs As Object, p As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    Dim psPath As String
    psPath = LCase$(Environ$("USERPROFILE") & "\onedrive\documents\2025\powershell\toastwatcherrt.ps1")
    For Each p In procs
        If InStr(LCase$(p.CommandLine), psPath) > 0 Then
            IsToastWatcherRunning = True
            Exit Function
        End If
    Next
    IsToastWatcherRunning = False
End Function

' === Attempt graceful termination of ToastWatcherRT ===
Public Function KillToastWatcher() As Boolean
    On Error Resume Next
    Dim wmi As Object, procs As Object, p As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    Dim psPath As String
    psPath = LCase$(Environ$("USERPROFILE") & "\onedrive\documents\2025\powershell\toastwatcherrt.ps1")
    For Each p In procs
        If InStr(LCase$(p.CommandLine), psPath) > 0 Then
            p.Terminate
            KillToastWatcher = True
        End If
    Next
End Function

' === Quick diagnostic output ===
Public Sub DebugToastStatus()
    Dim status As String
    status = "=== MsgBoxUI/ToastWatcher Diagnostic ===" & vbCrLf & _
             "MsgBoxUI Version: " & MSGBOXUI_VERSION & vbCrLf & _
             "PowerShell Toasts Enabled: " & UsePowerShellToasts & vbCrLf & _
             "Temp JSON Fallback Enabled: " & UseTempJsonFallback & vbCrLf & _
             "Pipe Name: " & ToastPipeName & vbCrLf & _
             "Listener Running: " & IsToastWatcherRunning & vbCrLf & _
             "Pending Temp JSON: " & ToastJsonPending & vbCrLf & _
             "========================================"
    MsgBox status, vbInformation, "MsgBoxUI Diagnostic"
End Sub

' === Internal low-level pipe write ===
Public Function WritePipeMessage(ByVal PipeName As String, ByVal text As String) As Boolean
    On Error Resume Next
    Dim hPipe As LongPtr
    Dim written As Long
    Dim bytes() As Byte
    bytes = StrConv(text, vbFromUnicode)
    
    hPipe = CreateFile(StrPtr(PipeName), GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hPipe <> INVALID_HANDLE_VALUE Then
        Call WriteFile(hPipe, bytes(0), UBound(bytes) + 1, written, 0)
        CloseHandle hPipe
        WritePipeMessage = (written > 0)
    Else
        WritePipeMessage = False
    End If
End Function

' ============================================================
' CONSOLE HELPERS (for PowerShell debug)
' ============================================================
Public Sub ConsoleLog(ByVal text As String)
    On Error Resume Next
    Debug.Print "[MsgBoxUI] " & text
End Sub


