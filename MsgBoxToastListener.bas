Attribute VB_Name = "MsgBoxToastListener"
'***************************************************************
' Module: MsgBoxToastListener.bas
' Purpose: Asynchronous PowerShell toast listener
'***************************************************************
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
         ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long, _
         ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
#End If

Private Const TEMP_DIR As String = "C:\Temp"
Private Const RESPONSE_FILE As String = "ToastResponse.txt"
Private m_TimerID As LongPtr
Private m_Interval As Long
Private m_Active As Boolean
Private m_LastResponseKey As String
Private m_LastResponseTime As Date
Private Const DEBOUNCE_SECONDS As Double = 2

Public Sub InitializeToastListener()
    StartToastListener 1
    Logs.Log "[ToastListener] Auto-started during app initialization.", "INFO"
End Sub

Public Sub ShutdownToastListener()
    StopToastListener
    Logs.Log "[ToastListener] Auto-stopped during app shutdown.", "INFO"
End Sub

Public Sub StartToastListener(Optional ByVal IntervalSeconds As Long = 2)
    If m_Active Then Exit Sub
    If IntervalSeconds < 1 Then IntervalSeconds = 2
    m_Interval = IntervalSeconds * 1000
    m_TimerID = SetTimer(0, 0, m_Interval, AddressOf TimerCallback)
    m_Active = (m_TimerID <> 0)
    Logs.Log "[ToastListener] Started (asynchronous, " & IntervalSeconds & "s)", "INFO"
End Sub

Public Sub StopToastListener()
    On Error Resume Next
    If m_Active Then
        KillTimer 0, m_TimerID
        m_Active = False
        Logs.Log "[ToastListener] Stopped.", "INFO"
    End If
End Sub

Private Sub TimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                          ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    CheckToastResponse
End Sub

Private Sub CheckToastResponse()
    On Error Resume Next
    Dim fso As Object, ts As Object, txt As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(TEMP_DIR & "\" & RESPONSE_FILE) Then Exit Sub
    Set ts = fso.OpenTextFile(TEMP_DIR & "\" & RESPONSE_FILE, 1)
    txt = Trim$(ts.ReadAll)
    ts.Close
    fso.DeleteFile TEMP_DIR & "\" & RESPONSE_FILE, True
    If Len(txt) > 0 Then HandleToastResponse txt
End Sub

Private Sub HandleToastResponse(ByVal responseText As String)
    Logs.Log "[ToastListener] Response received: " & responseText, "DEBUG"
    Dim dict As Object: Set dict = ParseSimpleJson(responseText)
    If dict Is Nothing Then Exit Sub

    Dim CallbackName As String, userInput As String
    CallbackName = dict("CallbackMacro")
    If CallbackName = "" Then CallbackName = dict("OnClickCallback")
    userInput = dict("UserInput")
    
    Dim responseKey As String
    responseKey = CallbackName & "|" & userInput
    
    If responseKey = m_LastResponseKey Then
        If DateDiff("s", m_LastResponseTime, Now) < DEBOUNCE_SECONDS Then Exit Sub
    End If

    m_LastResponseKey = responseKey
    m_LastResponseTime = Now

    If CallbackName <> "" Then
        On Error Resume Next
        Application.Run CallbackName, userInput
    End If
End Sub

Private Function ParseSimpleJson(ByVal txt As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    txt = Replace(Replace(Replace(Replace(txt, "{", ""), "}", ""), """", ""), "'", "")
    Dim parts() As String, kv() As String, i As Long
    parts = Split(txt, ",")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), ":")
        If UBound(kv) >= 1 Then dict(Trim$(kv(0))) = Trim$(kv(1))
    Next i
    Set ParseSimpleJson = dict
End Function


