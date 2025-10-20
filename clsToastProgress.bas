VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToastProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsToastProgress
' Purpose: Modern progress toast notifications using MSHTA
' Version: 1.0
' Date: October 2025
'***************************************************************
Option Explicit

Private Const TOAST_WIDTH As Long = 350
Private Const TOAST_HEIGHT As Long = 150
Private Const TOAST_MARGIN As Long = 20
Private Const TOAST_STACK_OFFSET As Long = 10

Private m_Title As String
Private m_Message As String
Private m_Icon As String
Private m_Position As String
Private m_Duration As Long
Private m_LinkUrl As String
Private m_HTAPath As String
Private m_ToastCount As Long
Private m_ProgressValue As Long

'----------------------------
' Initialize toast
Public Sub Init(ByVal Title As String, ByVal Message As String, _
                Optional ByVal Icon As String = "?", _
                Optional ByVal Position As String = "BR", _
                Optional ByVal DurationSec As Long = 0, _
                Optional ByVal LinkUrl As String = "")
    m_Title = Title
    m_Message = Message
    m_Icon = Icon
    m_Position = Position
    m_Duration = DurationSec
    m_LinkUrl = LinkUrl
    m_ProgressValue = 0
End Sub

'----------------------------
' Show toast (initial)
Public Sub Show()
    Dim posX As Long, posY As Long
    CalculateToastPosition m_Position, posX, posY
    
    m_HTAPath = Environ$("TEMP") & "\toast_progress_" & Format(Now, "yyyymmddhhnnss") & "_" & m_ToastCount & ".hta"
    m_ToastCount = m_ToastCount + 1
    
    Dim html As String
    html = BuildHTAHTML(m_ProgressValue, posX, posY)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.CreateTextFile(m_HTAPath, True, True)
    ts.Write html
    ts.Close
    
    ' Launch HTA
    shell "mshta.exe """ & m_HTAPath & """", vbHide
End Sub

'----------------------------
' Update progress (0-100)
Public Sub UpdateProgress(ByVal Percent As Long, Optional ByVal NewMessage As String = "")
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    m_ProgressValue = Percent
    If Len(NewMessage) > 0 Then m_Message = NewMessage
    
    ' Rewrite HTA to update bar
    Dim posX As Long, posY As Long
    CalculateToastPosition m_Position, posX, posY
    Dim html As String
    html = BuildHTAHTML(m_ProgressValue, posX, posY)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(m_HTAPath) Then
        Dim ts As Object: Set ts = fso.CreateTextFile(m_HTAPath, True, True)
        ts.Write html
        ts.Close
    End If
End Sub

'----------------------------
' Close toast
Public Sub Close()
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(m_HTAPath) Then fso.DeleteFile m_HTAPath
End Sub

'----------------------------
' Build HTA HTML
Private Function BuildHTAHTML(ByVal Progress As Long, ByVal posX As Long, ByVal posY As Long) As String
    Dim html As String
    html = "<html><head><meta charset='UTF-8'><title>" & EscapeHtml(m_Title) & "</title>" & vbCrLf
    html = html & "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no' SCROLL='no'>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body{margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;background:transparent;overflow:hidden;}" & vbCrLf
    html = html & "#toast{position:fixed;width:" & TOAST_WIDTH & "px;height:" & TOAST_HEIGHT & "px;"
    html = html & "padding:15px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);"
    html = html & "background:linear-gradient(135deg,#4caf50,#2e7d32);color:#ffffff;}"
    html = html & "h3{margin:0 0 8px;font-size:18px;font-weight:600;display:flex;align-items:center;}"
    html = html & ".icon{font-size:24px;margin-right:8px;}"
    html = html & "#progressBar{width:100%;background:rgba(255,255,255,0.2);border-radius:4px;height:20px;margin-top:10px;}"
    html = html & "#progress{width:" & Progress & "%;height:100%;background:#ffffff;border-radius:4px;transition:width 0.2s;}"
    html = html & "#dismissBtn{position:absolute;top:8px;right:8px;padding:4px 8px;font-size:16px;font-weight:bold;background:rgba(0,0,0,0.3);border-radius:4px;cursor:pointer;}" & vbCrLf
    html = html & "</style></head><body>" & vbCrLf
    
    html = html & "<div id='toast'>" & vbCrLf
    html = html & "<button id='dismissBtn' onclick='window.close()'>×</button>" & vbCrLf
    html = html & "<h3><span class='icon'>" & m_Icon & "</span>" & EscapeHtml(m_Title) & "</h3>" & vbCrLf
    html = html & "<p>" & EscapeHtml(m_Message) & "</p>" & vbCrLf
    html = html & "<div id='progressBar'><div id='progress'></div></div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    html = html & "<script>" & vbCrLf
    html = html & "window.resizeTo(" & TOAST_WIDTH + 20 & "," & TOAST_HEIGHT + 20 & ");" & vbCrLf
    html = html & "window.moveTo(" & posX & "," & posY & ");" & vbCrLf
    html = html & "</script></body></html>"
    
    BuildHTAHTML = html
End Function

'----------------------------
' Escape HTML
Private Function EscapeHtml(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")
    EscapeHtml = text
End Function

'----------------------------
' Calculate toast position
Private Sub CalculateToastPosition(ByVal posCode As String, ByRef outX As Long, ByRef outY As Long)
    Dim screenW As Long, screenH As Long
    On Error Resume Next
    screenW = CLng(CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\ScreenWidth"))
    screenH = CLng(CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\ScreenHeight"))
    If screenW = 0 Then screenW = 1920
    If screenH = 0 Then screenH = 1080
    
    Select Case UCase$(posCode)
        Case "TL": outX = TOAST_MARGIN: outY = TOAST_MARGIN
        Case "TR": outX = screenW - TOAST_WIDTH - TOAST_MARGIN: outY = TOAST_MARGIN
        Case "BL": outX = TOAST_MARGIN: outY = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "BR": outX = screenW - TOAST_WIDTH - TOAST_MARGIN: outY = screenH - TOAST_HEIGHT - TOAST_MARGIN
        Case "CR": outX = screenW - TOAST_WIDTH - TOAST_MARGIN: outY = (screenH - TOAST_HEIGHT) \ 2
        Case "C": outX = (screenW - TOAST_WIDTH) \ 2: outY = (screenH - TOAST_HEIGHT) \ 2
        Case Else: outX = screenW - TOAST_WIDTH - TOAST_MARGIN: outY = screenH - TOAST_HEIGHT - TOAST_MARGIN
    End Select
End Sub


