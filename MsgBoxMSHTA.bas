Attribute VB_Name = "MsgBoxMSHTA"
'***************************************************************
' Module: MsgBoxMSHTA
' Version: 7.0
' Purpose: Display Excel/Office Toasts via HTA
' Dependencies: clsToastNotification v7.1
' Features:
'   - Queued & stacked toasts
'   - Progress updates
'   - Native sound alerts
'   - Callback macro auto-invoker
'   - Optional OnTime timer cleanup
'***************************************************************
Option Explicit

Private SharedQueue As Collection

Private Sub InitQueue()
    If SharedQueue Is Nothing Then Set SharedQueue = New Collection
End Sub

Public Sub ShowToast(ByVal Title As String, ByVal Message As String, Optional ByVal Level As String = "INFO", _
                     Optional ByVal Duration As Long = 5, Optional ByVal Position As String = "BR", _
                     Optional ByVal SoundName As String = "", Optional ByVal CallbackMacro As String = "")
    Dim toast As clsToastNotification
    Set toast = New clsToastNotification
    
    toast.Title = Title
    toast.Message = Message
    toast.Level = Level
    toast.Duration = Duration
    toast.Position = Position
    toast.SoundName = SoundName
    toast.CallbackMacro = CallbackMacro
    
    InitQueue
    SharedQueue.Add toast
    If SharedQueue.count = 1 Then DisplayNextToast
End Sub

Private Sub DisplayNextToast()
    If SharedQueue.count = 0 Then Exit Sub
    Dim currentToast As clsToastNotification
    Set currentToast = SharedQueue(1)
    
    currentToast.Show "hta"
    
    Dim waitSeconds As Long
    waitSeconds = currentToast.Duration
    If currentToast.IsShowing Then
        Application.Wait Now + TimeSerial(0, 0, waitSeconds)
    End If
    
    currentToast.Close
    SharedQueue.Remove 1
    
    If SharedQueue.count > 0 Then DisplayNextToast
End Sub


