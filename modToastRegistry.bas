Attribute VB_Name = "modToastRegistry"
' modToastRegistry - small registry + OnTime handler for clsToastNotification
Option Explicit

Public ToastRegistry As Object ' Scripting.Dictionary

Private Sub EnsureRegistry()
    If ToastRegistry Is Nothing Then
        Set ToastRegistry = CreateObject("Scripting.Dictionary")
    End If
End Sub

' Register an instance against an ID
Public Sub RegisterToast(ByVal ID As String, ByRef instance As clsToastNotification)
    On Error Resume Next
    EnsureRegistry
    If Not ToastRegistry.Exists(ID) Then
        ToastRegistry.Add ID, instance
    Else
        ToastRegistry(ID) = instance
    End If
End Sub

' Unregister
Public Sub UnregisterToast(ByVal ID As String)
    On Error Resume Next
    EnsureRegistry
    If ToastRegistry.Exists(ID) Then ToastRegistry.Remove ID
End Sub

' Lookup
Public Function GetToastByHash(ByVal ID As String) As clsToastNotification
    On Error Resume Next
    EnsureRegistry
    If ToastRegistry.Exists(ID) Then
        Set GetToastByHash = ToastRegistry(ID)
    Else
        Set GetToastByHash = Nothing
    End If
End Function

' This is the procedure that Application.OnTime will call.
' OnTime will call a string expression like:
'   "'" & ThisWorkbook.Name & "'!RefreshToastTimerHandler(""Toast_2025..._1234"")"
' Therefore this handler must be Public and accept 1 string param.
Public Sub RefreshToastTimerHandler(ByVal HashID As String)
    On Error Resume Next
    Dim t As clsToastNotification
    Set t = GetToastByHash(HashID)
    If t Is Nothing Then
        ' nothing to do
        Exit Sub
    End If
    t.TimerTick
End Sub

