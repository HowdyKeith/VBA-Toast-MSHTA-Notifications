VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallbacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsCallbacks
' Purpose: Handle toast callback macros and link clicks
' Version: 1.0
' Date: October 2025
'***************************************************************
Option Explicit

' Dictionary of callbacks
Private Callbacks As Object

'----------------------------
' Initialize callbacks dictionary
Private Sub Class_Initialize()
    Set Callbacks = CreateObject("Scripting.Dictionary")
End Sub

'----------------------------
' Register a callback
Public Sub RegisterCallback(ByVal name As String, ByVal macroName As String)
    If Len(name) > 0 And Len(macroName) > 0 Then
        Callbacks(name) = macroName
    End If
End Sub

'----------------------------
' Execute a callback by name
Public Sub ExecuteCallback(ByVal name As String)
    If Callbacks.Exists(name) Then
        Dim macro As String
        macro = Callbacks(name)
        On Error Resume Next
        Application.Run macro
        If Err.Number <> 0 Then
            MsgBox "Error executing callback '" & name & "': " & Err.description, vbExclamation, "Callback Error"
        End If
        On Error GoTo 0
    Else
        MsgBox "Callback '" & name & "' not registered.", vbExclamation, "Callback Error"
    End If
End Sub

'----------------------------
' Remove a callback
Public Sub RemoveCallback(ByVal name As String)
    If Callbacks.Exists(name) Then Callbacks.Remove name
End Sub

'----------------------------
' Clear all callbacks
Public Sub ClearCallbacks()
    Callbacks.RemoveAll
End Sub

