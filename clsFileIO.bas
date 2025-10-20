VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsFileIO
' Purpose: File and JSON helper for toasts, progress updates, and temporary files
' Version: 1.0
' Date: October 2025
'***************************************************************
Option Explicit

Private Const TEMP_PREFIX As String = "toast_"

'----------------------------
' Get TEMP folder path
Public Function tempPath() As String
    tempPath = Environ$("TEMP")
    If Right$(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
End Function

'----------------------------
' Write text to file (overwrite)
Public Sub WriteTextFile(ByVal FilePath As String, ByVal Content As String, Optional ByVal Unicode As Boolean = True)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(FilePath, True, Unicode)
    ts.Write Content
    ts.Close
End Sub

'----------------------------
' Read text from file
Public Function ReadTextFile(ByVal FilePath As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        Dim ts As Object: Set ts = fso.OpenTextFile(FilePath, 1, True)
        ReadTextFile = ts.ReadAll
        ts.Close
    Else
        ReadTextFile = ""
    End If
End Function

'----------------------------
' Delete file if exists
Public Sub DeleteFile(ByVal FilePath As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then fso.DeleteFile FilePath
    On Error GoTo 0
End Sub

'----------------------------
' Create unique temporary file path
Public Function TempFile(Optional ByVal Suffix As String = ".txt") As String
    TempFile = tempPath() & TEMP_PREFIX & Format(Now, "yyyymmddhhnnss") & "_" & Int((1000 * Rnd) + 1) & Suffix
End Function

'----------------------------
' Simple JSON escape (for strings)
Public Function EscapeJson(ByVal text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, "/", "\/")
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbLf, "\n")
    EscapeJson = text
End Function

