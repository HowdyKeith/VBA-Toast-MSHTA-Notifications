VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToastItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
' Class: ToastItem.cls
' Purpose: Toast notification data object
' Version: 2.4
' Changes: Added Position property for toast positioning
'***************************************************************

Private m_Title As String
Private m_Message As String
Private m_DurationSec As Long
Private m_ToastType As String
Private m_LinkUrl As String
Private m_Icon As String
Private m_Sound As String
Private m_ImagePath As String
Private m_ImageSize As String
Private m_CallbackMacro As String
Private m_NoDismiss As Boolean
Private m_Position As String

' Properties
Public Property Let Title(v As String): m_Title = v: End Property
Public Property Get Title() As String: Title = m_Title: End Property

Public Property Let Message(v As String): m_Message = v: End Property
Public Property Get Message() As String: Message = m_Message: End Property

Public Property Let DurationSec(v As Long): m_DurationSec = v: End Property
Public Property Get DurationSec() As Long: DurationSec = m_DurationSec: End Property

Public Property Let ToastType(v As String): m_ToastType = v: End Property
Public Property Get ToastType() As String: ToastType = m_ToastType: End Property

Public Property Let LinkUrl(v As String): m_LinkUrl = v: End Property
Public Property Get LinkUrl() As String: LinkUrl = m_LinkUrl: End Property

Public Property Let Icon(v As String): m_Icon = v: End Property
Public Property Get Icon() As String: Icon = m_Icon: End Property

Public Property Let Sound(v As String): m_Sound = v: End Property
Public Property Get Sound() As String: Sound = m_Sound: End Property

Public Property Let ImagePath(v As String): m_ImagePath = v: End Property
Public Property Get ImagePath() As String: ImagePath = m_ImagePath: End Property

Public Property Let ImageSize(v As String): m_ImageSize = v: End Property
Public Property Get ImageSize() As String: ImageSize = m_ImageSize: End Property

Public Property Let CallbackMacro(v As String): m_CallbackMacro = v: End Property
Public Property Get CallbackMacro() As String: CallbackMacro = m_CallbackMacro: End Property

Public Property Let NoDismiss(v As Boolean): m_NoDismiss = v: End Property
Public Property Get NoDismiss() As Boolean: NoDismiss = m_NoDismiss: End Property

Public Property Let Position(v As String): m_Position = v: End Property
Public Property Get Position() As String: Position = m_Position: End Property

' JSON Serialization
Public Function ToJson() As String
    Dim json As String
    json = "{"
    json = json & """Title"":""" & JsonEscape(m_Title) & ""","
    json = json & """Message"":""" & JsonEscape(m_Message) & ""","
    json = json & """DurationSec"":" & m_DurationSec & ","
    json = json & """ToastType"":""" & m_ToastType & ""","
    json = json & """LinkUrl"":""" & JsonEscape(m_LinkUrl) & ""","
    json = json & """Icon"":""" & JsonEscape(m_Icon) & ""","
    json = json & """Sound"":""" & m_Sound & ""","
    json = json & """ImagePath"":""" & JsonEscape(m_ImagePath) & ""","
    json = json & """ImageSize"":""" & m_ImageSize & ""","
    json = json & """CallbackMacro"":""" & m_CallbackMacro & ""","
    json = json & """NoDismiss"":" & IIf(m_NoDismiss, "true", "false") & ","
    json = json & """Position"":""" & m_Position & """"
    json = json & "}"
    ToJson = json
End Function

Private Function JsonEscape(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n")
    txt = Replace(txt, vbCr, "\n")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbTab, "\t")
    JsonEscape = txt
End Function

