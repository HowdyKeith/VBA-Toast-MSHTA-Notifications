Attribute VB_Name = "WinRTMethods"
Option Explicit
'***************************************************************
' Module: WinRTMethods
' Version: 2.6 (October 2025) - Pure VBA WinRT bridge (optional/guarded)
' Purpose: Attempt WinRT activation from VBA using RoGetActivationFactory /
'          RoActivateInstance + pointer->IDispatch bridging (unsafe by default).
' IMPORTANT: The unsafe bridge is OFF by default. Enable only after backing up.
'***************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function WindowsCreateString Lib "api-ms-win-core-winrt-string-l1-1-0.dll" ( _
        ByVal sourceString As LongPtr, ByVal Length As Long, ByRef hString As LongPtr) As Long

    Private Declare PtrSafe Function WindowsDeleteString Lib "api-ms-win-core-winrt-string-l1-1-0.dll" ( _
        ByVal hString As LongPtr) As Long

    Private Declare PtrSafe Function RoGetActivationFactory Lib "combase.dll" ( _
        ByVal activatableClassId As LongPtr, ByRef iid As GUID, ByRef factory As LongPtr) As Long

    Private Declare PtrSafe Function RoActivateInstance Lib "combase.dll" ( _
        ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long

    Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" ( _
        ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long

    Private Declare PtrSafe Function CoCreateInstance Lib "ole32.dll" ( _
        ByRef clsid As GUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, _
        ByRef iid As GUID, ByRef ppv As LongPtr) As Long

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Function WindowsCreateString Lib "api-ms-win-core-winrt-string-l1-1-0.dll" ( _
        ByVal sourceString As Long, ByVal length As Long, ByRef hString As Long) As Long

    Private Declare Function WindowsDeleteString Lib "api-ms-win-core-winrt-string-l1-1-0.dll" ( _
        ByVal hString As Long) As Long

    Private Declare Function RoGetActivationFactory Lib "combase.dll" ( _
        ByVal activatableClassId As Long, ByRef iid As GUID, ByRef factory As Long) As Long

    Private Declare Function RoActivateInstance Lib "combase.dll" ( _
        ByVal activatableClassId As Long, ByRef instance As Long) As Long

    Private Declare Function IIDFromString Lib "ole32.dll" ( _
        ByVal lpsz As Long, ByRef lpiid As GUID) As Long

    Private Declare Function CoCreateInstance Lib "ole32.dll" ( _
        ByRef clsid As GUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, _
        ByRef iid As GUID, ByRef ppv As Long) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const S_OK As Long = 0&
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const IID_IInspectable_STR As String = "{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}"
Private Const IID_IDISPATCH_STR As String = "{00020400-0000-0000-C000-000000000046}"

' -- Unsafe bridge toggle: default False. Set to True ONLY after taking a backup and testing.
Public gAllowUnsafeBridge As Boolean

' === cached objects (optional) ===
Private m_httpClient As Object
Private m_toastMgr As Object
Private m_xmlDom As Object

' -------------------------
' Helper: CreateHString (wraps WindowsCreateString)
' -------------------------
Private Function CreateHStringFromString(ByVal s As String) As LongPtr
    On Error GoTo EH
    Dim h As LongPtr
    Dim p As LongPtr
    If Len(s) = 0 Then
        CreateHStringFromString = 0
        Exit Function
    End If
    ' WindowsCreateString takes (PCWSTR sourceString, UINT32 length, HSTRING *hstring)
    ' Pass pointer to unicode string: use StrPtr(s)
    #If VBA7 Then
        p = StrPtr(s)
    #Else
        p = StrPtr(s)
    #End If
    If WindowsCreateString(p, Len(s), h) = 0 Then
        CreateHStringFromString = h
        Exit Function
    End If
EH:
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] CreateHStringFromString error: " & Err.description, True
    CreateHStringFromString = 0
End Function

' -------------------------
' Helper: FreeHString
' -------------------------
Private Sub FreeHString(ByVal h As LongPtr)
    On Error Resume Next
    If h <> 0 Then WindowsDeleteString h
End Sub

' -------------------------
' WinRTCreateObjectPure (safe attempts, no pointer bridging)
'  - tries moniker: GetObject("winrt:...") (safe)
'  - optionally tries RoActivateInstance without pointer->VBA conversion (returns Nothing)
' -------------------------
Public Function WinRTCreateObjectPure(ByVal className As String) As Object
    On Error GoTo EH
    Dim obj As Object
    Dim moniker As String

    If Len(Trim$(className)) = 0 Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure: empty className", True
        Set WinRTCreateObjectPure = Nothing
        Exit Function
    End If

    ' 1) Try winrt: moniker via GetObject (safe; may succeed in some hosts)
    On Error Resume Next
    moniker = "winrt:" & className
    Set obj = Nothing
    Set obj = GetObject(moniker)
    If Not obj Is Nothing Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure: Activated via winrt: moniker: " & className, False
        Set WinRTCreateObjectPure = obj
        Exit Function
    End If
    On Error GoTo EH

    ' 2) As a last resort, attempt RoActivateInstance to validate availability (but do not return raw pointer)
    Dim hStr As LongPtr, instPtr As LongPtr, hr As Long
    hStr = CreateHStringFromString(className)
    If hStr = 0 Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure: CreateHString failed for " & className, True
        Set WinRTCreateObjectPure = Nothing
        Exit Function
    End If

    hr = RoActivateInstance(hStr, instPtr)
    FreeHString hStr

    If hr = S_OK Then
        ' We can't safely convert instPtr to VBA object here; return Nothing but log availability
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure: RoActivateInstance succeeded but pointer bridging disabled: " & className, False
        Set WinRTCreateObjectPure = Nothing
        Exit Function
    Else
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure: RoActivateInstance hr=" & Hex$(hr) & " for " & className, True
        Set WinRTCreateObjectPure = Nothing
        Exit Function
    End If

    Exit Function
EH:
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectPure error: " & Err.description, True
    Set WinRTCreateObjectPure = Nothing
End Function

' -------------------------
' WinRTCreateObjectUnsafe (uses pointer->IDispatch bridge) - only runs when gAllowUnsafeBridge = True
' -------------------------
Public Function WinRTCreateObjectUnsafe(ByVal className As String) As Object
    On Error GoTo EH
    If Not gAllowUnsafeBridge Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] Unsafe bridge disabled. Use WinRTCreateObjectPure first.", True
        Set WinRTCreateObjectUnsafe = Nothing
        Exit Function
    End If

    Dim hStr As LongPtr, factoryPtr As LongPtr, instPtr As LongPtr
    Dim iidInspectable As GUID, iidDispatch As GUID
    Dim hr As Long

    ' Prepare GUIDs
    IIDFromString StrPtr(IID_IInspectable_STR), iidInspectable
    IIDFromString StrPtr(IID_IDISPATCH_STR), iidDispatch

    hStr = CreateHStringFromString(className)
    If hStr = 0 Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe: CreateHString failed", True
        Set WinRTCreateObjectUnsafe = Nothing
        Exit Function
    End If

    ' Try RoActivateInstance first (factoryless)
    hr = RoActivateInstance(hStr, instPtr)
    If hr <> S_OK Then
        ' Try factory
        hr = RoGetActivationFactory(hStr, iidInspectable, factoryPtr)
        If hr = S_OK And factoryPtr <> 0 Then
            ' Call factory->ActivateInstance via IActivationFactory (vtable index)
            ' IActivationFactory::ActivateInstance is vtable slot 3 (zero-based), but calling via pointer is complex.
            ' Instead attempt QueryInterface to IDispatch on the factory pointer -> risky but we try.
            ' For many WinRT classes RoActivateInstance will be successful; fallback may fail.
            Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe: RoGetActivationFactory succeeded, factoryPtr=" & CStr(factoryPtr), False
            ' For safety, use factoryPtr as instance pointer if possible (best-effort)
            instPtr = factoryPtr
        Else
            FreeHString hStr
            Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe: activation failed hr=" & Hex$(hr), True
            Set WinRTCreateObjectUnsafe = Nothing
            Exit Function
        End If
    End If

    ' If we have an instance pointer, attempt QueryInterface for IDispatch and then copy pointer into VBA object
    If instPtr = 0 Then
        FreeHString hStr
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe: no instance pointer to convert", True
        Set WinRTCreateObjectUnsafe = Nothing
        Exit Function
    End If

    ' QueryInterface for IDispatch: implement by calling QueryInterface on IUnknown pointer (vtable idx 0x00 is QueryInterface)
    ' Build a temporary IUnknown pointer wrapper to call QueryInterface. Use CallWindowProc-style call? Not possible here.
    ' Simpler approach: attempt to coerce pointer into VBA object via CopyMemory into object variable (risky).
    Dim tmp As Object
    Dim ptrVal As LongPtr
    ptrVal = instPtr

    On Error Resume Next
    If LenB(ptrVal) = 0 Then
        ' compute pointer size
        #If VBA7 Then
            ' assume 8 bytes
            ptrVal = instPtr
        #Else
            ptrVal = instPtr
        #End If
    End If

    ' Copy pointer into object variable memory - STRICTLY DANGEROUS
    CopyMemory ByVal VarPtr(tmp), ByVal ptrVal, LenB(ptrVal) ' <- risky operation
    If Err.Number <> 0 Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] CopyMemory to object failed: " & Err.description, True
        On Error GoTo EH
        Set WinRTCreateObjectUnsafe = Nothing
        FreeHString hStr
        Exit Function
    End If

    ' Validate tmp
    If tmp Is Nothing Then
        Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] Pointer bridged but tmp is Nothing", True
        ' attempt to zero out
        On Error Resume Next
        Set tmp = Nothing
        FreeHString hStr
        Set WinRTCreateObjectUnsafe = Nothing
        Exit Function
    End If

    ' Success - return tmp
    Set WinRTCreateObjectUnsafe = tmp
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe: object bridged for " & className, False

    ' Do NOT free instPtr here; ownership rules are delicate. We leave it to COM.
    FreeHString hStr
    Exit Function

EH:
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] WinRTCreateObjectUnsafe error: " & Err.description, True
    On Error Resume Next
    Set WinRTCreateObjectUnsafe = Nothing
    FreeHString hStr
End Function

' -------------------------
' Public wrapper that picks the safe or unsafe path
' -------------------------
Public Function WinRTCreateObject(ByVal className As String) As Object
    ' Prefer safe call first
    Dim obj As Object
    Set obj = WinRTCreateObjectPure(className)
    If Not obj Is Nothing Then
        Set WinRTCreateObject = obj
        Exit Function
    End If

    ' If safe path returned Nothing but unsafe allowed, try unsafe
    If gAllowUnsafeBridge Then
        Set obj = WinRTCreateObjectUnsafe(className)
        If Not obj Is Nothing Then
            Set WinRTCreateObject = obj
            Exit Function
        End If
    End If

    ' No activation
    Set WinRTCreateObject = Nothing
End Function

' -------------------------
' Utilities and cached getters (non-unsafe)
' -------------------------
Public Function WinRTGetXmlDocument() As Object
    If m_xmlDom Is Nothing Then
        Set m_xmlDom = WinRTCreateObjectPure("Windows.Data.Xml.Dom.XmlDocument")
        If m_xmlDom Is Nothing Then
            ' fallback to MSXML
            Set m_xmlDom = CreateObject("MSXML2.DOMDocument")
        End If
    End If
    Set WinRTGetXmlDocument = m_xmlDom
End Function

Public Function WinRTGetToastManager() As Object
    If m_toastMgr Is Nothing Then
        Set m_toastMgr = WinRTCreateObject("Windows.UI.Notifications.ToastNotificationManager")
    End If
    Set WinRTGetToastManager = m_toastMgr
End Function

Public Function WinRTGetHttpClient() As Object
    If m_httpClient Is Nothing Then
        Set m_httpClient = WinRTCreateObject("Windows.Web.Http.HttpClient")
    End If
    Set WinRTGetHttpClient = m_httpClient
End Function

' -------------------------
' SafeRelease
' -------------------------
Public Sub SafeRelease()
    On Error Resume Next
    Set m_httpClient = Nothing
    Set m_xmlDom = Nothing
    Set m_toastMgr = Nothing
End Sub

' -------------------------
' Helper: GetStardate (existing format)
' -------------------------
Private Function GetStardate() As String
    GetStardate = "2025." & Format(Now, "ddd.hh")
End Function

' -------------------------
' Test & Rollback helpers
' -------------------------
Public Sub Test_PureBridge_Check()
    On Error GoTo EH
    Dim obj As Object
    Set obj = WinRTCreateObjectPure("Windows.UI.Notifications.ToastNotificationManager")
    If obj Is Nothing Then
        MsgBox "Pure activation not available in this host (no crash).", vbInformation
    Else
        MsgBox "WinRT object returned (unexpected). Type: " & TypeName(obj), vbInformation
    End If
    Exit Sub
EH:
    MsgBox "Test failed with error: " & Err.description, vbCritical
End Sub

Public Sub Test_PureBridge_EnableAndActivate()
    ' WARNING: enables unsafe bridge and tries to activate a WinRT object.
    ' 1) BACKUP your workbook before running.
    ' 2) Run this in a disposable copy first.
    Dim resp As VbMsgBoxResult
    resp = MsgBox("This will ENABLE the unsafe pointer bridge (dangerous). Are you sure you backed up and want to continue?", vbYesNo + vbExclamation)
    If resp <> vbYes Then Exit Sub

    gAllowUnsafeBridge = True
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] Unsafe bridge ENABLED by user", True

    On Error GoTo EH
    Dim obj As Object
    Set obj = WinRTCreateObjectUnsafe("Windows.UI.Notifications.ToastNotificationManager")
    If obj Is Nothing Then
        MsgBox "Unsafe activation returned Nothing. Either activation failed or bridging failed.", vbExclamation
    Else
        MsgBox "Unsafe activation succeeded. Type: " & TypeName(obj), vbInformation
    End If
    Exit Sub
EH:
    MsgBox "Test encountered error: " & Err.description, vbCritical
End Sub

Public Sub Test_PureBridge_ShowToastIfPossible()
    ' Attempts to show a simple toast using the bridge (unsafe). Backup recommended.
    On Error GoTo EH
    Dim toastMgr As Object
    Set toastMgr = WinRTCreateObject("Windows.UI.Notifications.ToastNotificationManager")
    If toastMgr Is Nothing Then
        MsgBox "ToastNotificationManager not available.", vbExclamation
        Exit Sub
    End If

    Dim xmlDoc As Object
    Set xmlDoc = WinRTGetXmlDocument()
    xmlDoc.LoadXML "<toast><visual><binding template='ToastGeneric'><text>VBA Bridge Test</text><text>Toast from VBA bridge.</text></binding></visual></toast>"

    Dim toast As Object
    Set toast = WinRTCreateObject("Windows.UI.Notifications.ToastNotification")
    If toast Is Nothing Then
        MsgBox "Failed to create ToastNotification object.", vbExclamation
        Exit Sub
    End If

    ' Use late-bound set of Content property if available, else attempt CallByName
    On Error Resume Next
    CallByName toast, "Initialize", VbMethod, xmlDoc
    On Error GoTo EH

    Dim notifier As Object
    Set notifier = CallByName(toastMgr, "CreateToastNotifier", VbMethod)
    CallByName notifier, "Show", VbMethod, toast
    MsgBox "If no errors occurred, toast was invoked (may or may not appear).", vbInformation
    Exit Sub
EH:
    MsgBox "ShowToastIfPossible error: " & Err.description, vbCritical
End Sub

Public Sub Rollback_DisableUnsafeBridge()
    gAllowUnsafeBridge = False
    Logs.DebugLog "[" & GetStardate & "] [WinRTMethods] Unsafe bridge DISABLED by user", True
    SafeRelease
    MsgBox "Unsafe bridge disabled and cached objects released.", vbInformation
End Sub

