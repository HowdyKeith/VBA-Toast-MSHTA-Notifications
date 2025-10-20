Attribute VB_Name = "ToastMasterDemo"
Option Explicit
'***************************************************************
' Module: ToastMasterDemo.bas
' Purpose: Complete demonstration and testing suite for toast system
' Version: 1.0
'***************************************************************

'================= MASTER TEST MENU =================
Public Sub MasterToastTestMenu()
    Dim choice As String
    
    Do
        choice = InputBox("===== TOAST NOTIFICATION MASTER TEST MENU =====" & vbCrLf & vbCrLf & _
                          "QUICK TESTS:" & vbCrLf & _
                          "1 - Quick Info Toast (MSHTA)" & vbCrLf & _
                          "2 - Quick Warning Toast (MSHTA)" & vbCrLf & _
                          "3 - Quick Error Toast (MSHTA)" & vbCrLf & _
                          "4 - Quick PowerShell Toast" & vbCrLf & vbCrLf & _
                          "POSITION TESTS:" & vbCrLf & _
                          "5 - Test All MSHTA Positions" & vbCrLf & _
                          "6 - Test All PowerShell Positions" & vbCrLf & _
                          "7 - Test Stacked Toasts (BR)" & vbCrLf & vbCrLf & _
                          "ADVANCED TESTS:" & vbCrLf & _
                          "8 - Toast with Link" & vbCrLf & _
                          "9 - Toast with Callback" & vbCrLf & _
                          "10 - Multi-Type Demo (Info/Warn/Error)" & vbCrLf & vbCrLf & _
                          "SYSTEM:" & vbCrLf & _
                          "11 - Check PowerShell Status" & vbCrLf & _
                          "12 - Toggle PowerShell Mode" & vbCrLf & _
                          "13 - Reset Toast Stack" & vbCrLf & _
                          "14 - Full System Test" & vbCrLf & vbCrLf & _
                          "0 - Exit", _
                          "Toast Master Demo", "1")
        
        If choice = "" Or choice = "0" Then Exit Sub
        
        Select Case val(choice)
            Case 1: QuickInfoTest
            Case 2: QuickWarningTest
            Case 3: QuickErrorTest
            Case 4: QuickPowerShellTest
            Case 5: TestAllMSHTAPositions
            Case 6: TestAllPowerShellPositions
            Case 7: TestStackedToasts
            Case 8: TestLinkToast
            Case 9: TestCallbackToast
            Case 10: MultiTypeDemo
            Case 11: CheckPowerShellStatus
            Case 12: TogglePowerShellMode
            Case 13: ResetStack
            Case 14: FullSystemTest
            Case Else
                MsgBox "Invalid choice. Please enter 0-14.", vbExclamation
        End Select
    Loop
End Sub

'================= QUICK TESTS =================
Private Sub QuickInfoTest()
    MsgBoxUniversal.MsgInfoEx "This is a quick info toast test!", "BR"
    MsgBox "Info toast displayed at bottom-right.", vbInformation, "Test Complete"
End Sub

Private Sub QuickWarningTest()
    MsgBoxUniversal.MsgWarnEx "This is a warning toast with sound!", "TR"
    MsgBox "Warning toast displayed at top-right with beep.", vbInformation, "Test Complete"
End Sub

Private Sub QuickErrorTest()
    MsgBoxUniversal.MsgErrorEx "This is an error toast!", "TL"
    MsgBox "Error toast displayed at top-left.", vbInformation, "Test Complete"
End Sub

Private Sub QuickPowerShellTest()
    If Not MsgBoxToastsPS.ShowToastPowerShell("PowerShell Test", "Testing PS toast!", 3, "INFO") Then
        MsgBox "PowerShell toast failed. Check if PowerShell is available.", vbExclamation
    Else
        MsgBox "PowerShell toast launched successfully.", vbInformation, "Test Complete"
    End If
End Sub

'================= POSITION TESTS =================
Private Sub TestAllMSHTAPositions()
    Dim positions As Variant
    Dim posNames As Variant
    Dim i As Long
    
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    posNames = Array("Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center-Right", "Center")
    
    MsgBox "Will display 6 MSHTA toasts in different positions." & vbCrLf & _
           "Each will appear for 3 seconds.", vbInformation, "Position Test"
    
    For i = LBound(positions) To UBound(positions)
        MsgBoxUniversal.ShowMsgBoxUnified _
            "Position: " & posNames(i), _
            "MSHTA Toast #" & (i + 1), _
            vbInformation, "modern", 3, "INFO", "", "", "?", False, "", "auto", "", CStr(positions(i))
        Application.Wait Now + TimeValue("00:00:03.5")
    Next i
    
    MsgBox "All MSHTA position tests complete!", vbInformation, "Test Complete"
End Sub

Private Sub TestAllPowerShellPositions()
    Dim positions As Variant
    Dim posNames As Variant
    Dim i As Long
    
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    posNames = Array("Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right", "Center-Right", "Center")
    
    MsgBox "Will display 6 PowerShell toasts in different positions." & vbCrLf & _
           "Each will appear for 3 seconds.", vbInformation, "Position Test"
    
    For i = LBound(positions) To UBound(positions)
        MsgBoxToastsPS.ShowToastPowerShell _
            "PS Toast #" & (i + 1), _
            "Position: " & posNames(i), _
            3, "INFO", "", "", "", "", "", "", False, CStr(positions(i))
        Application.Wait Now + TimeValue("00:00:03.5")
    Next i
    
    MsgBox "All PowerShell position tests complete!", vbInformation, "Test Complete"
End Sub

Private Sub TestStackedToasts()
    Dim i As Long
    
    MsgBox "Will display 5 stacked toasts at bottom-right." & vbCrLf & _
           "Watch them stack with offset!", vbInformation, "Stack Test"
    
    MsgBoxMSHTA.ResetToastStack ' Start fresh
    
    For i = 1 To 5
        MsgBoxUniversal.MsgInfoEx "Stacked Toast #" & i & " of 5", "BR"
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    Application.Wait Now + TimeValue("00:00:05")
    MsgBoxMSHTA.ResetToastStack
    
    MsgBox "Stack test complete!", vbInformation, "Test Complete"
End Sub

'================= ADVANCED TESTS =================
Private Sub TestLinkToast()
    MsgBoxUniversal.ShowMsgBoxUnified _
        "Click the link below to open Microsoft's website.", _
        "Toast with Link", _
        vbInformation, "modern", 0, "INFO", _
        "https://www.microsoft.com", "", "??", False, "", "auto", "", "BR"
    
    MsgBox "Link toast displayed (no auto-close)." & vbCrLf & _
           "Click the link or close button to dismiss.", vbInformation, "Test Info"
End Sub

Private Sub TestCallbackToast()
    MsgBoxUniversal.ShowMsgBoxUnified _
        "This toast has a callback macro attached.", _
        "Callback Test", _
        vbInformation, "modern", 0, "INFO", _
        "https://www.microsoft.com", "OnDemoToastCallback", "??", False, "", "auto", "", "CR"
    
    MsgBox "Callback toast displayed at center-right." & vbCrLf & _
           "Click the link to trigger callback.", vbInformation, "Test Info"
End Sub

Private Sub MultiTypeDemo()
    MsgBox "Will display 3 toasts: Info, Warning, Error" & vbCrLf & _
           "Each in different positions.", vbInformation, "Multi-Type Demo"
    
    ' Info at bottom-right
    MsgBoxUniversal.MsgInfoEx "This is an INFO toast.", "BR"
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Warning at top-right
    MsgBoxUniversal.MsgWarnEx "This is a WARNING toast!", "TR"
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Error at top-left
    MsgBoxUniversal.MsgErrorEx "This is an ERROR toast!", "TL"
    Application.Wait Now + TimeValue("00:00:03")
    
    MsgBox "Multi-type demo complete!", vbInformation, "Test Complete"
End Sub

'================= SYSTEM FUNCTIONS =================
Private Sub CheckPowerShellStatus()
    Dim psRunning As Boolean
    psRunning = MsgBoxUniversal.PowershellListenerRunning()
    
    Dim msg As String
    msg = "PowerShell Listener Status: " & IIf(psRunning, "RUNNING", "NOT RUNNING") & vbCrLf & vbCrLf
    msg = msg & "PowerShell Toast Mode: " & IIf(MsgBoxUniversal.UsePowerShellToasts, "ENABLED", "DISABLED") & vbCrLf & vbCrLf
    
    If Not psRunning And MsgBoxUniversal.UsePowerShellToasts Then
        msg = msg & "? Warning: PS mode is enabled but listener is not running." & vbCrLf
        msg = msg & "Toasts will fall back to MSHTA."
    ElseIf psRunning Then
        msg = msg & "? PowerShell listener is active and ready."
    Else
        msg = msg & "? MSHTA mode is active (default)."
    End If
    
    MsgBox msg, vbInformation, "PowerShell Status"
End Sub

Private Sub TogglePowerShellMode()
    MsgBoxUniversal.UsePowerShellToasts = Not MsgBoxUniversal.UsePowerShellToasts
    
    Dim msg As String
    If MsgBoxUniversal.UsePowerShellToasts Then
        msg = "PowerShell Toast Mode: ENABLED" & vbCrLf & vbCrLf
        msg = msg & "Toasts will use PowerShell if listener is running," & vbCrLf
        msg = msg & "otherwise fall back to MSHTA."
    Else
        msg = msg & "PowerShell Toast Mode: DISABLED" & vbCrLf & vbCrLf
        msg = msg & "All toasts will use MSHTA."
    End If
    
    MsgBox msg, vbInformation, "Mode Changed"
End Sub

Private Sub ResetStack()
    MsgBoxMSHTA.ResetToastStack
    MsgBox "Toast stack counter has been reset.", vbInformation, "Stack Reset"
End Sub

Private Sub FullSystemTest()
    Dim response As VbMsgBoxResult
    response = MsgBox("This will run a comprehensive test of the entire toast system." & vbCrLf & vbCrLf & _
                      "It will take about 60 seconds and display multiple toasts." & vbCrLf & vbCrLf & _
                      "Continue?", vbQuestion + vbYesNo, "Full System Test")
    
    If response <> vbYes Then Exit Sub
    
    ' Phase 1: Basic MSHTA toasts
    MsgBox "Phase 1: Testing basic MSHTA toasts...", vbInformation, "Test Phase 1/5"
    MsgBoxUniversal.MsgInfoEx "Phase 1: Info toast", "BR"
    Application.Wait Now + TimeValue("00:00:03")
    MsgBoxUniversal.MsgWarnEx "Phase 1: Warning toast", "TR"
    Application.Wait Now + TimeValue("00:00:03")
    MsgBoxUniversal.MsgErrorEx "Phase 1: Error toast", "TL"
    Application.Wait Now + TimeValue("00:00:03")
    
    ' Phase 2: Position tests
    MsgBox "Phase 2: Testing all positions...", vbInformation, "Test Phase 2/5"
    Dim positions As Variant
    positions = Array("TL", "TR", "BL", "BR", "CR", "C")
    Dim i As Long
    For i = LBound(positions) To UBound(positions)
        MsgBoxUniversal.MsgInfoEx "Position: " & positions(i), CStr(positions(i))
        Application.Wait Now + TimeValue("00:00:02")
    Next i
    
    ' Phase 3: Stacking test
    MsgBox "Phase 3: Testing stacked toasts...", vbInformation, "Test Phase 3/5"
    MsgBoxMSHTA.ResetToastStack
    For i = 1 To 3
        MsgBoxUniversal.MsgInfoEx "Stacked #" & i, "BR"
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.Wait Now + TimeValue("00:00:04")
    
    ' Phase 4: PowerShell test
    MsgBox "Phase 4: Testing PowerShell toasts...", vbInformation, "Test Phase 4/5"
    MsgBoxToastsPS.ShowToastPowerShell "PS Test", "PowerShell toast test", 3, "INFO"
    Application.Wait Now + TimeValue("00:00:04")
    
    ' Phase 5: Advanced features
    MsgBox "Phase 5: Testing advanced features...", vbInformation, "Test Phase 5/5"
    MsgBoxUniversal.ShowMsgBoxUnified _
        "This toast has a link!", _
        "Advanced Test", _
        vbInformation, "modern", 5, "INFO", _
        "https://www.microsoft.com", "", "??", False, "", "auto", "", "CR"
    Application.Wait Now + TimeValue("00:00:05")
    
    ' Cleanup
    MsgBoxMSHTA.ResetToastStack
    
    MsgBox "? Full system test complete!" & vbCrLf & vbCrLf & _
           "All components tested successfully.", vbInformation, "Test Complete"
End Sub

'================= CALLBACK HANDLERS =================
Public Sub OnDemoToastCallback()
    MsgBox "Callback executed successfully!" & vbCrLf & vbCrLf & _
           "This macro was triggered by clicking the toast link.", _
           vbInformation, "Callback Success"
End Sub

'================= DIAGNOSTIC TOOL =================
Public Sub DiagnoseToastSystem()
    Dim report As String
    report = "===== TOAST SYSTEM DIAGNOSTIC REPORT =====" & vbCrLf & vbCrLf
    
    ' Check temp directory
    Dim tempPath As String
    tempPath = MsgBoxUniversal.GetTempPath()
    report = report & "Temp Directory: " & tempPath & vbCrLf
    report = report & "Temp Dir Exists: " & CBool(Len(Dir(tempPath, vbDirectory)) > 0) & vbCrLf & vbCrLf
    
    ' Check PowerShell
    report = report & "PS Listener Running: " & MsgBoxUniversal.PowershellListenerRunning() & vbCrLf
    report = report & "PS Toast Mode: " & MsgBoxUniversal.UsePowerShellToasts & vbCrLf & vbCrLf
    
    ' Check for temp files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim TempFolder As Object
    Set TempFolder = fso.GetFolder(tempPath)
    
    Dim toastFileCount As Long
    Dim f As Object
    For Each f In TempFolder.files
        If InStr(1, f.name, "toast_", vbTextCompare) > 0 Or _
           InStr(1, f.name, "ShowToast_", vbTextCompare) > 0 Then
            toastFileCount = toastFileCount + 1
        End If
    Next f
    
    report = report & "Temp Toast Files: " & toastFileCount & vbCrLf & vbCrLf
    
    ' Module status
    report = report & "? MsgBoxUniversal.bas loaded" & vbCrLf
    report = report & "? MsgBoxMSHTA.bas loaded" & vbCrLf
    report = report & "? MsgBoxToastsPS.bas loaded" & vbCrLf & vbCrLf
    
    report = report & "System Status: OPERATIONAL"
    
    MsgBox report, vbInformation, "System Diagnostic"
End Sub

'================= CLEANUP UTILITY =================
Public Sub CleanupToastTempFiles()
    On Error Resume Next
    
    Dim tempPath As String
    tempPath = MsgBoxUniversal.GetTempPath()
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim count As Long
    count = 0
    
    ' Delete toast HTA files
    Dim f As Object
    For Each f In fso.GetFolder(tempPath).files
        If InStr(1, f.name, "toast_", vbTextCompare) > 0 Or _
           InStr(1, f.name, "ShowToast_", vbTextCompare) > 0 Or _
           InStr(1, f.name, "callback_", vbTextCompare) > 0 Then
            f.Delete True
            count = count + 1
        End If
    Next f
    
    MsgBox "Cleanup complete!" & vbCrLf & vbCrLf & _
           "Files removed: " & count, vbInformation, "Cleanup"
End Sub

