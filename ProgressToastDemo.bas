Attribute VB_Name = "ProgressToastDemo"

Public Sub ProgressToastDemo()
    Dim i As Long
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String: tmp = Environ$("TEMP")
    Dim requestFile As String: requestFile = tmp & "\ToastRequest.json"
    
    ' Start PowerShell listener if not running
    If Not MsgBoxUniversal.PowershellListenerRunning() Then
        MsgBoxUniversal.StartToastListener
        Application.Wait Now + TimeValue("00:00:02") ' Wait for listener to start
    End If
    
    ' Display initial progress toast
    If ShowToastPowerShell("Processing Task", "Starting process...", 0, "INFO", , , "BEEP", , , , True, "C", 0) Then
        Debug.Print "Progress toast initiated"
    Else
        MsgBox "Failed to initiate progress toast - is listener running?", vbExclamation
        Exit Sub
    End If
    
    ' Simulate progress updates
    For i = 10 To 100 Step 10
        ' Update JSON with new progress
        Dim tJson As String
        tJson = "{"
        tJson = tJson & """Title"":""" & EscapeJson("Processing Task") & ""","
        tJson = tJson & """Message"":""" & EscapeJson("Processing: " & i & "%") & ""","
        tJson = tJson & """DurationSec"":0,"
        tJson = tJson & """ToastType"":""INFO"","
        tJson = tJson & """LinkUrl"":"""","
        tJson = tJson & """Icon"":"""","
        tJson = tJson & """Sound"":""BEEP"","
        tJson = tJson & """ImagePath"":"""","
        tJson = tJson & """ImageSize"":""Small"","
        tJson = tJson & """CallbackMacro"":"""","
        tJson = tJson & """NoDismiss"":true,"
        tJson = tJson & """Position"":""C"","
        tJson = tJson & """Progress"":" & i
        tJson = tJson & "}"
        
        ' Write updated JSON
        Dim ts As Object
        Set ts = fso.CreateTextFile(requestFile, True, True)
        ts.Write tJson
        ts.Close
        
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] Updated progress to " & i & "%"
        
        ' Wait for PowerShell to process
        Dim maxWait As Long: maxWait = 30
        Dim j As Long
        For j = 1 To maxWait
            If Not fso.FileExists(requestFile) Then
                Exit For
            End If
            Sleep 100
            DoEvents
        Next j
        
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    ' Final toast to indicate completion
    ShowToastPowerShell "Task Complete", "Processing finished!", 3, "INFO", , , "BEEP", , , , , "C", 100
    MsgBox "Progress demo complete!", vbInformation
End Sub

