Attribute VB_Name = "Split_Join"
Option Explicit

' =============================================================
' VBA Module Split / Join Utilities + Diff Checker
' =============================================================
' Provides:
'   - ExportAndSplit: Export a module into text chunks
'   - ImportAndJoinAuto: Auto-rejoin chunks back into a .bas file
'   - TestRoundTrip: Export ? Split ? Rejoin ? Compare
'   - CompareModuleFiles: Line-by-line diff check
'   - ExportProcedureFromModule: Export single procedures
'
' Files saved in Documents folder.
' =============================================================

' --- Export and split a module ---
Sub ExportAndSplit(ModuleName As String)
    Dim vbComp As Object
    Dim FilePath As String, exportPath As String
    Dim fileNum As Integer, textData As String
    Dim chunkSize As Long, pos As Long, partNum As Long
    
    On Error GoTo ErrorHandler
    
    ' Locate the module
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ModuleName)
    
    ' Export module to temp file
    FilePath = Environ$("USERPROFILE") & "\Documents\" & ModuleName & "_full.bas"
    vbComp.Export FilePath
    
    ' Read module text
    fileNum = FreeFile
    Open FilePath For Input As #fileNum
    textData = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' Split into chunks
    chunkSize = 20000 ' safe size for ChatGPT paste
    pos = 1
    partNum = 1
    Do While pos <= Len(textData)
        exportPath = Environ$("USERPROFILE") & "\Documents\" & _
                     ModuleName & "_Part" & partNum & ".txt"
        fileNum = FreeFile
        Open exportPath For Output As #fileNum
        Print #fileNum, Mid$(textData, pos, chunkSize)
        Close #fileNum
        
        Debug.Print "Created: " & exportPath
        pos = pos + chunkSize
        partNum = partNum + 1
    Loop
    
    MsgBox "Done! Exported " & ModuleName & " into " & (partNum - 1) & " chunks.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportAndSplit: " & Err.Description, vbCritical
    If fileNum > 0 Then Close #fileNum
End Sub

' --- Auto re-join module parts ---
Sub ImportAndJoinAuto(ModuleName As String)
    Dim fso As Object, folder As Object
    Dim joinedPath As String, partPath As String
    Dim fileNum As Integer, outNum As Integer
    Dim partNum As Long, chunkData As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Environ$("USERPROFILE") & "\Documents")
    
    ' Where to save the joined file
    joinedPath = folder.path & "\" & ModuleName & "_Joined.bas"
    outNum = FreeFile
    Open joinedPath For Output As #outNum
    
    ' Loop until missing file
    partNum = 1
    Do
        partPath = folder.path & "\" & ModuleName & "_Part" & partNum & ".txt"
        If Dir(partPath) = "" Then Exit Do
        
        fileNum = FreeFile
        Open partPath For Input As #fileNum
        chunkData = Input$(LOF(fileNum), fileNum)
        Close #fileNum
        
        Print #outNum, chunkData;
        Debug.Print "Joined: " & partPath
        partNum = partNum + 1
    Loop
    
    Close #outNum
    MsgBox "Joined " & (partNum - 1) & " parts into: " & joinedPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ImportAndJoinAuto: " & Err.Description, vbCritical
    If fileNum > 0 Then Close #fileNum
    If outNum > 0 Then Close #outNum
End Sub

' --- Diff checker: Compare full vs joined file ---
Sub CompareModuleFiles(ModuleName As String)
    Dim folder As String
    Dim fullPath As String, joinedPath As String
    Dim file1 As Integer, file2 As Integer
    Dim line1 As String, line2 As String
    Dim lineNum As Long
    
    On Error GoTo ErrorHandler
    
    folder = Environ$("USERPROFILE") & "\Documents\"
    fullPath = folder & ModuleName & "_full.bas"
    joinedPath = folder & ModuleName & "_Joined.bas"
    
    If Dir(fullPath) = "" Or Dir(joinedPath) = "" Then
        MsgBox "Missing _full.bas or _Joined.bas for " & ModuleName, vbExclamation
        Exit Sub
    End If
    
    file1 = FreeFile
    Open fullPath For Input As #file1
    file2 = FreeFile
    Open joinedPath For Input As #file2
    
    lineNum = 0
    Do While Not EOF(file1) And Not EOF(file2)
        Line Input #file1, line1
        Line Input #file2, line2
        lineNum = lineNum + 1
        
        If line1 <> line2 Then
            MsgBox "? Difference at line " & lineNum & ":" & vbCrLf & _
                   "Full:   " & line1 & vbCrLf & _
                   "Joined: " & line2, vbCritical
            Close #file1: Close #file2
            Exit Sub
        End If
    Loop
    
    ' Check if one file has extra lines
    If Not EOF(file1) Or Not EOF(file2) Then
        MsgBox "? Files differ in length at line " & lineNum + 1, vbCritical
    Else
        MsgBox "? Success! Files are identical.", vbInformation
    End If
    
    Close #file1
    Close #file2
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in CompareModuleFiles: " & Err.Description, vbCritical
    If file1 > 0 Then Close #file1
    If file2 > 0 Then Close #file2
End Sub

' --- Round-trip test: split ? rejoin ? compare ---
Sub TestRoundTrip(ModuleName As String)
    Dim fso As Object, folder As String
    Dim fullPath As String, joinedPath As String
    
    On Error GoTo ErrorHandler
    
    folder = Environ$("USERPROFILE") & "\Documents\"
    fullPath = folder & ModuleName & "_full.bas"
    joinedPath = folder & ModuleName & "_Joined.bas"
    
    ' Step 1: Export + Split
    ExportAndSplit ModuleName
    
    ' Step 2: Rejoin
    ImportAndJoinAuto ModuleName
    
    ' Step 3: Quick file size comparison
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fullPath) And fso.FileExists(joinedPath) Then
        Debug.Print "Original size: " & fso.GetFile(fullPath).size & _
                    " vs Joined size: " & fso.GetFile(joinedPath).size
    End If
    
    ' Step 4: Run diff check
    CompareModuleFiles ModuleName
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in TestRoundTrip: " & Err.Description, vbCritical
End Sub

' ==========================================================
' Export Single Procedure from Module
' ==========================================================
' Scans a module for Sub/Function/Property blocks and exports chosen one
' ==========================================================

Sub ExportProcedureFromModule(ModuleName As String)
    Dim vbComp As Object
    Dim codeMod As Object
    Dim lineCount As Long, lineNum As Long
    Dim procName As String, procType As Long
    Dim procStart As Long, procBody As String
    Dim procList As Collection
    Dim i As Long, choice As String
    Dim msg As String
    
    On Error GoTo ErrorHandler
    
    ' Get module
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ModuleName)
    Set codeMod = vbComp.codeModule
    
    lineCount = codeMod.CountOfLines
    Set procList = New Collection
    
    ' Collect all procedure names
    lineNum = 1
    Do While lineNum < lineCount
        procName = codeMod.ProcOfLine(lineNum, procType)
        If procName <> "" Then
            On Error Resume Next
            procList.Add procName, procName
            On Error GoTo 0
            ' Jump past the proc body so we don't repeat
            lineNum = lineNum + codeMod.ProcCountLines(procName, procType)
        Else
            lineNum = lineNum + 1
        End If
    Loop
    
    If procList.count = 0 Then
        MsgBox "No procedures found in module " & ModuleName, vbExclamation
        Exit Sub
    End If
    
    ' Build a display string
    msg = "Procedures in " & ModuleName & ":" & vbCrLf & vbCrLf
    For i = 1 To procList.count
        msg = msg & i & ". " & procList(i) & vbCrLf
    Next
    
    ' Let user pick
    choice = InputBox(msg & vbCrLf & "Enter the number of the procedure to export:", _
                      "Pick Procedure")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    If CLng(choice) < 1 Or CLng(choice) > procList.count Then
        MsgBox "Number out of range.", vbExclamation
        Exit Sub
    End If
    
    ' Get the chosen procedure code
    procName = procList(CLng(choice))
    procStart = codeMod.ProcStartLine(procName, vbext_pk_Proc)
    procBody = codeMod.lines(procStart, codeMod.ProcCountLines(procName, vbext_pk_Proc))
    
    ' Export to Documents folder
    Dim FilePath As String, f As Integer
    FilePath = Environ$("USERPROFILE") & "\Documents\" & procName & ".bas"
    f = FreeFile
    Open FilePath For Output As #f
    Print #f, procBody
    Close #f
    
    MsgBox "Exported " & procName & " to:" & vbCrLf & FilePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportProcedureFromModule: " & Err.Description, vbCritical
    If f > 0 Then Close #f
End Sub

' ==========================================================
' Alternative Procedure Export (Using VBE)
' ==========================================================

Sub ExportProcedureFromActiveModule()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim lineNum As Long, procName As String
    Dim procList As Collection, procType As Long
    Dim i As Long, choice As String
    Dim outFile As String, f As Integer
    
    On Error GoTo ErrorHandler
    
    Set vbProj = Application.VBE.ActiveVBProject
    Set vbComp = Application.VBE.SelectedVBComponent
    Set codeMod = vbComp.codeModule
    
    Set procList = New Collection
    
    lineNum = 1
    Do While lineNum < codeMod.CountOfLines
        procName = codeMod.ProcOfLine(lineNum, procType)
        If procName <> "" Then
            On Error Resume Next
            procList.Add procName, procName
            On Error GoTo 0
            ' Skip to next proc
            lineNum = codeMod.ProcStartLine(procName, procType) + _
                       codeMod.ProcCountLines(procName, procType)
        Else
            lineNum = lineNum + 1
        End If
    Loop
    
    If procList.count = 0 Then
        MsgBox "No procedures found in this module.", vbExclamation
        Exit Sub
    End If
    
    ' Show list in Immediate Window
    Debug.Print "Procedures in " & vbComp.name & ":"
    For i = 1 To procList.count
        Debug.Print i & ": " & procList(i)
    Next
    
    ' Ask user to pick
    choice = InputBox("Enter the number of the procedure to export (see Immediate Window list):", _
                      "Pick Procedure")
    If choice = "" Then Exit Sub
    If Not IsNumeric(choice) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    i = CLng(choice)
    If i < 1 Or i > procList.count Then
        MsgBox "Number out of range.", vbExclamation
        Exit Sub
    End If
    
    procName = procList(i)
    
    ' Export selected proc
    Dim startLine As Long, lineCount As Long, procText As String
    
    startLine = codeMod.ProcStartLine(procName, vbext_pk_Proc)
    lineCount = codeMod.ProcCountLines(procName, vbext_pk_Proc)
    
    procText = codeMod.lines(startLine, lineCount)
    
    outFile = Environ$("USERPROFILE") & "\Documents\" & procName & ".bas"
    f = FreeFile
    Open outFile For Output As #f
    Print #f, procText
    Close #f
    
    MsgBox "Procedure '" & procName & "' exported to: " & outFile, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportProcedureFromActiveModule: " & Err.Description, vbCritical
    If f > 0 Then Close #f
End Sub

' ==========================================================
' Test Procedures
' ==========================================================

Sub TestExportProc()
    ' Example usage - replace "Module1" with your actual module name
    'ExportProcedureFromModule "Module1"
    ExportAndSplit ("Server")
'     ExportAndSplit ("AppLaunch")
   
End Sub

Sub TestSplitJoin()
    ' Example usage - replace "Module1" with your actual module name
    TestRoundTrip "Module1"
End Sub

' ==========================================================
' Export ALL modules in the project, split into chunks
' ==========================================================
Sub ExportAllModulesAndSplit()
    Dim vbComp As Object
    Dim exportedCount As Long
    
    On Error GoTo ErrorHandler
    
    exportedCount = 0
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Only process modules, class modules, and forms
        Select Case vbComp.Type
            Case 1, 2, 3 ' 1=StdModule, 2=ClassModule, 3=Form
                Debug.Print "Exporting module: " & vbComp.name
                ExportAndSplit vbComp.name
                exportedCount = exportedCount + 1
        End Select
    Next vbComp
    
    MsgBox "Done! Exported " & exportedCount & " modules into chunks.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportAllModulesAndSplit: " & Err.Description, vbCritical
End Sub

' ==========================================================
' Export ALL modules (whole, not split)
' ==========================================================
Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim exportedCount As Long
    
    On Error GoTo ErrorHandler
    
    exportedCount = 0
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Only process standard modules, class modules, and forms
        Select Case vbComp.Type
            Case 1, 2, 3 ' 1=StdModule, 2=ClassModule, 3=Form
                exportPath = Environ$("USERPROFILE") & "\Documents\" & vbComp.name & ".bas"
                vbComp.Export exportPath
                Debug.Print "Exported: " & exportPath
                exportedCount = exportedCount + 1
        End Select
    Next vbComp
    
    MsgBox "Done! Exported " & exportedCount & " modules (whole).", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ExportAllModules: " & Err.Description, vbCritical
End Sub

' ==========================================================
' Restore ALL exported modules (whole) into the project
' ==========================================================
Sub ImportAllModules()
    Dim fso As Object, folder As Object, file As Object
    Dim importPath As String, importedCount As Long
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Environ$("USERPROFILE") & "\Documents")
    
    importedCount = 0
    For Each file In folder.files
        ' Only handle VBA export file types
        If LCase(fso.GetExtensionName(file.path)) Like "bas" _
           Or LCase(fso.GetExtensionName(file.path)) Like "cls" _
           Or LCase(fso.GetExtensionName(file.path)) Like "frm" Then
            
            ThisWorkbook.VBProject.VBComponents.Import file.path
            Debug.Print "Imported: " & file.path
            importedCount = importedCount + 1
        End If
    Next
    
    MsgBox "Done! Imported " & importedCount & " modules/forms.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ImportAllModules: " & Err.Description, vbCritical
End Sub

' ==========================================================
' Full Restore of VBA Project from exported files
'   - Deletes all existing code modules, classes, forms
'   - Imports .bas / .cls / .frm from Documents folder
' ==========================================================
Sub RestoreAllModules()
    Dim vbComp As Object
    Dim fso As Object, folder As Object, file As Object
    Dim importedCount As Long, deletedCount As Long
    Dim ext As String
    
    On Error GoTo ErrorHandler
    
    ' 1. Delete all existing modules (but NOT workbook/worksheet objects)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' StdModule=1, ClassModule=2, Form=3
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
                deletedCount = deletedCount + 1
        End Select
    Next vbComp
    
    ' 2. Import everything from Documents
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Environ$("USERPROFILE") & "\Documents")
    
    importedCount = 0
    For Each file In folder.files
        ext = LCase(fso.GetExtensionName(file.path))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            ThisWorkbook.VBProject.VBComponents.Import file.path
            Debug.Print "Imported: " & file.path
            importedCount = importedCount + 1
        End If
    Next
    
    MsgBox "Restore complete!" & vbCrLf & _
           "Deleted: " & deletedCount & " old modules" & vbCrLf & _
           "Imported: " & importedCount & " from backup.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in RestoreAllModules: " & Err.Description, vbCritical
End Sub


