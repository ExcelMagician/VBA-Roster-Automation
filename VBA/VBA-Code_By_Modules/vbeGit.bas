Attribute VB_Name = "vbeGit"
Public Sub GitSave()
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
End Sub

Public Sub DeleteAndMake()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim desktopPath As String: desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Dim parentFolder As String: parentFolder = desktopPath & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
    
    Debug.Print "Parent Folder Path: " & parentFolder
    
    On Error Resume Next
    fso.DeleteFolder parentFolder, True
    On Error GoTo 0
    
    On Error GoTo MkDirError
    If Not fso.FolderExists(parentFolder) Then MkDir parentFolder
    If Not fso.FolderExists(childA) Then MkDir childA
    If Not fso.FolderExists(childB) Then MkDir childB
    On Error GoTo 0
    Exit Sub

MkDirError:
    MsgBox "Error creating directory: " & Err.Description & " (Path: " & parentFolder & ")", vbCritical
    Exit Sub
End Sub

Public Sub PrintAllCode()
    Dim item As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not item.codeModule Is Nothing Then
            Dim lineCount As Long
            lineCount = item.codeModule.CountOfLines
            If lineCount > 0 Then
                lineToPrint = item.codeModule.lines(1, lineCount)
                Debug.Print lineToPrint
                textToPrint = textToPrint & vbCrLf & lineToPrint
            Else
                Debug.Print item.Name & " has no code lines"
            End If
        Else
            Debug.Print item.Name & " has no accessible code module"
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
End Sub

Public Sub PrintAllContainers()
    Dim item As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
End Sub

Public Sub ExportModules()
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As Object
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case 3 ' vbext_ct_ClassModule or vbext_ct_MSForm
                filePath = filePath & ".cls"
            Case 1 ' vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case 100 ' vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
End Sub

Public Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    Dim fileSystem As Object
    Dim textObject As Object
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
End Sub

