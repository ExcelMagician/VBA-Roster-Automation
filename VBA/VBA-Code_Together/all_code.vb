
Sub SwapStaff()
    Dim wsRoster As Worksheet
    Dim wsPersonnel As Worksheet
    Dim wsSwap As Worksheet
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim dateCell As Range
    Dim dateRange As Range
    Dim oriName As String
    Dim newName As String
    Dim r As Long
    Dim i As Long
    Dim lastRow As Long
    Dim col As Variant
    Dim oriNameExists As Boolean
    Dim cellValue As String
    Dim lines() As String
    Dim nameExists As Boolean
    Dim currentName As String
    Dim cumulativeLength As Long
    Dim startPos As Integer
    
    ' Set worksheet references
    Set wsRoster = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")

    ' Get names from Swap sheet
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))
    
    ' Check for empty values
    If Len(oriName) = 0 Then
        MsgBox "Error: Original staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Error: New staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    
    ' Prompt user to select date cells in Column A
    On Error Resume Next
    Set dateRange = Application.InputBox("Select date cells (Column A)", Type:=8)
    On Error GoTo 0
    If dateRange Is Nothing Then Exit Sub

    ' Define slot columns: F, H, J, L, N (6, 8, 10, 12, 14)
    slotCols = Array(6, 8, 10, 12, 14)
    
    ' Check if oriName exists in the selected date rows across slot columns
    oriNameExists = False
    For Each dateCell In dateRange
        r = dateCell.Row
        For Each col In slotCols
            cellValue = wsRoster.Cells(r, col).Value
            If InStr(cellValue, vbNewLine) > 0 Then
                If Trim(Split(cellValue, vbNewLine)(0)) = oriName Then
                    oriNameExists = True
                End If
            Else
                If UCase(Trim(cellValue)) = oriName Then
                    oriNameExists = True
                End If
            End If
            If oriNameExists Then Exit For
        Next col
        If oriNameExists Then Exit For
    Next dateCell
    
    If Not oriNameExists Then
        MsgBox "Error: " & oriName & " not found in the selected rows. Swap not allowed.", vbCritical
    End If

    ' Loop over selected date rows
    For Each dateCell In dateRange
        r = dateCell.Row

        ' Check if newName exists in the same row across all slot columns
        nameExists = False
        For Each col In slotCols
            cellValue = wsRoster.Cells(r, col).Value
            If InStr(cellValue, vbNewLine) > 0 Then
                If Trim(Split(cellValue, vbNewLine)(0)) = newName Then
                    nameExists = True
                End If
            Else
                If UCase(Trim(cellValue)) = newName Then
                    nameExists = True
                End If
            End If
            If nameExists Then Exit For
        Next col
        
        If nameExists Then
            MsgBox "Error: " & newName & " already exists in row " & r & ". Swap not allowed.", vbCritical
        Else
            For Each slotCol In slotCols
                With wsRoster.Cells(r, slotCol)
                    ' Determine the current name based on whether there is a line break
                    If InStr(.Value, vbNewLine) > 0 Then
                        currentName = Trim(Split(.Value, vbNewLine)(0)) ' First unstriked line for subsequent swaps
                    Else
                        currentName = Trim(.Value) ' Entire value for initial swap
                    End If
                    
                    If UCase(currentName) = oriName Then ' Check the current name
                        ' Add new name first (unstriked) and preserve existing content
                        .Value = newName & vbNewLine & .Value
                        .VerticalAlignment = xlTop ' Align text to the top
                        .WrapText = True
                        
                        ' Split into lines to apply strikethrough to all previous names
            
                        lines = Split(.Value, vbNewLine)
                        cumulativeLength = Len(newName) + 2 ' Start with newName and its vbNewLine
                        
                        ' Apply strikethrough to all lines except the first one
                        For k = 1 To UBound(lines)
                            startPos = cumulativeLength
                            .Characters(startPos, Len(lines(k)) + 1).Font.Strikethrough = True
                            cumulativeLength = cumulativeLength + Len(lines(k)) + 2 ' Update for next line
                        Next k
                        
                        ' Explicitly increase row height by 15 points per swap
                        .RowHeight = .RowHeight + 15
                        
                        ' Update personnel counter for the new staff
                        lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
                        ' Deduct duties from the old staff
                        For i = 12 To lastRow
                            If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = oriName Then
                                wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value - 1 ' Decrement Weekly Duties Counter
                                If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                                    wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value - 1 ' Decrement AOH Counter
                                End If
                                Exit For
                            End If
                        Next i
                        ' Update duties for the new staff
                        For i = 12 To lastRow
                            If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = newName Then
                                wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value + 1 ' Increment Weekly Duties Counter
                                If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                                    wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value + 1 ' Increment AOH Counter
                                End If
                                Exit For
                            End If
                        Next i
                    End If
                End With
            Next slotCol
        End If
    Next dateCell

    MsgBox "Swap complete.", vbInformation
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim table As ListObject
    Set table = Me.ListObjects("Desk_PersonnelList")

    ' Check if the change happened in the Duties Counter column
    If Not Intersect(Target, table.ListColumns("Weekly Duties Counter").DataBodyRange) Is Nothing Then
        ' Sort by Duties Counter column (ascending)
        With table.Sort
            .SortFields.Clear
            .SortFields.Add key:=table.ListColumns("Weekly Duties Counter").DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
End Sub


Sub AssignFirstEmployeeToFirstSlot()
    Dim wsRoster As Worksheet
    Dim wsPersonnel As Worksheet
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim slotCell As Range
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long 'weekly duties
    Dim currAOH As Long
    Dim lastRow As Long
    Dim currRow As Long
    Dim found As Boolean
    Dim dateRow As Long
    Dim currDateSlotRange As Range
    Dim isAohSlot As Boolean
    Dim alreadyAssigned As Boolean 'already assigned on current day
    Dim canAssign As Boolean
    'haha

    ' Set references to sheets
    Set wsRoster = Sheets("Master")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")

    ' Find last row number of the employee list
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
    found = False
    
    ' Array of columns for Morning, Afternoon, AOH
    slotCols = Array(6, 8, 10)  ' F, H, J columns
    
    'Find current date and slot range
    
     ' Loop through each date row
     For dateRow = 6 To 186
        ' Loop through each slot column for this date
        For Each slotCol In slotCols
            Set slotCell = wsRoster.Cells(dateRow, slotCol)
            isAohSlot = (slotCol = 10) ' 10 is column J
            found = False
            
            'Loop through each staff
            For currRow = 12 To lastRow
                staffName = wsPersonnel.Cells(currRow, "B").Value
                maxDuties = wsPersonnel.Cells(currRow, "D").Value
                currDuties = wsPersonnel.Cells(currRow, "E").Value
                currAOH = wsPersonnel.Cells(currRow, "F").Value
                
                'Check if this staff already assigned today
                alreadyAssigned = False
                Set currDateSlotRange = wsRoster.Range("F" & dateRow & ":J" & dateRow)
                For Each cell In currDateSlotRange
                    If cell.Value = staffName Then
                        alreadyAssigned = True
                        Exit For
                    End If
                Next cell
                
                'Determine the criteria
                If isAohSlot Then
                    canAssign = (currAOH < 1) And (currDuties < maxDuties) And Not alreadyAssigned
                Else
                    canAssign = (currDuties < maxDuties) And Not alreadyAssigned
                End If
                    
                'Assign the staff and do counter increment if meet th criteria
                If canAssign Then
                    'Assign staff to a slot
                    slotCell.Value = staffName
                    
                    'Do increment
                    If isAohSlot Then
                        wsPersonnel.Cells(currRow, "F").Value = currAOH + 1
                    End If
                    wsPersonnel.Cells(currRow, "E").Value = currDuties + 1
                    
                    found = True
                    Exit For
                End If
            Next currRow
                
            ' If no staff found who can still take duties
            If Not found Then
                slotCell.Value = "Not Available"
            End If
            
        Next slotCol
    Next dateRow
    
    MsgBox "Roster filled"
    
End Sub

Sub AssignFirstEmployeeToFirstSlotCopy()
    Dim wsRosterCopy As Worksheet
    Dim wsPersonnel As Worksheet
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim slotCell As Range
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long 'weekly duties
    Dim currAOH As Long
    Dim lastRow As Long
    Dim currRow As Long
    Dim found As Boolean
    Dim dateRow As Long
    Dim currDateSlotRange As Range
    Dim isAohSlot As Boolean
    Dim alreadyAssigned As Boolean 'already assigned on current day
    Dim canAssign As Boolean
    Dim currDate As Date
    Dim isSaturday As Boolean
    Dim isVacation As Boolean
    Dim lastRowRoster As Integer

    Set wsRosterCopy = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSettings = Sheets("Settings")

    ' Find last row number of the employee list
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
    found = False
    
    If wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" And wsRosterCopy.Cells(2, 13).Value Mod 4 = 0 Then
        lastRowRoster = 187
    ElseIf wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" Then
        lastRowRoster = 186
    Else
        lastRowRoster = 189
    End If
    
    
     'Loop through each date row
     For dateRow = 6 To lastRowRoster
     
        currDate = wsRosterCopy.Cells(dateRow, 2).Value
        
        If Weekday(currDate, vbMonday) = 7 Or _
            Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0 Then
            
            ' Skip this date by marking all slots as "CLOSED"
            wsRosterCopy.Cells(dateRow, 4).Value = "CLOSED" ' D column
            wsRosterCopy.Cells(dateRow, 4).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, 6).Value = "CLOSED" ' F column
            wsRosterCopy.Cells(dateRow, 6).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, 8).Value = "CLOSED" ' H column
            wsRosterCopy.Cells(dateRow, 8).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, 10).Value = "CLOSED" ' J column
            wsRosterCopy.Cells(dateRow, 10).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, 12).Value = "CLOSED" ' L column
            wsRosterCopy.Cells(dateRow, 12).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, 14).Value = "CLOSED" ' N column
            wsRosterCopy.Cells(dateRow, 14).Interior.Color = vbRed
            GoTo NextDate ' Skip to the next date
        End If
        
        For Each slotCol In Array(4, 6, 8, 10, 12, 14) ' D, F, H, J, L, N columns
            Set slotCell = wsRosterCopy.Cells(dateRow, slotCol)
            slotCell.Interior.ColorIndex = xlNone ' Reset to no fill (default)
            slotCell.Font.Strikethrough = False
        Next slotCol
        
        isSaturday = (Weekday(currDate, vbMonday) = 6)
        
        isVacation = (wsRosterCopy.Cells(dateRow, 1).Value = "Vacation")
        
        If isSaturday Then
            slotCols = Array(12, 14) ' L, N for Saturday
        ElseIf isVacation Then
            slotCols = Array(6, 8) ' F, H only for vacation weekdays (no J AOH)
        Else
            slotCols = Array(6, 8, 10) ' F, H, J for Sem Time weekdays
        End If
            
        
        ResetAOHCounter.ResetAOHCounter
        
        ' Loop through each slot column for this date
        For Each slotCol In slotCols
            Set slotCell = wsRosterCopy.Cells(dateRow, slotCol)
            isAohSlot = (slotCol = 10 Or isSaturday) And Not isVacation ' J, L, or N as AOH
            found = False
            
            'Loop through each staff
            For currRow = 12 To lastRow
                staffName = wsPersonnel.Cells(currRow, "B").Value
                maxDuties = wsPersonnel.Cells(currRow, "D").Value
                currDuties = wsPersonnel.Cells(currRow, "E").Value
                currAOH = wsPersonnel.Cells(currRow, "F").Value
                
                'Check if this staff already assigned today
                alreadyAssigned = False
                If isSaturday Then
                    Set currDateSlotRange = wsRosterCopy.Range("L" & dateRow & ":N" & dateRow)
                ElseIf isVacation Then
                    Set currDateSlotRange = wsRosterCopy.Range("F" & dateRow & ":H" & dateRow)
                Else
                    Set currDateSlotRange = wsRosterCopy.Range("F" & dateRow & ":J" & dateRow)
                End If
                
                For Each cell In currDateSlotRange
                    If cell.Value = staffName Then
                        alreadyAssigned = True
                        Exit For
                    End If
                Next cell
                
                'Determine the criteria
                If isAohSlot Then
                    canAssign = (currAOH < 1) And (currDuties < maxDuties) And Not alreadyAssigned
                Else
                    canAssign = (currDuties < maxDuties) And Not alreadyAssigned
                End If
                    
                'Assign the staff and do counter increment if meet th criteria
                If canAssign Then
                    'Assign staff to a slot
                    slotCell.Value = staffName
                    
                    'Do increment
                    If isAohSlot Then
                        wsPersonnel.Cells(currRow, "F").Value = currAOH + 1
                    End If
                    wsPersonnel.Cells(currRow, "E").Value = currDuties + 1
                    
                    found = True
                    Exit For
                End If
            Next currRow
                
            ' If no staff found who can still take duties
            If Not found Then
                slotCell.Value = "Not Available"
            End If
            
        Next slotCol
        
NextDate:
    Next dateRow
    
    MsgBox "Roster filled"
    
End Sub






Public Sub InsertStaffCounter()
    Dim ws As Worksheet
    Dim newRow As Long
    Dim lastRow As Long
    Dim staffName As String, dept As String
    Dim maxDuties As Variant
    Dim checkRow As Long
    Dim i As Long
    Dim matchRow As Long
    
    matchRow = 0

    Set ws = ThisWorkbook.Sheets("PersonnelList (AOH & Desk)")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Read correct cell values
    staffName = UCase(Trim(ws.Range("C3").Value)) ' Name
    dept = Trim(ws.Range("C4").Value)          ' Department
    maxDuties = Trim(ws.Range("C5").Value)      ' Max Duties

    ' Validation
    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in for both Name and Department.", vbExclamation
        Exit Sub
    End If
    
    ' Check for duplicate names
    For checkRow = 10 To lastRow
        If ws.Cells(checkRow, 2).Value = staffName Then
            MsgBox "This staff name already exists. ", vbExclamation
            Exit Sub
        End If
    Next checkRow


    If Not IsNumeric(maxDuties) Or maxDuties < 0 Then
        MsgBox "Max Duties must be more than 0.", vbExclamation
        Exit Sub
    End If
    
    If Trim(ws.Range("C6").Value) > Trim(ws.Range("C5").Value) Then
        MsgBox "Duties Counter must be less than Max Duties per week.", vbExclamation
        Exit Sub
    End If

    
    ' Find next empty row based on column B (Name)
    newRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1

    ' Insert data into row
    ws.Cells(newRow, 2).Value = staffName    ' Name
    ws.Cells(newRow, 3).Value = dept               ' Dept
    ws.Cells(newRow, 4).Value = maxDuties      ' Max Duties
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Set Duties Counter
    If Trim(ws.Range("C6").Value) = "" Then
        ws.Cells(newRow, 5).Value = 0
    Else
        ws.Cells(newRow, 5).Value = ws.Range("C6").Value
    End If
    
    For i = 10 To lastRow
        If UCase(Trim(ws.Cells(i, 2).Value)) = UCase(Trim(staffName)) Then
            matchRow = i
        Exit For
        End If
    Next i

    ' Set AOH Counter
    If Trim(ws.Range("C7").Value) = "" Then
        ws.Cells(matchRow, 6).Value = 0
    Else
        If Trim(ws.Range("C7").Value) > 1 Then
            MsgBox "AOH Counter must not be more than 1.", vbExclamation
            Exit Sub
        End If
        ws.Cells(matchRow, 6).Value = ws.Range("C7").Value
    End If

    ' Clear input
    ws.Range("C3:C7").ClearContents

    MsgBox "Staff added successfully!", vbInformation
End Sub

Private Sub Worksheet_Calculate()
    Dim rng1 As Range, rng2 As Range, rng3 As Range
    Dim rng4 As Range, rng5 As Range, rng6 As Range

    ' Qualify ranges with the worksheet (Me refers to the worksheet containing this code)
    Set rng1 = Me.Range("B187:O187")
    Set rng2 = Me.Range("B188:O188")
    Set rng3 = Me.Range("B189:O189")
    Set rng4 = Me.Range("D187:O187")
    Set rng5 = Me.Range("D188:O188")
    Set rng6 = Me.Range("D189:O189")

    ' Check if B188 and B189 are empty
    If Me.Range("B188").Value = "" And Me.Range("B189").Value = "" Then
        If Me.Range("B187").Value = "" Then
            
            rng1.Borders.LineStyle = xlNone
            rng2.Borders.LineStyle = xlNone
            rng3.Borders.LineStyle = xlNone
            
            rng1.Interior.ColorIndex = xlNone
            rng2.Interior.ColorIndex = xlNone
            rng3.Interior.ColorIndex = xlNone
        Else
            
            rng1.Borders.LineStyle = xlContinuous
            rng2.Borders.LineStyle = xlNone
            rng3.Borders.LineStyle = xlNone
            
            For Each cell In rng4
                If cell.Value <> "" Then
                    cell.Interior.Color = vbWhite ' White for cells with content
                Else
                    cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
                End If
            Next cell
            
            rng2.Interior.ColorIndex = xlNone
            rng3.Interior.ColorIndex = xlNone
        End If
    Else
       
        rng1.Borders.LineStyle = xlContinuous
        rng2.Borders.LineStyle = xlContinuous
        rng3.Borders.LineStyle = xlContinuous

        ' Check each cell in the ranges and set color based on content
        For Each cell In rng4
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
        For Each cell In rng5
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
        For Each cell In rng6
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Define the target cell (e.g., A1, adjust as needed)
    Dim toggleCell As Range
    Set toggleCell = Me.Range("J2") ' Change to the actual cell reference
    
    ' Debug: Log that the event has started
    Debug.Print "Worksheet_Change event triggered for Target: " & Target.Address
    
    ' Check if the changed cell is the toggle cell
    If Not Intersect(Target, toggleCell) Is Nothing Then
        Dim currentValue As String
        Static previousValue As String ' Static variable to store the previous value
        currentValue = UCase(Trim(toggleCell.Value))
        
        ' Debug: Log the current and previous values
        Debug.Print "Toggle cell (J2) current value: " & currentValue
        Debug.Print "Toggle cell (J2) previous value: " & previousValue
        Debug.Print "Target value before trim: |" & Target.Value & "|"
        
        ' Check for toggle between "JAN-JUN" and "JUL-DEC" using previous value
        If ((previousValue = "JAN-JUN" And currentValue = "JUL-DEC") Or _
            (previousValue = "JUL-DEC" And currentValue = "JAN-JUN")) Then
            ' Debug: Log that the toggle condition is met
            Debug.Print "Toggle condition met: From " & previousValue & " to " & currentValue
            Call ClearTableContent.ClearTableContent
        Else
            ' Debug: Log when the toggle condition is not met
            Debug.Print "Toggle condition not met: Previous = " & previousValue & ", Current = " & currentValue
        End If
        
        ' Update the previous value for the next change
        previousValue = currentValue
    Else
        ' Debug: Log when the change is outside the toggle cell
        Debug.Print "Change detected outside toggle cell (J2): " & Target.Address
    End If
End Sub
Public Sub ResetAOHCounter()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim isAllOne As Boolean

    Set ws = ThisWorkbook.Sheets("PersonnelList (AOH & Desk)")

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    isAllOne = True

    For i = 12 To lastRow
        If ws.Cells(i, 6).Value <> 1 Then
            isAllOne = False
            
            Exit For
        End If
    Next i

    ' Reset if all have AOH = 1
    If isAllOne Then
        For i = 12 To lastRow
            ws.Cells(i, 6).Value = 0
        Next i
    End If
End Sub

Sub ResetDutiesAOHCounter()
'
' Reset_Duties_AOH_Counter Macro
'

'
    Sheets("PersonnelList (AOH & Desk)").Select
    Range("E12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E12:F12").Select
    Selection.AutoFill Destination:=Range( _
        "Desk_PersonnelList[[Weekly Duties Counter]:[AOH Counter]]")
    Sheets("MasterCopy").Select
End Sub

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

Public Sub ClearTableContent()
'
' ClearTableContent Macro
'

'
    Range("D6:O189").Select
    Selection.ClearContents
    Selection.Rows.AutoFit
End Sub
