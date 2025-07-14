Attribute VB_Name = "Module1"
Sub InsertStaff()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim morningtbl As ListObject
    Dim newRow As ListRow
    Dim staffName As String, dept As String
    Dim availType As String, workDays As String, percentage As String
    Dim checkRow As Long
    Dim specificDaysTbl As ListObject
    
    ' Set worksheet and tables
    Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
    If ws Is Nothing Then
        MsgBox "Worksheet 'Morning PersonnelList' not found.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Set morningtbl = ws.ListObjects("MorningMainList")
    Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
    On Error GoTo ErrHandler
    If morningtbl Is Nothing Then
        MsgBox "Table 'MorningMainList' not found on 'Morning PersonnelList'.", vbExclamation
        Exit Sub
    End If
    If specificDaysTbl Is Nothing And availType = "SPECIFIC DAYS" Then
        MsgBox "Table 'MorningSpecificDaysWorkingStaff' not found on 'Morning PersonnelList'.", vbExclamation
        Exit Sub
    End If

    ' Read correct cell values
    staffName = UCase(Trim(ws.Range("D5").Value)) ' Name
    dept = Trim(ws.Range("D6").Value)             ' Department
    availType = UCase(Trim(ws.Range("D7").Value)) ' Availability Type (converted to uppercase for consistency)
    workDays = Trim(ws.Range("D8").Value)         ' Working Days
    percentage = Trim(ws.Range("D9").Value)       ' Duties Percentage

    ' Auto-fill logic based on Availability Type
    If availType = "ALL DAYS" Then
        percentage = "100" ' Auto-fill Duties Percentage to 100%
        workDays = ""     ' Ignore Working Days for All Days
    ElseIf availType = "SPECIFIC DAYS" Then
        If workDays = "" Then
            MsgBox "Please enter Working Days for Specific Days availability.", vbExclamation
            Exit Sub
        End If
        If percentage = "" Or Not IsNumeric(percentage) Or Val(percentage) <= 0 Or Val(percentage) > 100 Then
            MsgBox "Please enter a valid Duties Percentage (1-100) for Specific Days.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Availability Type must be 'All Days' or 'Specific Days'.", vbExclamation
        Exit Sub
    End If

    ' Validation
    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in for both Name and Department.", vbExclamation
        Exit Sub
    End If
    
    ' Check for duplicate names in the table
    For checkRow = 1 To morningtbl.ListRows.Count
        If UCase(Trim(morningtbl.ListRows(checkRow).Range.Cells(1, GetColumnIndex(morningtbl, "Name")).Value)) = staffName Then
            MsgBox "This staff name already exists.", vbExclamation
            Exit Sub
        End If
    Next checkRow

    ' Insert new row at the top of the MorningMainList table
    Set newRow = morningtbl.ListRows.Add(AlwaysInsert:=True)
    
    ' Populate the new row with data using column names
    With newRow.Range
        .Cells(1, GetColumnIndex(morningtbl, "Name")).Value = staffName         ' Name
        .Cells(1, GetColumnIndex(morningtbl, "Department")).Value = dept        ' Department
        .Cells(1, GetColumnIndex(morningtbl, "Availability Type")).Value = availType ' Use entered Availability Type
        .Cells(1, GetColumnIndex(morningtbl, "Duties Percentage (%)")).Value = Val(percentage) ' Duties Percentage
        .Cells(1, GetColumnIndex(morningtbl, "Max Duties")).Value = 0           ' Temporary placeholder
        .Cells(1, GetColumnIndex(morningtbl, "Duties Counter")).Value = 0       ' Default Duties Counter
    End With

    ' Handle specific days workers by inserting into MorningSpecificDaysWorkingStaff table
    If availType = "SPECIFIC DAYS" Then
        Dim specificRow As ListRow
        Set specificRow = specificDaysTbl.ListRows.Add(AlwaysInsert:=True)
        With specificRow.Range
            .Cells(1, GetColumnIndex(specificDaysTbl, "Name")).Value = staffName   ' Name
            .Cells(1, GetColumnIndex(specificDaysTbl, "Working Days")).Value = workDays ' Working Days
        End With
    End If

    ' Call CalculateMaxDuties to update Max Duties for the entire table
    CalculateMaxDuties "MORNING"

    ' Update the newly added row with calculated Max Duties
    With newRow.Range
        Dim maxDutiesIndex As Long
        maxDutiesIndex = GetColumnIndex(morningtbl, "Max Duties")
        If maxDutiesIndex <> -1 Then
            .Cells(1, maxDutiesIndex).Value = morningtbl.ListRows(1).Range.Cells(1, maxDutiesIndex).Value
        Else
            MsgBox "Column 'Max Duties' not found.", vbExclamation
            newRow.Delete
            Exit Sub
        End If
    End With

    ' Clear input
    ws.Range("D5:D9").ClearContents

    MsgBox "Staff added and Max Duties calculated successfully!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    If Not newRow Is Nothing Then newRow.Delete ' Clean up if row was added
    Exit Sub
End Sub

' Helper function to get column index safely
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then
        MsgBox "Column '" & columnName & "' not found in table '" & tbl.Name & "'.", vbExclamation
        GetColumnIndex = -1
    End If
    On Error GoTo 0
End Function

' CalculateMaxDuties subroutine
Sub CalculateMaxDuties(dutyType As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim totalDuties As Long
    Dim totalStaff As Long
    Dim fullDuties As Long
    Dim i As Long
    Dim remaining As Long
    Dim totalAssigned As Long
    Dim dutiesPercentage As Double
    Dim eligibleCount As Long
    Dim eligible100() As Long 'Store the indices of staff with 100% duty
    Dim j As Long
    Dim rounded() As Long

    ' Set worksheet and table based on dutyType
    Select Case UCase(dutyType)
        Case "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
        Case "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("AfternoonPersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
        Case "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
        Case "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
        Case Else
            MsgBox "Invalid duty type. Use 'Morning', 'Afternoon', 'AOH', or 'Sat_AOH'.", vbExclamation
            Exit Sub
    End Select

    totalStaff = tbl.ListRows.Count
    totalDuties = ws.Range("H6").Value
    fullDuties = WorksheetFunction.RoundDown(totalDuties / totalStaff, 0)
    remaining = 0
    eligibleCount = 0
    
    ReDim eligible100(1 To totalStaff)
    ReDim rounded(1 To totalStaff)
    
    ' Calculate initial duties and max cap
    For i = 1 To totalStaff
        dutiesPercentage = tbl.ListRows(i).Range.Cells(GetColumnIndex(tbl, "Duties Percentage (%)")).Value
        
        If dutiesPercentage < 100 Then
            rounded(i) = CLng(fullDuties * (dutiesPercentage / 100))
        Else
            rounded(i) = fullDuties
            ' Mark eligible 100% staff for distribution
            eligibleCount = eligibleCount + 1
            eligible100(eligibleCount) = i
        End If
        
        totalAssigned = totalAssigned + rounded(i)
    Next i
    
    ' Distribute remaining slots to 100% staff
    remaining = totalDuties - totalAssigned
    
    If remaining > 0 Then
        If eligibleCount > 0 Then
            For j = 1 To remaining
                i = eligible100(((j - 1) Mod eligibleCount) + 1) ' Rotate among 100% staff
                rounded(i) = rounded(i) + 1
            Next j
        Else
            MsgBox "No available staff to assign remaining duties for " & dutyType, vbExclamation
        End If
    End If
    
    ' Write results back to sheet
    For i = 1 To totalStaff
        tbl.ListRows(i).Range.Cells(GetColumnIndex(tbl, "Max Duties")).Value = rounded(i)
    Next i
    
    Debug.Print "Max Duties calculated for " & dutyType & " with total duties: " & totalDuties & ", total staff: " & totalStaff
End Sub
