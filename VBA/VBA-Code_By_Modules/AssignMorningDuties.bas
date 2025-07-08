Attribute VB_Name = "AssignMorningDuties"
'declare worksheet and table
    Private wsRosterCopy As Worksheet
    Private wsPersonnel As Worksheet
    Private wsSettings As Worksheet
    Private morningtbl As ListObject
    Private spectbl As ListObject
    
'declare roster column number
    Private Const VAC_COL As Long = 1
    Private Const DATE_COL As Long = 2
    Private Const DAY_COL As Long = 3
    Private Const LMB_COL As Long = 4
    Private Const MOR_COL As Long = 6
    Private Const AFT_COL As Long = 8
    Private Const AOH_COL As Long = 10
    Private Const SAT_AOH_COL1 As Long = 12
    Private Const SAT_AOH_COL2 As Long = 14
    Private Const START_ROW As Long = 6

Sub AssignMorningDuties()
    Set wsRosterCopy = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Morning PersonnelList")
    Set morningtbl = wsPersonnel.ListObjects("MorningMainList")
    Set spectbl = wsPersonnel.ListObjects("MorningSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim candidates() As String
    Dim staffName As String
    Dim workDays As Variant

    totalDays = wsRosterCopy.Range(wsRosterCopy.Cells(START_ROW, DATE_COL), wsRosterCopy.Cells(LAST_ROW_ROSTER, DATE_COL)).Rows.Count
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.Count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from MorningMainList
        For r = 1 To morningtbl.ListRows.Count
            If morningtbl.DataBodyRange(r, morningtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = morningtbl.DataBodyRange(r, morningtbl.ListColumns("Max Duties").Index).Value
                Exit For
            End If
        Next r
        
        ' Build candidate pool of eligible rows
        Dim eligibleRows As Collection
        Set eligibleRows = GetEligibleRows(totalDays, workDays)
        
        ' Shuffle eligibleRows randomly
        Dim tmpRows() As Long
        ReDim tmpRows(1 To eligibleRows.Count)
        For j = 1 To eligibleRows.Count
            tmpRows(j) = eligibleRows(j)
        Next j
        Call ShuffleArray(tmpRows)
        
        ' Assign staff
        For j = 1 To Application.Min(maxDuties, eligibleRows.Count)
            wsRosterCopy.Cells(tmpRows(j), MOR_COL).Value = staffName
            
            ' increment the Duties Counter
            Call IncrementDutiesCounter(staffName)
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To LAST_ROW_ROSTER
        If wsRosterCopy.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRosterCopy.Cells(r, MOR_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To morningtbl.ListRows.Count
            staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
            If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
            currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
            'check if the staff already reach his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            
            ' Assign from top
            If wsRosterCopy.Cells(r, MOR_COL).Value = "" Then
                wsRosterCopy.Cells(r, MOR_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    MsgBox "Duties assignment completed!", vbInformation
End Sub

' Helper to shuffle array
Sub ShuffleArray(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    Randomize
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int(Rnd() * (i - LBound(arr) + 1)) + LBound(arr)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
    Next i
End Sub

Function GetEligibleRows(totalDays As Long, workDays As Variant) As Collection
    Dim eligibleRows As New Collection
    Dim r As Long, j As Long
    Dim dayName As String

    Debug.Print "=== Checking Eligible Rows ==="
    Debug.Print "WorkDays:"
    For j = LBound(workDays) To UBound(workDays)
        Debug.Print "[" & j & "]: " & workDays(j)
    Next j
    
    
    For r = START_ROW To LAST_ROW_ROSTER
        dayName = Trim(wsRosterCopy.Cells(r, DAY_COL).Value)
        ' Debug: show what day we are checking
        Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRosterCopy.Cells(r, MOR_COL)) Then
            Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                Debug.Print "  -> Added (Matched with " & workDays(j) & ")"
                Exit For
            End If
        Next j
        
SkipRow:
    Next r
    Debug.Print "Total Eligible: " & eligibleRows.Count
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = morningtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.Row - morningtbl.HeaderRowRange.Row

        ' Increment Duties Counter
        With morningtbl.ListRows(rowIdx).Range.Cells(morningtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub


