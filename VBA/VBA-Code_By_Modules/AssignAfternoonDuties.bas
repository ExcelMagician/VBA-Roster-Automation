Attribute VB_Name = "AssignAfternoonDuties"
' Declare worksheet and table
Private wsRoster As Worksheet
Private wsPersonnel As Worksheet
Private wsSettings As Worksheet
Private afternoontbl As ListObject
Private spectbl As ListObject

Sub AssignAfternoonDuties()
    Set wsRoster = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Afternoon PersonnelList")
    Set afternoontbl = wsPersonnel.ListObjects("AfternoonMainList")
    Set spectbl = wsPersonnel.ListObjects("AfternoonSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim candidates() As String
    Dim staffName As String
    Dim workDays As Variant

    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(LAST_ROW_ROSTER, DATE_COL)).Rows.Count
    Debug.Print "afternoon assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.Count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from MorningMainList
        For r = 1 To afternoontbl.ListRows.Count
            If afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Max Duties").Index).Value
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
        Dim assignedCount As Long
        assignedCount = 0
        
        For j = 1 To eligibleRows.Count
            If assignedCount >= maxDuties Then Exit For
        
            If Not IsWorkingOnSameDay(tmpRows(j), staffName) Then
                wsRoster.Cells(tmpRows(j), AFT_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
            End If
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To LAST_ROW_ROSTER
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To afternoontbl.ListRows.Count
            staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
            If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
            Dim currDuties As Long
            currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top
            If wsRoster.Cells(r, AFT_COL).Value = "" Then
                wsRoster.Cells(r, AFT_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    ' Step 3: Reassign All Days Staff for last 2 weeks with random assignment
    Dim needsReassignment As Boolean
    Do
        needsReassignment = False
        Dim lastTwoWeeksStart As Long
        lastTwoWeeksStart = LAST_ROW_ROSTER - 13 ' Last 14 days (2 weeks)
        If lastTwoWeeksStart < START_ROW Then lastTwoWeeksStart = START_ROW
        
        ' Check if any All Days staff need reassignment
        For i = 1 To afternoontbl.ListRows.Count
            staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
            maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
            currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
            If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And currDuties < maxDuties Then
                needsReassignment = True
                Debug.Print "Staff " & staffName & " needs reassignment (Current: " & currDuties & ", Max: " & maxDuties & ")"
            End If
        Next i
        
        If needsReassignment Then
            ' Remove existing assignments for All Days staff in last 2 weeks and collect staff needing reassignment
            Dim removedStaff() As String
            ReDim removedStaff(0)
            For r = lastTwoWeeksStart To LAST_ROW_ROSTER
                If wsRoster.Cells(r, AFT_COL).Value <> "" And wsRoster.Cells(r, AFT_COL).Value <> "CLOSED" Then
                    staffName = wsRoster.Cells(r, AFT_COL).Value
                    For i = 1 To afternoontbl.ListRows.Count
                        If afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value = staffName And _
                           UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                            wsRoster.Cells(r, AFT_COL).Value = ""
                            Call DecrementDutiesCounter(staffName)
                            Debug.Print "Removed " & staffName & " from row " & r & " (Last 2 weeks)"
                            ' Add to removedStaff array if not already present
                            Dim found As Boolean
                            found = False
                            For j = LBound(removedStaff) To UBound(removedStaff)
                                If removedStaff(j) = staffName Then found = True
                            Next j
                            If Not found Then
                                ReDim Preserve removedStaff(UBound(removedStaff) + 1)
                                removedStaff(UBound(removedStaff)) = staffName
                            End If
                            Exit For
                        End If
                    Next i
                End If
            Next r
            
            ' Build staff pool with only All Days staff needing reassignment (currDuties < maxDuties)
            Dim staffPool() As String
            ReDim staffPool(0)
            For i = 1 To afternoontbl.ListRows.Count
                staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
                If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                    maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
                    currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
                    If currDuties < maxDuties Then
                        found = False
                        For j = LBound(staffPool) To UBound(staffPool)
                            If staffPool(j) = staffName Then found = True
                        Next j
                        If Not found Then
                            ReDim Preserve staffPool(UBound(staffPool) + 1)
                            staffPool(UBound(staffPool)) = staffName
                        End If
                    End If
                End If
            Next i
            
            ' Combine removed staff with staff pool (avoid duplicates)
            For j = LBound(removedStaff) To UBound(removedStaff)
                If removedStaff(j) <> "" Then
                    found = False
                    For k = LBound(staffPool) To UBound(staffPool)
                        If staffPool(k) = removedStaff(j) Then found = True
                    Next k
                    If Not found Then
                        ReDim Preserve staffPool(UBound(staffPool) + 1)
                        staffPool(UBound(staffPool)) = removedStaff(j)
                    End If
                End If
            Next j
            
            ' Shuffle staff pool
            Dim tmpStaff() As String
            If UBound(staffPool) > 0 Then
                ReDim tmpStaff(1 To UBound(staffPool))
                For j = 1 To UBound(staffPool)
                    tmpStaff(j) = staffPool(j)
                Next j
                Call ShuffleArrayString(tmpStaff)
            End If
            
            ' Reassign randomly from staff pool
            Dim assignedInThisPass As Boolean
            assignedInThisPass = False
            For r = lastTwoWeeksStart To LAST_ROW_ROSTER
                If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipReassignDay
                If wsRoster.Cells(r, AFT_COL).Value <> "" Or wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Then GoTo SkipReassignDay
                
                For i = 1 To UBound(tmpStaff)
                    staffName = tmpStaff(i)
                    For j = 1 To afternoontbl.ListRows.Count
                        If afternoontbl.DataBodyRange(j, afternoontbl.ListColumns("Name").Index).Value = staffName And _
                           UCase(afternoontbl.DataBodyRange(j, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                            maxDuties = afternoontbl.DataBodyRange(j, afternoontbl.ListColumns("Max Duties").Index).Value
                            currDuties = afternoontbl.DataBodyRange(j, afternoontbl.ListColumns("Duties Counter").Index).Value
                            If currDuties < maxDuties And Not IsWorkingOnSameDay(r, staffName) Then
                                wsRoster.Cells(r, AFT_COL).Value = staffName
                                Call IncrementDutiesCounter(staffName)
                                Debug.Print "Reassigned All Days staff " & staffName & " to row " & r & " (Last 2 weeks)"
                                assignedInThisPass = True
                                Exit For
                            Else
                                Debug.Print "  Skipped " & staffName & " at row " & r & " (Limit reached or same-day conflict)"
                            End If
                        End If
                    Next j
                    If wsRoster.Cells(r, AFT_COL).Value <> "" Then Exit For
                Next i
SkipReassignDay:
            Next r
            
            ' If no assignments were made in this pass, break the loop to avoid infinite cycling
            If Not assignedInThisPass Then needsReassignment = False
        End If
    Loop Until Not needsReassignment
    
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

' Helper to shuffle array of strings
Sub ShuffleArrayString(arr() As String)
    Dim i As Long, j As Long, tmp As String
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

    'Debug.Print "=== Checking Eligible Rows ==="
    'Debug.Print "WorkDays:"
    For j = LBound(workDays) To UBound(workDays)
        'Debug.Print "[" & j & "]: " & workDays(j)
    Next j
    
    For r = START_ROW To LAST_ROW_ROSTER
        dayName = Trim(wsRoster.Cells(r, DAY_COL).Value)
        ' Debug: show what day we are checking
        'Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, AFT_COL)) Then
            'Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                'Debug.Print "  -> Added (Matched with " & workDays(j) & ")"
                Exit For
            End If
        Next j
        
SkipRow:
    Next r
    'Debug.Print "Total Eligible: " & eligibleRows.Count
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row

        ' Increment Duties Counter
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row

        ' Decrement Duties Counter
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    ' Check if staff is working on AOH on the same day
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
        
    IsWorkingOnSameDay = False
End Function

