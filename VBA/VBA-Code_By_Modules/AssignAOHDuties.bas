Attribute VB_Name = "AssignAOHDuties"
' Declare worksheet and table
Private wsRoster As Worksheet
Private wsPersonnel As Worksheet
Private wsSettings As Worksheet
Private aohtbl As ListObject
Private spectbl As ListObject

Sub AssignAOHDuties()
    Set wsRoster = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("AOH PersonnelList")
    Set aohtbl = wsPersonnel.ListObjects("AOHMainList")
    Set spectbl = wsPersonnel.ListObjects("AOHSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long, k As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim staffName As String
    Dim workDays As Variant

    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(LAST_ROW_ROSTER, DATE_COL)).Rows.Count
    
    ' Step 1: Assign Specific Days Staff with weekly limit
    For i = 1 To spectbl.ListRows.Count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print "Processing Specific Days staff: " & staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from AOHMainList
        For r = 1 To aohtbl.ListRows.Count
            If aohtbl.DataBodyRange(r, aohtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = aohtbl.DataBodyRange(r, aohtbl.ListColumns("Max Duties").Index).Value
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
        
        ' Assign staff with weekly limit check
        Dim assigned As Long
        assigned = 0
        For j = 1 To eligibleRows.Count
            If assigned >= maxDuties Then Exit For
            r = tmpRows(j)
            Debug.Print "Considering row " & r & " for " & staffName & " (Shuffled index: " & j & ")"
            
            Dim weekStart As Long, weekEnd As Long
            weekStart = r - (Weekday(wsRoster.Cells(r, DATE_COL).Value, vbMonday) - 1)
            If weekStart < START_ROW Then weekStart = START_ROW
            weekEnd = weekStart + 6
            If weekEnd >= LAST_ROW_ROSTER Then weekEnd = LAST_ROW_ROSTER - 1
            Debug.Print "  Week boundaries: Start = " & weekStart & ", End = " & weekEnd
            
            Dim dutyCount As Long
            dutyCount = 0
            For k = weekStart To weekEnd
                If k >= START_ROW And k < LAST_ROW_ROSTER And wsRoster.Cells(k, AOH_COL).Value = staffName And _
                   UCase(Trim(wsRoster.Cells(k, VAC_COL).Value)) = "SEM TIME" Then
                    dutyCount = dutyCount + 1
                End If
            Next k
            Debug.Print "  Current duty count in week: " & dutyCount
            
            If wsRoster.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, LAST_ROW_ROSTER) Then
                wsRoster.Cells(r, AOH_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assigned = assigned + 1
                Debug.Print "  Assigned " & staffName & " to row " & r & " (Duty count: " & assigned & ")"
            Else
                Debug.Print "  Skipped row " & r & " for " & staffName & " (Limit reached or slot taken)"
            End If
        Next j
        Debug.Print "Total assigned to " & staffName & ": " & assigned
    Next i
    
    ' Step 2: Assign All Days Staff with weekly limit
    For r = START_ROW To LAST_ROW_ROSTER
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, AOH_COL).Value = "CLOSED" Then GoTo SkipDay
        ' Check if the day is sem time (not vacation)
        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then GoTo SkipDay
        
        For i = 1 To aohtbl.ListRows.Count
            staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
            If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
            Dim currDuties As Long
            currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            
            ' Assign from top with weekly limit check
            If wsRoster.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, LAST_ROW_ROSTER) Then
                wsRoster.Cells(r, AOH_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Debug.Print "Assigned All Days staff " & staffName & " to row " & r
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
        For i = 1 To aohtbl.ListRows.Count
            staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
            maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
            currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
            If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And currDuties < maxDuties Then
                needsReassignment = True
                Debug.Print "Staff " & staffName & " needs reassignment (Current: " & currDuties & ", Max: " & maxDuties & ")"
            End If
        Next i
        
        If needsReassignment Then
            ' Remove existing assignments for All Days staff in last 2 weeks and collect staff needing reassignment
            Dim removedStaff() As String
            ReDim removedStaff(0)
            For r = lastTwoWeeksStart To LAST_ROW_ROSTER
                If wsRoster.Cells(r, AOH_COL).Value <> "" And wsRoster.Cells(r, AOH_COL).Value <> "CLOSED" Then
                    staffName = wsRoster.Cells(r, AOH_COL).Value
                    For i = 1 To aohtbl.ListRows.Count
                        If aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value = staffName And _
                           UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                            wsRoster.Cells(r, AOH_COL).Value = ""
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
            For i = 1 To aohtbl.ListRows.Count
                staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                    maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                    currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
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
                If wsRoster.Cells(r, AOH_COL).Value <> "" Or wsRoster.Cells(r, AOH_COL).Value = "CLOSED" Then GoTo SkipReassignDay
                If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then GoTo SkipReassignDay
                
                For i = 1 To UBound(tmpStaff)
                    staffName = tmpStaff(i)
                    For j = 1 To aohtbl.ListRows.Count
                        If aohtbl.DataBodyRange(j, aohtbl.ListColumns("Name").Index).Value = staffName And _
                           UCase(aohtbl.DataBodyRange(j, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
                            maxDuties = aohtbl.DataBodyRange(j, aohtbl.ListColumns("Max Duties").Index).Value
                            currDuties = aohtbl.DataBodyRange(j, aohtbl.ListColumns("Duties Counter").Index).Value
                            If currDuties < maxDuties And CheckWeeklyLimit(staffName, r, START_ROW, LAST_ROW_ROSTER) Then
                                wsRoster.Cells(r, AOH_COL).Value = staffName
                                Call IncrementDutiesCounter(staffName)
                                Debug.Print "Reassigned All Days staff " & staffName & " to row " & r & " (Last 2 weeks)"
                                assignedInThisPass = True
                                Exit For
                            Else
                                Debug.Print "  Skipped " & staffName & " at row " & r & " (Limit reached or same-day conflict)"
                            End If
                        End If
                    Next j
                    If wsRoster.Cells(r, AOH_COL).Value <> "" Then Exit For
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
        'Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, AOH_COL)) Then
            'Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is within sem time
        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then
            'Debug.Print "  -> Skipped (Not Sem Time)"
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
    Debug.Print "Total Eligible: " & eligibleRows.Count
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row

        ' Increment Duties Counter
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
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
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row

        ' Decrement Duties Counter
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function CheckWeeklyLimit(staffName As String, rowNum As Long, startRow As Long, endRow As Long) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim weekStart As Long
    Dim weekEnd As Long
    Dim dutyCount As Long
    
    Set ws = wsRoster
    weekStart = rowNum - (Weekday(ws.Cells(rowNum, DATE_COL).Value, vbMonday) - 1)
    If weekStart < startRow Then weekStart = startRow
    weekEnd = weekStart + 6
    If weekEnd >= endRow Then weekEnd = endRow - 1
    
    Debug.Print "Row " & rowNum & ": Week Start = " & weekStart & ", End = " & weekEnd
    dutyCount = 0
    For i = weekStart To weekEnd
        If i >= startRow And i < endRow And (ws.Cells(i, AOH_COL).Value = staffName Or ws.Cells(i, SAT_AOH_COL1).Value = staffName Or ws.Cells(i, SAT_AOH_COL2).Value = staffName) And UCase(Trim(ws.Cells(i, VAC_COL).Value)) = "SEM TIME" Then
            dutyCount = dutyCount + 1
            If dutyCount >= 1 Then
                CheckWeeklyLimit = False
                Exit Function
            End If
        End If
    Next i
    CheckWeeklyLimit = True
End Function

