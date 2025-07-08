Attribute VB_Name = "AssignAOHDuties"
' Declare worksheet and table
Private wsRosterCopy As Worksheet
Private wsPersonnel As Worksheet
Private wsSettings As Worksheet
Private aohtbl As ListObject
Private spectbl As ListObject

' Declare roster column numbers
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

Sub AssignAOHDuties()
    Set wsRosterCopy = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("AOH PersonnelList")
    Set aohtbl = wsPersonnel.ListObjects("AOHMainList")
    Set spectbl = wsPersonnel.ListObjects("AOHSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim staffName As String
    Dim workDays As Variant

    totalDays = wsRosterCopy.Range(wsRosterCopy.Cells(START_ROW, DATE_COL), wsRosterCopy.Cells(186, DATE_COL)).Rows.Count
    
    ' Step 1: Assign Specific Days Staff with weekly limit
    For i = 1 To spectbl.ListRows.Count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print "Processing Specific Days staff: " & staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from MorningMainList
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
            weekStart = r - (Weekday(wsRosterCopy.Cells(r, DATE_COL).Value, vbMonday) - 1)
            If weekStart < START_ROW Then weekStart = START_ROW
            weekEnd = weekStart + 6
            If weekEnd >= 186 Then weekEnd = 186 - 1
            Debug.Print "  Week boundaries: Start = " & weekStart & ", End = " & weekEnd
            
            Dim dutyCount As Long
            dutyCount = 0
            For k = weekStart To weekEnd
                If k >= START_ROW And k < 186 And wsRosterCopy.Cells(k, AOH_COL).Value = staffName And _
                   UCase(Trim(wsRosterCopy.Cells(k, VAC_COL).Value)) = "SEM TIME" Then
                    dutyCount = dutyCount + 1
                End If
            Next k
            Debug.Print "  Current duty count in week: " & dutyCount
            
            If wsRosterCopy.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, 186) Then
                wsRosterCopy.Cells(r, AOH_COL).Value = staffName
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
    For r = START_ROW To 186
        If wsRosterCopy.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRosterCopy.Cells(r, AOH_COL).Value = "CLOSED" Then GoTo SkipDay
        ' Check if the day is sem time (not vacation)
        If UCase(Trim(wsRosterCopy.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then GoTo SkipDay
        
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
            If wsRosterCopy.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, 186) Then
                wsRosterCopy.Cells(r, AOH_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                ' Debug.Print "Assigned All Days staff " & staffName & " to row " & r
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
    
    For r = START_ROW To 186
        dayName = Trim(wsRosterCopy.Cells(r, DAY_COL).Value)
        Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRosterCopy.Cells(r, AOH_COL)) Then
            Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is within sem time
        If UCase(Trim(wsRosterCopy.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then
            Debug.Print "  -> Skipped (Not Sem Time)"
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
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.Row - aohtbl.HeaderRowRange.Row

        ' Increment Duties Counter
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
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
    
    Set ws = wsRosterCopy
    weekStart = rowNum - (Weekday(ws.Cells(rowNum, DATE_COL).Value, vbMonday) - 1)
    If weekStart < startRow Then weekStart = startRow
    weekEnd = weekStart + 6
    If weekEnd >= endRow Then weekEnd = endRow - 1
    
    ' Debug.Print "Row " & rowNum & ": Week Start = " & weekStart & ", Week End = " & weekEnd
    dutyCount = 0
    For i = weekStart To weekEnd
        If i >= startRow And i < endRow And ws.Cells(i, AOH_COL).Value = staffName And UCase(Trim(ws.Cells(i, VAC_COL).Value)) = "SEM TIME" Then
            dutyCount = dutyCount + 1
            If dutyCount >= 1 Then
                CheckWeeklyLimit = False
                Exit Function
            End If
        End If
    Next i
    CheckWeeklyLimit = True
End Function

