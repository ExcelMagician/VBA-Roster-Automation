Attribute VB_Name = "AssignEmployeeCopy"
'declare worksheet and table
    Private wsRosterCopy As Worksheet
    Private wsPersonnel As Worksheet
    Private wsSettings As Worksheet
    Private morningtbl As ListObject
    
'declare roster column number
    Private dateCol As Long
    Private dayCol As Long
    Private LMBCol As Long
    Private morCol As Long
    Private aftCol As Long
    Private AOHCol As Long
    Private satAOHCol1 As Long
    Private satAOHCol2 As Long
    
Sub AssignFirstEmployeeToFirstSlotCopy()
    Set wsRosterCopy = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSettings = Sheets("Settings")

    dateCol = 2
    dayCol = 3
    LMBCol = 4
    morCol = 6
    aftCol = 8
    AOHCol = 10
    satAOHCol1 = 12
    satAOHCol2 = 14
    
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
     
        currDate = wsRosterCopy.Cells(dateRow, dateCol).Value
        
        If Weekday(currDate, vbMonday) = 7 Or _
            Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0 Then
            
            ' Skip this date by marking all slots as "CLOSED"
            wsRosterCopy.Cells(dateRow, LMBCol).Value = "CLOSED" ' D column
            wsRosterCopy.Cells(dateRow, LMBCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, morCol).Value = "CLOSED" ' F column
            wsRosterCopy.Cells(dateRow, morCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, aftCol).Value = "CLOSED" ' H column
            wsRosterCopy.Cells(dateRow, aftCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, AOHCol).Value = "CLOSED" ' J column
            wsRosterCopy.Cells(dateRow, AOHCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, satAOHCol1).Value = "CLOSED" ' L column
            wsRosterCopy.Cells(dateRow, satAOHCol1).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, satAOHCol2).Value = "CLOSED" ' N column
            wsRosterCopy.Cells(dateRow, satAOHCol2).Interior.Color = vbRed
            GoTo NextDate ' Skip to the next date
        End If
        
        For Each slotCol In Array(LMBCol, morCol, aftCol, AOHCol, satAOHCol1, satAOHCol2) ' D, F, H, J, L, N columns
            Set slotCell = wsRosterCopy.Cells(dateRow, slotCol)
            slotCell.Interior.ColorIndex = xlNone ' Reset to no fill (default)
            slotCell.Font.Strikethrough = False
        Next slotCol
        
        isSaturday = (Weekday(currDate, vbMonday) = 6)
        
        isVacation = (wsRosterCopy.Cells(dateRow, 1).Value = "Vacation")
        
        If isSaturday Then
            slotCols = Array(satAOHCol1, satAOHCol2) ' L, N for Saturday
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
            found = Falsez
            
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

Function countMorningOrAfternoonSlotsUDF() As Long
    Application.Volatile
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim r As Long
    Dim lastRow As Long
    Dim holidayCell As Range
    Dim isHoliday As Boolean
    Dim ws As Worksheet
    
    ' Initialize counter
    countMorningOrAfternoonSlotsUDF = 0
    
    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets("MasterCopy")
    
    ' Get the fixed start and end dates from H3 and K3
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Function ' Exit if dates are invalid
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    ' Determine the last row with a valid date in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Do While lastRow >= 6
        If IsDate(Trim(ws.Cells(lastRow, 2).Value)) Then
            Exit Do
        Else
            lastRow = lastRow - 1
        End If
    Loop
    If lastRow < 6 Then Exit Function ' No valid dates found
    
    ' Loop through each row in column B
    For r = 6 To lastRow
        currentDate = ws.Cells(r, 2).Value ' Date from column B
        If IsDate(Trim(currentDate)) Then
            ' Check if the date is within the custom period
            If currentDate >= startDate And currentDate <= endDate Then
                ' Check if it's not Sunday (1) or Saturday (7)
                If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
                    ' Check if it's not a public holiday using the named range
                    isHoliday = False
                    For Each holidayCell In Range("Settings_Holidays")
                        If IsDate(holidayCell.Value) Then
                            If DateValue(currentDate) = DateValue(holidayCell.Value) Then
                                isHoliday = True
                                Exit For
                            End If
                        End If
                    Next holidayCell
                    ' If not a holiday, increment counter
                    If Not isHoliday Then
                        countMorningOrAfternoonSlotsUDF = countMorningOrAfternoonSlotsUDF + 1
                    End If
                End If
            End If
        End If
    Next r
End Function
Function countAOHslotsUDF() As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim r As Long
    Dim lastRow As Long
    Dim holidayCell As Range
    Dim isHoliday As Boolean
    Dim ws As Worksheet
    
    ' Initialize counter
    countAOHslotsUDF = 0
    
    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets("MasterCopy")
    
    ' Get the fixed start and end dates from H3 and K3
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Function ' Exit if dates are invalid
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    ' Determine the last row with a valid date in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Do While lastRow >= 6
        If IsDate(Trim(ws.Cells(lastRow, 2).Value)) Then
            Exit Do
        Else
            lastRow = lastRow - 1
        End If
    Loop
    If lastRow < 6 Then Exit Function ' No valid dates found
    
    ' Loop through each row in column B
    For r = 6 To lastRow
        currentDate = ws.Cells(r, 2).Value ' Date from column B
        If IsDate(Trim(currentDate)) Then
            ' Check if the date is within the custom period
            If currentDate >= startDate And currentDate <= endDate Then
                ' Check if it's not Sunday (1) or Saturday (7)
                If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
                    ' Check if it's not a public holiday using the named range
                    isHoliday = False
                    For Each holidayCell In Range("Settings_Holidays")
                        If IsDate(holidayCell.Value) Then
                            If DateValue(currentDate) = DateValue(holidayCell.Value) Then
                                isHoliday = True
                                Exit For
                            End If
                        End If
                    Next holidayCell
                    ' Check if the corresponding marker in column A is "sem time"
                    If Not isHoliday And LCase(Trim(ws.Cells(r, 1).Value)) = "sem time" Then
                        countAOHslotsUDF = countAOHslotsUDF + 1
                    End If
                End If
            End If
        End If
    Next r
End Function

Function countSatAOH() As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim r As Long
    Dim lastRow As Long
    Dim holidayCell As Range
    Dim isHoliday As Boolean
    Dim ws As Worksheet
    
    ' Initialize counter
    countSatAOH = 0
    
    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets("MasterCopy")
    
    ' Get the fixed start and end dates from H3 and K3
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Function ' Exit if dates are invalid
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    ' Determine the last row with a valid date in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Do While lastRow >= 6
        If IsDate(Trim(ws.Cells(lastRow, 2).Value)) Then
            Exit Do
        Else
            lastRow = lastRow - 1
        End If
    Loop
    If lastRow < 6 Then Exit Function ' No valid dates found
    
    ' Loop through each row in column B
    For r = 6 To lastRow
        currentDate = ws.Cells(r, 2).Value ' Date from column B
        If IsDate(Trim(currentDate)) Then
            ' Check if the date is within the custom period
            If currentDate >= startDate And currentDate <= endDate Then
                ' Check if it's a Saturday (7)
                If Weekday(currentDate) = 7 Then
                    ' Check if it's not a public holiday using the named range
                    isHoliday = False
                    For Each holidayCell In Range("Settings_Holidays")
                        If IsDate(holidayCell.Value) Then
                            If DateValue(currentDate) = DateValue(holidayCell.Value) Then
                                isHoliday = True
                                Exit For
                            End If
                        End If
                    Next holidayCell
                    ' If not a holiday, increment counter
                    If Not isHoliday Then
                        countSatAOH = countSatAOH + 1
                    End If
                End If
            End If
        End If
    Next r
End Function
