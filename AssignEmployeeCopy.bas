Attribute VB_Name = "AssignEmployeeCopy"
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
            
        
        Call ResetAOHCounter
        
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






