Attribute VB_Name = "AssignEmployee"

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

