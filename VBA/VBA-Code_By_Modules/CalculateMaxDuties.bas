Attribute VB_Name = "CalculateMaxDuties"
Sub CalculateMaxDuties()
    Dim ws As Worksheet
    Dim morningtbl As ListObject
    Set ws = ThisWorkbook.Sheets("PersonnelList Copy")
    Set morningtbl = ws.ListObjects("MorningMainList")

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

    totalStaff = morningtbl.ListRows.Count
    totalDuties = ws.Range("H6").Value
    fullDuties = WorksheetFunction.RoundDown(totalDuties / totalStaff, 0)
    remaining = 0
    eligibleCount = 0
    
    ReDim eligible100(1 To totalStaff)
    ReDim rounded(1 To totalStaff)
    
    'Calculate initial duties and max cap
    For i = 1 To totalStaff
        dutiesPercentage = morningtbl.ListRows(i).Range.Cells(morningtbl.ListColumns("Duties Percentage (%)").Index).Value
        
        If dutiesPercentage < 100 Then
            rounded(i) = CLng(fullDuties * (dutiesPercentage / 100))
        Else
            rounded(i) = fullDuties
            'Mark eligible 100% staff for distribution
            eligibleCount = eligibleCount + 1
            eligible100(eligibleCount) = i
        End If
        
        totalAssigned = totalAssigned + rounded(i)
    Next i
    
    'Distribute remaining slots to 100% staff
    remaining = totalDuties - totalAssigned
    
    If remaining > 0 Then
        If eligibleCount > 0 Then
            For j = 1 To remaining
                i = eligible100(((j - 1) Mod eligibleCount) + 1) 'rotate among 100% staff
                rounded(i) = rounded(i) + 1
            Next j
        Else
            MsgBox "No available staff to assign remaining duties"
        End If
    End If
    
    'Write results back to sheet
    For i = 1 To totalStaff
        morningtbl.ListRows(i).Range.Cells(morningtbl.ListColumns("Max Duties").Index).Value = rounded(i)
    Next i
    
End Sub
