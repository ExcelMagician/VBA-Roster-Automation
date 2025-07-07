Attribute VB_Name = "InsertStaffCounter"
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

