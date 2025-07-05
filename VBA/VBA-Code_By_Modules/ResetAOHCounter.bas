Attribute VB_Name = "ResetAOHCounter"
Public Sub ResetAOHCounter()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim isAllOne As Boolean

    Set ws = ThisWorkbook.Sheets("PersonnelList (AOH & Desk)")

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    isAllOne = True

    For i = 12 To lastRow
        If ws.Cells(i, 6).Value <> 1 Then
            isAllOne = False
            
            Exit For
        End If
    Next i

    ' Reset if all have AOH = 1
    If isAllOne Then
        For i = 12 To lastRow
            ws.Cells(i, 6).Value = 0
        Next i
    End If
End Sub

