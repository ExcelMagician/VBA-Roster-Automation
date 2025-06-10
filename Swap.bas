Attribute VB_Name = "Swap"
Sub SwapStaff()
    Dim wsRoster As Worksheet, wsPersonnel As Worksheet, wsSwap As Worksheet
    Dim slotCols As Variant, slotCol As Variant
    Dim dateCell As Range, dateRange As Range
    Dim oriName As String, newName As String
    Dim r As Long, i As Long
    Dim lastRow As Long

    Set wsRoster = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")

    ' Get names
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))

    ' Prompt user to select date cells in Column A
    On Error Resume Next
    Set dateRange = Application.InputBox("Select date cells (Column A)", Type:=8)
    On Error GoTo 0
    If dateRange Is Nothing Then Exit Sub

    ' Columns: F, H, J, L, N
    slotCols = Array(6, 8, 10, 12, 14)

    ' Loop over selected date rows
    For Each dateCell In dateRange
        r = dateCell.Row

        For Each slotCol In slotCols
            With wsRoster.Cells(r, slotCol)
                If UCase(Trim(.Value)) = oriName Then
                    ' Add new name on a new line and apply strikethrough to original name
                    .Value = .Value & vbNewLine & newName
                    .VerticalAlignment = xlTop ' Align text to the top
                    .WrapText = True
                    .Characters(1, Len(oriName)).Font.Strikethrough = True
                    .Characters(Len(.Value) - Len(newName) + 1).Font.Strikethrough = False
                    .Rows.AutoFit

                    ' Update personnel counter for the new staff
                    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
                    For i = 12 To lastRow
                        If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = newName Then
                            wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value + 1 ' Weekly Duties Counter
                            If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                                wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value + 1
                            End If
                            Exit For
                        End If
                    Next i
                End If
            End With
        Next slotCol
    Next dateCell

    MsgBox "Swap complete.", vbInformation
End Sub
