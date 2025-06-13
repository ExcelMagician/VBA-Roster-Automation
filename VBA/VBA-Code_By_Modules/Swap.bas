Attribute VB_Name = "Swap"
Sub SwapStaff()
    Dim wsRoster As Worksheet
    Dim wsPersonnel As Worksheet
    Dim wsSwap As Worksheet
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim dateCell As Range
    Dim dateRange As Range
    Dim oriName As String
    Dim newName As String
    Dim r As Long
    Dim i As Long
    Dim lastRow As Long

    ' Set worksheet references
    Set wsRoster = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")

    ' Get names from Swap sheet
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))
    
    ' Check for empty values
    If Len(oriName) = 0 Then
        MsgBox "Error: Original staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Error: New staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    
    ' Prompt user to select date cells in Column A
    On Error Resume Next
    Set dateRange = Application.InputBox("Select date cells (Column A)", Type:=8)
    On Error GoTo 0
    If dateRange Is Nothing Then Exit Sub

    ' Define slot columns: F, H, J, L, N (6, 8, 10, 12, 14)
    slotCols = Array(6, 8, 10, 12, 14)
    
    ' Check if oriName exists in the selected date rows across slot columns
    Dim col As Variant
    Dim oriNameExists As Boolean
    oriNameExists = False
    For Each dateCell In dateRange
        r = dateCell.Row
        For Each col In slotCols
            If UCase(Trim(wsRoster.Cells(r, col).Value)) = oriName Then
                oriNameExists = True
                Exit For
            End If
        Next col
        If oriNameExists Then Exit For
    Next dateCell
    
    If Not oriNameExists Then
        MsgBox "Error: " & oriName & " not found in the selected rows. Swap not allowed.", vbCritical
        Exit Sub
    End If

    ' Loop over selected date rows
    For Each dateCell In dateRange
        r = dateCell.Row

        ' Check if newName exists in the same row across all slot columns
        Dim nameExists As Boolean
        nameExists = False
        For Each col In slotCols
            If UCase(Trim(wsRoster.Cells(r, col).Value)) = newName Then
                nameExists = True
                Exit For
            End If
        Next col
        
        If nameExists Then
            MsgBox "Error: " & newName & " already exists in row " & r & ". Swap not allowed.", vbCritical
        Else
            For Each slotCol In slotCols
                With wsRoster.Cells(r, slotCol)
                    Dim currentName As String
                    ' Determine the current name based on whether there’s a line break
                    If InStr(.Value, vbNewLine) > 0 Then
                        currentName = Trim(Split(.Value, vbNewLine)(0)) ' First unstriked line for subsequent swaps
                    Else
                        currentName = Trim(.Value) ' Entire value for initial swap
                    End If
                    
                    If UCase(currentName) = oriName Then ' Check the current name
                        ' Add new name first (unstriked) and preserve existing content
                        .Value = newName & vbNewLine & .Value
                        .VerticalAlignment = xlTop ' Align text to the top
                        .WrapText = True
                        
                        ' Split into lines to apply strikethrough to all previous names
                        Dim lines() As String
                        lines = Split(.Value, vbNewLine)
                        Dim k As Integer
                        Dim cumulativeLength As Long
                        cumulativeLength = Len(newName) + 2 ' Start with newName and its vbNewLine
                        
                        ' Apply strikethrough to all lines except the first one
                        For k = 1 To UBound(lines)
                            Dim startPos As Integer
                            startPos = cumulativeLength
                            .Characters(startPos, Len(lines(k)) + 1).Font.Strikethrough = True
                            cumulativeLength = cumulativeLength + Len(lines(k)) + 2 ' Update for next line
                        Next k
                        
                        ' Explicitly increase row height by 15 points per swap
                        .RowHeight = .RowHeight + 15
                        
                        ' Update personnel counter for the new staff
                        lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).Row
                        ' Deduct duties from the old staff
                        For i = 12 To lastRow
                            If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = oriName Then
                                wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value - 1 ' Decrement Weekly Duties Counter
                                If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                                    wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value - 1 ' Decrement AOH Counter
                                End If
                                Exit For
                            End If
                        Next i
                        ' Update duties for the new staff
                        For i = 12 To lastRow
                            If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = newName Then
                                wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value + 1 ' Increment Weekly Duties Counter
                                If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                                    wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value + 1 ' Increment AOH Counter
                                End If
                                Exit For
                            End If
                        Next i
                    End If
                End With
            Next slotCol
        End If
    Next dateCell

    MsgBox "Swap complete.", vbInformation
End Sub
