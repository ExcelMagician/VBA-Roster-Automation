Attribute VB_Name = "Swap"
Option Explicit

' Main subroutine to handle staff swapping
Sub SwapStaff()
    Dim wsPersonnel As Worksheet
    Dim wsSwap As Worksheet
    Dim slotCols As Variant
    Dim dateRange As Range
    Dim oriName As String
    Dim newName As String
    Dim dateCell As Range
    
    InitializeWorksheets wsPersonnel, wsSwap
    GetSwapNames wsSwap, oriName, newName
    ValidateNames oriName, newName
    
    Set dateRange = GetDateRange
    If dateRange Is Nothing Then Exit Sub
    If Not IsValidDateColumn(dateRange) Then Exit Sub
    
    slotCols = Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
    Dim r As Long
    For Each dateCell In dateRange
        r = dateCell.row
        Dim oriNameFound As Boolean
        CheckOriginalNameExists wsRoster, r, slotCols, oriName, oriNameFound
        If Not oriNameFound Then
            DisplayError "Error: " & oriName & " not found on date " & wsRoster.Cells(r, DATE_COL).Value & ". Swap not allowed.", vbExclamation
        Else
            Dim nameExists As Boolean
            CheckNewNameExists wsRoster, r, slotCols, newName, nameExists
            If nameExists Then
                DisplayError "Error: " & newName & " already exists on date " & wsRoster.Cells(r, DATE_COL).Value & ". Swap not allowed.", vbExclamation
            Else
                PerformSwap wsRoster, r, slotCols, oriName, newName, wsPersonnel
            End If
        End If
    Next dateCell
    
    wsRoster.Activate
    MsgBox "Swap completed.", vbInformation
End Sub

' Initialize worksheet references
Private Sub InitializeWorksheets(wsPersonnel As Worksheet, wsSwap As Worksheet)
    Set wsRoster = Sheets("Roster")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")
End Sub

' Get original and new staff names from Swap sheet
Private Sub GetSwapNames(wsSwap As Worksheet, oriName As String, newName As String)
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))
End Sub

' Validate that names are not empty
Private Sub ValidateNames(oriName As String, newName As String)
    If Len(oriName) = 0 Then
        MsgBox "Error: Original staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Error: New staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
End Sub

' Prompt user to select date range and return it
Private Function GetDateRange() As Range
    On Error Resume Next
    Set GetDateRange = Application.InputBox("Select date cells (Column B)", Type:=8)
    On Error GoTo 0
End Function

' Validate that the selected range is from column A (column 1)
Private Function IsValidDateColumn(dateRange As Range) As Boolean
    If Not dateRange.Columns.Count = 1 Or dateRange.Column <> 2 Then
        MsgBox "Please only select dates from Date column.", vbExclamation
        IsValidDateColumn = False
    Else
        IsValidDateColumn = True
    End If
End Function

' Check if the original name exists in the row
Private Sub CheckOriginalNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, ByRef oriNameFound As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    Dim currStaff As String
    oriNameFound = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
            Debug.Print "--multiple value"
        Else
            currStaff = UCase(Trim(cellValue))
        End If

        Debug.Print "Checking slot " & col & ": " & currStaff & " vs " & oriName
        
        If currStaff = oriName Then
            oriNameFound = True
            Exit For
        End If
    Next col
End Sub

' Check if the new name exists in the same row
Private Sub CheckNewNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, newName As String, ByRef nameExists As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    nameExists = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            If UCase(Trim(Split(cellValue, vbNewLine)(0))) = newName Then
                nameExists = True
            End If
        Else
            If UCase(Trim(cellValue)) = newName Then
                nameExists = True
            End If
        End If
        If nameExists Then Exit For
    Next col
End Sub

' Display an error message
Private Sub DisplayError(message As String, messageType As VbMsgBoxStyle)
    MsgBox message, messageType
End Sub

' Perform the swap operation for a given row
Private Sub PerformSwap(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, newName As String, wsPersonnel As Worksheet)
    Dim slotCol As Variant
    Dim currentName As String
    Dim lines() As String
    Dim i As Long
    Dim lastRow As Long
    Dim cumulativeLength As Long
    Dim startPos As Integer
    
    For Each slotCol In slotCols
        With wsRoster.Cells(r, slotCol)
            ' Determine the current name based on whether there is a line break
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
                lines = Split(.Value, vbNewLine)
                cumulativeLength = Len(newName) + 2 ' Start with newName and its vbNewLine
                
                ' Apply strikethrough to all lines except the first one
                For i = 1 To UBound(lines)
                    startPos = cumulativeLength
                    .Characters(startPos, Len(lines(i)) + 1).Font.Strikethrough = True
                    cumulativeLength = cumulativeLength + Len(lines(i)) + 2 ' Update for next line
                Next i
                
                ' Explicitly increase row height by 15 points per swap
                .RowHeight = .RowHeight + 15
                
                ' Update personnel counter for the new staff
                lastRow = wsPersonnel.Cells(wsPersonnel.Rows.Count, "B").End(xlUp).row
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
End Sub



