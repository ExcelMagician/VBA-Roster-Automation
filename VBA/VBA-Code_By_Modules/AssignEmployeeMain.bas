Attribute VB_Name = "AssignEmployeeMain"
' Global variable
Public LAST_ROW_ROSTER As Long
Public wsRoster As Worksheet
Public wsSettings As Worksheet

' Declare roster column numbers
Public Const VAC_COL As Long = 1
Public Const DATE_COL As Long = 2
Public Const DAY_COL As Long = 3
Public Const LMB_COL As Long = 4
Public Const MOR_COL As Long = 6
Public Const AFT_COL As Long = 8
Public Const AOH_COL As Long = 10
Public Const SAT_AOH_COL1 As Long = 12
Public Const SAT_AOH_COL2 As Long = 14
Public Const START_ROW As Long = 6

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "M\n14"
    Set wsRoster = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    
    Dim dateRow As Long
    Dim currDate As Date
    Dim slotCol As Variant
    Dim slotCell As Range
    
    'Find last row of roster
    If wsRoster.Cells(2, 10).Value = "Jan-Jun" And wsRoster.Cells(2, 13).Value Mod 4 = 0 Then
        LAST_ROW_ROSTER = 187
    ElseIf wsRoster.Cells(2, 10).Value = "Jan-Jun" Then
        LAST_ROW_ROSTER = 186
    Else
        LAST_ROW_ROSTER = 189
    End If
       
    ' Loop through each date row
    For dateRow = START_ROW To LAST_ROW_ROSTER
        currDate = wsRoster.Cells(dateRow, DATE_COL).Value
        ' Reset formatting for all slots
        For Each slotCol In Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
            Set slotCell = wsRoster.Cells(dateRow, slotCol)
            slotCell.Interior.ColorIndex = xlNone ' Reset to no fill (default)
            slotCell.Font.Strikethrough = False
        Next slotCol
        
        'Check for Closed date
        If IsClosedDay(currDate) Then
            Call MarkAllSlotsClosed(dateRow)
        End If
        
    Next dateRow
    
    'Call ResetAllCounters.ResetAllCounters
    
    Call AssignSatAOHDuties.AssignSatAOHDuties
    Call AssignAOHDuties.AssignAOHDuties
    Call AssignAfternoonDuties.AssignAfternoonDuties
    Call AssignMorningDuties.AssignMorningDuties
    Call AssignLoanMailBoxDuties.AssignLoanMailBoxDuties
    
    Call DuplicateSystemRoster.DuplicateSystemRoster
End Sub

Function IsClosedDate(currDate As Date) As Boolean
    IsClosedDate = (Weekday(currDate, vbMonday) = 7) Or _
        Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0
End Function

Sub MarkAllSlotsClosed(dateRow As Long)
    Dim col As Variant
    For Each col In Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
        With wsRoster.Cells(dateRow, col)
            .Value = "CLOSED"
            .Interior.Color = vbRed
        End With
    Next col
End Sub

