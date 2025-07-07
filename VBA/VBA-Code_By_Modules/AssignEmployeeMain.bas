Attribute VB_Name = "AssignEmployeeMain"
'declare worksheet and table
    Private wsRosterCopy As Worksheet
    Private wsPersonnel As Worksheet
    Private wsSettings As Worksheet
    Private morningtbl As ListObject
    Private spectbl As ListObject
    
'declare roster column number
    Private Const VAC_COL As Long = 1
    Private Const DATE_COL As Long = 2
    Private Const DAY_COL As Long = 3
    Private Const LMB_COL As Long = 4
    Private Const MOR_COL As Long = 6
    Private Const AFT_COL As Long = 8
    Private Const AOH_COL As Long = 10
    Private Const SAT_AOH_COL1 As Long = 12
    Private Const SAT_AOH_COL2 As Long = 14
    Private Const START_ROW As Long = 6
   
Sub Main()
    Set wsRosterCopy = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    
    Dim currDate As Date
    
    If wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" And wsRosterCopy.Cells(2, 13).Value Mod 4 = 0 Then
        lastRowRoster = 187
    ElseIf wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" Then
        lastRowRoster = 186
    Else
        lastRowRoster = 189
    End If
       
     'Loop through each date row
     For dateRow = 6 To lastRowRoster
     
        currDate = wsRosterCopy.Cells(dateRow, DATE_COL).Value
        
        If Weekday(currDate, vbMonday) = 7 Or _
            Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0 Then
            
            ' Skip this date by marking all slots as "CLOSED"
            wsRosterCopy.Cells(dateRow, LMB_COL).Value = "CLOSED" ' D column
            wsRosterCopy.Cells(dateRow, LMB_COL).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, MOR_COL).Value = "CLOSED" ' F column
            wsRosterCopy.Cells(dateRow, MOR_COL).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, AFT_COL).Value = "CLOSED" ' H column
            wsRosterCopy.Cells(dateRow, AFT_COL).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, AOH_COL).Value = "CLOSED" ' J column
            wsRosterCopy.Cells(dateRow, AOH_COL).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, SAT_AOH_COL1).Value = "CLOSED" ' L column
            wsRosterCopy.Cells(dateRow, SAT_AOH_COL1).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, SAT_AOH_COL2).Value = "CLOSED" ' N column
            wsRosterCopy.Cells(dateRow, SAT_AOH_COL2).Interior.Color = vbRed
            GoTo NextDate ' Skip to the next date
        End If
        
NextDate:
    Next dateRow
    
    Call AssignMorningDuties.AssignMorningDuties
End Sub

