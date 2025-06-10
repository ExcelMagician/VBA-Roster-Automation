Attribute VB_Name = "ResetDutiesAOHCounter"
Sub ResetDutiesAOHCounter()
Attribute ResetDutiesAOHCounter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reset_Duties_AOH_Counter Macro
'

'
    Sheets("PersonnelList (AOH & Desk)").Select
    Range("E12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E12:F12").Select
    Selection.AutoFill Destination:=Range( _
        "Desk_PersonnelList[[Weekly Duties Counter]:[AOH Counter]]")
    Range("Desk_PersonnelList[[Weekly Duties Counter]:[AOH Counter]]").Select
    Sheets("MasterCopy").Select
End Sub
Sub ClearTableContent()
Attribute ClearTableContent.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearTableContent Macro
'

'
    Range("D6:O189").Select
    Selection.ClearContents
    Cells.Select
    Selection.Rows.AutoFit
End Sub
