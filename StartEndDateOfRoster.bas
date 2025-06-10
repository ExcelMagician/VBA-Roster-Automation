Attribute VB_Name = "StartEndDateOfRoster"
Sub Generate_StartDate_OfRoster()
Attribute Generate_StartDate_OfRoster.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Generate_StartDate_OfRoster Macro
'

'
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(DATEVALUE(RC[-3]&R[-1]C[2]&R[-1]C[5]),"""")"
    Range("H4").Select
End Sub
Sub Generate_EndDate_OfRoster()
Attribute Generate_EndDate_OfRoster.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Generate_EndDate_OfRoster Macro
'

'
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=EOMONTH(RC[-3],0)"
    Range("K4").Select
End Sub
