Attribute VB_Name = "ClearTableContent"
Public Sub ClearTableContent()
'
' ClearTableContent Macro
'
    Range("D6:O189").Select
    Selection.ClearContents
    Selection.Rows.AutoFit
End Sub
