VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpecificDays 
   Caption         =   "Select Specific Days"
   ClientHeight    =   3120
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   6010
   OleObjectBlob   =   "frmSpecificDays.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpecificDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public targetSheetName As String

Private Sub chkMon_Click()

End Sub

Private Sub chkSat_Click()

End Sub

Private Sub chkWed_Click()

End Sub

Private Sub commandCancel_Click()
    Dim currentDays As String
    currentDays = Trim(ThisWorkbook.Sheets(targetSheetName).Range("D8").Value)
    
    If currentDays = "" Then
        MsgBox "Please select at least one day before proceeding.", vbExclamation
    Else
        ' Keep existing selection and hide the form
        Me.Hide
    End If
End Sub
Private Sub commandOK_Click()
    Dim selectedDays As String
    Dim firstDay As Boolean
    
    firstDay = True
    If chkMon.Value Then
        selectedDays = "Mon"
        firstDay = False
    End If
    If chkTues.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Tue"
        firstDay = False
    End If
    If chkWed.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Wed"
        firstDay = False
    End If
    If chkThurs.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Thu"
        firstDay = False
    End If
    If chkFri.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Fri"
        firstDay = False
    End If
    If chkSat.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Sat"
        firstDay = False
    End If
    
    If selectedDays = "" Then
        MsgBox "Please select at least one day.", vbExclamation
    Else
        ThisWorkbook.Sheets(targetSheetName).Range("D8").Value = selectedDays
        Me.Hide
    End If
End Sub

Private Sub UserForm_Click()

End Sub

