VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'Sorts duties counter by ascending
    Dim table As ListObject
    Set table = Me.ListObjects("MorningMainList")

    ' Check if the change happened in the Duties Counter column
    If Not Intersect(Target, table.ListColumns("Duties Counter").DataBodyRange) Is Nothing Then
        ' Sort by Duties Counter column (ascending)
        With table.Sort
            .SortFields.Clear
            .SortFields.Add key:=table.ListColumns("Duties Counter").DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
    
    
    'Sets Working Days for All Days and show form for Specific Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Mon, Tue, Wed, Thu, Fri, Sat"
                Debug.Print "Set D9 to 100% for All Days"
                
            Case "SPECIFIC DAYS"
                ' Show the multiselect form
                frmSpecificDays.Show
                Debug.Print "Showed form for Specific Days"
                
            Case Else
                ' Clear C2 if invalid selection
                Me.Range("D9").Value = ""
                Debug.Print "Invalid selection, D9 cleared"
        End Select
    End If
    
    Exit Sub
    
End Sub
