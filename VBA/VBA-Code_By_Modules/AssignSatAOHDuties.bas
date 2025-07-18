Attribute VB_Name = "AssignSatAOHDuties"
' Declare worksheet and table
    Private wsRoster As Worksheet
    Private wsPersonnel As Worksheet
    Private wsSettings As Worksheet
    Private aohtbl As ListObject
    Private spectbl As ListObject

Sub AssignSatAOHDuties()
    Set wsRoster = Sheets("MasterCopy (2)")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Sat AOH PersonnelList")
    Set aohtbl = wsPersonnel.ListObjects("SatAOHMainList")
    
    Dim i As Long, r As Long
    Dim maxDuties As Long
    Dim staffName As String
    Dim assignedStaff1 As String

    ' Pass 1: Assign staff to SAT_AOH_COL1
    For r = START_ROW To LAST_ROW_ROSTER
        Dim dayValue As String
        dayValue = Trim(wsRoster.Cells(r, DAY_COL).Text)
        Debug.Print "Row " & r & " Day Value: '" & dayValue & "'"
        If dayValue = "Sat" Then
            Debug.Print "Processing row " & r & " (Saturday found)"
            If wsRoster.Cells(r, SAT_AOH_COL1).Value = "" Then
                For i = 1 To aohtbl.ListRows.Count
                    staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                    maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                    Dim currDuties As Long
                    currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
                    Debug.Print "  Checking staff " & staffName & " (Max Duties: " & maxDuties & ", Curr Duties: " & currDuties & ")"
                    If currDuties < maxDuties Then
                        wsRoster.Cells(r, SAT_AOH_COL1).Value = staffName
                        Call IncrementDutiesCounter(staffName)
                        assignedStaff1 = staffName
                        Debug.Print "Assigned All Days staff " & staffName & " to row " & r & " (SAT_AOH_COL1)"
                        Exit For
                    Else
                        Debug.Print "    Skipped: Max duties reached or weekly limit exceeded"
                    End If
                Next i
            End If
        End If
    Next r

    ' Pass 2: Assign different staff to SAT_AOH_COL2
    For r = START_ROW To LAST_ROW_ROSTER
        Dim dayValue2 As String
        dayValue2 = Trim(wsRoster.Cells(r, DAY_COL).Text)
        Debug.Print "Row " & r & " Day Value: '" & dayValue2 & "'"
        If dayValue2 = "Sat" And wsRoster.Cells(r, SAT_AOH_COL1).Value <> "" And wsRoster.Cells(r, SAT_AOH_COL2).Value = "" Then
            Debug.Print "Processing row " & r & " for SAT_AOH_COL2"
            For i = 1 To aohtbl.ListRows.Count
                staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                Dim currDuties2 As Long
                currDuties2 = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
                Debug.Print "  Checking staff " & staffName & " (Max Duties: " & maxDuties & ", Curr Duties: " & currDuties2 & ")"
                If currDuties2 < maxDuties And staffName <> wsRoster.Cells(r, SAT_AOH_COL1).Value Then
                    wsRoster.Cells(r, SAT_AOH_COL2).Value = staffName
                    Call IncrementDutiesCounter(staffName)
                    Debug.Print "Assigned All Days staff " & staffName & " to row " & r & " (SAT_AOH_COL2)"
                    Exit For
                Else
                    Debug.Print "    Skipped: Max duties reached, weekly limit exceeded, or same as SAT_AOH_COL1"
                End If
            Next i
        End If
    Next r

    MsgBox "Sat AOH duties assignment completed!", vbInformation
End Sub

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row

        ' Increment Duties Counter
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub


