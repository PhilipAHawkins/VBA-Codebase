Sub AddCheckBoxes()
    Dim i As Long
    Dim nRows As Long
    Dim cbxColumn As Long
    Dim tmpColumn As Long
    Dim objOLE As OLEObject
'
    cbxColumn = 10  ' (Column J)
    tmpColumn = 12 ' (Column L)
' Find the last row of data on the sheet.  CAUTION:  Assumes just one table of data on the sheet!
    nRows = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'
    With ActiveSheet.OLEObjects
        For i = .Count To 1 Step -1
            .Item(i).Delete
        Next i
    End With
    
    For i = 2 To nRows  ' Assumes row 1 is a header row
        ActiveSheet.Cells(i, cbxColumn).ColumnWidth = 4#
        Set objOLE = ActiveSheet.OLEObjects.Add(ClassType:="Forms.CheckBox.1")
        With objOLE
            .Top = ActiveSheet.Cells(i, cbxColumn).Top + 1
            .Width = 10.5
            .Height = ActiveSheet.Cells(i, cbxColumn).Height - 2
            .Left = ActiveSheet.Cells(i, cbxColumn).Left + _
                    0.5 * ActiveSheet.Cells(i, cbxColumn).Width - _
                    0.5 * .Width
            .Name = "cbx" & i
            .LinkedCell = ActiveSheet.Name & "!" & Cells(i, tmpColumn).Address
            .Object.BackStyle = 0
        .Object.Value = 1
            .Object.TripleState = False
            .Object.Caption = ""
        End With
        ActiveSheet.Cells(i, cbxColumn + 1).Formula = "=IF(ISERROR(" & _
                ActiveSheet.Cells(i, tmpColumn).Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
                "), ""N/A"", IF(" & _
                ActiveSheet.Cells(i, tmpColumn).Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
                "=TRUE,""YES"",""NO""))"
    Next i
End Sub

