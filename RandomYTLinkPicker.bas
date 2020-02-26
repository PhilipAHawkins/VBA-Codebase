Public Sub OpenUrl() 

    Dim lSuccess As Long
    Dim a As Long
    Dim b As Range, d As Range
    Dim thisLink As String
    Dim foundRow As Long
    Dim VisRng As Range, cll As Range
    Dim i As Integer
    Dim visRows() As Variant
    Dim fewest_hits As Long
    Dim FoundCell As Range
    Application.ScreenUpdating = False
   
    ' results weren't finding links that got selected less than others
    ' so i changed it to look first for results played the least and pic from them
    ' so everything will eventually get played an equal number of times.

  ' Columns in worksheet looks like:
  'A: Name of video
  'B: Link to page
  'C: Type of video
  'D: # times the video has been picked  

    Set b = ThisWorkbook.Worksheets("Sheet1").Range("A:A")
    a = Application.WorksheetFunction.CountA(b)
    Debug.Print a

    ' find whatever the min value is in Col D so that we only pick a row with that number. Only grab a link from the visible rows.
    Set d = ThisWorkbook.Worksheets("Sheet1").Range("D2:D" & a)
    fewest_hits = Application.WorksheetFunction.Min(d)

   ' sort on that column
   ThisWorkbook.Worksheets("Sheet1").Range("A1:d" & a).AutoFilter Field:=4, Criteria1:=CStr(fewest_hits)

   ' get the filtered rows' row numbers as an array
    Set VisRng = Range(Range("b2:b" & a), Range("b2:b" & a).End(xlDown)).SpecialCells(xlCellTypeVisible)
    i = 0
    For Each cll In VisRng
        ReDim Preserve visRows(i) ' lol this is what VBA does instead of .append
        visRows(i) = cll.Value
        i = i + 1
    Next

    ' pick one of the array members at random
    thisLink = visRows(Int(Rnd() * 3) + 1)
    Debug.Print thisLink
    Set FoundCell = Range("B:B").Find(What:=thisLink)
    foundRow = FoundCell.Row
    ThisWorkbook.Worksheets("Sheet1").Range("d" & foundRow).Value = ThisWorkbook.Worksheets("Sheet1").Range("d" & foundRow).Value + 1
    ThisWorkbook.Worksheets("Sheet1").Range("A1:d" & a).AutoFilter Field:=4
   
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Sheet1").Range("e8").Value = ThisWorkbook.Worksheets("Sheet1").Range("a" & foundRow)
    lSuccess = ShellExecute(0, "Open", thisLink)
   
End Sub
