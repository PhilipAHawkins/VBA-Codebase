' Takes a sheet with links to streaming music on youtube and randomly selects one, and then opens it
' name of channel is column a, link is column b. activated by button.
' "Private Declare PtrSafe Function ShellExecute _" for 64-bit systems,
' Private Declare Function ShellExecute _" for 32.

Option Explicit

Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
 

Public Sub OpenUrl()

    Dim lSuccess As Long
    Dim a As Long
    Dim b As Range
    Dim thisLink As String
    Dim thisNum As Long
    Set b = ThisWorkbook.Worksheets("Sheet1").Range("A:A")
    a = Application.WorksheetFunction.CountA(b)
    Debug.Print a
    With ThisWorkbook.Worksheets("Sheet1").Range("b2:b" & a) '<==change to your sheet with data and button
        .ClearFormats
        ' .Cells(Application.WorksheetFunction.RandBetween(1, 10)).Interior.Color = vbRed
        thisNum = Application.WorksheetFunction.RandBetween(1, a)
        thisLink = .Cells(thisNum).Value
        Debug.Print thisLink
    End With
    lSuccess = ShellExecute(0, "Open", thisLink)
    ThisWorkbook.Worksheets("Sheet1").Range("e8").Value = ThisWorkbook.Worksheets("Sheet1").Range("a" & thisNum)
    'Debug.Print thisLink
End Sub
