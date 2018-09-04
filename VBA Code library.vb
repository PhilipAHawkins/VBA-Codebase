
'Metrics

Option Explicit

Sub WeeklyMetrics()

Dim Scores As Worksheet, my_ws As Worksheet, wsGS As Worksheet, aClient As Worksheet, enr As Worksheet
Dim b As Long, r As Long
Dim s As Range, t As Range, u As Range, v As Range,  As Range
Dim errorhandle As String, notready As String, frm As String, msg As String
Dim PTcache As PivotCache
Dim PT As PivotTable
Dim ScorR As String
Dim fd As FileDialog
Dim myFile As String, fpath As String, wbGS As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.StatusBar = True


Set Scores = ActiveWorkbook.Sheets("Scores")
Set my_ws = ActiveWorkbook.Sheets("All")
ScorR = "\\path\file.xlsx"

'errors to handle
If Len(Dir(ScorR)) = 0 Then
    MsgBox "Error: Scoring Reference has moved on the network drive. Please edit the macro to the new location.", vbCritical, "Macro stopped"
    Exit Sub
End If

If Scores.Range("a1") = "" Then
    MsgBox "Looks like you're not ready for this yet. Did you add data from the Employee ID Scores query?", _
        vbInformation, "Macro stopped."
    Application.StatusBar = False
    Exit Sub
End If

Application.StatusBar = "Please select Enrollment Data"
'select enrollment data sheet and import
Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Please select the file that contains your Enrollment Data for this week."
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "You did not select a file. Please click again and select a file to run the macro.", vbInformation
            Application.StatusBar = False
            Exit Sub
        End If
        fpath = .SelectedItems(1)
        wbGS = Dir(.SelectedItems(1))
        
    End With
    
If MsgBox("You selected " & wbGS & ". Continue?", vbYesNo) = vbNo Then
    MsgBox "Please start over."
    Exit Sub
ElseIf vbYes Then
    'do nothing
End If

Application.StatusBar = "1/10 Importing Scoring Reference Data"

Workbooks.Open (ScorR)

'copy over info needed for formulas instead of referring to outside workbooks
ActiveWorkbook.Worksheets(Array("Cut Score Reference", "Client Alpha", "User Alpha", _
    "my_exam Pilot", "GC-II Pilot", "my_exam Ops. 3")).Copy after:=ThisWorkbook.Worksheets("Pie Slices")
Workbooks("Scoring Reference.xlsx").Close

Set aClient = Sheets("Client Alpha")
'Prep Scores tab, only fills in columns J:L for speed
Scores.Activate

With Scores
    .Range("j2") = "=HLOOKUP($B2,'Cut Score Reference'!$A:$XFC,29,FALSE)"
    .Range("k2") = "=IFERROR(IF($B2=1,VLOOKUP($C2,'my_exam Pilot'!$A:$H,3,FALSE),IF($B2=2,VLOOKUP($C2,'GC-II Pilot'!$A:$H,5,FALSE),IF($B2=3,VLOOKUP($C2,'my_exam Ops. 3'!$A:$H,3,FALSE),IF(HLOOKUP($B2,'Cut Score Reference'!$A:$XCZ,32,FALSE)=""In Pilot"",""Completed"",IF($E2<HLOOKUP($B2,'Cut Score Reference'!$A:$XCZ,32,FALSE),""Did Not Pass"",""Pass""))))),""Error: Check"")"
    .Range("l2") = "=IF(COUNTIFS($C:$C,$C2,$J:$J,$J2,$K:$K,""Pass"")>0,""Pass"",IF(COUNTIFS($C:$C,$C2,$J:$J,$J2,$K:$K,""Completed"")>0,""Completed"",""Did Not Pass""))"
End With

b = Scores.UsedRange.Rows.Count
Application.StatusBar = "2/10 Filling in the formulas in the Scores tab"
Range("J2:L" & b).FillDown

Application.StatusBar = "3/10 Pasting as values"
Range("J2:L" & b).Copy
Range("J2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

Application.StatusBar = "4/10 Building list of "
aClient.Range("$A:$AN").AutoFilter Field:=14, Criteria1:="<>* Analyst*", Operator:=xlFilterValues
aClient.UsedRange.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
aClient.Range("$A:$AN").AutoFilter Field:=14
Set  = aClient.Range("$A:$A")
.Copy Destination:=my_ws.Range("a1")

r = Application.WorksheetFunction.CountA()

ActiveWorkbook.Worksheets("Dashboard").Visible = xlHidden
ActiveWorkbook.Worksheets("Cut Score Reference").Visible = xlHidden
ActiveWorkbook.Worksheets("Client Alpha").Visible = xlHidden
ActiveWorkbook.Worksheets("User Alpha").Visible = xlHidden
ActiveWorkbook.Worksheets("Eligibility Reference").Visible = xlHidden

'Prep  tab
Application.StatusBar = "5/10 Preparing the All   tab for formulas"

With my_ws
    .Range("b2") = "=VLOOKUP($A2,'Client Alpha'!$A:$AN,2,FALSE)"
    .Range("d2") = "=IF(COUNTIFS(Scores!$C:$C,$A2,Scores!$J:$J,""my_exam"")=0,""Not Assessed"",IF(COUNTIFS(Scores!$C:$C,'All  '!$A2,Scores!$J:$J,""my_exam"",Scores!$K:$K,""Pass"")>0,""Pass"",""Did Not Pass""))"
    .Range("N2") = "=VLOOKUP($A2,'Client Alpha'!$A:$AN,13,FALSE)"
    .Range("O2") = "=VLOOKUP($A2,'Client Alpha'!$A:$AN,23,FALSE)"
    .Range("e2") = "=IF($O2<($P$1-365),""Eligible"",""Not Eligible"")"
    .Range("f2") = "=COUNTIFS(Scores!$C:$C,'All  '!$A2,Scores!$J:$J,""my_exam"")"
    .Range("H2") = "=IFERROR(VLOOKUP($N2,'Eligibility Reference'!$A:$E,3,FALSE),""Error: Check Work Role"")"
    .Range("I2") = "=IFERROR(VLOOKUP($N2,'Eligibility Reference'!$A:$F,6,FALSE),""Error: Check Work Role"")"
    .Range("J2") = "=IF($O2<VLOOKUP($N2,'Eligibility Reference'!$A:$E,5,FALSE),""Eligible"",""Not Eligible"")"
    .Range("K2") = "=COUNTIFS(Scores!$C:$C,'All  '!$A2,Scores!$J:$J,K2)"
    
End With

'Filldown Block
Application.StatusBar = "6/10 Filling in formulas in All   tab (long step)"
my_ws.Range("b2:o" & r).FillDown

Application.StatusBar = "7/10 Pasting as values"
' now 4 pastevalues blocks
Set v = my_ws.Range("b:o")
v.Copy
v.PasteSpecial xlPasteValues

Application.StatusBar = "8/10 Importing enrollment data"
Workbooks.Open (fpath)
ActiveWorkbook.ActiveSheet.Copy after:=ThisWorkbook.Worksheets("Pie Slices")
Workbooks(wbGS).Close
ActiveSheet.Name = "Enroll"
Set enr = Worksheets("Enroll")

Application.StatusBar = "9/10 Preparing the Enrollments for each exam"

'4. my_exam is ready now
my_ws.Range("G2") = "=IFERROR(IF(OR($D2=""Pass"",$D2=""Did Not Pass""),""Completed"",IF(COUNTIFS(Enroll!$A:$A,$A2,Enroll!$B:$B,""101-1021"")>0,""Enrolled"",""Not Assessed"")),""Not Assessed"")"
my_ws.Range("G2:G" & r).FillDown

'Column M ready now
Application.StatusBar = "Step 5 of x: Now counting each exam."
my_ws.Range("M2") = "=IFERROR(IF(OR($L2=""Pass"",$L2=""Did Not Pass""),""Completed"",IF(COUNTIFS(Enroll!$A:$A,$A2,Enroll!$B:$B,$I2)>0,""Enrolled"",""Not Assessed"")),""Not Assessed"")"
my_ws.Range("M2:P" & r).FillDown

Application.StatusBar = "10/10 PivotTables"

'PivotTables for future metrics as needed

    On Error GoTo 0
    
'   Create a Pivot Cache
    Set PTcache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=my_ws.Range("A1").CurrentRegion.Address)

'   Add the PL-II worksheet
    Worksheets.Add
    ActiveSheet.Name = "pvt  PL-II"
'   Create the Pivot Table from the Cache
    Set PT = ActiveSheet.PivotTables.Add( _
      PivotCache:=PTcache, _
      TableDestination:=Range("A1"), _
      TableName:="PL-II")
    
    With PT
'       Add fields
        .PivotFields("PL-II Assessment").Orientation = xlRowField
        .PivotFields("PL-II Enrollment").Orientation = xlRowField
        With .PivotFields("ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of ID"
            End With
        With .PivotFields("PL-II Eligibility")
            .Orientation = xlPageField
            .PivotItems("Not Eligible").Visible = False
            End With
       End With
    
'   Add the my_exam worksheet
    Worksheets.Add
    ActiveSheet.Name = "pvt my_exam Enrollment"
'   Create the Pivot Table from the Cache
    Set PT = ActiveSheet.PivotTables.Add( _
      PivotCache:=PTcache, _
      TableDestination:=Range("A1"), _
      TableName:="my_exam")
      
          With PT
'       Add fields
            .PivotFields("PL-I Enrollment").Orientation = xlRowField
            With .PivotFields("ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of ID"
            End With
            With .PivotFields("PL-I Eligibility")
            .Orientation = xlPageField
            .PivotItems("Not Eligible").Visible = False
            End With
            With .PivotFields("my_exam C.S.E.")
            .Orientation = xlPageField
            .PivotItems("Honorary Certification").Visible = False
            End With

          End With
Application.StatusBar = False
MsgBox "Complete!"

End Sub

'Customer


Option Explicit


Sub Demo()
'incorporated into RawData sub
     
    WelcomeMat.Show
    
    If WelcomeMat.MyString <> "" Then
        If MsgBox("You entered " & WelcomeMat.MyString & ". Continue?", vbYesNo) = vbYes Then
            'do nothing
        ElseIf vbNo Then
            MsgBox "Macro stopped. Please start over."
            Unload WelcomeMat
            Exit Sub
        End If
        Cells(1, 1) = WelcomeMat.MyDate
        Cells(1, 2) = "This concludes the macro demonstration."
    Else: MsgBox "You did not enter a date."
    End If
    Unload WelcomeMat
End Sub

Sub opener()
'incorporated into GStand sub
Dim wb As Workbook
Dim ScorR As String

ScorR = "<path>"

Set wb = Workbooks.Open(ScorR)

End Sub


Sub GetADate()
'incorporated into RawData sub
Dim TheString As String, TheDate As Date
TheString = Application.InputBox("Enter A Date")
If IsDate(TheString) Then
    TheDate = DateValue(TheString)
    MsgBox TheDate & vbNewLine & TheString
Else
    MsgBox "Invalid date"
End If

End Sub



Sub RawData()

Dim a As Long
Dim b As Range
Dim Denied As Range
Dim currentColumn As Integer
Dim columnHeading As String
Dim wsName As String
Dim wSheets As Long
Dim TheString As String
Dim TheDate As Date

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = True

Set b = Range("A:A")
a = Application.WorksheetFunction.CountA(b)


'error block
    If Range("a2") = 0 Then
        MsgBox ("Error 1: Please enter the Scoring data before continuing")
        Exit Sub
    End If
    If Range("a3") = 0 Then
        MsgBox ("Error 1: Please enter the Scoring data before continuing")
        Exit Sub
    End If
    If Sheets("Past Customers").Range("a1") = 0 Then
        MsgBox ("Error 2: Please copy the Past Customer data to the Past Customers tab.")
        Exit Sub
    End If


'user determines date to delete everything after //turned off at the moment
'WelcomeMat.Show
    
 '   If WelcomeMat.MyString <> "" Then
  '      If MsgBox("You entered " & WelcomeMat.MyString & ", which the macro will interpret as" & WelcomeMat.MyDate & " . Continue?", vbYesNo) = vbYes Then
   '         'do nothing
    '    ElseIf vbNo Then
     '       MsgBox "Macro stopped. Please start over."
      '      Unload WelcomeMat
       '     Exit Sub
        'End If
    'Else: MsgBox "You did not enter a date."
    'Exit Sub
    'End If

'begin filters
Application.StatusBar = "Step 1 of 14: Now filtering out Column K: Result Attempt. Records that Did Not Pass or Completed will be moved to Denials tab."

'delete all records after user-entered MyDate
'ActiveSheet.Range("$A$2:$BM$2").AutoFilter Field:=6, Criteria1:=">" & WelcomeMat.MyDate, _
 '   Operator:=xlFilterValues
  '  ActiveSheet.UsedRange.Offset(2).SpecialCells(xlCellTypeVisible).EntireRow.Delete
   ' ActiveSheet.Range("$A$2:$BM$2").AutoFilter Field:=6

'Removing "Completed" and "Did Not Pass" from Column K
    Range("K3:K" & a).FillDown
    ActiveSheet.Range("$A$2:$BM$2").AutoFilter Field:=11, Criteria1:=Array( _
        "Completed", "Did Not Pass"), Operator:=xlFilterValues
    Set Denied = ActiveSheet.UsedRange.Offset(2).SpecialCells(xlCellTypeVisible)
    Denied.Copy Destination:=Worksheets("Denials").Range("A2")
    Denied.EntireRow.Delete
    ActiveSheet.Range("A:AD").AutoFilter Field:=11
    Set b = Range("A:A")
    a = Application.WorksheetFunction.CountA(b)
    Set Denied = Nothing

    Application.StatusBar = "Step 2 of 14: Removing previous Customers. Records that are on the Past Customers list will be moved to the Denials tab."
'Removing "Filter Out" from Column H
    Range("H3") = "=IFERROR(VLOOKUP($A3,'Past Customers'!$A:$Q,12,FALSE),""Verify Qualifications"")"
    Range("H3:H" & a).FillDown
    ActiveSheet.Range("$A$2:$BM$2").AutoFilter Field:=8, Criteria1:=Array( _
        "Filter Out"), Operator:=xlFilterValues
    Set Denied = ActiveSheet.UsedRange.Offset(2).SpecialCells(xlCellTypeVisible)
    Denied.Copy Destination:=Worksheets("Denials").Range("a" & Worksheets("Denials").UsedRange.Rows.Count).End(xlUp).Offset(1)
    Denied.EntireRow.Delete
    ActiveSheet.Range("A:AD").AutoFilter Field:=8
    Set b = Range("A:A")
    a = Application.WorksheetFunction.CountA(b)
    Set Denied = Nothing
    
Application.StatusBar = "Step 3 of 14: Removing duplicates"
'Find AnswerSheetID's that are duplicates from a long time ago and remove.
  
Application.StatusBar = "Step 4 of 14: Checking formulas in columns I through W."
'Input formulas into cells since the formulas may have been deleted
If Range("i3") = "" Then 'the rest of the formula cells will be empty too, so we need to put them back in.
    'do code
  End If

'Filldowns can start now
Application.StatusBar = "Step 5 of 14: Filling in Columns I through W."

Set b = Range("A:A")
a = Application.WorksheetFunction.CountA(b)

Range("I3:J" & a).FillDown
Range("H3:K" & a).Copy
Range("H3").PasteSpecial (xlPasteValues)

Range("L3:R" & a).FillDown
Range("L3:R" & a).Copy
Range("L3").PasteSpecial (xlPasteValues)

Range("T3:t" & a).FillDown
Range("t3:t" & a).Copy
Range("t3").PasteSpecial (xlPasteValues)

Application.CutCopyMode = False
Range("a2").Select


Application.StatusBar = "Step 6 of 14: Preparing for Good Standing check."
'filter organizations, copy data to new sheets, export.
Sheets.Add.Name = "<name>"

Sheets("1. Customer List_RawData").Activate
Application.StatusBar = "Step 7 of 14: Creating Client 1 file for Good Standing Check."

'although this goes pretty quick as written, could probably be looped to go faster.

'Client 1
    ActiveSheet.Range("A:J").AutoFilter Field:=4, Criteria1:= _
        "<name>"
    ActiveSheet.UsedRange.Offset(1).Copy _
    Destination:=Worksheets("Client 1").Range("A1")
    ActiveSheet.Range("A:J").AutoFilter Field:=4
Application.StatusBar = "Step 8 of 14: Creating Client 2 file for Good Standing Check."

'data has been filtered to individual worksheets, now export, format, and save.
For wSheets = ActiveWorkbook.Worksheets.Count To 1 Step -1
    Sheets(wSheets).Activate
    wsName = ActiveSheet.Name

    Select Case wsName
    
        Case "<cases>"
        Application.StatusBar = "Step 13 of 14: Formatting files created for Good Standing Check. Now formatting " & wsName & " (" & wSheets & " remaining)."
        'If a worksheet name is one of our organizations, do the following:
            For currentColumn = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
            'first, check every column in the worksheet
                columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value
            
                'second, delete unneeded columns
                Select Case columnHeading
                        Case "<cases>"
                        'Do nothing
                        Case Else
                        'Delete column
                        ActiveSheet.Columns(currentColumn).Delete
                End Select
            Next
            'once columns are deleted, copy the data, paste into a new workbook, format it, save it, then close it and return.
            ActiveSheet.UsedRange.Copy
            Workbooks.Add
            Selection.PasteSpecial Paste:=xlValues
            Cells.EntireColumn.AutoFit
            ActiveWorkbook.ActiveSheet.Name = wsName
            ActiveWorkbook.Sheets("Sheet2").Delete
            ActiveWorkbook.Sheets("Sheet3").Delete
                ActiveWorkbook.SaveAs Filename:= _
                "<path>" & wsName & " Customer List Good Standing Check " _
                & Format(Date, "yyyymmdd"), FileFormat:=51
            ActiveWorkbook.Close
            ThisWorkbook.Activate
            
        Case Else
        'do nothing
    End Select
Next 'proceeds to go to the next worksheet and repeat until all worksheets have been finished.

Application.StatusBar = "Step 14 of 14: Finishing"
Application.DisplayAlerts = False


Application.DisplayAlerts = True
Application.StatusBar = False

'Unload WelcomeMat
MsgBox "Complete!"

End Sub
Public Sub GStand()

'an importer

Dim r As Range
Dim str As String
Dim valform As String
Dim fd As FileDialog
Dim wb As Workbook
Dim CL As Worksheet
Dim a As Long
Dim b As Range
Dim myPath As String
Dim myFile As String
Dim c As Long
Dim d As Boolean
Dim wbGS As String
Dim wsGS As Worksheet
Dim fpath As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.StatusBar = True
Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False
'Because Anthea hid some columns and it threw an error.

c = 1
d = False
str = "VALIDATED"
'a batch of variables
Set CL = Worksheets("1. Customer List_RawData")

Application.StatusBar = "Step 1: User selects Good Standing workbook"

MsgBox "In the box that is about to appear, please select the completed Good Standing workbook.", vbExclamation, "Macro ready to run"

If Range("v4").Text <> "" Then 'the macro has been run before
    If MsgBox("The macro sees that it has already run on this spreadsheet. Do you want to continue?", vbYesNo) = vbNo Then
        MsgBox "Macro stopped. Please click again when ready."
        Application.StatusBar = False
        Exit Sub
    ElseIf vbYes Then d = True
    End If
End If

'user picks folder with validated files
Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Please select the file that contains your COMPLETED Client Good Standing workbook."
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo nextcode
        fpath = .SelectedItems(1)
        wbGS = Dir(.SelectedItems(1))
        myPath = Left$(fpath, (Len(fpath)) - Len(wbGS))
    End With
    
nextcode:
    myPath = myPath
    If myPath = "" Then
    MsgBox "You did not select a file. Please click again and select a file to run the macro.", vbInformation
    Application.StatusBar = False
    Exit Sub
    End If

myFile = Dir(myPath)

'error handler in case user picks wrong folder
If MsgBox("You selected " & vbNewLine & vbNewLine & wbGS & vbNewLine & vbNewLine & " which is located in " _
            & vbNewLine & vbNewLine & myPath & _
            vbNewLine & vbNewLine & " The macro assumes that the Good Standing Workbook is in " & _
            "the same folder as the VALIDATED pre-requisite workbooks. Continue?" _
            , vbYesNo + vbInformation) = vbNo Then
    MsgBox "Please select the correct folder.", vbInformation, "Macro did not run."
    Application.StatusBar = False
    Exit Sub
    ElseIf vbYes Then
    
        Workbooks.Open (wbGS)
        Set wsGS = Workbooks(wbGS).Worksheets("Client")
        wsGS.Copy after:=ThisWorkbook.Worksheets("Pre-requisites")
        Workbooks(wbGS).Close
               
        Set b = CL.Range("A:A")
        a = Application.WorksheetFunction.CountA(b)
        
        CL.Range("S3") = "=IFERROR(VLOOKUP(A3,Client!A:H,8,FALSE),""---"")"
        CL.Range("s3:s" & a).FillDown
        CL.Range("s3:s" & a).Copy
        CL.Range("s3").PasteSpecial (xlPasteValues)
                            
Application.StatusBar = "Step 2: Preparing to import..."

'do loop opens file and copies relevant data to Pre-reqs tab in next free row
        Do While myFile <> ""
            Set wb = Workbooks.Open(Filename:=myPath & myFile)
            If InStr(wb.Name, str) Then
                c = c + 1
                Application.StatusBar = "Step 2: Now importing. " & c & " Validated workbooks completed so far."
                ActiveWorkbook.ActiveSheet.UsedRange.Offset(1).Copy _
                Destination:=ThisWorkbook.Sheets("Pre-requisites").Range("a" & Rows.Count).End(xlUp).Offset(1)
            Workbooks(myFile).Close
            Else
            Workbooks(myFile).Close
            End If
            myFile = Dir
        Loop
End If

Application.StatusBar = "Step 3: Preparing formulas for Column U (Training Pre-Req)."

'set the formula for Col U of Tab 1, filldown
    With CL
        If d = True Then 'clear data from previous run of macro
            .Range("U3:U" & a + 1000).Clear
        End If
        .Range("U3") = "=vlookup(a3,'Pre-requisites'!$A:$H,8,FALSE)"
        .Range("U3:U" & a).FillDown
        .Activate
    End With

Application.StatusBar = False
MsgBox "Copied " & c & " workbooks. Columns S and U filled in.", vbInformation, "Macro Complete"



End Sub


Sub GovCouncPhase()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.StatusBar = True

'for faster processing, message suppression, status updates

Dim kcount As Long, vcount As Long, i As Long
Dim Denied As Range, Accept As Range, v As Range
Dim raw As Worksheet, cou As Worksheet
Dim pvt As String
Dim PTcache As PivotCache
Dim PT As PivotTable
Dim objPic As Shape
Dim objChart As Chart
Dim vform As String
Dim couCSV As Range

Set raw = Worksheets("name")
Set cou = Worksheets("name")
pvt = "pvt G. Council"

Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False

Application.StatusBar = "Step 1: User verification"

If MsgBox("Has this list been verified to the instructions for preparing the Validation List?", vbYesNo) = vbNo Then
        MsgBox "Macro stopped. Please click again when ready."
        Application.StatusBar = False
        Exit Sub
    ElseIf vbYes Then
        If Range("v3") <> "" Then 'clear out previous iterations' data
            raw.Range("v3:v100000").Clear
            cou.Range("a5:e100000").Clear
        End If
End If


'set some vars
    vform = "=formula"
    vcount = WorksheetFunction.CountA(Range("a:a"))
    
Application.StatusBar = "Step 2: Preparing formulas for Column V"
    raw.Range("v3") = vform
    raw.Range("v3:v" & vcount).FillDown

If WorksheetFunction.CountIf(Range("v3:v" & vcount), "Missing data in cells") > 1 Then
    If MsgBox("The macro detects that some records have errors in them, which" & vbNewLine & _
            "means that validation checks are not complete." & vbNewLine & vbNewLine & _
            "Continue? You will not be able to undo.", vbYesNo + vbCritical, _
            "USER: Please confirm that the Validation check is not complete.") = vbNo Then 'the user needs to stop and work needs to be undone.
        raw.Range("v3:v" & vcount + 1000000).Clear
        MsgBox "Macro stopped. Any data in Column V has been cleared. Please click again when ready."
        Application.StatusBar = False
        Exit Sub
    ElseIf vbYes Then
    'do nothing
    End If
End If


Application.StatusBar = "Step 3: Removing all ineligible records to Denials tab"

    raw.Range("$A$2:$BM$2").AutoFilter Field:=22, Criteria1:=Array( _
        "Not Eligible", "Missing Data in cells"), Operator:=xlFilterValues
    Set Denied = raw.UsedRange.Offset(2).SpecialCells(xlCellTypeVisible)

                Denied.Copy Destination:=Worksheets("Denials").Range("a2").CurrentRegion.End(xlUp).Offset(1) 'WATCH
                Denied.EntireRow.Delete
                Set Denied = Nothing
                raw.Range("$A$2:$BM$2").AutoFilter Field:=22
                
Application.StatusBar = "Step 4: Generating validation list (longest step)"

    'Copy the AnswerSheetID's over
    Set v = raw.Range("a3:a" & raw.UsedRange.Rows.Count)
    kcount = v.Rows.Count + 4 'because the header is 4 rows at the top of cou
    v.Copy Destination:=cou.Range("e5")
    
    'Fill in names with formula, then sort alphabetically, then fill in other columns, then export to PDF
    
    With cou
        .Range("d5") = "=VLOOKUP($e5,'1. Customer List_RawData'!$A:$W,9,FALSE)" 'Name
        .Range("d5:d" & kcount).FillDown
        .Range("d5:e" & kcount).Sort key1:=cou.Range("d5"), order1:=xlAscending 'successfully sorted by alphabetical
        .Range("c5") = "=VLOOKUP($e5,'1. Customer List_RawData'!$A:$W,4,FALSE)"
        .Range("b5") = "=HLOOKUP(VLOOKUP($e5,'1. Customer List_RawData'!$A:$W,2,FALSE),'Reference'!$A:$XFC,30,FALSE)"
        .Range("b5:c" & kcount).FillDown
        .Range("b5:d" & kcount).EntireColumn.AutoFit
        .Range("a5") = 1
        .Range("a6") = "=sum(a5+1)"
        .Range("a6:a" & kcount).FillDown
        .Range("a5").Copy
        .Range("a6:a" & kcount).PasteSpecial xlPasteFormats
        .PageSetup.PrintArea = "A1:E" & kcount
        .PageSetup.LeftMargin = Application.InchesToPoints(1#)
        .ExportAsFixedFormat xlTypePDF, Filename:="Validation List " & Format(Now(), "yyyymmdd") & ".pdf", _
            quality:=xlQualityStandard, includedocproperties:=True, ignoreprintareas:=False, openafterpublish:=False
        .Cells.Copy
    End With
    
Application.StatusBar = "Step 5: Exporting Validation list to PDF and CSV for sharing"

Workbooks.Add
Selection.PasteSpecial xlValues
Range("F:L").EntireColumn.Delete

With ActiveWorkbook.ActiveSheet
    .Range("a1") = " Professional Certification - 2nd Quarter 2015 Customer List"
    .Range("a2") = "Ordered by Certification, Organization, then Name"
    .Range("a3") = "Certificate Numbers will be assigned upon Customer"
    .Range("b:e").EntireColumn.AutoFit
    .SaveAs Filename:="Validation " & Format(Now(), "yyyymmdd"), FileFormat:=xlCSV
End With

ActiveWorkbook.Close

Application.StatusBar = "Step 6: Generating Pivot Table for statistics"

    On Error GoTo 0
    
'   Create a Pivot Cache
    Set PTcache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=cou.Range("A5").CurrentRegion.Address)

'   Add the PL-II worksheet
    Worksheets.Add
    ActiveSheet.Name = pvt
'   Create the Pivot Table from the Cache
    Set PT = ActiveSheet.PivotTables.Add( _
      PivotCache:=PTcache, _
      TableDestination:=Range("A4"), _
      TableName:="pvt Gov. Council")
    
    With PT
'       Add fields
        With .PivotFields("Certification")
            .Orientation = xlRowField
            .Name = "Exams"
        End With
        
        With .PivotFields("Name")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Name"
        End With
  
    End With

Application.StatusBar = False
MsgBox "Macro is done." & vbNewLine & vbNewLine _
        & "The Validation list was created and exported as both PDF " & vbNewLine _
        & "and CSV to the same folder as this workbook. Find the two new files with Validation in the name." & vbNewLine & vbNewLine _
        & "A pivot table was also created to show the number of exams." & vbNewLine & vbNewLine _
        & "You will now be taken to the pivot table.", vbInformation, "Validation phase complete!"
        
Exit Sub

errhandle:
    MsgBox "Sorry, an error occurred. Please check that you do not have any columns hidden."
    
End Sub




Sub getrange()
Dim v As Range
Dim raw As Worksheet

Set raw = ActiveWorkbook.Sheets("1. Customer List_RawData")
Set v = raw.Range("a3:a" & raw.UsedRange.Rows.Count)
MsgBox v.Rows.Count
End Sub



Sub copier()

Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlValues

End Sub

Sub CCL()

Dim raw As Worksheet, CCL As Worksheet
Dim a As Long, cert1 As String, cert2 As String, adate1 As String, adate2 As String
Dim ddate As Date
Dim ScorR As String
Dim wsCodes As String, wsMil As String, wsClient As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.StatusBar = True


Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False
ScorR = "<path>"
wsCodes = "Exploded Office Codes Master"
wsClient = "Client Alpha"
wsMil = "User Alpha"
Set raw = Worksheets("1. Customer List_RawData")
Set CCL = Worksheets("3. Concatenated Customer List")

If raw.Range("a5") = "" Then
    MsgBox "Error: Please run the macro in " & raw.Name & " .", vbCritical
    Exit Sub
End If

start:
cert1 = InputBox("Step 1 of 6: Please enter the first new Certificate Number.") 'user will enter, then confirm.
cert2 = InputBox("Step 2 of 6: Please re-enter the new Certificate Number for verification.")

If cert1 <> cert2 Or Not IsNumeric(cert1) Then
    MsgBox "The numbers did not match or your number was invalid. Please try again.", vbQuestion
    Exit Sub
End If

If MsgBox("Step 3 of 6: You entered " & cert1 & ". Is this correct?", vbYesNo) = vbYes Then
        'do nothing
ElseIf vbNo Then
    If MsgBox("Ok, then the macro will not use that number. Do you want to quit?", vbYesNo) = vbYes Then 'halt the macro and clear data
        Exit Sub
    ElseIf vbNo Then 'they can try again
        GoTo start
    End If
End If

setDate:
adate1 = InputBox("Step 4 of 6: Please enter the Award Date.")
adate2 = InputBox("Step 5 of 6: Please re-enter the new Certificate Number for verification.")

If adate1 <> adate2 Or Not IsDate(adate1) Then
    MsgBox "The dates did not match or your date was invalid. Please try again.", vbQuestion
    Exit Sub
End If

If MsgBox("Step 6 of 6: You entered " & adate1 & ". Is this correct?", vbYesNo) = vbYes Then
    ddate = CDate(adate1)
ElseIf vbNo Then
    If MsgBox("Ok, then the macro will not use that date. Do you want to quit?", vbYesNo) = vbYes Then 'halt the macro and clear data
        Exit Sub
    ElseIf vbNo Then 'they can try again
        GoTo setDate
    End If
End If

Workbooks.Open (ScorR)

ActiveWorkbook.Worksheets(wsCodes).Copy after:=ThisWorkbook.Worksheets("Pre-requisites")
Workbooks("<name>").Activate
    
a = raw.UsedRange.Offset(1).Application.WorksheetFunction.CountA(raw.Range("a:a")) - 2

With raw
    .Range("A3:A" & a).Copy Destination:=CCL.Range("A2") 'name
    .Range("J3:J" & a).Copy Destination:=CCL.Range("B2") 'exam
    .Range("C3:C" & a).Copy Destination:=CCL.Range("C2") 'ID#
    .Range("D3:D" & a).Copy Destination:=CCL.Range("D2") 'Org
End With

CCL.Activate

With CCL
    .Range("e2") = "=VLOOKUP($A2,'1. Customer List_RawData'!$A$2:$I$1048576,9,FALSE)"
    .Range("E2:E" & a).FillDown
    .Range("a2:e" & a).Sort key1:=CCL.Range("e2"), order1:=xlAscending 'successfully sorted by alphabetical
    .Range("f2") = "=MID(E2&"" ""&E2,FIND("","",E2)+1,LEN(E2))"
    .Range("g2") = "=IFERROR(VLOOKUP($C2,'Client Alpha'!$A:$D,4,FALSE),""Not Client"")"
    .Range("h2") = "=IFERROR(VLOOKUP($C2,'Client Alpha'!$A:$E,5,FALSE),""Not Client"")"
    .Range("i2") = "=IFERROR(VLOOKUP($C2,'User Alpha'!$A:$K,9,FALSE),"""")"
    .Range("j2") = "=VLOOKUP(A2,'1. Customer List_RawData'!A:N,14,FALSE)"
    .Range("k2") = "=VLOOKUP(A2,'1. Customer List_RawData'!A:J,10,FALSE)"
    .Range("F2:K" & a).FillDown
    .Cells.Copy
    .Cells.PasteSpecial xlPasteValues
    .Range("L2") = cert1
    .Range("L2:L200000").NumberFormat = "general"
    .Range("L3").Formula = "=sum(L2+1)"
    .Range("L3:L" & a).FillDown
    .Range("L3:L" & a).Copy
    .Range("L3").PasteSpecial xlPasteValues
    .Range("m2:m" & a) = ddate 'error here
End With
CCL.Activate
Cells(1, 1).Select
Application.CutCopyMode = False

    MsgBox "Concatenated Customer List ready."

End Sub



'Template

Sub Step5a()
'a basic macro that just autofills the data needed in Step 5, taking it from Step 4
Application.ScreenUpdating = False

Dim Step5Row As Long
Dim Step5myrange As Range
On Error GoTo wrong

Sheets("Step 4 - importAccess").Activate
    If Range("a3") = "0" Then GoTo wrong
    If Range("a3") = "" Then GoTo wrong
    Set Step5myrange = Range("a:a")
    Step5Row = WorksheetFunction.Application.CountA(Step5myrange) + 1
    Sheets("Step 5 - MailMerge").Activate
    Range("a3", "BM" & Step5Row).Select
    Selection.FillDown
    On Error GoTo wrong
    
Exit Sub
    
wrong:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbExclamation

End Sub



Sub Step2()
Application.ScreenUpdating = False
Dim Step2Row As Long
Dim Step2myrange As Range

If Range("a3") = "" Then GoTo wrong
If Range("a3") = 0 Then GoTo wrong
On Error GoTo wrong

Sheets("Step1 - importREMARK").Activate
    Set Step2myrange = Range("c:c")
    Step2Row = WorksheetFunction.Application.CountA(Step2myrange) + 1
    Sheets("Step 2 - tblAnswerSheet").Activate
    Range("a3", "EC" & Step2Row).Select
    Selection.FillDown
Exit Sub


wrong:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbExclamation
    Exit Sub

End Sub

Sub Step3()
    
    Application.ScreenUpdating = False
    
    Dim Step2Row As Long
    Dim Step2myrange As Range
    
    Sheets("Step 2 - tblAnswerSheet").Activate
    Range("a2").Select
    Range(Selection, Selection.End(xlDown)).Select
        
    Step2Row = Selection.Count + 1
    
    If Range("a2") = 0 Then GoTo badentry
    
    Range("A2:A" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("A1").PasteSpecial xlPasteValues
    Range("B2:B" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("B1").PasteSpecial xlPasteValues
    Range("G2:G" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("C1").PasteSpecial xlPasteValues
    Range("K2:K" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("D1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Sheets("Step 3 - tblAnswers").Activate
    
    'begins phase of repeating data

Dim largearray As Range
Dim i As Integer
Dim formulamissing As String

Set r = Range("a:a")
b = WorksheetFunction.Application.CountA(r)
' the variable "b" is something we'll use to count how many rows we need
c = (b - 1) * 120 + 1
'some math we'll use at the end of the sub

'error handle block:
If b < 2 Then GoTo badentry
If Range("a2") = 0 Then GoTo badentry
If Range("f2").HasFormula = False Then GoTo formulamissing
If Range("f3").HasFormula = True Then GoTo once

Set largearray = Range("A2", "D" & b)
 
For i = 1 To 119
    largearray.Copy Destination:=Range("a" & Rows.Count).End(xlUp).Offset(1)
Next i

'sets up Column E
Range("e2", "e" & b).Value = 1
Range("e" & b + 1).Formula = "=$E2+1"
Range("e" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown

'sets up Column F
Range("f" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown
    
MsgBox "The macro determined that there were " & b & " rows of data. Therefore, the last row number should be " & c & "." _
    & vbNewLine & vbNewLine & "Please ensure that this is correct before continuing.", vbInformation, "Macro complete!"

Exit Sub

formulamissing:
    Dim neededformula As String
        msg5 = "Hi there, it looks like someone has tampered with this template and didn't inform the Test Ops team. There was supposed to be a formula in cell F2." _
        & vbNewLine & vbNewLine & "Would you be kind enough to let the Test Ops team know that this occurred? You won't be able to use the macro until it's fixed, but you can still do this tab by hand. Thanks."
    MsgBox msg5, vbInformation, "Oh no! Better call Phil!"
    Exit Sub
        
badentry:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbCritical
    Exit Sub
    
once:
    msg = "Looks like you've already run this."
    MsgBox msg, vbCritical
End Sub

Sub selector()

Range(Selection & Selection.End(xlDown)).Select
End Sub

Sub step6b()
'on click, this calls a macro in a word template that starts a mailmerge

'uses late binding to call Word

Dim wdApp As Object
Dim newDoc As Object
Dim strfile As String
Dim merger As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = True
Application.AskToUpdateLinks = False

merger = "<path>"
strfile = "<path>"

Application.StatusBar = "Filtering the non-Pilot Client records"

If Range("j4") <> "X" Then
    MsgBox "Looks like you're not ready for this yet.", vbInformation, "Please run the first macro before this one."
    Sheets("Step 6 - Send Emails").Activate
    Exit Sub
End If

If Range("j10") = "X" Then
    MsgBox "Looks like you've already run this macro. If you intend to run it again, please delete the X in cell J10", vbInformation, "Please do not click this button more than once."
    Sheets("Step 6 - Send Emails").Activate
    Exit Sub
End If

Sheets("Step 5 - MailMerge").Activate

If Range("a1") = "" Then
    MsgBox "Looks like you're not ready for this yet.", vbInformation, "Please run the first macro before this one."
    Sheets("Step 6 - Send Emails").Activate
    Exit Sub
End If
    

ActiveSheet.Range("$A:$BM").AutoFilter Field:=17, Criteria1:="<$>", Operator:=xlFilterValues
ActiveSheet.Range("$A:$BM").AutoFilter Field:=8, Criteria1:=Array("Pass", "Did Not Pass"), Operator:=xlFilterValues

    'clear data from target, copy rows over
    Workbooks.Open (merger)
    ActiveWorkbook.Sheets("Selection").Delete
    ActiveWorkbook.Sheets.Add.Name = "Selection"
    ThisWorkbook.Activate
    ActiveSheet.UsedRange.Cells.SpecialCells(xlCellTypeVisible).Copy
    Workbooks("<$>").Activate
    Range("A1").PasteSpecial xlPasteAll
    Workbooks("<$>").Save
    Workbooks("<$>").Close

Sheets("Step 6 - Send Emails").Activate

Application.StatusBar = "Beginning Hand-off to Word"

Set wdApp = CreateObject("Word.Application")
Set newDoc = wdApp.Documents.Add(strfile)
Call wdApp.Run("startmerger")
Application.StatusBar = False

With Range("J10")
    .Value = "X"
    .Font.Size = 72
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

'delete PDFs
Kill "<path>\*.*"

End Sub


Sub Step6a()

Application.ScreenUpdating = False
Dim r As Integer
Dim iMsg As Object
Dim iConf As Object
Dim strbody As String
Dim Flds As Variant

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
On Error GoTo wrong
Application.StatusBar = True

If Range("j4") = "X" Then
    MsgBox "Looks like you've already run this macro. If you intend to run it again, please delete the X in cell J4", vbInformation, "Please do not click this button more than once."
    Exit Sub
End If

Sheets("Step 5 - MailMerge").Activate
If Range("a4") = "" Then
    If MsgBox("There appears to be only one record in Step 4. Are you sure you want to continue?", vbYesNo) = vbNo Then
        Sheets("Step 6 - Send Emails").Activate
        GoTo wrong
    ElseIf vbYes Then
    'do nothing
    End If
End If

Rows(1).EntireRow.Delete
Application.StatusBar = "Removing Pilot exam records"
For r = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
    If Range("a3") = 0 Then GoTo wrong
    If Cells(r, 8) = "Pending: Test in Pilot" Then
        ActiveSheet.Rows(r).EntireRow.Delete
    End If
Next

Application.StatusBar = "Filtering User"

ActiveSheet.Range("$A:$BM").AutoFilter Field:=17, Criteria1:=Array( _
    "Array member 1", "Array member 2"), Operator:=xlFilterValues
    If Cells.SpecialCells(xlCellTypeVisible).Rows.Count = 1 Then
        Application.StatusBar = False
        MsgBox "There were no User records to send. Pilot exams have been filtered out, and the macro is done."
        Application.StatusBar = False
        Exit Sub
    End If
    Cells.SpecialCells(xlCellTypeVisible).Copy
    Workbooks.Add
    
Application.StatusBar = "Moving User records to new sheet"
Application.DisplayAlerts = False
    With ActiveWorkbook
        .ActiveSheet.Range("a1").PasteSpecial (xlPasteValuesAndNumberFormats)
        .ActiveSheet.Name = "Mil"
        .Sheets("Sheet2").Delete
        .Sheets("Sheet3").Delete
        .ActiveSheet.Cells.Select
Application.DisplayAlerts = True
    End With

Application.StatusBar = "Saving User results as new workbook"

Selection.EntireColumn.AutoFit
    ActiveWorkbook.SaveAs Filename:="<path>" _
        & Format(Date, "yyyymmdd") & "<$>", FileFormat:=51
    Cells(1, 1).Select
    ActiveWorkbook.Close
    
ActiveSheet.Range("$A:$BM").AutoFilter Field:=17

Sheets("Step 6 - Send Emails").Activate
    
Application.StatusBar = "Emailing workbook of User results to request that they be moved low side"

    iConf.Load -1   ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
            = "mailnce.titanium.rttitanium.nima.ic.gov"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "<address>"
        .Update
    End With
    
strbody = "Supplier," & vbNewLine & vbNewLine & _
        "Please find this week's User results attached to this email to move to the low side." & vbNewLine & vbNewLine & _
        "Anthea, this is what the email to the suppplier to send results to the low side would look like. --Phil"

With iMsg
    Set .configuration = iConf
    .To = "<address>"
    .CC = "<address>"
    .bcc = ""
    .from = "<address>"
    .Subject = "User Results for Low Side"
    .textbody = strbody
    .addattachment ("<path>" _
        & Format(Date, "yyyymmdd") & "Access.xlsx")
    .send
End With
Application.StatusBar = False
MsgBox "<$>"

With Range("J4")
    .Value = "X"
    .Font.Size = 72
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

Exit Sub
    
wrong:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbExclamation
    Application.StatusBar = False
    
End Sub


Sub step6c()
'on click, this calls a macro in a word template that starts a mailmerge

'uses late binding to call Word

Dim wdApp As Object
Dim newDoc As Object
Dim supers As String
Dim merger As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = True


If Range("j16") = "X" Then
    MsgBox "Looks like you've already run this macro. If you intend to run it again, please delete the X in cell J16", vbInformation, "Please do not click this button more than once."
    Exit Sub
End If

If Range("j10") <> "X" Then
    MsgBox "Looks like you're not ready for this yet.", vbInformation, "Please run the first macro before this one."
    Exit Sub
End If

supers = "<path>"

Application.StatusBar = "Beginning Hand-off to Word"

Set wdApp = CreateObject("Word.Application")
Set newDoc = wdApp.Documents.Add(supers)
Call wdApp.Run("supervisormacro")

'delete PDFs
Kill "<path>*.*"

'close word
killwd = "TASKKILL /F /IM WINWORD.EXE"
Shell killwd, vbHide

With Range("J16")
    .Value = "X"
    .Font.Size = 72
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

If MsgBox("All emails have been sent. Would you like to save and close now?", vbYesNo) = vbYes Then
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Exit Sub
ElseIf vbNo Then
    'do nothing
End If

Application.StatusBar = False


End Sub



'Candidate

Sub startmerger()

Dim Mname As String
Dim Fname As String
Dim Mexam As String
Dim iMsg As Object
Dim iConf As Object
Dim strbody As String
Dim Flds As Variant
Dim i As Long
'CHANGE MERGEFIELDS, FILEPATHS, EMAIL SERVERS, AND EMAIL ADDRESSES
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

ActiveDocument.SaveAs2 FileName:="<filepath>" _
    & Format(Date, "yyyymmdd") & "Macro Mail Merge.docm", Fileformat:=wdFormatXMLDocumentMacroEnabled

    ActiveDocument.MailMerge.OpenDataSource Name:="<filepath and filename>", _
        ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
        AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
        WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
        Format:=wdOpenFormatAuto, SQLStatement:="SELECT * FROM `Selection$`", SQLStatement1:="", SubType:= _
        wdMergeSubTypeAccess
    
'error handling
    
i = ActiveDocument.MailMerge.DataSource.RecordCount
If MsgBox("Before the macro begins, there appears to be " & i & " records. If this is correct, press OK to Continue.", vbOKCancel) = vbOK Then
    MsgBox "Now generating Candidate Feedback."
    ElseIf vbCancel Then
    MsgBox "Please ensure that the number of records is correct. If this error persists, contact the macro developer."
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    Exit Sub
End If
    
If Len(ActiveDocument.MailMerge.DataSource.DataFields("AssessmentDate")) = 5 Then
    MsgBox "Error: Please format Assessment Dates and Retest Dates as Dates. They are currently formatted to some other data type." _
        , vbCritical, "Macro found error"
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    Exit Sub
End If

    'open the first record
    With ActiveDocument.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            .LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            Fname = ActiveDocument.MailMerge.DataSource.DataFields("Name_Concat").Value
            Mexam = ActiveDocument.MailMerge.DataSource.DataFields("AssessmentShortHand").Value
            End With
        .Execute Pause:=False
    End With

'save as PDF
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "<filepath>" _
        & Fname & " - " & Mexam & " 101- Program Feedback", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, from:=1, to:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
        
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    
'email
    
    iConf.Load -1   ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
            = "mailnce.titanium.rttitanium.nima.ic.gov"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "<email>"
        .Update
    End With

    strbody = "Dear Test Ops:" & vbNewLine & vbNewLine & _
            "This is a test of the Skynet macro package. Please respond if you received this message with your exam feedback attached. You will receive a second email shortly, too."
    'strbody = "Please see the attached document for feedback on your recently completed 101- exam."

    With iMsg
        Set .configuration = iConf
        .to = ActiveDocument.MailMerge.DataSource.DataFields("CandidateEmailAddress").Value
        .CC = "<address>"
        .bcc = ""
        .from = "<address>"
        .Subject = "Sample feedback email using macro to Test Ops"
        .textbody = strbody
        .addattachment ("<filepath>" _
            & Fname & " - " & Mexam & " 101- Program Feedback.pdf")
        .send
    End With
    Set iMsg = Nothing
    'go to next record
    
If ActiveDocument.MailMerge.DataSource.RecordCount <> 1 Then 'multiple records
Do
'go to next record
            With ActiveDocument.MailMerge
            .DataSource.ActiveRecord = wdNextRecord
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
                With .DataSource
                    .FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
                    .LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
                    Fname = ActiveDocument.MailMerge.DataSource.DataFields("Name_Concat").Value
                    Mexam = ActiveDocument.MailMerge.DataSource.DataFields("AssessmentShortHand").Value
                End With
            .Execute Pause:=False
            End With

    'save as PDF
 ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "<filepath>" _
        & Fname & " - " & Mexam & " 101- Program Feedback", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, from:=1, to:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False        'close and return focus to mail mailmerge document
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    
'email
      strbody = "Dear Test Ops:" & vbNewLine & vbNewLine & _
            "This is a test of the Skynet macro package. Please respond if you received this message with your exam feedback attached. You will receive a second email shortly, too."
 '   strbody = "Please see the attached document for feedback on your recently completed 101- exam."
    Set iMsg = CreateObject("CDO.Message")
    With iMsg
        Set .configuration = iConf
        .to = ActiveDocument.MailMerge.DataSource.DataFields("CandidateEmailAddress").Value
        .CC = "<address>"
        .bcc = ""
        .from = "<address>"
        .Subject = "Test CDO SMTP Email"
        .textbody = strbody
        .addattachment ("<filepath>" _
            & Fname & " - " & Mexam & " 101- Program Feedback.pdf")
        .send
    End With
    Set iMsg = Nothing
    
Loop Until ActiveDocument.MailMerge.DataSource.ActiveRecord = ActiveDocument.MailMerge.DataSource.RecordCount

End If

MsgBox "<textbody>"

ActiveDocument.Close wdSaveChanges

End Sub




'Supervisor

Sub supervisormacro()

Dim Mname As String
Dim Fname As String
Dim Mexam As String
Dim iMsg As Object
Dim iConf As Object
Dim strbody As String
Dim Flds As Variant
Dim i As Long

'CHANGE ALL MERGEFIELDS --Phil

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

ActiveDocument.SaveAs2 FileName:="<file path here>" _
    & Format(Date, "yyyymmdd") & "Supervisor Mail Merge.docm", Fileformat:=wdFormatXMLDocumentMacroEnabled

ActiveDocument.MailMerge.OpenDataSource Name:="<file path here>", _
    ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
    AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
    WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
    Format:=wdOpenFormatAuto, SQLStatement:="SELECT * FROM `Selection$`", SQLStatement1:="", SubType:= _
    wdMergeSubTypeAccess
    
i = ActiveDocument.MailMerge.DataSource.RecordCount
If MsgBox("Before the macro begins, there appear to be " & i & " records. If this is correct, press OK to Continue.", vbOKCancel) = vbOK Then
    MsgBox "Now generating Supervisor Feedback."
    ElseIf vbCancel Then
    MsgBox "Please ensure that the number of records is correct. If this error persists, contact the macro developer."
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    Exit Sub
End If
    
If Len(ActiveDocument.MailMerge.DataSource.DataFields("AssessmentDate")) = 5 Then
    MsgBox "Error: Please format Assessment Dates and Retest Dates as Dates. They are currently formatted to some other data type." _
        , vbCritical, "Macro found error"
    Exit Sub
End If

'open the first record
    With ActiveDocument.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            .LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            Fname = ActiveDocument.MailMerge.DataSource.DataFields("Name_Concat").Value
            Mexam = ActiveDocument.MailMerge.DataSource.DataFields("AssessmentShortHand").Value
            End With
        .Execute Pause:=False
    End With

'save as PDF
     ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "<file path here>" _
        & Fname & " - " & Mexam & "<text here>", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, from:=1, to:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
        
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    
'email
    
    iConf.Load -1   ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
            = "mailnce.titanium.rttitanium.nima.ic.gov"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "<email address path here>"
        .Update
    End With

    strbody = "<email message here>"

    With iMsg
        Set .configuration = iConf
        .to = ActiveDocument.MailMerge.DataSource.DataFields("<mergefield here>").Value
        .CC = "<email here>"
        .bcc = ""
        .from = "<email here>"
        .Subject = "Test CDO SMTP Email"
        .textbody = strbody
        .addattachment ("<file path here>" _
            & Fname & " - " & Mexam & " <file here>.pdf")
        .send
    End With
    Set iMsg = Nothing
    'go to next record
    
If ActiveDocument.MailMerge.DataSource.RecordCount <> 1 Then 'multiple records
Do
'go to next record
            With ActiveDocument.MailMerge
            .DataSource.ActiveRecord = wdNextRecord
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
                With .DataSource
                    .FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
                    .LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
                    Fname = ActiveDocument.MailMerge.DataSource.DataFields("Name_Concat").Value
                    Mexam = ActiveDocument.MailMerge.DataSource.DataFields("AssessmentShortHand").Value
                End With
            .Execute Pause:=False
            End With

    'save as PDF
  ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "<file path here>" _
        & Fname & " - " & Mexam & " <name here>", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, from:=1, to:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False        'close and return focus to mail mailmerge document
    ActiveDocument.Close Savechanges:=wdDoNotSaveChanges
    
'email
    
    strbody = "Please see the attached document for feedback on your recently employee's completed 101- exam."
    Set iMsg = CreateObject("CDO.Message")
    With iMsg
        Set .configuration = iConf
        .to = ActiveDocument.MailMerge.DataSource.DataFields("<mergefield here>").Value
        .CC = "<email here>"
        .bcc = ""
        .from = "<email here>"
        .Subject = "Test CDO SMTP Email"
        .textbody = strbody
        .addattachment ("<file path here>" _
            & Fname & " - " & Mexam & " <file here>.pdf")
        .send
    End With
    Set iMsg = Nothing
    
Loop Until ActiveDocument.MailMerge.DataSource.ActiveRecord = ActiveDocument.MailMerge.DataSource.RecordCount

End If

MsgBox "Emails sent. Document containing results saved to" & vbNewLine & vbNewLine & _
        "<text here>", vbInformation, "Have a nice day!"

ActiveDocument.Close wdSaveChanges
End Sub

'Others
Sub fizzbuzz()
For i = 1 To 100
    Select Case True
        Case i Mod 15 = 0
            Cells(1, i) = "FizzBuzz"
        Case i Mod 3 = 0
            Cells(1, i) = "Fizz"
        Case i Mod 5 = 0
            Cells(1, i) = "Buzz"
        Case Else
            Cells(1, i) = i
    End Select
Next i
End Sub
'if Option Explicit isn't declared, this is solved in 14 lines.

Sub ChangeRangeToPicture()

Dim i  As Integer
Dim intCount As Integer
Dim objPic As Shape
Dim objChart As Chart

'copy the range as an image
Call Sheet1.Range("A4:b7").CopyPicture(xlPrinter, xlPicture)

'remove all previous shapes in sheet2
intCount = Sheet2.Shapes.Count
For i = 1 To intCount
    Sheet2.Shapes.Item(1).Delete
Next i

'create an empty chart in sheet2
Sheet2.Shapes.AddChart

'activate sheet2
Sheet2.Activate

'select the shape in sheet2
Sheet2.Shapes.Item(1).Select
Set objChart = ActiveChart

'paste the range into the chart
objChart.Paste

'save the chart as a JPEG
objChart.Export ("C:\StuffBusinessTempExample.Jpeg")

End Sub

'Add-ins

Sub disabler()
Dim olapp As Object
Set olapp = CreateObject("outlook.application")


Outlook.
Outlook.Application.COMAddIns("AACGOutlook.Connect").Connect = False

End Sub


Sub ListCOMAddins()

'successfully lists the add-ins installed in Outlook.
'uses late binding. remove Outlook pieces to get Excel information.

Dim olapp As Object
Set olapp = CreateObject("outlook.application")

Dim lngRow As Long, objCOMAddin As COMAddIn

lngRow = 1

With ActiveSheet
      For Each objCOMAddin In olapp.COMAddIns
         .Cells(lngRow, "A").Value = objCOMAddin.Description
         .Cells(lngRow, "B").Value = objCOMAddin.Connect
         .Cells(lngRow, "C").Value = objCOMAddin.progID
         lngRow = lngRow + 1
      Next objCOMAddin
End With
End Sub


'Stocks

1: Sub TrendStat()

'Sample Excel VBA program for processing trend statistics of the form presented in Tables 3.1–3.8.

2: ‘columns: A=Date, B=Time, C=Open, D=High, E=Low, F=Close, G=Volume 
3: ‘columns: H=Test#1, I=Test #2A,  J=Test #2B, K=Test#3A, L=Test#3B

4: On Error GoTo ProgramExit

5: Dim DataRow As Long
6: Dim SummaryRow As Long
7: Dim Param1 As Double
8: Dim Param2 As Double
9: Dim Start_1 As Double
10: Dim End_1 As Double
11: Dim Step_1 As Double

12: ‘summary table headings
13: Cells(1, “P”) = “Param1”
14: Cells(1, “Q”) = “Param2”
15: Cells(1, “R”) = Cells(1, “I”)
16: Cells(1, “S”) = Cells(1, “J”)
17: Cells(1, “T”) = Cells(1, “J”) & “ Avg”
18: Cells(1, “U”) = Cells(1, “K”)
19: Cells(1, “V”) = Cells(1, “L”)
20: Cells(1, “W”) = Cells(1, “L”) & “ Avg”

21: DataRow = 2
22: SummaryRow = 2

23: ‘set loop parameters here
24: Start_1 = 0
25: End_1 = 0.02
26: Step_1 = 0.001 
27: For Param1 = Start_1 To End_1 Step Step_1

28: ‘set param2 options here
29: ‘Param2 = Param1
30: Param2 = 0

31: Cells(2, “N”) = Param1
32: Cells(3, “N”) = Param2
33: Cells(SummaryRow, “P”) = Param1
34: Cells(SummaryRow, “Q”) = Param2

35: Cells(SummaryRow, “R”) = Cells(1, “Y”)
36: Cells(SummaryRow, “S”) = Cells(1, “Z”)
37: Cells(SummaryRow, “T”) = Cells(1, “AA”)
38: Cells(SummaryRow, “U”) = Cells(1, “AB”)
39: Cells(SummaryRow, “V”) = Cells(1, “AC”)
40: Cells(SummaryRow, “W”) = Cells(1, “AD”)

41: SummaryRow = SummaryRow + 1

42: Next Param1

43: ProgramExit:

44: End Sub



1: Sub spikes()
'Excel VBA price-spike-summary program.

2: Dim DataRow As Long
3: Dim SummaryRow As Long
4: Dim LastDataRow As Long
5: Dim ReferenceTicker As String
6: Dim Count As Integer
7: Dim Criterion As Double
8: Dim IterationIndex As Integer
9: Dim ScoreColumnName As String
10: Dim SummaryColumn_2 As Integer
11: Dim RangeString As String

12: IterationIndex = 0
13: For Criterion = 0.5 To 4 Step 0.5
14: DataRow = 2
15: Count = 0
16: ReferenceTicker = Cells(DataRow, “A”)

‘Create a sum for each ticker and store in column G
17: While Cells(DataRow, “A”) <> “”
18:  If Cells(DataRow, “A”) = ReferenceTicker Then
19:   If Abs(Cells(DataRow, “F”)) > Criterion Then
20:    Count = Count + 1
21:   End If
22:  Else
23:   ReferenceTicker = Cells(DataRow, “A”)
24:   DataRow = DataRow - 1
25:   Cells(DataRow, “G”) = Count
26:   Count = 0
27:  End If
28: DataRow = DataRow + 1
29: Wend
30: Cells(DataRow - 1, “G”) = Count

‘Create summary table-add data after each pass through the file
31: LastDataRow = DataRow - 1
32: SummaryRow = DataRow + 2
33: ScoreColumnName = “>” & Criterion & “ StdDev”
34: SummaryColumn_2 = IterationIndex + 2
35: Cells(SummaryRow, SummaryColumn_2) = ScoreColumnName
36: If IterationIndex = 0 Then
37:  Cells(SummaryRow, “A”) = “Ticker”
38: End If
39: SummaryRow = SummaryRow + 1
40: For DataRow = 2 To LastDataRow
41:  If Cells(DataRow, “G”) <> “” Then
42:   Cells(SummaryRow, “A”) = Cells(DataRow, “A”)
43:   Cells(SummaryRow, SummaryColumn_2) = Cells(DataRow, “G”)
44:   SummaryRow = SummaryRow + 1
45:  End If
46: Next DataRow

‘Begin next pass through the file
47: IterationIndex = IterationIndex + 1
48: Next Criterion

‘Format summary table
49: RangeString = “A” & (LastDataRow + 4) & “:” & “I” & (SummaryRow - 1)
50: Range(RangeString).Select
51: Selection.NumberFormat = “0”
52: End Sub



Sub PriceSpikes ()

'incomplete macro copied from text

1: Dim WindowLength As Integer
2: Dim StdDevRange As Range

3: WindowLength = 20

‘Column headings
4: Cells(1, “A”) = “Symbol”
5: Cells(1, “B”) = “Date”
6: Cells(1, “C”) = “Close”
7: Cells(1, “D”) = “Price Change”
8: Cells(1, “E”) = “Log Change”
9: Cells(1, “F”) = “Spike”

‘Calculate price change and log of price change
10: DataRow = 3
11: While Cells(DataRow, “A”) <> “”
12:  If Cells(DataRow, “A”) = Cells(DataRow - 1, “A”) Then
13:   Cells(DataRow, “D”) = Cells(DataRow, “C”) - Cells(DataRow - 1, “C”)
14:   Cells(DataRow, “E”) = Log(Cells(DataRow, “C”) / Cells(DataRow - 1, “C”))
15:  End If

16: DataRow = DataRow + 1
17: Wend

‘Calculate price spikes in standard deviations
18: DataRow = WindowLength + 3
19: While Cells(DataRow, “A”) <> “”
20:  If Cells(DataRow, “A”) = Cells(DataRow - WindowLength - 1, “A”) Then
21:   Set StdDevRange = Range(“E” & (DataRow - WindowLength) & “:” & “E” & (DataRow - 1))
22:   Cells(DataRow, “F”) = Cells(DataRow, “D”) / (Application.WorksheetFunction.StDev(StdDevRange) * Cells(DataRow - 1, “C”))
23:  End If
24:  DataRow = DataRow + 1
25:  Wend

end sub




1: Sub AlignMultipleTickers()

'Excel VBA record alignment program for multiple tickers.
2: Dim SymCol As Integer
3: Dim DateCol As Integer
4: Dim CloseCol As Integer
5: Dim Sym_Col As String
6: Dim Date_Col As String
7: Dim Close_Col As String
8: Dim Iterations As Integer
9: Dim Row As Long
10: Dim TestRow As Long
11: Dim Direction As Integer
12: Dim LeftCell As String
13: Dim RightCell As String
14: Dim RangeString As String
‘determine date direction
15: If Cells(2, “B”) > Cells(3, “B”) Then
16:  Direction = 1
17: Else: Direction = 2
18: End If
19: For Iterations = 1 To 2
20:  SymCol = 4
21:  DateCol = 5
22:  CloseCol = 6
‘exit when blank column is encountered
23: While Cells(2, SymCol) <> “”
24:  Sym_Col = ColumnLetter(SymCol)
25:  Date_Col = ColumnLetter(DateCol)
26:  Close_Col = ColumnLetter(CloseCol)
27:  Row = 2
28: While (Cells(Row, “B”) <> “”) And (Cells(Row, Date_Col) <> “”)
29: Select Case Direction
30: Case 1  ‘decreasing date order
31: If Cells(Row, “B”) > Cells(Row, Date_Col) Then
32:  LeftCell = Sym_Col & Row
33:  RightCell = Close_Col & Row
34:  RangeString = LeftCell & “:” & RightCell
35:  Range(RangeString).Select
36:  Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
37: Else
38:  If Cells(Row, Date_Col) > Cells(Row, “B”) Then
39:   LeftCell = “A” & Row
40:   RightCell = “C” & Row
41:   RangeString = LeftCell & “:” & RightCell
42:   Range(RangeString).Select
43:   Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
44:  End If
45: End If
46: Case 2     ‘increasing date order
47: If Cells(Row, “B”) < Cells(Row, Date_Col) Then
48:  LeftCell = Sym_Col & Row
49:  RightCell = Close_Col & Row
50:  RangeString = LeftCell & “:” & RightCell
51:  Range(RangeString).Select
52:  Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
53: Else
54: If Cells(Row, Date_Col) < Cells(Row, “B”) Then
55:  LeftCell = “A” & Row
56:  RightCell = “C” & Row
57:  RangeString = LeftCell & “:” & RightCell
58:  Range(RangeString).Select
59:  Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
60: End If
61: End If
62: End Select
63: Row = Row + 1
64: Wend
‘continue to last record then back up and delete blanks
65: While Cells(Row, “B”) <> “” Or Cells(Row, Date_Col) <> “”
66: Row = Row + 1
67: Wend
68: For TestRow = Row To 1 Step -1
69: If Cells(TestRow, “B”) = “” Or Cells(TestRow, Date_Col) = “” Then
70:  LeftCell = “A” & TestRow
71:  RightCell = “C” & TestRow
72:  RangeString = LeftCell & “:” & RightCell
73:  Range(RangeString).Select
74:  Selection.Delete Shift:=xlUp
75:  LeftCell = Sym_Col & TestRow
76:  RightCell = Close_Col & TestRow
77:  RangeString = LeftCell & “:” & RightCell
78:  Range(RangeString).Select
79:  Selection.Delete Shift:=xlUp
80: End If
81: Next TestRow
‘move to next ticker
82: SymCol = SymCol + 3
83: DateCol = DateCol + 3
84: CloseCol = CloseCol + 3
85: Wend
86: Next Iterations ‘2nd pass through all tickers
87: End Sub
88: Function ColumnLetter(ColumnNumber As Integer) As String
89:  If ColumnNumber > 26 Then
90:   ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
Chr(((ColumnNumber - 1) Mod 26) + 65)
91:  Else:  ColumnLetter = Chr(ColumnNumber + 64)
92:  End If
93:End Function



1: Sub AlignRecordsComplete()
'Excel VBA record alignment program with extensions for handling ascending and descending dates and deletion of unmatched records.

2: Dim Row As Long
3: Dim TestRow As Long
4: Dim Direction As Integer
5: Dim LeftCell As String
6: Dim RightCell As String
7: Dim RangeString As String
8: If Cells(2, “B”) > Cells(3, “B”) Then
9:  Direction = 1
10: Else: Direction = 2
11: End If
12: Row = 2
13: While (Cells(Row, “B”) <> “”) And (Cells(Row, “E”) <> “”)
14:  Select Case Direction
15:  Case 1
16:  If Cells(Row, “B”) > Cells(Row, “E”) Then
17:   LeftCell = “D” & Row
18:   RightCell = “F” & Row
19:   RangeString = LeftCell & “:” & RightCell
20:   Range(RangeString).Select
21:   Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
22:  Else
23:   If Cells(Row, “E”) > Cells(Row, “B”) Then
24:    LeftCell = “A” & Row
25:    RightCell = “C” & Row
26:    RangeString = LeftCell & “:” & RightCell
27:    Range(RangeString).Select
28:    Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
29:   End If
30:  End If
31:  Case 2
32:  If Cells(Row, “B”) < Cells(Row, “E”) Then
33:   LeftCell = “D” & Row
34:   RightCell = “F” & Row
35:   RangeString = LeftCell & “:” & RightCell
36:   Range(RangeString).Select
37:   Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
38:  Else
39:   If Cells(Row, “E”) < Cells(Row, “B”) Then
40:    LeftCell = “A” & Row
41:    RightCell = “C” & Row
42:    RangeString = LeftCell & “:” & RightCell
43:    Range(RangeString).Select
44:    Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
45:   End If
46:  End If
47:  End Select
48: Row = Row + 1
49: Wend
50: While Cells(Row, “B”) <> “” Or Cells(Row, “E”) <> “”
51:  Row = Row + 1
52: Wend
53: For TestRow = Row To 1 Step -1
54:  If Cells(TestRow, “B”) = “” Or Cells(TestRow, “E”) = “” Then
55:   LeftCell = “A” & TestRow
56:   RightCell = “F” & TestRow
57:   RangeString = LeftCell & “:” & RightCell
58:   Range(RangeString).Select
59:   Selection.Delete Shift:=xlUp
60:  End If
61: Next TestRow
62: End Sub



1: Sub RemoveSpaces()

'Excel VBA program for removing unmatched records after date alignment.
2: Dim Row As Long
3: Dim TestRow As Long
4: Dim LeftCell As String
5: Dim RightCell As String
6: Dim RangeString As String
7: Row = 1
8: While Cells(Row, “B”) <> “” Or Cells(Row, “E”) <> “”
9:  Row = Row + 1
10: Wend
11: For TestRow = Row To 1 Step -1
12:  If Cells(TestRow, “B”) = “” Or Cells(TestRow, “E”) = “” Then
13:   LeftCell = “A” & TestRow
14:   RightCell = “F” & TestRow
15:   RangeString = LeftCell & “:” & RightCell
16:   Range(RangeString).Select
17:   Selection.Delete Shift:=xlUp
18:  End If
19: Next TestRow
20: End Sub



1: Sub AlignRecords()
'Excel VBA program for aligning records by date.
2: Dim Row As Long
3: Dim LeftCell As String
4: Dim RightCell As String
5: Dim RangeString As String
6: Row = 2
7: While (Cells(Row, “B”) <> “”) And (Cells(Row, “E”) <> “”)
8: If Cells(Row, “B”) > Cells(Row, “E”) Then
9:  LeftCell = “D” & Row
10:  RightCell = “F” & Row
11:  RangeString = LeftCell & “:” & RightCell
12:  Range(RangeString).Select
13:  Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
14: Else
15:  If Cells(Row, “E”) > Cells(Row, “B”) Then
16:   LeftCell = “A” & Row
17:   RightCell = “C” & Row
18:   RangeString = LeftCell & “:” & RightCell
19:   Range(RangeString).Select
20:   Selection.Insert Shift:=xlDown, _
CopyOrigin:=xlFormatFromLeftOrAbove
21:  End If
22: End If
23: Row = Row + 1
24: Wend
25: End Sub


'Other

Sub Autofiller()
'a basic macro that just autofills the data needed in Step 5, taking it from Step 4
Application.ScreenUpdating = False
Dim Step5Row As Long
Dim Step5myrange As Range

ans = MsgBox("Is Scoring References.xlsx open?", vbYesNo + vbQuestion)
If ans = vbNo Then Exit Sub

Sheets("Step 4 - importAccess").Activate
    
    Set Step5myrange = Range("a:a")
    Step5Row = WorksheetFunction.Application.counta(Step5myrange)
    Sheets("Step 5 - MailMerge").Activate
    Range("a3", "BM" & Step5Row).Select
    Selection.FillDown
    ActiveCell.Offset(0, 1).Select

        
Exit Sub

'successfully autofills

End Sub

Sub opensref()

' Opens a workbook on a network drive.
sref = "\\My_Networkdrive_Name\path\path\my_file.xlsx"
Workbooks.Open Filename:=sref

End Sub

Sub step5copy()
Dim a As Range
Dim b As Long

Sheets("Step 5 - MailMerge").Activate

'basic counta counting process leads to copying a range

Set a = Sheets("Step 5 - MailMerge").Range("a:a")
b = WorksheetFunction.Application.counta(a)

Range("a3", "BM" & b).Copy
Range("a3").PasteSpecial xlPasteValues

End Sub

Sub Step5b()

Application.ScreenUpdating = False

Dim r As Integer
'uses usedrange to do capture what range to perform a loop in, in this case deleting rows that have content in a cell.
'utilizes a "1 to step -1" structure for the loop.

For r = ActiveSheet.UsedRange.Rows.count To 1 Step -1
   If Cells(r, 8) = "Pending: Test in Pilot" Then
        ActiveSheet.Rows(r).EntireRow.Delete
    End If
Next
Exit Sub
    
    End Sub


Sub Step2()
Application.ScreenUpdating = False
Dim Step2Row As Long
Dim Step2myrange As Range

If Range("a2") = "" Then GoTo wrong
If Range("a2") = 0 Then GoTo wrong
On Error GoTo wrong

'uses counta and performs a filldown

Sheets("Step1 - import").Activate
    Set Step2myrange = Range("c:c")
    Step2Row = WorksheetFunction.Application.counta(Step2myrange)
    Sheets("Step 2 - tblAnswerSheet").Activate
    Range("a2", "EC" & Step2Row).Select
    Selection.FillDown
Exit Sub


wrong:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbExclamation
    Exit Sub

End Sub



Sub Step3original()
   'begin the phase of copying from Step 2
    
    Application.ScreenUpdating = False
    'makes it so the macro executes in the background instead of making visible changes on screen
    
    Sheets("Step 2 - tblAnswerSheet").Select
    'selects the sheet with our data
    If Range("a2") = 0 Then GoTo badentry
    'making sure the user didn't skip Step 2
    Range("A:A,B:B,G:G,K:K").Copy
    'copies the 4 columns we want to move to Step 3
    Sheets("Step 3 - tblAnswers").Range("A1").PasteSpecial xlPasteValues
    'goes to Step 3 and pastes the information into cell A1. Note that this pastes as values
    Application.CutCopyMode = False
    'because I don't want to copy anything anymore
    Sheets("Step 3 - tblAnswers").Activate
    
    'begins phase of repeating data

Dim largearray As Range
Dim i As Integer
Dim formulamissing As String

Set r = Range("a:a")

b = WorksheetFunction.Application.counta(r)
' the variable "b" is something we'll use to count how many rows we need
c = (b - 1) * 120 + 1
'some math we'll use at the end of the sub

If b < 2 Then GoTo badentry
If Range("a2") = 0 Then GoTo badentry
If Range("f2").HasFormula = False Then GoTo formulamissing
If Range("f3").HasFormula = True Then GoTo once
'a lot of spaghetti code specifically just for errors and stuff

'Error handling. we want to ensure that the user ran step 2 before doing this, and we want to ensure that the template hasn't been tampered with.

Set largearray = Range("A2", "D" & b)
'this sets an array using a2 through d(whatever the last row the user input was)
 
For i = 1 To 119
'ok, this command starts our ForNext loop. We need this copied and pasted 120 times, and we already have it once, so we need it 119 more times.
'for that reason, we set the variable "i" as our loop's "counter". This code sets the loop to occur until i equals 119 (aka, until the loop has looped 119 times).

largearray.Copy Destination:=Range("a" & Rows.count).End(xlUp).Offset(1)
'This copies the array we've named, and choses a destination we've specified where the array will be pasted.
'our destination is the next blank row in Column A.
'This says to go to the end of the array (Range("a" & Rows.count).End(xlUp)), to the end of column A, and go down one more row (.Offset(1))
'this works since we know the next row down is empty.

Next i
'now that the command is issued to do what we need it to, this command says that everything in between "For i" and "Next i" is what gets looped

Range("e2", "e" & b).Value = 1
Range("e" & b + 1).Formula = "=$E2+1"
'since we need to have question numbers to put into Access, this sets up the first "1's" you need and the formula for the first "2's".
'it then uses a formula to add the remaining question numbers, which we will do now.

Range("e" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown
'we've already calculated what the last row number should be, so we can select the cell and everything above it and filldown.
'this fills in the remaining question numbers

'BEGIN OLD CODE
'Range("D1").Select
 '   Selection.End(xlDown).Select
  '  Selection.Offset(0, 1).Select
   ' Range(Selection, Selection.End(xlUp)).Select
    'Selection.FillDown
'since we had everything to the left of the Questions column filled in, all I had to do to fill in the question numbers
'was just move to the bottom of the sheet, go one cell to the right into the Questions column, select everything from that cell to the first cell with the formula,
'and Filldown. This filled in all formulas for me.
'END OLD CODE

Range("f" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown
'same thing, but for the formulas

'BEGIN OLD CODE
'Range("E2").Select
 '   Selection.End(xlDown).Select
  '  Selection.Offset(0, 1).Select
   ' Range(Selection, Selection.End(xlUp)).Select
    'Selection.FillDown
    'does essentially the same thing, but for the Column F answers. Since the whole column is nothing but headers, we just select the appropriate cell and filldown.
'END OLD CODE
    
MsgBox "The macro determined that there were " & b & " rows of data. Therefore, the last row number should now be " & c & "." _
    & vbNewLine & vbNewLine & "Please ensure that this is correct before continuing.", vbInformation, "Macro complete!"

Exit Sub

formulamissing:
    Dim neededformula As String
        msg5 = "Hi there, it looks like someone has tampered with this template and didn't inform the Test Ops team. There was supposed to be a formula in cell F2." _
        & vbNewLine & vbNewLine & "Would you be kind enough to let the Test Ops team know that this occurred? You won't be able to use the macro until it's fixed, but you can still do this tab by hand. Thanks."
    MsgBox msg5, vbInformation, "Oh no! Better call Phil!"
    Exit Sub
        
badentry:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbCritical
    'last bit of error handling needed to ensure that this doesn't croak whenever someone accidentally enters the wrong thing.
    Exit Sub
    
once:
    msg = "Looks like you've already run this."
    MsgBox msg, vbCritical
    'last bit of error handling needed to ensure that this doesn't croak whenever someone accidentally enters the wrong thing.
    Exit Sub
End Sub



Sub deleteIrrelevantColumns()
   Dim currentColumn As Integer
   Dim columnHeading As String
   
   'uses usedrange to perform a select case loop.

   For currentColumn = ActiveSheet.UsedRange.Columns.count To 1 Step -1
   
   'this is the range. "For the range that you're already using, counting from the bottom row's value up to 1,..."
   'also note that it's 1 to -1. Since we're deleting rows, it's unwise to start at the top and work down
   'this is because the row numbers below a deleted row get a new row number.
   'eg when Row 2 is deleted, Row 3 is now Row 2. If we used Step 1 instead of -1, row 3 wouldn't get deleted.
   
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value
        
    '"...look in the first column of each cell in the range and..."
        
        Select Case columnHeading
        '"...if it says..."
            Case "EmplID", "Organization", "ExamDate", "Name", "Assessemnt", "Result_Attempt", "#Attempts per test", "Result_Overall", "#total attempts"
                '"...Do nothing. But if it says..."
                'Note that this is the most likely case and is listed first.
                'also note that the instruction to "do nothing" had nothing given to perform, so VBA assumed that you wanted nothing done.
            Case Else
                '"...anything else, literally anything else besides what was in the above case, delete if the entire column."
                ActiveSheet.Columns(currentColumn).Delete

        End Select
    Next

End Sub


Sub splitservices()
'
' cuts and pastes data for me to correct sheet
' at last test, worked instantly

Application.ScreenUpdating = False

Sheets.Add.Name = "Client 1"
Sheets.Add.Name = "Client 2"
Sheets.Add.Name = "Client 3"
Sheets.Add.Name = "Client 4"

Sheets("Scores").Activate

'Client 1
    ActiveSheet.Range("A:J").AutoFilter Field:=3, Criteria1:= _
        "US Client 1"
    ActiveSheet.UsedRange.Copy _
    Destination:=Worksheets("Client 1").Range("A1")
    ActiveSheet.Range("A:J").AutoFilter Field:=3
'Client 2
    ActiveSheet.Range("A:J").AutoFilter Field:=3, Criteria1:= _
        "US Client 2"
    ActiveSheet.UsedRange.Copy _
    Destination:=Worksheets("Client 2").Range("A1")
    ActiveSheet.Range("A:J").AutoFilter Field:=3
'Client 3
    ActiveSheet.Range("A:J").AutoFilter Field:=3, Criteria1:= _
        "US Client 3"
    ActiveSheet.UsedRange.Copy _
    Destination:=Worksheets("Client 3").Range("A1")
    ActiveSheet.Range("A:J").AutoFilter Field:=3
'Client 4
    ActiveSheet.Range("A:J").AutoFilter Field:=3, Criteria1:= _
        "US Client 4"
    ActiveSheet.UsedRange.Copy _
    Destination:=Worksheets("Client 4").Range("A1")
    ActiveSheet.Range("A:J").AutoFilter Field:=3
'done, now format
    For Each Worksheet In Worksheets
        Worksheet.UsedRange.WrapText = False
        Worksheet.UsedRange.entirecolumn.AutoFit
        Next Worksheet
    Sheets("Scores").Activate
       
End Sub


Sub makesnewfiles()
'
' copier Macro
'
'ended up not using this one, it kept thinking that the workbook it went to wasn't where it started, so it returned a Subscript Out of Range error.

'Client 4
    Sheets("Client 4").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlValues
    ActiveWorkbook.ActiveSheet.Name = "Client 4"
    ActiveWorkbook.Sheets("Sheet2").Delete
    ActiveWorkbook.Sheets("Sheet3").Delete
    ActiveWorkbook.ActiveSheet.Cells.Select
    Selection.entirecolumn.AutoFit
        ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & ActiveSheet.Range("c2") & "_ Professional Certification_All Testing" _
        & Format(Date, "yyyymmdd"), FileFormat:=51
    Workbooks("User Monthly Report.xlsm").Activate
   
'Client 3
    Sheets("Client 3").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & Range("c2").Value & "_ Professional Certification_All Testing" _
        & Format(Date, "yyyymmdd"), FileFormat:=51
   
    
'Client 2
    Sheets("Client 2").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & Range("c2").Value & "_ Professional Certification_All Testing" _
        & Format(Date, "yyyymmdd"), FileFormat:=51
   
    
'Client 1
    Sheets("Client 1").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & Range("c2").Value & "_ Professional Certification_All Testing" _
        & Format(Date, "yyyymmdd"), FileFormat:=51
   
End Sub

Sub wsprocess()
'this is a sub that calls other subs.
'I kept getting error 400 when performing each step as one macro, but by splitting
'them up I was able to find my mistakes easier. That's why I did it this way
'I know it jumps around, but honestly it works really quick already and this isn't that hard.
splitservices
Client 2
Client 3
AirForce
Client 4
ActiveWorkbook.Sheets("Client 1").Delete
'this is the command to delete specific worksheets.
'I believe that I could just say "Sheets(1)" "Sheets(2)" if that was the position the sheets were in
ActiveWorkbook.Sheets("Client 2").Delete
ActiveWorkbook.Sheets("Client 3").Delete
ActiveWorkbook.Sheets("Client 4").Delete
Sheets("Scores").Rows("2:" & Rows.count).ClearContents
MsgBox "Complete!"

End Sub

Sub Client 4()
'
' copier Macro
'
Application.ScreenUpdating = False

'Client 4
    Sheets("Client 4").Select
    Cells.Select
    'this selects all cells in a sheet
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteAll
    ' xlpasteall pastes both values and formulas.
    'the line above it, workbooks.add, creates a new workbook and makes it the active window.
    'ergo, these commands copy data on a worksheet and paste it into a new workbook.
    ActiveWorkbook.ActiveSheet.Name = "Client 4"
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Sheets("Sheet2").Delete
    ActiveWorkbook.Sheets("Sheet3").Delete
    ActiveWorkbook.ActiveSheet.Cells.Select
    Selection.entirecolumn.AutoFit
    'this line and the one above it select all cells and autofit the columns
        ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & "Client 4  Professional Certification All Testing " _
        & Format(Date, "yyyymmdd"), FileFormat:=51
        'saves the new workbook using the naming convention we're giving it here.
        'Notes on this procedure:
        'be sure to have that last \ after "monthly report" or it'll think you want to name the file "Monthly Report"
        'if we only had "\" as the filename, it would save the workbook in the same folder as the template (which is what it does now, I just want i specified)
        'The I drive is a network drive; that poses no issue.
        'using the & operator we can append more text to the filename
        'including today's date in YYMMDD format
        'fileformat:=51 is the xlsx type, so you'll usually use that unless you want to keep macros in the workbook
    Workbooks("User Monthly Report.xlsm").Activate

End Sub

Sub Client 3()
'
' copier Macro
'

Application.ScreenUpdating = False
'Client 3
    Sheets("Client 3").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteAll
    ActiveWorkbook.ActiveSheet.Name = "Client 3"
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Sheets("Sheet2").Delete
    ActiveWorkbook.Sheets("Sheet3").Delete
    ActiveWorkbook.ActiveSheet.Cells.Select
    Selection.entirecolumn.AutoFit
        ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & "Client 3  Professional Certification All Testing " _
        & Format(Date, "yyyymmdd"), FileFormat:=51
    Workbooks("User Monthly Report.xlsm").Activate

End Sub

Sub Client 2()
'
' copier Macro
'


'Client 2
    Sheets("Client 2").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteAll
    ActiveWorkbook.ActiveSheet.Name = "Client 2"
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Sheets("Sheet2").Delete
    ActiveWorkbook.Sheets("Sheet3").Delete
    ActiveWorkbook.ActiveSheet.Cells.Select
    Selection.entirecolumn.AutoFit
    ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" _
        & "Client 2  Professional Certification All Testing " _
        & Format(Date, "yyyymmdd"), FileFormat:=51
    Workbooks("User Monthly Report.xlsm").Activate

End Sub

Sub AirForce()
'
' copier Macro
'

Application.ScreenUpdating = False
'Client 1
    Sheets("Client 4").Select
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteAll
    ActiveWorkbook.ActiveSheet.Name = "Client 1"
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Sheets("Sheet2").Delete
    ActiveWorkbook.Sheets("Sheet3").Delete
    ActiveWorkbook.ActiveSheet.Cells.Select
    Selection.entirecolumn.AutoFit
        ActiveWorkbook.SaveAs Filename:= _
        "C:\101-\User Results Notification\Jobs\Monthly Report\" & "Client 1  Professional Certification All Testing " _
        & Format(Date, "yyyymmdd"), FileFormat:=51
    Workbooks("User Monthly Report.xlsm").Activate

End Sub




Sub selector()
'selects a range from whichever cell is selected down to the last cell in a column with data
Range(Selection & Selection.End(xlDown)).Select
End Sub


Sub Step3()

'this is how it ended up once they provided feedback for me.
'the team asked if buttons in various places could go in the top row
'this ended up making me rewrite a bunch of code, and below is the result.
'I'm fully aware that I'll find better ways to do this some day, this is just to preserve it for posterity
    
    Application.ScreenUpdating = False
    Dim Step2Row As Long
    Dim Step2myrange As Range
    
    Sheets("Step 2 - tblAnswerSheet").Activate
    Range("a2").Select
    'going forward I'd like to find ways to avoid using select
    Range(Selection, Selection.End(xlDown)).Select
        
    Step2Row = Selection.count + 1
    'had to add the +1 because of the top row containing buttons.
            
    If Range("a2") = 0 Then GoTo badentry
    Range("A2:A" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("A1").PasteSpecial xlPasteValues
    'probably could've just used Destination instead of a whole new line
    Range("B2:B" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("B1").PasteSpecial xlPasteValues
    Range("G2:G" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("C1").PasteSpecial xlPasteValues
    Range("K2:K" & Step2Row).Copy
    Sheets("Step 3 - tblAnswers").Range("D1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    'because I don't want to copy anything anymore
    Sheets("Step 3 - tblAnswers").Activate
    
    'begins phase of repeating data

Dim largearray As Range
Dim i As Integer
Dim formulamissing As String

Set r = Range("a:a")
b = WorksheetFunction.Application.counta(r)
' the variable "b" is something we'll use to count how many rows we need
c = (b - 1) * 120 + 1
'some math we'll use at the end of the sub

If b < 2 Then GoTo badentry
If Range("a2") = 0 Then GoTo badentry
If Range("f2").HasFormula = False Then GoTo formulamissing
If Range("f3").HasFormula = True Then GoTo once

Set largearray = Range("A2", "D" & b)
'this sets an array using a2 through d(whatever the last row the user input was)
 
For i = 1 To 119
'ok, this command starts our ForNext loop. We need this copied and pasted 120 times, and we already have it once, so we need it 119 more times.
'for that reason, we set the variable "i" as our loop's "counter". This code sets the loop to occur until i equals 119 (aka, until the loop has looped 119 times).

largearray.Copy Destination:=Range("a" & Rows.count).End(xlUp).Offset(1)
'This copies the array we've named, and choses a destination we've specified where the array will be pasted.
'our destination is the next blank row in Column A.
'This says to go to the end of the array (Range("a" & Rows.count).End(xlUp)), to the end of column A, and go down one more row (.Offset(1))
'this works since we know the next row down is empty.

Next i
'now that the command is issued to do what we need it to, this command says that everything in between "For i" and "Next i" is what gets looped

Range("e2", "e" & b).Value = 1
Range("e" & b + 1).Formula = "=$E2+1"
'since we need to have question numbers to put into Access, this sets up the first "1's" you need and the formula for the first "2's".
'it then uses a formula to add the remaining question numbers, which we will do now.

Range("e" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown
'we've already calculated what the last row number should be, so we can select the cell and everything above it and filldown.
'this fills in the remaining question numbers

'BEGIN OLD CODE
'Range("D1").Select
 '   Selection.End(xlDown).Select
  '  Selection.Offset(0, 1).Select
   ' Range(Selection, Selection.End(xlUp)).Select
    'Selection.FillDown
'since we had everything to the left of the Questions column filled in, all I had to do to fill in the question numbers
'was just move to the bottom of the sheet, go one cell to the right into the Questions column, select everything from that cell to the first cell with the formula,
'and Filldown. This filled in all formulas for me.
'END OLD CODE

Range("f" & c).Select
   Range(Selection, Selection.End(xlUp)).Select
   Selection.FillDown
'same thing, but for the formulas

'BEGIN OLD CODE
'Range("E2").Select
 '   Selection.End(xlDown).Select
  '  Selection.Offset(0, 1).Select
   ' Range(Selection, Selection.End(xlUp)).Select
    'Selection.FillDown
    'does essentially the same thing, but for the Column F answers. Since the whole column is nothing but headers, we just select the appropriate cell and filldown.
'END OLD CODE
    
MsgBox "The macro determined that there were " & b & " rows of data. Therefore, the last row number should be " & c & "." _
    & vbNewLine & vbNewLine & "Please ensure that this is correct before continuing.", vbInformation, "Macro complete!"

Exit Sub

formulamissing:
    Dim neededformula As String
        msg5 = "Hi there, it looks like someone has tampered with this template and didn't inform the Test Ops team. There was supposed to be a formula in cell F2." _
        & vbNewLine & vbNewLine & "Would you be kind enough to let the Test Ops team know that this occurred? You won't be able to use the macro until it's fixed, but you can still do this tab by hand. Thanks."
    MsgBox msg5, vbInformation, "Oh no! Better call Phil!"
    Exit Sub
        
badentry:
    msg = "Looks like you're not ready for this yet."
    MsgBox msg, vbCritical
    'last bit of error handling needed to ensure that this doesn't croak whenever someone accidentally enters the wrong thing.
    Exit Sub
    
once:
    msg = "Looks like you've already run this."
    MsgBox msg, vbCritical
    'last bit of error handling needed to ensure that this doesn't croak whenever someone accidentally enters the wrong thing.
    Exit Sub
End Sub

Sub copytodiffworksheet2()
'developer notes:
'OBSOLETE. SHORTER METHOD FOUND. SEE MACRO ENTIRECOLUMN
'Hi everyone! Please enjoy this macro -Phil Hawkins
'this sub copies the selected data to another worksheet

Application.ScreenUpdating = False

Dim rng1 As Range
Dim rng2 As Range
Dim rng3 As Range
Dim rng4 As Range
'the above are declarations needed to complete this process.

Sheets("Step 2 - tblAnswerSheet").Activate
'this ensures that the correct worksheet that has our data is selected so that Excel knows which sheet to copy from

Range("a1").Select

Set rng1 = Range(ActiveCell, ActiveCell.End(xlDown))
'this sets a range for Excel to look in. This Macro is set to support up to 30,000 rows of data
'starting with the selected cell, this does the same as "control + down" on a keyboard, then names the selection as "rng1"
    rng1.Copy
'ready to copy
    Sheets("Step 3 - tblAnswers").Activate
'selects the right worksheet
    Range("a1").Select
'selects the cell we want to paste into
    Selection.PasteSpecial xlPasteValues
'pastes the data we copied.
'NOTE: this uses xlPasteValues, which will only paste the data, not any formula that is included in the cell
    Sheets("Step 2 - tblAnswerSheet").Activate
'returns to the worksheet with our data in order to do this again for the other columns of data
    
Range("b1").Select
    Set rng2 = Range(ActiveCell, ActiveCell.End(xlDown))
    rng2.Copy
    Sheets("Step 3 - tblAnswers").Activate
    Range("b1").Select
    Selection.PasteSpecial xlPasteValues
    Sheets("Step 2 - tblAnswerSheet").Activate
Range("g1").Select
    Set rng3 = Range(ActiveCell, ActiveCell.End(xlDown))
    rng3.Copy
    Sheets("Step 3 - tblAnswers").Activate
    Range("c1").Select
    Selection.PasteSpecial xlPasteValues
    Sheets("Step 2 - tblAnswerSheet").Activate
Range("k1").Select
    Set rng4 = Range(ActiveCell, ActiveCell.End(xlDown))
    rng4.Copy
    Sheets("Step 3 - tblAnswers").Activate
    Range("d1").Select
    Selection.PasteSpecial xlPasteValues
    
Application.CutCopyMode = False
'This stops copying the last selection.
  
End Sub


Sub entirecolumn()
'Hi everyone, hope you enjoy this macro -Phil Hawkins
'This copies the data in Step 2 that we need into Step 3.
'Pretty simple, utilizes the .entirecolumn command to select the data. It's pasted in as Values.

Application.ScreenUpdating = False

'1
Sheets("Step 2 - tblAnswerSheet").Activate
Range("a1").entirecolumn.Copy
Sheets("Step 3 - tblAnswers").Activate
Range("a1").PasteSpecial xlPasteValues

'2
Sheets("Step 2 - tblAnswerSheet").Activate
Range("b1").entirecolumn.Copy
Sheets("Step 3 - tblAnswers").Activate
Range("b1").PasteSpecial xlPasteValues
'3
Sheets("Step 2 - tblAnswerSheet").Activate
Range("g1").entirecolumn.Copy
Sheets("Step 3 - tblAnswers").Activate
Range("c1").PasteSpecial xlPasteValues
'4
Sheets("Step 2 - tblAnswerSheet").Activate
Range("k1").entirecolumn.Copy
Sheets("Step 3 - tblAnswers").Activate
Range("d1").PasteSpecial xlPasteValues

Application.CutCopyMode = False
Range("e1").Select


End Sub



Sub goodloop()
Dim SVal As Integer
Dim NumToFill As Integer
Dim Cnt As Integer

SVal = 1
NumToFill = 100
'basic for next loop

For Cnt = 0 To NumToFill - 1
ActiveCell.Offset(Cnt, 0).Value = SVal + Cnt
Next Cnt

End Sub


Sub findmaxval()
'uses excel's MAX function to find the maximum value in a range
Range("A:A").Find(Application.WorksheetFunction.Max(Range("A:A"))).Activate
'willing to bet this exact same line could work for similar functions
End Sub


Sub getmeouttahere()
'code for in case you need a message box that cancels if the user clicks no
If MsgBox("TEXT HERE", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

End Sub


Sub killa()
Kill "C:\docs\*.*" 'deletes all files in folder so that it's empty

RmDir "C:\docs" 'deletes an empty folder

MkDir "C:\docs" 'creates a folder that doesn't exist

End Sub


    Function sPing(sHost) As String
     
        Dim oPing As Object, oRetStatus As Object
     
        Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
          ("select * from Win32_PingStatus where address = '" & sHost & "'")
        On Error GoTo errh
        For Each oRetStatus In oPing
            If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
                sPing = "Status code is " & oRetStatus.StatusCode
            Else
                sPing = "Pinging " & sHost & " with " & oRetStatus.BufferSize & " bytes of data:" & Chr(10) & Chr(10)
                sPing = sPing & "Time (ms) = " & vbTab & oRetStatus.ResponseTime & Chr(10)
                sPing = sPing & "TTL (s) = " & vbTab & vbTab & oRetStatus.ResponseTimeToLive
            End If
        Next
errh:
            MsgBox "error!"
    End Function

    Sub TestPing()
        MsgBox sPing("1")
    End Sub


	Function SingleQuoteWrap(i_Value As Variant) As String
    SingleQuoteWrap = "'" & i_Value & "'"
End Function


'comes from http://peltiertech.com/change-series-formula-improved-routines/

Sub ChangeSeriesFormulaAllCharts()
    ''' Do all charts in sheet
    Dim oChart As ChartObject
    Dim OldString As String, NewString As String
    Dim mySrs As Series
    Dim iChartType As XlChartType
    Dim sFormula As String

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString = InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        For Each oChart In ActiveSheet.ChartObjects
            For Each mySrs In oChart.Chart.SeriesCollection
                sFormula = ""
                On Error Resume Next
                sFormula = mySrs.Formula
                On Error GoTo 0
                ' change to column chart if series is inaccessible
                If Len(sFormula) = 0 Then
                    iChartType = mySrs.ChartType
                    mySrs.ChartType = xlColumnClustered
                End If
                mySrs.Formula = WorksheetFunction.Substitute(mySrs.Formula, _
                    OldString, NewString)
                If Len(sFormula) = 0 Then mySrs.ChartType = iChartType
            Next
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
	End If
End Sub
    

Sub ChangeSeriesFormulaAllChartsAllSheets()
    ''' Do all charts in all sheets
    Dim oWksht As Worksheet
    Dim oChart As ChartObject
    Dim OldString As String, NewString As String
    Dim mySrs As Series

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString = InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        For Each oWksht In ActiveWorkbook.Worksheets
            For Each oChart In oWksht.ChartObjects
                For Each mySrs In oChart.Chart.SeriesCollection
                    mySrs.Formula = WorksheetFunction.Substitute(mySrs.Formula, _
                        OldString, NewString)
                Next
            Next
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
    End If
End Sub


Sub DeathToApostrophe() 
    Dim s As Range, temp As String 
    If MsgBox("Are you sure you want to remove all leading apostrophes from the entire sheet?", _ 
    vbOKCancel + vbQuestion, "Remove Apostrophes") = vbCancel Then Exit Sub 
    Application.ScreenUpdating = False 
    For Each s In ActiveSheet.UsedRange 
        If s.HasFormula = False Then 
             'Gets text and rewrites to same cell without   the apostrophe.
            s.Value = s.Text 
        End If 
    Next s 
    Application.ScreenUpdating = True 
End Sub 

Sub splitByColtoRows()
	' Stack Overflow's method takes fewer lines than mine.
    Application.ScreenUpdating = False
	Dim r As Range, i As Long, ar
    Set r = Worksheets("Sheet1").Range("B999999").End(xlUp)
    Do While r.row > 1
        if r.row mod 100 = 0 then debug.print(r.row)
		ar = Split(r.value, ",")
        If UBound(ar) >= 0 Then r.value = ar(0)
        For i = UBound(ar) To 1 Step -1
            r.EntireRow.Copy
            r.Offset(1).EntireRow.Insert
            r.Offset(1).value = ar(i)
        Next
        Set r = r.Offset(-1)
    Loop
	Application.CutCopyMode = false
End Sub
