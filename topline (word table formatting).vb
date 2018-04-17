Sub topline2 () 
Dim oCell As Word.Cell 
Dim strCellString As Long 
Application.ScreenUpdating = False 
'applies a style to every table in the document where if a cell in the table has text in it, a black top border is applied. 'version that only does a few tables, the print statement doesn't loop for some reason. For i = 1 To ActiveDocument.Tables.Count 
for i = 1 to ActiveDocument.Tables.Count
	For Each oCell In ActiveDocument.Tables(i).Range.Cells 
		'Debug .Print (oCell.Columnindex & oCell.Rowindex & oCell.Range.Text) 
		strCellString = Len(oCell.Range.Text) 
		If strCellString >= 3 Then 
			'got a weird artifact in blank cells, using a Len method to get around it. 
			With oCell.Borders(wdBorderTop) 
				.LineStyle = Options.DefaultBorderLineStyle 
				.LineWidth = Options.DefaultBorderLineWidth 
				.Color= Options.DefaultBorderColor 
			End With 
		End If 
		If Left(oCell.Range.Text, 2) = "FY" Then 
			Tables(i) .Rows(oCell.Rowindex) .Shading.BackgroundPatternColor -738132071 
		End If 
	Next oCell Debug.Print ("finished table" & i) 
Next i 
Debug. Print ("EWIR")
For j = 1 To 6 
	ActiveDocument.Tables(j).Select 
	Selection.InsertRowsAbove 1 
	Selection.TypeText Text:= "Quarter" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:= "Action" 
	Debug.Print (Tables (j) .Cell( l , 1) .Range.Text) 
	Selection.MoveRight Unit: =wdCell 
	Selection.TypeText Text: ="ELNOT" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:= "Emitter Name" 
	Selection. MoveRight Unit: =wdCell 
	Selection.TypeText Text:= "Emitter Function" 
	Tables(j) .Rows(l) .HeadingFormat = wdToggle 
	Tables(j) .Rows(l) .Shading.BackgroundPatternColor &HABABAB 
	CYear = CYear + 1 
Next j 
Debug.Print ("Signatures" ) 
For j = 7 To 12 
	ActiveDocument.Tables(j).Select 
	Selection.InsertRowsAbove 1 
	Selection.TypeText Text:= "Quarter" 
	Selection.MoveRight Unit: =wdCell 
	Selection.TypeText Text:= "Action" 
	Selection.MoveRight Unit: =wdCell 
	Selection.TypeText Text: ="Equipment Code" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:= "Emitter Name" 
	Tables (j) .Rows (l).HeadingFormat = wdToggle 
	Tables(j) .Rows(l) .Shading.BackgroundPatternColor &HABABAB 
	CYear = CYear + 1 
Next j 
Debug. Print ( "C&P 11 ) 
For j = 13 To 18 
	ActiveDocument .Tables (j) .Select 
	Selection.InsertRowsAbove 1 
	Selection.TypeText Text:="Quarter" 
	Selection.MoveRight Unit: =wdCell 
	Selection.TypeText Text: ="Action" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:="Equipment Code" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text: ="Emitter Name" 
	Tables(j).Rows(l).HeadingFormat = wdToggle 
	Tables (j) .Rows (l ).Shading.BackgroundPatternColor = &HABABAB 
	CYear = CYear + 1 
Next j 
End Sub 

Sub runtheactions ()
Dim oCell As Word.Cell 
Dim CurrPage As Long 
Dim NewPage As Long 
Dim RowPage As Long 
Dim CurrAct As String 
Dim NewAct As Boolean 
Dim i As Long 
'get the page number of the first row 
CurrPage = 1 
Debug . Print (CurrPage) 
'start the loop 
For i = 1 To ActiveDocument.Tables.Count 
	For Each oCell In ActiveDocument.Tables (i ) . Columns (2) .Cells 'set the text in col 2 as var but only if not blank 
		If Len (oCell.Range. Text) > 2 Then 
			NewAct = True 
			CurrAct = oCell.Range . Text 
		End If 
		'Check what page the row is on 
		oCell.Select 
		RowPage = Selection.Information(wdActiveEndPageNumber) 
		'if next row has different page info than page number of first row, 
		If RowPage <> CurrPage Then 'increase first row's page number by 1 
			CurrPage = CurrPage + 1 'if blank, repeat the var with" (cont.)" 
			If NewAct = False Then 
				oCell.Range.Text = CurrAct & " (cont .)" 
			End If 
		End If 
		NewAct = False 
		Debug.Print (RowPage) 
	Next oCell 
Next i 
End Sub 

Sub runtheyears() 
'same but Columns(l) because pseudocode. 
Dim oCell As Word.Cell 
Dim CurrPage As Long 
Dim NewPage As Long 
Dim RowPage As Long 
Dim CurrAct As String 
Dim NewAct As Boolean 
Dim i As Long 
'get the page number of the first row 
CurrPage = 1 
Debug.Print (CurrPage) 
'start the loop 
For i = 1 To ActiveDocument.Tables.Count 
	For Each oCell In ActiveDocument.Tables(i) .Columns(l) .Cells 'set the text in col 2 as var but only if not blank 
		If Len(oCell.Range.Text) > 2 Then 
			NewAct = True 
			CurrAct oCell.Range.Text 
		End If 
		'Check what page the row is on 
		oCell.Select RowPage = Selection.Information(wd.ActiveEndPageNumber) 
		'if next row has different page info than page number of first row, 
		If RowPage <> CurrPage Then 'increase first row's page number by 1 
			CurrPage = CurrPage + 1 'if blank, repeat the var with" (cont.)" 
			If NewAct = False Then 
				oCell.Range.Text CurrAct & " (cont .) " 
			End If 
		End If 
		NewAct = False 
		Debug.Print (RowPage) 
	Next oCell 
Next i 
End Sub

Sub topline3 () 
Dim oCell As Word.Cell 
Dim strCellString As Long 
Application.ScreenUpdating = False 
'applies a s tyle to every table in the document where if a cell in the table has text in it, a black top border is applied. 
'version that only does a few tables, the print statement doesn't loop for some reason. 
For i = 19 To ActiveDocument.Tables . Count 
	For Each oCell In ActiveDocument.Tables(i) .Range.Cells 
		'Debug.Print (oCell.Columnindex & oCell.Rowindex & oCell.Range.Text) 
		strCellString = Len (oCell .Range .Text) 
		If strCellString >= 3 Then 
			'got a weird artifact in blank cells, using a Len method to get around it. 
			With oCell.Borders(wdBorderTop) 
				.LineStyle = Options.DefaultBorderLineStyle 
				.LineWidth = Options.DefaultBorderLineWidth 
				.Color = Options.DefaultBorderColor 
			End With
		End If 
		If Left(oCell.Range.Text , 2) = "FY" Then 
			Tables(i) .Rows(oCell.Rowindex) . Shading .BackgroundPatternColor -738132071 
		End If 
	Next oCell 
	Debug.Print ("f inished table " & i ) 
Next i
For j = 19 To 22 
	ActiveDocument.Tables (j) . Select 
	Selection.InsertRowsAbove 1 
	Selection.TypeText Text:="Quarter" 
	Selection.MoveRight Unit: =wdCell 
	Selection.TypeText Text:="Action" 
	Selection. MoveRight Unit: =wdCell 
	Selection.TypeText Text:="Requirement Object" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:="Fidelity" 
	Selection.MoveRight Unit:=wdCell 
	Selection.TypeText Text:="Reactivity" 
	Tables (j) .Rows( l ) .HeadingFormat = wdToggle 
	Tables(j ) .Rows( l ) .Shading.BackgroundPatternColor &HABABAB 
	CYear = CYear + 1 
Next j 
end sub 

Sub fontmakeplz () 
Selection.WholeStory 
Selection.Font.Name= "Arial" 'why the he ck is that the only way to change font? 
For i = 1 To ActiveDocument .Tables.Count Tables(i) . Select 
	With Selection.Font 
		.Size = 11 .Bold = False 
	End With 
Next i 
Debug.Print ("Font formatted" ) 
End Sub 

Sub split_ comma_cells_into_new_ rows() 
' s e t Ce lls (i, 3) to the column to split on where "3" is the column number (eg, Column C). 
' on about 2300 rows this took about 10 seconds. 
Dim myString As String 
Dim Values_array() As String 
Dim z As Integer 
Dim x as Integer 
Dim i As Integer 
Application.StatusBar = True 
Application.ScreenUpdating False 
i = 2 
Do While Range( "A" & i) <> "" 
	myString = Cells(i, 3) z = Int (countchrs (myString, " , ")) 
	If z > 0 Then 
		myString = Cells(i, 3) .Value 
		Values_array = Split(myString, " , " ) 
		' t he array starts at 0 . 
		For x = 0 To z 
			Values_array(x) = Trim(Values_array(x)) Rows (i) . Copy 
			Rows(i + 1) .Insert Shift:=xlUp 
			Range( "C" & i) .Value= Values_array(x) 
			i = i + 1 
			If x = z Then 
				Range( "C" & i) .EntireRow.Delete 
				i = i -1 
			End If 
		Next x 
	Else 'do nothing End If 
	i = i + 1 
Loop 
Application.CutCopyMode = False 
Application.StatusBar False 
End Sub 

Public Function countchrs(expression As String, character As String) As Long 
Dim iResult As Long 
Dim sParts() As String 
sParts = Split(expression, character) 
iResult = UBound(sParts, 1) 
If (iResult = -1) Then 
	iResult = 0 
End If 
countchrs = iResult 
End Function 
