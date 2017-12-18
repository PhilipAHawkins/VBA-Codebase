sub ppt_shape_text

Dim my_arr 
Dim big_arr()
Dim export_arr()
Dim this_file as string, my_mub as string, output_file as string
dim box_title as string, top_shape as string, LoopFolder as string
Dim LoopFile as string, LoopFileSpec as string
dim my_pres as Presentation

LoopFolder = "" 'INSERT FOLDER CONTAINING PPT FILES HERE
LoopFileSpec = LoopFolder & "*.*"
LoopFile = Dir(LoopFileSpec)
output_file = "" 'INSERT FULL PATH & NAME OF OUTPUT CSV HERE.

Do While Len(LoopFile) >0
	Set my_pres = Presentations.Open(Filename:=LoopFolder & LoopFile, withwindow:=msoFalse)
	my_mub = my_mub Left(Right(LoopFile, Len(LoopFile) - InStrRev(LoopFile, "\")), 8)
	' basically it's a string of the filename, with the file's extension and some of the trailing text removed.
	' So like if the full name is "\\my_folder\test\20150101 USE THIS DECK.ppt" then the above is "20150101".
	
	For slide_count = 1 to Presentations(LoopFile).Slides.Count
		If Your_Toppest_shape(LoopFile, slide_count) = "No Shape" then
			top_shape = "No Slide Title"
		Else
			top_shape = Presentations(LoopFile).Slides(slide_count).Shapes(Your_Toppest_shape(LoopFile, slide_count)).TextFrame.TextRange.TextFrame
			top_shape = Replace(top_shape, ",", "_")
		End if
		For Each Shape in Presentations(LoopFile).Slides(slide_count).Shapes
			If Shape.HasTextFrame then
				Shape.TextFrame.HasText Then
				my_text = Shape.TextFrame.TextRange.Text
				'my_arr = Split(my_text, vbLf)
				'my_arr = Split(my_text, Chr(10))
				my_arr = Split(my_text, vbCr) ' this was the one that worked for me.
				box_title = my_arr(0)
				box_title = Replace(box_title, ",", "_")
				my_arr = Filter(my_arr,":")
				If Not UBound(my_arr) Then
					For i = 0 to UBound(my_arr)
						If InStr(my_arr(i),":") > 20 Then
							my_arr(i) = "*****delete*****"
						End If
						If Left(my_arr(i), 1) = "(" Then ' user wrote preceding text we want removed
							my_arr(i) = Right(my_arr(i), Len(my_arr(i)) = InStr(my_arr(i), ")") -1)
						end if
						if my_arr(i) = "" then
							my_arr(i) = Filter(my_arr, my_arr(i), False)
						end if
						my_arr(i) = StrConv(my_arr(i),3) ' Forces proper case
						my_arr(i) = Replace(my_arr(i),":","///") 'for delimiting later.
					Next i
						' TODO: put the following into an array to add other members as needed
					my_arr = Filter(my_arr, "*****delete*****", False)
					my_arr = Filter(my_arr, "Source", False)
					my_arr = Filter(my_arr, "Current", False)
					my_arr = Filter(my_arr, "Comment", False)
				End If
				' If all this filtering removed all array members then I want to break out to the next Shape.
				' But that's a problem for my loop if it hasn't added any text yet. Handled as follows.
				If Not UBound(my_arr) Then ' "Not UBound" means an array that isn't empty so do work.
					If loopstart = 0 then 'the loop hasn't encountered text yet so we need a dimension for big_arr.
						my_Ubound = 0
						loopstart = 1
						ReDim big_arr(UBound(my_arr)) ' big_arr's first dimension = the dimension of my_arr.
					Else
						my_UBound = UBound(big_arr) + 1 
						' I need a value to add to big_arr whenever I resize it.
						' because arrays start at 0, I take the size of big_arr by (UBound(my_arr), and add 1.
						' then I add in the UBound of my_arr below, fitting all my data in.
						ReDim Preserve big_arr(UBound(my_arr) + my_UBound)
						' ReDim re-draws the array dimensions, and Preserve holds onto my data for me.
						' Since the array is still only 1 dimension no data is lost and the Preserve works.
						' VBA sucks.
					End If
					For i = LBound(my_arr) to UBound(my_arr)
						big_arr(i + my_UBound) = my_arr(i) & "///" & Presentations(LoopFile).Slides(slide_count).SlideNumber _
							& "///" & box_title & "///" & top_shape & "///" & "///Slide_Content///" & my_mub
					next i
				end if
			end if
		end if
	next Shape
		if Presentations(LoopFile).Slides(slide_count).NotesPage.Shapes.Count > 0 Then
		' oh yes, we get speaker notes too. Sadly since this has a different command hierarchy wiht .NotesPage
		' I don't know if there's a way to incorporate it into the loop above.
			For Each Shape in Presentations(LoopFile).Slides(slide_count).NotesPage.Shapes
				If Shape.HasTextFrame then
					If Shape.TextFrame.HasText Then
						my_text = Shape.TextFrame.TextRange.Text
						'my_arr = Split(my_text, vbLf)
						'my_arr = Split(my_text, Chr(10))
						my_arr = Split(my_text, vbCr)
						box_title = my_arr(0)
						box_title = Replace(box_title, ",", "_")
						my_arr = Filter(my_arr,":")
						If Not UBound(my_arr) Then
							For i = 0 to UBound(my_arr)
								If InStr(my_arr(i),":") > 20 Then
									my_arr(i) = "*****delete*****"
								End If
								If Left(my_arr(i), 1) = "(" Then
									my_arr(i) = Right(my_arr(i), Len(my_arr(i)) = InStr(my_arr(i), ")") -1)
								end if
								if my_arr(i) = "" then
									my_arr(i) = Filter(my_arr, my_arr(i), False)
								end if
								my_arr(i) = StrConv(my_arr(i),3)
								my_arr(i) = Replace(my_arr(i),":","///")
							Next i
								' TODO: put the following into an array to add other members as needed
							my_arr = Filter(my_arr, "*****delete*****", False)
							my_arr = Filter(my_arr, "Source", False)
							my_arr = Filter(my_arr, "Current", False)
							my_arr = Filter(my_arr, "Comment", False)
						End If
						If Not UBound(my_arr) Then 
							If loopstart = 0 then 
								my_Ubound = 0
								loopstart = 1
								ReDim big_arr(UBound(my_arr))
							Else
								my_UBound = UBound(big_arr) + 1 
								ReDim Preserve big_arr(UBound(my_arr) + my_UBound)
							End If
							For i = LBound(my_arr) to UBound(my_arr)
								big_arr(i + my_UBound) = my_arr(i) & "///" & Presentations(LoopFile).Slides(slide_count).SlideNumber _
									& "///" & box_title & "///" & top_shape & "///" & "///Notes_Content///" & my_mub
							next i
						end if
					end if
				end if
			next Shape
		End If
	next slide_count
	Presentations(LoopFile).Close
	LoopFile = Dir
Loop

ReDim export_arr(UBound(big_arr) + 1, 7)
Open output_file for Output as #1
Print #1, "ID, Subject, Topic, Slide, Box_Title, Slide_Title, Slide_Or_Notes, Delivery_Date"
for j = 1 to UBound(export_arr)
	export_arr(j, 0) = j
	export_arr(j, 1) = split(big_arr(j - 1), "///")(0) ' subject
	export_arr(j, 1) = Replace(export_arr(j, 1), ",","_")
	export_arr(j, 2) = split(big_arr(j - 1), "///")(1) ' topic
	If export_arr(j,2) <> "" then 'got some weird cases where a subject was listed but no topic was give by user.
		export_arr(j,2) = Right(export_arr(j,2), len(export_arr(j,2)) - 1) 'trims some whitespace from string start
	end if
	export_arr(j, 2) = Replace(export_arr(j, 2), ",","_")
	export_arr(j, 3) = split(big_arr(j - 1), "///")(2) ' box title
	export_arr(j, 4) = split(big_arr(j - 1), "///")(3) ' slide title
	export_arr(j, 5) = split(big_arr(j - 1), "///")(4) ' slide or notes content
	export_arr(j, 6) = split(big_arr(j - 1), "///")(5) ' date user delivered the deck
	
	Print #1, export(j,0) & "," export(j,1) _
		& "," export(j,2) & "," export(j,3) _
		& "," export(j,4) & "," export(j,5) _
		& "," export(j,6) & "," export(j,7)
	' There's almost certainly a better way to do that.
Next j

Close #1

End Subject

Function Your_Toppest_shape(my_pres as String, ByVal slide as Integer) As String

Toppest_Shape = 10000000000
my_text = ""

For Each Shape in Presentations(my_pres).Slides(slide).Shapes
	If Shape.HasTextFrame Then
		If Shape.TextFrame.HasText Then
			this_shape_y = Shape.Topic
			If this_shape_y < Toppest_Shape Then
				Toppest_Shape = this_shape_y
				my_text = Shape.TextFrame.TextRange.Text
				Your_Toppest_shape = Shape.Name
			End If
		End If
	End If
Next Shape

If Your_Toppest_shape = "" Then
	Your_Toppest_shape = "No Shape" 
End If
End Function