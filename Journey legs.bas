Sub pcfn_counter()

' Make sure data is sorted ahead of time.
' Put a column named Path in Col J.

' numbers the legs of a journey taken in a large dataset. in the data, various
' journeys are taken by various ships over a time series, and their info is listed
' in the data. Col G identifies the tracking itinerary number. data is ordered by 
' date, descending. if the code sees a new itinerary number, it places "1" in Col J.
' if the itinerary number isn't new, it increments Col J by 1.
Dim a As Long
Application.ScreenUpdating = False

a = Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A"))

For i = 2 To a
    If i Mod 1000 = 0 Then Debug.Print (i)
    If Range("G" & i) = Range("G" & i - 1) Then
        Range("j" & i).Value = Range("j" & i - 1).Value + 1
    Else
        Range("j" & i).Value = 1
    End If
Next i
    
End Sub

' this is what my formulas looked like trying to do this w/o VBA. Nasty.
'=IF(L3=1,"Start",INDEX(C:C,MATCH(1,(L3-1=L:L)*(I3=I:I),0)))
'INDEX(C:C,MATCH(1,(L3-1=L:L)*(I3=I:I),0))

Sub coming_from()

' Make sure data is sorted ahead of time.
' Put a column named Path in Col J.

' Finds the port that a ship just came from on its journey.
' journeys are identified with a tracking number in Col I.
' Previous ports for each journey leg are identified in Col L.
' Places the name of this port in Col D.
' if a row is a ship's 1st leg, outputs "start".
' forgot to add an output for "end", so it's handled in Sub going_to.
Dim a As Long
Application.ScreenUpdating = False

a = Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A"))

For i = 2 To a
    If i Mod 1000 = 0 Then Debug.Print (i)
    'max_legs = Application.WorksheetFunction.CountIf(Range("I:I"), Range("i" & i))
    If Range("l" & i) = 1 Then
        Range("d" & i).Value = "Start"
    'ElseIf Range("l" & i) = max_legs Then
        'Range("d" & i).Value = "End"
    Else
        my_pcfn = (ActiveSheet.Range("i" & i).Value)
        If Range("i" & i - 1) = my_pcfn Then
            Range("d" & i).Value = Range("c" & i - 1)
        Else:
            MsgBox ("data out of order, plz start over")
            GoTo a
        End If
    End If
Next i
a:
End Sub


Sub fix_end()
Dim a As Long
Application.ScreenUpdating = False

a = Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A"))

For i = 2 To a
    If Range("d" & i).Value = "End" Then
        Range("d" & i).Value = Range("c" & i - 1)
    Else
        ' do nothing
    End If
Next i
    
End Sub


Sub going_to()

Dim a As Long
Application.ScreenUpdating = False

a = Application.WorksheetFunction.CountA(ActiveSheet.Range("A:A"))
my_pcfn = 0
max_legs = 0

For i = 2 To a
    If i Mod 1000 = 0 Then Debug.Print (i)
    If my_pcfn <> ActiveSheet.Range("j" & i).Value Then
        my_pcfn = ActiveSheet.Range("j" & i).Value
        max_legs = Application.WorksheetFunction.CountIf(Range("j:j"), my_pcfn)
    End If
    If Range("m" & i) = max_legs Then
        Range("e" & i).Value = "End"
    Else
        Range("e" & i).Value = Range("c" & i + 1)
    End If
Next i
End Sub


=IF($D2="Start","",INDEX(J:J,MATCH($D2,$C:$C,0)))
=IF($D2="Start","",INDEX(K:K,MATCH($D2,$C:$C,0)))
=IF($E2="End","",INDEX(J:J,MATCH($E2,$C:$C,0)))
=IF($E2="End","",INDEX(K:K,MATCH($E2,$C:$C,0)))