Sub coming_from()

' Make sure data is sorted ahead of time.
' Put a column named Path in Col J.

' Determine steps of a path when the data doesn't provide it but 
' does provide dates when steps occurred.
' Paths are identified with a unique number in Col I.
' Previous destinations for each step are identified in Col L.
' Current destination is Col D.
' if a row is a step's 1st leg, outputs "start".
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
        Else
            MsgBox ("data out of order, plz start over")
            Exit Sub
        End If
    End If
Next i

End Sub

Sub coming_from()

' Make sure data is sorted ahead of time.
' Put a column named Path in Col J.

' Determine steps of a path when the data doesn't provide it but 
' does provide dates when steps occurred.
' Paths are identified with a unique number in Col I.
' Previous destinations for each step are identified in Col L.
' Current destination is Col D.
' if a row is a step's 1st leg, outputs "start".
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
        Else
            MsgBox ("data out of order, plz start over")
            Exit Sub
        End If
    End If
Next i

End Sub