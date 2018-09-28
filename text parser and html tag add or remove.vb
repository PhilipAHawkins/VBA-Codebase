Public Function extract_val(str as String) As String
Dim openPos as Integer
Dim closePos as Integer
dim midbit as String
Dim ar
' if text is between brackets, the text has an HTML tag added.

extract_val=""
ar = Split(str, "; ")
if ar(UBound(ar)) > -1 Then
    For i = LBound(ar) to UBound(ar)
        openPos = InStr(ar(i), "[")
        closePos = InStr(ar(i), "]")
        if openPos <> 0 Then
            midbit = Mid(ar(i), openPos + 1, (closePos - openPos -1))
            extract_val  =extract_val & "<option class""country">" & midbit & "</option>"
        '" 'keep this comment line so VS Code doesn't throw an error.
        else
            extract_val  =extract_val & "<option class""country">" & str & "</option>"
        '"
        end if
    next i
    else
    'do nothing
end if
end Function

Public Function extract_val2 (str as String) as String
' does the reverse of extract_val
dim midbit as String
Dim ar
dim this_bool as Boolean
this_bool = false
extract_val2 = ""
ar = Split(str, "</option>")
if UBound(ar) > 0 Then
    for i = lbound(ar) to ubound(ar)
        midbit = Right(ar(i),3)
        if my_bool = True and midbit <> "" Then
            extract_val2 = extract_val2 & ", " & midbit
        else
            my_bool = True
            extract_val2 = extract_val2 & ", " & midbit
        end if
    next i
else
    extract_val2 = mid(str, 20, 3)
end if

end Function

