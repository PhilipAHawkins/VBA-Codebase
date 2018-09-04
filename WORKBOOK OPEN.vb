' Goes in "THIS WORKBOOK"

Private Sub Workbook_Open()
    Call Disable_Error_Checking
    Call Comment_Indicator
    Call Hide_Toolbars
    Call Auto_Calc_Itera
    Call Edit_Move
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

End Sub

