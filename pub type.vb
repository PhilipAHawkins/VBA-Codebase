Public Type RECT
    left As Integer
    right As Integer
    top As Integer
    bottom As Integer
End Type

Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal HWnd As Long, ByVal lpRECT As RECT) As Long
Sub AAA()
    Dim WinText As String
    Dim HWnd As Long
    Dim L As Long
    HWnd = GetForegroundWindow()
    WinText = String(255, vbNullChar)
    L = GetWindowText(HWnd, WinText, 255)
    WinText = left(WinText, InStr(1, WinText, vbNullChar) - 1)
    Debug.Print L, WinText
End Sub
Sub Attempt()
    Dim HWnd As Long
    Dim WndRect As RECT
    HWnd = FindWindow(vbNullString, "Untitled - Notepad")
    x = GetWindowRect(HWnd, WndRect)
End Sub