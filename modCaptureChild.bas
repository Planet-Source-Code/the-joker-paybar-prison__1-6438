Attribute VB_Name = "Module1"

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As Where) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Type Where
    Pointa As Long
    Pointb As Long
End Type

    Public Const WM_CLOSE = &H10
    Public Const SW_HIDE = 0
    Public Const SW_MAXIMIZE = 3
    Public Const SW_SHOW = 5
    Public Const SW_MINIMIZE = 6

Public wincount As Integer

Sub WindowHandle(win, cas As Long)
    'by storm
    'Case 0 = CloseWindow
    'Case 1 = Show Win
    'Case 2 = Hide Win
    'Case 3 = Max Win
    'Case 4 = Min Win
    Select Case cas
        Case 0:
        Dim X%
        X% = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
        X = ShowWindow(win, SW_SHOW)
        Case 2:
        X = ShowWindow(win, SW_HIDE)
        Case 3:
        X = ShowWindow(win, SW_MAXIMIZE)
        Case 4:
        X = ShowWindow(win, SW_MINIMIZE)
    End Select

End Sub
