Attribute VB_Name = "modAlwaysONTop"
Option Explicit

Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private bOnTopState     As Boolean

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, _
    ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long)

Public Sub AlwaysOnTop(frmName As Form, bState As Boolean)
Dim lFlag As Long
    On Error Resume Next
    If bState = True Then
       lFlag = HWND_TOPMOST
    Else
       lFlag = HWND_NOTOPMOST
    End If
    bOnTopState = bState
    Call SetWindowPos(frmName.hWnd, lFlag, 0&, 0&, 0&, 0&, _
        (SWP_NOSIZE Or SWP_NOMOVE))
End Sub


