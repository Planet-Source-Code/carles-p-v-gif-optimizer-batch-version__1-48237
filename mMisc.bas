Attribute VB_Name = "mMisc"
'================================================
' Module:        mMisc.bas
' Author:        -
' Last revision: 2003.09.03
'================================================



Option Explicit

'-- API:

Public Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Public Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

'//

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HWND_TOPMOST   As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE     As Long = &H1
Private Const SWP_NOMOVE     As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const BDR_RAISED As Long = &H5
Private Const BDR_SUNKEN As Long = &HA
Private Const BF_RECT    As Long = &HF

Private Const DFC_SCROLL      As Long = &H3
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_PUSHED     As Long = &H200

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal un1 As Long, ByVal un2 As Long) As Long

'//

Public Sub SetFormPos(oForm As Form, ByVal bTop As Boolean)
    
    '-- Enable/Disable TopMost window pos.
    Call SetWindowPos(oForm.hWnd, IIf(bTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub RemoveButtonBorderEnhance(oButton As CommandButton)
    
    '-- 'C' button (Button.Style = 'Graphical')
    Call SendMessage(oButton.hWnd, &HF4&, &H0&, 0&)
End Sub

Public Sub DrawPopupButton(ByVal hDC As Long, lpRect As RECT2, ByVal bDown As Boolean, Optional ByVal bDisabled As Boolean = 0)
    
    '-- 'Popup button'
    Call DrawFrameControl(hDC, lpRect, DFC_SCROLL, DFCS_SCROLLDOWN Or -bDown * DFCS_PUSHED Or -bDisabled * &H100)
End Sub
