Attribute VB_Name = "mBFF"
'================================================
' Module:        mBFF.bas
' Author:        (*)
' Last revision: 2003.08.23
'================================================
'
' (*) From original work 'Bobo System Treeview Thievery' by MrBobo
'     http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=40007&lngWId=1
'     Copyright PSST Software 2002

Option Explicit

'-- API:

Private Type BROWSEINFO
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const WM_MOVE              As Long = &H3
Private Const WM_USER              As Long = &H400
Private Const BIF_STATUSTEXT       As Long = &H4&
Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const BFFM_INITIALIZED     As Long = 1
Private Const BFFM_SELCHANGED      As Long = 2
Private Const BFFM_SETSELECTION    As Long = (WM_USER + 102)
Private Const MAX_PATH             As Long = 260

Private Const GWL_WNDPROC          As Long = (-4)

Private Const GW_NEXT              As Long = 2
Private Const GW_CHILD             As Long = 5
Private Const WM_CLOSE             As Long = &H10

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'//

Private m_lpPrevWndProc    As Long
Private m_sCurrentFolder   As String
Private m_sInitFolder      As String
Private m_hDialogWindow    As Long
Private m_hSysTreeWindow   As Long
Private m_oDialogContainer As Object
Private m_bNormalDialog    As Boolean

'//

'========================================================================================
' Properties
'========================================================================================

Public Property Get CurrentFolder() As String
    
    CurrentFolder = pvTrimNull(m_sCurrentFolder)
End Property

Public Property Let InitFolder(ByVal New_InitFolder As String)

    m_sInitFolder = New_InitFolder
End Property

'========================================================================================
' Methods
'========================================================================================

Public Sub SetBFFContainer(oContainer As Object)
    
    Set m_oDialogContainer = oContainer ' Container for the Treeview
    Call pvBrowseForFolder              ' Spawn the dialog
End Sub

Public Sub SizeBFF(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, bRedraw As Boolean)
    
    '-- Called on the resize event of the Container holding the Treeview
    Call MoveWindow(m_hSysTreeWindow, nLeft, nTop, nWidth, nHeight, bRedraw)
End Sub

Public Sub ChangeBFFFolder(ByVal NewFolder As String)
    
    '-- We call this sub to change the path of the Treeview
    m_sCurrentFolder = NewFolder
    '-- Tell BrowseForFolder what to do
    Call SendMessage(m_hDialogWindow, BFFM_SETSELECTION, 1, m_sCurrentFolder)
End Sub

Public Sub CloseUpBFF()
    
    '-- Send the Treeview back to the BrowseForFolder dialog
    Call SetParent(m_hSysTreeWindow, m_hDialogWindow)
    '-- Close the dialog
    Call SendMessage(m_hDialogWindow, WM_CLOSE, 1, 0)
    '-- Just to be sure...
    Call DestroyWindow(m_hDialogWindow)
End Sub

Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal Title As String) As String
'-- Standard BrowseForFolder dialog
    
  Dim tBI      As BROWSEINFO
  Dim lpIDList As Long
  Dim sTitle As String
  Dim sBuffer  As String
  
    sTitle = StrConv(Title, vbFromUnicode)
    With tBI
        .hWndOwner = hWndOwner
        .lpszTitle = StrPtr(sTitle)
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfnCallback = pvGetAddressOfFunction(AddressOf pvBrowseCallbackProc)
    End With
    
    '-- Open browse for folder dialog...
    m_bNormalDialog = -1
    lpIDList = SHBrowseForFolder(tBI)
    m_bNormalDialog = 0
    '   ... and return selected folder
    sBuffer = Space(MAX_PATH)
    Call SHGetPathFromIDList(lpIDList, sBuffer)
    BrowseForFolder = pvTrimNull(sBuffer)
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvBrowseForFolder()

  Dim tBI As BROWSEINFO
    
    With tBI
        .hWndOwner = GetDesktopWindow
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfnCallback = pvGetAddressOfFunction(AddressOf pvBrowseCallbackProc)
    End With
    
    '-- Prevents app. title bar flickering (deact./act.)
    Call LockWindowUpdate(tBI.hWndOwner)
    '-- Open browse for folder dialog...
    Call SHBrowseForFolder(tBI)
End Sub

Private Function pvBrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList  As Long
  Dim lRet      As Long
  Dim sBuffer   As String
  Dim hWndT     As Long
  Dim ClWind    As String * 14
  Dim ClCaption As String * 100
    
    If (m_bNormalDialog = 0) Then m_hDialogWindow = hWnd
    
    Select Case uMsg
    
        Case BFFM_INITIALIZED
        
            If (m_bNormalDialog = 0) Then
            
                '-- Enumerate child windows
                hWndT = GetWindow(hWnd, GW_CHILD)
                
                Do While hWndT <> 0
                
                    GetClassName hWndT, ClWind, 14
                    
                    '-- Here's what we're really after - it's Treeview!
                    If (Left$(ClWind, 13) = "SysTreeView32") Then
                        m_hSysTreeWindow = hWndT
                        Exit Do
                    End If
                    hWndT = GetWindow(hWndT, GW_NEXT)
                Loop
                '-- Steal the Treeview for our own use
                pvGrabBFF m_oDialogContainer
            End If
            
            '-- Change to current folder
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_sInitFolder)
            
        Case BFFM_SELCHANGED
        
            If (m_bNormalDialog = 0) Then
            
                '-- Path has changed - better tell our form
                sBuffer = Space(MAX_PATH)
                lRet = SHGetPathFromIDList(lp, sBuffer)
                m_sCurrentFolder = sBuffer
                
                '-- Raise event...
                Call fMain.PathChange
            End If
    End Select
    
    pvBrowseCallbackProc = 0
End Function

Private Function pvGetAddressOfFunction(lAdd As Long) As Long

    pvGetAddressOfFunction = lAdd
End Function

Private Sub pvGrabBFF(oNewOwner As Object)
    
    '-- Thievery in progress...
    '   It's mine now!
    Call SetParent(m_hSysTreeWindow, oNewOwner.hWnd)
    '-- Put it where we want it
    Call SizeBFF(0, 0, oNewOwner.ScaleWidth, oNewOwner.ScaleHeight, 0)
    '-- Temporary hook to catch the move event
    Call pvDialogHook
End Sub

Private Sub pvTaskbarHide()

    '-- Hide the BrowseForFolder dialog from the Taskbar
    Call ShowWindow(m_hDialogWindow, 0)
    '-- Done with hooking
    Call pvDialogUnhook
End Sub

Private Sub pvDialogHook()
'-- This hook is very temporary and is self-cancelling
'   It is needed because we need to hide the BrowseForFolder
'   dialog from the Taskbar, but we cant do that until
'   it is fully opened. So we hook it for it's move
'   event (which we call). When it moves we hide it from
'   the Taskbar and then unhook - bit messy but it works.

    m_lpPrevWndProc = SetWindowLong(m_hDialogWindow, GWL_WNDPROC, AddressOf pvWindowProc)
End Sub

Private Sub pvDialogUnhook()

    Call SetWindowLong(m_hDialogWindow, GWL_WNDPROC, m_lpPrevWndProc)
    Call LockWindowUpdate(ByVal 0&)
End Sub

Private Function pvWindowProc(ByVal mHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If (uMsg = WM_MOVE) Then Call pvTaskbarHide
    pvWindowProc = CallWindowProc(m_lpPrevWndProc, mHwnd, uMsg, wParam, lParam)
End Function

Private Function pvTrimNull(StartString As String) As String
  
  Dim lPos As Long
  
    lPos = InStr(StartString, Chr$(0))
    If (lPos) Then
        pvTrimNull = Left$(StartString, lPos - 1)
      Else
        pvTrimNull = StartString
    End If
End Function

