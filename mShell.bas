Attribute VB_Name = "mShell"
'================================================
' Module:        mShell.bas
' Author:        (*)
' Last revision: 2003.08.28
'================================================
'
' (*) From original work 'SHFileOperation Demo' by MrBobo
'     http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=38982&lngWId=1
'
'     Copyright PSST Software 2002
'     Submitted to Planet Source Code - September 2002
'     If you got it elsewhere - they stole it from PSC.
'
'     Written by MrBobo - enjoy
'     Please visit our website - www.psst.com.au



Option Explicit

'-- API:

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH             As Long = 260

Private Type SHITEMID
    cb   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Public Enum eSpecialFolderID
    [Default] = 0
    [Programs] = 2
    [Control Panel] = 3
    [Printers] = 4
    [My Documents] = 5
    [Favourites] = 6
    [StartUp] = 7
    [Recent] = 8
    [SendTo] = 9
    [Recycle Bin] = 10
    [Start Menu] = 11
    [Desktop] = 16
    [My Computer] = 17
    [Network] = 18
    [NetHood] = 19
    [Fonts] = 20
    [Templates] = 21
    [All users / Desktop] = 25
    [Application Data] = 26
    [PrintHood] = 27
    [Temporary Internet Files] = 32
    [Cookies] = 33
    [History] = 34
'-- See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/enums/csidl.asp
End Enum

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'//

Public Function GetSpecialFolder(ByVal SpecialFolder As eSpecialFolderID) As String
    
  Dim lRet  As Long
  Dim sPath As String
  Dim tIDL  As ITEMIDLIST
  
  Const NOERROR    As Long = 0
  Const MAX_LENGTH As Long = 260
  
    lRet = SHGetSpecialFolderLocation(GetDesktopWindow, SpecialFolder, tIDL)
    If (lRet = NOERROR) Then
        sPath = Space$(MAX_LENGTH)
        lRet = SHGetPathFromIDList(ByVal tIDL.mkid.cb, ByVal sPath)
        If (lRet) Then
            GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
End Function
