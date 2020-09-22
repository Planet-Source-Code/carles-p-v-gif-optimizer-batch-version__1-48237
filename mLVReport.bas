Attribute VB_Name = "mLVReport"
'================================================
' Module:        mLVReport.bas
' Author:        (*)
' Last revision: 2003.08.23
'================================================
'
' (*) From original work 'Explorer v1.3' by Nelly
'     http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=31549&lngWId=1



Option Explicit

'-- API:

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH  As Long = 260
Private Const AMAX_PATH As Long = 260

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes   As Long
    ftCreationTime     As FILETIME
    ftLastAccessTime   As FILETIME
    ftLastWriteTime    As FILETIME
    nFileSizeHigh      As Long
    nFileSizeLow       As Long
    dwReserved0        As Long
    dwReserved1        As Long
    cFileName          As String * MAX_PATH
    cAlternateFileName As String * AMAX_PATH
End Type

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL    As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT   As Long = 10

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function PathFindExtension Lib "Shlwapi" Alias "PathFindExtensionW" (ByVal pPath As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Type SHFILEINFO
    hIcon         As Long
    iIcon         As Long
    dwAttributes  As Long
    szDisplayName As String * MAX_PATH
    szTypeName    As String * 80
End Type

Private Const SHGFI_LARGEICON     As Long = &H0
Private Const SHGFI_SMALLICON     As Long = &H1
Private Const SHGFI_SHELLICONSIZE As Long = &H4
Private Const ILD_TRANSPARENT     As Long = &H1
Private Const ILD_NORMAL          As Long = &H0

Private Const SHGFI_TYPENAME     As Long = &H400
Private Const SHGFI_SYSICONINDEX As Long = &H4000
Private Const SHGFI_DISPLAYNAME  As Long = &H200
Private Const SHGFI_EXETYPE      As Long = &H2000

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
                                                                             
'//

'-- Private variables:

Private m_GIFIconGot As Boolean ' GIF associated icon has been got
Private m_FilesCount As Long    ' Current folder files count
Private m_FilesSize  As Long    ' Current folder files total size
            
'========================================================================================
' Properties
'========================================================================================
            
Public Property Get FilesCount() As Long
    FilesCount = m_FilesCount
End Property

Public Property Get FilesSize() As Long
    FilesSize = m_FilesSize
End Property
            
'========================================================================================
' Methods
'========================================================================================

Public Sub FillLV(oFileList As ListView, ByVal sFolderPath As String, ByVal sMask As String)

  Dim lReturn    As Long
  Dim lNextFile  As Long
  Dim sPath      As String
  Dim sFilename  As String
  Dim lstItem    As ListItem
  Dim lstSubItem As ListSubItem
  Dim WFD        As WIN32_FIND_DATA
  Dim SYSTIME    As SYSTEMTIME
    
    With oFileList
        
        '-- Reset module vars.
        m_FilesCount = 0
        m_FilesSize = 0
        
        '-- Hide listview (faster filling) and clear it
        .Visible = 0
        .ListItems.Clear
        
        '-- Build masked path
        sPath = sFolderPath & sMask
        '-- Search for first file
        lReturn = FindFirstFile(sPath, WFD) & Chr$(0)
                        
        '-- No files: end
        If (lReturn <= 0) Then
            .Visible = -1
            Exit Sub
        End If
        
        Do  '-- Skip directories...
            If (Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory) Then
                
                '-- Get filename...
                sFilename = pvStripNullChar(WFD.cFileName)
                If (IsEmpty(sFilename) = 0) Then
                        
                    '-- GIF icon extracted [?]
                    If (m_GIFIconGot = 0) Then
                        m_GIFIconGot = -1
                        pvGetAssociatedIcon sFolderPath & sFilename
                    End If
                    '-- Add listview item/subitems
                    Set lstItem = .ListItems.Add(, , sFilename, , 2): lstItem.Checked = -1
                    Set lstSubItem = lstItem.ListSubItems.Add(, , Format$(WFD.nFileSizeLow, "@@@@@@@@@@"))
                    Set lstSubItem = lstItem.ListSubItems.Add(, , Chr$(45))
                    Set lstSubItem = lstItem.ListSubItems.Add(, , Chr$(45))
                End If
            End If
            '-- Files count and total files size
            m_FilesCount = m_FilesCount + 1
            m_FilesSize = m_FilesSize + WFD.nFileSizeLow
            
            '-- Next file
            lNextFile = FindNextFile(lReturn, WFD)
        
        Loop Until lNextFile <= 0
    
        '-- Close search handle
        lNextFile = FindClose(lReturn&)
        
        '-- Make listview visible
        .Visible = -1
    End With
End Sub

Public Function ShellOpenFile(sCorrectPath As String, sFilename As String, lhWnd As Long) As Long
    
    '-- Open file through associated application
    ShellOpenFile = ShellExecute(lhWnd, "open", sFilename, vbNullString, sCorrectPath, SW_SHOWNORMAL)
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvGetAssociatedIcon(ByVal sFilename As String)
 
  Dim lSysImHndle As Long
  Dim lRet        As Long
  Dim liItem      As ListImage
  Dim SHI         As SHFILEINFO
                
    lSysImHndle = SHGetFileInfo(sFilename, 0&, SHI, Len(SHI), SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
    lRet = ImageList_Draw(lSysImHndle, SHI.iIcon, fMain.picBuff.hDC, 0, 0, ILD_NORMAL)
    
    With fMain.ilLV
        Set liItem = .ListImages.Add(, , fMain.picBuff.Image)
    End With
End Sub

Private Function pvStripNullChar(sInput As String) As String

  Dim x As Integer
    
    x = InStr(1, sInput, Chr$(0))
    If (x > 0) Then
        pvStripNullChar = Left$(sInput, x - 1)
    End If
End Function

