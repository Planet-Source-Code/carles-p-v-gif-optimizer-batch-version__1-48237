Attribute VB_Name = "mSettings"
'================================================
' Module:        mSettings.bas
' Author:        -
' Last revision: 2003.09.07
'================================================



Option Explicit
Option Compare Text

'-- API:

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'//

Private m_sBrowseFolder As String
Private m_sBackupFolder As String
Private m_sTargetFolder As String
Private m_sFileMask     As String

Public Property Get BrowseFolder() As String
    BrowseFolder = pvTrimNull(m_sBrowseFolder)
End Property
Public Property Let BrowseFolder(ByVal New_BrowseFolder As String)
    m_sBrowseFolder = New_BrowseFolder
End Property

Public Property Get BackupFolder() As String
    BackupFolder = pvTrimNull(m_sBackupFolder)
End Property
Public Property Let BackupFolder(ByVal New_BackupFolder As String)
    m_sBackupFolder = New_BackupFolder
End Property

Public Property Get TargetFolder() As String
    TargetFolder = pvTrimNull(m_sTargetFolder)
End Property
Public Property Let TargetFolder(ByVal New_TargetFolder As String)
    m_sTargetFolder = New_TargetFolder
End Property

Public Property Get FileMask() As String
    FileMask = m_sFileMask
End Property
Public Property Let FileMask(ByVal New_FileMask As String)
    m_sFileMask = New_FileMask
End Property

'//

Public Sub LoadSettings()

    With fMain
    
        '-- Window pos./size and state
        .Width = pvGetINI("GOb.ini", "Form", "Width", .Width)
        .Height = pvGetINI("GOb.ini", "Form", "Height", .Height)
        .Top = pvGetINI("GOb.ini", "Form", "Top", (Screen.Height - .Height) \ 2)
        .Left = pvGetINI("GOb.ini", "Form", "Left", (Screen.Width - .Width) \ 2)
        .WindowState = pvGetINI("GOb.ini", "Form", "WindowState", .WindowState)
        .mnuView(0).Checked = pvGetINI("GOb.ini", "Form", "WindowOnTop", 0)
        
        '-- Splitters
        .ucSplitterH.Left = pvGetINI("GOb.ini", "Splitters", "SplitterH", .ucSplitterH.Left)
        .ucSplitterV.Top = pvGetINI("GOb.ini", "Splitters", "SplitterV", .ucSplitterV.Top)
        
        '-- GIF viewer
        .ucGIFViewer.AutoPlay = pvGetINI("GOb.ini", "GIF viewer", "AutoPlay", -1)
        .mnuGIFViewerContext(0).Checked = .ucGIFViewer.AutoPlay
        .ucGIFViewer.BackColor = pvGetINI("GOb.ini", "GIF viewer", "Background color", vbWindowBackground)
        
        '-- LV Report (headers withs)
        .lvReport.ColumnHeaders(1).Width = pvGetINI("GOb.ini", "LV Report", "Header 1", 125)
        .lvReport.ColumnHeaders(2).Width = pvGetINI("GOb.ini", "LV Report", "Header 2", 75)
        .lvReport.ColumnHeaders(3).Width = pvGetINI("GOb.ini", "LV Report", "Header 3", 75)
        .lvReport.ColumnHeaders(4).Width = pvGetINI("GOb.ini", "LV Report", "Header 4", 60)
        
        '-- Optimization options
        .optTarget(pvGetINI("GOb.ini", "Options", "Target", 0)) = -1
        .chkBackup = pvGetINI("GOb.ini", "Options", "Backup", 0)
        .mnuOptimizationOptions(pvGetINI("GOb.ini", "Options", "Optimization method", 0)).Checked = -1
        .mnuOptimizationOptions(4).Checked = pvGetINI("GOb.ini", "Options", "Second palette optimization", -1)
        .mnuOptimizationOptions(6).Checked = pvGetINI("GOb.ini", "Options", "Disable interlacing", -1)
    
        '-- Paths / File mask
        m_sBrowseFolder = pvGetINI("GOb.ini", "Folders", "BrowseFolder", mShell.GetSpecialFolder([My Documents]))
        m_sBackupFolder = pvGetINI("GOb.ini", "Folders", "BackupFolder", mShell.GetSpecialFolder([My Documents]))
        m_sTargetFolder = pvGetINI("GOb.ini", "Folders", "TargetFolder", mShell.GetSpecialFolder([My Documents]))
        m_sFileMask = pvGetINI("GOb.ini", "File mask", "FileMask", "*")
        .txtBackupFolder = m_sBackupFolder
        .txtTargetFolder = m_sTargetFolder
        .txtFileMask = m_sFileMask
    End With
End Sub
    
Public Sub SaveSettings()

    With fMain
    
        '-- Window pos./size and state
        If (.WindowState = vbNormal) Then
            Call pvPutINI("GOb.ini", "Form", "Width", .Width)
            Call pvPutINI("GOb.ini", "Form", "Height", .Height)
            Call pvPutINI("GOb.ini", "Form", "Top", .Top)
            Call pvPutINI("GOb.ini", "Form", "Left", .Left)
        End If
        Call pvPutINI("GOb.ini", "Form", "WindowState", .WindowState)
        Call pvPutINI("GOb.ini", "Form", "WindowOnTop", .mnuView(0).Checked)
        
        '-- Splitters
        Call pvPutINI("GOb.ini", "Splitters", "SplitterH", .ucSplitterH.Left)
        Call pvPutINI("GOb.ini", "Splitters", "SplitterV", .ucSplitterV.Top)
        
        '-- GIF viewer
        Call pvPutINI("GOb.ini", "GIF viewer", "AutoPlay", .ucGIFViewer.AutoPlay)
        Call pvPutINI("GOb.ini", "GIF viewer", "Background color", .ucGIFViewer.BackColor)
        
        '-- LV Report (headers withs)
        Call pvPutINI("GOb.ini", "LV Report", "Header 1", .lvReport.ColumnHeaders(1).Width)
        Call pvPutINI("GOb.ini", "LV Report", "Header 2", .lvReport.ColumnHeaders(2).Width)
        Call pvPutINI("GOb.ini", "LV Report", "Header 3", .lvReport.ColumnHeaders(3).Width)
        Call pvPutINI("GOb.ini", "LV Report", "Header 4", .lvReport.ColumnHeaders(4).Width)
        
        '-- Optimization options
        Call pvPutINI("GOb.ini", "Options", "Target", IIf(.optTarget(0), 0, 1))
        Call pvPutINI("GOb.ini", "Options", "Backup", .chkBackup)
        Call pvPutINI("GOb.ini", "Options", "Optimization method", IIf(.mnuOptimizationOptions(0).Checked, 0, 1))
        Call pvPutINI("GOb.ini", "Options", "Second palette optimization", .mnuOptimizationOptions(4).Checked)
        Call pvPutINI("GOb.ini", "Options", "Disable interlacing", .mnuOptimizationOptions(6).Checked)
    
        '-- Paths / File mask
        pvPutINI "GOb.ini", "Folders", "BrowseFolder", m_sBrowseFolder
        pvPutINI "GOb.ini", "Folders", "BackupFolder", m_sBackupFolder
        pvPutINI "GOb.ini", "Folders", "TargetFolder", m_sTargetFolder
        pvPutINI "GOb.ini", "File mask", "FileMask", m_sFileMask
    End With
End Sub

'//

Private Function pvAppPath() As String

    pvAppPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString)
End Function

Private Sub pvPutINI(INIFile As String, INIHead As String, INIKey As String, INIVal As String)
  
  Dim sINIFileName As String
  Dim sRet         As String
  
    sINIFileName = pvAppPath & INIFile
    sRet = WritePrivateProfileString(INIHead, INIKey, INIVal, sINIFileName)
End Sub

Private Function pvGetINI(INIFile As String, INIHead As String, INIKey As String, INIDefault As String) As String

  Dim sINIFileName As String
  Dim sTemp        As String * 260
  Dim sRet         As String
    
    sINIFileName = pvAppPath & INIFile
    sRet = GetPrivateProfileString(INIHead, INIKey, INIDefault, sTemp, Len(sTemp), sINIFileName)
    pvGetINI = Trim$(sTemp)
    pvGetINI = Left$(pvGetINI, Len(pvGetINI) - 1)
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
