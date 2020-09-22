VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fMain 
   Caption         =   "GO - batch"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Tag             =   "Preserve"
      Top             =   6735
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3307
            MinWidth        =   3307
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3307
            MinWidth        =   3307
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7541
         EndProperty
      EndProperty
   End
   Begin GObatch.ucSplitter ucSplitterH 
      Height          =   6735
      Left            =   2910
      Top             =   0
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   11880
   End
   Begin GObatch.ucSplitter ucSplitterV 
      Height          =   90
      Left            =   0
      Top             =   4125
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   159
      Orientation     =   1
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3540
      Left            =   3000
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Preserve"
      Top             =   2700
      Width           =   5280
      Begin VB.CommandButton cmdFileMask 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Apply file mask"
         Top             =   150
         Width           =   285
      End
      Begin VB.TextBox txtFileMask 
         Height          =   300
         Left            =   3810
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   150
         Width           =   1050
      End
      Begin GObatch.ucProgress ucProgressCurrent 
         Height          =   180
         Left            =   1785
         Tag             =   "Preserve"
         Top             =   2610
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   318
         BorderStyle     =   1
      End
      Begin GObatch.ucProgress ucProgress 
         Height          =   180
         Left            =   1785
         Tag             =   "Preserve"
         Top             =   2880
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   318
         BorderStyle     =   1
      End
      Begin VB.CommandButton cmdBFFTarget 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Browse..."
         Top             =   2025
         Width           =   285
      End
      Begin VB.CommandButton cmdBFFBackup 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Browse..."
         Top             =   1290
         Width           =   285
      End
      Begin VB.TextBox txtTargetFolder 
         Height          =   300
         Left            =   375
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2025
         Width           =   4485
      End
      Begin VB.TextBox txtBackupFolder 
         Height          =   300
         Left            =   375
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1290
         Width           =   4485
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "Create &backup:"
         Height          =   210
         Left            =   375
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2325
      End
      Begin VB.OptionButton optTarget 
         Caption         =   "&Save optimized files to:"
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1710
         Width           =   3000
      End
      Begin VB.OptionButton optTarget 
         Caption         =   "&Modify original files"
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   675
         Width           =   3000
      End
      Begin VB.CommandButton cmdOptimize 
         Caption         =   "&Optimize!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "Preserve"
         ToolTipText     =   "Start optimization"
         Top             =   2595
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "Preserve"
         ToolTipText     =   "Stop optimization"
         Top             =   2595
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblSavedV 
         Height          =   210
         Left            =   1785
         TabIndex        =   18
         Tag             =   "Preserve"
         Top             =   3135
         Width           =   1800
      End
      Begin VB.Label lblSaved 
         Caption         =   "Saved:"
         Height          =   210
         Left            =   1125
         TabIndex        =   17
         Tag             =   "Preserve"
         Top             =   3135
         Width           =   615
      End
      Begin VB.Label lblFileMask 
         Caption         =   "File mask"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2940
         TabIndex        =   4
         Top             =   195
         Width           =   1410
      End
      Begin VB.Label lblOptimizationOptions 
         Caption         =   "Optimization options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   3
         Top             =   195
         Width           =   2250
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total"
         Height          =   210
         Left            =   1125
         TabIndex        =   16
         Tag             =   "Preserve"
         Top             =   2865
         Width           =   765
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Current"
         Height          =   210
         Left            =   1125
         TabIndex        =   15
         Tag             =   "Preserve"
         Top             =   2595
         Width           =   780
      End
      Begin VB.Label lblProgress 
         Caption         =   "Progress:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Tag             =   "Preserve"
         Top             =   2595
         Width           =   840
      End
   End
   Begin VB.PictureBox picBFF 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3210
      Left            =   0
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2565
   End
   Begin MSComctlLib.ImageList ilLV 
      Left            =   7695
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin GObatch.ucGIFViewer ucGIFViewer 
      Height          =   2265
      Left            =   0
      Top             =   4215
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   3995
      BackColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lvReport 
      Height          =   2220
      Left            =   2985
      TabIndex        =   1
      Tag             =   "Preserve"
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3916
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilLV"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   3308
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Original"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Optimized"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saved"
         Object.Width           =   1589
      EndProperty
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8070
      ScaleHeight     =   16
      ScaleMode       =   0  'User
      ScaleWidth      =   16
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "Always on &top"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGIFViewerContextTop 
      Caption         =   "GIFViewerContext"
      Visible         =   0   'False
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "&Autoplay"
         Index           =   0
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "&Background color..."
         Index           =   1
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "&Play"
         Index           =   3
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "&Next frame"
         Index           =   4
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "&Copy image"
         Index           =   6
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuGIFViewerContext 
         Caption         =   "Cancel"
         Index           =   8
      End
   End
   Begin VB.Menu mnuLVReportContextTop 
      Caption         =   "LVReportContext"
      Visible         =   0   'False
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "&Select all"
         Index           =   0
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "&Unselect all"
         Index           =   1
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "&Invert selection"
         Index           =   2
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "Select &optimized"
         Index           =   4
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "S&elect from current"
         Index           =   5
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "U&nselect from current"
         Index           =   6
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "&Refresh list"
         Index           =   8
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "Open..."
         Index           =   10
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuLVReportContext 
         Caption         =   "Cancel"
         Index           =   12
      End
   End
   Begin VB.Menu mnuOptimizationOptionsTop 
      Caption         =   "OptimizationOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Use &Minimum Bounding Rectangle"
         Index           =   0
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Use &Frame Differencing"
         Index           =   1
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Use &best"
         Index           =   2
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Apply second &palette optimization"
         Index           =   4
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Disable &interlacing"
         Index           =   6
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Remove &comments"
         Index           =   7
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuOptimizationOptions 
         Caption         =   "Cancel"
         Index           =   9
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       GO batch [1.0 rev. 09]
' Author:        Carles P.V.
' Last revision: 2004.06.25
'================================================
'
' LOG:
'
'  - 2003.08.29: Fixed paths bug for W98-.
'  - 2003.08.30: File operations now through 'FileSystemObject'
'                instead of 'SHFileOperation API'.
'                Problems under W98- (?).
'  - 2003.09.03: Added: Optimization options:
'                · Optimization method
'                · Second palette optimization
'                · Disable interlacing
'  - 2003.09.04: Added: File masking.
'  - 2003.09.06: Added: 'Use best' optimization option.
'  - 2003.09.07: Added: Preserved last column headers widths.
'  - 2003.09.09: Added: 'Next frame' and 'Copy image' (current frame image)
'  - 2003.09.09: Fixed: 'Scan for redundant disposal modes'!!!.
'  - 2003.09.18: Improved: 'Crop transparent images'.
'  - 2003.09.18: Improved: 'Crop transparent images'.
'  - 2004.06.25: Added: Comments supported.
'
' To do...
'
'  - Support command line input.
'  - Improve error handling.
'  - Allow asynchronous GIF loading/playing (for large GIFs).
'  - Allow manual optimization.
'  - Solve focus issues with BFF treeview
'  - [?]



Option Explicit

Private m_bCancel As Boolean
Private m_mnuRect As RECT2, m_bDwnInRect As Boolean

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Load()
    
    '-- Load app. settings...
    Call mSettings.LoadSettings

    '-- Define popup menu button rect. (optim. options)
    Call mMisc.SetRect(m_mnuRect, 5, 10, 25, 30)
    
    '-- Show form
    Call Me.Show: DoEvents
    Call mMisc.SetFormPos(Me, Me.mnuView(0).Checked)

    '-- Initialize splitters
    Call ucSplitterH.Initialize(Me)
    Call ucSplitterV.Initialize(Me)
    
    '-- Change buttons appearance
    Call mMisc.RemoveButtonBorderEnhance(cmdFileMask)
    Call mMisc.RemoveButtonBorderEnhance(cmdBFFBackup)
    Call mMisc.RemoveButtonBorderEnhance(cmdBFFTarget)
    Call mMisc.RemoveButtonBorderEnhance(cmdOptimize)
    Call mMisc.RemoveButtonBorderEnhance(cmdCancel)
    
    '-- Set BFF-TV container
    mBFF.InitFolder = mSettings.BrowseFolder
    Call mBFF.SetBFFContainer(Me.picBFF)
End Sub

Private Sub Form_Resize()

  Const DMAX As Long = 75

    On Error Resume Next
    
    '-- Resize splitters
    Call ucSplitterH.Move(ucSplitterH.Left, 0, ucSplitterH.Width, ScaleHeight - sbInfo.Height - 3)
    Call ucSplitterV.Move(0, ucSplitterV.Top, ucSplitterH.Left, ucSplitterV.Height)
    
    '-- Update their min/max pos.
    ucSplitterH.xMax = Me.ScaleWidth - DMAX
    ucSplitterH.xMin = DMAX
    ucSplitterV.yMax = Me.ScaleHeight - sbInfo.Height - DMAX
    ucSplitterV.yMin = DMAX
    
    '-- Relocate splitters
    If (WindowState = vbNormal) Then
        If (ucSplitterH.Left > ucSplitterH.xMax) Then ucSplitterH.Left = ucSplitterH.xMax
        If (ucSplitterV.Top > ucSplitterV.yMax) Then ucSplitterV.Top = ucSplitterV.yMax
    End If
    
    '-- Resize controls
    Call picBFF.Move(0, 0, ucSplitterH.Left, ucSplitterV.Top)
    Call ucGIFViewer.Move(0, ucSplitterV.Top + ucSplitterV.Height, ucSplitterH.Left, ScaleHeight - sbInfo.Height - ucSplitterV.Height - picBFF.Height - 4)
    Call lvReport.Move(ucSplitterH.Left + ucSplitterH.Width - 1, -1, ScaleWidth - ucSplitterH.Left - ucSplitterH.Width + 2, ScaleHeight - sbInfo.Height - 230)
    Call picOptions.Move(ucSplitterH.Left + ucSplitterH.Width, lvReport.Top + lvReport.Height, ScaleWidth - ucSplitterH.Left - ucSplitterH.Width, ScaleHeight - sbInfo.Height - lvReport.Height - 3)

    On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '-- Optimizing [?]
    If (cmdCancel.Visible) Then
        Cancel = 1
        Call MsgBox("Unable to quit application now!", vbExclamation, "GO - batch")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Save settings
    BrowseFolder = mBFF.CurrentFolder
    Call mSettings.SaveSettings
    
    '-- Return the Treeview to BrowseForFolder dialog
    '   and close the hidden dialog
    Call mBFF.CloseUpBFF
End Sub

'========================================================================================
' Splitters
'========================================================================================

Private Sub ucSplitterV_Release()
    Call Form_Resize
End Sub

Private Sub ucSplitterH_Release()
    Call Form_Resize
End Sub

'========================================================================================
' BFF Treview
'========================================================================================

Private Sub picBFF_Resize()

    '-- Resize the Treeview as needed
    On Error Resume Next
    Call mBFF.SizeBFF(0, 0, picBFF.ScaleWidth, picBFF.ScaleHeight, -1)
    On Error GoTo 0
End Sub

Public Sub PathChange()
  
  Dim oFSO As New FileSystemObject
    
    '-- Valid folder [?]
    If (Not oFSO.FolderExists(mBFF.CurrentFolder)) Then Exit Sub
    
    '-- Store new folder
    mSettings.BrowseFolder = mBFF.CurrentFolder
    
    '-- Destroy GIF (viewer) and fill listview
    DoEvents
    Screen.MousePointer = vbArrowHourglass
    Call ucGIFViewer.Destroy
    Call mLVReport.FillLV(lvReport, mBFF.CurrentFolder & IIf(Right$(mBFF.CurrentFolder, 1) <> "\", "\", vbNullString), mSettings.FileMask & ".gif")
    Screen.MousePointer = vbNormal
    
    '-- Select first item without loading file
    If (lvReport.ListItems.Count) Then
        lvReport.ListItems(1).Selected = -1
        Call lvReport.SetFocus
    End If
    
    '-- Show search info
    sbInfo.Panels(1).Text = Format$(mLVReport.FilesCount, "#,0 file/s found")
    sbInfo.Panels(2).Text = Format$(mLVReport.FilesSize, "#,0 bytes")
    sbInfo.Panels(3).Text = vbNullString
    lblSavedV.Caption = "0 bytes"
End Sub

'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)

    Select Case Index
        
        Case 0 '-- Exit
            Call Unload(Me)
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Always on top
            With mnuView(0)
                .Checked = Not .Checked: mMisc.SetFormPos Me, .Checked
            End With
    End Select
End Sub

Private Sub mnuGIFViewerContext_Click(Index As Integer)
    
  Dim lRet As Long
  
    Select Case Index
        
        Case 0 '-- Autoplay
            With mnuGIFViewerContext(0)
                .Checked = Not .Checked: ucGIFViewer.AutoPlay = .Checked
            End With
            
        Case 1 '-- Background color...
            lRet = mDialogColor.SelectColor(Me.hWnd, ucGIFViewer.BackColor)
            If (lRet <> -1) Then
                ucGIFViewer.BackColor = lRet
            End If
        
        Case 3 '-- Play/Pause
            With ucGIFViewer
                If (.GIFIsPlaying) Then
                    mnuGIFViewerContext(3).Caption = "&Play": .Pause
                  Else
                    mnuGIFViewerContext(3).Caption = "&Pause": .Play
                End If
            End With
            
        Case 4 '-- Next frame
            Call ucGIFViewer.NextFrame
            
        Case 6 '-- Copy image
            With VB.Clipboard
                Call .Clear
                Call .SetData(ucGIFViewer.GIFCurrentFrameImage, vbCFBitmap)
            End With
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)

    Select Case Index
        
        Case 0 '-- About
            Call MsgBox(vbCrLf & _
                        "GO - batch v" & App.Major & "." & App.Minor & " (rev. " & App.Revision & ")" & vbCrLf & vbCrLf & _
                        "GIF file size optimizer / batch version" & vbCrLf & _
                        "Carles P.V. - © 2003" & vbCrLf & vbCrLf & _
                        "GIF optimization based on TGIFImage by Anders Melander" & vbCrLf & _
                        "GIF decoding based on work by Vlad Vissoultchev" & vbCrLf & _
                        "GIF encoding based on work by Ron van Tilburg" & vbCrLf & vbCrLf, , "About")
    End Select
End Sub

'========================================================================================
' ListView (lvReport)
'========================================================================================

Private Sub lvReport_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If (cmdOptimize.Visible) Then
    
        DoEvents
        Screen.MousePointer = vbArrowHourglass
        
        With ucGIFViewer
            If (.LoadFromFile(mBFF.CurrentFolder & "\" & lvReport.SelectedItem) = 0) Then
                Call ucGIFViewer.Destroy
                lvReport.SelectedItem.SmallIcon = 1
                sbInfo.Panels(3).Text = lvReport.SelectedItem & ": Unexpected error"
              Else
                sbInfo.Panels(3).Text = lvReport.SelectedItem & ": " & .GIFWidth & "x" & .GIFHeight & " [" & .GIFFrames & " frame/s]"
            End If
        End With
        
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub lvReport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (cmdOptimize.Visible) Then
        If (Button = vbRightButton And lvReport.ListItems.Count) Then
            Call Me.PopupMenu(Me.mnuLVReportContextTop, , , , mnuLVReportContext(10))
        End If
    End If
End Sub

Private Sub lvReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If (cmdOptimize.Visible) Then
    
        With lvReport
            If (.ListItems.Count) Then
                If (.SortKey = ColumnHeader.Index - 1) Then
                    If (.SortOrder = lvwAscending) Then
                        .SortOrder = lvwDescending
                      Else
                        .SortOrder = lvwAscending
                    End If
                  Else
                    .SortKey = ColumnHeader.Index - 1
                    .SortOrder = lvwAscending
                End If
                .Sorted = -1
                .SelectedItem.EnsureVisible
            End If
        End With
    End If
End Sub

Private Sub mnuLVReportContext_Click(Index As Integer)
    
  Dim lItm As Long
  Dim lRet As Long
    
    Select Case Index
        
        Case 0  '-- Select all
            For lItm = 1 To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = -1
            Next lItm
        
        Case 1  '-- Unselect all
            For lItm = 1 To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = 0
            Next lItm
    
        Case 2  '-- Invert selection
            For lItm = 1 To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = Not lvReport.ListItems(lItm).Checked
            Next lItm
            
        Case 4  '-- Select optimized
            For lItm = 1 To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = (lvReport.ListItems(lItm).SubItems(3) <> Chr$(45))
            Next lItm
            
        Case 5  '-- Select from current
            For lItm = lvReport.SelectedItem.Index To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = -1
            Next lItm
            
        Case 6  '-- Unselect from current
            For lItm = lvReport.SelectedItem.Index To lvReport.ListItems.Count
                lvReport.ListItems(lItm).Checked = 0
            Next lItm
            
        Case 8  '-- Refresh (force path change)
            Call Me.PathChange
            
        Case 10 '-- Open...
            lRet = mLVReport.ShellOpenFile(mSettings.BrowseFolder, lvReport.SelectedItem, Me.hWnd)
            If (lRet <= 32) Then
                Call MsgBox("Unexpected error opening " & lvReport.SelectedItem, vbExclamation, "GO - batch")
            End If
    End Select
End Sub

'========================================================================================
' GIF viewer (ucGIFViewer)
'========================================================================================

Private Sub ucGIFViewer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (Button = vbRightButton) Then
        
        With ucGIFViewer
            Me.mnuGIFViewerContext(3).Caption = IIf(.GIFIsPlaying, "&Pause", "&Play")
            Me.mnuGIFViewerContext(3).Enabled = (.GIFLoaded And .GIFFrames > 1)
            Me.mnuGIFViewerContext(4).Enabled = (.GIFLoaded And .GIFFrames > 1 And Not .GIFIsPlaying)
            Me.mnuGIFViewerContext(6).Enabled = (.GIFLoaded And Not .GIFIsPlaying)
        End With
        Call Me.PopupMenu(Me.mnuGIFViewerContextTop)
    End If
End Sub

'========================================================================================
' Selecting Backup and Target folders
'========================================================================================

Private Sub cmdBFFBackup_Click()

  Dim sRet As String

    '-- Browse for Backup folder
    mBFF.InitFolder = mSettings.BackupFolder
    sRet = mBFF.BrowseForFolder(Me.hWnd, "Select Backup folder...")
    If (sRet <> vbNullString) Then
        mSettings.BackupFolder = sRet
        txtBackupFolder.Text = sRet
    End If
End Sub

Private Sub cmdBFFTarget_Click()

  Dim sRet As String
    
    '-- Browse for Target folder
    mBFF.InitFolder = mSettings.TargetFolder
    sRet = mBFF.BrowseForFolder(Me.hWnd, "Select Target folder...")
    If (sRet <> vbNullString) Then
        mSettings.TargetFolder = sRet
        txtTargetFolder.Text = sRet
    End If
End Sub

'========================================================================================
' Optimization options
'========================================================================================

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (cmdOptimize.Visible) Then
        If (Button = vbLeftButton And m_bDwnInRect) Then
            Call mMisc.DrawPopupButton(picOptions.hDC, m_mnuRect, IIf(mMisc.PtInRect(m_mnuRect, x, y), -1, 0))
        End If
    End If
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (cmdOptimize.Visible) Then
        If (mMisc.PtInRect(m_mnuRect, x, y) And Button = vbLeftButton) Then
            mMisc.DrawPopupButton picOptions.hDC, m_mnuRect, -1
            m_bDwnInRect = -1
          Else
            m_bDwnInRect = 0
        End If
    End If
End Sub

Private Sub picOptions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (cmdOptimize.Visible) Then
        mMisc.DrawPopupButton picOptions.hDC, m_mnuRect, 0
        If (m_bDwnInRect And mMisc.PtInRect(m_mnuRect, x, y) And Button = vbLeftButton) Then
            Call Me.PopupMenu(Me.mnuOptimizationOptionsTop, , picOptions.Left + m_mnuRect.x1, picOptions.Top + m_mnuRect.y2)
        End If
    End If
End Sub

Private Sub picOptions_Paint()
   Call mMisc.DrawPopupButton(picOptions.hDC, m_mnuRect, 0, Not cmdOptimize.Visible)
End Sub

'//

Private Sub mnuOptimizationOptions_Click(Index As Integer)
    
  Dim nItm As Integer
  
    Select Case Index
    
        '-- Optimization method
        Case 0, 1, 2
            For nItm = 0 To 2
                mnuOptimizationOptions(nItm).Checked = 0
            Next nItm
            mnuOptimizationOptions(Index).Checked = -1
            
        '-- Second palette optimization / Disable interlacing / Remove comments
        Case 4, 6, 7
            mnuOptimizationOptions(Index).Checked = Not mnuOptimizationOptions(Index).Checked
    End Select
End Sub

'========================================================================================
' File mask
'========================================================================================

Private Sub txtFileMask_Change()

    '-- Update file mask
    mSettings.FileMask = Trim$(txtFileMask.Text)
End Sub

Private Sub txtFileMask_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- Force mask change
    cmdFileMask = (KeyCode = vbKeyReturn)
End Sub

Private Sub cmdFileMask_Click()
    
    '-- Force path change
    Call Me.PathChange
End Sub

'========================================================================================
' Optimize! / Cancel
'========================================================================================

Private Sub cmdOptimize_Click()
    
  Dim lItm As Long
  Dim oGIF As New cGIF
  Dim nFrm As Integer
  
  Dim oFSO      As New FileSystemObject
  Dim oFile     As File
  Dim sCurGIF   As String
  Dim sBUpGIF   As String
  Dim sTrgGIF   As String
  Dim sTmpGIF   As String
  Dim sTmpGIF_0 As String
  Dim sTmpGIF_1 As String
  Dim lLenGIF_0 As Long
  Dim lLenGIF_1 As Long
  
  Dim bSkip     As Boolean
  Dim bRet      As Boolean
  Dim lSaved    As Long
      
    '-- Empty list [?]
    If (lvReport.ListItems.Count = 0) Then Exit Sub
    
    '-- Check if folders exist
    If (optTarget(0) And chkBackup) Then
        If (Not oFSO.FolderExists(mSettings.BackupFolder)) Then
            Call MsgBox("Backup folder does not exist. Please, select a valid folder.", vbExclamation, "GO - batch")
            Exit Sub
        End If
    End If
    If (optTarget(1)) Then
        If (Not oFSO.FolderExists(mSettings.TargetFolder)) Then
            Call MsgBox("Target folder does not exist. Please, select a valid folder.", vbExclamation, "GO - batch")
            Exit Sub
        End If
    End If
    
    '-- Check if source and target folders are not the same folder
    If (optTarget(0) And chkBackup) Then
        If (mSettings.BackupFolder = mSettings.BrowseFolder) Then
            Call MsgBox("Backup folder can not be current source folder.", vbExclamation, "GO - batch")
            Exit Sub
        End If
    End If
    If (optTarget(1)) Then
        If (mSettings.TargetFolder = mSettings.BrowseFolder) Then
            Call MsgBox("Target folder can not be current source folder.", vbExclamation, "GO - batch")
            Exit Sub
        End If
    End If
    
    '-- Temporary full GIF paths
    sTmpGIF_0 = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "oGIF_0.tmp"
    sTmpGIF_1 = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "oGIF_1.tmp"
    
    '//
    
    '-- Set focus on ListView, set progress max. and resed 'Saved' label
    Call lvReport.SetFocus
    ucProgress.Max = mLVReport.FilesCount
    lblSavedV.Caption = "0 bytes"
    
    '-- Starting optimization...
    Call pvSetControlsStatus(-1)
    cmdOptimize.Visible = 0
    cmdCancel.Visible = -1
    
    '//
    
    On Error Resume Next
    
    For lItm = 1 To ucProgress.Max
        
        '-- Progress (Total)
        ucProgress = lItm
        
        '-- Load GIF [?]...
        If (lvReport.ListItems(lItm).Checked) Then
            
            With lvReport.ListItems(lItm)
            
                '-- Select current item
                .Selected = -1
                .EnsureVisible
            
                '-- Get source/target GIF paths
                sCurGIF = mSettings.BrowseFolder & IIf(Right$(mSettings.BrowseFolder, 1) <> "\", "\", vbNullString) & .Text
                sBUpGIF = mSettings.BackupFolder & IIf(Right$(mSettings.BackupFolder, 1) <> "\", "\", vbNullString) & .Text
                sTrgGIF = mSettings.TargetFolder & IIf(Right$(mSettings.TargetFolder, 1) <> "\", "\", vbNullString) & .Text
            End With
            
            DoEvents: bSkip = 0
            
            If (oGIF.LoadFromFile(sCurGIF)) Then
                
                Select Case True
                
                    Case mnuOptimizationOptions(0).Checked '-- M.B.R.
                        
                        bRet = pvSaveOptimizedTmpGIF(oGIF, sTmpGIF_0, [fomMinimumBoundingRectangle], mnuOptimizationOptions(4).Checked, mnuOptimizationOptions(6).Checked, mnuOptimizationOptions(7).Checked, ucProgressCurrent)
                        lLenGIF_0 = FileLen(sTmpGIF_0)
                        
                        If (lLenGIF_0 < FileLen(sCurGIF)) Then
                            sTmpGIF = sTmpGIF_0
                          Else
                            bSkip = -1
                        End If
                        
                    Case mnuOptimizationOptions(1).Checked '-- F.D.
                        
                        bRet = pvSaveOptimizedTmpGIF(oGIF, sTmpGIF_1, [fomFrameDifferencing], mnuOptimizationOptions(4).Checked, mnuOptimizationOptions(6).Checked, mnuOptimizationOptions(7).Checked, ucProgressCurrent)
                        lLenGIF_1 = FileLen(sTmpGIF_1)
                        
                        If (lLenGIF_1 < FileLen(sCurGIF)) Then
                            sTmpGIF = sTmpGIF_1
                          Else
                            bSkip = -1
                        End If
                    
                    Case mnuOptimizationOptions(2).Checked '-- Best
                        
                        bRet = pvSaveOptimizedTmpGIF(oGIF, sTmpGIF_0, [fomMinimumBoundingRectangle], mnuOptimizationOptions(4).Checked, mnuOptimizationOptions(6).Checked, mnuOptimizationOptions(7).Checked, ucProgressCurrent)
                        
                        If (bRet) Then
                            If (oGIF.LoadFromFile(sCurGIF)) Then
                                
                                bRet = pvSaveOptimizedTmpGIF(oGIF, sTmpGIF_1, [fomFrameDifferencing], mnuOptimizationOptions(4).Checked, mnuOptimizationOptions(6).Checked, mnuOptimizationOptions(7).Checked, ucProgressCurrent)
                                lLenGIF_0 = FileLen(sTmpGIF_0)
                                lLenGIF_1 = FileLen(sTmpGIF_1)
                                
                                Select Case True
                                
                                    Case lLenGIF_0 <= lLenGIF_1
                                        If (lLenGIF_0 < FileLen(sCurGIF)) Then
                                            sTmpGIF = sTmpGIF_0
                                          Else
                                            bSkip = -1
                                        End If
                                        
                                    Case lLenGIF_0 > lLenGIF_1
                                        If (lLenGIF_1 < FileLen(sCurGIF)) Then
                                            sTmpGIF = sTmpGIF_1
                                          Else
                                            bSkip = -1
                                        End If
                                End Select
                            End If
                        End If
                End Select
                
                '//
                
                If (bRet) Then
                
                    Select Case True
                        
                        Case optTarget(0) '-- Overwrite original + Backup [?]
                        
                            If (chkBackup) Then
                                Set oFile = oFSO.GetFile(sCurGIF): oFile.Copy sBUpGIF
                            End If
                            If (Not bSkip) Then
                                Set oFile = oFSO.GetFile(sTmpGIF): oFile.Copy sCurGIF
                            End If
                            
                        Case optTarget(1) '-- Custom target folder
                        
                            If (Not bSkip) Then
                                Set oFile = oFSO.GetFile(sTmpGIF): oFile.Copy sTrgGIF
                            End If
                    End Select
                    
                    '-- Show new optimized file size and % saved
                    With lvReport.ListItems(lItm)
                    
                        If (Not bSkip) Then
                        
                            '-- Show new optimized file size and saved bytes
                            .SubItems(2) = Format$(FileLen(sTmpGIF), "@@@@@@@@@@")
                            .SubItems(3) = Format$(Format$(1 - Val(.SubItems(2)) / Val(.SubItems(1)), "0.00%"), "@@@@@@@@")
                            '-- Show current total bytes saved
                            lSaved = lSaved + Val(.SubItems(1)) - Val(.SubItems(2))
                            lblSavedV.Caption = Format$(lSaved, "#,# bytes")
                          
                          Else
                            '-- Skipped
                            .SubItems(2) = Format$("Skipped", "@@@@@@@@@@")
                            .SubItems(3) = Chr$(45)
                        End If
                    End With
                    
                  Else
                    lvReport.ListItems(lItm).SmallIcon = 1
                End If
              
              Else
                lvReport.ListItems(lItm).SmallIcon = 1
            End If
        End If
        If (m_bCancel) Then Exit For

ErrH: Next lItm
On Error GoTo 0
    
    '-- Restore controls status
    Call pvSetControlsStatus(0)
    cmdOptimize.Visible = -1
    cmdCancel.Visible = 0: m_bCancel = 0
    
    '-- Kill temp. files
    On Error Resume Next
    Set oFile = oFSO.GetFile(sTmpGIF_0): oFile.Delete
    Set oFile = oFSO.GetFile(sTmpGIF_1): oFile.Delete
    On Error GoTo 0
    
    '-- End
    ucProgressCurrent = 0
    ucProgress = 0
End Sub

Private Sub cmdCancel_Click()
    m_bCancel = -1
End Sub

'//

Private Sub pvSetControlsStatus(ByVal bOptimizing As Boolean)
    
  Dim oCtl As Control
    
    '-- 'Real' controls
    On Error Resume Next
    For Each oCtl In Controls
        If (oCtl.Tag <> "Preserve") Then
            oCtl.Enabled = Not bOptimizing
        End If
    Next oCtl
    On Error GoTo 0
    
    '-- 'Owner drawn' controls
    mMisc.DrawPopupButton picOptions.hDC, m_mnuRect, 0, bOptimizing
End Sub

Private Function pvSaveOptimizedTmpGIF( _
                 oGIF As cGIF, _
                 ByVal sPath As String, _
                 ByVal eOM As eFrameOptimizationMethod, _
                 ByVal bSecondPaletteOptimization As Boolean, _
                 ByVal bDisableInterlacing As Boolean, _
                 ByVal bRemoveComments As Boolean, _
                 oProgress As ucProgress) As Boolean
                 
  Dim nFrm As Integer
  
    On Error GoTo ErrH
    
    '-- Pre-optimization checking
    Call mGIFOptimizer.CheckGIF(oGIF, oProgress)
    
    '-- Optimize Global palette
    Call mGIFOptimizer.OptimizeGlobalPalette(oGIF, oProgress)
    
    '-- Optimize Local palette/s
    ucProgressCurrent.Max = 0
    For nFrm = 1 To oGIF.FramesCount
        If (oGIF.LocalPaletteUsed(nFrm)) Then oProgress.Max = oProgress.Max + 1
    Next nFrm
    For nFrm = 1 To oGIF.FramesCount
        oProgress = nFrm
        Call mGIFOptimizer.OptimizeLocalPalette(oGIF, nFrm)
    Next nFrm
    
    '-- Remove Local palette/s
    If (oGIF.FramesCount > 1) Then
        Call mGIFOptimizer.RemoveLocalPalettes(oGIF, oProgress)
    End If
    
    '-- Optimize frames
    If (oGIF.FramesCount > 1) Then
        Call mGIFOptimizer.OptimizeFrames(oGIF, eOM, oProgress)
    End If
    
    '-- Crop transparent images
    Call mGIFOptimizer.CropTransparentImages(oGIF, oProgress)
    
    '-- Second palette optimization
    If (bSecondPaletteOptimization) Then
    
        If (oGIF.FramesCount > 1) Then
        
            '-- Optimize Global palette
            Call mGIFOptimizer.OptimizeGlobalPalette(oGIF, oProgress)
            
            '-- Optimize Local palette/s
            ucProgressCurrent.Max = 0
            For nFrm = 1 To oGIF.FramesCount
                If (oGIF.LocalPaletteUsed(nFrm)) Then oProgress.Max = oProgress.Max + 1
            Next nFrm
            For nFrm = 1 To oGIF.FramesCount
                oProgress = nFrm
                Call mGIFOptimizer.OptimizeLocalPalette(oGIF, nFrm)
            Next nFrm
            
            '-- Remove Local palette/s (2nd pass)
            Call mGIFOptimizer.RemoveLocalPalettes(oGIF, oProgress)
        End If
    End If
    
    '-- Disable interlacing
    If (bDisableInterlacing) Then
    
        For nFrm = 1 To oGIF.FramesCount
            oGIF.FrameInterlaced(nFrm) = 0
        Next nFrm
    End If
    
    '-- Remove comments
    If (bRemoveComments) Then
        
        For nFrm = 1 To oGIF.FramesCount
            Call oGIF.FrameCommentRemoveAll(nFrm)
        Next nFrm
        Call oGIF.TrailingCommentRemoveAll
    End If
    
    '-- Save optimized GIF
    pvSaveOptimizedTmpGIF = oGIF.Save(sPath)
    
ErrH: On Error GoTo 0
End Function
