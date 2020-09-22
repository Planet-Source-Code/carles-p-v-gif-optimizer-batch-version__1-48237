VERSION 5.00
Begin VB.UserControl ucGIFViewer 
   CanGetFocus     =   0   'False
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   ClipControls    =   0   'False
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ucGIFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User Control:  ucGIFViewer.ctl
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
'                cDIB.cls
'                mGIFLZWDec.bas
' Last revision: 2003.09.09
'================================================
'
' Notes:
' - Infinite looping.
'
' LOG:
'
' 2003.06.27: - Added NextFrame() method.
'             - Added GIFIsPlaying() property.
' 2003.06.30: - NextFrame() method only available when GIF not playing.
' 2003.08.14: - Added best fit functionality.
' 2003.08.20: - Improved paint routines.
' 2003.09.09: - Added GIFCurrentFrameImage property (read only).



Option Explicit

'-- API:

Private Const RGN_DIFF As Long = 4

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'//

'-- Public enums.:
Public Enum ucBorderStyleConstants
    [None] = 0
    [Fixed Single]
End Enum

'-- Private Constants:
Private Const m_def_AutoPlay As Boolean = 0

'-- Property Variables:
Private m_AutoPlay As Boolean

'-- Event Declarations:
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event GIFFrameRendered(ByVal Frame As Integer)

'-- Private Variables:
Private m_oGIF            As cGIF
Attribute m_oGIF.VB_VarHelpID = -1
Private m_Frames          As Integer
Private m_Frame           As Integer
Private m_FrameBuffDIB    As New cDIB
Private m_RestoringDIB    As New cDIB
Private m_BackgroundDIB   As New cDIB
Private m_xOffset         As Long
Private m_yOffset         As Long
Private m_W               As Long
Private m_H               As Long



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Create GIF object
    Set m_oGIF = New cGIF
End Sub

Private Sub UserControl_Terminate()

    '-- Stop timer
    tmrDelay.Enabled = 0
    '-- Destroy GIF object and buffers
    Set m_oGIF = Nothing
    Set m_FrameBuffDIB = Nothing
    Set m_RestoringDIB = Nothing
    Set m_BackgroundDIB = Nothing
End Sub

Private Sub UserControl_Resize()
       
    '-- Screen offsets/bounds
    Call m_FrameBuffDIB.GetBestFitInfo(ScaleWidth, ScaleHeight, m_xOffset, m_yOffset, m_W, m_H)
    '-- Erase background and refresh
    Call pvEraseBackground
    Call UserControl_Paint
End Sub

Private Sub UserControl_Paint()

    '-- Paint current rendered frame
    If (m_FrameBuffDIB.hDIB <> 0) Then
        Call m_FrameBuffDIB.Stretch(hDC, m_xOffset, m_yOffset, m_W, m_H)  ', , , , , , [smHalftone]
    End If
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function LoadFromFile(ByVal Filename As String) As Boolean
     
    '-- Stop animation, and reset frames count
    tmrDelay.Enabled = 0: m_Frames = 0
    
    '-- Load from file...
    If (m_oGIF.LoadFromFile(Filename)) Then
    
        '-- Get number of frames/ Success [?]
        m_Frames = m_oGIF.FramesCount
        LoadFromFile = (m_Frames > 0)
    End If
    Call pvStart(m_AutoPlay)
End Function

Public Function LoadFromResource(ByVal ResourceID As String, ByVal ResourceType As String) As Boolean
    
    '-- Stop animation, and reset frames count
    tmrDelay.Enabled = 0: m_Frames = 0
    
    '-- Load from resource...
    If (m_oGIF.LoadFromStream(LoadResData(ResourceID, ResourceType))) Then
     
        '-- Get number of frames/ Success [?]
        m_Frames = m_oGIF.FramesCount
        LoadFromResource = (m_Frames > 0)
    End If
    Call pvStart(m_AutoPlay)
End Function

'//

Public Sub Play()
    '-- Start/Continue animation
    tmrDelay.Enabled = (m_Frames > 1)
End Sub

Public Sub Pause()
    '-- Pause animation
    tmrDelay.Enabled = 0
End Sub

Public Sub Rewind()
    '-- Goto first frame
    m_Frame = 1
End Sub

Public Sub NextFrame()
    '-- Render next frame
    If (m_Frames > 0 And Not tmrDelay.Enabled) Then Call tmrDelay_Timer
End Sub

Public Sub Destroy()

    '-- Stop timer
    tmrDelay.Enabled = 0: m_Frames = 0
    '-- Destroy GIF object and buffers
    Call m_oGIF.Destroy
    Call m_FrameBuffDIB.Destroy
    Call m_RestoringDIB.Destroy
    Call m_BackgroundDIB.Destroy
    
    '-- Erase background
    m_W = 0: m_H = 0: Call pvEraseBackground
End Sub

'========================================================================================
' UseControl events
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get GIFLoaded() As Boolean
    GIFLoaded = (m_Frames > 0)
End Property

Public Property Get GIFWidth() As Integer
    GIFWidth = m_oGIF.ScreenWidth
End Property

Public Property Get GIFHeight() As Integer
    GIFHeight = m_oGIF.ScreenHeight
End Property

Public Property Get GIFFrames() As Integer
    GIFFrames = m_Frames
End Property

Public Property Get GIFCurrentFrame() As Integer
    GIFCurrentFrame = m_Frame
End Property

Public Property Get GIFCurrentFrameImage() As StdPicture
    Set GIFCurrentFrameImage = m_FrameBuffDIB.Image
End Property

Public Property Get GIFIsPlaying() As Boolean
    GIFIsPlaying = tmrDelay.Enabled
End Property

'//

Public Property Get AutoPlay() As Boolean
    AutoPlay = m_AutoPlay
End Property
Public Property Let AutoPlay(ByVal New_AutoPlay As Boolean)
    m_AutoPlay = New_AutoPlay
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."

  Dim bPlaying As Boolean
 
    UserControl.BackColor() = New_BackColor
    
    bPlaying = tmrDelay.Enabled
    tmrDelay.Enabled = 0
    Call pvStart(bPlaying)
End Property

Public Property Get BorderStyle() As ucBorderStyleConstants
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As ucBorderStyleConstants)
    UserControl.BorderStyle() = New_BorderStyle
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

'//

Private Sub UserControl_InitProperties()
    m_AutoPlay = m_def_AutoPlay
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_AutoPlay = .ReadProperty("AutoPlay", m_def_AutoPlay)
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", 0)
        UserControl.Enabled = .ReadProperty("Enabled", -1)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("AutoPlay", m_AutoPlay, m_def_AutoPlay)
        Call .WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
        Call .WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
        Call .WriteProperty("Enabled", UserControl.Enabled, -1)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub tmrDelay_Timer()
    
    '-- Render current frame
    Call pvRenderFrame
    
    '-- Next frame / First
    If (m_Frame < m_Frames) Then
        m_Frame = m_Frame + 1
      Else
        m_Frame = 1
    End If
    
    '-- Paint frame
    Call UserControl_Paint
End Sub

Private Sub pvRenderFrame()

    With m_oGIF
    
        '-- Set current frame delay
        Select Case .FrameDelay(m_Frame)
            Case Is < 0
                tmrDelay.Interval = 60000 ' Max.: 1 min.
            Case Is = 0
                tmrDelay.Interval = 100   ' Def.: 0.1 sec.
            Case Is < 5
                tmrDelay.Interval = 50    ' Min.: 0.05 sec.
            Case Else
                tmrDelay.Interval = .FrameDelay(m_Frame) * 10
        End Select
        
        '-- Restore:
        If (m_Frame = 1) Then
            Call m_RestoringDIB.Cls(UserControl.BackColor)
        End If
        Call m_FrameBuffDIB.LoadBlt(m_RestoringDIB.hDC)
        
        '-- Draw current frame:
        Call .FrameDraw(m_FrameBuffDIB.hDC, m_Frame)
        
        '-- Update restoring buffer:
        Select Case .FrameDisposalMethod(m_Frame)
            Case [dmNotSpecified], [dmDoNotDispose]
                '-- Update from current
                Call m_RestoringDIB.LoadBlt(m_FrameBuffDIB.hDC)
            Case [dmRestoreToBackground]
                '-- Update from background
                Call m_BackgroundDIB.Stretch(m_RestoringDIB.hDC, .FrameLeft(m_Frame), .FrameTop(m_Frame), .FrameDIBXOR(m_Frame).Width, .FrameDIBXOR(m_Frame).Height)
            Case [dmRestoreToPrevious]
                '-- Preserve buffer
        End Select
    End With
    
    '-- Raise event
    RaiseEvent GIFFrameRendered(m_Frame)
End Sub

Private Sub pvInitializePreviewBuffers()

    With m_oGIF
                
        '-- Create buffer DIBs
        Call m_FrameBuffDIB.Create(.ScreenWidth, .ScreenHeight, [32_bpp])
        Call m_RestoringDIB.Create(.ScreenWidth, .ScreenHeight, [32_bpp])
        Call m_BackgroundDIB.Create(.ScreenWidth, .ScreenHeight, [32_bpp])
        
        '-- Initialize them
        Call m_FrameBuffDIB.Cls(UserControl.BackColor)
        Call m_RestoringDIB.Cls(UserControl.BackColor)
        Call m_BackgroundDIB.Cls(UserControl.BackColor)
       
        '-- Screen offsets/bounds
        Call m_FrameBuffDIB.GetBestFitInfo(ScaleWidth, ScaleHeight, m_xOffset, m_yOffset, m_W, m_H)
    End With
End Sub

Private Sub pvStart(Optional ByVal bPlay As Boolean = 0)
    
    If (m_Frames > 0) Then
    
        '-- Initialize rendering buffers
        Call pvInitializePreviewBuffers
        m_Frame = 1
    
        '-- Erase background
        Call pvEraseBackground
    
        '-- Enable timer [?]
        If (m_Frames > 1 And bPlay) Then
            tmrDelay.Interval = 1
            tmrDelay.Enabled = -1
          Else
            Call tmrDelay_Timer
        End If
    End If
End Sub

Private Sub pvEraseBackground()
  
  Dim lClr   As Long
  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    Call OleTranslateColor(UserControl.BackColor, 0, lClr)
    hBrush = CreateSolidBrush(lClr)
    
    '-- Create Cls region
    hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    hRgn_2 = CreateRectRgn(m_xOffset, m_yOffset, m_xOffset + m_W, m_yOffset + m_H)
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub
