VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cGIF.cls
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                mGIFLZWDec.bas
'                mGIFLZWEnc.bas
' Last revision: 2003.12.16
'================================================
'
' LOG:
'
' - 2003.08.20: · Fixed 'Out of stream' bug.
' - 2003.12.04: · Fixed 'Unexpected block skiping'.
' - 2003.12.16: · Added comment/s support. (*)
' - 2003.12.19: · Added FrameClone() method.
' - 2003.12.22: · Added FrameImage() property.
'
' (*) Multiple trailing large-comments and
'     multiple frame large-comments.

Option Explicit

'-- API:

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

'-- Private Constants:

Private Const GIF_SIGNATURE                 As String = "GIF"
Private Const GIF_VERSION_87a               As String = "87a"
Private Const GIF_VERSION_89a               As String = "89a"
Private Const GIF_NETSCAPE_ID               As String = "NETSCAPE2.0"
Private Const GIF_TRAILER                   As Byte = &H3B

Private Const GIF_EXTENSION_INTRODUCER      As Byte = &H21
Private Const GIF_APPLICATION_EXTENSION     As Byte = &HFF
Private Const GIF_GRAPHIC_CONTROL_EXTENSION As Byte = &HF9
Private Const GIF_COMMENT_EXTENSION         As Byte = &HFE
Private Const GIF_IMAGE_DESCRIPTOR          As Byte = &H2C
Private Const GIF_BLOCK_TERMINATOR          As Byte = &H0

'-- Private Types:

Private Type tGIFScreenDescriptor
    sdLogicalScreenWidth     As Integer
    sdLogicalScreenHeight    As Integer
    sdFlags                  As Byte
    sdBackgroundColorIndex   As Byte
    sdPixelAspectRatio       As Byte
End Type

Private Type tGIFImageDescriptor
    idImageLeftPosition      As Integer
    idImageTopPosition       As Integer
    idImageWidth             As Integer
    idImageHeight            As Integer
    idFlags                  As Byte
End Type

Private Type tGIFGraphicControlExtension
    gceBlockSize             As Byte ' = &H4
    gceFlags                 As Byte
    gceDelayTime             As Integer
    gceTransparentColorIndex As Byte
End Type

Private Type tGIFApplicationExtension
    aeBlockSize              As Byte ' = &HB
    aeApplicationIdentifier  As String * 8
    aeAuthenticationCode     As String * 3
End Type

Private Type tGIFNetscapeApplicationExtension
    naeBlockSize             As Byte ' = &HB
    naeNestcapeID            As String * 11
    naeSubBlockSize          As Byte ' = &H3
    naeFlags                 As Byte ' = &H1
    naeLoops                 As Integer
End Type

'//

Private Type tFrame
    frDIBXOR                 As cDIB
    frDIBAND                 As cDIB
    frLeft                   As Integer
    frTop                    As Integer
    frInterlaced             As Boolean
    frUseTransparentColor    As Boolean
    frTransparentColorIndex  As Integer
    frDelay                  As Integer
    frDisposalMethod         As Byte
    frUserInput              As Boolean
    frLocalPaletteUsed       As Boolean
    frLocalPaletteSorted     As Boolean
    frLocalPaletteEntries    As Integer
    frLocalPalette(1023)     As Byte
    frComment()              As String
End Type

'-- Public Enums.:

Public Enum GIFDisposalMethod
    [dmNotSpecified]
    [dmDoNotDispose]
    [dmRestoreToBackground]
    [dmRestoreToPrevious]
End Enum

'-- Private Enums.:

Private Enum GIFflags
    '-- Animation flags
    gfGlobalPaletteExists = &H80
    gfGlobalPaletteBpp = &H70
    gfGlobalPaletteSorted = &H8
    gfGlobalPaletteSize = &H7
    '-- Image descriptor flags
    gfLocalPaletteExists = &H80
    gfLocalPaletteSorted = &H20
    gfLocalPaletteSize = &H7
    gfInterlaced = &H40
    '-- Graphic control flags
    gfDisposalMethod = &H1C
    gfUserInput = &H2
    gfUseTransparentColor = &H1
End Enum

'-- Private Property Variables:

Private m_Signature            As String * 3
Private m_Version              As String * 3
Private m_ScreenDescriptor     As tGIFScreenDescriptor
Private m_Loops                As Integer

Private m_GlobalPaletteExists  As Boolean
Private m_GlobalPaletteSorted  As Boolean
Private m_GlobalPaletteEntries As Integer
Private m_GlobalPalette(1023)  As Byte

Private m_TrailingComment()    As String

Private m_FramesCount          As Integer
Private m_Frame()              As tFrame



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()

    '-- Initialize mGIFLZWDec/Enc ext. modules
    Call mGIFLZWDec.InitPowers
    Call mGIFLZWEnc.InitMasks
End Sub

Private Sub Class_Terminate()

    '-- Destroy current GIF
    Call Me.Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function LoadFromFile(ByVal Filename As String) As Boolean

  Dim hFile     As Long
  Dim aStream() As Byte
  
    On Error GoTo ErrH
    
    If (FileLen(Filename)) Then
    
        '-- Open file...
        hFile = FreeFile
        Open Filename For Binary Access Read As #hFile
        
        '-- Get data...
        ReDim aStream(0 To FileLen(Filename) - 1)
        Get #hFile, , aStream()
        
        '-- Close file
        Close #hFile
        
        '-- Load (build) from loaded stream
        LoadFromFile = LoadFromStream(aStream())
    End If
    Exit Function
    
ErrH:
    Call Me.Destroy
End Function

Public Function LoadFromStream(Stream() As Byte) As Boolean
  
  Dim tApplicationExtension    As tGIFApplicationExtension
  Dim tGraphicControlExtension As tGIFGraphicControlExtension
  Dim tImageDescriptor         As tGIFImageDescriptor
  
  Dim lpPos        As Long
  Dim lpEnt        As Long
  Dim nImg         As Integer
  Dim aLZWCodeSize As Byte
  Dim aData()      As Byte
  
  Dim nCmnt        As Integer
  Dim sComments()  As String
    
    '-- Destroy previous GIF
    Call Me.Destroy

    '-- Check signature
    Call CopyMemory(ByVal m_Signature, Stream(0), 3)
    If (m_Signature <> GIF_SIGNATURE) Then
        Exit Function
    End If
    
    '-- Check version
    Call CopyMemory(ByVal m_Version, Stream(3), 3)
    If (m_Version <> GIF_VERSION_87a And m_Version <> GIF_VERSION_89a) Then
        Exit Function
    End If
    
    On Error Resume Next

    '-- Go to Screen Descriptor pos. and extract it
    lpPos = lpPos + 6
    Call CopyMemory(m_ScreenDescriptor, Stream(lpPos), Len(m_ScreenDescriptor))
    
    '-- Go to Global Palette pos. and extract it [?]
    lpPos = lpPos + Len(m_ScreenDescriptor)
    
    With m_ScreenDescriptor
    
        '-- Does exist [?]
        m_GlobalPaletteExists = -(pvGetFlag(.sdFlags, gfGlobalPaletteExists))
        If (m_GlobalPaletteExists) Then
        
            '-- Get palette flags
            m_GlobalPaletteSorted = -(pvGetFlag(.sdFlags, gfGlobalPaletteSorted))
            m_GlobalPaletteEntries = 2 ^ (pvGetFlag(.sdFlags, gfGlobalPaletteSize) + 1)
            
            '-- Get palette
            For lpEnt = 0 To m_GlobalPaletteEntries - 1
                m_GlobalPalette(4 * lpEnt + 0) = Stream(lpPos + 3 * lpEnt + 2)
                m_GlobalPalette(4 * lpEnt + 1) = Stream(lpPos + 3 * lpEnt + 1)
                m_GlobalPalette(4 * lpEnt + 2) = Stream(lpPos + 3 * lpEnt + 0)
            Next lpEnt
            lpPos = lpPos + m_GlobalPaletteEntries * 3
        End If
    End With
    
    '-- Initialize temp. vars.
    ReDim sComments(0)

    '-- Start extracting blocks...
    Do: Select Case Stream(lpPos)
        
            Case GIF_EXTENSION_INTRODUCER
                
                '-- Move forward one byte (Extension label)
                lpPos = lpPos + 1
            
                '-- Get data
                Select Case Stream(lpPos)
                
                    Case GIF_APPLICATION_EXTENSION
                 
                        '-- Move forward one byte (Extension label)
                        lpPos = lpPos + 1
                        
                        '-- Get structure
                        Call CopyMemory(tApplicationExtension, Stream(lpPos), Len(tApplicationExtension))
                        lpPos = lpPos + Len(tApplicationExtension)
                        
                        '-- Get data sub-block/s
                        Call pvGetDataSubBlocks(Stream(), lpPos, aData())
                        
                        '-- NETSCAPE 2.0 application extension [?]
                        With tApplicationExtension
                            If (.aeApplicationIdentifier & .aeAuthenticationCode = GIF_NETSCAPE_ID) Then
                                Call CopyMemory(m_Loops, aData(2), 2)
                            End If
                        End With
                        
                    Case GIF_GRAPHIC_CONTROL_EXTENSION
                        
                        '-- Move forward one byte (Extension label)
                        lpPos = lpPos + 1
                        
                        '-- Get structure
                        Call CopyMemory(tGraphicControlExtension, Stream(lpPos), Len(tGraphicControlExtension))
                        lpPos = lpPos + Len(tGraphicControlExtension)
                        
                    Case GIF_COMMENT_EXTENSION
                        
                        '-- Move forward one byte (Extension label)
                        lpPos = lpPos + 1
                        
                        '-- New comment
                        ReDim Preserve sComments(0 To UBound(sComments) + 1)
                        
                        '-- Get byte array and convert to string
                        Call pvGetDataSubBlocks(Stream(), lpPos, aData())
                        sComments(UBound(sComments)) = Space$(UBound(aData))
                        Call CopyMemory(ByVal sComments(UBound(sComments)), aData(1), UBound(aData))
                    
                    Case Else '-- NOT SUPPORTED or UNEXPECTED extension
                        
                        '-- Move forward one byte (Extension label)
                        lpPos = lpPos + 1
                        
                        '-- Skip data sub-blocks
                        Call pvSkipDataSubBlocks(Stream(), lpPos)
                End Select
                
            Case GIF_IMAGE_DESCRIPTOR
                
                '-- Move forward one byte (Extension label)
                lpPos = lpPos + 1
                        
                '-- Image counter
                nImg = nImg + 1
                ReDim Preserve m_Frame(nImg)
                
                '-- Get structure
                Call CopyMemory(tImageDescriptor, Stream(lpPos), Len(tImageDescriptor))
                lpPos = lpPos + Len(tImageDescriptor)
                
                '-- Store last graphic control extension data
                With tGraphicControlExtension
                    m_Frame(nImg).frDelay = .gceDelayTime
                    m_Frame(nImg).frDisposalMethod = pvGetFlag(.gceFlags, gfDisposalMethod)
                    m_Frame(nImg).frUserInput = pvGetFlag(.gceFlags, gfUserInput)
                    m_Frame(nImg).frUseTransparentColor = pvGetFlag(.gceFlags, gfUseTransparentColor)
                    If (m_Frame(nImg).frUseTransparentColor) Then
                        m_Frame(nImg).frTransparentColorIndex = .gceTransparentColorIndex
                    End If
                End With
                
                '-- Store last comment/s
                ReDim m_Frame(nImg).frComment(0 To UBound(sComments))
                For nCmnt = 1 To UBound(sComments)
                    m_Frame(nImg).frComment(nCmnt) = sComments(nCmnt)
                Next nCmnt
                ReDim sComments(0)
                
                '//
                
                With tImageDescriptor
                
                    '-- Image descriptor
                    m_Frame(nImg).frLeft = .idImageLeftPosition
                    m_Frame(nImg).frTop = .idImageTopPosition
                    m_Frame(nImg).frInterlaced = -(pvGetFlag(.idFlags, gfInterlaced))
                    
                    '-- Local palette [?]
                    m_Frame(nImg).frLocalPaletteUsed = -(pvGetFlag(.idFlags, gfLocalPaletteExists))
                    If (m_Frame(nImg).frLocalPaletteUsed) Then
                    
                        '-- Get palette flags
                        m_Frame(nImg).frLocalPaletteSorted = -(pvGetFlag(.idFlags, gfLocalPaletteSorted))
                        m_Frame(nImg).frLocalPaletteEntries = 2 ^ (pvGetFlag(.idFlags, gfLocalPaletteSize) + 1)
                        
                        '-- Get palette
                        For lpEnt = 0 To m_Frame(nImg).frLocalPaletteEntries - 1
                            m_Frame(nImg).frLocalPalette(4 * lpEnt + 0) = Stream(lpPos + 3 * lpEnt + 2)
                            m_Frame(nImg).frLocalPalette(4 * lpEnt + 1) = Stream(lpPos + 3 * lpEnt + 1)
                            m_Frame(nImg).frLocalPalette(4 * lpEnt + 2) = Stream(lpPos + 3 * lpEnt + 0)
                        Next lpEnt
                        lpPos = lpPos + m_Frame(nImg).frLocalPaletteEntries * 3
                      
                      Else
                        '-- Store a local copy
                        m_Frame(nImg).frLocalPaletteSorted = m_GlobalPaletteSorted
                        m_Frame(nImg).frLocalPaletteEntries = m_GlobalPaletteEntries
                        Call CopyMemory(m_Frame(nImg).frLocalPalette(0), m_GlobalPalette(0), 1024)
                    End If
                    
                    '-- Create Image (8-bpp XOR) and Mask (1-bpp AND) DIBs
                    Set m_Frame(nImg).frDIBXOR = New cDIB
                    Set m_Frame(nImg).frDIBAND = New cDIB
                    '-- Get DIB bits
                    Call m_Frame(nImg).frDIBXOR.Create(.idImageWidth, .idImageHeight, [08_bpp])
                    Call m_Frame(nImg).frDIBAND.Create(.idImageWidth, .idImageHeight, [01_bpp])
                End With

                '-- Get LZW code size
                aLZWCodeSize = Stream(lpPos)
                lpPos = lpPos + 1
                '-- Get data sub-block/s (Encoded image bytes)
                Call pvGetDataSubBlocks(Stream(), lpPos, aData())
                '-- Build XOR and AND DIBs
                Call pvDecodeFrameImage(nImg, aLZWCodeSize, aData())
            
            '-- End of data:
            Case GIF_TRAILER: Exit Do
        
        End Select
        
        '-- Move forward one byte (Next extension label):
        lpPos = lpPos + 1
        
    Loop Until lpPos > UBound(Stream)
    
    '-- Recolect trailing comments
    ReDim m_TrailingComment(0 To UBound(sComments))
    For nCmnt = 1 To UBound(sComments)
        m_TrailingComment(nCmnt) = sComments(nCmnt)
    Next nCmnt
    
    '//
        
    '-- Get image/s
    m_FramesCount = nImg
    '-- Success
    LoadFromStream = (m_FramesCount > 0)
    
On Error GoTo 0
End Function

Public Function Save(ByVal Filename As String) As Boolean

  Dim hFile As Long
  
  Dim tApplicationExtension    As tGIFNetscapeApplicationExtension
  Dim tGraphicControlExtension As tGIFGraphicControlExtension
  Dim tImageDescriptor         As tGIFImageDescriptor
  
  Dim aGPal()      As Byte
  Dim aLPal()      As Byte
  Dim aGPalBPP     As Byte
  Dim aLPalBPP     As Byte
  Dim aLZWCodeSize As Byte
  Dim lpEnt        As Long
  Dim nFrm         As Integer
  
    On Error Resume Next
    Kill Filename
    On Error GoTo 0
  
    On Error GoTo ErrH
    
    '-- Create file...
    hFile = FreeFile
    Open Filename For Binary Access Write As #hFile
    
    '-- Write signature:
    Put #hFile, , GIF_SIGNATURE
      
    '-- Write version:
    If (m_FramesCount > 0 Or m_Frame(1).frUseTransparentColor) Then
        Put #hFile, , GIF_VERSION_89a
      Else
        Put #hFile, , GIF_VERSION_87a
    End If
    
    '-- Prepare Global Palette [?]:
    If (m_GlobalPaletteExists) Then
    
        '-- Get palette color depth
        Do: aGPalBPP = aGPalBPP + 1
        Loop Until 2 ^ aGPalBPP >= m_GlobalPaletteEntries
        
        '-- Build palette (RGB triples)
        ReDim aGPal(0 To 3 * (2 ^ aGPalBPP) - 1)
        For lpEnt = 0 To m_GlobalPaletteEntries - 1
            aGPal(3 * lpEnt + 0) = m_GlobalPalette(4 * lpEnt + 2)
            aGPal(3 * lpEnt + 1) = m_GlobalPalette(4 * lpEnt + 1)
            aGPal(3 * lpEnt + 2) = m_GlobalPalette(4 * lpEnt + 0)
        Next lpEnt
    End If
    
    '-- Prepare Screen Descriptor flags:
    With m_ScreenDescriptor
        .sdFlags = (&H1 * (aGPalBPP - 1) Or _
                    &H8 * -m_GlobalPaletteSorted Or _
                    &H10 * (aGPalBPP - 1) Or _
                    &H80 * -m_GlobalPaletteExists _
                    ) * -m_GlobalPaletteExists
    End With
    
    '-- Write Screen Descriptor and Global Palette [?]:
    Put #hFile, , m_ScreenDescriptor
    If (m_GlobalPaletteExists) Then
        Put #hFile, , aGPal()
    End If
 
    '-- Prepare and write Netscape Application block [?]:
    If (m_FramesCount > 1) Then
    
        With tApplicationExtension
            .naeBlockSize = &HB
            .naeNestcapeID = GIF_NETSCAPE_ID
            .naeSubBlockSize = &H3
            .naeFlags = &H1
            .naeLoops = m_Loops
        End With
        Put #hFile, , GIF_EXTENSION_INTRODUCER
        Put #hFile, , GIF_APPLICATION_EXTENSION
        Put #hFile, , tApplicationExtension
        Put #hFile, , GIF_BLOCK_TERMINATOR
    End If
    
    '-- Write frame block/s
    For nFrm = 1 To m_FramesCount
        
        '-- Prepare and write frame Comment/s:
        Call pvWriteCommentBlocks(hFile, m_Frame(nFrm).frComment())
        
        '-- Prepare and write Graphic Control structure:
        With tGraphicControlExtension
            .gceBlockSize = &H4
            .gceDelayTime = m_Frame(nFrm).frDelay
            .gceTransparentColorIndex = IIf(m_Frame(nFrm).frUseTransparentColor, m_Frame(nFrm).frTransparentColorIndex, 0)
            .gceFlags = &H1 * -m_Frame(nFrm).frUseTransparentColor Or _
                        &H2 * -m_Frame(nFrm).frUserInput Or _
                        &H4 * m_Frame(nFrm).frDisposalMethod
        End With
        Put #hFile, , GIF_EXTENSION_INTRODUCER
        Put #hFile, , GIF_GRAPHIC_CONTROL_EXTENSION
        Put #hFile, , tGraphicControlExtension
        Put #hFile, , GIF_BLOCK_TERMINATOR
        
        '-- Prepare Local Palette [?]:
        If (m_Frame(nFrm).frLocalPaletteUsed) Then
        
            '-- Get palette color depth
            aLPalBPP = 0
            Do: aLPalBPP = aLPalBPP + 1
            Loop Until 2 ^ aLPalBPP >= m_Frame(nFrm).frLocalPaletteEntries
            
            '-- Build palette (RGB triples)
            ReDim aLPal(0 To 3 * (2 ^ aLPalBPP) - 1)
            With m_Frame(nFrm)
                For lpEnt = 0 To .frLocalPaletteEntries - 1
                    aLPal(3 * lpEnt + 0) = .frLocalPalette(4 * lpEnt + 2)
                    aLPal(3 * lpEnt + 1) = .frLocalPalette(4 * lpEnt + 1)
                    aLPal(3 * lpEnt + 2) = .frLocalPalette(4 * lpEnt + 0)
                Next lpEnt
            End With
        End If
        
        '-- Prepare Image Descriptor structure:
        With tImageDescriptor
            .idImageWidth = m_Frame(nFrm).frDIBXOR.Width
            .idImageHeight = m_Frame(nFrm).frDIBXOR.Height
            .idImageLeftPosition = m_Frame(nFrm).frLeft
            .idImageTopPosition = m_Frame(nFrm).frTop
            .idFlags = &H1 * IIf(m_Frame(nFrm).frLocalPaletteUsed, aLPalBPP - 1, aGPalBPP - 1) Or _
                       &H20 * -m_Frame(nFrm).frLocalPaletteSorted Or _
                       &H40 * -m_Frame(nFrm).frInterlaced Or _
                       &H80 * -m_Frame(nFrm).frLocalPaletteUsed
        End With
        
        '-- Write Image Descriptor and Palette [?]:
        Put #hFile, , GIF_IMAGE_DESCRIPTOR
        Put #hFile, , tImageDescriptor
        If (m_Frame(nFrm).frLocalPaletteUsed) Then
            Put #hFile, , aLPal()
        End If
        
        '-- Write LZW code size:
        If (m_Frame(nFrm).frLocalPaletteUsed) Then
            aLZWCodeSize = IIf(aLPalBPP > 1, aLPalBPP, 2)
          Else
            aLZWCodeSize = IIf(aGPalBPP > 1, aGPalBPP, 2)
        End If

        '-- Compress/Write frame Image:
        Put #hFile, , aLZWCodeSize
        With m_Frame(nFrm)
            Call mGIFLZWEnc.Encode(hFile, .frDIBXOR, (aLZWCodeSize + 1), .frInterlaced)
        End With
        Put #hFile, , GIF_BLOCK_TERMINATOR
    Next nFrm
    
    '-- Prepare and write trailing Comment/s:
    Call pvWriteCommentBlocks(hFile, m_TrailingComment())
    
    '-- Finaly, write GIF Trailer label and close file:
    Put #hFile, , GIF_TRAILER
    Close #hFile
    
    '-- Success:
    Save = True
    
ErrH:
End Function

Public Sub Destroy()
   
  Dim nFrm As Integer
    
    '-- Clear Signature/Version
    m_Signature = vbEmpty
    m_Version = vbEmpty
    
    '-- Clear global variables
    Call ZeroMemory(m_ScreenDescriptor, Len(m_ScreenDescriptor))
    Call ZeroMemory(m_GlobalPalette(0), 1024)
    m_GlobalPaletteExists = False
    m_GlobalPaletteSorted = False
    m_GlobalPaletteEntries = 0
    m_Loops = 0
    
    '-- Clear frame/s variables
    For nFrm = 1 To m_FramesCount
        Set m_Frame(nFrm).frDIBXOR = Nothing
        Set m_Frame(nFrm).frDIBAND = Nothing
    Next nFrm
    ReDim m_Frame(0)
    ReDim m_TrailingComment(0)
    
    '-- Reset frames counter
    m_FramesCount = 0
End Sub

'//

Public Sub FrameDraw(ByVal hDC As Long, ByVal nFrame As Integer, Optional ByVal Left As Integer = 0, Optional ByVal Top As Integer = 0, Optional ByVal ScaleFactor As Single = 1, Optional ByVal ScreenCrop As Boolean = False)
  
  Dim nScrW As Integer
  Dim nScrH As Integer
  
    With m_Frame(nFrame)
    
        If (ScreenCrop) Then
            '-- Crop to screen rectangle
            If (.frLeft + .frDIBXOR.Width > m_ScreenDescriptor.sdLogicalScreenWidth) Then
                nScrW = m_ScreenDescriptor.sdLogicalScreenWidth - .frLeft
              Else
                nScrW = .frDIBXOR.Width
            End If
            If (.frTop + .frDIBXOR.Height > m_ScreenDescriptor.sdLogicalScreenHeight) Then
                nScrH = m_ScreenDescriptor.sdLogicalScreenHeight - .frTop
              Else
                nScrH = .frDIBXOR.Height
            End If
          Else
            '-- Whole frame
            nScrW = .frDIBXOR.Width
            nScrH = .frDIBXOR.Height
        End If
        
        '-- Check negatives (avoids flipping)
        If (nScrW < 0) Then nScrW = 0
        If (nScrH < 0) Then nScrH = 0
        
        '-- Paint frame
        If (.frUseTransparentColor) Then
            Call .frDIBAND.Stretch(hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH, vbSrcAnd)
            Call .frDIBXOR.Stretch(hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH, vbSrcPaint)
          Else
            Call .frDIBXOR.Stretch(hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH)
        End If
    End With
End Sub

Public Sub FrameMask(ByVal nFrame As Integer, ByVal TransparentColorIndex As Byte)
  
  Dim aPalXOR()  As Byte
  Dim aPalMsk()  As Byte
  Dim aPalAND(7) As Byte
    
    With m_Frame(nFrame)
        
        '-- Temp. palettes
        aPalXOR() = .frLocalPalette()
        aPalMsk() = .frLocalPalette()
        Call FillMemory(aPalAND(4), 3, &HFF)
        
        '-- Set new transp. entry [?]
        If (.frUseTransparentColor) Then
        
            Call FillMemory(aPalXOR(TransparentColorIndex * 4), 3, &H0)
            Call FillMemory(aPalMsk(0), 1024, &H0)
            Call FillMemory(aPalMsk(TransparentColorIndex * 4), 3, &HFF)

            '-- AND DIB (Use transparent):
            Call .frDIBXOR.SetPalette(aPalMsk())
            Call .frDIBAND.SetPalette(aPalAND())
            Call .frDIBAND.LoadBlt(.frDIBXOR.hDC)
            '-- XOR DIB:
            Call .frDIBXOR.SetPalette(aPalXOR())
            
          Else
            '-- AND DIB (Don't use transparent):
            Call .frDIBAND.SetPalette(aPalAND())
            Call .frDIBAND.Cls(&H0)
            '-- XOR DIB:
            Call .frDIBXOR.SetPalette(aPalXOR())
        End If
    End With
End Sub

'//

Public Sub FrameInsert(ByVal nFrame As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal UseGlobalPalette As Boolean = False)

  Dim nFrm     As Integer
  Dim bFrmAdd  As Boolean
  Dim aPalBW() As Byte
  
    '-- Add [?]
    bFrmAdd = (nFrame > m_FramesCount Or m_FramesCount = 0)
  
    '-- Increase frames count
    m_FramesCount = m_FramesCount + 1
    
    '-- Redim. frame array
    ReDim Preserve m_Frame(m_FramesCount)
    
    '-- Move frames [?]
    If (bFrmAdd) Then
        nFrm = m_FramesCount
      Else
        For nFrm = m_FramesCount - 1 To nFrame Step -1
            Call Me.FrameMoveDown(nFrm)
        Next nFrm
        nFrm = nFrame
    End If
    
    '-- Create B&W palette
    ReDim aPalBW(7) As Byte
    Call FillMemory(aPalBW(4), 3, &HFF)
    
    '-- Create XOR and AND DIBs (1-bpp BW palette)
    With m_Frame(nFrm)
        Set .frDIBXOR = New cDIB
        Call .frDIBXOR.Create(Width, Height, [08_bpp])
        Call .frDIBXOR.SetPalette(aPalBW())
        Set .frDIBAND = New cDIB
        Call .frDIBAND.Create(Width, Height, [01_bpp])
        Call .frDIBAND.SetPalette(aPalBW())
    End With
    
    '-- Basic default
    With m_Frame(nFrm)
        If (UseGlobalPalette And m_GlobalPaletteExists) Then
            Call CopyMemory(.frLocalPalette(0), m_GlobalPalette(0), 1024)
            .frLocalPaletteUsed = False
            .frLocalPaletteEntries = m_GlobalPaletteEntries
          Else
            Call CopyMemory(.frLocalPalette(0), aPalBW(0), 8)
            .frLocalPaletteUsed = True
            .frLocalPaletteEntries = 2
        End If
        ReDim .frComment(0)
    End With
    
    '-- Frame #1 [?]
    If (m_FramesCount = 1) Then
        m_ScreenDescriptor.sdLogicalScreenWidth = Width
        m_ScreenDescriptor.sdLogicalScreenHeight = Height
        ReDim m_TrailingComment(0)
    End If
End Sub

Public Sub FrameClone(ByVal nFrame As Integer)

  Dim nFrm As Integer

    '-- Increase frames count
    m_FramesCount = m_FramesCount + 1

    '-- Redim. frame array
    ReDim Preserve m_Frame(m_FramesCount)

    '-- Move frames
    For nFrm = m_FramesCount - 1 To nFrame Step -1
        Call Me.FrameMoveDown(nFrm)
    Next nFrm

    '-- Clone...
    With m_Frame(nFrame + 1)

        '-- DIBs
        Set m_Frame(nFrame).frDIBXOR = New cDIB
        Set m_Frame(nFrame).frDIBAND = New cDIB
        Call .frDIBXOR.CloneTo(m_Frame(nFrame).frDIBXOR)
        Call .frDIBAND.CloneTo(m_Frame(nFrame).frDIBAND)

        '-- Local palette copy
        Call CopyMemory(m_Frame(nFrame).frLocalPalette(0), .frLocalPalette(0), 1024)
        m_Frame(nFrame).frLocalPaletteEntries = .frLocalPaletteEntries
        m_Frame(nFrame).frLocalPaletteSorted = .frLocalPaletteSorted
        m_Frame(nFrame).frLocalPaletteUsed = .frLocalPaletteUsed

        '-- Props.:
        m_Frame(nFrame).frLeft = .frLeft
        m_Frame(nFrame).frTop = .frTop
        m_Frame(nFrame).frDelay = .frDelay
        m_Frame(nFrame).frDisposalMethod = .frDisposalMethod
        m_Frame(nFrame).frUseTransparentColor = .frUseTransparentColor
        m_Frame(nFrame).frTransparentColorIndex = .frTransparentColorIndex
        m_Frame(nFrame).frInterlaced = .frInterlaced
        m_Frame(nFrame).frUserInput = .frUserInput
        m_Frame(nFrame).frComment() = .frComment()
    End With
End Sub

Public Sub FrameRemove(ByVal nFrame As Integer)

  Dim nFrm As Integer
  
    '-- Destroy DIBs
    Set m_Frame(nFrame).frDIBXOR = Nothing
    Set m_Frame(nFrame).frDIBAND = Nothing
   
    '-- Move frames [?]
    If (nFrame < m_FramesCount) Then
        For nFrm = nFrame To m_FramesCount - 1
            m_Frame(nFrm) = m_Frame(nFrm + 1)
        Next nFrm
    End If
    '-- Decrease frames count
    m_FramesCount = m_FramesCount - 1
    
    '-- Redim. frame arrays [?]
    If (m_FramesCount > 0) Then
        ReDim Preserve m_Frame(m_FramesCount)
      Else
        ReDim m_Frame(0)
    End If
End Sub

Public Sub FrameMoveUp(ByVal nFrame As Integer)

  Dim swpFrame As tFrame
    
    '-- Swap frame
    swpFrame = m_Frame(nFrame - 1)
    m_Frame(nFrame - 1) = m_Frame(nFrame)
    m_Frame(nFrame) = swpFrame
End Sub

Public Sub FrameMoveDown(ByVal nFrame As Integer)

  Dim swpFrame As tFrame
  
    '-- Swap frame
    swpFrame = m_Frame(nFrame + 1)
    m_Frame(nFrame + 1) = m_Frame(nFrame)
    m_Frame(nFrame) = swpFrame
End Sub

'//

Public Sub FrameCommentAdd(ByVal nFrame As Integer)
    ReDim Preserve m_Frame(nFrame).frComment(0 To UBound(m_Frame(nFrame).frComment) + 1)
End Sub
Public Sub FrameCommentRemoveAll(ByVal nFrame As Integer)
    ReDim m_Frame(nFrame).frComment(0)
End Sub

Public Sub TrailingCommentAdd()
    ReDim Preserve m_TrailingComment(0 To UBound(m_TrailingComment) + 1)
End Sub
Public Sub TrailingCommentRemoveAll()
    ReDim m_TrailingComment(0)
End Sub

'========================================================================================
' Properties
'========================================================================================

'----------------------------------------------------------------------------------------
' Animation
'----------------------------------------------------------------------------------------

Public Property Get AnimationLoops() As Integer
    AnimationLoops = m_Loops
End Property
Public Property Let AnimationLoops(ByVal New_AnimationLoops As Integer)
    m_Loops = New_AnimationLoops
End Property

Public Property Get TrailingCommentsCount() As Integer
    TrailingCommentsCount = UBound(m_TrailingComment)
End Property

Public Property Get TrailingComment(ByVal nComment As Integer) As String
    TrailingComment = m_TrailingComment(nComment)
End Property

Public Property Let TrailingComment(ByVal nComment As Integer, New_TrailingComment As String)
    m_TrailingComment(nComment) = New_TrailingComment
End Property

'----------------------------------------------------------------------------------------
' Screen
'----------------------------------------------------------------------------------------

Public Property Get ScreenWidth() As Integer
    ScreenWidth = m_ScreenDescriptor.sdLogicalScreenWidth
End Property
Public Property Let ScreenWidth(ByVal New_ScreenWidth As Integer)
    m_ScreenDescriptor.sdLogicalScreenWidth = New_ScreenWidth
End Property

Public Property Get ScreenHeight() As Integer
    ScreenHeight = m_ScreenDescriptor.sdLogicalScreenHeight
End Property
Public Property Let ScreenHeight(ByVal New_ScreenHeight As Integer)
    m_ScreenDescriptor.sdLogicalScreenHeight = New_ScreenHeight
End Property

Public Property Get ScreenBackgroundColorIndex() As Byte
    ScreenBackgroundColorIndex = m_ScreenDescriptor.sdBackgroundColorIndex
End Property
Public Property Let ScreenBackgroundColorIndex(ByVal New_ScreenBackgroundColorIndex As Byte)
    m_ScreenDescriptor.sdBackgroundColorIndex = New_ScreenBackgroundColorIndex
End Property

Public Property Get ScreenPixelAspectRatio() As Single
    ScreenPixelAspectRatio = (m_ScreenDescriptor.sdPixelAspectRatio + 15) / 64
End Property
Public Property Let ScreenPixelAspectRatio(ByVal New_ScreenPixelAspectRatio As Single)
    m_ScreenDescriptor.sdPixelAspectRatio = CByte(New_ScreenPixelAspectRatio * 64 - 15)
End Property

'----------------------------------------------------------------------------------------
' Global palette
'----------------------------------------------------------------------------------------

Public Property Get lpGlobalPalette() As Long
    lpGlobalPalette = VarPtr(m_GlobalPalette(0))
End Property

Public Property Get GlobalPaletteColor(ByVal nEntry As Byte) As Long
    If (m_GlobalPaletteExists) Then
        GlobalPaletteColor = RGB(m_GlobalPalette(nEntry * 4 + 2), m_GlobalPalette(nEntry * 4 + 1), m_GlobalPalette(nEntry * 4))
    End If
End Property

Public Property Get GlobalPaletteExists() As Boolean
    GlobalPaletteExists = m_GlobalPaletteExists
End Property
Public Property Let GlobalPaletteExists(ByVal New_GlobalPaletteExists As Boolean)
    m_GlobalPaletteExists = New_GlobalPaletteExists
End Property

Public Property Get GlobalPaletteSorted() As Boolean
    GlobalPaletteSorted = m_GlobalPaletteSorted
End Property
Public Property Let GlobalPaletteSorted(ByVal New_GlobalPaletteSorted As Boolean)
    m_GlobalPaletteSorted = New_GlobalPaletteSorted
End Property

Public Property Get GlobalPaletteEntries() As Integer
    GlobalPaletteEntries = m_GlobalPaletteEntries
End Property
Public Property Let GlobalPaletteEntries(ByVal New_GlobalPaletteEntries As Integer)
    m_GlobalPaletteEntries = New_GlobalPaletteEntries
End Property

'----------------------------------------------------------------------------------------
' Frame
'----------------------------------------------------------------------------------------

Public Property Get FramesCount() As Integer
    FramesCount = m_FramesCount
End Property

'-- Use it carefully!
Public Property Get FrameDIBXOR(ByVal nFrame As Integer) As cDIB
    Set FrameDIBXOR = m_Frame(nFrame).frDIBXOR
End Property

'-- Use it carefully!
Public Property Get FrameDIBAND(ByVal nFrame As Integer) As cDIB
    Set FrameDIBAND = m_Frame(nFrame).frDIBAND
End Property

Public Property Get FrameImage(ByVal nFrame As Integer) As StdPicture
    
  Dim aPalCopy() As Byte
  
    With m_Frame(nFrame)
    
        '-- Transparent [?]: Unmask frame (XOR DIB)
        If (.frUseTransparentColor) Then
            Call .frDIBXOR.GetPalette(aPalCopy())
            Call .frDIBXOR.SetPalette(.frLocalPalette())
        End If
        
        '-- Now, get image
        Set FrameImage = .frDIBXOR.Image

        '-- Transparent [?]: Remask frame (XOR DIB)
        If (.frUseTransparentColor) Then
            Call .frDIBXOR.SetPalette(aPalCopy())
        End If
    End With
End Property

'//

Public Property Get FrameDelay(ByVal nFrame As Integer) As Integer
    FrameDelay = m_Frame(nFrame).frDelay
End Property
Public Property Let FrameDelay(ByVal nFrame As Integer, ByVal New_Delay As Integer)
    m_Frame(nFrame).frDelay = New_Delay
End Property

Public Property Get FrameInterlaced(ByVal nFrame As Integer) As Boolean
    FrameInterlaced = m_Frame(nFrame).frInterlaced
End Property
Public Property Let FrameInterlaced(ByVal nFrame As Integer, ByVal New_Interlaced As Boolean)
    m_Frame(nFrame).frInterlaced = New_Interlaced
End Property

Public Property Get FrameLeft(ByVal nFrame As Integer) As Integer
    FrameLeft = m_Frame(nFrame).frLeft
End Property
Public Property Let FrameLeft(ByVal nFrame As Integer, ByVal New_FrameLeft As Integer)
    m_Frame(nFrame).frLeft = New_FrameLeft
End Property

Public Property Get FrameTop(ByVal nFrame As Integer) As Integer
    FrameTop = m_Frame(nFrame).frTop
End Property
Public Property Let FrameTop(ByVal nFrame As Integer, ByVal New_FrameTop As Integer)
    m_Frame(nFrame).frTop = New_FrameTop
End Property

Public Property Get FrameDisposalMethod(ByVal nFrame As Integer) As GIFDisposalMethod
    FrameDisposalMethod = m_Frame(nFrame).frDisposalMethod
End Property
Public Property Let FrameDisposalMethod(ByVal nFrame As Integer, ByVal New_DisposalMethod As GIFDisposalMethod)
    m_Frame(nFrame).frDisposalMethod = New_DisposalMethod
End Property

Public Property Get FrameUseTransparentColor(ByVal nFrame As Integer) As Boolean
    FrameUseTransparentColor = m_Frame(nFrame).frUseTransparentColor
End Property
Public Property Let FrameUseTransparentColor(ByVal nFrame As Integer, ByVal New_UseTransparentColor As Boolean)
    m_Frame(nFrame).frUseTransparentColor = New_UseTransparentColor
End Property

Public Property Get FrameTransparentColorIndex(ByVal nFrame As Integer) As Byte
    FrameTransparentColorIndex = m_Frame(nFrame).frTransparentColorIndex
End Property
Public Property Let FrameTransparentColorIndex(ByVal nFrame As Integer, ByVal New_TransparentColorIndex As Byte)
    m_Frame(nFrame).frTransparentColorIndex = New_TransparentColorIndex
End Property

Public Property Get FrameUserInput(ByVal nFrame As Integer) As Boolean
    FrameUserInput = m_Frame(nFrame).frUserInput
End Property
Public Property Let FrameUserInput(ByVal nFrame As Integer, ByVal New_FrameUserInput As Boolean)
    m_Frame(nFrame).frUserInput = New_FrameUserInput
End Property

'----------------------------------------------------------------------------------------
' Local palette
'----------------------------------------------------------------------------------------

Public Property Get lpFramePalette(ByVal nFrame As Integer) As Long
    lpFramePalette = VarPtr(m_Frame(nFrame).frLocalPalette(0))
End Property

Public Property Get lpFramePaletteColor(ByVal nFrame As Integer, ByVal nEntry As Byte) As Long
    With m_Frame(nFrame)
        lpFramePaletteColor = RGB(.frLocalPalette(nEntry * 4 + 2), .frLocalPalette(nEntry * 4 + 1), .frLocalPalette(nEntry * 4))
    End With
End Property

Public Property Get LocalPaletteUsed(ByVal nFrame As Integer) As Boolean
    LocalPaletteUsed = m_Frame(nFrame).frLocalPaletteUsed
End Property
Public Property Let LocalPaletteUsed(ByVal nFrame As Integer, ByVal New_LocalPaletteUsed As Boolean)
    m_Frame(nFrame).frLocalPaletteUsed = New_LocalPaletteUsed
End Property

Public Property Get LocalPaletteSorted(ByVal nFrame As Integer) As Boolean
    LocalPaletteSorted = m_Frame(nFrame).frLocalPaletteSorted
End Property
Public Property Let LocalPaletteSorted(ByVal nFrame As Integer, ByVal New_LocalPaletteSorted As Boolean)
    m_Frame(nFrame).frLocalPaletteSorted = New_LocalPaletteSorted
End Property

Public Property Get LocalPaletteEntries(ByVal nFrame As Integer) As Integer
    LocalPaletteEntries = m_Frame(nFrame).frLocalPaletteEntries
End Property
Public Property Let LocalPaletteEntries(ByVal nFrame As Integer, ByVal New_LocalPaletteEntries As Integer)
    m_Frame(nFrame).frLocalPaletteEntries = New_LocalPaletteEntries
End Property

'----------------------------------------------------------------------------------------
' Comment/s
'----------------------------------------------------------------------------------------

Public Property Get FrameCommentsCount(ByVal nFrame As Integer) As Integer
    FrameCommentsCount = UBound(m_Frame(nFrame).frComment)
End Property

Public Property Get FrameComment(ByVal nFrame As Integer, ByVal nComment As Integer) As String
    FrameComment = m_Frame(nFrame).frComment(nComment)
End Property

Public Property Let FrameComment(ByVal nFrame As Integer, ByVal nComment As Integer, ByVal New_FrameComment As String)
    m_Frame(nFrame).frComment(nComment) = New_FrameComment
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvGetDataSubBlocks(aStream() As Byte, lpPos As Long, aBuffer() As Byte)
                    
  ReDim aBuffer(0) As Byte
  Dim lBufferMax   As Long
  Dim aBlockSize   As Byte
    
    '-- Get data sub-blocks...
    Do
        aBlockSize = aStream(lpPos)
        lBufferMax = UBound(aBuffer)
        ReDim Preserve aBuffer(lBufferMax + aBlockSize)
        
        lBufferMax = lBufferMax + aBlockSize
        Call CopyMemory(aBuffer(lBufferMax - aBlockSize + 1), aStream(lpPos + 1), aBlockSize)
        
        lpPos = lpPos + aBlockSize + 1
    Loop While aStream(lpPos) > 0
End Sub

Private Sub pvSkipDataSubBlocks(aStream() As Byte, lpPos As Long)
                    
  Dim aBlockSize As Byte
                    
    '-- Skip data sub-blocks...
    Do
        aBlockSize = aStream(lpPos)
        lpPos = lpPos + aBlockSize + 1
    Loop While aStream(lpPos) > 0
End Sub

Private Function pvGetFlag(ByVal lFlags As Long, ByVal lMask As GIFflags) As Byte

    '-- Extract packed flag
    If (lMask > 0) Then
        pvGetFlag = (lFlags And lMask)
        Do While (lMask And 1) = 0
            lMask = lMask \ 2
            pvGetFlag = pvGetFlag \ 2
        Loop
    End If
End Function

'//

Private Sub pvDecodeFrameImage(ByVal nFrame As Integer, ByVal LZWCodeSize As Byte, EncodedBytes() As Byte)
  
  Dim aPalXOR() As Byte
  
    With m_Frame(nFrame)
        
        '-- Decode XOR bits
        Call mGIFLZWDec.Decode(.frDIBXOR, .frInterlaced, LZWCodeSize, EncodedBytes())
        
        '-- Set XOR palette:
        aPalXOR() = .frLocalPalette()
        If (.frUseTransparentColor) Then
            Call FillMemory(aPalXOR(.frTransparentColorIndex * 4), 3, &H0)
        End If
        Call .frDIBXOR.SetPalette(aPalXOR())
          
        '-- Mask frame [?]
        If (.frUseTransparentColor) Then
            Call Me.FrameMask(nFrame, .frTransparentColorIndex)
        End If
    End With
End Sub

'//

Private Sub pvWriteCommentBlocks(ByVal hFile As Long, sComments() As String)
  
  Dim nCmnt      As Integer
  Dim aBuff()    As Byte
  Dim lBuffSize  As Long
  Dim lBuffPos   As Long
  Dim aBlock()   As Byte
  Dim aBlockSize As Byte
  
  Const aNULL    As Byte = &H0
           
    For nCmnt = 1 To UBound(sComments)
    
        '-- Extension/type labels
        Put #hFile, , GIF_EXTENSION_INTRODUCER
        Put #hFile, , GIF_COMMENT_EXTENSION
        
        '-- Full-lenght comment. Store as byte buffer
        lBuffSize = Len(sComments(nCmnt))
        
        If (lBuffSize) Then
        
            ReDim aBuff(1 To lBuffSize) As Byte
            Call CopyMemory(aBuff(1), ByVal sComments(nCmnt), lBuffSize)
            
            '-- Split to data sub-blocks
            lBuffPos = 1
            Do While lBuffPos < lBuffSize
                
                '-- Get block
                If (lBuffPos + 255 < lBuffSize) Then
                    aBlockSize = 255
                  Else
                    aBlockSize = (lBuffSize - lBuffPos) + 1
                End If
                ReDim aBlock(1 To aBlockSize) As Byte
                Call CopyMemory(aBlock(1), aBuff(lBuffPos), aBlockSize)
                '-- Move forward block size
                lBuffPos = lBuffPos + aBlockSize
                
                '-- Write block
                Put #hFile, , aBlockSize
                Put #hFile, , aBlock()
            Loop
            
          Else
            Put #hFile, , aNULL
            Put #hFile, , aNULL
        End If
        
        '-- Terminator
        Put #hFile, , GIF_BLOCK_TERMINATOR
    Next nCmnt
End Sub
