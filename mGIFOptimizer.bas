Attribute VB_Name = "mGIFOptimizer"
'================================================
' Module:        mGIFOptimizer.bas
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
' Last revision: 2004.01.05
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT2, lpRect2 As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

Public Enum eFrameOptimizationMethod
    [fomMinimumBoundingRectangle] = 0
    [fomFrameDifferencing]
End Enum

'//

'========================================================================================
' Methods
'========================================================================================

Public Function CheckGIF(oGIF As cGIF, oProgress As ucProgress) As Boolean
 
  Dim nCurr        As Integer
  Dim nPrev        As Integer
  Dim tSACurr      As SAFEARRAY2D
  Dim aBitsCurr()  As Byte
  
  Dim bIsTrns      As Boolean
  Dim aTrnsIdxCurr As Byte
  Dim aMaxIdx      As Byte
                    
  Dim rRectCurr    As RECT2
  Dim rRectNext    As RECT2
  Dim rMerge       As RECT2
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
  
    On Error GoTo ErrH
  
    With oGIF
    
        oProgress.Max = 2 * .FramesCount
    
        '== Scan for transparent images without transparent pixels:
                
        For nCurr = 1 To .FramesCount
            
            '-- Current progress
            oProgress.Value = nCurr
            
            If (.FrameUseTransparentColor(nCurr)) Then
                
                '-- Map current 8-bpp DIB bits
                Call pvBuild_08bppSA(tSACurr, .FrameDIBXOR(nCurr))
                Call CopyMemory(ByVal VarPtrArray(aBitsCurr()), VarPtr(tSACurr), 4)
                
                '-- Get frame transparent index and dimensions
                aTrnsIdxCurr = .FrameTransparentColorIndex(nCurr)
                W = .FrameDIBXOR(nCurr).Width - 1
                H = .FrameDIBXOR(nCurr).Height - 1
                
                '-- Scan for a not-transparent pixel
                bIsTrns = False
                For y = 0 To H
                    For x = 0 To W
                        If (aBitsCurr(x, y) = aTrnsIdxCurr) Then bIsTrns = True: GoTo ScanDone
                    Next x
                Next y
            
ScanDone:       '-- Unmap DIB bits
                Call CopyMemory(ByVal VarPtrArray(aBitsCurr()), 0&, 4)
                
                '-- No transparent pixel/s:
                '   Turn off transparency and reset index
                If (bIsTrns = False) Then
                    Let .FrameUseTransparentColor(nCurr) = False
                    Let .FrameTransparentColorIndex(nCurr) = 0
                End If
                
              Else
                '-- Reset index
                .FrameTransparentColorIndex(nCurr) = 0
            End If
        Next nCurr
    
        '== Scan for redundant disposal modes:
       
        For nCurr = 2 To .FramesCount

            oProgress.Value = oProgress.Max + nCurr
            nPrev = nCurr - 1

            '-- Get current (n-1) and next frame (n) rects.
            Call SetRect(rRectCurr, .FrameLeft(nPrev), .FrameTop(nPrev), .FrameLeft(nPrev) + .FrameDIBXOR(nPrev).Width, .FrameTop(nPrev) + .FrameDIBXOR(nPrev).Height)
            Call SetRect(rRectNext, .FrameLeft(nCurr), .FrameTop(nCurr), .FrameLeft(nCurr) + .FrameDIBXOR(nCurr).Width, .FrameTop(nCurr) + .FrameDIBXOR(nCurr).Height)

            '-- If current frame is completely covered by next frame...
            If (.FrameDisposalMethod(nPrev) = [dmRestoreToBackground] Or .FrameDisposalMethod(nPrev) = [dmRestoreToPrevious]) Then
                If (.FrameUseTransparentColor(nCurr) = False) Then
                    If (IntersectRect(rMerge, rRectCurr, rRectNext)) Then
                        If (EqualRect(rMerge, rRectCurr)) Then
                            Let .FrameDisposalMethod(nPrev) = [dmNotSpecified]
                        End If
                    End If
                End If
            End If
        Next nCurr: .FrameDisposalMethod(.FramesCount) = [dmNotSpecified]
        
        '== Scan for 'out of palette bounds':
        
        If (.GlobalPaletteExists) Then
            If (.ScreenBackgroundColorIndex > .GlobalPaletteEntries - 1) Then
                Let .GlobalPaletteEntries = .ScreenBackgroundColorIndex + 1
            End If
        End If
        
        If (.GlobalPaletteExists) Then
            aMaxIdx = .GlobalPaletteEntries - 1
            For nCurr = 1 To .FramesCount
                If (Not .LocalPaletteUsed(nCurr)) Then
                    aMaxIdx = .FrameTransparentColorIndex(nCurr)
                End If
            Next nCurr
            If (aMaxIdx > .GlobalPaletteEntries - 1) Then
                Let .GlobalPaletteEntries = aMaxIdx + 1
                If (Not .LocalPaletteUsed(nCurr)) Then
                    .LocalPaletteEntries(nCurr) = .GlobalPaletteEntries
                End If
           End If
        End If
        
        For nCurr = 1 To .FramesCount
            If (.LocalPaletteUsed(nCurr)) Then
                If (.FrameUseTransparentColor(nCurr) And .FrameTransparentColorIndex(nCurr) > .LocalPaletteEntries(nCurr) - 1) Then
                    Let .LocalPaletteEntries(nCurr) = .FrameTransparentColorIndex(nCurr) + 1
                End If
            End If
        Next nCurr
    End With
    
    '-- Success
    CheckGIF = True
    
ErrH:
    oProgress = 0
    On Error GoTo 0
End Function

Public Function OptimizeGlobalPalette(oGIF As cGIF, oProgress As ucProgress) As Integer

  Dim aPal(1023)  As Byte
  Dim bUseGlobal  As Boolean
  
  Dim nBefore     As Integer
  Dim aNPal(1023) As Byte
  Dim aGPal(1023) As Byte
  Dim bUsed()     As Boolean
  Dim aTrnIdx()   As Byte
  Dim aInvIdx()   As Byte
  Dim lIdx        As Long
  Dim lMax        As Long
  
  Dim nCurr       As Integer
  Dim tSA08       As SAFEARRAY2D
  Dim aBits08()   As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    On Error GoTo ErrH
    
    With oGIF
        
        oProgress.Max = 2 * .FramesCount
        
        If (.GlobalPaletteExists) Then
        
            '-- Check if at least one frame uses Global
            bUseGlobal = False
            For nCurr = 1 To .FramesCount
                If (Not .LocalPaletteUsed(nCurr)) Then bUseGlobal = True
            Next nCurr
            '-- Global not used, destroy it
            If (Not bUseGlobal) Then
                Let .GlobalPaletteExists = False
                Let .GlobalPaletteEntries = 0
                Call FillMemory(ByVal .lpGlobalPalette, 1024, 0)
            End If
        End If
        
        If (.GlobalPaletteExists) Then
        
            '-- Store current number of entries and initialize arrays
            nBefore = .GlobalPaletteEntries
            ReDim bUsed(nBefore - 1)
            ReDim aTrnIdx(nBefore - 1)
            ReDim aInvIdx(nBefore - 1)
            
            '-- Store Global palette as RGB byte array
            CopyMemory aGPal(0), ByVal .lpGlobalPalette, 1024
            
            '-- Check all used entries
            For nCurr = 1 To .FramesCount
                
                If (Not .LocalPaletteUsed(nCurr)) Then
                
                    oProgress.Value = nCurr
                
                    '-- Get frame dimensions
                    W = .FrameDIBXOR(nCurr).Width - 1
                    H = .FrameDIBXOR(nCurr).Height - 1
                
                    '-- Map current 8-bpp DIB bits
                    Call pvBuild_08bppSA(tSA08, .FrameDIBXOR(nCurr))
                    Call CopyMemory(ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4)
                    
                    '-- Check used entries...
                    For y = 0 To H
                        For x = 0 To W
                            bUsed(aBits08(x, y)) = True
                        Next x
                    Next y
                    
                    '-- Unmap DIB bits
                    Call CopyMemory(ByVal VarPtrArray(aBits08()), 0&, 4)
                End If
            Next nCurr
        
            '-- 'Strecth' palette...
            For lIdx = 0 To .GlobalPaletteEntries - 1
                If (bUsed(lIdx)) Then
                    aTrnIdx(lMax) = lIdx ' New index
                    aInvIdx(lIdx) = lMax ' Inverse index
                    lMax = lMax + 1      ' Current count
                End If
            Next lIdx
            
            '-- Any entry removed [?]
            If (lMax < .GlobalPaletteEntries) Then
        
                If (lMax > 0) Then lMax = lMax - 1
                '-- Build temp. palette with only used entries
                For lIdx = 0 To lMax
                    Call CopyMemory(aNPal(4 * lIdx), aGPal(4 * aTrnIdx(lIdx)), 4)
                Next lIdx
                
                '-- Set as new Global palette
                Call CopyMemory(ByVal .lpGlobalPalette, aNPal(0), 1024)
                .GlobalPaletteEntries = lMax + 1
                
                '-- Update Background color index
                .ScreenBackgroundColorIndex = aInvIdx(.ScreenBackgroundColorIndex)
                
                '-- Set as XOR DIBs palette (and update indexes)
                For nCurr = 1 To .FramesCount
                    
                    If (Not .LocalPaletteUsed(nCurr)) Then
                    
                        oProgress.Value = .FramesCount + nCurr
                        
                        '-- Set new frame DIB palette
                        Call .FrameDIBXOR(nCurr).SetPalette(aNPal())
                        '-- Store temp. copy
                        Call CopyMemory(ByVal .lpFramePalette(nCurr), aNPal(0), 1024)
                        Let .LocalPaletteEntries(nCurr) = .GlobalPaletteEntries
                        
                        '-- Get dimensions
                        W = .FrameDIBXOR(nCurr).Width - 1
                        H = .FrameDIBXOR(nCurr).Height - 1
                        
                        '-- Map current 8-bpp DIB bits
                        Call pvBuild_08bppSA(tSA08, .FrameDIBXOR(nCurr))
                        Call CopyMemory(ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4)

                        '-- Update indexes...
                        For y = 0 To H
                            For x = 0 To W
                                aBits08(x, y) = aInvIdx(aBits08(x, y))
                            Next x
                        Next y

                        '-- Unmap DIB bits
                        Call CopyMemory(ByVal VarPtrArray(aBits08()), 0&, 4)
                         
                        '-- Update transparent index
                        Let .FrameTransparentColorIndex(nCurr) = aInvIdx(.FrameTransparentColorIndex(nCurr))
                    End If
                Next nCurr
            End If
            oProgress.Value = 0

            '-- Return removed entries
            OptimizeGlobalPalette = (nBefore - .GlobalPaletteEntries)
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErrH:
    oProgress = 0
    OptimizeGlobalPalette = -1
    On Error GoTo 0
End Function

Public Function OptimizeLocalPalette(oGIF As cGIF, ByVal nFrame As Integer) As Integer

  Dim nBefore     As Integer
  Dim aNPal(1023) As Byte
  Dim aLPal(1023) As Byte
  Dim bUsed()     As Boolean
  Dim aTrnIdx()   As Byte
  Dim aInvIdx()   As Byte
  Dim lIdx        As Long
  Dim lMax        As Long
  
  Dim tSA08       As SAFEARRAY2D
  Dim aBits08()   As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    On Error GoTo ErrH
    
    With oGIF
        
        If (.LocalPaletteUsed(nFrame)) Then
        
            '-- Store current number of entries and initialize arrays
            nBefore = .LocalPaletteEntries(nFrame)
            ReDim bUsed(nBefore - 1)
            ReDim aTrnIdx(nBefore - 1)
            ReDim aInvIdx(nBefore - 1)
            
            '-- Store Local palette as RGB byte array
            Call CopyMemory(aLPal(0), ByVal .lpFramePalette(nFrame), 1024)
            
            '-- Get frame dimensions
            W = .FrameDIBXOR(nFrame).Width - 1
            H = .FrameDIBXOR(nFrame).Height - 1
            
            '-- Map current 8-bpp DIB bits
            Call pvBuild_08bppSA(tSA08, .FrameDIBXOR(nFrame))
            Call CopyMemory(ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4)
            
            '-- Check used entries...
            For y = 0 To H
                For x = 0 To W
                    bUsed(aBits08(x, y)) = True
                Next x
            Next y
            
            '-- 'Strecth' palette...
            For lIdx = 0 To .LocalPaletteEntries(nFrame) - 1
                If (bUsed(lIdx)) Then
                    aTrnIdx(lMax) = lIdx ' New index
                    aInvIdx(lIdx) = lMax ' Inverse index
                    lMax = lMax + 1      ' Current count
                End If
            Next lIdx
            
            '-- Any entry removed [?]
            If (lMax < .LocalPaletteEntries(nFrame)) Then
        
                If (lMax > 0) Then lMax = lMax - 1
                '-- Build temp. palette with only used entries
                For lIdx = 0 To lMax
                    Call CopyMemory(aNPal(4 * lIdx), aLPal(4 * aTrnIdx(lIdx)), 4)
                Next lIdx
                
                '-- Set as XOR DIB palette
                Call .FrameDIBXOR(nFrame).SetPalette(aNPal())
                        
                '-- Store temp. copy
                Call CopyMemory(ByVal .lpFramePalette(nFrame), aNPal(0), 1024)
                Let .LocalPaletteEntries(nFrame) = lMax + 1
            
                '-- Update indexes...
                For y = 0 To H
                    For x = 0 To W
                        aBits08(x, y) = aInvIdx(aBits08(x, y))
                    Next x
                Next y
                 
                '-- Update transparent index
                Let .FrameTransparentColorIndex(nFrame) = aInvIdx(.FrameTransparentColorIndex(nFrame))
            End If
            
            '-- Unmap DIB bits
            Call CopyMemory(ByVal VarPtrArray(aBits08()), 0&, 4)
            
            '-- Return removed entries
            OptimizeLocalPalette = (nBefore - .LocalPaletteEntries(nFrame))
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErrH:
    OptimizeLocalPalette = -1
    On Error GoTo 0
End Function

Public Function RemoveLocalPalettes(oGIF As cGIF, oProgress As ucProgress) As Integer

  Dim nLocalsOut     As Integer
  
  Dim lGPal(255)     As Long
  Dim aGPal(1023)    As Byte
  Dim lLPal(255)     As Long
  
  Dim nGMaxEnt       As Integer
  Dim lGEnt          As Long
  Dim lLEnt          As Long
  Dim bLEntExists()  As Boolean
  Dim bGEntChecked() As Boolean
  Dim nLCount        As Integer
  Dim aLTrnEnt()     As Byte
  
  Dim nTrnsIdx       As Integer
  Dim bTrnsExists    As Boolean
  
  Dim nCurr          As Integer
  Dim tSA08          As SAFEARRAY2D
  Dim aBits08()      As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    On Error GoTo ErrH
    
    With oGIF
        
        oProgress.Max = .FramesCount
        
        '-- Check if Global palette exists, store temp. copy and get max. index
        If (.GlobalPaletteExists) Then
            Call CopyMemory(lGPal(0), ByVal .lpGlobalPalette, 1024)
            nGMaxEnt = .GlobalPaletteEntries
          Else
            nGMaxEnt = 0
        End If
        
        For nCurr = 1 To .FramesCount
            
            '-- Current progress
            oProgress.Value = nCurr
            
            '-- Local palette used [?]
            If (.LocalPaletteUsed(nCurr)) Then
                
                '-- Store temp. copy (Long format)
                Call CopyMemory(lLPal(0), ByVal .lpFramePalette(nCurr), 1024)
                
                '-- Redim arrays
                ReDim bLEntExists(.LocalPaletteEntries(nCurr) - 1)
                ReDim aLTrnEnt(.LocalPaletteEntries(nCurr) - 1)
                ReDim bGEntChecked(nGMaxEnt)
                nLCount = 0
                
                '-- Transparent index/flag
                nTrnsIdx = IIf(.FrameUseTransparentColor(nCurr), .FrameTransparentColorIndex(nCurr), -1)
                
                '-- Check which Local colors exist in Global palette
                If (nGMaxEnt > 0) Then
                    For lLEnt = 0 To .LocalPaletteEntries(nCurr) - 1
                        For lGEnt = 0 To nGMaxEnt - 1
                            If (lLPal(lLEnt) = lGPal(lGEnt) And bGEntChecked(lGEnt) = False) Then
                                bGEntChecked(lGEnt) = True ' Global entry checked!
                                bLEntExists(lLEnt) = True  ' Color in Global
                                aLTrnEnt(lLEnt) = lGEnt    ' Translated index
                                nLCount = nLCount + 1
                                Exit For
                            End If
                        Next lGEnt
                    Next lLEnt
                End If
                
                '-- Check if we can store all Local entries in Global palette
                If (.LocalPaletteEntries(nCurr) - nLCount <= 256 - nGMaxEnt) Then
                
                    '-- Remove this Local palette
                    .LocalPaletteUsed(nCurr) = False: nLocalsOut = nLocalsOut + 1
                    
                    '-- Add colors to Global...
                    For lLEnt = 0 To .LocalPaletteEntries(nCurr) - 1
                        If (Not bLEntExists(lLEnt)) Then
                            lGPal(nGMaxEnt) = lLPal(lLEnt)
                            aLTrnEnt(lLEnt) = nGMaxEnt
                            nGMaxEnt = nGMaxEnt + 1
                        End If
                    Next lLEnt
                    If (nGMaxEnt > 256) Then nGMaxEnt = 256
                    
                    '-- Get dimensions
                    W = .FrameDIBXOR(nCurr).Width - 1
                    H = .FrameDIBXOR(nCurr).Height - 1
                    
                    '-- Map current 8-bpp DIB bits
                    Call pvBuild_08bppSA(tSA08, .FrameDIBXOR(nCurr))
                    Call CopyMemory(ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4)

                    '-- Update indexes...
                    For y = 0 To H
                        For x = 0 To W
                            aBits08(x, y) = aLTrnEnt(aBits08(x, y))
                        Next x
                    Next y

                    '-- Unmap DIB bits
                    Call CopyMemory(ByVal VarPtrArray(aBits08()), 0&, 4)
                    
                    '-- Update transparent index
                    Let .FrameTransparentColorIndex(nCurr) = aLTrnEnt(.FrameTransparentColorIndex(nCurr))
                        
                    '-- Update Global
                    Let .GlobalPaletteEntries = nGMaxEnt
                    Let .GlobalPaletteExists = True
                    Call CopyMemory(ByVal .lpGlobalPalette, lGPal(0), 1024)
               End If
            End If
        Next nCurr
        
        '-- Translate new Global palette to a byte array
        Call CopyMemory(aGPal(0), lGPal(0), 1024)
        
        '-- Update image palettes (using new Global)
        If (nLocalsOut > 0) Then
            For nCurr = 1 To .FramesCount
               If (Not .LocalPaletteUsed(nCurr)) Then
                   Let .LocalPaletteEntries(nCurr) = .GlobalPaletteEntries
                   Call .FrameDIBXOR(nCurr).SetPalette(aGPal())
                   Call CopyMemory(ByVal .lpFramePalette(nCurr), lGPal(0), 1024)
               End If
            Next nCurr
        End If
    End With
    
    '-- Return number of removed Locals
    oProgress.Value = 0
    RemoveLocalPalettes = nLocalsOut
    On Error GoTo 0
    Exit Function
    
ErrH:
    oProgress = 0
    RemoveLocalPalettes = -1
    On Error GoTo 0
End Function

Public Function OptimizeFrames(oGIF As cGIF, ByVal OptimizationMethod As eFrameOptimizationMethod, oProgress As ucProgress) As Boolean
 
  Dim tSAPrev       As SAFEARRAY2D
  Dim tSACurr       As SAFEARRAY2D
  Dim aBitsPrev()   As Byte
  Dim aBitsCurr()   As Byte
  
  Dim lPalPrev(255) As Long
  Dim lPalCurr(255) As Long
  Dim aPalXOR(1023) As Byte
  
  Dim bTrnsAdded    As Boolean
  Dim aTrnsAddedIdx As Byte
  Dim nTrnsIdxPrev  As Integer
  Dim nTrnsIdxCurr  As Integer
  Dim bIsTrnsCurr   As Boolean
  
  Dim xOffPrev      As Long
  Dim yOffPrev      As Long
  Dim xOffCurr      As Long
  Dim yOffCurr      As Long
  Dim xLUVPrev      As Long
  Dim yLUVPrev      As Long
  Dim xLUVCurr      As Long
  Dim yLUVCurr      As Long
                    
  Dim nPrev         As Integer
  Dim nCurr         As Integer
  Dim nNext         As Integer
  Dim nRect         As Integer
  Dim nCurrPal      As Integer
  Dim rInit()       As RECT2
  Dim rCrop()       As RECT2
  Dim bCrop()       As Boolean
  Dim rMerge        As RECT2
  
  Dim x As Long, xi As Long, xf As Long
  Dim y As Long, yi As Long, yf As Long
  
    On Error GoTo ErrH
  
    With oGIF
    
        '== Initialize
        
        '-- Redim. Crop rectangles arrays
        ReDim rInit(1 To .FramesCount)
        ReDim rCrop(1 To .FramesCount)
        ReDim bCrop(1 To .FramesCount)
        '-- Set max. prog.
        oProgress.Max = .FramesCount
        
        '== Define initial frame rects.
        
        For nCurr = 1 To .FramesCount
            Call SetRect(rInit(nCurr), .FrameLeft(nCurr), .FrameTop(nCurr), .FrameLeft(nCurr) + .FrameDIBXOR(nCurr).Width, .FrameTop(nCurr) + .FrameDIBXOR(nCurr).Height)
        Next nCurr
        
        '== Prepare transparency (if necessary)
        
        If (OptimizationMethod = [fomFrameDifferencing]) Then
            
            For nCurr = 2 To .FramesCount
            
                '-- Current progress
                oProgress.Value = nCurr
                nPrev = nCurr - 1
                
                '-- Current frame will 'remain' on screen
                If (.FrameDisposalMethod(nPrev) <> [dmRestoreToBackground] And .FrameDisposalMethod(nPrev) <> [dmRestoreToPrevious]) Then
                    
                    '-- Only scan common (intersected) rectangle
                    If (IntersectRect(rMerge, rInit(nCurr), rInit(nPrev))) Then
                
                        '-- Try to get a transparent entry...
                        If (Not .FrameUseTransparentColor(nCurr)) Then
                            
                            '-- Frame has Local palette
                            If (.LocalPaletteUsed(nCurr)) Then
                                
                                '-- Is there space for an extra entry
                                If (.LocalPaletteEntries(nCurr) < 256) Then
                                    
                                    '-- Add to palette a blank entry
                                    Let .LocalPaletteEntries(nCurr) = .LocalPaletteEntries(nCurr) + 1
                                    '-- Use it as transparent entry
                                    Let .FrameTransparentColorIndex(nCurr) = .LocalPaletteEntries(nCurr) - 1
                                    '-- And 'turn on' transparency
                                    Let .FrameUseTransparentColor(nCurr) = True
                                    
                                    '-- Restore/Set palette (un-mask frame)
                                    Call CopyMemory(aPalXOR(0), ByVal .lpFramePalette(nCurr), 1024)
                                    Call .FrameDIBXOR(nCurr).SetPalette(aPalXOR())
                                End If
                                
                            '-- Frame uses Global palette
                            Else
                            
                                '-- Is there space for an extra entry, or does already
                                '   exists (previously added) one?
                                If (.GlobalPaletteEntries() < 256 Or bTrnsAdded) Then
                                    
                                    If (bTrnsAdded) Then
                                    
                                        '-- A transparent entry already added. Use it
                                        Let .FrameTransparentColorIndex(nCurr) = aTrnsAddedIdx
                                        
                                      Else
                                        '-- No transparent entry. Add one...
                                        Let .GlobalPaletteEntries = .GlobalPaletteEntries + 1
                                        
                                        '-- This is the Global palette: update all temp. palette copies
                                        For nCurrPal = 1 To .FramesCount
                                            '-- Update number of entries
                                            If (Not .LocalPaletteUsed(nCurrPal)) Then
                                                Let .LocalPaletteEntries(nCurrPal) = .GlobalPaletteEntries
                                            End If
                                            '-- Restore/Set palette (un-mask frame)
                                            Call CopyMemory(aPalXOR(0), ByVal .lpGlobalPalette, 1024)
                                            Call .FrameDIBXOR(nCurr).SetPalette(aPalXOR())
                                        Next nCurrPal
                                        '-- Set new transparent entry
                                        Let .FrameTransparentColorIndex(nCurr) = .GlobalPaletteEntries - 1
                                        
                                        '-- Store this index. Enable 'bTrnsAdded' flag
                                        aTrnsAddedIdx = .GlobalPaletteEntries - 1
                                        bTrnsAdded = True
                                    End If
                                    
                                    '-- 'Turn on' transparency
                                    Let .FrameUseTransparentColor(nCurr) = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next nCurr
        End If
        
        '== Start removing redundant pixels...
        
        '-- From last to first to avoid interfering
        For nCurr = .FramesCount - 1 To 1 Step -1
            
            '-- Current progress
            oProgress.Value = .FramesCount - nCurr
            nNext = nCurr + 1
            
            '-- Get current intersection rectangle
            Call IntersectRect(rMerge, rInit(nCurr), rInit(nNext))
            
            '-- Initialize initial Crop rectangle
            Select Case OptimizationMethod
                Case [fomMinimumBoundingRectangle]: rCrop(nNext) = rInit(nNext)
                Case [fomFrameDifferencing]:        rCrop(nNext) = rMerge
            End Select
                
            '-- Current frame will 'remain' on screen
            If (.FrameDisposalMethod(nCurr) <> [dmRestoreToBackground] And .FrameDisposalMethod(nCurr) <> [dmRestoreToPrevious]) Then
                
                '-- Only scan common (intersected) rectangle
                If (IntersectRect(rMerge, rInit(nCurr), rInit(nNext))) Then
                    
                    '-- Map frame bits
                    Call pvBuild_08bppSA(tSAPrev, .FrameDIBXOR(nCurr))
                    Call CopyMemory(ByVal VarPtrArray(aBitsPrev()), VarPtr(tSAPrev), 4)
                    Call pvBuild_08bppSA(tSACurr, .FrameDIBXOR(nNext))
                    Call CopyMemory(ByVal VarPtrArray(aBitsCurr()), VarPtr(tSACurr), 4)
                    
                    '-- Store temp. copy of palettes (RGB Long format)
                    Call CopyMemory(lPalPrev(0), ByVal .lpFramePalette(nCurr), 1024)
                    Call CopyMemory(lPalCurr(0), ByVal .lpFramePalette(nNext), 1024)
            
                    '-- Get scan bounds and offsets, and transparent index/flag of both frame
                    xi = rMerge.x1
                    xf = rMerge.x2 - 1
                    yi = rMerge.y1
                    yf = rMerge.y2 - 1
                    xOffPrev = .FrameLeft(nCurr)
                    yOffPrev = .FrameTop(nCurr)
                    xOffCurr = .FrameLeft(nNext)
                    yOffCurr = .FrameTop(nNext)
                    nTrnsIdxPrev = IIf(.FrameUseTransparentColor(nCurr), .FrameTransparentColorIndex(nCurr), -1)
                    nTrnsIdxCurr = IIf(.FrameUseTransparentColor(nNext), .FrameTransparentColorIndex(nNext), -1)
                    
                    '-- Frame differencing
                    If (OptimizationMethod = [fomFrameDifferencing]) Then
                        
                        '-- Current frame has transparent entry. Can be processed:
                        If (.FrameUseTransparentColor(nNext)) Then
                            '-- Start scan...
                            For y = yi To yf
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev) Then
                                        If (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) = lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then aBitsCurr(xLUVCurr, yLUVCurr) = nTrnsIdxCurr
                                    End If
                                Next x
                            Next y
                        End If
                        
                    '-- Minimum Bounding Rectangle
                    Else
                        
                        '-- Previous frame contains current frame:
                        If (EqualRect(rMerge, rInit(nNext))) Then
                            
                            '-- Top:
Check_y1:                   For y = yi To yf
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nNext).y1 = y: GoTo Check_y2
                                Next x
                            Next y
                            Call SetRectEmpty(rCrop(nNext)): GoTo ScanDone
                            
Check_y2:                   '-- Bottom:
                            For y = yf To yi Step -1
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nNext).y2 = y + 1: GoTo Check_x1
                                Next x
                            Next y
                            Call SetRectEmpty(rCrop(nNext)): GoTo ScanDone
                            
Check_x1:                   yi = rCrop(nNext).y1
                            yf = rCrop(nNext).y2 - 1
                            
                            '-- Left:
                            For x = xi To xf
                                xLUVPrev = x - xOffPrev
                                xLUVCurr = x - xOffCurr
                                For y = yi To yf
                                    yLUVPrev = y - yOffPrev
                                    yLUVCurr = y - yOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nNext).x1 = x: GoTo Check_x2
                                Next y
                            Next x
                            Call SetRectEmpty(rCrop(nNext)): GoTo ScanDone
                            
Check_x2:                   '-- Right:
                            For x = xf To xi Step -1
                                xLUVPrev = x - xOffPrev
                                xLUVCurr = x - xOffCurr
                                For y = yi To yf
                                    yLUVPrev = y - yOffPrev
                                    yLUVCurr = y - yOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nNext).x2 = x + 1: GoTo ScanDone
                                Next y
                            Next x
                            Call SetRectEmpty(rCrop(nNext)): GoTo ScanDone
                            
ScanDone:                   '-- Need to crop [?]
                            bCrop(nNext) = Not (-EqualRect(rCrop(nNext), rMerge))
                        End If
                    End If
                    
                    '-- Unmap frame DIB bits
                    Call CopyMemory(ByVal VarPtrArray(aBitsPrev()), 0&, 4)
                    Call CopyMemory(ByVal VarPtrArray(aBitsCurr()), 0&, 4)
                End If
            End If
        Next nCurr
         
        '== Remove redundant frames...
        If (OptimizationMethod = [fomMinimumBoundingRectangle]) Then
        
            For nCurr = .FramesCount To 2 Step -1
            
                '-- Null rectangle [?]
                If (IsRectEmpty(rCrop(nCurr))) Then
    
                    If (.FrameDisposalMethod(nCurr) <> [dmRestoreToBackground] And .FrameDisposalMethod(nCurr) <> [dmRestoreToPrevious]) Then
                        
                        '-- Add delay time of removed frame to previous one
                        Let .FrameDelay(nCurr - 1) = .FrameDelay(nCurr - 1) + .FrameDelay(nCurr)
                        '-- Remove redundant frame
                        Call .FrameRemove(nCurr)
                        
                        '-- Update temp. bounding rectangles array
                        For nRect = nCurr To .FramesCount
                            rCrop(nRect) = rCrop(nRect + 1)
                            bCrop(nRect) = bCrop(nRect + 1)
                        Next nRect
                    
                      Else
                        bCrop(nCurr) = False
                    End If
                End If
            Next nCurr
        End If
        
        '== Crop frames...
        If (OptimizationMethod = [fomMinimumBoundingRectangle]) Then
            For nCurr = 2 To .FramesCount
                If (bCrop(nCurr)) Then
                    Call OffsetRect(rCrop(nCurr), -.FrameLeft(nCurr), -.FrameTop(nCurr))
                    Call pvCropFrame(oGIF, nCurr, rCrop(nCurr))
                End If
            Next nCurr
        End If

        '== Success
        OptimizeFrames = True
    End With
    
ErrH:
    oProgress = 0
    On Error GoTo 0
End Function

Public Function CropTransparentImages(oGIF As cGIF, oProgress As ucProgress) As Boolean

  Dim nPrev       As Integer
  Dim nCurr       As Integer
  Dim tSA         As SAFEARRAY2D
  Dim aBits()     As Byte
  Dim aTrnsIdx    As Byte
  
  Dim nRect       As Integer
  Dim rInit()     As RECT2
  Dim rCrop()     As RECT2
  Dim bCrop()     As Boolean
  
  Dim x As Long, xi As Long, xf As Long
  Dim y As Long, yi As Long, yf As Long
    
    On Error GoTo ErrH
    
    With oGIF
    
        '== Initialize
        '-- Redim. Crop rectangles arrays
        ReDim rInit(1 To .FramesCount)
        ReDim rCrop(1 To .FramesCount)
        ReDim bCrop(1 To .FramesCount)
        '-- Set max. prog.
        oProgress.Max = .FramesCount
        
        '== Define initial frame rects.
        For nCurr = 1 To .FramesCount
            Call SetRect(rInit(nCurr), 0, 0, .FrameDIBXOR(nCurr).Width, .FrameDIBXOR(nCurr).Height)
        Next nCurr
        
        '== Calc. minimum bounding rectangles
        For nCurr = 1 To .FramesCount
            
            '-- Current progress
            oProgress.Value = nCurr
            
            '-- Initialize crop rect.
            Call CopyRect(rCrop(nCurr), rInit(nCurr))
            
            '-- Only try to crop transparent frames
            If (.FrameUseTransparentColor(nCurr)) Then
            
                '-- Map frame bits
                Call pvBuild_08bppSA(tSA, .FrameDIBXOR(nCurr))
                Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(tSA), 4)
            
                '-- Get bounds
                xi = 0: xf = .FrameDIBXOR(nCurr).Width - 1
                yi = 0: yf = .FrameDIBXOR(nCurr).Height - 1
                
                '-- Transparent color index
                aTrnsIdx = .FrameTransparentColorIndex(nCurr)
                
                '-- Top:
                For y = yi To yf
                    For x = xi To xf
                        If (aBits(x, y) <> aTrnsIdx) Then rCrop(nCurr).y1 = y: GoTo Check_y2
                    Next x
                Next y
                Call SetRectEmpty(rCrop(nCurr)): GoTo ScanDone
                
Check_y2:       '-- Bottom:
                For y = yf To yi Step -1
                    For x = xi To xf
                        If (aBits(x, y) <> aTrnsIdx) Then rCrop(nCurr).y2 = y + 1: GoTo Check_x1
                    Next x
                Next y
                Call SetRectEmpty(rCrop(nCurr)): GoTo ScanDone
                
Check_x1:       yi = rCrop(nCurr).y1
                yf = rCrop(nCurr).y2 - 1
                            
                '-- Left:
                For x = xi To xf
                    For y = yi To yf
                        If (aBits(x, y) <> aTrnsIdx) Then rCrop(nCurr).x1 = x: GoTo Check_x2
                    Next y
                Next x
                Call SetRectEmpty(rCrop(nCurr)): GoTo ScanDone
                
Check_x2:       '-- Right:
                For x = xf To xi Step -1
                    For y = yi To yf
                        If (aBits(x, y) <> aTrnsIdx) Then rCrop(nCurr).x2 = x + 1: GoTo ScanDone
                    Next y
                Next x
                Call SetRectEmpty(rCrop(nCurr)): GoTo ScanDone
                
ScanDone:       '-- Scan done
                Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
               
                '-- Need to crop [?]
                bCrop(nCurr) = Not -EqualRect(rCrop(nCurr), rInit(nCurr))
            End If
        Next nCurr
        
        '== Remove redundant frames...
        For nCurr = .FramesCount To 2 Step -1
        
            '-- Null rectangle [?]
            If (IsRectEmpty(rCrop(nCurr)) <> 0) Then

                If (.FrameDisposalMethod(nCurr) <> [dmRestoreToBackground] And .FrameDisposalMethod(nCurr) <> [dmRestoreToPrevious]) Then

                    '-- Add delay time of removed frame to previous one
                    Let .FrameDelay(nCurr - 1) = .FrameDelay(nCurr - 1) + .FrameDelay(nCurr)
                    '-- Remove redundant frame
                    Call .FrameRemove(nCurr)
                    
                    '-- Update temp. bounding rectangles array
                    For nRect = nCurr To .FramesCount
                        rCrop(nRect) = rCrop(nRect + 1)
                        bCrop(nRect) = bCrop(nRect + 1)
                    Next nRect
                  
                  Else
                    bCrop(nCurr) = False
                End If
            End If
        Next nCurr
        
        '== Crop frames (skip first frame if NULL)...
        For nCurr = 1 To .FramesCount
            If (bCrop(nCurr) And (IsRectEmpty(rCrop(nCurr)) = 0)) Then
                Call pvCropFrame(oGIF, nCurr, rCrop(nCurr))
            End If
        Next nCurr
        
        '== Success
        CropTransparentImages = True
    End With
    
ErrH:
    oProgress.Value = 0
    On Error GoTo 0
End Function

Public Function RemaskFrames(oGIF As cGIF, oProgress As ucProgress) As Boolean
  
  Dim nFrm As Integer
  
    On Error GoTo ErrH
    
    With oGIF
    
        oProgress.Max = .FramesCount
        
        '-- Re-mask frames
        For nFrm = 1 To .FramesCount
            oProgress.Value = nFrm
            Call .FrameMask(nFrm, .FrameTransparentColorIndex(nFrm))
        Next nFrm
        '-- Success
        RemaskFrames = True
    End With
    
ErrH:
    oProgress.Value = 0
    On Error GoTo 0
End Function

'========================================================================================
' Private
'========================================================================================

Private Function pvCropFrame(oGIF As cGIF, ByVal nFrame As Integer, rCrop As RECT2) As Boolean

  Dim oDIBBuff      As New cDIB
  Dim lpBitsDst     As Long
  Dim lpBitsSrc     As Long
  Dim lWidthDst     As Long
  Dim lScanWidthDst As Long
  Dim lScanWidthSrc As Long
  Dim y             As Long
  
  Dim aPalXOR(1023) As Byte

    On Error GoTo ErrH
    
    With oGIF

        '-- Prepare palette
        Call CopyMemory(aPalXOR(0), ByVal .lpFramePalette(nFrame), 1024)
        
        '-- XOR DIB
        With .FrameDIBXOR(nFrame)
            
            '-- Make a temp. copy
            Call .CloneTo(oDIBBuff)
            '-- Resize current frame
            Call .Create(rCrop.x2 - rCrop.x1, rCrop.y2 - rCrop.y1, [08_bpp])
            Call .SetPalette(aPalXOR())
            
            '-- Get some props. for cropping process
            lpBitsDst = .lpBits
            lpBitsSrc = oDIBBuff.lpBits
            lScanWidthDst = .BytesPerScanline
            lScanWidthSrc = oDIBBuff.BytesPerScanline
            lWidthDst = .Width
            
            '-- Crop...
            For y = rCrop.y1 To rCrop.y2 - 1
                Call CopyMemory(ByVal lpBitsDst + lScanWidthDst * (y - rCrop.y1), ByVal lpBitsSrc + (y * lScanWidthSrc) + rCrop.x1, lWidthDst)
            Next y
        End With
        
        '-- AND DIB
        With .FrameDIBAND(nFrame)
            Call .Create(rCrop.x2 - rCrop.x1, rCrop.y2 - rCrop.y1, [01_bpp])
        End With
        
        '-- Update frame position
        Let .FrameLeft(nFrame) = .FrameLeft(nFrame) + rCrop.x1
        Let .FrameTop(nFrame) = .FrameTop(nFrame) + rCrop.y1
    End With
    
    '-- Success
    pvCropFrame = True
    
ErrH:
    On Error GoTo 0
End Function

'//

Private Sub pvBuild_08bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 8-bpp DIB mapping
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
End Sub
