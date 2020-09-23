Attribute VB_Name = "bitBltMAsk"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub TransBLT(SrcHDC As Long, xSrc As Long, ySrc As Long, MaskHDC As Long, ByVal Xmsk As Long, ByVal Ymsk As Long, width As Long, height As Long, DestHDC As Long, x As Long, y As Long, Optional UseSrcAsMask As Boolean)
    Const SRCPAINT = &HEE0086
    If UseSrcAsMask Then Xmsk = xSrc: Ymsk = ySrc
    BitBlt DestHDC, x, y, width, height, MaskHDC, Xmsk, Ymsk, SRCPAINT
    BitBlt DestHDC, x, y, width, height, SrcHDC, xSrc, ySrc, vbSrcAnd
End Sub

Public Sub TransFlipBlt(SrcHDC As Long, MaskHDC As Long, xSrc As Long, ySrc As Long, width As Long, height As Long, DestHDC As Long, x As Long, y As Long, Optional UseSrcAsMask As Boolean, Optional FlipX As Boolean, Optional FlipY As Boolean, Optional ByVal Xmsk As Long, Optional ByVal Ymsk As Long)
    Const SRCPAINT = &HEE0086
    If UseSrcAsMask Then Xmsk = xSrc: Ymsk = ySrc
    FlipBlt MaskHDC, Xmsk, Ymsk, width, height, DestHDC, x, y, FlipX, FlipY, SRCPAINT
    FlipBlt SrcHDC, xSrc, ySrc, width, height, DestHDC, x, y, FlipX, FlipY, vbSrcAnd
End Sub

Public Sub MakeMask(SrcHDC As Long, InvHdc As Long, MskHdc, x As Long, y As Long, width As Long, height As Long, TransParent As Long)
    CreateMask InvHdc, x, y, width, height, SrcHDC, x, y, TransParent
    BitBlt MskHdc, x, y, width, height, InvHdc, x, y, vbSrcInvert
End Sub
Public Sub CreateMask(hDestDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hSrcDC As Long, xSrc As Long, ySrc As Long, TransColor As Long)
    Dim OrigColor As Long ' Holds source original background color
    Dim DestBKColor As Long ' Holds destination original background color
    Dim OrigTextColor As Long
    Dim hMaskBmp As Long
    Dim hMaskPrevBmp As Long
    Dim MaskDC As Long    'Masks must be inverted
    
    MaskDC = CreateCompatibleDC(hDestDC)
    hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&) '//Create a monochrome bitmap For our mask
    hMaskPrevBmp = SelectObject(MaskDC, hMaskBmp)
    
    OrigColor = SetBkColor(hSrcDC, TransColor)
    Call BitBlt(MaskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy) '//Copy hSrcDc into our mask bitmap
    SetBkColor hSrcDC, OrigColor '//Restore the original color
    
    DestBKColor = SetBkColor(hDestDC, vbWhite) '//All the white In our bitmap hasto be white
    OrigTextColor = SetTextColor(hDestDC, vbBlack)
    BitBlt hDestDC, x, y, nWidth, nHeight, MaskDC, 0, 0, vbSrcCopy
    SetTextColor hDestDC, OrigTextColor
    SetBkColor hDestDC, DestBKColor '//Restore the original back color bak
    
    Call SelectObject(MaskDC, hMaskPrevBmp) 'Select our original bitmap bak
    Call DeleteObject(hMaskBmp) 'Delete our mask bitmap
    Call DeleteDC(MaskDC) 'Delete MaskDC
End Sub

Public Sub FlipBlt(SrcHDC As Long, SrcX As Long, SrcY As Long, width As Long, height As Long, DestHDC As Long, DestX As Long, DestY As Long, Optional FlipX As Boolean, Optional FlipY As Boolean, Optional dwRop As Long = vbSrcCopy)
    Dim temp As Long, temp2 As Long
    If Not FlipX And Not FlipY Then BitBlt DestHDC, DestX, DestY, width, height, SrcHDC, SrcX, SrcY, dwRop 'Don't flip along X or Y
    If FlipX And Not FlipY Then 'Flip along X
        For temp = 0 To width - 1
            BitBlt DestHDC, DestX + temp, DestY, 1, height, SrcHDC, SrcX + width - 1 - temp, SrcY, dwRop
        Next
    End If
    If Not FlipX And FlipY Then 'Flip along Y
        For temp = 0 To height - 1
            BitBlt DestHDC, DestX, DestY + temp, width, 1, SrcHDC, SrcX, SrcY + height - 1 - temp, dwRop
        Next
    End If
    If FlipX And FlipY Then 'Flip along X and Y
        For temp = 0 To width - 1
            For temp2 = 0 To height - 1
                BitBlt DestHDC, DestX + temp, DestY + temp2, 1, 1, SrcHDC, SrcX + width - 1 - temp, SrcY + height - 1 - temp2, dwRop
            Next
        Next
    End If
End Sub

Public Function SysToLNG(ByVal lColor As Long) As Long 'Special thanks to redbird77 for this code and realizing what the bug was
If (lColor And &H80000000) Then SysToLNG = GetSysColor(lColor And &HFFFFFF) Else SysToLNG = lColor ' If hi-bit is set, then it is a system color.
End Function

Public Function Red(color As Long)
    Red = color Mod 256
End Function

Public Function Green(color As Long)
    Green = ((color And &HFF00) / 256) Mod 256
End Function

Public Function Blue(color As Long)
    Blue = (color And &HFF0000) / 65536
End Function

Public Function AlphaBlend(colorA As Long, colorB As Long, Alpha As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = Blend(Red(colorA), Red(colorB), Alpha)
    g = Blend(Green(colorA), Green(colorB), Alpha)
    b = Blend(Blue(colorA), Blue(colorB), Alpha)
    AlphaBlend = RGB(r, g, b)
End Function

Public Function Blend(colorA As Long, colorB As Long, Alpha As Double) As Long
    Blend = Abs((colorA - colorB) * Alpha + colorB) Mod 256
End Function
