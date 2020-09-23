Attribute VB_Name = "PaintFunctions"
'''''''''''''''''''''''''''''''''''''''''''''''''
'''            Graphix  Functions             '''
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
'Timer procedure halts operation [vbmodal]
''''''''''''''Bit Blt Function''''''''''''''''''''
Public Declare Function BitBlt Lib "gdi32" _
       (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, _
       ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
       ByVal dwRop As Long) As Long

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046
'Draws an image to a pictuer box.

'Inputs:
'hDestDC = The .hDC (Device context) property of the picturebox/Form which is the destination of the image
'nWidth = The width of the image
'n height = The height of the image
'hSrcDC = The .hdc property of the picture which is the source of the image
'XSrc = If you only want a part of the image to be drawn use this to specify it's X location
'YSrc = If you only want a part of the image to be drawn use this to specify it's Y location
'dwRop = The Raster Operation to preform (Use Copy)

'Returns:
'Sets the Destination's Image (NOT picture) property as the icon you specify

'Tip:
'A conbination of Src and of a mask and ontop of it an src Invert
'Makes the image transparant based on the mask

'Limitations:
'The Source & destination Control MUST have
'Auto Redraw = True
'Scale mode = 3 'VbPixels
'Use Destination.Refresh to make the image drawn visible.
'''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Sterch Blt Function''''''''''''''''
Public Declare Function StretchBlt Lib "gdi32" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''
'Same as BitBlt but with the ability to strech the image
''''''''''''Transparent Blt Function'''''''''''''
Public Declare Function TransparentBlt Lib "msimg32.dll" _
   (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'Same as StrechBlt with the ability of specifing a color as
'Transparent to the crTransparent variable (RGB(0,255,0))
'Remarks:
'The TransparentBlt function is supported
'for source bitmaps of 4 bits per pixel
'and 8 bits per pixel
'''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''Alpha Blend Function''''''''''''''
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" _
    (ByVal hdc As Long, ByVal lInt As Long, _
    ByVal lInt As Long, ByVal lInt As Long, _
    ByVal lInt As Long, ByVal hdc As Long, _
    ByVal lInt As Long, ByVal lInt As Long, _
    ByVal lInt As Long, ByVal lInt As Long, _
    ByVal BLENDFUNCT As Long) As Long
'See the assosiated subroutine "Alpha_Blend" below

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
(Destination As Any, Source As Any, ByVal Length As Long)
'copy a part of memory to a variable

'_______________________________________________


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Alpha_Blend(DestinationHdc&, SourceHdc&, _
           X&, Y&, CropX&, CropY&, CropW&, CropH&, _
           Width&, Height&, Optional TransparancyLevel& = 128)
   Dim BF As BLENDFUNCTION, lBF As Long
   Const AC_SRC_OVER = 0
    'set the parameters
    With BF
        .BlendOp = AC_SRC_OVER '0
        .BlendFlags = 0
        .SourceConstantAlpha = TransparancyLevel&
        .AlphaFormat = AC_SRC_ALPHA
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend DestinationHdc&, X&, Y&, Width&, Height&, SourceHdc&, CropX&, CropY&, CropW&, CropH&, lBF
    'If even one pixel of the range you specify,
    'doesn't exist in the pictuer you give it then
    'this function doesn't work
    '
End Sub
'Draws a semi-transparent image
'TransparancyLevel& is the ALPHA of the image
'BTW: WTF is ALPHA
'(Other than the 1st letter of the greek alphabet)
'''''''''''''''''''''''''''''''''''''''''''''''
