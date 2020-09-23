Attribute VB_Name = "DIBits"

Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_PAL_COLORS = 1
Public Const DIB_PAL_INDICES = 2
Public Const DIB_PAL_LOGINDICES = 4
Public Const DIB_PAL_PHYSINDICES = 2
Public Const DIB_RGB_COLORS = 0
Public Const SRCCOPY = &HCC0020
Public Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

Public Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Bits() As Byte             '(Colors)
End Type

'What is returned is srcBM which holds the bitmap's 'header and R,G,B,A data given by srcBits
Public Sub CreateBitmap(ByRef srcBM As BITMAPINFO, ByRef srcBits() As Byte, ByVal W&, ByVal H&)
    ReDim srcBM.Bits(3, W - 1, H - 1) As Byte
    With srcBM.Header
        .biSize = 40
        .biBitCount = 32 '(R,G,B,Alpha - Reserved)
        .biPlanes = 1
        .biWidth = W
        .biHeight = -H
    End With
    srcBM.Bits = srcBits
End Sub
Public Sub LoadBitmapInformation(ByRef Bits() As Byte, ByVal srcWidth&, ByVal srcHeight&, ByVal PictureHandle&)
  Dim HanlingDC&
  Dim Bitmap As BITMAPINFO
  HanlingDC = CreateCompatibleDC(0)
  If (HanlingDC <> 0) Then
    'If our context was created
    SelectObject HanlingDC, PictureHandle
    CreateBitmap Bitmap, Bitmap.Bits, srcWidth, srcHeight
    'Bitmap.Bits given is all Zero (Black Image) 3D Array
    GetDIBits HanlingDC, PictureHandle, 0, srcHeight, Bitmap.Bits(0, 0, 0), Bitmap, DIB_RGB_COLORS
    DeleteObject HanlingDC
    Bits = Bitmap.Bits
    Erase Bitmap.Bits
  End If
End Sub
Public Sub DrawDIB(DestContext&, Left&, Top&, Width&, Height&, SourceLeft&, SourceRight&, SourceWidth&, SourceHeight&, ByRef Bits() As Byte)
Dim TempBitmap As BITMAPINFO
  CreateBitmap TempBitmap, Bits, Width&, Height&
  StretchDIBits DestContext&, Left&, Top&, Width&, Height&, SourceLeft&, SourceRight&, SourceWidth&, SourceHeight&, Bits(0, 0, 0), TempBitmap, DIB_RGB_COLORS, SRCCOPY
End Sub

