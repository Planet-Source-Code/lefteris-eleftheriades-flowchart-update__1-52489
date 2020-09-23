Attribute VB_Name = "FormShaper"
Global Const winding = 2
Global Const alternate = 1
' CombineRgn() Styles
Public Const RGN_AND = 1 'Shows the part when both regions are touched
Public Const RGN_OR = 2 'Shows the part when one or both regions are touched
Public Const RGN_XOR = 3 'Shows the part when one of both regions are touched
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const RGN_MIN = RGN_AND
Public Const RGN_MAX = RGN_COPY

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum RgnCombStyle
  RCS_AND = 1 'Shows the part when both regions are touched
  RCS_OR = 2 'Shows the part when one or both regions are touched
  RCS_XOR = 3 'Shows the part when one of both regions are touched
  RCS_diff = 4
  RCS_COPY = 5
End Enum

Public Enum SScaleMode
 S_Twip = 1
 S_Point = 2
 S_Pixel = 3
 S_Inch = 5
 S_Milimeter = 6
 S_Centimeter = 7
End Enum

Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long 'The only difference from CreateRectRgn is it is destinated thru a RECT variable
Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Declare Function CreatePolyPolygonRgn& Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyfillMode As Long, lpPolyCount As Long)
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'—————————————————————————————————————————————————————
'  COOL FUNCTION ... Shape To Region ... COOL FUNCTION
'  You draw a shape on the form, call this function and
'  The shape data will be used to create a region, to
'  later shape the form in one Shape or a combination
'  of them.
'————————————————————————————————————————————————————

Public Function ShapeToRegion(Shapei As Object, frm As Object) As Long
Dim RRound As Long
If TypeOf Shapei Is Shape Then

  Select Case Shapei.Shape
     Case 0
           'Rectangle
           ShapeToRegion = CreateRectRgn(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel))
     Case 1
           'Squair
           'TODO: Add some center FX
           ShapeToRegion = CreateRectRgn(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel))
                
     Case 2
            'Oval
            ShapeToRegion = CreateEllipticRgn&(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel))
     Case 3
            'Circle
            'TODO: Add some center FX
            ShapeToRegion = CreateEllipticRgn&(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel))
     Case 4
           'Rounded Rectangle
           'Things are getting more and more complicated
           'The arc of the curve is 25% of the smallest dimension between Width and height
           
           If Shapei.Width < Shapei.Height Then
              RRound = StS(Shapei.Width, frm.ScaleMode, S_Pixel) / 4
           Else
              RRound = StS(Shapei.Height, frm.ScaleMode, S_Pixel) / 4
           End If
           ShapeToRegion = CreateRoundRectRgn(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel), RRound, RRound)
     Case 5
           'Rounded squair
           'TODO: Add some center FX
           'The arc of the curve is 25% of the smallest dimension between Width and height
           
           If Shapei.Width < Shapei.Height Then
              RRound = (Shapei.Width / 15) / 4
           Else
              RRound = (Shapei.Height / 15) / 4
           End If
        
           ShapeToRegion = CreateRoundRectRgn(StS(Shapei.Left, frm.ScaleMode, S_Pixel), StS(Shapei.Top, frm.ScaleMode, S_Pixel), StS(Shapei.Left + Shapei.Width, frm.ScaleMode, S_Pixel), StS(Shapei.Top + Shapei.Height, frm.ScaleMode, S_Pixel), RRound, RRound)
  End Select
End If
End Function

Public Function ModCombineRegion(Region1&, Region2&, Style As RgnCombStyle, DeleteSource As Boolean, frm As Object) As Long
Dim ROut&
  ROut& = CreateRectRgn(0, 0, StS(frm.Width, frm.ScaleMode, S_Pixel), StS(frm.Height, frm.ScaleMode, S_Pixel))
  CombineRgn ROut&, Region1&, Region2&, Style
  ModCombineRegion = ROut&
  If DeleteSource Then
     DeleteObject Region1&
     DeleteObject Region2&
  End If
End Function

Public Function StS(Value As Double, ScaleFrom As SScaleMode, ScaleTo As SScaleMode) As Double
' —————————————————————————————————
'|Any Scale Mode To Any Scale Mode |
' —————————————————————————————————
  Dim ScaleFromValue As Long
  Dim SMFDec As Long
  Dim SMTDec As Long
  
  Const TwipsPerPointXY = 20
  Const TwipsPerPixelXY = 15
  Const TwipsPerCharacterX = 120 '\
  Const TwipsPerCharacterY = 240 '/
  Const TwipsPerInchXY = 1439
  Const TwipsPerMilimeterXY = 56.7
  Const TwipsPerCentimeterXY = 567
  
  Select Case ScaleFrom
         Case 1: SMFDec = 1
         Case 2: SMFDec = TwipsPerPointXY
         Case 3: SMFDec = TwipsPerPixelXY
         Case 5: SMFDec = TwipsPerInchXY
         Case 6: SMFDec = TwipsPerMilimeterXY
         Case 7: SMFDec = TwipsPerCentimeterXY
         Case Else: MsgBox "Error", vbCritical
  End Select
  
  Select Case ScaleTo
         Case 1: SMTDec = 1
         Case 2: SMTDec = TwipsPerPointXY
         Case 3: SMTDec = TwipsPerPixelXY
         Case 5: SMTDec = TwipsPerInchXY
         Case 6: SMTDec = TwipsPerMilimeterXY
         Case 7: SMTDec = TwipsPerCentimeterXY
         Case Else: MsgBox "Error", vbCritical
  End Select
  StS = (Value / SMTDec) * SMFDec
End Function
