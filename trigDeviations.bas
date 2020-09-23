Attribute VB_Name = "trigDeviations"
'Taken from: mk:@MSITStore:C:\Program Files\Microsoft Visual Studio\MSDN98\98VSa\1033\office95.chm::/html/S11624.HTM
Public Function Arcsin(X#) As Double
    If Abs(X) = 1 Then
        Arcsin = X * 1.5707963267949
    Else
        Arcsin = Atn(X / Sqr(-X * X + 1))
    End If
End Function

Public Function Arccos(X#) As Double
    If X = -1 Then
        Arccos = 3.14159265359
    ElseIf X = 1 Then
        Arccos = 0
    Else
        Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End If
End Function

'It's true you can know if a line is clicked, by calling this function on its container control
Public Function LineMouseEvent(LineName As Line, X As Single, Y As Single, Optional Aura% = 0) As Boolean
  'You are propably wandering wtf is aura. if Aura is big then the line will be considered as clicked even if it is not exactly clicked on. if it's zero the event will only trigger on exact click
  On Error GoTo ExitF
  Dim XMin&, XMax&, YMin&, YMax&, Gradient As Currency, HalfBordWid&
  If LineName.X1 < LineName.X2 Then
     XMin = LineName.X1
     XMax = LineName.X2
  Else
     XMin = LineName.X2
     XMax = LineName.X1
  End If
  If LineName.Y1 < LineName.Y2 Then
     YMin = LineName.Y1
     YMax = LineName.Y2
  Else
     YMin = LineName.Y2
     YMax = LineName.Y1
  End If
  HalfBordWid = (LineName.BorderWidth + Aura) / 2
  If X >= XMin - HalfBordWid And X <= XMax + HalfBordWid And Y >= YMin - HalfBordWid And Y <= YMax + HalfBordWid Then
       'calculate the line vector equation and check the Y values
       If LineName.X2 - LineName.X1 = 0 Then
          LineMouseEvent = True
       Else
          Gradient = (LineName.Y2 - LineName.Y1) / (LineName.X2 - LineName.X1)
          'Line Equation is: Y - LineName.Y2 = Gradient * (X - LineName.X2)
           LineMouseEvent = CBool(Abs(Gradient * (X - LineName.X2) - (Y - LineName.Y2)) < (LineName.BorderWidth + Aura))
       End If
  End If
ExitF:
End Function

