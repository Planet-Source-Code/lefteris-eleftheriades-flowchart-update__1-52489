VERSION 5.00
Begin VB.UserControl FlowShape 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   114
End
Attribute VB_Name = "FlowShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum sShapes
   Oval = 0
   Rectangle = 1
   Parallel = 2
   Lozenge = 3
   Diamond = 4
   DoubleRectangle = 5
End Enum
'Default Property Values:
Const m_def_Shape = 0
'Property Variables:
Dim m_Shape As Integer
Dim m_Caption As String
Dim m_Caption2 As String
Dim mControlData As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Dim QSS(3) As POINTAPI
Dim QSS2(5) As POINTAPI

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    ShapeControl
    PropertyChanged "Caption"
End Property

Public Property Get Title() As String
    Title = m_Caption2
End Property

Public Property Let Title(ByVal New_Title As String)
    m_Caption2 = New_Title
    RefreshCaption
    PropertyChanged "Title"
End Property

Private Sub RefreshCaption()
    'UserControl.Cls
    If m_Caption2 = "" Or m_Caption = "" Then
       UserControl.CurrentY = (UserControl.Height / 15 - UserControl.TextHeight(m_Caption)) / 2
    Else
       UserControl.CurrentY = (UserControl.Height / 15 - UserControl.TextHeight(m_Caption) - UserControl.TextHeight(m_Caption2)) / 2
    End If
    
    If m_Caption2 <> "" Then
      UserControl.CurrentX = (UserControl.Width / 15 - UserControl.TextWidth(m_Caption2)) / 2
      UserControl.Print m_Caption2
    End If
    UserControl.CurrentX = (UserControl.Width / 15 - UserControl.TextWidth(m_Caption)) / 2
    UserControl.Print m_Caption
    'Polygon UserControl.hdc, QSS(0), 4
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    RefreshCaption
End Property

Public Property Get ControlData() As String
    ControlData = mControlData
End Property

Public Property Let ControlData(ByVal New_ControlData As String)
    mControlData = New_ControlData
    PropertyChanged "ControlData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = UserControl.FontTransparent
    RefreshCaption
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    UserControl.FontTransparent() = New_FontTransparent
    PropertyChanged "FontTransparent"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'The Underscore following "Circle" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    UserControl.Circle (X, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As Long
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
    Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    RefreshCaption
End Property

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    RefreshCaption
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Palette
Public Property Get Palette() As Picture
Attribute Palette.VB_Description = "Returns/sets an image that contains the palette to use on an object when PaletteMode is set to Custom"
    Set Palette = UserControl.Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
    Set UserControl.Palette = New_Palette
    PropertyChanged "Palette"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaletteMode
Public Property Get PaletteMode() As Integer
Attribute PaletteMode.VB_Description = "Returns/sets a value that determines which palette to use for the controls on a object."
    PaletteMode = UserControl.PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As Integer)
    UserControl.PaletteMode() = New_PaletteMode
    PropertyChanged "PaletteMode"
    RefreshCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    RefreshCaption
End Property

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point = UserControl.Point(X, Y)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

'The Underscore following "PSet" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    UserControl.PSet Step(X, Y), Color
End Sub

Private Sub UserControl_Resize()
    ShapeControl
    RefreshCaption
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
    TextWidth = UserControl.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Shape() As sShapes
    Shape = m_Shape
End Property

Public Property Let Shape(ByVal New_Shape As sShapes)
    m_Shape = New_Shape
    PropertyChanged "Shape"
    ShapeControl
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Shape = m_def_Shape
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Caption = PropBag.ReadProperty("Caption", "")
    m_Caption2 = PropBag.ReadProperty("Title", "")
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 8)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Shape = PropBag.ReadProperty("Shape", m_def_Shape)
    mControlData = PropBag.ReadProperty("ControlData", "")
    ShapeControl
    RefreshCaption
End Sub
Public Sub ShapeControl()
    UserControl.DrawWidth = 2
    UserControl.Cls
    Select Case m_Shape
     'Oval
     Case 0
            SetWindowRgn UserControl.hwnd, CreateEllipticRgn&(0, 0, 0 + UserControl.Width / 15, 0 + UserControl.Height / 15), True
            DoEvents
            UserControl.Circle (UserControl.Width / 30 - 1, UserControl.Height / 30 - 1), UserControl.Width / 30 - 1, , , , ((UserControl.Height) / (UserControl.Width + 60))
     'Rectangle
     Case 1, 5
            QSS(0).X = 0: QSS(0).Y = 0
            QSS(1).X = UserControl.Width / 15: QSS(1).Y = 0
            QSS(2).X = UserControl.Width / 15: QSS(2).Y = UserControl.Height / 15
            QSS(3).X = 0: QSS(3).Y = UserControl.Height / 15
            SetWindowRgn UserControl.hwnd, CreatePolygonRgn(QSS(0), 4, RGN_OR), True
            DoEvents
            Polygon UserControl.hdc, QSS(0), 4
     'Parallel
     Case 2
            QSS(0).X = UserControl.Height / 30: QSS(0).Y = 0
            QSS(1).X = UserControl.Width / 15: QSS(1).Y = 0
            QSS(2).X = UserControl.Width / 15 - UserControl.Height / 30: QSS(2).Y = UserControl.Height / 15
            QSS(3).X = 0: QSS(3).Y = UserControl.Height / 15
            SetWindowRgn UserControl.hwnd, CreatePolygonRgn(QSS(0), 4, RGN_OR), True
            DoEvents
            Polygon UserControl.hdc, QSS(0), 4
     'Lozenge
     Case 3
            QSS(0).X = UserControl.Width / 30: QSS(0).Y = 0
            QSS(1).X = UserControl.Width / 15: QSS(1).Y = UserControl.Height / 30
            QSS(2).X = UserControl.Width / 30: QSS(2).Y = UserControl.Height / 15 - 1
            QSS(3).X = 0: QSS(3).Y = UserControl.Height / 30
            SetWindowRgn UserControl.hwnd, CreatePolygonRgn(QSS(0), 4, RGN_OR), True
            DoEvents
            Polygon UserControl.hdc, QSS(0), 4
     'Diamond
     Case 4
            QSS2(0).X = UserControl.Height / 30: QSS2(0).Y = 0
            QSS2(1).X = UserControl.Width / 15 - UserControl.Height / 30: QSS2(1).Y = 0
            QSS2(2).X = UserControl.Width / 15: QSS2(2).Y = UserControl.Height / 30
            QSS2(3).X = UserControl.Width / 15 - UserControl.Height / 30: QSS2(3).Y = UserControl.Height / 15
            QSS2(4).X = UserControl.Height / 30: QSS2(4).Y = UserControl.Height / 15
            QSS2(5).X = 0: QSS2(5).Y = UserControl.Height / 30
            SetWindowRgn UserControl.hwnd, CreatePolygonRgn(QSS2(0), 6, RGN_OR), True
            DoEvents
            Polygon UserControl.hdc, QSS2(0), 6
    End Select
    If m_Shape = 5 Then
        UserControl.DrawWidth = 1
        UserControl.Line (UserControl.Width / 15 - UserControl.Height / 30, 0)-(UserControl.Width / 15 - UserControl.Height / 30, UserControl.Height / 15)
        UserControl.Line (UserControl.Height / 30, 0)-(UserControl.Height / 30, UserControl.Height / 15)
    End If
    RefreshCaption
    UserControl.Refresh
    DoEvents
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("Title", m_Caption2, "")
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
    Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Shape", m_Shape, m_def_Shape)
    Call PropBag.WriteProperty("ControlData", mControlData, "")
End Sub

