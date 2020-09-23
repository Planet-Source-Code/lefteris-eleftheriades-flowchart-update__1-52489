VERSION 5.00
Begin VB.UserControl CoolerBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   Begin VB.PictureBox Se 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF5522&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   885
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2985
      Top             =   615
   End
   Begin VB.PictureBox S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   1
      Top             =   -15
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox D 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      Top             =   435
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "CoolerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This toolbar control Manipulates the buttons graphic file so it is drawn over a non-rectangular surface and converts it to grayscale just as in IE

'Created With the help of Mictosoft ActiveX Ctl Interface Wizard Add-in

'Api Types
Private Type POINTAPI
        X As Long
        Y As Long
End Type

'Api Declarations
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Enumerations
Enum tStyles
   OfficeXP = 0
   InternetExplorer = 1
End Enum

Enum cAlignment
   cTop = 0
   cLeft = 1
End Enum

'Default Property Values:
Const m_def_Margin = 8
Const m_def_Align = 0
Const m_def_Style = 0
Const m_def_MaskColor = &HFF00FF
'Const m_def_Align = 0
Const m_def_ButtonHeight = 32
Const m_def_ButtonWidth = 32
'Property Variables:
Dim m_Margin As Long
Dim m_Align As Byte
Dim m_Style As tStyles
Dim m_PTmp As Picture
Dim m_MaskColor As OLE_COLOR
'Dim m_Align As Byte
Dim m_ButtonHeight As Long
Dim m_ButtonWidth As Long
Dim Disabled(-2 To 100) As Boolean
'Event Declarations:
Event AlignmentChanged(newAlignment As Byte)
Event Click(buttonIndex As Integer) 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick(buttonIndex As Integer) 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(buttonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(buttonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(buttonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_UserMemId = -607
Event MouseOut()

Private bisMouseOver As Boolean
Private MOIv As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl_MouseOut
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
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
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
End Property

Public Property Get ButtonDisabled(ByVal Index As Integer) As Boolean
    ButtonDisabled = Disabled(Index)
End Property

Public Property Let ButtonDisabled(ByVal Index As Integer, ByVal New_ButtonDisabled As Boolean)
    Disabled(Index) = New_ButtonDisabled
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Function isMouseOver() As Boolean
Dim pt As POINTAPI
GetCursorPos pt
isMouseOver = (WindowFromPoint(pt.X, pt.Y) = hwnd)
End Function

Private Sub tmrOut_Timer()
   If bisMouseOver Then
      If Not isMouseOver Then
         bisMouseOver = False
         UserControl_MouseOut
         RaiseEvent MouseOut
      End If
   End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click(MOIv)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick(MOIv)
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
    RaiseEvent MouseDown(MOIv, Button, Shift, X, Y)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MIX&
    
    
    MIX = Int(X / (m_ButtonWidth + m_Margin))
    If MIX < -2 Then Exit Sub
    If MIX > (S.Width \ ButtonWidth) - 1 Then MIX = -2
      If m_Style = 1 Then
       If m_Align = 0 Then
         If Disabled(MIX) = False Then
           UserControl.Cls
           For I = 0 To (S.Width \ ButtonWidth) - 1
             If I <> MIX Or (Not isMouseOver) Then
               BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
               BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 1, SRCINVERT
             End If
           Next
           If Button = 0 Then
             BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 2, SRCAND
             BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 0, SRCINVERT
             
             UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX, S.Height + m_Margin), vbWhite
             UserControl.Line ((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), &H707070
             UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, 0), vbWhite
             UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, S.Height + m_Margin)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), &H707070
           Else
             If isMouseOver Then
               BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 + 1, m_Margin / 2 + 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 2, SRCAND
               BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 + 1, m_Margin / 2 + 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 0, SRCINVERT
      
               UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX, S.Height + m_Margin), &H707070
               UserControl.Line ((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), vbWhite
               UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, 0), &H707070
               UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, S.Height + m_Margin)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), vbWhite
             End If
           End If
         End If
       Else
         MIX = Int(Y / (m_ButtonHeight + m_Margin))
         If Disabled(MIX) = False Then
           UserControl.Cls
           If MIX > (S.Width \ ButtonWidth) - 1 Then MIX = -2
           For I = 0 To (S.Width \ ButtonWidth) - 1
             If I <> MIX Or (Not isMouseOver) Then
               BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
               BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 1, SRCINVERT
             End If
           Next
           If Button = 0 Then
             BitBlt UserControl.hdc, m_Margin / 2 - 1, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 2, SRCAND
             BitBlt UserControl.hdc, m_Margin / 2 - 1, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 0, SRCINVERT
             
             UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX), vbWhite
             UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), &H707070
             UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(0, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), vbWhite
             UserControl.Line (S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), &H707070
           Else
             If isMouseOver Then
                BitBlt UserControl.hdc, m_Margin / 2 + 1, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 + 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 2, SRCAND
                BitBlt UserControl.hdc, m_Margin / 2 + 1, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 + 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 0, SRCINVERT

                UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX), &H707070
                UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), vbWhite
                UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(0, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), &H707070
                UserControl.Line (S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), vbWhite
             End If
           End If
         End If
       End If
    Else
       If m_Align = 0 Then
         If Disabled(MIX) = False Then
           UserControl.Cls
           For I = 0 To (S.Width \ ButtonWidth) - 1
             BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
             BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 0, SRCINVERT
           Next
           If Button = 0 Then
              BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 2, SRCAND
              BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2 - 1, m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 0, SRCINVERT
              Alpha_Blend UserControl.hdc, Se.hdc, (m_ButtonWidth + m_Margin) * MIX, 0, 0, 0, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, 90
              UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), Se.BackColor, B
           Else
              If isMouseOver Then
                BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 2, SRCAND
                BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * MIX + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(X / (m_ButtonWidth + m_Margin)), S.Height * 0, SRCINVERT
                Alpha_Blend UserControl.hdc, Se.hdc, (m_ButtonWidth + m_Margin) * MIX, 0, 0, 0, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, 120
                UserControl.Line ((m_ButtonWidth + m_Margin) * MIX, 0)-((m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin, S.Height + m_Margin), Se.BackColor, B
              End If
           End If
         End If
       Else
         MIX = Int(Y / (m_ButtonHeight + m_Margin))
         If Disabled(MIX) = False Then
           UserControl.Cls
           If MIX > (S.Width \ ButtonWidth) - 1 Then MIX = -2
           For I = 0 To (S.Width \ ButtonWidth) - 1
             BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
             BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 0, SRCINVERT
           Next
           If Button = 0 Then
              BitBlt UserControl.hdc, m_Margin / 2 - 1, (m_ButtonHeight + m_Margin) * MIX + m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 2, SRCAND
              BitBlt UserControl.hdc, m_Margin / 2 - 1, (m_ButtonHeight + m_Margin) * MIX + m_Margin / 2 - 1, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 0, SRCINVERT
              Alpha_Blend UserControl.hdc, Se.hdc, 0, (m_ButtonWidth + m_Margin) * MIX, 0, 0, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, 90
              UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), Se.BackColor, B
           Else
             If isMouseOver Then
                BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonHeight + m_Margin) * MIX + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 2, SRCAND
                BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonHeight + m_Margin) * MIX + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, (m_ButtonWidth) * Int(Y / (m_ButtonHeight + m_Margin)), S.Height * 0, SRCINVERT
                Alpha_Blend UserControl.hdc, Se.hdc, 0, (m_ButtonWidth + m_Margin) * MIX, 0, 0, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, m_ButtonWidth + m_Margin, m_ButtonHeight + m_Margin, 120
                UserControl.Line (0, (m_ButtonWidth + m_Margin) * MIX)-(S.Height + m_Margin, (m_ButtonWidth + m_Margin) * MIX + m_ButtonWidth + m_Margin), Se.BackColor, B
             End If
           End If
         End If
       End If
    End If
    
    RaiseEvent MouseMove(MOIv, Button, Shift, X, Y)
    
    If Disabled(MIX) = False Then
      MOIv = MIX
      tmrOut.Enabled = True
      bisMouseOver = True
    Else
      MOIv = -2
      tmrOut.Enabled = True
      UserControl_MouseOut
    End If
End Sub

Private Sub UserControl_MouseOut()
    UserControl.Cls
    If m_Style = 1 Then
      If m_Align = 0 Then
          For I = 0 To (S.Width \ ButtonWidth) - 1
            BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
            BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 1, SRCINVERT
          Next
      Else
          For I = 0 To (S.Width \ ButtonWidth) - 1
            BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
            BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 1, SRCINVERT
          Next
      End If
    Else
      If m_Align = 0 Then
          For I = 0 To (S.Width \ ButtonWidth) - 1
            BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
            BitBlt UserControl.hdc, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 0, SRCINVERT
          Next
      Else
          For I = 0 To (S.Width \ ButtonWidth) - 1
            BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 2, SRCAND
            BitBlt UserControl.hdc, m_Margin / 2, (m_ButtonWidth + m_Margin) * I + m_Margin / 2, m_ButtonWidth, m_ButtonHeight, D.hdc, m_ButtonWidth * I, S.Height * 0, SRCINVERT
          Next
      End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(MOIv, Button, Shift, X, Y)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonHeight() As Long
    ButtonHeight = m_ButtonHeight
End Property

Public Property Let ButtonHeight(ByVal New_ButtonHeight As Long)
    m_ButtonHeight = New_ButtonHeight
    PropertyChanged "ButtonHeight"
    UserControl_MouseOut
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonWidth() As Long
    ButtonWidth = m_ButtonWidth
End Property

Public Property Let ButtonWidth(ByVal New_ButtonWidth As Long)
    m_ButtonWidth = New_ButtonWidth
    PropertyChanged "ButtonWidth"
    UserControl_MouseOut
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
'    Set m_ButtonImages = LoadPicture("")
    m_ButtonHeight = m_def_ButtonHeight
    m_ButtonWidth = m_def_ButtonWidth
'    m_Align = m_def_Align
    m_MaskColor = m_def_MaskColor
    Set m_PTmp = LoadPicture("")
    m_Style = m_def_Style
    m_Align = m_def_Align
    m_Margin = m_def_Margin
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_ButtonHeight = PropBag.ReadProperty("ButtonHeight", m_def_ButtonHeight)
    m_ButtonWidth = PropBag.ReadProperty("ButtonWidth", m_def_ButtonWidth)
    Set S.Picture = PropBag.ReadProperty("ButtonImages", Nothing)
'    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    D.Width = S.Width
    D.Height = S.Height * 3
    Set D.Picture = PropBag.ReadProperty("PTmp", Nothing)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    Se.BackColor = PropBag.ReadProperty("SelectionColor", &HFF5522)
    m_Margin = PropBag.ReadProperty("Margin", m_def_Margin)
    Se.Width = ButtonWidth + m_Margin
    Se.Height = ButtonHeight + m_Margin
    UserControl_MouseOut
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ButtonHeight", m_ButtonHeight, m_def_ButtonHeight)
    Call PropBag.WriteProperty("ButtonWidth", m_ButtonWidth, m_def_ButtonWidth)
    Call PropBag.WriteProperty("ButtonImages", S.Picture, Nothing)
'    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("PTmp", D.Picture, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("SelectionColor", Se.BackColor, &HFF5522)
    Call PropBag.WriteProperty("Margin", m_Margin, m_def_Margin)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=S,S,-1,Picture
Public Property Get ButtonImages() As Picture
Attribute ButtonImages.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set ButtonImages = S.Picture
End Property

Public Property Set ButtonImages(ByVal New_ButtonImages As Picture)
    Set S.Picture = New_ButtonImages
    D.Width = S.Width
    MakeAdvBitmap m_MaskColor And &HFF&, (m_MaskColor And &HFF00&) \ &H100&, (m_MaskColor And &HFF0000) \ &H10000
    PropertyChanged "ButtonImages"
    Set m_PTmp = D.Picture
    PropertyChanged "PTmp"
    m_ButtonWidth = S.Height
    m_ButtonHeight = S.Height
    UserControl_MouseOut
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Align() As cAlignment
    Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As cAlignment)
    Dim OW&
    If m_Align <> New_Align Then
       OW = UserControl.Width
       UserControl.Width = UserControl.Height
       UserControl.Height = OW
    End If
    m_Align = New_Align
    PropertyChanged "Align"
    UserControl_MouseOut
End Property

Sub MakeAdvBitmap(MaskR As Byte, MaskG As Byte, MaskB As Byte)
  'Draws 3 images in a jiffy thanks to my DIBits Module
  'See Ultra Fast Image Processing.Doc (if i included it any ways, it explains all about fast image processing)
  Dim PictureBits() As Byte
  Dim GrayBits() As Byte
  Dim MaskBits() As Byte
  Dim SpriteBits() As Byte
  Dim X&, Y&, A&, R&, G&, B&
  LoadBitmapInformation PictureBits, S.Width, S.Height, S.Picture.Handle
  ReDim GrayBits(3, S.Width - 1, S.Height - 1)
  ReDim MaskBits(3, S.Width - 1, S.Height - 1)
  ReDim SpriteBits(3, S.Width - 1, S.Height - 1)
  For X = 0 To S.Width - 1
     For Y = 0 To S.Height - 1
         R = PictureBits(0, X, Y)
         G = PictureBits(1, X, Y)
         B = PictureBits(2, X, Y)

         If R = MaskR And B = MaskB And G = MaskG Then
            MaskBits(0, X, Y) = 255
            MaskBits(1, X, Y) = 255
            MaskBits(2, X, Y) = 255
         Else
            SpriteBits(0, X, Y) = R
            SpriteBits(1, X, Y) = G
            SpriteBits(2, X, Y) = B
            A = (R + G + B) / 3
            GrayBits(0, X, Y) = A
            GrayBits(1, X, Y) = A
            GrayBits(2, X, Y) = A
         End If
     Next
  Next
  D.Height = S.Height * 3
  D.Picture = LoadPicture("")
  D.Cls
  DrawDIB D.hdc, 0, 0, S.Width, S.Height, 0, 0, S.Width, S.Height, SpriteBits
  DrawDIB D.hdc, 0, S.Height, S.Width, S.Height, 0, 0, S.Width, S.Height, GrayBits
  DrawDIB D.hdc, 0, S.Height * 2, S.Width, S.Height, 0, 0, S.Width, S.Height, MaskBits
  D.Refresh
  D.Picture = D.Image
  Erase GrayBits
  Erase PictureBits
  Erase MaskBits
  Erase SpriteBits
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    D.Width = S.Width
    MakeAdvBitmap m_MaskColor And &HFF&, (m_MaskColor And &HFF00&) \ &H100&, (m_MaskColor And &HFF0000) \ &H10000
    PropertyChanged "MaskColor"
    Set m_PTmp = D.Picture
    PropertyChanged "PTmp"
    UserControl_MouseOut
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Style() As tStyles
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As tStyles)
    m_Style = New_Style
    PropertyChanged "Style"
    UserControl_MouseOut
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
    UserControl_MouseOut
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Se,Se,-1,BackColor
Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    SelectionColor = Se.BackColor
End Property

Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
    Se.BackColor() = New_SelectionColor
    PropertyChanged "SelectionColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,8
Public Property Get Margin() As Long
    Margin = m_Margin
End Property

Public Property Let Margin(ByVal New_Margin As Long)
    m_Margin = New_Margin
    PropertyChanged "Margin"
    Se.Width = ButtonWidth + m_Margin
    Se.Height = ButtonHeight + m_Margin
    UserControl_MouseOut
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
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, ByVal Width1 As Variant, ByVal Height1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant, ByVal Width2 As Variant, ByVal Height2 As Variant, ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

