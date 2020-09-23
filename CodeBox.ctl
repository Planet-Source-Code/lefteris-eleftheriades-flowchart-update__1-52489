VERSION 5.00
Begin VB.UserControl CodeBox 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ToolboxBitmap   =   "CodeBox.ctx":0000
   Begin VB.PictureBox picTextBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   195
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   424
      TabIndex        =   0
      Top             =   255
      Visible         =   0   'False
      Width           =   6360
   End
   Begin VB.PictureBox PicTextArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   0
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   424
      TabIndex        =   3
      Top             =   0
      Width           =   6360
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   6510
      Top             =   1080
   End
   Begin VB.PictureBox Selector 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   2
      Top             =   4530
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.PictureBox LineBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   1
      Top             =   4845
      Visible         =   0   'False
      Width           =   6570
   End
End
Attribute VB_Name = "CodeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Advanced Code Edit control by: Lefteris Eleftheriades
'Portion of projects VBScript Pad & Advanced Flowchart Programming
'———————————————————————————Advanced Code Edit Control——————————————————————————
'This control acts like a regular textbox  but automatically colors blue
'each word that matches any of the given in the Keywords property.
'it does not leave traces and it has all the functionalities of a regular textbox
'plus can set a background image.
'This control was designed for any person that wishes to make a programing interface
'for any purpose
'The control could be used in HTML editors, VBScript Editors,Javascript Editors
'your own programming language, as a VIP name hilighter or as a batch file editor,
'to hilight keywords.
'———————————————————————————————————————————————————————————————————————————————
'YOU MAY USE THIS CODE TO ANY OF YOUR PROGRAMS AS LONG AS YOU MENTION THE
'CREATOR'S NAME IN THE ABOUT BOX OF YOUR APPLICATION(if you have one).
'THE CONTROL CAN BE MODIFIED AS DESIRED BUT CAN NOT BE CLAIMED AS YOUR OWN.
'You are allowed to take any function of this code if you mention me in your
'about box. 2K+3 AdvCodeBox.Ctl SRA OS 76 1000
'———————————————————————————————————————————————————————————————————————————————
Option Explicit
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, _
               ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, _
               ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
               ByVal BLENDFUNCT As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCINVERT As Long = &H660046

Dim TextBoxLines(1000) As String
Dim Lineproperties(1000) As Byte
Dim CurretLeft As Long, CurerentLine As Long
Dim BlinkingLineX As Long
Dim BlinkState As Boolean
Dim sSelStart As Long
Dim vbsKeywords() As String
Dim NoOfLines As Long
Dim SelectedAreaFrom(1000) As Long
Dim SelectedAreaLength(1000) As Long
Dim LineSelected(1000) As Boolean

Dim SelectionX1&, SelectionX2&
Dim SelectionY1&, SelectionY2&
Dim PrevEventsHandled As Boolean

Const CommentSymbol As String = "'" ' "//"
Const StringSymbol As String = """" ' "'"
Dim DownChr&, MoveChr&
Dim DownLne&, MoveLne&
'Event Declarations:
Event Click() 'MappingInfo=PicTextArea,PicTextArea,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=PicTextArea,PicTextArea,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=PicTextArea,PicTextArea,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=PicTextArea,PicTextArea,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=PicTextArea,PicTextArea,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicTextArea,PicTextArea,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicTextArea,PicTextArea,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicTextArea,PicTextArea,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event LineChanged(PreviewsLine As Integer, NewLine As Integer)

'Default Property Values:
Const m_def_Number_of_Lines = 0
Const m_def_Top_Line = 0
Const m_def_Keywords = "Me,Mod,Line,GoTo,True,False,While,Until,Wend,Private,Public,Global,Sub,Function,End,If,Then,Else,Dim,For,Next,To,TypeOf,Is,Long,Integer,Boolean,Single,Double,Byte,Currency,String,Object,With,ByVal,(ByVal,ByRef,(ByRef,As,Type,Enum,Const,ReDim,Declare,Lib,Alias,Static,UBound,LBound,Do,Loop,Open,Close,And,Or,Not,Xor,Append,BF,Binary,Call,Select,Case,CBool,CByte,CCur,CDate,CDbl,CDec,CInt,CLng,AddressOf,Collection,Control,CSng,CStr,Cstr,CVar,ElseIf,Error,Exit,Explicit,Friend,Get,Let,Set,Input,New,Nothing,On,Option,Optional,(Optional,Output,Print,Property,Random,Step,Tokens,VarPtr"
Const m_def_KeywordColor = 0
Const m_def_CommentColor = 0
Const m_def_TextboxLine = ""
'Property Variables:
Dim m_Number_of_Lines As Integer
Dim m_Top_Line As Integer
Dim m_Keywords As String
Dim m_KeywordColor As OLE_COLOR
Dim m_CommentColor As OLE_COLOR
Dim m_TextboxLine As String
Dim VisibleLines&
Public Sub Alpha_Blend(DestinationHdc&, SourceHdc&, X&, Y&, CropX&, CropY&, CropW&, CropH&, Width&, Height&, Optional TransparancyLevel& = 128)
    AlphaBlend DestinationHdc&, X&, Y&, Width&, Height&, SourceHdc&, CropX&, CropY&, CropW&, CropH&, TransparancyLevel * &H10000
End Sub

Private Sub PicTextArea_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
If PrevEventsHandled Then
PrevEventsHandled = False
  Dim TL&, TT&, CL&
 'Part 1: The special keys handler
  DeselectAll
  RedrawLine
  Select Case KeyCode
      Case vbKeyLeft
           If CurretLeft > 1 Then CurretLeft = CurretLeft - 1
      Case vbKeyRight
           If CurretLeft <= Len(TextBoxLines(CurerentLine)) Then
              CurretLeft = CurretLeft + 1
           End If
      Case vbKeyUp
           If CurerentLine > 0 Then
              CurerentLine = CurerentLine - 1
              RaiseEvent LineChanged(CurerentLine + 1, (CurerentLine))
              DoEvents
           End If
           If CurretLeft > Len(TextBoxLines(CurerentLine)) Then CurretLeft = Len(TextBoxLines(CurerentLine)) + 1
           PropertyChanged "Current_Line"
      Case vbKeyDown
           CurerentLine = CurerentLine + 1
           RaiseEvent LineChanged(CurerentLine - 1, (CurerentLine))
           If CurretLeft > Len(TextBoxLines(CurerentLine)) Then CurretLeft = Len(TextBoxLines(CurerentLine)) + 1
           PropertyChanged "Current_Line"
           DoEvents
      Case vbKeyBack
           If CurretLeft > 1 Then
             TextBoxLines(CurerentLine) = Mid(TextBoxLines(CurerentLine), 1, CurretLeft - 2) & Mid(TextBoxLines(CurerentLine), CurretLeft)
             CurretLeft = CurretLeft - 1
           ElseIf CurretLeft = 1 And CurerentLine > 0 Then
              'The tricky part move the text 1 line up
              TT = CurerentLine
              CurerentLine = CurerentLine - 1
              TL = Len(TextBoxLines(CurerentLine))
              TextBoxLines(CurerentLine) = TextBoxLines(CurerentLine) & TextBoxLines(CurerentLine + 1)
              RedrawLine False
              For CL = TT To 999
                  TextBoxLines(CL) = TextBoxLines(CL + 1)
                  If CL < 30 Then
                     RedrawLine False, CL
                  End If
              Next
              PicTextArea.Refresh
              TextBoxLines(1000) = ""
              RedrawLine False
              CurerentLine = TT - 1
              CurretLeft = TL + 1
              RaiseEvent LineChanged((TT), TT - 1)
              PropertyChanged "Current_Line"
           End If
           
           If CurretLeft > Len(TextBoxLines(CurerentLine)) Then CurretLeft = Len(TextBoxLines(CurerentLine)) + 1
           RedrawLine
      Case vbKeyDelete
            'Todo: code the shit for delete again
           If CurretLeft <= Len(TextBoxLines(CurerentLine)) Then
             TextBoxLines(CurerentLine) = Mid(TextBoxLines(CurerentLine), 1, CurretLeft - 1) & Mid(TextBoxLines(CurerentLine), CurretLeft + 1)
           Else
              TT = CurerentLine
              TL = Len(TextBoxLines(CurerentLine))
              TextBoxLines(CurerentLine) = TextBoxLines(CurerentLine) & TextBoxLines(CurerentLine + 1)
              RedrawLine False
              For CL = TT + 1 To 999
                  TextBoxLines(CL) = TextBoxLines(CL + 1)
                  If CL < TT + PicTextArea.Height \ PicTextArea.TextHeight("|") Then
                        RedrawLine False, CL
                  End If
              Next
              PicTextArea.Refresh
              TextBoxLines(1000) = ""
              RedrawLine False
              CurerentLine = TT - 1
              If CurerentLine < 0 Then CurerentLine = 0
              RaiseEvent LineChanged((TT), (CurerentLine))
              CurretLeft = TL + 1
              PropertyChanged "Current_Line"
           End If
           RedrawLine
      Case vbKeyReturn
           'Shift all lines under the currentline down
           TT = CurerentLine
                      
           For CL = 999 To TT + 2 Step -1
                TextBoxLines(CL) = TextBoxLines(CL - 1)
                'scroll affected
                If CL < TT + PicTextArea.Height \ PicTextArea.TextHeight("|") Then
                   RedrawLine False, CL
                End If
            Next
            DoEvents
            PicTextArea.Refresh
            TextBoxLines(TT + 1) = Mid(TextBoxLines(TT), CurretLeft)
            RedrawLine False, TT + 1
            TextBoxLines(TT) = Mid(TextBoxLines(TT), 1, CurretLeft - 1)
            RedrawLine False, TT
            PicTextArea.Refresh
            CurerentLine = CurerentLine + 1
            NoOfLines = NoOfLines + 1
            RaiseEvent LineChanged((TT), (CurerentLine))
            PropertyChanged "Current_Line"
      Case vbKeyHome
            CurretLeft = 1
      Case vbKeyEnd
            CurretLeft = Len(TextBoxLines(CurerentLine)) + 1
  End Select
  RedrawLine True
  BitBlt picTextBuffer.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, PicTextArea.hdc, 0, 0, SRCCOPY
  DoEvents
  PrevEventsHandled = True
End If
End Sub

Sub RedrawLine(Optional BlinkSt As Boolean = False, Optional ByVal lLine As Long = -1)
   Dim WordsInLine() As String, CurrX As Long
   Dim ChrX&, IsComment As Boolean
   Dim IsStringBlock As Boolean, cSelection As Boolean
   Dim I&
   'Dim WordsInLine2() As String,TempLine As String
   If lLine = -1 Then lLine = CurerentLine
   LineBuffer.Cls
      
'      TempLine = Trim(TextBoxLines(CurerentLine))
'      TempLine = Replace(TempLine, "(", " ")
'      TempLine = Replace(TempLine, ")", " ")
'      TempLine = Replace(TempLine, ".", " ")
'      TempLine = Replace(TempLine, ",", " ")
      
   If CurerentLine < 0 Then Exit Sub
      
   If Lineproperties(lLine) = 0 Then
      If LineSelected(lLine) Then
         LineBuffer.BackColor = BlendColors(Selector.Point(1, 1), PicTextArea.BackColor, 90)

         '&HFFA5A5
         LineBuffer.ForeColor = &H5A0000
         'If TextBoxLines(lLine) <> "" Then
            cSelection = True
         'End If
      Else
         LineBuffer.BackColor = PicTextArea.BackColor: LineBuffer.ForeColor = 0
         cSelection = False
      End If
      WordsInLine = Split(" " & TextBoxLines(lLine), " ")
      'WordsInLine2 = Split(" " & TextBoxLines(CurerentLine) & " ", " ")
      CurrX = 0
      For I = 1 To UBound(WordsInLine())
         LineBuffer.CurrentX = CurrX
         LineBuffer.CurrentY = 0
         If IsKeyword(WordsInLine(I)) And Not (IsComment Or IsStringBlock) Then
            LineBuffer.ForeColor = &HFF0000
         Else
            If CountChars(WordsInLine(I), StringSymbol) + 2 Mod 2 = 1 Then IsStringBlock = Not IsStringBlock
            
            If InStr(1, WordsInLine(I), CommentSymbol) <> 0 And Not IsStringBlock Then
               'LineBuffer.ForeColor = &H9000&
               'IsComment = True
            
            ElseIf IsComment And Not IsStringBlock Then
               'LineBuffer.ForeColor = &H9000&
            Else
               If LineSelected(lLine) Then
                  LineBuffer.ForeColor = &H5A0000
               Else
                  ''''The next foked things up. it cleared the buffer'''''LineBuffer.BackColor = vbWhite
                  LineBuffer.ForeColor = vbBlack
                  DoEvents
               End If
            End If
         End If
'         Debug.Print LineBuffer.CurrentX; LineBuffer.CurrentY
         LineBuffer.Print WordsInLine(I)
         LineBuffer.Refresh
         CurrX = CurrX + LineBuffer.TextWidth(WordsInLine(I)) + LineBuffer.TextWidth(" ")
         'ChrX = ChrX + Len(WordsInLine(I))
      Next I
   Else
         Select Case Lineproperties(CurerentLine)
                Case 1: LineBuffer.BackColor = vbYellow: LineBuffer.ForeColor = 0
                Case 2: LineBuffer.BackColor = &H80: LineBuffer.ForeColor = vbWhite
                Case 3: LineBuffer.BackColor = vbWhite: LineBuffer.ForeColor = vbRed
         End Select
         LineBuffer.Print TextBoxLines(lLine)
   End If
      DrawBlinkingLine BlinkSt, cSelection
      CopyBufferToTextbox lLine
End Sub

Function CountChars(String1$, Charac$)
   Dim I&, C&
   For I = 1 To Len(String1) - Len(Charac$) + 1
       If Mid(String1, I, Len(Charac)) = Charac Then C = C + 1
   Next I
   CountChars = C
End Function

Function IsKeyword(ByVal Word As String) As Boolean
Attribute IsKeyword.VB_Description = "Returns if a word is a keyword"
   Dim Flag As Boolean, I&
   Flag = False
   On Error GoTo Skoops
   If Word <> "" Then
     Word = UCase(Word)
     For I = 1 To UBound(vbsKeywords())
        If UCase(vbsKeywords(I)) = Word Then
           Flag = True
           Exit For
        End If
     Next I
     IsKeyword = Flag
   End If
Exit Function
Skoops:
If Err.Number <> 0 Then IsKeyword = False
End Function

Private Sub PicTextArea_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
   On Error Resume Next
   DeselectAll
   If KeyAscii >= 32 And CurretLeft > 0 Then
      TextBoxLines(CurerentLine) = Mid(TextBoxLines(CurerentLine), 1, CurretLeft - 1) & Chr(KeyAscii) & Mid(TextBoxLines(CurerentLine), CurretLeft)
      CurretLeft = CurretLeft + 1
      RedrawLine True
      BitBlt picTextBuffer.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, PicTextArea.hdc, 0, 0, SRCCOPY
   End If
   
   'PicTextArea.Refresh
End Sub
'
'Sub DrawBlinkingLine(lVisible As Boolean, Optional cSelection As Boolean)
'
' If CurretLeft > 0 Then
'   BlinkingLineX = LineBuffer.TextWidth(Mid(TextBoxLines(CurerentLine), 1, CurretLeft - 1))
'   '(-DrawBlinkingLine * vbWhite) is a combination of boolean algebra with maths
'   If cSelection Then
'      LineBuffer.Line (BlinkingLineX, 0)-(BlinkingLineX, LineBuffer.Height), &HFFA5A5
'   Else
'      LineBuffer.Line (BlinkingLineX, 0)-(BlinkingLineX, LineBuffer.Height), ((lVisible + 1) * vbWhite)
'   End If
' End If
'End Sub

Private Sub PicTextArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
   Dim I&, FoundCharOver As Boolean
If Button = 1 Then DeselectAll
  
   RedrawLine
   CurerentLine = Int(Y / (LineBuffer.TextHeight("|") + 0))
   If CurerentLine < 0 Then Exit Sub
   If CurerentLine > NoOfLines Then CurerentLine = NoOfLines
   For I = 1 To Len(TextBoxLines(CurerentLine))
      If X < LineBuffer.TextWidth(Mid(TextBoxLines(CurerentLine), 1, I)) Then
         CurretLeft = I
         FoundCharOver = True
         Exit For
      End If
   Next I
      
   If Not FoundCharOver Then CurretLeft = Len(TextBoxLines(CurerentLine)) + 1

If Button = 1 Then
   DownChr& = CurretLeft
   DownLne& = CurerentLine

   'Alpha_Blend PicTextArea.hdc, Selector.hdc, (X), (Y), 0, 0, 20, Selector.Height, 20, Selector.Height, 128
   RedrawLine
   Timer1.Enabled = False
End If
MoveChr = 0
MoveLne = 0
End Sub

Private Sub PicTextArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
On Error Resume Next
Dim FoundCharOver As Boolean
Dim XLeft&, XRight&, XWidth&, XTmp&, I&, ML&, YLn&
Dim X2Left&, X2Right&, X2Width&
If DownChr > 0 And Button = 1 Then
   For I = 1 To Len(TextBoxLines(CurerentLine))
      If X < LineBuffer.TextWidth(Mid(TextBoxLines(CurerentLine), 1, I)) Then
         ML = I
         FoundCharOver = True
         Exit For
      End If
   Next I
   YLn = Int(Y / (LineBuffer.TextHeight("|") + 0))
   If YLn < 0 Then YLn = 0
   
   If Not FoundCharOver Then ML = Len(TextBoxLines(YLn))
  ' If YLn > DownLne Then ML = Len(TextBoxLines(DownLne))

   'Debug.Print YLn; DownLne
   XLeft = LineBuffer.TextWidth(Mid(TextBoxLines(DownLne), 1, DownChr - 1))
   XRight = LineBuffer.TextWidth(Mid(TextBoxLines(YLn), 1, ML))
   PicTextArea.Refresh
   If XLeft > XRight Then
      XTmp = XLeft
      XLeft = XRight
      If X < 2 Then XLeft = 0
      XRight = XTmp
   End If
   If YLn > DownLne Then
      XLeft = LineBuffer.TextWidth(Mid(TextBoxLines(DownLne), 1, DownChr - 1))
      XRight = LineBuffer.TextWidth(TextBoxLines(DownLne))
      
      X2Left = 0
      X2Right = LineBuffer.TextWidth(Mid(TextBoxLines(YLn), 1, ML))
   End If
   
   XWidth = XRight - XLeft
   X2Width = X2Right - X2Left

   BitBlt PicTextArea.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, picTextBuffer.hdc, 0, 0, SRCCOPY
   If YLn > DownLne Then
     For I = DownLne + 1 To YLn - 1
       LineSelected(I) = True
       RedrawLine False, I
     Next
   Else
     For I = YLn + 1 To DownLne
       LineSelected(I) = True
       RedrawLine False, I
     Next
   End If
   MoveChr = ML
   MoveLne = YLn
   Alpha_Blend PicTextArea.hdc, Selector.hdc, XLeft + 1, DownLne * LineBuffer.TextHeight("|"), 0, 0, XWidth, Selector.Height, XWidth, Selector.Height, 90
   Alpha_Blend PicTextArea.hdc, Selector.hdc, X2Left + 1, YLn * LineBuffer.TextHeight("|"), 0, 0, X2Width, Selector.Height, X2Width, Selector.Height, 90
   'if YLn = DownLne and DownChr = XRight
   PicTextArea.Refresh
End If
End Sub

Private Sub PicTextArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
   If MoveChr = 0 And MoveLne = 0 Then
      Timer1.Enabled = True
   End If
End Sub

Private Sub Timer1_Timer()
  BlinkState = Not BlinkState
  'If GetFocus = PicTextArea.hWnd Or GetFocus = UserControl.hWnd Then
     RedrawLine BlinkState
  'Else
  '   RedrawLine False
  'End If
  LineBuffer.Refresh
End Sub

Sub CopyBufferToTextbox(Optional ByVal lLine As Long = -1)
   If lLine = -1 Then lLine = CurerentLine
   PicTextArea.Line (0, lLine * (LineBuffer.TextHeight("|") + 0))-(1000, lLine * (LineBuffer.TextHeight("|") + 0) + (LineBuffer.TextHeight("|") - 1)), PicTextArea.BackColor, BF
   PicTextArea.PaintPicture LineBuffer.Image, 1, lLine * (LineBuffer.TextHeight("|") + 0), LineBuffer.TextWidth(TextBoxLines(lLine)) + 1, , , , LineBuffer.TextWidth(TextBoxLines(lLine)) + 1
End Sub
'75mins

Private Sub UserControl_Initialize()

   'LineSelected(1) = True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicTextArea.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicTextArea.BackColor() = New_BackColor
    picTextBuffer.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = PicTextArea.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    PicTextArea.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub PicTextArea_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    PicTextArea.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = PicTextArea.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    PicTextArea.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = PicTextArea.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    PicTextArea.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

Private Sub PicTextArea_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = PicTextArea.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    PicTextArea.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = PicTextArea.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    PicTextArea.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LineBuffer,LineBuffer,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = LineBuffer.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set LineBuffer.Font = New_Font
    Set PicTextArea.Font = New_Font
    Set picTextBuffer.Font = New_Font
    Set Selector.Font = New_Font
    LineBuffer.Height = LineBuffer.TextHeight("|")
    Selector.Height = LineBuffer.TextHeight("|")
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display normal text and graphics in an object."
    ForeColor = PicTextArea.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    PicTextArea.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = PicTextArea.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = PicTextArea.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = PicTextArea.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/Sets the number of seconds the curret will be blinking"
    Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    Timer1.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

Private Sub PicTextArea_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,Line
Public Sub Line(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
Attribute Line.VB_Description = "Draws lines and rectangles on an object."
    PicTextArea.Line (X1, Y1)-(X2, Y2), Color
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = PicTextArea.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PicTextArea.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = PicTextArea.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    PicTextArea.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    PicTextArea.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,Point
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point = PicTextArea.Point(X, Y)
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
'MappingInfo=PicTextArea,PicTextArea,-1,PSet
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    PicTextArea.PSet Step(X, Y), Color
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    PicTextArea.Refresh
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    PicTextArea.Move 0, 0, UserControl.Width / 15, UserControl.Height / 15
    picTextBuffer.Move 0, 0, UserControl.Width / 15, UserControl.Height / 15
    LineBuffer.Width = UserControl.Width / 15
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LineBuffer,LineBuffer,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    TextHeight = LineBuffer.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LineBuffer,LineBuffer,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
    TextWidth = LineBuffer.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicTextArea,PicTextArea,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = PicTextArea.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    PicTextArea.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DrawBlinkingLine(lVisible As Boolean, Optional cSelection As Boolean) As Variant
   'To do: make it blink
   If CurretLeft > 0 Then
      BlinkingLineX = LineBuffer.TextWidth(Mid(TextBoxLines(CurerentLine), 1, CurretLeft - 1))
   End If
   '(-DrawBlinkingLine * vbWhite) is a combination of boolean algebra with maths
   If cSelection Then
      LineBuffer.Line (BlinkingLineX, 0)-(BlinkingLineX, LineBuffer.Height), &HFFA5A5
   Else
      LineBuffer.Line (BlinkingLineX, 0)-(BlinkingLineX, LineBuffer.Height), ((lVisible + 1) * vbWhite)
   End If
End Function

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim I&
    
    UserControl.ScaleMode = vbPixels
    PicTextArea.AutoRedraw = True
    PicTextArea.ScaleMode = vbPixels
    LineBuffer.AutoRedraw = True
    LineBuffer.ScaleMode = vbPixels
    
    PicTextArea.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picTextBuffer.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    PicTextArea.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    PicTextArea.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    PicTextArea.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    PicTextArea.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    PicTextArea.Enabled = PropBag.ReadProperty("Enabled", True)
    Set LineBuffer.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set PicTextArea.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set picTextBuffer.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Selector.Font = PropBag.ReadProperty("Font", Ambient.Font)
    PicTextArea.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Timer1.Interval = PropBag.ReadProperty("Interval", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    PicTextArea.MousePointer = PropBag.ReadProperty("MousePointer", 3)
    PicTextArea.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Keywords = PropBag.ReadProperty("Keywords", m_def_Keywords)
    m_KeywordColor = PropBag.ReadProperty("KeywordColor", m_def_KeywordColor)
    m_CommentColor = PropBag.ReadProperty("CommentColor", m_def_CommentColor)
    
    For I = 0 To UBound(TextBoxLines)
         TextBoxLines(I) = PropBag.ReadProperty("TextboxLine" & I, m_def_TextboxLine)
         If I < 40 Then RedrawLine False, I
    Next I
    
    Erase vbsKeywords
    vbsKeywords = Split("," & m_Keywords & ",", ",")
       
    CurerentLine = 0
    CurretLeft = 1
    LineBuffer.Height = LineBuffer.TextHeight("|")
    Selector.Height = LineBuffer.TextHeight("|")
     
    BitBlt picTextBuffer.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, PicTextArea.hdc, 0, 0, SRCCOPY
    NoOfLines = 4
   
   
    CurerentLine = PropBag.ReadProperty("Current_Line", 0)
    m_Top_Line = PropBag.ReadProperty("Top_Line", m_def_Top_Line)
    m_Number_of_Lines = PropBag.ReadProperty("Number_of_Lines", m_def_Number_of_Lines)
    
    PrevEventsHandled = True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim I&
    Call PropBag.WriteProperty("BackColor", PicTextArea.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", PicTextArea.BorderStyle, 0)
    Call PropBag.WriteProperty("CurrentX", PicTextArea.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", PicTextArea.CurrentY, 0)
    Call PropBag.WriteProperty("DrawWidth", PicTextArea.DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", PicTextArea.Enabled, True)
    Call PropBag.WriteProperty("Font", LineBuffer.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", PicTextArea.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Interval", Timer1.Interval, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", PicTextArea.MousePointer, 3)
    Call PropBag.WriteProperty("ToolTipText", PicTextArea.ToolTipText, "")
    Call PropBag.WriteProperty("Keywords", m_Keywords, m_def_Keywords)
    Call PropBag.WriteProperty("KeywordColor", m_KeywordColor, m_def_KeywordColor)
    Call PropBag.WriteProperty("CommentColor", m_CommentColor, m_def_CommentColor)
    For I = 0 To UBound(TextBoxLines)
         If I < 40 Then Call PropBag.WriteProperty("TextboxLine" & I, TextBoxLines(I), m_def_TextboxLine)
    Next I
    Call PropBag.WriteProperty("Current_Line", CurerentLine, 0) 'CurrentLine =0
    Call PropBag.WriteProperty("Top_Line", m_Top_Line, m_def_Top_Line)
    Call PropBag.WriteProperty("Number_of_Lines", m_Number_of_Lines, m_def_Number_of_Lines)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Keywords() As String
Attribute Keywords.VB_Description = "Returns/Sets The keywords that will be hilighted. Seperate each word with a comma"
    Keywords = m_Keywords
End Property

Public Property Let Keywords(ByVal New_Keywords As String)
    m_Keywords = New_Keywords
    PropertyChanged "Keywords"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get KeywordColor() As OLE_COLOR
Attribute KeywordColor.VB_Description = "Returns/Sets the hilight color of the keywordwords"
    KeywordColor = m_KeywordColor
End Property

Public Property Let KeywordColor(ByVal New_KeywordColor As OLE_COLOR)
    m_KeywordColor = New_KeywordColor
    PropertyChanged "KeywordColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CommentColor() As OLE_COLOR
Attribute CommentColor.VB_Description = "Returns/Sets the color Commented line will appear in"
    CommentColor = m_CommentColor
End Property

Public Property Let CommentColor(ByVal New_CommentColor As OLE_COLOR)
    m_CommentColor = New_CommentColor
    PropertyChanged "CommentColor"
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,""
Public Property Get TextboxLine(ByVal Index As Integer) As String
    TextboxLine = TextBoxLines(Index)
End Property

Public Property Let TextboxLine(ByVal Index As Integer, ByVal New_TextboxLine As String)
    TextBoxLines(Index) = New_TextboxLine
    RedrawLine False, Index
    BitBlt picTextBuffer.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, PicTextArea.hdc, 0, 0, SRCCOPY
    PropertyChanged "TextboxLine" & Index
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Keywords = m_def_Keywords
    m_KeywordColor = m_def_KeywordColor
    m_CommentColor = m_def_CommentColor
    m_TextboxLine = m_def_TextboxLine
    m_Top_Line = m_def_Top_Line
    m_Number_of_Lines = m_def_Number_of_Lines
End Sub

Sub DeselectAll()
   Dim I&
   For I = 0 To 1000
       LineSelected(I) = False
   Next
   BitBlt PicTextArea.hdc, 0, 0, PicTextArea.Width, PicTextArea.Height, picTextBuffer.hdc, 0, 0, SRCCOPY
   PicTextArea.Refresh
   Timer1.Enabled = True
End Sub

Function BlendColors(Color1&, Color2&, Strength As Byte, Optional Brightness As Byte = 0) As Long
   'Function to blend two colors.
   'Splits Both color R, G, B's
   'Add up the colors ( 0.4 * Red1 + (1-0.4) * Red2 = RedOut) and so on
   Dim Power#
   Dim B1&, G1&, R1&
   Dim B2&, G2&, R2&
   
   Power# = Strength / 255
   
   B1 = (Color1 And &HFF&) * Power
   G1 = ((Color1 And &HFF00&) \ &H100&) * Power
   R1 = ((Color1 And &HFF0000) \ &H10000) * Power
   
   B2 = (Color2 And &HFF&) * (1 - Power)
   G2 = ((Color2 And &HFF00&) \ &H100&) * (1 - Power)
   R2 = ((Color2 And &HFF0000) \ &H10000) * (1 - Power)
   
   BlendColors = RGB(B1 + B2 + Brightness, G1 + G2 + Brightness, R1 + R2 + Brightness)
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Current_Line() As Integer
Attribute Current_Line.VB_Description = "The Line The Curret is currently over"
    Current_Line = CurerentLine
End Property

Public Property Let Current_Line(ByVal New_Current_Line As Integer)
    CurerentLine = New_Current_Line
    PropertyChanged "Current_Line"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Top_Line() As Integer
    Top_Line = m_Top_Line
End Property

Public Property Let Top_Line(ByVal New_Top_Line As Integer)
    m_Top_Line = New_Top_Line
    PropertyChanged "Top_Line"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Number_of_Lines() As Integer
    Number_of_Lines = m_Number_of_Lines
End Property

Public Property Let Number_of_Lines(ByVal New_Number_of_Lines As Integer)
    m_Number_of_Lines = New_Number_of_Lines
    PropertyChanged "Number_of_Lines"
End Property

Sub LineSyntaxHilight(Line_No As Integer, HilightType As Byte)
    Lineproperties(Line_No) = HilightType
    RedrawLine False, Line_No
    PicTextArea.Refresh
End Sub

