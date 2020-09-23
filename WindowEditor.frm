VERSION 5.00
Begin VB.Form WindowEditor 
   Caption         =   "Window Editor"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   Picture         =   "WindowEditor.frx":0000
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   6
      Top             =   2925
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   15
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2550
      Width           =   885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      Height          =   5025
      Left            =   930
      ScaleHeight     =   4965
      ScaleWidth      =   6570
      TabIndex        =   0
      Top             =   45
      Width           =   6630
      Begin VB.PictureBox frmSize 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   100
         Index           =   2
         Left            =   1605
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   4
         Top             =   3180
         Width           =   100
      End
      Begin VB.PictureBox frmSize 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   100
         Index           =   1
         Left            =   3345
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   3
         Top             =   3180
         Width           =   100
      End
      Begin VB.PictureBox frmSize 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   100
         Index           =   0
         Left            =   3345
         MousePointer    =   9  'Size W E
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   1530
         Width           =   100
      End
      Begin VB.PictureBox PicFrm 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   15
         ScaleHeight     =   3165
         ScaleWidth      =   3330
         TabIndex        =   1
         Top             =   15
         Width           =   3330
      End
   End
End
Attribute VB_Name = "WindowEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Under Development
Const HorRes = 0
Const DiagRes = 1
Const VertRes = 2
Dim FrmSizeMD(2) As POINTAPI

Enum CountBy
   ByRows = 1
   ByColums = 0
End Enum

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Index%
  Index = DrawButtons(Me, X, Y, 25, 25, 2, 6, 4, 3, ByRows, 2, 2)
End Sub

Function DrawButtons(DrawDest As Object, X As Single, Y As Single, ButWidth&, ButHeight&, Cols&, Rows&, SpaceH&, SpaceV&, IndexCountType As CountBy, Optional BeginOffcetX&, Optional BeginOffcetY&) As Integer
  Dim XX&, YY&, Col&, Row&, hIndex&, vIndex&
  
  Col = (X - BeginOffcetX) \ (ButWidth + SpaceH)
  XX = Col * (ButWidth + SpaceH) + BeginOffcetX
  
  Row = (Y - BeginOffcetY) \ (ButHeight + SpaceV)
  YY = Row * (ButHeight + SpaceV) + BeginOffcetY
  
  If IndexCountType = 0 Then
     DrawButtons = Col * Rows + Row
  Else
     DrawButtons = Col + Row * Cols
  End If
  
  If Row < Rows And Col < Cols Then
    DrawDest.Cls
    DrawDest.Line (XX, YY)-(XX + ButWidth, YY), &H606060
    DrawDest.Line (XX, YY)-(XX, YY + ButHeight), &H606060
    DrawDest.Line (XX, YY + ButHeight)-(XX + ButWidth, YY + ButHeight), &HFFFFFF
    DrawDest.Line (XX + ButWidth, YY)-(XX + ButWidth, YY + ButHeight), &HFFFFFF
  End If
End Function

Private Sub frmSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   FrmSizeMD(Index).X = X
   FrmSizeMD(Index).Y = Y
End Sub

Private Sub frmSize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then
    Select Case Index
      Case HorRes
           frmSize(Index).Left = frmSize(Index).Left + X - FrmSizeMD(Index).X
           frmSize(VertRes).Left = frmSize(Index).Left / 2 - 50
           PicFrm.Width = frmSize(Index).Left - 1
           frmSize(DiagRes).Move frmSize(Index).Left, frmSize(VertRes).Top
      Case DiagRes
           frmSize(Index).Move frmSize(Index).Left + X - FrmSizeMD(Index).X, frmSize(Index).Top + Y - FrmSizeMD(Index).Y
           PicFrm.Move 0, 0, frmSize(Index).Left - 1, frmSize(Index).Top - 1
           frmSize(HorRes).Move frmSize(Index).Left, frmSize(Index).Top / 2 - 50
           frmSize(VertRes).Move frmSize(Index).Left / 2 - 50, frmSize(Index).Top
      Case VertRes
           frmSize(Index).Top = frmSize(Index).Top + Y - FrmSizeMD(Index).Y
           frmSize(HorRes).Top = frmSize(Index).Top / 2 - 50
           PicFrm.Height = frmSize(Index).Top - 1
           frmSize(DiagRes).Move frmSize(HorRes).Left, frmSize(Index).Top
    End Select
  End If
End Sub
