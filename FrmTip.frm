VERSION 5.00
Begin VB.Form FrmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "FrmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "More"
      Height          =   330
      Left            =   1485
      TabIndex        =   6
      Top             =   1530
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   630
      Top             =   1470
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   2
      Left            =   120
      Picture         =   "FrmTip.frx":000C
      ScaleHeight     =   825
      ScaleWidth      =   720
      TabIndex        =   4
      Top             =   285
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   1
      Left            =   4425
      Picture         =   "FrmTip.frx":1F3E
      ScaleHeight     =   825
      ScaleWidth      =   720
      TabIndex        =   3
      Top             =   2580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   0
      Left            =   3705
      Picture         =   "FrmTip.frx":3E70
      ScaleHeight     =   825
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   2580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1410
      Left            =   75
      ScaleHeight     =   1350
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   60
      Width           =   4515
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   780
         TabIndex        =   5
         Top             =   45
         Width           =   3645
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   330
      Left            =   2355
      TabIndex        =   0
      Top             =   1530
      Width           =   855
   End
End
Attribute VB_Name = "FrmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   Dim B As String
   'Get a different tip
   Do
   B = LoadResString(101 + Int(Rnd * 6))
   Loop Until B <> Label1.Caption
   'Show it
   Label1.Caption = B
End Sub

Private Sub Form_Load()
   Randomize
   Label1.Caption = LoadResString(101 + Int(Rnd * 7))
End Sub

Private Sub Timer1_Timer()
Static K As Boolean
K = Not K
Picture2(2).Picture = Picture2(K + 1).Picture
End Sub
