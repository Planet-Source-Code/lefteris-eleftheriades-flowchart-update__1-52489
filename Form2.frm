VERSION 5.00
Begin VB.Form Display 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2805
      Left            =   45
      ScaleHeight     =   183
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   15
      Width           =   3900
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   915
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2925
      End
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 And RunningFlow Then
      Picture1.CurrentX = Text1.Left
      Picture1.CurrentY = Text1.Top + 0
      Picture1.Print Text1.Text
      'FlowChart.Command1_Click
      FlowChart.CoolerBar2_Click 2
      'Running = True
   End If
End Sub
