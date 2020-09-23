VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   360
      Left            =   2415
      TabIndex        =   2
      Top             =   720
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   345
      Width           =   3120
   End
   Begin VB.Label Label1 
      Caption         =   "Prompt (use %Variable to display variables):"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   3075
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  FlowChart.FlowShape1(SelectedShape).Caption = Text1.Text
  Unload Me
End Sub
