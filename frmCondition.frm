VERSION 5.00
Begin VB.Form frmCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condition"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   300
      Left            =   1365
      TabIndex        =   1
      Top             =   435
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   3105
   End
End
Attribute VB_Name = "frmCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   FlowChart.FlowShape1(SelectedShape).Caption = Text1.Text
   Unload Me
End Sub

