VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "As Command"
      Height          =   225
      Left            =   135
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   360
      Left            =   2430
      TabIndex        =   3
      Top             =   855
      Width           =   825
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmInput.frx":0000
      Left            =   1200
      List            =   "frmInput.frx":0002
      TabIndex        =   1
      Top             =   420
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   765
      TabIndex        =   0
      Text            =   """Enter Value: """
      Top             =   75
      Width           =   2460
   End
   Begin VB.Label Label2 
      Caption         =   "Store variable:"
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   510
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Prompt:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   840
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  FlowChart.FlowShape1(SelectedShape).Caption = Combo1.Text & " = " & Text1.Text
  FlowChart.FlowShape1(SelectedShape).ControlData = Text1.Text
  Unload Me
End Sub

