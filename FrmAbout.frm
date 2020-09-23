VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flowchart Assistant 2004"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1185
      TabIndex        =   3
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   $"FrmAbout.frx":000C
      Height          =   1095
      Left            =   285
      TabIndex        =   2
      Top             =   510
      Width           =   3045
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FlowChart Assistant 2004"
      BeginProperty Font 
         Name            =   "Tiranti Solid LET"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8000&
      Height          =   405
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   3270
   End
   Begin VB.Label Label1 
      Caption         =   "FlowChart Assistant 2004"
      BeginProperty Font 
         Name            =   "Tiranti Solid LET"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   60
      Width           =   3270
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub
