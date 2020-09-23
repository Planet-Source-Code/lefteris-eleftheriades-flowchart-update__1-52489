VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1380
      Width           =   3345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   1365
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   360
      Left            =   3750
      TabIndex        =   0
      Top             =   1725
      Width           =   825
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1305
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   2302
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Define"
      TabPicture(0)   =   "frmProcess.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Arithmetic Operation"
      TabPicture(1)   =   "frmProcess.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(1)=   "Label3(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "String Operation"
      TabPicture(2)   =   "frmProcess.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4(1)"
      Tab(2).Control(1)=   "Label6(1)"
      Tab(2).Control(2)=   "Label3(1)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Boolean"
      TabPicture(3)   =   "frmProcess.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3(2)"
      Tab(3).Control(1)=   "Label4(2)"
      Tab(3).ControlCount=   2
      Begin VB.Label Label3 
         Caption         =   $"frmProcess.frx":0070
         Height          =   870
         Index           =   3
         Left            =   135
         TabIndex        =   12
         Top             =   375
         Width           =   4440
      End
      Begin VB.Label Label4 
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   -73500
         TabIndex        =   5
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label4 
         Caption         =   "+ - * / ( )  ^"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   -73455
         TabIndex        =   7
         Top             =   780
         Width           =   1725
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Note: All strings should be included in double quotes   ""Hello"""
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   -71910
         TabIndex        =   2
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   $"frmProcess.frx":0104
         Height          =   870
         Index           =   0
         Left            =   -74880
         TabIndex        =   8
         Top             =   390
         Width           =   4440
      End
      Begin VB.Label Label3 
         Caption         =   $"frmProcess.frx":0199
         Height          =   840
         Index           =   1
         Left            =   -74910
         TabIndex        =   6
         Top             =   390
         Width           =   4440
      End
      Begin VB.Label Label3 
         Caption         =   $"frmProcess.frx":0231
         Height          =   840
         Index           =   2
         Left            =   -74895
         TabIndex        =   4
         Top             =   420
         Width           =   3210
      End
      Begin VB.Label Label4 
         Caption         =   "AND OR NOT XOR "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -73500
         TabIndex        =   3
         Top             =   780
         Width           =   2265
      End
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1095
      TabIndex        =   11
      Top             =   1365
      Width           =   105
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Equation As String
Private Sub Command1_Click()
  Equation = Combo1.Text & " = " & Text1.Text
  FlowChart.FlowShape1(SelectedShape).Caption = Equation
  Unload Me
End Sub

