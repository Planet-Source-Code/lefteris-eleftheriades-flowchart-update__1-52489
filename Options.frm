VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2340
      Left            =   15
      TabIndex        =   1
      Top             =   30
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   4128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Options.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Toolbar"
      TabPicture(1)   =   "Options.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture5"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   -74955
         ScaleHeight     =   1920
         ScaleWidth      =   4215
         TabIndex        =   20
         Top             =   360
         Width           =   4215
         Begin VB.CheckBox Check8 
            Caption         =   "Show Status bar"
            Height          =   210
            Left            =   45
            TabIndex        =   27
            Top             =   1560
            Width           =   1485
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Show quick modifier"
            Height          =   240
            Left            =   45
            TabIndex        =   26
            Top             =   1290
            Width           =   1770
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Show Size boxes"
            Height          =   240
            Left            =   45
            TabIndex        =   25
            Top             =   1035
            Width           =   1545
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Tooltips"
            Height          =   210
            Left            =   45
            TabIndex        =   23
            Top             =   810
            Width           =   900
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Office XP"
            Height          =   210
            Index           =   2
            Left            =   30
            TabIndex        =   22
            Top             =   495
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Internet Explorer"
            Height          =   210
            Index           =   3
            Left            =   30
            TabIndex        =   21
            Top             =   270
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   2055
            X2              =   2055
            Y1              =   0
            Y2              =   1920
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   2100
            X2              =   2100
            Y1              =   0
            Y2              =   1920
         End
         Begin VB.Label Label4 
            Caption         =   "Style:"
            Height          =   210
            Left            =   15
            TabIndex        =   24
            Top             =   45
            Width           =   420
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   30
         ScaleHeight     =   1920
         ScaleWidth      =   4215
         TabIndex        =   2
         Top             =   360
         Width           =   4215
         Begin VB.Frame Frame1 
            Caption         =   "Grid"
            Height          =   1305
            Index           =   0
            Left            =   2340
            TabIndex        =   12
            Top             =   555
            Width           =   1845
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "Options.frx":0038
               Left            =   945
               List            =   "Options.frx":004E
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   180
               Width           =   825
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00404040&
               Height          =   270
               Left            =   945
               ScaleHeight     =   210
               ScaleWidth      =   735
               TabIndex        =   14
               Top             =   540
               Width           =   795
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               Height          =   270
               Left            =   945
               ScaleHeight     =   210
               ScaleWidth      =   720
               TabIndex        =   13
               Top             =   870
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "Space:"
               Height          =   225
               Left            =   75
               TabIndex        =   18
               Top             =   255
               Width           =   570
            End
            Begin VB.Label Label2 
               Caption         =   "Color:"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   17
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Label2 
               Caption         =   "BackColor:"
               Height          =   225
               Index           =   1
               Left            =   75
               TabIndex        =   16
               Top             =   930
               Width           =   795
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1890
            Left            =   30
            TabIndex        =   4
            Top             =   -30
            Width           =   2250
            Begin VB.CheckBox Check1 
               Caption         =   "Source Code Display"
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   30
               Value           =   1  'Checked
               Width           =   1875
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Pascal"
               Height          =   210
               Index           =   0
               Left            =   135
               TabIndex        =   10
               Top             =   285
               Width           =   1065
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Visual Basic"
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   9
               Top             =   525
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Keyword Highliting"
               Height          =   225
               Left            =   150
               TabIndex        =   8
               Top             =   810
               Value           =   1  'Checked
               Width           =   1635
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00FF0000&
               Height          =   255
               Left            =   1815
               ScaleHeight     =   195
               ScaleWidth      =   225
               TabIndex        =   7
               Top             =   795
               Width           =   285
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Dock in main Window"
               Height          =   195
               Left            =   150
               TabIndex        =   6
               Top             =   1125
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Export Source code on save"
               Height          =   360
               Left            =   150
               TabIndex        =   5
               Top             =   1395
               Width           =   1875
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2355
            TabIndex        =   3
            Text            =   "FixedSys"
            Top             =   225
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Defaut Font:"
            Height          =   210
            Left            =   2355
            TabIndex        =   19
            Top             =   0
            Width           =   900
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   3405
      TabIndex        =   0
      Top             =   2415
      Width           =   915
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
