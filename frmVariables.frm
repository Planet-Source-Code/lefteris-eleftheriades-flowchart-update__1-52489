VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variable Declaration"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2445
      TabIndex        =   6
      Top             =   1710
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2820
      TabIndex        =   5
      Top             =   690
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2820
      TabIndex        =   4
      Top             =   1320
      Width           =   360
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frmVariables.frx":0000
      Left            =   60
      List            =   "frmVariables.frx":0007
      TabIndex        =   3
      Top             =   30
      Width           =   2715
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmVariables.frx":0017
      Left            =   1740
      List            =   "frmVariables.frx":002D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1335
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Variable:                      Type:"
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   1110
      Width           =   2055
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Combo1_Click()
    If List1.ListIndex > 0 Then
       List1.List(List1.ListIndex) = Text1.Text & " As " & Combo1.Text
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      List1.AddItem Text1.Text & " As " & Combo1.Text
   Else
      If List1.ListIndex > 0 Then List1.RemoveItem List1.ListIndex
   End If
End Sub

Private Sub Command2_Click()
   Dim VarLst As String, VList2 As String
   On Error Resume Next
   If List1.ListCount > 0 Then
      For I = 1 To List1.ListCount - 1
          VarLst = VarLst & ",  " & List1.List(I)
          VList2 = VList2 & ", " & Mid(List1.List(I), 1, InStr(1, List1.List(I), " ") - 1)
      Next
   End If
   
   FlowChart.FlowShape1(SelectedShape).Tag = "Dim " & Trim(Mid(VList2, 2))
   FlowChart.FlowShape1(SelectedShape).ControlData = "Dim " & Trim(Mid(VarLst, 2))
   Unload Me
End Sub

Private Sub Form_Load()
   Combo1.ListIndex = 0
End Sub

Private Sub List1_Click()
   Dim Words() As String
   If List1.ListIndex > 0 Then
      Words = Split(List1.List(List1.ListIndex), " ")
      Text1.Text = Words(0)
      Combo1.ListIndex = SendMessage(Combo1.hwnd, CB_FINDSTRINGEXACT, -1, ByVal CStr(Words(2)))
   End If
End Sub

Private Sub Text1_Change()
    If List1.ListIndex > 0 Then
       List1.List(List1.ListIndex) = Text1.Text & " As " & Combo1.Text
    End If
End Sub
