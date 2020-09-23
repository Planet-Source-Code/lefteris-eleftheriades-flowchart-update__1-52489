VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FlowChart 
   Caption         =   "FlowChart Assistant - [Flow1.flc]"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8625
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   StartUpPosition =   3  'Windows Default
   Begin FlowChartEditor.CoolerBar CoolerBar3 
      Height          =   255
      Left            =   8340
      TabIndex        =   28
      Top             =   540
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHeight    =   14
      ButtonWidth     =   15
      ButtonImages    =   "Form1.frx":1042
      PTmp            =   "Form1.frx":1334
      Style           =   1
      Margin          =   2
   End
   Begin FlowChartEditor.CoolerBar CoolerBar2 
      Height          =   450
      Left            =   5535
      TabIndex        =   27
      Top             =   0
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHeight    =   27
      ButtonWidth     =   27
      ButtonImages    =   "Form1.frx":1B66
      PTmp            =   "Form1.frx":57A0
      Style           =   1
      Margin          =   2
   End
   Begin FlowChartEditor.CoolerBar CoolerBar1 
      Height          =   435
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHeight    =   26
      ButtonWidth     =   44
      ButtonImages    =   "Form1.frx":10BAA
      PTmp            =   "Form1.frx":1773C
      Style           =   1
      Margin          =   2
   End
   Begin VB.PictureBox ModeImg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   8265
      Picture         =   "Form1.frx":2B94E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   2265
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3795
      Top             =   3045
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Flow Diagrams|*.flc"
   End
   Begin VB.PictureBox ModeImg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8265
      Picture         =   "Form1.frx":2BC90
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   2550
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox ModeImg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   8265
      Picture         =   "Form1.frx":2BFD2
      ScaleHeight     =   300
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   1935
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSScriptControlCtl.ScriptControl VBS1 
      Left            =   6795
      Top             =   915
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
      UseSafeSubset   =   -1  'True
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2745
      TabIndex        =   5
      Top             =   555
      Width           =   5580
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   630
      TabIndex        =   15
      Top             =   510
      Width           =   630
      Begin VB.CommandButton Command4 
         Height          =   150
         Index           =   1
         Left            =   420
         Picture         =   "Form1.frx":2C3D4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Text            =   "0"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton Command4 
         Height          =   150
         Index           =   0
         Left            =   420
         Picture         =   "Form1.frx":2C476
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   345
      ScaleHeight     =   270
      ScaleWidth      =   630
      TabIndex        =   12
      Top             =   510
      Width           =   630
      Begin VB.CommandButton Command3 
         Height          =   150
         Index           =   1
         Left            =   420
         Picture         =   "Form1.frx":2C518
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   225
      End
      Begin VB.CommandButton Command3 
         Height          =   150
         Index           =   0
         Left            =   420
         Picture         =   "Form1.frx":2C5BA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Text            =   "0"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.CommandButton ToolButton 
      Enabled         =   0   'False
      Height          =   435
      Index           =   7
      Left            =   8160
      Picture         =   "Form1.frx":2C65C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Toggle Display Window"
      Top             =   2775
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Picture3 
      Height          =   5685
      Left            =   0
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   18
      Top             =   870
      Width           =   8055
      Begin VB.VScrollBar VScroll1 
         Height          =   5385
         LargeChange     =   10
         Left            =   7755
         Max             =   500
         TabIndex        =   6
         Top             =   0
         Value           =   250
         Width           =   240
      End
      Begin VB.PictureBox WorkSpace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   5400
         Left            =   0
         ScaleHeight     =   360
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   517
         TabIndex        =   0
         Top             =   0
         Width           =   7755
         Begin FlowChartEditor.FlowShape FlowShape1 
            Height          =   510
            Index           =   0
            Left            =   3285
            TabIndex        =   25
            Top             =   90
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   900
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8,25
            CurrentY        =   24
            DrawWidth       =   2
            FillColor       =   16777215
            FillStyle       =   0
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   1710
            Left            =   6195
            ScaleHeight     =   1710
            ScaleWidth      =   1530
            TabIndex        =   22
            Top             =   3660
            Visible         =   0   'False
            Width           =   1530
            Begin VB.ListBox List1 
               Height          =   1620
               Left            =   45
               TabIndex        =   23
               Top             =   45
               Width           =   1440
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00404040&
               Index           =   3
               X1              =   1530
               X2              =   0
               Y1              =   1695
               Y2              =   1695
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00404040&
               Index           =   2
               X1              =   1515
               X2              =   1515
               Y1              =   1710
               Y2              =   0
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   1530
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   0
               X2              =   0
               Y1              =   1710
               Y2              =   0
            End
         End
         Begin VB.Shape oSelection 
            BorderStyle     =   3  'Dot
            Height          =   1275
            Left            =   2835
            Top             =   15
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Line arrDown 
            BorderWidth     =   2
            Index           =   0
            Visible         =   0   'False
            X1              =   244
            X2              =   255
            Y1              =   67
            Y2              =   80
         End
         Begin VB.Line arrUp 
            BorderWidth     =   2
            Index           =   0
            Visible         =   0   'False
            X1              =   265
            X2              =   255
            Y1              =   67
            Y2              =   81
         End
         Begin VB.Line ConnectingLine 
            BorderWidth     =   2
            Index           =   0
            Visible         =   0   'False
            X1              =   255
            X2              =   255
            Y1              =   44
            Y2              =   79
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lefteris Eleftheriades"
            Height          =   210
            Left            =   6240
            TabIndex        =   19
            Top             =   5160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Shape sSelection 
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            Height          =   585
            Left            =   3255
            Top             =   60
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   10
         Left            =   0
         Max             =   500
         TabIndex        =   7
         Top             =   5400
         Value           =   250
         Width           =   7770
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5700
      Left            =   0
      ScaleHeight     =   376
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   20
      Top             =   870
      Visible         =   0   'False
      Width           =   8055
      Begin FlowChartEditor.CodeBox txtCode 
         Height          =   5640
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   9948
         BackColor       =   16777215
         CurrentX        =   1000
         CurrentY        =   639
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keywords        =   $"Form1.frx":2D29E
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   12
      X1              =   135
      X2              =   135
      Y1              =   32
      Y2              =   54
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   11
      X1              =   90
      X2              =   90
      Y1              =   31
      Y2              =   53
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   66
      X2              =   66
      Y1              =   32
      Y2              =   54
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   9
      X1              =   21
      X2              =   21
      Y1              =   32
      Y2              =   54
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   21
      X2              =   66
      Y1              =   53
      Y2              =   53
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   91
      X2              =   136
      Y1              =   53
      Y2              =   53
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   91
      X2              =   136
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   22
      X2              =   67
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   574
      X2              =   574
      Y1              =   35
      Y2              =   54
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   181
      X2              =   181
      Y1              =   36
      Y2              =   54
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   182
      X2              =   574
      Y1              =   53
      Y2              =   53
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   182
      X2              =   574
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Data:"
      Height          =   165
      Left            =   2220
      TabIndex        =   17
      Top             =   555
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "H:"
      Height          =   180
      Index           =   1
      Left            =   1125
      TabIndex        =   14
      Top             =   555
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "W:"
      Height          =   165
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   585
      Width           =   195
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewItm 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu zhyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveItm 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAsItm 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu zhyp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Print P&review"
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu zhyp3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHolderExport 
         Caption         =   "&Export Source"
         Visible         =   0   'False
         Begin VB.Menu mnuExportVisualBasicCode 
            Caption         =   "&Visual Basic"
         End
         Begin VB.Menu mnuExportPascalCode 
            Caption         =   "&Pascal"
         End
      End
      Begin VB.Menu mnuMakeExeItm 
         Caption         =   "&Make .EXE"
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu zhyp4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExitItm 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuFlowView 
         Caption         =   "Flowcharts"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSourceCodeView 
         Caption         =   "Source Code"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTipItm 
         Caption         =   "&Tip"
      End
      Begin VB.Menu mnuAboutItm 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FlowChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function PutFocus Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type ConnectLine
   PreviewsShape As Integer
   NextShape As Integer
   ParaShape As Integer
   Exists As Boolean
End Type

Private Type aVariable
   vName As String
   vValue As String
End Type
Dim MDX&, MDY&, sInd%, GridSpace As Long
Dim NoOfObjects As Integer, NoOfLines As Integer
Dim ConnLne(1 To 150) As ConnectLine
Dim ConLne(1 To 150, 0 To 1) As Integer
Dim LineInvolved(150, 150) As Byte
Dim SelectedConnLine%
Dim ClickRec(1) As Integer
Dim LineAffectedIndex(1 To 3) As Integer
Dim SelectedLineIndex As Integer
Dim oSelX As Single, oSely As Single
Dim NotSaved As Boolean
Dim CStep As Integer
Dim SaveFileName As String
Dim VEd As String
Dim NextStepP As Integer
Dim Watches() As String
Private Type SaveShape
   X As Integer '2
   Y As Integer '2
   W As Byte '1
   H As Byte '1
   Shape As Byte '1
   Title As String 'X+1
   Caption As String 'X+1
   ControlData As String 'X+1
   Next1 As Integer
   Next2 As Integer
End Type 'Total 10 + length of strings bytes
Private Type SaveLineInvolved
   FlowShape1Index As Byte
   FlowShape2Index As Byte
   LineIndex As Byte
End Type
Dim ShitToSave() As SaveShape
Dim LineShitToSave() As SaveLineInvolved
Dim OriginalTop(100) As Long
Dim OriginalLeft(100) As Long
Dim PrevIndex As Integer, PrevIndex2 As Integer
Dim ToolTipHandles(20) As Long
Dim MuseDn As Boolean
Dim VScrollV&, HScrollV&
Dim InExternalRun As Boolean
Sub RunFlowChart(Optional UnloadEditor As Boolean = False)
Dim Eflag As Boolean
RunningFlow = True
 If CStep = 0 Then
    Display.Show
    SetWindowPos Display.hwnd, -1, 0, 0, 0, 0, 83

    For I = 1 To FlowShape1().Count - 1
        If FlowShape1(I).Title = "Start" Then
           CStep = I
           Exit For
        End If
    Next I
    Display.Picture1.Cls
    If CStep = 0 Then
       MsgBox "The Program must begin with a ""Start""", vbInformation
       Exit Sub
    End If
 End If
 Do
 sSelection.Move FlowShape1(CStep).Left, FlowShape1(CStep).Top - 2, FlowShape1(CStep).Width + 2, FlowShape1(CStep).Height + 5
 If FlowShape1(CStep).Title = "Input" Then Eflag = True
 CStep = ExecuteStep(CStep)
 If CStep = 0 Then RunningFlow = False
 Loop Until CStep = 0 Or Eflag
End Sub

Function ExecuteStep(Step As Integer) As Integer
   Dim IFRet As Boolean, Fret$
   Dim wIsolator() As String
   'On Error Resume Next
   If Display.Text1.Visible = True Then
      Display.Picture1.CurrentX = Display.Text1.Left
      Display.Picture1.CurrentY = Display.Text1.Top + 0
      Display.Picture1.Print Display.Text1.Text
      Display.Text1.Visible = False
      If Display.Text1.Text <> "" Then
         If IsNumeric(Display.Text1.Text) Then
            VBS1.ExecuteStatement "On Error resume next: " & VEd & " = " & Display.Text1.Text
         Else
            VBS1.ExecuteStatement "On Error resume next: " & VEd & " = """ & Display.Text1.Text & """"
         End If
      Else
         VBS1.ExecuteStatement "On Error resume next: " & VEd & " = " & """" & """"
      End If
      Debug.Print VEd & " = " & Display.Text1.Text
      Display.Text1.Text = ""
   End If
                
   Select Case FlowShape1(Step).Title
       Case "Output"
             If FlowShape1(Step).Caption <> "" Then
                VBS1.ExecuteStatement "On Error resume next: Out = " & FlowShape1(Step).Caption
                Display.Picture1.Print CStr(VBS1.Eval("Out"))
                Debug.Print VBS1.Eval("Out")
                Display.Picture1.Refresh
             End If
        Case "Process"
             If FlowShape1(Step).Caption <> "" Then
                VBS1.ExecuteStatement "On Error resume next: " & FlowShape1(Step).Caption
             
             End If
        Case "Input"
             If FlowShape1(Step).Caption <> "" Then
                VBS1.ExecuteStatement "On Error resume next: Out =  " & FlowShape1(Step).ControlData
                Fret = VBS1.Eval("Out")
                Display.Text1.Left = Display.Picture1.TextWidth(Fret)
                Display.Text1.Top = Display.Picture1.CurrentY '- 1
                Display.Picture1.Print Fret
                If Fret = "" Then MsgBox "Error"
                Display.Picture1.Refresh
                Display.Text1.Visible = True
                DoEvents
                Display.Text1.SetFocus
                'PutFocus Display.Text1.hwnd
                'SetActiveWindow Display.Text1.hwnd
                wIsolator = Split(FlowShape1(Step).Caption, "=")
                VEd = Trim(wIsolator(0))
                'Running = False
                'Do
                '   DoEvents
                'Loop Until Running
                
                'Display.Text1.Visible = False
             End If
         Case "IF"
               VBS1.ExecuteStatement "On Error resume next: vIF = " & FlowShape1(Step).Caption
               IFRet = VBS1.Eval("vIF")
               Debug.Print FlowShape1(Step).Caption & "-->" & IFRet
   End Select
   List1.Clear
   On Error Resume Next
   For I = 0 To UBound(Watches)
       List1.AddItem Watches(I) & " = " & VBS1.Eval(Watches(I))
   Next
   
   If FlowShape1(Step).Title = "IF" Then
      ExecuteStep = GetNextControl(Step, IFRet)
   Else
      ExecuteStep = GetNextControl(Step)
   End If
   
   
End Function
Function GetNextControl(ByVal Current As Integer, Optional TrueConditionCase As Boolean = False)
   GetNextControl = ConLne(Current, -TrueConditionCase)
   'For I = 1 To NoOfLines
   '    If ConnLne(I).PreviewsShape = Current Then
   '       GetNextControl = ConnLne(I).NextShape
   '       Exit For
   '    End If
   'Next I
End Function

Private Sub Command3_Click(Index As Integer)
     If Index = 0 Then
        FlowShape1(SelectedShape).Move FlowShape1(SelectedShape).Left - 10, FlowShape1(SelectedShape).Top, FlowShape1(SelectedShape).Width + 20
     Else
        If FlowShape1(SelectedShape).Width > 40 Then
           FlowShape1(SelectedShape).Move FlowShape1(SelectedShape).Left + 10, FlowShape1(SelectedShape).Top, FlowShape1(SelectedShape).Width - 20
        End If
     End If
     sSelection.Move FlowShape1(SelectedShape).Left - 2, FlowShape1(SelectedShape).Top - 2, FlowShape1(SelectedShape).Width + 4, FlowShape1(SelectedShape).Height + 4
     Text1(0).Text = FlowShape1(SelectedShape).Width
     FlowShape1(SelectedShape).ShapeControl
End Sub

Private Sub Command4_Click(Index As Integer)
     If Index = 0 Then
        FlowShape1(SelectedShape).Move FlowShape1(SelectedShape).Left, FlowShape1(SelectedShape).Top - 10, FlowShape1(SelectedShape).Width, FlowShape1(SelectedShape).Height + 20
     Else
        If FlowShape1(SelectedShape).Height > 40 Then
           FlowShape1(SelectedShape).Move FlowShape1(SelectedShape).Left, FlowShape1(SelectedShape).Top + 10, FlowShape1(SelectedShape).Width, FlowShape1(SelectedShape).Height - 20
        End If
     End If
     sSelection.Move FlowShape1(SelectedShape).Left - 2, FlowShape1(SelectedShape).Top - 2, FlowShape1(SelectedShape).Width + 4, FlowShape1(SelectedShape).Height + 4
     Text1(1).Text = FlowShape1(SelectedShape).Height
     FlowShape1(SelectedShape).ShapeControl
End Sub

Function EndOfIf(IfIndex%) As Integer
Dim AAr(20) As Integer, BAr(20) As Integer
A = GetNextControl(IfIndex, True)
AAr(0) = A
B = GetNextControl(IfIndex, False)
BAr(0) = B
For I = 1 To 20
    A = GetNextControl(A, False)
    AAr(I) = A
    B = GetNextControl(B, False)
    BAr(I) = B
Next
For I = 1 To 20
   For J = 1 To 20
      If AAr(I) = BAr(I) Then
         EndOfIf = I
         GoTo Z
      End If
   Next
Next
Z:
End Function

Sub SwitchView()
   Dim wIsolator() As String
   Dim CLine&
   Static ProgMode As Byte
      
   ProgMode = ProgMode + 1
   If ProgMode > 2 Then ProgMode = 0
'   Command5.Picture = ModeImg(ProgMode).Picture
   For I = 1 To 30
       txtCode.TextboxLine(I) = ""
   Next
   If ProgMode = 0 Then
      Picture2.Visible = False
      Picture3.Visible = True

   ElseIf ProgMode = 1 Then
      For I = 1 To FlowShape1().Count - 1
        If FlowShape1(I).Title = "Start" Then
           FlowPoss = I
           Exit For
        End If
      Next
      txtCode.TextboxLine(0) = FlowShape1(FlowPoss).ControlData & vbCrLf
      txtCode.TextboxLine(1) = "Sub Main()" & vbCrLf
      CLine = 1
      On Error Resume Next
      Do
         CLine = CLine + 1
         Erase wIsolator()
         wIsolator = Split(FlowShape1(GetNextControl(FlowPoss)).Caption, " ")

         If FlowShape1(GetNextControl(FlowPoss)).Title = "Stop" Then Exit Do
         If FlowShape1(GetNextControl(FlowPoss)).Title = "" Then
         '   'Flowchart incomplete
         '   Exit Do
         End If
         Select Case FlowShape1(GetNextControl(FlowPoss)).Title
            Case "Output": txtCode.TextboxLine(CLine) = "   MsgBox " & FlowShape1(GetNextControl(FlowPoss)).Caption & vbCrLf
            Case "Process": txtCode.TextboxLine(CLine) = "   " & FlowShape1(GetNextControl(FlowPoss)).Caption & vbCrLf
            Case "Input": txtCode.TextboxLine(CLine) = "   " & wIsolator(0) & " = Inputbox(" & FlowShape1(GetNextControl(FlowPoss)).ControlData & ")" & vbCrLf
            Case "IF"
                  txtCode.TextboxLine(CLine) = "   If " & FlowShape1(GetNextControl(FlowPoss)).Caption & " Then"
           '       MsgBox FlowShape1(EndOfIf(GetNextControl(FlowPoss))).Title
         End Select
         FlowPoss = GetNextControl(FlowPoss)
      Loop
      txtCode.TextboxLine(CLine) = "End Sub"
      
      Picture2.Visible = True
      Picture3.Visible = False
   ElseIf ProgMode = 2 Then
      
      For I = 1 To FlowShape1().Count - 1
        If FlowShape1(I).Title = "Start" Then
           FlowPoss = I
           Exit For
        End If
      Next
      txtCode.TextboxLine(0) = "Program PascalFlow;"
      txtCode.TextboxLine(1) = Replace(Replace(FlowShape1(FlowPoss).ControlData, "Dim", "Var", , , vbTextCompare), ",", ";")
      txtCode.TextboxLine(2) = "Begin " & vbCrLf
      CLine = 2
      On Error Resume Next
      Do
         CLine = CLine + 1
         Erase wIsolator()
         wIsolator = Split(FlowShape1(GetNextControl(FlowPoss)).Caption, " ")

         If FlowShape1(GetNextControl(FlowPoss)).Title = "Stop" Then Exit Do
         If FlowShape1(GetNextControl(FlowPoss)).Title = "" Then
            'Flowchart incomplete
            'Exit Do
         End If
         Select Case FlowShape1(GetNextControl(FlowPoss)).Title
            Case "Output": txtCode.TextboxLine(CLine) = "    Writeln(" & Replace(Replace(FlowShape1(GetNextControl(FlowPoss)).Caption, """", "'"), "&", ",") & ");" & vbCrLf
            Case "Process": txtCode.TextboxLine(CLine) = "    " & Replace(Replace(Replace(FlowShape1(GetNextControl(FlowPoss)).Caption, "=", ":="), "&", "+"), """", "'") & ";" & vbCrLf
            Case "Input"
                         txtCode.TextboxLine(CLine) = "    Write(" & Replace(Replace(FlowShape1(GetNextControl(FlowPoss)).ControlData, """", "'"), "&", ",") & ");"
                         CLine = CLine + 1
                         txtCode.TextboxLine(CLine) = "    Readln(" & wIsolator(0) & ");" & vbCrLf
                         
            Case "IF"
                  txtCode.TextboxLine(CLine) = "    If " & Replace(FlowShape1(GetNextControl(FlowPoss)).Caption, """", "'") & " Then"
           '       MsgBox FlowShape1(EndOfIf(GetNextControl(FlowPoss))).Title
         End Select
         FlowPoss = GetNextControl(FlowPoss)
      Loop
      txtCode.TextboxLine(CLine) = "End."
      
      Picture2.Visible = True
      Picture3.Visible = False
   End If
End Sub

Sub StepFlowChart()
 If NextStepP = 0 Then
    Display.Show
    SetWindowPos Display.hwnd, -1, 0, 0, 0, 0, 83

    For I = 1 To FlowShape1().Count - 1
        If FlowShape1(I).Title = "Start" Then
           NextStepP = I
           Exit For
        End If
    Next I
    Display.Picture1.Cls
    If NextStepP = 0 Then
       MsgBox "The Program must begin with a ""Start""", vbInformation
       Exit Sub
    End If
 End If
 sSelection.Move FlowShape1(NextStepP).Left, FlowShape1(NextStepP).Top - 2, FlowShape1(NextStepP).Width + 2, FlowShape1(NextStepP).Height + 5
 NextStepP = ExecuteStep(NextStepP)
End Sub

Private Sub CoolerBar1_Click(Index As Integer)
   If Index <= 5 Then
     NotSaved = True
     NoOfObjects = FlowShape1.Count
     Load FlowShape1(FlowShape1.Count)
     Select Case Index
        Case 0
                FlowShape1(NoOfObjects).Shape = Oval
                FlowShape1(NoOfObjects).Title = "Start"
        Case 1
                FlowShape1(NoOfObjects).Shape = Oval
                FlowShape1(NoOfObjects).Title = "Stop"
        Case 2
                FlowShape1(NoOfObjects).Width = FlowShape1(0).Width * 1.55
                FlowShape1(NoOfObjects).Shape = Parallel
                FlowShape1(NoOfObjects).Title = "Input"
        Case 3
                FlowShape1(NoOfObjects).Shape = Rectangle
                FlowShape1(NoOfObjects).Title = "Process"
        
        Case 4
                FlowShape1(NoOfObjects).Width = FlowShape1(0).Width * 1.55
                FlowShape1(NoOfObjects).Shape = Parallel
                FlowShape1(NoOfObjects).Title = "Output"
        Case 5
                FlowShape1(NoOfObjects).Title = "IF"
                FlowShape1(NoOfObjects).Caption = ""
                FlowShape1(NoOfObjects).Width = FlowShape1(0).Width * 2.1
                FlowShape1(NoOfObjects).Height = FlowShape1(0).Height * 1.5
                FlowShape1(NoOfObjects).Shape = Lozenge
        Case 6
                'FlowShape1(NoOfObjects).Title = "Function"
                'FlowShape1(NoOfObjects).Width = FlowShape1(0).Width * 1.55
                'FlowShape1(NoOfObjects).Shape = Diamond
        Case 7
                'FlowShape1(NoOfObjects).Title = "Start"
                'FlowShape1(NoOfObjects).Width = FlowShape1(0).Width * 1.55
                'FlowShape1(NoOfObjects).Shape = DoubleRectangle
        Case 8
                'FlowShape1(NoOfObjects).Shape = Rectangle
                'FlowShape1(NoOfObjects).Title = "Window"
                'FlowShape1(NoOfObjects).Caption = "Toggle"
     End Select
     FlowShape1(NoOfObjects).Left = (44 + 4) * Index - FlowShape1(NoOfObjects).Width / 2 + 8
     If Index = 0 Then FlowShape1(NoOfObjects).Left = 3
     sSelection.Move FlowShape1(NoOfObjects).Left, FlowShape1(NoOfObjects).Top - 2, FlowShape1(NoOfObjects).Width + 2, FlowShape1(NoOfObjects).Height + 5
     FlowShape1(NoOfObjects).ZOrder
     DoEvents
     FlowShape1(NoOfObjects).ShapeControl
     SelectedShape = NoOfObjects
     FlowShape1(NoOfObjects).Visible = True
     sSelection.Visible = True
   End If
End Sub

Private Sub CoolerBar1_MouseMove(buttonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If PrevIndex <> X \ (42 + 4) Then
    TipForeColor ToolTipHandles(0), &H0
    Select Case X \ (42 + 4)
      Case 0
             TipIconTitle ToolTipHandles(0), TTIconInfo, "Start"
             TipText "Use to begin your main code." & vbCrLf & "DoubleClick to set variables", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 1
             TipIconTitle ToolTipHandles(0), TTIconInfo, "Stop"
             TipText "Use to indicate the end of the" & vbCrLf & "main code or a function", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 2
             TipIconTitle ToolTipHandles(0), TTIconInfo, "Input"
             TipText "Request Data Input." & vbCrLf & "Can save both numbers, strings" & vbCrLf & "and mathematical expressions", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 3
             TipIconTitle ToolTipHandles(0), TTIconInfo, "Process"
             TipText "Request Data Input." & vbCrLf & "Performs mathematical, string and" & vbCrLf & "logical operations", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 4
             TipIconTitle ToolTipHandles(0), TTIconInfo, "Output"
             TipText "Request Data Input." & vbCrLf & "Displays the desired text or" & vbCrLf & "variable contents.", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 5
             TipIconTitle ToolTipHandles(0), TTIconInfo, "If"
             TipText "Changes the flow of the " & vbCrLf & "code in to an other root.", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 6
             TipForeColor ToolTipHandles(0), &H808080
             TipIconTitle ToolTipHandles(0), TTIconWarning, "Call"
             TipText "Calls external/internal function" & vbCrLf & "(Not Availiable)", False, CoolerBar1.hwnd, ToolTipHandles(0)
      Case 7
             TipForeColor ToolTipHandles(0), &H808080
             TipIconTitle ToolTipHandles(0), TTIconWarning, "Start Function"
             TipText "Declares an internal function" & vbCrLf & "(Not Availiable)", False, CoolerBar1.hwnd, ToolTipHandles(0)
    End Select
    PrevIndex = X \ (42 + 4)
  End If
End Sub

Sub CoolerBar2_Click(buttonIndex As Integer)
  Select Case buttonIndex
   Case 0
          NotSaved = True
          NoOfObjects = NoOfObjects + 1
          Load FlowShape1(NoOfObjects)
          
          FlowShape1(NoOfObjects).Width = 12
          FlowShape1(NoOfObjects).Height = 15
          FlowShape1(NoOfObjects).Shape = Oval
          
          FlowShape1(NoOfObjects).Left = CoolerBar2.Left
          sSelection.Move FlowShape1(NoOfObjects).Left, FlowShape1(NoOfObjects).Top - 2, FlowShape1(NoOfObjects).Width + 2, FlowShape1(NoOfObjects).Height + 5
          FlowShape1(NoOfObjects).ZOrder
          DoEvents
          FlowShape1(NoOfObjects).ShapeControl
          SelectedShape = NoOfObjects
          FlowShape1(NoOfObjects).Visible = True
          sSelection.Visible = True
   Case 1
          If ClickRec(0) <> ClickRec(1) And ClickRec(0) > 0 And ClickRec(1) > 0 Then
             If ConLne(ClickRec(1), 0) = 0 Then
                NoOfLines = NoOfLines + 1
                Load ConnectingLine(NoOfLines)
                Load arrUp(NoOfLines)
                Load arrDown(NoOfLines)
                ConLne(ClickRec(1), 0) = ClickRec(0)
                LineInvolved(ClickRec(1), ClickRec(0)) = NoOfLines
                ConnectingLineController ClickRec(1), ClickRec(0), NoOfLines
             Else
                'There already is a connection to an other shape
                If FlowShape1(ClickRec(1)).Shape = Lozenge And ConLne(ClickRec(1), 1) = 0 Then
                   'If the source shape is a condition, then you are allowed one more connection
                   NoOfLines = NoOfLines + 1
                   Load ConnectingLine(NoOfLines)
                   Load arrUp(NoOfLines)
                   Load arrDown(NoOfLines)
                   ConLne(ClickRec(1), 1) = ClickRec(0)
                   LineInvolved(ClickRec(1), ClickRec(0)) = NoOfLines
                   ConnectingLineController ClickRec(1), ClickRec(0), NoOfLines
                Else
                   MsgBox "You can not make an other connection for this shape", vbExclamation
                End If
             End If
             ConnectingLine(NoOfLines).Visible = True
             arrUp(NoOfLines).Visible = True
             arrDown(NoOfLines).Visible = True
           Else
             MsgBox "You have made an invalid selection" & vbCrLf & "Select the source shape, then the destination" & vbCrLf & "shape, and click this button again", vbInformation
           End If
    Case 2: RunFlowChart
    Case 3: StepFlowChart
    Case 4: FrmTip.Show vbModal, Me
    Case 5: SwitchView
    Case 6: Picture4.Visible = Not Picture4.Visible
  End Select
End Sub

Private Sub CoolerBar2_MouseMove(buttonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If PrevIndex2 <> X \ (27 + 4) Then
    Select Case X \ (27 + 4)
      Case 0
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Connection Pole"
             TipText "Use it as an intermediate pole for" & vbCrLf & "better presentation of your flowchart", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 1
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Connect"
             TipText "Select the source shape, then the" & vbCrLf & "destination shape and click here to connect them", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 2
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Run"
             TipText "Executes the code", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 3
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Step"
             TipText "Executes the code a shape each time", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 4
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Tip of the Day"
             TipText "Displays the tip of the day dialogue", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 5
             TipIconTitle ToolTipHandles(1), TTIconInfo, "View"
             TipText "Switch between viewing flowchart," & vbCrLf & "Visual Basic code or Pascal code.", False, CoolerBar2.hwnd, ToolTipHandles(1)
      Case 6
             TipIconTitle ToolTipHandles(1), TTIconInfo, "Watches"
             TipText "Shows the variable watch window" & vbCrLf & "Right-Click on the listbox to set variables", False, CoolerBar2.hwnd, ToolTipHandles(1)
    End Select
    PrevIndex2 = X \ (27 + 4)
  End If
End Sub

Private Sub FlowShape1_Click(Index As Integer)
   SelectedShape = Index
   
   Select Case FlowShape1(Index).Title
      Case "Process"
                    Text2.Text = FlowChart.FlowShape1(Index).Caption
      Case "Input"
                    Text2.Text = FlowChart.FlowShape1(Index).Caption
      Case "Output"
                    Text2.Text = FlowChart.FlowShape1(Index).Caption
      Case "Start": Text2.Text = FlowChart.FlowShape1(Index).ControlData
      Case "IF"
                    Text2.Text = FlowChart.FlowShape1(Index).Caption
   End Select

End Sub

Private Sub FlowShape1_DblClick(Index As Integer)
   SelectedShape = Index
   
   Select Case FlowShape1(Index).Title
      Case "Process"
                    frmProcess.Show , Me
                    frmProcess.Text1.Text = Trim(Replace(LastPart(FlowShape1(Index).Caption, " = "), "=", ""))
                    frmProcess.Combo1.Text = Trim(FirstPart(FlowShape1(Index).Caption, " = "))
      Case "Input"
                    frmInput.Show , Me
                    frmInput.Text1.Text = Trim(Replace(LastPart(FlowShape1(Index).Caption, " = "), "=", ""))
                    frmInput.Combo1.Text = Trim(FirstPart(FlowShape1(Index).Caption, " = "))
      Case "Output"
                    frmOutput.Show , Me
                    frmOutput.Text1.Text = FlowShape1(Index).Caption
      Case "Start": frmVariables.Show vbModal, Me
      Case "IF"
                    frmCondition.Show , Me
                    frmCondition.Text1.Text = FlowShape1(Index).Caption
   End Select
End Sub

Private Sub FlowShape1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim I%, J%
  'Delete a shape
  If KeyCode = vbKeyJ Then
     CoolerBar2_Click 1
  End If
  If KeyCode = vbKeyDelete Then
     MsgBox "Can Not Delete Anything", vbExclamation
     Exit Sub
     Unload FlowShape1(Index)
     
     On Error Resume Next
     If ConLne(Index, 0) <> 0 Then
        Unload ConnectingLine(LineInvolved(Index, ConLne(Index, 0)))
        Unload arrDown(LineInvolved(Index, ConLne(Index, 0)))
        Unload arrUp(LineInvolved(Index, ConLne(Index, 0)))
        ConLne(Index, 0) = 0
     End If
     
     If ConLne(Index, 1) <> 0 Then
        Unload ConnectingLine(LineInvolved(Index, ConLne(Index, 1)))
        Unload arrDown(LineInvolved(Index, ConLne(Index, 1)))
        Unload arrUp(LineInvolved(Index, ConLne(Index, 1)))
        ConLne(Index, 1) = 0
     End If
     
     For I = 1 To FlowShape1.Count - 1
       'For J = 1 To FlowShape1.Count - 1
         If (ConLne(I, 0) = Index) Or (ConLne(I, 1) = Index) Then
           Unload ConnectingLine(LineInvolved(I, Index))
           Unload arrDown(LineInvolved(I, Index))
           Unload arrUp(LineInvolved(I, Index))
           NoOfLines = NoOfLines - 1
           ConLne(I, 0) = 0
           ConLne(I, 1) = 0
         End If
      ' Next J
     Next I
  End If
End Sub

Private Sub FlowShape1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  LineAffectedIndex1 = 0
  LineAffectedIndex2 = 0
  FlowShape1(Index).ShapeControl
  MuseDn = False
End Sub

Private Sub Form_Load()
  CoolerBar1.ButtonDisabled(6) = True
  CoolerBar1.ButtonDisabled(7) = True
  'Create CoolTips for the two Toolbars
  ToolTipHandles(0) = CreateTip(CoolerBar1.hwnd, False, 0, -1, "Coolbar", "This is a coolbar", TTIconInfo, TTBalloon)
  ToolTipHandles(1) = CreateTip(CoolerBar2.hwnd, False, 0, -1, "Coolbar", "This is a coolbar", TTIconInfo, TTBalloon)
  PrevIndex = -2
  
  'CreateTip ToolButton(0).hwnd, False, 0, -1, "Begin", "Use to begin your main code." & vbCrLf & "DoubleClick to set variables", TTIconInfo, TTBalloon
  'CreateTip ToolButton(1).hwnd, False, 0, -1, "Stop", "Use to show the end of the" & vbCrLf & "main code or a function", TTIconInfo, TTBalloon
  'CreateTip ToolButton(2).hwnd, False, 0, -1, "Input", "Request For Data Input." & vbCrLf & "Can save both numbers, strings" & vbCrLf & "and mathematical expressions", TTIconInfo, TTBalloon
  
  'CreateTip Check1.hwnd, False, 0, -1, "Watches", "Rightclick on the listbox" & vbCrLf & "to add a new watch" & vbCrLf & "and mathematical expressions", TTIconInfo, TTBalloon
  '
  Randomize
  If Int(Rnd * 5) = 0 Then Label1.Visible = True
  
  GridSpace = 10
  On Error GoTo Skoops
  DoEvents
  WorkSpace.Width = Screen.Width / 15
  WorkSpace.Height = Screen.Height / 15
  DoEvents
  ClickRec(0) = -1
  ClickRec(1) = -1
  If GridSpace > 1 Then
    For X = 0 To WorkSpace.Width Step GridSpace
      For Y = 0 To WorkSpace.Height Step GridSpace
         WorkSpace.PSet (X, Y), &H707070 '&H909090
      Next Y
      DoEvents
    Next X
  End If
'  WorkSpace.Picture = WorkSpace.Image
  
  WorkSpace.Width = 498
  WorkSpace.Height = 358
  
  Me.Show
  Form_Resize
  'FrmTip.Show vbModal, Me
  txtCode.Interval = 400
  
  VScrollV = VScroll1.Value
  HScrollV = HScroll1.Value
  If Len(Interaction.Command) > 4 Then
     Select Case Mid(Interaction.Command, 1, 2)
        Case "/o": OpenFlowChart Mid(Interaction.Command, 4)
     End Select
   End If
Skoops:
  If Err.Number <> 0 Then
     GridSpace = 1
     WorkSpace.Width = 498
     WorkSpace.Height = 358
  End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMinimized Then
   Picture3.Move 1, 32 + 27, Me.Width / 15 - 10, Me.Height / 15 - 80 - 27
   Picture2.Move 1, 32 + 27, Me.Width / 15 - 10, Me.Height / 15 - 80 - 27
   txtCode.Move 0, 0, Picture2.Width - 4, Picture2.Height - 4
   VScroll1.Move Picture3.Width - VScroll1.Width - 4, 0, VScroll1.Width, Picture3.Height - HScroll1.Height - 4
   HScroll1.Move 0, Picture3.Height - HScroll1.Height - 4, Picture3.Width - VScroll1.Width - 4
   WorkSpace.Move 0, 0, Picture3.Width - VScroll1.Width - 4, Picture3.Height - HScroll1.Height - 4
   Label1.Move WorkSpace.Width - Label1.Width, WorkSpace.Height - Label1.Height
   Picture4.Move WorkSpace.Width - Picture4.Width, WorkSpace.Height - Picture4.Height
End If
End Sub

Function FirstPart(Stri$, Delminer$) As String
   On Error Resume Next
   FirstPart = Trim(Mid(Stri, 1, InStr(1, Stri, Delminer) - 1))
End Function

Function LastPart(Stri$, Delminer$) As String
   On Error Resume Next
   LastPart = Trim(Mid(Stri, InStr(1, Stri, Delminer) + 1))
End Function

Private Sub HScroll1_Change()
For I = 1 To FlowShape1().Count - 1
   FlowShape1(I).Left = FlowShape1(I).Left - (HScroll1.Value - HScrollV) * 3
Next
sSelection.Left = sSelection.Left - (HScroll1.Value - HScrollV) * 3
HScrollV = HScroll1.Value
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Watches() = Split(InputBox("Enter Variables to Watch" & vbCrLf & "Seperate by commas", "Watches"), ",")
End Sub

Private Sub mnuAboutItm_Click()
  frmabout.Show vbModal, Me
End Sub

Private Sub mnuExitItm_Click()
  End
End Sub

Private Sub mnuNewItm_Click()
On Error Resume Next
For I = 0 To arrUp.Count - 1
   Unload arrUp(I)
   Unload arrDown(I)
   Unload ConnectingLine(I)
Next
For I = 1 To FlowShape1.Count - 1
    Unload FlowShape1(I)
Next
For I = 0 To 150
  For J = 0 To 150
     LineInvolved(I, J) = 0
  Next
  ConLne(I, 0) = 0
  ConLne(I, 1) = 0
Next

NoOfLines = 0
NoOfObjects = 0
End Sub

Private Sub mnuOpen_Click()
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName <> "" Then OpenFlowChart CommonDialog1.FileName
End Sub


Sub OpenFlowChart(OpenFileName As String)
Dim Nol As Integer
ReDim ShitToSave(1 To 100)
ReDim LineShitToSave(1 To 100)
Dim LoadIndex&
If NotSaved Then
   If MsgBox("Flowchart not saved. Opening a flowchart" & vbCrLf & "will clear your existing one, continue anyway?", vbYesNo Or vbExclamation, "Open File") = vbNo Then Exit Sub
End If
mnuNewItm_Click

Open OpenFileName For Binary As #1
    Get #1, , ShitToSave()
Close #1

For I = 1 To UBound(ShitToSave())
    If ShitToSave(I).W = 0 And ShitToSave(I).H = 0 Then Exit For
    Load FlowShape1(FlowShape1.Count)
    FlowShape1(FlowShape1.Count - 1).Left = ShitToSave(I).X
    FlowShape1(FlowShape1.Count - 1).Top = ShitToSave(I).Y
    FlowShape1(FlowShape1.Count - 1).Width = ShitToSave(I).W
    FlowShape1(FlowShape1.Count - 1).Height = ShitToSave(I).H
    FlowShape1(FlowShape1.Count - 1).Shape = ShitToSave(I).Shape
    FlowShape1(FlowShape1.Count - 1).Title = ShitToSave(I).Title
    FlowShape1(FlowShape1.Count - 1).Caption = ShitToSave(I).Caption
    FlowShape1(FlowShape1.Count - 1).ControlData = ShitToSave(I).ControlData
    FlowShape1(FlowShape1.Count - 1).Visible = True
    FlowShape1(FlowShape1.Count - 1).ShapeControl
    
    If ShitToSave(I).Next1 <> 0 Then
        ConLne(I, 0) = ShitToSave(I).Next1
    
        L = ConnectingLine.Count
    
        Load ConnectingLine(L)
        Load arrDown(L)
        Load arrUp(L)
        LineInvolved(I, ShitToSave(I).Next1) = L
        Nol = Nol + 1
    End If
    
    If ShitToSave(I).Shape = 3 And ShitToSave(I).Next2 <> 0 Then
       ConLne(I, 1) = ShitToSave(I).Next2
       
       L = ConnectingLine.Count
       
       Load ConnectingLine(L)
       Load arrDown(L)
       Load arrUp(L)
       LineInvolved(I, ShitToSave(I).Next2) = L
       Nol = Nol + 1
    End If
Next

For I = 1 To UBound(ShitToSave())
    If ShitToSave(I).W = 0 And ShitToSave(I).H = 0 Then Exit For
    If ShitToSave(I).Next1 <> 0 Then
       ConnectingLineController (I), ConLne(I, 0), LineInvolved(I, ShitToSave(I).Next1)
       ConnectingLine(LineInvolved(I, ShitToSave(I).Next1)).Visible = True
       arrDown(LineInvolved(I, ShitToSave(I).Next1)).Visible = True
       arrUp(LineInvolved(I, ShitToSave(I).Next1)).Visible = True
    End If
    If ShitToSave(I).Shape = 3 And ShitToSave(I).Next2 <> 0 Then
       ConnectingLineController (I), ConLne(I, 1), LineInvolved(I, ShitToSave(I).Next2)
       ConnectingLine(LineInvolved(I, ShitToSave(I).Next2)).Visible = True
       arrDown(LineInvolved(I, ShitToSave(I).Next2)).Visible = True
       arrUp(LineInvolved(I, ShitToSave(I).Next2)).Visible = True
    End If
Next

NoOfObjects = I
NoOfLines = Nol
End Sub
Private Sub mnuSaveAsItm_Click()
If FlowShape1.Count = 1 Then Exit Sub
ReDim ShitToSave(1 To FlowShape1.Count - 1)

If ConnectingLine.Count - 1 >= 1 Then
ReDim LineShitToSave(1 To ConnectingLine.Count - 1)
End If
Dim I%, X As Byte, Y As Byte
For I = 1 To FlowShape1.Count - 1
    ShitToSave(I).X = FlowShape1(I).Left
    ShitToSave(I).Y = FlowShape1(I).Top
    ShitToSave(I).W = FlowShape1(I).Width
    ShitToSave(I).H = FlowShape1(I).Height
    ShitToSave(I).Shape = FlowShape1(I).Shape
    ShitToSave(I).Title = FlowShape1(I).Title
    ShitToSave(I).Caption = FlowShape1(I).Caption
    ShitToSave(I).ControlData = FlowShape1(I).ControlData
    ShitToSave(I).Next1 = ConLne(I, 0)
    ShitToSave(I).Next2 = ConLne(I, 1)
Next

I = 0

CommonDialog1.ShowSave


Open CommonDialog1.FileName For Binary As #1
    Put #1, , ShitToSave()
Close #1
SaveFileName = CommonDialog1.FileName
NotSaved = False
End Sub

Private Sub mnuSaveItm_Click()
   mnuSaveAsItm_Click
End Sub

Private Sub mnuTipItm_Click()
  FrmTip.Show vbModal, Me
End Sub

Private Sub FlowShape1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    SelectedShape = Index
    If Index <> ClickRec(0) Then
       ClickRec(1) = ClickRec(0)
       ClickRec(0) = Index
    End If
    MDX = X
    MDY = Y
    sInd = Index
    For I = 1 To NoOfLines
       If ConnLne(I).PreviewsShape = Index Then
          LineAffectedIndex(1) = I
       ElseIf ConnLne(I).NextShape = Index Then
          LineAffectedIndex(2) = I
       ElseIf ConnLne(I).ParaShape = Index Then
          LineAffectedIndex(3) = I
       End If
    Next I
    sSelection.Move FlowShape1(Index).Left, FlowShape1(Index).Top - 2, FlowShape1(Index).Width + 2, FlowShape1(Index).Height + 5
    
    ConnectingLine(SelectedLineIndex).BorderColor = &H80000008
    arrUp(SelectedLineIndex).BorderColor = &H80000008
    arrDown(SelectedLineIndex).BorderColor = &H80000008
    
    sSelection.Visible = True
    
    Text1(0).Text = FlowShape1(SelectedShape).Width
    Text1(1).Text = FlowShape1(SelectedShape).Height
    FlowShape1(Index).ZOrder
    MuseDn = True
End Sub

Private Sub FlowShape1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Static cx&, cy&, I%
If MuseDn Then
If Button = 1 And Index = sInd Then
 cx = cx + (X - MDX)
 cx = Round(cx / GridSpace) * GridSpace
 cy = cy + (Y - MDY)
 cy = Round(cy / GridSpace) * GridSpace
 If cx < 1 Then cx = 1
 If cy < 1 Then cy = 1
 If cx > WorkSpace.Width - FlowShape1(Index).Width - 4 Then cx = WorkSpace.Width - FlowShape1(Index).Width - 4
 If cy > WorkSpace.Height - FlowShape1(Index).Height - 4 Then cy = WorkSpace.Height - FlowShape1(Index).Height - 4
 FlowShape1(Index).Move cx, cy
 'ConnectingLineController Index, Index + 1, 0
 
 ConnectingLineController Index, ConLne(Index, 0), LineInvolved(Index, ConLne(Index, 0))
 If FlowShape1(Index).Shape = Lozenge And ConLne(Index, 1) <> 0 Then
    ConnectingLineController Index, ConLne(Index, 1), LineInvolved(Index, ConLne(Index, 1))
 End If
 For I = 1 To FlowShape1.Count - 1
    If (ConLne(I, 0) = Index) Or (ConLne(I, 1) = Index) Then
       ConnectingLineController I, Index, LineInvolved(I, Index)
    End If
 Next I
 
 'If LineAffectedIndex(1) <> 0 Then
 '   If ConnLne(LineAffectedIndex(1)).Exists Then ConnectingLineController ConnLne(LineAffectedIndex(1)).PreviewsShape, ConnLne(LineAffectedIndex(1)).NextShape, LineAffectedIndex(1)
 'End If
 'If LineAffectedIndex(2) <> 0 Then
 '   If ConnLne(LineAffectedIndex(2)).Exists Then ConnectingLineController ConnLne(LineAffectedIndex(2)).PreviewsShape, ConnLne(LineAffectedIndex(2)).NextShape, LineAffectedIndex(2)
 'End If
 'If LineAffectedIndex(3) <> 0 Then
 '   If ConnLne(LineAffectedIndex(3)).Exists Then ConnectingLineController ConnLne(LineAffectedIndex(3)).PreviewsShape, ConnLne(LineAffectedIndex(3)).NextShape, LineAffectedIndex(3)
 'End If
 sSelection.Move FlowShape1(Index).Left, FlowShape1(Index).Top - 2, FlowShape1(Index).Width + 2, FlowShape1(Index).Height + 5
 OriginalLeft(Index) = FlowShape1(Index).Left
 OriginalTop(Index) = FlowShape1(Index).Top
 FlowShape1(Index).ZOrder
End If
End If
End Sub
'''''||||||||||||||||||'''''''||||||||||||||||||||'''''''''''''''|||||||||||||
Private Sub VScroll1_Change()
For I = 1 To FlowShape1().Count - 1
   FlowShape1(I).Top = FlowShape1(I).Top - (VScroll1.Value - VScrollV) * 3
Next
sSelection.Top = sSelection.Top - (VScroll1.Value - VScrollV) * 3
VScrollV = VScroll1.Value
End Sub


Sub ConnectingLineController(StartIndex%, StopIndex%, ByVal LineIndex%)
On Error GoTo ExitF
    Dim Xa&, Xb&, Ya&, Yb&
    'Sub By: Alexander Popov
    If FlowShape1(StopIndex).Left > (FlowShape1(StartIndex).Left + FlowShape1(StartIndex).Width) Then
        Xa = FlowShape1(StartIndex).Left + FlowShape1(StartIndex).Width
        Ya = FlowShape1(StartIndex).Top + FlowShape1(StartIndex).Height / 2
    ElseIf FlowShape1(StopIndex).Top > (FlowShape1(StartIndex).Top + FlowShape1(StartIndex).Height) Then
        Xa = FlowShape1(StartIndex).Left + FlowShape1(StartIndex).Width / 2
        Ya = FlowShape1(StartIndex).Top + FlowShape1(StartIndex).Height
    ElseIf (FlowShape1(StopIndex).Left + FlowShape1(StopIndex).Width) < FlowShape1(StartIndex).Left Then
        Xa = FlowShape1(StartIndex).Left
        Ya = FlowShape1(StartIndex).Top + FlowShape1(StartIndex).Height / 2
    Else
        Xa = FlowShape1(StartIndex).Left + FlowShape1(StartIndex).Width / 2
        Ya = FlowShape1(StartIndex).Top
    End If
            
    If FlowShape1(StartIndex).Left > (FlowShape1(StopIndex).Left + FlowShape1(StopIndex).Width) Then
        Xb = FlowShape1(StopIndex).Left + FlowShape1(StopIndex).Width
        Yb = FlowShape1(StopIndex).Top + FlowShape1(StopIndex).Height / 2
    ElseIf FlowShape1(StartIndex).Top > (FlowShape1(StopIndex).Top + FlowShape1(StopIndex).Height) Then
        Xb = FlowShape1(StopIndex).Left + FlowShape1(StopIndex).Width / 2
        Yb = FlowShape1(StopIndex).Top + FlowShape1(StopIndex).Height
    ElseIf (FlowShape1(StartIndex).Left + FlowShape1(StartIndex).Width) < FlowShape1(StopIndex).Left Then
        Xb = FlowShape1(StopIndex).Left
        Yb = FlowShape1(StopIndex).Top + FlowShape1(StopIndex).Height / 2
    Else
        Xb = FlowShape1(StopIndex).Left + FlowShape1(StopIndex).Width / 2
        Yb = FlowShape1(StopIndex).Top
    End If
    ConnectingLine(LineIndex).X1 = Xa: ConnectingLine(LineIndex).Y1 = Ya
    ConnectingLine(LineIndex).X2 = Xb: ConnectingLine(LineIndex).Y2 = Yb
    showArrow LineIndex, Xa, Ya, Xb, Yb
ExitF:
End Sub


Sub showArrow(arINDEX, X1&, Y1&, X2&, Y2&, Optional Angle& = 30, Optional arrow_len# = 15)
   Const DR = 57.2957795130823
   Dim Opposite#, Adjacent#, Hypotenuse#, arrow_angle#
   Dim mSin#, mCos#, mSin1#, mCos1#, mSin2#, mCos2#, ArcS#, ArcC#
  
   Adjacent = (X1 - X2)
   Opposite = (Y1 - Y2)
   Hypotenuse = Sqr(Adjacent * Adjacent + Opposite * Opposite) 'Pythagoras's theorim
    
   If Hypotenuse <> 0 Then
       mSin = Opposite / Hypotenuse
       mCos = Adjacent / Hypotenuse
   Else
       mSin = 0
       mCos = 0
   End If
    
   arrow_angle = Angle / DR
   ArcS = Arcsin(mSin)
   ArcC = Arccos(mCos)
   If (mSin >= 0) And (mCos >= 0) Then
        mSin1 = Sin(ArcS - arrow_angle)
        mSin2 = Sin(ArcS + arrow_angle)
        mCos1 = Cos(ArcC - arrow_angle)
        mCos2 = Cos(ArcC + arrow_angle)
   ElseIf (mSin <= 0) And (mCos >= 0) Then
        mSin1 = Sin(ArcS - arrow_angle)
        mSin2 = Sin(ArcS + arrow_angle)
        mCos1 = Cos(ArcC + arrow_angle)
        mCos2 = Cos(ArcC - arrow_angle)
   ElseIf (mSin >= 0) And (mCos <= 0) Then
        mSin1 = Sin(ArcS + arrow_angle)
        mSin2 = Sin(ArcS - arrow_angle)
        mCos1 = Cos(ArcC - arrow_angle)
        mCos2 = Cos(ArcC + arrow_angle)
   ElseIf (mSin <= 0) And (mCos <= 0) Then
        mSin1 = Sin(ArcS + arrow_angle)
        mSin2 = Sin(ArcS - arrow_angle)
        mCos1 = Cos(ArcC + arrow_angle)
        mCos2 = Cos(ArcC - arrow_angle)
   End If
      
   arrUp(arINDEX).X1 = X2
   arrUp(arINDEX).Y1 = Y2
   arrDown(arINDEX).X1 = X2
   arrDown(arINDEX).Y1 = Y2

   arrUp(arINDEX).X2 = mCos1 * arrow_len + X2
   arrUp(arINDEX).Y2 = mSin1 * arrow_len + Y2
   arrDown(arINDEX).X2 = mCos2 * arrow_len + X2
   arrDown(arINDEX).Y2 = mSin2 * arrow_len + Y2
End Sub

Function DataString(ByVal ObjectIndex As Byte) As String
    DataString = Chr(ObjectIndex) & Chr(FlowShape1(ObjectIndex).Shape)
End Function

Private Sub WorkSpace_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim K%, L%
   If KeyCode = vbKeyDelete Then
      MsgBox "Can Not Delete Anything", vbExclamation
      Exit Sub
      For K = 1 To FlowShape1.Count - 1
         For L = 1 To FlowShape1.Count - 1
           If LineInvolved(K, L) = SelectedConnLine Then
             ConLne(K, 0) = 0
             ConLne(K, 1) = 0
            
             LineInvolved(K, L) = 0
             LineInvolved(L, K) = 0
           End If
         Next L
      Next K
      If SelectedConnLine <> 0 Then
         Unload ConnectingLine(SelectedConnLine)
         Unload arrDown(SelectedConnLine)
         Unload arrUp(SelectedConnLine)
      End If
   End If
End Sub

Private Sub WorkSpace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LineFound As Boolean
On Error Resume Next
  For I = 1 To NoOfLines 'ConnectingLine.Count - 1
      If LineMouseEvent(ConnectingLine(I), X, Y, 8) Then
         ConnectingLine(I).BorderColor = &HBB0000
         arrUp(I).BorderColor = &HBB0000
         arrDown(I).BorderColor = &HBB0000
         sSelection.Visible = False
         SelectedLineIndex = I
         LineFound = True
         SelectedConnLine = I
      Else
         ConnectingLine(I).BorderColor = &H80000008
         arrUp(I).BorderColor = &H80000008
         arrDown(I).BorderColor = &H80000008
      End If

  Next I
  If Not LineFound Then
     'show selector
     oSelX = X
     oSely = Y
     oSelection.Move 0, 0, 0, 0 'hide it until mouse is moved
     oSelection.Visible = True
  End If
End Sub

Private Sub WorkSpace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim L&, T&
'MsgBox ""
    If Button <> 0 Then
      'Nested if should be always used in events triggered this often
      If oSelection.Visible Then
         If X > oSelX Then
            L = oSelX
         Else
            L = X
         End If
         If Y > oSely Then
            T = oSely
         Else
            T = Y
         End If
         oSelection.Move L, T, Abs(X - oSelX), Abs(Y - oSely)
      End If
    End If
End Sub

Private Sub WorkSpace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   oSelection.Visible = False
End Sub
