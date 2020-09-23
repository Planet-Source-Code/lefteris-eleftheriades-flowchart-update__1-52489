Attribute VB_Name = "AdvTip"
'Modification of class module code
'Original code: PictureWindow Software
Option Explicit


Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'private data

Public Function CreateTip(ParentControl_hWnd&, Centered As Boolean, ForeColor&, BackColor&, mTitle$, TipText$, mIcon As ttIconType, Style As ttStyleEnum) As Long
'Returns the Tip Handle (store it in a long variable inorder to modify the tooltip later in your code
    Dim lpRect As RECT
    Dim lWinStyle As Long
    Dim ti As TOOLINFO
    Dim lHwnd As Long
    'If lHwnd <> 0 Then
    '    DestroyWindow lHwnd
    'End If
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    ''create baloon style if desired
    If Style = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
    ''the parent control has to have been set first
    If Not ParentControl_hWnd = 0 Then
        lHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, _
                    CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
                    ParentControl_hWnd, 0&, App.hInstance, 0&)
                    
        ''make our tooltip window a topmost window
        SetWindowPos lHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
                    
        ''get the rect of the parent control
        GetClientRect ParentControl_hWnd, lpRect
        
        ''now set our tooltip info structure
        With ti
            ''if we want it centered, then set that flag
            If Centered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If
            
            ''set the hwnd prop to our parent control's hwnd
            .lHwnd = ParentControl_hWnd
            .lId = 0
            .hInstance = App.hInstance
            .lpRect = lpRect
            .lpStr = TipText
        End With
        
        ''add the tooltip structure
        SendMessage lHwnd, TTM_ADDTOOLA, 0&, ti
        
        ''if we want a title or we want an icon
        If mTitle <> vbNullString Or mIcon <> TTNoIcon Then
            SendMessage lHwnd, TTM_SETTITLE, CLng(mIcon), ByVal mTitle
        End If
        
        If ForeColor <> -1 Then
            SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, ForeColor, 0&
        End If
        
        If BackColor <> -1 Then
            SendMessage lHwnd, TTM_SETTIPBKCOLOR, BackColor, 0&
        End If
        
    End If
    CreateTip = lHwnd
End Function

Sub TipIconTitle(TipHandle&, ByVal mIcon As ttIconType, Title$)
    If TipHandle <> 0 And Title <> "" And mIcon <> TTNoIcon Then
        SendMessage TipHandle, TTM_SETTITLE, CLng(mIcon), ByVal Title
    End If
End Sub

Sub TipForeColor(TipHandle&, ByVal ForeColor As Long)
    If TipHandle <> 0 Then
        SendMessage TipHandle, TTM_SETTIPTEXTCOLOR, ForeColor, 0&
    End If
End Sub

Sub TipBackColor(TipHandle&, ByVal BackColor As Long)
    If TipHandle <> 0 Then
        SendMessage TipHandle, TTM_SETTIPBKCOLOR, BackColor, 0&
    End If
End Sub

Sub TipText(ByVal mTipText As String, TipCentered As Boolean, ParentControl_hWnd&, TipHandle&)
    Dim ti As TOOLINFO
    Dim lpRect As RECT
    GetClientRect ParentControl_hWnd, lpRect
    With ti
        ''if we want it centered, then set that flag
        If TipCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
        Else
            .lFlags = TTF_SUBCLASS
        End If
        .lHwnd = ParentControl_hWnd
        .lId = 0
        .hInstance = App.hInstance
        .lpRect = lpRect
        .lpStr = mTipText
    End With
    If TipHandle& <> 0 Then
        SendMessage TipHandle&, TTM_UPDATETIPTEXTA, 0&, ti
    End If
End Sub
