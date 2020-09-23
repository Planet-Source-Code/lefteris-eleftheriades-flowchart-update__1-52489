Attribute VB_Name = "Module1"
'What's So Great Here Is That The One Extra Ocx File Needed For The Project is Self-Extracted
'Best of all, the packed application size is smaller than 200Kb!
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public SelectedShape As Integer
Public CurrentStep As Long
Public Running As Boolean
Public RunningFlow As Boolean


Public Function StringFromBuffer(Buffer As String) As String
    Dim nPos As Long

    nPos = InStr(Buffer, vbNullChar)
    If nPos > 0 Then
        StringFromBuffer = Left(Buffer, nPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function

Public Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(255)
    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, 255) Then
        GetWindowsSysDir = StringFromBuffer(strBuf)
    End If
End Function


Sub Main()
On Error Resume Next
If Dir(GetWindowsSysDir & "\Msscript.ocx") = "" Then
  If MsgBox("MsScript.ocx Not found." & vbCrLf & "Install?", vbExclamation Or vbYesNo Or vbDefaultButton1, "MsScript") = vbYes Then
     'Extract and register Microsoft Scripting Control
     ExtractMsScriptOcx GetWindowsSysDir
     DoEvents
     Shell GetWindowsSysDir & "\Regsvr32 " & GetWindowsSysDir & "\Msscript.ocx"
     DoEvents
  End If
End If
DoEvents
FlowChart.Show
End Sub

Sub ExtractMsScriptOcx(ByVal WinSysDirs$)
  Dim MsScriptd() As Byte
  
  If Dir(GetWindowsSysDir & "\Msscript.ocx") = "" Then
     'Load the ocx file
     MsScriptd = LoadResData(101, "ocx")
     Open WinSysDirs & "\Msscript.ocx" For Binary As #1
        Put 1, , MsScriptd()
     Close #1
    Erase MsScriptd
  End If
  
  DoEvents
  
  If Dir(GetWindowsSysDir & "\Msscript.hlp") = "" Then
     'Load it's help files
     MsScriptd = LoadResData(102, "ocx")
     Open WinSysDirs & "\Msscript.hlp" For Binary As #1
        Put 1, , MsScriptd()
     Close #1
     Erase MsScriptd
  End If
  
  DoEvents
  
  If Dir(GetWindowsSysDir & "\Msscript.cnt") = "" Then
     'Load the helpfile's table of contents
     MsScriptd = LoadResData(103, "ocx")
     Open WinSysDirs & "\Msscript.cnt" For Binary As #1
        Put 1, , MsScriptd()
     Close #1
     Erase MsScriptd
  End If
  
  DoEvents
End Sub
