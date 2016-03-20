Attribute VB_Name = "basMsgBox"
Option Explicit
Public gMessageBoxResult As VbMsgBoxResult

'
Public Function InputBox(Prompt As String, Optional Title, Optional Default, Optional XPos, Optional YPos, Optional HelpFile, Optional Context) As String
    Load frmInPut
    frmInPut.lblPrompt.Caption = Prompt
    If Not IsMissing(Default) Then
       frmInPut.txtResult.Text = CStr(Default)
    End If
    frmInPut.Caption = Title
    If Not IsMissing(XPos) Then
       If XPos = "" Then
             XPos = Screen.Height - frmInPut.Height / 2
       ElseIf IsNumeric(XPos) Then
             XPos = Val(XPos)
       Else  'If it's not a value then
             Err.Raise 13, "InputBox", "Type Mismatch"
       End If
    Else
       XPos = Screen.Height / 2 - frmInPut.Height / 2
    End If
    If Not IsMissing(YPos) Then
          If YPos = "" Then
             YPos = Screen.Width - frmInPut.Width / 2
          ElseIf IsNumeric(YPos) Then
             YPos = Val(XPos)
          Else
             Err.Raise 13, "InputBox", "Type Mismatch"
          End If
    Else
          YPos = Screen.Width / 2 - frmInPut.Width / 2
    End If
    frmInPut.Move CSng(YPos), CSng(XPos)
    frmInPut.Show vbModal
    InputBox = frmInPut.txtResult.Text
    Set frmInPut = Nothing
       
End Function

Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional HelpFile As String, Optional Context As String) As VbMsgBoxResult
    Dim Ctrl As Control
    Dim MSgType As Long
    Dim Btns As Byte, Icon As Byte, Default As Byte, allign As Byte
    Load frmMsgBox
    MSgType = Buttons
    'Call GetMessageProps(Buttons, Btns, Icon, Default, allign)
    On Error Resume Next
    Call frmMsgBox.ShowMessage(Prompt, MSgType, Title)
    frmMsgBox.lblPrompt.Caption = Prompt
    frmMsgBox.Show vbModal
    If gMessageBoxResult = -1 Then
       ' MsgBox "No result"
    Else
        MsgBox = gMessageBoxResult
    End If
    Unload frmMsgBox

End Function

