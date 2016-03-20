VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   4125
   ClientTop       =   3450
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Ignore"
      Height          =   375
      Index           =   3
      Left            =   1860
      TabIndex        =   3
      Top             =   660
      Width           =   775
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Retry"
      Height          =   375
      Index           =   2
      Left            =   990
      TabIndex        =   2
      Top             =   660
      Width           =   775
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ok"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   775
   End
   Begin VB.Image img 
      Height          =   435
      Left            =   60
      Stretch         =   -1  'True
      Top             =   90
      Width           =   435
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   215
      Left            =   240
      TabIndex        =   0
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   570
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event result()
'Private Btn As Byte
Private Default As Byte
Private Button As Byte

Public Event ButtonClicked(BtnCaption As String)

Public Sub ShowMessage(Prompt As String, MSgType As Long, Caption As String)
Dim Btn As Byte
Dim Icon As Byte
Dim allign As Byte
Dim I As Integer
Dim Kantext As String

Const LM = 200
Const TM = 200
Const RM = 200
Const BM = 200
Const Gap = 75
Const CHAR_PER_LINE = 80

Alligns:
    If MSgType >= 4096 Then   'There is an allignment required
        If MSgType >= 1048576 Then
            allign = 6
            MSgType = MSgType - 1048576
        ElseIf MSgType >= 524288 Then
            allign = 5
            MSgType = MSgType - 524288
        ElseIf MSgType >= 65536 Then
            allign = 4
            MSgType = MSgType - 65536
        ElseIf MSgType >= 16384 Then
            allign = 3
            MSgType = MSgType - 16384
        ElseIf MSgType >= 4096 Then
            allign = 2
            MSgType = MSgType - 4096
        ElseIf MSgType >= 4096 Then
            allign = 1
            MSgType = MSgType - 4096
        End If
        GoTo Defaults:
    End If

Defaults:
    Default = 1
    If MSgType >= 256 Then
        If MSgType >= 768 Then
            Default = 4
            MSgType = MSgType - 768
        ElseIf MSgType >= 512 Then
            Default = 3
            MSgType = MSgType - 512
        ElseIf MSgType >= 256 Then
            Default = 2
            MSgType = MSgType - 256
        End If
        GoTo Icons:
    End If

Icons:
    Icon = 0
    If MSgType >= 16 Then
        If MSgType >= 64 Then
            Icon = 4
            MSgType = MSgType - 64
        ElseIf MSgType >= 48 Then
            Icon = 3
            MSgType = MSgType - 48
        ElseIf MSgType >= 32 Then
            Icon = 2
            MSgType = MSgType - 32
        ElseIf MSgType >= 16 Then
            Icon = 1
            MSgType = MSgType - 16
        End If
        GoTo Buttons:
    End If

Buttons:
    Btn = 1
    Button = 1
  
    cmd(1).Caption = GetResourceString(1)
    If MSgType >= 1 Then
        If MSgType >= 5 Then
            Btn = 2
            MSgType = MSgType - 5
            cmd(1).Caption = GetResourceString(59)
            cmd(2).Caption = GetResourceString(2) '"Cancel"
            Button = 5
        ElseIf MSgType >= 4 Then
            Btn = 2
            MSgType = MSgType - 4
            cmd(1).Caption = GetResourceString(22) '"Yes"
            cmd(2).Caption = GetResourceString(50) '"No"
            Button = 4
        ElseIf MSgType >= 3 Then
            Btn = 3
            cmd(1).Caption = GetResourceString(22) '"Yes"
            cmd(2).Caption = GetResourceString(50) ' "no"
            cmd(3).Caption = GetResourceString(2) '"Cancel"
            MSgType = MSgType - 3
            Button = 3
        ElseIf MSgType >= 2 Then
            Btn = 3
            cmd(1).Caption = GetResourceString(34): cmd(2).Caption = GetResourceString(59): cmd(3).Caption = GetResourceString(97)
            MSgType = MSgType - 2
            Button = 2
        ElseIf MSgType >= 1 Then
            Btn = 2
            cmd(1).Caption = GetResourceString(1): cmd(2).Caption = GetResourceString(2)   'Cancel
            MSgType = MSgType - 1
            Button = 1
        ElseIf MSgType >= 0 Then
            Btn = 1
            cmd(1).Caption = GetResourceString(1)  ' OK
            Button = 1
        End If
    End If

        
'Set the Icon
If Icon > 0 Then
    'img.Picture = LoadPicture(App.Path & "\msgbox0" & Icon & ".ico")
    img.Picture = LoadResPicture(150 + Icon, vbResIcon)
End If

'Set the caption
    Dim Buf As String
    Dim NextChar As String
    Dim TmpStr As String
    lblPrompt.Caption = ""
    
    Buf = Prompt
    TmpStr = Prompt

    'PromptLen = Len(Buf)
    While Len(Buf) > CHAR_PER_LINE And TmpStr <> ""
        TmpStr = Left(Buf, CHAR_PER_LINE)
        Buf = Right(Buf, Len(Buf) - CHAR_PER_LINE)
        NextChar = Left(Buf, 1)
        I = 0
        While (NextChar <> " " And I < 2 And Buf <> "")
            TmpStr = TmpStr & NextChar
            Buf = Right(Buf, Len(Buf) - 1)
            I = I + 1
        Wend
        'MsgBox lblPrompt.Width
        If lblPrompt.Width < TextWidth(TmpStr) Then
            lblPrompt.Width = TextWidth(TmpStr)
        End If
        If lblPrompt.Caption <> "" Then
            lblPrompt.Caption = lblPrompt.Caption & vbCrLf
        End If
        lblPrompt.Caption = lblPrompt.Caption & TmpStr
    Wend
    
    
    lblPrompt.Width = IIf(lblPrompt.Width < TextWidth(Buf), TextWidth(Buf), lblPrompt.Width)
    If lblPrompt.Caption <> "" Then
        lblPrompt.Caption = lblPrompt.Caption & vbCrLf
    End If
    lblPrompt.Caption = lblPrompt.Caption & Buf


'Arrange the buttons
'Calculate the total space req
Dim TotWid As Integer
    For I = 1 To Btn
        TotWid = TotWid + cmd(I).Width
        TotWid = TotWid + Gap
    Next I
    
    TotWid = TotWid - Gap
    
'Arrange all the controls on the form
    If Icon > 0 Then
        lblPrompt.Left = img.Left + img.Width + RM
    Else
        lblPrompt.Left = RM
    End If
    
    lblPrompt.Top = TM
    For I = 1 To 3
        cmd(I).Top = lblPrompt.Top + lblPrompt.Height + TM
    Next I
    Me.Width = RM + LM + IIf(lblPrompt.Width > TotWid, lblPrompt.Width, TotWid) + IIf(Icon > 0, img.Width, 0)
    Me.Height = RM + lblPrompt.Top + lblPrompt.Height + TM + cmd(1).Height + TM + BM
    
Dim LPos As Integer
    LPos = (Me.Width - TotWid) \ 2
    For I = 1 To Btn
        cmd(I).Left = LPos
        cmd(I).Visible = True
        LPos = LPos + cmd(I).Width + Gap
    Next I
    For I = I To 3
        cmd(I).Visible = False
    Next I

Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Me.Caption = IIf(Caption = "", App.ProductName, Caption)
        
End Sub

Private Sub cmd_Click(Index As Integer)
gMessageBoxResult = -1
If Button = 1 Then
    If Index = 1 Then
        gMessageBoxResult = vbOK
    Else
        gMessageBoxResult = vbCancel
    End If
    'gMessageBoxResult = vbOK
ElseIf Button = 2 Then
    If Index = 1 Then
        gMessageBoxResult = vbAbort
    ElseIf Index = 2 Then
        gMessageBoxResult = vbRetry
    Else
        gMessageBoxResult = vbIgnore
    End If
ElseIf Button = 3 Then
    If Index = 1 Then
        gMessageBoxResult = vbYes
    ElseIf Index = 2 Then
        gMessageBoxResult = vbNo
    ElseIf Index = 3 Then
        gMessageBoxResult = vbCancel
    End If
ElseIf Button = 4 Then
    If Index = 1 Then
        gMessageBoxResult = vbYes
    ElseIf Index = 2 Then
        gMessageBoxResult = vbNo
    ElseIf Index = 3 Then
        gMessageBoxResult = vbCancel
    End If
ElseIf Button = 5 Then
    If Index = 1 Then
        gMessageBoxResult = vbRetry
    ElseIf Index = 2 Then
        gMessageBoxResult = vbCancel
    End If
ElseIf Button = 6 Then
    If Index = 1 Then
        gMessageBoxResult = vbRetry
    ElseIf Index = 2 Then
        gMessageBoxResult = vbCancel
    End If
End If
' Here Once again the Assign the returned value

Unload Me
'End
End Sub


Private Sub Form_Activate()
If Default < 1 Or Default > cmd.count Then
    cmd(1).SetFocus
    Exit Sub
End If
cmd(Default).SetFocus
End Sub



Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
  
Call SetKannadaCaption
End Sub

Private Sub SetKannadaCaption()
  Call SetFontToControls(Me)

End Sub



