VERSION 5.00
Begin VB.Form frmRptDt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIS Report date..."
   ClientHeight    =   1590
   ClientLeft      =   6570
   ClientTop       =   3300
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3090
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1080
      TabIndex        =   2
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2100
      TabIndex        =   3
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txtEndDate 
      Height          =   390
      Left            =   1530
      TabIndex        =   1
      Top             =   570
      Width           =   1455
   End
   Begin VB.TextBox txtStDate 
      Height          =   390
      Left            =   1530
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "End Date :"
      Height          =   390
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblStDate 
      AutoSize        =   -1  'True
      Caption         =   "&Start Date :"
      Height          =   390
      Left            =   60
      TabIndex        =   4
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmRptDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event OKClick(StDate As String, EndDate As String)
Public Event CancelClick()
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
lblStDate.Caption = GetResourceString(109)
lblEndDate.Caption = GetResourceString(110)

End Sub



Private Sub cmdCancel_Click()
RaiseEvent CancelClick
Unload Me


End Sub


Private Sub cmdOk_Click()
Dim StartDateUS As String
Dim EndDateUS As String


'Check For Validate of Dates
    
If Not DateValidate(txtEndDate, "/", True) Then Exit Sub
    
EndDateUS = GetSysFormatDate(txtEndDate.Text)
    
If txtStDate.Enabled Then
    If Not TextBoxDateValidate(txtStDate, "/", True, True) Then Exit Sub
    If Not TextBoxDateValidate(txtEndDate, "/", True, True) Then Exit Sub
    StartDateUS = GetSysFormatDate(txtStDate.Text)
    If DateDiff("d", CDate(StartDateUS), CDate(EndDateUS)) < 0 Then
        MsgBox "Start date should be earlier than the end date ", _
            vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If


Me.Hide

Screen.MousePointer = vbHourglass

RaiseEvent OKClick(txtStDate.Text, txtEndDate.Text)

Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
'Center the form
CenterMe Me

Call SetKannadaCaption

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

End Sub

Private Sub txtEndDate_GotFocus()
txtEndDate.SelStart = 0
txtEndDate.SelLength = Len(txtEndDate)
End Sub


Private Sub txtStDate_GotFocus()
txtStDate.SelStart = 0
txtStDate.SelLength = Len(txtStDate)

End Sub


