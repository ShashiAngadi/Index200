VERSION 5.00
Begin VB.Form frmCancel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceessing ......."
   ClientHeight    =   1560
   ClientLeft      =   1755
   ClientTop       =   1845
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4320
   Begin VB.PictureBox PicStatus 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   30
      ScaleHeight     =   225
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "End"
      Height          =   400
      Left            =   2310
      TabIndex        =   2
      Top             =   1110
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   400
      Left            =   3300
      TabIndex        =   0
      Top             =   1110
      Width           =   945
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message"
      Height          =   705
      Left            =   60
      TabIndex        =   1
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   4185
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CancelClicked()
Public Event EndClicked()



Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
cmdCancel.Caption = GetResourceString(2) 'Cancel
End Sub

Private Sub cmdCancel_Click()
Dim MousePointer As Integer
MousePointer = Screen.MousePointer
Screen.MousePointer = vbDefault
'If MsgBox("Are you sure U want to cancel this process", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbYes Then
If MsgBox(GetResourceString(613), vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
    RaiseEvent CancelClicked
    gCancel = 1
    Unload Me
End If
Screen.MousePointer = MousePointer
End Sub



Private Sub cmdEnd_Click()
gCancel = 1
Dim MousePointer As Integer
MousePointer = Screen.MousePointer
Screen.MousePointer = vbDefault
'If MsgBox("Are you sure U want to cancel this process", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbYes Then
If MsgBox(GetResourceString(613), vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
    RaiseEvent EndClicked
    gCancel = 2
End If
Screen.MousePointer = MousePointer
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
Call SetKannadaCaption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   ""(Me.hwnd, False)
End Sub


Private Sub lblMessage_Change()
On Error Resume Next
'Me.prg.Visible = True
'Call UpdateStatus(PicStatus, prg.Value / prg.Max)
Me.Refresh
Err.Clear
End Sub

