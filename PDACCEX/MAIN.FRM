VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   2055
   ClientTop       =   2985
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   1050
      TabIndex        =   1
      Top             =   1650
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch Express Deposits"
      Height          =   735
      Left            =   1050
      TabIndex        =   0
      Top             =   330
      Width           =   4275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_User As clsUsers
Attribute m_User.VB_VarHelpID = -1
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()

gUser.MaxRetries = 2
gUser.ShowLoginDialog

If Not gUser.LoginStatus Then
   GoTo Err_Line
End If
gUserID = gUser.UserID
'Load the Express Form
frmPDAccEx.Show vbModal

Exit Sub
Err_Line:
    MsgBox "The system could not log you on !", vbExclamation, gAppName & " - Error"
    End
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
gDBTrans.CloseDB

End Sub



