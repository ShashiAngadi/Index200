VERSION 5.00
Begin VB.Form frmPassWord 
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   2850
   ClientTop       =   3435
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   465
      Left            =   720
      TabIndex        =   8
      Top             =   2430
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   2370
      TabIndex        =   7
      Top             =   2430
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      Caption         =   "Password"
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   3825
      Begin VB.TextBox txtPaswd2 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   750
         Width           =   2145
      End
      Begin VB.TextBox txtPaswd1 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   330
         Width           =   2145
      End
      Begin VB.Label Label2 
         Caption         =   "Confirm password:"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Change password:"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   1470
      Width           =   3825
      Begin VB.CheckBox Check1 
         Caption         =   "Automatically hide all entries made for the day"
         Height          =   465
         Left            =   240
         TabIndex        =   1
         Top             =   270
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Attempts As Integer
Const MaxAttempts = 3
Private Function ChangePassword(NewPwd As String) As Boolean
On Error GoTo ERRLINE
Dim UserName As String
Dim Password As String


UserName = Trim$(frmLogin.txtUserName.Text)
Password = Trim$(frmLogin.txtUserPassword.Text)

gDBTrans.SQLStmt = "Select * from UserTab where " & _
                       " LoginName = '" & UserName & "' and " & _
                       " Password = '" & Password & "'"

Dim RetVal As Integer

RetVal = gDBTrans.SQLFetch
If RetVal <= 0 Or RetVal > 1 Then
      GoTo ERRLINE
End If

'Now Change Rename With New Password

gDBTrans.SQLStmt = "UPDATE UserTab SET password = '" & NewPwd & "'" & " where " & _
                       " LoginName = '" & UserName & "' and " & _
                       " Password = '" & Password & "'"
      
gDBTrans.BeginTrans
      If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            GoTo ERRLINE
      End If

gDBTrans.CommitTrans
ChangePassword = True
Exit Function

ERRLINE:
      ChangePassword = False
      
      
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim NewPwd As String
If Trim$(Me.txtPaswd1.Text) <> Trim$(Me.txtPaswd2.Text) Then
      MsgBox "Passwords Do Not Match Unable To Change Password ", vbInformation, wis_MESSAGE_TITLE
      Attempts = Attempts + 1
      If Attempts > MaxAttempts Then Call cmdCancel_Click
      ActivateTextBox txtPaswd1
      Exit Sub
End If
If Trim$(txtPaswd1.Text) = Trim$(txtPaswd2.Text) Then
NewPwd = Trim$(txtPaswd1.Text)
If Not ChangePassword(NewPwd) Then
      MsgBox "Unable To Change The Password ", vbInformation, wis_MESSAGE_TITLE
End If
End If
End Sub


Private Sub Form_Load()
'centre the form
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
txtPaswd1.Text = ""
txtPaswd2.Text = ""
txtPaswd1.PasswordChar = "*"
txtPaswd2.PasswordChar = "*"
End Sub


