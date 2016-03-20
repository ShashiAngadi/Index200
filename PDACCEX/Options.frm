VERSION 5.00
Begin VB.Form frmPassWord 
   Caption         =   "Form3"
   ClientHeight    =   2085
   ClientLeft      =   2805
   ClientTop       =   1980
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2760
      TabIndex        =   6
      Top             =   1590
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   1590
      TabIndex        =   5
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Frame fraPassword 
      Caption         =   "Password"
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3825
      Begin VB.TextBox txtPaswd2 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   750
         Width           =   2145
      End
      Begin VB.TextBox txtPaswd1 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1590
         TabIndex        =   1
         Top             =   330
         Width           =   2145
      End
      Begin VB.Label lblConfirmPassword 
         Caption         =   "Confirm password:"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label lblChangePassword 
         Caption         =   "Change password:"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function ChangePassword(NewPasswd As String) As Boolean
On Error GoTo ErrLine
Dim UserName As String
Dim Password As String
 

UserName = gCurrUser.UserName
Password = gCurrUser.Userpassword
'Password = Trim$(frmLogin.txtUserPassword.Text)
    gDBTrans.SQLStmt = "Select * from UserTab where " & _
                        " LoginName = '" & UserName & "' and " & _
                        " Password = '" & Password & "'"
Dim Retval As Integer
Dim Rst As Recordset
Retval = gDBTrans.Fetch(Rst, adOpenForwardOnly)
If Retval <= 0 Or Retval > 1 Then GoTo ErrLine

'Update Table With NewPAssword
gDBTrans.SQLStmt = "Update UserTab SET Password = '" & NewPasswd & "' " & " where " & _
                        " LoginName = '" & UserName & "' and " & _
                        " Password = '" & Password & "'"
gDBTrans.BeginTrans
If Not gDBTrans.SQLExecute Then
   gDBTrans.RollBack
   GoTo ErrLine
End If

gDBTrans.CommitTrans
If Not gCurrUser.Login(UserName, NewPasswd, gStrDate) Then GoTo ErrLine
ChangePassword = True
Exit Function

ErrLine:
ChangePassword = False
End Function


Private Sub SetKannadaCaption()
Dim ctrl As Control
On Error Resume Next
For Each ctrl In Me
    ctrl.Font.Name = gFontName
    If Not TypeOf ctrl Is ComboBox Then
        ctrl.Font.Size = gFontSize
    End If
Next
Me.fraPassword.Caption = LoadResString(gLangOffSet + 153)
Me.lblChangePassword.Caption = LoadResString(gLangOffSet + 497)
Me.lblConfirmPassword.Caption = LoadResString(gLangOffSet + 154)
'Me.chkHide.Caption = LoadResString(gLangOffSet + 498)
End Sub



Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()

If Trim$(Me.txtPaswd1.Text) <> Trim$(Me.txtPaswd1.Text) Then
   ''MsgBox "Passwords Do Not Match " & vbCrLf & "Please try Again ", vbInformation, wis_MESSAGE_TITLE
   MsgBox LoadResString(gLangOffSet + 615) & vbCrLf & "Please try Again ", vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtPaswd1
   Exit Sub
End If

If Trim$(Me.txtPaswd1.Text) = Trim$(Me.txtPaswd1.Text) Then
      Dim NewPasswd As String
      NewPasswd = Trim$(txtPaswd1.Text)
      If Not ChangePassword(NewPasswd) Then
         ''MsgBox "Unable To Change The password ", vbInformation, wis_MESSAGE_TITLE
         MsgBox LoadResString(gLangOffSet + 798), vbInformation, wis_MESSAGE_TITLE
      Else
         'MsgBox "Password Changed Successfully", vbInformation, wis_MESSAGE_TITLE
         MsgBox LoadResString(gLangOffSet + 799), vbInformation, wis_MESSAGE_TITLE
      End If
End If
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    Call SetKannadaCaption
    Me.txtPaswd1.Text = ""
    Me.txtPaswd2.Text = ""
    Me.txtPaswd2.PasswordChar = "*"
    Call SetKannadaCaption
    
End Sub


