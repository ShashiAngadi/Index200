VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please identify yourself"
   ClientHeight    =   2310
   ClientLeft      =   2220
   ClientTop       =   1860
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1770
      Width           =   1275
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1770
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4695
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   300
         Left            =   4020
         TabIndex        =   9
         Top             =   1080
         Width           =   405
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   2250
         TabIndex        =   8
         Top             =   1080
         Width           =   1665
      End
      Begin VB.TextBox txtUserPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   690
         Width           =   2175
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   2250
         TabIndex        =   2
         Top             =   300
         Width           =   2175
      End
      Begin VB.Image img 
         Height          =   555
         Left            =   240
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblUserDate 
         Caption         =   "Date:"
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label lblUserPassword 
         Caption         =   "User password:"
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblUserName 
         Caption         =   "User name: "
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LoginClicked(UserName As String, Userpassword As String, LoginDate As String, ByRef UnloadDialog As Boolean)
Public Event CancelClicked()

Private Sub cmdCancel_Click()
RaiseEvent CancelClicked
Unload Me

End Sub

Private Sub cmdLogin_Click()
Dim UnloadDialog As Boolean
Dim Rst As ADODB.Recordset

'If the login date is earlier then Balancesheet date, the application will exit
gDbTrans.SQLStmt = "Select ValueData FROM Install WHERE KeyData = 'BalanceSheetDate' "
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    If CDate(txtDate.Text) > CDate(Rst(0)) Then
        gDbTrans.SQLStmt = "Select ValueData FROM Install WHERE KeyData = 'BalanceSheetEndDate' "
        If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
            If CDate(txtDate.Text) > CDate(Rst(0)) Then
                Unload Me
            End If
        End If
    End If
End If

RaiseEvent LoginClicked(Trim$(txtUserName.Text), Trim$(txtUserPassword.Text), "", UnloadDialog)
Dim DBPath As String
If UnloadDialog Then
      If DateValidate(Me.txtDate.Text, "/", True) Then
            gStrDate = FormatDate(txtDate.Text)
      Else
            gStrDate = Format(Now, "MM/DD/YYYY")
      End If
      DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "server")
      If DBPath = "" Then
      'Give the local path of the MDB FILE
        DBPath = App.Path
      'Call MakeBackUp(DBPath & "\Index 2000.mdb")
      Else
        DBPath = "\\" & DBPath & "\Index 2000"
      End If
    Unload Me
End If
'gDBTrans.CompactDataBase
End Sub

Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + Frame1.Left + cmdDate.Left - 1400
        .Top = Me.Top + Frame1.Top + cmdDate.Top
    If DateValidate(txtDate.Text, "/", True) Then
        .SelDate = txtDate.Text
    Else
        .SelDate = FormatDate(gStrDate)
    End If
    .Show vbModal
    If .SelDate <> "" Then txtDate.Text = .SelDate
End With

End Sub





Private Sub Command1_Click()
    frmDayBegin.Show 1
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    'set icon for the form caption
    Me.Icon = LoadResPicture(161, vbResIcon)
 
    img.Picture = LoadResPicture(108, vbResIcon)
    Me.Caption = gAppName & " - Please identify yourself"
    Me.Icon = LoadResPicture(108, vbResIcon)
    'Load today's date
    Me.txtDate = FormatDate(gStrDate)
    'set kannada fonts
    Call SetKannadaCaption
End Sub




Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

lblUserName.Caption = LoadResString(gLangOffSet + 151) & " " & LoadResString(gLangOffSet + 35)
lblUserPassword.Caption = LoadResString(gLangOffSet + 151) & " " & LoadResString(gLangOffSet + 153)
lblUserDate.Caption = LoadResString(gLangOffSet + 37)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
cmdLogin.Caption = LoadResString(gLangOffSet + 151)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
End Sub

Private Sub img_DblClick()
txtUserName = "ADMIN"
txtUserPassword = "ADMIN"
Call cmdLogin_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub


