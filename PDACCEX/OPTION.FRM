VERSION 5.00
Begin VB.Form frmOption 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   2865
   ClientTop       =   5310
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   4170
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2670
      TabIndex        =   3
      Top             =   1770
      Width           =   1395
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options.."
      Height          =   1665
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   3975
      Begin VB.OptionButton optChangePassword 
         Caption         =   "Change Password"
         Height          =   495
         Left            =   330
         TabIndex        =   2
         Top             =   900
         Width           =   2175
      End
      Begin VB.OptionButton optUsrAdd 
         Caption         =   " Add Users"
         Height          =   495
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetKannadaCaption()
Dim ctrl As Control
On Error Resume Next
For Each ctrl In Me
    ctrl.Font.Name = gFontName
    If Not TypeOf ctrl Is ComboBox Then
        ctrl.Font.Size = gFontSize
    End If
Next
Me.optUsrAdd.Caption = LoadResString(gLangOffSet + 496)
Me.optChangePassword.Caption = LoadResString(gLangOffSet + 497)
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)
Me.fraOptions.Caption = LoadResString(gLangOffSet + 491)
End Sub


Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
'centre the form
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
Call SetKannadaCaption
optChangePassword.Value = False
optUsrAdd.Value = False
End Sub

Private Sub optChangePassword_Click()
If optChangePassword.Value = True Then
frmPassWord.Show vbModal, Me
optChangePassword.Value = False
End If
 End Sub

Private Sub optUsrAdd_Click()
If optUsrAdd.Value = True Then
frmUsrAdd.Show vbModal, Me
optUsrAdd.Value = False
End If
End Sub


