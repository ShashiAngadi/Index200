VERSION 5.00
Begin VB.Form frmUninst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "uninstall"
   ClientHeight    =   2385
   ClientLeft      =   2250
   ClientTop       =   1560
   ClientWidth     =   6585
   Icon            =   "Uninst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   2475
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue with uninstall "
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   1920
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6555
      Begin VB.Image img 
         Height          =   444
         Left            =   288
         Top             =   264
         Width           =   468
      End
      Begin VB.Label lblWarn 
         Alignment       =   2  'Center
         Caption         =   "You are about to remove iris from your computer are you sure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1152
         Left            =   156
         TabIndex        =   1
         Top             =   276
         Width           =   6228
      End
   End
End
Attribute VB_Name = "frmUninst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End
End Sub

Private Sub cmdContinue_Click()
cmdContinue.Enabled = False
frmProgress.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = (Screen.Width - Me.Width) \ 2
Me.Top = (Screen.Height - Me.Height) \ 2
Me.Caption = gAppName & " - Uninstall"
img.Picture = LoadResPicture(102, vbResIcon)
lblWarn.Caption = "Caution !!!" & vbCrLf & vbCrLf & "You are about to " & _
"remove " & gAppName & " from your computer." & vbCrLf & "Are you sure ?"


End Sub


