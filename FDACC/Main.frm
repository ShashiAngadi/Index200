VERSION 5.00
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FD Deposit "
   ClientHeight    =   4470
   ClientLeft      =   1470
   ClientTop       =   1845
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6600
   Begin VB.CommandButton Command3 
      Caption         =   "Launch Depositloans  Module"
      Height          =   555
      Left            =   990
      TabIndex        =   4
      Top             =   1470
      Width           =   4545
   End
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   750
      Top             =   3030
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Add New User "
      Height          =   525
      Left            =   2280
      TabIndex        =   3
      Top             =   3330
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   555
      Left            =   990
      TabIndex        =   1
      Top             =   2160
      Width           =   4545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch FD Module"
      Height          =   555
      Left            =   990
      TabIndex        =   0
      Top             =   810
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Index 2000 - Fixed Deposit Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   6225
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FD As clsFDAcc

Private Sub cmdUser_Click()
gCurrUser.ShowUserDialog
End Sub


Private Sub Command1_Click()
    Set FD = New clsFDAcc
    FD.Show
    Set FD = Nothing
End Sub
Private Sub Command2_Click()
    If gWindowHandle > 0 Then Exit Sub
    NudiStopKeyboardEngine
    Unload Me
End Sub


Private Sub Command3_Click()
Dim ClsDeploan As New ClsDeploan
ClsDeploan.Show
End Sub

Private Sub Form_Load()
'MsgBox "I HAve Taken PDAcc.frm As Dummy Form For Development" & vbCrLf & _
    " which has to remove after the implementation " & vbCrLf & " SHASHI"

End Sub
Private Sub Form_Unload(Cancel As Integer)
gDBTrans.CloseDB
If Not gDBTrans Is Nothing Then Set gDBTrans = Nothing
End
End Sub


