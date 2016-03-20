VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FD Deposit "
   ClientHeight    =   3045
   ClientLeft      =   1590
   ClientTop       =   2835
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   585
      Left            =   1380
      TabIndex        =   1
      Top             =   2040
      Width           =   3525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch FD Module"
      Height          =   585
      Left            =   1380
      TabIndex        =   0
      Top             =   1200
      Width           =   3525
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
      Height          =   495
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   6075
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim FD As New clsFDAcc

FD.Show

End Sub


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
'MsgBox "I HAve Taken PDAcc.frm As Dummy Form For Development" & vbCrLf & _
    " which has to remove after the implementation " & vbCrLf & " SHASHI"
End Sub


