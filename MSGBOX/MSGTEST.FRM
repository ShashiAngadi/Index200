VERSION 5.00
Begin VB.Form frmMsgTest 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   1500
   ClientTop       =   2565
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   6120
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   1140
      TabIndex        =   1
      Top             =   1110
      Width           =   3915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   1140
      TabIndex        =   0
      Top             =   300
      Width           =   3885
   End
End
Attribute VB_Name = "frmMsgTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Retval As Integer
'    Retval = MsgBox(" Test ", vbOKCancel, "")
frmMsgBox.Show vbModal
Unload frmMsgBox
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub


