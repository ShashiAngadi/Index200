VERSION 5.00
Begin VB.Form frmInPut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   1680
   ClientTop       =   3555
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6480
      TabIndex        =   3
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6450
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Prompt messag will appear here "
      Height          =   735
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6165
   End
End
Attribute VB_Name = "frmInPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
End Sub

Private Sub cmdCancel_Click()

txtResult.Text = ""
Unload Me
'Me.Hide
End Sub

Private Sub cmdOk_Click()
'Unload Me
Me.Hide
End Sub


Private Sub Form_Load()
Call SetKannadaCaption
End Sub



