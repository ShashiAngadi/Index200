VERSION 5.00
Begin VB.Form frmDataMain 
   Caption         =   "Index Data Entry"
   ClientHeight    =   3165
   ClientLeft      =   615
   ClientTop       =   915
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdSB 
      Caption         =   "&Savings"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdMember 
      Caption         =   "&Members"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMember_Click()
    frmDataEntry.ShowForm (wis_Members)
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdSB_Click()
    frmDataEntry.ShowForm (wis_SBAcc)
End Sub
