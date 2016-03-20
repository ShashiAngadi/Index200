VERSION 5.00
Begin VB.Form WisMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX 2000"
   ClientHeight    =   3915
   ClientLeft      =   2835
   ClientTop       =   2970
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5115
   Begin VB.CommandButton cmdRDTfr 
      Caption         =   "Transfer Rd Dtails"
      Height          =   615
      Left            =   270
      TabIndex        =   4
      Top             =   2100
      Width           =   4605
   End
   Begin VB.CommandButton cmdCustTrf 
      Caption         =   "Transfer Name details"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1380
      Width           =   4605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   270
      TabIndex        =   2
      Top             =   2880
      Width           =   4605
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Reccuring Deposit Account Module"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   4605
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Reccuring Deposit Account Module Ver."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   4575
   End
End
Attribute VB_Name = "WISMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCustTrf_Click()
Dim OldDBName As String
Dim NewDBName As String

OldDBName = App.Path & "\..\Index 2000.mdb"
NewDBName = App.Path & "\RDAcc.mdb"
If TransferNameTab(OldDBName, NewDBName) Then
   MsgBox "Customer details transferred"
Else
    MsgBox "unable to transafer Customer details "
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdLaunch_Click()
Dim RDAcc As New clsRDAcc
RDAcc.Show
End Sub

Private Sub cmdRDTfr_Click()
Dim OldDBName As String
Dim NewDBName As String

OldDBName = App.Path & "\..\Index 2000.mdb"
NewDBName = App.Path & "\RDAcc.mdb"

If TransferRD(OldDBName, NewDBName) Then
    MsgBox "RD Details transferred"
Else
    MsgBox "Unable to transfer the RD detail"
End If

End Sub

Private Sub Form_Load()
lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub




