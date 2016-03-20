VERSION 5.00
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   2835
   ClientTop       =   3000
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5085
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   60
      Top             =   4200
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   90
      TabIndex        =   2
      Top             =   1680
      Width           =   4605
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Members Module"
      Height          =   615
      Left            =   150
      TabIndex        =   0
      Top             =   870
      Width           =   4605
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Members Module Ver."
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
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   4875
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdLaunch_Click()
Dim MMAcc As New clsMMAcc
MMAcc.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub




