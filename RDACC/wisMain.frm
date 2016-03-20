VERSION 5.00
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   3000
   ClientTop       =   2850
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5145
   Begin WIS_GRID_Print.GridPrint GridPrint 
      Left            =   660
      Top             =   1320
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Reccuring Deposit Account Module"
      Height          =   615
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   4605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   330
      TabIndex        =   1
      Top             =   1710
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
      Left            =   300
      TabIndex        =   2
      Top             =   180
      Width           =   4575
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdCustTrf_Click()
Dim OldDBName As String
Dim NewDBName As String

OldDBName = App.Path & "\..\AppMain\Index 2000.mdb"
NewDBName = App.Path & "\RDAcc.mdb"
If Not TransferNameTab(OldDBName, NewDBName) Then
'   MsgBox "Customer details transferred"
'Else
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

OldDBName = App.Path & "\..\AppMain\Index 2000.mdb"
NewDBName = App.Path & "\RDAcc.mdb"

StartTimer

If Not TransferNameTab(OldDBName, NewDBName) Then
'   MsgBox "Customer details transferred"
'Else
    MsgBox "unable to transafer Customer details "
    Exit Sub
End If

If TransferRD(OldDBName, NewDBName) Then
    If gDBTrans.WISCompactDB(NewDBName, "WIS!@#", "WIS!@#") Then
        StopTimer
        MsgBox "RD Details transferred"
    Else
        MsgBox "Problem While Compacting", vbCritical
    End If
Else
    MsgBox "Unable to transfer the RD detail"
End If

Screen.MousePointer = vbNormal

End Sub
Private Sub Form_Load()
lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub





