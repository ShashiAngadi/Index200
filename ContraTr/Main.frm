VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX 2000"
   ClientHeight    =   3990
   ClientLeft      =   3090
   ClientTop       =   1755
   ClientWidth     =   5490
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5490
   Begin WIS_GRID_Print.GridPrint gfrdPrint 
      Left            =   1080
      Top             =   3540
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Add User"
      Height          =   465
      Left            =   3930
      TabIndex        =   6
      Top             =   3390
      Width           =   1005
   End
   Begin VB.PictureBox grdPrint 
      Height          =   480
      Left            =   2520
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3360
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   60
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCustTransfer 
      Caption         =   "Transfer Customer Details"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1950
      Width           =   4605
   End
   Begin VB.CommandButton cmdCATransfer 
      Caption         =   "Transfer CA Accounts"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1290
      Width           =   4605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   4605
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch CURRENT Account Module"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   630
      Width           =   4605
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "CURRENT Account Module Ver."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   5325
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCATransfer_Click()
Dim OldDb As String
Dim NewDB As String

gDBTrans.CloseDB
OldDb = App.Path & "\..\APPMAIN\Index 2000.mdb"
NewDB = App.Path & "\CAAcc.mdb"
'StartTimer

If Not TransferCA(OldDb, NewDB) Then
    MsgBox "Unable to Transfer the Current accounts", vbInformation, wis_MESSAGE_TITLE
Else
    If Not TransferNameTab(OldDb, NewDB) Then Exit Sub
    Dim je As JRO.JetEngine
    Set je = New JRO.JetEngine
    ' Make sure there isn't already a file with the
    ' name of the compacted database.
    Debug.Print FileLen(App.Path & "\CaAcC.mdb")
    
    Screen.MousePointer = vbHourglass
    If Dir(App.Path & "\NewCaAcC.mdb") <> "" Then _
       Kill App.Path & "\NewCaAcC.mdb"
    ' Compact the database specifying the new database password
    je.CompactDatabase "Data Source=" & App.Path & "\CaAcC.mdb;" & _
       "Jet OLEDB:Database Password=WIS!@#;", _
       "Data Source=" & App.Path & "\NewCaAcC.mdb;" & _
       "Jet OLEDB:Database Password=WIS!@#;"
    ' Delete the original database
    Kill App.Path & "\CaAcC.mdb"
    ' Rename the file back to the original name
    Name App.Path & "\NewCaAcC.mdb" As App.Path & "\CaAcC.mdb"
    
    Debug.Print FileLen(App.Path & "\CaAcC.mdb")
    
    Screen.MousePointer = vbNormal
    
'    StopTimer
    
    MsgBox "Current accounts transaferred ", vbInformation, wis_MESSAGE_TITLE
End If

If Not gDBTrans.OpenDB(App.Path & "\CAAcc.mdb", "WIS!@#") Then _
    MsgBox "Cannot Open Database", vbInformation
End Sub
Private Sub cmdCustTransfer_Click()
Dim OldDb As String
Dim NewDB As String
OldDb = App.Path & "\..\APPMAIN\Index 2000.mdb"
NewDB = App.Path & "\CAAcc.mdb"

If TransferNameTab(OldDb, NewDB) Then
    MsgBox "Customer details transferred"
Else
    MsgBox "Unable to transferr the Customer details "
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End
End Sub
Private Sub cmdLaunch_Click()
Dim CAAcc As New clsCAAcc
CAAcc.Show

End Sub
Private Sub cmdUser_Click()
gCurrUser.ShowUserDialog
End Sub
Private Sub Form_Load()
lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
lbl.ForeColor = &H8000000D
lbl.BackColor = &HC0FFC0
End Sub


