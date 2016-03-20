VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   3030
   ClientTop       =   2070
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5115
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   180
      Top             =   2850
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Add User"
      Height          =   435
      Left            =   780
      TabIndex        =   5
      Top             =   2790
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer Name,Place.Caste"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2460
      TabIndex        =   4
      Top             =   2760
      Width           =   2355
   End
   Begin VB.CommandButton cmdTfr 
      Caption         =   "Transfer Sb Accounts"
      Height          =   555
      Left            =   300
      TabIndex        =   3
      Top             =   1380
      Width           =   4605
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2070
      Width           =   4605
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch SB Account Module"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   690
      Width           =   4605
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   90
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "SB Account Module Ver."
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
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SbAcc As New clsSBAcc

Private Sub cmdExit_Click()
gDBTrans.CloseDB

Set gDBTrans = Nothing
Set gCurrUser = Nothing
NudiStopKeyboardEngine
Unload Me
Set SbAcc = Nothing
End Sub

Private Sub cmdLaunch_Click()

SbAcc.Show

End Sub


Private Sub cmdTfr_Click()
On Error GoTo Errline

StartTimer

Command1_Click

Call gDBTrans.WISCompactDB("C:\Index 2000 Total New\SBACC\SBAcC.mdb", "WIS!@#", "WIS!@#")

If Not TransferSB(App.Path & "\..\Appmain\Index 2000.mdb", "C:\Index 2000 Total New\SBACC\SBAcC.mdb") Then
    MsgBox "Unable to transafer the sbdetails"
End If

If Not gDBTrans.WISCompactDB("C:\Index 2000 Total New\SBACC\SBAcC.mdb", "WIS!@#", "WIS!@#") Then
    MsgBox "Unable to Caomact the sbdetails"
End If

Screen.MousePointer = vbNormal

StopTimer

MsgBox "Sb Details data transferred"

Errline:
If Err Then _
    MsgBox "SB Data Transfer Failed", vbCritical, "Transferring SB Data"

End Sub

Private Sub cmdUser_Click()
gCurrUser.ShowUserDialog
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Not TransferNameTab(App.Path & "\..\Appmain\Index 2000.mdb", "C:\Index 2000 Total New\SBACC\SBAcC.mdb") Then
    MsgBox "Unable to transafer the Name Details"
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
lbl.Caption = lbl.Caption & App.Major & "." & App.Minor & "." & App.Revision
End Sub




