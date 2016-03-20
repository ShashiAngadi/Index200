VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   1500
   ClientTop       =   2835
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Modify User"
      Height          =   645
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   3285
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Login"
      Height          =   435
      Left            =   1170
      TabIndex        =   2
      Top             =   210
      Width           =   3315
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   465
      Left            =   1200
      TabIndex        =   1
      Top             =   1470
      Width           =   3315
   End
   Begin MSCommLib.MSComm Comm 
      Left            =   240
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   210
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   465
      Left            =   1230
      TabIndex        =   0
      Top             =   2100
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_Users As clsUsers
Attribute m_Users.VB_VarHelpID = -1

Private Sub cmdSearch_Click()
    frmSearch.Show vbModal
End Sub

Private Sub Command1_Click()
Comm.CommPort = 2
MsgBox Comm.Settings
Comm.PortOpen = True

MsgBox Comm.PortOpen

Comm.PortOpen = False

End Sub


Private Sub Command2_Click()
'm_Users.CancelError = True
m_Users.MaxRetries = 2
m_Users.CancelError = True
m_Users.ShowLoginDialog
'If Not m_Users.Login("admin", "admin", "") Then
'    MsgBox "INvalid login !", vbExclamation
'End If

End Sub

Private Sub Command3_Click()
    m_Users.ShowUserDialog
End Sub

Private Sub Form_Load()
    Set m_Users = New clsUsers
End Sub


