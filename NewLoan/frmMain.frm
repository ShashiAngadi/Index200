VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Management module"
   ClientHeight    =   6045
   ClientLeft      =   2070
   ClientTop       =   1305
   ClientWidth     =   4200
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4200
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   -90
      Top             =   2880
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   -30
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   5595
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   3705
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   5055
         Left            =   210
         ScaleHeight     =   4995
         ScaleWidth      =   3225
         TabIndex        =   9
         Top             =   300
         Width           =   3285
         Begin VB.CommandButton cmdLaunch 
            Caption         =   "Launch &Members Module..."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   60
            TabIndex        =   7
            Top             =   3450
            Width           =   3015
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   60
            TabIndex        =   8
            Top             =   4140
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00000000&
            Caption         =   "&Launch Bkcc Loans module..."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   60
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   6
            Top             =   2880
            Width           =   3015
         End
         Begin VB.CommandButton cmdReport 
            Caption         =   "&Reports"
            Height          =   465
            Left            =   60
            TabIndex        =   5
            Top             =   2340
            Width           =   3015
         End
         Begin VB.CommandButton cmdLoanDet 
            Caption         =   "Show Loan &Details"
            Height          =   465
            Left            =   60
            TabIndex        =   4
            Top             =   1785
            Width           =   3015
         End
         Begin VB.CommandButton cmdCustInfo 
            Caption         =   "&Customer Information"
            Height          =   465
            Left            =   60
            TabIndex        =   1
            Top             =   90
            Width           =   3015
         End
         Begin VB.CommandButton cmdLoanAcc 
            Caption         =   "Create &Loan Account"
            Height          =   495
            Left            =   60
            TabIndex        =   3
            Top             =   1215
            Width           =   3015
         End
         Begin VB.CommandButton cmdNewLoan 
            Caption         =   "Create &New Loan Scheme"
            Height          =   465
            Left            =   60
            TabIndex        =   2
            Top             =   660
            Width           =   3015
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            Height          =   4185
            Left            =   3390
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   2010
         TabIndex        =   10
         Top             =   270
         Width           =   75
      End
   End
   Begin VB.Menu mnuBtnContextMenu 
      Caption         =   "zz"
      Visible         =   0   'False
      Begin VB.Menu mnuBtnWhatsThis 
         Caption         =   "What's This?"
      End
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_clsLoan As clsLoan
Attribute m_clsLoan.VB_VarHelpID = -1
'Private WithEvents m_ClsCustDet As clsCustDet
Private m_ClsCustDet As clsCustReg
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_BankId As Integer
Private ThisControl As Control
Private m_OldDb As String

Private Sub cmdClose_Click()

Set frmLoanReport = Nothing
Set m_clsLoan = Nothing
Set m_ClsCustDet = Nothing


gDbTrans.CloseDB
Set gDbTrans = Nothing
Unload Me
End
End Sub

Private Sub cmdCustInfo_Click()
    m_ClsCustDet.ShowDialog
    gDbTrans.BeginTrans
    m_ClsCustDet.ModuleID = wis_Members
    m_ClsCustDet.SaveCustomer
    gDbTrans.CommitTrans
End Sub

Private Sub cmdDb_Click()
'Select Th DataBase From Where
'the data has to be transferred
With cdb
    .CancelError = False
    .Filter = "Data File (*.*)|*.mdb|AllFiles (*.*)|*.*"
    .ShowOpen
    m_OldDb = .FileName
End With
End Sub

Private Sub cmdLaunch_Click()
Dim MMAcc As New clsMMAcc
MMAcc.Show
End Sub

Private Sub Command1_Click()
Dim LoanClass As clsBkcc
Set LoanClass = New clsBkcc
LoanClass.Show
End Sub


Private Sub mnuBtnWhatsThis_Click()
    ThisControl.ShowWhatsThis
End Sub
Private Sub Command2_Click()
gDbTrans.CloseDB
Unload Me
End Sub

Private Sub Command3_Click()
m_clsLoan.ShowCreateLoanAccount
'frmLoanMaster.Show vbModal
End Sub

Private Sub Command4_Click()
frmLoanDetail.Show 1
End Sub


Private Sub cmdLoanAcc_Click()
m_clsLoan.ShowCreateLoanAccount
End Sub

Private Sub cmdLoanDet_Click()
m_clsLoan.ShowLoanAccountDetail
End Sub

Private Sub cmdNewLoan_Click()
m_clsLoan.ShowLoanSchemes
End Sub

Private Sub CMDrEPORTS_Click()

End Sub

Private Sub cmdReport_Click()
frmLoanReport.Show
End Sub


Private Sub Form_Load()

lbltitle.Caption = "NewLoan Module Version : " & App.Major _
                  & "." & App.Minor & "." & App.Revision
lbltitle.FontBold = True

cmdCustInfo.Enabled = False
Me.Caption = Me.Caption '& " - " & gBankName
Call CenterMe(Me)

Set m_ClsCustDet = New clsCustReg
Set m_clsLoan = New clsLoan

End Sub


