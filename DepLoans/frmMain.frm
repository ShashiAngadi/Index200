VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDX2000 - Loan Management module"
   ClientHeight    =   1935
   ClientLeft      =   2460
   ClientTop       =   1800
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   2850
      Top             =   990
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   2850
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   810
      TabIndex        =   1
      Top             =   1080
      Width           =   3705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Launch Loans module..."
      Height          =   495
      Left            =   810
      TabIndex        =   0
      Top             =   240
      Width           =   3705
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoanDlg As frmDepLoan

Private Sub cmdName_Click()
Dim OldDb As String
Dim NewDb As String

OldDb = App.Path & "\..\Index 2000.mdb"
NewDb = App.Path & "\DepositLoans.mdb"
If TransferNameTab(OldDb, NewDb) Then
    MsgBox "Customer Details transferrd"
End If

End Sub


Private Sub Command1_Click()
Dim LoanClass As clsDepLoan
Set LoanClass = New clsDepLoan
LoanClass.Show
End Sub

Private Sub Command2_Click()
gDBTrans.CloseDB
NudiStopKeyboardEngine
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not LoanDlg Is Nothing Then
    Unload LoanDlg
    Set LoanDlg = Nothing
End If
If Not gDBTrans Is Nothing Then Set gDBTrans = Nothing
        
End Sub

Private Sub GridPrint1_ProcessCount(Count As Long)

End Sub


