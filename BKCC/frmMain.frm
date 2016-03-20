VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#5.0#0"; "GRDPRINT.OCX"
Begin VB.Form wisMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDX2000 - Loan Management module"
   ClientHeight    =   1710
   ClientLeft      =   3705
   ClientTop       =   2940
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   4200
      Top             =   990
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   90
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   465
      TabIndex        =   1
      Top             =   900
      Width           =   3705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Launch Loans module..."
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   3705
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LoanDlg As frmBKCCAcc

Private Sub Command1_Click()
Dim LoanClass As clsBkcc
Set LoanClass = New clsBkcc
LoanClass.Show
End Sub

Private Sub Command2_Click()
'    gDBTrans.CloseDB
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Not LoanDlg Is Nothing Then
    Unload LoanDlg
    Set LoanDlg = Nothing
End If
If Not gDBTrans Is Nothing Then Set gDBTrans = Nothing
        
End Sub

