VERSION 5.00
Begin VB.Form frmRepTemp 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   2370
   ClientTop       =   2265
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5610
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   4230
      TabIndex        =   11
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   345
      Left            =   2820
      TabIndex        =   4
      Top             =   2130
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      Begin VB.OptionButton optAcc 
         Caption         =   "Account reports"
         Height          =   435
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Width           =   1905
      End
      Begin VB.OptionButton optStock 
         Caption         =   "Stock Reports"
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   1935
      End
      Begin VB.OptionButton optFin 
         Caption         =   "Financial Reports"
         Height          =   315
         Left            =   2310
         TabIndex        =   1
         Top             =   90
         Width           =   1875
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   3165
         Begin VB.OptionButton optTriBal 
            Caption         =   "Trial Balance"
            Height          =   255
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optTrading 
            Caption         =   "Trading report"
            Height          =   285
            Left            =   180
            TabIndex        =   9
            Top             =   1470
            Width           =   1605
         End
         Begin VB.OptionButton optBal 
            Caption         =   "Balance Sheet"
            Height          =   375
            Left            =   180
            TabIndex        =   8
            Top             =   1110
            Width           =   2625
         End
         Begin VB.OptionButton optPL 
            Caption         =   "Profit && Loss"
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Top             =   840
            Width           =   2625
         End
         Begin VB.OptionButton optRP 
            Caption         =   "Receipt &&& Payament"
            Height          =   345
            Left            =   180
            TabIndex        =   6
            Top             =   540
            Width           =   2625
         End
      End
   End
End
Attribute VB_Name = "frmRepTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmdShow_Click()

If optStock Then
    frmRepSelect.Show 1
ElseIf optAcc Then
    Dim ReportsClass As clsReports

    Set ReportsClass = New clsReports
    
    Me.MousePointer = vbHourglass
    
    Do
        If Not ReportsClass.ShowReportForm Then
            Me.MousePointer = vbDefault
            Exit Do
        End If
    Loop
    
    Me.MousePointer = vbDefault
    
    'ReportsClass.ShowCurrentReport
    
    Set ReportsClass = Nothing

Else
    

    Set ReportsClass = New clsReports
    Me.MousePointer = vbHourglass
    If Not ReportsClass.ShowReportDate Then Exit Sub

    If optRP Then ReportsClass.ShowRPReport
    If optPL Then ReportsClass.ShowPandLAccount
    If optTriBal Then ReportsClass.ShowTrialBalance
    If optTrading Then ReportsClass.ShowTradingAccount
    If optBal Then ReportsClass.ShowBalanceSheet
    
    ReportsClass.ShowCurrentReport
    Me.MousePointer = vbDefault
    
    'ReportsClass.ShowCurrentReport
    
    Set ReportsClass = Nothing

    
End If

End Sub


Private Sub optFin_Click()

Frame2.Enabled = optFin

End Sub


