VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmDepLoan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX-2000   -   Deposit Loan Wizard"
   ClientHeight    =   8325
   ClientLeft      =   1605
   ClientTop       =   1095
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   6720
      TabIndex        =   31
      Top             =   7740
      Width           =   1215
   End
   Begin VB.Frame fraRepayments 
      Height          =   6930
      Left            =   300
      TabIndex        =   32
      Top             =   525
      Width           =   7455
      Begin VB.CommandButton cmdLoanInst 
         Caption         =   "Loan Payment"
         Height          =   400
         Left            =   180
         TabIndex        =   30
         Top             =   6375
         Width           =   1335
      End
      Begin VB.ComboBox cmbDeposit 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   300
         Width           =   2175
      End
      Begin VB.CommandButton cmdRepayDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3240
         TabIndex        =   15
         Top             =   2580
         Width           =   315
      End
      Begin VB.TextBox txtRepayDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   16
         Top             =   2580
         Width           =   1305
      End
      Begin VB.TextBox txtAccNo 
         Height          =   315
         Left            =   5280
         MaxLength       =   9
         TabIndex        =   4
         Top             =   250
         Width           =   840
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6300
         TabIndex        =   5
         Top             =   250
         Width           =   1020
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Undo Last"
         Enabled         =   0   'False
         Height          =   400
         Left            =   4200
         TabIndex        =   29
         Top             =   6375
         Width           =   1575
      End
      Begin VB.CommandButton cmdRepay 
         Caption         =   "&Repay"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   400
         Left            =   6000
         TabIndex        =   25
         Top             =   6375
         Width           =   1335
      End
      Begin WIS_Currency_Text_Box.CurrText txtPrincAmount 
         Height          =   345
         Left            =   5670
         TabIndex        =   18
         Top             =   1260
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtRegInterest 
         Height          =   345
         Left            =   5670
         TabIndex        =   20
         Top             =   1695
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMisc 
         Height          =   345
         Left            =   5670
         TabIndex        =   22
         Top             =   2145
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtRepayAmt 
         Height          =   345
         Left            =   5670
         TabIndex        =   24
         Top             =   2580
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtIntBalance 
         Height          =   345
         Left            =   1860
         TabIndex        =   13
         Top             =   2115
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraLoanGrid 
         Height          =   2700
         Left            =   210
         TabIndex        =   26
         Top             =   3480
         Width           =   6970
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6420
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2130
            Width           =   435
         End
         Begin VB.CommandButton cmdPrevTrans 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6420
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   435
         End
         Begin VB.CommandButton cmdNextTrans 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6420
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   660
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2280
            Left            =   30
            TabIndex        =   33
            Top             =   270
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   4022
            _Version        =   393216
            Rows            =   10
            Cols            =   5
            AllowBigSelection=   0   'False
            ScrollBars      =   2
         End
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   3165
         Left            =   120
         TabIndex        =   71
         Top             =   3120
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   5583
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Instructions"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Account Statement"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2460
         Left            =   240
         TabIndex        =   72
         Top             =   3600
         Width           =   6885
         Begin VB.CommandButton cmdAddNote 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6450
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   210
            Width           =   450
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2355
            Left            =   60
            TabIndex        =   74
            Top             =   150
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   4154
            _Version        =   393217
            TextRTF         =   $"DepLnAcc.frx":0000
         End
      End
      Begin VB.Label lblCustName 
         Caption         =   "Label1"
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Width           =   1305
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   7200
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label lblDepositType 
         Caption         =   "DepositType "
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label lblRegInterest 
         Caption         =   "Reg Interest"
         Height          =   300
         Left            =   4005
         TabIndex        =   19
         Top             =   1755
         Width           =   1545
      End
      Begin VB.Label lblIntBlance 
         Caption         =   "Interest Balance"
         Height          =   300
         Left            =   225
         TabIndex        =   12
         Top             =   2205
         Width           =   1485
      End
      Begin VB.Label lblPrincipal 
         Caption         =   "Princ Amount"
         Height          =   300
         Left            =   4005
         TabIndex        =   17
         Top             =   1305
         Width           =   1485
      End
      Begin VB.Label lblMisc 
         Caption         =   "Miscellaneous"
         Height          =   300
         Left            =   4005
         TabIndex        =   21
         Top             =   2205
         Width           =   1545
      End
      Begin VB.Label lblRepayDate 
         AutoSize        =   -1  'True
         Caption         =   "Date of repayment : "
         Height          =   300
         Left            =   225
         TabIndex        =   14
         Top             =   2655
         Width           =   1650
      End
      Begin VB.Label lblRepayAmt 
         Caption         =   "Repaid amount :"
         Height          =   300
         Left            =   4005
         TabIndex        =   23
         Top             =   2655
         Width           =   1500
      End
      Begin VB.Label txtBalance 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         Top             =   1695
         Width           =   1665
      End
      Begin VB.Label txtLoanAmt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1860
         TabIndex        =   9
         Top             =   1260
         Width           =   1665
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance amount :"
         Height          =   300
         Left            =   225
         TabIndex        =   10
         Top             =   1770
         Width           =   1710
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Loan amount :"
         Height          =   300
         Left            =   225
         TabIndex        =   8
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1740
         TabIndex        =   7
         Top             =   750
         Width           =   5460
      End
      Begin VB.Label lblLoanNo 
         AutoSize        =   -1  'True
         Caption         =   "Loan Acc No"
         Height          =   300
         Left            =   4020
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   6810
      Left            =   300
      TabIndex        =   36
      Top             =   525
      Width           =   7455
      Begin VB.Frame fraOrder 
         Height          =   2205
         Left            =   210
         TabIndex        =   70
         Top             =   3900
         Width           =   6975
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advance"
            Height          =   400
            Left            =   5550
            TabIndex        =   67
            Top             =   1620
            Width           =   1215
         End
         Begin VB.TextBox txtStartDate 
            Height          =   315
            Left            =   1545
            TabIndex        =   63
            Top             =   1020
            Width           =   1230
         End
         Begin VB.TextBox txtEndDate 
            Height          =   315
            Left            =   4845
            TabIndex        =   66
            Top             =   1005
            Width           =   1230
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   315
            Left            =   2880
            TabIndex        =   62
            Top             =   1005
            Width           =   315
         End
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   315
            Left            =   6465
            TabIndex        =   65
            Top             =   1005
            Width           =   315
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   315
            Left            =   3600
            TabIndex        =   59
            Top             =   180
            Width           =   2925
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   345
            Left            =   330
            TabIndex        =   58
            Top             =   180
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   90
            X2              =   6540
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "From date :"
            Height          =   315
            Left            =   150
            TabIndex        =   61
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "To date :"
            Height          =   315
            Left            =   3585
            TabIndex        =   64
            Top             =   1065
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3675
         Left            =   210
         TabIndex        =   69
         Top             =   330
         Width           =   6975
         Begin VB.OptionButton optReport 
            Caption         =   "Monthly Balance"
            Height          =   315
            Index           =   6
            Left            =   3600
            TabIndex        =   56
            Top             =   2220
            Width           =   2805
         End
         Begin VB.ComboBox cmbRepDeposit 
            Height          =   315
            Left            =   2640
            TabIndex        =   50
            Top             =   390
            Width           =   3825
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Daily Cash Book(Contra)"
            Height          =   315
            Index           =   5
            Left            =   300
            TabIndex        =   52
            Top             =   1590
            Width           =   2655
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Loan Ledger"
            Height          =   555
            Index           =   4
            Left            =   330
            TabIndex        =   53
            Top             =   2190
            Width           =   3015
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Over Due Loans"
            Height          =   315
            Index           =   3
            Left            =   3600
            TabIndex        =   57
            Top             =   2850
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Sub day Book"
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   51
            Top             =   1050
            Width           =   2565
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Loan Details"
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   55
            Top             =   1650
            Width           =   2655
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Loan Balance"
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   54
            Top             =   1080
            Width           =   2865
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   6510
            X2              =   90
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label lblDep 
            Caption         =   "Loan Deposit Types :"
            Height          =   315
            Left            =   360
            TabIndex        =   49
            Top             =   420
            Width           =   1785
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   400
         Left            =   6000
         TabIndex        =   68
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame fraLoanAccounts 
      Caption         =   "Loan Accounts..."
      Height          =   6810
      Left            =   300
      TabIndex        =   35
      Top             =   525
      Width           =   7455
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   400
         Left            =   6060
         TabIndex        =   75
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoanUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6060
         TabIndex        =   48
         Top             =   5610
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoanIssueClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   6060
         TabIndex        =   60
         Top             =   6105
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveLoan 
         Caption         =   "C&reate"
         Height          =   400
         Left            =   6060
         TabIndex        =   47
         Top             =   5130
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   1110
         Left            =   270
         ScaleHeight     =   1050
         ScaleWidth      =   5655
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   375
         Width           =   5715
         Begin VB.Image Image2 
            Height          =   375
            Left            =   135
            Picture         =   "DepLnAcc.frx":0082
            Stretch         =   -1  'True
            Top             =   60
            Width           =   345
         End
         Begin VB.Label lblLoanIssueDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   990
            TabIndex        =   46
            Top             =   420
            Width           =   4710
         End
         Begin VB.Label lblLoanIssueHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            TabIndex        =   45
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox picLoanIssueViewPort 
         BackColor       =   &H00FFFFFF&
         Height          =   5160
         Left            =   270
         ScaleHeight     =   5100
         ScaleWidth      =   5670
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1500
         Width           =   5730
         Begin VB.PictureBox picLoanIssueSlider 
            Height          =   1740
            Left            =   -45
            ScaleHeight     =   1680
            ScaleWidth      =   5400
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   15
            Width           =   5460
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   4770
               TabIndex        =   42
               Top             =   0
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   -30
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.TextBox txtLoanIssue 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   0
               Left            =   2895
               TabIndex        =   40
               Top             =   0
               Width           =   2490
            End
            Begin VB.Label lblLoanIssue 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "lblIssuePrompt"
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   0
               Left            =   30
               TabIndex        =   43
               Top             =   0
               Width           =   2850
            End
         End
         Begin VB.VScrollBar vscLoanIssue 
            Height          =   1755
            Left            =   5415
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7485
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   13203
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Repayments"
            Key             =   "Repayments"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Loan Accounts"
            Key             =   "LoanAccounts"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDepLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_AccID As Long
Private m_LoanID As Long
Private m_CustID As Long

Private m_DepositType As Integer

Private m_rstLoanTrans As Recordset
Private m_rstLoanMast As Recordset
Private m_rstPledge As Recordset
Private m_TransID As Long

Private m_SelItem
Private m_Notes As New clsNotes
Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private m_clsRepOption As clsRepOption

'Varible to store the Pledge Deposit Details
Private WithEvents m_frmPldegeDeposits As frmDepSelect
Attribute m_frmPldegeDeposits.VB_VarHelpID = -1
Private m_PledgeDeposits As String
Private m_PledgeValue As Currency
Private m_AccArr() As String
Private m_DepTypeArr() As Integer

Const CTL_MARGIN = 15

' Declare events.
Public Event SetStatus(strMsg As String)
Public Event WindowClosed()
Public Event ShowReport(ReportType As wis_DepLoanReports, ReportOrder As wis_ReportOrder, _
            DepositType As Integer, fromDate As String, toDate As String, RepOption As clsRepOption)
            
Public Event AccountChanged(ByVal AccId As Long)
Public Event AccountTransaction(transType As wisTransactionTypes)
Private Function GetInterstBalance() As Currency

Err.Clear
On Error GoTo Err_line

Dim IntBalance As Currency
Dim bkMark

If m_LoanID = 0 Then GoTo Exit_Line
If m_rstLoanTrans Is Nothing Then GoTo Exit_Line
If m_rstLoanTrans.EOF Then m_rstLoanTrans.MoveLast

bkMark = m_rstLoanTrans.Bookmark

Do
    If m_rstLoanTrans.BOF Or m_rstLoanTrans.EOF Then Exit Do
    If m_rstLoanTrans(0) = "INTEREST" Then
        IntBalance = FormatField(m_rstLoanTrans("Balance"))
        If m_rstLoanTrans("Amount") > 0 Then Exit Do
    End If
    m_rstLoanTrans.MovePrevious
Loop

m_rstLoanTrans.Bookmark = bkMark

Exit_Line:
GetInterstBalance = IntBalance
Exit Function
Err_line:
    If Err.Number Then MsgBox "Error in ComputeInterestBalance :" & Err.Description & _
         vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Err.Clear
    
End Function

'
Private Function LoanRepay() As Boolean
On Error GoTo Err_line

' Variables used in this procedure...
' -----------------------------------
Dim TransID As Long
Dim inTransaction As Boolean
Dim NewBalance As Currency
Dim Balance As Currency
Dim PrincAmount  As Currency
Dim IntBalance As Currency
Dim PayAmount As Currency
Dim IntAmount As Currency
Dim C_IntAmount As Currency
Dim PrevIntBalance As Currency
Dim MiscAmount As Currency
Dim IntPaidDate As Date
Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim RetMsg As Integer
Dim rst As ADODB.Recordset
' -----------------------------------
'Check for the date specified
TransDate = GetSysFormatDate(txtRepayDate.Text)
 
'Calculate the RegInt & Penal Int on the Specified date
IntAmount = txtRegInterest
C_IntAmount = ComputeDepLoanRegularInterest(TransDate, m_LoanID)
IntBalance = GetInterstBalance

MiscAmount = txtMisc
PayAmount = txtRepayAmt

If PayAmount <= 0 Then
    'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    txtRepayAmt.SetFocus
    Exit Function
End If

PrincAmount = PayAmount - IntAmount - MiscAmount '- IntBalance
'if Princpal amount is negative then
If PrincAmount < 0 Then
    'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    txtRepayAmt.SetFocus
    Exit Function
End If

'If interest amount he is paying is lees than
'his interest balance then warn about this
If IntAmount > 0 Then
    If IntAmount < IntBalance Then
        'RetMsg = MsgBox("Repayment amount is less than the Interest balance" & _
            vbCrLf & "Do you want to contine?", vbYesNoCancel, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(670) & _
            vbCrLf & GetResourceString(541), vbYesNo, wis_MESSAGE_TITLE)
        If RetMsg = vbNo Then Exit Function
    ElseIf IntAmount < C_IntAmount + IntBalance Then
    'If interest amount he is paying is lees than
    'his interest till date then confirm whether to
    'Keep the difference amount as interest balance
        'RetMsg = MsgBox("Repayment amount is less than Interest till date" & _
            vbCrLf & "Do want to keep the difference amount as Balance", vbYesNoCancel, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(668) & vbCrLf & GetResourceString(669) _
            , vbYesNoCancel + vbDefaultButton1, wis_MESSAGE_TITLE)
        If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then C_IntAmount = IntAmount - IntBalance
    ElseIf IntAmount > C_IntAmount + IntBalance Then
      'If interest amount he is paying is more than
      'his interest till date then confirm whether to
      'Keep the difference amount as extra balance
        'RetMsg = MsgBox("Repayment amount is more than Interest till date" & _
            vbCrLf & "Do want to keep the difference amount as Balance", vbYesNoCancel, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(666) & vbCrLf & GetResourceString(667) _
            , vbYesNoCancel + vbDefaultButton2, wis_MESSAGE_TITLE)
        If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then C_IntAmount = IntAmount - IntBalance
    End If
End If


' Get a balance and new transactionID.
Dim LastTransDate As Date

gDbTrans.SqlStmt = "SELECT Balance,TransDate,TransID FROM DepositLOantrans " & _
    " WHERE LoanID = " & m_LoanID & " Order By TransID Desc "
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then GoTo Exit_Line
Balance = Val(FormatField(rst("Balance")))
TransID = FormatField(rst("Transid"))
LastTransDate = rst("TransDate")

gDbTrans.SqlStmt = "SELECT * FROM DepositLoanIntTrans " & _
        " WHERE LoanID = " & m_LoanID & " ORDER By TransID Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If FormatField(rst("TransID")) > TransID Then
        TransID = FormatField(rst("TransID"))
        LastTransDate = rst("TransDate")
    End If
End If
NewBalance = Balance - PrincAmount

TransID = TransID + 1

'validate the date here if he has specified the previous date
If DateDiff("d", LastTransDate, TransDate) < 0 Then
        MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    Exit Function
End If

' Begin the transaction
Dim bankClass As clsBankAcc
Dim LoanHeadID As Long
Dim IntHeadID As Long
Dim MiscHeadId As Long
Dim headName As String
Dim LoanHeadName As String
Dim LoanHeadNameEnglish  As String

If Not gDbTrans.BeginTrans Then GoTo Exit_Line
inTransaction = True
Set bankClass = New clsBankAcc

If m_DepositType Then
    LoanHeadName = GetDepositTypeTextEnglish(m_DepositType, LoanHeadNameEnglish)
    LoanHeadName = LoanHeadName & " " & GetResourceString(58)
    If Len(LoanHeadNameEnglish) > 0 Then LoanHeadNameEnglish = LoanHeadNameEnglish & " " & LoadResString(58)
Else
    LoanHeadName = GetResourceString(43, 58)
    If Len(LoanHeadNameEnglish) > 0 Then LoanHeadNameEnglish = LoadResourceStringS(43, 58)
End If

LoanHeadID = bankClass.GetHeadIDCreated(LoanHeadName, LoanHeadNameEnglish, parMemDepLoan, 0, wis_DepositLoans + m_DepositType)
MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), parBankIncome, 0, wis_None)

LoanHeadName = LoanHeadName & " " & GetResourceString(483)
If Len(LoanHeadNameEnglish) > 0 Then LoanHeadNameEnglish = LoanHeadNameEnglish & " " & LoadResString(483)
IntHeadID = bankClass.GetHeadIDCreated(LoanHeadName, LoanHeadNameEnglish, parMemDepLoan, 0, wis_DepositLoans + m_DepositType)

transType = wDeposit

''Very First update the principle amount
' Update the Princple Amount.
gDbTrans.SqlStmt = "INSERT INTO  DepositLoantrans " _
        & "(LoanID, TransID, TransType, Amount, " _
        & "TransDate, Balance, Particulars ) " _
        & "VALUES (" & m_LoanID & ", " & TransID & ", " _
        & transType & ", " & PrincAmount & ", " _
        & "#" & TransDate & "#, " _
        & NewBalance & ",'" & GetResourceString(76) & "')"
'If he is paying any Principle amount then Count Make the transaction
If PrincAmount Then
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    If Not bankClass.UpdateCashDeposits(LoanHeadID, PrincAmount, _
        TransDate) Then GoTo Exit_Line
End If

' Update the interest amount.
If IntAmount + MiscAmount > 0 Then
    'Get th interest balance
    IntBalance = IntBalance + C_IntAmount - IntAmount
    gDbTrans.SqlStmt = "INSERT INTO  DepositLoanIntTrans  " _
            & "(LoanID, TransID,TransType, Amount, " _
            & "PenalAmount,MiscAmount,TransDate, Balance, Particulars ) " _
            & "VALUES (" & m_LoanID & ", " & TransID & ", " _
            & transType & ", " & IntAmount & ", " _
            & "0 " & ", " & MiscAmount & ", #" _
            & TransDate & "#, " _
            & IntBalance & ", '" & GetResourceString(47) & "')"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    ' Update the loan master with interest balance.
    ' In this case, the total repaid amt is less than the interest payable.
    ' Therefore, put this difference amt, to loan master table.
    gDbTrans.SqlStmt = "UPDATE DepositLoanMaster SET LastIntDate = " _
            & "#" & TransDate & "# WHERE loanid = " & m_LoanID
    ' Execute the updation.
    If IntAmount > 0 Then
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
        If Not bankClass.UpdateCashDeposits(IntHeadID, IntAmount, _
            TransDate) Then GoTo Exit_Line
    End If
    If MiscAmount > 0 Then
        'If Not gDbTrans.SQLExecute Then GoTo Exit_line
    
        If Not bankClass.UpdateCashDeposits(MiscHeadId, MiscAmount, _
            TransDate) Then GoTo Exit_Line
    End If
End If


    ' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.
If NewBalance = 0 Then
    gDbTrans.SqlStmt = "UPDATE DepositLoanMaster SET " & _
        " LoanClosedDate = #" & TransDate & "# Where LoanID =  " & m_LoanID
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If
   
    ' Commit the transaction.
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False
    
LoanRepay = True
'MsgBox "Loan repayment accepted.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(706), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    Exit Function

Err_line:
    If Err Then
        MsgBox "AcceptPayment: " & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(707) & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line
End Function

'
Private Sub ArrangePropSheet()

Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean
lblLoanIssueDesc.BorderStyle = 0
lblLoanIssueHeading.BorderStyle = 0
fraLoanAccounts.Caption = ""

' Arrange the Slider panel.
With picLoanIssueSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCountLoanIssue
    .Height = txtLoanIssue(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    ' If the height is greater than viewport height,
    ' the scrollbar needs to be displayed.  So,
    ' reduce the width accordingly.
    If .Height > picLoanIssueViewPort.ScaleHeight Then
        NeedsScrollbar = True
        .Width = picLoanIssueViewPort.ScaleWidth - _
                vscLoanIssue.Width
    Else
        .Width = picLoanIssueViewPort.ScaleWidth
    End If

End With

' Set/Reset the properties of scrollbar.
With vscLoanIssue
    .Height = picLoanIssueViewPort.ScaleHeight
    .Min = 0
    .Max = picLoanIssueSlider.Height - picLoanIssueViewPort.ScaleHeight
    If .Max < 0 Then .Max = 0
    .SmallChange = txtLoanIssue(0).Height
    .LargeChange = picLoanIssueViewPort.ScaleHeight / 2
End With

' Adjust the text controls on this panel.
Dim I As Integer
For I = 0 To txtLoanIssue.count - 1
    txtLoanIssue(I).Width = picLoanIssueSlider.ScaleWidth _
            - lblLoanIssue(I).Width - CTL_MARGIN
Next


If NeedsScrollbar Then
    vscLoanIssue.Visible = True
End If

For I = 0 To txtLoanIssue.count - 1
    txtLoanIssue(I).Width = picLoanIssueSlider.ScaleWidth - _
        (lblLoanIssue(I).Left + lblLoanIssue(I).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For I = 0 To cmb.count - 1
    cmb(I).Width = txtLoanIssue(I).Width
Next
For I = 0 To cmd.count - 1
    cmd(I).Left = txtLoanIssue(I).Left _
        + txtLoanIssue(I).Width - cmd(I).Width
Next

'' Draw lines for the remaining portions of the viewport.
'With picLoanIssueViewPort
'    .CurrentX =


End Sub

'
' Fills the details of members to report window.
' Instead of standard fillview function, this is
' written to handle the specific requirements of fill members...
'
Private Function FillMembers(ListViewCtl As ListView, rs As Recordset, Optional AutoWidth As Boolean) As Boolean
On Error GoTo fillmembers_error
Const FIELD_MARGIN = 1.5

' Check if there are any records in the recordset.
rs.MoveLast
rs.MoveFirst
If rs.RecordCount = 0 Then
    FillMembers = True
    GoTo Exit_Line
End If

Dim I As Integer
Dim itmX As ListItem
With ListViewCtl
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    For I = 1 To rs.Fields.count - 1
        .ColumnHeaders.Add , rs.Fields(I).name, rs.Fields(I).name, _
                    ListViewCtl.Parent.TextWidth(rs.Fields(I).name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                    rs.Fields(I).Type = adInteger Or _
                    rs.Fields(I).Type = adBigInt Or _
                    rs.Fields(I).Type = adDouble Or _
                    rs.Fields(I).Type = adCurrency Then
                .ColumnHeaders(I).Alignment = lvwColumnRight
            End If
        End If
    Next

    ' Begin a loop for processing rows.
    Do While Not rs.EOF
        ' Add the details.
        Set itmX = .ListItems.Add(, , rs.Fields(1))
        ' If the 'Autowidth' property is enabled,
        ' then check if the width needs to be expanded.
        If AutoWidth Then
            If .ColumnHeaders(1).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.Fields(1))) Then
                .ColumnHeaders(1).Width = _
                    .Parent.TextWidth(FormatField(rs.Fields(1))) * FIELD_MARGIN
            End If
        End If
        ' Add sub-items.
        For I = 2 To rs.Fields.count - 1
            ' If the field name is Category
            If StrComp(rs.Fields(I).name, "Category", vbTextCompare) = 0 Then
                itmX.SubItems(I - 1) = IIf((FormatField(rs.Fields(I)) = 1), _
                        GetResourceString(440), GetResourceString(441))
            Else
                itmX.SubItems(I - 1) = FormatField(rs.Fields(I))
            End If

            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then
                If .ColumnHeaders(I).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(itmX.SubItems(I - 1)) Then
                    .ColumnHeaders(I).Width = _
                        .Parent.TextWidth(itmX.SubItems(I - 1)) * FIELD_MARGIN
                End If
            End If
        Next
        rs.MoveNext
    Loop
End With
FillMembers = True

Exit_Line:
ListViewCtl.Visible = True
ListViewCtl.view = lvwReport
Exit Function

fillmembers_error:
    If Err.Number = 3265 Then
        On Error Resume Next
        Resume
    ElseIf Err Then
        MsgBox "FillMembers: The following error occurred." _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(710) _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function
'
Public Function LoanLoad(lLoanID As Long) As Boolean
On Error GoTo LoanLoad_Error

' Declare variables needed for this procedure...
Dim Lret As Long
Dim LastDate As String
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes
Dim IntAmount As Currency
Dim txtIndex As Byte
Dim rst As ADODB.Recordset
Dim Offset As Long
Dim PledgeDesc As String
Dim strArr() As String

' Get the loan details from the LoanMaster table.
gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE " _
    & " LoanID = " & lLoanID
Lret = gDbTrans.Fetch(m_rstLoanMast, adOpenStatic)
If Lret < 1 Then GoTo Exit_Line
m_LoanID = lLoanID

Call m_Notes.LoadNotes(wis_DepositLoans, lLoanID)
Me.TabStrip1.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True

' Check wheather loanclosed or not
If FormatField(m_rstLoanMast("Loanclosed")) Then
    cmdLoanInst.Enabled = False
    cmdLoanUpdate.Enabled = False
    cmdRepay.Enabled = True
    cmdUndo.Caption = GetResourceString(313) '"Reopen"
    cmdAddNote.Enabled = False
    rtfNote.BackColor = wisGray
    rtfNote.Enabled = False
Else
    cmdLoanInst.Enabled = True
    cmdLoanUpdate.Enabled = True
    cmdRepay.Enabled = True
    cmdUndo.Enabled = gCurrUser.IsAdmin
    cmdUndo.Caption = GetResourceString(14) '"Delete"
    cmdAddNote.Enabled = True
    rtfNote.BackColor = vbWhite
    rtfNote.Enabled = True
End If

' Fill the loan amount.
txtLoanAmt.Caption = FormatField(m_rstLoanMast("LoanAmount"))
'Now Load the Same Loan detials into loan issue loan grid
' Check for member details.
Dim CustClass As New clsCustReg
txtIndex = GetIndex("MemberID")
frmPhoto.setAccNo (FormatField(m_rstLoanMast("CustomerId")))
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("CustomerId"))
txtLoanIssue(txtIndex + 1).Text = CustClass.CustomerName(FormatField(m_rstLoanMast("CustomerId")))

Set CustClass = Nothing
txtCustNAme = txtLoanIssue(txtIndex + 1).Text
'Loan Details
txtIndex = GetIndex("LoanAccNo")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("AccNum"))

m_DepositType = FormatField(m_rstLoanMast("DepositType"))
m_DepositType = m_DepositType Mod 100

'Get the Deposit Type & Name
txtIndex = GetIndex("DepositType")
txtIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
Dim I As Integer
With cmb(txtIndex)
    If .ListCount > 0 Then .ListIndex = 0: cmbDeposit.ListIndex = 0
    For I = 0 To .ListCount - 1
        If m_DepositType = .ItemData(I) Then
            .ListIndex = I
            txtLoanIssue(Val(GetIndex("DepositType"))).Text = .List(I)
            cmbDeposit.ListIndex = I
            Exit For
        End If
    Next
End With

' Pledge item
txtIndex = GetIndex("PledgeDeposit")
m_PledgeDeposits = FormatField(m_rstLoanMast("PledgeDescription"))
txtLoanIssue(txtIndex).Text = m_PledgeDeposits

'Number of Pldge
txtIndex = GetIndex("NoOfDep")
txtLoanIssue(txtIndex).Text = GetStringArray(m_PledgeDeposits, strArr, gDelim) + 1

' Pledge Value
txtIndex = GetIndex("PledgeValue")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("PledgeValue"))
m_PledgeValue = Val(txtLoanIssue(txtIndex))

' Loan Amount.
txtIndex = GetIndex("LoanAmount")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("LoanAmount"))

' Loan due date.
txtIndex = GetIndex("DueDate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("LoanDueDate"))

'Interest Rate
txtIndex = GetIndex("InterestRate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("InterestRate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "InterestRate", FormatField(m_rstLoanMast("InterestRate")))

'Penal Interest Rate
txtIndex = GetIndex("PenalInterestRate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("PenalInterestRate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "PenalInterestRate", FormatField(m_rstLoanMast("PenalInterestRate")))

' Create Date
txtIndex = GetIndex("IssueDate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("LoanIssueDate"))
'txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "IssueDate", FormatField(m_rstLoanMast("IssueDate")))

'Remarks
txtIndex = GetIndex("Remarks")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("Remarks"))

' Get the balance amount on this loan.
gDbTrans.SqlStmt = "SELECT TOP 1 Balance,Transid " & _
        " FROM DepositLoanTrans WHERE" & _
        " LoanID = " & lLoanID & " ORDER BY transID Desc"
Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Lret <= 0 Then GoTo Exit_Line

txtBalance.Caption = FormatField(rst(0))
cmdUndo.Caption = GetResourceString(5) '"uNDO"


' Get all the transactions on this loan.
gDbTrans.SqlStmt = "SELECT 'PRINCIPLE',TransDate,TransID, TransType,Amount,Balance " & _
    "FROM DepositLoantrans WHERE " & _
    "loanid = " & lLoanID & " " & _
    " UNION " & _
    "SELECT 'INTEREST',TransDate,TransID, TransType,Amount,Balance " & _
    "FROM DepositLoanIntTrans WHERE " & _
    "Loanid = " & lLoanID & " " & _
    "ORDER BY TransID "

Lret = gDbTrans.Fetch(m_rstLoanTrans, adOpenForwardOnly)
If Lret <= 0 Then GoTo Exit_Line

If m_rstLoanTrans.RecordCount = 1 Then _
        cmdUndo.Caption = GetResourceString(14)  '' "Delete"

m_rstLoanTrans.MoveLast
m_TransID = FormatField(m_rstLoanTrans.Fields("TransID"))
m_rstLoanTrans.MoveFirst
cmdPrevTrans.Enabled = False
cmdNextTrans.Enabled = False
If m_TransID > 10 Then cmdPrevTrans.Enabled = True: cmdPrevTrans = True

'Load the recordset details to loan grid.
Call LoanLoadGrid

cmdSaveLoan.Enabled = False

' Compute and display the regular interest.
IntAmount = ComputeDepLoanRegularInterest(GetSysFormatDate(txtRepayDate.Text), lLoanID)
txtRegInterest = Val(FormatCurrency(IntAmount \ 1))

'Get the Interest Balance And Loan TransCtion
txtIntBalance = GetInterstBalance

'Display the total instalment amount.
txtRepayAmt = 0

'Now get the Pledge Deposit Information
gDbTrans.SqlStmt = "SELECt * From PledgeDeposit" & _
        " WHERE LoanId = " & lLoanID
        
If gDbTrans.Fetch(m_rstPledge, adOpenDynamic) <= 0 Then GoTo Exit_Line
Lret = m_rstPledge.RecordCount

ReDim m_AccArr(Lret - 1)
ReDim m_DepTypeArr(Lret - 1)
Lret = 0
If m_frmPldegeDeposits Is Nothing Then
    Set m_frmPldegeDeposits = New frmDepSelect
    Load m_frmPldegeDeposits
End If
m_frmPldegeDeposits.LoadDeposits (m_rstLoanMast("CustomerId"))

While Not m_rstPledge.EOF
    m_AccArr(Lret) = m_rstPledge("AccId")
    m_DepTypeArr(Lret) = m_rstPledge("DepositType")
    Lret = Lret + 1
    m_rstPledge.MoveNext
Wend

txtIndex = GetIndex("DepositType")
txtIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
If Not m_rstLoanTrans Is Nothing Then
    cmb(txtIndex).Locked = True
    cmbDeposit.Locked = True
End If

'Get the Intrest Balance
txtIndex = GetIndex("InterestBalance")

'Now load the inttrestBalance to the registration Details.
Dim rstIntTrans As Recordset
Dim TransDate As String

gDbTrans.SqlStmt = "SELECT * from DepositLoanIntTrans where Loanid=" & m_LoanID & _
                   " And TransID = (Select Max(TransID) From " & _
                   " DepositLoanIntTrans Where LoanID = " & m_LoanID & ")"

Call gDbTrans.Fetch(rstIntTrans, adOpenForwardOnly)
txtIndex = GetIndex("InterestBalance")
txtLoanIssue(txtIndex).Text = FormatField(rstIntTrans("Balance"))

RaiseEvent AccountChanged(lLoanID)
LoanLoad = True
cmdPhoto.Enabled = Len(gImagePath)
m_LoanID = lLoanID

Exit_Line:
Exit Function

LoanLoad_Error:
    If Err Then
        MsgBox "LoanLoad: " & Err.Description, vbCritical
'        MsgBox GetResourceString(713) & Err.Description, vbCritical
        Resume
    End If
    GoTo Exit_Line
End Function

Private Sub cmdPhoto_Click()
If m_CustID > 0 Then
    frmPhoto.setAccNo (m_CustID)
        If (m_CustID > 0) Then frmPhoto.Show vbModal
End If
End Sub

'
Private Sub LoanLoadGrid()
Dim transType As wisTransactionTypes
Dim Balance As Currency
Err.Clear
On Error GoTo Err_line

' If no recordset, exit.
If m_rstLoanTrans Is Nothing Then Exit Sub

' If no records, exit.
If m_rstLoanTrans.RecordCount = 0 Then Exit Sub

'Show 10 records or till eof of the page being pointed to
With grd
    ' Initialize the grid.
    .Clear: .AllowUserResizing = flexResizeBoth
    ' Set the format string for the loan details grid.
    .Cols = 5
    .Rows = 11
    .FixedRows = 1
    .FixedCols = 0
    .Row = 0
    .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1200  'Date
    .Col = 1: .Text = GetResourceString(81): .ColWidth(1) = 1150 'Loan issued
    .Col = 2: .Text = GetResourceString(82): .ColWidth(2) = 1200 'LOan Repayment
    .Col = 3: .Text = GetResourceString(47): .ColWidth(3) = 950 'Interest
    .Col = 4: .Text = GetResourceString(42): .ColWidth(4) = 1100  'Balance
    .Visible = False

    Dim NextRow As Boolean
    Dim TransDate As Date
    Dim TransID As Long
    'Dim TransType As wisTransactionTypes
    
    .Row = .FixedRows: TransID = m_rstLoanTrans("TransID")
    m_TransID = TransID
    Do
        If TransID <> m_rstLoanTrans("TransID") Then
            If .Row = .Rows - 1 Then Exit Do
            .Row = .Row + 1
        End If
        TransID = m_rstLoanTrans("TransID")
        If m_rstLoanTrans(0) = "INTEREST" Then
            .Col = 0: .Text = FormatField(m_rstLoanTrans("TransDate"))
            .Col = 3: .Text = FormatCurrency(Val(.Text) + m_rstLoanTrans("Amount"))
        Else
            .Col = 0: .Text = FormatField(m_rstLoanTrans("TransDate"))
            transType = m_rstLoanTrans("transType")
            
            .Col = IIf(transType = wDeposit Or transType = wContraDeposit, 2, 1)
            .Text = FormatCurrency(Val(.Text) + m_rstLoanTrans("Amount"))
            .Col = 4: .Text = FormatField(m_rstLoanTrans("Balance"))
        End If
nextRecord:
        m_rstLoanTrans.MoveNext
        If m_rstLoanTrans.EOF Then Exit Do
    Loop
    .Visible = True
    .Row = 1
    cmdNextTrans.Enabled = Not CBool(m_rstLoanTrans.EOF)
    cmdPrevTrans.Enabled = CBool(m_TransID >= 10)
    
End With

'Enable the UNDO Button
cmdUndo.Enabled = gCurrUser.IsAdmin

Exit Sub

Err_line:
    If Err Then
        MsgBox "LoanLoadGrid: " & Err.Description
        'MsgBox GetResourceString(714) & Err.Description
    End If
'Resume
End Sub
'
Private Function LoanSave() As Boolean

' Setup error handler.
On Error GoTo LoanSave_Error

' Declare variables for this procedure...
' ------------------------------------------
Dim strDep As String
Dim txtIndex As Integer
Dim inTransaction As Boolean
Dim Lret As Long, nRet As Integer
Dim NewLoanID As Long
Dim newTransID As Long
Dim Balance As Currency
Dim LoanAmount As Currency
Dim PledgeAmount As Currency
Dim count As Integer
Dim custId As Long
Dim LoanAccNum As String
Dim SbACCID As Long
Dim IssueDate As Date
Dim DueDate As Date
Dim rst As ADODB.Recordset
Dim SqlStr As String
Dim DepositType As Integer

' ------------------------------------------

DepositType = GetDepositType(GetVal("DepositType"))
DepositType = DepositType Mod 100

' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)
    custId = Val(.Text)
End With

'Check for loan account number detaisl.
txtIndex = GetIndex("LoanAccNo")
With txtLoanIssue(txtIndex)
    LoanAccNum = Trim$(.Text)
End With

txtIndex = GetIndex("PledgeValue")
With txtLoanIssue(txtIndex)
    PledgeAmount = Val(.Text)
End With

'Loan Amount.
txtIndex = GetIndex("LoanAmount")
With txtLoanIssue(txtIndex)
    LoanAmount = (.Text)
End With

'Loan issue date.
txtIndex = GetIndex("IssueDate")
With txtLoanIssue(txtIndex)
    IssueDate = GetSysFormatDate(.Text)
End With

'Loan due date.
txtIndex = GetIndex("DueDate")
With txtLoanIssue(txtIndex)
    DueDate = GetSysFormatDate(.Text)
End With

' Get a new loanid.
gDbTrans.SqlStmt = "SELECT MAX(LoanID) FROM DepositLoanMaster"
Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Lret <= 0 Then GoTo Exit_Line

NewLoanID = Val(FormatField(rst(0))) + 1

' Get the TransID and balance for updating LoanTrans Table.
gDbTrans.SqlStmt = "SELECT TOP 1 TransID, BALANCE " & _
                    " FROM DepositLoanTrans" & _
                    " WHERE LoanID = " & NewLoanID
Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

If Lret = 0 Then
    newTransID = 1
    LoanAmount = Val(GetVal("LoanAmount"))
    Balance = LoanAmount
Else
    newTransID = Val(FormatField(rst(0))) + 1
    LoanAmount = Val(GetVal("LoanAmount"))
    Balance = Val(FormatField(rst(1))) + LoanAmount
End If

' Begin transaction.
inTransaction = True
gDbTrans.BeginTrans

'Dim Sqlstr As String

' Put an entry into LoanMaster table.
SqlStr = "INSERT INTO DepositLoanMaster " & _
        " (LoanID, AccNum,DepositType, CustomerID," & _
            " LoanIssueDate, PledgeValue, PledgeDescription, LoanAmount," & _
            "LoanDueDate,InterestRate,PenalInterestRate, Remarks,UserID)"

SqlStr = SqlStr & "  VALUES (" & NewLoanID & ", " & _
        AddQuotes(LoanAccNum, True) & "," & DepositType & ", " & _
        GetVal("MemberID") & ", " & _
        " #" & IssueDate & "#, " & _
        Val(GetVal("PledgeValue")) & ", " & _
        AddQuotes(m_PledgeDeposits, True) & ", " & _
        LoanAmount & ", " & _
        "#" & DueDate & "#, " & _
        Val(GetVal("InterestRate")) & "," & _
        Val(GetVal("PenalInterestRate")) & "," & _
        AddQuotes(GetVal("Remarks"), True) & _
        "," & gUserID & " )"
        
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
'here we are creating loan account only
'and we are not performing any transaction to the transtable
    
' Commit the transaction
'If Not gDbTrans.CommitTrans Then GoTo Exit_line
'inTransaction = False

'Now set the mark to deposits as Pledge Deposits
Dim TableName As String
Dim PledgeAccId As Long


For count = 0 To UBound(m_AccArr)
    DepositType = m_DepTypeArr(count)
    DepositType = DepositType Mod 100
    If DepositType = wisDeposit_RD Then
        TableName = "RDMASTER"
    ElseIf DepositType = wisDeposit_PD Then
        TableName = "PDMASTER"
    Else
        TableName = "FDMASTER"
    End If
    SqlStr = "SELECT AccID From " & TableName & " WHERE " & _
                " AccNum = " & AddQuotes(m_AccArr(count), True)
    If TableName = "FDMASTER" Then _
                SqlStr = SqlStr & " AND DepositType = " & DepositType
            
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                            PledgeAccId = FormatField(rst(0))
    
    gDbTrans.SqlStmt = "INSERT INTO PledgeDeposit " & _
            "( LOanID,AccId,DepositType,PledgeNum) Values " & _
            "(" & NewLoanID & "," & PledgeAccId & ", " & _
            m_DepTypeArr(count) & "," & count + 1 & ")"
    
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    gDbTrans.SqlStmt = "UPDATE " & TableName & _
        " Set LoanId = " & NewLoanID & " WHERE AccID = " & PledgeAccId
    
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
Next

gDbTrans.CommitTrans
inTransaction = False

'MsgBox "New loan created successfully.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(728), vbInformation, wis_MESSAGE_TITLE

'Now Load the Same loan ito repayment tab
txtAccNo.Text = GetVal("LoanAccno")
Me.cmdSaveLoan.Enabled = False
Me.cmdLoanUpdate.Enabled = True
LoanSave = True

'Clear the all text boxes
Call LoanClear

Exit_Line:
    If inTransaction Then gDbTrans.RollBack

Exit Function

LoanSave_Error:
    If Err Then
        MsgBox "LoanSave: " & Err.Description, vbCritical
        'MsgBox GetResourceString(729) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Function

Private Sub LoanClear()

'Set m_CustReg = Nothing
On Error Resume Next
Unload m_frmPldegeDeposits
On Error GoTo Exit_Line
 
Set m_rstLoanMast = Nothing
Set m_rstLoanTrans = Nothing
Set m_frmPldegeDeposits = Nothing
Set m_rstPledge = Nothing

ReDim m_AccArr(0)
ReDim m_DepTypeArr(0)

If m_LoanID = 0 And m_CustID = 0 Then Exit Sub
m_LoanID = 0
RaiseEvent AccountChanged(0)

m_PledgeDeposits = ""
m_PledgeValue = 0
txtLoanAmt.Caption = ""
txtCustNAme = ""

cmbDeposit.ListIndex = -1
txtBalance.Caption = ""
txtRegInterest = 0
txtIntBalance = 0
txtPrincAmount.Value = 0
txtRepayDate.Text = gStrDate
txtRepayAmt.Value = 0

grd.Clear
' tabLoans.Tabs.Clear
cmdLoanInst.Enabled = False
cmdRepay.Enabled = False
cmdUndo.Enabled = False

' Clear the Loan Issue Grid Also
'Clear the Tab1
Call ClearLoanDep
Call ClearLoanDepGrd
'Clear the tab2
Dim I As Integer
For I = 0 To txtLoanIssue.count - 1
   txtLoanIssue(I).Text = ""
Next

' If a date field, display today's date.
I = GetIndex("IssueDate")
If I >= 0 Then
   txtLoanIssue(I).Text = gStrDate
End If

I = GetIndex("DepositType")
I = ExtractToken(lblLoanIssue(I).Tag, "TextIndex")
cmb(I).Locked = False
cmbDeposit.Locked = False

cmdSaveLoan.Enabled = True
cmdLoanUpdate.Enabled = False

m_CustID = 0

Exit_Line:
Err.Clear

End Sub
' Reverts the last transaction in loan transactions table.
Private Function LoanUndoLastTransaction() As Boolean
' Variables of the procedure...
Dim TransID As Long
Dim IntTransID As Long
Dim inTransaction As Boolean
Dim Lret As Long
Dim rst As ADODB.Recordset
Dim IntAmount As Currency
Dim MiscAmount As Currency
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes

' Setup the error handler.
On Error GoTo Err_line

' Check if a loan account is loaded.
If m_rstLoanMast Is Nothing Then GoTo Exit_Line

' Get the last transaction from the loan transaction table.
' Fetch the last Record.
gDbTrans.SqlStmt = "SELECT Top 1 Transid,TransType,Amount,TransDate,Balance " & _
    " FROM DepositLoanTrans WHERE loanid = " & m_LoanID & _
    " ORDER BY TransId DESC"

Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf rst("Transid") < 1 Then
    'MsgBox "There are no transaction to undo.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(661), vbInformation, wis_MESSAGE_TITLE
    LoanUndoLastTransaction = True
    GoTo Exit_Line
End If


' Check the transaction types of the records to undo.
TransID = Val(FormatField(rst("Transid")))
transType = Val(FormatField(rst("TransType")))
TransDate = rst("TransDate")
Amount = FormatField(rst("Amount"))

'Now Check For THe LastTransaction in Intreste Trans
' Fetch the last Record.
gDbTrans.SqlStmt = "SELECT Top 1 Transid,Transtype,Amount,MiscAmount, " & _
    " TransDate FROM DepositLoanIntTrans WHERE loanid = " & m_LoanID & _
    " ORDER BY TransId DESC"

Dim rstTemp As ADODB.Recordset
If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then
    If rstTemp("Transid") >= TransID Then
        If rstTemp("Transid") > TransID Then Amount = 0
        TransID = rstTemp("Transid")
        IntAmount = FormatField(rstTemp("amount"))
        MiscAmount = FormatField(rstTemp("MiscAmount"))
        TransDate = rstTemp("TransDate")
        IntTransType = rstTemp("TransType")
    End If
End If
' Begin the transaction
Dim bankClass As clsBankAcc
Dim LoanHeadID As Long
Dim IntHeadID As Long
Dim MiscHeadId As Long
Dim headName As String
Dim LoanHeadName As String
Dim LoanHeadNameEnglish As String

Set bankClass = New clsBankAcc


If m_DepositType Then
    LoanHeadName = GetDepositTypeTextEnglish(m_DepositType, LoanHeadNameEnglish)
    LoanHeadName = LoanHeadName & " " & GetResourceString(58)
    If Len(LoanHeadNameEnglish) > 0 Then LoanHeadNameEnglish = LoanHeadNameEnglish & " " & LoadResString(58)
Else
    LoanHeadName = GetResourceString(43, 58)
    LoanHeadNameEnglish = LoadResourceStringS(43, 58)
End If

LoanHeadID = bankClass.GetHeadIDCreated(LoanHeadName, LoanHeadNameEnglish, parMemDepLoan, 0, wis_DepositLoans + m_DepositType)

LoanHeadName = LoanHeadName & " " & GetResourceString(483)
If Len(LoanHeadNameEnglish) > 0 Then LoanHeadNameEnglish = LoanHeadNameEnglish & " " & LoadResString(483)
IntHeadID = bankClass.GetHeadIDCreated(LoanHeadName, LoanHeadNameEnglish, parMemDepLoan, 0, wis_DepositLoans + m_DepositType)
MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), parBankIncome, 0, wis_None)

If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & LoanHeadID & _
            " And AccID = " & m_LoanID & " ANd TransID = " & TransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
            LoanUndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If

' Begin transaction
inTransaction = gDbTrans.BeginTrans

'First Delete the interest transaction then
gDbTrans.SqlStmt = "DELETE FROM DepositLoanIntTrans WHERE transid = " & TransID _
        & " AND loanid = " & m_LoanID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line


' Delete the entry of principle amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM DepositLoanTrans WHERE TransId = " & TransID _
        & " AND loanid = " & m_LoanID
If Not gDbTrans.SQLExecute Then
    MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'UpDate The LoanMAster
gDbTrans.SqlStmt = "UPDATE Depositloanmaster SET " & _
    " LoanClosed = 0, LastIntDate = NULL  Where Loanid =  " & m_LoanID
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line


'Now Update the Ledger
If transType = wDeposit Then
    If Not bankClass.UndoCashDeposits(LoanHeadID, Amount, TransDate) Then GoTo Exit_Line
ElseIf transType = wWithdraw Then
    If Not bankClass.UndoCashWithdrawls(LoanHeadID, Amount, TransDate) Then GoTo Exit_Line
End If

'Take the ledger of Interest amount
If IntTransType = wDeposit Then
    If IntAmount > 0 Then _
        If Not bankClass.UndoCashDeposits(IntHeadID, IntAmount, TransDate) Then GoTo Exit_Line
    If MiscAmount > 0 Then _
        If Not bankClass.UndoCashDeposits(MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
ElseIf IntTransType = wWithdraw Then
    If IntAmount > 0 Then _
        If Not bankClass.UndoCashWithdrawls(IntHeadID, IntAmount, TransDate) Then GoTo Exit_Line
    If MiscAmount > 0 Then _
        If Not bankClass.UndoCashWithdrawls(MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
End If


If Amount = IntAmount + MiscAmount And transType = wContraWithdraw Then
    If Not bankClass.UndoContraTrans(LoanHeadID, IntHeadID, _
            Amount, TransDate) Then GoTo Exit_Line
ElseIf transType = wContraWithdraw Or transType = wContraDeposit Then
    MsgBox "Unable to do the transcation"
End If

' Commit the transaction.
If inTransaction Then gDbTrans.CommitTrans
inTransaction = False

'MsgBox "The last transaction is deleted.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE

LoanUndoLastTransaction = True
Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    Exit Function

Err_line:
    If Err Then
        MsgBox "LoanUndoLastTransaction: " & vbCrLf _
                '& Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(731) & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    Else
        MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function




'
Private Function LoanUpDate() As Boolean

' Setup error handler.
On Error GoTo LoanUpdate_Error:

' Declare variables for this procedure...
' ------------------------------------------
Dim txtIndex As Integer
Dim inTransaction As Boolean
Dim Lret As Long, nRet As Integer
Dim lMemberID As Long
Dim LoanAmount As Currency
Dim rst As ADODB.Recordset
Dim DepositType As Integer
'Dim TransType As wisTransactionTypes

' ------------------------------------------

' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the member id of the person availing the loan.", vbInformation
        MsgBox GetResourceString(715), vbInformation
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

'Check for loan account number detaisl.
txtIndex = GetIndex("LoanAccNo")
Dim LoanAccNum As String
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan account number.", vbInformation
        MsgBox GetResourceString(606), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    LoanAccNum = Trim$(.Text)
  '"check whether this loan account already exists
    gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE " & _
        " AccNum = " & AddQuotes(Trim$(.Text), True) & _
        " AND DepositType = " & GetDepositType(GetVal("DepositType")) & _
        " AND LoanId <> " & m_LoanID
        
    If gDbTrans.Fetch(rst, adOpenForwardOnly) Then
        'msgbox "This loan account already exists " & "Please specify loan account n0
        MsgBox GetResourceString(58, 545) & _
            vbCrLf & GetResourceString(606), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
    End If
End With

' Pledge item validation.
txtIndex = GetIndex("PledgeDeposit")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the value of pledged items.", _
                    vbInformation , wis_MESSAGE_TITLE
        MsgBox GetResourceString(719), _
                    vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

Dim PledgeAmount As Currency
txtIndex = GetIndex("PledgeValue")
With txtLoanIssue(txtIndex)
    ' Ensure the value of pledge items is valid.
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for pledge item.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(720), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    PledgeAmount = Val(.Text)
End With

' Loan Amount.
txtIndex = GetIndex("LoanAmount")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for loan amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(506), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
    End If
LoanAmount = Val(.Text)
End With

'Loan issue date.
Dim IssueDate As Date
txtIndex = GetIndex("IssueDate")
With txtLoanIssue(txtIndex)
    If Not DateValidate(.Text, "/", True) Then
        'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
    End If
    IssueDate = GetSysFormatDate(.Text)
End With

'Loan due date.
Dim DueDate As Date
txtIndex = GetIndex("DueDate")
With txtLoanIssue(txtIndex)
    If Not DateValidate(.Text, "/", True) Then
        'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
    End If
    DueDate = GetSysFormatDate(.Text)
End With
If DateDiff("d", IssueDate, DueDate) <= 0 Then
    'MsgBox "Due date is earlier than the issue date", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(580), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
End If

'Loan Interest Rate
txtIndex = GetIndex("InterestRate")
If Not IsNumeric(txtLoanIssue(txtIndex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    Exit Function
End If

'loan Penal Interest Rate
txtIndex = GetIndex("PenalInterestRate")
If Not IsNumeric(txtLoanIssue(txtIndex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    Exit Function
End If

'Get the IntrestBalance form The transaction
Dim IntbalAmount As Currency
txtIndex = GetIndex("InterestBalance")
With txtLoanIssue(txtIndex)
'    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
'            'MsgBox "Invalid value for loan amount.", _
'                    vbInformation, wis_MESSAGE_TITLE
'            MsgBox GetResourceString(506), _
'                    vbInformation, wis_MESSAGE_TITLE
'            ActivateTextBox txtLoanIssue(txtindex)
'            GoTo Exit_Line
'    End If
IntbalAmount = Val(.Text)
End With

'Get the deposit selected.

DepositType = GetDepositType(GetVal("DepositType"))
'Here cHeck Whether DepositType Has Changed
Debug.Print "Kannada"
If m_DepositType <> DepositType Then _
    If MsgBox("You have changed the depositType" & vbCrLf & _
        "Do you want to continue?", vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function


'Now check Whether Pldeged DepositHas Changed
Dim DepChanged As Boolean
Dim count As Integer
nRet = UBound(m_AccArr)

m_rstPledge.MoveFirst
Do While Not m_rstPledge.EOF
    For count = 0 To nRet
        If m_rstPledge("AccId") = m_AccArr(count) And _
            m_rstPledge("DepositType") = m_DepTypeArr(count) Then Exit For
    Next
    If count > nRet Then DepChanged = True: Exit Do
    m_rstPledge.MoveNext
Loop

Debug.Print "Kannada"
If DepChanged Then
    If MsgBox("You have changed the pledged deposits" & vbCrLf & _
        "Do you want to continue?", vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If

' Begin transaction.
gDbTrans.BeginTrans
inTransaction = True
Dim SqlStr As String

' Put an entry into LoanMaster table.
SqlStr = "UPDATE DepositLoanMaster Set " & _
        " AccNum = " & AddQuotes(LoanAccNum, True) & " ," & _
        " DepositType = " & m_DepositType & ", " & _
        " LoanIssueDate = #" & IssueDate & "#, " & _
        " PledgeValue = " & Val(GetVal("PledgeValue")) & ", " & _
        " PledgeDescription = " & AddQuotes(GetVal("PledgeDeposit"), True) & ", " & _
        " LoanAmount = " & LoanAmount & ", " & _
        " LoanDueDate = #" & DueDate & "#, " & _
        " InterestRate = " & Val(GetVal("InterestRate")) & "," & _
        " PenalInterestRate = " & Val(GetVal("PenalInterestRate")) & "," & _
        " Remarks = " & AddQuotes(GetVal("Remarks"), True) & _
        " Where LoanID = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

'Now Insert the Pledged Deposits
If DepChanged Then
    SqlStr = "Delete * From PledgeDeposit Where LoanID= " & m_LoanID
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    For count = 0 To nRet
        If m_AccArr(count) = 0 Then Exit For
        SqlStr = "Insert INTO PledgeDeposit " & _
            "(LoanId,accId,DepositType,PledgeNum) VAlues " & _
            "(" & m_LoanID & "," & m_AccArr(count) & "," & _
            m_DepTypeArr(count) & "," & count + 1 & ")"
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    Next
End If

'Now update into the IntrestTable ?What to update

gDbTrans.SqlStmt = "UPDATE DepositLoanintTrans Set " & _
        " Balance =" & GetVal("InterestBalance") & _
        " Where LoanID = " & m_LoanID & _
        " And TransID = (Select Max(TransID) From " & _
        " DepositLoanIntTrans Where LoanID = " & m_LoanID & ")"

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

gDbTrans.CommitTrans
MsgBox GetResourceString(707), vbInformation, wis_MESSAGE_TITLE
'MsgBox "Loan updated  successfully.", vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    'If Val(txtLoanIssue(GetIndex("InstalmentNames")).Text) > 1 Then Unload frmLoanName
    If inTransaction Then gDbTrans.RollBack
    Exit Function

LoanUpdate_Error:
    If Err Then
        MsgBox "LoanSave: " & Err.Description, vbCritical
        'MsgBox GetResourceString(728) & Err.Description, vbCritical
    End If
Resume
    GoTo Exit_Line
End Function


' Returns the memberID of Guarantor1, selected by the user.
Private Function PropGuarantorID(GuarantorNum As Integer) As Long
Dim strTmp As String
Dim txtIndex As Integer
txtIndex = GetIndex("Guarantor" & GuarantorNum)
If txtIndex >= 0 Then
    strTmp = ExtractToken(txtLoanIssue(txtIndex).Tag, "GuarantorID")
    PropGuarantorID = Val(strTmp)
End If

End Function
'
Private Sub PropInitializeForm()

' Set the proerties for tab strip control.
TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True

' Load the properties for loan issue panel.
LoadLoanIssueProp

' Remove all tabs of tabloans.
'tabLoans.Tabs.Clear


End Sub
' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetVal(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
For I = 0 To txtLoanIssue.count - 1
    strTxt = ExtractToken(lblLoanIssue(I).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        GetVal = txtLoanIssue(I).Text
        Exit For
    End If
Next
End Function


'
Private Function GetIndex(strDataSrc As String) As Integer
GetIndex = -1
Dim strTmp As String
Dim I As Integer
For I = 0 To lblLoanIssue.count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(lblLoanIssue(I).Tag, "DataSource")
    If StrComp(strDataSrc, strTmp, vbTextCompare) = 0 Then
        GetIndex = I
        Exit For
    End If
Next

End Function
'
Private Sub DisplayResults(MoveDirection As Integer)
#If RAVI_COMMENTED Then
'On Error Resume Next
'Uses the record set to display results
    If m_rstSearchResults Is Nothing Then
        Exit Sub
    End If
    
    If MoveDirection = 1 Then
        m_rstSearchResults.MoveNext
    ElseIf MoveDirection = -1 Then
        m_rstSearchResults.MovePrevious
    End If

    
    Dim customerID As Long
    customerID = FormatField(m_rstSearchResults("SBMaster.CustomerID"))
    If Not m_AccHolder.LoadCustomerInfo(customerID) Then
        Exit Sub
    End If
    txtNewAccNo.Text = m_rstSearchResults("AccID")
    m_AccNo = FormatField(m_rstSearchResults("AccID"))
    txtName.Text = m_AccHolder.FullName
    
    cmdDetails.Enabled = True

    'txtNewDate.Text = m_rstsearchresults("CreateDate")
    txtNewDate.Text = FormatField(m_rstSearchResults("CreateDate"))
    txtLedger.Text = FormatField(m_rstSearchResults("LedgerNo"))
    txtFolio.Text = FormatField(m_rstSearchResults("FolioNo"))
    
    Dim I As Integer
    Dim Arr() As String
    Call GetStringArray(FormatField(m_rstSearchResults("Nominee")), Arr, ";")
    txtNominee.Text = "": txtAge.Text = "": cmbRelation.ListIndex = -1
    For I = 0 To UBound(Arr)
        Select Case I
            Case 0: txtNominee.Text = Arr(I)
            Case 1: txtAge.Text = Arr(I)
            Case 2: cmbRelation.ListIndex = Val(Arr(I))
        End Select
    Next I
    
    Call GetStringArray(FormatField(m_rstSearchResults("JointHolder")), Arr, ";")
    cmbHolders.Clear
    For I = 0 To UBound(Arr)
        cmbHolders.AddItem Arr(I)
    Next I
    
    txtIntroAccNo.Text = FormatField(m_rstSearchResults("Introduced"))
    If Val(txtIntroAccNo.Text) = 0 Then
        txtIntroAccNo.Text = ""
    End If
    
    'Load the name now
    Call m_AccHolder.LoadCustomerInfo(FormatField(m_rstSearchResults("SBMaster.CustomerID")))
    txtIntroName.Text = m_AccHolder.FullName
    
    m_rstSearchResults.MovePrevious
    If m_rstSearchResults.BOF Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If
    If m_rstSearchResults.BOF Then
        m_rstSearchResults.MoveFirst
    Else
        m_rstSearchResults.MoveNext
    End If
    
    m_rstSearchResults.MoveNext
    If m_rstSearchResults.EOF Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If m_rstSearchResults.EOF Then
        m_rstSearchResults.MoveLast
    Else
        m_rstSearchResults.MovePrevious
    End If
    
#End If
End Sub

'
Private Function LoadLoanIssueProp() As Boolean

'
' Read the data from Loans.ini and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\DepLn_" & gLangOffSet & ".PRP"

If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet <> wis_NoLangOffset Then
        PropFile = App.Path & "\DepLnkan.PRP"
    Else
        PropFile = App.Path & "\DepLoan.PRP"
    End If
End If
If Dir(PropFile, vbNormal) = "" Then

    'MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox GetResourceString(602) _
            & PropFile & "' !", vbExclamation
    Exit Function
End If

' Declare required variables...
Dim strTmp As String
Dim strPropType As String
Dim FirstImgCtl As Boolean
Dim FirstControl As Boolean
Dim I As Integer, CtlIndex As Integer
Dim strRet As String, imgCtlIndex As Integer
FirstControl = True
FirstImgCtl = True
Dim strTag As String

' Read all the prompts and load accordingly...
Do
    ' Read a line.
    strTag = Trim(ReadFromIniFile("LoanIssue", _
                "Prop" & I + 1, PropFile))
    If strTag = "" Then Exit Do

    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load lblLoanIssue(lblLoanIssue.count)
        Load txtLoanIssue(txtLoanIssue.count)
    End If
    CtlIndex = lblLoanIssue.count - 1
'    Debug.Assert I <> 15
    ' Get the property type.
    strPropType = Trim$(ExtractToken(strTag, "PropType"))
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With lblLoanIssue(CtlIndex)
                .FontBold = True
                '.Text = ""
                .Caption = ""
            End With
            txtLoanIssue(CtlIndex).Enabled = False
        Case "EDITABLE"
            ' Add 4 spaces for indentation purposes.
            With lblLoanIssue(CtlIndex)
                .Caption = Space(2)
                .FontBold = False
            End With
            txtLoanIssue(CtlIndex).Enabled = True
        Case Else
            'MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            MsgBox GetResourceString(603) _
                    & "in Property file!", vbCritical
            Exit Function
    End Select
    ' Set the PROPERTIES for controls.
'    Debug.Assert I <> 12
    With lblLoanIssue(CtlIndex)
        strRet = PutToken(strTag, "Visible", "True")
        .Tag = strRet
        .Caption = .Caption & ExtractToken(.Tag, "Prompt")
        If CtlIndex = 0 Then
            .Top = 0
        Else
            .Top = lblLoanIssue(CtlIndex - 1).Top _
                + lblLoanIssue(CtlIndex - 1).Height + CTL_MARGIN
        End If
        .Left = 0
        .Visible = True
    End With
    With txtLoanIssue(CtlIndex)
        .Top = lblLoanIssue(CtlIndex).Top
        .Left = lblLoanIssue(CtlIndex).Left + _
            lblLoanIssue(CtlIndex).Width + CTL_MARGIN
        .Visible = True
        ' Check the LockEdit property.
        strRet = ExtractToken(strTag, "LockEdit")
        If StrComp(strRet, "True", vbTextCompare) = 0 Then
            .Locked = True
        Else
            .Locked = False
        End If
    End With

    ' Get the display type. If its a List or Browse,
    ' then load a combo or a cmd button.
    Dim CmdLoaded As Boolean
    Dim ListLoaded As Boolean
    strPropType = ExtractToken(strTag, "DisplayType")
    Select Case UCase$(strPropType)
        Case "LIST"
            'Load a combo.
            If Not ListLoaded Then
                ListLoaded = True
            Else
                Load cmb(cmb.count)
            End If
            ' Set the alignment.
            With cmb(cmb.count - 1)
                '.Index = i
                .Left = txtLoanIssue(I).Left
                .Top = txtLoanIssue(I).Top
                .Width = txtLoanIssue(I).Width
                ' Set it's tab order.
                .TabIndex = txtLoanIssue(I).TabIndex + 1
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                lblLoanIssue(I).Tag = PutToken(lblLoanIssue(I).Tag, _
                        "TextIndex", CStr(cmb.count - 1))
                'txtData(i).Visible = False
                ' If the list data is given, load it.
                Dim List() As String, j As Integer
                Dim strListData As String
                strListData = ExtractToken(strTag, "ListData")
                If strListData <> "" Then
                    ' Break up the data into array elements.
                    GetStringArray strListData, List(), ","
                    cmb(cmb.count - 1).Clear
                    For j = 0 To UBound(List)
                        cmb(cmb.count - 1).AddItem List(j)
                    Next
                End If
            End With

        Case "BROWSE"
            'Load a command button.
            If Not CmdLoaded Then CmdLoaded = True _
                    Else Load cmd(cmd.count)
            With cmd(cmd.count - 1)
                '.Index = i
                .Width = txtLoanIssue(I).Height
                .Height = .Width
                .Left = txtLoanIssue(I).Left + txtLoanIssue(I).Width - .Width
                .Top = txtLoanIssue(I).Top
                .TabIndex = txtLoanIssue(I).TabIndex + 1
                .ZOrder 0
                
                ' If width and caption are mentioned for
                ' the command button, apply them.
                strTmp = ExtractToken(strTag, "ButtonWidth")
                If strTmp <> "" Then
                    If Val(strTmp) > txtLoanIssue(I).Width / 2 Then
                        ' Restrict the width to half the textbox width.
                        .Width = txtLoanIssue(I).Width / 2
                    ElseIf Val(strTmp) < .Width Then
                        .Width = txtLoanIssue(I).Height
                    Else
                        .Width = Val(strTmp)
                    End If
                End If
                strTmp = ExtractToken(strTag, "ButtonCaption")
                If strTmp <> "" Then
                    .Caption = strTmp
                Else
                    .Caption = "..."
                End If
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                lblLoanIssue(I).Tag = PutToken(lblLoanIssue(I).Tag, _
                        "TextIndex", CStr(cmd.count - 1))
            End With
    End Select

    ' Increment the loop count.
    I = I + 1
Loop
Call ArrangePropSheet

' Display today's date, for date field.
I = GetIndex("IssueDate")
If I >= 0 Then
    txtLoanIssue(I).Text = gStrDate
End If

Dim cmbIndex As Integer

    ' Find out the textbox bound to DepositType.
    I = GetIndex("DepositType")
    ' Get the combobox index for this text.
    cmbIndex = ExtractToken(lblLoanIssue(I).Tag, "TextIndex")
    Call LoadDepositTypes(cmb(cmbIndex))
    
End Function
'
Private Sub PropScrollLoanIssueWindow(Ctl As Control)

If picLoanIssueSlider.Top + Ctl.Top + Ctl.Height > picLoanIssueViewPort.ScaleHeight Then
    ' The control is below the viewport.
    Do While picLoanIssueSlider.Top + Ctl.Top + Ctl.Height > _
                    picLoanIssueViewPort.ScaleHeight
        ' scroll down by one row.
        With vscLoanIssue
            If .Value + .SmallChange <= .Max Then
                    .Value = .Value + .SmallChange
            Else
                    .Value = .Max
            End If
        End With
    Loop

ElseIf picLoanIssueSlider.Top + Ctl.Top < 0 Then
    ' The control is above the viewport.
    ' Keep scrolling until it is in viewport.
    Do While picLoanIssueSlider.Top + Ctl.Top < 0
        With vscLoanIssue
            If .Value - .SmallChange >= .Min Then
                .Value = .Value - .SmallChange
            Else
                .Value = .Min
            End If
        End With
    Loop
End If

End Sub
Private Sub PropSetIssueDescription(Ctl As Control)
' Extract the description title.
lblLoanIssueHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblLoanIssueDesc.Caption = ExtractToken(Ctl.Tag, "Description")

End Sub
Private Function Validate() As Boolean
' Setup error handler.
On Error GoTo Err_line

' Declare variables for this procedure...
' ------------------------------------------
Dim txtIndex As Integer
Dim Balance As Currency
Dim LoanAmount As Currency
Dim PledgeAmount As Currency
Dim IssueDate As Date
Dim DueDate As Date
Dim rst As ADODB.Recordset
Dim DepositType As Integer

' ------------------------------------------

DepositType = GetDepositType(GetVal("DepositType"))
DepositType = DepositType Mod 100

txtIndex = GetIndex("DepositType")
With txtLoanIssue(txtIndex)
    If Trim$(.Text) = "" Then
        'MsgBox "Please select loan type", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(718), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(GetIndex("DepositType"))
        Exit Function
    End If
End With

' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the member id of the person availing the loan.", vbInformation
        MsgBox GetResourceString(715), vbInformation
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

'Check for loan account number detaisl.
txtIndex = GetIndex("LoanAccNo")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan account number.", vbInformation
        MsgBox GetResourceString(606), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
  
  '"check whether this loan account already exists
    gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE " & _
                    " AccNum = " & AddQuotes(Trim$(.Text), True) & _
                    " AND DepositType = " & DepositType
        
    If gDbTrans.Fetch(rst, adOpenForwardOnly) Then
        'msgbox "This loan account already exists " & "Please specify loan account n0
        MsgBox GetResourceString(58, 545) & _
            vbCrLf & GetResourceString(606), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        Exit Function
    End If
End With

' Pledge item validation.
txtIndex = GetIndex("PledgeDeposit")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the value of pledged items.", _
                    vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(719), _
                    vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

txtIndex = GetIndex("PledgeValue")
With txtLoanIssue(txtIndex)
    ' Ensure the value of pledge items is valid.
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for pledge item.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(720), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    PledgeAmount = Val(.Text)
End With

'Loan Amount.
txtIndex = GetIndex("LoanAmount")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    If LoanAmount > PledgeAmount Then
        'MsgBox "Amount do not tally." & Do Yo want to continue, vbInformation , wis_MESSAGE_TITLE
        If MsgBox(GetResourceString(797) & vbCrLf & _
            GetResourceString(541), _
            vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
    End If
End With

'Loan issue date.
txtIndex = GetIndex("IssueDate")
With txtLoanIssue(txtIndex)
    If Not DateValidate(.Text, "/", True) Then
        'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
    End If
    IssueDate = GetSysFormatDate(.Text)
End With

'Loan due date.
txtIndex = GetIndex("DueDate")
With txtLoanIssue(txtIndex)
    If Not DateValidate(.Text, "/", True) Then
        'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        Exit Function
    End If
    DueDate = GetSysFormatDate(.Text)
End With
If DateDiff("d", IssueDate, DueDate) <= 0 Then
    'MsgBox "Due date is earlier than the issue date", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(580), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
End If

'Loan Interest Rate
txtIndex = GetIndex("InterestRate")
If Not IsNumeric(txtLoanIssue(txtIndex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    Exit Function
End If

'loan Penal Interest Rate
txtIndex = GetIndex("PenalInterestRate")
If Not IsNumeric(txtLoanIssue(txtIndex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    Exit Function
End If


Validate = True

Exit_Line:
Err_line:

End Function

' Returns the number of items that are visible for a control array.
' Looks in the control's tag for visible property, rather than
' depend upon the control's visible property for some obvious reasons.
Private Function VisibleCountLoanIssue() As Integer
On Error GoTo Err_line
Dim I As Integer
Dim strVisible As String

For I = 0 To lblLoanIssue.count - 1
    strVisible = ExtractToken(lblLoanIssue(I).Tag, "Visible")
    If StrComp(strVisible, "True", vbTextCompare) = 0 Then
        VisibleCountLoanIssue = VisibleCountLoanIssue + 1
    End If
Next

Err_line:
End Function


Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String

txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" Then txtLoanIssue(Val(txtIndex)).Text = cmb(Index).Text

End Sub

Private Sub cmdAddNote_Click()
If m_Notes.ModuleID = 0 Then
    Exit Sub
End If

Call m_Notes.Show
Call m_Notes.DisplayNote(rtfNote)
End Sub

Private Sub cmdAdvance_Click()
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    m_clsRepOption.ShowDialog
    
End Sub

Private Sub cmdEndDate_Click()
With Calendar
    .Left = Me.Left + fraReports.Left + fraOrder.Left + cmdEndDate.Left - .Width
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdEndDate.Top + 300
    .selDate = txtEndDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtEndDate.Text = Calendar.selDate
End With

End Sub

'
Private Sub cmdLoad_Click()
Me.MousePointer = vbDefault
Dim ret As Integer
Dim rst As ADODB.Recordset
Dim AccNum As String
Dim RecFetch As Boolean
Dim SqlStr As String
'
'If cmbDeposit.ListIndex = -1 Then Exit Sub
txtRepayDate.Tag = txtRepayDate
Me.MousePointer = vbHourglass
AccNum = Trim$(txtAccNo)
'Get the Loan Id
SqlStr = "SELECT * FROM DepositLoanMaster " & _
    " WHERE AccNum = " & AddQuotes(Trim$(txtAccNo), True)
    
'Check for the depositType which is existing
'If cmbDeposit.ListIndex >= 0 Then _
'    SqlStr = SqlStr & " AND Deposittype = " & cmbDeposit.ItemData(cmbDeposit.ListIndex)
gDbTrans.SqlStmt = SqlStr

ret = gDbTrans.Fetch(rst, adOpenStatic)
If ret < 1 Then  'RecFetch = True 'GoTo Exit_Line
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
     Me.MousePointer = vbNormal
    GoTo Exit_Line
End If

If ret > 1 Then
    Dim Deptype As wis_DepositType
    If cmbDeposit.ListIndex < 0 Then
        MsgBox "Please select the Deposit type", vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    Deptype = cmbDeposit.ItemData(cmbDeposit.ListIndex)
    gDbTrans.SqlStmt = "SELECT * FROM DepositLOanMaster " & _
        " WHERE AccNum = " & AddQuotes(Trim$(txtAccNo), True) & _
        " And DepositType = " & Deptype
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo Exit_Line
End If

If Not LoanLoad(FormatField(rst("LoanID"))) Then GoTo Exit_Line

'Get the intrest balance to the custreg
txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)

Exit_Line:
    Me.MousePointer = vbDefault
    Exit Sub
Err_line:
    Call LoanClear
    If Err Then MsgBox "Error In LoanLoad :" & Err.Number & vbCrLf & _
            Err.Description, vbInformation, wis_MESSAGE_TITLE
    Err.Clear
End Sub

'
Private Sub cmd_Click(Index As Integer)
'Me.MousePointer = vbHourglass
' Variables of this routine...
Dim txtIndex As String
Dim strField As String
Dim Lret As Long
Dim NameStr As String
Dim SqlStr As String
Dim custId As Long
Dim rst As ADODB.Recordset
Dim custReg As clsCustReg


' Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
strField = ExtractToken(lblLoanIssue(Val(txtIndex)).Tag, "DataSource")

Select Case UCase$(strField)
    Case "MEMBERID"
        If custReg Is Nothing Then Set custReg = New clsCustReg
        strField = InputBox("Enter the name whom you are searching", wis_MESSAGE_TITLE)
        'Added
       
        With gDbTrans
            .SqlStmt = "SELECT CustomerId, " & _
                " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name, " & _
                " Profession FROM NameTab "
            If Trim(strField) <> "" Then
                .SqlStmt = .SqlStmt & " Where (FirstName like '" & strField & "%' " & _
                    " Or MiddleName like '" & strField & "%' " & _
                    " Or LastName like '" & strField & "%')"
                .SqlStmt = .SqlStmt & " Order by IsciName"
            Else
                .SqlStmt = .SqlStmt & " Order by FirstName"
            End If
            Lret = .Fetch(rst, adOpenStatic)
            If Lret <= 0 Then
                'MsgBox "No data available!", vbExclamation
                MsgBox GetResourceString(278), vbExclamation
                Exit Sub
            End If
        End With
        
        ' Create a report dialog.
        If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp
        
        With m_LookUp
            .m_SelItem = ""
            ' Fill the data to report dialog.
            If Not FillView(.lvwReport, rst, True) Then
                'MsgBox "Error filling the customer details.", vbCritical
                MsgBox "Error filling the customer details.", vbCritical
                Exit Sub
            End If
            Screen.MousePointer = vbDefault
            ' Display the dialog.
            .Show vbModal
        End With
        If m_CustID <> Val(m_SelItem) Then Set m_frmPldegeDeposits = Nothing
        m_CustID = Val(m_SelItem)
        txtLoanIssue(GetIndex("MemberName")).Text = custReg.CustomerName(Val(m_SelItem))
        txtLoanIssue(GetIndex("MemberID")).Text = Val(m_SelItem)
        
        
    Case "MEMBERNAME"
        If custReg Is Nothing Then Set custReg = New clsCustReg
        custReg.ShowDialog
        txtLoanIssue(GetIndex("MemberName")).Text = custReg.FullName
        txtLoanIssue(GetIndex("MemberID")).Text = custReg.customerID
    
    Case "DUEDATE"
        With Calendar
            .Left = txtLoanIssue(txtIndex).Left + Me.Left _
                    + picLoanIssueViewPort.Left + fraLoanAccounts.Left + 50
            .Top = Me.Top + txtLoanIssue(txtIndex).Top _
                    + picLoanIssueSlider.Top + picLoanIssueViewPort.Top _
                    + fraLoanAccounts.Top + 300
            .Width = txtLoanIssue(txtIndex).Width
            .Height = .Width
            .selDate = txtLoanIssue(txtIndex).Text
            .Show vbModal, Me
            If .selDate <> "" Then txtLoanIssue(txtIndex).Text = .selDate
        End With

    Case "ISSUEDATE"
        With Calendar
            .Left = txtLoanIssue(txtIndex).Left + Me.Left _
                    + picLoanIssueViewPort.Left + fraLoanAccounts.Left + 50
            .Top = Me.Top + txtLoanIssue(txtIndex).Top _
                    + picLoanIssueSlider.Top + picLoanIssueViewPort.Top _
                    + fraLoanAccounts.Top + 300
            .Width = txtLoanIssue(txtIndex).Width
            .Height = .Width
            .selDate = txtLoanIssue(txtIndex).Text
            .Show vbModal, Me
            txtLoanIssue(txtIndex).Text = .selDate
        End With
    Case "PLEDGEDEPOSIT"
        'If m_CustReg Is Nothing Then Exit Sub
        If m_CustID = 0 Then Exit Sub
        If m_frmPldegeDeposits Is Nothing Then
            Set m_frmPldegeDeposits = New frmDepSelect
            Load m_frmPldegeDeposits
            m_frmPldegeDeposits.LoadDeposits (m_CustID)
        End If
        Dim ArrCount As Integer
        Dim LstCount As Integer
        LstCount = 1
        For ArrCount = 0 To UBound(m_AccArr)
          With m_frmPldegeDeposits.lstDeposits
            If LstCount < .ListItems.count Then LstCount = 1
            Do '
                If LstCount > .ListItems.count Then Exit Do
                If .ListItems(LstCount).SubItems(1) = m_AccArr(ArrCount) Then
                    .ListItems(LstCount).Selected = True
                    Exit Do
                End If
                LstCount = LstCount + 1
            Loop
          End With
        Next
        With m_frmPldegeDeposits
            .Left = Me.Left + fraLoanAccounts.Left + txtLoanIssue(txtIndex).Width _
                    + txtLoanIssue(txtIndex).Left + picLoanIssueViewPort.Left - .Width
            .Top = Me.Top + txtLoanIssue(txtIndex).Top _
                    + picLoanIssueSlider.Top + picLoanIssueViewPort.Top _
                    + fraLoanAccounts.Top - .Height / 2
            .Show vbModal
        End With
        
        txtIndex = GetIndex("PledgeValue")
        txtLoanIssue(txtIndex) = m_PledgeValue
        txtIndex = GetIndex("PledgeDeposit")
        txtLoanIssue(txtIndex) = m_PledgeDeposits
        txtIndex = GetIndex("NoOfDep")
        txtLoanIssue(txtIndex) = UBound(m_AccArr) + 1
        
End Select

ExitLine:
Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoanInst_Click()
frmDepLoanInst.LoanID = m_LoanID

frmDepLoanInst.Show 1
Call LoanLoad(m_LoanID)
End Sub

Private Sub cmdLoanIssueClear_Click()
Call LoanClear

End Sub

Private Sub ClearLoanDep()

'cmbDeposit.Clear

'txtAccNo.Text = ""
txtRegInterest = 0
txtIntBalance = 0
txtLoanAmt.Caption = ""
txtMisc = 0
txtBalance.Caption = ""
txtPrincAmount = 0

 

End Sub

Private Sub ClearLoanDepGrd()
  Dim Offset As Single
    Offset = 100
    With grd
        .Clear: .Cols = 6:  .Rows = 21: .FixedRows = 1: .FixedCols = 0
        .Row = 0
        .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = .Width / .Cols - Offset '"Date"
        .Col = 1: .Text = GetResourceString(81): .ColWidth(1) = .Width / .Cols - Offset '"Sold"
        .Col = 2: .Text = GetResourceString(82): .ColWidth(2) = .Width / .Cols - Offset '"Returned"
        .Col = 3: .Text = GetResourceString(47): .ColWidth(3) = .Width / .Cols - Offset ''"Net Worth"
        .Col = 5: .Text = GetResourceString(42) '"Balance"
    End With
End Sub
Private Sub cmdNextTrans_Click()
'm_rstLoanTrans.MoveFirst
'm_rstLoanTrans.Find " TransID = " & m_TransID
Call LoanLoadGrid
End Sub

Private Sub cmdOk_Click()

Dim Cancel As Boolean

'Ask the user wen closing application

Call LoanClear
Unload Me

End Sub

Private Sub cmdPrevTrans_Click()
    m_TransID = m_TransID - 10
    m_rstLoanTrans.MoveFirst
    m_rstLoanTrans.Find "TransID >= " & m_TransID
    Call LoanLoadGrid
End Sub
Private Sub cmdPrint_Click()
    If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    
    m_frmPrintTrans.Show vbModal
      
End Sub
Private Sub cmdRepay_Click()

If Not LoanRepay Then Exit Sub

Call LoanLoad(m_LoanID)

txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)

End Sub

Private Sub cmdRepayDate_Click()
With Calendar
    .Left = Me.Left + Me.fraRepayments.Left + Me.cmdRepayDate.Left
    .Top = Me.Top + Me.fraRepayments.Top + Me.cmdRepayDate.Top
    .selDate = txtRepayDate.Text
    .Show vbModal
    If .selDate <> "" Then txtRepayDate.Text = .selDate
End With

End Sub

Private Sub cmdSaveLoan_Click()
If Not Validate Then Exit Sub
If Not LoanSave Then Exit Sub

Call LoanLoad(m_LoanID)

End Sub


Private Sub cmdStDate_Click()
With Calendar
    .selDate = gStrDate
    .Left = Me.Left + fraReports.Left + fraOrder.Left + cmdStDate.Left
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdStDate.Top + 300
    .selDate = txtStartDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtStartDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdUndo_Click()

Dim nRet As Integer

If cmdUndo.Caption = GetResourceString(5) Then  '"Undo
   'nRet = MsgBox("This will undo the last transaction for this loan account. " _
        & vbCrLf & "Do you want to continue ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(583), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   Me.MousePointer = vbHourglass
   
   Call LoanUndoLastTransaction

ElseIf cmdUndo.Caption = GetResourceString(313) Then  '"Reopen
   'nRet = MsgBox("Are you sure to reopen this loan ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(538), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   Me.MousePointer = vbHourglass
   gDbTrans.SqlStmt = "UpDate LoanMaster set LoanClosed = 1, InterestBalance = 0 where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to reopen the loan"
      MsgBox GetResourceString(536), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.CommitTrans
   'MsgBox "Account reopened succefully"
   MsgBox GetResourceString(522), vbInformation, wis_MESSAGE_TITLE
   
ElseIf cmdUndo.Caption = GetResourceString(14) Then  '"Delete
   'nRet = MsgBox("Are you sure to delete this accoutnt ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(539), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   If Not LoanUndoLastTransaction Then
         'MsgBox "Unable to delete the loan"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   gDbTrans.SqlStmt = "Delete * From DepositLoanTrans Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete the loan"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   
   'Delete the Details In Pledge Deposit
   gDbTrans.SqlStmt = "Delete * From PledgeDeposit Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   
   'Delete the Details Of Pldge Deposit
   gDbTrans.SqlStmt = "Delete * From DepositLoanMaster Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   
   gDbTrans.CommitTrans
   'MsgBox "Account Deleted succefully"
   MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE
   Me.MousePointer = vbDefault
   'Exit Sub
End If

' Reload the loan details.
Dim LoanID As Long
LoanID = m_LoanID
Call LoanClear
Call LoanLoad(LoanID)


Me.MousePointer = vbDefault
End Sub
'
Private Sub cmdLoanUpdate_Click()
    Call LoanUpDate
End Sub
'
Private Sub cmdView_Click()

Dim fromDate As String
Dim toDate As String

' Validate the user input.
' Check for starting date.
With txtStartDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        'MsgBox "Enter a valid starting date.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtStartDate
        Exit Sub
    End If
    If .Enabled Then fromDate = .Text
End With

' Check for ending date.
With txtEndDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        'MsgBox "Enter a valid ending date.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEndDate
        Exit Sub
    End If
    If .Enabled Then toDate = .Text
End With

Dim ReportOrder As wis_ReportOrder
Dim ReportType As wis_DepLoanReports
Dim DepositType As Integer

ReportOrder = IIf(optName.Value, wisByName, wisByAccountNo)

If optReport(0).Value Then ReportType = repDepLnBalance
If optReport(1).Value Then ReportType = repDepLnDetail
If optReport(2).Value Then ReportType = repDepSubDayBook 'repDepLnTransaction
If optReport(3).Value Then ReportType = repDepLnOverDue
If optReport(4).Value Then ReportType = repDepLnGenLedger
If optReport(5).Value Then ReportType = repDepLnCashBook
If optReport(6).Value Then ReportType = repDepLnMonthlyBalance

If cmbRepDeposit.ListIndex >= 0 Then DepositType = cmbRepDeposit.ItemData(cmbRepDeposit.ListIndex)

    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    

gCancel = 0
RaiseEvent ShowReport(ReportType, ReportOrder, DepositType, _
        fromDate, toDate, m_clsRepOption)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
If KeyCode <> vbKeyTab Then Exit Sub
If Not CtrlDown Then Exit Sub

Dim I As Byte
With TabStrip
    I = .SelectedItem.Index
    If Shift = 2 Then
        I = I + 1
        If I > .Tabs.count Then I = 1
    Else
        I = I - 1
        If I = 0 Then I = .Tabs.count
    End If
    .Tabs(I).Selected = True
End With

End Sub

Private Sub Form_Load()
      Screen.MousePointer = vbHourglass
      'set icon for the form caption
      Me.Icon = LoadResPicture(161, vbResIcon)
      cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
      Call SetKannadaCaption
      
      'Centre the form
          Me.Move (Screen.Width - Me.Width) \ 2, _
                  (Screen.Height - Me.Height) \ 2
      
      PropInitializeForm
      
      Me.txtRepayDate.Text = gStrDate
      ReDim m_LoanAmount(0)
      
      
      Call LoadDepositTypes(Me.cmbDeposit)
      
      'Clear the Loan Detial
      m_LoanID = 1
      Call LoanClear
      
      Call LoadDepositTypes(cmbRepDeposit)
      optReport(0).Value = True
      'Individual Report types for the Loans
      Screen.MousePointer = vbDefault
      
      txtEndDate = gStrDate
      
    cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
    cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Now Set the module level Reference to nothing before unloading
Set m_rstLoanMast = Nothing
Set m_rstLoanTrans = Nothing

' Report form.
If Not m_LookUp Is Nothing Then
    Unload m_LookUp
    Set m_LookUp = Nothing
End If
'If Not m_CustReg Is Nothing Then Set m_CustReg = Nothing
gWindowHandle = 0
RaiseEvent WindowClosed
End Sub

Private Sub lblLoanIssue_DblClick(Index As Integer)
On Error Resume Next
txtLoanIssue(Index).SetFocus

End Sub


Private Sub m_frmPldegeDeposits_OKClicked(AccList() As String, _
                            DepList() As Integer, TotalBalance As Currency)

ReDim m_AccArr(0)
ReDim m_DepTypeArr(0)
'Now Assign the deposit Type & AccOunt Id
m_AccArr = AccList
m_DepTypeArr = DepList
'Now Get the List Of Accounts & thier Account NOs
Dim count As Integer
Dim MaxCount As Integer
Dim loopCount As Integer

MaxCount = UBound(m_AccArr)
m_PledgeDeposits = ""

If MaxCount = 0 And m_AccArr(0) = "" Then Exit Sub
loopCount = 1
For count = 0 To MaxCount
    Do
        If m_frmPldegeDeposits.lstDeposits.ListItems(loopCount).SubItems(1) = m_AccArr(count) Then
            m_PledgeDeposits = m_PledgeDeposits & gDelim & m_frmPldegeDeposits.lstDeposits.ListItems(loopCount)
            Exit Do
        End If
        loopCount = loopCount + 1
        If loopCount > m_frmPldegeDeposits.lstDeposits.ListItems.count Then Exit For
    Loop
Next

m_PledgeDeposits = Mid(m_PledgeDeposits, Len(gDelim) + 1)
'm_PledgeAccounts = StrAccList
m_PledgeValue = TotalBalance

If Val(GetVal("InterestRate")) > 0 Then Exit Sub

Dim IntRate As Single
Dim rst As Recordset

For count = 0 To MaxCount
    If m_DepTypeArr(count) = wisDeposit_PD Then
        gDbTrans.SqlStmt = "Select RateOfInterest From " & _
                " PDMaster Where Accid = " & m_AccArr(count)
    
    ElseIf m_DepTypeArr(count) = wisDeposit_RD Then
        gDbTrans.SqlStmt = "Select RateOfInterest From " & _
                " RDMaster Where AccID = " & (m_AccArr(count))
    
    Else
        gDbTrans.SqlStmt = "Select RateOfInterest From " & _
                " FDMaster Where AccID = " & (m_AccArr(count)) & _
                " And DepositType = " & m_DepTypeArr(count)
    End If
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        If IntRate < FormatField(rst(0)) Then IntRate = FormatField(rst(0))

Next

IntRate = IntRate + 2
txtLoanIssue(GetIndex("InterestRate")) = IntRate

End Sub


Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)

Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset

SqlStr = "SELECT * From DepositLoanTrans WHERE LoanId = " & m_LoanID & _
    " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
    " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Set clsPrint = New clsTransPrint

'Printer.PaperSize = 9
Printer.Font.name = gFontName
Printer.Font.Size = 12 'gFontSize
Dim clsCust As New clsCustReg
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & clsCust.CustomerName(m_CustID)
    Set clsCust = Nothing
    
    .Cols = 5
    .ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    .ColWidth(1) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    .ColWidth(2) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    .ColWidth(3) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    .ColWidth(4) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    While Not rst.EOF
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(2) = FormatField(rst("Amount"))
            .ColText(3) = " "
        Else
            .ColText(2) = " "
            .ColText(3) = FormatField(rst("Amount"))
        End If
        .ColText(4) = FormatField(rst("Balance"))
        .PrintText
        rst.MoveNext
    Wend
    .newPage
End With

Set rst = Nothing
Set clsPrint = Nothing

End Sub

Private Sub m_frmPrintTrans_TransClick(bNewPassbook As Boolean)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintId, lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer


'First get the last printed txnId and last printed row From the DEPLoanMaster
SqlStr = "SELECT LastPrintID, LastPrintRow From DepositLoanMaster WHERE LoanID = " & m_LoanID
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintId"))

' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From DepositLoanTrans WHERE LoanID = " & m_LoanID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

' Print the first page of passbook, if newPassbook option is chosen.
If bNewPassbook Then
    clsPrint.printPassbookPage wis_DepositLoans, m_AccID
    
        'Update the print rows as 0 and lst transid as -5
    'Now Update the Last Print Id to the master
    SqlStr = "UPDATE DepositLoanMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
            ", LastPrintRow = 0 Where LoanID = " & m_LoanID
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
    Else
        gDbTrans.CommitTrans
    End If
    Exit Sub

End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
If (rst(0) = 0) Then
     MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

' Fetch records for txns that have been created after lasttxnId.
SqlStr = "SELECT * From DepositLoanTrans WHERE LoanID = " & m_LoanID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Print [or don't print] header part
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If


'Printer.PaperSize = 9
'Printer.Font.Name = gFontName
'Printer.Font.Size = 12 'gFontSize
Printer.Font = "Courier New"
Printer.FONTSIZE = 9
With clsPrint
   ' .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 4
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .COlHeader(1) = GetResourceString(275) 'Cheque
    '.ColWidth(2) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(3) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(4) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(5) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    
    If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
    End If
    
    ' column widths for printing txn rows.
        .ColWidth(0) = 17
        .ColWidth(1) = 24
        .ColWidth(2) = 15
        .ColWidth(3) = 17
        .ColWidth(4) = 17
        '.ColWidth(5) = 15

    While Not rst.EOF
        If .isNewPage Then
            .printHeader3
            .isNewPage = False
        End If

        TransID = FormatField(rst("TransID"))
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        
       If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(2) = FormatField(rst("Amount"))
            .ColText(3) = " "
        Else
            .ColText(2) = " "
            .ColText(3) = FormatField(rst("Amount"))
        End If
             
        .ColText(4) = FormatField(rst("Balance"))
        .printRow1
        
        ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE) Then
        
            ' since we have to print now in a new page,
            ' we need to print the header.
            ' So, set columns widths for header.
            
            .newPage
           ' MsgBox "plz insert new page"
            curPrintRow = 1
        End If
        rst.MoveNext
    Wend
    .newPage
End With
Printer.EndDoc

Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing

'Now Update the Last Print Id to the master
SqlStr = "UPDATE DepositLoanMaster set LastPrintId = " & TransID & _
        ", LastPrintRow = " & curPrintRow - 1 & _
        " Where LoanID = " & m_LoanID
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If


End Sub


Private Sub m_LookUp_SelectClick(strSelection As String)
m_SelItem = strSelection
End Sub


Private Sub optReport_Click(Index As Integer)

Dim Dt1 As Boolean
Dim Amt As Boolean

Dim Gender As Boolean


Select Case Index
    Case 0
        Dt1 = False
        Amt = True
        Gender = True
    Case 1
        Dt1 = False
        Amt = False
        Gender = True
    Case 2
        Dt1 = True
        Amt = True
        Gender = True
    Case 3
        Dt1 = False
        Amt = True
    Case 4
        Dt1 = True
        Amt = False
        Gender = False
    Case 5
        Dt1 = True
        Amt = True
        Gender = True
    Case 6
        Dt1 = True
        Amt = False
        Gender = False
        
End Select

    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = Gender
        .EnableAmountRange = Amt
    End With

    With txtStartDate
        .Enabled = Dt1:
        .BackColor = IIf(Dt1, wisWhite, wisGray)
    End With
    lblDate1.Enabled = Dt1

    cmdStDate.Enabled = Dt1

End Sub

Private Sub picLoanIssueSlider_Resize()

'If Not arrangerequired Then Exit Sub
On Error Resume Next

' Arrange the controls on this container.
Dim Ctl As Control
For Each Ctl In Me.Controls
    If Ctl.Container.name = picLoanIssueSlider.name Then
        If TypeOf Ctl Is Label Then
            
        ElseIf TypeOf Ctl Is TextBox Then
        
        ElseIf TypeOf Ctl Is ComboBox Then
        
        ElseIf TypeOf Ctl Is CommandButton Then
        
        End If
    End If
Next


End Sub

Private Sub TabStrip_Click()
On Error Resume Next
Dim txtIndex As Integer
cmdRepay.Default = False
Select Case UCase$(TabStrip.SelectedItem.Key)
    
    Case "LOANACCOUNTS"
        fraLoanAccounts.Visible = True
        fraRepayments.Visible = False
        fraReports.Visible = False
        
        txtIndex = GetIndex("MemberID")
        txtLoanIssue(txtIndex).SetFocus
        cmdSaveLoan.Default = cmdSaveLoan.Enabled
        cmdLoanUpdate.Default = cmdLoanUpdate.Enabled
                
    Case "REPAYMENTS"
        
        fraLoanAccounts.Visible = False
        fraRepayments.Visible = True
        fraReports.Visible = False
        
        cmdLoanInst.Default = True
        cmdRepay.Default = cmdRepay.Enabled
        
    Case "REPORTS"
        fraLoanAccounts.Visible = False
        fraRepayments.Visible = False
        fraReports.Visible = True
        
        cmdView.Default = True
        
End Select

End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    fraInstructions.Visible = True
    fraInstructions.ZOrder 0
    fraLoanGrid.Visible = False
Else
    fraLoanGrid.ZOrder 0
    fraLoanGrid.Visible = True
    fraInstructions.Visible = False
End If

End Sub

Private Sub txtAccNo_Change()

Call LoanClear
If Trim$(txtAccNo.Text) <> "" Then
    cmdLoad.Enabled = True
    txtRepayAmt.Enabled = True
    txtRepayDate.Enabled = True
    If gOnLine Then cmdRepayDate.Enabled = False
    
Else
    cmdLoad.Enabled = False
    txtRepayAmt.Enabled = False
    txtRepayDate.Enabled = False
End If
End Sub

Private Sub txtAccNo_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'Me.ActiveControl
End Sub


Private Sub txtEndAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtEndDate_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtIntBalance_GotFocus()
With txtIntBalance
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtLoanIssue_DblClick(Index As Integer)
Dim strDispType As String
' Get the display type.
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    txtLoanIssue_KeyPress Index, vbKeyReturn
End If

End Sub
Private Sub txtLoanIssue_GotFocus(Index As Integer)

lblLoanIssue(Index).ForeColor = vbBlue
PropSetIssueDescription lblLoanIssue(Index)

Dim DepositType As Integer

DepositType = GetDepositType(GetVal("DepositType"))
' Scroll the window, so that the
' control in focus is visible.
PropScrollLoanIssueWindow txtLoanIssue(Index)

' Select the text, if any.
With txtLoanIssue(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' If the display type is Browse, then
' show the command button for this text.
Dim strDispType As String
Dim TextIndex As String
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    'Get the cmdbutton index.
    TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "TextIndex")
    If TextIndex <> "" Then cmd(Val(TextIndex)).Visible = True
    strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")
    If StrComp(strDispType, "PledgeDeposit", vbTextCompare) = 0 Then
        cmd(Val(TextIndex)).Enabled = True
    End If
End If

' Hide all other command buttons...
Dim I As Integer
For I = 0 To cmd.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmd(I).Visible = False
    End If
Next

If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "textindex")
    ' Get the cmdbutton index.
    On Error Resume Next
    If TextIndex <> "" Then
        If cmb(Val(TextIndex)).Visible Then Exit Sub
        cmb(Val(TextIndex)).Visible = True
        cmb(Val(TextIndex)).SetFocus
    End If
End If
' Hide all other combo boxes.
For I = 0 To cmb.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmb(I).Visible = False
    End If
Next

End Sub
Private Sub txtLoanIssue_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(lblLoanIssue(Index).Tag, "TextIndex")
        If Trim$(strIndex) <> "" Then
            cmb(Val(strIndex)).Visible = True
            cmb(Val(strIndex)).SetFocus
            cmb(Val(strIndex)).ZOrder 0
        End If
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub txtLoanIssue_LostFocus(Index As Integer)
lblLoanIssue(Index).ForeColor = vbBlack

' Declare the reqd. member variables here...
Dim strDataSrc As String
Dim Lret As Long
Dim txtIndex As Integer
' Get the name of the data source bound to this control.
strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")

Dim strDispType As String
Dim TextIndex As String
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
Select Case UCase(strDispType)
    Case "BROWSE"
        ' Get the cmdbutton index.
        TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "TextIndex")
        If TextIndex <> "" And ActiveControl.name <> cmd(0).name Then _
            cmd(Val(TextIndex)).Visible = False
    
    Case "LIST"
        ' Get the cmdbutton index.
        TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "TextIndex")
        If TextIndex <> "" And ActiveControl.name <> cmb(0).name Then _
            cmb(Val(TextIndex)).Visible = False
End Select

strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")
TextIndex = GetIndex(strDispType)
Select Case UCase(strDataSrc)
    Case "MEMBERID"
        'Get the CustomerName
        Dim custId As Long
        Dim custReg As clsCustReg
        TextIndex = GetIndex("MemberID")
        m_CustID = Val(txtLoanIssue(TextIndex).Text)
        If m_CustID <= 0 Then Exit Sub
        'Now Search the Customer
        Set custReg = New clsCustReg
        txtLoanIssue(GetIndex("MemberName")).Text = custReg.CustomerName(m_CustID)
        If Trim$(GetVal("MemberName")) = "" Then txtLoanIssue(TextIndex) = ""
        'm_CustReg.LoadCustomerInfo (CustId)
        Set custReg = Nothing
    Case "PLEDGEDEPOSIT"
        If Trim$(txtLoanIssue(txtIndex).Text) = "" Then Exit Sub
        
End Select


End Sub



Private Sub txtMisc_Change()
    Call txtRegInterest_Change
   'txtRepayAmt = txtRegInterest + txtIntBalance + txtMisc + txtPrincAmount
End Sub

Private Sub txtPenalInterest_Change()
   'txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest) + Val(txtMisc.Text))
End Sub

Private Sub txtPenalInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtMisc_GotFocus()
With txtMisc
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtPrincAmount_GotFocus()
With txtPrincAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtRegInterest_Change()
    txtRepayAmt = txtRegInterest + txtMisc + txtPrincAmount
End Sub

Private Sub txtRegInterest_GotFocus()
With txtRegInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtRepayAmt_Change()
If txtRepayAmt <> 0 Then
    cmdRepay.Enabled = True
Else
    cmdRepay.Enabled = False
End If

If Me.ActiveControl.name = txtRepayAmt.name Then _
    txtPrincAmount = txtRepayAmt - txtMisc - txtRegInterest
    
End Sub

Private Sub txtRepayAmt_GotFocus()
With txtRepayAmt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtRepayDate_LostFocus()
If Not DateValidate(txtRepayDate, "/", True) Then Exit Sub
Dim TransDate As Date
Dim IntAmount As Currency

TransDate = GetSysFormatDate(txtRepayDate.Text)
IntAmount = ComputeDepLoanRegularInterest(TransDate, m_LoanID)
txtRegInterest = IntAmount
txtIntBalance = GetInterstBalance

'IntAmount = ComputeDepLoanPenalInterest(transDate, m_LoanID)
'txtPenalInterest = IntAmount
End Sub


Private Sub txtStartAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtTotInst_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub vscLoanIssue_Change()
' Move the picLoanissueSlider.
picLoanIssueSlider.Top = -vscLoanIssue.Value
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

'Set kannada caption for the generally used  controls
cmdOk.Caption = GetResourceString(1)

'Set captions for the all the tabs
TabStrip.Tabs(1).Caption = GetResourceString(216)
TabStrip.Tabs(2).Caption = GetResourceString(80, 36) & GetResourceString(92)
TabStrip.Tabs(3).Caption = GetResourceString(283) & GetResourceString(92)

TabStrip1.Tabs(1).Caption = GetResourceString(219) 'Instructions
TabStrip1.Tabs(2).Caption = GetResourceString(28) 'Instructions
TabStrip1.Tabs(2).Selected = True


'Set captions for the Loan Accounts
fraLoanAccounts.Caption = GetResourceString(80, 36) & GetResourceString(92)
cmdSaveLoan.Caption = GetResourceString(15)
cmdLoanUpdate.Caption = GetResourceString(7)
cmdLoanIssueClear.Caption = GetResourceString(8)
cmdPhoto.Caption = GetResourceString(415)

'Set kannada captions for the Repayments
lblDepositType.Caption = GetResourceString(45) '"Deposit Type
lblMisc.Caption = GetResourceString(327)
lblLoanNo.Caption = GetResourceString(58, 60)
lblCustName.Caption = GetResourceString(35)
lblLoanAmount.Caption = GetResourceString(58)
lblBalance.Caption = GetResourceString(67, 58)  'Balance Loan
lblRegInterest.Caption = GetResourceString(344)
lblIntBlance.Caption = GetResourceString(67, 47)  'Balance Interest
lblPrincipal.Caption = GetResourceString(310) 'Principal
lblRepayAmt.Caption = GetResourceString(52, 42) 'Total Amount
'fraLoanGrid.Caption = GetResourceString(58,38)
cmdLoad.Caption = GetResourceString(3)
cmdLoanInst.Caption = GetResourceString(290) 'Loan Payment
cmdUndo.Caption = GetResourceString(5)
cmdRepay.Caption = GetResourceString(20)
lblRepayDate.Caption = GetResourceString(38, 37)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)

'Set kannadacaption for Reports tab
fraReports.Caption = GetResourceString(283) & GetResourceString(92)
'fraDaterange.Caption = GetResourceString(106)

optReport(0).Caption = GetResourceString(80, 42) 'Loan BAlance
optReport(1).Caption = GetResourceString(80, 295) 'Loan details
optReport(2).Caption = GetResourceString(390, 63) 'Loan Sub day book
optReport(3).Caption = GetResourceString(84, 18)     'Over due loans
optReport(4).Caption = GetResourceString(43, 80, 93)
optReport(5).Caption = GetResourceString(85)
optReport(6).Caption = GetResourceString(463, 42) 'Monthly balance
optAccId.Caption = GetResourceString(36, 60) '"lOAN aCCOUNT nO
optName.Caption = GetResourceString(69) ' Name

lblDep.Caption = GetResourceString(58, 45) 'Loan Deposit
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)

cmdView.Caption = GetResourceString(13)
cmdAdvance.Caption = GetResourceString(491)    'Options
End Sub

