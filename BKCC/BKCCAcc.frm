VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmBKCCAcc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX-2000 - BKCC"
   ClientHeight    =   8985
   ClientLeft      =   3915
   ClientTop       =   465
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   6480
      TabIndex        =   39
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Frame fraRepayments 
      Height          =   7725
      Left            =   240
      TabIndex        =   1
      Top             =   540
      Width           =   7425
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1890
         TabIndex        =   17
         Top             =   2925
         Width           =   1725
      End
      Begin VB.CommandButton cmdMisc 
         Caption         =   "..."
         Height          =   315
         Left            =   6750
         TabIndex        =   109
         Top             =   2505
         Width           =   315
      End
      Begin VB.CommandButton cmdAbn 
         Caption         =   "..."
         Height          =   315
         Left            =   3300
         TabIndex        =   108
         Top             =   1575
         Width           =   315
      End
      Begin VB.TextBox txtVoucherNo 
         Height          =   315
         Left            =   5400
         TabIndex        =   38
         Top             =   4020
         Width           =   1395
      End
      Begin VB.CommandButton cmdCheque 
         Caption         =   "..."
         Height          =   315
         Left            =   6810
         TabIndex        =   107
         Top             =   4020
         Width           =   315
      End
      Begin VB.ComboBox cmbCheque 
         Height          =   315
         Left            =   5430
         TabIndex        =   106
         Top             =   4020
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtRepayDate 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1890
         TabIndex        =   101
         Top             =   2505
         Width           =   1335
      End
      Begin VB.CommandButton cmdMember 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   6750
         TabIndex        =   7
         Top             =   300
         Width           =   330
      End
      Begin VB.TextBox txtMemberID 
         Height          =   350
         Left            =   5385
         MaxLength       =   9
         TabIndex        =   6
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Top             =   4020
         Width           =   315
      End
      Begin VB.TextBox txtSbAccNum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2790
         TabIndex        =   36
         Top             =   4020
         Width           =   645
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1665
         TabIndex        =   19
         Top             =   3465
         Width           =   2115
      End
      Begin VB.CommandButton cmdRepayDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3270
         TabIndex        =   15
         Top             =   2505
         Width           =   315
      End
      Begin VB.TextBox txtAccNo 
         Height          =   350
         Left            =   1650
         MaxLength       =   9
         TabIndex        =   3
         Top             =   300
         Width           =   1110
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   400
         Left            =   2880
         TabIndex        =   4
         Top             =   230
         Width           =   840
      End
      Begin VB.CheckBox chkSb 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit To SB Account"
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   3930
         Width           =   2355
      End
      Begin WIS_Currency_Text_Box.CurrText txtRegInterest 
         Height          =   345
         Left            =   5520
         TabIndex        =   24
         Top             =   1605
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPenalInterest 
         Height          =   345
         Left            =   5520
         TabIndex        =   27
         Top             =   2085
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMiscAmount 
         Height          =   345
         Left            =   5520
         TabIndex        =   29
         Top             =   2505
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtRepayAmt 
         Height          =   345
         Left            =   5520
         TabIndex        =   33
         Top             =   2925
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotal 
         Height          =   345
         Left            =   5400
         TabIndex        =   31
         Top             =   3465
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtIntBalance 
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   1200
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraLoanGrid 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2595
         Left            =   270
         TabIndex        =   40
         Top             =   4890
         Width           =   6975
         Begin VB.CommandButton cmdNotice 
            Caption         =   "Notice"
            Height          =   400
            Left            =   0
            TabIndex        =   142
            Top             =   2160
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdUndo 
            Caption         =   "Undo Last"
            Enabled         =   0   'False
            Height          =   400
            Left            =   3960
            TabIndex        =   141
            Top             =   2160
            Width           =   1545
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "&Accept"
            Height          =   400
            Left            =   5640
            TabIndex        =   140
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6330
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   1710
            Width           =   435
         End
         Begin VB.CommandButton cmdPrevTrans 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6330
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   105
            Width           =   435
         End
         Begin VB.CommandButton cmdNextTrans 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6330
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   630
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2010
            Left            =   60
            TabIndex        =   41
            Top             =   90
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   3545
            _Version        =   393216
            Rows            =   10
            Cols            =   5
            AllowBigSelection=   0   'False
            ScrollBars      =   2
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2985
         Left            =   120
         TabIndex        =   139
         Top             =   4560
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   5265
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
               Caption         =   "Pass book"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2505
         Left            =   240
         TabIndex        =   143
         Top             =   4920
         Width           =   6975
         Begin VB.CommandButton cmdAddNote 
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
            Left            =   6360
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   90
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2325
            Left            =   60
            TabIndex        =   145
            Top             =   120
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   4101
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"BKCCAcc.frx":0000
         End
      End
      Begin VB.Label lblTrans 
         Caption         =   "Trans"
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   2925
         Width           =   1755
      End
      Begin VB.Label txtIssueDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "qqq"
         Height          =   345
         Left            =   5445
         TabIndex        =   104
         Top             =   5355
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblIssueDate 
         AutoSize        =   -1  'True
         Caption         =   "Issued On :"
         Height          =   315
         Left            =   3930
         TabIndex        =   11
         Top             =   5265
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label txtLoanAmt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   105
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label txtBalance 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   103
         Top             =   1605
         Width           =   1365
      End
      Begin VB.Label txtLastInt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   102
         Top             =   2085
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         X1              =   7200
         X2              =   60
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   7140
         X2              =   90
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Label lblChequeNo 
         Caption         =   "Cheque No"
         Height          =   315
         Left            =   3840
         TabIndex        =   37
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label lblMemberId 
         AutoSize        =   -1  'True
         Caption         =   "Member Number :"
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Top             =   315
         Width           =   1260
      End
      Begin VB.Label lblLastInt 
         Caption         =   "Last Int Paid date"
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   2085
         Width           =   1755
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotInst 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         Height          =   315
         Left            =   4140
         TabIndex        =   30
         Top             =   3465
         Width           =   1530
      End
      Begin VB.Label lblParticulars 
         Caption         =   "Particulars"
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   3465
         Width           =   1455
      End
      Begin VB.Label txtMemberName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1650
         TabIndex        =   9
         Top             =   690
         Width           =   5445
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   7200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblLoanAccNO 
         AutoSize        =   -1  'True
         Caption         =   "Loan Account No"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label lblIntBalance 
         Caption         =   "Interest Balance :"
         Height          =   315
         Left            =   3810
         TabIndex        =   20
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label lblRegInterest 
         AutoSize        =   -1  'True
         Caption         =   "Regular Interest :"
         Height          =   315
         Left            =   3810
         TabIndex        =   23
         Top             =   1605
         Width           =   1410
      End
      Begin VB.Label lblPenalInterest 
         AutoSize        =   -1  'True
         Caption         =   "Penal Interest :"
         Height          =   315
         Left            =   3810
         TabIndex        =   26
         Top             =   2085
         Width           =   1410
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   "Misc Amount"
         Height          =   315
         Left            =   3810
         TabIndex        =   28
         Top             =   2505
         Width           =   1410
      End
      Begin VB.Label lblRepayDate 
         AutoSize        =   -1  'True
         Caption         =   "Date of repayment : "
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   2505
         Width           =   1785
      End
      Begin VB.Label lblRepayAmt 
         AutoSize        =   -1  'True
         Caption         =   "Repaid amount :"
         Height          =   315
         Left            =   3810
         TabIndex        =   32
         Top             =   2925
         Width           =   1410
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         Caption         =   "Balance amount :"
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Top             =   1605
         Width           =   1605
      End
      Begin VB.Label lblLoanAmount 
         AutoSize        =   -1  'True
         Caption         =   "Loan amount :"
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   1440
      End
   End
   Begin VB.Frame fraLoanAccounts 
      Caption         =   "Loan Accounts..."
      Height          =   7725
      Left            =   240
      TabIndex        =   98
      Top             =   540
      Width           =   7425
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   400
         Left            =   6000
         TabIndex        =   138
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoanUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6030
         TabIndex        =   53
         Top             =   6270
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoanClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   6030
         TabIndex        =   54
         Top             =   6855
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoanSave 
         Caption         =   "C&reate"
         Height          =   400
         Left            =   6030
         TabIndex        =   52
         Top             =   5700
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   1110
         Left            =   270
         ScaleHeight     =   1050
         ScaleWidth      =   5595
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   465
         Width           =   5655
         Begin VB.Image Image2 
            Height          =   405
            Left            =   135
            Picture         =   "BKCCAcc.frx":0082
            Stretch         =   -1  'True
            Top             =   90
            Width           =   345
         End
         Begin VB.Label lblLoanIssueDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   960
            TabIndex        =   44
            Top             =   450
            Width           =   4590
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
            Height          =   315
            Left            =   990
            TabIndex        =   43
            Top             =   45
            Width           =   135
         End
      End
      Begin VB.PictureBox picLoanIssueViewPort 
         BackColor       =   &H00FFFFFF&
         Height          =   5640
         Left            =   270
         ScaleHeight     =   5580
         ScaleWidth      =   5610
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1620
         Width           =   5670
         Begin VB.PictureBox picLoanIssueSlider 
            Height          =   1740
            Left            =   -15
            ScaleHeight     =   1680
            ScaleWidth      =   5280
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   15
            Width           =   5340
            Begin VB.CommandButton cmdLoanIssue 
               Caption         =   "..."
               Height          =   345
               Index           =   0
               Left            =   4890
               TabIndex        =   50
               Top             =   0
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2880
               Style           =   2  'Dropdown List
               TabIndex        =   49
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
               Left            =   2685
               TabIndex        =   48
               Top             =   0
               Width           =   2580
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
               TabIndex        =   47
               Top             =   0
               Width           =   2610
            End
         End
         Begin VB.VScrollBar vscLoanIssue 
            Height          =   1755
            Left            =   5355
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.Frame fraProps 
      Caption         =   "Properties"
      Height          =   7725
      Left            =   240
      TabIndex        =   110
      Top             =   540
      Width           =   7425
      Begin VB.TextBox txtNabardRate 
         Height          =   315
         Left            =   4020
         TabIndex        =   158
         Top             =   5520
         Width           =   705
      End
      Begin VB.TextBox txtRebateRate2 
         Height          =   315
         Left            =   4020
         TabIndex        =   156
         Top             =   6600
         Width           =   705
      End
      Begin VB.TextBox txtRebateRate1 
         Height          =   315
         Left            =   4020
         TabIndex        =   154
         Top             =   6180
         Width           =   705
      End
      Begin VB.TextBox txtSubsidyRate 
         Height          =   315
         Left            =   4020
         TabIndex        =   153
         Top             =   4800
         Width           =   705
      End
      Begin VB.TextBox txtRebateRate 
         Height          =   315
         Left            =   4020
         TabIndex        =   151
         Top             =   5160
         Width           =   705
      End
      Begin VB.TextBox txtEffectiveDate 
         Height          =   345
         Left            =   1770
         TabIndex        =   22
         Top             =   7110
         Width           =   1545
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   400
         Left            =   3600
         TabIndex        =   25
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   7
         Left            =   4020
         TabIndex        =   137
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   4230
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   6
         Left            =   4020
         TabIndex        =   134
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   5
         Left            =   4020
         TabIndex        =   131
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   3300
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   4
         Left            =   4020
         TabIndex        =   128
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   2790
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   3
         Left            =   4020
         TabIndex        =   125
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   2340
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   2
         Left            =   4020
         TabIndex        =   122
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   1860
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   1
         Left            =   4020
         TabIndex        =   119
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   1410
         Width           =   705
      End
      Begin VB.TextBox txtLoanIntRate 
         Height          =   345
         Index           =   0
         Left            =   4020
         TabIndex        =   116
         ToolTipText     =   "Rate of Inteest for Loan limit"
         Top             =   900
         Width           =   705
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   0
         Left            =   1770
         TabIndex        =   115
         Top             =   900
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   1
         Left            =   1770
         TabIndex        =   118
         Top             =   1410
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   2
         Left            =   1770
         TabIndex        =   121
         Top             =   1860
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   3
         Left            =   1770
         TabIndex        =   124
         Top             =   2340
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   4
         Left            =   1770
         TabIndex        =   127
         Top             =   2790
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   5
         Left            =   1770
         TabIndex        =   130
         Top             =   3300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   6
         Left            =   1770
         TabIndex        =   133
         Top             =   3750
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMaxLimit 
         Height          =   345
         Index           =   7
         Left            =   1770
         TabIndex        =   136
         Top             =   4260
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblNabardRate 
         Alignment       =   1  'Right Justify
         Caption         =   "Nabard Rebate Rate of Interest"
         Height          =   315
         Left            =   240
         TabIndex        =   159
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label lblRebateRate2 
         Alignment       =   1  'Right Justify
         Caption         =   "Rebate Rate of Interest"
         Height          =   315
         Left            =   240
         TabIndex        =   157
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label lblRebateRate1 
         Alignment       =   1  'Right Justify
         Caption         =   "Rebate Rate of Interest"
         Height          =   315
         Left            =   360
         TabIndex        =   155
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label lblSubsidyRate 
         Alignment       =   1  'Right Justify
         Caption         =   "Quarterly Subsidy Rate of Interest"
         Height          =   315
         Left            =   240
         TabIndex        =   152
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label lblRebateRate 
         Alignment       =   1  'Right Justify
         Caption         =   "Quarterly Rebate Rate of Interest"
         Height          =   315
         Left            =   240
         TabIndex        =   150
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label lblEffectiveDate 
         Caption         =   "Effective Date"
         Height          =   255
         Left            =   180
         TabIndex        =   96
         Top             =   7140
         Width           =   1815
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   210
         TabIndex        =   135
         Top             =   4230
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   210
         TabIndex        =   132
         Top             =   3750
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   210
         TabIndex        =   129
         Top             =   3300
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   210
         TabIndex        =   126
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   210
         TabIndex        =   123
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   210
         TabIndex        =   120
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   210
         TabIndex        =   117
         Top             =   1410
         Width           =   1395
      End
      Begin VB.Label lblMaxLimit 
         Caption         =   "Loan Maximum Limit"
         Height          =   255
         Left            =   1770
         TabIndex        =   112
         Top             =   450
         Width           =   1695
      End
      Begin VB.Label txtMinLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   114
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label lblLoanIntRate 
         Caption         =   "Loan Interest Rate"
         Height          =   285
         Left            =   4020
         TabIndex        =   113
         Top             =   450
         Width           =   1785
      End
      Begin VB.Label lblMinLimit 
         Caption         =   "Loan minimum limit"
         Height          =   315
         Left            =   120
         TabIndex        =   111
         Top             =   450
         Width           =   1455
      End
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   7725
      Left            =   240
      TabIndex        =   56
      Top             =   540
      Width           =   7425
      Begin VB.Frame fraDepReports 
         Height          =   4575
         Left            =   1200
         TabIndex        =   146
         Top             =   1680
         Visible         =   0   'False
         Width           =   7215
         Begin VB.OptionButton optReports 
            Caption         =   "Deposit Balance"
            Height          =   315
            Index           =   8
            Left            =   400
            TabIndex        =   82
            Top             =   120
            Value           =   -1  'True
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Deposit's Sub Day Book"
            Height          =   315
            Index           =   9
            Left            =   400
            TabIndex        =   83
            Top             =   753
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Interest Received"
            Height          =   315
            Index           =   11
            Left            =   400
            TabIndex        =   86
            Top             =   2655
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Deposit's Daily Cash Book"
            Height          =   315
            Index           =   12
            Left            =   3900
            TabIndex        =   88
            Top             =   753
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Deposits' General Ledger"
            Height          =   315
            Index           =   13
            Left            =   3900
            TabIndex        =   87
            Top             =   120
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Deposit holder list"
            Height          =   315
            Index           =   14
            Left            =   400
            TabIndex        =   84
            Top             =   1386
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Deposit Monthly Balance"
            Height          =   315
            Index           =   15
            Left            =   400
            TabIndex        =   85
            Top             =   2019
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Guarntor List"
            Height          =   315
            Index           =   10
            Left            =   3900
            TabIndex        =   89
            Top             =   1386
            Width           =   3165
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Monthy Deposit Receipt && Paymet"
            Height          =   315
            Index           =   20
            Left            =   3900
            TabIndex        =   90
            Top             =   2019
            Width           =   3225
         End
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Claim Bill"
         Height          =   315
         Index           =   25
         Left            =   480
         TabIndex        =   66
         Top             =   4800
         Width           =   3165
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   80
         TabIndex        =   147
         Top             =   300
         Width           =   7095
         Begin VB.OptionButton optLoanReports 
            Caption         =   "Loan Reports"
            Height          =   300
            Left            =   400
            TabIndex        =   149
            Top             =   0
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton optDepositReports 
            Caption         =   "Deposit Reports"
            Height          =   300
            Left            =   3900
            TabIndex        =   148
            Top             =   30
            Width           =   3015
         End
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Loan &advanced"
         Height          =   315
         Index           =   22
         Left            =   3990
         TabIndex        =   70
         Top             =   1939
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Loans &Recovery"
         Height          =   315
         Index           =   23
         Left            =   3990
         TabIndex        =   74
         Top             =   3807
         Width           =   3165
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Other Receivables"
         Height          =   315
         Index           =   21
         Left            =   3990
         TabIndex        =   72
         Top             =   2873
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Member Tranasaction"
         Height          =   315
         Index           =   19
         Left            =   480
         TabIndex        =   64
         Top             =   4290
         Width           =   3285
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Schedule &1"
         Height          =   315
         Index           =   18
         Left            =   3990
         TabIndex        =   73
         Top             =   3340
         Width           =   3015
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Monthly Register "
         Height          =   315
         Index           =   17
         Left            =   3990
         TabIndex        =   68
         Top             =   1005
         Width           =   3165
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Monthly Receipt && Payment"
         Height          =   315
         Index           =   16
         Left            =   480
         TabIndex        =   63
         Top             =   3807
         Width           =   3285
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Loan Monthly Balance"
         Height          =   315
         Index           =   7
         Left            =   480
         TabIndex        =   62
         Top             =   3330
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "General Ledger"
         Height          =   315
         Index           =   5
         Left            =   480
         TabIndex        =   59
         Top             =   1899
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "SUb Cash Book"
         Height          =   315
         Index           =   4
         Left            =   3990
         TabIndex        =   69
         Top             =   1472
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Loan Holder List"
         Height          =   315
         Index           =   6
         Left            =   480
         TabIndex        =   61
         Top             =   2853
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Interest Collected"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   60
         Top             =   2376
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Over due Loans"
         Height          =   315
         Index           =   2
         Left            =   3990
         TabIndex        =   71
         Top             =   2406
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Sub Day  Book"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   58
         Top             =   1422
         Width           =   3075
      End
      Begin VB.OptionButton optReports 
         Caption         =   "LoanBalance"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   57
         Top             =   945
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   400
         Left            =   5910
         TabIndex        =   97
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Frame fraOrder 
         Height          =   1815
         Left            =   120
         TabIndex        =   75
         Top             =   5370
         Width           =   7155
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   400
            Left            =   5700
            TabIndex        =   95
            Top             =   1250
            Width           =   1215
         End
         Begin VB.ComboBox cmbFarmer 
            Height          =   315
            Left            =   1740
            TabIndex        =   79
            Top             =   1350
            Width           =   1770
         End
         Begin VB.TextBox txtStartDate 
            Height          =   345
            Left            =   1725
            TabIndex        =   91
            Top             =   800
            Width           =   1350
         End
         Begin VB.TextBox txtEndDate 
            Height          =   345
            Left            =   5235
            TabIndex        =   94
            Top             =   800
            Width           =   1290
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   315
            Left            =   3180
            TabIndex        =   81
            Top             =   800
            Width           =   315
         End
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   315
            Left            =   6585
            TabIndex        =   93
            Top             =   800
            Width           =   315
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   375
            Left            =   510
            TabIndex        =   76
            Top             =   180
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   375
            Left            =   3900
            TabIndex        =   77
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label lblfarmer 
            Caption         =   "Farmer"
            Height          =   315
            Left            =   150
            TabIndex        =   78
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   6930
            X2              =   60
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "Starting date :"
            Height          =   315
            Left            =   120
            TabIndex        =   80
            Top             =   850
            Width           =   1440
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "Ending date :"
            Height          =   315
            Left            =   3840
            TabIndex        =   92
            Top             =   850
            Width           =   1305
         End
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Claim Bill"
         Height          =   315
         Index           =   24
         Left            =   3990
         TabIndex        =   65
         Top             =   4320
         Width           =   3165
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Claim Bill"
         Height          =   315
         Index           =   26
         Left            =   3990
         TabIndex        =   67
         Top             =   4800
         Width           =   3165
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   8325
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   14684
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
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
            Object.ToolTipText     =   "Creation of New account"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Report of loan balance,etc."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Properties"
            Key             =   "Props"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Show the properties of Kissan Cedit card"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBKCCAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_LoanID As Long
Private m_rstLoanTrans As Recordset
Private m_rstLoanMast As Recordset
Private m_TransID As Long
Private m_retVar As Variant
Private m_Notes As New clsNotes

Private m_KCCDepositTrans As Boolean

Private m_DepHeadID As Long
Private m_LoanHeadID As Long
Private m_DepIntHeadID As Long
Private m_RegIntHeadID As Long
Private m_PenalHeadID As Long

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private WithEvents m_frmLoanReport As frmBkccReport
Attribute m_frmLoanReport.VB_VarHelpID = -1
Private WithEvents m_frmAbn As frmBKCCAbn
Attribute m_frmAbn.VB_VarHelpID = -1
Private WithEvents m_frmAsset As frmAsset
Attribute m_frmAsset.VB_VarHelpID = -1
Private m_clsRepOption  As clsRepOption
'Private WithEvents m_frmReportOption As frmReportview


Private m_clsReceivable As clsReceive
Private m_CustReg As clsCustReg
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
'Private WithEvents m_frmCheque As frmCheque
Private m_frmCheque As frmCheque

Const CTL_MARGIN = 15
Private m_dbOperation As wis_DBOperation

'Declare events.
Public Event SetStatus(strMsg As String)
Public Event WindowClosed()
Public Event AccountChanged(ByVal LoanID As Long)

Public Event ShowReport(ReportType As wis_BKCCReports, ReportOrder As wis_ReportOrder, _
        FarmerType As wisFarmerClassification, fromDate As String, toDate As String, _
        RepOptClass As clsRepOption)

Private Function AddInterestToDeposit() As Boolean
Dim MaxTransID As Long
Dim Balance As Currency
Dim LastTransDate As Date
Dim Deposit As Boolean
Dim TransDate As Date
Dim rst As Recordset

TransDate = GetSysFormatDate(txtRepayDate)

' Get a new transactionID.
MaxTransID = GetKCCMaxTransID(m_LoanID)

gDbTrans.SqlStmt = "SELECT Balance,TransID,TransDate " & _
            " FROM BKCCTrans WHERE loanid = " & m_LoanID & _
            " And TransID = " & MaxTransID & " Order By TransID desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Balance = Val(FormatField(rst("Balance")))
    LastTransDate = rst("TransDate")
End If
'Deposit = IIf(Balance < 0, True, False)
Deposit = True

'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, LastTransDate) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo Exit_Line
End If

'Get the Transaction Id From The Interest table

'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, GetKCCLastTransDate(m_LoanID)) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo Exit_Line
End If

'Dim TransID As Long
MaxTransID = MaxTransID + 1

Dim IntAmount As Currency
Dim transType As wisTransactionTypes
IntAmount = txtTotal.Value

' Begin the transaction
If Not gDbTrans.BeginTrans Then GoTo Exit_Line
'First do The Interest TransCtion
    transType = wContraWithdraw
    Deposit = True
    gDbTrans.SqlStmt = "INSERT INTO BKCCIntTrans (LoanID, TransID," _
            & " TransDate, TransType, IntAmount, PenalIntAmount,  " _
            & "MiscAmount, IntBalance,Deposit, Particulars,UserID ) " _
            & "VALUES (" & m_LoanID & ", " & MaxTransID & ", " _
            & " #" & TransDate & "#," & transType & ", " _
            & txtRegInterest.Value & ", 0 ," _
            & txtMiscAmount.Value & ", 0 ," & Deposit & "," _
            & AddQuotes(GetResourceString(47), True) _
            & "," & gUserID & ")"
            
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

'Now insert The Deposit Interest As Debit to the account
transType = wContraDeposit
Deposit = CBool(Balance < 0)
Balance = Balance - IntAmount
gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID," _
            & " TransDate, TransType, Amount,Balance,Deposit, Particulars, UserID ) " _
            & "VALUES (" & m_LoanID & ", " & MaxTransID & ", " _
            & " #" & TransDate & "#," & transType & ", " _
            & IntAmount & "," _
            & Balance & "," & Deposit & "," _
            & AddQuotes(GetResourceString(233)) _
            & "," & gUserID & ")"
    
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
Dim bankClass As New clsBankAcc

'Now Update the Amount to ledger heads
If Not bankClass.UpdateContraTrans(m_DepIntHeadID, _
            IIf(Balance > 0, m_LoanHeadID, m_DepHeadID), IntAmount, TransDate) Then GoTo Exit_Line
    
gDbTrans.CommitTrans
AddInterestToDeposit = True
Exit Function

Exit_Line:

gDbTrans.RollBack

End Function

Private Function CheckForDueAmount() As Currency

'Now check whether he has any amount due
'like Insurence amount, Vehicle insusrence,Abn ,EP, etc
'First Get the Max TransID From this acount
Dim TransID As Integer

'Now Check whether there is any transction
'in Amount receivable & whs transid is moter than TransID
TransID = GetReceivAbleAmountID(m_LoanHeadID, m_LoanID)
If TransID Then 'There is AMount due int
    Dim rstTemp As Recordset
    gDbTrans.SqlStmt = "Select * From AmountReceivAble" & _
            " Where AccHeadID = " & m_LoanHeadID & _
            " And AccID = " & m_LoanID '& _
            " And TransID > (Select Max(TransID) From" & _
                " AmountReceivAble Where AccHeadID = " & m_LoanHeadID & _
                " And AccID = " & m_LoanID & _
                " AND Balance = 0" & ")"
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        Set m_clsReceivable = New clsReceive
        While Not rstTemp.EOF
            Call m_clsReceivable.AddHeadAndAmount(rstTemp("DueHeadID"), rstTemp("Amount"))
            rstTemp.MoveNext
        Wend
        
        With txtMiscAmount
            .Value = m_clsReceivable.TotalAmount * -1
            .Locked = True
            CheckForDueAmount = m_clsReceivable.TotalAmount
        End With
        
    End If
End If

End Function

Private Function CheckValidity() As Boolean
CheckValidity = False

On Error GoTo Exit_Line

Dim txtIndex As Integer
Dim CtrlIndex As Integer
Dim Lret As Long
Dim lGuarantorID As Long
Dim NewLoanID As Long
Dim newTransID As Long
Dim Balance As Currency
Dim Amount As Currency
Dim count As Integer
Dim MemID As Long
Dim custId As Long
Dim SbACCID As Long
Dim rst As Recordset

' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the member id of the person availing the loan.", vbInformation
        MsgBox GetResourceString(715), vbInformation
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    'Check if the specified memberid is valid.
    gDbTrans.SqlStmt = "SELECT * FROM MemMaster " & _
                " WHERE AccNum = " & AddQuotes(.Text, True)
    
    Lret = gDbTrans.Fetch(rst, adOpenDynamic)
    If Lret <= 0 Then
        'MsgBox "MemberID " & .Text & " does not exist.", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(716), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    MemID = FormatField(rst("AccId"))
    custId = FormatField(rst("CustomerID"))
    
    ' ------------------------------------------------
    ' Check for the eligibility of this member...
    ' A member is elibible for taking loan, only if does not have any
    ' over due loans.
    If GetValue("IssueDate") <> "" Then
        If HasOverdueLoans(custId, GetSysFormatDate(GetValue("IssueDate"))) Then
            'nRet = MsgBox("The member hasoverdue loans.  Loan cannot be issued." _
                    & vbCrLf & "Do you want to continue anyway?", vbQuestion + _
                    vbYesNo, wis_MESSAGE_TITLE)
            Lret = MsgBox(GetResourceString(717) _
                    & vbCrLf & GetResourceString(541), vbQuestion + _
                    vbYesNo, wis_MESSAGE_TITLE)
            If Lret = vbNo Then
                ActivateTextBox txtLoanIssue(txtIndex)
                GoTo Exit_Line
            End If
        End If
    End If
End With

'Loan Account No
txtIndex = GetIndex("AccNum")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        'MsgBox "Account NO not Specified.", vbInformation , wis_MESSAGE_TITLE
        MsgBox GetResourceString(36, 60, 296), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for loan amount." & vbcrlf & "Do You want to coninue", _
                    vbInformation, wis_MESSAGE_TITLE
        If MsgBox(GetResourceString(500) & vbCrLf & GetResourceString(501), _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
            Exit Function
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
    End If
    'Now Chack whether any account is there with the same account no
    gDbTrans.SqlStmt = "SELECT * FROM BKCCMaster Where " & _
            " AccNum = " & AddQuotes(Trim(.Text), True)
    If m_dbOperation = Update Then _
            gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND LoanId <> " & m_LoanID
    
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'MsgBox "This Account No already axists Please specify othe account", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(545) & vbCrLf & GetResourceString(641), _
                vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

'LoanSanction Amount
txtIndex = GetIndex("SanctionAmount")
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
End With

'Current Year's Loan Sanction
'LoanSanction Amount
txtIndex = GetIndex("CurrentSanction")
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
End With

'in Current Year's Extra Amount sanctioned then
'LoanSanction Amount
txtIndex = GetIndex("ExtraSanction")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then .Text = "0.00"
    If Not IsNumeric(.Text) Or Val(.Text) < 0 Then
            'MsgBox "Invalid value for loan amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(506), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
    End If
End With

'Memeber type
'Now Get the Farmer Type
txtIndex = GetIndex("FarmerType")
'Now get the index of the combo box which represents the farmer type
CtrlIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
If cmb(CtrlIndex).ListIndex < 0 Then
    MsgBox "Select the farmer Type", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    GoTo Exit_Line
End If

' Loan issue date.
txtIndex = GetIndex("IssueDate")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        MsgBox "Specify the Card issued date of for this loan.", _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

' Loan renew date.
txtIndex = GetIndex("RenewDate")
With txtLoanIssue(txtIndex)
    If Trim(.Text) = "" Then
        MsgBox "Specify the next date of renewal for this loan.", _
                    vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtIndex)
        GoTo Exit_Line
    End If
End With

'Loan Interest Rate
txtIndex = GetIndex("InterestRate")
If Val(txtLoanIssue(txtIndex)) = 0 Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'loan Penal Interest Rate
txtIndex = GetIndex("PenalInterestRate")
If Val(txtLoanIssue(txtIndex)) = 0 Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Interest Rate for deposits
txtIndex = GetIndex("DepositInterestRate")
If Val(txtLoanIssue(txtIndex)) = 0 Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(646), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Guarantors...
txtIndex = GetIndex("Guarantor1")
With txtLoanIssue(txtIndex)
    lGuarantorID = PropGuarantorID(1)
    If Trim(.Text) = "" Then
        'nRet = MsgBox("Guarantor not specified. Do you want to continue?", _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        Lret = MsgBox(GetResourceString(723) & GetResourceString(541), _
                vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE)
        If Lret = vbNo Then GoTo Exit_Line
    Else
        ' Check if the guarantor is the same as the loan claimer.
        lGuarantorID = PropGuarantorID(1)
        If lGuarantorID = custId Then
            'MsgBox "A person cannot stand guarantee for his own loan !", _
                    vbExclamation , wis_MESSAGE_TITLE
            MsgBox GetResourceString(724), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
        ' Check if the guarantor is eligible for standing guarantee.
        If HasOverdueLoans(lGuarantorID) Then  '//TODO//
            'MsgBox "Guarantor1 " & GetValue("Guarantor1") _
                    & " has loan overdues.  Please select another " _
                    & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
            Lret = MsgBox(GetResourceString(725), vbYesNo + _
                vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE)
            If Lret = vbNo Then
                ActivateTextBox txtLoanIssue(txtIndex)
                GoTo Exit_Line
            End If
        End If
            
        If Not IsMemberExistsForCustomerID(lGuarantorID) Then  '//TODO//
            MsgBox "Guarantor1 " & GetValue("Guarantor1") _
                    & " is not a active member. Please select another guarantor." _
                    & "", vbExclamation, wis_MESSAGE_TITLE
            
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
    
    End If
End With

' Guarantor 2.
txtIndex = GetIndex("Guarantor2")
With txtLoanIssue(txtIndex)
    
    If Trim(.Text) <> "" Then
        ' Check if both the guarantors are the same.
        If lGuarantorID <> 0 Then
            If lGuarantorID = PropGuarantorID(2) Then
                'MsgBox "Both guarantors should be different persons.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(726), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtLoanIssue(txtIndex)
                GoTo Exit_Line
            End If
        End If
    
        ' Check if the guarantor is the same as the loan claimer.
        lGuarantorID = PropGuarantorID(2)
        If lGuarantorID = custId Then
            'MsgBox "A person cannot stand guarantee for his own loan!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(724), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
        
        ' Check if the guarantor is eligible for standing guarantee.
        If HasOverdueLoans(lGuarantorID) Then
            'MsgBox "Guarantor2 " & GetValue("Guarantor2") _
                    & " has loan overdues.  Please select another " _
                    & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
            Lret = MsgBox(GetResourceString(727), vbYesNo + _
                vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE)
            If Lret = vbNo Then
                ActivateTextBox txtLoanIssue(txtIndex)
                GoTo Exit_Line
            End If
        End If
    End If
End With

'Now VeryFy the Details of Land
txtIndex = GetIndex("DryLand")
CtrlIndex = GetIndex("WetLand")
If Val(txtLoanIssue(txtIndex)) = 0 And Val(txtLoanIssue(CtrlIndex)) = 0 Then
    MsgBox "Please specify the Land details of the customer", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanIssue(txtIndex)
    GoTo Exit_Line
End If

With txtLoanIssue(txtIndex)
    If Val(.Text) Then
        txtIndex = GetIndex("dryIncome")
        If Val(txtLoanIssue(txtIndex)) = 0 Then
            MsgBox "Please specify the income from the Dry land", vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
    End If
End With

With txtLoanIssue(CtrlIndex)
    If Val(.Text) Then
        txtIndex = GetIndex("WetIncome")
        If Val(txtLoanIssue(txtIndex)) = 0 Then
            MsgBox "Please specify the income from the irrigation land", vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            GoTo Exit_Line
        End If
    End If
End With

CheckValidity = True


Exit_Line:

End Function

Private Sub InsertAmountReceiveAble()

If m_LoanID = 0 Then Exit Sub
If m_clsReceivable Is Nothing Then Exit Sub
If m_LoanHeadID = 0 Then Exit Sub

On Error GoTo Err_line
 
'Dim LastTransDate As Date
Dim TransID As Long
Dim inTransaction As Boolean

Dim NewBalance As Currency
Dim Balance As Currency

Dim MiscAmount As Currency
Dim PrincAmount  As Currency

Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim Deposit As Boolean
Dim rst As Recordset
Dim Amount As Currency
    
Dim bankClass As clsBankAcc
Dim MiscHeadId As Long

Dim VoucherNo As String
Dim UserID As Integer
Dim strParticulars As String
Dim rstTemp As Recordset


'Get the Voucher No and Cheque No
UserID = gCurrUser.UserID
VoucherNo = "xxx"

MiscAmount = m_clsReceivable.TotalAmount

'Check the Date of transaction
If Not DateValidate(txtRepayDate.Text, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If MsgBox(GetResourceString(38) & " " & _
        GetResourceString(37) & " = " & txtRepayDate.Text & _
        vbCrLf & GetResourceString(541), _
        vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
         
transType = wContraWithdraw
TransDate = GetSysFormatDate(Me.txtRepayDate.Text)


'Now check whether the amount collecting is
'Deducting or addinng to the Loan head or not
Dim IsAddingToAccount As Boolean
Dim I As Integer
Dim OtherTransType As wisTransactionTypes

IsAddingToAccount = m_clsReceivable.IsAddToAccount

'If all the amount is adding to the Account then
OtherTransType = wContraDeposit

'now get the Balance and Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Top 1 Balance,TransID,TransDate " & _
            " FROM BKCCTrans WHERE Loanid = " & m_LoanID & _
            " ORDER BY TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Balance = Val(FormatField(rst("Balance")))
    TransID = Val(FormatField(rst("TransID")))
End If

'Check the Transaction Date
'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, GetKCCLastTransDate(m_LoanID)) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo ExitLine
End If

'Begin the transactionhere
gDbTrans.BeginTrans
inTransaction = True

'First insert then Amount ot the Accounthead
TransID = GetKCCMaxTransID(m_LoanID) + 1

If IsAddingToAccount Then
    Deposit = IIf(Balance < 0, True, False)
    'If customer is paying the amount to the bank,reduece the balance
    Balance = Balance + m_clsReceivable.TotalAmount
    
    strParticulars = GetResourceString(327)
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans " _
                & "(LoanID, TransID, TransType, Amount, " _
                & " TransDate, Balance, Deposit,VoucherNO,UserId, Particulars) " _
                & "VALUES (" & m_LoanID & ", " & TransID & ", " _
                & transType & ", " & Amount & ", " _
                & "#" & TransDate & "#, " _
                & Balance & ", " _
                & Deposit & ", " _
                & AddQuotes(VoucherNo) & "," & UserID & "," _
                & AddQuotes(strParticulars) & " )"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    'Now get the Contra ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID + 1
    gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                    "( ContraID,AccHeadId,AccID," & _
                    " TransType,TransID,Amount," & _
                    " UserID,VoucherNo) VALUES (" & _
                    ContraID & "," & _
                    m_LoanHeadID & "," & m_LoanID & _
                    transType & "," & TransID & "," & _
                    m_clsReceivable.TotalAmount & "," & UserID & "," & _
                    AddQuotes(VoucherNo) & ")"
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
End If

'Now Get the pending Balance From the Previous account
Balance = 0
gDbTrans.SqlStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & m_LoanHeadID & _
        " ANd AccId = " & m_LoanID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    rstTemp.MoveLast
    Balance = FormatField(rstTemp("Balance"))
End If

I = 0
Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
Do While MiscHeadId > 0
    If MiscHeadId = 0 Then Exit Do
    
    If IsAddingToAccount Then
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(MiscHeadId) & "," & MiscHeadId & _
                        OtherTransType & "," & TransID & "," & _
                        MiscAmount & "," & UserID & "," & _
                        AddQuotes(VoucherNo) & ")"
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        'Insert the Same Into Head Transaction class
        If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                        MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
    End If
    
    Balance = Balance + MiscAmount
    
    'Now insert this details Into amount receivable table
    'If Not AddToAmountReceivable   Then GoTo ExitLine
    If Not AddToAmountReceivable(m_LoanHeadID, m_LoanID, _
                TransID, TransDate, MiscAmount, MiscHeadId) Then GoTo ExitLine
    I = I + 1
    Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
Loop

gDbTrans.CommitTrans

inTransaction = False

Call ResetTransactionForm


Exit Sub

ExitLine:
    
    If inTransaction Then gDbTrans.RollBack
    Exit Sub
    
Err_line:

    MsgBox "Error in Misc transaction", vbInformation, wis_MESSAGE_TITLE
 
End Sub

Private Sub LoadChequeNos()

If m_LoanID = 0 Then Exit Sub

Debug.Assert m_LoanID <> 0

Dim RstCheque As ADODB.Recordset

'Get the Loan Balance
'Now Check whether to load cheque book of deposit
'If Balance is Positive then load Loan cheque book
'If Balance is negative then load Deposit Cheque
'If Balance is 0 then load all the cheques

gDbTrans.SqlStmt = ""
If Val(txtBalance.Tag) >= 0 Then
    'Get the cheque details
    gDbTrans.SqlStmt = "Select AccID,Trans from ChequeMaster" & _
        " Where AccID = " & m_LoanID & _
        " AND AccHeadID = " & m_LoanHeadID
End If

If m_DepHeadID <> m_LoanHeadID Then
    If Val(txtBalance.Tag) <= 0 Then
        If Len(gDbTrans.SqlStmt) Then
            gDbTrans.SqlStmt = "Select AccID,Trans from ChequeMaster" & _
                " Where AccID = " & m_LoanID & _
                " AND (AccHeadID = " & m_LoanHeadID & " OR AccHeadID = " & m_DepHeadID & ")"
        Else
            gDbTrans.SqlStmt = "Select AccID,Trans from ChequeMaster" & _
                    " Where AccID = " & m_LoanID & _
                    " AND AccHeadID = " & m_DepHeadID
        End If
    End If
End If

cmbCheque.Clear
cmbCheque.AddItem ""
If gDbTrans.Fetch(RstCheque, adOpenForwardOnly) <= 0 Then
    cmbCheque.Tag = "InVisible"
    Set RstCheque = Nothing
    Exit Sub
End If

With cmbCheque
    .Tag = "Visible"
    .Clear
    If Not RstCheque Is Nothing Then
        .AddItem ""
        While Not RstCheque.EOF
            If FormatField(RstCheque("Trans")) = wischqIssue Then _
                        .AddItem RstCheque("ChequeNO")
            RstCheque.MoveNext
        Wend
    End If
End With

End Sub

Private Function LoanTransaction() As Boolean

On Error GoTo Err_line
 
'Dim LastTransDate As Date
Dim TransID As Long
Dim inTransaction As Boolean

Dim NewBalance As Currency
Dim Balance As Currency
Dim ContraTrans As Boolean

Dim DepInt As Currency
Dim RegInt As Currency
Dim PenalInt As Currency
Dim MiscAmount As Currency
Dim DueAmount As Currency
Dim PrincAmount  As Currency
Dim DepAmount As Currency
Dim IntBalance As Currency
Dim PenalIntBalance As Currency

Dim TotalAmount As Currency
Dim transType As wisTransactionTypes
Dim SelectedTransType As wisTransactionTypes
Dim TransDate As Date
Dim Deposit As Boolean
Dim rst As Recordset
Dim Amount As Currency
    
Dim SBClass As clsSBAcc
Dim bankClass As clsBankAcc
Dim MiscHeadId As Long

Dim VoucherNo As String
Dim ChequeNo  As String
Dim UserID As Integer
Dim strParticulars As String


'Get the Voucher No and Cheque No
UserID = gCurrUser.UserID
If cmbCheque.Visible Then
    If cmbCheque.ListIndex >= 0 Then ChequeNo = cmbCheque.Text
    VoucherNo = Trim(cmbCheque.Text)
Else
    VoucherNo = Trim(txtVoucherNo.Text)
End If

' -----------------------------------
'Get the Interest Balance as on TransCtion Date
gDbTrans.SqlStmt = "SELECT Top 1 * From BKCCIntTrans " & _
            " where LoanId = " & m_LoanID & " ORDER By TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    IntBalance = FormatField(rst("IntBalance"))
    PenalIntBalance = FormatField(rst("PenalIntBalance"))
End If

'Regular & penal interest Of the loan
RegInt = txtRegInterest
PenalInt = txtPenalInterest
MiscAmount = txtMiscAmount 'txtPenalInt = txtPenalInterest.Value

If txtRegInterest.Tag < 0 Then
    'in case the interest is negative it means
    'that interest is of the deposit amount
    'so make the regular interest as zero
    'consider the amount as deposit intereset
    DepInt = txtRegInterest
    RegInt = 0
End If

TransDate = GetSysFormatDate(txtRepayDate.Text)

TotalAmount = txtTotal
PrincAmount = txtRepayAmt

'if he  is adding the interest to the bkcc Deposit then
'do that transcctin
If cmbTrans.ListIndex = 2 Then
    'Add the Interest to the deposit
    LoanTransaction = AddInterestToDeposit
    GoTo Exit_Line
    'Exit Function
End If

transType = IIf(cmbTrans.ListIndex = 0, wDeposit, wWithdraw)
SelectedTransType = transType

If TotalAmount = 0 And PrincAmount = MiscAmount Then
    'if the transaction amount is only the misc amount
    'then check for the Misclenaeous heads are exists or not
    'If not then warn the user and exit
    If m_clsReceivable Is Nothing Then
        MsgBox "Unable to do this transaction", vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    'If m_clsReceivable.Count = 0 Then Err.Raise vbObjectError + 1, "", "Unable to transact"
    LoanTransaction = TransactMiscAmount
    GoTo Exit_Line
    'Exit Function
End If

'If he is withdrawingthe amount from the Loan Account
'Check for the due amount if it is there
DueAmount = GetReceivAbleAmount(m_LoanHeadID, m_LoanID)
If (DueAmount > 0 And m_clsReceivable Is Nothing) Or (DueAmount > 0 And DueAmount <> MiscAmount) Then
    If MsgBox("This account holder has to pay due amount" & _
            vbCrLf & GetResourceString(541), vbQuestion + _
            vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If

If (transType = wWithdraw) And (TotalAmount = RegInt + PenalInt + MiscAmount) Then
    transType = wContraWithdraw
    If RegInt > 0 Or PenalInt > 0 Then
      If chkSb.Value = vbChecked Then
          If MsgBox("No amount will be transferred to SBAccount. " & _
              "Paid amount will be received back as interest" & _
              vbCrLf & "Do you want to continue?" _
              , vbYesNo, wis_MESSAGE_TITLE) = vbNo Then GoTo Exit_Line
          chkSb.Value = vbUnchecked
      Else
          If MsgBox("No amount will given to the customer " & _
              "Paid amount will be received back as interest" & vbCrLf & _
              "Do you want to continue?", vbYesNo, _
              wis_MESSAGE_TITLE) = vbNo Then GoTo Exit_Line
      End If
    Else
      chkSb.Value = vbUnchecked
    End If
End If

'If transaction is payment & amount is crediting to the sb account
If chkSb.Value = vbChecked And SelectedTransType = wWithdraw Then
    Dim SbACCID As Long
    Dim ContraID As Integer
    Dim SbHeadID As Long

    Set SBClass = New clsSBAcc
    
    SbHeadID = GetIndexHeadID(GetResourceString(421))
    'Check For the SB Number Existace
    SbACCID = SBClass.GetAccountID(txtSbAccNum)
    If Len(Trim(txtSbAccNum)) = 0 Or SbACCID = 0 Then
        'MsgBox "Account not existance", vbInformation, "SB Account"
        MsgBox GetResourceString(421, 525), _
            vbInformation, "SB Account"
        GoTo Exit_Line
    End If
    
    Set SBClass = Nothing
    'Get the MAx Contra TransID
    ContraID = GetMaxContraTransID + 1
    
    transType = wContraWithdraw
End If

If PrincAmount < 0 Then
    'Invalid Amount Specified
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayAmt
    Exit Function
End If

If TotalAmount < Val(txtIntBalance.Tag) And Val(txtIntBalance.Tag) > 0 Then
    'If MsgBox("The amount he is paying is less than his Previous Interest Balance " & _
        vbCrLf & " Do you want to continue with transaction", vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    If MsgBox(GetResourceString(670) & _
        vbCrLf & GetResourceString(541), vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If

'Calculate the RegInt & Penal Int on the Specified date
Dim Cal_RegInt As Currency
Dim Cal_PenalInt As Currency
Dim IntBalance_pr As Currency

Cal_RegInt = BKCCRegularInterest(TransDate, m_LoanID) \ 1
Cal_PenalInt = BKCCPenalInterest(TransDate, m_LoanID) \ 1
IntBalance_pr = Cal_RegInt + Cal_PenalInt

'now get the Interest Balance difference
'Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Balance,TransID,TransDate " & _
            " FROM BKCCTrans WHERE loanid = " & m_LoanID & _
            " ORDER BY TransID Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
            Balance = Val(FormatField(rst("Balance")))

Deposit = IIf(Balance < 0, True, False)

'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, GetKCCLastTransDate(m_LoanID)) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo Exit_Line
End If

If RegInt > 0 Or PenalInt > 0 Then
    If (transType = wDeposit Or transType = wContraDeposit) Then
        'If he is repaying the amount then check whether the interest
        'he is paying is equal to the actual interest accrued
        'Or he is collecting the more/less interest
        'If he collecting more/less interest then keep it as balance
        If (RegInt - 1 > (Cal_RegInt + IntBalance)) Or _
            (PenalInt - 1 > (Cal_PenalInt + PenalIntBalance)) Then
           'If MsgBox("The amount he is paying is more than his Interest till date " & _
                vbCrLf & " Do you want to keep this Extra amount as advance interest", vbYesNo + vbQuestion + _
                vbDefaultButton1, wis_MESSAGE_TITLE) = vbNo Then
            If MsgBox(GetResourceString(666) & _
                vbCrLf & GetResourceString(667), vbYesNo + vbQuestion + _
                vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
                    'IntBalance = 0: PenalIntBalance = 0
            Else
                IntBalance = IntBalance + Cal_RegInt - txtRegInterest
                PenalIntBalance = PenalIntBalance + Cal_PenalInt - txtPenalInterest
            End If
        ElseIf (RegInt < Cal_RegInt Or PenalInt < Cal_PenalInt) Then
           'If MsgBox("The amount he is paying is less than his Interest till date " & _
                vbCrLf & " Do you want to keep this difference amount as interest balance ", vbYesNo + vbQuestion + _
                vbDefaultButton1, wis_MESSAGE_TITLE) = vbNo Then
            If MsgBox(GetResourceString(668) & _
                vbCrLf & GetResourceString(669), vbYesNo + vbQuestion + _
                vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
                    'IntBalance = 0: PenalIntBalance = 0
            Else
                IntBalance = IntBalance + Cal_RegInt - txtRegInterest
                PenalIntBalance = PenalIntBalance + Cal_PenalInt - txtPenalInterest
            End If
        End If
    
    Else
        'If he is withdrawing the amount
        'and He is collecting the more/less interest
        If Abs(RegInt - (Cal_RegInt + IntBalance)) > 1 Or _
            Abs(PenalInt - (Cal_PenalInt + PenalIntBalance)) > 1 Then
            MsgBox "The interest amount collecting is more or less than his Interest till date" & _
                vbCrLf & "so this transaction is not possible", vbInformation, wis_MESSAGE_TITLE
            'MsgBox GetResourceString(666) & _
                vbCrLf & GetResourceString(667), vbInformation, wis_MESSAGE_TITLE
            Exit Function
        End If
    End If
End If

'Update the Regular & Penal interest amount.
TransID = GetKCCMaxTransID(m_LoanID) + 1
transType = IIf(cmbTrans.ListIndex = 0, wDeposit, wWithdraw)
If transType = wWithdraw And SbACCID Then transType = wContraWithdraw

Set bankClass = New clsBankAcc

'BEGIN THE TRANSACTION IN DATABASE
If Not gDbTrans.BeginTrans Then GoTo Exit_Line
inTransaction = True

If RegInt > 0 Or PenalInt > 0 Or MiscAmount > 0 Then
    
    MiscHeadId = parIncome + 1 'Misceleneous
    If MiscAmount > 0 Then
        MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), parBankIncome, 0, wis_None)
    End If
    
    transType = wDeposit
    If (SelectedTransType = wWithdraw) And _
        TotalAmount = RegInt + PenalInt + MiscAmount Then transType = wContraDeposit
        
        
    Deposit = False
    'strParticulars = GetResourceString(47)
    strParticulars = cmbParticulars.Text
    gDbTrans.SqlStmt = "INSERT INTO BKCCIntTrans " & _
                " (LoanID, TransID,TransDate, TransType," & _
                " IntAmount, PenalIntAmount, MiscAmount, " & _
                " IntBalance,Deposit, VoucherNo,UserID,Particulars ) " _
            & "VALUES (" & m_LoanID & ", " & TransID & ", " _
            & " #" & TransDate & "#," & transType & ", " _
            & RegInt & ", " & PenalInt & "," _
            & MiscAmount & ", 0 ," & Deposit & "," _
            & AddQuotes(VoucherNo) & "," & UserID & "," _
            & AddQuotes(strParticulars) & " )"
            
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    gDbTrans.SqlStmt = "UPDATE BKCCmaster SET [LastIntDate] = " & _
                    "#" & TransDate & "# WHERE Loanid = " & m_LoanID
    ' Execute the updation.
    If RegInt + PenalInt > 0 Then If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If TotalAmount = RegInt + PenalInt + MiscAmount And (SelectedTransType = wWithdraw) Then
        If Not bankClass.UpdateContraTrans(m_LoanHeadID, m_RegIntHeadID, RegInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UpdateContraTrans(m_LoanHeadID, m_PenalHeadID, PenalInt, TransDate) Then GoTo Exit_Line
        If MiscAmount > 0 Then
            If m_clsReceivable Is Nothing Then
                If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                        MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
            Else
                Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
                Do While MiscHeadId > 0
                    If MiscHeadId = 0 Then Exit Do
                    If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                                MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
                    'do the same transaction in amountReceivable amount
                    If Not RemoveFromAmountReceivable(m_LoanHeadID, m_LoanID, _
                            TransID, TransDate, MiscAmount, MiscHeadId) Then GoTo Exit_Line
                    
                    Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
                Loop
            End If
        End If
        If PrincAmount = 0 Then PrincAmount = TotalAmount
    Else
        If Not bankClass.UpdateCashDeposits(m_RegIntHeadID, RegInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UpdateCashDeposits(m_PenalHeadID, PenalInt, TransDate) Then GoTo Exit_Line
'        If Not BankClass.UpdateCashDeposits(MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
        If MiscAmount > 0 Then
            If (RegInt + PenalInt) = 0 And Not m_clsReceivable Is Nothing Then
                'This means he is paying only the amount receivable from him
                'so check whether he has any amount due as receivable
                'if not then don't do this transaction
                If GetReceivAbleAmount(m_LoanHeadID, m_LoanID) = 0 Then
                    MsgBox "Unable to do this transaction", vbInformation, wis_MESSAGE_TITLE
                    GoTo Exit_Line
                End If
            End If
        
            If m_clsReceivable Is Nothing Then
                If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                            MiscAmount, TransDate) Then GoTo Exit_Line
            Else
                Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
                Do While MiscHeadId
                    If MiscHeadId = 0 Then Exit Do
                    If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                             MiscAmount, TransDate) Then GoTo Exit_Line
                    'do the same transaction in amountReceivable account
                    If Not RemoveFromAmountReceivable(m_LoanHeadID, m_LoanID, _
                            TransID, TransDate, MiscAmount, MiscHeadId) Then GoTo Exit_Line
                    
                    Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
                Loop
            End If
        End If
    End If

ElseIf DepInt > 0 Then
    transType = wContraWithdraw
    Deposit = True
    strParticulars = GetResourceString(47)
    gDbTrans.SqlStmt = "INSERT INTO BKCCIntTrans (LoanID, TransID," _
            & " TransDate, TransType, IntAmount, PenalIntAmount,  " _
            & "MiscAmount, IntBalance,Deposit, VoucherNo,UserID,Particulars ) " _
            & "VALUES (" & m_LoanID & ", " & TransID & ", " _
            & " #" & TransDate & "#," & transType & ", " _
            & DepInt & ", " & PenalInt & "," _
            & MiscAmount & ", 0 ," & Deposit & "," _
            & AddQuotes(VoucherNo) & "," & UserID & "," _
            & AddQuotes(strParticulars) & ")"
            
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    gDbTrans.SqlStmt = "UPDATE BKCCmaster SET [LastIntDate] = " & _
                    "#" & TransDate & "# WHERE loanid = " & m_LoanID
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    'if we are paying interest to the deposit amount
    'then the interest should credit to the respective account
    'So add the interest amount to deposit
    
    transType = wContraDeposit
    Deposit = (Balance < 0)
    Balance = Balance - DepInt
    strParticulars = cmbParticulars.Text
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID," _
                    & " TransDate, TransType,Amount," _
                    & " Balance,Deposit,VoucherNo,UserID, Particulars ) " _
                    & " VALUES (" & m_LoanID & ", " & TransID & ", " _
                    & " #" & TransDate & "#," & transType & ", " _
                    & DepInt & "," _
                    & Balance & "," & Deposit & "," _
                    & AddQuotes(VoucherNo) & "," & UserID & "," _
                    & AddQuotes(strParticulars) & ")"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not bankClass.UpdateContraTrans(m_DepIntHeadID, m_DepHeadID, DepInt, TransDate) Then GoTo Exit_Line
    
    'Add this amount to the Principle amount
    'if in case he is withdrawing the all the amount
    'from the deposit
    If cmbTrans.ListIndex > 0 Then PrincAmount = PrincAmount + DepInt
End If

' Update the principal amount.
'If this account is having loan and Customer Paying the amount
'More than that of loan balance OR account is having Deposit
'and Customer withdrwing the Amount More than that of balance
'In such cases we have to do
'two transaction one for the loan(deposit)  other one as deposit(loan)

'Get the whether the amount transaction is related to the Deposit OR LOan
Deposit = CBool(Balance < 0)
transType = IIf(cmbTrans.ListIndex = 0, wDeposit, wWithdraw)
If transType = wWithdraw And SbACCID Then transType = wContraWithdraw
    
'If customer is paying the amount to the bank,reduece the balance
If transType = wDeposit Or transType = wContraDeposit Then
    NewBalance = Balance - PrincAmount
Else
    NewBalance = Balance + PrincAmount
End If

'if customer is having negtive balance then he is antaining the deposit account
'So the amounthe trascating the deposit amount
Amount = PrincAmount
If Balance < 0 Then DepAmount = PrincAmount: PrincAmount = 0
If Balance = 0 And (transType = wDeposit Or transType = wContraDeposit) Then DepAmount = PrincAmount: PrincAmount = 0

'Now Check whetehr There are two transaction will take place
'If He is withdrwaing the amount more than his deposit balance then
'He has to make two transaction one for deposit amount as trans in deposit '
'next one for Loan trans
'If He is Depositing the amount more than his Loan balance then also two transac
Dim TwoTransactions As Boolean
TwoTransactions = (Balance > 0 And NewBalance < 0) Or (Balance < 0 And NewBalance > 0)

'if there is no transaction in Deposit
' then there only one transaction
'If TwoTransactions Then TwoTransactions = m_KCCDepositTrans

If TwoTransactions Then
    'Then Check whether user allowed to do two transactions at a time
    Dim SetUp As New clsSetup
    If UCase(SetUp.ReadSetupValue("General", "KCCTwoTransaction", "True")) = "TRUE" Then
        MsgBox "You can not make this transaction now" & vbCrLf & _
            "Please make saperate transaction for Loan & deposit" _
            , vbInformation, wis_MESSAGE_TITLE
        gDbTrans.RollBack
        Exit Function
    End If
    
    'suppose we have to do two transaction
    'Then do the first transaction here
    If Balance < 0 Then
        PrincAmount = Abs(NewBalance)
        DepAmount = Abs(Balance)
    Else
        PrincAmount = Abs(Balance)
        DepAmount = Abs(NewBalance)
    End If
    'NOw Ge the amount to be transact in the next transaction
    Amount = Abs(NewBalance)
    'strParticulars = GetResourceString(IIf(Balance > 0, 58, 43))
    strParticulars = cmbTrans.Text
    strParticulars = IIf(transType = wDeposit Or transType = wContraDeposit, "By ", "To ") & "Cash"
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans" _
            & " (LoanID, TransID, TransType, Amount," _
            & " TransDate, Balance,Deposit,VoucherNo,UserID,Particulars) " _
            & "VALUES (" & m_LoanID & ", " & TransID & ", " _
            & transType & "," & Abs(Balance) & "," _
            & "#" & TransDate & "#, " _
            & " 0 , " & Deposit & ", " _
            & AddQuotes(VoucherNo) & "," & UserID & "," _
            & AddQuotes(strParticulars) & ")"
    
    'Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If transType = wContraWithdraw Then
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            " (ContraID, AccHeadID,AccID,TransType,TransID,Amount,UserID )" & _
            " VALUES (" & ContraID & "," & _
            IIf(Deposit, m_DepHeadID, m_LoanHeadID) & "," & _
            m_LoanID & "," & transType & ", " & _
            TransID & "," & Abs(Balance) & "," & gUserID & " )"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        Set SBClass = New clsSBAcc
        If SBClass.DepositAmount(SbACCID, Abs(Balance), _
                "from KCC AccNo " & txtAccNo.Text, TransDate, " ") = 0 Then GoTo Exit_Line
        If Not bankClass.UpdateContraTrans(IIf(Deposit, m_DepHeadID, m_LoanHeadID), _
                SbHeadID, Abs(Balance), TransDate) Then GoTo Exit_Line
    End If
    
    Deposit = Not Deposit
End If

If NewBalance < 0 Then Deposit = True
If NewBalance > 0 Then Deposit = False
'If Not m_KCCDepositTrans Then Deposit = False

'The Deposit Interest amount alrady added to the the deposit
'So if the difference is more than the deposit interest
'then only do this transction

transType = IIf(cmbTrans.ListIndex = 0, wDeposit, wWithdraw)
If transType = wWithdraw And chkSb.Value = vbChecked Then transType = wContraWithdraw

If PrincAmount Or DepAmount Then
    If TotalAmount = RegInt + PenalInt + MiscAmount Then transType = wContraWithdraw
    
    'Now Get the Amount he is transacting
    'If Amount = 0 Then Amount = IIf(NewBalance > 0, PrincAmount, DepAmount)
    'Amount = PrincAmount
    'strParticulars = GetResourceString(IIf(Balance > 0, 58, 43))
    strParticulars = cmbParticulars.Text
    'strParticulars = IIf(TransType = wDeposit Or TransType = wContraDeposit, "By ", "To ") & "Cash"
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans " _
                & "(LoanID, TransID, TransType, Amount, " _
                & " TransDate, Balance, Deposit,VoucherNO,UserId, Particulars) " _
                & "VALUES (" & m_LoanID & ", " & TransID & ", " _
                & transType & ", " & Amount & ", " _
                & "#" & TransDate & "#, " _
                & NewBalance & ", " & Deposit & "," _
                & AddQuotes(VoucherNo) & "," & UserID & "," _
                & AddQuotes(strParticulars) & " )"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    If chkSb.Value = vbChecked And transType = wContraWithdraw Then
        'The user debiting the withdrawn amount ot his Sb account then
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                " (ContraID, AccHeadID,AccId," & _
                " TransType,TransID,Amount,UserID )" & _
                " VALUES (" & ContraID & "," & _
                IIf(Deposit, m_DepHeadID, m_LoanHeadID) & "," & _
                m_LoanID & "," & transType & ", " & _
                TransID & "," & Amount & "," & gUserID & " )"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
        'Post the Deposit Transaction to SB account
        If SBClass Is Nothing Then Set SBClass = New clsSBAcc
        If SBClass.DepositAmount(SbACCID, Amount, _
                "from Kcc AccNo " & txtAccNo, TransDate, " ") = 0 Then GoTo Exit_Line
        
        'Update the Contra transaction details to the ContrTrans table
        If Not bankClass.UpdateContraTrans(IIf(Deposit, m_DepHeadID, m_LoanHeadID), _
                SbHeadID, PrincAmount, TransDate) Then GoTo Exit_Line
        
        If Not m_clsReceivable Is Nothing Then
            Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
            Do While MiscHeadId
                If MiscHeadId = 0 Then Exit Do
                gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                        " (ContraID, AccHeadID,AccId," & _
                        " TransType,TransID,Amount,UserID )" & _
                        " VALUES (" & ContraID & "," & _
                        GetParentID(MiscHeadId) & "," & _
                        MiscHeadId & "," & wContraDeposit & ", " & _
                        0 & "," & Amount & "," & gUserID & " )"
                If Not gDbTrans.SQLExecute Then GoTo Exit_Line
                
                If Not bankClass.UpdateContraTrans(IIf(Deposit, m_DepHeadID, m_LoanHeadID), _
                                MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
                
                Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
            Loop
        End If
    
    Else
        If transType = wWithdraw Then
            If Not bankClass.UpdateCashWithDrawls(m_DepHeadID, _
                        DepAmount, TransDate) Then GoTo Exit_Line
            If Not bankClass.UpdateCashWithDrawls(m_LoanHeadID, _
                        PrincAmount, TransDate) Then GoTo Exit_Line
        ElseIf transType = wDeposit Then
            If Not bankClass.UpdateCashDeposits(m_DepHeadID, _
                        DepAmount, TransDate) Then GoTo Exit_Line
            If Not bankClass.UpdateCashDeposits(m_LoanHeadID, _
                        PrincAmount, TransDate) Then GoTo Exit_Line
        End If
    End If
End If
    
'If NewBalance = 0 And Balance > 0 Then cmdClose.Visible = True

' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.

'    gDbTrans.SQLStmt = "UPDATE Bkccmaster SET loanclosed = 1, " & _
'                "InterestBalance = 0 Where LoanID =  " & m_LoanID
'    ' Execute the updation.
'    If Not gDbTrans.SQLExecute Then GoTo Exit_line
'End If
   
'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(IIf(Balance >= 0, m_LoanHeadID, m_DepHeadID), _
            m_LoanID, TransDate, TransID, (PrincAmount - RegInt - PenalInt)) < 1 Then GoTo Exit_Line
    
End If
'if he has withdrawn from the Cheque book
'then Mark the cheque in cheque book as issued
If Len(ChequeNo) Then
    gDbTrans.SqlStmt = "Update ChequeMaster Set Trans = " & wischqPay & "," & _
                "TransDate = #" & TransDate & "#, Amount = " & Amount & "," & _
                "Particulars = " & AddQuotes(strParticulars, True) & _
                " Where AccID = " & m_LoanID & " AND AccHeadID = " & m_LoanHeadID & _
                " AND ChequeNo = " & ChequeNo
    'Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
End If


' Commit the transaction.
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False

LoanTransaction = True
'MsgBox "Loan repayment accepted.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(706), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    Set SBClass = Nothing
    Set Cashclass = Nothing
    If inTransaction Then gDbTrans.RollBack
    'Writ the Particulars
    Call WriteParticularstoFile(cmbParticulars.Text, App.Path & "\BKCC.ini")
    
    'Load the Particulars
    Call LoadParticularsFromFile(cmbParticulars, App.Path & "\BKCC.ini")
    
    Exit Function


Err_line:
    If Err Then
        MsgBox "LoanRepay: " & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(707) & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
    End If
    GoTo Exit_Line

End Function

Private Sub ArrangeLoanIssuePropSheet()

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


If NeedsScrollbar Then vscLoanIssue.Visible = True

For I = 0 To txtLoanIssue.count - 1
    txtLoanIssue(I).Width = picLoanIssueSlider.ScaleWidth - _
        (lblLoanIssue(I).Left + lblLoanIssue(I).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For I = 0 To cmb.count - 1
    cmb(I).Width = txtLoanIssue(I).Width
Next
For I = 0 To cmdLoanIssue.count - 1
    cmdLoanIssue(I).Left = txtLoanIssue(I).Left _
        + txtLoanIssue(I).Width - cmdLoanIssue(I).Width
Next

'' Draw lines for the remaining portions of the viewport.
'With picLoanIssueViewPort
'    .CurrentX =


End Sub
' Fills in the details of the Account holder to the form.
Private Sub FillAcHolderDetails()
#If COMMENTED Then
On Error GoTo Err_line
Dim I As Integer
Dim strField As String

' Extract the nominee info from the field.
Dim NomineeInfo() As String
If Not IsNull(gDbTrans.rst("Nominee")) Then _
    GetStringArray FormatField(gDbTrans.rst("Nominee")), NomineeInfo(), ";"
If UBound(NomineeInfo) = 0 Then
    ReDim NomineeInfo(2)
    NomineeInfo(0) = " "
    NomineeInfo(1) = " "
    NomineeInfo(2) = " "
End If

' fill in the details of ac-holder.
For I = 0 To txtPrompt.count - 1
    ' Read the bound field of this control.
    strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If strField <> "" Then
        With txtData(I)
            Select Case UCase$(strField)
                Case "ACCID"
                    .Text = gDbTrans.rst("Accid")
                Case "ACNAME"
                    .Text = m_AccHolder.FullName
                Case "NOMINEE_NAME"
                    .Text = NomineeInfo(0)
                Case "NOMINEE_AGE"
                    .Text = NomineeInfo(1)
                Case "NOMINEE_RELATION"
                    .Text = NomineeInfo(2)
                Case "JOINTHOLDER"
                    .Text = gDbTrans.rst("JointHolder")
                Case "INTRODUCERID"
                    .Text = IIf(gDbTrans.rst("Introduced") = 0, "", gDbTrans.rst("Introduced"))
                Case "INTRODUCERNAME"

                Case "LEDGERNO"
                    .Text = gDbTrans.rst("LedgerNo")
                Case "FOLIONO"
                    .Text = gDbTrans.rst("FolioNO")
                Case "CREATEDATE"
                    .Text = FormatField(gDbTrans.rst("CreateDate"))
            End Select
        End With
    End If
Next


Err_line:
    If Err.Number = 9 Then  'Subscript out of range.
        Resume Next
    ElseIf Err Then
        MsgBox "FillAcHolderDetails: " & vbCrLf _
                & Err.Description, vbCritical
        'MsgBox GetResourceString(709) & vbCrLf _
                & Err.Description, vbCritical
    End If

#End If
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
                    rs.Fields(I).Type = adLongVarChar Or _
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
Public Function HasOverdueLoans(CustomerID As Long, Optional AsOnDate As Date) As Boolean
' Variables...
Dim Lret As Long

' Check how many loans this member has taken.
gDbTrans.SqlStmt = "SELECT A.loanID,Balance,TransDate FROM BKCCMaster A, " & _
    " BKCCTrans B WHERE CustomerID = " & CustomerID & _
    " AND A.LoanID = B.LoanID ANd Transid = (Select Max(TransID) From " & _
        " BKCCTrans C WHERE C.LoanID = A.LoanID " & _
        " And TransDate < #" & DateAdd("yyyy", -1, AsOnDate) & "# )"
        
Dim LoansRS As Recordset
Dim rst As Recordset

If gDbTrans.Fetch(LoansRS, adOpenDynamic) < 1 Then Exit Function
' No loans taken previously. Return False.


Dim Balance As Currency
Dim RepaidAmount As Currency


' Get the balances remaining for each of the above loans...
Do While Not LoansRS.EOF
    Balance = Val(FormatField(LoansRS("Balance")))
    If Balance > 0 Then
        ' Check the due date.
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as RepaidAmount " & _
            " FROM BKCCTrans WHERE loanID = " & LoansRS("LoanID") & _
            " AND TransDate > #" & LoansRS("TransDate") & "# " & _
            " AND TransDate <= #" & AsOnDate & "# " & _
            " AND TransType > 0 "
        
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            Balance = Balance - FormatField(rst("RepaidAmount"))
            If Balance > 0 Then
                HasOverdueLoans = True
                Exit Do
            End If
        End If
    End If
    ' Move to next record.
    LoansRS.MoveNext
Loop

End Function

Public Function LoanLoad(ByVal lLoanID As Long) As Boolean
On Error GoTo LoanLoad_Error

' Declare variables needed for this procedure...
Dim Lret As Long
Dim transType As wisTransactionTypes
Dim MemClass As New clsMMAcc
Dim MaxTransID As Long
Dim IntAmount As Currency
Dim txtIndex As Byte
Dim CtrlIndex As Byte
Dim InstalmentMode As Byte
Dim rstTemp As Recordset
Dim TransDate As Date
Dim I As Integer
Dim SqlStr As String
Dim memberTYpe As Integer
Dim MemberName As String


' Get the loan details from the LoanMaster table.
If m_LoanID = lLoanID Then
    Call ResetTransactionForm
Else
    Call ResetUserIntereface
End If

gDbTrans.SqlStmt = "SELECT * FROM BKCCMaster WHERE loanID = " & lLoanID

Lret = gDbTrans.Fetch(m_rstLoanMast, adOpenDynamic)
If Lret <= 0 Then GoTo Exit_Line

'm_LoanID = lLoanID

cmdLoanSave.Enabled = False
TransDate = vbNull
If DateValidate(txtRepayDate, "/", True) Then _
            TransDate = GetSysFormatDate(txtRepayDate.Text)
' Check wheather loanclosed or not
If FormatField(m_rstLoanMast("LoanClosed")) Then
    cmdLoanUpdate.Enabled = False
    cmdUndo.Caption = GetResourceString(313) '"Reopen"
Else
    cmdLoanUpdate.Enabled = True
    cmdUndo.Caption = GetResourceString(5) '"uNDO"
End If


'Get the total repaid amount on this loan.
'IF HE HAS ALREADY REPAID A LOAN COMPLTELY THEN WE HAVE TO
'FETCH THE RECORDS OF NEW LOAN ONLY

' Get the balance amount on this loan.
'Lret = GetKCCMaxTransID(CInt(lLoanID))
gDbTrans.SqlStmt = "SELECT Top 1 Balance,Transid " & _
                " FROM BKCCTrans WHERE" & _
                " loanID = " & lLoanID & _
                " Order By TransID Desc"
Lret = gDbTrans.Fetch(rstTemp, adOpenDynamic)

If Lret > 0 Then 'GoTo LoanLoad_Error
    rstTemp.MoveLast
    txtBalance.Caption = FormatCurrency(Abs(FormatField(rstTemp(0))))
    txtBalance.Tag = FormatCurrency(FormatField(rstTemp(0)))
    txtBalance.ToolTipText = txtBalance.Tag
    txtBalance.ForeColor = IIf(rstTemp(0) < 1, vbBlue, vbRed)
    
    SqlStr = "SELECT 'PRINCIPAL',Amount,Particulars ," & _
            " TransID,TransDate,TransType, Balance," & _
            " Deposit,VoucherNo FROM BKCCTrans WHERE loanID = " & lLoanID
    SqlStr = SqlStr & " UNION " & _
            " SELECT 'INTEREST',(IntAmount+PenalIntAmount) as Amount,MiscAmount," & _
            " TransID,TransDate,TransType, IntBalance as Balance," & _
            " Deposit,VoucherNo FROM BKCCIntTrans " & _
            " WHERE loanID = " & lLoanID
    gDbTrans.SqlStmt = SqlStr & " ORDER BY TransID"
    SqlStr = ""
    Lret = gDbTrans.Fetch(m_rstLoanTrans, adOpenDynamic)
    If Lret <= 0 Then GoTo LoanLoad_Error
    
    If m_rstLoanTrans.RecordCount = 1 Then _
                cmdUndo.Caption = GetResourceString(14) ' "Delete"
    m_rstLoanTrans.MoveLast
    m_TransID = m_rstLoanTrans("TransID")
    m_rstLoanTrans.MoveFirst
    If m_TransID > 10 Then
        m_TransID = m_TransID - (m_TransID Mod 10)
        m_rstLoanTrans.Find "TransID >= " & m_TransID
        If m_rstLoanTrans.EOF Then m_rstLoanTrans.MoveFirst
    End If
End If
'Load the Notes if he has any
With rtfNote
    '.BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    '.Enabled = IIf(ClosedDate = "", True, False)
    Call m_Notes.LoadNotes(wis_BKCCLoan, m_LoanID)
End With
Call m_Notes.DisplayNote(rtfNote)
Me.TabStrip2.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True

'Now Load the Transaction details of the account
Call LoanLoadGrid
cmdPrint.Enabled = True

'Now Get the Cheque details of the account
Call LoadChequeNos

'Get the Balance Interest if Any
txtIntBalance.Tag = FormatField(m_rstLoanMast("InterestBalance"))
txtIntBalance.Text = FormatCurrency(FormatField(m_rstLoanMast("InterestBalance")))
txtLoanIssue(Val(GetIndex("InterestBalance"))) = txtIntBalance.Text
txtLastInt = FormatField(m_rstLoanMast("LastIntDate"))

If TransDate <> vbNull Then
    
    ' Compute and display the regular interest.
    IntAmount = BKCCRegularInterest(TransDate, lLoanID)
    lblRegInterest = GetResourceString(IIf(IntAmount < 0, 233, 344))
    txtRegInterest.ToolTipText = IntAmount \ 1
    txtRegInterest.Tag = IntAmount \ 1
    txtRegInterest = Abs(IntAmount \ 1)
    
    ' Compute and display the penal interest for defaulted payments.
    IntAmount = BKCCPenalInterest(TransDate, lLoanID)
    txtPenalInterest.ToolTipText = IntAmount / 1

End If

txtPenalInterest = 0
If IntAmount >= 0 Then txtPenalInterest = IntAmount \ 1
'Display the total instalment amount.
txtRepayAmt = 0 'txtTotal.Text

'Get the Intrest Balance
txtIndex = GetIndex("InterestBalance")
'Now load the inttrestBalance to the registration Details.
Dim rstIntTrans As Recordset

gDbTrans.SqlStmt = "SELECT * from BkccIntTrans where Loanid=" & lLoanID & _
                   " And TransID = (Select Max(TransID) From " & _
                   " BKCCIntTrans Where LoanID = " & lLoanID & ")"

Call gDbTrans.Fetch(rstIntTrans, adOpenForwardOnly)
txtIndex = GetIndex("InterestBalance")
txtLoanIssue(txtIndex).Text = FormatField(rstIntTrans("intBalance"))


'CHECK FOR THE LOAID OF ALREADY LOADED AND THIS LOANID
'If The loanid Has Not changed then need not to load the below details
'BEACAUSE ALL THE SETAILSARE LOADED ONCE
If lLoanID = m_LoanID Then LoanLoad = True: Exit Function


'Fill The Customer Details
If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
m_CustReg.LoadCustomerInfo (m_rstLoanMast("CustomerID"))
txtMemberName = m_CustReg.FullName

' Fill the loan amount.
txtLoanAmt.Caption = Val(FormatField(m_rstLoanMast("CurrentSanction"))) + Val(FormatField(m_rstLoanMast("ExtraSanction")))

' Fill the loan issue date.
txtIssueDate.Caption = FormatField(m_rstLoanMast("IssueDate"))

'Get the Loan Account.
Call GetMemberNameNumberByCustID(m_rstLoanMast("CustomerID"), MemberName, memberTYpe)
txtIndex = GetIndex("AccNum")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("AccNum"))
txtLoanIssue(txtIndex + 1).Text = MemberName 'MemClass.memberName(FormatField(m_rstLoanMast("AccNum")))

txtIndex = GetIndex("MemberType")
CtrlIndex = Val(ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex"))
If cmb(CtrlIndex).ListCount > 1 Then
    Call SetComboIndex(cmb(CtrlIndex), , CLng(memberTYpe))
    txtLoanIssue(txtIndex).Text = cmb(CtrlIndex).Text
End If
'Check for member details.
txtIndex = GetIndex("MemberID")
Dim MemID As Long

MemID = FormatField(m_rstLoanMast("MemId"))
gDbTrans.SqlStmt = "Select AccNum from MemMaster Where AccID = " & MemID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) Then
    txtLoanIssue(txtIndex).Text = FormatField(rstTemp("AccNum"))
    txtLoanIssue(txtIndex + 1).Text = MemClass.MemberName(txtLoanIssue(txtIndex).Text)
End If

'Get the Farmer Type
txtIndex = GetIndex("FarmerType")
I = FormatField(m_rstLoanMast("FarmerType"))
CtrlIndex = Val(ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex"))
Call SetComboIndex(cmb(CtrlIndex), , CLng(I))
txtLoanIssue(txtIndex) = cmb(CtrlIndex).Text

Dim LoanLimit As Currency
'3 year sanction Amounts
txtIndex = GetIndex("SanctionAmount")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("SanctionAmount"))
'current year sanction Amounts
txtIndex = GetIndex("CurrentSanction")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("CurrentSanction"))
LoanLimit = Val(txtLoanIssue(txtIndex).Text)
'extra sanction Amounts
txtIndex = GetIndex("ExtraSanction")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("ExtraSanction"))
LoanLimit = LoanLimit + Val(txtLoanIssue(txtIndex).Text)
'Maximum LOan limit
txtIndex = GetIndex("TotalAmount")
txtLoanIssue(txtIndex).Text = FormatCurrency(LoanLimit)

' Loan Issue date. OR ' Create Date
txtIndex = GetIndex("IssueDate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("IssueDate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "IssueDate", FormatField(m_rstLoanMast("IssueDate")))

' Loan renew Date
txtIndex = GetIndex("RenewDate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("RenewDate"))

'Interest Rate
txtIndex = GetIndex("InterestRate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("IntRate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "InterestRate", FormatField(m_rstLoanMast("IntRate")))

'Penal Interest Rate
txtIndex = GetIndex("PenalInterestRate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("PenalIntRate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "PenalInterestRate", FormatField(m_rstLoanMast("PenalIntRate")))

'Deposit Interest Rate
txtIndex = GetIndex("DepositInterestRate")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("DepIntRate"))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "DepositInterestRate", FormatField(m_rstLoanMast("DepIntRate")))

' Guarantors ..
txtIndex = GetIndex("Guarantor1")
Lret = FormatField(m_rstLoanMast("Guarantor1"))
txtLoanIssue(txtIndex).Text = IIf(Lret, m_CustReg.CustomerName(Lret), "")
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "GuarantorID", CStr(Lret))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "CustomerID", CStr(Lret))

txtIndex = GetIndex("Guarantor2")
Lret = FormatField(m_rstLoanMast("Guarantor2"))
txtLoanIssue(txtIndex).Text = IIf(Lret, m_CustReg.CustomerName(Lret), "")
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "GuarantorID", CStr(Lret))
txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "CustomerID", CStr(Lret))

'Remarks
'txtindex = GetIndex("Remarks")
'txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("Remarks"))

'Dry Land
txtIndex = GetIndex("DryLand")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("DryLand"))

'Irrigatted land
txtIndex = GetIndex("WetLand")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("IrrigationLand"))

'Dry Land Income
txtIndex = GetIndex("DryIncome")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("DryIncome"))

'Irrigatted land
txtIndex = GetIndex("wetIncome")
txtLoanIssue(txtIndex).Text = FormatField(m_rstLoanMast("IrrigationIncome"))

'Here Check for the due amount if any
'If CheckForDueAmount > 0 Then
Dim DueAmount As Currency
DueAmount = GetReceivAbleAmount(m_LoanHeadID, lLoanID)
If DueAmount Then
    cmbTrans.ListIndex = 0
    txtMiscAmount = DueAmount
    txtTotal = txtMiscAmount
End If


LoanLoad = True
m_LoanID = lLoanID

m_dbOperation = Update

RaiseEvent AccountChanged(m_LoanID)

Exit_Line:
    Exit Function

LoanLoad_Error:
    If Err Then
        MsgBox "LoanLoad: " & Err.Description, vbCritical
        'Resume
        Err.Clear
    End If
'Resume
    GoTo Exit_Line
End Function

'
Private Sub LoanLoadGrid()
Dim transType As wisTransactionTypes
Dim Balance As Currency
Dim Amount As Currency

Err.Clear
On Error GoTo Err_line

' If no recordset, exit.
If m_rstLoanTrans Is Nothing Then Exit Sub
'If Not NextClick Then Rst.MoveLast

' If no records, exit.
If m_rstLoanTrans.RecordCount = 0 Then Exit Sub

'Show 10 records or till eof of the page being pointed to
With grd
    ' Initialize the grid.
    .Visible = False
    .Clear: .AllowUserResizing = flexResizeBoth
    .Cols = IIf(m_KCCDepositTrans, 10, 6)
    .Rows = 12
    .FixedRows = 1
    .FixedCols = 0
    .Font.Size = 11
    .Row = 0
    .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1150  'Date
    .Col = 1: .Text = GetResourceString(39): .ColWidth(1) = 1000  'Particulars
    .Col = 2: .Text = GetResourceString(272): .ColWidth(3) = 800  'Withdrawls
    .Col = 3: .Text = GetResourceString(271): .ColWidth(2) = 800  'Deposits
    .Col = 4: .Text = GetResourceString(47): .ColWidth(4) = 700  'Interest Received
    .Col = 5: .Text = GetResourceString(67, 58): .ColWidth(5) = 1100 'Balance
    If m_KCCDepositTrans Then
        .Col = 6: .Text = GetResourceString(271): .ColWidth(6) = 800  'Deposits
        .Col = 7: .Text = GetResourceString(272): .ColWidth(7) = 800  'Withdrawls
        .Col = 8: .Text = GetResourceString(47): .ColWidth(8) = 700  'Interest
        .Col = 9: .Text = GetResourceString(43, 42): .ColWidth(9) = 800 'Balance
    End If
    
    
    .ScrollBars = flexScrollBarBoth
    
    Dim TransID As Long
    .Row = .FixedRows
    TransID = m_rstLoanTrans("TransID")
    .Col = 0: .Text = FormatField(m_rstLoanTrans("TransDate"))

    Do
        If m_rstLoanTrans.EOF Then Exit Do
        transType = m_rstLoanTrans("TransType")
        If TransID <> m_rstLoanTrans("TransID") Then
            Amount = 0
            If .Row = 11 Then m_rstLoanTrans.MovePrevious: Exit Do
            .Row = .Row + 1
            TransID = m_rstLoanTrans("TransID")
            .Col = 0: .Text = FormatField(m_rstLoanTrans("TransDate"))
        End If
        If m_rstLoanTrans(0) = "INTEREST" Then
            .Col = 4
            If m_rstLoanTrans("Deposit") And m_KCCDepositTrans Then .Col = 8
            .Text = FormatField(m_rstLoanTrans("Amount"))
'            .CellForeColor = IIf(FormatField(m_rstLoanTrans("Deposit")), vbRed, vbBlack)
        Else
            If m_KCCDepositTrans Then Amount = 0
            Amount = Amount + FormatField(m_rstLoanTrans("Amount"))
            .Col = 1: .Text = FormatField(m_rstLoanTrans("Particulars"))
            '.Col = 1: .Text = FormatField(m_rstLoanTrans("VoucherNo"))
            If transType = wDeposit Or transType = wContraDeposit Then
                .Col = 3
            Else
                .Col = 2
            End If
            If m_rstLoanTrans("Deposit") And m_KCCDepositTrans Then .Col = .Col + 4
            .Text = FormatCurrency(Amount) 'FormatField(m_rstLoanTrans("Amount"))
            .Col = IIf(m_rstLoanTrans("Deposit") And m_KCCDepositTrans, 9, 5)
            .Text = FormatCurrency(Abs(m_rstLoanTrans("Balance")))
        End If
        
nextRecord:
        m_rstLoanTrans.MoveNext
        If m_rstLoanTrans.EOF Then Exit Do
    Loop
    .Visible = True
    .Row = 1
End With

cmdNextTrans.Enabled = True
cmdPrevTrans.Enabled = True
If m_rstLoanTrans.EOF Then cmdNextTrans.Enabled = False
m_rstLoanTrans.MoveFirst
TransID = m_rstLoanTrans("TransID")
If TransID >= m_TransID Then cmdPrevTrans.Enabled = False

cmdUndo.Enabled = gCurrUser.IsAdmin

Exit Sub

Err_line:
    If Err Then
        MsgBox "LoanLoadGrid: " & Err.Description
        'MsgBox GetResourceString(714) & Err.Description
        grd.Visible = True
        Err.Clear
    End If
'Resume
End Sub
Private Function LoanSave() As Boolean

'FIrst check the validity of all controls & required fields
If Not CheckValidity Then Exit Function

'Setup error handler.
On Error GoTo LoanSave_Error

' Declare variables for this procedure...
' ------------------------------------------
Dim txtIndex As Integer
Dim CtrlIndex As Integer
Dim inTransaction As Boolean
Dim Lret As Long, nRet As Integer
Dim NewLoanID As Long
Dim newTransID As Long
Dim Balance As Currency
Dim Amount As Currency
Dim count As Integer
Dim MemID As Long
Dim custId As Long
Dim SbACCID As Long
Dim rst As Recordset


' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)
    gDbTrans.SqlStmt = "SELECT * FROM MemMaster " & _
                " WHERE AccNum = " & AddQuotes(.Text, True)
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    MemID = FormatField(rst("AccId"))
    custId = FormatField(rst("CustomerID"))
End With


'Memeber type
'Now Get the Farmer Type
Dim FarmerType As Integer
txtIndex = GetIndex("FarmerType")
'Now get the index of the combo box which represents the farmer type
CtrlIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
FarmerType = cmb(CtrlIndex).ItemData(cmb(CtrlIndex).ListIndex)

' Get a new loanid.
gDbTrans.SqlStmt = "SELECT MAX(LoanID) FROM BKCCMaster"
Lret = gDbTrans.Fetch(rst, adOpenDynamic)
If Lret <= 0 Then GoTo Exit_Line
NewLoanID = Val(FormatField(rst(0))) + 1

' Begin the transaction.
gDbTrans.BeginTrans
inTransaction = True

' Put an entry into LoanMaster table.
gDbTrans.SqlStmt = "INSERT INTO BKCCMaster ( LoanID, AccNum," _
        & "MemID, CustomerID,IssueDate, " _
        & "SanctionAmount,CurrentSanction, ExtraSanction, " _
        & "RenewDate,Guarantor1, Guarantor2,FarmerType, " _
        & "IntRate,PenalIntRate,DepIntRate," _
        & "DryLand, IrrigationLand," _
        & "DryIncome, IrrigationIncome,Remarks,UserID )"

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " VALUES (" & _
        NewLoanID & ", " _
        & AddQuotes(Trim$(GetValue("AccNum")), True) & "," _
        & MemID & ", " & custId & ", " _
        & "#" & GetSysFormatDate(GetValue("IssueDate")) & "#, " _
        & Val(GetValue("SanctionAmount")) & ", " _
        & Val(GetValue("CurrentSanction")) & ", " _
        & Val(GetValue("ExtraSanction")) & ", " _
        & "#" & GetSysFormatDate(GetValue("RenewDate")) & "#, " _
        & PropGuarantorID(1) & ", " & PropGuarantorID(2) & ", " _
        & FarmerType & ", " _
        & Val(GetValue("InterestRate")) & "," _
        & Val(GetValue("PenalInterestRate")) & "," _
        & Val(GetValue("DepositInterestRate")) & "," _
        & AddQuotes(GetValue("DryLand")) & "," _
        & AddQuotes(GetValue("WetLand")) & "," _
        & Val(GetValue("DryIncome")) & "," _
        & Val(GetValue("WetIncome")) & "," _
        & AddQuotes(GetValue("Remarks"), True) _
        & "," & gUserID & ")"
    
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
' Commit the transaction
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False

'MsgBox "New loan Account created successfully.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(728), vbInformation, wis_MESSAGE_TITLE

Call LoanLoad(NewLoanID)

cmdLoanSave.Enabled = False
cmdLoanUpdate.Enabled = True

Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    
Exit Function

LoanSave_Error:
    If Err Then
        MsgBox "LoanSave: " & Err.Description, vbCritical
        'MsgBox GetResourceString(729) & Err.Description, vbCritical
        Err.Clear
    End If
'Resume
    GoTo Exit_Line
End Function

Private Sub LoanTabClear()

 txtLoanAmt.Caption = ""
 txtIssueDate.Caption = ""
' txtRepaidAmt.Caption = ""
 txtBalance.Caption = ""
 txtBalance.Tag = ""
 txtIntBalance.Text = ""
' txtInstAmt.Text = ""
 txtRegInterest = 0
 txtPenalInterest = 0
 txtTotal = 0
 txtRepayDate.Text = gStrDate
 txtRepayAmt = 0
 
 grd.Clear
 
' cmdRepay.Enabled = False
 Me.cmdUndo.Enabled = False
 ' Cler the Loan Issue Grid Also
 Call cmdLoanClear_Click
End Sub
' Reverts the last transaction in loan transactions table.
Private Function LoanUndoLastTransaction() As Boolean

If (gCurrUser.UserPermissions And perBankAdmin) = 0 Then
    'No permission
    MsgBox GetResourceString(685), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


'Variables of the procedure...
Dim lastTransID As Long
Dim inTransaction As Boolean
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes
Dim rst As Recordset

Dim RegInt As Currency
Dim PenalInt As Currency
Dim LoanAmount As Currency
Dim DepAmount As Currency
Dim MiscAmount As Currency
Dim DepInt As Currency
Dim boolDeposit As Boolean
Dim MiscHeadId As Long
Dim ChequeNo As String

Dim TransDate As Date

' Setup the error handler.
On Error GoTo Err_line

' Check if a loan account is loaded.
If m_rstLoanMast Is Nothing Then GoTo Exit_Line

'Now Get the Last Transaction which is supposed to delete
lastTransID = GetKCCMaxTransID(m_LoanID)

'Get the last transaction from the loan transaction table.
gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans " & _
        " Where Loanid = " & m_LoanID & _
        " And TransID = " & lastTransID
        
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    TransDate = rst("TransDate")
    transType = FormatField(rst("TransType"))
    'if Transaction is of withdrawn
    'then user might have used the cheque
    'so to mark the cheque get the cheque no
    If transType = wWithdraw Or transType = wContraWithdraw Then _
                                ChequeNo = FormatField(rst("VoucherNo"))
    
    boolDeposit = FormatField(rst("Deposit"))
    If boolDeposit Then
        DepAmount = FormatField(rst("Amount"))
    Else
        LoanAmount = FormatField(rst("Amount"))
    End If
    If rst.RecordCount > 1 Then
        rst.MoveNext
        boolDeposit = FormatField(rst("Deposit"))
        If boolDeposit Then
            DepAmount = FormatField(rst("Amount"))
        Else
            LoanAmount = FormatField(rst("Amount"))
        End If
    End If
End If

gDbTrans.SqlStmt = "SELECT *  FROM BKCCIntTrans " & _
            " WHERE loanid = " & m_LoanID & _
            " AND TransID = " & lastTransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    TransDate = rst("TransDate")
    IntTransType = FormatField(rst("TransType"))

    boolDeposit = FormatField(rst("Deposit"))
    If boolDeposit Then
        DepInt = FormatField(rst("IntAmount"))
    Else
        RegInt = FormatField(rst("IntAmount"))
        PenalInt = FormatField(rst("PenalIntAmount"))
    End If
    MiscAmount = FormatField(rst("MiscAmount"))
End If

If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & IIf(boolDeposit, m_DepHeadID, m_LoanHeadID) & _
            " And AccID = " & m_LoanID & " ANd TransID = " & lastTransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
                        LoanUndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If
    
'there might be a transaction in amount receivable
'So undo that transation
gDbTrans.SqlStmt = "Select * From AmountReceivable " & _
            " Where AccHeadID = " & m_LoanHeadID & _
            " And AccId = " & m_LoanID & " And AccTransID = " & lastTransID
Dim MiscHeadCount As Integer
Dim I As Integer

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'Now Load All those Heads And Amounts to the Varable
    'Where ever the transaction Has happened
    Dim DueAmount() As Currency
    Dim DueHeadID() As Long
    TransDate = rst("TransDate")
    Do
        ReDim Preserve DueAmount(MiscHeadCount)
        ReDim Preserve DueHeadID(MiscHeadCount)
        DueHeadID(UBound(DueAmount)) = FormatField(rst("DueHeadID"))
        DueAmount(UBound(DueAmount)) = FormatField(rst("Amount"))
        
        rst.MoveNext
        MiscHeadCount = MiscHeadCount + 1
        If rst.EOF Then Exit Do
    Loop
End If

' Begin transaction
gDbTrans.BeginTrans
inTransaction = True

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
MiscHeadId = GetHeadID(GetResourceString(327), parBankIncome)
If MiscHeadId = 0 Then MiscHeadId = parIncome + 1

' Delete the entry of principle amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM BKCCTrans " & _
            " WHERE Transid = " & lastTransID & _
            " AND Loanid = " & m_LoanID

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

' Delete the entry of Interest amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM BKCCIntTrans " & _
                " WHERE Transid = " & lastTransID & _
                " AND Loanid = " & m_LoanID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
'Now Delete the transaction that has happened in Amount receivable Field due to this
If MiscHeadCount Then _
    If Not UndoAmountReceivable(m_LoanHeadID, m_LoanID, lastTransID) Then GoTo Exit_Line


If transType = wContraDeposit Or transType = wDeposit Then
    If DepInt > 0 And (DepInt = DepAmount Or DepInt = LoanAmount) Then
        'iF amount has been Deposited to the Loan Account then
        If Not bankClass.UndoContraTrans(m_DepIntHeadID, IIf(LoanAmount > 0, m_LoanHeadID, m_DepHeadID), DepInt, TransDate) Then GoTo Exit_Line
    Else
        'iF amount has been Deposited to the Loan Account then
        If Not bankClass.UndoCashDeposits(m_LoanHeadID, LoanAmount, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoCashDeposits(m_DepHeadID, DepAmount, TransDate) Then GoTo Exit_Line
        'When amoun tis deposited to the Account then Interst & Penal int
        'Also Must be added to the account if any
        ''Now Update the interest and other
        If Not bankClass.UndoCashDeposits(m_RegIntHeadID, RegInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoCashDeposits(m_PenalHeadID, PenalInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoCashWithdrawls(m_DepIntHeadID, DepInt, TransDate) Then GoTo Exit_Line
        If MiscHeadCount Then
            For I = 0 To MiscHeadCount - 1
                If Not bankClass.UndoCashDeposits(DueHeadID(I), DueAmount(I), TransDate) Then GoTo Err_line
            Next I
        Else
            If Not bankClass.UndoCashDeposits(MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
        End If
    End If

Else 'If TransType = wContraWithdraw Then
    If LoanAmount = RegInt + PenalInt + MiscAmount And LoanAmount > 0 Then
        ' The amount is debitted to the loan account and credted to the interest acount
        If Not bankClass.UndoContraTrans(m_LoanHeadID, m_RegIntHeadID, RegInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoContraTrans(m_LoanHeadID, m_PenalHeadID, PenalInt, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoContraTrans(m_LoanHeadID, MiscHeadId, MiscAmount, TransDate) Then GoTo Exit_Line
    Else
        If Not bankClass.UndoCashWithdrawls(m_LoanHeadID, LoanAmount, TransDate) Then GoTo Exit_Line
        If Not bankClass.UndoCashWithdrawls(m_DepHeadID, DepAmount, TransDate) Then GoTo Exit_Line
        ''Now update the Interest and mIsc amount
        'PArt amount might have been receipted as interest
        If Not bankClass.UndoCashDeposits(m_RegIntHeadID, RegInt, TransDate) Then GoTo Err_line
        If Not bankClass.UndoCashDeposits(m_PenalHeadID, PenalInt, TransDate) Then GoTo Err_line
        'If Not BankClass.UndoCashWithdrawls(m_DepIntHeadID, DepInt, TransDate) Then GoTo Err_line
        'Now Undo the Miscleneous amouunt from the Misc headid
        'Check Whether there or any other heads wheremisc amount has been transferred
        If MiscHeadCount Then
            For I = 0 To MiscHeadCount - 1
                If Not bankClass.UndoCashDeposits(DueHeadID(I), DueAmount(I), TransDate) Then GoTo Err_line
            Next I
        Else
            If Not bankClass.UndoCashDeposits(MiscHeadId, MiscAmount, TransDate) Then GoTo Err_line
        End If
    End If
End If

'UpDate The LoanMAster
gDbTrans.SqlStmt = "UPDATE BkCCMaster SET LoanClosed = 0 " & _
    " Where Loanid =  " & m_LoanID
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
If DateDiff("d", TransDate, m_rstLoanMast("LastIntDate")) = 0 Then
    gDbTrans.SqlStmt = "UPDATE BkccMaster SET LastIntDate = NULL " & _
        " Where Loanid =  " & m_LoanID
        ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

If Len(ChequeNo) Then
    'if user has used the cheque for this transactioj
    'then mark the che que as unused
    gDbTrans.SqlStmt = "Update ChequeMaster Set Trans = " & wischqIssue & "," & _
                "TransDate = NULL, Amount = NULL ," & _
                "Particulars = ''" & _
                " Where AccID = " & m_LoanID & " AND AccHeadID = " & m_LoanHeadID & _
                " AND ChequeNo = " & ChequeNo
    'Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

' Commit the transaction.
gDbTrans.CommitTrans
inTransaction = False

'MsgBox "The last transaction is deleted.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    If inTransaction Then gDbTrans.RollBack: inTransaction = False
    Exit Function

Err_line:

    If Err Then
        MsgBox "LoanUndoLastTransaction: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    GoTo Exit_Line

End Function
Private Function LoanUpDate() As Boolean

If Not CheckValidity Then Exit Function
' Setup error handler.
On Error GoTo LoanUpdate_Error:

' Declare variables for this procedure...
' ------------------------------------------
Dim txtIndex As Integer
Dim inTransaction As Boolean
Dim Lret As Long
Dim MemID As Long
Dim custId As Long
Dim lGuarantorID As Long
Dim rst As Recordset
' ------------------------------------------

' Check for member details.
txtIndex = GetIndex("MemberID")
With txtLoanIssue(txtIndex)

    ' Check if the specified memberid is valid.
    gDbTrans.SqlStmt = "SELECT * FROM MemMaster " & _
            " WHERE AccNum = " & AddQuotes(.Text, True)
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    MemID = FormatField(rst("AccId"))
    custId = FormatField(rst("CustomerID"))
End With


'Memeber type
'Now Get the Farmer Type
Dim FarmerType As Integer
Dim CtrlIndex As Integer
txtIndex = GetIndex("FarmerType")

'Now get the index of the combo box which represents the farmer type
CtrlIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
With cmb(CtrlIndex)
    FarmerType = .ItemData(.ListIndex)
End With

txtIndex = GetIndex("InterestBalance")
With txtLoanIssue(txtIndex)
    If .Text <> "" Then
        If Not CurrencyValidate(txtLoanIssue(txtIndex), True) Then
            MsgBox "Please specify the interst balance", vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtIndex)
            Exit Function
        End If
    End If
End With

'Get the Interest Balance
Dim IntBalance As Double
gDbTrans.SqlStmt = "Select IntBalance From BKCCIntTrans " & _
        " Where LoanID = " & m_LoanID & _
        " Order By TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then IntBalance = FormatField(rst(0))

'Put into the bkcc int trans
If IntBalance <> Val(txtLoanIssue(txtIndex)) Then
    If MsgBox("You are changing the interest balance" & vbCrLf & _
        "Do you want to coninue?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
        ActivateTextBox txtLoanIssue(txtIndex)
        Exit Function
    End If
End If

'Dim FarmerType As Integer
txtIndex = GetIndex("FarmerType")
'Now get the index of the combo box which represents the farmer type
CtrlIndex = ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex")
FarmerType = cmb(CtrlIndex).ItemData(cmb(CtrlIndex).ListIndex)

'Begin transaction.
gDbTrans.BeginTrans
inTransaction = True

' Put an entry into LoanMaster table.
gDbTrans.SqlStmt = "UPDATE BKCCMaster Set " & _
        " AccNum = " & AddQuotes(Trim$(GetValue("AccNum")), True) & ", " & _
        " MemId = " & MemID & ", CustomerId = " & custId & ", " & _
        " IssueDate = #" & GetSysFormatDate(GetValue("IssueDate")) & "#, " & _
        " SanctionAmount = " & GetValue("SanctionAmount") & ", " & _
        " CurrentSanction = " & GetValue("CurrentSanction") & ", " & _
        " ExtraSanction = " & GetValue("ExtraSanction") & ", " & _
        " RenewDate = #" & GetSysFormatDate(GetValue("RenewDate")) & "#, " & _
        " Guarantor1 = " & PropGuarantorID(1) & ", " & _
        " Guarantor2 = " & PropGuarantorID(2) & ", " & _
        " IntRate = " & Val(GetValue("InterestRate")) & "," & _
        " PenalIntRate = " & Val(GetValue("PenalInterestRate")) & "," & _
        " DepIntRate = " & Val(GetValue("DepositInterestRate")) & "," & _
        " DryLand = " & AddQuotes(GetValue("DryLand")) & "," & _
        " IrrigationLand = " & AddQuotes(GetValue("WetLand")) & "," & _
        " DryIncome = " & Val(GetValue("DryIncome")) & "," & _
        " IrrigationIncome = " & Val(GetValue("WetIncome")) & " ," & _
        " InterestBalance= " & Val(GetValue("InterestBalance")) & "," & _
        " FarmerType = " & FarmerType & _
        " WHERE LoanID = " & m_LoanID

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

'Put into the bkccinttra
gDbTrans.SqlStmt = "UPDATE BKCCintTrans set " & _
        " intBalance =" & GetValue("InterestBalance") & _
        " Where LoanID = " & m_LoanID & _
        " And TransID = (Select Max(TransID) From " & _
        " BkccIntTrans Where LoanID = " & m_LoanID & ")"
If Not gDbTrans.CommitTrans Then GoTo Exit_Line

inTransaction = False
'm_LoanID = 0
MsgBox GetResourceString(707), vbInformation, wis_MESSAGE_TITLE
'MsgBox "Loan updated  successfully.", vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    'If Val(txtLoanIssue(GetIndex("InstalmentNames")).Text) > 1 Then Unload frmLoanName
    If inTransaction Then gDbTrans.RollBack: m_LoanID = 0
    Exit Function

LoanUpdate_Error:
    If Err Then
        MsgBox "LoanUPDATE: " & Err.Description, vbCritical
        'MsgBox GetResourceString(728) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line

End Function
' Returns the memberID of Guarantor1, selected by the user.
Private Function PropGuarantorID(GuarantorNum As Integer) As Long
Dim strTmp As String
Dim txtIndex As Integer
txtIndex = GetIndex("Guarantor" & GuarantorNum)
If txtIndex >= 0 Then
    'strTmp = ExtractToken(txtLoanIssue(txtIndex).Tag, "GuarantorID")
    strTmp = ExtractToken(txtLoanIssue(txtIndex).Tag, "CustomerID")
    PropGuarantorID = Val(strTmp)
End If

End Function
Private Sub PropInitializeForm()

' Set the proerties for tab strip control.
TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True

' Load the properties for loan issue panel.
LoadLoanIssueProp

' Remove all tabs of tabloans.


End Sub
' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetValue(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
For I = 0 To txtLoanIssue.count - 1
    strTxt = ExtractToken(lblLoanIssue(I).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        GetValue = txtLoanIssue(I).Text
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
'****************************************************************************************
'Returns a new account number
'Author: Girish
'Date : 29th Dec, 1999
'Modified by Ravindra on 25th Jan, 2000
'****************************************************************************************
Private Function GetNewAccountNumber() As String
    Dim rst As Recordset
    Dim NewAccNo As String
    'gDBTrans.SQLStmt = "Select TOP 1 AccID from SBMaster order by AccID desc"
    gDbTrans.SqlStmt = "SELECT MAX(AccID) FROM SBMaster"
    If gDbTrans.Fetch(rst, adOpenDynamic) = 0 Then
        NewAccNo = "0"
    Else
        NewAccNo = FormatField(rst(0))
    End If
    If IsNumeric(NewAccNo) Then NewAccNo = Val(NewAccNo) + 1
    
    GetNewAccountNumber = NewAccNo

End Function

'
Private Function LoadLoanIssueProp() As Boolean
' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\Bkcc_" & gLangOffSet & ".PRP"

If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet <> wis_NoLangOffset Then
        PropFile = App.Path & "\Bkcckan.PRP"
    Else
        PropFile = App.Path & "\bkcc.PRP"
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
    strTag = Trim(ReadFromIniFile("LoanIssue", "Prop" & I + 1, PropFile))
    If strTag = "" Then Exit Do
    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load lblLoanIssue(lblLoanIssue.count)
        Load txtLoanIssue(txtLoanIssue.count)
    End If
    CtlIndex = lblLoanIssue.count - 1
    ' Get the property type.
    strPropType = Trim$(ExtractToken(strTag, "PropType"))
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With lblLoanIssue(CtlIndex)
                .FontBold = True
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
        .Locked = IIf(StrComp(strRet, "True", vbTextCompare) = 0, True, False)
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
                    Else Load cmdLoanIssue(cmdLoanIssue.count)
            With cmdLoanIssue(cmdLoanIssue.count - 1)
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
                        "TextIndex", CStr(cmdLoanIssue.count - 1))
            End With
    End Select

    ' Increment the loop count.
    I = I + 1
Loop
ArrangeLoanIssuePropSheet

' Display today's date, for date field.
I = GetIndex("IssueDate")
If I >= 0 Then txtLoanIssue(I).Text = gStrDate


End Function
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
Private Sub ResetTransactionForm()

'First Clear the Transaction Form
grd.Clear
Set m_rstLoanMast = Nothing
Set m_rstLoanTrans = Nothing

txtBalance = ""
txtIntBalance.Tag = ""
cmbTrans.ListIndex = -1
cmbParticulars.ListIndex = -1
cmbParticulars.Text = ""

cmbCheque.Visible = False
cmbCheque.Tag = "InVisible"

cmdAbn.Enabled = False
cmdAccept.Enabled = False
txtIntBalance = 0
txtIntBalance.Enabled = False
txtRegInterest = 0
txtRegInterest.Enabled = False
txtPenalInterest = 0
txtPenalInterest.Enabled = False
txtTotal = 0
txtTotal.Enabled = False
txtRepayAmt = 0
txtRepayAmt.Enabled = False
txtSbAccNum = ""
txtSbAccNum.Enabled = False
chkSb.Value = vbUnchecked

cmdPrevTrans.Enabled = False
cmdNextTrans.Enabled = False
cmdPrint.Enabled = False

'If Not m_CustReg Is Nothing Then m_CustReg.NewCustomer

'Now Clear the Account creation Controls
Dim count As Integer
'For Count = 0 To txtLoanIssue.Count - 1
'    txtLoanIssue(Count).Text = ""
'Next
'For Count = 0 To cmb.Count - 1
'    cmb(Count).ListIndex = -1
'Next

With cmdLoanSave
    .Enabled = False
End With
With txtMiscAmount
    .Locked = False
    .Value = 0
End With

Set m_clsReceivable = Nothing

End Sub

Private Sub ResetUserIntereface()
On Error GoTo ErrLine
Call ResetTransactionForm

'Get the Head ids
Dim ClsBank As clsBankAcc
Dim headName As String
Dim headNameEnglish As String
Dim SetUp As New clsSetup
If m_LoanHeadID = 0 Then
    Set ClsBank = New clsBankAcc
    gDbTrans.BeginTrans
    headName = GetResourceString(229)
    headNameEnglish = LoadResString(229)
    
    m_LoanHeadID = ClsBank.GetHeadIDCreated(headName & " " & GetResourceString(58), _
                    headNameEnglish & " " & LoadResString(58), parMemberLoan, 0, wis_BKCCLoan)
    m_RegIntHeadID = ClsBank.GetHeadIDCreated(headName & " " & GetResourceString(58, 344), _
            headNameEnglish & " " & LoadResourceStringS(58, 344), parMemLoanIntReceived, 0, wis_BKCCLoan)
    m_PenalHeadID = ClsBank.GetHeadIDCreated(headName & " " & GetResourceString(58, 345), _
            headNameEnglish & " " & LoadResourceStringS(58, 345), parMemLoanPenalInt, 0, wis_BKCCLoan)
    
    m_DepHeadID = m_LoanHeadID
    m_DepIntHeadID = m_RegIntHeadID

'Check whether Negtive balance is considered as Deposit or NOt
    m_KCCDepositTrans = IIf(UCase(SetUp.ReadSetupValue("General", "KCCDeposit", "True")) = "FALSE", False, True)

    If m_KCCDepositTrans Then
        m_DepHeadID = ClsBank.GetHeadIDCreated(headName & " " & GetResourceString(43), _
                    headNameEnglish & " " & LoadResString(43), parMemberDeposit, 0, wis_BKCC)
        m_DepIntHeadID = ClsBank.GetHeadIDCreated(headName & " " & GetResourceString(43, 487), _
                    headNameEnglish & " " & LoadResourceStringS(43, 487), parMemDepIntPaid, 0, wis_BKCC)
    End If
    Set SetUp = Nothing
    
    gDbTrans.CommitTrans
End If

m_LoanID = 0
m_dbOperation = Insert

txtMemberName = ""
txtLoanAmt = ""
txtIssueDate = ""

cmdAbn.Enabled = False

cmdLoanUpdate.Enabled = False

If Not m_CustReg Is Nothing Then m_CustReg.NewCustomer

'Now Clear the Account creation Controls
Dim count As Integer
For count = 0 To txtLoanIssue.count - 1
    txtLoanIssue(count).Text = ""
Next
For count = 0 To cmb.count - 1
    cmb(count).ListIndex = -1
Next
Dim txtIndex As Integer
txtIndex = GetIndex("MemberType")
txtIndex = Val(ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex"))
If cmb(txtIndex).ListCount = 1 Then cmb(txtIndex).ListIndex = 0
cmdLoanSave.Enabled = True


RaiseEvent AccountChanged(0)

ErrLine:
End Sub


Private Function TransactMiscAmount() As Boolean

On Error GoTo Err_line
 
'Dim LastTransDate As Date
Dim TransID As Long
Dim inTransaction As Boolean

Dim NewBalance As Currency
Dim Balance As Currency

Dim MiscAmount As Currency
Dim PrincAmount  As Currency
'Dim DepAmount As Currency

Dim TotalAmount As Currency
Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim Deposit As Boolean
Dim rst As Recordset
Dim Amount As Currency
    
Dim bankClass As clsBankAcc
Dim MiscHeadId As Long

Dim VoucherNo As String
Dim UserID As Integer
Dim strParticulars As String
Dim rstTemp As Recordset


'Get the Voucher No and Cheque No
UserID = gCurrUser.UserID
VoucherNo = Trim(txtVoucherNo.Text)

MiscAmount = txtMiscAmount 'txtPenalInt = txtPenalInterest.Value

transType = IIf(cmbTrans.ListIndex = 0, wContraDeposit, wContraWithdraw)
TransDate = GetSysFormatDate(txtRepayDate.Text)

TotalAmount = txtMiscAmount.Value


'Now check whether the amount collecting is
'Deducting or addinng to the Loan head or not
Dim IsAddingToAccount As Boolean
Dim I As Integer
Dim OtherTransType As wisTransactionTypes

'Check Whether the amoutisnowraferring to the misc account
IsAddingToAccount = m_clsReceivable.IsAddToAccount
'If all the amount is adding to the Account then
OtherTransType = IIf(transType = wContraDeposit, wContraWithdraw, wContraDeposit)

'now get the Balance and Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Top 1 Balance,TransID,TransDate " & _
            " FROM BKCCTrans WHERE loanid = " & m_LoanID & _
            " ORDER BY TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Balance = Val(FormatField(rst("Balance")))
    TransID = Val(FormatField(rst("TransID")))
End If

'Begin the transactionhere
gDbTrans.BeginTrans
inTransaction = True

TransID = TransID + 1

'First insert then Amount ot the Accounthead
If IsAddingToAccount Then
    Deposit = IIf(Balance < 0, True, False)
    'If customer is paying the amount to the bank,reduece the balance
    If transType = wDeposit Or transType = wContraDeposit Then
        Balance = Balance - TotalAmount
    Else
        Balance = Balance + TotalAmount
    End If

    strParticulars = GetResourceString(327)
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans " _
                & "(LoanID, TransID, TransType, Amount, " _
                & " TransDate, Balance, Deposit,VoucherNO,UserId, Particulars) " _
                & "VALUES (" & m_LoanID & ", " & TransID & ", " _
                & transType & ", " & Amount & ", " _
                & "#" & TransDate & "#, " _
                & Balance & ", " & Deposit & "," _
                & AddQuotes(VoucherNo) & "," & UserID & "," _
                & AddQuotes(strParticulars) & " )"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    'Now get the Contra ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID + 1
    gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                    "( ContraID,AccHeadId,AccID," & _
                    " TransType,TransID,Amount," & _
                    " UserID,VoucherNo) VALUES (" & _
                    ContraID & "," & _
                    m_LoanHeadID & "," & m_LoanID & _
                    transType & "," & TransID & "," & _
                    TotalAmount & "," & UserID & "," & _
                    AddQuotes(VoucherNo) & ")"
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
End If

'Now Get the pending Balance From the Previous account
Balance = 0
gDbTrans.SqlStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & m_LoanHeadID & _
        " ANd AccId = " & m_LoanID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    rstTemp.MoveLast
    Balance = FormatField(rstTemp("Balance"))
End If

I = 0
Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)

Do While MiscHeadId > 0

    If MiscHeadId = 0 Then Exit Do
    
    If IsAddingToAccount Then
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(MiscHeadId) & "," & MiscHeadId & _
                        OtherTransType & "," & TransID & "," & _
                        MiscAmount & "," & UserID & "," & _
                        AddQuotes(VoucherNo) & ")"
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        'Insert the Same Into Head Transaction class
        If transType = wContraWithdraw Then
            If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                    MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
        Else
            If Not bankClass.UpdateContraTrans(MiscHeadId, _
                    m_LoanHeadID, MiscAmount, TransDate) Then GoTo ExitLine
        End If
    Else
        If transType = wContraDeposit Then
            transType = wDeposit
            If Not bankClass.UndoCashDeposits(MiscHeadId, _
                                 MiscAmount, TransDate) Then GoTo ExitLine
        End If
    End If
    
    If OtherTransType = wContraWithdraw Then
        Balance = Balance - MiscAmount
    Else
        Balance = Balance + MiscAmount
    End If
    'Now insert this details Into amount receivable table
    Debug.Assert Amount = 0
    If Not AddToAmountReceivable(m_LoanHeadID, m_LoanID, _
                    TransID, TransDate, MiscAmount, MiscHeadId) Then GoTo ExitLine
    I = I + 1

    Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, Amount, MiscAmount)
Loop

gDbTrans.CommitTrans
inTransaction = False

Call ResetTransactionForm

TransactMiscAmount = True
Exit Function

ExitLine:
    If inTransaction Then gDbTrans.RollBack
    Exit Function
    
Err_line:
    
    MsgBox "Error in Misc transaction", vbInformation, wis_MESSAGE_TITLE
    
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


Private Sub chkSb_Click()

If cmbTrans.ListIndex = 0 Then
    'Transaction is of Receipt
    txtSbAccNum.Enabled = False
    txtSbAccNum.BackColor = wisGray
    txtSbAccNum.ZOrder 0
    cmbCheque.Visible = False
    txtRepayDate.Tag = ""
    Call txtRepayDate_Change
    Exit Sub
End If
If chkSb.Value = vbChecked And cmbTrans.ListIndex = 1 Then
    'Add Intrest transaction
    cmdSB.Enabled = True
    txtSbAccNum.Enabled = True
    txtSbAccNum.BackColor = wisWhite
Else
    'Add Intrest transaction
    cmdSB.Enabled = False
    txtSbAccNum.Enabled = False
    txtSbAccNum.BackColor = wisGray
    
    Exit Sub
End If

'Now Case if chksb.value=true
txtRegInterest = 0
txtPenalInterest = 0

If m_LoanID = 0 Then Exit Sub

txtSbAccNum.ZOrder 0
cmbCheque.Visible = False
txtVoucherNo.Visible = True
txtVoucherNo.ZOrder 0
Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT AccNum From SBMASTER " & _
    " WHERE CustomerId = (SELECT CustomerID From BKCCMAster " & _
        " WHERE LoanID = " & m_LoanID & ");"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                        txtSbAccNum = FormatField(rst(0))

End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" And cmb(Index).ListIndex >= 0 Then
    txtLoanIssue(Val(txtIndex)).Text = cmb(Index).Text
    txtLoanIssue(Val(txtIndex)).Tag = PutToken(txtLoanIssue(Val(txtIndex)).Tag, "ComboValue", cmb(Index).ItemData(cmb(Index).ListIndex))
End If

End Sub





Private Sub cmbTrans_Click()

If cmbTrans.ListIndex < 0 Then Exit Sub
If m_LoanID = 0 Then Exit Sub

cmdMisc.Enabled = False
cmdAccept.Enabled = True
txtRegInterest.Enabled = True
txtTotal.Enabled = True

lblChequeNo.Caption = GetResourceString(41) & _
                    " " & GetResourceString(60) '"Voucher Num
lblRegInterest.Caption = GetResourceString(344)
With cmbTrans
    chkSb.Enabled = False
    cmbCheque.Visible = False
    txtVoucherNo.ZOrder 0
    If .ListIndex = 0 Then 'Receipt
        .Tag = 1
        lblTotInst.Caption = GetResourceString(52) & _
                        " " & GetResourceString(40)
        chkSb.Caption = GetResourceString(47) & _
                    " " & GetResourceString(10) 'Interest add
        chkSb.TabIndex = cmbParticulars.TabIndex + 1
        'Now Check the Receivable Amount
        txtMiscAmount = GetReceivAbleAmount(m_LoanHeadID, m_LoanID)
        If txtMiscAmount.Value > 0 Then _
            Set m_clsReceivable = LoadReceivableAmounts(m_LoanHeadID, m_LoanID)
        
        chkSb.Enabled = True
        txtSbAccNum.Enabled = False
        txtSbAccNum.BackColor = wisGray
        cmdSB.Enabled = False
        cmdMisc.Enabled = True

    ElseIf .ListIndex = 1 Then 'Payment
        .Tag = "-1"
        lblTotInst.Caption = GetResourceString(375) & _
                            " " & GetResourceString(40)
        chkSb.Caption = GetResourceString(421) & _
                        " " & GetResourceString(271) ' Savings bank receipt
        lblChequeNo.Caption = GetResourceString(275) & _
                    " " & GetResourceString(60) '"Sb acc NUm
        chkSb.TabIndex = 34
        chkSb.Enabled = True
        txtSbAccNum.Enabled = True
        lblChequeNo.Enabled = True
        txtSbAccNum.BackColor = wisWhite
        cmdSB.Enabled = True
        'cmdCheque.Enabled = True
        If UCase(cmbCheque.Tag) = "VISIBLE" Then
            cmbCheque.Visible = True
            cmbCheque.ZOrder 0
        End If
    Else
    'If .ListIndex = 2 Then 'interest on deposit
        .Tag = 1
        lblRegInterest.Caption = GetResourceString(233)
        chkSb.Enabled = False
        txtSbAccNum.Enabled = False
        txtSbAccNum.BackColor = wisGray
        cmdSB.Enabled = False
        txtMiscAmount.Locked = False
    End If
    If Not chkSb.Enabled Then chkSb.Value = vbUnchecked
End With


Call chkSb_Click
txtVoucherNo.Visible = Not cmbCheque.Visible

txtRepayDate.Tag = ""
Call txtRepayDate_Change

Dim TransDate As Date

If cmbTrans.ListIndex = 2 Then
'If cmbTrans.ListIndex = 2 Or (cmbTrans.ListIndex = 1 And chkSb.Value = vbUnchecked) Then
    'If He is giving deposit Interest then
    'Disable the All other controls
    If DateValidate(txtRepayDate.Text, "/", True) Then
        TransDate = GetSysFormatDate(txtRepayDate)
        txtRegInterest = BKCCDepositInterest(m_LoanID, TransDate) \ 1
    End If
    
    With txtIntBalance
        .Text = "0"
    End With
    With txtPenalInterest
        .Text = "0"
        .Locked = True
    End With
     
    With txtMiscAmount
        .Text = "0"
        .Locked = True
    End With
    If cmbTrans.ListIndex = 1 Then CheckForDueAmount
    
    With txtRepayAmt
        .Text = "0"
        .Enabled = False
    End With
    With txtTotal
        .Enabled = chkSb.Enabled
        '.Enabled = True
    End With

Else
    txtIntBalance.Enabled = True
    txtPenalInterest.Enabled = True
    txtMiscAmount.Enabled = True
    txtRepayAmt.Enabled = True
    txtTotal.Enabled = True
End If

End Sub

Private Sub cmdAbn_Click()

cmdAbn.Enabled = False
If m_LoanID <= 0 Then Exit Sub
If Not DateValidate(txtRepayDate, "/", True) Then Exit Sub
    

Dim obj As New clsBkcc
cmdAbn.Enabled = (obj.LoanDueDate(m_LoanID) < GetSysFormatDate(txtRepayDate))
Set obj = Nothing

If m_frmAbn Is Nothing Then _
            Set m_frmAbn = New frmBKCCAbn

m_frmAbn.LoanAccountID = m_LoanID
m_frmAbn.Show vbModal

cmdAbn.Enabled = True

End Sub

Private Sub cmdAccept_Click()

' Check if a valid amount is entered.
If txtTotal <= 0 And m_clsReceivable Is Nothing Then
    'MsgBox "Enter valid amount.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    txtTotal.SetFocus
    GoTo Exit_Line
End If

If Not DateValidate(txtRepayDate.Text, "/", True) Then
    'iNVALID dATE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo Exit_Line
End If

'Check for the Previos transction date
If DateDiff("d", GetKCCLastTransDate(m_LoanID), GetSysFormatDate(txtRepayDate.Text)) < 0 Then
    'iNVALID dATE
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayDate
    GoTo Exit_Line
End If

'Check The Trasnction Type
If cmbTrans.ListIndex < 0 Then
    'Specify transaction
    MsgBox GetResourceString(588), vbInformation, wis_MESSAGE_TITLE
    cmbTrans.SetFocus
    GoTo Exit_Line
End If

If cmbCheque.Visible Then
    If Trim(cmbCheque.Text) = "" Then
    MsgBox GetResourceString(755), vbInformation, wis_MESSAGE_TITLE
    cmbCheque.SetFocus
    Exit Sub
    End If
ElseIf Trim(txtVoucherNo.Text) = "" Then
    MsgBox GetResourceString(755), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtVoucherNo
    Exit Sub
End If

'Get the Particulars
If Trim$(cmbParticulars.Text) = "" Then
    'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(621), vbExclamation, gAppName & " - Error"
    cmbParticulars.SetFocus
    Exit Sub
End If

If LoanTransaction Then LoanLoad (m_LoanID)

Exit_Line:

End Sub

Private Sub cmdAddNote_Click()

If m_Notes.ModuleId = 0 Then Exit Sub

If m_Notes.AccId = 0 Then Exit Sub

Call m_Notes.Show
Call m_Notes.DisplayNote(rtfNote)

End Sub

Private Sub cmdAdvance_Click()
 If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
 
 m_clsRepOption.ShowDialog
 
End Sub

Private Sub cmdApply_Click()
'Now Check The whether the limits are correct
Dim I As Integer
Dim MaxI As Integer
Dim TransDate As Date
Dim SetupClass As New clsSetup
Dim IntClass As New clsInterest

'Now chack the date validity
If Not DateValidate(txtEffectiveDate.Text, "/", True) Then
    'Invalid date specified
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtEffectiveDate
    Exit Sub
End If
TransDate = GetSysFormatDate(txtEffectiveDate)

If Val(Trim$(txtSubsidyRate)) > 100 Then
    'Invalid Interest rate specified
    MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtSubsidyRate
    Exit Sub
End If
Call IntClass.SaveInterest(wis_BKCCLoan, "Subsidy", CSng(txtSubsidyRate.Text), , , FinUSFromDate)
             
If Val(Trim$(txtRebateRate)) > 100 Then
    'Invalid Interest rate specified
    MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRebateRate
    Exit Sub
End If

Call IntClass.SaveInterest(wis_BKCCLoan, "Rebate", CSng(txtRebateRate.Text), , , FinUSFromDate)
             
If Val(Trim$(txtRebateRate1)) > 100 Then
    'Invalid Interest rate specified
    MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRebateRate1
    Exit Sub
End If
Call IntClass.SaveInterest(wis_BKCCLoan, "Rebate1", CSng(txtRebateRate1.Text), , , FinUSFromDate)

If Val(Trim$(txtRebateRate2)) > 100 Then
    'Invalid Interest rate specified
    MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRebateRate2
    Exit Sub
End If
Call IntClass.SaveInterest(wis_BKCCLoan, "Rebate2", CSng(txtRebateRate2.Text), , , FinUSFromDate)
             
If Val(Trim$(txtNabardRate)) > 100 Then
    'Invalid Interest rate specified
    MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtNabardRate
    Exit Sub
End If
Call IntClass.SaveInterest(wis_BKCCLoan, "NabardRate", CSng(txtNabardRate.Text), , , FinUSFromDate)

MaxI = txtMaxLimit.UBound
For I = 1 To MaxI
    If txtMaxLimit(I).Value = 0 Then Exit For
    'Now Check for the this limit is more than previos 1 or not
    If txtMaxLimit(I) <= txtMaxLimit(I - 1) And txtMaxLimit(I).Value > 0 Then
        'Invalid amount specified
        MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtMaxLimit(I)
        Exit Sub
    End If
    If Val(txtLoanIntRate(I)) <= 0 Then
        'Invalid Interest rate specified
        MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIntRate(I)
        Exit Sub
    End If
Next
MaxI = I

'Now Save the Iterest Rate
Dim strScheme As String

If MaxI > txtLoanIntRate.UBound Then MaxI = txtLoanIntRate.UBound
For I = 0 To MaxI
    
    strScheme = txtMinLimit(I).Caption & "-" & txtMaxLimit(I).Value
    If txtMaxLimit(I).Value <= 0 Then Exit For
    
    'Now store the Intrest rate to interest class
    If IntClass.SaveInterest(wis_BKCCLoan, "Slab" & I, _
             CSng(txtLoanIntRate(I).Text), , , FinUSFromDate) Then
        
        'then Store the Slab infomation int the set up then
        Call SetupClass.WriteSetupValue("KCCLoanInt", "Slab" & I, strScheme)
    
    End If
Next


End Sub

Private Sub cmdCheque_Click()
'if he is with drawing through the cheque
If Val(txtBalance) = 0 Then Exit Sub

Set m_frmCheque = New frmCheque
m_frmCheque.p_AccID = m_LoanID
m_frmCheque.AccHeadID = IIf(Val(txtBalance.Tag) > 0, m_LoanHeadID, m_DepHeadID)
m_frmCheque.Show vbModal

Call LoadChequeNos
    
End Sub

'
Private Sub cmdEndDate_Click()
With Calendar
    .Left = Me.Left + fraReports.Left + fraOrder.Left + cmdEndDate.Left - .Width
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdEndDate.Top + 300
    .selDate = txtEndDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtEndDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdLoad_Click()
Me.MousePointer = vbHourglass

Dim LoanID As Long
Dim rst As Recordset
Dim BkccFetch As Integer

gDbTrans.SqlStmt = "SELECT * From BKCCMaster Where " & _
        " AccNum = " & AddQuotes(txtAccNo, True)
BkccFetch = gDbTrans.Fetch(rst, adOpenDynamic)
If BkccFetch < 1 Then
    MsgBox GetResourceString(525), vbInformation, wis_MESSAGE_TITLE
    'Account does not exdists
    Call ResetUserIntereface
    cmdAccept.Enabled = False
    GoTo Exit_Line
Else
    cmdAccept.Enabled = True
End If

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
m_CustReg.LoadCustomerInfo (FormatField(rst("CustomerID")))

txtMemberName.Visible = True
txtMemberName = m_CustReg.FullName
LoanID = FormatField(rst("LoanId"))
gDbTrans.SqlStmt = "SELECT AccNum From MemMaster Where Accid= " & rst("MemID")
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then txtMemberID = FormatField(rst("AccNum"))

If LoanLoad(LoanID) Then fraLoanGrid.Visible = True
'txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)

Exit_Line:
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdLoanIssue_Click(Index As Integer)
Me.MousePointer = vbHourglass
' Variables of this routine...
Dim txtIndex As String
Dim strField As String
Dim Lret As Long
Dim NameStr As String
Dim rst As Recordset
Dim DueAmount As Currency

' Check to which text index it is mapped.
txtIndex = ExtractToken(cmdLoanIssue(Index).Tag, "TextIndex")

' Extract the Bound field name.
strField = ExtractToken(lblLoanIssue(Val(txtIndex)).Tag, "DataSource")

Select Case UCase$(strField)
    Case "MEMBERID"
        If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
        m_CustReg.ShowDialog
        'Now Check Whether selected Customer Is member Or not
        gDbTrans.SqlStmt = "SELECT * From MemMaster Where CustomerId = " & m_CustReg.CustomerID
        If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
            MsgBox "Selected customer is not member ", vbInformation, wis_MESSAGE_TITLE
            Me.MousePointer = vbNormal
            Exit Sub
        End If
        txtIndex = GetIndex("MemberID")
        txtLoanIssue(txtIndex).Text = FormatField(rst("AccNum"))
        txtIndex = GetIndex("MemberName")
        txtLoanIssue(txtIndex).Text = m_CustReg.FullName
        txtIndex = GetIndex("MemberType")
        txtIndex = Val(ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex"))
        If cmb(txtIndex).ListCount > 1 Then
            Call SetComboIndex(cmb(txtIndex), , FormatField(rst("MemberType")))
        End If
    
    Case "RENEWDATE"
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
            If .selDate <> "" Then txtLoanIssue(txtIndex).Text = .selDate
        End With

    Case "GUARANTOR1", "GUARANTOR2"
        ' Query for member details...
        gDbTrans.SqlStmt = "SELECT AccNum,AccId, title + space(1) + " _
                & "firstname + space(1) + middlename" _
                & " + space(1) + lastname AS name FROM NameTab, " _
                & "MemMaster WHERE nametab.customerID = MemMaster.customerID"
        
        'now Check Whether He Want Search   any particular name
         'NameStr = InputBox("Enter customer name , You want search", "Name Search")
         NameStr = InputBox(GetResourceString(785), "Name Search")
         If NameStr <> "" Then
             gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And ( FirstNAme like '" & NameStr & "%' " & _
                    " Or MiddleName like '" & NameStr & "%' Or LAstName like '" & NameStr & "%' )"
         End If
             gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order By IsciName"

        Lret = gDbTrans.Fetch(rst, adOpenDynamic)
        If Lret <= 0 Then
            MsgBox GetResourceString(569), vbInformation, wis_MESSAGE_TITLE
            'MsgBox "No members present. Create members using the members module.", vbInformation, wis_MESSAGE_TITLE
            GoTo ExitLine
        End If
        'Fill the details to report dialog and display it.
        If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp
        
        If Not FillView(m_LookUp.lvwReport, rst, True) Then
           MsgBox GetResourceString(562), _
                    vbCritical, wis_MESSAGE_TITLE
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
           GoTo ExitLine
        End If
        With m_LookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
            .lvwReport.ColumnHeaders(3).Width = 3750
            .lvwReport.ColumnHeaders(2).Width = 0
            .Title = "Select the Guarantor..."
            .m_SelItem = ""
            .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem <> "" Then
                txtLoanIssue(txtIndex).Text = .lvwReport.SelectedItem.SubItems(2)
                ' Add the guarantorID to the tag property.
                txtLoanIssue(txtIndex).Tag = PutToken(txtLoanIssue(txtIndex).Tag, "GuarantorID", .lvwReport.SelectedItem.Text)
            End If
        End With
    
    Case "RECEIVABLE"
        If m_LoanID = 0 Then Exit Sub
        If m_clsReceivable Is Nothing Then _
                    Set m_clsReceivable = New clsReceive Else DueAmount = m_clsReceivable.TotalAmount
            
        m_clsReceivable.Show
        
        If m_clsReceivable.TotalAmount = 0 Then GoTo ExitLine
        If DueAmount Then GoTo ExitLine
        'Now Call the Procedure to insert the Receivable Amount
        'in Recivable table
        Call InsertAmountReceiveAble
        
    Case "DRYLAND"
        If m_dbOperation = Insert Then GoTo ExitLine
        If m_CustReg Is Nothing Then GoTo ExitLine
        If m_CustReg.CustomerID = 0 Then GoTo ExitLine
        If m_frmAsset Is Nothing Then Set m_frmAsset = New frmAsset
        m_frmAsset.CustomerID = m_CustReg.CustomerID
        m_frmAsset.Show 1

End Select
ExitLine:
Me.MousePointer = vbDefault
End Sub


Private Sub cmdLoanClear_Click()

Call ResetUserIntereface
cmdLoanSave.Default = True

Exit Sub
Dim I As Integer
For I = 0 To txtLoanIssue.count - 1
    txtLoanIssue(I).Text = ""
Next

' If a date field, display today's date.
ReDim m_LoanAmount(0)
I = GetIndex("IssueDate")
If I >= 0 Then
    txtLoanIssue(I).Text = gStrDate
End If
cmdLoanSave.Enabled = True
cmdLoanUpdate.Enabled = False
End Sub

Private Sub cmdLoanSave_Click()

If Not LoanSave Then Exit Sub

'Now Load the Same loan ito repayment tab
txtAccNo.Text = GetValue("AccNum")
'fill to the grid
Call cmdLoad_Click
End Sub

Private Sub cmdMember_Click()

If Val(txtMemberID.Text) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

Dim LoanID As Long
Dim rst As Recordset
Dim AccNum As String
Dim memberTYpe As Integer

'Get the Member details
Call GetMemberNameCustIDByMemberNum(Trim(txtMemberID.Text), LoanID, memberTYpe)

'gDbTrans.SqlStmt = "SELECT * From BKCCMaster Where " & _
        " MemID = (SELECT AccID From MemMaster Where AccNum = '" & Trim$(txtMemberID.Text) & "')"

gDbTrans.SqlStmt = "SELECT * From BKCCMaster Where CustomerId = " & LoanID
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(525), vbInformation, wis_MESSAGE_TITLE
    'Account does not exdists
    Call ResetUserIntereface
    GoTo Exit_Line
Else
    cmdAccept.Enabled = True
End If

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
m_CustReg.LoadCustomerInfo (FormatField(rst("CustomerID")))

txtMemberName.Visible = True
txtMemberName = m_CustReg.FullName

LoanID = FormatField(rst("LoanId"))
txtAccNo = FormatField(rst("accNum"))

If LoanLoad(LoanID) Then fraLoanGrid.Visible = True
'txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)

Exit_Line:
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdMisc_Click()

With cmbTrans
    If .ListIndex < 0 Or .ListIndex = 2 Then Exit Sub
End With

If m_LoanID = 0 Then Exit Sub
If m_clsReceivable Is Nothing Then _
            Set m_clsReceivable = New clsReceive
    
m_clsReceivable.Show
    
If m_clsReceivable.TotalAmount Then
    txtMiscAmount.Locked = True
    txtMiscAmount.Value = m_clsReceivable.TotalAmount
Else
    txtMiscAmount.Locked = False
    Set m_clsReceivable = Nothing
End If


End Sub

Private Sub cmdNextTrans_Click()
m_TransID = m_TransID + 10
m_rstLoanTrans.Find "TransID >= " & m_TransID
Call LoanLoadGrid

End Sub

Private Sub cmdNotice_Click()
Dim MyDoc As Object
Dim Filename As String

Dim rstMaster As Recordset
gDbTrans.SqlStmt = "Select * From BKCCMaster" & _
        " Where LoanID = " & m_LoanID

With wisMain.cdb
    .Filter = "Word Files(*.doc)|*.doc|All Files(*.*)|*.* "
    .DefaultExt = "*.doc"
    .CancelError = True
    .ShowOpen
    Filename = .Filename
End With
If Filename = "" Then Exit Sub
If Dir(Filename, vbNormal) <> "" Then
    If MsgBox("The file " & Filename & " already exists, do you want to overwrite it?", _
        vbYesNo, "Saving the file ...") = vbNo Then Exit Sub
    Kill Filename
End If

Screen.MousePointer = vbHourglass

Set MyDoc = GetObject(App.Path & "\NOTICE.DOC")
'Set MyDoc = New Word.Document
'Set MyDoc = Word.Application.Documents.Open(FileName:=App.Path & "\NOTICE.DOC")
'Word.Application.Visible = False
MyDoc.Activate

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Name", ReplaceWith:=txtMemberName, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    '.Execute FindText:="LoanType", ReplaceWith:=FormatField(m_rstScheme("SchemeName")), _
        Format:=True, Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Loan", ReplaceWith:=txtBalance.Caption, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Date", ReplaceWith:=FormatField(rstMaster("IssueDate")), Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Balance", ReplaceWith:=txtBalance.Caption, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Interest", ReplaceWith:=txtRegInterest, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Penal", ReplaceWith:=txtPenalInterest, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Installment", ReplaceWith:=Me.txtBalance, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Total", ReplaceWith:=txtTotal.Text, Format:=True, _
        Replace:=wdReplaceAll
End With

With ActiveDocument.content.Find
    .ClearFormatting
    .Font.Bold = True
    With .Replacement
        .ClearFormatting
        .Font.Bold = True
        .Font.name = gFontName
    End With
    .Execute findText:="Today", ReplaceWith:=txtRepayDate.Text, Format:=True, _
        Replace:=wdReplaceAll
End With

ActiveDocument.SaveAs Filename:=Filename
'''ActiveDocument.Close savechanges:=wdDoNotSaveChanges

MyDoc.Close
''Documents.Close savechanges:=wdDoNotSaveChanges
Set MyDoc = Nothing

Screen.MousePointer = vbNormal

End Sub
Private Sub cmdOk_Click()
Dim Cancel As Boolean
Call ResetUserIntereface

Unload Me
End Sub

Private Sub cmdPhoto_Click()
    If Not m_CustReg Is Nothing Then
        frmPhoto.setAccNo (m_CustReg.CustomerID)
        If (m_CustReg.CustomerID > 0) Then _
            frmPhoto.Show vbModal
    End If
End Sub

Private Sub cmdPrevTrans_Click()
    
    m_TransID = m_TransID - 10
    If m_TransID < 1 Then m_TransID = 1
    m_rstLoanTrans.Find "TransID >= " & m_TransID
    Call LoanLoadGrid

End Sub


Private Sub cmdPrint_Click()

If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
m_frmPrintTrans.Show vbModal

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


Private Sub cmdSB_Click()

If m_LoanID = 0 Then Exit Sub

'if the transaction is payment
    'then check for for the whether he is traferring
    'the amount withdrwan to the sb account
'if he not tranferring the amount to the SB then
'check whether he is withdrawing the amount
'through the cheque book if he has received

Dim SearchName As String
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT FirstName,MiddleName,LastName " & _
    " FROM NAMETAB WHERE CustomerID = " & m_CustReg.CustomerID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    SearchName = FormatField(rst(0))
    If SearchName = "" Then SearchName = FormatField(rst(1))
    If SearchName = "" Then SearchName = FormatField(rst(2))
End If
SearchName = InputBox("Enter the Name to search", "Account Search", SearchName)

Screen.MousePointer = vbHourglass
On Error GoTo Err_line
Dim Lret As Long

' Query the database to get all the customer names...
With gDbTrans
    .SqlStmt = "SELECT AccNum," & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name " & _
        " FROM SBMaster A, NameTab B WHERE A.CustomerID = B.CustomerID "
    If Trim(SearchName) <> "" Then
        .SqlStmt = .SqlStmt & " AND (FirstName like '" & SearchName & "%' " & _
            " Or MiddleName like '" & SearchName & "%' " & _
            " Or LastName like '" & SearchName & "%')"
        .SqlStmt = .SqlStmt & " Order by IsciName"
    Else
        .SqlStmt = .SqlStmt & " Order by AccNum"
    End If
    Lret = .Fetch(rst, adOpenStatic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        GoTo Exit_Line
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
        GoTo Exit_Line
    End If
    Screen.MousePointer = vbDefault
    ' Display the dialog.
    m_retVar = ""
    .Show vbModal
    txtSbAccNum = m_retVar
   End With

Exit_Line:
    Screen.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
        'MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
        MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
    End If
    GoTo Exit_Line

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
Dim rstFetch As ADODB.Recordset
Dim transType As wisTransactionTypes
Dim transAmount As Currency
Dim Balance As Currency
Dim TransDate As Date
Dim bankClass As clsBankAcc
Dim SBClass As clsSBAcc
Dim ContraID As Long


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
   gDbTrans.SqlStmt = "UpDate BkccMaster set LoanClosed = 0," & _
                    "InterestBalance = 0 where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to reopen the loan"
      MsgBox GetResourceString(536), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.CommitTrans
   'MsgBox "Account reopened successfully"
   MsgBox GetResourceString(522), vbInformation, wis_MESSAGE_TITLE
   
ElseIf cmdUndo.Caption = GetResourceString(14) Then  '"Delete
    ''19 June 2013
    Dim lastTransID As Long
    
   'nRet = MsgBox("Are you sure to delete this account ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(539), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   'START Shashi Angadi 18 Jun 2013
   ''Check the Transaction type to delete it from the Account Heads
   gDbTrans.SqlStmt = "Select A.Amount,A.Balance,A.TransType,A.TransID,A.TransDate From BKCCTrans A Where " & _
    " A.LoanID = " & m_LoanID & " order  by TransID desc"
    If gDbTrans.Fetch(rstFetch, adOpenForwardOnly) > 0 Then
        ''Check the Transtype
        transType = FormatField(rstFetch("TransType"))
        transAmount = FormatField(rstFetch("Amount"))
        Balance = FormatField(rstFetch("Balance"))
        TransDate = FormatField(rstFetch("TransDate"))
        lastTransID = FormatField(rstFetch("TransID"))
    Else
        nRet = MsgBox("Unable to delete this transaction ?", vbOKOnly, wis_MESSAGE_TITLE)
        Exit Sub
    End If
    
    Dim exitSub As Boolean
    
    If transType = wContraDeposit Then
        exitSub = True
    ElseIf transType = wContraWithdraw Then
        Set SBClass = New clsSBAcc
        Dim SBAccountID As Long
        Dim SbHeadID As Long
            
        Dim rstContra As Recordset
        'In case of contra transaction
        gDbTrans.SqlStmt = "Select * from ContraTrans Where ContraID = " & _
                "(SELECT distinct ContraID From ContraTrans " & _
                " WHERE AccHeadID = " & m_LoanHeadID & _
                " And Accid = " & m_LoanID & " And  TransID = " & lastTransID & " )"
        
        ''Check for the Relative SB Account of this transaction
        If gDbTrans.Fetch(rstContra, adOpenDynamic) > 0 Then
            Do While rstFetch.EOF
                ContraID = FormatField(rstContra("ContraID"))
                If FormatField(rstContra("AccHeadID")) <> m_LoanHeadID Then
                    SBAccountID = FormatField(rstContra("AccID"))
                    SbHeadID = FormatField(rstContra("AccHeadID"))
                    Exit Do
                End If
            Loop
        Else
            exitSub = True
        End If
        
        ''Check is this Amount is reversable in the SB trans
        If Not exitSub Then _
            If Not SBClass.IsContraTransactionRemovable(SBAccountID, transAmount, TransDate) Then exitSub = True
        
    End If
    
    If exitSub Then
            'Undo has to be done from where the amount is withdrawn
        MsgBox "Unable to delete the contra Deposit transaction", vbOKOnly, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
   'END Shashi Angadi 18 Jun 2013
   Dim InTrans As Boolean
   gDbTrans.BeginTrans
   InTrans = True
   Me.MousePointer = vbHourglass
   gDbTrans.SqlStmt = "Delete * From BKCCTrans Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then exitSub = True
   
   'Shashi 19 2013
   ''Correct the Amount in the Account Heads
   If Not exitSub Then
        If bankClass Is Nothing Then Set bankClass = New clsBankAcc
        If transType = wDeposit Then
            If Not bankClass.UndoCashDeposits(m_LoanHeadID, transAmount, TransDate) Then GoTo ExitLine
        ElseIf transType = wWithdraw Then
            If Not bankClass.UndoCashWithdrawls(m_LoanHeadID, transAmount, TransDate) Then GoTo ExitLine
        ElseIf transType = wContraWithdraw Then
            ''Remove the transaction od Account heads
            If Not bankClass.UndoContraTrans(m_LoanHeadID, SbHeadID, transAmount, TransDate) Then GoTo ExitLine
            
            ''Remove the Amount Debited to SB account
            If Not SBClass.UndoContraDepositAmount(SBAccountID, transAmount, TransDate) Then GoTo ExitLine
            
            'Delete the Transcation in ContraTransTable
            gDbTrans.SqlStmt = "Delete * from ContraTrans Where ContraID = " & ContraID
            If Not gDbTrans.SQLExecute Then GoTo ExitLine
            
        End If
   End If
   'Shashi 19 2013
   
   gDbTrans.SqlStmt = "Delete * From BKCcMaster Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then GoTo ExitLine
   
   gDbTrans.CommitTrans
   InTrans = False
   'MsgBox "Account reopened succefully"
   MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE
   Me.MousePointer = vbDefault
   'Exit Sub
End If

ExitLine:
If InTrans Then
      'MsgBox "Unable to delete the loan"
    gDbTrans.RollBack
    MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
End If

' Reload the loan details.
If Not m_rstLoanMast Is Nothing Then
    LoanLoad FormatField(m_rstLoanMast("loanid"))
End If
Me.MousePointer = vbDefault
End Sub
Private Function GetCustomerSBAccountNumber() As String
    GetCustomerSBAccountNumber = ""
    
    gDbTrans.SqlStmt = "SELECT AccNum From SBMASTER " & _
        " WHERE CustomerId = (SELECT CustomerID From BKCCMAster " & _
        " WHERE LoanID = " & m_LoanID & ");"
    
    Dim rst As Recordset
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then GetCustomerSBAccountNumber = FormatField(rst(0))

End Function
Private Function GetCustomerSBAccountID() As Long
    GetCustomerSBAccountID = 0
    
    gDbTrans.SqlStmt = "SELECT AccID From SBMASTER " & _
        " WHERE CustomerId = (SELECT CustomerID From BKCCMAster " & _
        " WHERE LoanID = " & m_LoanID & ");"
    Dim rst As Recordset
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then GetCustomerSBAccountID = FormatField(rst(0))

End Function

Private Sub cmdLoanUpdate_Click()
    Call LoanUpDate
End Sub

Private Sub cmdView_Click()
'Call SetKannadaCaption
' Validate the user input.
' Check for starting date.

Dim fromDate As String
Dim toDate As String

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
    If .Enabled Then
        If Not DateValidate(.Text, "/", True) Then
            'MsgBox "Enter a valid ending date.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(501), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtEndDate
            Exit Sub
        End If
        If txtStartDate.Enabled Then
            If WisDateDiff(.Text, txtStartDate.Text) > 0 Then
                MsgBox "invalid date specified", vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtEndDate
                Exit Sub
            End If
        End If
        toDate = .Text
    End If
End With



If cmbFarmer.ListIndex < 0 Then cmbFarmer.ListIndex = 0
'If cmbGender.ListIndex < 0 Then cmbGender.ListIndex = 0

Dim ReportType As wis_BKCCReports
'0,1,2,3,4,5,,6,7,16,17,18,19,21,22,23,24,
If optDepositReports.Value Then
    If optReports(8) Then ReportType = repBkccDepBalance
    If optReports(9) Then ReportType = repBKCCDepDayBook
    If optReports(14) Then ReportType = repBkccDepHolder
    If optReports(11) Then ReportType = repBkccDepIntPaid
    If optReports(15) Then ReportType = repBkccDepMonBal
    If optReports(10) Then ReportType = repBkccGuarantor
    If optReports(12) Then ReportType = repBkccDepDailyCash
    If optReports(13) Then ReportType = repBkccDepGLedger
    If optReports(20) Then ReportType = repBkccDepMonTrans

Else
    If optReports(0) Then ReportType = repBKCCLoanBalance
    If optReports(1) Then ReportType = repBKCCLoanDayBook
    If optReports(2) Then ReportType = repBkccOD
    If optReports(3) Then ReportType = repBKCCLoanIntCol
    If optReports(4) Then ReportType = repBKCCLoanDailyCash
    If optReports(5) Then ReportType = repBKCCLoanGLedger
    If optReports(6) Then ReportType = repBkccLoanHolder
    If optReports(7) Then ReportType = repBkccLoanMonBal
    If optReports(16) Then ReportType = repBkccMonTrans
    If optReports(17) Then ReportType = repBKCCMonthlyRegister
    If optReports(18) Then ReportType = repBKCCShedule_1
    If optReports(19) Then ReportType = repBKCCMemberTrans
    If optReports(21) Then ReportType = repBKCCReceivable
    If optReports(22) Then ReportType = repBKCCLoanIssued
    If optReports(23) Then ReportType = repBKCCLoanReturned
    If optReports(24) Then ReportType = repBKCCLoanClaimBill
    If optReports(25) Then ReportType = repBKCCLoanClaimBill_Yearly
    If optReports(26) Then ReportType = repBKCCLoanClaimBill_PrevYearly
End If

gCancel = 0

If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption

RaiseEvent ShowReport(ReportType, IIf(optAccId, wisByAccountNo, wisByName), _
        cmbFarmer.ItemData(cmbFarmer.ListIndex), fromDate, toDate, _
        m_clsRepOption)

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
Select Case KeyCode
    Case vbKeyUp
        ' Scroll up.
        With Me.vscLoanIssue
            If .Value - .SmallChange > .Min Then
                .Value = .Value - .SmallChange
            Else
                .Value = .Min
            End If
        End With
    Case vbKeyDown
        ' Scroll down.
        With vscLoanIssue
            If .Value + .SmallChange < .Max Then
                .Value = .Value + .SmallChange
            Else
                .Value = .Max
            End If
        End With
   Case vbKeyTab
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
End Select


End Sub

Private Sub Form_Load()
Dim rst As Recordset
Dim cmbIndex As Integer
Dim recCount As Integer

    Screen.MousePointer = vbHourglass
    'set icon for the form caption
    Me.Icon = LoadResPicture(161, vbResIcon)
    cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
    'SetKannadaCaption
    Call SetKannadaCaption
    'Centre the form
    Call CenterMe(Me)
    
    PropInitializeForm
    fraLoanGrid.Visible = True
    Me.TabStrip2.Tabs(2).Selected = True
    'm_Notes.ModuleID = wis_BKCCLoan
   
'Member Type
cmbIndex = GetIndex("MemberType")
cmbIndex = Val(ExtractToken(lblLoanIssue(Val(GetIndex("MemberType"))).Tag, "TextIndex"))
cmb(cmbIndex).Clear
Call LoadMemberTypes(cmb(cmbIndex))
If cmb(cmbIndex).ListCount <= 1 Then cmb(cmbIndex).Locked = True

'Report Combo
With cmbFarmer
    .Clear
    .AddItem GetResourceString(338, 378)
    .ItemData(.newIndex) = 0
    .AddItem GetResourceString(379, 378)
    .ItemData(.newIndex) = SmallFarmer
    .AddItem GetResourceString(380, 378)
    .ItemData(.newIndex) = BigFarmer
    .AddItem GetResourceString(381, 378)
    .ItemData(.newIndex) = MarginalFarmer
End With

'Now Load the farmer Type
On Error Resume Next
cmbIndex = Val(ExtractToken(lblLoanIssue(Val(GetIndex("FarmerType"))).Tag, "TextIndex"))
With cmb(cmbIndex)
    .Clear
    .AddItem GetResourceString(379, 378)
    .ItemData(.newIndex) = SmallFarmer
    .AddItem GetResourceString(380, 378)
    .ItemData(.newIndex) = BigFarmer
    .AddItem GetResourceString(381, 378)
    .ItemData(.newIndex) = MarginalFarmer
End With

On Error GoTo 0
         
gDbTrans.SqlStmt = "SELECT * FROM Install WHERE KeyData = 'NOTICE'"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then cmdNotice.Visible = True
Screen.MousePointer = vbDefault
txtRepayDate.Text = gStrDate
txtEndDate = txtRepayDate

With cmbTrans
    .Clear
    .AddItem GetResourceString(271)
    .AddItem GetResourceString(272)
    .AddItem GetResourceString(43, 47)
End With


If gOnLine Then
    txtRepayDate.Locked = True
    cmdRepayDate.Enabled = False
End If

'Load the Particulars
Call LoadParticularsFromFile(cmbParticulars, App.Path & "\BKCC.ini")

cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)

cmdAbn.Tag = CInt((gCurrUser.UserPermissions And perBankAdmin) Or (gCurrUser.UserPermissions And perOnlyWaves))

'Now Load the Properties of the KCC Loan
Dim IntClass As New clsInterest
Dim SetUp As New clsSetup
Dim I As Integer
Dim pos As Integer
Dim SchemeName As String

Do
    'First Get the Limit range for Fro slab
    SchemeName = SetUp.ReadSetupValue("KCCLoanInt", "Slab" & I, "0")
    If SchemeName = "0" Then Exit Do
    
    'Now Get the LImit range
    pos = InStr(1, SchemeName, "-")
    txtMinLimit(I).Caption = Left(SchemeName, pos - 1)
    txtMaxLimit(I).Value = Val(Mid(SchemeName, pos + 1))
    
    'Now Get the LOan Int Rate for this range
    SchemeName = Val(IntClass.InterestRate(wis_BKCCLoan, "Slab" & I, FinUSFromDate))
    txtLoanIntRate(I).Text = SchemeName
    I = I + 1
Loop

txtSubsidyRate.Text = Val(IntClass.InterestRate(wis_BKCCLoan, "Subsidy", FinUSFromDate))
txtRebateRate.Text = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate", FinUSFromDate))
txtNabardRate.Text = Val(IntClass.InterestRate(wis_BKCCLoan, "NabardRate", FinUSFromDate))

txtRebateRate1.Text = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate1", FinUSFromDate))
txtRebateRate2.Text = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate2", FinUSFromDate))

fraReports.Width = fraProps.Width
Me.Width = TabStrip.Width + 350
fraDepReports.Left = 80
fraDepReports.Top = optReports(0).Top - 100
fraDepReports.Width = fraReports.Width - 180
fraDepReports.BorderStyle = 0
cmdPhoto.Enabled = Len(gImagePath)

optReports(1).Value = True
optReports(9).Value = True

optReports(0).Value = True
optReports(8).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Report form.
If Not m_LookUp Is Nothing Then
    Unload m_LookUp
    Set m_LookUp = Nothing
End If
gWindowHandle = 0
Set m_CustReg = Nothing
Set m_Notes = Nothing

RaiseEvent WindowClosed

End Sub

Private Sub lblLoanIssue_DblClick(Index As Integer)
On Error Resume Next
txtLoanIssue(Index).SetFocus

End Sub


Private Sub m_frmAbn_WindowClosed()
Set m_frmAbn = Nothing
End Sub

Private Sub m_frmLoanReport_Initialise(Min As Long, Max As Long)

On Error Resume Next
With frmCancel
    If Max <> 0 Then
        .PicStatus.Visible = True
        .Refresh
    End If
End With

End Sub

Private Sub m_frmLoanReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS: " & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
End With
End Sub


Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 4
Dim curPrintRow As Integer
Dim TransID As Long

'1. Fetch last print row from sb master table.
'First get the last printed txnID From the SbMaster
SqlStr = "SELECT  LastPrintRow From BKCCMaster WHERE LoanId = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintRow = IIf(IsNull(metaRst("LastPrintrow")), 0, metaRst("LastPrintrow"))

'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From BKCCTrans WHERE LoanId = " & m_LoanID
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    
    ' If there are no records to print, since the last printed txn,
    ' display a message and exit.
        If (rst(0) = 0) Then
            MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
    
     'Print [or don't print] header part
     'lastPrintRow = IIf(IsNull(Rst("LastPrintRow")), 0, Rst("LastPrintRow"))
         If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
          'clsPrint.newPage
           clsPrint.isNewPage = True
         End If
  '3. Getting matching records for passbook printing

SqlStr = "SELECT LoanId,TransDate,TransId, 0 as IntAmount, 0 as PenalIntAmount, Amount, Balance, TransType, Particulars, 'PRINCIPAL' as Tabletype,Deposit" & _
    " FROM BKCCTrans WHERE LoanId = " & m_LoanID & _
    " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
    " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
    " UNION " & _
    " SELECT LoanId,TransDate,TransId,IntAmount,PenalIntAmount, MiscAmount as Amount, IntBalance as Balance, TransType, Particulars, 'INTEREST' as TableType,Deposit" & _
    " FROM BKCCIntTrans WHERE LoanId = " & m_LoanID & _
    " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
    " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
    " ORDER BY TransID"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Set clsPrint = New clsTransPrint

'Printer.PaperSize = 9
Printer.Font.name = gFontName
Printer.Font.Size = 9 'gFontSize

With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 9
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .ColHeader(1) = GetResourceString(37) 'Date
    '.ColWidth(1) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(2) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(3) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(4) = 8: .COlHeader(4) = GetResourceString(344) 'Interest
    '.ColWidth(5) = 8: .COlHeader(5) = GetResourceString(345) 'Interest
    '.ColWidth(6) = 15: .COlHeader(6) = GetResourceString(42) 'Balance
    
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
        .ColWidth(0) = 9
        .ColWidth(1) = 5
        .ColWidth(2) = 5
        .ColWidth(3) = 5
        .ColWidth(4) = 5
        .ColWidth(5) = 3
        .ColWidth(6) = 17
        .ColWidth(7) = 5
        .ColWidth(8) = 5
        .ColWidth(9) = 5
    
       Dim bHeaderPrinted As Boolean
  bHeaderPrinted = False
  
  '''
  Dim Receipt As Currency
  Dim payment As Currency
  Dim IntAmount As Currency
  Dim MicAmount As Currency
  Dim TransDate As String
  Dim Particulars As String
  Dim PenalInt As Currency
  Dim Balance As Currency
  Dim Total As Currency
  Dim DepReceipt As Currency
  Dim DepPayment As Currency
  Dim DepBalance As Currency
  
       
       
    While Not rst.EOF
        If .isNewPage Then
            If (bHeaderPrinted = False) Then
               ' .printHead
                .isNewPage = False
                bHeaderPrinted = True
                End If
        Else
            '.printHead
            '.isNewPage = False
        End If
       
       If TransID <> 0 And TransID <> FormatField(rst("TransID")) Then
         .ColText(0) = TransDate
         .ColText(1) = Particulars
         .ColText(2) = Receipt
         .ColText(3) = payment
         .ColText(4) = IntAmount
         .ColText(5) = PenalInt
         .ColText(6) = Balance
         .ColText(7) = DepReceipt
         .ColText(8) = DepPayment
         .ColText(9) = DepBalance
        ' .ColText(6) = Total
         If Receipt <> 0 Then
         .ColText(2) = Receipt
        Else
         .ColText(2) = " "
        End If
           
        If payment <> 0 Then
         .ColText(3) = payment
        Else
         .ColText(3) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(4) = IntAmount
        Else
         .ColText(4) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(5) = PenalInt
        Else
         .ColText(5) = " "
        End If
      
        If DepReceipt <> 0 Then
         .ColText(7) = DepReceipt
        Else
         .ColText(7) = " "
        End If
        
        If DepPayment <> 0 Then
         .ColText(8) = DepPayment
        Else
         .ColText(8) = " "
        End If
        
        If DepBalance <> 0 Then
         .ColText(9) = DepBalance
        Else
         .ColText(9) = " "
        End If
         
          Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & Balance
         .PrintRow2
         
         
          ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE1) Then
            .newPage
            For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
           ' MsgBox "plz insert new page"
            curPrintRow = 1
        End If
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    
    
    TransID = FormatField(rst("TransID"))
        
        If rst("TableType") = "INTEREST" Then
       
            TransDate = FormatField(rst("TransDate"))
            Particulars = FormatField(rst("Particulars"))
            'MicAmount = FormatField(Rst("Amount"))
            IntAmount = FormatField(rst("IntAmount"))
            PenalInt = FormatField(rst("PenalIntAmount"))
            Balance = FormatField(rst("Balance"))
            'Total = IntAmount + PenalInt + Receipt
            
        ElseIf rst("TableType") = "PRINCIPAL" Then
            TransDate = FormatField(rst("TransDate"))
            
            If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
                   payment = FormatField(rst("Amount"))
            Else
                   Receipt = FormatField(rst("Amount"))
            End If
            Particulars = FormatField(rst("Particulars"))
            Balance = FormatField(rst("Balance"))
            
            'If FormatField(Rst("Deposit")) Then
                DepReceipt = Receipt ': Receipt = 0
                DepPayment = payment ': payment = 0
                DepBalance = Balance ': Balance = 0
                'Debug.Print DepReceipt & "|" & DepPayment & "|" & DepBalance
            'End If
               Debug.Print DepReceipt & "|" & DepPayment & "|" & DepBalance
               'Total = IntAmount + PenalInt + payment
               Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & DepBalance
        End If
           
       ' .PrintRows
        
        
        rst.MoveNext
       
    Wend
    
    If TransID > 0 Then
         .ColText(0) = TransDate
         .ColText(1) = Particulars
         .ColText(2) = Receipt
         .ColText(3) = payment
         .ColText(4) = IntAmount
         .ColText(5) = PenalInt
         .ColText(6) = Balance
         .ColText(7) = DepReceipt
         .ColText(8) = DepPayment
         .ColText(9) = DepBalance
         If Receipt <> 0 Then
         .ColText(2) = Receipt
        Else
         .ColText(2) = " "
        End If
           
        If payment <> 0 Then
         .ColText(3) = payment
        Else
         .ColText(3) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(4) = IntAmount
        Else
         .ColText(4) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(5) = PenalInt
        Else
         .ColText(5) = " "
        End If
      
        If DepReceipt <> 0 Then
         .ColText(7) = DepReceipt
        Else
         .ColText(7) = " "
        End If
        
        If DepPayment <> 0 Then
         .ColText(8) = DepPayment
        Else
         .ColText(8) = " "
        End If
        
        If DepBalance <> 0 Then
         .ColText(9) = DepBalance
        Else
         .ColText(9) = " "
        End If
        
         
          Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & Balance
         'Debug.Print Receipt & " | " & payment & " | " & IntAmount
         .PrintRow2
          Receipt = 0: payment = 0:
          IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    .newPage
End With
Printer.EndDoc
Set rst = Nothing
Set clsPrint = Nothing
'Now Update the Last Print Id to the master
SqlStr = "UPDATE BKCCMaster set LastPrintID =" & TransID & _
           ", LastPrintRow = " & curPrintRow - 1 & _
           " Where LoanId = " & m_LoanID
        
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If
End Sub
    
Private Sub m_frmPrintTrans_TransClick(bNewPassbook As Boolean)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintId, lastPrintRow As Integer
Const HEADER_ROWS = 6
Dim curPrintRow As Integer
       
'First get the last printed txnId and last printed row From the BkCCMaster
SqlStr = "SELECT  LastPrintID, LastPrintRow From BKCCMaster WHERE LoanId = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintID"))
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If IsNull(metaRst("LastPrintRow")) And lastPrintId = 1 Then lastPrintId = 0


' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From BkccTrans WHERE LoanId = " & m_LoanID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
If (rst(0) = 0) Then
    
    Dim iRet As Integer
    iRet = MsgBox("There are no transactions available for printing." & vbCrLf & _
    "Do you want to reset priting for this account?", vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE)
    If (iRet = vbYes) Then
        'Now Update the Last Print Id to the master
        SqlStr = "UPDATE BKCCMaster set LastPrintID =" & TransID & _
                   ", LastPrintRow = " & curPrintRow - 1 & _
                   " Where LoanId = " & m_LoanID
                
        gDbTrans.BeginTrans
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
        Else
            gDbTrans.CommitTrans
        End If
    End If
    
    
    Exit Sub
End If


'First grecords for txns From the BKCCTrans and BkccIntTrans
SqlStr = "SELECT LoanId,TransDate,TransId, 0 as IntAmount, 0 as PenalIntAmount, Amount, Balance, TransType, Particulars, 'PRINCIPAL' as Tabletype,Deposit " & _
    " FROM BKCCTrans WHERE LoanId = " & m_LoanID & _
    " UNION " & _
    " SELECT LoanId, TransDate,TransId,IntAmount,PenalIntAmount, MiscAmount as Amount, IntBalance as Balance, TransType, Particulars, 'INTEREST' as TableType,Deposit" & _
    " FROM BKCCIntTrans WHERE LoanId = " & m_LoanID & _
    " ORDER BY TransID"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


 clsPrint.isNewPage = True
'Print [or don't print] header part
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE1 - 1) Then
   ' clsPrint.newPage
    clsPrint.isNewPage = True
    
End If


Set clsPrint = New clsTransPrint
'Printer.PaperSize = 9
Printer.Font = gFontName
Printer.Font.Size = 9 'gFontSize
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 9
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .ColHeader(1) = GetResourceString(37) 'Date
    '.ColWidth(1) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(2) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(3) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(4) = 8: .COlHeader(4) = GetResourceString(344) 'Interest
    '.ColWidth(5) = 8: .COlHeader(5) = GetResourceString(345) 'Interest
    '.ColWidth(6) = 15: .COlHeader(6) = GetResourceString(42) 'Balance
    
     If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE1) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
        For count = 0 To (HEADER_ROWS + 3 + lastPrintRow)
            Printer.Print ""
        Next count
    End If
    
      ' column widths for printing txn rows.
        .ColWidth(0) = 8
        .ColWidth(1) = 10
        .ColWidth(2) = 6
        .ColWidth(3) = 7
        .ColWidth(4) = 4
        .ColWidth(5) = 4
        .ColWidth(6) = 20
        .ColWidth(7) = 7
        .ColWidth(8) = 6
        .ColWidth(9) = 5
        
       
    Dim bHeaderPrinted As Boolean
  bHeaderPrinted = False
  
  '''
  Dim Receipt As Currency
  Dim payment As Currency
  Dim IntAmount As Currency
  Dim MicAmount As Currency
  Dim TransDate As String
  Dim Particulars As String
  Dim PenalInt As Currency
  Dim Balance As Currency
  Dim Total As Currency
  Dim DepReceipt As Currency
  Dim DepPayment As Currency
  Dim DepBalance As Currency
  
       
       
    While Not rst.EOF
        If .isNewPage Then
            If (bHeaderPrinted = False) Then
               ' .printHead
                .isNewPage = False
                bHeaderPrinted = True
            End If
        Else
            '.printHead
            '.isNewPage = False
        End If
       
       If TransID <> 0 And TransID <> FormatField(rst("TransID")) Then
         .ColText(0) = TransDate
         .ColText(1) = Particulars
         .ColText(2) = Receipt
         .ColText(3) = payment
         .ColText(4) = IntAmount
         .ColText(5) = PenalInt
         .ColText(6) = Balance
         .ColText(7) = DepReceipt
         .ColText(8) = DepPayment
         .ColText(9) = DepBalance
        ' .ColText(6) = Total
        
        If Receipt <> 0 Then
         .ColText(2) = Receipt
        Else
         .ColText(2) = " "
        End If
           
        If payment <> 0 Then
         .ColText(3) = payment
        Else
         .ColText(3) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(4) = IntAmount
        Else
         .ColText(4) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(5) = PenalInt
        Else
         .ColText(5) = " "
        End If
      
        If DepReceipt <> 0 Then
         .ColText(7) = DepReceipt
        Else
         .ColText(7) = " "
        End If
        
        If DepPayment <> 0 Then
         .ColText(8) = DepPayment
        Else
         .ColText(8) = " "
        End If
        
        If DepBalance <> 0 Then
         .ColText(9) = DepBalance
        Else
         .ColText(9) = " "
        End If
   
   Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & Balance
         .PrintRow2
         GoTo tmpLabel:
          ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE1) Then
            .newPage
            For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
           ' MsgBox "plz insert new page"
            curPrintRow = 1
        End If
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    
    
    TransID = FormatField(rst("TransID"))
        
        If rst("TableType") = "INTEREST" Then
       
            TransDate = FormatField(rst("TransDate"))
            TransDate = Left(TransDate, Len(TransDate) - 5)
            Particulars = FormatField(rst("Particulars"))
            'MicAmount = FormatField(Rst("Amount"))
            IntAmount = FormatField(rst("IntAmount"))
            PenalInt = FormatField(rst("PenalIntAmount"))
            Balance = FormatField(rst("Balance"))
        
            'Total = IntAmount + PenalInt + Receipt
            
        ElseIf rst("TableType") = "PRINCIPAL" Then
            TransDate = FormatField(rst("TransDate"))
            TransDate = Left(TransDate, Len(TransDate) - 5)
            If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
                   payment = FormatField(rst("Amount"))
            Else
                   Receipt = FormatField(rst("Amount"))
            End If
            Particulars = FormatField(rst("Particulars"))
            Balance = FormatField(rst("Balance"))
            
            'If FormatField(Rst("Deposit")) Then
                DepReceipt = Receipt ': Receipt = 0
                DepPayment = payment ': payment = 0
                DepBalance = Balance ': Balance = 0
                'Debug.Print DepReceipt & "|" & DepPayment & "|" & DepBalance
            'End If
               Debug.Print DepReceipt & "|" & DepPayment & "|" & DepBalance
               'Total = IntAmount + PenalInt + payment
               Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & DepBalance
        End If
           
       ' .PrintRows
        
        
        rst.MoveNext
       
    Wend
    If TransID > 0 Then
         .ColText(0) = TransDate
         .ColText(1) = Particulars
         .ColText(2) = Receipt
         .ColText(3) = payment
         .ColText(4) = IntAmount
         .ColText(5) = PenalInt
         .ColText(6) = Balance
         .ColText(7) = DepReceipt
         .ColText(8) = DepPayment
         .ColText(9) = DepBalance
         
         
         'To skip tha 0's not to print Mrudula 25th july 2014.....
         If Receipt <> 0 Then
         .ColText(2) = Receipt
        Else
         .ColText(2) = " "
        End If
           
        If payment <> 0 Then
         .ColText(3) = payment
        Else
         .ColText(3) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(4) = IntAmount
        Else
         .ColText(4) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(5) = PenalInt
        Else
         .ColText(5) = " "
        End If
      
        If DepReceipt <> 0 Then
         .ColText(7) = DepReceipt
        Else
         .ColText(7) = " "
        End If
        
        If DepPayment <> 0 Then
         .ColText(8) = DepPayment
        Else
         .ColText(8) = " "
        End If
        
        If DepBalance <> 0 Then
         .ColText(9) = DepBalance
        Else
         .ColText(9) = " "
        End If
         
         
         
         Debug.Print TransDate & "|" & Particulars & "|" & Receipt & " | " & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & DepReceipt & "|" & DepPayment & "|" & Balance
         'Debug.Print Receipt & " | " & payment & " | " & IntAmount
         
         .PrintRow2
         
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    .newPage
End With
tmpLabel:
Printer.EndDoc
Exit Sub
Set rst = Nothing
Set clsPrint = Nothing
'Now Update the Last Print Id to the master
SqlStr = "UPDATE BKCCMaster set LastPrintID =" & TransID & _
           ", LastPrintRow = " & curPrintRow - 1 & _
           " Where LoanId = " & m_LoanID
        
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If
End Sub

Private Sub m_frmReportLoan_Initialise(Min As Long, Max As Long)
On Error Resume Next
With frmCancel
    UpdateStatus frmCancel.PicStatus, 0
    If Max <> 0 Then
        .PicStatus.Visible = True
        .Refresh
    End If
End With
End Sub

Private Sub m_LookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub


Private Sub optDepositReports_Click()
    fraDepReports.Visible = True
    fraDepReports.ZOrder 0
End Sub

Private Sub optLoanReports_Click()
    fraDepReports.Visible = False
    fraDepReports.ZOrder 1
End Sub

Private Sub optReports_Click(Index As Integer)

Dim Dt As Boolean
Dim Amt As Boolean
Dim Cst As Boolean

If Index = 0 Then Amt = True: Cst = True
If Index = 1 Then Dt = True: Amt = True: Cst = True
If Index = 2 Then Amt = True: Cst = True
If Index = 3 Then Dt = True: Cst = True
If Index = 4 Then Dt = True: Cst = True: Amt = True
If Index = 5 Then Dt = True
If Index = 6 Then Cst = True: Amt = True

If Index = 7 Then Cst = False: Dt = True
If Index = 8 Then Amt = True: Cst = True
If Index = 9 Then Dt = True: Amt = True: Cst = True
'If Index = 10 Then Dt = false
If Index = 11 Then Dt = True: Amt = True
If Index = 12 Then Dt = True: Cst = True: Amt = True
If Index = 13 Then Dt = True
If Index = 14 Then Amt = True: Cst = True
If Index = 18 Then Cst = True
If Index = 19 Then Amt = True: Cst = True: Dt = True

If Index = 22 Or Index = 23 Or Index = 24 Or Index = 25 Or Index = 26 Then Amt = True: Cst = True: Dt = True


If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
With m_clsRepOption
    .EnableCasteControls = Cst
    .EnableAmountRange = Amt
End With

txtStartDate.Enabled = Dt
txtStartDate.BackColor = IIf(Dt, wisWhite, wisGray)
cmdStDate.Enabled = Dt

cmbFarmer.Enabled = Cst

cmbFarmer.BackColor = IIf(Cst, wisWhite, wisGray)

End Sub

Private Sub optReports_DblClick(Index As Integer)
If Index <> 18 Then Exit Sub
Dim strReport As String
strReport = optReports(18).Caption
strReport = InputBox("Enter the report Name", "Report Name", strReport)
If Len(strReport) > 5 Then _
        optReports(18).Caption = strReport

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
On Error Resume Next
Select Case UCase$(TabStrip.SelectedItem.Key)
    Case "LOANACCOUNTS"
        fraLoanAccounts.ZOrder 0
        txtIndex = GetIndex("MemberID")
        txtLoanIssue(txtIndex).SetFocus
        cmdLoanSave.Default = cmdLoanSave.Enabled
        cmdLoanUpdate.Default = cmdLoanUpdate.Enabled
        
        fraReports.Visible = False
        fraRepayments.Visible = False
        fraLoanAccounts.Visible = True
        fraProps.Visible = False
        
    Case "REPAYMENTS"
        fraRepayments.ZOrder 0
        cmdAccept.Default = True
        
        fraReports.Visible = False
        fraRepayments.Visible = True
        fraLoanAccounts.Visible = False
        fraProps.Visible = False
        
    Case "REPORTS"
        fraReports.ZOrder 0
        cmdView.Default = True
        
        fraReports.Visible = True
        fraRepayments.Visible = False
        fraLoanAccounts.Visible = False
        fraProps.Visible = False

    Case "PROPS"
        fraProps.ZOrder 0
        cmdApply.Default = True
        
        fraReports.Visible = False
        fraRepayments.Visible = False
        fraLoanAccounts.Visible = False
        fraProps.Visible = True
        
End Select

End Sub

Private Sub TabStrip2_Click()
    
    If TabStrip2.SelectedItem.Index = 1 Then
        fraLoanGrid.Visible = False
        fraInstructions.Visible = True
        fraInstructions.ZOrder 0
    Else
        fraLoanGrid.Visible = True
        fraInstructions.Visible = False
        fraLoanGrid.ZOrder 0
    End If
End Sub

Private Sub txtAccNo_Change()
If Trim$(txtAccNo.Text) <> "" Then
    cmdLoad.Enabled = True
    txtRepayAmt.Enabled = True
    txtRepayDate.Enabled = True

Else
    cmdLoad.Enabled = False
    txtRepayAmt.Enabled = False
    txtRepayDate.Enabled = False
End If
Call ResetUserIntereface
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



Private Sub txtIntBalance_Change()
On Error Resume Next
If ActiveControl.name <> txtTotal.name Then _
    txtTotal = (txtIntBalance.Value + txtRegInterest.Value + _
            txtPenalInterest.Value + txtMiscAmount.Value) * Val(cmbTrans.Tag) + txtRepayAmt
Err.Clear
End Sub


Private Sub txtIntBalance_GotFocus()
With txtIntBalance
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtLoanIntRate_GotFocus(Index As Integer)
    On Error Resume Next
    With txtLoanIntRate(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtLoanIssue_Change(Index As Integer)

' In case of fields bound to loanduedate and
' instalment mode, compute the instalment amount
' and display it.
Dim strDataSrc As String
Dim Amt As Currency
Dim txtIndex As Integer

strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")
If StrComp(strDataSrc, "LoanDueDate", vbTextCompare) = 0 Or _
        StrComp(strDataSrc, "InstalmentMode", vbTextCompare) = 0 Then
'    Amt = ComputeInstalmentAmount
    
    ' Display it in the relevant fieldbox.
    txtIndex = GetIndex("InstalmentAmount")
    txtLoanIssue(txtIndex).Text = Amt
End If

End Sub
Private Sub txtLoanIssue_DblClick(Index As Integer)
Dim strDispType As String
' Get the display type.
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")

If StrComp(strDispType, "List", vbTextCompare) = 0 Then _
            txtLoanIssue_KeyPress Index, vbKeyReturn


End Sub
Private Sub txtLoanIssue_GotFocus(Index As Integer)

lblLoanIssue(Index).ForeColor = vbBlue
PropSetIssueDescription lblLoanIssue(Index)

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
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "textindex")
    If TextIndex <> "" Then cmdLoanIssue(Val(TextIndex)).Visible = True
End If


' Hide all other command buttons...
Dim I As Integer
For I = 0 To cmdLoanIssue.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmdLoanIssue(I).Visible = False
    End If
Next

' Hide all other combo boxes.
cmdLoanSave.Default = cmdLoanSave.Enabled
cmdLoanUpdate.Default = cmdLoanUpdate.Enabled
If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    cmdLoanSave.Default = False
    cmdLoanUpdate.Default = False
    TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "textindex")
    ' Get the cmdbutton index.
    On Error Resume Next
    If TextIndex <> "" Then
        If cmb(Val(TextIndex)).Visible Then Exit Sub
        cmb(Val(TextIndex)).Visible = True
        cmb(Val(TextIndex)).SetFocus
    End If
End If

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
            cmb(Val(strIndex)).ZOrder
            cmb(Val(strIndex)).SetFocus
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
Dim rst As Recordset
Dim memberTYpe As Integer
Dim cmbIndex As Integer

' Get the name of the data source bound to this control.
strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")
Dim MMObj As New clsMMAcc
    
Select Case UCase(strDataSrc)
  Case "MEMBERID"
' For the textbox bound to MemberID, we need to
' fill in the corresponding member name.
    txtIndex = GetIndex("MemberType")
    cmbIndex = Val(ExtractToken(lblLoanIssue(txtIndex).Tag, "TextIndex"))
    If cmb(cmbIndex).ListIndex >= 0 Then memberTYpe = cmb(cmbIndex).ItemData(cmb(cmbIndex).ListIndex)
    ' Get the name of the member for the given memberid.
    txtLoanIssue(Index + 1).Text = MMObj.MemberName(txtLoanIssue(Index).Text, memberTYpe)
      
  Case "GUARANTOR1", "GUARANTOR2"
    ' Get the index of the textbox bound to thisfield.
    If IsNumeric(txtLoanIssue(Index).Text) Then
        gDbTrans.SqlStmt = "SELECT A.CustomerID,Title + ' ' + FirstName " & _
            " + ' ' + MiddleName +' '+ LastName as Name " & _
            " From MemMaster A, NameTab B Where " & _
            " AccNum = " & AddQuotes(txtLoanIssue(Index).Text, True) & _
            " ANd A.CustomerID = B.CustomerID "
        If gDbTrans.Fetch(rst, adOpenDynamic) Then
            txtLoanIssue(Index).Text = FormatField(rst("Name"))
            txtLoanIssue(Index).Tag = PutToken(txtLoanIssue(Index).Tag, _
                "CustomerID", FormatField(rst("CustomerID")))
        End If
    End If
End Select

Set MMObj = Nothing
    
'for Loan Installment Vinay

End Sub

Private Sub txtMaxLimit_GotFocus(Index As Integer)
On Error Resume Next
With txtMaxLimit(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMaxLimit_LostFocus(Index As Integer)

If Index >= txtMaxLimit.UBound Then Exit Sub
'If maximum limit is 0
'then it is the limit for all above slabs
'then need not to put the limit further
If txtMaxLimit(Index).Value = 0 Then _
    txtEffectiveDate.TabIndex = txtMaxLimit(Index).TabIndex + 1: Exit Sub


'If the text box is not the last Text Box
'Then Add the maximum Limit
'as minimum limit for next Slab
    
txtMinLimit(Index + 1).Caption = txtMaxLimit(Index).Value + 1

End Sub


Private Sub txtMemberID_Change()
If Trim$(txtMemberID.Text) <> "" Then
    cmdMember.Enabled = True
    txtRepayAmt.Enabled = True
    txtRepayDate.Enabled = True
Else
    cmdMember.Enabled = False
    txtRepayAmt.Enabled = False
    txtRepayDate.Enabled = False
End If
Call ResetUserIntereface

End Sub



Private Sub txtMiscAmount_Change()
Call txtIntBalance_Change
End Sub

Private Sub txtMiscAmount_GotFocus()
With txtMiscAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtPenalInterest_Change()
Call txtIntBalance_Change
End Sub

Private Sub txtPenalInterest_GotFocus()
With txtPenalInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtRegInterest_Change()
    Call txtIntBalance_Change
End Sub

Private Sub txtRegInterest_GotFocus()
With txtRegInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

'Me.ActiveControl.SelStart = 0
'Me.ActiveControl.SelLength = Len(Me.ActiveControl)
End Sub


Private Sub txtRepayAmt_Change()
    Call txtIntBalance_Change
End Sub

Private Sub txtRepayAmt_GotFocus()
With txtRepayAmt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtRepayDate_Change()

If m_LoanID = 0 Then txtRepayDate.Tag = "": Exit Sub

If Not DateValidate(txtRepayDate.Text, "/", True) Then Exit Sub
If txtRepayDate.Tag = txtRepayDate.Text Then Exit Sub

Dim TotalAmount As Currency
Dim IntAmount As Currency
Dim TransDate As Date

TransDate = GetSysFormatDate(txtRepayDate)

TotalAmount = Val(txtTotal.Text)
'Get the Balance Interest if Any
txtIntBalance.Tag = FormatField(m_rstLoanMast("InterestBalance"))
txtIntBalance.Text = txtIntBalance.Tag 'FormatCurrency(FormatField(m_rstLoanMast("InterestBalance")))
txtLoanIssue(Val(GetIndex("InterestBalance"))) = txtIntBalance.Text
txtLastInt = FormatField(m_rstLoanMast("LastIntDate"))

'Compute and display the regular interest.
If chkSb.Value = vbChecked Then
    IntAmount = BKCCRegularInterest(TransDate, m_LoanID)
    If IntAmount < 0 Then IntAmount = BKCCDepositInterest(m_LoanID, TransDate)
    lblRegInterest = GetResourceString(IIf(IntAmount < 0, 233, 344))
End If
'If IntAmount > 0 And cmbTrans.ListIndex = 2 Then cmbTrans.ListIndex = -1

txtRegInterest.Tag = IntAmount \ 1
txtRegInterest.ToolTipText = IntAmount \ 1
txtRegInterest = Abs(IntAmount \ 1)

'Compute and display the penal interest for defaulted payments.
IntAmount = BKCCPenalInterest(TransDate, m_LoanID)
txtPenalInterest.Tag = IntAmount / 1
If IntAmount >= 0 Then txtPenalInterest = IntAmount \ 1
'Display the total instalment amount.

'txtTotal = FormatCurrency(TotalAmount)
'Now Calulate the Principal amount
txtRepayAmt = txtTotal - txtIntBalance - txtRegInterest - txtPenalInterest - txtMiscAmount

'Now Load the Same Loan detials into loan issue loan grid
txtRepayDate.Tag = txtRepayDate.Text

cmdAbn.Enabled = Val(cmdAbn.Tag)

End Sub

Private Sub txtRepayDate_GotFocus()
With txtRepayDate
    .SelStart = 0
    .SelLength = InStr(1, .Text, "/")
End With

End Sub


Private Sub txtStartAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtStartDate_GotFocus()
On Error Resume Next
With txtStartDate
    .SelStart = 0
    .SelLength = InStr(1, .Text, "/") - 1
End With

End Sub


Private Sub txtTotal_Change()

On Error Resume Next
If ActiveControl.name = txtTotal.name Then _
    txtRepayAmt.Value = txtTotal.Value - (txtIntBalance.Value + _
            txtRegInterest.Value + txtPenalInterest.Value + txtMiscAmount.Value) * Val(cmbTrans.Tag)

Err.Clear

End Sub

Private Sub txtTotal_GotFocus()
With txtTotal
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub vscLoanIssue_Change()
' Move the picLoanissueSlider.
picLoanIssueSlider.Top = -vscLoanIssue.Value
End Sub

Private Sub SetKannadaCaption()
'to trap an error
Call SetFontToControlsSkipGrd(Me)


'Raise the error if occurs
On Error GoTo 0
'Set kannada caption for the generally used  controls
cmdOk.Caption = GetResourceString(1)

'Set captions for the all the tabs
TabStrip.Tabs(1).Caption = GetResourceString(216)
TabStrip.Tabs(2).Caption = GetResourceString(80) & " " & _
                            GetResourceString(36) & GetResourceString(92)
TabStrip.Tabs(3).Caption = GetResourceString(283) & GetResourceString(92)
TabStrip.Tabs(4).Caption = GetResourceString(213)

TabStrip2.Tabs(1).Caption = GetResourceString(219)
TabStrip2.Tabs(2).Caption = GetResourceString(218)
    
'Set Kannada/English captions for the Repayments
lblMisc.Caption = GetResourceString(327)
lblLoanAccNo.Caption = GetResourceString(58) & " " & _
    GetResourceString(60) 'Loan No
lblMemberId.Caption = GetResourceString(49) & " " & _
    GetResourceString(60)
cmdLoad.Caption = GetResourceString(3)
lblName.Caption = GetResourceString(35)
lblLoanAmount.Caption = GetResourceString(58)
lblIssueDate.Caption = GetResourceString(340)
lblBalance.Caption = GetResourceString(67, 40)  'Interest Balance
lblLastInt.Caption = GetResourceString(391, 47, 37)
lblRegInterest.Caption = GetResourceString(344)
lblPenalInterest.Caption = GetResourceString(345)
lblIntBalance.Caption = GetResourceString(67, 47)
lblTotInst.Caption = GetResourceString(52, 40) '"Total Amount
lblRepayAmt.Caption = GetResourceString(310, 40) '"Principal Amount
lblParticulars.Caption = GetResourceString(39)  '"Particulars

cmdUndo.Caption = GetResourceString(5)
'cmdInstalment.Caption = GetResourceString(57)
cmdAccept.Caption = GetResourceString(4)
chkSb.Caption = GetResourceString(421, 271)
lblChequeNo.Caption = GetResourceString(275, 60) '"Sb acc NUm

lblfarmer.Caption = GetResourceString(378, 253)
lblRepayDate.Caption = GetResourceString(216, 37)
lblTrans.Caption = GetResourceString(38) 'Transction

'Set the Caption create account
cmdLoanSave.Caption = GetResourceString(7) 'Save
cmdLoanUpdate.Caption = GetResourceString(171) 'pdate
cmdLoanClear.Caption = GetResourceString(8)  'Clear
cmdPhoto.Caption = GetResourceString(415)

'Set kannada caption for Reports tab
optReports(0).Caption = GetResourceString(80, 42) 'Loan Balance
optReports(1).Caption = GetResourceString(390, 63) 'Sub Day Book
optReports(2).Caption = GetResourceString(84, 18) 'Over Due Loans
optReports(3).Caption = GetResourceString(80, 483) 'Loan Interest Recceived
optReports(4).Caption = GetResourceString(390, 85) 'Loan Cash Book
optReports(5).Caption = GetResourceString(80, 93) 'Loan generalledger
optReports(6).Caption = GetResourceString(83)  'Loan List
optReports(7).Caption = GetResourceString(80, 463, 42) 'Monthly Balance

optReports(8).Caption = GetResourceString(43, 42) 'Deposit Balance
optReports(9).Caption = GetResourceString(43, 390, 63) 'Deposit Sub day Book
optReports(10).Caption = GetResourceString(389, 49, 295)
optReports(11).Caption = GetResourceString(43, 487) 'Deposit Interest Recceived
optReports(12).Caption = GetResourceString(43, 390, 85) 'Deposit Cash Book
optReports(13).Caption = GetResourceString(43, 93) 'Deposit general
optReports(14).Caption = GetResourceString(43)  'Deposit Transaction
optReports(15).Caption = GetResourceString(43, 463, 42) 'Monthly Balance
optReports(16).Caption = GetResourceString(231) 'Monthly Payament,Receipt,PAyament

optReports(17).Caption = GetResourceString(463, 430) 'Monthly Register
optReports(19).Caption = GetResourceString(49, 28) 'Member Transction
optReports(20).Caption = GetResourceString(43, 231) 'Monthly Payament,Receipt,PAyament
optReports(21).Caption = GetResourceString(237, 364) 'Other Receivables
optReports(22).Caption = GetResourceString(290)  'LOan Issue
optReports(23).Caption = GetResourceString(82) 'repayment Made
optReports(24).Caption = GetResourceString(414, 438, 430)
optReports(25).Caption = GetResourceString(374, 431, 438, 430)
optReports(26).Caption = GetResourceString(250, 431, 438, 430)

optAccId.Caption = GetResourceString(68)
optName.Caption = GetResourceString(69)
'lblGender.Caption = GetResourceString(125)

fraReports.Caption = GetResourceString(283) & GetResourceString(92)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
'lblAmt1.Caption = GetResourceString(147,42)
'lblAmt2.Caption = GetResourceString(148,42)
cmdView.Caption = GetResourceString(13)
lblfarmer.Caption = GetResourceString(378, 253)
'lblPlace.Caption = GetResourceString(112)
'lblCaste.Caption = GetResourceString(111)
cmdNotice.Caption = "Notice"


'Captions for Properties frame
fraProps.Caption = GetResourceString(213)
lblMinLimit.Caption = GetResourceString(147, 42)
lblMaxLimit.Caption = GetResourceString(148, 42)
lblLoanIntRate.Caption = GetResourceString(186)

lblEffectiveDate.Caption = GetResourceString(38, 37)
txtEffectiveDate.Text = gStrDate
cmdApply.Caption = GetResourceString(6)
cmdAdvance.Caption = GetResourceString(491)    'Options

optLoanReports.Caption = GetResourceString(80, 283)
optDepositReports.Caption = GetResourceString(43, 283)

lblSubsidyRate.Caption = GetResourceString(414, 436, 186)
lblRebateRate.Caption = GetResourceString(414, 437, 186)
lblNabardRate.Caption = "Nabard " & GetResourceString(437, 186)

lblRebateRate1.Caption = GetResourceString(431, 437, 186) & " 1"
lblRebateRate2.Caption = GetResourceString(431, 437, 186) & " 2"

End Sub

