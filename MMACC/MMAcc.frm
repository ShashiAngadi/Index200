VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmMMAcc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX-2000   -   Members Module"
   ClientHeight    =   7635
   ClientLeft      =   1545
   ClientTop       =   780
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   415
      Left            =   6450
      TabIndex        =   30
      Top             =   7020
      Width           =   1215
   End
   Begin VB.Frame fraReports 
      Height          =   5985
      Left            =   270
      TabIndex        =   80
      Top             =   720
      Width           =   7245
      Begin VB.Frame fraMemType 
         Height          =   1980
         Left            =   165
         TabIndex        =   88
         Top             =   3420
         Width           =   6975
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   430
            Left            =   5640
            TabIndex        =   99
            Top             =   1400
            Width           =   1215
         End
         Begin VB.TextBox txtDate1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1590
            TabIndex        =   93
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtDate2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5160
            TabIndex        =   98
            Top             =   810
            Width           =   1305
         End
         Begin VB.CommandButton cmdDate1 
            Caption         =   "..."
            Height          =   315
            Left            =   2850
            TabIndex        =   92
            Top             =   840
            Width           =   315
         End
         Begin VB.CommandButton cmdDate2 
            Caption         =   "..."
            Height          =   315
            Left            =   6540
            TabIndex        =   97
            Top             =   810
            Width           =   315
         End
         Begin VB.ComboBox cmbRepMemType 
            Height          =   315
            Left            =   1590
            TabIndex        =   95
            Top             =   1440
            Width           =   2415
         End
         Begin VB.OptionButton optName 
            Caption         =   " By name "
            Height          =   315
            Left            =   3540
            TabIndex        =   90
            Top             =   330
            Width           =   2115
         End
         Begin VB.OptionButton optMemId 
            Caption         =   "By Member Id :"
            Height          =   315
            Left            =   360
            TabIndex        =   89
            Top             =   270
            Value           =   -1  'True
            Width           =   2250
         End
         Begin VB.Label lblDate2 
            Caption         =   "and before (dd/mm/yyyy)"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3300
            TabIndex        =   96
            Top             =   870
            Width           =   1845
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            TabIndex        =   91
            Top             =   870
            Width           =   1485
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6900
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblMemType 
            Caption         =   "Member Type"
            Height          =   315
            Left            =   90
            TabIndex        =   94
            Top             =   1470
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Default         =   -1  'True
         Height          =   430
         Left            =   5910
         TabIndex        =   44
         Top             =   5500
         Width           =   1215
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report"
         Height          =   3345
         Left            =   165
         TabIndex        =   81
         Top             =   195
         Width           =   6975
         Begin VB.OptionButton optReports 
            Caption         =   "no-Loan members List"
            Height          =   315
            Index           =   11
            Left            =   3630
            TabIndex        =   110
            Top             =   2760
            Width           =   2850
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Loan Memebr List"
            Height          =   315
            Index           =   10
            Left            =   300
            TabIndex        =   109
            Top             =   2760
            Width           =   3000
         End
         Begin VB.OptionButton optReports 
            Caption         =   "General Ledger"
            Height          =   315
            Index           =   9
            Left            =   300
            TabIndex        =   108
            Top             =   1830
            Width           =   2865
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Show Monthly Balance"
            Height          =   315
            Index           =   8
            Left            =   3630
            TabIndex        =   107
            Top             =   2310
            Width           =   3000
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Share Certificate List"
            Height          =   315
            Index           =   7
            Left            =   3630
            TabIndex        =   106
            Top             =   1815
            Width           =   3000
         End
         Begin VB.OptionButton optReports 
            Caption         =   "List of members as on"
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   40
            Top             =   2310
            Width           =   2850
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Member Balances where"
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   2865
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Sub day book"
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   38
            Top             =   870
            Width           =   2865
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Sub Cash Book"
            Height          =   315
            Index           =   3
            Left            =   300
            TabIndex        =   39
            Top             =   1350
            Width           =   2865
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Members Admitted"
            Height          =   315
            Index           =   4
            Left            =   3630
            TabIndex        =   41
            Top             =   360
            Width           =   2955
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Members Cancelled"
            Height          =   315
            Index           =   5
            Left            =   3630
            TabIndex        =   42
            Top             =   840
            Width           =   2955
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Member/Share  Fee collected"
            Height          =   315
            Index           =   6
            Left            =   3630
            TabIndex        =   43
            Top             =   1335
            Width           =   3000
         End
      End
   End
   Begin VB.Frame fraProps 
      Height          =   5985
      Left            =   270
      TabIndex        =   78
      Top             =   720
      Width           =   7245
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   415
         Left            =   5790
         TabIndex        =   74
         Top             =   5490
         Width           =   1215
      End
      Begin VB.Frame fraPropMemCharges 
         Height          =   1125
         Left            =   255
         TabIndex        =   79
         Top             =   225
         Width           =   6735
         Begin VB.TextBox txtPropShareValue 
            Height          =   350
            Left            =   5670
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox txtPropShareFee 
            Height          =   350
            Left            =   5670
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txtPropMemFee 
            Height          =   350
            Left            =   2520
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox txtPropMemCancel 
            Height          =   350
            Left            =   2520
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblPropShareValue 
            Caption         =   "Share Value :"
            Height          =   300
            Left            =   3540
            TabIndex        =   54
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label lblPropShareFee 
            Caption         =   "Share Fee :"
            Height          =   300
            Left            =   3540
            TabIndex        =   52
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lblPropMemFee 
            Caption         =   "Membership Fee :"
            Height          =   300
            Left            =   180
            TabIndex        =   48
            Top             =   255
            Width           =   1800
         End
         Begin VB.Label lblPropMemCancel 
            Caption         =   "MemberShip Cancellation : "
            Height          =   300
            Left            =   165
            TabIndex        =   50
            Top             =   660
            Width           =   2115
         End
      End
      Begin VB.Frame fraPropDivedend 
         Caption         =   "Divedend"
         Height          =   3975
         Left            =   240
         TabIndex        =   56
         Top             =   1425
         Width           =   6765
         Begin VB.TextBox txtPropNextDivedendOn 
            Height          =   315
            Left            =   5400
            TabIndex        =   61
            Top             =   690
            Width           =   1245
         End
         Begin VB.CommandButton cmdDividend 
            Caption         =   "Add Dividend"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4980
            TabIndex        =   68
            Top             =   1920
            Width           =   1665
         End
         Begin VB.CommandButton cmdUndoInterests 
            Caption         =   "Undo Dividend"
            Enabled         =   0   'False
            Height          =   400
            Left            =   4110
            TabIndex        =   71
            Top             =   2580
            Width           =   2535
         End
         Begin VB.TextBox txtUndoInterest 
            Height          =   315
            Left            =   2055
            TabIndex        =   70
            Top             =   2610
            Width           =   1245
         End
         Begin VB.TextBox txtToDate 
            Height          =   315
            Left            =   5400
            TabIndex        =   65
            Top             =   1110
            Width           =   1245
         End
         Begin VB.TextBox txtFromDate 
            Height          =   315
            Left            =   2055
            TabIndex        =   63
            Top             =   1155
            Width           =   1245
         End
         Begin VB.TextBox txtPropDivedend 
            Height          =   315
            Left            =   2535
            TabIndex        =   58
            Text            =   "0.00"
            Top             =   270
            Width           =   735
         End
         Begin ComctlLib.ProgressBar prg 
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   1980
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Line Line4 
            X1              =   6720
            X2              =   0
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Label txtPropLastDivedendOn 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2070
            TabIndex        =   104
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label txtFailAccID 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   90
            TabIndex        =   73
            Top             =   3510
            Width           =   6555
         End
         Begin VB.Label lblUndoInterest 
            Caption         =   "Undo interest added on: "
            Height          =   300
            Left            =   150
            TabIndex        =   69
            Top             =   2610
            Width           =   2025
         End
         Begin VB.Label lblFailAccIDs 
            Caption         =   "Accounts where undo was not possible:"
            Height          =   300
            Left            =   150
            TabIndex        =   72
            Top             =   3180
            Width           =   5850
         End
         Begin VB.Label lblStatus 
            Caption         =   "x"
            Height          =   300
            Left            =   180
            TabIndex        =   66
            Top             =   1590
            Width           =   5535
         End
         Begin VB.Label lblToDate 
            Caption         =   "ToDate :"
            Height          =   300
            Left            =   3510
            TabIndex        =   64
            Top             =   1170
            Width           =   1785
         End
         Begin VB.Label lblFromDate 
            Caption         =   "FromDate : "
            Height          =   300
            Left            =   120
            TabIndex        =   62
            Top             =   1185
            Width           =   1815
         End
         Begin VB.Label lblpropDivedendRate 
            Caption         =   "Divedend Rate :"
            Height          =   300
            Left            =   120
            TabIndex        =   57
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label lblPropLastDivedendOn 
            Caption         =   "Last divedend debited on : "
            Height          =   300
            Left            =   105
            TabIndex        =   59
            Top             =   765
            Width           =   1875
         End
         Begin VB.Label lblPropNextDivedendOn 
            Caption         =   "Last divedend due on : "
            Height          =   300
            Left            =   3525
            TabIndex        =   60
            Top             =   750
            Width           =   1815
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   6630
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11695
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transactions"
            Key             =   "Transactions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New / Modify Account"
            Key             =   "AddModify"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Properties"
            Key             =   "Properties"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTransact 
      Height          =   6000
      Left            =   270
      TabIndex        =   82
      Top             =   720
      Width           =   7245
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   315
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   1
         Top             =   210
         Width           =   1065
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1650
         Width           =   2535
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5880
         TabIndex        =   26
         Top             =   5550
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1230
         TabIndex        =   8
         Top             =   1200
         Width           =   1365
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3870
         TabIndex        =   27
         Top             =   5550
         Width           =   1935
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   400
         Left            =   2400
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.ComboBox cmbLeaves 
         Height          =   315
         Left            =   1230
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Frame fraPassBook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2265
         Left            =   270
         TabIndex        =   86
         Top             =   3120
         Width           =   6705
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1830
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grdTrans 
            Height          =   2205
            Left            =   90
            TabIndex        =   23
            Top             =   90
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   5
            AllowUserResizing=   1
         End
         Begin VB.CommandButton cmdNextTrans 
            Height          =   375
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   600
            Width           =   435
         End
         Begin VB.CommandButton cmdPrevTrans 
            Height          =   375
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   90
            Width           =   435
         End
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2265
         Left            =   270
         TabIndex        =   83
         Top             =   3120
         Width           =   6705
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   90
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2055
            Left            =   90
            TabIndex        =   84
            Top             =   90
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   3625
            _Version        =   393217
            TextRTF         =   $"MMAcc.frx":0000
         End
      End
      Begin VB.CommandButton cmdLeaves 
         Caption         =   "Leaves..."
         Height          =   375
         Left            =   2970
         TabIndex        =   12
         Top             =   2075
         Width           =   795
      End
      Begin VB.Frame fraShare 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2265
         Left            =   270
         TabIndex        =   100
         Top             =   3120
         Width           =   6705
         Begin VB.CommandButton cmdSharePrint 
            Height          =   375
            Left            =   6210
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1800
            Width           =   435
         End
         Begin VB.CommandButton cmdSharePrev 
            Height          =   375
            Left            =   6210
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   105
            Width           =   435
         End
         Begin VB.CommandButton cmdShareNext 
            Height          =   375
            Left            =   6210
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   510
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grdShare 
            Height          =   2205
            Left            =   90
            TabIndex        =   101
            Top             =   90
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   5
            AllowUserResizing=   1
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2865
         Left            =   180
         TabIndex        =   5
         Top             =   2640
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   5054
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Instructions"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Transactions"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Certificate details"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin WIS_Currency_Text_Box.CurrText txtShareFees 
         Height          =   345
         Left            =   5700
         TabIndex        =   14
         Top             =   1170
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   5700
         TabIndex        =   17
         Top             =   1620
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtTotal 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5700
         TabIndex        =   20
         Top             =   2070
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1230
         TabIndex        =   105
         Top             =   660
         Width           =   5805
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   90
         X2              =   6990
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   210
         X2              =   7110
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance : Rs. 00.00"
         Height          =   300
         Left            =   4740
         TabIndex        =   3
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label lblMemNo 
         Caption         =   "Member No. : "
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblName 
         Caption         =   "Name(s) : "
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   300
         Left            =   150
         TabIndex        =   10
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs.) : "
         Height          =   300
         Left            =   4020
         TabIndex        =   16
         Top             =   1650
         Width           =   1395
      End
      Begin VB.Label lblDate 
         Caption         =   "Date : "
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   825
      End
      Begin VB.Label lblLeaves 
         Caption         =   "Leaves:"
         Height          =   300
         Left            =   210
         TabIndex        =   11
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblShareFee 
         Caption         =   "Share fees:"
         Height          =   300
         Left            =   4020
         TabIndex        =   13
         Top             =   1230
         Width           =   1395
      End
      Begin VB.Label lblTOtal 
         Caption         =   "Total (Rs.) :"
         Height          =   300
         Left            =   4020
         TabIndex        =   19
         Top             =   2100
         Width           =   1365
      End
   End
   Begin VB.Frame fraNew 
      Height          =   6000
      Left            =   270
      TabIndex        =   28
      Top             =   720
      Width           =   7245
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   400
         Left            =   5940
         TabIndex        =   111
         Top             =   3720
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   990
         Left            =   150
         ScaleHeight     =   930
         ScaleWidth      =   5655
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   315
         Width           =   5715
         Begin VB.Image imgNewAcc 
            Height          =   375
            Left            =   135
            Stretch         =   -1  'True
            Top             =   60
            Width           =   345
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   1020
            TabIndex        =   47
            Top             =   360
            Width           =   4620
         End
         Begin VB.Label lblHeading 
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
            TabIndex        =   46
            Top             =   45
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5940
         TabIndex        =   33
         Top             =   4260
         Width           =   1215
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   4290
         Left            =   150
         ScaleHeight     =   4230
         ScaleWidth      =   5655
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1335
         Width           =   5715
         Begin VB.VScrollBar VScroll1 
            Height          =   4185
            Left            =   5430
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picSlider 
            Height          =   2625
            Left            =   -30
            ScaleHeight     =   2565
            ScaleWidth      =   5400
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   -30
            Width           =   5460
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2340
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   -30
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   285
               Index           =   0
               Left            =   4830
               TabIndex        =   75
               Top             =   30
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.TextBox txtData 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
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
               Left            =   2460
               TabIndex        =   29
               Top             =   0
               Width           =   2910
            End
            Begin VB.Label txtPrompt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Account Holder"
               ForeColor       =   &H80000008&
               Height          =   350
               Index           =   0
               Left            =   30
               TabIndex        =   87
               Top             =   0
               Width           =   2385
            End
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   5940
         TabIndex        =   32
         Top             =   5265
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   5940
         TabIndex        =   31
         Top             =   4755
         Width           =   1215
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
         Height          =   195
         Left            =   135
         TabIndex        =   77
         Top             =   5670
         Width           =   1230
      End
   End
   Begin VB.Label lblMemberTypeName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Member Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2670
      TabIndex        =   112
      Top             =   7080
      Width           =   1635
   End
End
Attribute VB_Name = "frmMMAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------------------------
Private m_AccID As Long
Private m_rstPassBook As Recordset
Private m_rstShareBook As Recordset
Private m_CustReg As New clsCustReg
Private m_Notes As New clsNotes
Private M_setUp As clsSetup
Private m_MemberType As Integer
Private m_MemberTypeName As String
Private m_MemberTypeNameEnglish As String
Private m_MemberLocked As Boolean
Private M_ModuleID As wisModules

Const CTL_MARGIN = 15
Private m_accUpdatemode As wis_DBOperation

Private m_clsRepOption As clsRepOption

Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmMMShare As frmMMShare
Attribute m_frmMMShare.VB_VarHelpID = -1
Private m_frmdividend As frmIntPayble
Public Event AccountChanged(ByVal AccId As Long)
Public Event WindowClosed()
Public Event ShowReport(ReportType As wis_MemReports, ReportOrder As wis_ReportOrder, _
        memberTYpe As wis_MemberType, fromDate As String, toDate As String, RepOptionClass As clsRepOption)
       
Public Property Let memberTYpe(NewValue As Integer)
    
If NewValue = 0 Or m_MemberType = NewValue Then Exit Property
    
m_MemberType = NewValue

Dim rstMemType As Recordset
Dim txtIndex  As Integer
Dim cmbIndex As Integer
    
    gDbTrans.SqlStmt = "Select * From MemberTypeTab where MemberType = " & m_MemberType
    M_ModuleID = wis_Members + m_MemberType
    
    If gDbTrans.Fetch(rstMemType, adOpenDynamic) > 0 Then
        'CLEAR all the controls
        ResetUserInterface
        'Load the Membertype details
        m_MemberTypeName = FormatField(rstMemType("MemberTypeName"))
        m_MemberTypeNameEnglish = FormatField(rstMemType("MemberTypeNameEnglish"))
        lblMemberTypeName.Caption = m_MemberTypeName
        lblMemberTypeName.Tag = m_MemberTypeNameEnglish
        
        txtIndex = GetIndex("AccNum")
        txtData(txtIndex).Text = GetNewAccountNumber
        
        'MemType
        txtIndex = GetIndex("MemberType")
        txtData(txtIndex).Text = m_MemberTypeName
        cmbIndex = ExtractToken(txtPrompt(txtIndex).Tag, "TextIndex")
        Call SetComboIndex(cmb(cmbIndex), , m_MemberType)
        cmb(cmbIndex).Locked = True
        Call SetComboIndex(cmbRepMemType, , m_MemberType)
        cmbRepMemType.Locked = True
    End If
    
End Property

Public Property Get memberTYpe() As Integer
    memberTYpe = m_MemberType
End Property
Public Property Let SingleMemberModule(NewValue As Boolean)
    m_MemberLocked = NewValue
    If m_MemberLocked Then lblMemberTypeName.Visible = False
End Property

Public Property Get SingleMemberModule() As Boolean
    SingleMemberModule = m_MemberLocked
End Property
       
'Save The Account
Private Function AccountSave() As Boolean

Dim txtIndex As Byte
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim TransDate As Date
Dim MemFee As Currency
Dim PrevMemFee As Double
Dim PrevTransDate As Date

Dim strAccNum As String

' Check for a valid Account number.
    txtIndex = GetIndex("AccNum")
    strAccNum = Trim$(txtData(txtIndex).Text)
    
    
    'Check for the existance of the MemberNum
    gDbTrans.SqlStmt = "Select * FROM MemMaster " & _
        " Where AccNum = " & AddQuotes(strAccNum)
    If m_MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & m_MemberType
    
    If m_accUpdatemode = Update Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND AccID <> " & m_AccID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        '"This Account number already exists"
        MsgBox GetResourceString(545), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If
    
    'Check for the Account Group
    If GetAccGroupID = 0 Then
        'MsgBox "You have not selected the Account Group", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(749), vbInformation, wis_MESSAGE_TITLE
        txtIndex = GetIndex("AccGroup")
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If

    'See if account already exists if it is new
    If m_accUpdatemode = Insert Then
        AccId = 1
        gDbTrans.SqlStmt = "Select Max(AccID) from MemMaster"
        If gDbTrans.Fetch(rst, adOpenDynamic) Then AccId = FormatField(rst(0)) + 1
        
        'Check Whether This Customer already a member
        gDbTrans.SqlStmt = "Select * from MemMaster Where " & _
                            " CustomerID = " & m_CustReg.CustomerID
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            Debug.Print "Kannada"
            If MsgBox("This Customer already has Member Number" & _
                FormatField(rst("AccNum")) & vbCrLf & _
                "Do you want to continue!", vbYesNo + vbQuestion _
                + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then Exit Function
        End If
        'Now Check for the of existance of the asccount Number
        gDbTrans.SqlStmt = "Select * from MemMaster Where AccNum = " & AddQuotes(strAccNum)
        If m_MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & m_MemberType
    Else
        'PrevTransDate = Rst("CreateDate")
        gDbTrans.SqlStmt = "Select * from MemIntTrans Where " & _
                            " Accid = " & m_AccID & " AND TransID = 1"
        Call gDbTrans.Fetch(rst, adOpenDynamic)
        PrevMemFee = FormatField(rst("Amount"))
        PrevTransDate = rst("TransDate")
        AccId = m_AccID
        gDbTrans.SqlStmt = "Select * from MemMaster Where " & _
                        " AccNum = " & AddQuotes(strAccNum, True) & _
                        " And Accid <> " & m_AccID
    End If
    If gDbTrans.Fetch(rst, adOpenDynamic) >= 1 Then
        'MsgBox "Account number " & .Text & "already exists." & vbCrLf & vbCrLf & _
            "Please specify another account number !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(545), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

    ' Check for account holder name.
    txtIndex = GetIndex("AccName")
    With txtData(txtIndex)
        If Trim$(.Text) = "" Then
            'MsgBox "Member name not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(516), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With
    
    'Check if Member type is specified
    Dim memberTYpe As Byte
    If m_MemberType = 0 Then
        txtIndex = GetIndex("MemberType")
        With txtData(txtIndex)
            If Trim$(.Text) = "" Then
                'MsgBox "Member type not specified !", vbExclamation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(761), vbExclamation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            txtIndex = Val(ExtractToken(txtPrompt(txtIndex).Tag, "TextIndex"))
            memberTYpe = cmb(txtIndex).ItemData(cmb(txtIndex).ListIndex)
        End With
    Else
        memberTYpe = m_MemberType
    End If
    
    'Check For Mebership Fee 'Code By Shashi on 18/2/2000
    txtIndex = GetIndex("MemberFee")
    With txtData(txtIndex)
        If Trim(.Text) = "" Then
            'If MsgBox("You have not specified the Membersip Fee" & vbCrLf & vbCrLf & _
                "Do you want to Continue ?", vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            If MsgBox(GetResourceString(782) & vbCrLf & vbCrLf & _
                GetResourceString(541), vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
                ActivateTextBox txtData(txtIndex)
                Exit Function
            End If
        ElseIf Not IsNumeric(.Text) Then
            'MsgBox "Invalid amount specified", , wis_MESSAGE_TITLE
            MsgBox GetResourceString(506), , wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If
    End With
    
    ' Check for nominee name.
    Dim NomineeId As Long
    NomineeId = Val(txtData(GetIndex("NomineeID")))
    txtIndex = GetIndex("NomineeName")
        
    If NomineeId = 0 And txtData(txtIndex) = "" Then
        'MsgBox "Nominee name not specified!", _
                vbExclamation, wis_MESSAGE_TITLE
        If MsgBox(GetResourceString(558) & vbCrLf & GetResourceString(541), _
                vbYesNo + vbInformation, wis_MESSAGE_TITLE) = vbNo Then
                    ActivateTextBox txtData(txtIndex)
                    GoTo Exit_Line
        End If
    End If
    
    ' Check for nominee relationship.
    txtIndex = GetIndex("NomineeRelation")
    With txtData(txtIndex)
        If Trim$(.Text) = "" And NomineeId Then
            'MsgBox "Specify nominee relationship.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(559), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With

    txtIndex = GetIndex("IntroducerID")
    With txtData(txtIndex)
        ' Check if an introducerID has been specified.
        If Trim$(.Text) = "" Then
            'If MsgBox("No introducer has been specified!" _
                & vbCrLf & "Add this Member anyway?", vbQuestion + vbYesNo) = vbNo Then
            If MsgBox(GetResourceString(560) _
                & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
                GoTo Exit_Line
            End If
        Else
            ' Check if the introducer exists.
            gDbTrans.SqlStmt = "SELECT CustomerID FROM NameTab WHERE CustomerID = " & Val(.Text)
            If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
                'MsgBox "The introducer account number " & .Text & " is invalid.", _
                        vbExclamation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(514), _
                        vbExclamation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
        End If
    End With

'Check for a valid creation date
    If Not DateValidate(GetVal("CreateDate"), "/", True) Then
        'MsgBox "Invalid create date specified !" & vbCrLf & "Please specify in DD/MM/YYYY format!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & GetResourceString(573), vbExclamation, gAppName & " - Error"
        txtIndex = GetIndex("CreateDate")
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If

'Confirm before proceeding
    If m_accUpdatemode = Update Then
        'If MsgBox("This will update the account " & GetVal("AccID") _
                & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox(GetResourceString(520) & " " & GetVal("AccNum") _
                & "." & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
            GoTo Exit_Line
        End If
    ElseIf m_accUpdatemode = Insert Then
        'If MsgBox("This will create the new member." _
                & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox(GetResourceString(540) & " " & GetVal("AccNum") _
                 & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
            GoTo Exit_Line
        End If
    End If


'Start Transactions to Data base
    m_CustReg.ModuleId = wis_Members
    
    gDbTrans.BeginTrans
    If Not m_CustReg.SaveCustomer Then
        'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(617), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
    
    ' Insert/update to database.
    Dim SqlStmt As String
    Dim SqlSecond As String
    Dim transType As wisTransactionTypes
    Dim bankClass As clsBankAcc
    Dim headID As Long
    Set bankClass = New clsBankAcc
    
    headID = bankClass.GetHeadIDCreated(m_MemberTypeName & " " & GetResourceString(79, 191), _
                m_MemberTypeNameEnglish & " " & LoadResourceStringS(79, 191), parBankIncome, 0, M_ModuleID)
    
    TransDate = GetSysFormatDate(GetVal("CreateDate"))
    
    If m_accUpdatemode = Insert Then
        TransDate = GetSysFormatDate(GetVal("CreateDate"))
        MemFee = Val(GetVal("MemberFee"))
        
        SqlStmt = "Insert into MemMaster (AccNum,AccID, CustomerID, " & _
                " CreateDate, NomineeID,NomineeName,NomineeRelation, IntroducerID," & _
                "LedgerNo, FolioNo,MemberType,AccGroupID,UserID) " & _
                "values (" & AddQuotes(GetVal("AccNum"), True) & "," & _
                AccId & "," & _
                m_CustReg.CustomerID & "," & _
                "#" & TransDate & "#," & _
                NomineeId & ", " & _
                AddQuotes(GetVal("NomineeName"), True) & "," & _
                AddQuotes(GetVal("NomineeRelation"), True) & ", " & _
                Val(GetVal("IntroducerID")) & "," & _
                AddQuotes(GetVal("LedgerNo"), True) & ", " & _
                AddQuotes(GetVal("FolioNo"), True) & ", " & _
                memberTYpe & "," & _
                GetAccGroupID & "," & gUserID & " )"
        'Build Sql To Insert values into MemTrans
        transType = wDeposit
        SqlSecond = "Insert Into MemIntTrans(AccId,TransId,TransDate," & _
                " Amount,TransType,Balance,Particulars,UserId)" & _
                " Values(" & _
                AccId & ", 1, " & _
                "#" & TransDate & "#" & _
                ", " & MemFee & ", " & _
                transType & ", 0,'Memberdhip Fee'," & gUserID & " ) "
        
    ElseIf m_accUpdatemode = Update Then
        'First Undo the Member Fee
        If Not bankClass.UndoCashDeposits(headID, PrevMemFee, PrevTransDate) Then
            gDbTrans.RollBack
            Exit Function
        End If
        ' The user has selected updation.
        ' Build the SQL update statement.
        AccId = m_AccID
        MemFee = Val(GetVal("MemberFee"))
        TransDate = GetSysFormatDate(GetVal("CreateDate"))
        
        SqlStmt = "Update MemMaster set " & _
                " AccNum = " & AddQuotes(strAccNum) & ", " & _
                " NomineeID = " & NomineeId & ", " & _
                " NomineeName = " & AddQuotes(GetVal("NomineeName"), True) & ", " & _
                " NomineeRelation = " & AddQuotes(GetVal("NomineeRelation"), True) & "," & _
                " IntroducerID = " & Val(GetVal("IntroducerID")) & "," & _
                " LedgerNo = " & AddQuotes(GetVal("LedgerNo"), True) & "," & _
                " FolioNo = " & AddQuotes(GetVal("FolioNo"), True) & ", " & _
                " MemberType = " & memberTYpe & ", " & _
                " CreateDate = #" & TransDate & "#," & _
                " AccGroupId = " & GetAccGroupID & _
                " where AccID = " & m_AccID
        SqlSecond = "UpDate MemIntTrans Set Amount = " & MemFee & "," & _
            " TransDate = #" & TransDate & "#" & _
            " Where TransId = 1 and AccId = " & AccId
    End If
    ' Insert/update the data.
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStmt
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo Exit_Line
    End If
    gDbTrans.SqlStmt = SqlSecond
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo Exit_Line
    End If
    ''update the Member fee
    If Not bankClass.UpdateCashDeposits(headID, MemFee, TransDate) Then
        gDbTrans.RollBack
        Exit Function
    End If
        
    MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
    '"Saved the Member details."
    
    gDbTrans.CommitTrans
    AccountSave = True
    
Exit_Line:
    Exit Function

SaveAccount_error:
    If Err Then
        MsgBox "SaveAccount: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
    End If
    GoTo Exit_Line
    
End Function
Private Function GetAccGroupID() As Byte

Dim cmbIndex As Integer
cmbIndex = GetIndex("AccGroup")
If cmbIndex < 0 Then Exit Function
cmbIndex = Val(ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex"))
With cmb(cmbIndex)
    If .ListCount = 1 Then .ListIndex = 0
    If .ListIndex < 0 Then Exit Function
    GetAccGroupID = .ItemData(.ListIndex)
End With
End Function


Private Function AccountTransaction() As Boolean

Dim Amount As Currency
Dim I As Long
Dim FaceValue As Currency
Dim rst As ADODB.Recordset

'Prelim check
If m_AccID <= 0 Then
    'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
    cmdUndo.Enabled = False
    Exit Function
End If

'Check if account exists
Dim ClosedON As String
If Not AccountExists(m_AccID, ClosedON) Then
    'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If
If ClosedON <> "" Then
    'MsgBox "This account has been closed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(524), vbExclamation, gAppName & " - Error"
    Exit Function
End If

'Validate the date and assign to variable
Dim TransDate As Date
If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
    'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
Else
    TransDate = GetSysFormatDate(txtDate.Text)
End If

'Get the Transaction Type
Dim transType As wisTransactionTypes
If cmbTrans.ListIndex = -1 Then
    MsgBox GetResourceString(588), vbExclamation, gAppName & " - Error"
    cmbTrans.SetFocus
    Exit Function
Else
    With cmbTrans
        If .ListIndex = 0 Then transType = wDeposit
        If .ListIndex = 1 Then transType = wWithdraw
    End With
End If

'Check if leaves have been specified
If cmbLeaves.ListCount = 0 Then Call cmdLeaves_Click

If cmbLeaves.ListCount = 0 Then Exit Function
    
'The value of the shares needs to be determined previously
If transType = wDeposit Or transType = wContraDeposit Then
    If txtAmount = 0 Then
        'MsgBox "Unable to compute total share worth !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(764), vbExclamation, gAppName & " - Error"
        Exit Function
    Else
        Amount = txtAmount
    End If
ElseIf transType = wWithdraw Or transType = wContraWithdraw Then
    'YOu need to compute this using the database
    For I = 0 To cmbLeaves.ListCount - 1
        gDbTrans.SqlStmt = "Select FaceValue from ShareTrans " & _
            " where CertNo = " & AddQuotes(cmbLeaves.List(I), True)
        If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
            'MsgBox "Unable to locate the Certificate Number : " & cmbLeaves.List(i)
            MsgBox GetResourceString(587) & " " & cmbLeaves.List(I)
            Exit Function
        Else
            Amount = Amount + FormatField(rst("FaceValue"))
        End If
    Next I
End If

'Get the last TransID,TransDate and Balance
    Dim TransID As Long
    Dim Balance As Long
'Check for the last transaction date
If DateDiff("D", TransDate, GetMemberLastTransDate(m_AccID)) > 0 Then
    'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If

'get the last transaction Date
TransID = GetMemberMaxTransID(m_AccID) + 1

If Val(txtAmount) = 0 Then txtAmount = Amount
If (transType = wWithdraw Or transType = wContraWithdraw) Then
    If Amount = 0 Then Amount = txtAmount
    If Amount > 0 And Amount <> txtAmount Then
        If MsgBox("The amount you enteres & selected share amount are not equal" & _
                vbCrLf & GetResourceString(541), vbYesNo + vbQuestion _
                + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo _
            Then Exit Function
    End If
End If
If Val(txtAmount) > 0 Then Amount = txtAmount

'Balance = Amount
gDbTrans.SqlStmt = "Select TOP 1 Balance  " & _
    " from MemTrans where Accid = " & m_AccID & _
    " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Balance = FormatField(rst("Balance"))

If transType = wDeposit Or transType = wContraDeposit Then
    Balance = Balance + Amount
Else
    If Amount > Balance Then Amount = Balance
    Balance = Balance - Amount
End If

'Obtain the Face Value from current settings
Dim bankClass As clsBankAcc
Dim SetUp As New clsSetup
FaceValue = Val(SetUp.ReadSetupValue("MMAcc", "ShareValue", "10.00"))
Set SetUp = Nothing

Set bankClass = New clsBankAcc
Dim headID As Long
Dim FeeHeadId As Long
Dim headName As String
Dim headNameEnglish As String
Dim InTrans As Boolean

'Perform DB Transactions
gDbTrans.BeginTrans
InTrans = True

headName = m_MemberTypeName & " " & GetResourceString(53, 36) 'Share account
headNameEnglish = m_MemberTypeNameEnglish & " " & LoadResourceStringS(53, 36) 'Share account
headID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemberShare, 0, M_ModuleID)
'Create the share Fee
headName = GetResourceString(53, 191) 'Share Fee
headNameEnglish = LoadResourceStringS(53, 191) 'Share Fee
FeeHeadId = bankClass.GetHeadIDCreated(headName, headNameEnglish, parBankIncome, 0, M_ModuleID)

'First perform the transaction for Share fee(charges)
'Note that this will not have any effect on the Balance.
'IT is here only to determine the profit made by the bank
gDbTrans.SqlStmt = "INSERT into MemIntTrans (AccID, TransID, TransDate, " & _
                " Amount, TransType, Balance,UserID) values ( " & _
                m_AccID & "," & _
                TransID & ", " & _
                "#" & TransDate & "#, " & _
                txtShareFees & ", " & _
                wDeposit & ", " & _
                Balance & "," & gUserID & " )"
If TransID = 1 Or txtShareFees > 0 Then
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    If Not bankClass.UpdateCashDeposits(FeeHeadId, txtShareFees, TransDate) Then _
        GoTo Exit_Line
End If

'Insert / Update the the Share leaves table depending upon the transaction type
For I = 0 To cmbLeaves.ListCount - 1
    If transType = wDeposit Or transType = wContraDeposit Then
        'Insert the share numbers into the ShareLeaves tab
        gDbTrans.SqlStmt = "Insert into ShareTrans (AccID, SaleTransID, " & _
                " CertNo, FaceValue) values (" & _
                m_AccID & ", " & _
                TransID & ", " & _
                AddQuotes(cmbLeaves.List(I), True) & ", " & _
                FaceValue & ")"
    Else
        'Update the PurchaseTransID of the ShareLeaves
        gDbTrans.SqlStmt = "Update ShareTrans set ReturnTransID = " & TransID & _
            " where AccId = " & m_AccID & " And CertNo = " & AddQuotes(cmbLeaves.List(I), True)
    End If
    
    'Fire the query just built
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
Next I

If Balance = 0 Then
    ' If Balnce is 0 then set all the Shares as returned
    If transType = wWithdraw Or transType = wContraWithdraw Then
        gDbTrans.SqlStmt = "Update ShareTrans set ReturnTransID = " & _
            TransID & " where AccId = " & m_AccID & " And ReturnTransID = 0 "
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
End If

'Insert into MemTrans Tab
gDbTrans.SqlStmt = "INSERT INTO MemTrans (AccID, TransID, TransDate, " & _
            " Leaves, Amount, TransType, Balance,UserID) values ( " & _
            m_AccID & "," & _
            TransID & ", " & _
            "#" & TransDate & "#, " & _
            cmbLeaves.ListCount & ", " & _
            Amount & ", " & _
            transType & ", " & _
            Balance & "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then
    'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
    gDbTrans.RollBack
    Exit Function
End If

If transType = wDeposit Then
    If Not bankClass.UpdateCashDeposits(headID, Amount, TransDate) Then GoTo Exit_Line
ElseIf transType = wWithdraw Then
    If Not bankClass.UpdateCashWithDrawls(headID, Amount, TransDate) Then GoTo Exit_Line

    'If transaction is cash withdraw & there is casier window
    'then transfer the While Amount cashier window
    If gCashier Then
        Dim Cashclass As clsCash
        Set Cashclass = New clsCash
        If Cashclass.TransferToCashier(headID, _
                m_AccID, TransDate, TransID, Amount) < 1 Then GoTo Exit_Line
        Set Cashclass = Nothing
    End If
End If
    
'IF The Balance Falls Below Zero Ask For account Closure
If Balance <= 0 Then
    'If MsgBox(" This Transaction will Result In Negative Balance " & _
            "Do You Wish To Close The Account ", _
            vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
     If MsgBox(GetResourceString(656, 541), _
               vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
        gDbTrans.SqlStmt = "Update MemMaster set " & _
                " ClosedDate = #" & TransDate & "#" & _
                " where AccID = " & m_AccID
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
End If

gDbTrans.CommitTrans
InTrans = False

If Not txtAmount.Locked Then txtAmount.Locked = True

AccountTransaction = True

Exit Function

Exit_Line:
 
    If InTrans Then gDbTrans.RollBack
    InTrans = False
    'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"

End Function

Public Function AccountUndoLastTransaction() As Boolean

'Prelim check
If m_AccID <= 0 Then
    'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
    cmdUndo.Enabled = False
    Exit Function
End If

'Check if account exists
Dim ClosedON As String
If Not AccountExists(m_AccID, ClosedON) Then
    'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If
If ClosedON <> "" Then
    'if MsgBox ("This account has been closed ! " & vbCrLf & " Do You Wish To Reopen It ", vbExclamation, gAppName & " - Error"
    If MsgBox(GetResourceString(548), vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
        If Not AccountReopen(m_AccID) Then
            MsgBox GetResourceString(536), vbInformation, wis_MESSAGE_TITLE ' unable to reopen the account
        Else
            MsgBox GetResourceString(522), vbInformation, wis_MESSAGE_TITLE 'account reopened successfully
            AccountUndoLastTransaction = True
        End If
        Exit Function
     End If
End If
    
'Confirm UNDO
'If MsgBox("Are you sure you want to undo the last transaction ?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(583), vbYesNo + vbQuestion, _
                            gAppName & " - Error") = vbNo Then Exit Function
    
    
'Get last transaction record
Dim Amount As Currency
Dim ShareFee As Currency
Dim ret As Integer
Dim TransID As Long
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim rst As Recordset

'Get the Last Transction Id
TransID = GetMemberMaxTransID(m_AccID)

If TransID = 0 Then
    'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(645), vbInformation, gAppName & " - Error"
    Exit Function
End If

gDbTrans.SqlStmt = "Select TOP 1 * from MemTrans where " & _
                " AccID = " & m_AccID & " ANd TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Amount = FormatField(rst("Amount"))
    transType = rst("TransType")
    TransDate = rst("TransDate")
End If

gDbTrans.SqlStmt = "Select TOP 1 * from MemIntTrans where " & _
                " AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    transType = rst("TransType")
    ShareFee = FormatField(rst("amount"))
    TransDate = rst("TransDate")
End If

Dim headID As Long
Dim FeeHeadId As Long
Dim headName As String
    
headName = GetResourceString(53, 36) 'Share account
headID = GetIndexHeadID(headName)
'Create the share Fee
headName = GetResourceString(53, 191) 'Share Fee
FeeHeadId = GetIndexHeadID(headName)

  
If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & headID & _
            " And AccID = " & m_AccID & " And TransID = " & TransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
            AccountUndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If
    
    Dim bankClass As clsBankAcc

'Delete record from Data base
Dim InTrans As Boolean
gDbTrans.BeginTrans
InTrans = True

'Delete the last two records(one will be share purchase/sale and the other will be FEES)
gDbTrans.SqlStmt = "Delete from MemTrans " & _
                " where AccID = " & m_AccID & _
                " and TransID = " & TransID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

gDbTrans.SqlStmt = "Delete from MemIntTrans " & _
            " Where AccID = " & m_AccID & _
            " And TransID = " & TransID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

Set bankClass = New clsBankAcc

'Build SQL
If transType = wWithdraw Then
    gDbTrans.SqlStmt = "Update ShareTrans set ReturnTransID = null " & _
                " where ReturnTransID = " & TransID & ";"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not bankClass.UndoCashWithdrawls(headID, Amount, TransDate) Then GoTo Exit_Line
    If Not bankClass.UndoCashWithdrawls(FeeHeadId, ShareFee, TransDate) Then GoTo Exit_Line
    
ElseIf transType = wDeposit Then
    gDbTrans.SqlStmt = "Delete from ShareTrans where Accid = " & m_AccID & _
                       " And SaleTransID = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

    If Not bankClass.UndoCashDeposits(headID, Amount, TransDate) Then GoTo Exit_Line
    If Not bankClass.UndoCashDeposits(FeeHeadId, ShareFee, TransDate) Then GoTo Exit_Line
    
ElseIf transType = wContraWithdraw Or transType = wContraDeposit Then
    MsgBox "Under Developemnet "
    GoTo Exit_Line
Else
    'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(645), vbInformation, gAppName & " - Error"
    GoTo Exit_Line
End If


gDbTrans.CommitTrans
InTrans = False

AccountUndoLastTransaction = True
Exit Function
    
Exit_Line:
    
    If InTrans Then gDbTrans.RollBack
    'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"

End Function

'  Arrange The Property Sheet
Private Sub ArrangePropSheet()

Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean

' Arrange the Slider panel.
With picSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCount()
    .Height = txtData(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    ' If the height is greater than viewport height,
    ' the scrollbar needs to be displayed.  So,
    ' reduce the width accordingly.
    If .Height > picViewport.ScaleHeight Then
        NeedsScrollbar = True
        .Width = picViewport.ScaleWidth - _
                VScroll1.Width
    Else
        .Width = picViewport.ScaleWidth
    End If

End With

' Set/Reset the properties of scrollbar.
With VScroll1
    .Height = picViewport.ScaleHeight
    .Min = 0
    .Max = picSlider.Height - picViewport.ScaleHeight
    If .Max < 0 Then .Max = 0
    .SmallChange = txtData(0).Height
    .LargeChange = picViewport.ScaleHeight / 2
End With

' Adjust the text controls on this panel.
Dim I As Integer
For I = 0 To txtData.count - 1
    txtData(I).Width = picSlider.ScaleWidth _
            - txtPrompt(I).Width - CTL_MARGIN
Next


If NeedsScrollbar Then
    VScroll1.Visible = True
End If

' Need to adjust the width of text boxes, due to
' change in width of the slider box.
Dim CtlIndex As Integer
For I = 0 To txtData.count - 1
    txtData(I).Width = picSlider.ScaleWidth - _
        (txtPrompt(I).Left + txtPrompt(I).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For I = 0 To cmb.count - 1
    cmb(I).Width = txtData(I).Width
Next
For I = 0 To cmd.count - 1
    cmd(I).Left = txtData(I).Left + txtData(I).Width - cmd(I).Width
Next

End Sub
' This Function Deletes The Account Permanently
Private Function AccountDelete() As Boolean
Dim AccId As Long
Dim rst As Recordset

    If m_AccID <= 0 Then Exit Function
    AccId = m_AccID
    
'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from MemMaster where AccID = " & AccId
    If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    If FormatField(rst("ClosedDate")) <> "" Then
        'MsgBox "This account has already been closed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(524), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

'Check if transactions are there
    gDbTrans.SqlStmt = "Select TOP 1 * from MemTrans where " & _
        "AccID = " & AccId & " and TransID > 1"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        'MsgBox "You cannot delete an account having transactions !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Delete account from DB
    gDbTrans.BeginTrans
    
    gDbTrans.SqlStmt = "Delete from MemIntPayable where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    'Delete from IntTrans Table (Delete the Member fee)
    gDbTrans.SqlStmt = "Delete from MemIntTrans where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    'Delete from Trans Table (Delete the Member fee)
    gDbTrans.SqlStmt = "Delete from MemTrans where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    'Delete from the master
    gDbTrans.SqlStmt = "Delete from MemMaster where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
        
    gDbTrans.CommitTrans
AccountDelete = True
End Function


' Returns the index of the control bound to "strDatasrc".
Private Function GetIndex(strDataSrc As String) As Integer
GetIndex = -1
Dim strTmp As String
Dim I As Integer
For I = 0 To txtPrompt.count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(txtPrompt(I).Tag, "DataSource")
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
    Dim NewAccNo As Long
    Dim rst As ADODB.Recordset
    'gDBTrans.SQLStmt = "Select TOP 1 AccID from MemMaster order by AccID desc"
    gDbTrans.SqlStmt = "SELECT MAX(val(AccNum)) FROM MemMaster"
    If m_MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where MemberType = " & m_MemberType
    
    If gDbTrans.Fetch(rst, adOpenDynamic) = 0 Then
        NewAccNo = 1
    Else
        NewAccNo = FormatField(rst(0))
        If IsNumeric(NewAccNo) Then NewAccNo = Val(NewAccNo) + 1
    End If
    GetNewAccountNumber = NewAccNo
End Function
' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetVal(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
For I = 0 To txtData.count - 1
    strTxt = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        GetVal = txtData(I).Text
        Exit For
    End If
Next
End Function
Private Function PassBookPageInitialize()
    Dim Offset As Single
    Offset = 100
    With grdTrans
        .Clear: .Cols = 4:  .Rows = 11: .FixedRows = 1: .FixedCols = 0
        .Row = 0
        .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1150 '.Width / .Cols - Offset '"Date"
        .Col = 1: .Text = GetResourceString(334): .ColWidth(1) = .Width / .Cols - Offset '"Sold"
        .Col = 2: .Text = GetResourceString(335): .ColWidth(2) = .Width / .Cols - Offset '"Returned"
        .Col = 3: .Text = GetResourceString(42): .ColWidth(3) = .Width / .Cols - Offset ''"Balance"
    End With
    
    'Exit Function
    Dim ColCount As Integer
    With grdTrans
        .Clear: .Cols = 6:  .Rows = 11: .FixedRows = 1: .FixedCols = 0
        ColCount = .Cols
        .Row = 0
        .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1150 '.Width / .Cols - Offset '"Date"
        .Col = 1: .Text = GetResourceString(53, 212): .ColWidth(1) = .Width / ColCount - Offset '"Share(Face) value"
        .Col = 2: .Text = GetResourceString(334): .ColWidth(2) = .Width / .Cols - Offset '"Sold"
        .Col = 3: .Text = GetResourceString(53, 40): .ColWidth(3) = .Width / ColCount - Offset '"Share Amount"
        .Col = 4: .Text = GetResourceString(335): .ColWidth(4) = .Width / .Cols - Offset '"Quantity"
        .Col = 5: .Text = GetResourceString(53, 42): .ColWidth(5) = .Width / ColCount - Offset ''"Balance"
    End With
    
End Function

Private Function LoadPropSheet() As Boolean

TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True
lblDesc.BorderStyle = 0
lblHeading.BorderStyle = 0
lblOperation.Caption = GetResourceString(54) '"Operation Mode : <INSERT>"

'
' Read the data from SBSetup.ini and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\MMACC_" & gLangOffSet & ".PRP"
If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet Then
        PropFile = App.Path & "\MMACCKan.PRP"
    Else
        PropFile = App.Path & "\MMACC.PRP"
    End If
End If
If Dir(PropFile, vbNormal) = "" Then
    MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox GetResourceString(816) & PropFile & "' !", vbExclamation
    Exit Function
End If

'Load the CLIP Icon
    imgNewAcc.Picture = LoadResPicture(105, vbResIcon)
 
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
    strTag = ReadFromIniFile("Property Sheet", "Prop" & I + 1, PropFile)
    If strTag = "" Then Exit Do

    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load txtPrompt(txtPrompt.count)
        Load txtData(txtData.count)
    End If
    CtlIndex = txtPrompt.count - 1

    ' Get the property type.
    strPropType = ExtractToken(strTag, "PropType")
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With txtPrompt(CtlIndex)
                .FontBold = True
                .Caption = ""
            End With
            txtData(CtlIndex).Enabled = False

        Case "EDITABLE"
            ' Add 4 spaces for indentation purposes.
            With txtPrompt(CtlIndex)
                .Caption = IIf(gLangOffSet, Space(2), Space(4))
                .FontBold = False
                .Enabled = True
            End With
            txtData(CtlIndex).Enabled = True
        Case Else
            'MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            Exit Function

    End Select

    ' Set the PROPERTIES for controls.
    With txtPrompt(CtlIndex)
        strRet = PutToken(strTag, "Visible", "True")
        .Tag = strRet
        .Caption = .Caption & ExtractToken(.Tag, "Prompt")
        If CtlIndex = 0 Then
            .Top = 0
        Else
            .Top = txtPrompt(CtlIndex - 1).Top _
                + txtPrompt(CtlIndex - 1).Height + CTL_MARGIN
        End If
        .Left = 0
        .Visible = True
    End With
    With txtData(CtlIndex)
        .Top = txtPrompt(CtlIndex).Top
        .Left = txtPrompt(CtlIndex).Left + _
            txtPrompt(CtlIndex).Width + CTL_MARGIN
        .Visible = True
        ' Check the LockEdit property.
        strRet = ExtractToken(strTag, "LockEdit")
        If StrComp(strRet, "True", vbTextCompare) = 0 Then
            .Locked = True
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
                .Left = txtData(I).Left
                .Top = txtData(I).Top
                .Width = txtData(I).Width
                .ZOrder 0
                ' Set it's tab order.
                .TabIndex = txtData(I).TabIndex + 1
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
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
                'If it is a member type then load the combo box from Resource string
                'refer end og this function
            End With

        Case "BROWSE"
            'Load a command button.
            If Not CmdLoaded Then
                CmdLoaded = True
            Else
                Load cmd(cmd.count)
            End If
            With cmd(cmd.count - 1)
                '.Index = i
                .Width = txtData(I).Height
                .Height = txtData(I).Height
                .Left = txtData(I).Left + txtData(I).Width - .Width
                .Top = txtData(I).Top
                .TabIndex = txtData(I).TabIndex + 1
                .ZOrder 0
                '.Visible = True
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                        "TextIndex", CStr(cmd.count - 1))
                If I = 1 Then
                    .Caption = GetResourceString(294) '"Reset"
                    .Width = 1000
                ElseIf I = 2 Then
                    .Caption = GetResourceString(295) '"Details..."
                    .Width = 1000
                Else
                    .Caption = "..."
                    .Width = 350
                End If
            End With

    End Select

    ' Increment the loop count.
    I = I + 1
Loop

ArrangePropSheet

' Get a new account number and display it to accno textbox.
Dim txtIndex As Integer
txtIndex = GetIndex("AccNum")
txtData(txtIndex).Text = GetNewAccountNumber

' Show the current date wherever necessary.
txtIndex = GetIndex("CreateDate")
txtData(txtIndex).Text = gStrDate

' Set the default updation mode.
m_accUpdatemode = Insert

End Function
Private Sub ResetUserInterface()

'If m_Accid = 0 Then Exit Sub
If m_AccID = 0 And m_CustReg.CustomerID = 0 Then Exit Sub
Set m_rstPassBook = Nothing
Set m_rstShareBook = Nothing

m_accUpdatemode = Insert

'First the TAB 1
    'Disable the UI if you are unable to load the specified account number
    lblBalance.Caption = ""
    With txtCustName
        .BackColor = wisGray: .Enabled = False: .Caption = ""
    End With
    With txtDate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With cmdDate
        .Enabled = False
    End With
    With txtShareFees
        .BackColor = wisGray: .Enabled = False
    End With
    
    With txtAmount
        .BackColor = wisGray: .Enabled = False: .Value = 0
    End With
    
    With txtTotal
        .BackColor = wisGray: .Enabled = False: .Caption = "0.00"
    End With
    
    With cmbTrans
        .BackColor = wisGray: .Enabled = False
    End With

    With cmbLeaves
        .BackColor = wisGray: .Enabled = False: .Clear
    End With

    With cmdLeaves
        .Enabled = False
    End With

    With Me.rtfNote
        .BackColor = wisGray: .Enabled = False:
        .Text = GetResourceString(259) '"< No notes defined >"
        If gLangOffSet <> wis_NoLangOffset Then
            .Font.name = gFontName: .Font.Size = gFontSize
        Else
            .Font.Size = 10: .Font = "Arial"
        End If
    End With

    With cmdAccept
        .Enabled = False
    End With
    With cmdUndo
        .Enabled = False
    End With
    
    Call PassBookPageInitialize
    
    cmdAddNote.Enabled = False
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    
    
'Now the Tab 2
    Dim I As Integer
    Dim strField As String
    Dim txtIndex As Integer
    
    'Show the Share TransDetail Frame
    'TabStrip2.Tabs.Item(1).Index = (0)
    TabStrip2.Tabs.Item(1).Selected = True
    'Enable the reset (auto acc no generator button)
    cmd(0).Enabled = True
    
    For I = 0 To txtData.count - 1
        txtData(I).Text = ""
        ' If its Createdate field, then put today's left.
        strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
        If StrComp(strField, "CreateDate", vbTextCompare) = 0 Then
            txtData(I).Text = gStrDate
        End If
        ' If it is Membership Fee Then Readvalue Form Setup & Write
        If StrComp(strField, "MemberFee", vbTextCompare) = 0 Then
            Dim SetUp As New clsSetup
            txtData(I).Text = SetUp.ReadSetupValue("MMAcc", "MemberShipFee", "10.00")
        End If
        
    Next
    lblOperation.Caption = GetResourceString(54) '"Operation Mode : <INSERT>"
    txtIndex = GetIndex("AccNum")
    txtData(txtIndex).Text = GetNewAccountNumber
    txtData(txtIndex).Locked = False
    cmdDelete.Enabled = False
    
'The form level variables
    m_accUpdatemode = Insert
    m_CustReg.NewCustomer
    m_AccID = 0
    Set m_rstPassBook = Nothing
RaiseEvent AccountChanged(0)

End Sub

Private Sub ScrollWindow(Ctl As Control)

If picSlider.Top + Ctl.Top + Ctl.Height > picViewport.ScaleHeight Then
    ' The control is below the viewport.
    Do While picSlider.Top + Ctl.Top + Ctl.Height > picViewport.ScaleHeight
        ' scroll down by one row.
        With VScroll1
            If .Value + .SmallChange <= .Max Then
                .Value = .Value + .SmallChange
            Else
                .Value = .Max
            End If
        End With
    Loop

ElseIf picSlider.Top + Ctl.Top < 0 Then
    ' The control is above the viewport.
    ' Keep scrolling until it is in viewport.
    Do While picSlider.Top + Ctl.Top < 0
        With VScroll1
            If .Value - .SmallChange >= .Min Then
                .Value = .Value - .SmallChange
            Else
                .Value = .Min
            End If
        End With
    Loop
End If

End Sub
Private Sub SetDescription(Ctl As Control)

' Extract the description title.
lblHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblDesc.Caption = ExtractToken(Ctl.Tag, "Description")
End Sub
Private Sub PassBookPageShow()
Dim I As Integer
Dim transType As wisTransactionTypes

'Check if Rec Set has been set
If m_rstPassBook Is Nothing Then Exit Sub
    
'Show records till eof
With grdTrans
    Call PassBookPageInitialize
    '.Visible = False
    I = 0
    m_rstPassBook.MoveFirst
    Do
         If m_rstPassBook.EOF Then Exit Do
        'TransType = m_rstPassBook("Transtype")
        I = I + 1
        If .Rows < I + 1 Then .Rows = I + 1
        .Row = I
        .Col = 0: .Text = FormatField(m_rstPassBook("TransDate"))
        '.Col = IIf(TransType = wDeposit Or TransType = wContraDeposit, 1, 2)
        '.Text = FormatField(m_rstPassBook("ShareCount"))
        .Col = 3: .Text = FormatField(m_rstPassBook("Balance"))
        ''New PassBok Vaidynathan
        .Col = 1: .Text = FormatField(m_rstPassBook("FaceValue"))
        .Col = IIf(m_rstPassBook("Trans") = "Sales", 2, 4)
            .Text = FormatField(m_rstPassBook("ShareCount"))
        .Col = 3: .Text = FormatCurrency(FormatField(m_rstPassBook("FaceValue")) * FormatField(m_rstPassBook("ShareCount")))
        .Col = 5: .Text = FormatField(m_rstPassBook("Balance"))
nextRecord:
        m_rstPassBook.MoveNext
       
    Loop
    '.Visible = True
    .Row = 1
End With
m_rstPassBook.MoveFirst

End Sub



Private Sub ShareLeavesInitialize()

'Initialize the grid to show share details
With grdShare
    .Clear
    .Rows = 10: .Cols = 5
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(31) '"Date"
    .Col = 1: .Text = GetResourceString(53, 337, 60) '"Share Certificate"
    .Col = 2: .Text = GetResourceString(53, 140) ''"Share Value"
    .Col = 3: .Text = GetResourceString(334, 37) '"Sold date"
    .Col = 4: .Text = GetResourceString(335) '"Returned On"
    .ColWidth(0) = 600: .ColWidth(1) = 1400
    .ColWidth(2) = 1300: .ColWidth(3) = 1100
    .ColWidth(4) = 1100 ': .ColWidth(0) = 500
End With

End Sub

' Returns the number of items that are visible for a control array.
' Looks in the control's tag for visible property, rather than
' depend upon the control's visible property for some obvious reasons.
Private Function VisibleCount() As Integer
On Error GoTo Err_line
Dim I As Integer
Dim strVisible As String
For I = 0 To txtPrompt.count - 1
    strVisible = ExtractToken(txtPrompt(I).Tag, "Visible")
    If StrComp(strVisible, "True") = 0 Then
        VisibleCount = VisibleCount + 1
    End If
Next
Err_line:
End Function


'
Private Sub ShareLeavesShow()

Call ShareLeavesInitialize
If m_rstShareBook Is Nothing Then Exit Sub

Dim rst As Recordset

cmdSharePrev.Enabled = IIf(m_rstShareBook.AbsolutePosition > 1, True, False)

Do While Not m_rstShareBook.EOF
    With grdShare
        '.Visible = True
        If .Row = 11 Then Exit Do
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = .Row
        .Col = 1: .Text = FormatField(m_rstShareBook("CertNO"))
        If FormatField(m_rstShareBook("ReturnTransID")) <> 0 Then 'Share has been sold
            'Get Sold date
            gDbTrans.SqlStmt = "Select TransDate from MemTrans where TransID = " & _
                        m_rstShareBook("ReturnTransID") & " And AccId = " & m_rstShareBook("AccID")
            If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
                .Col = 4: .Text = FormatField(rst("TransDate"))
            End If
        End If
        'Get Face Value
        gDbTrans.SqlStmt = "Select TransDate from MemTrans where TransID = " & _
                            m_rstShareBook("SaleTransID") & " And AccId = " & m_rstShareBook("AccID")
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
            .Col = 3: .Text = FormatField(rst("TransDate"))
        .Col = 2: .Text = FormatField(m_rstShareBook("FaceValue"))
    End With
    m_rstShareBook.MoveNext
Loop

cmdShareNext.Enabled = IIf(m_rstShareBook.EOF, False, True)
    
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" Then
    txtData(Val(txtIndex)).Text = cmb(Index).Text
End If

End Sub



Private Sub cmbTrans_Click()

If cmbTrans.ListCount = 0 Then
    'MsgBox "Initialization Error"
    MsgBox GetResourceString(608)
    Exit Sub
End If

If cmbTrans.ListIndex = 0 Then  'A case of share purchase
Dim SetUp As New clsSetup
    txtShareFees = Val(SetUp.ReadSetupValue("MMAcc", "ShareFee", "10.00"))
    Exit Sub
End If

If cmbTrans.ListIndex = 1 Then  'A case of share sales
    txtShareFees = 0
    txtAmount.Locked = False
    'txtCheque.Visible = False
    'cmbCheque.Visible = True
    'cmdCheque.Visible = True
    'Exit Sub
Else
    txtAmount.Locked = True
End If



End Sub


Private Sub cmd_Click(Index As Integer)
Dim txtIndex As String
Dim rst As Recordset

' Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")

Select Case UCase$(strField)
    Case "ACCNUM"
        If m_accUpdatemode = Insert Then
            txtData(txtIndex).Text = GetNewAccountNumber
        End If

    Case "ACCNAME"
        m_CustReg.ModuleId = wis_Members
        m_CustReg.ShowDialog
        txtData(txtIndex).Text = m_CustReg.FullName

    Case "CREATEDATE"
        With Calendar
            .Left = txtData(txtIndex).Left + Me.Left _
                    + Me.picViewport.Left + fraNew.Left + 50
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fraNew.Top + 300
            .Width = txtData(txtIndex).Width
            If .Top + .Height > Screen.Height Then .Top = .Top - .Height - txtData(txtIndex).Height
            .Height = .Width
            .selDate = txtData(txtIndex).Text
            .Show vbModal, Me
            If .selDate <> "" Then txtData(txtIndex).Text = .selDate
        End With
    
    Case "INTRODUCERID"
        ' Build a query for getting introducer details.
        ' If an account number specified, exclude it from the list.
        gDbTrans.SqlStmt = "SELECT MemMaster.AccID as [Acc No], " _
                    & "Title + FirstName + Space(1) + Middlename " _
                    & "+ space(1) + LastName as Name, HomeAddress, " _
                    & "OfficeAddress FROM MemMaster, NameTab WHERE " _
                    & "MemMaster.CustomerID = NameTab.CustomerID"
        Dim intIndex As Integer
        intIndex = GetIndex("AccNum")
        If txtData(intIndex).Text <> "" And _
                IsNumeric(txtData(intIndex).Text) Then
            gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND " _
                & "AccID <> " & txtData(intIndex).Text
        End If
        Dim Lret As Long
        Lret = gDbTrans.Fetch(rst, adOpenDynamic)
        If Lret <= 0 Then
            'MsgBox "No accounts present!", vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(525), vbExclamation, wis_MESSAGE_TITLE
            Exit Sub
        End If
        'Fill the details to report dialog and display it.
        If m_frmLookUp Is Nothing Then
            Set m_frmLookUp = New frmLookUp
        End If
        If Not FillView(m_frmLookUp.lvwReport, rst) Then
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
            MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
            Exit Sub
        End If
        With m_frmLookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
            .lvwReport.ColumnHeaders(2).Width = 3750
            .lvwReport.ColumnHeaders(3).Width = 3750
            .Title = "Select Introducer..."
            .Show vbModal, Me
          
             txtData(txtIndex).Text = .lvwReport.SelectedItem.Text
             txtData(txtIndex + 1).Text = .lvwReport.SelectedItem.SubItems(1)
        End With
End Select
End Sub
Private Sub cmdAccept_Click()

'Check and perform appropriate transaction
If Not AccountTransaction() Then Exit Sub
    
'Reload the account
    If Not AccountLoad(m_AccID) Then Exit Sub

'Show the Pass bok details
    TabStrip2.Tabs(2).Selected = True
End Sub

Private Sub cmdAddNote_Click()
If m_Notes.ModuleId = 0 Then
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

Private Sub cmdApply_Click()
'Cheque values
If Not CurrencyValidate(Me.txtPropShareFee.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropShareFee
    Exit Sub
End If
If Not CurrencyValidate(Me.txtPropShareValue.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropShareValue
    Exit Sub
End If

If Not CurrencyValidate(Me.txtPropMemCancel.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropMemCancel
    Exit Sub
End If
If Not CurrencyValidate(Me.txtPropMemFee.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropMemFee
    Exit Sub
End If

If Val(Me.txtPropDivedend.Text) < 0 Then
    'MsgBox "Invalid interest value specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(518), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropDivedend
    Exit Sub
End If
If Trim$(txtPropLastDivedendOn) <> "" Then
    If Not DateValidate(Me.txtPropLastDivedendOn, "/", True) Then
        'MsgBox "Invalid date specified", , wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), , wis_MESSAGE_TITLE
        'ActivateTextBox Me.txtPropLastDivedendOn
        Exit Sub
    End If
End If
If Trim$(txtPropNextDivedendOn) <> "" Then
    If Not DateValidate(txtPropNextDivedendOn, "/", True) Then
        'MsgBox "Invalid date specified", , wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), , wis_MESSAGE_TITLE
        'ActivateTextBox Me.txtPropNextDivedendOn
        Exit Sub
    End If
End If

Dim SetUp As New clsSetup

Call SetUp.WriteSetupValue("MMAcc", "MemberShipFee", Me.txtPropMemFee.Text)
Call SetUp.WriteSetupValue("MMAcc", "Cancellation", txtPropMemCancel.Text)
Call SetUp.WriteSetupValue("MMAcc", "ShareFee", txtPropShareFee.Text)
Call SetUp.WriteSetupValue("MMAcc", "ShareValue", txtPropShareValue.Text)
Call SetUp.WriteSetupValue("MMAcc", "RateOfDivedend", txtPropDivedend.Text)
Call SetUp.WriteSetupValue("MMAcc", "LastDivedendOn", txtPropLastDivedendOn)
Call SetUp.WriteSetupValue("MMAcc", "NextDivedendOn", txtPropNextDivedendOn)
cmdApply.Enabled = False
End Sub

Private Sub cmdDate_Click()
 With Calendar
    .Left = Me.Top + Me.fraTransact.Left + cmdDate.Left - .Left / 2
    .Top = Me.Top + Me.fraTransact.Top + cmdDate.Top
    If DateValidate(txtDate.Text, "/", True) Then
        .selDate = txtDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    Me.txtDate.Text = .selDate
 End With
End Sub

Private Sub cmdDate1_Click()
With Calendar
    .Top = Me.Top + Me.fraReports.Top + fraMemType.Top + cmdDate1.Top
    .Left = Me.Left + Me.fraReports.Left + fraMemType.Left + cmdDate1.Left - .Width / 2
    If Not DateValidate(txtDate1.Text, "/", True) Then
        .selDate = gStrDate
    Else
        .selDate = txtDate1.Text
    End If
    .Show vbModal
    txtDate1.Text = .selDate
End With

End Sub

Private Sub cmdDate2_Click()
With Calendar
    .Top = Me.Top + Me.fraReports.Top + fraMemType.Top + cmdDate2.Top
    .Left = Me.Left + Me.fraReports.Left + fraMemType.Left + cmdDate2.Left - .Width / 2
    If Not DateValidate(txtDate2.Text, "/", True) Then
        .selDate = gStrDate
    Else
        .selDate = txtDate2.Text
    End If
    .Show vbModal
    txtDate2.Text = .selDate
End With

End Sub


Private Sub cmdDelete_Click()
'If MsgBox("Are you sure you want to delete this member ?", vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
If MsgBox(GetResourceString(539), vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
    Exit Sub
End If

If AccountDelete() Then
    Call ResetUserInterface
End If
    
End Sub

Private Sub cmdDividend_Click()
Dim Mon As Integer
Dim Yr As Integer
If Not DateValidate(txtFromDate.Text, "/", True) Then Exit Sub
If Not DateValidate(txtToDate.Text, "/", True) Then Exit Sub

If MsgBox("This will put member DIVIDEND FROM " & txtFromDate.Text & _
    " TO " & txtToDate.Text & " On " & gStrDate, vbYesNo, "MEMBER DIVIDEND") = vbNo Then Exit Sub
Me.Refresh

If Not ComputeShareDividend(GetSysFormatDate(txtFromDate), GetSysFormatDate(txtToDate), CDate(gStrDate), CSng(txtPropDivedend.Text)) Then
    MsgBox "Error in calulating Share Dividend", vbInformation, "Share Dividend"
End If

End Sub

Private Function ComputeShareDividend(ByVal fromDate As Date, _
            ByVal toDate As Date, ByVal TransDate As Date, Rate As Single) As Boolean
        
Dim rst_Main As Recordset
Dim rstTemp As Recordset

'Dim ArrAccID() As Long
Dim AccId As Long
Dim MaxAccID As Long
Dim count As Integer
Dim DateLimit As String
Dim YearLimit As Integer
Dim Dividend As Currency
Dim TotalDividend As Currency
Dim Balance As Currency
Dim PayableAccID As Long
Dim DividendAccID As Long
Dim TransID As Long
Dim ConsiderAllMonths As Boolean
Dim DivFact As Integer

Dim name As String
Dim nameEnglish As String

prg.Visible = True
lblStatus.Visible = True
Me.MousePointer = vbHourglass
lblStatus = "Collecting the account Information"
prg.Value = 0
gDbTrans.SqlStmt = "SELECT AccID,AccNum FROM MemMaster " & _
    " Where AccID in (Select Distinct AccID From MemTrans) ORDER BY val(AccNum)"

If m_MemberType > 0 Then
    gDbTrans.SqlStmt = "SELECT AccID,AccNum,MemberTypeName FROM MemMaster A Inner Join MemberTypeTab B" & _
        " On A.MemberType = b.MemberType " & _
        " Where A.AccID in (Select Distinct AccID From MemTrans) ORDER BY MemberType, val(AccNum)"
End If

If gDbTrans.Fetch(rst_Main, adOpenStatic) <= 0 Then Exit Function
MaxAccID = rst_Main.RecordCount
                                                
lblStatus = "Collecting the account Information"
prg.Max = MaxAccID + 1
prg.Min = 0

'Validate date
If fromDate > toDate Then Exit Function
    
TransID = 100
rst_Main.MoveFirst

'Load All the Details to a form
Set m_frmdividend = New frmIntPayble

Load m_frmdividend

With m_frmdividend
    Call .LoadContorls(CInt(MaxAccID) + 1, 20)
    .lblTitle = GetResourceString(420, 200)
    .Title(0) = GetResourceString(36, 60)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(200)
    .Title(3) = GetResourceString(42, 200)
    .PutTotal = False
    .TotalColoumn = False
    .BalanceColoumn = True
End With
Dim tempdate As Date

'Set the date to for the calculation
'first set the todate to the end of given month
toDate = GetSysLastDate(toDate)

'Now set the Fromdate
If MsgBox("Do you want to consider the share balance of each month ?" _
    , vbQuestion + vbDefaultButton2 + vbYesNo, wis_MESSAGE_TITLE) = vbYes Then _
        ConsiderAllMonths = True

If ConsiderAllMonths Then fromDate = Month(fromDate) & "/15/" & Year(fromDate)
If ConsiderAllMonths Then DivFact = 12 Else DivFact = 1

For count = 1 To MaxAccID
    AccId = rst_Main("AccID")
    lblStatus = "Calculating Divedend for Memeber Number " & AccId 'ArrAccID(Count)
    prg.Value = count
    Me.Refresh
    tempdate = IIf(ConsiderAllMonths, fromDate, toDate)
    Dividend = 0
    
    'While (Mon > 3 And Yr < YearLimit) Or (Mon <= 3 And Yr = YearLimit)
    While tempdate <= toDate
        'DateLimit = Mon & "/15/" & Yr
        gDbTrans.SqlStmt = "Select Max(Balance) as MaxBalance FROM MemTrans as A " & _
            " where TransID = (Select MAX(TransID) FROM MemTrans B WHERE " & _
                    " A.AccID = B.AccID and TransDate <= #" & tempdate & "# ) " & _
             "AND AccID = " & AccId 'ArrAccID(Count)
        If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
            Balance = FormatField(rstTemp(0))
            Dividend = Dividend + (Balance * 1 * Rate) / (100 * DivFact)
        End If
        tempdate = DateAdd("m", 1, tempdate)
    Wend
    
    gDbTrans.SqlStmt = "Select Title + Space(1) + FirstName + " & _
            "Space(1) + MiddleName + Space(1) + LastName as Name " & _
            " from NameTab, MemMaster  where MemMaster.AccId  = " & rst_Main("AccID") & _
            " And  NameTab.CustomerID =  MemMaster.CustomerId"
    name = ""
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then name = FormatField(rstTemp(0))
    
    Dividend = Dividend \ 1
    With m_frmdividend
        .KeyData(count) = rst_Main("AccID")
        .AccNum(count) = rst_Main("AccNum")
        .CustName(count) = name
        .Amount(count) = Dividend
        .Balance(count) = Balance
        TotalDividend = TotalDividend + Dividend
    End With
    
    rst_Main.MoveNext
Next

With m_frmdividend
    .CustName(count) = GetResourceString(52, 200)
    .Amount(count) = TotalDividend
    .ShowForm
End With
Me.Refresh

If m_frmdividend.grd.Rows = 1 Then GoTo Exit_Line

Me.Refresh

gDbTrans.BeginTrans


rst_Main.MoveFirst
TotalDividend = 0
gDbTrans.SqlStmt = "select Max(TransID) From MemIntPayable"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then TransID = FormatField(rstTemp(0))
TransID = TransID + 1

Dim VoucherNo As String
Dim UserID As Integer
UserID = gCurrUser.UserID

rst_Main.MoveFirst
prg.Max = prg.Max + 4

For count = 1 To MaxAccID

    With m_frmdividend
        Dividend = Val(.Amount(count))
        TotalDividend = TotalDividend + Dividend
    End With

    Balance = Balance + Dividend
    If Dividend > 0 Then
        gDbTrans.SqlStmt = "Insert into MemIntPayable(AccID, TransID, " & _
                "TransDate, Amount,Balance, " & _
                " TransType,UserID, VoucherNo )" & _
                " values ( " & _
                rst_Main("AccId") & "," & _
                TransID & "," & _
                "#" & TransDate & "#," & _
                Dividend & "," & _
                Balance & "," & _
                wContraDeposit & "," & _
                UserID & "," & _
                AddQuotes(VoucherNo, True) & ")"
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
            
        Dividend = 0
    End If
    rst_Main.MoveNext
    prg.Value = count
    lblStatus = "Inserting divident to Member No " & count
Next
Me.Refresh

Dim bankClass As clsBankAcc

Set bankClass = New clsBankAcc
name = m_MemberTypeName & " " & GetResourceString(49, 200) 'Member Dividend
nameEnglish = m_MemberTypeNameEnglish & " " & LoadResourceStringS(49, 200) 'Member Dividend
PayableAccID = bankClass.GetHeadIDCreated(name, nameEnglish, parPayAble, 0, M_ModuleID)

name = name & " " & GetResourceString(267) 'Divedend Paid
nameEnglish = nameEnglish & " " & LoadResString(267) 'Diiveden Paid
DividendAccID = bankClass.GetHeadIDCreated(name, nameEnglish, parBankExpense, 0, M_ModuleID)

If Not bankClass.UpdateContraTrans(DividendAccID, PayableAccID, _
        TotalDividend, TransDate) Then GoTo Exit_Line

Set bankClass = Nothing

gDbTrans.CommitTrans

Me.MousePointer = vbDefault
MsgBox "Dividend Calculation Is Over", vbInformation, "Dividend"
ComputeShareDividend = True

prg.Visible = False
lblStatus.Visible = False

gDbTrans.BeginTrans

Exit_Line:

gDbTrans.RollBack

Me.MousePointer = vbDefault
If Not m_frmdividend Is Nothing Then
    On Error Resume Next
    Unload m_frmdividend
    Err.Clear
    On Error GoTo 0
    Set m_frmdividend = Nothing
End If

End Function

Private Function UndoDividendPayable(OnIndianDate As String) As Boolean

lblFailAccIDs = ""

Dim OnDate As Date
Dim rst As Recordset

OnDate = GetSysFormatDate(OnIndianDate)

'Before undoing check whether he has already added the interestpayble amount or not
gDbTrans.SqlStmt = "Select *  from MemIntPayable Where " & _
    " TransDate = #" & GetSysFormatDate(OnIndianDate) & "# " & _
    " And Transtype = " & wContraDeposit

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    UndoDividendPayable = True
    Exit Function
End If
  
Me.MousePointer = vbHourglass
  On Error GoTo ErrLine
  'declare the variables necessary
  
Dim Amount As Currency
gDbTrans.SqlStmt = "Select Sum(A.Amount) From MemIntPayable A" & _
    " WHERE TransID = (SELECT TransID FROM MemIntPayable B WHERE " & _
       " B.AccID = A.AccID AND TransDate = #" & OnDate & "# " & _
       " AND TransType = " & wContraDeposit & ") " & _
       " AND TransID = (SELECT Max(TransID) FROM MemIntPayable B WHERE " & _
            " B.AccID = A.AccID )"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Amount = FormatField(rst(0))

If Amount Then
    If MsgBox("You are withdrawing the dividend payble Rs." & Amount & _
        vbCrLf & "Do You want to continue?", vbYesNo, _
        wis_MESSAGE_TITLE) = vbNo Then Exit Function
Else
    Exit Function
End If

gDbTrans.SqlStmt = "DELETE A.* From MemIntPayable A," & _
        " WHERE A.TransID = (SELECT TransID FROM MemIntPayable C WHERE " & _
            " C.AccID = A.AccID AND TransDate = #" & OnDate & "# " & _
            " AND TransType = " & wContraDeposit & ") " & _
        " AND TransID = (SELECT Max(TransID) FROM MemIntPayable B WHERE " & _
                " B.AccID = A.AccID )"

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If


Dim bankClass As clsBankAcc
Dim PayableAccID As Long
Dim DividendAccID As Long
Dim headName As String
Dim headNameEnglish As String

Set bankClass = New clsBankAcc
headName = m_MemberTypeName & " " & GetResourceString(49, 200) 'Member Dividend
headNameEnglish = m_MemberTypeNameEnglish & " " & LoadResourceStringS(49, 200) 'Member Dividend
PayableAccID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parPayAble, 0, M_ModuleID)
headName = headName & " " & GetResourceString(267) 'Diiveden Paid
headNameEnglish = headNameEnglish & " " & LoadResString(267) 'Diiveden Paid
DividendAccID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parBankExpense, 0, M_ModuleID)

If Not bankClass.UndoContraTrans(DividendAccID, PayableAccID, _
        Amount, OnDate) Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

gDbTrans.CommitTrans


UndoDividendPayable = True

ErrLine:
    Set bankClass = Nothing
    
End Function
Private Sub cmdLeaves_Click()
Dim rst As ADODB.Recordset
'Check if trans type is spec
If cmbTrans.ListIndex = -1 Then
    'MsgBox "Please specify the type of transaction.", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(588), vbInformation, gAppName & " - Error"
    cmbTrans.SetFocus
    Exit Sub
End If

'Display share form
    Set m_frmMMShare = New frmMMShare
    'Load m_frmShareLeaves
    With m_frmMMShare
        .Caption = GetResourceString(53, 295) '"Share Details"
        If cmbTrans.ListIndex = 0 Then
            .txtStartNo = 1
            gDbTrans.SqlStmt = "Select max(val(CertNo)) from ShareTrans where AccId = " & m_AccID
            If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                .txtStartNo = Val(FormatField(rst(0))) + 1
            .fraPurchase.Caption = GetResourceString(314) '"Issue Of Shares"
            .fraPurchase.Visible = True
            .cmdSelect.Visible = False
            .cmdAdd.Visible = True: .cmdAdd.Default = True
            .cmdInvert.Visible = False
            .fraSale.Visible = False
        ElseIf cmbTrans.ListIndex = 1 Then
            .fraPurchase.Visible = False
            .fraSale.Caption = GetResourceString(315) '"Current Share Nos"
            .fraSale.Visible = True
            .cmdAdd.Visible = False
            .cmdSelect.Visible = True: .cmdSelect.Default = True
            .cmdInvert.Visible = True
            'Fill up the leaves that this guy has got
            gDbTrans.SqlStmt = "Select * from ShareTrans " & _
                " where AccID = " & m_AccID & " order by CertNo"
            If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
                Beep
                txtAmount.Locked = False
                Exit Sub
            End If
            'FIll up the list box in the form
            Dim I As Long
            .lstCheque.Clear
            For I = 1 To rst.RecordCount
                If FormatField(rst("ReturnTransID")) = 0 Then _
                .lstCheque.AddItem FormatField(rst("CertNo"))
                rst.MoveNext
            Next I
        End If
    End With
    m_frmMMShare.Show vbModal

End Sub

Private Sub cmdLoad_Click()

Dim rst As Recordset
gDbTrans.SqlStmt = "Select AccId from MemMaster where " & _
            " AccNum = " & AddQuotes(txtAccNo, True)
If m_MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt + " And MemberType = " & m_MemberType

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
    'MsgBox "Account does not exists"
    MsgBox GetResourceString(716), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
   
If Not AccountLoad(rst("AccID")) Then
    ActivateTextBox txtAccNo
    Exit Sub
End If


End Sub

Private Sub cmdNextTrans_Click()

If m_rstPassBook Is Nothing Then Exit Sub

Dim CurPos As Integer

'Position cursor to start of next page
    If m_rstPassBook.EOF Then m_rstPassBook.MoveLast
    CurPos = m_rstPassBook.AbsolutePosition
    CurPos = 10 - (CurPos Mod 10)
    If m_rstPassBook.AbsolutePosition + CurPos >= m_rstPassBook.RecordCount Then
        Beep
        Exit Sub
    Else
        m_rstPassBook.Move CurPos
    End If

Call PassBookPageShow

#If junk Then
If m_rstPassBook.AbsolutePosition < m_rstPassBook.RecordCount - 10 Then
    If m_rstPassBook.AbsolutePosition Mod 10 <> 0 Then
        m_rstPassBook.Move 10 - m_rstPassBook.AbsolutePosition Mod 10
        If m_rstPassBook.AbsolutePosition >= m_rstPassBook.RecordCount - 10 Then
            cmdNextTrans.Enabled = False
        End If
    End If
Else
    cmdNextTrans.Enabled = False
End If
Call ShowPassBookPage
If m_rstPassBook.AbsolutePosition >= m_rstPassBook.RecordCount Then
    cmdPrevTrans.Enabled = False
Else
    cmdPrevTrans.Enabled = True
End If
#End If

End Sub

Private Sub cmdOk_Click()
Dim Cancel As Boolean
'ask for the user

Unload Me
End Sub

Private Sub cmdPhoto_Click()
If Not m_CustReg Is Nothing Then
    frmPhoto.setAccNo (m_CustReg.CustomerID)
        If (m_CustReg.CustomerID > 0) Then frmPhoto.Show vbModal
End If
End Sub

Private Sub cmdPrevTrans_Click()

If m_rstPassBook Is Nothing Then Exit Sub

Dim CurPos As Integer

'Position cursor to previous page
    If m_rstPassBook.EOF Then m_rstPassBook.MoveLast
    CurPos = m_rstPassBook.AbsolutePosition
    CurPos = CurPos - CurPos Mod 10 - 10
    If CurPos < 0 Then
        Beep
        Exit Sub
    Else
        m_rstPassBook.MoveFirst
        m_rstPassBook.Move (CurPos)
    End If
    Call PassBookPageShow
    
#If junk Then
If m_rstPassBook.AbsolutePosition > 10 Then
    If m_rstPassBook.AbsolutePosition Mod 10 = 0 Then
        'm_rstpassbook.MovePrevious
        m_rstPassBook.Move -1 * (m_rstPassBook.AbsolutePosition Mod 10 + 20)
    Else
        m_rstPassBook.Move -1 * (m_rstPassBook.AbsolutePosition Mod 10 + 10)
    End If
    
    If m_rstPassBook.AbsolutePosition < 10 Then
        cmdPrevTrans.Enabled = False
    End If
End If
Call ShowPassBookPage
If m_rstPassBook.AbsolutePosition < 10 Then
    cmdNextTrans.Enabled = False
Else
    cmdNextTrans.Enabled = True
End If
#End If
End Sub

Private Sub cmdPrint_Click()
    If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    
    m_frmPrintTrans.Show vbModal

End Sub

Private Sub cmdReset_Click()

Call ResetUserInterface

End Sub
Private Sub cmdSave_Click()

'SaveAccount
    If Not AccountSave Then Exit Sub
    
'Reload the account details once saved
    Dim AccNo As Long
    txtAccNo = GetVal("AccNum")
    Call cmdLoad_Click
End Sub



Private Function AccountReopen(AccId As Long) As Boolean
Dim rst As Recordset

'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from MemMaster where AccID = " & AccId
    If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Update MemMaster set ClosedDate = NULL " & _
        " where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans
    
AccountReopen = True
End Function

Private Sub cmdShareNext_Click()

If m_rstShareBook Is Nothing Then Exit Sub
Call ShareLeavesShow

End Sub

Private Sub cmdSharePrev_Click()

If m_rstShareBook Is Nothing Then Exit Sub

Dim CurPos As Integer

'Position cursor to previous page
    If m_rstShareBook.EOF Then m_rstShareBook.MoveLast
    CurPos = m_rstShareBook.AbsolutePosition
    CurPos = CurPos - CurPos Mod 10 - 20
    If CurPos < 0 Then
        Beep
        Exit Sub
    Else
        m_rstShareBook.MoveFirst
        m_rstShareBook.Move (CurPos)
    End If
    Call ShareLeavesShow

End Sub

Private Sub cmdUndo_Click()

If Not AccountUndoLastTransaction() Then Exit Sub

If Not AccountLoad(m_AccID) Then
    'MsgBox "Unable to undo transaction !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(609), vbCritical, gAppName & " - Error"
    Exit Sub
End If
Me.TabStrip2.Tabs(2).Selected = True
End Sub


Private Sub cmdUndoInterests_Click()
'Call UndoDividendPayable
End Sub

Private Sub cmdView_Click()

'Make one round of date checks
If Not optReports(7) Then
    If txtDate1.Enabled And Not DateValidate(txtDate1.Text, "/", True) Then
        'MsgBox "Invalid date specified !"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate1
        Exit Sub
    End If
    
    If txtDate2.Enabled And Not DateValidate(txtDate2.Text, "/", True) Then
        'MsgBox "Invalid date specified !" , vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate2
        Exit Sub
    End If
End If


If optReports(3).Value Then
    If Not DateValidate(txtDate1.Text, "/", True) Then
        'MsgBox "You must specify a from date in DD/MM/YYYY format for General Ledger report !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate1
        Exit Sub
    End If
End If

Dim ReportType As wis_MemReports
Dim fromDate As String
Dim toDate As String
If optReports(9) Then If Not txtDate1.Enabled Then fromDate = FinIndianFromDate

If txtDate1.Enabled Then fromDate = txtDate1
If txtDate2.Enabled Then toDate = txtDate2


If optReports(0) Then ReportType = repMemBalance
If optReports(1) Then ReportType = repMembers
If optReports(2) Then ReportType = repMemDayBook
If optReports(3) Then ReportType = repMemSubCashBook
If optReports(4) Then ReportType = repMemOpen
If optReports(5) Then ReportType = repMemClose
If optReports(6) Then ReportType = repFeeCol
If optReports(7) Then ReportType = repMemShareCert
If optReports(8) Then ReportType = repMonthlyBalance
If optReports(9) Then ReportType = repMemLedger
If optReports(10) Then ReportType = repMemLoanMembers
If optReports(11) Then ReportType = repMemNonLoanMembers


If m_clsRepOption Is Nothing Then Set m_clsRepOption = New clsRepOption


If cmbRepMemType.ListIndex < 0 Then cmbRepMemType.ListIndex = 0

RaiseEvent ShowReport(ReportType, IIf(Me.optMemId, wisByAccountNo, wisByName), _
    cmbRepMemType.ItemData(cmbRepMemType.ListIndex), _
    fromDate, toDate, m_clsRepOption)
    
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
        With VScroll1
            If .Value - .SmallChange > .Min Then
                .Value = .Value - .SmallChange
            Else
                .Value = .Min
            End If
        End With
    Case vbKeyDown
        ' Scroll down.
        With VScroll1
            If .Value + .SmallChange < .Max Then
                .Value = .Value + .SmallChange
            Else
                .Value = .Max
            End If
        End With
   Case vbKeyTab
        Dim I As Byte
        If TabStrip.SelectedItem.Index = 1 And Me.ActiveControl.name = TabStrip2.name Then
            With TabStrip2
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
            Exit Sub
        End If
     
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

Public Function AccountExists(AccId As Long, Optional ClosedON As String) As Boolean
Dim ret As Integer
Dim rst As Recordset

'Query Database
    gDbTrans.SqlStmt = "Select * from MemMaster where " & _
                        " AccID = " & AccId
    ret = gDbTrans.Fetch(rst, adOpenDynamic)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbCritical, gAppName & " - Error"
        Exit Function
    End If
    
'Check the closed status
    If Not IsMissing(ClosedON) Then
        ClosedON = FormatField(rst("ClosedDate"))
    End If

AccountExists = True
End Function



Public Function AccountLoad(AccId As Long) As Boolean

Dim rstMaster As Recordset
'Dim rstTemp As Recordset
Dim ClosedDate As String
Dim ret As Integer
Dim I As Integer
Dim txtIndex As Byte

'Check if account number is valid
    If AccId <= 0 Then GoTo DisableUserInterface
    
'Check if account number exists
    If Not AccountExists(AccId) Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        GoTo DisableUserInterface
    End If

'Query data base 'Set record set to local rec set
    gDbTrans.SqlStmt = "Select * from MemMaster where AccID = " & AccId
    If gDbTrans.Fetch(rstMaster, adOpenDynamic) <= 0 Then GoTo DisableUserInterface
    
'Load the Name details
    If Not m_CustReg.LoadCustomerInfo(FormatField(rstMaster("CustomerID"))) Then
        'MsgBox "Unable to load customer information !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        GoTo DisableUserInterface
    End If
    m_CustReg.ModuleId = wis_Members
'Get the transaction details of this account holder
    
    gDbTrans.SqlStmt = " Select Count(*) as ShareCount, A.AccId,A.TransID,TransDate,FaceValue,Balance, 'Sales' as Trans " & _
        " from MemTrans A Left join ShareTrans B " & _
        " On  A.AccID=B.AccID and A.TransID=B.SaleTransID " & _
        " where A.AccID= " & AccId & _
        " And (A.TransType= " & wDeposit & " or A.TransType= " & wContraDeposit & ")" & _
        " Group By A.AccID,TransID,TransDate,FaceValue,Balance"
        
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " UNION " & _
        " Select Count(*) as ShareCunt, A.AccId,A.TransID,TransDate,FaceValue,Balance, 'Returns' as Trans " & _
        " from MemTrans A Left join ShareTrans B " & _
        " On  A.AccID=B.AccID and A.TransID=B.ReturnTransID " & _
        " where A.AccID = " & AccId & _
        " And (A.TransType= " & wWithdraw & " or A.TransType= " & wContraWithdraw & ")" & _
        " Group By A.AccID,TransID,TransDate,FaceValue,Balance"
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " order by A.TransID"
    
    'gDbTrans.SQLStmt = "Select * from MemTrans where " & _
        " AccID = " & AccID & " order by TransID"
    ret = gDbTrans.Fetch(m_rstPassBook, adOpenDynamic)
    
    If ret < 0 Then GoTo DisableUserInterface
    If ret > 0 Then
        Dim BalanceAmount As Currency
        Dim MembershipFee As Double
        m_rstPassBook.MoveLast
        BalanceAmount = m_rstPassBook("Balance")
        'Position to first record of last page
        With m_rstPassBook
            .Move -1 * (.AbsolutePosition Mod 10)
            .MoveNext
        End With
    Else
        Set m_rstPassBook = Nothing
        PassBookPageInitialize
    End If
    
    grdTrans.Visible = False
    Call PassBookPageShow
    Me.grdTrans.Visible = True
'Get the Share details of this account holder
   gDbTrans.SqlStmt = "Select AccID,CertNo,SaleTransID,ReturnTransID," & _
        "FaceValue From ShareTrans A WHERE" & _
        " A.AccId = " & AccId & " Order by val(CertNO)"

    Call gDbTrans.Fetch(m_rstShareBook, adOpenStatic)
    
    grdShare.Visible = False
    Call ShareLeavesShow

    grdShare.Visible = True
    
'Do not allow for delete
    cmdDelete.Enabled = gCurrUser.IsAdmin
    If ret > 1 Then cmdDelete.Enabled = False
    
'Assign to some module level variables
    m_AccID = AccId
    m_accUpdatemode = Update
    
'Load account to the User Interface
'TAB 1
ClosedDate = FormatField(rstMaster("ClosedDate"))
With Me
    With .lblBalance
        .Caption = GetResourceString(42) & " " & _
            GetResourceString(312) & " " & FormatCurrency(BalanceAmount) '"Balance:  Rs. "
        If ClosedDate <> "" Then _
            .Caption = .Caption & " " & GetResourceString(524)
        
    End With
    
    With .txtCustName
        .Enabled = True: .BackColor = vbWhite
        .Caption = m_CustReg.FullName
    End With
    With .txtDate
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .Text = gStrDate
    End With
    With cmdDate
        .Enabled = True
        If gOnLine Then .Enabled = False
    End With
    With .cmbTrans
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .ListIndex = -1
    End With

    With .txtAmount
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .Value = 0
        '.Locked = True
    End With
    With .txtTotal
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .Caption = "0.00"
        '.Locked = True
    End With
    With .txtShareFees
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .Value = 0
    End With

    With .cmbLeaves
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        .Clear
    End With
    With cmdLeaves
        .Enabled = IIf(ClosedDate = "", True, False)
    End With
    cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
    cmdPrevTrans.Enabled = IIf(ClosedDate = "", True, False)
    cmdNextTrans.Enabled = IIf(ClosedDate = "", True, False)
    
    With .rtfNote
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
        Call m_Notes.LoadNotes(M_ModuleID, AccId)
    End With
    Call m_Notes.DisplayNote(.rtfNote)
    .cmdAccept.Enabled = IIf(ClosedDate = "", True, False)
    ' Code changed on 22/8/2000
    .cmdUndo.Enabled = IIf(ClosedDate = "", True, True) And gCurrUser.IsAdmin
    .cmdUndo.Caption = IIf(ClosedDate = "", GetResourceString(5), GetResourceString(313))
    
End With
    
    'TAB 2
    'Update labels and other buttons
    TabStrip2.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True
    'TabStrip2.Tabs.Item(1).Selected = True
    lblOperation.Caption = GetResourceString(56) '"Operation Mode : <UPDATE>"
    cmdDelete.Enabled = gCurrUser.IsAdmin

    Dim NomineeId As Long
    Dim IntroId As Long
    NomineeId = FormatField(rstMaster("NomineeID"))
    IntroId = FormatField(rstMaster("IntroducerID"))
    
    Dim strField As String
    Dim rstTemp As Recordset
    
    For I = 0 To txtPrompt.count - 1
        'Read the bound field of this control.
        On Error Resume Next
        strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
        If strField <> "" Then
            With txtData(I)
                Select Case UCase$(strField)
                    Case "ACCNUM"
                        .Text = rstMaster("AccNum")
                        '.Locked = True
                    Case "ACCNAME"
                        .Text = m_CustReg.FullName
                    Case "MEMBERFEE"
                        gDbTrans.SqlStmt = "Select * from MemIntTrans " & _
                            " where AccID = " & m_AccID & " ANd TransID =1"
                        If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                            .Text = FormatField(rstTemp("Amount"))
                        ''.Text = rstMaster("MembershipFee")
                    Case "NOMINEEID"
                        .Text = IIf(NomineeId, m_CustReg.CustomerName(NomineeId), "")
                    Case "NOMINEENAME"
                        .Text = IIf(NomineeId, m_CustReg.CustomerName(NomineeId), FormatField(rstMaster("NomineeName")))
                    Case "NOMINEERELATION"
                        .Text = FormatField(rstMaster("NomineeRelation"))
                    Case "INTRODUCERID"
                        .Text = IIf(IntroId, IntroId, "")
                    Case "INTRODUCERNAME"
                        .Text = IIf(IntroId, m_CustReg.CustomerName(IntroId), "")
                    Case "LEDGERNO"
                        .Text = rstMaster("LedgerNo")
                    Case "FOLIONO"
                        .Text = rstMaster("FolioNO")
                    Case "ACCGROUP"
                        gDbTrans.SqlStmt = "SELECT GroupName FROM AccountGroup WHERE " & _
                                "AccGroupID = " & FormatField(rstMaster("AccGroupId"))
                        If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
                        .Text = FormatField(rstTemp("GroupName"))
                    Case "CREATEDATE"
                        .Text = FormatField(rstMaster("CreateDate"))
                    Case "MEMBERTYPE"
                        txtIndex = ExtractToken(.Tag, "TextIndex")
                        If txtIndex <> "" Then _
                            cmb(txtIndex).ListIndex = Val(FormatField(rstMaster("MemberType"))) - 1
                        
                        .Text = cmb(txtIndex).Text
                    Case Else:
                        MsgBox "Label not found !", vbCritical, gAppName & " - Error"
                End Select
            End With
        End If
        
        
        Dim CtlIndex As Integer
        Dim CtlCount As Integer
        
        strField = ExtractToken(txtPrompt(I).Tag, "DisplayType")
        CtlIndex = Val(ExtractToken(txtPrompt(I).Tag, "TextIndex"))
        CtlCount = 0
        If strField <> "" Then
            With txtData(I)
              Select Case UCase$(strField)
                Case "LIST"
                    Do
                        If CtlCount = cmb(CtlIndex).ListCount Then Exit Do
                        If cmb(CtlIndex).List(CtlCount) = txtData(I).Text Then
                            cmb(CtlIndex).ListIndex = CtlCount
                            Exit Do
                        End If
                        CtlCount = CtlCount + 1
                    Loop
                
                Case "BOOLEAN"
                    'chk(CtlIndex).value = IIf(txtData(I).Text = True, vbChecked, vbUnchecked)
                    
              End Select
            End With
        End If
    Next

cmdPhoto.Enabled = Len(gImagePath)
AccountLoad = True
RaiseEvent AccountChanged(m_AccID)

Exit Function

DisableUserInterface:
    Call ResetUserInterface
    

Exit Function
    
ErrLine:
MsgBox "Account Load:" & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"

End Function

Private Sub Form_Load()

M_ModuleID = wis_Members

Screen.MousePointer = vbHourglass
Call CenterMe(Me)
'Centre the form
 Me.Move (Screen.Width - Me.Width) \ 2, _
            (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
cmdSharePrev.Picture = LoadResPicture(101, vbResIcon)
cmdShareNext.Picture = LoadResPicture(102, vbResIcon)
cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)
 
 'set kannada fonts
 Call SetKannadaCaption
 
'Fill up transaction Types
    
With cmbTrans
    .AddItem GetResourceString(302) '"Share Issued"   'Something like a deposit
    .AddItem GetResourceString(303) '"Share Returned"   'Something like a withdrawal (Don't get confused)
End With


'Load ICONS
cmdAddNote.Picture = LoadResPicture(103, vbResBitmap)

'Adjust the Grid for Pass book
Call PassBookPageInitialize
Call ShareLeavesInitialize

Call LoadPropSheet

Dim cmbIndex As Integer
cmbIndex = GetIndex("AccGroup")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
Call LoadAccountGroups(cmb(cmbIndex))

'load memebr Types
Dim rstMemType As Recordset
Dim recCount As Integer

cmbIndex = GetIndex("MemberType")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
    
    With cmb(cmbIndex)
        .Clear
        .AddItem GetResourceString(102, 49)
        .ItemData(.newIndex) = 1
        .AddItem GetResourceString(103, 49)
        .ItemData(.newIndex) = 2
        .AddItem GetResourceString(104, 49)
        .ItemData(.newIndex) = 3
    End With
    With cmbRepMemType
        .Clear
        .AddItem ""
        .ItemData(.newIndex) = 0
        .AddItem GetResourceString(102, 49)
        .ItemData(.newIndex) = 1
        .AddItem GetResourceString(103, 49)
        .ItemData(.newIndex) = 2
        .AddItem GetResourceString(104, 49)
        .ItemData(.newIndex) = 3
    End With
gDbTrans.SqlStmt = "Select * From MemberTypeTab order by membertype"
recCount = gDbTrans.Fetch(rstMemType, adOpenDynamic)
If recCount > 1 Or rstMemType("MemberType") > 0 Then
    cmbRepMemType.Locked = True
    With cmb(cmbIndex)
        .Clear
        cmbRepMemType.Clear
        While rstMemType.EOF = False
            .AddItem FormatField(rstMemType("MemberTypeName"))
            .ItemData(.newIndex) = FormatField(rstMemType("MemberType"))
            
            'Add to the Reports member type
            cmbRepMemType.AddItem FormatField(rstMemType("MemberTypeName"))
            cmbRepMemType.ItemData(cmbRepMemType.newIndex) = FormatField(rstMemType("MemberType"))
        
            'Move to next record
            rstMemType.MoveNext
        Wend
        
    End With
End If

'Load the Setup values
    Dim SetUp As New clsSetup
    If SetUp Is Nothing Then Set SetUp = New clsSetup
    txtPropMemFee.Text = SetUp.ReadSetupValue("MMAcc", "MemberShipFee", "0.00")
    txtPropMemCancel.Text = SetUp.ReadSetupValue("MMAcc", "Cancellation", "0.00")
    txtPropShareFee.Text = SetUp.ReadSetupValue("MMAcc", "ShareFee", "0.00")
    txtPropShareValue.Text = SetUp.ReadSetupValue("MMAcc", "ShareValue", "0.00")
    txtPropDivedend.Text = SetUp.ReadSetupValue("MMAcc", "RateOfDivedend", "0.00")
    txtPropLastDivedendOn = SetUp.ReadSetupValue("MMAcc", "LastDivedendOn", "")
    txtPropNextDivedendOn = SetUp.ReadSetupValue("MMAcc", "NextDivedendOn", "")
    cmdApply.Enabled = False

Dim rst As Recordset
'Reset the User Interface
Dim lstIndex As Integer
    Call ResetUserInterface
    Call optReports_Click(0)
    Call TabStrip_Click
  
Screen.MousePointer = vbDefault
txtDate2 = gStrDate
txtDate2.Enabled = True
txtDate2.BackColor = vbWhite

If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'gWindowHandle = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

' Notes object.
Set m_Notes = Nothing

' Customer Registration object.
Set m_CustReg = Nothing
gWindowHandle = 0
RaiseEvent WindowClosed

End Sub



Private Sub m_frmMMShare_ShareIssued(ShareNos() As String, Cancel As Boolean)
Dim I As Long
Dim MaxI As Integer
Dim rst As Recordset

MaxI = UBound(ShareNos)
'Check for duplicate certificate numbers
For I = 0 To MaxI
    gDbTrans.SqlStmt = "Select * from ShareTrans where " & _
        " CertNo = " & AddQuotes(ShareNos(I), True) & " And Accid = " & m_AccID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        'MsgBox "Some of the Share Certificate numbers you have specified have already been issued !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(589) & _
            "The Shares are already given to Member No. : " & _
            FormatField(rst("AccID")), vbExclamation, gAppName & " - Error"
        Cancel = True
        Exit Sub
    End If
Next I

'Fill up the leaf box
    cmbLeaves.Clear
    For I = 0 To MaxI
        cmbLeaves.AddItem ShareNos(I)
    Next I

'Calculate the total amount
Dim SetupClass As New clsSetup
Dim ShareValue As Currency
    ShareValue = Val(SetupClass.ReadSetupValue("MMAcc", "ShareValue", "10"))
    txtAmount = ShareValue * (MaxI + 1)

End Sub

Private Sub m_frmMMShare_ShareReturned(Leaves() As String)
Dim I As Long
'Fill up the combo
    cmbLeaves.Clear
    For I = 0 To UBound(Leaves)
        cmbLeaves.AddItem Leaves(I)
    Next I


End Sub

Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer
'1. Fetch last print row from PDmaster table.
'First get the last printed txnID From the memMaster
SqlStr = "SELECT  LastPrintRow From MemMaster WHERE AccId = " & m_AccID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintRow = IIf(IsNull(metaRst("LastPrintrow")), 0, metaRst("LastPrintrow"))


'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From MemTrans WHERE AccId = " & m_AccID
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
    SqlStr = "SELECT * From MemTrans WHERE AccId = " & m_AccID & _
        " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
        " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#"
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If


'Printer.PaperSize = 9
'Printer.Font.Name = gFontName
'Printer.Font.Size = 12 'gFontSize
Printer.Font = "Courier New"
Printer.FONTSIZE = 9
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
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
       ' .ColWidth(5) = 15

    While Not rst.EOF
        If .isNewPage Then
            .printHeader1
            .isNewPage = False
        End If
        .ColText(0) = FormatField(rst("TransDate"))
        '.ColText(1) = FormatField(Rst("ChequeNo"))
        .ColText(1) = FormatField(rst("Particulars"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(2) = FormatField(rst("Amount"))
        Else
            .ColText(3) = FormatField(rst("Amount"))
        End If
        .ColText(4) = FormatField(rst("Balance"))
   .PrintRow
        
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
SqlStr = "UPDATE MemMaster set LastPrintRow = " & curPrintRow - 1 & _
        " Where Accid = " & m_AccID
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
Const HEADER_ROWS = 3
Dim curPrintRow As Integer


On Error Resume Next

'First get the last printed txnId and last printed row From the MemMaster
SqlStr = "SELECT LastPrintID, LastPrintRow From MemMaster WHERE AccId = " & m_AccID
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintId"))

' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From MemTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
' Print the first page of passbook, if newPassbook option is chosen.
If bNewPassbook Then
    clsPrint.printPassbookPage wis_Members, m_AccID
SqlStr = "UPDATE MemMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
        ", LastPrintRow = 0 Where Accid = " & m_AccID
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
SqlStr = "SELECT * From MemTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
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
Printer.Font = "Courier New"
Printer.FONTSIZE = 9
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
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
     .ColWidth(0) = 20
        .ColWidth(1) = 10
        .ColWidth(2) = 25
        .ColWidth(3) = 15
        .ColWidth(4) = 18

   While Not rst.EOF
        If .isNewPage Then
            .printHeader1
            .isNewPage = False
        End If

        TransID = FormatField(rst("TransID"))
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        '.ColText(2) = FormatField(Rst("Particulars"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(2) = FormatField(rst("Amount"))
        Else
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
SqlStr = "UPDATE MemMaster set LastPrintId = " & TransID & _
        ", LastPrintRow = " & curPrintRow - 1 & _
        " Where Accid = " & m_AccID
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If


End Sub

Private Sub optReports_Click(Index As Integer)

Dim Dt1 As Boolean
Dim Amt1 As Boolean
Dim Amt2 As Boolean
Dim Gender As Boolean
Dim MemType As Boolean

'Disable the Other Optiomn Buttons
Select Case Index

    Case 0
        Amt1 = True: Amt2 = True
        Gender = True
        MemType = True
    Case 1
        optMemId.Enabled = True
        optName.Enabled = True
        Gender = True
        MemType = True
    Case 2, 3
        Dt1 = True
        Gender = True
        MemType = True
    
    Case 4
        Dt1 = True
        Gender = True
        MemType = True
    Case 5
        Dt1 = True
        Gender = True
        MemType = True
    Case 6
        Dt1 = True
        Gender = True
        MemType = True
    Case 7
        'lblAmt1 = GetResourceString(147) & " " & _
                GetResourceString(53) & " " & _
                GetResourceString(60)
        'lblAmt2 = GetResourceString(148) & " " & _
                GetResourceString(53) & " " & _
                GetResourceString(60)
        Amt1 = True: Amt2 = True: Dt1 = True
        Gender = True
        MemType = True
    Case 8
        Gender = False
        MemType = False
        Amt1 = False: Amt2 = False: Dt1 = False
    
    Case 8, 9
        Dt1 = True
    Case 10, 11
        Amt1 = False: Amt2 = False
        Gender = True
        MemType = True
    
    
    
End Select


    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = Gender
        .EnableAmountRange = Amt1
    End With

    With txtDate1
        .Enabled = Dt1:
        .BackColor = IIf(Dt1, wisWhite, wisGray)
    End With
    lblDate1.Enabled = Dt1
    cmdDate1.Enabled = Dt1


    With cmbRepMemType
        .Enabled = MemType
        .BackColor = IIf(MemType, wisWhite, wisGray)
    End With
    lblMemType.Enabled = MemType
    
End Sub

Private Sub TabStrip_Click()

fraTransact.Visible = False
fraNew.Visible = False
fraReports.Visible = False
fraProps.Visible = False

Select Case TabStrip.SelectedItem.Index
    Case 1
        fraTransact.Visible = True
        cmdAccept.Default = True
    
    Case 2
        fraNew.Visible = True
        txtData(1).SetFocus
        cmdSave.Default = True
    
    Case 3
        fraReports.Visible = True
        cmdView.Default = True
        
    Case 4
        fraProps.Visible = True
        cmdApply.Default = True
        
End Select

End Sub

Private Sub TabStrip2_Click()
Dim BoolAccept As Boolean
Dim BoolTabstrip As Boolean
Dim BoolUndo As Boolean

BoolAccept = cmdAccept.Enabled
BoolTabstrip = TabStrip2.Enabled
BoolUndo = cmdUndo.Enabled


cmdAccept.Enabled = False
TabStrip2.Enabled = False
cmdUndo.Enabled = False

fraInstructions.Visible = False
fraPassBook.Visible = False
fraShare.Visible = False

If TabStrip2.SelectedItem.Index = 1 Then fraInstructions.Visible = True: fraInstructions.ZOrder 0
If TabStrip2.SelectedItem.Index = 2 Then fraPassBook.Visible = True: fraPassBook.ZOrder 0
If TabStrip2.SelectedItem.Index = 3 Then fraShare.Visible = True: fraShare.ZOrder 0


ExitLine:

    cmdAccept.Enabled = BoolAccept
    TabStrip2.Enabled = BoolTabstrip
    cmdUndo.Enabled = BoolUndo

End Sub

Private Sub txtAccNo_Change()
cmdLoad.Enabled = IIf(Trim$(txtAccNo.Text) <> "", True, False)

Call ResetUserInterface
End Sub


Private Sub txtAmount_Change()

txtTotal = FormatCurrency(txtShareFees + txtAmount)

End Sub

Private Sub txtData_DblClick(Index As Integer)
txtData_KeyPress Index, vbKeyReturn
End Sub
Private Sub txtData_GotFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlue
SetDescription txtPrompt(Index)

' Scroll the window, so that the
' control in focus is visible.
ScrollWindow txtData(Index)

' Select the text, if any.
With txtData(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' If the display type is Browse, then
' show the command button for this text.
Dim strDispType As String
Dim TextIndex As String
strDispType = ExtractToken(txtPrompt(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    If TextIndex <> "" Then cmd(Val(TextIndex)).Visible = True
End If

' Hide all other command buttons...
Dim I As Integer
For I = 0 To cmd.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then cmd(I).Visible = False
Next

If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
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

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(txtPrompt(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(txtPrompt(Index).Tag, "TextIndex")
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
Private Sub txtData_LostFocus(Index As Integer)

txtPrompt(Index).ForeColor = vbBlack
Dim strDatSrc As String
Dim Lret As Long
Dim txtIndex As Integer
Dim rst As Recordset

' If the item is IntroducerID, validate the
' ID and name.
strDatSrc = ExtractToken(txtPrompt(Index).Tag, "DataSource")
If StrComp(strDatSrc, "IntroducerID", vbTextCompare) = 0 Then
    ' Check if any data is found in this text.
    If Trim$(txtData(Index).Text) <> "" Then
        If Val(txtData(Index).Text) <= 0 Then
            'MsgBox "Invalid member ID specified !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(760), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtData(Index)
            Exit Sub
        End If
        gDbTrans.SqlStmt = "SELECT MemMaster.CustomerID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM MemMaster, " _
                & "NameTab WHERE MemMaster.AccID = " & Val(txtData(Index).Text) _
                & " AND MemMaster.CustomerID = NameTab.CustomerID"
        Lret = gDbTrans.Fetch(rst, adOpenDynamic)
        If Lret > 0 Then
            txtIndex = GetIndex("IntroducerName")
            txtData(txtIndex).Text = FormatField(rst("Name"))
        End If
    Else
        txtIndex = GetIndex("IntroducerName")
        txtData(txtIndex).Text = ""
    End If
End If
End Sub

Private Sub txtFromDate_Change()
    cmdDividend.Enabled = True
End Sub

Private Sub txtPrompt_GotFocus(Index As Integer)
    txtPrompt(Index).ForeColor = vbBlue
End Sub

Private Sub txtPrompt_LostFocus(Index As Integer)
    txtPrompt(Index).ForeColor = vbBlack
End Sub













Private Sub txtPropDivedend_Change()
    cmdApply.Enabled = True
End Sub



Private Sub txtPropLastDivedendOn_Change()
    'cmdApply.Enabled = True
End Sub

Private Sub txtPropMemCancel_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPropMemFee_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPropNextDivedendOn_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPropShareFee_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPropShareValue_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtShareFees_Change()

Call txtAmount_Change

End Sub

Private Sub VScroll1_Change()
' Move the picSlider.
picSlider.Top = -VScroll1.Value
End Sub


Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

'Set kannada fonts for all the tabs
With TabStrip
    .Tabs(1).Caption = GetResourceString(38)
    .Tabs(2).Caption = GetResourceString(211)
    .Tabs(3).Caption = GetResourceString(283) & GetResourceString(92)
    .Tabs(4).Caption = GetResourceString(213)
End With
With TabStrip2
    .Tabs(1).Caption = GetResourceString(219)
    .Tabs(2).Caption = GetResourceString(38)
    .Tabs(3).Caption = GetResourceString(53) & " " & _
            GetResourceString(295)  ' "Share Details"
End With

'for general form
cmdLoad.Caption = GetResourceString(3)
cmdLeaves.Caption = GetResourceString(53)
cmdOk.Caption = GetResourceString(1)
'for Tabstrip1 i.e Transaction Tab

lblMemNo.Caption = GetResourceString(49) + GetResourceString(60)
lblName.Caption = GetResourceString(35)
lblDate.Caption = GetResourceString(37)
lblTrans.Caption = GetResourceString(38)
lblLeaves.Caption = GetResourceString(53)
lblShareFee.Caption = GetResourceString(53, 191)
lblAmount.Caption = GetResourceString(40)
lblTOtal.Caption = GetResourceString(52) '
lblBalance.Caption = GetResourceString(42)
cmdUndo.Caption = GetResourceString(5)
cmdAccept.Caption = GetResourceString(4)
cmdSharePrint.Picture = LoadResPicture(120, vbResBitmap)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)

'for frame New/modify Account Tab2
cmdDelete.Caption = GetResourceString(14)
cmdSave.Caption = GetResourceString(7) '
cmdReset.Caption = GetResourceString(8)
lblOperation.Caption = GetResourceString(54)
cmdPhoto.Caption = GetResourceString(415)
'for frame Reports tab3
fraChooseReport.Caption = GetResourceString(288)
optReports(0).Caption = GetResourceString(61)  '
optReports(1).Caption = GetResourceString(49, 295) '
optReports(2).Caption = GetResourceString(390, 63) 'sub day book
optReports(3).Caption = GetResourceString(390, 85) ' Sub cash book
optReports(4).Caption = GetResourceString(94) '
optReports(5).Caption = GetResourceString(95) '
optReports(6).Caption = GetResourceString(96) '
optReports(7).Caption = GetResourceString(53, 337)
optReports(8).Caption = GetResourceString(463, 42) 'Monthly Balance
optReports(9).Caption = GetResourceString(49, 93) 'Memebr Gen ledger
optReports(10).Caption = GetResourceString(426) 'Loan Members
optReports(11).Caption = GetResourceString(427) 'No-Loan Member
optName.Caption = GetResourceString(69) '
optMemId.Caption = GetResourceString(68)

lblMemType.Caption = GetResourceString(101)

fraMemType.Caption = GetResourceString(101)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
cmdView.Caption = GetResourceString(13) '
lblPropMemFee.Caption = GetResourceString(79, 191)
lblPropMemCancel.Caption = GetResourceString(79, 195)
'fraShare.Caption = GetResourceString(53,295)
lblPropShareFee.Caption = GetResourceString(53, 191) '"
lblPropShareValue.Caption = GetResourceString(53, 140)
fraPropDivedend.Caption = GetResourceString(200) '
lblpropDivedendRate.Caption = GetResourceString(200, 305)
lblPropLastDivedendOn.Caption = GetResourceString(202)
lblPropNextDivedendOn = GetResourceString(203)
cmdApply.Caption = GetResourceString(6)
lblFromDate.Caption = GetResourceString(109)
lblToDate.Caption = GetResourceString(110)
cmdDividend.Caption = GetResourceString(200)

lblUndoInterest.Caption = GetResourceString(228)
lblFailAccIDs.Caption = GetResourceString(190)
cmdUndoInterests.Caption = GetResourceString(188)
cmdAdvance.Caption = GetResourceString(491)    'Options
End Sub

