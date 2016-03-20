VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmUtils 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilities"
   ClientHeight    =   6090
   ClientLeft      =   2175
   ClientTop       =   1530
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Uitls"
      Height          =   4600
      Index           =   0
      Left            =   200
      TabIndex        =   62
      Top             =   600
      Width           =   6400
      Begin VB.OptionButton optMember 
         Caption         =   "New Member Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   390
         TabIndex        =   72
         Top             =   1185
         Width           =   2805
      End
      Begin VB.OptionButton optFarmer 
         Caption         =   "New Farmer Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   390
         TabIndex        =   71
         Top             =   2640
         Width           =   2805
      End
      Begin VB.OptionButton optPlace 
         Caption         =   "New Place"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3390
         TabIndex        =   70
         Top             =   480
         Width           =   2865
      End
      Begin VB.OptionButton optCustomer 
         Caption         =   "New Customer Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3390
         TabIndex        =   69
         Top             =   1920
         Width           =   2865
      End
      Begin VB.OptionButton optCaste 
         Caption         =   "New Caste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   68
         Top             =   480
         Width           =   2865
      End
      Begin VB.OptionButton optAccount 
         Caption         =   "New Account Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   390
         TabIndex        =   67
         Top             =   1980
         Width           =   2865
      End
      Begin VB.OptionButton optLoanPurpose 
         Caption         =   "New Loan Purpose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3390
         TabIndex        =   66
         Top             =   2640
         Width           =   2805
      End
      Begin VB.OptionButton optDeposit 
         Caption         =   "New Deposit Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3390
         TabIndex        =   65
         Top             =   1200
         Width           =   2865
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4950
         TabIndex        =   63
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblNewMessage 
         Caption         =   "X"
         Height          =   495
         Left            =   480
         TabIndex        =   64
         Top             =   3210
         Width           =   5565
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Uitls"
      Height          =   4600
      Index           =   1
      Left            =   200
      TabIndex        =   0
      Top             =   600
      Width           =   6400
      Begin VB.OptionButton optPdRepair 
         Caption         =   "Repair Pigmy Trans"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   2205
         Width           =   3135
      End
      Begin VB.OptionButton optCompareDb 
         Caption         =   "Compare the database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   1282
         Width           =   3255
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4950
         TabIndex        =   25
         Top             =   3960
         Width           =   1215
      End
      Begin VB.OptionButton optPrintOrder 
         Caption         =   "Print Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optCompact 
         Caption         =   "Compact the data base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   1282
         Width           =   2775
      End
      Begin VB.OptionButton optBackUp 
         Caption         =   "Back Up the data base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   2205
         Width           =   2775
      End
      Begin VB.OptionButton optEndOFDay 
         Caption         =   "End Of The Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblMessage 
         Caption         =   "X"
         Height          =   615
         Left            =   360
         TabIndex        =   26
         Top             =   2970
         Width           =   5565
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5490
      TabIndex        =   1
      Top             =   5550
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Year end function"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4600
      Index           =   6
      Left            =   200
      TabIndex        =   38
      Top             =   600
      Width           =   6405
      Begin VB.CommandButton cmdYearEnd 
         Caption         =   "Perform year end operation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1230
         TabIndex        =   39
         Top             =   1650
         Width           =   3825
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Other Set Up"
      Height          =   4600
      Index           =   4
      Left            =   200
      TabIndex        =   28
      Top             =   600
      Width           =   6400
      Begin VB.CheckBox chkNewBalanceSheet 
         Caption         =   "Show Progressive Balance Sheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   61
         Top             =   1176
         Width           =   4875
      End
      Begin VB.CheckBox chkTradeHeads 
         Caption         =   "Show all expense heads in trading"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   60
         Top             =   4065
         Width           =   4755
      End
      Begin VB.TextBox txtImagePath 
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Text            =   "c:\Index2000\ImagePath"
         Top             =   3480
         Width           =   3615
      End
      Begin VB.CheckBox chkKccSapTrans 
         Caption         =   "Saperate transaction for KCC Deposit && KCC Loan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   37
         Top             =   723
         Width           =   5475
      End
      Begin VB.ComboBox cmbDateFormat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2700
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   2880
         Width           =   2085
      End
      Begin VB.CheckBox chkSubHeadTotal 
         Caption         =   "Show Sub head total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   33
         Top             =   2355
         Width           =   4935
      End
      Begin VB.CheckBox chkRPFormat 
         Caption         =   "Show heads in Receipt && payment both side"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   32
         Top             =   2025
         Width           =   4875
      End
      Begin VB.CheckBox chkNegBal 
         Caption         =   "Show Nagative amount in Balance sheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   31
         Top             =   1629
         Width           =   4695
      End
      Begin VB.CheckBox chkKCC 
         Caption         =   "Consider Negative balance as deposit in KCC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   330
         Width           =   5355
      End
      Begin VB.CommandButton cmdApply4 
         Caption         =   "&Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5040
         TabIndex        =   29
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label lblImagePath 
         Caption         =   "Path to store Images"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblDateFormat 
         Caption         =   "Select the Date format "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Other Set Up"
      Height          =   4600
      Index           =   5
      Left            =   200
      TabIndex        =   44
      Top             =   600
      Width           =   6400
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Toggle key"
         Height          =   315
         Left            =   2760
         TabIndex        =   59
         Top             =   3360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         Left            =   2760
         TabIndex        =   57
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtTaluka 
         Height          =   375
         Left            =   2760
         TabIndex        =   54
         Text            =   "Gadag"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtGuarnteer 
         Height          =   375
         Left            =   2760
         TabIndex        =   52
         Text            =   "2"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox txtState 
         Height          =   375
         Left            =   2760
         TabIndex        =   50
         Text            =   "Karnataka"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtPinCode 
         Height          =   375
         Left            =   2760
         TabIndex        =   48
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton cmdApplyExtra 
         Caption         =   "&Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4920
         TabIndex        =   46
         Top             =   3900
         Width           =   1335
      End
      Begin VB.TextBox txtDistrict 
         Height          =   375
         Left            =   2760
         TabIndex        =   45
         Text            =   "Gadag"
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblFont 
         Caption         =   "Select Font"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblGuranteer 
         Caption         =   "Guaranteers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPinCode 
         Caption         =   "Pin code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblDistrict 
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblTaluka 
         Caption         =   "Taluka :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Cashier Set Up"
      Height          =   4600
      Index           =   3
      Left            =   200
      TabIndex        =   17
      Top             =   600
      Width           =   6400
      Begin VB.CheckBox chkTradingCash 
         Caption         =   "Registar Trading's Cash Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   34
         Top             =   3390
         Width           =   5895
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4860
         TabIndex        =   23
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CheckBox chkCashIndex 
         Caption         =   "Show Cash Denomination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   22
         Top             =   2202
         Width           =   3915
      End
      Begin VB.CheckBox chkCashPayment 
         Caption         =   "Transfer payment to cashier window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   21
         Top             =   2796
         Width           =   3915
      End
      Begin VB.CheckBox chkPassing 
         Caption         =   "Show Passing Officer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   20
         Top             =   1608
         Width           =   3915
      End
      Begin VB.CheckBox chkCashReceipt 
         Caption         =   "Receipts at cashier window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   19
         Top             =   1014
         Width           =   3915
      End
      Begin VB.CheckBox chkCashWindow 
         Caption         =   "Show cashier windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   18
         Top             =   420
         Width           =   3915
      End
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4600
      Index           =   2
      Left            =   200
      TabIndex        =   6
      Top             =   600
      Width           =   6405
      Begin VB.ComboBox cmbEnglish 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdAdvance 
         Caption         =   "&Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4950
         TabIndex        =   27
         Top             =   3930
         Width           =   1335
      End
      Begin VB.ComboBox cmb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3420
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         TabIndex        =   13
         Text            =   "22/22/2222"
         Top             =   2070
         Width           =   1485
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   315
         Left            =   5820
         TabIndex        =   12
         Top             =   2070
         Width           =   315
      End
      Begin VB.OptionButton optTransId 
         Caption         =   "Update the Transid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2730
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.OptionButton optUndoInt 
         Caption         =   "Undo interest receivable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   2115
         Width           =   3195
      End
      Begin VB.OptionButton optIntReceivable 
         Caption         =   "Interest Receivable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   3195
      End
      Begin VB.OptionButton optOpBalance 
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   885
         Width           =   3195
      End
      Begin VB.OptionButton optClgBank 
         Caption         =   "Set Clearing Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   3195
      End
      Begin VB.Label lblAccName 
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3390
         TabIndex        =   16
         Top             =   390
         Width           =   2655
      End
      Begin VB.Label lblObDate 
         Caption         =   "OB Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3420
         TabIndex        =   15
         Top             =   1500
         Width           =   2175
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5235
      Left            =   90
      TabIndex        =   24
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9234
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Create New"
            Key             =   "New"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Utilities"
            Key             =   "Utils"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advance"
            Key             =   "Advance"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cashier"
            Key             =   "Cash"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Other"
            Key             =   "Other"
            Object.Tag             =   ""
            Object.ToolTipText     =   "shows other set up"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Extra"
            Key             =   "Extra"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Extra"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_AddGroup As clsAddGroup
Attribute m_AddGroup.VB_VarHelpID = -1

Public Event WindowClosed()

Private Function AddNewGroup()

Dim grpType As Long
If optAccount Then grpType = grpAccount
If optCaste Then grpType = grpCaste
If optPlace Then grpType = grpPlace
If optCustomer Then grpType = grpCustomer
If optFarmer Then grpType = grpFarmer
If optDeposit Then grpType = grpDeposit
If optLoanPurpose Then grpType = grpLoanPurpose
If optMember Then grpType = grpMember

'If optproduct Then grpType = grpProduct
'If optunit Then grdpType = grpUnit

If grpType = 0 Then
    Call DatabaseFunctions
    Exit Function
End If

If m_AddGroup Is Nothing Then Set m_AddGroup = New clsAddGroup

m_AddGroup.ShowAddGroup (grpType)

AddNewGroup = True

End Function

Private Sub BankSetup()

Dim LstCount As Integer
Dim rst As Recordset

If optClgBank Then
    With cmb
        If .ListCount < 1 Then Exit Sub
        If .ListIndex < 0 Then
            MsgBox "Please select the bank " & _
                "to recognise as clearing bank", vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
        LstCount = .ItemData(.ListIndex)
    End With
    
    gDbTrans.SqlStmt = "Select * From install Where KeyData = 'ClearingBankId'"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        gDbTrans.SqlStmt = "Update Install Set ValueData = '" & LstCount & "'" & _
            " Where KeyData = 'ClearingBankId'"
    Else
        gDbTrans.SqlStmt = "Insert Into Install(KeyData , ValueData) " & _
            " Values ( 'ClearingBankId', '" & LstCount & "')"
    End If
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        MsgBox "Unable to set the Clearing Bank"
        Exit Sub
    End If
    gDbTrans.CommitTrans

ElseIf optTransId Then
    Call UpdateTransactionID

ElseIf optOpBalance Then
    Call SetOpeningBalance
ElseIf optIntReceivable Then
    Call LoanInterestReceivable
ElseIf optUndoInt Then
    Call UndoInterestReceivable
End If


End Sub

Private Sub DatabaseFunctions()

Dim dstFile As String
Dim srcFile As String
Dim DBPath As String
Dim DbName As String
Dim DbUtilClass As clsDBUtilities
Set DbUtilClass = New clsDBUtilities

srcFile = GetDataBaseName
    
If Left(srcFile, 2) = "\\" Then Exit Sub

    
If optBackUp Then

    If Not DbUtilClass.MakeBackUp(srcFile) Then
        MsgBox "Can not take the Back up ", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
  #If TargetFile Then
    'Select The Path Name to copy the INdex 2000
    With wisMain.cdb
        .CancelError = False
        .DefaultExt = "MDB"
        .DialogTitle = "Copy the database"
        .ShowSave
        dstFile = .Filename
        If dstFile = "" Then Exit Sub
        count = InStr(1, .Filename, .FileTitle)
        dstPath = Left(.Filename, count - 2)
    End With
    
    'Check the PAth where data base is keeping
    If UCase(dstPath) = UCase(App.Path) Then
        MsgBox "Back up can not be copied to the application path", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    'before copying the database CLose the database
    Call gDbTrans.CloseDB
    'SetAttr
    'Copy the database tothe data
    
    FileCopy srcFile, dstFile
    
    'Now open the Data base
    If Not gDbTrans.OpenDB(srcFile, "PRAGMANS") Then
        MsgBox "Can not take the Back up ", vbInformation, wis_MESSAGE_TITLE
        End
    End If
    
  #End If
  
    'Now write the information about the back up
    'to the registry
    'Write the PAth of the backup
    'Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "BackupPath", dstPath)
    'Write the date of the backup
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "BackupDate", Format(Now, "dd/mm/yyyy"))
    'SetRegistryValue
    Screen.MousePointer = vbDefault
    
    Exit Sub
End If

If optCompact Then
    Screen.MousePointer = vbHourglass
    If Not DbUtilClass.CompactTheDataBase(srcFile, "WIS!@#") Then
        Screen.MousePointer = vbDefault
        MsgBox "Can not compact the database", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
        
    'Now write the information about the Compact
    'into the registry
    'Write the date of the Compact
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "CompactDate", Format(Now, "dd/mm/yyyy"))
    'SetRegistryValue
    Screen.MousePointer = vbDefault
End If


If optCompareDb Then
    Screen.MousePointer = vbHourglass
    Dim pos As Integer
    Do
        pos = InStr(pos + 1, srcFile, "\", vbTextCompare)
        If pos = 0 Then Exit Do
        DbName = Mid(srcFile, pos + 1)
        dstFile = Left(srcFile, pos - 1)
    Loop
    If dstFile = "" Then GoTo Exit_Line
    'Move one folder behind
    pos = InstrRev(dstFile, "\")
    If pos Then dstFile = Left(dstFile, pos - 1)
    DbName = Mid(DbName, 1, Len(DbName) - 4) & " BLANK.mdb"
    dstFile = dstFile & "\BlankDataBase\" & DbName
    
    If Not DbUtilClass.CompareDBFromDB(dstFile, "WIS!@#") Then
        MsgBox "Can not compact the database", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
        
    'Now write the information about the Compact
    'into the registry
    'Write the date of the Compact
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "CompareDate", Format(Now, "dd/mm/yyyy"))
    'SetRegistryValue
    Screen.MousePointer = vbDefault
End If

Exit_Line:
Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
Set DbUtilClass = Nothing
End Sub

Private Sub GetInfo()

'fra.Visible = False
Dim strProcess As String

If optBackUp Then strProcess = "BackUp"
If optCompact Then strProcess = "Compact"
'If optZip Then strProcess = "Compress"

If strProcess = "" Then lblMessage = "": Exit Sub

Dim strDate As String

'Get the last date of the backup
strDate = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, strProcess & "DATE")
If strDate = "" Then
    lblMessage = "System does not have any information " & _
        "about when last " & strProcess & " was taken"
Else
    lblMessage = "The last " & strProcess & " was taken on " & strDate
End If
End Sub

Private Sub RepairPigmyDeposits()
    Dim bankClass As New clsBankAcc
    Call bankClass.RepairPigmyDeposits
    Set bankClass = Nothing
End Sub

Private Sub WriteSetUp3()

Dim SetUp As clsSetup

Set SetUp = New clsSetup

Dim ValueData As String

ValueData = IIf(chkCashWindow.Value = vbChecked, "True", "False")
Call SetUp.WriteSetupValue("General", "CashierWindow", UCase(ValueData))

ValueData = IIf(chkCashIndex.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "CashIndex", UCase(ValueData))


ValueData = IIf(chkCashReceipt.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "CashReceipt", UCase(ValueData))

ValueData = IIf(chkCashPayment.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "CashPayment", UCase(ValueData))

ValueData = IIf(chkPassing.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "Passing", UCase(ValueData))

ValueData = IIf(chkTradingCash.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("Trading", "RegisterCashTransaction", UCase(ValueData))

Set SetUp = Nothing

End Sub

Private Sub WriteSetUp4()

Dim SetUp As clsSetup

Set SetUp = New clsSetup

Dim ValueData As String

ValueData = IIf(chkNegBal.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "NegativeInBalanceSheet", UCase(ValueData))

'KCC Deposit
'Check whether ther are two ledger heads for KCC
'1 for Deposit & one for KCC loan
ValueData = IIf(chkKCC.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "KCCDeposit", UCase(ValueData))
'Check whethe at a time we can make two transaction for KCc Deposit & KCc loan
ValueData = IIf(chkKCC.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "KCCTwoTransaction", UCase(ValueData))

ValueData = IIf(chkNegBal.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "NegativeInBalanceSheet", UCase(ValueData))

ValueData = IIf(chkRPFormat.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "RPorPLBothHeads", UCase(ValueData))

ValueData = IIf(chkSubHeadTotal.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "ShowRPorPLTotal", UCase(ValueData))

ValueData = IIf(chkNewBalanceSheet.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("General", "ShowNewBalSheet", UCase(ValueData))

If cmbDateFormat.ListCount >= 0 Then
    DateFormat = cmbDateFormat.Text
    Call SetUp.WriteSetupValue("General", "DateFormat", DateFormat)
End If
If Len(txtImagePath.Text) > 0 Then
    DateFormat = cmbDateFormat.Text
    gImagePath = Trim$(txtImagePath.Text)
    If Mid$(gImagePath, Len(gImagePath) - 1) <> "\" Then gImagePath = gImagePath + "\"
    
    Call SetUp.WriteSetupValue("General", "ImagePath", gImagePath)
End If

ValueData = IIf(chkTradeHeads.Value = vbChecked, "TRUE", "False")
Call SetUp.WriteSetupValue("Trading", "ShowAllExpenseHeads", UCase(ValueData))

Set SetUp = Nothing

End Sub


Private Sub ReadSetUp()

Dim SetUp As clsSetup
Set SetUp = New clsSetup

Dim ValueData As String
Dim count As Integer

ValueData = UCase(SetUp.ReadSetupValue("General", "CashierWindow", "False"))
chkCashWindow.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "CashIndex", "False"))
chkCashIndex.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "CashReceipt", "False"))
chkCashReceipt.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "CashPayment", "False"))
chkCashPayment.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "Passing", "False"))
chkPassing.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "KCCDeposit", "True"))
chkKCC.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "KCCTwoTransaction", "True"))
chkKccSapTrans.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)


ValueData = UCase(SetUp.ReadSetupValue("General", "NegativeInBalanceSheet", "False"))
chkNegBal.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)


ValueData = UCase(SetUp.ReadSetupValue("General", "RPorPLBothHeads", "True"))
chkRPFormat.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)
Call chkRPFormat_Click

ValueData = UCase(SetUp.ReadSetupValue("Trading", "ShowAllExpenseHeads", "True"))
chkTradeHeads.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "ShowRPorPLTotal", "True"))
chkSubHeadTotal.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = UCase(SetUp.ReadSetupValue("General", "ShowNewBalSheet", "True"))
chkNewBalanceSheet.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)



ValueData = UCase(SetUp.ReadSetupValue("Trading", "RegisterCashTransaction", "False"))
chkTradingCash.Value = IIf(ValueData = "TRUE", vbChecked, vbUnchecked)

ValueData = SetUp.ReadSetupValue("General", "ImagePath", "")
txtImagePath.Text = ValueData

ValueData = UCase(SetUp.ReadSetupValue("General", "DateFormat", "dd/mm/yyyy"))
With cmbDateFormat
    count = 0
    Do
        If count > .ListCount Then Exit Do
        If ValueData = UCase(.List(count)) Then .ListIndex = count: Exit Do
        count = count + 1
    Loop
End With

'Extra

txtTaluka = SetUp.ReadSetupValue("Customer", "Taluka", "")
txtDistrict = SetUp.ReadSetupValue("Customer", "District", "")
txtPinCode = SetUp.ReadSetupValue("Customer", "PINCODE", "")
txtState = SetUp.ReadSetupValue("Customer", "State", "Karnataka")
txtGuarnteer = SetUp.ReadSetupValue("Loan", "NoOfGuaranteers", "2")

Call SetComboIndex(cmbFont, SetUp.ReadSetupValue("General", "FontName", cmbFont.List(0)))

Set SetUp = Nothing


End Sub


Private Function SetOpeningBalance() As Boolean

If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If

Dim TransDate As Date
Dim ModuleId As wisModules
Dim headID As Long
Dim SchemeID As Long
Dim AccountType As wis_AccountType

TransDate = GetSysFormatDate(txtDate)

With cmb
    If .ListIndex < 0 Then
        MsgBox "Please select the account type", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    headID = .ItemData(.ListIndex)
End With
ModuleId = GetModuleIDFromHeadID(headID)

If ModuleId = wis_None Then Exit Function

If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

gDbTrans.SqlStmt = ""

Dim TableName As String
If ModuleId = wis_Members Then
    TableName = "Mem"
ElseIf ModuleId = wis_SBAcc Then
    TableName = "SB"
ElseIf ModuleId = wis_CAAcc Then
    TableName = "CA"
ElseIf ModuleId = wis_RDAcc Then
    TableName = "RD"
ElseIf ModuleId = wis_PDAcc Then
    TableName = "PD"
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    TableName = "BKCC"
ElseIf ModuleId = wis_Loans Then
    TableName = "Loan"
ElseIf ModuleId = wis_DepositLoans Then
    TableName = "DepositLoan"
End If

If ModuleId = wis_Members Or ModuleId = wis_SBAcc _
    Or ModuleId = wis_CAAcc Or ModuleId = wis_PDAcc _
    Or ModuleId = wis_RDAcc Then
    AccountType = Liability
    gDbTrans.SqlStmt = "Select AccNum,Balance,A.AccID,TransID,TransType," & _
        "Amount,FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From " & TableName & "Master A," & TableName & "Trans B,NameTab C " & _
        "Where A.AccID = B.AccID And C.CustomerID = A.CustomerID And TransID = " & _
            "(Select Min(TransID) From " & TableName & "Trans C " & _
            " Where C.AccID = B.AccID And TransDate >= #" & TransDate & "#)" & _
        " Order By val(AccNum)"

ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,Balance,A.LoanID as AccID,TransID,TransType," & _
        "Amount,FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From BkccMaster A,BkccTrans B,NameTab C Where A.LoanID = B.LoanID " & _
        "And C.CustomerID = A.CustomerID And TransID = " & _
            "(Select Min(TransID) From BkccTrans C " & _
            " Where C.LoanID = B.LoanID And TransDate >= #" & TransDate & "#)" & _
        " ORder By val(AccNum)"

ElseIf ModuleId = wis_Loans Then
    
    SchemeID = ModuleId Mod 100
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,Balance,A.LoanID as AccID,TransID,TransType," & _
        "Amount,FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From LoanMaster A,LoanTrans B,NameTab C Where A.LoanID = B.LoanID " & _
        "And C.CustomerID = A.CustomerID And SchemeID = " & SchemeID & " " & _
        "And TransID = (Select Min(TransID) From LoanTrans C " & _
            " Where C.LoanID = B.LoanID And TransDate >= #" & TransDate & "#)" & _
        " ORder By val(AccNum)"

ElseIf ModuleId = wis_DepositLoans Then
    
    SchemeID = ModuleId Mod 100
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,Balance,A.LoanID as AccID,TransID,TransType," & _
        "Amount,FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From DepositLoanMaster A,DepositLoanTrans B,NameTab C Where A.LoanID = B.LoanID " & _
        "And C.CustomerID = A.CustomerID And DepositType = " & SchemeID & " " & _
        "And TransID = (Select Min(TransID) From DepositLoanTrans C " & _
            " Where C.LoanID = B.LoanID And TransDate >= #" & TransDate & "#)" & _
        " ORder By val(AccNum)"
End If

If gDbTrans.SqlStmt = "" Then
    MsgBox "This facility is not availiable for this account type", _
        vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Screen.MousePointer = vbHourglass

Dim rst As Recordset

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    SetOpeningBalance = SetFirstOpeningBalance
    Exit Function
End If

Dim MaxCount As Integer

Unload frmIntPayble
Load frmIntPayble
With frmIntPayble
    .Caption = "Opening Balance"
    MaxCount = rst.RecordCount + 1
    Call .LoadContorls(MaxCount, 20)
    .Title(0) = GetResourceString(36, 60)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(67)
    .Title(3) = GetResourceString(67)
    .Title(4) = GetResourceString(67)
    .BalanceColoumn = False
    .TotalColoumn = False
    .PutTotal = True
End With

Dim TotalBalance As Currency
Dim Balance As Currency
Dim transType As wisTransactionTypes
Dim Mult As Integer

TotalBalance = 0
MaxCount = 0
While Not rst.EOF
    MaxCount = MaxCount + 1
    
    Balance = rst("Balance")
    transType = rst("TransType")
    
    Mult = IIf(AccountType = Asset, 1, -1)
    Mult = Mult * IIf(transType = wContraDeposit Or transType = wDeposit, -1, 1)
    
    Balance = Balance - rst("Amount") * Mult
    With frmIntPayble
        .AccNum(MaxCount) = rst("AccNum")
        .CustName(MaxCount) = FormatField(rst("CustName"))
        .Amount(MaxCount) = Balance
        TotalBalance = TotalBalance + Balance
    End With
    rst.MoveNext
Wend
    
Screen.MousePointer = vbDefault
With frmIntPayble
    MaxCount = MaxCount + 1
    .CustName(MaxCount) = GetResourceString(286)
    .Amount(MaxCount) = TotalBalance
    .ShowForm
    If .grd.Rows < MaxCount Then Exit Function
End With

Screen.MousePointer = vbHourglass

Dim loopCount As Integer
Dim DiffAmount As Currency
Dim TransID As Integer

TableName = TableName & "Trans"
gDbTrans.BeginTrans

rst.MoveFirst
For loopCount = 1 To MaxCount - 1

    TransID = rst("TransID")
    Balance = rst("Balance")
    transType = rst("TransType")
    
    Mult = IIf(AccountType = Asset, 1, -1)
    Mult = Mult * IIf(transType = wContraDeposit Or transType = wDeposit, -1, 1)
    Balance = Balance - rst("Amount") * Mult
    
    DiffAmount = Balance - frmIntPayble.Amount(loopCount)
    If DiffAmount = 0 Then GoTo NextCount
    
    Balance = frmIntPayble.Amount(loopCount)
    transType = rst("TransType")
    
    If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
        Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
        gDbTrans.SqlStmt = "UpDate " & TableName & _
            " SET Balance = Balance - " & DiffAmount & _
            " Where AccID = " & rst("Accid") & _
            " ANd TransID >= " & rst("TransID") & _
            " And TransDate >= #" & TransDate & "#"
            
    ElseIf ModuleId = wis_BKCC Or ModuleId = wis_Loans _
        Or ModuleId = wis_BKCCLoan Or ModuleId = wis_DepositLoans Then
        
        gDbTrans.SqlStmt = "UpDate " & TableName & _
            " SET Balance = Balance - " & DiffAmount & _
            " Where LoanID = " & rst("Accid") & _
            " And TransID >= " & rst("TransID") & _
            " And TransDate >= #" & TransDate & "#"
    End If
    
    If Not gDbTrans.SQLExecute Then GoTo LastLine

NextCount:

    rst.MoveNext
Next

gDbTrans.CommitTrans
SetOpeningBalance = True
Screen.MousePointer = vbDefault
Exit Function

LastLine:
Screen.MousePointer = vbDefault
gDbTrans.RollBack
MsgBox "Unable update the opening balance ", vbInformation, wis_MESSAGE_TITLE

End Function

Private Function SetFirstOpeningBalance() As Boolean

If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If

Dim TransDate As Date
Dim ModuleId As wisModules
Dim headID As Long
Dim SchemeID As Long
Dim AccountType As wis_AccountType

TransDate = GetSysFormatDate(txtDate)

With cmb
    If .ListIndex < 0 Then
        MsgBox "Please select the account type", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    headID = .ItemData(.ListIndex)
End With
ModuleId = GetModuleIDFromHeadID(headID)

If ModuleId = wis_None Then Exit Function

gDbTrans.SqlStmt = ""

Dim TableName As String
If ModuleId = wis_Members Then
    TableName = "Mem"
ElseIf ModuleId = wis_SBAcc Then
    TableName = "SB"
ElseIf ModuleId = wis_CAAcc Then
    TableName = "CA"
ElseIf ModuleId = wis_RDAcc Then
    TableName = "RD"
ElseIf ModuleId = wis_PDAcc Then
    TableName = "PD"
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    TableName = "BKCC"
ElseIf ModuleId = wis_Loans Then
    TableName = "Loan"
ElseIf ModuleId = wis_DepositLoans Then
    TableName = "DepositLoan"
End If


If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
        Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
    AccountType = Liability
    gDbTrans.SqlStmt = "Select AccNum, A.AccID," & _
        "FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From " & TableName & "Master A, NameTab C " & _
        "Where C.CustomerID = A.CustomerID " & _
        " Order By val(AccNum)"

ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,A.LoanID as AccID," & _
        "FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From BkccMaster A, NameTab C " & _
        "Where C.CustomerID = A.CustomerID " & _
        " ORder By val(AccNum)"

ElseIf ModuleId = wis_Loans Then
    
    SchemeID = ModuleId Mod 100
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,A.LoanID as AccID, " & _
        "FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From LoanMaster A, NameTab C Where " & _
        "C.CustomerID = A.CustomerID And SchemeID = " & SchemeID & " " & _
        " ORder By val(AccNum)"

ElseIf ModuleId = wis_DepositLoans Then
    
    SchemeID = ModuleId Mod 100
    AccountType = Asset
    gDbTrans.SqlStmt = "Select AccNum,A.LoanID as AccID," & _
        "FirstName+ ' ' + MiddleNAme +' '+ LastName as CustName " & _
        "From DepositLoanMaster A,NameTab C Where " & _
        "C.CustomerID = A.CustomerID And DepositType = " & SchemeID & " " & _
        " ORder By val(AccNum)"
End If

If gDbTrans.SqlStmt = "" Then
    MsgBox "This facility is not availiable for this account type", _
        vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Screen.MousePointer = vbHourglass

Dim rst As Recordset
Dim MaxCount As Integer

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

Unload frmIntPayble
Load frmIntPayble
With frmIntPayble
    MaxCount = rst.RecordCount + 1
    Call .LoadContorls(MaxCount, 20)
    .lblTitle.Caption = cmb.Text & " " & GetResourceString(284)
    .Title(0) = GetResourceString(36, 60)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(67)
    .Title(3) = GetResourceString(67)
    .Title(4) = GetResourceString(67)
    .BalanceColoumn = False
    .TotalColoumn = False
    .PutTotal = True
End With

Dim TotalBalance As Currency
Dim Balance As Currency
Dim transType As wisTransactionTypes
Dim Mult As Integer
Dim rstTemp As Recordset


TableName = TableName & "Trans"
TotalBalance = 0
MaxCount = 0

Dim TransID As Integer
Dim Amount  As Currency

While Not rst.EOF
    MaxCount = MaxCount + 1
    Amount = 0
    Balance = 0
    TransID = 0
    If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
            Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
        gDbTrans.SqlStmt = "Select Balance,Amount,TransType,TransID " & _
            " From " & TableName & _
            " Where AccID = " & rst("AccID") & _
            " And TransDate < #" & TransDate & "# Order By TransId Desc"
    
    Else 'If (ModuleId - (ModuleId Mod 100)) = wis_DepositLoans _
            Or ModuleId = wis_BKCC Or (ModuleId - (ModuleId Mod 100)) = wis_Loans Then
        gDbTrans.SqlStmt = "Select Balance,Amount,TransType,TransID " & _
            " From " & TableName & _
            " Where LoanID = " & rst("AccID") & _
            " And TransDate < #" & TransDate & "# Order By TransId Desc"
        
    End If
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) Then
        transType = rstTemp("TransType")
        Mult = IIf(AccountType = Asset, 1, -1)
        Mult = Mult * IIf(transType = wContraDeposit Or transType = wDeposit, -1, 1)
    
        Balance = FormatField(rstTemp("Balance"))
        transType = FormatField(rstTemp("TransType"))
        TransID = FormatField(rstTemp("TransID"))
        Balance = Balance - rstTemp("Amount") * Mult
    End If

    With frmIntPayble
        .AccNum(MaxCount) = rst("AccNum")
        .CustName(MaxCount) = FormatField(rst("CustName"))
        .Amount(MaxCount) = Balance
        .KeyData(MaxCount) = TransID
        TotalBalance = TotalBalance + Balance
    End With
    rst.MoveNext
Wend
    
Screen.MousePointer = vbDefault
With frmIntPayble
    MaxCount = MaxCount + 1
    .CustName(MaxCount) = GetResourceString(286)
    .Amount(MaxCount) = TotalBalance
    .ShowForm
    If .grd.Rows < MaxCount Then Exit Function
End With

Screen.MousePointer = vbHourglass

Dim loopCount As Integer
Dim DiffAmount As Currency


gDbTrans.BeginTrans
TransDate = DateAdd("d", -1, TransDate)

rst.MoveFirst
For loopCount = 1 To MaxCount - 1
    
    
    DiffAmount = Balance - frmIntPayble.Amount(loopCount)
    If DiffAmount = 0 Then GoTo NextCount
    
    With frmIntPayble
        Balance = .Amount(loopCount)
        TransID = .KeyData(loopCount)
    End With
    
    If TransID Then
        If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
                Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
            gDbTrans.SqlStmt = "Select Balance,Amount,TransType,TransID " & _
                " From " & TableName & _
                " Where AccID = " & rst("AccID") & _
                " And TransID = " & TransID
        
        Else
            gDbTrans.SqlStmt = "Select Balance,Amount,TransType,TransID " & _
                " From " & TableName & _
                " Where LoanID = " & rst("AccID") & _
                " And TransID = " & TransID
            
        End If
        
        DiffAmount = 0
        If gDbTrans.Fetch(rstTemp, adOpenDynamic) Then
            
            transType = rstTemp("TransType")
            TransID = rstTemp("TransID")
            Balance = rstTemp("Balance")
            transType = rstTemp("TransType")

            Mult = IIf(AccountType = Asset, 1, -1)
            Mult = Mult * IIf(transType = wContraDeposit Or transType = wDeposit, -1, 1)
            
            Balance = Balance - rstTemp("Amount") * Mult
            DiffAmount = Balance - frmIntPayble.Amount(loopCount)
        End If
        
        If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
            Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
            gDbTrans.SqlStmt = "UpDate " & TableName & _
                " SET Balance = Balance - " & DiffAmount & _
                " Where AccID = " & rst("Accid") & _
                " ANd TransID >= " & rst("TransID") & _
                " And TransDate >= #" & TransDate & "#"
                
        Else
            gDbTrans.SqlStmt = "UpDate " & TableName & _
                " SET Balance = Balance - " & DiffAmount & _
                " Where LoanID = " & rst("Accid") & _
                " And TransID >= " & rst("TransID") & _
                " And TransDate >= #" & DateAdd("d", 1, TransDate) & "#"
        End If
    Else
        
        TransID = 100
        If ModuleId = wis_Members Or ModuleId = wis_SBAcc Or ModuleId = wis_CAAcc _
            Or ModuleId = wis_PDAcc Or ModuleId = wis_RDAcc Then
            transType = wDeposit
            gDbTrans.SqlStmt = "Insert Into " & TableName & _
                " (AccId, TransID,Amount,TransType,TransDate,Balance) " & _
                " VALUES (" & rst("Accid") & ", " & TransID & _
                "," & Balance & "," & transType & "," & _
                "#" & TransDate & "#," & Balance & ")"
        Else
            transType = wWithdraw
            gDbTrans.SqlStmt = "Insert Into " & TableName & _
                " (LoanId, TransID,Amount,TransType,TransDate,Balance) " & _
                " VALUES (" & rst("Accid") & ", " & TransID & _
                "," & Balance & "," & transType & "," & _
                "#" & TransDate & "#," & Balance & ")"
        End If
    End If
    
    If Not gDbTrans.SQLExecute Then GoTo LastLine

NextCount:

    rst.MoveNext
Next

gDbTrans.CommitTrans
SetFirstOpeningBalance = True
Screen.MousePointer = vbDefault
Exit Function

LastLine:
Screen.MousePointer = vbDefault
gDbTrans.RollBack
MsgBox "Unable update the opening balance ", vbInformation, wis_MESSAGE_TITLE


End Function

Private Function UpdateTransactionID() As Boolean
If cmb.ListIndex < 1 Then GoTo EndLine

If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If

Dim TransDate As Date
TransDate = GetSysFormatDate(txtDate.Text)
Dim StrFirst As String
Dim strTable As String

StrFirst = UCase(Left(cmb.Text, 2))
strTable = cmb.Text

Select Case StrFirst
Case "FD", "SB", "RD", "PD", "ME"
    gDbTrans.SqlStmt = "Select TransID,AccID From " & strTable & " " & _
        "WHere TransDate > #" & TransDate & "# Order By TransID Desc"

End Select

Dim rst As ADODB.Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo EndLine


While Not rst.EOF
    rst("TransId") = rst("TransId") + 1
    rst.MoveNext
Wend

EndLine:
Exit Function

ErrLine:
    MsgBox "Error in Update the Trans id of " * cmb.Text, vbInformation, wis_MESSAGE_TITLE



End Function

Private Sub chkRPFormat_Click()
chkSubHeadTotal.Enabled = chkRPFormat.Value
End Sub

Private Sub cmbFont_Click()
    If Len(cmbFont.Text) < 1 Then Exit Sub
    If cmbFont.ListIndex < 0 Then Exit Sub
    lblTaluka.FontName = cmbFont.Text
    lblDistrict.FontName = cmbFont.Text
    lblPinCode.FontName = cmbFont.Text
    lblState.FontName = cmbFont.Text
    lblGuranteer.FontName = cmbFont.Text
    lblFont.FontName = cmbFont.Text
End Sub


Private Sub cmdAdvance_Click()
    Call BankSetup
End Sub

Private Sub cmdApply_Click()
Call WriteSetUp3
End Sub

Private Sub cmdApply4_Click()
Call WriteSetUp4
End Sub

Private Sub cmdApplyExtra_Click()
Dim SetUp As New clsSetup

Call SetUp.WriteSetupValue("Customer", "Taluka", Trim$(txtTaluka.Text))
Call SetUp.WriteSetupValue("Customer", "District", Trim$(txtDistrict.Text))
Call SetUp.WriteSetupValue("Customer", "PINCODE", Trim$(txtPinCode.Text))
Call SetUp.WriteSetupValue("Customer", "State", Trim$(txtState.Text))
Call SetUp.WriteSetupValue("Loan", "NoOfGuaranteers", CStr(CInt(txtGuarnteer.Text)))

If cmbFont.ListCount > 0 Then
    Call SetUp.WriteSetupValue("General", "FontName", cmbFont.Text)
    Call WriteToIniFile("Language", "FontName", cmbFont.Text, App.Path & "\" & constFINYEARFILE)
    gFontName = cmbFont.Text
End If

Set SetUp = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    RaiseEvent WindowClosed
End Sub

Private Sub cmdDate_Click()
With Calendar
    .Left = Left + fra(2).Left + cmdDate.Left + cmdDate.Width
    .Top = Top + fra(2).Top + cmdDate.Top
    If DateValidate(txtDate, "/", True) Then _
        .selDate = txtDate
    .Show 1
    txtDate = .selDate
End With

End Sub

Private Sub cmdLoad_Click()
    Me.MousePointer = vbHourglass
    Call AddNewGroup
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdOk_Click()

Dim frmOpBalance As Object
Dim rst As Recordset
Dim count As Long

If optEndOFDay.Value Then
    gDbTrans.SqlStmt = "SELECT A.AccID, A.MaturityDate " & _
        " From FDMaster A, FDTrans B Where A.AccID = B.AccId" & _
        " AND A.MaturityDate <= #" & gStrDate & "# " & _
        " And (A.ClosedDate is Null Or A.ClosedDate = #1/1/100#) " & _
        " And (A.MaturedOn is Null Or A.MaturedOn = #1/1/100#) " & _
        " ORDER BY val(A.AccNum) "
        
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        Do While Not rst.EOF
            Set frmOpBalance = New frmFdClose
            With frmOpBalance
                .txtDate = FormatField(rst("maturitydate"))
                .AccountId = rst("AccID")
                .optMature = True
            End With
            frmOpBalance.Show vbModal
            Set frmOpBalance = Nothing
            rst.MoveNext
            If Not rst.EOF Then If MsgBox("GO to next account", vbYesNo, wis_MESSAGE_TITLE) = vbNo Then rst.MoveLast: rst.MoveNext
        Loop
        Set rst = Nothing
    End If

ElseIf optPrintOrder Then
    
    frmParentOrder.Show 1

ElseIf optCompareDb Or optCompact Or optBackUp Then
    Call DatabaseFunctions
ElseIf optPdRepair Then
    Call RepairPigmyDeposits
End If


End Sub

Private Sub cmdOptions_Click()
      'API call to activate Keyboard Layout
  Call SHREE_KBD_SETUP(Pass1, Pass2)
End Sub

Private Sub cmdYearEnd_Click()
    
    If MsgBox("Do you want perform year end operation", vbYesNo, _
        wis_MESSAGE_TITLE) = vbNo Then Exit Sub

    MsgBox "This may take saveral minutes to complete" & vbCrLf & _
        "Please close all other application before start this"
    Screen.MousePointer = vbHourglass
    Dim FinClass As New clsFinChange
    FinClass.YearEndFunctions
    Set FinClass = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

If KeyCode <> vbKeyTab Then Exit Sub
If Shift And vbCtrlMask = 0 Then Exit Sub

'Dim CtrlDown
'CtrlDown = (Shift And vbCtrlMask) > 0

Dim ShiftDown
ShiftDown = (Shift And vbShiftMask)

Dim I As Byte
With TabStrip1
    I = .SelectedItem.Index
    If ShiftDown Then
        I = I - 1
        If I = 0 Then I = .Tabs.count
    Else
        I = I + 1
        If I > .Tabs.count Then I = 1
    End If
    .Tabs(I).Selected = True
End With

End Sub

Private Sub Form_Load()

SetKannadaCaption
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

'the operation which will carry on the database
'will perform if database is at loac mmachine
'Get the Server name

If GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "server") = "" Then
    'If database is local then
    optCompact.Enabled = True
    'Check whether compacting (Zipping) exe is there or not
    'optZip.Enabled = CBool(Dir(App.Path & "\PKZip.exe") <> "")
    optBackUp.Enabled = True
Else
    optCompact.Enabled = False
    optBackUp.Enabled = False
End If

optPdRepair.Visible = CBool(gCurrUser.UserPermissions = perOnlyWaves)


With cmbDateFormat
    .Clear
    .Text = ""
    .AddItem "dd/mm/yyyy"
    .AddItem "dd-mm-yyyy"
    .AddItem "dd.mm.yyyy"
    .AddItem "d/m/yyyy"
    .AddItem "d-m-yyyy"
    .AddItem "d.m.yyyy"
    .AddItem "dd/mm/yy"
    .AddItem "dd-mm-yy"
    .AddItem "dd.mm.yy"
    .AddItem "d/m/yy"
    .AddItem "d-m-yy"
    .AddItem "d.m.yy"
End With
If gLangShree Then
    lblFont.Visible = True
    cmbFont.Visible = True
    cmdOptions.Visible = True
    AddBilingualFont cmbFont
    ''Select the FOnt
End If


Call ReadSetUp
TabStrip1.Tabs(1).Selected = True

''IF this the Date of YEAR END, then add the tab for yearr end operations
If DayBeginUSDate > FinUSEndDate Or (Month(DayBeginUSDate) = 3 And Day(DayBeginUSDate) = 31) Then
    TabStrip1.Tabs.Add , , "Year end"
End If

Me.optFarmer.Visible = CBool(GetConfigValue("FarmerType", "True"))

Me.Width = TabStrip1.Width + 400

End Sub

Private Function LoanInterestReceivable() As Boolean

Dim SchemeID As Integer, AsOnDate As Date
Dim rstMaster As Recordset
Dim rstTrans As Recordset
Dim rstIntTrans As Recordset
Dim rstReceivAble As Recordset

If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If

AsOnDate = GetSysFormatDate(txtDate.Text)
With cmb
    If .ListIndex < 0 Then
        MsgBox "Please select the account type", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    SchemeID = .ItemData(.ListIndex)
End With

gDbTrans.SqlStmt = "SELECT A.LoanID,AccNum,LastIntDate," & _
                " IssueDate,TransDate,Balance,TransId," & _
                " Title+' '+FirstName +' '+ MiddleName +' '+LastName As CustName" & _
                " From LoanMaster A,LoanTrans B, NameTab C " & _
                " Where A.SchemeID = " & SchemeID & _
                " And B.LoanID = A.LoanID And C.CustomerID = A.CustomerID" & _
                " AND TransID = (Select MAx(TransID) From LoanTrans E Where" & _
                    " E.LoanID = A.LoanID And TransDate <= #" & AsOnDate & "# )" & _
                " And ClosedDate is NULL ORDER BY val(AccNum)"

If gDbTrans.Fetch(rstMaster, adOpenDynamic) < 1 Then Exit Function


'Get the Loan int Trans
gDbTrans.SqlStmt = "SELECT A.LoanID,TransDate,IntBalance," & _
                " PenalIntBalance,TransID " & _
                " From LoanMaster A,LoanIntTrans B " & _
                " Where A.SchemeID = " & SchemeID & _
                " And A.LoanID = B.LoanID " & _
                " AND TransID = (Select MAx(TransID) From LoanIntTrans C Where" & _
                    " C.LoanID = A.LoanID And TransDate <= #" & AsOnDate & "# )" & _
                " And ClosedDate is NULL ORDER BY val(AccNum)"

If gDbTrans.Fetch(rstIntTrans, adOpenDynamic) < 1 Then Set rstIntTrans = Nothing

'"Loan receivable
gDbTrans.SqlStmt = "SELECT A.LoanID,TransDate,Balance,TransID,Amount" & _
                " From LoanMaster A,LoanIntReceivAble B " & _
                " Where A.SchemeID = " & SchemeID & " And A.LoanID = B.LoanID " & _
                " AND TransID = (Select MAx(TransID) From LoanIntReceivAble C Where" & _
                    " C.LoanID = A.LoanID And TransDate <= #" & AsOnDate & "# )" & _
                " And ClosedDate is NULL ORDER BY val(AccNum)"
If gDbTrans.Fetch(rstReceivAble, adOpenDynamic) < 1 Then Set rstReceivAble = Nothing

With frmIntPayble
    Call .LoadContorls(rstMaster.RecordCount + 1, 20)
    .lblTitle = cmb.Text & " " & GetResourceString(80) & _
                GetResourceString(376, 47)
    .BalanceColoumn = True
    .TotalColoumn = True
    .Title(0) = GetResourceString(36, 60)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(250, 376, 450)
    .Title(3) = GetResourceString(376, 47)
    .Title(4) = GetResourceString(52, 450)
End With


Dim IntBalance As Currency
Dim PenalBalance As Currency
Dim ReceivableBalance
Dim IntAmount As Currency
Dim PenalInt As Currency
Dim LastIntDate As Date
Dim LoanID As Long
Dim TransID As Integer
Dim TotalIntAmount As Currency
Dim TotalBalance As Currency

Dim loanClass As clsLoan
Set loanClass = New clsLoan

Dim count As Integer
count = 1


frmCancel.PicStatus.Visible = True
frmCancel.lblMessage.Visible = True
frmCancel.Show

While Not rstMaster.EOF
    LoanID = rstMaster("LoanID")
    LastIntDate = rstMaster("Issuedate")
    TransID = rstMaster("TransID")
    'Get the Int Trasaction details
    If Not rstIntTrans Is Nothing Then
        rstIntTrans.MoveFirst
        rstIntTrans.Find "LoanID = " & LoanID
        If Not rstIntTrans.EOF Then
            LastIntDate = rstIntTrans("TransDate")
            IntBalance = FormatField(rstIntTrans("IntBalance"))
            If rstIntTrans("Transid") > TransID Then TransID = rstIntTrans("transID")
            'PenalBalance = FormatField(rstIntTrans("PenalIntBalance"))
        End If
    End If
    With loanClass
        IntAmount = .RegularInterest(LoanID, , AsOnDate)
        'PenalInt = .PenalInterest(LoanID, , AsOnDate)
    End With
    If Not rstReceivAble Is Nothing Then
        rstReceivAble.MoveFirst
        rstReceivAble.Find "LoanID = " & LoanID
        If Not rstReceivAble.EOF Then
            LastIntDate = IIf(DateDiff("d", rstReceivAble("TransDate"), LastIntDate) < 0, _
                    rstReceivAble("TransDate"), LastIntDate)
            ReceivableBalance = FormatField(rstReceivAble("Balance"))
            If rstReceivAble("Transid") > TransID Then TransID = rstReceivAble("transID")
        End If
    End If
    
    IntAmount = IntBalance + IntAmount
    With frmIntPayble
        .AccNum(count) = FormatField(rstMaster("accNum"))
        .CustName(count) = FormatField(rstMaster("CustName"))
        .Balance(count) = ReceivableBalance
        .Amount(count) = IntAmount
        .Total(count) = ReceivableBalance + IntAmount
        .KeyData(count) = TransID
        TotalIntAmount = TotalIntAmount + IntAmount
        TotalBalance = TotalBalance + ReceivableBalance
    End With
    
    DoEvents
    If gCancel Then rstMaster.MoveLast
    With frmCancel
        .lblMessage = "Calculationg interest of AccNo: " & rstMaster("AccNum")
        UpdateStatus .PicStatus, count / rstMaster.RecordCount
    End With
    
    
NextCount:
    count = count + 1
    rstMaster.MoveNext
Wend

UpdateStatus frmCancel.PicStatus, 1, True

Set loanClass = Nothing
Unload frmCancel

With frmIntPayble
    .CustName(count) = GetResourceString(286)
    .Balance(count) = TotalBalance
    .Amount(count) = TotalIntAmount
    .Total(count) = TotalBalance + TotalIntAmount
    .KeyData(count) = TransID
        
    TotalBalance = TotalBalance + ReceivableBalance
    .ShowForm
    If .grd.Rows < rstMaster.RecordCount Then GoTo Exit_Line
End With

Dim InTrans As Boolean
Dim Balance As Currency
Dim UserID As Integer
UserID = gCurrUser.UserID

InTrans = gDbTrans.BeginTrans
TotalIntAmount = 0
TotalBalance = 0
rstMaster.MoveFirst
count = 1
With frmCancel
    .lblMessage.Visible = True
    .PicStatus.Visible = True
    .Show
    UpdateStatus .PicStatus, 0, True
End With

While Not rstMaster.EOF
    LoanID = rstMaster("LoanID")
    Balance = FormatField(rstMaster("Balance"))
    If Balance = 0 Then GoTo NextPut
    
    IntAmount = 0
    With frmIntPayble
        IntAmount = .Amount(count)
        TransID = .KeyData(count) + 1
    End With
    If IntAmount <= 0 Then GoTo NextPut
    
    ReceivableBalance = 0
    'Get the Int Trasaction details
    If Not rstReceivAble Is Nothing Then
        rstReceivAble.MoveFirst
        rstReceivAble.Find "LoanID = " & LoanID
        If Not rstReceivAble.EOF Then _
            ReceivableBalance = FormatField(rstReceivAble("Balance"))
    End If
    
    'Now Insert Into the loanTrans
    gDbTrans.SqlStmt = "Insert INTO LoanTrans ( " & _
        " LoanId,TransId,TransDate,Amount, " & _
        " TransType,Balance,UserID ) " & _
        " VALUES ( " & _
        LoanID & "," & TransID & "," & _
        "#" & AsOnDate & "#," & _
        IntAmount & ", " & wContraWithdraw & "," & _
        Balance + IntAmount & "," & UserID & ")"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    'Mark the Int Paid Date
    gDbTrans.SqlStmt = "Update LoanMaster Set LastIntDate = #" & AsOnDate & "#" & _
        " WHERE LoanID = " & LoanID
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

    'Now Insert Into the loan receivable trans
    gDbTrans.SqlStmt = "Insert INTO LoanIntReceivable( " & _
        " LoanId,TransId,TransDate,Amount, " & _
        " TransType,Balance,UserID ) " & _
        " VALUES ( " & _
        LoanID & "," & TransID & "," & _
        "#" & AsOnDate & "#," & _
        IntAmount & ", " & wContraDeposit & "," & _
        ReceivableBalance + IntAmount & "," & UserID & ")"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    TotalIntAmount = TotalIntAmount + IntAmount

    'Make the Interest balance as zero
    gDbTrans.SqlStmt = "Update LoanIntTrans Set IntBalance = 0, " & _
        " PenalIntBalance = 0 WHERE LoanID = " & LoanID & _
        " And TransID = (Select Top 1 TransID From LoanIntTrans " & _
                " Where LoanID = " & LoanID & " ORDEr By TransId desc) ;"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

    DoEvents
    If gCancel Then rstMaster.MoveLast
    With frmCancel
        .lblMessage = "Calculationg interest of AccNo: " & rstMaster("AccNum")
        UpdateStatus .PicStatus, count / rstMaster.RecordCount
    End With
    
NextPut:
    count = count + 1
    rstMaster.MoveNext
    
Wend

UpdateStatus frmCancel.PicStatus, 1, True

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
Dim headID As Long
Dim RecHeadId As Long
Dim headName As String
Dim headNameEnglish As String


cmbEnglish.ListIndex = cmb.ListIndex
headName = cmb.Text
headNameEnglish = cmbEnglish.Text
headID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemberLoan, 0, wis_Loans + SchemeID)

headName = cmb.Text & " " & GetResourceString(376, 47)   ' Loan Interest Receivable
headNameEnglish = cmbEnglish.Text & " " & LoadResourceStringS(376, 47)   ' Loan Interest Receivable
RecHeadId = bankClass.GetHeadIDCreated(headName, headNameEnglish, parLoanIntProv, 0, wis_Loans + SchemeID)

'Now make the transaction to the ledger heads
If Not bankClass.UpdateContraTrans(headID, RecHeadId, TotalIntAmount, AsOnDate) Then _
        GoTo Exit_Line

gDbTrans.CommitTrans
InTrans = False

LoanInterestReceivable = True

Exit_Line:
    gCancel = 0
    If InTrans Then gDbTrans.RollBack
    Debug.Assert Err.Number = 0
    
    If Err Then _
        MsgBox GetResourceString(535) & vbCrLf & Err.Description
    
    'Resume
    On Error Resume Next
    Unload frmCancel
    Err.Clear
    
End Function



Private Function UndoInterestReceivable() As Boolean

Dim SchemeID As Integer
Dim AsOnDate As Date
Dim rst As Recordset
Dim TotalAmount As Currency

If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If
AsOnDate = GetSysFormatDate(txtDate.Text)

With cmb
    If .ListIndex < 0 Then
        MsgBox "Please select the account type", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    SchemeID = .ItemData(.ListIndex)
End With

gDbTrans.SqlStmt = "SELECT Sum(A.Amount) as TotalAmount" & _
    " From LoanTrans A, LoanIntReceivable B " & _
    " Where B.LoanID = A.LoanID And A.LoanID in (Select LoanID From " & _
        " LoanMaster Where SchemeID = " & SchemeID & ")" & _
    " AND A.TransID = (Select Max(TransID) From LoanTrans E Where" & _
        " E.LoanID = A.LoanID) And B.TransID = A.TransID "

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        TotalAmount = FormatField(rst("TotalAmount"))
If TotalAmount = 0 Then
    MsgBox "No receivable interest is provided on this date", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If MsgBox("Are you sure you want delete the Interest provision Rs." & TotalAmount, _
        vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function


Dim InTrans As Boolean
gDbTrans.BeginTrans
InTrans = True

gDbTrans.SqlStmt = "Delete A.* ,B.* " & _
    " From LoanTrans A, LoanIntReceivable B " & _
    " Where B.LoanID = A.LoanID And A.LoanID In (Select LoanID From " & _
        " LoanMaster Where SchemeID = " & SchemeID & ")" & _
    " AND A.TransID = (Select Max(TransID) From LoanTrans E Where" & _
        " E.LoanID = A.LoanID) And B.TransID = A.TransID "

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
Dim headID As Long
Dim RecHeadId As Long
Dim headName As String

Debug.Assert Len(cmb.Text) = 0
'DOUBT IN whether to use getheadid or getindexheadid
headName = cmb.Text
headID = GetIndexHeadID(headName)
If headID = 0 Then headID = GetHeadID(headName, parMemberLoan)

headName = cmb.Text & " " & GetResourceString(376) & _
                    " " & GetResourceString(47) ' Loan Receivable Interest

RecHeadId = GetIndexHeadID(headName)
If RecHeadId = 0 Then RecHeadId = GetHeadID(headName, parLoanIntProv)

'Now make the transaction to the ledger heads
If Not bankClass.UndoContraTrans(headID, RecHeadId, TotalAmount, AsOnDate) Then _
        GoTo Exit_Line
InTrans = False

Exit_Line:
    gCancel = 0
    If InTrans Then gDbTrans.RollBack
    Debug.Assert Err.Number = 0
    
    If Err Then
        MsgBox "Error UndoIntReceivable" & vbCrLf & Err.Description
        MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    End If
    'Resume
    On Error Resume Next
    Unload frmCancel
    Err.Clear
    
End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

optOpBalance.Caption = GetResourceString(284)
optIntReceivable.Caption = GetResourceString(376, 450)
optUndoInt.Caption = GetResourceString(19, 376, 450)

optEndOFDay.Caption = GetResourceString(349)
optAccount.Caption = GetResourceString(260, 36, 253)
optCaste.Caption = GetResourceString(260, 111)
optPlace.Caption = GetResourceString(260, 112)
optMember.Caption = GetResourceString(260, 101)
optDeposit.Caption = GetResourceString(260, 43, 253)
optCustomer.Caption = GetResourceString(260, 205, 253)
optLoanPurpose.Caption = GetResourceString(260, 80, 221)
optFarmer.Caption = GetConfigValue("FarmerTypeName", GetResourceString(378)) _
                    & " " & GetResourceString(253)
        
        
lblImagePath = GetResourceString(415, 409)
lblAccName = GetResourceString(36, 35)
lblObDate = GetResourceString(284, 37)

'fra(2).Visible = False
cmdLoad.Caption = GetResourceString(3)
cmdOK.Caption = GetResourceString(4)
cmdAdvance.Caption = GetResourceString(6)
cmdApply.Caption = GetResourceString(6)
cmdApply4.Caption = GetResourceString(6)

cmdCancel.Caption = GetResourceString(11)

lblDistrict = GetResourceString(132)
lblPinCode = GetResourceString(134)
lblState = GetResourceString(133)
lblGuranteer = GetResourceString(389, 60)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_AddGroup = New clsAddGroup
    gWindowHandle = 0
    
End Sub

Private Sub optAccount_Click()
'fra.Visible = False
End Sub

Private Sub optBackUp_Click()
    'Get the Information about the
    Call GetInfo
End Sub

Private Sub optCaste_Click()
'fra.Visible = False
End Sub

Private Sub optClgBank_Click()

'Fra.Visible = True
Dim rst As Recordset
Dim ClgBankID As Long

gDbTrans.SqlStmt = "Select * from Install Where KeyData = 'ClearingBankID'"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then ClgBankID = FormatField(rst("ValueData"))

gDbTrans.SqlStmt = "Select * from Heads Where Parentid = " & parBankAccount

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    optClgBank.Enabled = False
    Exit Sub
End If
With cmb
    .Clear
    cmbEnglish.Clear
    While Not rst.EOF
        .AddItem FormatField(rst("HeadName"))
        .ItemData(.newIndex) = rst("HeadID")
        cmbEnglish.AddItem FormatField(rst("HeadNameEnglish"))
        cmbEnglish.ItemData(cmbEnglish.newIndex) = rst("HeadID")
        
        If rst("HeadID") = ClgBankID Then ClgBankID = .newIndex

        rst.MoveNext
    Wend
    If ClgBankID < 1000 Then .ListIndex = ClgBankID
End With

End Sub

Private Sub optCompact_Click()
    'Get the Information about the
    Call GetInfo
End Sub

Private Sub optCustomer_Click()
'fra.Visible = False
End Sub

Private Sub optDeposit_Click()
'fra.Visible = False
End Sub

Private Sub optEndOFDay_Click()
    'Get the Information about the
    Call GetInfo
End Sub

Private Sub optIntReceivable_Click()
    
'Fra.Visible = True
Dim rst As Recordset

cmb.Clear
cmbEnglish.Clear
gDbTrans.SqlStmt = "Select * from LoanScheme " 'Where Parentid = " & parMemberLoan
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
  With cmb
    While Not rst.EOF
        .AddItem FormatField(rst("SchemeName"))
        .ItemData(.newIndex) = rst("SchemeID")
        cmbEnglish.AddItem FormatField(rst("SchemeNameEnglish"))
        cmbEnglish.ItemData(cmbEnglish.newIndex) = rst("SchemeID")
        rst.MoveNext
    Wend
  End With
End If

txtDate = FinIndianFromDate

End Sub


Private Sub optOpBalance_Click()
    
'Fra.Visible = True
Dim rst As Recordset

Dim headID As Long
With cmb
    .Clear
    cmbEnglish.Clear
    headID = GetIndexHeadID(GetResourceString(53, 36))
    .AddItem GetResourceString(53, 36)
    .ItemData(.newIndex) = headID
    
    cmbEnglish.AddItem LoadResourceStringS(53, 36)
    cmbEnglish.ItemData(cmbEnglish.newIndex) = headID
End With


gDbTrans.SqlStmt = "Select * from Heads Where Parentid = " & parMemberDeposit
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    With cmb
    While Not rst.EOF
        .AddItem FormatField(rst("HeadName"))
        .ItemData(.newIndex) = rst("HeadID")
        
        cmbEnglish.AddItem FormatField(rst("HeadNameEnglish"))
        cmbEnglish.ItemData(cmbEnglish.newIndex) = rst("HeadID")
        rst.MoveNext
    Wend
    End With
End If

gDbTrans.SqlStmt = "Select * from Heads Where Parentid = " & parMemberLoan
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    With cmb
    While Not rst.EOF
        .AddItem FormatField(rst("HeadName"))
        .ItemData(.newIndex) = rst("HeadID")
        
        cmbEnglish.AddItem FormatField(rst("HeadNameEnglish"))
        cmbEnglish.ItemData(cmbEnglish.newIndex) = rst("HeadID")
        rst.MoveNext
    Wend
    End With
End If

txtDate = FinIndianFromDate

End Sub

Private Sub optPlace_Click()
'fra.Visible = False
End Sub

Private Sub optPrintOrder_Click()
'Fra.Visible = False

End Sub


Private Sub optTransId_Click()

'Fra.Visible = True
Dim rst As Recordset


With cmb
    .Clear
    .AddItem ""
    .AddItem "MemTrans"
    .AddItem "MemIntPayable"
    .AddItem "SBTrans"
    .AddItem "SBPLTrans"
    .AddItem "CATrans"
    .AddItem "CAPLTrans"
    .AddItem "FDTrans"
    .AddItem "FdIntTrans"
    .AddItem "FdIntPayable"
    .AddItem "RDTrans"
    .AddItem "RDIntTrans"
    .AddItem "RDIntPayable"
    .AddItem "PDTrans"
    .AddItem "PDIntTrans"
    .AddItem "PDIntPayable"
    .AddItem "LoanTrans"
    .AddItem "LoanIntTrans"
    .AddItem "LoanIntReceivable"
    .AddItem "BKCCTrans"
    .AddItem "BKCCIntTrans"
    .AddItem "DepositLoanTrans"
    .AddItem "DepositLoanIntTrans"
    .AddItem "FdTrans"

End With

With cmbEnglish
    .Clear
    .AddItem ""
    .AddItem "MemTrans"
    .AddItem "MemIntPayable"
    .AddItem "SBTrans"
    .AddItem "SBPLTrans"
    .AddItem "CATrans"
    .AddItem "CAPLTrans"
    .AddItem "FDTrans"
    .AddItem "FdIntTrans"
    .AddItem "FdIntPayable"
    .AddItem "RDTrans"
    .AddItem "RDIntTrans"
    .AddItem "RDIntPayable"
    .AddItem "PDTrans"
    .AddItem "PDIntTrans"
    .AddItem "PDIntPayable"
    .AddItem "LoanTrans"
    .AddItem "LoanIntTrans"
    .AddItem "LoanIntReceivable"
    .AddItem "BKCCTrans"
    .AddItem "BKCCIntTrans"
    .AddItem "DepositLoanTrans"
    .AddItem "DepositLoanIntTrans"
    .AddItem "FdTrans"

End With

'txtDate = FinIndianFromDate
txtDate = GetIndianDate(DateAdd("d", -1, FinUSFromDate))

End Sub

Private Sub optUndoInt_Click()
Call optIntReceivable_Click
End Sub


Private Sub TabStrip1_Click()
Dim SelIndex As Byte

SelIndex = TabStrip1.SelectedItem.Index
fra(4).Visible = False

Dim I As Byte
Dim MaxI As Byte

MaxI = TabStrip1.Tabs.count
For I = 1 To MaxI
    fra(I - 1).Visible = CBool(I = SelIndex)
Next

fra(SelIndex - 1).ZOrder 0

End Sub


