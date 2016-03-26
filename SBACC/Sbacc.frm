VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmSBAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX-2000   -   SB Account Wizard"
   ClientHeight    =   7935
   ClientLeft      =   1065
   ClientTop       =   720
   ClientWidth     =   8265
   Icon            =   "Sbacc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDepositType 
      Appearance      =   0  'Flat
      Caption         =   ".."
      Height          =   315
      Left            =   4560
      TabIndex        =   109
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6900
      TabIndex        =   35
      Top             =   7290
      Width           =   1215
   End
   Begin VB.Frame fraTransact 
      Height          =   6500
      Left            =   360
      TabIndex        =   70
      Top             =   570
      Width           =   7635
      Begin VB.TextBox txtVoucherNo 
         Height          =   330
         Left            =   5310
         TabIndex        =   30
         Top             =   2220
         Width           =   2115
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   300
         Left            =   2970
         TabIndex        =   16
         Top             =   1230
         Width           =   315
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   2220
         Width           =   2145
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
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
         Left            =   2910
         TabIndex        =   10
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   450
         Left            =   4920
         TabIndex        =   32
         Top             =   5970
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         TabIndex        =   17
         Top             =   1230
         Width           =   1545
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   450
         Left            =   6210
         TabIndex        =   31
         Top             =   5970
         Width           =   1215
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1770
         Width           =   2115
      End
      Begin VB.TextBox txtAccNo 
         Height          =   315
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   9
         Top             =   210
         Width           =   1515
      End
      Begin VB.ComboBox cmbCheque 
         Height          =   315
         Left            =   5310
         TabIndex        =   26
         Top             =   1740
         Width           =   1725
      End
      Begin VB.CommandButton cmdCheque 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   28
         Top             =   1770
         Width           =   315
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   5310
         TabIndex        =   27
         Top             =   1740
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   660
         Width           =   6075
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   5310
         TabIndex        =   23
         Top             =   1230
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraPassBook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2445
         Left            =   240
         TabIndex        =   71
         Top             =   3450
         Width           =   7065
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   1800
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2325
            Left            =   120
            TabIndex        =   72
            Top             =   30
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   5
            AllowUserResizing=   1
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
         Begin VB.CommandButton cmdPrevTrans 
            Height          =   405
            Left            =   6630
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   75
            Width           =   435
         End
         Begin VB.CommandButton cmdNextTrans 
            Height          =   405
            Left            =   6630
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   510
            Width           =   435
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2925
         Left            =   150
         TabIndex        =   33
         Top             =   3000
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5159
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2325
         Left            =   270
         TabIndex        =   73
         Top             =   3510
         Width           =   7035
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
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   90
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2235
            Left            =   60
            TabIndex        =   74
            Top             =   30
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   3942
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Sbacc.frx":000C
         End
      End
      Begin VB.Line Line4 
         X1              =   7530
         X2              =   120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line3 
         X1              =   7500
         X2              =   150
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Label lblVoucher 
         Caption         =   "Voucher No: "
         Height          =   300
         Left            =   3930
         TabIndex        =   29
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label lblDate 
         Caption         =   "Date : "
         Height          =   300
         Left            =   180
         TabIndex        =   15
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lblParticular 
         Caption         =   "Particulars : "
         Height          =   300
         Left            =   180
         TabIndex        =   20
         Top             =   2220
         Width           =   945
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   300
         Left            =   3930
         TabIndex        =   22
         Top             =   1260
         Width           =   1425
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   300
         Left            =   180
         TabIndex        =   18
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblName 
         Caption         =   "Name(s) : "
         Height          =   300
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No. : "
         Height          =   300
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance : Rs. 00.00"
         Height          =   300
         Left            =   4830
         TabIndex        =   13
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label lblInstrNo 
         Caption         =   "Instument no:"
         Height          =   300
         Left            =   3930
         TabIndex        =   25
         Top             =   1770
         Width           =   1425
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7000
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   12356
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraProps 
      Height          =   6500
      Left            =   360
      TabIndex        =   75
      Top             =   570
      Width           =   7635
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Undo Interest"
         Height          =   435
         Left            =   4200
         TabIndex        =   107
         Top             =   6000
         Width           =   1935
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6270
         TabIndex        =   12
         Top             =   5950
         Width           =   1215
      End
      Begin VB.Frame fraInterest 
         Caption         =   "Interest"
         Height          =   3405
         Left            =   90
         TabIndex        =   89
         Top             =   2490
         Width           =   7425
         Begin VB.TextBox txtUndoInterest 
            Height          =   360
            Left            =   2760
            TabIndex        =   7
            Top             =   1980
            Width           =   1425
         End
         Begin VB.CommandButton cmdUndoInterests 
            Caption         =   "Undo interests now"
            Enabled         =   0   'False
            Height          =   450
            Left            =   5280
            TabIndex        =   6
            Top             =   1980
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddInterests 
            Caption         =   "Add interests now"
            Enabled         =   0   'False
            Height          =   450
            Left            =   5280
            TabIndex        =   5
            Top             =   1140
            Width           =   2055
         End
         Begin VB.TextBox txtInterestTo 
            Height          =   330
            Left            =   3960
            TabIndex        =   96
            Top             =   300
            Width           =   1245
         End
         Begin VB.TextBox txtInterestFrom 
            Height          =   330
            Left            =   2130
            TabIndex        =   94
            Top             =   300
            Width           =   1245
         End
         Begin VB.TextBox txtPropInterest 
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
            Left            =   6660
            TabIndex        =   99
            Text            =   "0.00"
            Top             =   300
            Width           =   675
         End
         Begin ComctlLib.ProgressBar prg 
            Height          =   300
            Left            =   240
            TabIndex        =   91
            Top             =   1200
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   529
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label txtFailAccIds 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            TabIndex        =   11
            Top             =   2910
            Width           =   7125
         End
         Begin VB.Line Line1 
            X1              =   7320
            X2              =   90
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblStatus 
            Caption         =   "x"
            Height          =   225
            Left            =   270
            TabIndex        =   100
            Top             =   750
            Width           =   4845
         End
         Begin VB.Label lblFailAccIDs 
            Caption         =   "Accounts where undo was not possible:"
            Height          =   330
            Left            =   150
            TabIndex        =   97
            Top             =   2460
            Width           =   6030
         End
         Begin VB.Label lblUndoInterest 
            Caption         =   "Undo interest added on: "
            Height          =   300
            Left            =   150
            TabIndex        =   95
            Top             =   2010
            Width           =   2325
         End
         Begin VB.Label lblInttoDate 
            Caption         =   "to"
            Height          =   240
            Left            =   3480
            TabIndex        =   93
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblIntFromDate 
            Caption         =   "Add interest from date : "
            Height          =   330
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lblRateofInterest 
            Caption         =   "Rate of interest : "
            Height          =   300
            Left            =   5340
            TabIndex        =   98
            Top             =   330
            Width           =   1245
         End
      End
      Begin VB.Frame fraCharges 
         Caption         =   "Charges"
         Height          =   2385
         Left            =   4080
         TabIndex        =   85
         Top             =   180
         Width           =   3405
         Begin VB.CheckBox chkMinBalInt 
            Caption         =   "No Interest on Minimum Balance"
            Height          =   615
            Left            =   120
            TabIndex        =   106
            Top             =   1560
            Width           =   3135
         End
         Begin VB.TextBox txtPropPassBook 
            Height          =   300
            Left            =   2430
            TabIndex        =   90
            Text            =   "0.00"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtPropCancel 
            Height          =   300
            Left            =   2430
            TabIndex        =   87
            Text            =   "0.00"
            Top             =   450
            Width           =   855
         End
         Begin VB.Label lblDupPassbook 
            Caption         =   "Duplicate pass book : "
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
            Left            =   150
            TabIndex        =   88
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblCancellation 
            Caption         =   "Cancellation : "
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
            Left            =   150
            TabIndex        =   86
            Top             =   450
            Width           =   1635
         End
      End
      Begin VB.Frame fraMinBalance 
         Caption         =   "Minimum balance"
         Height          =   2385
         Left            =   90
         TabIndex        =   76
         Top             =   180
         Width           =   4005
         Begin VB.TextBox txtPropMaxTrans 
            Height          =   330
            Left            =   2970
            TabIndex        =   84
            Text            =   "0.00"
            Top             =   1860
            Width           =   945
         End
         Begin VB.TextBox txtPropBalance 
            Height          =   330
            Left            =   2970
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   1370
            Width           =   945
         End
         Begin VB.TextBox txtPropNoCheque 
            Height          =   330
            Left            =   2970
            TabIndex        =   80
            Text            =   "0.00"
            Top             =   880
            Width           =   945
         End
         Begin VB.TextBox txtPropCheque 
            Height          =   330
            Left            =   2970
            TabIndex        =   78
            Text            =   "0.00"
            Top             =   390
            Width           =   945
         End
         Begin VB.Label lblPropMaxTrans 
            Caption         =   "Max Withdrawls per week :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   83
            Top             =   1860
            Width           =   2925
         End
         Begin VB.Label lblInsuffBal 
            Caption         =   "Insufficient balance : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   81
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label lblWithoutCheque 
            Caption         =   "Without Cheque Book :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   79
            Top             =   900
            Width           =   2355
         End
         Begin VB.Label lblWithCheque 
            Caption         =   "With Cheque Book :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   77
            Top             =   390
            Width           =   2565
         End
      End
   End
   Begin VB.Frame fraNew 
      Height          =   6500
      Left            =   360
      TabIndex        =   38
      Top             =   570
      Width           =   7635
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   450
         Left            =   6300
         TabIndex        =   104
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Clear"
         Height          =   450
         Left            =   6300
         TabIndex        =   4
         Top             =   5445
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6300
         TabIndex        =   2
         Top             =   4375
         Width           =   1215
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   4620
         Left            =   150
         ScaleHeight     =   4560
         ScaleWidth      =   5955
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1305
         Width           =   6015
         Begin VB.VScrollBar VScroll1 
            Height          =   4545
            Left            =   5700
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picSlider 
            Height          =   3645
            Left            =   -45
            ScaleHeight     =   3585
            ScaleWidth      =   5670
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   0
            Width           =   5730
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
               Height          =   270
               Index           =   0
               Left            =   2670
               TabIndex        =   1
               Top             =   30
               Width           =   2970
            End
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Caption         =   "chk"
               Height          =   300
               Index           =   0
               Left            =   2820
               TabIndex        =   24
               Top             =   2070
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2340
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   720
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   4860
               TabIndex        =   67
               Top             =   870
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.Label txtPrompt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Account Holder"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   30
               TabIndex        =   103
               Top             =   30
               Width           =   2625
            End
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   450
         Left            =   6300
         TabIndex        =   3
         Top             =   4910
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   990
         Left            =   150
         ScaleHeight     =   930
         ScaleWidth      =   5955
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   255
         Width           =   6015
         Begin VB.Image imgNewAcc 
            Height          =   375
            Left            =   135
            Stretch         =   -1  'True
            Top             =   120
            Width           =   345
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   990
            TabIndex        =   66
            Top             =   360
            Width           =   4920
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
            Height          =   255
            Left            =   990
            TabIndex        =   65
            Top             =   30
            Width           =   135
         End
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
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
         Left            =   135
         TabIndex        =   69
         Top             =   6060
         Width           =   1545
      End
   End
   Begin VB.Frame fraReports 
      Height          =   6500
      Left            =   360
      TabIndex        =   101
      Top             =   570
      Width           =   7635
      Begin VB.Frame fraReport 
         Caption         =   "Choose a report"
         Height          =   3345
         Left            =   270
         TabIndex        =   102
         Top             =   210
         Width           =   7095
         Begin VB.OptionButton optReports 
            Caption         =   "Cheque book accounts"
            Height          =   315
            Index           =   9
            Left            =   300
            TabIndex        =   43
            Top             =   2550
            Width           =   3375
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Accounts Created"
            Height          =   315
            Index           =   4
            Left            =   4080
            TabIndex        =   44
            Top             =   360
            Width           =   2745
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Show Monthly Balance"
            Height          =   315
            Index           =   8
            Left            =   4080
            TabIndex        =   48
            Top             =   2550
            Width           =   2745
         End
         Begin VB.OptionButton optReports 
            Caption         =   "SB General Ledger "
            Height          =   315
            Index           =   7
            Left            =   300
            TabIndex        =   42
            Top             =   1995
            Width           =   3375
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Joint accounts"
            Height          =   315
            Index           =   0
            Left            =   4080
            TabIndex        =   47
            Top             =   1995
            Width           =   2745
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Balances as on"
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   39
            Top             =   360
            Width           =   3375
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Sub Day Book"
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   40
            Top             =   900
            Width           =   3375
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Sub Cash Book"
            Height          =   315
            Index           =   3
            Left            =   300
            TabIndex        =   41
            Top             =   1455
            Width           =   3375
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Account Closed"
            Height          =   315
            Index           =   5
            Left            =   4080
            TabIndex        =   45
            Top             =   900
            Width           =   2745
         End
         Begin VB.OptionButton optReports 
            Caption         =   "Products && Interests"
            Height          =   315
            Index           =   6
            Left            =   4080
            TabIndex        =   46
            Top             =   1455
            Width           =   2745
         End
      End
      Begin VB.Frame fraOrder 
         Caption         =   " List Order"
         Height          =   1875
         Left            =   270
         TabIndex        =   49
         Top             =   3900
         Width           =   7095
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   450
            Left            =   5760
            TabIndex        =   58
            Top             =   1290
            Width           =   1215
         End
         Begin VB.TextBox txtDate1 
            Height          =   315
            Left            =   1590
            TabIndex        =   54
            Top             =   780
            Width           =   1245
         End
         Begin VB.TextBox txtDate2 
            Height          =   315
            Left            =   5400
            TabIndex        =   57
            Top             =   780
            Width           =   1245
         End
         Begin VB.CommandButton cmdDate1 
            Caption         =   "..."
            Height          =   315
            Left            =   2910
            TabIndex        =   53
            Top             =   780
            Width           =   315
         End
         Begin VB.CommandButton cmdDate2 
            Caption         =   "..."
            Height          =   315
            Left            =   6660
            TabIndex        =   56
            Top             =   780
            Width           =   315
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   315
            Left            =   4080
            TabIndex        =   51
            Top             =   210
            Width           =   2085
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   315
            Left            =   270
            TabIndex        =   50
            Top             =   240
            Value           =   -1  'True
            Width           =   2325
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   225
            Left            =   120
            TabIndex        =   52
            Top             =   810
            Width           =   1395
         End
         Begin VB.Label lblDate2 
            Caption         =   "and before (dd/mm/yyyy)"
            Height          =   315
            Left            =   3390
            TabIndex        =   55
            Top             =   810
            Width           =   2445
         End
         Begin VB.Line Line2 
            X1              =   6990
            X2              =   60
            Y1              =   630
            Y2              =   630
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   450
         Left            =   6180
         TabIndex        =   60
         Top             =   5850
         Width           =   1215
      End
   End
   Begin VB.Label lblDepositTypeName 
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
      Left            =   2760
      TabIndex        =   108
      Top             =   7320
      Width           =   1635
   End
End
Attribute VB_Name = "frmSBAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_FormLoaded As Boolean
Private m_SBHeadId As Long
Private m_AccID As Long
Private m_AccClosed As Boolean
Private m_CustReg As clsCustReg
'Private m_TransID As Long
Private m_rstTrans As Recordset
Private m_NoOfPages As Integer
Private m_DepositType As Integer
Private m_mulipleDeposit As Boolean
Private m_DepositTypeName As String
Private m_DepositTypeNameEnglish As String
        

Private m_JointCustID(3) As Long
Private M_UserPermission As wis_Permissions

Private m_Notes As New clsNotes
Private m_frmCheque As frmCheque
Private m_frmSBJoint As frmJoint

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_clsRepOption As clsRepOption

Const CTL_MARGIN = 15
Private m_accUpdatemode As Integer

'Public Event AccountLoad()
Public Event ShowReport(ReportType As wis_SBReports, _
        ReportOrder As wis_ReportOrder, _
        fromDate As String, toDate As String, _
        RepOption As clsRepOption)
        
Public Event AccountTransaction(transType As wisTransactionTypes, cancel As Integer)
Public Event AccountChanged(AccId As Long)
Public Event DateChanged(IndianDate As String)
Public Event UpdateStatus(strMsg As String)
Public Event WindowClosed()
Public Event SelectDeposit(ByRef DepositType As Integer, ByRef cancel As Boolean)



Private Function AccountDepositInterests(IndianFromDate As String, _
                    IndianToDate As String, IndianTransDate As String, _
                    FailAccIDArr() As Long) As Boolean

Me.Refresh
Dim SqlStr As String
Dim Mn As Integer
Dim Yr As Integer
Dim I As Integer
Dim count As Integer

Dim AccIDs() As Long
Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim rst As ADODB.Recordset
Dim This_date As Date
Dim Interest() As Currency
Dim PresentInterest As Currency
ReDim Interest(0)
Dim Products(0) As Currency
Dim RoundInterest As Boolean  'The interest amount whether has to be round or not
Dim frmInterest As frmIntPayble

'Get the month start and number of months
If WisDateDiff(IndianFromDate, IndianToDate) < 0 Then Exit Function

'GET THE DATE OF TRANSACTION
TransDate = GetSysFormatDate(IndianTransDate)


'Get the interest rate from setup
    Dim l_Setup As New clsSetup
    Dim Rate As Double
    Dim MaxTrans As Integer
    'Call l_Setup.WriteSetupValue("SBAcc" & m_DepositType, "RateOfInterest", txtPropInterest.Text)
    Call SaveInterest(wis_SBAcc, "SBAcc" & m_DepositType, Val(txtPropInterest.Text))
    'Rate = Val(l_Setup.ReadSetupValue("SBAcc", "RateOfInterest", 0))
    
    Rate = GetInterestRate(wis_SBAcc + m_DepositType, "SBAcc")
    If Rate = 0 Then Rate = txtPropInterest
    If Rate < 1 Or Rate >= 100 Then
        'MsgBox "Invalid interest rate specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
        Set l_Setup = Nothing
        Exit Function
    End If
    
    MaxTrans = Val(l_Setup.ReadSetupValue("SBAcc" & m_DepositType, "MaxTransactions", ""))
    Set l_Setup = Nothing

'Loop through all the months till To_Date starting with From_Date
    This_date = GetSysFormatDate(IndianFromDate)

'Initialize the status
    prg.Max = DateDiff("m", This_date, GetSysFormatDate(IndianToDate)) + 1
    prg.Min = 0
    prg.Value = 0
    
    ReDim AccIDs(0)
    'ReDim Products(0)
    Mn = Month(This_date)
    Yr = Year(This_date)
    'Call ComputeSBProducts_New(AccIDs(), Mn, Yr, Products(), MaxTrans, IIf(chkMinBalInt.Value = vbChecked, True, False))
    Call ComputeSBProducts_Daily(AccIDs(), Interest(), This_date, GetSysFormatDate(IndianToDate), Rate, IIf(chkMinBalInt.Value = vbChecked, True, False), m_DepositType)
    'Update the status
    prg.Value = 0 'prg.Value + 1
    
    'lblStatus.Caption = "Computing Interests for " & GetMonthString(Mn) & " ..."
    lblStatus.Caption = GetResourceString(906) & " " & GetMonthString(Mn) & " ..."
    Me.Refresh
    'Dimension the Interest array
    If UBound(Interest) < UBound(Products) Then _
        ReDim Preserve Interest(UBound(AccIDs))
    
    'Decide here abount the Whether to round off the interest or not
    Dim RetMsg As Integer
    'RetMsg = MsgBox("Do you want round off the interest amount", vbYesNoCancel + vbDefaultButton1, wis_MESSAGE_TITLE)
    RetMsg = MsgBox(GetResourceString(806), vbYesNoCancel + vbDefaultButton1, wis_MESSAGE_TITLE)
    If RetMsg = vbNo Then
        RoundInterest = False
    ElseIf RetMsg = vbYes Then
        RoundInterest = True
    Else
        Exit Function
    End If
    
    'Loop through all the products and collect sum of interests
    'For I = 0 To UBound(Products) - 1
    '    If Products(I) > 0 Then
    '        PresentInterest = (Products(I) * 1 * Rate) / (100 * 12)
    '        Interest(I) = Interest(I) + PresentInterest
    '    End If
    'Next I

    'Move to next month
    This_date = DateAdd("m", 1, This_date)
    DoEvents
    If gCancel Then Exit Function
    


    'First you need to get a new transid  and Balance array
    Dim TransID As Long
    Dim Balance As Currency
    Dim TotalIntAmount As Currency
    Dim TotalBalance As Currency
    
    'gDbTrans.SqlStmt = "Select max(TransID)as MaxTransID," & _
        " max(TransDate) as MAxTransDate,AccID from SBTrans Group BY AccID"
    
    gDbTrans.SqlStmt = "Select A.AccId,A.TransID,A.TransDate From SbTrans A " & _
        "inner Join ((Select max(TransID)as MaxTransID, AccID from SBTrans Group BY AccID)as B)" & _
        " On A.AccID=B.AccId and A.TransID=B.MaxTransID "
    If m_DepositType > 0 Then
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where A.AccID in " & _
            "(Select distinct AccID from SBMAster where DepositType = " & m_DepositType & " )"
    End If
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order by A.AccID"
    
    gDbTrans.CreateView ("qrySbMaxDateID")
    
    'Dim sqlStr As String
    SqlStr = "Select A.AccID, A.TransID as TransID," & _
        " A.TransDate as TransDate,Balance from qrySbMaxDateID A " & _
        " Inner Join SbTrans B on A.AccID = B.AccID And A.TransID=B.TransID" & _
        " Order By A.AccID"
    gDbTrans.SqlStmt = SqlStr
    gDbTrans.CreateView ("qrySbAccBalance")
    
     If RoundInterest Then Interest(0) = Interest(0) \ 1
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        For I = 0 To UBound(AccIDs) - 1
            
            rst.Find "AccID = " & AccIDs(I)
            
            If Not (rst.EOF Or rst.BOF) Then
                Balance = CCur(FormatField(rst("Balance")))
                If DateDiff("d", TransDate, rst("TransDate")) > 0 Then Balance = 0
                If Balance > 0 Then  'No interest if balance = 0 i.e if account closed
                    If RoundInterest Then
                        Interest(I) = Interest(I) \ 1
                    Else
                        Interest(I) = ((Interest(I) * 100 \ 1) \ 1) / 100
                    End If
                Else
                    FailAccIDArr(UBound(FailAccIDArr)) = AccIDs(I)
                    ReDim Preserve FailAccIDArr(UBound(FailAccIDArr) + 1)
                End If
            End If
        Next I
    End If
Me.Refresh

'Now Show the Interest to  as form
Set frmInterest = New frmIntPayble
Dim rstMain As Recordset
'qrySbMaxDateID
gDbTrans.SqlStmt = "Select AccNum,TransID,Balance, A.AccID,Title + ' ' " & _
        " + FirstNAme  +' '+ MiddleName  +' '+ LastNAme as CustName " & _
        " From SbMAster A, qrySbAccBalance B, Nametab C " & _
        " WHERE A.AccID = B.AccID " & _
        " AND C.CustomerID = A.CustomerID " & _
        " AND (ClosedDate is NULL OR ClosedDate > #" & TransDate & "#)"
If m_DepositType > 0 Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And A.DepositType = " & m_DepositType
End If
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER By A.AccID"

Call gDbTrans.Fetch(rstMain, adOpenStatic)

Dim rowNo As Integer
Load frmInterest
With frmInterest
    rowNo = 0
    Call .LoadContorls(rstMain.recordCount + 1, 20)
    .lblTitle = GetResourceString(421, 47)
    .Title(0) = GetResourceString(36, 60)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(36)
    .Title(3) = GetResourceString(47)
    .Title(4) = GetResourceString(42)
    
    .PutTotal = True
    .BalanceColoumn = False
    .TotalColoumn = True
    
    I = 0
    count = 0
    rowNo = 1
    While Not rstMain.EOF
        .AccNum(rowNo) = FormatField(rstMain("AccNum"))
        .CustName(rowNo) = FormatField(rstMain("custName"))
        .Balance(rowNo) = FormatCurrency(rstMain("Balance"))
        .Total(rowNo) = FormatCurrency(rstMain("Balance") + Interest(I))
        .KeyData(rowNo) = rstMain("TransID")
        Do
            If AccIDs(count) = rstMain("AccID") Then Exit Do
            count = count + 1
        Loop
        .Amount(rowNo) = FormatCurrency(Interest(count))
        
        TotalIntAmount = TotalIntAmount + Interest(I)
        TotalBalance = TotalBalance + rstMain("Balance")
        
        DoEvents
        If gCancel Then Exit Function
        rowNo = rowNo + 1
        rstMain.MoveNext: I = I + 1
    Wend
    
    .CustName(I) = GetResourceString(286)
    .Balance(I) = TotalBalance
    .Amount(I) = TotalIntAmount
    .Total(I) = TotalBalance + TotalIntAmount
    
    
'Now show the Form
    .ShowForm
    
    If .grd.Rows < 3 Then
        Unload frmInterest
        Set frmInterest = Nothing
        Exit Function
    End If
    
End With
    
Me.Refresh

 
    'Initialize the progress bar
    prg.Value = 0
    prg.Min = 0
    lblStatus.Caption = "Adding Interest to the accounts..."
    lblStatus.Caption = GetResourceString(907)
    
    'Get the balance of the this year
    SqlStr = "SELECT Top 1 BALANCE FROM SBPLTrans " & _
            " ORDER By TransDate Desc,AccID Desc"
    gDbTrans.SqlStmt = SqlStr
    Dim IntBalance As Currency
    IntBalance = 0
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then IntBalance = FormatField(rst("Balance"))

ReDim Interest(0)

'Start inserting interests
gDbTrans.BeginTrans
Dim IntAmount As Currency

TotalBalance = 0
TotalIntAmount = 0

I = 0
count = 0
rstMain.MoveFirst
rowNo = 0
While Not rstMain.EOF
    
    With frmInterest
        rowNo = rowNo + 1
        IntAmount = .Amount(rowNo)
        TransID = .KeyData(rowNo) + 2
    End With
    
    Do
        If AccIDs(count) = rstMain("AccID") Then Exit Do
        count = count + 1
        Debug.Assert count < UBound(AccIDs)
    Loop
    
'For i = 0 To UBound(AccIDs) - 1
    If IntAmount > 0 Then
        TotalIntAmount = TotalIntAmount + IntAmount
        transType = wContraDeposit
        'Initailly deposit that much amount to Sb Trans
        SqlStr = "Insert into SBTrans (AccID, TransID, TransDate, Amount, " & _
            " Balance, Particulars, TransType, ChequeNo,UserId ) values ( " & _
            AccIDs(count) & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            rstMain("Balance") + IntAmount & "," & _
            "'Int upto " & IndianToDate & "'," & _
            transType & ", 0," & gUserID & ")"
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
        'Now with draw that much amount from the sb int trans
        transType = wContraWithdraw
        IntBalance = IntBalance + IntAmount
        SqlStr = "Insert INTO SBPLTrans (AccID, TransID," & _
            " TransDate, Amount, Balance, Particulars, TransType, UserID)" & _
            " VALUES ( " & _
            AccIDs(count) & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            IntBalance & "," & _
            "'Int paid upto " & IndianToDate & "'," & _
            transType & _
            "," & gUserID & ")"
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
    End If
    
    DoEvents
    If gCancel Then Exit Function
    
    rstMain.MoveNext: I = I + 1
Wend

Me.Refresh

Dim IntHeadID As Long
Dim AccHeadName As String
Dim AccHeadNameEnglish As String
Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
AccHeadName = m_DepositTypeName & " " & GetResourceString(487)
AccHeadNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(487)
IntHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepIntPaid, 0, wis_SBAcc + m_DepositType)

Call bankClass.UpdateContraTrans(IntHeadID, m_SBHeadId, TotalIntAmount, TransDate)

gDbTrans.CommitTrans

'MsgBox "Total Accounts deposited " & count, vbInformation, gAppName & " - Message"
MsgBox GetResourceString(622) & count, vbInformation, gAppName & " - Message"

AccountDepositInterests = True

End Function
Private Function AccountDepositInterestsUNDO(IndianTransDate As String, FailAccIDArr() As Long) As Boolean
Dim SqlStr As String
Dim AccIDs() As Long
Dim transType As wisTransactionTypes
Dim ret As Long
Dim I As Long
Dim TransID() As Long
Dim FoundAtLeastOne As Boolean
Dim rst As ADODB.Recordset
Dim USdate As Date
Dim IntAmount As Currency
Dim rstInt As Recordset


USdate = GetSysFormatDate(IndianTransDate)
'First Check interet is deposited to any account or
SqlStr = "SELECT AccID,TransID,Amount From SBPLTRANS WHERE " & _
            " TransDate = #" & USdate & "# ORDER BY ACCID"
If m_DepositType > 0 Then
    SqlStr = "SELECT A.AccID,A.TransID,A.Amount From SBPLTRANS A, SBMAster B WHERE " & _
            " A.AccId= B.AccID and B.DepositType = " & m_DepositType & _
            " And TransDate = #" & USdate & "# ORDER BY A.ACCID"
End If
            
gDbTrans.SqlStmt = SqlStr
ret = gDbTrans.Fetch(rstInt, adOpenStatic)
If ret < 1 Then
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    AccountDepositInterestsUNDO = True
    Exit Function
End If

FoundAtLeastOne = True
ReDim AccIDs(ret)
ReDim TransID(ret)
For I = 0 To ret - 1
    AccIDs(I) = FormatField(rstInt("AccID"))
    TransID(I) = FormatField(rstInt("TransID"))
    IntAmount = IntAmount + FormatField(rstInt("Amount"))
    rstInt.MoveNext
Next I

Screen.MousePointer = vbHourglass
'Now Get the Max transId Of each account
    SqlStr = "SELECT MAX(TransID) as MaxTransId ,AccID From SBTrans" & _
        " Group By AccID ORDER BY AccID"

    gDbTrans.SqlStmt = SqlStr
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
'Check for each account if interest added was the last transaction
I = 0
ReDim FailAccIDArr(0)
While Not rst.EOF
    If AccIDs(I) < rst("AccID") Then GoTo Err_line
    If AccIDs(I) <> rst("AccID") Then GoTo nextRecord
    If TransID(I) <> rst("MaxTransID") Then
        FailAccIDArr(UBound(FailAccIDArr)) = AccIDs(I)
        ReDim Preserve FailAccIDArr(UBound(FailAccIDArr) + 1)
        'now reduced that much amount from inte total interest
        rstInt.MoveFirst
        rstInt.Find ("TransID = " & TransID(I) & " ANd AccID = " & AccIDs(I))
        If Not (rstInt.EOF Or rstInt.BOF) Then
            IntAmount = IntAmount - rstInt("Amount")
        End If
        AccIDs(I) = 0
        TransID(I) = 0
                
    End If
    I = I + 1
    If I >= UBound(AccIDs) Then rst.MoveLast
nextRecord:
    rst.MoveNext
Wend
'Now remove from data base
    Dim count As Long
    gDbTrans.BeginTrans
    'And also show toolbar about progress
    prg.Min = 0
    prg.Max = UBound(AccIDs) - 1
    For I = 0 To UBound(AccIDs) - 1
        prg.Value = I
        If TransID(I) <> 0 Then
            gDbTrans.SqlStmt = "Delete from SBPLTrans where Accid = " & _
                    AccIDs(I) & " and TransID = " & TransID(I)
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
            gDbTrans.SqlStmt = "Delete from SBTrans where Accid = " & _
                    AccIDs(I) & " and TransID = " & TransID(I)
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            count = count + 1
        End If
    Next I
    
    
'If TransType = wContraDeposit Or TransType = wContraWithdraw Then
    Dim IntHeadID As Long
    Dim AccHeadName As String
    Dim AccHeadNameEnglish As String
    Dim bankClass As New clsBankAcc
    AccHeadName = m_DepositTypeName & " " & GetResourceString(487)
    AccHeadNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(487)
    IntHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepIntPaid, 0, wis_SBAcc + m_DepositType)
    
    Call bankClass.UndoContraTrans(IntHeadID, m_SBHeadId, IntAmount, GetSysFormatDate(IndianTransDate))
'End If
    

'Check Failures
    If UBound(FailAccIDArr) >= 1 Then
        'MsgBox "Interests could not be withdrawn from some of the accounts !" & vbCrLf & vbCrLf & _
                "You will have to undo these transactions manually for each account using the Undo Transaction Command", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(624) & vbCrLf & vbCrLf & _
                GetResourceString(625), vbExclamation, gAppName & " - Error"
    End If
   
gDbTrans.CommitTrans
Screen.MousePointer = vbDefault
    'MsgBox "Deposits withdrawn from " & count & " deposits !", vbInformation, gAppName & " - Information"
    MsgBox GetResourceString(626) & count & GetResourceString(271), vbInformation, gAppName & " - Information"
    AccountDepositInterestsUNDO = True

Err_line:
Screen.MousePointer = vbDefault
End Function

Private Function AccountName(AccId As Long) As String

Dim Lret As Long
Dim rst As ADODB.Recordset

'Prelim checks
    If AccId <= 0 Then Exit Function

'Query DB
        gDbTrans.SqlStmt = "SELECT AccID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM SBMaster, " _
                & "NameTab WHERE SBmaster.AccID = " & AccId _
                & " AND SBMaster.CustomerID = NameTab.CustomerID"
        Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
        If Lret = 1 Then
            AccountName = FormatField(rst("Name"))
        ElseIf Lret > 1 Then
            'MsgBox "Data base error !", vbCritical, gAppName & " - Error"
            MsgBox GetResourceString(601), vbCritical, gAppName & " - Error"
            Exit Function
        End If

End Function

Private Function AccountSave() As Boolean

Dim txtIndex As Byte
Dim AccIndex As Byte
Dim count As Integer
Dim rst As ADODB.Recordset
Dim strAccNum As String

Dim SqlStr As String

On Error GoTo SaveAccount_error

Dim JointNo As Integer
    
' Check for a valid Account number.
    AccIndex = GetIndex("AccID")
    With txtData(AccIndex)
    
        'See if acc no has been specified
        strAccNum = Trim$(.Text)
        If strAccNum = "" Then
            'MsgBox "No Account number specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(500), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        
        'Check whetehr this account Num already exists or not
        
        gDbTrans.SqlStmt = "Select AccID from SBMaster where " & _
                " AccNum = " & AddQuotes(strAccNum, True)
        
        If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & m_DepositType
        'if Account is updating then
        If m_accUpdatemode = wis_UPDATE Then _
            gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AccID <> " & m_AccID
        
        If gDbTrans.Fetch(rst, adOpenForwardOnly) >= 1 Then
            'MsgBox "Account number " & .Text & "already exists." & vbCrLf & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(545) & vbCrLf & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
            Exit Function
        End If
    End With

    ' Check for account holder name.
    txtIndex = GetIndex("AccName")
    With txtData(txtIndex)
        If Trim$(.Text) = "" Then
            'MsgBox "Account holder name not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(529), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With

    ' Check for nominee name.
Dim NomineeSpecified As Boolean
txtIndex = GetIndex("NomineeName")
With txtData(txtIndex)
    NomineeSpecified = IIf(Trim$(.Text) = "", False, True)
    'if he is saving the account first time
    'then only ask the for the continuetion
    If m_accUpdatemode = wis_INSERT And Not NomineeSpecified Then
        'MsgBox "Nominee name not specified!", _
                vbExclamation, wis_MESSAGE_TITLE
        If MsgBox(GetResourceString(558) & vbCrLf & GetResourceString(541), _
                vbInformation + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End If
End With

'Check for nominee age.
txtIndex = GetIndex("NomineeAge")
With txtData(txtIndex)
    ' if has specified the name then only prompt for the age
    If Val(.Text) = 0 And NomineeSpecified Then
        'MsgBox "Nominee age not specified!", _
                vbExclamation, wis_MESSAGE_TITLE
       MsgBox GetResourceString(507), _
                vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
    If Trim$(.Text) <> "" And Not IsNumeric(Trim$(.Text)) Then
         'MsgBox "Invalid nominee age specified!", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(507), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
    If (Val(Trim$(.Text)) <= 0 Or Val(Trim$(.Text)) >= 100) And NomineeSpecified Then
        'MsgBox "Invalid nominee age specified!", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(507), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
End With
    
'Check for nominee relationship.
txtIndex = GetIndex("NomineeRelation")
With txtData(txtIndex)
    'check whether nominee name specified or not
    If Trim$(.Text) = "" And NomineeSpecified Then
        'MsgBox "Specify nominee relationship.", vbInformation , wis_MESSAGE_TITLE
        MsgBox GetResourceString(559), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
End With

'Check for the introducer ID
txtIndex = GetIndex("IntroducerID")
With txtData(txtIndex)
    ' Check if an introducerID has been specified.
    If Trim$(.Text) = "" Then
      If m_accUpdatemode = wis_INSERT Then
        'If MsgBox("No introducer has been specified!" _
            & vbCrLf & "Add this Account anyway?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox(GetResourceString(560) _
            & vbCrLf & GetResourceString(663), vbQuestion + vbYesNo) = vbNo Then
            GoTo Exit_Line
        End If
      End If
    Else
        'Check if the introducer exists.
        gDbTrans.SqlStmt = "SELECT AccID FROM SBMaster " & _
                " WHERE AccNum = " & AddQuotes(.Text, True)
        If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
            'MsgBox "The introducer account number " & .Text & " is invalid.", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(514), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        'Check if Accno clash
        If txtData(AccIndex).Text = .Text Then
           'MsgBox "The introducer account number cannot be the same as the account holder!", vbExclamation, wis_MESSAGE_TITLE
           MsgBox GetResourceString(564), vbExclamation, wis_MESSAGE_TITLE
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
'Check for the Account Group
If GetAccGroupID = 0 Then
    'MsgBox "You have not selected the Account Group", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(749), vbInformation, wis_MESSAGE_TITLE
    txtIndex = GetIndex("AccGroup")
    ActivateTextBox txtData(txtIndex)
    Exit Function
End If

'Confirm before proceeding
If m_accUpdatemode = wis_UPDATE Then
    'If MsgBox("This will update the account " & GetVal("AccID") _
            & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
   If MsgBox(GetResourceString(520) & " " & GetVal("AccID") _
            & "." & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
        GoTo Exit_Line
    End If
ElseIf m_accUpdatemode = wis_INSERT Then
    
    'If MsgBox("This will create a new account with an account number " & GetVal("AccID") _
            & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
    If MsgBox(GetResourceString(540) & " " & GetVal("AccID") _
            & "." & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
        GoTo Exit_Line
    End If
End If

' Insert/update to database.
Dim InOperative As Boolean
InOperative = IIf(UCase(GetVal("InOperative")) = "TRUE", True, False)

'Start Transactions to Data base
m_CustReg.ModuleID = wis_SBAcc

'Begin the Transaction Of Database
gDbTrans.BeginTrans

If Not m_CustReg.SaveCustomer Then
    'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
    Exit Function
End If

If m_accUpdatemode = wis_INSERT Then
    'nRet = MsgBox("Add this account to database?", vbQuestion + vbYesNo)
    'If nRet = vbNo Then GoTo exit_line
    ' Build the SQL insert statement.
    '"get the Account ID
    Dim AccId As Long
    AccId = 1
    'Get teh New account ID
    SqlStr = "SELECt Max(AccId) FROM SBMaster"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then AccId = FormatField(rst(0)) + 1
    
'    AccId = GetNewAccountNumber
    SqlStr = "Insert into SBMaster (AccID,AccNum, CustomerID, CreateDate, " _
            & " JointHolder, NomineeName,NomineeAge,NomineeRelation, " _
            & " IntroducerId, LedgerNo, FolioNo,AccGroupID,UserId  "
    SqlStr = SqlStr & IIf(m_DepositType > 0, " ,DepositTYpe ) ", ")")  'ADD CLosing Bracket
    
    'Add Values for parameters Below
    SqlStr = SqlStr & " Values (" & AccId & "," & _
            AddQuotes(strAccNum, True) & "," & _
            m_CustReg.CustomerID & "," & _
            "#" & GetSysFormatDate(GetVal("CreateDate")) & "#," & _
            AddQuotes(Left(GetVal("JointHolder"), 250), True) & ", " & _
            AddQuotes(Left(GetVal("Nomineename"), 25), True) & ", " & _
            Val(GetVal("NomineeAge")) & ", " & _
            AddQuotes(Left(GetVal("NomineeRelation"), 15), True) & ", " & _
            Val(GetVal("IntroducerID")) & "," & _
            AddQuotes(GetVal("LedgerNo"), True) & ", " & _
            AddQuotes(GetVal("FolioNo"), True) & ", " & _
            GetAccGroupID & _
            "," & gUserID & ""
    SqlStr = SqlStr & IIf(m_DepositType > 0, ", " & m_DepositType & ")", ")")   'ADD CLosing Bracket
    ' Insert/update the data.
    
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    ' Insert the data to the joint table
    JointNo = 0
    
    If GetVal("JointHolder") <> "" Then
        If m_frmSBJoint Is Nothing Then Set m_frmSBJoint = New frmJoint
        For count = 0 To m_frmSBJoint.txtName.count - 1
            JointNo = JointNo + 1
            m_JointCustID(count) = m_frmSBJoint.JointCustId(count)
            If m_JointCustID(count) = 0 Then Exit For
            SqlStr = "Insert into SBJoint (AccID,CustomerID, CustomerNum) " _
                & "values (" & AccId & "," & _
                m_JointCustID(count) & "," & _
                JointNo & _
                ")"
            gDbTrans.SqlStmt = SqlStr
            If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        Next
    End If

ElseIf m_accUpdatemode = wis_UPDATE Then
    ' The user has selected updation.
    ' Build the SQL update statement.
    SqlStr = "Update SBMaster set " & _
        " AccNum = " & AddQuotes(strAccNum, True) & "," & _
        " JointHolder = " & AddQuotes(Left(GetVal("JointHolder"), 250), True) & "," & _
        " NomineeName = " & AddQuotes(Left(GetVal("Nomineename"), 25), True) & "," & _
        " NomineeAge = " & Val(GetVal("NomineeAge")) & "," & _
        " NomineeRelation =" & AddQuotes(Left(GetVal("NomineeRelation"), 15), True) & "," & _
        " IntroducerId = " & Val(GetVal("IntroducerID")) & "," & _
        " LedgerNo = " & AddQuotes(GetVal("LedgerNo"), True) & "," & _
        " FolioNo = " & AddQuotes(GetVal("FolioNo"), True) & ", " & _
        " CreateDate = #" & GetSysFormatDate(GetVal("CreateDate")) & "#," & _
        " AccGroupID = " & GetAccGroupID & ", " & _
        " InOperative = " & InOperative
    If m_DepositType > 0 Then SqlStr = SqlStr & ", DepositType = " & m_DepositType
    SqlStr = SqlStr & " Where AccID = " & m_AccID
    
    ' Insert/update the data.
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    JointNo = 0
    If GetVal("JointHolder") <> "" Then
        For count = 0 To 3
            JointNo = JointNo + 1
            If m_JointCustID(count) <> m_frmSBJoint.JointCustId(count) Then
                'If MsgBox("You have changed the joint account holder" & _
                    vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, _
                    wis_MESSAGE_TITLE) = vbNo Then Exit Function
                Debug.Print "Kannada"
                If MsgBox(GetResourceString(675) & vbCrLf & _
                    GetResourceString(541), vbQuestion + vbYesNo, _
                    wis_MESSAGE_TITLE) = vbNo Then Exit Function
                If m_frmSBJoint.JointCustId(count) = 0 Then 'Delte the Joint account details
                    SqlStr = "DELETE * FROM SBJoint WHERE AccID = " & m_AccID & _
                        " AND CustomerNum >= " & JointNo
                    gDbTrans.SqlStmt = SqlStr
                    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
                    
                    m_JointCustID(count) = 0
                End If
            End If
            If m_frmSBJoint.JointCustId(count) = 0 Then Exit For
            If m_JointCustID(count) = 0 Then 'Insert the new record
                SqlStr = "Insert into SBJoint (AccID,CustomerID, CustomerNum) " _
                        & "values (" & m_AccID & "," & _
                        m_frmSBJoint.JointCustId(count) & "," & _
                        JointNo & _
                        ")"
            Else 'Update the existing record
                SqlStr = "UPDATE SBJoint Set CustomerID = " & m_frmSBJoint.JointCustId(count) & _
                    " WHERE AccID = " & m_AccID & " AND CustomerNum = " & JointNo
            End If
            gDbTrans.SqlStmt = SqlStr
            If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        Next
    End If
    
End If
    
'Now Commit the Transaction
gDbTrans.CommitTrans
    
    
    'MsgBox "Saved the account details.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
    AccountSave = True
    
    
Exit_Line:
    Exit Function

SaveAccount_error:
    If Err Then
        'MsgBox "SaveAccount: " & vbCrLf _
                & Err.Description, vbCritical
        MsgBox GetResourceString(519) & vbCrLf _
                & Err.Description, vbCritical
    End If
    GoTo Exit_Line
    
End Function

Private Function AccountUndoLastTransaction() As Boolean


If gCurrUser.UserPermissions And perBankAdmin = 0 Then
    'No permission
    MsgBox GetResourceString(685), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


Dim ClsBank As clsBankAcc

'Prelim check
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
        cmdUndo.Enabled = False
        Exit Function
    End If
    
'Check if account exists
    Dim ClosedON As String
    If Not SBAccountExists(m_AccID, ClosedON) Then
        'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    
    If ClosedON <> "" Then
        'MsgBox "Account has been closed previously. Ths ", vbExclamation, gAppName & " - Error"
        'If MsgBox("Account has been closed previously." & vbCrLf & _
                "This action will reopen the account." & vbCrLf & _
                "Do you want to continue ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
        If MsgBox(GetResourceString(524) & vbCrLf & _
                GetResourceString(548) & vbCrLf & _
                GetResourceString(541), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                Exit Function
        Else  'Reopen the account first
            If Not AccountReopen(m_AccID) Then
                'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
                Exit Function
            End If
        End If
    End If
    
    'Get last transaction record
    Dim Amount As Currency
    Dim ret As Integer
    Dim TransID As Long
    Dim TransDate As Date
    Dim ChequeNo As Long
    Dim transType As wisTransactionTypes
    Dim ConType As String
    Dim rst As ADODB.Recordset
    
    
    gDbTrans.SqlStmt = "Select TOP 1 * from SBTrans where " & _
                " AccID = " & m_AccID & " order by TransID desc"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret >= 1 Then
        cmbTrans.ListIndex = -1
        Amount = FormatField(rst("Amount"))
        TransID = FormatField(rst("TransID"))
        TransDate = rst("TransDate")
        ChequeNo = FormatField(rst("ChequeNo"))
        transType = FormatField(rst("TransType"))
        
    ElseIf ret <= 0 Then
        'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
        MsgBox GetResourceString(551), vbInformation, gAppName & " - Error"
        Exit Function
    End If
    
    'AND GET THE NAME OF THE TRANSACTION YOU ARE ABOUT TO DELETE..
    If transType = wDeposit Or transType = wContraDeposit Then _
            ConType = GetResourceString(271) '"Deposit"
    If transType = wWithdraw Or transType = wContraWithdraw Then _
            ConType = GetResourceString(272) '"WithDraw"
    
'Confirm UNDO
'If MsgBox("Are you sure you want to undo the last" &
'"transaction of Rs." & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(583) & vbCrLf & _
         Amount & " (" & ConType & ")", _
         vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    
If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & m_SBHeadId & _
            " And Accid = " & m_AccID & " And  TransID = " & TransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
                    AccountUndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If
    

'Delete record from Data base
    Dim InTrans As Boolean
    gDbTrans.BeginTrans
    InTrans = True
    
    'If it has transaction in the PL Trans w.r.t this transaction
    'dletete that transaction also
    gDbTrans.SqlStmt = "Delete from SBPLTrans where AccID = " & m_AccID & _
            " AND TransID = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    'Delete From the Sb Trans Table
    gDbTrans.SqlStmt = "Delete from SBTrans where AccID = " & m_AccID & _
            " AND TransID = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    'Prepare the cheque list
    If transType = wWithdraw Or transType = wContraWithdraw Then
        If ChequeNo > 0 Then
            gDbTrans.SqlStmt = "UPDATE ChequeMaster Set Trans = " & wischqIssue & _
                " WHERE ChequeNo = " & ChequeNo & " AND AccID = " & m_AccID & _
                " AND AccHeadID = " & m_SBHeadId
            If Not gDbTrans.SQLExecute Then GoTo ExitLine
            cmbCheque.AddItem ChequeNo
        End If
    End If
    
  Set ClsBank = New clsBankAcc

If transType = wContraDeposit Or transType = wContraWithdraw Then
    Dim IntHeadID As Long
    IntHeadID = GetIndexHeadID(m_DepositTypeName & " " & GetResourceString(487))
    If transType = wContraDeposit Then
        If Not ClsBank.UndoContraTrans(IntHeadID, m_SBHeadId, Amount, TransDate) Then GoTo ExitLine
    ElseIf transType = wContraWithdraw Then
        If Not ClsBank.UndoContraTrans(m_SBHeadId, IntHeadID, Amount, TransDate) Then GoTo ExitLine
    End If
End If

If transType = wDeposit Then
    If Not ClsBank.UndoCashDeposits(m_SBHeadId, Amount, TransDate) Then GoTo ExitLine
ElseIf transType = wWithdraw Then
    If Not ClsBank.UndoCashWithdrawls(m_SBHeadId, Amount, TransDate) Then GoTo ExitLine
End If
    
    
gDbTrans.CommitTrans
InTrans = False

AccountUndoLastTransaction = True

ExitLine:
If InTrans Then gDbTrans.RollBack

End Function

Private Sub ArrangePropSheet()

Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean

Dim Size As Single

' Arrange the Slider panel.
With picSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCount()
    .Height = txtData(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    Size = txtPrompt(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    
    .Height = IIf(.Height > Size, .Height, Size)
    
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
    .Max = picSlider.ScaleHeight - picViewport.ScaleHeight
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
Private Function AccountDelete() As Boolean
Dim AccNo As Long
Dim rst As ADODB.Recordset
AccNo = m_AccID

'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from SBMaster where AccID = " & AccNo
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
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
    gDbTrans.SqlStmt = "Select TOP 1 * from SBTrans where AccID = " & AccNo
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        'MsgBox "You cannot delete an account having transactions !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Delete account from DB
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Delete from SBMaster where AccID = " & AccNo
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans

AccountDelete = True
End Function

Private Sub LoadDepositType(Deptype As Integer)
    
    Dim rstDepType As Recordset
    Dim txtIndex  As Integer
    Dim cmbIndex As Integer

    gDbTrans.SqlStmt = "Select * From DepositTypeTab where ModuleID = " & wis_SBAcc & _
        " AND DepositType = " & Deptype
    
    If gDbTrans.Fetch(rstDepType, adOpenDynamic) > 0 Then
        'Load the Membertype details
        m_DepositTypeName = FormatField(rstDepType("DepositTypeName"))
        m_DepositTypeNameEnglish = FormatField(rstDepType("DepositTypeNameEnglish"))
        lblDepositTypeName.Caption = m_DepositTypeName
        lblDepositTypeName.Tag = m_DepositTypeNameEnglish
        
        'Position the Lable
        lblDepositTypeName.Left = (Me.Width - lblDepositTypeName.Width) / 2 - 100
        cmdDepositType.Left = lblDepositTypeName.Left + lblDepositTypeName.Width + 50
        m_SBHeadId = 0
        'CLEAR all the controls
        ResetUserInterface

        'txtIndex = GetIndex("AccNum")
        'txtData(txtIndex).Text = GetNewAccountNumber
        
        'DepType
        'txtIndex = GetIndex("DepositType")
        'txtData(txtIndex).Text = m_DepositTypeName
        'cmbIndex = ExtractToken(txtPrompt(txtIndex).Tag, "TextIndex")
        'Call SetComboIndex(cmb(cmbIndex), , m_depositType)
        'cmb(cmbIndex).Locked = True
        'Call SetComboIndex(cmbRepMemType, , m_depositType)
        'cmbRepMemType.Locked = True
    Else
        m_DepositTypeName = GetResourceString(421)
        m_DepositTypeNameEnglish = LoadResourceStringS(421)
    End If
    
End Sub

Private Sub LoadSetupValues()
'Load the Setup values
    Dim SetUp As New clsSetup
    Dim defValue As String
    defValue = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithChequeBook", "0.00")
    txtPropCheque.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "MinBalanceWithChequeBook", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithoutChequeBook", "0.00")
    txtPropNoCheque.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "MinBalanceWithoutChequeBook", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "InsufficientBalance", "0.00")
    txtPropBalance.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "InsufficientBalance", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "Cancellation", "0.00")
    txtPropCancel.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "Cancellation", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "DuplicatePassBook", "0.00")
    txtPropPassBook.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "DuplicatePassBook", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "RateOfInterest", "0.00")
    txtPropInterest.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "RateOfInterest", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "InterestFrom", "")
    txtInterestFrom.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "InterestFrom", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "MaxTransactions", "")
    txtPropMaxTrans.Text = SetUp.ReadSetupValue("SBAcc" & m_DepositType, "MaxTransactions", defValue)
    defValue = SetUp.ReadSetupValue("SBAcc", "NoInterestOnMinBalance", "False")
    chkMinBalInt.Value = IIf(CBool(SetUp.ReadSetupValue("SBAcc" & m_DepositType, "NoInterestOnMinBalance", defValue)), vbChecked, vbUnchecked)
End Sub

Public Property Let MultipleDeposit(NewValue As Boolean)
    m_mulipleDeposit = NewValue
    lblDepositTypeName.Visible = NewValue
    cmdDepositType.Visible = NewValue
End Property
Public Property Get IsFormLoaded() As Boolean
    IsFormLoaded = m_FormLoaded
End Property
Public Property Let DepositType(NewValue As Integer)
    m_DepositType = NewValue
    If m_FormLoaded Then
        lblDepositTypeName.Visible = NewValue
        cmdDepositType.Visible = NewValue
        LoadDepositType (NewValue)
        Call LoadSetupValues
    End If
    
End Property

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
Private Function GetNewAccountNumber() As Long
Dim NewAccNo As Long
Dim rst As ADODB.Recordset
    
    gDbTrans.SqlStmt = "SELECT MAX(val(AccNum)) FROM SBMaster"
    If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt + " Where DepositType = " & m_DepositType
    If gDbTrans.Fetch(rst, adOpenForwardOnly) = 0 Then
        NewAccNo = 1
    Else
        NewAccNo = Val(FormatField(rst(0))) + 1
    End If
    GetNewAccountNumber = NewAccNo
End Function
Private Function AccountTransaction() As Boolean

Dim AccountCloseFlag As Boolean
Dim ClosedON As String
Dim TransDate As Date
Dim Trans As wisTransactionTypes
Dim ChequeNo As Long

Dim ParticularsArr() As String
Dim I As Integer
Dim rst As ADODB.Recordset

Dim SqlStr As String

'Prelim check
If m_AccID <= 0 Then
    'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
    cmdUndo.Enabled = False
    Exit Function
End If

'Check if account exists
If Not SBAccountExists(m_AccID, ClosedON) Then
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
If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
    'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If
TransDate = GetSysFormatDate(txtDate.Text)

'Check if the date of transaction is earlier than account opening date itself
gDbTrans.SqlStmt = "Select * from SBMaster where AccID = " & m_AccID

If gDbTrans.Fetch(rst, adOpenForwardOnly) <> 1 Then
    'MsgBox "DB error !", vbCritical, gAppName & " - ERROR"
    MsgBox GetResourceString(601), vbCritical, gAppName & " - ERROR"
    Exit Function
End If

If DateDiff("d", TransDate, rst("CreateDate")) > 0 Then
    'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(568), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If

'Get the Transaction Type
If cmbTrans.ListIndex = -1 Then
     'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(588), vbExclamation, gAppName & " - Error"
    cmbTrans.SetFocus
    Exit Function
End If
'Validate the Amount
If txtAmount = 0 Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    'ActivateTextBox txtAmount
    txtAmount.SetFocus
    Exit Function
End If

Dim lstIndex As Byte
With cmbTrans
    lstIndex = .ListIndex
    If .ListIndex = 0 Then Trans = wDeposit
    If .ListIndex = 1 Then Trans = wWithdraw
    If .ListIndex = 2 Then Trans = wContraWithdraw
    If .ListIndex = 3 Then Trans = wContraDeposit
End With

'Validate the Cheque No
Dim IntBalance As Currency
'If Trans = wWithDraw Then

If lstIndex = 1 Then
    With cmbCheque
        If .ListIndex <= 0 Then
            For I = 1 To .ListCount - 1
                If Val(.Text) = Val(.List(I)) Then
                    cmbCheque.ListIndex = I
                    ChequeNo = .ItemData(I)
                    Exit For
                End If
            Next I
        End If
        
        If .ListIndex < 0 Then
            'See if cheque book has been issued to him
            If .ListCount > 1 Then
                'If MsgBox("This account has a cheque book but the specified cheque leaf " & cmbCheque.Text & _
                            " does not belong to the cheque book." & vbCrLf & _
                            "Do you want to continue with the transaction ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                If MsgBox(GetResourceString(628) & cmbCheque.Text & _
                    GetResourceString(629) & vbCrLf & _
                    GetResourceString(541), vbQuestion + vbYesNo, _
                    gAppName & " - Error") = vbNo Then Exit Function
            End If
        End If
    End With
Else 'If Trans = wDeposit Or Trans = wContraDeposit Then
    cmbCheque.ListIndex = -1
End If

'Get Voucher NO
If Trim$(txtVoucherNo.Text) = "" Then
    '"Voucher' Number not specified !", "please specify the voucher no"
    MsgBox GetResourceString(827) & vbCrLf & GetResourceString(828), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtVoucherNo
    Exit Function
End If
    
'Get the Particulars
If Trim$(cmbParticulars.Text) = "" Then
    'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(621), vbExclamation, gAppName & " - Error"
    cmbParticulars.SetFocus
    Exit Function
End If

'cHECK THE LAST dATE OF TRANSACTION
SqlStr = "Select TOP 1 * from SBTrans where AccID = " & m_AccID & _
    " order by TransID DESC"
'Get the Balance and new transid
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    'See if the date is earlier than last date of transaction
    If DateDiff("D", rst("TransDate"), TransDate) < 0 Then
        'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    End If
End If

'CHECK THE LAST dATE OF TRANSACTION
SqlStr = "Select TOP 1 * from SBPLTrans where AccID = " & m_AccID & _
    " order by TransID DESC"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    'See if the date is earlier than last date of transaction
    If DateDiff("D", rst("TransDate"), TransDate) < 0 Then
        'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    End If
End If

Dim cancel As Integer
RaiseEvent AccountTransaction(Trans, cancel)

If cancel Then Exit Function

AccountTransaction = True

'Update Cheque table
 If ChequeNo <> 0 And lstIndex = 1 Then
    With cmbCheque
        If .ListIndex >= 0 Then .RemoveItem .ListIndex
    End With
 End If

'Update the Particulars combo    'Read to part array
Call WriteParticularstoFile(cmbParticulars.Text, App.Path & "\SBAcc.ini")
Call LoadParticularsFromFile(cmbParticulars, App.Path & "\SbAcc.ini")

AccountTransaction = True

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
Private Sub LoadPreviousTransaction(TransID As Integer)

If m_AccID = 0 Then Exit Sub
Dim rstDetail As Recordset

gDbTrans.SqlStmt = "select * from SbTrans " & _
    "Where AccID = " & m_AccID & " AND TransID = " & TransID
    
If gDbTrans.Fetch(rstDetail, adOpenDynamic) < 1 Then Exit Sub

Dim transType As wisTransactionTypes

transType = FormatField(rstDetail("TransType"))

LoadLine:

cmdAccept.Caption = GetResourceString(171)
cmdUndo.Enabled = False
cmdDate.Enabled = False
With txtDate
    .Text = FormatField(rstDetail("TransDate"))
    .Locked = True
End With

With cmbTrans
    If transType = wDeposit Then .ListIndex = 0
    If transType = wWithdraw Then .ListIndex = 1
    If transType = wContraDeposit Then .ListIndex = 2
    If transType = wContraWithdraw Then .ListIndex = 3
    .Locked = True
End With
cmbParticulars.Text = FormatField(rstDetail("Particulars"))
txtAmount.Value = FormatField(rstDetail("Amount"))
txtVoucherNo = FormatField(rstDetail("VoucherNo"))
txtCheque = FormatField(rstDetail("ChequeNo"))
End Sub

Private Function PassBookPageInitialize()
    With grd
        If .Cols < 3 Then .Cols = 6
        .Clear: .Rows = 11: .FixedRows = 1: .FixedCols = 0
        .Row = 0
        .Col = 0: grd.Text = GetResourceString(37): .ColWidth(0) = 1200       ' "Date"
        .Col = 1: grd.Text = GetResourceString(39): .ColWidth(1) = 850       '"Particulars"
        .Col = 2: grd.Text = GetResourceString(275): .ColWidth(2) = 600       '"Cheque"
        .Col = 2: grd.Text = GetResourceString(41): .ColWidth(2) = 600       '"Voucher"
        .Col = 3: grd.Text = GetResourceString(276): .ColWidth(3) = 1100       '"Debit"
        .Col = 4: grd.Text = GetResourceString(277): .ColWidth(4) = 1100       '"Credit"
        .Col = 5: grd.Text = GetResourceString(42): .ColWidth(5) = 1100       '"Balance"
    End With
    
End Function

Private Function LoadPropSheet() As Boolean

TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True
lblDesc.BorderStyle = 0
lblHeading.BorderStyle = 0
lblOperation.Caption = GetResourceString(54) '"Operation Mode : <INSERT>"
'
' Read the data from SBAcc.PRP and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\SBAcc_" & gLangOffSet & ".PRP"
If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet Then
        PropFile = App.Path & "\SBAccKan.PRP"
    Else
        PropFile = App.Path & "\SBAcc.PRP"
    End If
End If
If Dir(PropFile, vbNormal) = "" Then
    'MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox GetResourceString(602) _
            & PropFile & "' !", vbExclamation
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
    strTag = ReadFromIniFile("Property Sheet", _
                "Prop" & I + 1, PropFile)
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
            MsgBox GetResourceString(603), vbCritical
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
    Dim ChkLoaded As Boolean
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
                .Height = .Width
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
                    .Caption = GetResourceString(294) ' "Reset"
                    .Width = 1000
                ElseIf I = 2 Then
                    .Caption = GetResourceString(295) '"Details..."
                    .Width = 1000
                Else
                    .Caption = "Specify..."
                    .Width = 350
                End If
            End With
        Case "BOOLEAN"
              ' Load a check box.
            If Not ChkLoaded Then
                ChkLoaded = True
            Else
                Load chk(chk.count)
            End If
            With chk(chk.count - 1)
                txtData(I).Text = "False"
                .Visible = True
                .Left = txtData(I).Left
                .Top = txtData(I).Top + CTL_MARGIN
                .Width = txtData(I).Width
                .Height = txtData(I).Height - 2 * CTL_MARGIN
                .Caption = String(txtData(I).Width / Me.TextWidth(" "), " ")
                .TabIndex = txtData(I).TabIndex + 1
                .ZOrder 0
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                        "TextIndex", CStr(chk.count - 1))
            End With
    End Select

    ' Increment the loop count.
    I = I + 1
Loop

ArrangePropSheet

' Get a new account number and display it to accno textbox.
Dim txtIndex As Integer
txtIndex = GetIndex("AccID")
txtData(txtIndex).Text = GetNewAccountNumber

' Show the current date wherever necessary.
txtIndex = GetIndex("CreateDate")
txtData(txtIndex).Text = gStrDate

' Set the default updation mode.
m_accUpdatemode = wis_INSERT

End Function
Private Sub ResetUserInterface()

'Get the Sb Head ID in  Heads AccTrans Table
Dim ClsBank As clsBankAcc

If m_SBHeadId = 0 And m_DepositType > -1 Then
    Set ClsBank = New clsBankAcc
    
    gDbTrans.BeginTrans
    m_SBHeadId = ClsBank.GetHeadIDCreated(m_DepositTypeName, _
            m_DepositTypeNameEnglish, parMemberDeposit, 0, wis_SBAcc + m_DepositType)
    gDbTrans.CommitTrans
    
    Set ClsBank = Nothing
End If

'First the TAB 1

    On Error Resume Next
    If Not m_frmSBJoint Is Nothing Then
        Unload m_frmSBJoint
        Set m_frmSBJoint = Nothing
    End If

'If m_Accid = 0 Then Exit Sub
If m_AccID = 0 And m_CustReg.CustomerID = 0 Then Exit Sub
    'Disable the UI if you are unable to load the specified account number
    lblBalance.Caption = ""
    
    With cmbAccNames
        .BackColor = wisGray: .Enabled = False: .Clear
    End With
    
    With txtDate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    
    With cmdDate
        .Enabled = False
    End With
    
    With cmdPrint
        .Enabled = False
    End With
    With cmdUndo
        .Enabled = False
    End With
    With txtAmount
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With cmbTrans
        .BackColor = wisGray: .Enabled = False
    End With
    With cmbCheque
        .BackColor = wisGray: .Enabled = False: .Clear
    End With
    With txtCheque
        .BackColor = wisGray: .Enabled = False
    End With
    With cmdCheque
        .BackColor = wisGray: .Enabled = False
    End With
    With txtVoucherNo
        .BackColor = wisGray: .Enabled = False
    End With
    With cmbParticulars
        .BackColor = wisGray: .Enabled = False
    End With
    With Me.rtfNote
        .BackColor = wisGray: .Enabled = False
        .Text = GetResourceString(259)   '"< No notes defined >"
        If gLangOffSet Then
            .Font.name = gFontName: .Font.Size = gFontSize
        Else
            .Font.Size = 10: .Font = "Arial"
        End If
    End With
    With cmdAccept
        .Enabled = False
        .Caption = GetResourceString(4)
    End With
    With cmdUndo
        .Enabled = False
    End With
    
    Call PassBookPageInitialize
    
    cmdAddNote.Enabled = False
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    Set m_frmSBJoint = Nothing
    
'Now the Tab 2
    Dim I As Integer
    Dim strField As String
    Dim txtIndex As Integer
    
    'Enable the reset (auto acc no generator button)
    cmd(0).Enabled = True

    For I = 0 To txtData.count - 1
        txtData(I).Text = ""
        ' If its Createdate field, then put today's left.
        strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
        If StrComp(strField, "CreateDate", vbTextCompare) = 0 Then
            txtData(I).Text = gStrDate
        End If
    Next
    lblOperation.Caption = GetResourceString(54)    '"Operation Mode : <INSERT>"
    
    txtIndex = GetIndex("AccID")
    txtData(txtIndex).Text = GetNewAccountNumber
    txtData(txtIndex).Locked = False
    'cmdTerminate.Enabled = False
    cmdDelete.Enabled = False
    
    'Now rest the combo box and Check box
    For I = 0 To cmb.count - 1
        cmb(I).ListIndex = -1
    Next
    For I = 0 To chk.count - 1
        chk(I).Value = vbUnchecked
    Next
'The form level variables
    m_accUpdatemode = wis_INSERT
    m_CustReg.NewCustomer
    m_AccID = 0
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
'
Private Sub PassBookPageShow()
Const RecorsdToShow = 10

' If no recordset, exit.
If m_rstTrans Is Nothing Then Exit Sub
' If no records, exit.
If m_rstTrans.recordCount = 0 Then Exit Sub
Dim transType As wisTransactionTypes
Dim TransID  As Long

TransID = m_rstTrans("TransID")
cmdPrevTrans.Enabled = False
If TransID > 10 Then cmdPrevTrans.Enabled = True

'Show 10 records or till eof of the page being pointed to
With grd
    Call PassBookPageInitialize
    .Visible = False
    .Row = 1
    Do
        .RowData(.Row) = m_rstTrans("TransId")
        .Col = 0: .Text = FormatField(m_rstTrans("TransDate"))
        .Col = 1: .Text = FormatField(m_rstTrans("Particulars"))
        .Col = 2: .Text = FormatField(m_rstTrans("ChequeNo"))
        If Trim(.Text) = "" Then .Text = FormatField(m_rstTrans("VoucherNo"))
        transType = m_rstTrans("TransType")
        .Col = IIf(transType = wWithdraw Or transType = wContraWithdraw, 4, 3)
        .Text = FormatField(m_rstTrans("Amount"))
        .Col = 5: .Text = FormatField(m_rstTrans("Balance"))
nextRecord:
        m_rstTrans.MoveNext
        If m_rstTrans.EOF Then Exit Do
        If .Row = .Rows - 1 Then Exit Do
        .Row = .Row + 1
    Loop
    .Visible = True
    .Row = 1
End With
    
cmdNextTrans.Enabled = Not m_rstTrans.EOF
cmdUndo.Enabled = gCurrUser.IsAdmin

If m_rstTrans.recordCount < 10 Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
End If
End Sub



Private Function UpdateTransaction() As Boolean
Dim TransID As Integer

TransID = Val(cmbTrans.Tag)

If TransID = 0 Then Exit Function
UpdateTransaction = False

Dim rstDetail As Recordset
gDbTrans.SqlStmt = "Select * From SbTrans " & _
    " WHERE AccID = " & m_AccID & " AND TransID = " & TransID

If gDbTrans.Fetch(rstDetail, adOpenDynamic) <> 1 Then GoTo ExitLine
If rstDetail("Amount") = txtAmount.Value Then GoTo ExitLine

Dim transType As wisTransactionTypes
Dim AccHeadID As Long

transType = rstDetail("TransType")

AccHeadID = wis_CashHeadID
If transType = wContraDeposit Or transType = wContraWithdraw Then
    'check the transction with interesdt trans
    Dim rstInt As Recordset
    gDbTrans.SqlStmt = "Select * From SbPLTrans " & _
        "Where AccID = " & m_AccID & " AND TransID = " & TransID
    
    If gDbTrans.Fetch(rstInt, adOpenDynamic) < 1 Then
        MsgBox "This Transaction can not be updated"
        GoTo ExitLine
    End If
    AccHeadID = GetIndexHeadID(m_DepositTypeName & " " & GetResourceString(487))
    
'if transtype=wContraDeposit t
End If

Dim PrevAmount As Currency
Dim Amount As Currency
Dim DiffAmount As Currency
Dim TransDate As Date

TransDate = rstDetail("TransDate")
PrevAmount = rstDetail("Amount")
Amount = txtAmount.Value

If transType = wDeposit Or transType = wContraDeposit Then
    DiffAmount = Amount - PrevAmount
Else
    DiffAmount = PrevAmount - Amount
End If

gDbTrans.BeginTrans

gDbTrans.SqlStmt = "UPDate SBTrans Set Amount = " & Amount & _
        " Where AccID = " & m_AccID & " AND TransID = " & TransID
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

gDbTrans.SqlStmt = "UPDate SBTrans Set Balance = Balance + " & DiffAmount & _
        " Where AccID = " & m_AccID & " AND TransID >= " & TransID
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

If transType = wContraDeposit Or transType = wContraWithdraw Then
    gDbTrans.SqlStmt = "UPDate SBPLTrans Set Amount = " & Amount & _
            " Where AccID = " & m_AccID & " AND TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo ErrLine
    End If
End If

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
Dim LedgerEntrySuccess As Boolean
LedgerEntrySuccess = False

If transType = wDeposit Then
    LedgerEntrySuccess = bankClass.UpdateCashDeposits(m_SBHeadId, Amount - PrevAmount, TransDate)
ElseIf transType = wWithdraw Then
    'Call BankClass.UndoCashWithdrawls(m_SBHeadId, Amount - PrvAmount, TransDate)
    'This is corrected By Shashi Angadi on Jun 17 2013
    LedgerEntrySuccess = bankClass.UndoCashWithdrawls(m_SBHeadId, PrevAmount - Amount, TransDate)
ElseIf transType = wContraDeposit Then
    LedgerEntrySuccess = bankClass.UpdateContraTrans(AccHeadID, m_SBHeadId, Amount - PrevAmount, TransDate)
Else
    'Call BankClass.UpdateContraTrans(m_SBHeadId, AccHeadID, Amount - PrevAmount, TransDate)
    'This is corrected By Shashi Angadi on Jun 17 2013
    LedgerEntrySuccess = bankClass.UpdateContraTrans(m_SBHeadId, AccHeadID, PrevAmount - Amount, TransDate)
End If

''Extra check Added By Shashi on 17 Jun 2013
If LedgerEntrySuccess = True Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If
    

ExitLine:
UpdateTransaction = True

ErrLine:
    If Err Then MsgBox "ERROR in UpdateTransaction"


Set bankClass = Nothing
End Function

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


Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub chk_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(chk(Index).Tag, "TextIndex")
If txtIndex <> "" Then
    txtData(Val(txtIndex)).Text = IIf(chk(Index).Value = vbChecked, True, False)
End If
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


If cmbTrans.ListIndex = 0 Then  'A case of deposit
    txtCheque.Visible = True
    cmbCheque.Visible = False
    cmdCheque.Visible = False
    Exit Sub
End If

If cmbTrans.ListIndex = 1 Then  'A case of withdraw
    txtCheque.Visible = False
    cmbCheque.Visible = True
    cmdCheque.Visible = True
    Exit Sub
End If
    


End Sub


Private Sub cmd_Click(Index As Integer)
Screen.MousePointer = vbHourglass
Dim txtIndex As String
Dim count As Integer
Dim rst As ADODB.Recordset

' Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")
Screen.MousePointer = vbDefault
Select Case UCase$(strField)
    Case "ACCID"
        If m_accUpdatemode = wis_INSERT Then txtData(txtIndex).Text = GetNewAccountNumber

    Case "ACCNAME"
        If m_CustReg Is Nothing Then
            Set m_CustReg = New clsCustReg
            m_CustReg.NewCustomer
            m_CustReg.ModuleID = wis_SBAcc
        End If
        m_CustReg.ShowDialog
        txtData(txtIndex).Text = m_CustReg.FullName
    Case "JOINTHOLDER"
        'IF HE has not enter the customer detaisl then
        'do not allow him to enter the joint account holders
        If Trim(m_CustReg.FullName) = "" Then Exit Sub
        
        Dim strJointID As String
        If m_frmSBJoint Is Nothing Then
            Set m_frmSBJoint = New frmJoint
            m_frmSBJoint.ModuleID = wis_SBAcc
        End If
        m_frmSBJoint.ModuleID = wis_SBAcc
        With m_frmSBJoint
            .Left = Me.Left + picViewport.Left + _
                txtData(txtIndex).Left + fraNew.Left + CTL_MARGIN
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fraNew.Top + 300
            Screen.MousePointer = vbDefault
            .Show vbModal
            
            If .Status = "OK" Then txtData(txtIndex).Text = .JointHolders
        End With

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
    
    Case "INTRODUCERNAME"
        Screen.MousePointer = vbDefault
        Dim AccNum As String
        AccNum = GetIndex("AccID")
        If m_CustReg.CustomerID = 0 Then Exit Sub
        
        Dim NameStr  As String
        Dim Lret As Long
        'now Check Whether He Want Search   any particular name
        'NameStr = InputBox("Enter customer name , You want search", "Name Search")
        NameStr = InputBox(GetResourceString(785), "Name Search")
        
        Screen.MousePointer = vbHourglass
        'if Has not select any name then exit the sub
        If NameStr = "" Then Screen.MousePointer = vbDefault: Exit Sub
        
        ' Build a query for getting introducer details.
        ' If an account number specified, exclude it from the list.
        gDbTrans.SqlStmt = "SELECT CustomerID, Title + FirstName + " _
            & "Space(1) + Middlename + space(1) + LastName as Name, " _
            & " HomeAddress, OfficeAddress FROM NameTab " _
            & " WHERE CustomerID <> " & m_CustReg.CustomerID _
            & " And ( FirstName like '" & NameStr & "%' " _
            & " Or MiddleName like '" & NameStr & "%' " _
            & " Or LastName like '" & NameStr & "%' )" _
            & " Order By IsciName"

        Lret = gDbTrans.Fetch(rst, adOpenStatic)
        If Lret <= 0 Then
           'MsgBox "No Customers present!", vbExclamation, wis_MESSAGE_TITLE
           MsgBox GetResourceString(525), vbExclamation, wis_MESSAGE_TITLE
           Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        'Fill the details to report dialog and display it.
        If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
        If Not FillView(m_frmLookUp.lvwReport, rst) Then
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
            MsgBox GetResourceString(526), _
                    vbCritical, wis_MESSAGE_TITLE
            Screen.MousePointer = vbDefault
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
            .m_SelItem = ""
            Screen.MousePointer = vbDefault
            .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem = "" Then Exit Sub
            txtData(txtIndex).Text = m_CustReg.CustomerName(Val(.m_SelItem))
            txtData(GetIndex("IntroducerID")).Text = .m_SelItem
        End With
        
End Select
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdAccept_Click()

If cmdAccept.Caption = GetResourceString(171) Then 'update
    If (M_UserPermission And perBankAdmin) = 0 Then
        MsgBox GetResourceString(685), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    If Not UpdateTransaction Then Exit Sub
Else
    'Checkthe permission
    If (M_UserPermission And perClerk) = 0 And (M_UserPermission And perCashier) = 0 Then
        If (M_UserPermission And perBankAdmin) = 0 Then
            MsgBox GetResourceString(685), vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
    
    'Check and perform appropriate transaction
    If Not AccountTransaction() Then Exit Sub
End If
txtAmount = 0

'Reload the account
    If Not AccountLoad(m_AccID) Then Exit Sub

    TabStrip2.Tabs(2).Selected = True
    
    If txtDate.Enabled Then
        txtDate.SelLength = Len(txtDate.Text)
        txtDate.SetFocus
    End If

End Sub

Private Sub cmdAddInterests_Click()
Dim rst As ADODB.Recordset

    If Not DateValidate(txtInterestFrom.Text, "/", True) Then
        'MsgBox "Invalid from date specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtInterestFrom
        Exit Sub
    End If
    If Not DateValidate(txtInterestTo.Text, "/", True) Then
       'MsgBox "Invalid to date specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtInterestTo
        Exit Sub
    End If

    'If MsgBox("WARNING !" & vbCrLf & vbCrLf & _
            "This will add interests to all the deposits !" & vbCrLf & _
            "Click OK only if you are sure about this operation !" & vbCrLf & vbCrLf & _
            "Are you sure you want to continue ?", vbQuestion + vbYesNo, gAppName & _
            " - Confirmation") = vbNo Then
     If MsgBox(GetResourceString(632) & vbCrLf & vbCrLf & _
            GetResourceString(633) & vbCrLf & _
            GetResourceString(634) & vbCrLf & vbCrLf & _
            GetResourceString(541), vbQuestion + vbYesNo, gAppName & _
            " - Confirmation") = vbNo Then
            Exit Sub
    End If

Dim FailIDs() As Long
ReDim FailIDs(0)
    
    If Not AccountDepositInterests(txtInterestFrom.Text, txtInterestTo.Text, gStrDate, FailIDs) Then
        'MsgBox "Unable to add interests!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
lblStatus.Caption = ""
Dim TmpStr As String
Dim I As Integer

For I = 0 To UBound(FailIDs) - 1
'Get the Account number From The account Id
    gDbTrans.SqlStmt = "SELECT AccNum From SBMaster Where AccID = " & FailIDs(I)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then TmpStr = TmpStr & ", " & FormatField(rst("AccNum"))
Next I

txtFailAccIDs = TmpStr
    
    If Not GetSBInterestChanged(Me.txtInterestFrom.Text, CSng(txtPropInterest)) Then
        'MsgBox "Unable To Add Interest ' ,vbcr ,wis_MESSAGE_TITLE"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

    'Update the SETUP
    Dim l_Setup As New clsSetup
    Dim DateStr As String
    
    DateStr = DateAdd("m", 1, GetSysFormatDate(txtInterestTo.Text))
    DateStr = 1 & "/" & Month(DateStr) & "/" & Year(DateStr)
    'DateStr = FormatDate(DateStr)
    txtInterestFrom.Text = DateStr
    txtInterestTo.Text = ""
    Call l_Setup.WriteSetupValue("SBAcc" & m_DepositType, "NextInterestOn", txtInterestFrom.Text)
    
    'Reset the progress
    prg.Value = 0
'Write To Interest Tab
End Sub
'
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

Private Sub cmdApply_Click()
'Cheque values
If Not CurrencyValidate(Me.txtPropCheque.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropCheque
    Exit Sub
End If
If Not CurrencyValidate(Me.txtPropNoCheque.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropNoCheque
    Exit Sub
End If

If Not CurrencyValidate(Me.txtPropBalance.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropBalance
    Exit Sub
End If
If Not CurrencyValidate(Me.txtPropCancel.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropCancel
    Exit Sub
End If
If Not CurrencyValidate(Me.txtPropPassBook.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropPassBook
    Exit Sub
End If
'If Val(Me.txtPropMaxTrans.Text) < 1 Then
'    MsgBox "Invalid Maximum transactions per day value specified !", vbExclamation, gAppName & " - Error"
'    ActivateTextBox txtPropMaxTrans
'    Exit Sub
'End If
If Val(Me.txtPropInterest.Text) < 0 Then
    'MsgBox "Invalid interest value specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(518), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtPropInterest
    Exit Sub
End If


Dim SetUp As New clsSetup

Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "MinBalanceWithChequeBook", txtPropCheque.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "MinBalanceWithoutChequeBook", txtPropNoCheque.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "InsufficientBalance", txtPropBalance.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "Cancellation", txtPropCancel.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "DuplicatePassBook", txtPropPassBook.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "MaxTransactions", txtPropMaxTrans.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "RateOfInterest", txtPropInterest.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "InterestFrom", txtInterestFrom.Text)
'Call Setup.WriteSetupValue("SBAcc" & m_DepositType, "NextInterestOn", txtPropNextInterestOn.Text)
Call SetUp.WriteSetupValue("SBAcc" & m_DepositType, "NoInterestOnMinBalance", chkMinBalInt.Value)
cmdApply.Enabled = False
End Sub

Private Sub cmdCheque_Click()
Set m_frmCheque = New frmCheque
m_frmCheque.p_AccID = m_AccID
m_frmCheque.AccHeadID = m_SBHeadId
m_frmCheque.Show vbModal
Dim RstCheque As ADODB.Recordset
'Get the cheque details
    gDbTrans.SqlStmt = "Select * from ChequeMaster where AccID = " & m_AccID _
            & " AND AccHeadID = " & m_SBHeadId
    cmbCheque.Clear
    cmbCheque.AddItem ""
    If gDbTrans.Fetch(RstCheque, adOpenForwardOnly) <= 0 Then
       Set RstCheque = Nothing
       Exit Sub
    End If
    
 With cmbCheque
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

Private Sub cmdDate_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtDate.Text, "/", True) Then .selDate = txtDate.Text
    .Left = Me.Left + Me.fraTransact.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + Me.fraTransact.Top + cmdDate.Top - 100
    .Show vbModal
    If .selDate <> "" Then
        txtDate.Text = .selDate
        RaiseEvent DateChanged(txtDate)
    End If
End With

End Sub

Private Sub cmdDate1_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtDate1.Text, "/", True) Then .selDate = txtDate1.Text
    .Left = Me.Left + Me.fraTransact.Left + cmdDate1.Left - .Width / 2
    .Top = Me.Top + Me.fraReport.Top + cmdDate2.Top + fraOrder.Top
    .Show vbModal
    If .selDate <> "" Then txtDate1.Text = .selDate
End With

End Sub

Private Sub cmdDate2_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtDate2.Text, "/", True) Then .selDate = txtDate2.Text
    .Left = Me.Left + Me.fraTransact.Left + cmdDate2.Left - .Width / 2
    .Top = Me.Top + Me.fraTransact.Top + cmdDate2.Top + fraOrder.Top
    .Show vbModal
    If .selDate <> "" Then txtDate2.Text = .selDate
End With

End Sub


Private Sub cmdDelete_Click()

If (M_UserPermission And perCreateAccount) = 0 Then
    If (M_UserPermission And perBankAdmin) = 0 Then
        'MsgBox "You have no permission to this operation", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(826), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If

'If MsgBox("Are you sure you want to delete this account ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(539), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
    Exit Sub
End If

Call AccountDelete
    
Call ResetUserInterface

End Sub

Private Sub cmdDepositType_Click()
    Dim cancel As Boolean
    Dim DepositType As Integer
    RaiseEvent SelectDeposit(DepositType, cancel)
    If Not cancel And DepositType <> m_DepositType Then txtAccNo = ""
    
End Sub

Private Sub cmdLoad_Click()
Dim rst As ADODB.Recordset

'get the account id from SBMaster
gDbTrans.SqlStmt = "SELECT ACCID From SBMaster WHere AccNum = " & AddQuotes(txtAccNo.Text, True)
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositTYpe = " & m_DepositType
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Call ResetUserInterface
    Exit Sub
End If
   
Dim AccId As Long
AccId = FormatField(rst("AccID"))

If txtDate.Enabled Then
    txtDate.SelLength = Len(txtDate.Text)
    Call PassBookPageInitialize
End If

If Not AccountLoad(AccId) Then
    ActivateTextBox txtAccNo
    Exit Sub
End If
'TabStrip2.Tabs(1).Selected = True

End Sub

Private Sub cmdNextTrans_Click()

If m_rstTrans Is Nothing Then Exit Sub
If m_rstTrans.EOF Then cmdNextTrans.Enabled = False: Exit Sub
If m_rstTrans Is Nothing Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    Exit Sub
End If

Call PassBookPageShow

End Sub


Private Sub cmdOk_Click()
Dim cancel As Boolean
Set m_Notes = Nothing

'If MsgBox(GetResourceString(750), vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
'    Exit Sub
'End If

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

If m_rstTrans Is Nothing Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    Exit Sub
End If

If m_rstTrans.EOF And m_rstTrans.BOF Then Exit Sub
Dim TransID As Long

If m_rstTrans.EOF Then
    m_rstTrans.MoveLast
    TransID = m_rstTrans.AbsolutePosition
    If TransID Mod 10 = 0 Then TransID = TransID - 10
    TransID = TransID - TransID Mod 10
    TransID = TransID - 10
    If TransID < 1 Then TransID = 1
Else
    TransID = m_rstTrans.AbsolutePosition
    TransID = TransID - TransID Mod 10
    TransID = TransID - 20
    If TransID < 1 Then TransID = 1
End If
m_rstTrans.MoveFirst
m_rstTrans.Move TransID - 1

Call PassBookPageShow

End Sub


Private Sub cmdPrint_Click()
    If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    
    m_frmPrintTrans.Show vbModal

End Sub
 
Private Sub cmdPrint1_Click()
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
    .ReportTitle = m_CustReg.FullName
    .GridObject = grd
    .PrintGrid
End With
Set m_grdPrint = Nothing
End Sub

Private Sub cmdReset_Click()

Call ResetUserInterface


End Sub
Private Sub cmdSave_Click()
Dim rst As ADODB.Recordset

If (M_UserPermission And perCreateAccount) = 0 Then
    If (M_UserPermission And perBankAdmin) = 0 Then
        MsgBox "You have no permission to this operation", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If

'SaveAccount
    If Not AccountSave Then Exit Sub

'Reload the account details once saved
    Dim AccNum As String
    Dim AccNo As Long
    AccNum = GetVal("AccID")
    'now get the account id of this account
    gDbTrans.SqlStmt = "SELECT ACCID From SBMAster WHERE AccNum = " & AddQuotes(AccNum, True)
    If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositTYpe = " & m_DepositType
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    AccNo = FormatField(rst("AccID"))
    If Not AccountLoad(AccNo) Then
       'If MsgBox("Are you sure you want to delete this account ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
       If MsgBox(GetResourceString(539), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
         Exit Sub
        End If
    End If
    txtAccNo.Text = AccNum

End Sub
Private Function AccountReopen(AccNo As Long) As Boolean
Dim rst As ADODB.Recordset

'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from SBMaster where AccID = " & AccNo
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If


    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Update SBMaster set ClosedDate = NULL where AccID = " & AccNo
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans
    
AccountReopen = True
End Function
 
Private Sub cmdTemp_Click()
'Select the Account Which are having Trans aftr the closed date

Dim rstTrans As Recordset
Dim TotalIntAmount As Currency
Dim TransDate As Date

gDbTrans.SqlStmt = "Select max(TransID)as MaxTransID," & _
        " max(TransDate) as MAxTransDate,AccID from SBTrans Group BY AccID"
gDbTrans.CreateView ("qrySBWrongInt")
    
gDbTrans.SqlStmt = "Select A.AccID, ClosedDate,A.AccNum,C.Amount,C.TransID,C.TransDate,C.Balance from (SBMAster A " & _
    " Inner Join qrySBWrongInt B on A.AccID = B.AccID) " & _
    " Inner Join (SBTrans C  inner join SBPLTrans D on C.AccId =D.AccID and C.TransID =D.TransID ) on A.AccID = C.AccID " & _
    " Where A.ClosedDate <  B.MaxTransDate and c.TransID=B.MAxTransID "
If m_DepositType > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " A.DepositType = " & m_DepositType
    
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " order BY A.AccID"
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Exit Sub

TransDate = rstTrans("TransDate")
gDbTrans.BeginTrans

While Not rstTrans.EOF
    ''Count the amount
    TotalIntAmount = TotalIntAmount + FormatField(rstTrans("Amount"))
    'Now Delete the Transcation in  SB Trans
    gDbTrans.SqlStmt = "Delete * from SBTrans where AccID= " & rstTrans("AccID") & " and TransID= " & rstTrans("TransID")
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        MsgBox "Error Occured"
    End If
    gDbTrans.SqlStmt = "Delete * from SBPLTrans where AccID= " & rstTrans("AccID") & " and TransID= " & rstTrans("TransID")
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        MsgBox "Error Occured in int"
    End If
    rstTrans.MoveNext
Wend

'''''''''''''''
Dim IntHeadID As Long
Dim AccHeadName As String
Dim AccHeadNameEnglish As String
Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
AccHeadName = m_DepositTypeName & " " & GetResourceString(487)
AccHeadNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(487)
IntHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepIntPaid, 0, wis_SBAcc + m_DepositType)

If Not bankClass.UndoContraTrans(IntHeadID, m_SBHeadId, TotalIntAmount, TransDate) Then
    gDbTrans.RollBack
    MsgBox "Error Occured in Undo contra"
End If


gDbTrans.CommitTrans
End Sub

Private Sub cmdUndo_Click()

If Not AccountUndoLastTransaction() Then Exit Sub

Call cmdLoad_Click

End Sub


Private Sub cmdUndoInterests_Click()
Dim rst As ADODB.Recordset

'Prelim check
    If Not DateValidate(txtUndoInterest.Text, "/", True) Then
        Exit Sub
    End If
    
    'If MsgBox("WARNING !" & vbCrLf & vbCrLf & _
            "This will withdraw interests to all the deposits of the specified date !" & vbCrLf & _
            "Click OK only if you are sure about this operation !" & vbCrLf & vbCrLf & _
            "Are you sure you want to continue ?", vbQuestion + vbYesNo, gAppName & _
            " - Confirmation") = vbNo Then
    If MsgBox(GetResourceString(632) & vbCrLf & vbCrLf & _
            GetResourceString(635) & vbCrLf & _
            GetResourceString(634) & vbCrLf & vbCrLf & _
            GetResourceString(541), vbQuestion + vbYesNo, gAppName & _
            " - Confirmation") = vbNo Then
            Exit Sub
    End If
    Me.Refresh

Dim FailIDs() As Long
ReDim FailIDs(0)
' Now show the satus onstatus label
    'lblStatus.Caption = "Deleteing Interest added on " & txtUndoInterest.Text
    lblStatus.Caption = GetResourceString(908) & " " & txtUndoInterest.Text
    
    If Not AccountDepositInterestsUNDO(txtUndoInterest.Text, FailIDs) Then
        'MsgBox "Unable to deposit interests !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        prg.Value = 0: lblStatus.Caption = ""
        Exit Sub
    End If
    prg.Value = 0
    lblStatus.Caption = ""
Dim TmpStr As String
Dim I As Integer

For I = 0 To UBound(FailIDs) - 1
'Get the Account number From The account Id
    gDbTrans.SqlStmt = "SELECT AccNum From SBMaster Where AccID = " & FailIDs(I)
    If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositTYpe = " & m_DepositType
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then TmpStr = TmpStr & ", " & FormatField(rst("AccNum"))
Next I

txtFailAccIDs = TmpStr

End Sub

Private Sub cmdView_Click()

'Make one round of date checks
If Me.txtDate1.Enabled Then
    If Not DateValidate(txtDate1.Text, "/", True) Then
        'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate1
        Exit Sub
    End If
End If
If txtDate2.Enabled Then
    If Not DateValidate(txtDate2.Text, "/", True) Then
        'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate2
        Exit Sub
    End If
End If
Dim ReportType As wis_SBReports
Dim ReportOrder As wis_ReportOrder

ReportOrder = IIf(optAccId, wisByAccountNo, wisByName)


If optReports(0).Value Then
    ReportType = repSBJoint
ElseIf optReports(1).Value Then
    ReportType = repSBBalance
ElseIf optReports(2).Value Then
    ReportType = repSBDayBook
ElseIf optReports(3).Value Then
    ReportType = repSBSubCashBook 'repSBLedger
ElseIf optReports(4).Value Then
    ReportType = repSBAccOpen
ElseIf optReports(5).Value Then
    ReportType = repSBAccClose
ElseIf optReports(6).Value Then
    ReportType = repSBProduct
ElseIf optReports(7).Value Then
    ReportType = repSBLedger
ElseIf optReports(8).Value Then
    ReportType = repSbMonthlyBalance
ElseIf optReports(9).Value Then
    ReportType = repSBCheque 'repSBSubCashBook
End If

    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption

RaiseEvent ShowReport(ReportType, ReportOrder, _
        txtDate1, txtDate2, m_clsRepOption)
        
        
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Initialize()
    m_DepositTypeName = GetResourceString(421)
    m_DepositTypeNameEnglish = LoadResourceStringS(421)
    m_DepositType = -1
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

Public Function AccountLoad(ByVal AccId As Long) As Boolean

Dim rstMaster As ADODB.Recordset
Dim rstJoint As ADODB.Recordset
Dim RstCheque As ADODB.Recordset
Dim rstTemp As ADODB.Recordset

Dim ClosedDate As String
Dim ret As Integer
Dim JointHolders() As String
Dim I As Integer

'Before Reloading the
Call ResetUserInterface
    
'Query data base
gDbTrans.SqlStmt = "Select * from SBMaster where AccID = " & AccId
If gDbTrans.Fetch(rstMaster, adOpenForwardOnly) <= 0 Then
    Call ResetUserInterface
    Exit Function
End If

'Load the Name details
    cmdPhoto.Enabled = True
    If m_CustReg Is Nothing Then
        Set m_CustReg = New clsCustReg
        m_CustReg.NewCustomer
        m_CustReg.ModuleID = wis_SBAcc
        cmdPhoto.Enabled = False
    End If
    If Not m_CustReg.LoadCustomerInfo(rstMaster("CustomerID")) Then
        'MsgBox "Unable to load customer information !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        Call ResetUserInterface
        Exit Function
    End If

'Get the Joint account holder details
    'If frm Joint already Loaded
    'and you load it once agian with out unload it
    'Then it will not free the memory Prevoiousley loaded
    'If Not m_frmSBJoint Is Nothing Then Unload m_frmSBJoint

    gDbTrans.SqlStmt = "Select * from SBJoint where AccID = " & AccId
    If gDbTrans.Fetch(rstJoint, adOpenForwardOnly) > 0 Then
        If m_frmSBJoint Is Nothing Then Set m_frmSBJoint = New frmJoint
        m_frmSBJoint.ModuleID = wis_SBAcc
        m_frmSBJoint.JointCustId(0) = m_CustReg.CustomerID
    End If

'Get the cheque details
    gDbTrans.SqlStmt = "Select * from ChequeMaster where AccID = " & AccId _
            & " AND AccHeadId = " & m_SBHeadId & " ORDER BY ChequeNo"
    cmbCheque.Clear
    cmbCheque.AddItem ""
    If gDbTrans.Fetch(RstCheque, adOpenForwardOnly) <= 0 Then Set RstCheque = Nothing

'Get the transaction details of this account holder
    Dim rst As ADODB.Recordset
    Dim BalanceAmount As Currency
    gDbTrans.SqlStmt = "Select TOP 1 Balance from SBTrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        cmdDelete.Enabled = IIf(M_UserPermission And perBankAdmin, True, False)  'True
        cmdUndo.Enabled = False
    Else
        BalanceAmount = FormatField(rst(0))
        Set rst = Nothing
    End If

'Assign to some module level variables
    m_AccID = AccId
    m_accUpdatemode = wis_UPDATE
    m_AccClosed = IIf(FormatField(rstMaster("ClosedDate")) <> "", True, False)
    
'Load account to the User Interface
    'TAB 1
    ClosedDate = FormatField(rstMaster("ClosedDate"))
    'If account is in operative then it is treates as same as closed account
    If ClosedDate = "" Then _
        If FormatField(rstMaster("InOperative")) Then ClosedDate = "1/1/100"

    With Me
        With .lblBalance
            .FontBold = True
            .ForeColor = vbBlue
            .Caption = GetResourceString(42) & " : " & FormatCurrency(BalanceAmount) '"Balance:  Rs. "
            If ClosedDate <> "" Then .Caption = GetResourceString(282) & ": " & ClosedDate: .ForeColor = vbRed '.Caption & " - Account Closed":
            .Alignment = vbRightJustify
        End With
        
        With .cmbAccNames
            .Enabled = True: .BackColor = vbWhite: .Clear
            .AddItem m_CustReg.CustomerName(rstMaster("CustomerID"))
            If Not rstJoint Is Nothing Then
                I = 0
                While Not rstJoint.EOF
                    .AddItem m_CustReg.CustomerName(rstJoint("CustomerID"))
                    m_JointCustID(I) = FormatField(rstJoint("CustomerID"))
                    m_frmSBJoint.JointCustId(I) = m_JointCustID(I)
                    rstJoint.MoveNext: I = I + 1
                Wend
            End If
            .ListIndex = 0
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
        With cmdPrint
            .Enabled = True
        End With
        With cmdUndo
            .Enabled = IIf(ClosedDate = "", True, False) And gCurrUser.IsAdmin
        End With
        
        With .cmbTrans
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            If gOnLine And (M_UserPermission <> perBankAdmin) Then
                .ListIndex = 1
                .Locked = True
            End If
            .Locked = False
            .Tag = 0
        End With

        With .cmbParticulars
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .ListIndex = -1
        End With
        
        With .txtAmount
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Value = 0
        End With
        
        With .cmbCheque
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Clear
            Dim ChequeNos() As Long
            If Not RstCheque Is Nothing Then
                .AddItem ""
                While Not RstCheque.EOF
                    If FormatField(RstCheque("Trans")) = wischqIssue Then
                        .AddItem RstCheque("ChequeNO")
                        .ItemData(.newIndex) = RstCheque("ChequeNO")
                    End If
                    RstCheque.MoveNext
                    I = I + 1
                Wend
            End If
        End With
        With txtVoucherNo
            .BackColor = vbWhite: .Enabled = True
        End With
        With .txtCheque
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Text = ""
        End With
        
        With .cmdCheque
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
        End With
        cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
        
        With .rtfNote
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
             Call m_Notes.LoadNotes(wis_SBAcc, AccId)
        End With
        
        Call m_Notes.DisplayNote(.rtfNote)
        
        .cmdAccept.Enabled = IIf(ClosedDate = "", True, False)
    End With
        
'TAB 2
'Update labels and other buttons
With Me
    .TabStrip2.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True
    lblOperation.Caption = GetResourceString(56) ' "Operation Mode : <UPDATE>"
    
    Dim strField As String
    Dim CtlIndex As Integer
    Dim CtlCount As Integer
        
    For I = 0 To txtPrompt.count - 1
        ' Read the bound field of this control.
        On Error Resume Next
        strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
        If strField <> "" Then
            With txtData(I)
                Select Case UCase$(strField)
                    Case "ACCID"
                        .Text = rstMaster("AccNum")
                        '.Locked = True
                    Case "ACCNAME"
                        .Text = m_CustReg.FullName
                    Case "NOMINEENAME"
                        .Text = FormatField(rstMaster("NomineeName"))
                    Case "NOMINEEAGE"
                        .Text = FormatField(rstMaster("NomineeAge"))
                    Case "NOMINEERELATION"
                        .Text = FormatField(rstMaster("NomineeRelation"))
                    Case "JOINTHOLDER"
                        If m_frmSBJoint Is Nothing Then Set m_frmSBJoint = New frmJoint
                        .Text = m_frmSBJoint.JointHolders  'FormatField(rstMaster("JointHolder"))
                    Case "INTRODUCERID"
                        .Text = IIf(rstMaster("IntroducerId") = "0", "", rstMaster("IntroducerId"))
                    Case "INTRODUCERNAME"
                        .Text = AccountName(rstMaster("IntroducerId"))
                    Case "LEDGERNO"
                        .Text = rstMaster("LedgerNo")
                    Case "FOLIONO"
                        .Text = rstMaster("FolioNO")
                    Case "CREATEDATE"
                        .Text = FormatField(rstMaster("CreateDate"))
                    Case "ACCGROUP"
                        gDbTrans.SqlStmt = "SELECT GroupName FROM AccountGroup WHERE " & _
                                "AccGroupID = " & FormatField(rstMaster("AccGroupId"))
                        If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
                            .Text = FormatField(rstTemp("GroupName"))
                         
                    Case "INOPERATIVE"
                        .Text = FormatField(rstMaster("InOperative"))
                    Case Else:
                        MsgBox "Label not found !", vbCritical, gAppName & " - Error"
                End Select
            End With
        End If
        
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
                    
                Case "BROWSE"
                    
                Case "BOOLEAN"
                    chk(CtlIndex).Value = IIf(txtData(I).Text = True, vbChecked, vbUnchecked)
                    
              End Select
            End With
        End If
    Next
End With
    
'Check is the DepositTYpe loaded or Not
If Len(lblDepositTypeName.Caption) = 0 And FormatField(rstMaster("DepositTYpe")) > 0 Then
    LoadDepositType (FormatField(rstMaster("DepositTYpe")))
End If
    
AccountLoad = True

Set rstJoint = Nothing
Set rstMaster = Nothing
Set RstCheque = Nothing

'enable the Control depending on the user permission
cmdDelete.Enabled = gCurrUser.IsAdmin 'IIf(cmdDelete.Enabled, True, IIf(M_UserPermission And perCreateAccount, True, False))
cmdAccept.Enabled = IIf(M_UserPermission And perClerk, True, False)
cmdUndo.Enabled = gCurrUser.IsAdmin 'IIf(M_UserPermission And perBankAdmin, True, False)
cmdView.Enabled = IIf(M_UserPermission And perReadOnly, True, False)

'Now Load the Transaction details
gDbTrans.SqlStmt = "Select * From SBTrans WHERE " & _
    " AccID = " & AccId & " ORDER By TransID"

If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
    Dim TransID As Long
    m_rstTrans.MoveLast
    TransID = m_rstTrans("TransID")
    m_rstTrans.MoveFirst
    If TransID Mod 10 = 0 Then TransID = TransID - 10
    m_rstTrans.Find "TransiD > " & TransID - TransID Mod 10
    If m_rstTrans.recordCount <= 10 Then m_rstTrans.MoveFirst
    PassBookPageShow
End If

If M_UserPermission And perBankAdmin Then
    cmdAccept.Enabled = True
    cmdUndo.Enabled = True
    cmdView.Enabled = True
End If

'cmdPhoto.Enabled = True
cmdPhoto.Enabled = Len(gImagePath)



RaiseEvent AccountChanged(m_AccID)
Exit Function

ErrLine:
    'MsgBox "Account Load:" & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(526), vbCritical, gAppName & " - Error"
End Function


Private Sub Form_Load()

m_FormLoaded = True
Screen.MousePointer = vbHourglass
'Caption = GetResourceString(1)

'Set icon for the form caption
'Icon = LoadResPicture(161, vbResIcon)
cmdPrint.Picture = LoadResPicture(102, vbResBitmap)

Call SetKannadaCaption
RaiseEvent UpdateStatus("Loading SB Account module...")
'Centre the form
Call CenterMe(Me)
lblStatus.Caption = ""

'Fill up transaction Types
With cmbTrans
   .AddItem GetResourceString(271) 'Deposit
   .AddItem GetResourceString(272) 'Withdraw
   .AddItem GetResourceString(273) 'Charges
   .AddItem GetResourceString(47) 'INterest
   
   'if it is on line put only WithDraw option
   If gOnLine And (M_UserPermission = perOnlyWaves) Then
        .ListIndex = 1 ''only witdraw
        .Locked = True
   End If
   
End With
     
'Fill up particulars with default values from SBAcc.INI
Call LoadParticularsFromFile(cmbParticulars, App.Path & "\SBACC.ini")
'Load ICONS
cmdAddNote.Picture = LoadResPicture(103, vbResBitmap)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)

fraInstructions.ZOrder (0)

TabStrip.Tabs(3).Tag = 1
Call LoadPropSheet

Call LoadSetupValues
cmdApply.Enabled = False


'Now Load the Account Groups
Dim cmbIndex As Byte
cmbIndex = GetIndex("AccGroup")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
Call LoadAccountGroups(cmb(cmbIndex))

txtData(Val(GetIndex("AccGroup"))).Text = cmb(cmbIndex).Text

Set m_CustReg = New clsCustReg

lblDepositTypeName.Caption = ""
'Reset the User Interface
Call ResetUserInterface

'Now enable the save button
'w.r.t user permission
M_UserPermission = gCurrUser.UserPermissions
cmdSave.Enabled = IIf(M_UserPermission And perCreateAccount, True, False)

txtDate2 = gStrDate

RaiseEvent UpdateStatus("")
Screen.MousePointer = vbDefault
If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

'Call SetWindowText(Me.hwnd, GetResourceString(1))


'''''TEMP CODE DLETE BELOW CODE
If (gCurrUser.UserPermissions And perOnlyWaves) Then cmdTemp.Visible = True Else cmdTemp.Visible = False

Dim rstTrans As Recordset
gDbTrans.SqlStmt = "Select max(TransID)as MaxTransID," & _
        " max(TransDate) as MAxTransDate,AccID from SBTrans Group BY AccID"
gDbTrans.CreateView ("qrySBWrongInt")
    
gDbTrans.SqlStmt = "Select A.AccID, ClosedDate,A.AccNum,C.Amount,C.TransID,C.TransDate,C.Balance from (SBMAster A " & _
    " Inner Join qrySBWrongInt B on A.AccID = B.AccID) " & _
    " Inner Join (SBTrans C  inner join SBPLTrans D on C.AccId =D.AccID and C.TransID =D.TransID ) on A.AccID = C.AccID " & _
    " Where A.ClosedDate <  B.MaxTransDate and c.TransID=B.MAxTransID " & _
    " order BY A.AccID"
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then cmdTemp.Visible = False


End Sub
Private Sub Form_Resize()
    lblDepositTypeName.Left = (Me.Width - lblDepositTypeName.Width) / 2 - 100
    cmdDepositType.Left = lblDepositTypeName.Left + lblDepositTypeName.Width = 50
End Sub

Private Sub Form_Unload(cancel As Integer)

m_FormLoaded = False
' Cheque form.
If Not m_frmCheque Is Nothing Then
    Unload m_frmCheque
    Set m_frmCheque = Nothing
End If

' Report form.
If Not m_frmLookUp Is Nothing Then
    Unload m_frmLookUp
    Set m_frmLookUp = Nothing
End If

' Notes object.
Set m_Notes = Nothing

' Customer Registration object.
Set m_CustReg = Nothing
If Not m_frmSBJoint Is Nothing Then
    Unload m_frmSBJoint
    Set m_frmSBJoint = Nothing
End If

gWindowHandle = 0

RaiseEvent WindowClosed

End Sub

Private Sub grd_DblClick()
Dim TransID As Integer
TransID = grd.RowData(grd.Row)
If TransID = 0 Then Exit Sub

cmbTrans.Tag = TransID
Call LoadPreviousTransaction(TransID)

End Sub



Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)

Call PrintBetweenDates(wis_SBAcc, m_AccID, StartIndiandate, EndIndianDate)

Exit Sub

Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer
Dim YLocation As Integer

 '1.select the details of theaccount holder
SqlStr = "SELECT A.Name,B.AccNum " & _
            " FROM QryName A, " & _
           " SBMaster as  B " & _
            "WHERE A.CustomerId = B.CustomerId " & _
            "AND B.AccId = " & m_AccID
            
            
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Set clsPrint = New clsTransPrint

'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From SBTrans WHERE AccId = " & m_AccID
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
    If (rst(0) = 0) Then
        MsgBox "There are no transactions available for printing."
        Exit Sub
    End If

'Print [or don't print] header part
'lastPrintRow = IIf(IsNull(Rst("LastPrintRow")), 0, Rst("LastPrintRow"))
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If

'3. Getting matching records for passbook printing
    SqlStr = "SELECT * From SBTrans WHERE AccId = " & m_AccID & _
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
Printer.Font = gFontName '"Courier New"
Printer.FONTSIZE = 12
Dim BeginY As Integer


With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    
    .Cols = 5
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .COlHeader(1) = GetResourceString(275) 'Cheque
    '.ColWidth(2) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(3) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(4) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(5) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    
    If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE_A4) Then
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
     .ColWidth(0) = 15
        .ColWidth(1) = 22
        .ColWidth(2) = 8
        .ColWidth(3) = 13
        .ColWidth(4) = 14
        .ColWidth(5) = 15
    BeginY = Printer.CurrentY
    While Not rst.EOF
        If .isNewPage Then
            Printer.Print " "
            ''Print the BANK Name
            Printer.CurrentY = 1000
            
            Printer.Font.name = gFontName
            Printer.Font.Size = Printer.Font.Size + 2
            Printer.CurrentX = 5000 - (Printer.TextWidth(gCompanyName) / 2)
            Printer.Font.Bold = True
            Printer.Print gCompanyName
            Printer.Font.Bold = False
            Printer.Font.Size = Printer.Font.Size - 2
            
            BeginY = Printer.CurrentY
            Printer.CurrentY = BeginY + 50
            Printer.Print (FormatField(metaRst("Name")))
            Printer.CurrentX = 9000
            Printer.CurrentY = BeginY + 50
            Printer.Print GetResourceString(421, 60); ":" & FormatField(metaRst("AccNum"))
            .printHeader
            BeginY = Printer.CurrentY
            .isNewPage = False
        End If
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        .ColText(2) = FormatField(rst("ChequeNo"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(3) = FormatField(rst("Amount"))
            .ColText(4) = " "
        Else
            .ColText(4) = FormatField(rst("Amount"))
            .ColText(3) = " "
        End If
        .ColText(5) = FormatField(rst("Balance"))
        .PrintRow
        YLocation = Printer.CurrentY + 100
        ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE_A4) Then
            ' since we have to print now in a new page,
            ' we need to print the header.' So, set columns widths for header.
            
            Printer.Line (1800, BeginY)-(1800, YLocation)
            Printer.Line (4500, BeginY)-(4500, YLocation)
            Printer.Line (5700, BeginY)-(5650, YLocation)
            Printer.Line (7400, BeginY)-(7400, YLocation)
            Printer.Line (9250, BeginY)-(9250, YLocation)
            'Printer.Line (11200, 1500)-(11200, YLocation)
            
            Printer.CurrentX = 0
            .newPage
           
            curPrintRow = 1
        End If
        rst.MoveNext
    Wend
    '.newPage
End With

    Printer.Line (1800, BeginY)-(1800, YLocation)
    Printer.Line (4500, BeginY)-(4500, YLocation)
    Printer.Line (5700, BeginY)-(5700, YLocation)
    Printer.Line (7400, BeginY)-(7400, YLocation)
    Printer.Line (9250, BeginY)-(9250, YLocation)
    'Printer.Line (11200, 1500)-(11200, YLocation)
    
Printer.EndDoc

Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing

'Now Update the Last Print Id to the master
SqlStr = "UPDATE SBMaster set LastPrintRow = " & curPrintRow - 1 & _
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
Const HEADER_ROWS = 4
Const ROWS_PER_PAGE2 = 17
Dim curPrintRow As Integer

Set clsPrint = New clsTransPrint

' Print the first page of passbook, if newPassbook option is chosen.

If bNewPassbook Then
    clsPrint.printPassbookPage wis_SBAcc, m_AccID '
    'Update the print rows as 0 and lst transid as -5
    'Now Update the Last Print Id to the master
    SqlStr = "UPDATE SBMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
            ", LastPrintRow = 0 Where Accid = " & m_AccID
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
    Else
        gDbTrans.CommitTrans
    End If
    
    MsgBox "First page printed successfully!", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


'First get the last printed txnId and last printed row From the SbMaster
SqlStr = "SELECT  LastPrintID, LastPrintRow From SBMaster WHERE AccId = " & m_AccID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'get position of the last Trans Printed in
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintId"))
'get position of the last print point
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If IsNull(metaRst("LastPrintRow")) And lastPrintId = 1 Then lastPrintId = 0

'Count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT * From SBTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
SqlStr = "SELECT count(*) From SBTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'If there are no records to print, since the last printed transaction,
'display a message and exit.
If (rst(0) = 0) Then
    Dim iRet As Integer
    iRet = MsgBox("There are no transactions available for printing." & vbCrLf & _
    "Do you want to reset the printing for this account?", vbYesNo + vbDefaultButton2, "Debug Message")
    
    If (iRet = vbYes) Then
           SqlStr = "UPDATE SBMaster set LastPrintId = " & TransID & _
                    ", LastPrintRow = " & curPrintRow - 1 & _
                    " Where Accid = " & m_AccID
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


' Fetch records for txns that have been created after lasttxnId.
SqlStr = "SELECT * From SBTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Printing starts here....
Dim count As Integer
Dim currPrintRow As Integer
If lastPrintRow <= 0 Then
    currPrintRow = 0
Else
    currPrintRow = lastPrintRow
End If

With clsPrint
    
    .Cols = 6
    
    'skip the lines already printed
    For count = 1 To currPrintRow
        'accommodate the header of the first part of passbook
        If (count = 1) Then .NewLines (4) 'header of the first portion of the passbook
        
        'accommodate the header of the second part of passbook
        If (count = (wis_ROWS_PER_PAGE / 2) + 1) Then
            .NewLines (4)   'new lines towards the end of the first page
            .NewLines (4)   'header of the second portion of the passbook
        End If
        .NewLines (1)
    Next count
    If currPrintRow > 0 Then
        Printer.CurrentY = Printer.CurrentY + IIf(currPrintRow > (wis_ROWS_PER_PAGE / 2), -300, 200)
    End If
    
    While Not rst.EOF
       'if we are at the end of page, prompt to turn to a new page...
        If (currPrintRow >= wis_ROWS_PER_PAGE) Then
            Printer.EndDoc
            'Pause Application til the printing completes
            PauseApplication (5)
            MsgBox "End of page reached. Please turn over to new page. ", vbOKOnly
            '.newPage
            currPrintRow = 0    'rest the last printed row
            
        End If
    
        'if we at the end of the top portion of the passbook, leave a few blank spaces to go to the second (lower) part
        If (currPrintRow = (wis_ROWS_PER_PAGE / 2)) Then
            'Printer.Print "...."
            .NewLines (1)
        End If
    
        If (currPrintRow = 0 Or currPrintRow = (wis_ROWS_PER_PAGE / 2)) Then
            .printHeader
            If currPrintRow = 0 Then
                Dim YLocation As Integer
                YLocation = Printer.CurrentY
                Printer.Line (1800, 200)-(1800, 10000)
                Printer.Line (4500, 200)-(4500, 10000)
                Printer.Line (5700, 200)-(5700, 10000)
                Printer.Line (7400, 200)-(7400, 10000)
                Printer.Line (9250, 200)-(9250, 10000)
                
                Printer.CurrentY = YLocation
                Printer.CurrentX = 0
            End If
        End If
        
        'set the column sizes for the row
        .ColWidth(0) = 15
        .ColWidth(1) = 22
        .ColWidth(2) = 8
        .ColWidth(3) = 13
        .ColWidth(4) = 14
        .ColWidth(5) = 15
    
        TransID = FormatField(rst("TransID"))
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        .ColText(2) = FormatField(rst("ChequeNo"))
        
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(3) = FormatField(rst("Amount"))
            .ColText(4) = " "
        Else
            .ColText(4) = FormatField(rst("Amount"))
            .ColText(3) = " "
        End If
        .ColText(5) = FormatField(rst("Balance"))
        
        .PrintRow
        currPrintRow = currPrintRow + 1
        'Debug.Print currPrintRow & " " & Printer.CurrentY & "   " & rst("TransDate") & "  "; rst("Amount")
        rst.MoveNext
        
    Wend
    .newPage
    
End With
Printer.EndDoc
'Printing ends here...


Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing

    'Now Update the Last Print Id to the master
    SqlStr = "UPDATE SBMaster set LastPrintId = " & TransID & _
            ", LastPrintRow = " & currPrintRow & _
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
Dim Amt As Boolean
Dim Gender As Boolean

Select Case Index
    Case 0
        
    Case 1
        Amt = True
        Gender = True
    Case 2
        Dt1 = True: Amt = True
        Gender = True
    Case 3
        Dt1 = True
        Amt = True
        Gender = True
    Case 4
        Dt1 = True
        Gender = True
    Case 5
        Dt1 = True
        Gender = True
    Case 6
        Dt1 = True
    Case 7
        Dt1 = True 'Gender = True
    Case 8
        Dt1 = True
    Case 9
        'Dt1 = True
        'Gender = True
        'Amt = True
End Select

    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = Gender
        .EnableAmountRange = Amt
    End With

    lblDate1.Enabled = Dt1
    With txtDate1
        .Enabled = Dt1
        .BackColor = IIf(Dt1, wisWhite, wisGray)
    End With
    
    
    cmdDate1.Enabled = Dt1

End Sub

Private Sub optReports_GotFocus(Index As Integer)
TabStrip.Tabs(3).Tag = Index
End Sub


Private Sub TabStrip_Click()
On Error Resume Next
fraTransact.Visible = False
fraNew.Visible = False
fraProps.Visible = False
fraReports.Visible = False

Select Case TabStrip.SelectedItem.Index

    Case 1
        fraTransact.Visible = True
        cmdAccept.Default = True
        txtAccNo.SetFocus
      Case 2
        fraNew.Visible = True
        cmdSave.Default = True
        txtData(Val(TabStrip.Tabs(2).Tag)).SetFocus
    Case 4
        fraProps.Visible = True
        cmdApply.Default = True
        txtPropBalance.SetFocus
    Case 3
        fraReports.Visible = True
        cmdView.Default = True
        Call optReports(Val(TabStrip.Tabs(3).Tag)).SetFocus
        
End Select

End Sub

Private Sub TabStrip2_Click()
    If TabStrip2.SelectedItem.Index = 1 Then
        fraInstructions.Visible = True
        fraInstructions.ZOrder 0
        fraPassBook.Visible = False
    End If
    If TabStrip2.SelectedItem.Index = 2 Then
        fraInstructions.Visible = False
        fraPassBook.Visible = True
        fraPassBook.ZOrder 0
        
    End If
End Sub



Private Sub txtAccNo_Change()
cmdLoad.Enabled = IIf(Trim$(txtAccNo.Text) <> "", True, False)
Call ResetUserInterface

End Sub


Private Sub txtAmount_Click()
On Error Resume Next
txtAmount.ToolTipText = txtAmount.TextInFigure
End Sub

Private Sub txtAmount_GotFocus()
txtAmount.SetFocus
End Sub

Private Sub txtAmount_LostFocus()
txtCheque.Text = 1
End Sub

Private Sub txtCheque_LostFocus()
cmbParticulars.Text = "BY CASH"
End Sub

Private Sub txtData_DblClick(Index As Integer)
    txtData_KeyPress Index, vbKeyReturn
End Sub
Private Sub txtData_GotFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlue
SetDescription txtPrompt(Index)
TabStrip.Tabs(2).Tag = Index
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
cmdSave.Default = True

strDispType = ExtractToken(txtPrompt(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    If TextIndex <> "" Then cmd(Val(TextIndex)).Visible = True
ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    cmdSave.Default = False
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    ' Get the cmdbutton index.
    On Error Resume Next
    If TextIndex <> "" Then
        If cmb(Val(TextIndex)).Visible Then Exit Sub
        cmb(Val(TextIndex)).Visible = True
        cmb(Val(TextIndex)).SetFocus
    End If
End If


' Hide all other command buttons...
Dim I As Integer
For I = 0 To cmd.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmd(I).Visible = False
    End If
Next

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
Dim rst As ADODB.Recordset

' If the item is IntroducerID, validate the
' ID and name.
strDatSrc = ExtractToken(txtPrompt(Index).Tag, "DataSource")
Select Case UCase(strDatSrc)
    Case "INTRODUCERID"
        ' Check if any data is found in this text.
        If Trim$(txtData(Index).Text) <> "" Then
            If Val(txtData(Index).Text) <= 0 Then
                'MsgBox "Invalid account number specified !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(500), vbExclamation, gAppName & " - Error"
                ActivateTextBox txtData(Index)
                Exit Sub
            End If
            gDbTrans.SqlStmt = "SELECT Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM " _
                & "NameTab WHERE CustomerID = " & Val(txtData(Index).Text)
                
            Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
            If Lret > 0 Then
                txtIndex = GetIndex("IntroducerName")
                txtData(txtIndex).Text = FormatField(rst("Name"))
            End If
        Else
            txtIndex = GetIndex("IntroducerName")
            txtData(txtIndex).Text = ""
        End If
    Case "ACCID"
    
End Select

End Sub

Private Sub txtDate_GotFocus()
TabStrip2.Tabs(2).Selected = True
End Sub


Private Sub txtDate_LostFocus()
RaiseEvent DateChanged(txtDate)
End Sub


Private Sub txtInterestFrom_Change()
    cmdAddInterests.Enabled = False
    If Not DateValidate(txtInterestFrom.Text, "/", True) Then
        'MsgBox "Invalid from date specified !", vbExclamation, gAppName & " - Error"
        'ActivateTextBox txtPropLastInterestOn
        Exit Sub
    End If
    If Not DateValidate(txtInterestTo.Text, "/", True) Then
        'MsgBox "Invalid to date specified !", vbExclamation, gAppName & " - Error"
        'ActivateTextBox txtPropNextInterestOn
        Exit Sub
    End If
    cmdAddInterests.Enabled = True

End Sub

Private Sub txtInterestTo_Change()
    cmdAddInterests.Enabled = False
    If Not DateValidate(txtInterestFrom.Text, "/", True) Then
        'MsgBox "Invalid from date specified !", vbExclamation, gAppName & " - Error"
        'ActivateTextBox txtPropLastInterestOn
        Exit Sub
    End If
    If Not DateValidate(txtInterestTo.Text, "/", True) Then
        'MsgBox "Invalid to date specified !", vbExclamation, gAppName & " - Error"
        'ActivateTextBox txtPropNextInterestOn
        Exit Sub
    End If
    cmdAddInterests.Enabled = True

End Sub






Private Sub txtPropBalance_Change()
cmdApply.Enabled = True
End Sub

Private Sub txtPropCancel_Change()
cmdApply.Enabled = True
End Sub


Private Sub txtPropCheque_Change()
cmdApply.Enabled = True
End Sub

Private Sub txtPropInterest_Change()
cmdApply.Enabled = True
End Sub



Private Sub txtPropNoCheque_Change()
cmdApply.Enabled = True
End Sub


Private Sub txtPropPassBook_Change()
cmdApply.Enabled = True
End Sub

Private Sub txtUndoInterest_Change()
    cmdUndoInterests.Enabled = False
    If Not DateValidate(txtUndoInterest.Text, "/", True) Then
        Exit Sub
    End If
    cmdUndoInterests.Enabled = gCurrUser.IsAdmin
End Sub

Private Sub VScroll1_Change()
' Move the picSlider.
picSlider.Top = -VScroll1.Value
End Sub


Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

'Now Assign The Names to the Controls
'The Below Code load From The the resource file

TabStrip.Tabs(1).Caption = GetResourceString(38)
TabStrip.Tabs(2).Caption = GetResourceString(211)
TabStrip.Tabs(3).Caption = GetResourceString(283) & GetResourceString(92)
TabStrip.Tabs(4).Caption = GetResourceString(213)

cmdAccept.Caption = GetResourceString(4)
cmdOk.Caption = GetResourceString(1)
cmdLoad.Caption = GetResourceString(3)

' TransCtion Frame
lblAccNo.Caption = GetResourceString(36, 60)
lblName.Caption = GetResourceString(35)
lblDate.Caption = GetResourceString(37)    '
lblTrans.Caption = GetResourceString(38)
lblParticular.Caption = GetResourceString(39)
lblAmount.Caption = GetResourceString(40)
lblBalance.Caption = GetResourceString(42)
lblInstrNo.Caption = GetResourceString(275, 60)
cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)    '
lblVoucher.Caption = GetResourceString(41)  'Voucher No
TabStrip2.Tabs(1).Caption = GetResourceString(219)
TabStrip2.Tabs(2).Caption = GetResourceString(218)

'Now Change the Font of New Account Frame
cmdDelete.Caption = GetResourceString(14)
cmdSave.Caption = GetResourceString(7)
cmdReset.Caption = GetResourceString(8)
lblOperation.Caption = GetResourceString(54)
cmdPhoto.Caption = GetResourceString(415)

'Now Changet The Caption of Report Frame
fraReport.Caption = GetResourceString(283) & GetResourceString(92)
optReports(0).Caption = GetResourceString(265, 36) 'Joint account
optReports(1).Caption = GetResourceString(67)
optReports(2).Caption = GetResourceString(390, 63) 'Sub Day Book
'optReports(3).Caption = GetResourceString(421,93)
optReports(3).Caption = GetResourceString(390, 85) 'Sub Cash BooK
optReports(4).Caption = GetResourceString(64)
optReports(5).Caption = GetResourceString(65)
optReports(6).Caption = GetResourceString(66)
'optReports(7).Caption = GetResourceString(177)  'Accounts with cheque
optReports(7).Caption = GetResourceString(421, 93) 'General Ledger
optReports(8).Caption = GetResourceString(463, 42) 'Monthly Balance
'optReports(9).Caption = GetResourceString(390,85) 'Sub Cash BooK
optReports(9).Caption = GetResourceString(177)  'Accounts with cheque

fraOrder.Caption = GetResourceString(287)
optAccId.Caption = GetResourceString(68)
optName.Caption = GetResourceString(69)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
cmdView.Caption = GetResourceString(13)

' Change the Captions Of Properites frame"
fraMinBalance.Caption = GetResourceString(147, 42)
lblWithCheque.Caption = GetResourceString(177)
lblWithoutCheque.Caption = GetResourceString(178)
lblInsuffBal.Caption = GetResourceString(179)    '"
fraCharges.Caption = GetResourceString(273)    '
lblCancellation.Caption = GetResourceString(181)
lblDupPassbook.Caption = GetResourceString(182)
fraInterest.Caption = GetResourceString(47, 295)
lblIntFromDate.Caption = GetResourceString(184)
lblIntToDate.Caption = GetResourceString(108)   '
lblRateOfInterest.Caption = GetResourceString(186)
cmdAddInterests.Caption = GetResourceString(187)
lblUnDoInterest.Caption = GetResourceString(228)
lblFailAccIDs.Caption = GetResourceString(190)    '
cmdUndoInterests.Caption = GetResourceString(188)
cmdApply.Caption = GetResourceString(6)

lblPropMaxTrans.Caption = GetResourceString(333)
cmdAdvance.Caption = GetResourceString(491)    'Options

End Sub

