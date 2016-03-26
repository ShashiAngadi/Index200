VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmRDAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX-2000   -   Reccuring Deposit  Account Wizard"
   ClientHeight    =   7845
   ClientLeft      =   1275
   ClientTop       =   1305
   ClientWidth     =   8325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDepositType 
      Appearance      =   0  'Flat
      Caption         =   ".."
      Height          =   315
      Left            =   4680
      TabIndex        =   102
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   450
      Left            =   6840
      TabIndex        =   33
      Top             =   7230
      Width           =   1335
   End
   Begin VB.Frame fraProps 
      Height          =   6465
      Left            =   390
      TabIndex        =   73
      Top             =   570
      Width           =   7575
      Begin VB.Frame fraInterest 
         Caption         =   "Interest rates (%)"
         Height          =   4335
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   7575
         Begin VB.TextBox txtIntDate 
            Height          =   345
            Left            =   1860
            TabIndex        =   95
            Top             =   3360
            Width           =   1365
         End
         Begin VB.TextBox txtSenInt 
            Height          =   315
            Left            =   2430
            TabIndex        =   94
            Top             =   2430
            Width           =   795
         End
         Begin VB.TextBox txtEmpInt 
            Height          =   315
            Left            =   2430
            TabIndex        =   93
            Top             =   2010
            Width           =   795
         End
         Begin VB.TextBox txtGenInt 
            Height          =   315
            Left            =   2430
            TabIndex        =   92
            Top             =   1590
            Width           =   795
         End
         Begin VB.ComboBox cmbTo 
            Height          =   315
            Left            =   1890
            TabIndex        =   91
            Top             =   1170
            Width           =   1335
         End
         Begin VB.ComboBox cmbFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   90
            Top             =   1214
            Width           =   1455
         End
         Begin VB.OptionButton optDays 
            Caption         =   "Days"
            Height          =   300
            Left            =   90
            TabIndex        =   89
            Top             =   390
            Width           =   1335
         End
         Begin VB.OptionButton optMon 
            Caption         =   "Month"
            Height          =   300
            Left            =   1770
            TabIndex        =   88
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtLoanInt 
            Height          =   315
            Left            =   2430
            TabIndex        =   87
            Text            =   "+"
            Top             =   2850
            Width           =   795
         End
         Begin VB.CommandButton cmdIntApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   450
            Left            =   1890
            TabIndex        =   86
            Top             =   3780
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid grdInt 
            Height          =   4125
            Left            =   3240
            TabIndex        =   96
            Top             =   150
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   7276
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            AllowUserResizing=   3
         End
         Begin VB.Label lblIntApply 
            Caption         =   "Int apply date"
            Height          =   555
            Left            =   90
            TabIndex        =   13
            Top             =   3390
            Width           =   1455
         End
         Begin VB.Label lblSenInt 
            Caption         =   "Senior Citizen"
            Height          =   300
            Left            =   90
            TabIndex        =   58
            Top             =   2465
            Width           =   1905
         End
         Begin VB.Label lblEmpInt 
            Caption         =   "Emplyees Interest Rate"
            Height          =   300
            Left            =   90
            TabIndex        =   59
            Top             =   2053
            Width           =   1965
         End
         Begin VB.Label lblGenlInt 
            Caption         =   "General Interest"
            Height          =   300
            Left            =   90
            TabIndex        =   100
            Top             =   1641
            Width           =   1995
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   300
            Left            =   1770
            TabIndex        =   99
            Top             =   765
            Width           =   1095
         End
         Begin VB.Label lblFrom 
            Caption         =   "from"
            Height          =   300
            Left            =   180
            TabIndex        =   98
            Top             =   802
            Width           =   1035
         End
         Begin VB.Label lblLoanInt 
            Caption         =   "Max loan percent:"
            Height          =   300
            Left            =   90
            TabIndex        =   97
            Top             =   2880
            Width           =   1965
         End
      End
      Begin VB.CommandButton cmdUndoPayble 
         Caption         =   "Command1"
         Height          =   450
         Left            =   4860
         TabIndex        =   80
         Top             =   5130
         Width           =   2475
      End
      Begin VB.TextBox txtFailAccIDs 
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   6000
         Width           =   6825
      End
      Begin VB.CommandButton cmdIntPayable 
         Caption         =   "&Interest Payable"
         Height          =   450
         Left            =   480
         TabIndex        =   76
         Top             =   5130
         Width           =   2745
      End
      Begin VB.TextBox txtIntPayable 
         Height          =   315
         Left            =   2340
         TabIndex        =   74
         Top             =   4530
         Width           =   1245
      End
      Begin ComctlLib.ProgressBar prg 
         Height          =   300
         Left            =   240
         TabIndex        =   77
         Top             =   6000
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label txtLastIntDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5940
         TabIndex        =   84
         Top             =   4530
         Width           =   1365
      End
      Begin VB.Label lblLastIntDate 
         Caption         =   "Last Interest Updated on :"
         Height          =   285
         Left            =   3780
         TabIndex        =   83
         Top             =   4590
         Width           =   1875
      End
      Begin VB.Label lblStatus 
         Caption         =   "x"
         Height          =   405
         Left            =   240
         TabIndex        =   78
         Top             =   5550
         Width           =   6795
      End
      Begin VB.Label lblIntPayable 
         AutoSize        =   -1  'True
         Caption         =   "Interest Payable Date"
         Height          =   345
         Left            =   210
         TabIndex        =   75
         Top             =   4590
         Width           =   1950
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7065
      Left            =   120
      TabIndex        =   18
      Top             =   60
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   12462
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
   Begin VB.Frame fraNew 
      Height          =   6465
      Left            =   390
      TabIndex        =   30
      Top             =   570
      Width           =   7575
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   450
         Left            =   6180
         TabIndex        =   101
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "&Terminate"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6180
         TabIndex        =   36
         Top             =   4470
         Width           =   1200
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   4650
         Left            =   150
         ScaleHeight     =   4590
         ScaleWidth      =   5895
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1305
         Width           =   5955
         Begin VB.VScrollBar VScroll1 
            Height          =   4635
            Left            =   5640
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picSlider 
            Height          =   4485
            Left            =   -45
            ScaleHeight     =   4425
            ScaleWidth      =   5610
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   15
            Width           =   5670
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Caption         =   "Check1"
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   81
               Top             =   630
               Width           =   1755
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2355
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   -30
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   4860
               TabIndex        =   64
               Top             =   0
               Visible         =   0   'False
               WhatsThisHelpID =   315
               Width           =   300
            End
            Begin VB.TextBox txtPrompt 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Index           =   0
               Left            =   30
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               Text            =   "Account Holder"
               Top             =   0
               Width           =   2475
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
               Left            =   2520
               TabIndex        =   32
               Top             =   0
               Width           =   3060
            End
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   960
         Left            =   150
         ScaleHeight     =   900
         ScaleWidth      =   5775
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   285
         Width           =   5835
         Begin VB.Image imgNewAcc 
            Height          =   435
            Left            =   180
            Stretch         =   -1  'True
            Top             =   150
            Width           =   375
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   1020
            TabIndex        =   63
            Top             =   330
            Width           =   4710
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
            TabIndex        =   62
            Top             =   45
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Clear"
         Height          =   450
         Left            =   6180
         TabIndex        =   35
         Top             =   5565
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   450
         Left            =   6180
         TabIndex        =   34
         Top             =   5025
         Width           =   1200
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
         Height          =   195
         Left            =   135
         TabIndex        =   66
         Top             =   6030
         Width           =   1290
      End
   End
   Begin VB.Frame fraTransact 
      Height          =   6465
      Left            =   390
      TabIndex        =   67
      Top             =   570
      Width           =   7575
      Begin VB.TextBox txtCheque 
         Height          =   345
         Left            =   5850
         TabIndex        =   17
         Top             =   2220
         Width           =   1305
      End
      Begin VB.TextBox TxtInstallmentNo 
         Height          =   345
         Left            =   5850
         TabIndex        =   15
         Top             =   1773
         Width           =   1305
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3105
         TabIndex        =   5
         Top             =   1326
         Width           =   315
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   1740
         MaxLength       =   9
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   717
         Width           =   5475
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1773
         Width           =   1965
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   450
         Left            =   6240
         TabIndex        =   27
         Top             =   5850
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1470
         TabIndex        =   6
         Top             =   1326
         Width           =   1515
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   450
         Left            =   4560
         TabIndex        =   28
         Top             =   5850
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   450
         Left            =   3180
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         ItemData        =   "Rdacc.frx":0000
         Left            =   1470
         List            =   "Rdacc.frx":0002
         TabIndex        =   10
         Top             =   2220
         Width           =   1965
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close A/C"
         Enabled         =   0   'False
         Height          =   450
         Left            =   3000
         TabIndex        =   29
         Top             =   5850
         Width           =   1335
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2325
         Left            =   240
         TabIndex        =   20
         Top             =   3225
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
            TabIndex        =   21
            Top             =   195
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2205
            Left            =   120
            TabIndex        =   68
            Top             =   150
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   3889
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Rdacc.frx":0004
         End
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   5850
         TabIndex        =   12
         Top             =   1326
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraPassBook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2265
         Left            =   270
         TabIndex        =   22
         Top             =   3240
         Width           =   7035
         Begin VB.CommandButton cmdPrint 
            Height          =   345
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1770
            Width           =   405
         End
         Begin VB.CommandButton cmdPrevTrans 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   105
            Width           =   405
         End
         Begin VB.CommandButton cmdNextTrans 
            Height          =   375
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   540
            Width           =   405
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2175
            Left            =   60
            TabIndex        =   23
            Top             =   90
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   5
            AllowUserResizing=   1
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2925
         Left            =   150
         TabIndex        =   19
         Top             =   2760
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
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   150
         X2              =   7440
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   7050
         Y1              =   1164
         Y2              =   1164
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No. : "
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblName 
         Caption         =   "Name(s) : "
         Height          =   300
         Left            =   150
         TabIndex        =   3
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   1773
         Width           =   1035
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   300
         Left            =   3900
         TabIndex        =   11
         Top             =   1326
         Width           =   1545
      End
      Begin VB.Label lblInstllNo 
         AutoSize        =   -1  'True
         Caption         =   "Installment Number :"
         Height          =   300
         Left            =   3870
         TabIndex        =   14
         Top             =   1773
         Width           =   1650
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance : Rs. 00.00"
         Height          =   300
         Left            =   5400
         TabIndex        =   70
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label lblParticular 
         Caption         =   "Particulars : "
         Height          =   300
         Left            =   150
         TabIndex        =   9
         Top             =   2220
         Width           =   945
      End
      Begin VB.Label lblDate 
         Caption         =   "Date : "
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   1326
         Width           =   735
      End
      Begin VB.Label lblInstrNo 
         Caption         =   "Iinstrument Number"
         Height          =   300
         Left            =   3870
         TabIndex        =   16
         Top             =   2220
         Width           =   1575
      End
   End
   Begin VB.Frame fraReports 
      Height          =   6465
      Left            =   390
      TabIndex        =   71
      Top             =   570
      Width           =   7575
      Begin VB.Frame fraOrder 
         Caption         =   "Order By"
         Height          =   2130
         Left            =   210
         TabIndex        =   48
         Top             =   3570
         Width           =   7125
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   450
            Left            =   5760
            TabIndex        =   57
            Top             =   1500
            Width           =   1215
         End
         Begin VB.CommandButton cmdDate1 
            Caption         =   "..."
            Height          =   315
            Left            =   2940
            TabIndex        =   52
            Top             =   1080
            Width           =   315
         End
         Begin VB.TextBox txtFromDate 
            Height          =   315
            Left            =   1620
            TabIndex        =   53
            Top             =   1080
            Width           =   1290
         End
         Begin VB.TextBox txtToDate 
            Height          =   315
            Left            =   5130
            TabIndex        =   56
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdDate2 
            Caption         =   "..."
            Height          =   315
            Left            =   6630
            TabIndex        =   55
            Top             =   1080
            Width           =   315
         End
         Begin VB.OptionButton optName 
            Caption         =   "Name "
            Height          =   300
            Left            =   3780
            TabIndex        =   50
            Top             =   360
            Width           =   1980
         End
         Begin VB.OptionButton optAccID 
            Caption         =   "Account No"
            Height          =   300
            Left            =   195
            TabIndex        =   49
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   225
            Left            =   90
            TabIndex        =   51
            Top             =   1140
            Width           =   1485
         End
         Begin VB.Label lblDate2 
            Caption         =   "but before (dd/mm/yyyy)"
            Height          =   225
            Left            =   3255
            TabIndex        =   54
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Line Line3 
            X1              =   90
            X2              =   6960
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   450
         Left            =   6105
         TabIndex        =   60
         Top             =   5865
         Width           =   1215
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report"
         Height          =   3390
         Left            =   180
         TabIndex        =   72
         Top             =   240
         Width           =   7125
         Begin VB.OptionButton optCashBook 
            Caption         =   "Sub Cash Book"
            Height          =   285
            Left            =   180
            TabIndex        =   82
            Top             =   1500
            Width           =   3135
         End
         Begin VB.OptionButton optMonthlyBalance 
            Caption         =   "Monthly Balance"
            Height          =   285
            Left            =   3780
            TabIndex        =   47
            Top             =   1485
            Width           =   2265
         End
         Begin VB.OptionButton optDepGLedger 
            Caption         =   "Deposit General Ledger"
            Height          =   285
            Left            =   180
            TabIndex        =   45
            Top             =   2070
            Width           =   3135
         End
         Begin VB.OptionButton optMature 
            Caption         =   "Deposits that mature"
            Height          =   285
            Left            =   3780
            TabIndex        =   43
            Top             =   915
            Width           =   2715
         End
         Begin VB.OptionButton optSubDayBook 
            Caption         =   "Sub day Book"
            Height          =   285
            Left            =   180
            TabIndex        =   41
            Top             =   930
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton optLiabilities 
            Caption         =   "Liabilities"
            Height          =   255
            Left            =   3780
            TabIndex        =   44
            Top             =   360
            Width           =   2370
         End
         Begin VB.OptionButton optOpened 
            Caption         =   "Deposits opened"
            Height          =   255
            Left            =   3780
            TabIndex        =   42
            Top             =   2070
            Width           =   2370
         End
         Begin VB.OptionButton optDepositBalance 
            Caption         =   " Deposit Balances As On "
            Height          =   285
            Left            =   180
            TabIndex        =   40
            Top             =   360
            Width           =   3135
         End
         Begin VB.OptionButton optClosed 
            Caption         =   "Deposits closed"
            Height          =   255
            Left            =   3780
            TabIndex        =   46
            Top             =   2670
            Width           =   2370
         End
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
      Left            =   2880
      TabIndex        =   103
      Top             =   7320
      Width           =   1635
   End
End
Attribute VB_Name = "frmRDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Interest As Double
Private m_AccID As Long
Private m_AccNum As String
Private m_AccClosed As Boolean
Private m_rstPassBook As Recordset
Private m_CustReg As New clsCustReg
Private m_Notes As New clsNotes

Private M_ModuleID As wisModules
Private m_retVar As Long
Private m_DepositType As Integer
Private m_mulipleDeposit As Boolean
Private m_FormLoaded As Boolean
Private m_DepositTypeName As String
Private m_DepositTypeNameEnglish As String
 
Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private m_clsRepOption As clsRepOption

Private m_frmRDJoint As frmJoint
Private m_JointCust(0 To 3) As Long
Private m_InstallmentRst As Recordset  'vinay
'Private m_AmountCounter As Integer       'vinay
Private m_Amount As Currency  'vinay

Const CTL_MARGIN = 15

Private m_accUpdatemode As Integer
Private m_MaturityDateUS As Date 'vinay
Private m_MaturityDateIND As Date 'vinay
Private m_ClosedOn As String

Public Event WindowClosed()
Public Event AccountChanged(ByVal AccId As Long)
Public Event ShowReport(ReportType As wis_RDReports, ReportOrder As wis_ReportOrder, _
            fromDate As String, toDate As String, _
            RepOption As clsRepOption)
Public Event SelectDeposit(ByRef DepositType As Integer, ByRef cancel As Boolean)

Public Property Let MultipleDeposit(NewValue As Boolean)
    m_mulipleDeposit = NewValue
    lblDepositTypeName.Visible = NewValue
    cmdDepositType.Visible = NewValue
End Property
Public Property Get IsFormLoaded() As Boolean
    IsFormLoaded = m_FormLoaded
End Property
Public Property Let DepositType(NewValue As Integer)
    
    If m_FormLoaded And m_DepositType <> NewValue Then
        m_DepositType = NewValue
        lblDepositTypeName.Visible = NewValue
        cmdDepositType.Visible = NewValue
        LoadDepositType (NewValue)
        'Call LoadSetupValues
        
    End If
End Property
Private Function AccountName(AccId As Long) As String

Dim Lret As Long
Dim rst As ADODB.Recordset
'Prelim checks
If AccId <= 0 Then Exit Function

'Query DB
    gDbTrans.SqlStmt = "SELECT AccID, Title + FirstName + space(1) + " _
            & "MiddleName + space(1) + Lastname AS Name FROM RDMaster, " _
            & "NameTab WHERE RDMaster.AccID = " & AccId _
            & " AND RDMaster.CustomerID = NameTab.CustomerID"

    Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
        
    If Lret = 1 Then AccountName = FormatField(rst("Name"))
    Set rst = Nothing
End Function

Private Function AccountSave() As Boolean
    Dim txtIndex As Byte
    Dim AccIndex As Byte
    Dim I As Integer
    Dim AccNum As String
    Dim rst As ADODB.Recordset

' Check for a valid Account number.
    AccIndex = GetIndex("AccNum")

    With txtData(AccIndex)

'See if acc no has been specified
        If Trim$(.Text) = "" Then
            'MsgBox "No Account number specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(525), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        AccNum = Trim$(.Text)

'See if account already exists if it is new
        
        If m_accUpdatemode = wis_INSERT Then
            gDbTrans.SqlStmt = "Select AccNum from RDMaster where AccNum = " & AddQuotes(AccNum, True)
            If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & m_DepositType
            If gDbTrans.Fetch(rst, adOpenForwardOnly) >= 1 Then
                'MsgBox "Account number " & .Text & "already exists." & vbCrLf & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(545) & vbCrLf & vbCrLf & GetResourceString(641), vbExclamation, gAppName & " - Error"
                Set rst = Nothing
                Exit Function
            End If
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
        NomineeSpecified = True
        If Trim$(.Text) = "" Then
            'MsgBox "Nominee name not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            If MsgBox(GetResourceString(558) & vbCrLf & GetResourceString(541), _
                    vbInformation + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            NomineeSpecified = False
        End If
    End With

' Check for nominee relationship.
        txtIndex = GetIndex("NomineeRelation")
        With txtData(txtIndex)
            If NomineeSpecified And Trim$(.Text) = "" Then
                'MsgBox "Specify nominee relationship.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(559), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
        End With
    
' Check For Installment Amount
        txtIndex = GetIndex("RDAmount")
        With txtData(txtIndex)
            If Trim$(.Text) = "" Then
                'MsgBox "Specify Installment Amount.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(642), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            
            If Not IsNumeric(Trim$(.Text)) Then
                MsgBox GetResourceString(642), _
                        vbInformation, wis_MESSAGE_TITLE
                    ActivateTextBox txtData(txtIndex)
                    GoTo Exit_Line
            End If
        End With
        
'  Check for Date Of Maturity
        
        If Not DateValidate(GetVal("MaturityDate"), "/", True) Then
             'MsgBox "Invalid create date specified !" & vbCrLf & "Please specify in DD/MM/YYYY format!", vbExclamation, gAppName & " - Error"
             MsgBox GetResourceString(501) & vbCrLf & GetResourceString(573), vbExclamation, gAppName & " - Error"
             txtIndex = GetIndex("MaturityDate")
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If

'Check For No Of Months
        txtIndex = GetIndex("NoOfInst")
        With txtData(txtIndex)
            If Trim$(.Text) = "" Then
                'MsgBox "Specify No Of Months.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(643), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            If Not IsNumeric(Trim$(.Text)) Then
                'MsgBox "Specify No Of Months In Numbers.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(644), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
        End With
    
'Check for Rate Of Interest
        txtIndex = GetIndex("RateOfInterest")
        With txtData(txtIndex)
            If Trim$(.Text) = "" Then
                'MsgBox "Specify Rate Of Interest.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(505), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            
            If Not IsNumeric(Trim$(.Text)) Then
                'MsgBox "Specify Rate Of Interest Real Numbers.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(646), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            
            If Val(Trim$(.Text)) > 20 Then
                MsgBox "Specify a decent rate of interest", _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If

        End With
        
' Check if an introducerID has been specified.
        txtIndex = GetIndex("IntroducerID")
        With txtData(txtIndex)
            If Trim$(.Text) = "" Then
                'If MsgBox("No introducer has been specified!" _
                    & vbCrLf & "Add this Account anyway?", vbQuestion + vbYesNo) = vbNo Then
                If MsgBox(GetResourceString(560) _
                    & vbCrLf & "Add this Account anyway?", vbQuestion + vbYesNo) = vbNo Then
                    GoTo Exit_Line
                End If
            Else
            ' Check if the introducer exists.
                gDbTrans.SqlStmt = "SELECT CustomerID FROM NAmeTab WHERE CustomerID = " & Val(.Text)
                If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
                    'MsgBox "The introducer account number " & .Text & " is invalid.", _
                            vbExclamation, wis_MESSAGE_TITLE
                    MsgBox GetResourceString(514), _
                            vbExclamation, wis_MESSAGE_TITLE
                    ActivateTextBox txtData(txtIndex)
                    GoTo Exit_Line
                End If

            'Check if accnos clash
                If Val(txtData(AccIndex).Text) = Val(.Text) Then
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
            Debug.Print "Kannada"
            MsgBox "You have not selected the Account Group", vbInformation, wis_MESSAGE_TITLE
            txtIndex = GetIndex("AccGroup")
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If

    'Confirm before proceeding
        Dim AccId As Long
        If m_accUpdatemode = wis_UPDATE Then
            'If MsgBox("This will update the account " & GetVal("AccID") _
                    & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
            If MsgBox(GetResourceString(520) & GetVal("AccID") _
                    & "." & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
                GoTo Exit_Line
            End If
            AccId = m_AccID
            If Not m_frmRDJoint Is Nothing Then
              With m_frmRDJoint
                For I = 0 To 3
                    If m_JointCust(I) <> .JointCustId(I) Then
                        If MsgBox("Joint Account holder have changed " & vbCrLf & _
                            " Do you want to continue?", vbQuestion + vbYesNo, _
                            wis_MESSAGE_TITLE) = vbNo Then Exit Function
                    End If
                Next
              End With
            End If
        
        ElseIf m_accUpdatemode = wis_INSERT Then
            'If MsgBox("This will create a new account with an account number " & GetVal("AccID") _
                    & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
            If MsgBox(GetResourceString(540) & GetVal("AccNum") _
                    & "." & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
                GoTo Exit_Line
            End If
            'Get the new account number
            gDbTrans.SqlStmt = "SELECt MAx(AccID) From RDMAster "
            AccId = 1
            If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then AccId = FormatField(rst(0)) + 1
        End If

'Start Transactions to Data base
        gDbTrans.BeginTrans
        If Not m_CustReg.SaveCustomer Then
            'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
            MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
            gDbTrans.RollBack
            Exit Function
        End If
    'Save the Joint Account Holders
        If Not m_frmRDJoint Is Nothing Then m_frmRDJoint.SaveJointCustomers
        Dim cumulativeID As Byte
        I = GetIndex("Cumulative")
        I = Val(ExtractToken(txtPrompt(I).Tag, "TextIndex"))
        cumulativeID = IIf(cmb(I).ListIndex = -1, 0, cmb(I).ItemData(cmb(I).ListIndex))
        
        gDbTrans.BeginTrans
        ' Insert/update to database.
        If m_accUpdatemode = wis_INSERT Then
            
            gDbTrans.SqlStmt = "Insert into RDMaster (AccID,AccNum, CustomerID," & _
                " CreateDate, NomineeID,NomineeName,NomineeRelation, IntroducerID, " & _
                " LedgerNo, FolioNo, InstallmentAmount," & _
                " NoOfInstallments, RateOfInterest,MaturityDate,AccGroupId,Cumulative,UserID,DepositType)" & _
                " values (" & AccId & "," & _
                AddQuotes(AccNum, True) & "," & _
                m_CustReg.CustomerID & "," & _
                "#" & GetSysFormatDate(GetVal("CreateDate")) & "#," & _
                Val(GetVal("NomineeID")) & ", " & _
                AddQuotes(GetVal("NomineeName"), True) & "," & _
                AddQuotes(GetVal("NomineeRelation"), True) & ", " & _
                Val(GetVal("IntroducerID")) & "," & _
                AddQuotes(GetVal("LedgerNo"), True) & ", " & _
                AddQuotes(GetVal("FolioNo"), True) & ", " & _
                CCur(GetVal("RDAmount")) & ", " & _
                Val(GetVal("NoOfInst")) & ", " & _
                GetVal("RateOfInterest") & ", " & _
                "#" & GetSysFormatDate(GetVal("MaturityDate")) & "# ," & _
                GetAccGroupID & "," & cumulativeID & "," & gUserID & "," & m_DepositType & " )"
            'Insert/update the data.
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                GoTo Exit_Line
            End If
        
            'Insert the data to the joint table
            I = 0
            Dim count As Integer
            If Not m_frmRDJoint Is Nothing Then
                For count = 0 To m_frmRDJoint.txtName.count - 1
                    m_JointCust(count) = m_frmRDJoint.JointCustId(count)
                    If m_JointCust(count) = 0 Then Exit For
                    gDbTrans.SqlStmt = "Insert into RDJoint " & _
                        "(AccNum,CustomerID, CustomerNum)" _
                        & " values (" & AddQuotes(GetVal("AccNum"), True) & "," & _
                        m_JointCust(count) & "," & _
                        count + 1 & _
                        ")"
                    If Not gDbTrans.SQLExecute Then
                        gDbTrans.RollBack
                        GoTo Exit_Line
                    End If
                Next
            End If
            
        ElseIf m_accUpdatemode = wis_UPDATE Then
            ' The user has selected updation.
            ' Build the SQL update statement.
            
            gDbTrans.SqlStmt = "Update RDMaster set " & _
                " NomineeId = " & Val(GetVal("NomineeID")) & ", " & _
                " NomineeName = " & AddQuotes(GetVal("NomineeName"), True) & ", " & _
                " NomineeRelation = " & AddQuotes(GetVal("NomineeRelation"), True) & ", " & _
                " IntroducerId = " & Val(GetVal("IntroducerID")) & "," & _
                " LedgerNo = " & AddQuotes(GetVal("LedgerNo"), True) & "," & _
                " FolioNo = " & AddQuotes(GetVal("FolioNo"), True) & ", " & _
                " CreateDate = #" & GetSysFormatDate(GetVal("CreateDate")) & "#," & _
                " InstallmentAmount = " & CCur(GetVal("RDAmount")) & "," & _
                " NoofInstallments = " & Val(GetVal("NoOfInst")) & "," & _
                " RateOfInterest = " & GetVal("RateOfInterest") & "," & _
                " AccGroupID = " & GetAccGroupID & "," & _
                " MaturityDate = " & "#" & GetSysFormatDate(GetVal("MaturityDate")) & "#" & _
                " where AccID = " & m_AccID
                
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                GoTo Exit_Line
            End If
            
            'm_JointCustID(0) = m_CustReg.CustomerId
            If Not m_frmRDJoint Is Nothing Then
              For count = 0 To 3
                If m_JointCust(count) <> m_frmRDJoint.JointCustId(count) Then
                    'If MsgBox("You have changed the joint account holder" & _
                        vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, _
                        wis_MESSAGE_TITLE) = vbNo Then Exit Function
                    If MsgBox(GetResourceString(675) & vbCrLf & _
                        GetResourceString(541), vbQuestion + vbYesNo, _
                        wis_MESSAGE_TITLE) = vbNo Then Exit Function
                    If m_frmRDJoint.JointCustId(count) = 0 Then 'Delte the Joint account details
                        gDbTrans.SqlStmt = "DELETE * FROM RDJoint WHERE AccNum = " & _
                            AddQuotes(AccNum, True) & _
                            " AND CustomerNum >= " & count + 1
                        'gDBTrans.SQLStmt = SqlStr
                        If Not gDbTrans.SQLExecute Then
                            gDbTrans.RollBack
                            GoTo Exit_Line
                        End If
                        m_JointCust(count) = 0
                    End If
                End If
                If m_frmRDJoint.JointCustId(count) = 0 Then Exit For
                If m_JointCust(count) = 0 Then 'Insert the new record
                    gDbTrans.SqlStmt = "Insert into RDJoint (AccNum,CustomerID, CustomerNum) " _
                            & "values (" & AddQuotes(AccNum, True) & "," & _
                            m_frmRDJoint.JointCustId(count) & "," & _
                            count + 1 & _
                            ")"
                Else 'Update the existing record
                    gDbTrans.SqlStmt = "UPDATE RDJoint Set CustomerID = " & _
                        m_frmRDJoint.JointCustId(count) & _
                        " WHERE AccNum = " & AddQuotes(AccNum, True) & _
                        " AND CustomerNum = " & count + 1
                End If
                If Not gDbTrans.SQLExecute Then
                    gDbTrans.RollBack
                    GoTo Exit_Line
                End If
              Next
            End If
        End If

    
        'MsgBox "Saved the account details.", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
        AccountSave = True
        gDbTrans.CommitTrans
    
        m_Interest = GetVal("RateOfInterest")
Exit_Line:
        Set rst = Nothing
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
'Prelim check
Dim Trans As Byte
Dim TransDate As Date

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
    'If MsgBox("Account has been closed previously." & vbCrLf & _
            "This action will reopen the account." & vbCrLf & _
            "Do you want to continue ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
    If MsgBox(GetResourceString(524) & vbCrLf & _
            GetResourceString(548) & vbCrLf & _
            GetResourceString(541), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
        Exit Function
    End If
    'Reopen the account first
    'Msgbox
    If Not AccountReopen(m_AccID) Then
        'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
End If
        
        'Get last transaction record
Dim Amount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim TransID As Long
Dim AmountTrans As Boolean
Dim IntTrans As Boolean
Dim PayableTrans As Boolean
Dim rst As ADODB.Recordset

Dim HeadID As Long
Dim IntHeadID As Long
Dim PayableID As Long


HeadID = GetHeadID(m_DepositTypeName, parMemberDeposit)
        
'Get the MAx Transid
TransID = GetRDMaxTransID(m_AccID)

gDbTrans.SqlStmt = "Select * from RDTrans " & _
                " where AccID = " & m_AccID & _
                " And TransID = " & TransID

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    AmountTrans = True
    Amount = FormatField(rst("Amount"))
    Trans = FormatField(rst("TransType"))
    TransDate = rst("TransDate")
End If


gDbTrans.SqlStmt = "Select * from RDIntTrans " & _
                " Where AccID = " & m_AccID & _
                " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    IntAmount = FormatField(rst("Amount"))
    TransDate = rst("TransDate")
    IntTrans = True
    IntHeadID = GetHeadID(m_DepositTypeName & " " & GetResourceString(487), parMemDepIntPaid)
End If

gDbTrans.SqlStmt = "Select TOP 1 * from RDIntPayable " & _
                " where AccID = " & m_AccID & _
                " And TransID = " & TransID

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    PayableAmount = FormatField(rst("Amount"))
    TransDate = rst("TransDate")
    PayableTrans = True
    PayableID = GetHeadID(m_DepositTypeName & " " & GetResourceString(375, 47), parDepositIntProv)
    If PayableTrans And (rst("Transtype") = wDeposit Or rst("TransType") = wContraDeposit) Then
        MsgBox "This will undo the INterest payble transaction", , wis_MESSAGE_TITLE
    End If
End If

'Confirm UNDO
'If MsgBox("Are you sure you want to undo the last transaction of Rs." & Amount & "?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(627) & Amount & "?", _
        vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then Exit Function

If Trans = wContraDeposit Or Trans = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & HeadID & _
            " And Accid = " & m_AccID & " And TransID = " & TransID
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
    gDbTrans.BeginTrans

    'Un do the amount
    gDbTrans.SqlStmt = "Delete from RDTrans where " & _
            " AccID = " & m_AccID & " and TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    'Un do the interest paid/received
    gDbTrans.SqlStmt = "Delete from RDIntTrans where " & _
                " AccID = " & m_AccID & " and TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    'Un do the interest Payable Withdraw or deposit if any
    gDbTrans.SqlStmt = "Delete from RDIntPayable where " & _
                " AccID = " & m_AccID & " and TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
Dim ClsBank As clsBankAcc

Set ClsBank = New clsBankAcc
    
'Delete the Last Transaction in the BankHeads
If Trans = wDeposit Or Trans = wContraDeposit Then
    If Amount Then Call ClsBank.UndoCashDeposits(HeadID, Amount, TransDate)
    If IntAmount Then Call ClsBank.UndoCashDeposits(IntHeadID, IntAmount, TransDate)
    If PayableAmount Then Call ClsBank.UndoCashDeposits(PayableID, PayableAmount, TransDate)
Else
    If Amount Then Call ClsBank.UndoCashWithdrawls(HeadID, Amount, TransDate)
    If IntAmount Then Call ClsBank.UndoCashWithdrawls(IntHeadID, IntAmount, TransDate)
    If PayableAmount Then Call ClsBank.UndoCashWithdrawls(PayableID, PayableAmount, TransDate)
End If
    
gDbTrans.CommitTrans
Set rst = Nothing
AccountUndoLastTransaction = True


End Function
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
' If the height is greater than viewport height, the scrollbar needs to be displayed.
'So,reduce the width accordingly.

        If .Height > picViewport.ScaleHeight Then
            NeedsScrollbar = True
            .Width = picViewport.ScaleWidth - VScroll1.Width
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

' Need to adjust the width of text boxes, due to change in width of the slider box.

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
Private Function AccountClose() As Boolean
'MsgBox "Wrong Code, Test it immediatly"
MsgBox "Wrong Code, Test it immediatly"

'Prelim checks
    
    If m_AccID <= 0 Then Exit Function
    
'Check if account exists
    If Not AccountExists(m_AccID) Then Exit Function
    
    Dim ret As Integer
    Dim AccNo As Long
    AccNo = m_AccID
'Check date format
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Invalid date specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Close the account
    
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Update RDMaster set " & _
        " ClosedDate = #" & GetSysFormatDate(Trim$(txtDate.Text)) & "#" & _
        " where AccID = " & AccNo
    
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
gDbTrans.CommitTrans
AccountClose = True

End Function
Private Function AccountDelete(AccId As Long) As Boolean
Dim rst As ADODB.Recordset
'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & AccId
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
    gDbTrans.SqlStmt = "Select TOP 1 * from RDTrans where AccID = " & AccId
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        'MsgBox "You cannot delete an account having transactions !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    Dim AccNum As String
    gDbTrans.SqlStmt = "SELECT AccNum From RDMAster WHERE AccID = " & AccId
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    AccNum = FormatField(rst("AccNum"))
    
'Delete account from DB
    gDbTrans.BeginTrans
    'First Delete The Joint account Information
    gDbTrans.SqlStmt = "Delete from RDJoint where AccNum = " & AddQuotes(AccNum, True)
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.SqlStmt = "Delete from RDMaster where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    gDbTrans.CommitTrans
    Set rst = Nothing
    AccountDelete = True

End Function
Private Function CheckInstallment(InstallmentNo As Integer) As Boolean

'----------------------------------------------------------------
'Function has false value when still installments are to be paid
'It has true value when all rd installments are paid
'----------------------------------------------------------------

    Dim Trans As wisTransactionTypes
    Dim ContraTrans As wisTransactionTypes
    Dim InstTrack As Integer
    Dim rst As ADODB.Recordset

    Trans = wDeposit
    ContraTrans = wContraDeposit
    gDbTrans.SqlStmt = "Select * from RDTrans where AccID =" & m_AccID & _
                      " and (TransType= " & Trans & " OR Transtype= " & Trans & ")"
    InstTrack = gDbTrans.Fetch(rst, adOpenForwardOnly)
    Set rst = Nothing

    If InstTrack < 0 Then Exit Function
    
    If InstTrack <= InstallmentNo - 1 Then Exit Function
    CheckInstallment = True

End Function

'This Function Checks For the Installment due
'If a instllment is due this will checkfor the
'whether the Carges of the late installment has paid or not
'IF IT IS PAID IT RETURNS TRUE ELSE FALSE
Private Function PayChargesForLatePayament(AccId As Long) As Boolean

    Dim MonDiff As Integer
    Dim NoOfInstall As Integer
    Dim CreateDate As Date
    Dim rst As ADODB.Recordset
    Dim LastTransDate As Date

    Dim transType As wisTransactionTypes
    Dim ContraTrans As wisTransactionTypes
    
    transType = wDeposit
    ContraTrans = wContraDeposit
    gDbTrans.SqlStmt = "Select * From RDTrans Where ACCId = " & AccId & " and " & _
            " (TransType = " & transType & " OR TransType = " & ContraTrans & ")"
        
    NoOfInstall = gDbTrans.Fetch(rst, adOpenForwardOnly)

    If NoOfInstall < 1 Then Exit Function
    
    gDbTrans.SqlStmt = "Select CreateDate From RDMaster Where AccId = " & AccId

    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    
    CreateDate = rst(0)
    MonDiff = DateDiff("M", CreateDate, GetSysFormatDate(txtDate.Text))
    
    transType = wDeposit
    ContraTrans = wContraDeposit
    
    If MonDiff > NoOfInstall Then
        gDbTrans.SqlStmt = "Select TOP 1 Transdate from RDIntTrans " & _
            " where AccID = " & AccId & _
            " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
            " Order by TransDate desc "
            If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
                PayChargesForLatePayament = True
                Exit Function
            End If
        
        LastTransDate = rst("TransDate")
        If DateDiff("M", LastTransDate, GetSysFormatDate(txtDate.Text)) = 0 Then Exit Function
        
        PayChargesForLatePayament = True
    End If
    Set rst = Nothing


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


Private Function GetIndex(strDataSrc As String) As Integer

' Returns the index of the control bound to "strDatasrc".

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
Private Function GetMonthName(Abs_Pos As Integer) As String

Dim Duplicate As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim count As Integer
Dim Trans As wisTransactionTypes
Dim CreateDate As Integer
Trans = wDeposit
Set Duplicate = m_rstPassBook.Clone
    Duplicate.MoveLast
    Duplicate.MoveFirst
While Duplicate.AbsolutePosition <= m_rstPassBook.AbsolutePosition And Duplicate.EOF = False

    If Duplicate("Transtype") = Trans Then count = count + 1
       
    Duplicate.MoveNext

Wend
gDbTrans.SqlStmt = "Select CreateDate from RDmaster where Accid = " & m_AccID
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
GetMonthName = "---"
Else
CreateDate = Month(rst("CreateDate"))
count = (count + CreateDate - 1)

If count > 12 Then
        count = count Mod 12
End If

GetMonthName = GetMonthString(count)

End If

Set rst = Nothing

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

    gDbTrans.SqlStmt = "SELECT MAX(val(AccNum)) FROM RDMaster"
    If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where DepositType = " & m_DepositType
    
        If gDbTrans.Fetch(rst, adOpenForwardOnly) = 0 Then
            NewAccNo = 1
        Else
            NewAccNo = Val(FormatField(rst(0))) + 1
        End If
    Set rst = Nothing
    GetNewAccountNumber = NewAccNo

End Function
Private Function AccountTransaction() As Boolean

Dim AccountCloseFlag As Boolean
Dim rst As ADODB.Recordset
Dim Trans As wisTransactionTypes
Dim lstIndex As Byte

'In This case(RD) Only one transaction is not occuring.  i,e Withdrawals ---Siddu

''Get the Transaction Type
Dim IsIntTrans As Boolean

With cmbTrans
   lstIndex = .ListIndex
   Trans = wDeposit
   If .ListIndex = 1 Then IsIntTrans = True
End With

Dim TransDate As Date
TransDate = GetSysFormatDate(txtDate.Text)

Dim Amount As Currency
'Get the Amount to be Deposit
Amount = txtAmount


Dim InstallmentNo As Integer
Dim InstallmentAmount As Currency

gDbTrans.SqlStmt = "Select * from RDMaster where Accid = " & m_AccID
Call gDbTrans.Fetch(rst, adOpenDynamic)
If Not IsIntTrans Then
    TxtInstallmentNo.Locked = True
    InstallmentAmount = CCur(rst("InstallmentAmount"))
    InstallmentNo = rst("NoOfInstallments")
End If

' get instrument no
Dim VoucherNo As String
VoucherNo = txtCheque

'Get the Particulars
Dim Particulars As String
Particulars = Trim$(cmbParticulars.Text)

' see the transaction type if it is wcharges change the particulars
If cmbTrans.ListIndex = 1 Then Particulars = "By Charges"

'Get the Balance and new transid
Dim Balance As Currency
Dim TransID As Long
    
gDbTrans.SqlStmt = "Select TOP 1 * FROM " & _
            IIf(IsIntTrans, "RDIntTrans", "RDTRans") & _
            " WHERE AccID = " & m_AccID & " order by TransID desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then Balance = FormatField(rst("Balance"))

'consider The Depoit case
TransID = GetRDMaxTransID(m_AccID) + 1

If TransID = 1 Then
    Balance = Val(InputBox("Please enter a balance to start " & _
            " with as this account has not transaction performed", "Initial Balance", "0.00"))
    If Balance < 0 Then
        'MsgBox "Invalid initial balance specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(517), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
End If

'check for the opening balance  and deposit
If Balance = 0 And Val(txtAmount) = 0 Then
    MsgBox "Either Opening balance or deposit amount should be specified ", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Dim UserID As Long
Dim BankCls As clsBankAcc
'UserId = gcurruser.UserId

gDbTrans.BeginTrans
'Inserting of
    gDbTrans.SqlStmt = "Insert into " & _
            IIf(IsIntTrans, "RDIntTrans", "RDTrans") & _
            " ( AccID, TransID,TransDate, Amount, TransType, " & _
            " Balance, Particulars, VoucherNo,UserID ) " & _
            " values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            Trans & "," & _
            Balance + Amount & "," & _
            AddQuotes(Particulars, True) & "," & _
            AddQuotes(VoucherNo, True) & "," & _
            UserID & ")"
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
 Set BankCls = New clsBankAcc

Dim HeadID As Long
HeadID = BankCls.GetHeadIDCreated(m_DepositTypeName, m_DepositTypeNameEnglish, parMemberDeposit, 0, wis_RDAcc + m_DepositType)
If lstIndex = 1 Then _
    HeadID = BankCls.GetHeadIDCreated(m_DepositTypeName & " " & GetResourceString(487), _
            m_DepositTypeNameEnglish & " " & LoadResourceStringS(487), parMemDepIntPaid, 0, wis_RDAcc + m_DepositType)

'Perform the tranaction in the RD Head
If Not BankCls.UpdateCashDeposits(HeadID, Amount, TransDate) Then
    gDbTrans.RollBack
    Exit Function
End If

Set BankCls = Nothing
    
gDbTrans.CommitTrans
AccountTransaction = True

'Update the Particulars combo
'Read to part array
    
    Dim ParticularsArr() As String
    ReDim ParticularsArr(20)
    
'Read elements of combo to array
    Dim I As Integer
    For I = 0 To cmbParticulars.ListCount - 1
        ParticularsArr(I) = cmbParticulars.List(I)
    Next I
    
'Update last accessed elements
    Call UpdateLastAccessedElements(Trim$(cmbParticulars.Text), ParticularsArr, True)
    
'Write to
    cmbParticulars.Clear
    
    For I = 0 To UBound(ParticularsArr)
        If Trim$(ParticularsArr(I)) <> "" Then
            Call WriteToIniFile("Particulars", "Key" & I, ParticularsArr(I), App.Path & "\RDAcc.ini")
            cmbParticulars.AddItem ParticularsArr(I)
        End If
    
    Next I

End Function

'
' Returns the text value from a control array
' bound the field "FieldName".
'
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
    
    With grd
        .Clear: .Rows = 11: .FixedRows = 1: .FixedCols = 0
        .Row = 0
        .Col = 0: .Text = GetResourceString(37)
        .CellFontBold = True: .ColWidth(0) = 1150 ' "Date"
        .Col = 1: .Text = GetResourceString(39)
        .CellFontBold = True: .ColWidth(1) = 750  '"Particulars"
        .Col = 2: .Text = GetResourceString(57)
        .CellFontBold = True: .ColWidth(0) = 1150  '"Installment"
        .Col = 4: .Text = GetResourceString(277)
        .CellFontBold = True: .ColWidth(0) = 950  '"Debit"
        .Col = 3: .Text = GetResourceString(276)
        .CellFontBold = True: .ColWidth(0) = 950  '"Credit"
        .Col = 5: .Text = GetResourceString(42)
        .CellFontBold = True: .ColWidth(0) = 1050  '"Balance"
    End With

End Function
Private Sub LoadDepositType(Deptype As Integer)
    
    Dim rstDepType As Recordset
    Dim txtIndex  As Integer
    Dim cmbIndex As Integer

    gDbTrans.SqlStmt = "Select * From DepositTypeTab where ModuleID = " & wis_RDAcc & _
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
      
        'CLEAR all the controls
        ResetUserInterface
    Else
        m_DepositType = 0
        m_DepositTypeName = GetResourceString(424)
        m_DepositTypeNameEnglish = LoadResourceStringS(424)
    End If
    
End Sub

Private Function LoadPropSheet() As Boolean

    TabStrip.ZOrder 1
    TabStrip.Tabs(1).Selected = True
    lblDesc.BorderStyle = 0
    lblHeading.BorderStyle = 0
    lblOperation.Caption = GetResourceString(54)    ' "Operation Mode : <INSERT>"

' Read the data from RDAcc.PRP and load the relevant data.
' Check for the existence of the file.
    
Dim PropFile As String
PropFile = App.Path & "\RDAcc_" & gLangOffSet & ".PRP"
If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet Then
        PropFile = App.Path & "\RDAcckan.PRP"
    Else
        PropFile = App.Path & "\RDAcc.PRP"
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
                    .Text = ""
                End With
                
                txtData(CtlIndex).Enabled = False
    
            Case "EDITABLE"
' Add 4 spaces for indentation purposes.
                
                With txtPrompt(CtlIndex)
                    .Text = IIf(gLangOffSet, Space(2), Space(4))
                    .FontBold = False
                    .Enabled = True
                End With
                
                txtData(CtlIndex).Enabled = True
            
            Case Else
                
                'MsgBox "Unknown Property type encountered " _
                        & "in Property file!", vbCritical
                MsgBox GetResourceString(603) _
                        & "in Property file!", vbCritical
                Exit Function
    
        End Select
    
' Set the PROPERTIES for controls.
        
        With txtPrompt(CtlIndex)
            
            strRet = PutToken(strTag, "Visible", "True")
            .Tag = strRet
            .Text = .Text & ExtractToken(.Tag, "Prompt")
            
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
    
' Get the display type. If its a List or Browse, then load a combo or a cmd button.
        
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
                    .Width = txtData(I).Height
                    .Height = .Width
                    .Left = txtData(I).Left + txtData(I).Width - .Width
                    .Top = txtData(I).Top
                    .TabIndex = txtData(I).TabIndex + 1
                    .ZOrder 0
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
            
            
            Case "BOOLEAN"

            'Load a command button.
                
                If Not ChkLoaded Then
                    ChkLoaded = True
                Else
                    Load chk(chk.count)
                End If
                With chk(chk.count - 1)
                    .Width = txtData(I).Height
                    .Height = .Width
                    .Left = txtData(I).Left + txtData(I).Width - .Width
                    .Top = txtData(I).Top
                    .Caption = ""
                    .TabIndex = txtData(I).TabIndex + 1
                    .ZOrder 0
                    .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                    txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                            "TextIndex", CStr(cmd.count - 1))
                    txtData(I).Text = False
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
'Load the Recurring Deposit
Dim cmbIndex As Integer

    ' Find out the textbox bound to DepositType.
    I = GetIndex("Cumulative")
    ' Get the combobox index for this text.
    cmbIndex = ExtractToken(txtPrompt(I).Tag, "TextIndex")
    With cmb(cmbIndex)
        .Clear
        .AddItem "None"
        .ItemData(.newIndex) = 0
        .AddItem GetResourceString(414) 'Quarterly
        .ItemData(.newIndex) = Inst_Quartery
        .AddItem GetResourceString(432) '"Half Yearly" '
        .ItemData(.newIndex) = Inst_HalfYearly
        .AddItem GetResourceString(431) 'Yearly
        .ItemData(.newIndex) = Inst_Yearly
        .ListIndex = 0
        txtData(I).Text = .Text
    End With
  
' Show the current date wherever necessary.
    txtIndex = GetIndex("CreateDate")
    txtData(txtIndex).Text = gStrDate
    
' Set the default updation mode.
    m_accUpdatemode = wis_INSERT
    
    
End Function
Private Sub LoadInterestRates()

With grdInt
    .Clear
    .Cols = 3
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(311)
    .Col = 2: .Text = GetResourceString(186)
    .ColWidth(0) = 400
    .ColWidth(1) = 2500
    .ColWidth(2) = 700
    
    Dim I As Integer, MinI As Integer, MaxI As Integer
    Dim retstr As String, Prevstr As String
    Dim strPrevFrom As String
    Dim strKey As String
    Dim SetUp As New clsSetup
    Dim StrFrom As String, strTo As String
    
    optDays.Value = True
    MaxI = cmbFrom.ListCount - 1
    StrFrom = cmbFrom.List(0)
    For I = 0 To MaxI
        strKey = "DAYS" & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
        'strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
        retstr = SetUp.ReadSetupValue("DEPOSIT" & wisDeposit_RD, strKey, "")
        retstr = SetUp.ReadSetupValue("DEPOSIT" & wis_RDAcc + m_DepositType, strKey, retstr)
        If retstr = "" Then Exit For
        strTo = cmbTo.List(I)
        If Val(Prevstr) <> Val(retstr) Then
            If .Rows = .Row + 1 Then .Rows = .Rows + 17
            .Row = .Row + 1
            .Col = 0: .Text = .Row
            .Col = 1: .Text = GetFromDateString(StrFrom, strTo)
            .Col = 2: .Text = Val(retstr)
            strPrevFrom = StrFrom
            StrFrom = cmbTo.List(I)
        Else
            .Col = 1: .Text = GetFromDateString(strPrevFrom, strTo)
            .Col = 2: .Text = Val(retstr)
            StrFrom = cmbTo.List(I)
        End If
        Prevstr = Val(retstr)
    Next
    
    optMon.Value = True
    MaxI = cmbFrom.ListCount - 1
    'strFrom = cmbFrom.List(0)
    For I = 0 To MaxI
        strKey = "YEAR" & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
        'strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
        'retstr = SetUp.ReadSetupValue("DEPOSIT" & m_DepositType, strKey, "")
        retstr = SetUp.ReadSetupValue("DEPOSIT" & wisDeposit_PD, strKey, "")
        If retstr = "" Then Exit For
        If Val(Prevstr) <> Val(retstr) Then
            strTo = cmbTo.List(I)
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 0: .Text = .Row
            .Col = 1: .Text = GetFromDateString(StrFrom, strTo)
            .Col = 2: .Text = Val(retstr)
            StrFrom = cmbTo.List(I)
        End If
        Prevstr = Val(retstr)
    Next
    
End With

End Sub
Private Sub ResetUserInterface()

'get HeadID in the HeadsAccTrans Table(RDHeadID)
Dim ClsBank As clsBankAcc

If m_AccID = 0 And m_CustReg.CustomerID = 0 Then Exit Sub

'First the TAB 1
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
    With txtAmount
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    
    With cmbTrans
        .BackColor = wisGray: .Enabled = False
    End With
    
    With txtCheque
        .BackColor = wisGray: .Enabled = False
    End With
    With cmbParticulars
        .BackColor = wisGray: .Enabled = False
    End With
    
    With Me.rtfNote
        .BackColor = wisGray: .Enabled = False: .Text = GetResourceString(259)  ' "< No notes defined >"
        If gLangOffSet Then
            .Font.name = gFontName: .Font.Size = gFontSize
        Else
            .Font.Size = 10: .Font = "Arial"
        End If
    End With
    
    With cmdClose
        .Enabled = False
    End With
    
    With cmdAccept
        .Enabled = False
    End With
    
    With cmdUndo
        .Enabled = False
    End With
    
    
    With TxtInstallmentNo
        .Enabled = False: .BackColor = wisGray: .Text = ""
    End With
    
    Call PassBookPageInitialize
    
    cmdAddNote.Enabled = False
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    cmdClose.Enabled = False
'Now the Tab 2
    
    Dim I As Integer
    Dim strField As String
    Dim txtIndex As Integer

'Enable the reset (auto acc no generator button)
    
    cmd(0).Enabled = True
    Dim SetUp As New clsSetup
    For I = 0 To txtData.count - 1
        txtData(I).Text = ""

    'If its Createdate field, then put today's left.
        strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
        
        If StrComp(strField, "Cumulative", vbTextCompare) = 0 Then
            Dim cmbIndex As Integer
            cmbIndex = ExtractToken(txtPrompt(I).Tag, "TextIndex")
            cmb(cmbIndex).ListIndex = 0
            txtData(I).Text = cmb(cmbIndex).Text
        End If
        If StrComp(strField, "CreateDate", vbTextCompare) = 0 Then
            txtData(I).Text = gStrDate
        End If
    ' If It is Rate of Interest then read Value from Setup Tab
        If StrComp(strField, "RateOfInterest", vbTextCompare) = 0 Then
            txtData(I).Text = SetUp.ReadSetupValue("RDAcc", "Interest On RDDeposit", 18#)
        End If
    Next
    
    On Error Resume Next
    Unload m_frmRDJoint
    Set m_frmRDJoint = Nothing
    For I = 0 To 3
        m_JointCust(I) = 0
    Next
    
    lblOperation.Caption = GetResourceString(54)    '"Operation Mode : <INSERT>"
    txtIndex = GetIndex("AccNum")
    txtData(txtIndex).Text = GetNewAccountNumber
    txtData(txtIndex).Locked = False
    cmdTerminate.Enabled = False
    
'The form level variables
    
    m_accUpdatemode = wis_INSERT
    m_CustReg.NewCustomer
    Set m_CustReg = Nothing
    Set m_CustReg = New clsCustReg
    m_AccID = 0
    
    Set m_rstPassBook = Nothing

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
    Dim MonthStr As String
    Dim ChargeCounter As Integer
    Dim transType As wisTransactionTypes
    Dim CurPos As Integer
    
'Check if Rec Set has been set
    MonthStr = "---"
        If m_rstPassBook Is Nothing Then Exit Sub
        CurPos = m_rstPassBook.AbsolutePosition
        If CurPos < 0 Then CurPos = 0
'Show 10 records or till eof of the page being pointed to
        With grd
             Call PassBookPageInitialize
            .Visible = False
            I = 0
            Do
                If m_rstPassBook.EOF Then Exit Do
                MonthStr = GetMonthName(m_rstPassBook.AbsolutePosition)
                I = I + 1
                If I > 10 Then Exit Do
                .Row = I
                .Col = 0: .Text = FormatField(m_rstPassBook("TransDate"))
                .Col = 1: .Text = FormatField(m_rstPassBook("Particulars"))
                
                .Col = 2: .Text = MonthStr
                transType = m_rstPassBook("TransType")
                .Col = IIf(transType = wDeposit Or transType = wContraDeposit, 3, 4)
                .Text = FormatField(m_rstPassBook("Amount"))
                
                .Col = 5: .Text = FormatField(m_rstPassBook("Balance"))
                
                m_rstPassBook.MoveNext
                If m_rstPassBook.EOF Then Exit Do
            Loop
            .Visible = True
            .Row = 1
        End With
        
        'Now set the Next & Prev Trans Button
        cmdNextTrans.Enabled = Not m_rstPassBook.EOF
        cmdPrevTrans.Enabled = IIf(CurPos > 1, -1, 0)
End Sub

Private Function ValidControls() As Boolean
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

If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
    'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If

If cmbTrans.ListIndex = -1 Then
    'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(588), vbExclamation, gAppName & " - Error"
    cmbTrans.SetFocus
    Exit Function
End If

Dim TransDate As Date
TransDate = GetSysFormatDate(txtDate.Text)

'See if the date is earlier than last date of transaction
If DateDiff("D", TransDate, GetRDLastTransDate(m_AccID)) > 0 Then
    'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If
Dim rst As Recordset
'Check if the date of transaction is earlier than account opening date itself
gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & m_AccID
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 0 Then Exit Function
If DateDiff("D", TransDate, rst("CreateDate")) > 0 Then
    'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(568), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If
''Get the Transaction Type

If cmbTrans.ListIndex <> 1 Then
    If Val(txtAmount) <> CCur(rst("InstallmentAmount")) Then
        'MsgBox "Amount Should Be Equal To Installment Amount", _
                vbExclamation, gAppName & "-ERROR"
        If MsgBox(GetResourceString(648) & vbCrLf & GetResourceString(541), _
                vbInformation + vbYesNo, gAppName & "-ERROR") = vbNo Then
            txtAmount.SetFocus
            Exit Function
        End If
    End If
End If

If Not IsNumeric(txtCheque.Text) Or txtCheque = "" Then
    'MsgBox "Instrument  in Number Should be  specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(650), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtCheque
    Exit Function
End If

If Trim$(cmbParticulars.Text) = "" Then
    'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(621), vbExclamation, gAppName & " - Error"
    cmbParticulars.SetFocus
    Exit Function
End If


ValidControls = True

End Function

Private Function VisibleCount() As Integer

    ' Returns the number of items that are visible for a control array.
    ' Looks in the control's tag for visible property, rather than
    ' depend upon the control's visible property for some obvious reasons.
    
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


Private Sub chk_LostFocus(Index As Integer)
' Update the current text to the data text

    Dim txtIndex As String
    
    txtIndex = ExtractToken(chk(Index).Tag, "TextIndex")
    If txtIndex <> "" Then txtData(Val(txtIndex)).Text = chk(Index).Value

End Sub


Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub cmb_LostFocus(Index As Integer)

' Update the current text to the data text

    Dim txtIndex As String
    
    txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
    
    If txtIndex <> "" Then txtData(Val(txtIndex)).Text = cmb(Index).Text
    
End Sub


Private Sub cmbTrans_Click()

    If cmbTrans.ListCount = 0 Then
        MsgBox GetResourceString(608)
        Exit Sub
    End If
    
    If cmbTrans.ListIndex = 0 Then
        TxtInstallmentNo.Text = Installment(m_AccID)
        TxtInstallmentNo.Locked = True
        'txtAmount.Locked = True
        txtAmount.Text = m_Amount
    End If
    
    If cmbTrans.ListIndex = 1 Then
        'txtAmount.Locked = False
        TxtInstallmentNo.Locked = True
        txtAmount.Text = 0#
        TxtInstallmentNo.Text = 0
    End If
    
    
'A case of deposit
    
    If cmbTrans.ListIndex = 0 Then
        txtCheque.Visible = True
'        cmbCheque.Visible = False
'        cmdCheque.Visible = False
        Exit Sub
    End If

End Sub

'
Private Sub cmbTrans_LostFocus()
Dim rst As Recordset
Dim InstAmount As Currency
Dim lstIndex As Byte

'if it is null then exit
If cmbTrans.ListIndex = -1 Then Exit Sub

'Get the index
With cmbTrans
    lstIndex = .ListIndex
End With

'Get the defualt instalment amount here
gDbTrans.SqlStmt = "SELECT * from RDMASTER where AccNum =" & AddQuotes(m_AccNum, True)
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & m_DepositType

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine

'Instalment amount from
InstAmount = FormatCurrency(rst("InstallmentAmount"))

If lstIndex = 0 Then txtAmount = InstAmount

Exit Sub

ErrLine:
        MsgBox "Instalmment Amount : " & Chr(13) + Chr(10) & Err.Description, vbInformation, wis_MESSAGE_TITLE

End Sub

'
Private Sub cmd_Click(Index As Integer)
    Dim txtIndex As String
    Dim I As Integer
    Dim rst As ADODB.Recordset

txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")
    Select Case UCase$(strField)
       
        Case "ACCID", "ACCNUM"
            If m_accUpdatemode = wis_INSERT Then
                txtData(txtIndex).Text = GetNewAccountNumber
            End If
       
        Case "ACCNAME"
            m_CustReg.ModuleID = wis_RDAcc
            m_CustReg.ShowDialog
            If m_CustReg.CustomerID <> 0 Then _
                txtData(txtIndex).Text = m_CustReg.FullName
            
        Case "JOINTHOLDER"
            If m_frmRDJoint Is Nothing Then Set m_frmRDJoint = New frmJoint
            With m_frmRDJoint
                .Left = Me.Left + picViewport.Left + _
                    txtData(txtIndex).Left + fraNew.Left + CTL_MARGIN
                .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                    + fraNew.Top + 300 - Me.VScroll1.Value
                .Show 1
                'If .Status = "OK" Then txtData(txtIndex).Text = .JointHolders
                txtData(Val(GetIndex("JointHolder"))).Text = .JointHolders
            End With
            
        Case "MATURITYDATE"
            With Calendar
            .Left = txtData(txtIndex).Left + Me.Left _
                    + Me.picViewport.Left + fraNew.Left + 50
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fraNew.Top + 300
            .Width = txtData(txtIndex).Width
            If .Top + .Height > Screen.Height Then .Top = .Top - .Height - txtData(txtIndex).Height
            .Height = .Width
            '.selDate = iif(DateValidate(txtData(txtIndex).Text,"/",True),get
            .Show vbModal, Me
            If .selDate <> "" Then txtData(txtIndex).Text = .selDate
        End With
    
        Case "CREATEDATE"
          With Calendar
            .Left = txtData(txtIndex).Left + Me.Left _
                    + Me.picViewport.Left + fraNew.Left + 50
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fraNew.Top + 300
            .Width = txtData(txtIndex).Width
            If .Top + .Height > Screen.Height Then _
                .Top = .Top - .Height - txtData(txtIndex).Height
            .Height = .Width
            .selDate = txtData(txtIndex).Text
            .Show vbModal, Me
            If .selDate <> "" Then txtData(txtIndex).Text = .selDate
          End With
    
        Case "INTRODUCERNAME", "NOMINEENAME"

        ' Build a query for getting introducer details.
        ' If an account number specified, exclude it from the list.
            
            gDbTrans.SqlStmt = "SELECT CustomerID as [Cust No], " _
                        & " Title +' '+ FirstName + ' ' + Middlename " _
                        & "+ ' ' + LastName as Name, HomeAddress, " _
                        & "OfficeAddress FROM NameTab WHERE " _
                        & " CustomerID <> " & m_CustReg.CustomerID
            
            Dim strSearch As String
            strSearch = InputBox("Enter name the customer, You are seaching", wis_MESSAGE_TITLE)
            If Trim(strSearch) <> "" Then
                gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND " & _
                    "( FirstName like '" & strSearch & "' OR " & _
                    " MiddleName like '" & strSearch & "' OR " & _
                    " LastName like '" & strSearch & "' )"
            End If
            If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
                'MsgBox "No accounts present!", vbExclamation, wis_MESSAGE_TITLE
                MsgBox GetResourceString(525), vbExclamation, wis_MESSAGE_TITLE
                Exit Sub
            End If

        'Fill the details to report dialog and display it.
            If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
            
            If Not FillView(m_frmLookUp.lvwReport, rst) Then
                'MsgBox "Error loading introducer accounts.", _
                        vbCritical, wis_MESSAGE_TITLE
                MsgBox GetResourceString(562), _
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
                .m_SelItem = ""
                .Show vbModal, Me
                If .m_SelItem <> "" Then
                    txtData(txtIndex - 1).Text = .lvwReport.SelectedItem.Text
                    txtData(txtIndex).Text = .lvwReport.SelectedItem.SubItems(1)
                End If
            End With
    End Select
End Sub

Private Sub cmdApply_Click()

'    Dim RdDepositInterest As Single
'    Dim RDLoanInterest As Single
'
'    If Not IsNumeric(Txt_int_on_rd_dep_tab4.Text) Then
'        'MsgBox "Enter A Valid Interest Value ", vbInformation, gAppName & "ERROR !! "
'        MsgBox GetResourceString(518), vbInformation, gAppName & "ERROR !! "
'        ActivateTextBox Txt_int_on_rd_dep_tab4
'        Exit Sub
'    End If
'
'    If Val(Txt_int_on_rd_dep_tab4.Text) > 20 Then
'        MsgBox "Specify Decent Rate of Interest", vbInformation, gAppName & "ERROR !! "
'        ActivateTextBox Txt_int_on_rd_dep_tab4
'        Exit Sub
'    End If
'
'    If Not IsNumeric(Me.Txt_int_on_rd_loan_tab4.Text) Then
'        'MsgBox "Enter A Valid Interest Value ", vbInformation, gAppName & "ERROR !! "
'        MsgBox GetResourceString(518), vbInformation, gAppName & "ERROR !! "
'        ActivateTextBox Txt_int_on_rd_loan_tab4
'        Exit Sub
'    End If
'
'    RdDepositInterest = Val(Me.Txt_int_on_rd_dep_tab4.Text)
'    RDLoanInterest = Val(Me.Txt_int_on_rd_loan_tab4.Text)
'
'    Dim clsSetup As New clsSetup
'    Dim ClsInt As New clsInterest
'    Call clsSetup.WriteSetupValue("RDAcc", "Interest On RDDeposit", CStr(RdDepositInterest))
'    Call clsSetup.WriteSetupValue("RDAcc", "Interest On RDLoan", CStr(RDLoanInterest))
'    Set clsSetup = Nothing
'    'Write to INterest Class
'    Call ClsInt.SaveInterest(wis_RDAcc, "Deposit", RdDepositInterest)
'    Call ClsInt.SaveInterest(wis_RDAcc, "Loan", RDLoanInterest)
'    Set ClsInt = Nothing
End Sub

Private Sub cmdAccept_Click()


If Not ValidControls Then Exit Sub
If Not AccountTransaction() Then Exit Sub

'Reload the account
If Not AccountLoad(m_AccID) Then Exit Sub

'Display the pass book
TabStrip2.Tabs(2).Selected = True

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

Private Sub cmdClose_Click()
    frmRDClose.p_AccID = m_AccID
    
    frmRDClose.DepositType = m_DepositType
    frmRDClose.DepositName = m_DepositTypeName
    frmRDClose.DepositNameEnglish = m_DepositTypeNameEnglish
    frmRDClose.Show vbModal
    AccountLoad (m_AccID)
End Sub

Private Sub cmdDate_Click()
With Calendar
    .Left = Left + fraNew.Left + cmdDate.Left '- .Width / 2
    .Top = Top + Me.fraNew.Top + cmdDate.Top / 2
    If DateValidate(Me.txtDate.Text, "/", True) Then
        .selDate = txtDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtDate.Text = .selDate
End With

End Sub

Private Sub cmdDate1_Click()
With Calendar
    .Left = Me.Left + Me.fraReports.Left + fraOrder.Left + cmdDate1.Left - .Width / 2
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdDate1.Top
    If DateValidate(txtFromDate.Text, "/", True) Then
        .selDate = txtFromDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtFromDate.Text = .selDate
End With
End Sub

Private Sub cmdDate2_Click()
With Calendar
    .Left = Me.Left + fraReports.Left + Me.fraOrder.Left + cmdDate2.Left - .Width / 2
    .Top = Me.Top + fraReports.Top + Me.fraOrder.Top + cmdDate2.Top
    If DateValidate(Me.txtToDate.Text, "/", True) Then
        .selDate = txtToDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtToDate.Text = .selDate
End With

End Sub


Private Sub cmdDepositType_Click()
   Dim cancel As Boolean
    Dim DepositType As Integer
    RaiseEvent SelectDeposit(DepositType, cancel)
    If Not cancel And DepositType <> m_DepositType Then txtAccNo = "": txtData(GetIndex("AccNum")).Text = GetNewAccountNumber
End Sub

Private Sub cmdIntApply_Click()

If cmbFrom.ListIndex < 0 Then Exit Sub
If cmbTo.ListIndex < 0 Then Exit Sub
If cmbFrom.ListIndex > cmbTo.ListIndex Then Exit Sub

If Not DateValidate(txtIntDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    'Invalid date specifid
    ActivateTextBox txtIntDate
    Exit Sub
End If


Dim strKey As String
Dim TransDate As Date
TransDate = GetSysFirstDate(txtIntDate)

strKey = IIf(optDays, "DAYS", "MNTH")

Dim FromIndex As Integer
Dim ToIndex As Integer
Dim I As Integer

'intBegin = cmbFrom.ItemData(cmbFrom.ListIndex)
'intEnd = cmbTo.ItemData(cmbTo.ListIndex)
FromIndex = cmbFrom.ListIndex
ToIndex = cmbTo.ListIndex

Dim SetUp As New clsSetup
Dim strModule As String
Dim strValue As String
Dim strDef As String


'strModule = "DEPOSIT" & m_DepositType
strModule = "DEPOSIT" & (wis_RDAcc + m_DepositType)
strDef = IIf(optDays, "DAYS", "YEAR")

'First check whether he has enter the previous slab interest rates or not
'if he has not entered the previous slab interest rates
'then enter the same rate for thse slabs

For I = 0 To FromIndex - 1
    strKey = strDef & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
    'strValue = Setup.ReadSetupValue(strModule, strKey, "")
    strValue = GetInterestRateOnDate(M_ModuleID, strKey, TransDate)
    If Len(strValue) = 0 Then
        'strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
        Call SetUp.WriteSetupValue(strModule, strKey, strValue)
        Call SaveInterest(M_ModuleID, strKey, _
                Val(txtGenInt), Val(txtEmpInt), Val(txtSenInt), TransDate)
    
    End If
Next

'Enter the Deatils of the slab interest rate
For I = FromIndex To ToIndex
    strKey = strDef & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
    strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
    Call SetUp.WriteSetupValue(strModule, strKey, strValue)
    Call SaveInterest(M_ModuleID, strKey, _
                Val(txtGenInt), Val(txtEmpInt), Val(txtSenInt), TransDate)
Next

'Then check whether he has enter the next slab interest rates or not
'if he has not entered the interest rates
'then enter the same rate for the next slabs also
FromIndex = ToIndex + 1
ToIndex = cmbTo.ListCount - 1
For I = FromIndex To ToIndex
    strKey = strDef & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
    strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
    Call SetUp.WriteSetupValue(strModule, strKey, strValue)
Next

Call LoadInterestRates
cmdIntApply.Enabled = False
End Sub

Private Sub cmdIntPayable_Click()
If Not DateValidate(txtIntPayable.Text, "/", True) Then
'''    MsgBox "Invalid Date Format Specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntPayable
    Exit Sub
End If
Me.Refresh
Call AddInterestPayableOfRD(txtIntPayable.Text)
End Sub

Private Function AddInterestPayableOfRD(OnIndianDate As String) As Boolean

Dim DimPos As Integer
Dim AsOnDate As Date
DimPos = InStr(1, OnIndianDate, "31/3/", vbTextCompare)
If DimPos = 0 Then DimPos = InStr(1, OnIndianDate, "31/03/", vbTextCompare)
If DimPos = 0 Then
    'MsgBox "Unable to perform the transactions", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

AsOnDate = GetSysFormatDate(OnIndianDate)

On Error GoTo ErrLine
  'declare the variables necessary
  
Dim transType As wisTransactionTypes
Dim rstMain As ADODB.Recordset
Dim rstPayble As ADODB.Recordset
Dim rstInt As ADODB.Recordset
Dim InterestRate As Currency
Dim LastIntDate As Date
Dim CreateDate As Date
Dim MatDate As Date
Dim Duration As Long
Dim IntAmount As Currency
Dim UserID As Long

Dim AccId As Long
Dim PayableBalance As Currency
Dim TransID As Long

'Dim BankClass As New clsBankAcc

Dim count As Integer
Dim IntPayble As Currency
Dim TotalIntPayble As Currency
Dim TotalIntAmount As Currency
Dim rst As ADODB.Recordset

AsOnDate = GetSysFormatDate(OnIndianDate)

'Before Adding check whether he has already added the amount
'TransType = wContraInterest
gDbTrans.SqlStmt = "Select * from RDIntPayable Where TransDate = #" & AsOnDate & "# "
If m_DepositType > 0 Then
    gDbTrans.SqlStmt = "Select top 10 * from RDIntPayable Where AccID in " & _
            " (Select distinct AccID from RDMaster where DepositType = " & m_DepositType & " ) " & _
        " And TransDate = #" & AsOnDate & "# "
End If
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    MsgBox "Interest Payable already added to the Accounts", vbInformation, wis_MESSAGE_TITLE
    Set rst = Nothing
    Exit Function
End If

'Build The Querry
gDbTrans.SqlStmt = "SELECT Title+' '+FirstName+' '+MiddleName +' '+LastName As CustNAme," & _
        " Balance, B.TransID,A.AccId,A.AccNum, CreateDate, MaturityDate, TransDate," & _
        " TransType,RateOfInterest From " & _
        " RDMaster A, RDTrans B, NameTab C Where " & _
        " A.AccId = B.AccId And A.CustomerID = C.CustomerID" & _
        " AND Balance <> 0 " & _
        " And (ClosedDate Is NULL Or CreateDate < #" & AsOnDate & "#) And TransID =  " & _
        " (Select Max(TransID) From RDTrans D Where A.AccId = D.AccId ) "
        If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And A.DepositType = " & m_DepositType
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order By val(A.AccNum)"

'lblStatus.Caption = "Computing Interests for
lblStatus.Caption = GetResourceString(906) & "  ............"
txtFailAccIDs = ""

count = gDbTrans.Fetch(rstMain, adOpenStatic)
If count < 1 Then GoTo ExitLine

gDbTrans.SqlStmt = "Select Balance, A.AccId,AccNum,TransId From  " & _
    " RDIntTrans A, RDMaster B Where B.accID = A.AccID " & _
    " AND TransID = (SELECT MAx(TransID) From RDIntTrans C " & _
        " WHERE C.AccId = B.AccID ) "
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.DepositType = " & m_DepositType
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Val(B.AccNum)"

Set rstInt = Nothing
Call gDbTrans.Fetch(rstInt, adOpenStatic)

gDbTrans.SqlStmt = "Select Balance As Payble, A.AccId,AccNum,TransId From  " & _
    " RDIntPayable A, RDMaster B Where B.accID = A.AccID " & _
    " AND TransID = (SELECT MAx(TransID) From RDIntPayable C " & _
        " WHERE C.AccId = B.AccID ) "
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.DepositType = " & m_DepositType
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Val(B.AccNum)"
Set rstPayble = Nothing
Call gDbTrans.Fetch(rstPayble, adOpenStatic)

Load frmIntPayble
With frmIntPayble
    Call .LoadContorls(count + 1, 20)
    .lblTitle.Caption = GetResourceString(424, 375, 47)
    .PutTotal = True
    .Title(0) = GetResourceString(36)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(250, 450)
    .Title(3) = GetResourceString(450)
    .Title(4) = GetResourceString(52, 450)
End With

Dim tmpTransID As Long
Dim AccTransID As Long

count = 1
prg.Min = 0: prg.Max = rstMain.recordCount + 2
While Not rstMain.EOF
    
    AccId = Val(FormatField(rstMain("AccId")))
    AccTransID = rstMain("TransID")
    
    CreateDate = rstMain("CreateDate")
    LastIntDate = CreateDate
    MatDate = rstMain("MaturityDate")
    InterestRate = Val(FormatField(rstMain("RateofInterest")))
    
    If FormatField(rstMain("Balance")) = 0 Then AccTransID = 0 'GoTo NextDeposit
    'Int the balance we have to store
    'the Payable Balance so Make It O
    PayableBalance = 0
    If Not rstInt Is Nothing Then
'        rstInt.MoveFirst
        rstInt.Find " AccID = " & rstMain("accId"), , adSearchForward
        If Not rstInt.EOF Then
            LastIntDate = FormatField(rstMain("TransDate"))
            tmpTransID = rstInt("TransID")
            If AccTransID Then _
                AccTransID = IIf(tmpTransID > AccTransID, tmpTransID, AccTransID)
        End If
    End If
    
    'Now Get The Date Difference
    'On that Difference Get the InterestRate
    Duration = DateDiff("D", LastIntDate, AsOnDate)
    If DateDiff("d", MatDate, AsOnDate) <= 1 Then _
                Duration = DateDiff("D", LastIntDate, MatDate)
    
    'Check The Last Trasnction Date
    If DateDiff("D", rstMain("TransDate"), AsOnDate) < 0 Then Duration = 0
    If Duration <= 0 Then
        AccTransID = 0 'GoTo NextDeposit
        frmIntPayble.KeyData(count) = 0
    End If
    
    'If InterestRate = 0 Then _
        InterestRate = CCur(GetRDDepositInterest(OnIndianDate))
    IntAmount = ComputeRDDepositInterestAmount(rstMain("AccId"), GetSysFormatDate(OnIndianDate))
    IntAmount = IntAmount \ 1
    
    If Not rstPayble Is Nothing Then
'        rstPayble.MoveFirst
        rstPayble.Find "AccId = " & rstMain("AccID"), , adSearchForward
        If Not rstPayble.EOF Then
            PayableBalance = FormatField(rstPayble("Payble"))
            If DateDiff("D", rstMain("TransDate"), AsOnDate) < 0 Then AccTransID = 0
            tmpTransID = rstPayble("TransID")
            If AccTransID Then _
                AccTransID = IIf(tmpTransID > AccTransID, tmpTransID, AccTransID)
        End If
    End If
    
    
    If AccTransID Then TransID = AccTransID + 1
    With frmIntPayble
        .grd.RowData(count) = TransID
        .AccNum(count) = rstMain("AccNum")
        .CustName(count) = rstMain("CustName")
        .Balance(count) = PayableBalance
        .Amount(count) = IntAmount
        .Total(count) = IntAmount + PayableBalance
        .KeyData(count) = TransID
        TotalIntAmount = TotalIntAmount + IntAmount
        TotalIntPayble = TotalIntPayble + PayableBalance
    End With
    
NextDeposit:
    rstMain.MoveNext: count = count + 1
    prg.Value = count
Wend

With frmIntPayble
    .CustName(count) = GetResourceString(47, 450, 52, 42) 'Total Interest Payble
    .Balance(count) = TotalIntPayble
    .Amount(count) = FormatCurrency(TotalIntAmount)
End With

Me.Refresh
frmIntPayble.ShowForm

Me.Refresh

prg.Value = 0
If frmIntPayble.grd.Rows = 1 Then Exit Function

'Now Update to the RDTrans
'lblStatus.Caption = "Computing Interests for
 lblStatus.Caption = GetResourceString(907)
    
gDbTrans.BeginTrans


TotalIntAmount = 0
Me.Refresh
count = 1
rstMain.MoveFirst
'For Count = 0 To UBound(TransID)
While Not rstMain.EOF
    AccId = rstMain("AccID")
    With frmIntPayble
        TransID = .KeyData(count)
        IntAmount = .Amount(count)
        PayableBalance = .Total(count)
    End With
    If TransID > 0 And IntAmount > 0 Then
        'Now withdraw the interest amount fromInterest Account
        transType = wContraWithdraw
        gDbTrans.SqlStmt = "INSERT INTO RDIntTrans " & _
            " (AccID,TransDate, TransID, Amount, " & _
            " TransType,Balance,Particulars,UserID) VALUES " & _
            " (" & AccId & "," & _
            "#" & AsOnDate & "#," & _
            TransID & "," & _
            IntAmount & "," & _
            transType & "," & _
            PayableBalance & "," & _
            "'Interest Payable'," & _
            UserID & " )"
            
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            GoTo ErrLine
        End If
        'Now Deposit the interest amount to Interest Payable Account
        transType = wContraDeposit
        gDbTrans.SqlStmt = "INSERT INTO RDIntPayable" & _
            " (AccID,TransDate, TransID, Amount, " & _
            " TransType,Balance,Particulars,UserID) VALUES " & _
            " (" & AccId & "," & _
            "#" & AsOnDate & "#," & _
            TransID & "," & _
            IntAmount & "," & _
            transType & "," & _
            PayableBalance & "," & _
            "'Interest Payable'," & _
            UserID & " )"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            GoTo ErrLine
        End If
        TotalIntAmount = TotalIntAmount + IntAmount
    ElseIf TransID = 0 Then
        txtFailAccIDs = txtFailAccIDs & rstMain("AccNum") & ", "
    End If
    prg.Value = count
    count = count + 1
    
    rstMain.MoveNext
    
Wend

'Next Count

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

'Now Get the Payble And INtere HeadID
Dim PayableHeadID As Long
Dim IntHeadID As Long
Dim headName As String
Dim headNameEnglish As String

headName = m_DepositTypeName & " " & GetResourceString(450)  'GetResourceString(424, 450)         'RD Interest Provision
headNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(450) 'LoadResourceStringS(424, 450)        'RD Interest Provision
IntHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemDepIntPaid, 0, wis_RDAcc + m_DepositType)

headName = m_DepositTypeName & " " & GetResourceString(375, 47) 'GetResourceString(424, 375, 47)             'Rd Payable Interest
headNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(375, 47)  'LoadResourceStringS(424, 375, 47)  'Rd Payable Interest
PayableHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parDepositIntProv, 0, wis_RDAcc + m_DepositType)

'Make the Payble TransCtion to the ledger heads
Call bankClass.UpdateContraTrans(IntHeadID, PayableHeadID, TotalIntAmount, AsOnDate)


gDbTrans.CommitTrans
    

lblStatus = ""
If Len(txtFailAccIDs) > 0 Then
    txtFailAccIDs.Visible = True
    lblStatus = GetResourceString(36) & GetResourceString(92, 544)
End If
'MsgBox " Interest payble  added success fully", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(47, 450) & " " & _
    GetResourceString(637), vbInformation, wis_MESSAGE_TITLE

GoTo ExitLine

ErrLine:

MsgBox "Error In RDAccount --Interest payble", vbCritical, wis_MESSAGE_TITLE
Resume

ExitLine:
'Set BankClass = Nothing
Unload frmIntPayble
Set frmIntPayble = Nothing
Set rst = Nothing
Set rstMain = Nothing
Set rstInt = Nothing
Set rstPayble = Nothing

AddInterestPayableOfRD = True

End Function

Private Sub cmdLoad_Click()
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim ret As Long

gDbTrans.SqlStmt = "SELECT AccID,LoanId From RDMaster " & _
    " WHERE AccNum = " & AddQuotes(txtAccNo, True)
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & m_DepositType

ret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If ret = 1 Then
    AccId = FormatField(rst("AccID"))
ElseIf ret < 1 Then
    'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

Set rst = Nothing

If Not AccountLoad(AccId) Then
    ActivateTextBox txtAccNo
    Exit Sub
End If

End Sub
Private Sub cmdNextTrans_Click()
    
    If m_rstPassBook Is Nothing Then Exit Sub
    
    Dim CurPos As Integer
    'Dim CurPosInst As Integer
'Position cursor to start of next page
        
        If m_rstPassBook.EOF Then m_rstPassBook.MoveLast
        
         Call PassBookPageShow
         
         Exit Sub
        CurPos = m_rstPassBook.AbsolutePosition
        CurPos = 10 - (CurPos Mod 10)
        If m_rstPassBook.AbsolutePosition + CurPos >= m_rstPassBook.recordCount Then
            Beep
            Exit Sub
        Else
            m_rstPassBook.Move CurPos
        End If
    Call PassBookPageShow
    
    #If junk Then
    
    If m_rstPassBook.AbsolutePosition < m_rstPassBook.recordCount - 10 Then
        
        If m_rstPassBook.AbsolutePosition Mod 10 <> 0 Then
            m_rstPassBook.Move 10 - m_rstPassBook.AbsolutePosition Mod 10
            
            If m_rstPassBook.AbsolutePosition >= m_rstPassBook.recordCount - 10 Then
                cmdNextTrans.Enabled = False
            End If
        
        End If
    
    Else
        cmdNextTrans.Enabled = False
    End If
    
    Call ShowPassBookPage
    
    If m_rstPassBook.AbsolutePosition >= m_rstPassBook.recordCount Then
        cmdPrevTrans.Enabled = False
    Else
        cmdPrevTrans.Enabled = True
    End If
    
    #End If

End Sub
Private Sub cmdOk_Click()
    
    Dim cancel As Boolean
        'Ask the user before Closing
'    If MsgBox(GetResourceString(750), vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
'        Exit Sub
'    End If
    
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
    If CurPos = 10 Then cmdPrevTrans.Enabled = False: Exit Sub
   ' CurPosInst = m_InstallmentRst.AbsolutePosition
    CurPos = CurPos - CurPos Mod 10 - 10
    'CurPosInst = CurPosInst - CurPosInst Mod 10 - 10
    
    If CurPos < 0 Then
        Beep
        Exit Sub
    Else
        m_rstPassBook.MoveFirst
        m_rstPassBook.Move (CurPos)
    End If
    
    Call PassBookPageShow
    
End Sub

Private Sub cmdPrint_Click()
     If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    
    m_frmPrintTrans.Show vbModal
      
Exit Sub
      

End Sub

Private Sub cmdReset_Click()

    Call ResetUserInterface

End Sub
Private Sub cmdSave_Click()

'SaveAccount
    If Not AccountSave Then
        Exit Sub
    End If

'Reload the account details once saved
    Dim AccNo As String
    Dim rst As ADODB.Recordset
    AccNo = Val(GetVal("AccNum"))
    Dim AccId As Long
    gDbTrans.SqlStmt = "SELECT AccID From RDMaster " & _
        " WHERE AccNum = " & AddQuotes(AccNo, True)
    If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & m_DepositType
    If gDbTrans.Fetch(rst, adOpenForwardOnly) = 1 Then
        AccId = FormatField(rst("AccID"))
    End If
    Set rst = Nothing

    If Not AccountLoad(AccId) Then
        'MsgBox  "Error loading account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(526), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    txtAccNo.Text = m_AccID

End Sub

Private Sub cmdTerminate_Click()
Dim I As Integer
Dim strField As String
Dim ret As Integer
Dim rst As ADODB.Recordset

'Prelim check
    If m_AccID = 0 Then
        'MsgBox "No account loaded !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
        Exit Sub
    End If

'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & m_AccID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
'Check if have to reopen the account
    If m_AccClosed Then
        'If MsgBox("Are you sure you want to reopen this account ?", vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
        If MsgBox(GetResourceString(538), vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
            Exit Sub
        End If
        If Not AccountReopen(m_AccID) Then
            Exit Sub
        End If
        'MsgBox "Account reopened successfully !", vbInformation, gAppName & " - Message"
        MsgBox GetResourceString(522), vbInformation, gAppName & " - Message"
        If Not AccountLoad(m_AccID) Then
            'MsgBox "Unable to reload the account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
            Exit Sub
        End If

        Exit Sub
    Else
        'Check if there are any transactions
        gDbTrans.SqlStmt = "Select TOP 1 * from RDTrans where AccID = " & m_AccID & " order by TransID desc"
        ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
        If ret <= 0 Then
            'Ret = MsgBox("You do not have any transactions on this account !" & _
                vbCrLf & "It is recommended that you delete this account permanently." & _
                vbCrLf & vbCrLf & _
                "Click Yes to delete this account permanently. (Recommended)" & _
                vbCrLf & "Click No to only close this account." & _
                vbCrLf & "Click Cancel to cancel the operation", _
                vbYesNoCancel + vbQuestion, gAppName & " - Confirmation")
            ret = MsgBox(GetResourceString(551) & _
                vbCrLf & GetResourceString(552) & _
                vbCrLf & vbCrLf & _
                GetResourceString(652) & _
                vbCrLf & GetResourceString(653) & _
                vbCrLf & GetResourceString(654), _
                vbYesNoCancel + vbQuestion, gAppName & " - Confirmation")
            If ret = vbCancel Then
                Exit Sub
            ElseIf ret = vbYes Then  'Proceed with delete
                If Not AccountDelete(m_AccID) Then
                    'MsgBox "Unable to delete account !", vbCritical, gAppName & " - Error"
                    MsgBox GetResourceString(532), vbCritical, gAppName & " - Error"
                    Exit Sub
                Else
                    Call ResetUserInterface
                End If
            End If
        Else
            'Check if balance is 0
            If FormatField(rst("Balance")) > 0 Then
                'MsgBox "This account has a balance of Rs. " & FormatField(gDBTrans.Rst("Balance")) & " and thus cannot be closed !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(549) & FormatField(rst("Balance")) & GetResourceString(655), vbExclamation, gAppName & " - Error"
                Exit Sub
            End If
            'If MsgBox("Are you sure you want to close this account ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
            If MsgBox(GetResourceString(656), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                Exit Sub
            End If
        End If
        
        'Close this account now
        If Not AccountClose() Then Exit Sub
        'MsgBox "Account closed successfully !", vbInformation, gAppName & " - Message"
        MsgBox GetResourceString(657), vbInformation, gAppName & " - Message"
        'Reload the account
        If Not AccountLoad(m_AccID) Then
            'MsgBox "Unable to reload the account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
            Exit Sub
        End If
        Exit Sub
    End If
    
     
End Sub

Private Function AccountReopen(AccId As Long) As Boolean

Dim ClosedDate As String
Dim rst As ADODB.Recordset
'Check if account number exists in data base
gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
    'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If
ClosedDate = FormatField(rst("ClosedDate"))
Set rst = Nothing

gDbTrans.BeginTrans
gDbTrans.SqlStmt = "Update RDMaster set ClosedDate = NULL where AccID = " & AccId
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If
gDbTrans.CommitTrans
    
' If While Closing this A/c if any Misc Amount collected that has to return(Undo)
'Dim BankClass As New clsBankAcc
'Call BankClass.UndoUPDatedMiscProfit(ClosedDate, "Closing RD " & "-" & m_AccID)
'Call BankClass.UndoRemoveInterestPayble(ClosedDate, "Closing RD " & "-" & m_AccID, 3)
'Set BankClass = Nothing

AccountReopen = True
End Function

Private Sub cmdUndo_Click()

If Not AccountUndoLastTransaction() Then Exit Sub
    
If Not AccountLoad(m_AccID) Then
    'MsgBox "Unable to undo transaction !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(609), vbCritical, gAppName & " - Error"
    Exit Sub
End If

End Sub


'Private Sub cmdUndoInterests_Click()
''Prelim check
'    If Not DateValidate(txtUndoInterest.Text, "/", True) Then
'        Exit Sub
'    End If
'
'    If MsgBox("WARNING !" & vbCrLf & vbCrLf & _
'            "This will withdraw interests to all the deposits of the specified date !" & vbCrLf & _
'            "Click OK only if you are sure about this operation !" & vbCrLf & vbCrLf & _
'            "Are you sure you want to continue ?", vbQuestion + vbYesNo, gAppName & _
'            " - Confirmation") = vbNo Then

Private Sub cmdUndoPayble_Click()
If Not DateValidate(txtIntPayable.Text, "/", True) Then
'''    MsgBox "Invalid Date Format Specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntPayable
    Exit Sub
End If
If UndoInterestPayableOfRD(txtIntPayable.Text) Then
    MsgBox "Added Interest payable transction was removed", vbInformation, wis_MESSAGE_TITLE
Else
    MsgBox "Unable to remove payable transction ", vbInformation, wis_MESSAGE_TITLE
End If
End Sub

Private Sub cmdView_Click()

'First check the dates specified
If txtFromDate.Enabled And Not DateValidate(txtFromDate.Text, "/", True) Then
    'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtFromDate
    Exit Sub
End If

If txtToDate.Enabled Then
    If Not DateValidate(txtToDate.Text, "/", True) Then
        'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtToDate
        Exit Sub
    End If
    If txtFromDate.Enabled Then
    If WisDateDiff(txtFromDate.Text, txtToDate.Text) < 0 Then
        'MsgBox "TO date is earlier that the specified FROM date!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtToDate
        Exit Sub
    End If
    End If
End If

Dim ReportType As wis_RDReports

If optOpened Then ReportType = repRDAccOpen
If optClosed Then ReportType = repRDAccClose
If optDepGLedger Then ReportType = repRDLedger
If optDepositBalance Then ReportType = repRDBalance
If optSubDayBook Then ReportType = repRDDayBook
If optLiabilities Then ReportType = repRDLaib
If optMature Then ReportType = repRDMat
If optMonthlyBalance Then ReportType = repRDMonbal
If optCashBook Then ReportType = repRDCashBook

If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    

RaiseEvent ShowReport(ReportType, IIf(optAccId, wisByAccountNo, wisByName), _
            IIf(txtFromDate.Enabled, txtFromDate, ""), IIf(txtToDate.Enabled, txtToDate, ""), _
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

Private Function AccountExists(AccId As Long, Optional ClosedON As String) As Boolean
Dim ret As Long
Dim rst As ADODB.Recordset
'Query Database
    gDbTrans.SqlStmt = "Select * from RDMaster where " & _
                        " AccID = " & AccId
   ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
       MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Check the closed status
    If Not IsMissing(ClosedON) Then _
       ClosedON = FormatField(rst("ClosedDate"))

    Set rst = Nothing

    AccountExists = True
End Function

'
Public Function AccountLoad(AccId As Long) As Boolean

Dim rstMaster As Recordset
Dim rstJoint As Recordset
Dim rstTemp As Recordset

Dim ClosedDate As String
Dim ret As Integer
Dim JointHolders() As String
Dim I As Integer

'Check if account number is valid
    If AccId <= 0 Then GoTo DisableUserInterface
    
'Query data base
    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & AccId
    Call gDbTrans.Fetch(rstMaster, adOpenForwardOnly)
    'Set record set to local rec set
    
    m_AccID = FormatField(rstMaster("AccID"))
    m_AccNum = FormatField(rstMaster("AccNum"))
    
    On Error Resume Next
    Unload m_frmRDJoint
    Set m_frmRDJoint = Nothing
    On Error GoTo DisableUserInterface
    
    gDbTrans.SqlStmt = "Select * from RDJoint where AccNUm = " & AddQuotes(txtAccNo, True)
    If gDbTrans.Fetch(rstJoint, adOpenStatic) > 0 Then _
        Set m_frmRDJoint = New frmJoint

'Load the Name details
    If Not m_CustReg.LoadCustomerInfo(FormatField(rstMaster("CustomerID"))) Then
        'MsgBox "Unable to load customer information !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        GoTo DisableUserInterface
    End If
    
    gDbTrans.SqlStmt = "Select * from RDTrans where " & _
                " Accid = " & m_AccID & " ORDER BY TransID"
    ret = gDbTrans.Fetch(m_rstPassBook, adOpenStatic)
    If ret > 0 Then
        Dim BalanceAmount As Currency
        m_rstPassBook.MoveLast
        BalanceAmount = m_rstPassBook("Balance")
    
        'Position to first record of last page
        m_rstPassBook.MoveFirst
        If m_rstPassBook.recordCount > 10 Then m_rstPassBook.Move ret - ret Mod 10
        cmdUndo.Enabled = gCurrUser.IsAdmin
    Else
        Set m_rstPassBook = Nothing
        PassBookPageInitialize
        cmdUndo.Enabled = False
    End If
    
'Assign to some module level variables
    m_accUpdatemode = wis_UPDATE
    m_AccClosed = IIf(FormatField(rstMaster("ClosedDate")) <> "", True, False)
    
'Load account to the User Interface
    'TAB 1

ClosedDate = FormatField(rstMaster("ClosedDate"))
With lblBalance
    .Caption = GetResourceString(42) & "  " & FormatCurrency(BalanceAmount)
    If ClosedDate <> "" Then
        .Caption = "Account Closed"
        cmdUndo.Caption = GetResourceString(313)
    Else
        cmdUndo.Caption = GetResourceString(19)
    End If
End With

With cmbAccNames
    .Enabled = True: .BackColor = vbWhite: .Clear
    .AddItem m_CustReg.FullName
    
    If Not rstJoint Is Nothing Then
              
        While Not rstJoint.EOF
            .AddItem m_CustReg.CustomerName(rstJoint("CustomerID"))
            m_JointCust(I) = rstJoint("CustomerID")
            rstJoint.MoveNext
        Wend
    End If
    .ListIndex = 0
End With
With txtDate
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    .Text = gStrDate
End With
With cmdDate
    .Enabled = IIf(ClosedDate = "", True, False)
    If gOnLine Then .Enabled = False

End With

With cmbTrans
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    .ListIndex = -1
End With

With cmbParticulars
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    .ListIndex = -1
End With

With txtAmount
    '.Enabled = True
    '.BackColor = vbWhite
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    .Text = FormatField(rstMaster("InstallmentAmount"))
    m_Amount = CCur(txtAmount)
End With

    cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
    cmdPrevTrans.Enabled = IIf(ClosedDate = "", True, True)
    cmdNextTrans.Enabled = IIf(ClosedDate = "", True, True)

With rtfNote
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    Call m_Notes.LoadNotes(wis_RDAcc, AccId)
End With

Call m_Notes.DisplayNote(rtfNote)
TabStrip2.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True
With txtCheque
    .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    .Enabled = IIf(ClosedDate = "", True, False)
    .Text = ""
End With

With TxtInstallmentNo
       .Text = Installment(AccId): .Locked = True
     '.BackColor = vbWhite
       .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
       .Enabled = IIf(ClosedDate = "", True, False)
End With

cmdAccept.Enabled = IIf(ClosedDate = "", True, False)
cmdClose.Enabled = IIf(ClosedDate = "", True, False)

Call PassBookPageShow
    
    'TAB 2
    'Update labels and other buttons
    'get teh rate of rd deposit  interest from SETup TAble
    
    Dim SetupClass  As New clsSetup
    Dim RateofDepInterest As Single
    Dim IntroId As Long
    Dim NomineeId As Long
    Dim cumulativeID As Byte
     
    NomineeId = FormatField(rstMaster("NomineeID"))
    IntroId = FormatField(rstMaster("IntroducerID"))
    cumulativeID = FormatField(rstMaster("Cumulative"))
    
    RateofDepInterest = CSng(SetupClass.ReadSetupValue("RDAcc", "Interest On RDDeposit", 12))
    Set SetupClass = Nothing
    
lblOperation.Caption = GetResourceString(56)  '"Operation Mode : <UPDATE>"
cmdTerminate.Caption = IIf(ClosedDate = "", "&Terminate", "&Reopen")
cmdTerminate.Enabled = True
'mallikpatil@usa.net

Dim strField As String
Dim txtIndex As String
Dim count As Integer
For I = 0 To txtPrompt.count - 1
    ' Read the bound field of this control.
    On Error Resume Next
    strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If strField <> "" Then
        With txtData(I)
            Select Case UCase$(strField)
                Case "ACCNUM"
                    .Text = rstMaster("AccNum")
                    .Locked = True
                Case "ACCNAME"
                    .Text = m_CustReg.FullName
                Case "JOINTHOLDER"
                    .Text = ""
                    If Not m_frmRDJoint Is Nothing Then .Text = m_frmRDJoint.JointHolders
                Case "NOMINEEID"
                    .Text = IIf(NomineeId, NomineeId, "")
                Case "NOMINEENAME"
                    .Text = IIf(NomineeId, m_CustReg.CustomerName(NomineeId), FormatField(rstMaster("NomineeName")))
                    txtIndex = ExtractToken(txtPrompt(I).Tag, "TextIndex")
                    For count = 0 To cmb(Val(txtIndex)).ListCount - 1
                        If cmb(Val(txtIndex)).List(count) = .Text Then
                            cmb(Val(txtIndex)).ListIndex = count
                            Exit For
                        End If
                    Next
                    cmb(Val(txtIndex)).Text = .Text
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
                Case "CREATEDATE"
                    .Text = FormatField(rstMaster("CreateDate"))
                Case "RDAMOUNT"
                    .Text = FormatField(rstMaster("InstallmentAmount"))
                Case "DATEOFPAYMENT"
                    .Text = FormatField(rstMaster("DateOfPayment"))
                Case "NOOFINST"
                    .Text = FormatField(rstMaster("NoOfInstallments"))
                Case "RATEOFINTEREST"
                    .Text = RateofDepInterest
                Case "MATURITYDATE"
                    .Text = FormatField(rstMaster("MaturityDate"))
                Case "CUMULATIVE"
                    txtIndex = ExtractToken(txtPrompt(I).Tag, "TextIndex")
                    For count = 0 To cmb(Val(txtIndex)).ListCount - 1
                        If cmb(Val(txtIndex)).ItemData(count) = cumulativeID Then
                            cmb(Val(txtIndex)).ListIndex = count
                            cmb(Val(txtIndex)).Text = cmb(Val(txtIndex)).List(count)
                            .Text = cmb(Val(txtIndex)).Text
                            Exit For
                        End If
                    Next
                    
                Case "ACCGROUP"
                    gDbTrans.SqlStmt = "SELECT GroupName FROM AccountGroup WHERE " & _
                            "AccGroupID = " & FormatField(rstMaster("AccGroupId"))
                    If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
                    .Text = FormatField(rstTemp("GroupName"))
                Case "NOTIFY"
                    .Text = FormatField(rstMaster("NotifyOnMaturity"))
                    chk(Val(ExtractToken(txtPrompt(I).Tag, "Notify"))).Value = Val(.Text)
                Case "OPERATIVE"
                    .Text = FormatField(rstMaster("InOperative"))
                    chk(Val(ExtractToken(txtPrompt(I).Tag, "OPERATIVE"))).Value = Val(.Text)
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
                chk(CtlIndex).Value = IIf(txtData(I).Text = True, vbChecked, vbUnchecked)
                
          End Select
        End With
    End If

    
Next
    
'Check is the DepositTYpe loaded or Not
If Len(lblDepositTypeName.Caption) = 0 And FormatField(rstMaster("DepositTYpe")) > 0 Then
    LoadDepositType (FormatField(rstMaster("DepositTYpe")))
End If
    
AccountLoad = True
cmdPhoto.Enabled = Len(gImagePath)
Exit Function

DisableUserInterface:
    Call ResetUserInterface

Exit Function
    
ErrLine:
'MsgBox "Account Load:" & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"
MsgBox GetResourceString(521) & vbCrLf & GetResourceString(526), vbCritical, gAppName & " - Error"
End Function

Private Sub Form_Load()

Screen.MousePointer = vbHourglass
'Centre the form
Call CenterMe(Me)

'Load the Default Head name a
If Len(m_DepositTypeName) = 0 Then
    m_DepositTypeName = GetResourceString(424)
    m_DepositTypeNameEnglish = LoadResourceStringS(424)
End If

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)

' Set kannada fonts
Call SetKannadaCaption
'Fill up transaction Types in RD Only Deposit Type Of transaction Takes Place
    With cmbTrans
        .AddItem GetResourceString(271)    '"Deposit"
        .AddItem GetResourceString(273)    '"Charges"
    End With
    
' fill the module type
    m_CustReg.ModuleID = wis_RDAcc

'Fill up particulars with default values from RDacc.INI
    Dim Particulars As String
    Dim I As Integer
    Do
        Particulars = ReadFromIniFile("Particulars", "Key" & I, gAppPath & "\RDAcc.INI")
        If Trim$(Particulars) <> "" Then
            cmbParticulars.AddItem Particulars
        End If
        I = I + 1
    Loop Until Trim$(Particulars) = ""

'Load ICONS
    cmdAddNote.Picture = LoadResPicture(103, vbResBitmap)

'Adjust the Grid for Pass book No Cheque For RDacc
With grd
    .Rows = 11
    .Cols = 6
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(37): .CellFontBold = True   ' "Date"
    .Col = 1: .Text = GetResourceString(39): .CellFontBold = True   '"Particulars"
    .Col = 2: .Text = GetResourceString(55): .CellFontBold = True   '"Installment"
    .Col = 2: .Text = GetResourceString(276): .CellFontBold = True   '"Debit"
    .Col = 3: .Text = GetResourceString(277): .CellFontBold = True   '"Credit"
    .Col = 4: .Text = GetResourceString(42): .CellFontBold = True   '"Balance"
End With

Call LoadPropSheet

Dim cmbIndex As Byte
cmbIndex = GetIndex("AccGroup")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
Call LoadAccountGroups(cmb(cmbIndex))

Call LoadInterestRates
'Dim SetUp As New clsSetup
'Txt_int_on_rd_dep_tab4 = SetUp.ReadSetupValue("RDacc", "Interest On RDDeposit", CStr(12#))
'Txt_int_on_rd_loan_tab4 = SetUp.ReadSetupValue("RDacc", "Interest On RDLoan", CStr(14#))

txtToDate = gStrDate
m_FormLoaded = True
lblDepositTypeName.Caption = ""
'Reset the User Interface
Call ResetUserInterface
Screen.MousePointer = vbDefault

If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
gWindowHandle = 0

End Sub

Private Sub Form_Resize()
lblDepositTypeName.Left = (Me.Width - lblDepositTypeName.Width) / 2 - 100
    cmdDepositType.Left = lblDepositTypeName.Left + lblDepositTypeName.Width = 50
End Sub

'
Private Sub Form_Unload(cancel As Integer)

' Report form.
If Not m_frmLookUp Is Nothing Then Unload m_frmLookUp
Set m_frmLookUp = Nothing


' Notes object.
Set m_Notes = Nothing

' Customer Registration object.
Set m_CustReg = Nothing
gWindowHandle = 0

RaiseEvent WindowClosed

End Sub

Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)
Call PrintBetweenDates(wis_RDAcc, m_AccID, StartIndiandate, EndIndianDate)
Exit Sub
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer
'1. Fetch last print row from RDmaster table.
'First get the last printed txnID From the RDMaster
SqlStr = "SELECT  LastPrintRow From RDMaster WHERE AccId = " & m_AccID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintRow = IIf(IsNull(metaRst("LastPrintrow")), 0, metaRst("LastPrintrow"))


'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From RDTrans WHERE AccId = " & m_AccID
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
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
clsPrint.isNewPage = True
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE_A4 - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If


'3. Getting matching records for passbook printing
    SqlStr = "SELECT * From RDTrans WHERE AccId = " & m_AccID & _
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

    While Not rst.EOF
        If .isNewPage Then
            .printHeader2
            .isNewPage = False
        End If
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        .ColText(2) = FormatField(rst("ChequeNo"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(3) = FormatField(rst("Amount"))
        Else
            .ColText(4) = FormatField(rst("Amount"))
        End If
        .ColText(5) = FormatField(rst("Balance"))
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

'Now Update the Last Print Id to the RDmaster
SqlStr = "UPDATE RDMaster set LastPrintRow = " & curPrintRow - 1 & _
        " Where Accid = " & m_AccID
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If


End Sub

'
Private Sub m_frmPrintTrans_TransClick(bNewPassbook As Boolean)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintId, lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer

'First get the last printed transaID From the RDMaster

SqlStr = "SELECT  LastPrintID, LastPrintRow From RDMaster WHERE AccId = " & m_AccID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenForwardOnly) < 1 Then
MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintId"))
'get position of the last print point
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If IsNull(metaRst("LastPrintRow")) And lastPrintId = 1 Then lastPrintId = 0

' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From RDTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

' Print the first page of passbook, if newPassbook option is chosen.
If bNewPassbook Then
    clsPrint.printPassbookPage wis_RDAcc, m_AccID
    'Update the print rows as 0 and lst transid as -5
    'Now Update the Last Print Id to the master
    SqlStr = "UPDATE RDMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
            ", LastPrintRow = 0 Where Accid = " & m_AccID
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
    Else
        gDbTrans.CommitTrans
    End If
    

    MsgBox "New Passbook page printed", vbInformation, "Index 2000 Print"
    Exit Sub
End If
' If there are no records to print, since the last printed txn,
' display a message and exit.
If (rst(0) = 0) Then
    
    Dim bRet As Integer
    bRet = MsgBox("There are no transactions available for printing." & vbCrLf & _
        "Do you want to reset the printing from beginning?", vbYesNo, "Debug message")
    If (bRet = vbYes) Then
        SqlStr = "UPDATE RDMaster set LastPrintId = " & 0 & _
                ", LastPrintRow = " & 0 & _
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
SqlStr = "SELECT * From RDTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Print [or don't print] header part
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
End If


'Printer.PaperSize = 9
Printer.Font = gFontName
'Printer.Font.Size = 12 'gFontSize
'Printer.Font = "Courier New"
Printer.Font.Size = 9
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 5
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .COlHeader(1) = GetResourceString(275) 'Cheque
    '.ColWidth(2) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(3) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(4) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(5) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    
    If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 1 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
    End If
    
    ' column widths for printing txn rows.
     .ColWidth(0) = 15
        .ColWidth(1) = 13
        .ColWidth(2) = 17
        .ColWidth(3) = 13
        .ColWidth(4) = 14
        .ColWidth(5) = 15

    While Not rst.EOF
        If .isNewPage Then
            .printHeader2
            .isNewPage = False
        End If

        TransID = FormatField(rst("TransID"))
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        .ColText(2) = FormatField(rst("ChequeNo"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(3) = FormatField(rst("Amount"))
            .ColText(4) = ""
        Else
            .ColText(3) = ""
            .ColText(4) = FormatField(rst("Amount"))
        End If
        .ColText(5) = FormatField(rst("Balance"))
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

'Now Update the Last Print Id to the RDmaster
SqlStr = "UPDATE RDMaster set LastPrintId = " & TransID & _
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

Private Sub optCashBook_Click()
Call optSUbDayBook_Click
End Sub

'
Private Sub optClosed_Click()
    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(False, False)
    
If optClosed.Value = True Then _
    txtFromDate.Enabled = True: txtFromDate.BackColor = vbWhite: cmdDate1.Enabled = True
End Sub

'
Private Sub optDepGLedger_Click()

'Enable/Disabale Place,Caste, Group and AMount Range Controls
Call EnableOptionDialogControls(False, False)
    
With txtFromDate
    .Enabled = True
    .BackColor = wisWhite
End With
cmdDate1.Enabled = True

End Sub

Private Sub optDepositBalance_Click()

'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, False)
    
txtFromDate.Enabled = False
cmdDate1.Enabled = False
txtFromDate.BackColor = wisGray
End Sub

Private Sub optLiabilities_Click()
'Enable/Disabale Place,Caste, Group and AMount Range Controls
Call EnableOptionDialogControls(True, False)

If optLiabilities.Value = True Then _
    txtFromDate.Enabled = False: cmdDate1.Enabled = False: txtFromDate.BackColor = wisGray
End Sub


'
Private Sub optLoanBalance_Click()

    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, True)
    
    txtToDate.Enabled = False
    txtToDate.BackColor = wisGray
End Sub

'
Private Sub optMature_Click()

    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, False)
    
    If optMature.Value = True Then _
        txtFromDate.Enabled = True: txtFromDate.BackColor = vbWhite: cmdDate1.Enabled = True
End Sub

Private Sub optMonthlyBalance_Click()

    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, True)
    
    txtFromDate.Enabled = False
    cmdDate1.Enabled = False
    txtFromDate.BackColor = wisGray
    txtFromDate.Enabled = True: cmdDate1.Enabled = True: txtFromDate.BackColor = wisWhite
End Sub

'
Private Sub optOpened_Click()

    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, False)
    
    If optOpened.Value = True Then _
        txtFromDate.Enabled = True: txtFromDate.BackColor = vbWhite: cmdDate1.Enabled = True
End Sub

Private Sub optSUbDayBook_Click()
    'Enable/Disabale Place,Caste, Group and AMount Range Controls
    Call EnableOptionDialogControls(True, True)

    With txtFromDate
        .Enabled = True
        .BackColor = wisWhite
    End With
    cmdDate1.Enabled = True

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
        cmdIntApply.Default = True
        
End Select


End Sub

Private Sub TabStrip2_Click()
    If TabStrip2.SelectedItem.Index = 1 Then
        fraInstructions.ZOrder 0
        fraInstructions.Visible = True
        fraPassBook.Visible = False
    Else
        fraInstructions.Visible = False
        fraPassBook.ZOrder 0
        fraPassBook.Visible = True
    End If
End Sub



Private Sub txtAccNo_Change()
cmdLoad.Enabled = IIf(Trim$(txtAccNo.Text) <> "", True, False)

If m_AccID Then Call ResetUserInterface

End Sub


Private Sub txtData_DblClick(Index As Integer)
txtData_KeyPress Index, vbKeyReturn
End Sub
'
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
    'Get the cmdbutton index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    If TextIndex <> "" Then cmd(Val(TextIndex)).Visible = True
ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
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
'
Private Sub txtData_LostFocus(Index As Integer)

txtPrompt(Index).ForeColor = vbBlack
Dim strDatSrc As String
Dim strDatVal As String
Dim Lret As Long
Dim txtIndex As Integer
Dim rst As ADODB.Recordset

' If the item is IntroducerID, validate the
' ID and name.
strDatSrc = ExtractToken(txtPrompt(Index).Tag, "DataSource")
Select Case UCase(strDatSrc)
    Case "INTRODUCERID", "NOMINEEID"
        ' Check if any data is found in this text.
        If Val(txtData(Index).Text) > 0 Then
            gDbTrans.SqlStmt = "SELECT CustomerID, Title + FirstName + space(1) + " _
                    & "MiddleName + space(1) + Lastname AS Name FROM " _
                    & "NameTab WHERE NameTab.CustomerID = " & Val(txtData(Index).Text)
            Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
            If Lret > 0 Then
                txtIndex = GetIndex("IntroducerName")
                txtData(txtIndex).Text = FormatField(rst("Name"))
            End If
        Else
            txtData(Index + 1).Text = ""
        End If

    'Find AutoMatically Maturity date with the help of NoOfInstallments
    Dim NoOFInstallments As Double
    
    Case "NOOfINST"
        strDatVal = txtData(GetIndex("CreateDate"))
        NoOFInstallments = Val(txtData(Index).Text)
        On Error Resume Next
        txtData(Index + 2).Text = Format(DateAdd("m", NoOFInstallments, GetSysFormatDate(strDatVal)), "dd/mm/yyyy")

    Case "CREATEDATE"
        strDatVal = txtData(GetIndex("CreateDate"))
        NoOFInstallments = Val(txtData(GetIndex("NoOfInst")))
        On Error Resume Next
        'txtData(GetIndex("MaturityDate")).Text = Format(DateAdd("m", NoOFInstallments, CDate(FormatDate(txtData(Index)))), "dd/mm/yyyy")
        Err.Clear
End Select
    Set rst = Nothing

End Sub

'
Private Sub txtPrompt_GotFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlue
End Sub


'
Private Sub txtPrompt_LostFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlack
End Sub

'
Private Function UndoInterestPayableOfRD(OnIndianDate As String) As Boolean

lblStatus = ""
txtFailAccIDs = ""
Dim DimPos As Integer
Dim OnDate As Date
Dim rst As ADODB.Recordset

OnDate = GetSysFormatDate(OnIndianDate)

DimPos = InStr(1, OnIndianDate, "31/3/", vbTextCompare)
If DimPos = 0 Then DimPos = InStr(1, OnIndianDate, "31/03/", vbTextCompare)
If DimPos = 0 Then
    'MsgBox "Unable to perform the transactions", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

OnDate = GetSysFormatDate(OnIndianDate)

'Before undoing check whether he has already added the interestpayble amount or not
gDbTrans.SqlStmt = "Select *  from RDIntTrans Where " & _
    " TransDate = #" & OnDate & "# " & _
    " And Particulars ='Interest Payable'"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    UndoInterestPayableOfRD = True
    Exit Function
End If
  
Screen.MousePointer = vbHourglass
  On Error GoTo ErrLine
  'declare the variables necessary


'Get the Payble Amount
gDbTrans.SqlStmt = "SELECT SUM(A.Amount) From RdIntPayable A" & _
    " WHERE A.TransID = " & _
        "(SELECT TransID FROM RDIntTrans C WHERE" & _
        " Particulars = 'Interest Payable' AND TransDate = #" & OnDate & "#" & _
        " AND C.AccID = A.AccID) AND TransDate = #" & OnDate & "#" & _
    " AND A.TransID > (SELECT Max(TransID) FROM RDTrans E WHERE " & _
        " A.AccID = E.AccID)"

'Dim Rst As Recordset
Dim Amount As Currency

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ErrLine
Amount = FormatField(rst(0))

Dim SqlStr As String

'DELETE THE TRANSCTION FROM Interest payable account _
'and respective transaction in RD Interest account
SqlStr = "DELETE A.*, B.* From RDIntPayable A," & _
    " RDIntTrans B WHERE A.AccID = B.AccID " & _
    " AND B.Particulars = 'Interest Payable' "

'Where The Interest payable Transction Should be the last transaction
SqlStr = SqlStr & " AND A.TransID = (SELECT Max(TransID) FROM " & _
    " RDIntTrans C WHERE TransDate = #" & OnDate & "# AND C.AccID = A.AccID)"

'And The Interest paid Transction Should also be the last transaction
SqlStr = SqlStr & " AND B.TransID = (SELECT Max(TransID) FROM " & _
    " RDIntPayable D WHERE TransDate = #" & OnDate & "# AND D.AccID = A.AccID)"

'And Transid's of bOthe Intrest payble interest accounte should be same
'After this Transction No Transacion should have taken place in the PD TRans
SqlStr = SqlStr & " AND B.TransID = A.TransID " & _
 " AND (B.TransID > (Select Max(TransID) From RDTrans E Where E.AccId = A.AccId)) "

gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
gDbTrans.SQLExecute

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

'Now Get the Payble And INtere HeadID
Dim PayableHeadID As Long
Dim IntHeadID As Long
Dim headName As String
Dim headNameEnglish As String
headName = m_DepositTypeName & " " & GetResourceString(450) 'GetResourceString(424, 450)        'RD Interest Provision
headNameEnglish = m_DepositTypeNameEnglish & " " & LoadResourceStringS(450) 'LoadResourceStringS(424, 450)  'RD Interest Provision
IntHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemDepIntPaid, 0, wis_RDAcc + m_DepositType)

headName = m_DepositTypeName & " " & GetResourceString(375, 47)  ' GetResourceString(424, 375, 47) 'Rd Payable Interest
headNameEnglish = m_DepositTypeName & " " & LoadResourceStringS(375, 47) 'LoadResourceStringS(424,375,47)               'Rd Payable Interest
PayableHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parDepositIntProv, 0, wis_RDAcc + m_DepositType)

'Make the Payble TransCtion to the ledger heads
If Not bankClass.UndoContraTrans(IntHeadID, PayableHeadID, Amount, OnDate) Then _
    gDbTrans.RollBack: GoTo ExitLine

gDbTrans.CommitTrans

Set bankClass = Nothing
UndoInterestPayableOfRD = True
'now Check If Any  Account are unable to the undo
gDbTrans.SqlStmt = "Select AccNum from RDMaster A,RDIntTrans B Where " & _
    " A.AccId = B.accID And TransDate = #" & GetSysFormatDate(OnIndianDate) & "# "
If m_DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND A.DepositType = " & m_DepositType
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.Particulars ='Interest Payable' ORDER BY Val(AccNum) "

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ExitLine

MsgBox "For some accounts Interest payable wat not removed", vbInformation, wis_MESSAGE_TITLE

While Not rst.EOF
    txtFailAccIDs = txtFailAccIDs & "," & rst("AccNum")
    rst.MoveNext
Wend

txtFailAccIDs = Mid(txtFailAccIDs, 2)
Set rst = Nothing

GoTo ExitLine

ErrLine:
    MsgBox "Error In RDAccount -- Remove Interest payble", vbCritical, wis_MESSAGE_TITLE
    'Resume

ExitLine:

Set bankClass = Nothing
Screen.MousePointer = vbDefault

End Function



'
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
lblDate.Caption = GetResourceString(37)
lblTrans.Caption = GetResourceString(38)
lblParticular.Caption = GetResourceString(39)
lblAmount.Caption = GetResourceString(40)
lblBalance.Caption = GetResourceString(42)
lblInstrNo.Caption = GetResourceString(41)
lblInstllNo.Caption = GetResourceString(55)
cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)
cmdClose.Caption = GetResourceString(11)
TabStrip2.Tabs(1).Caption = GetResourceString(219)
TabStrip2.Tabs(2).Caption = GetResourceString(218)

'Now Change the Font of New Account Frame
cmdTerminate.Caption = GetResourceString(14)
cmdSave.Caption = GetResourceString(7)
cmdReset.Caption = GetResourceString(8)
lblOperation.Caption = GetResourceString(54)
cmdPhoto.Caption = GetResourceString(415)

'Now Change The Caption of Report Frame
Me.fraChooseReport.Caption = GetResourceString(288)
optDepositBalance.Caption = GetResourceString(70)
optSubDayBook.Caption = GetResourceString(390, 85) '"Sub Day book
optMature.Caption = GetResourceString(72)   '"Sub Day Book
optDepGLedger.Caption = GetResourceString(43, 93) '"Deposit GeneralLegder
optOpened.Caption = GetResourceString(64)    '"
optClosed.Caption = GetResourceString(78)
optCashBook.Caption = GetResourceString(390, 63) '"Sub Day book

optLiabilities.Caption = GetResourceString(77)
fraOrder.Caption = GetResourceString(287)
optAccId.Caption = GetResourceString(36, 60)
optName.Caption = GetResourceString(35)
optMonthlyBalance.Caption = GetResourceString(463, 42) 'Monthly balance
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
lblFrom.Caption = GetResourceString(109)
lblTo.Caption = GetResourceString(110)
cmdView.Caption = GetResourceString(13) '
'fraDateRange.Caption = GetResourceString(106)  '"Specify a Date range"
cmdView.Caption = GetResourceString(13)


'Now change the caption of properties frame
'lblint_on_rd_dep.Caption = GetResourceString(424,47,305)
'lblint_on_rd_loan.Caption = GetResourceString(424,80,186)
cmdIntApply.Caption = GetResourceString(4)
lblStatus.Caption = "" 'GetResourceString(190)
optDays.Caption = GetResourceString(44) & GetResourceString(92)
optMon.Caption = GetResourceString(192) & GetResourceString(92)
lblGenlInt = GetResourceString(344)
lblEmpInt = GetResourceString(155, 47) & GetResourceString(305)

cmdIntApply.Caption = GetResourceString(6)

lblIntPayable.Caption = GetResourceString(450, 37)
cmdIntPayable.Caption = GetResourceString(450, 171)
Me.cmdUndoPayble.Caption = GetResourceString(5, 450)
cmdAdvance.Caption = GetResourceString(491)    'Options
End Sub


'
'Write Description
'Rename The Function Name
Private Function Installment(AccId As Long) As String 'GETINSTALLMENTMONTH
    
Dim InstallmentNo As Integer
Dim MonthNo As Integer
Dim CreateDate As Date
Dim NoOFInstallments As Integer
Dim rst As ADODB.Recordset

    gDbTrans.SqlStmt = "Select * from RDMaster where accid = " & AccId
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    CreateDate = rst("CreateDate")
    NoOFInstallments = FormatField(rst("NoOfInstallments"))
    
    ' the above statement is changed by lingappa on 26th sep 2000 to convert date to
    ' American format.
    MonthNo = Month(CreateDate)
    
    InstallmentNo = 0
    gDbTrans.SqlStmt = "select * from RDTrans where AccId = " & AccId & _
        " and ( Transtype = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
    InstallmentNo = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If NoOFInstallments < InstallmentNo Then
        Installment = "---"
     '   MsgBox " AllInstallments Are Paid By Client ", vbInformation, gAppName & "ERROR"
    Else
        If (InstallmentNo + MonthNo) > 12 Then
            Installment = GetMonthString((CInt(InstallmentNo) + MonthNo) Mod 12)
        Else
            Installment = GetMonthString(InstallmentNo + CByte(MonthNo))
        End If
    End If
    Set rst = Nothing

End Function
Private Sub EnableOptionDialogControls(EnablePlaceCaste As Boolean, EnableAmountRange As Boolean)
    
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = EnablePlaceCaste
        .EnableAmountRange = EnableAmountRange
    End With

End Sub
