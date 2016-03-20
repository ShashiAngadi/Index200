VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmFDAcc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FDAcc "
   ClientHeight    =   8385
   ClientLeft      =   945
   ClientTop       =   915
   ClientWidth     =   8430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   7020
      TabIndex        =   119
      Top             =   7770
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Deposit"
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
      Left            =   4710
      TabIndex        =   86
      Top             =   150
      Width           =   3195
   End
   Begin VB.Frame fraDeposits 
      Height          =   6900
      Left            =   240
      TabIndex        =   81
      Top             =   630
      Width           =   7800
      Begin VB.TextBox txtEffective 
         Height          =   315
         Left            =   1770
         TabIndex        =   14
         Top             =   1760
         Width           =   1275
      End
      Begin VB.CommandButton cmdEffective 
         Caption         =   "..."
         Height          =   315
         Left            =   3090
         TabIndex        =   13
         Top             =   1760
         Width           =   315
      End
      Begin VB.TextBox txtCertificate 
         Height          =   315
         Left            =   5730
         TabIndex        =   11
         Top             =   1335
         Width           =   1635
      End
      Begin VB.CommandButton cmdInterest 
         Caption         =   "Withdraw interest"
         Enabled         =   0   'False
         Height          =   400
         Left            =   180
         TabIndex        =   31
         Top             =   6330
         Width           =   2145
      End
      Begin VB.CommandButton cmdMatureDate 
         Caption         =   "..."
         Height          =   315
         Left            =   7080
         TabIndex        =   16
         Top             =   1760
         Width           =   315
      End
      Begin VB.CommandButton cmdDepositDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3090
         TabIndex        =   8
         Top             =   1335
         Width           =   315
      End
      Begin VB.TextBox txtAccNo 
         Height          =   315
         Left            =   1770
         MaxLength       =   9
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   785
         Width           =   5295
      End
      Begin VB.TextBox txtDepositDate 
         Height          =   315
         Left            =   1770
         TabIndex        =   9
         Top             =   1335
         Width           =   1275
      End
      Begin VB.TextBox txtMatureDate 
         Height          =   315
         Left            =   5730
         TabIndex        =   17
         Top             =   1760
         Width           =   1305
      End
      Begin VB.TextBox txtDays 
         Height          =   315
         Left            =   1770
         TabIndex        =   19
         Top             =   2185
         Width           =   705
      End
      Begin VB.TextBox txtInterest 
         Height          =   315
         Left            =   5730
         TabIndex        =   21
         Top             =   2185
         Width           =   705
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Accept"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6390
         TabIndex        =   28
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   525
         Left            =   4530
         TabIndex        =   29
         Top             =   6240
         Width           =   1785
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3240
         TabIndex        =   30
         Top             =   6360
         Width           =   1215
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepositAmount 
         Height          =   345
         Left            =   1770
         TabIndex        =   23
         Top             =   2610
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMatureAmount 
         Height          =   345
         Left            =   5730
         TabIndex        =   26
         Top             =   2610
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2325
         Left            =   240
         TabIndex        =   115
         Top             =   3720
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
            TabIndex        =   116
            Top             =   90
            Width           =   450
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2235
            Left            =   60
            TabIndex        =   117
            Top             =   30
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   3942
            _Version        =   393217
            TextRTF         =   $"FDAcc.frx":0000
         End
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   2805
         Left            =   150
         TabIndex        =   33
         Top             =   3360
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4948
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
               Caption         =   "Account Statement"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Account Details"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraAccStmt 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2235
         Left            =   180
         TabIndex        =   82
         Top             =   3840
         Width           =   7245
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   118
            Top             =   1680
            Width           =   435
         End
         Begin VB.CommandButton cmdPrevious 
            Enabled         =   0   'False
            Height          =   405
            Left            =   6690
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   30
            Width           =   435
         End
         Begin VB.CommandButton cmdNext 
            Enabled         =   0   'False
            Height          =   405
            Left            =   6690
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   510
            Width           =   435
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   2115
            Left            =   120
            TabIndex        =   113
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3731
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"FDAcc.frx":0082
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraLedger 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2265
         Left            =   180
         TabIndex        =   87
         Top             =   3810
         Width           =   7125
         Begin MSFlexGridLib.MSFlexGrid grdLedger 
            Height          =   2115
            Left            =   180
            TabIndex        =   88
            Top             =   120
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   3731
            _Version        =   393216
         End
      End
      Begin VB.Line Line3 
         X1              =   7335
         X2              =   180
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   7260
         Y1              =   1210
         Y2              =   1210
      End
      Begin VB.Label lblEffective 
         Caption         =   "Effecitve Date"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1760
         Width           =   1425
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         Height          =   195
         Left            =   6690
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblFDBalance 
         Caption         =   "Balance : "
         Height          =   315
         Left            =   5310
         TabIndex        =   4
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblCertificate 
         Caption         =   "Certificate No :"
         Height          =   315
         Left            =   3780
         TabIndex        =   10
         Top             =   1335
         Width           =   1725
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No. : "
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblName 
         Caption         =   "Name(s) : "
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   785
         Width           =   915
      End
      Begin VB.Label lblMatureAmount 
         Caption         =   "Maturity amount (Rs) : "
         Height          =   315
         Left            =   3780
         TabIndex        =   25
         Top             =   2610
         Width           =   1725
      End
      Begin VB.Label lblDepositDate 
         Caption         =   "Deposit date :"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1335
         Width           =   1425
      End
      Begin VB.Label lblMatureDate 
         Caption         =   "Matures on : "
         Height          =   315
         Left            =   3780
         TabIndex        =   15
         Top             =   1755
         Width           =   1755
      End
      Begin VB.Label lblDays 
         Caption         =   "Days : "
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   2185
         Width           =   1215
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest (%) : "
         Height          =   315
         Left            =   3780
         TabIndex        =   20
         Top             =   2190
         Width           =   1725
      End
      Begin VB.Label lblDepositAmount 
         Caption         =   "Deposit amount (Rs) : "
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   2610
         Width           =   1515
      End
   End
   Begin VB.Frame fraNew 
      Height          =   6900
      Left            =   270
      TabIndex        =   36
      Top             =   630
      Width           =   7800
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   400
         Left            =   6420
         TabIndex        =   114
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteAcc 
         Caption         =   "Delete"
         Height          =   400
         Left            =   6420
         TabIndex        =   85
         Top             =   4980
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   400
         Left            =   6420
         TabIndex        =   46
         Top             =   5550
         Width           =   1200
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   6420
         TabIndex        =   47
         Top             =   6120
         Width           =   1200
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   1140
         Left            =   150
         ScaleHeight     =   1080
         ScaleWidth      =   5775
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   345
         Width           =   5835
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
            TabIndex        =   45
            Top             =   45
            Width           =   135
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   990
            TabIndex        =   44
            Top             =   390
            Width           =   4770
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   135
            Picture         =   "FDAcc.frx":00F7
            Stretch         =   -1  'True
            Top             =   60
            Width           =   345
         End
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   4845
         Left            =   150
         ScaleHeight     =   4785
         ScaleWidth      =   6015
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1545
         Width           =   6075
         Begin VB.PictureBox picSlider 
            Height          =   4485
            Left            =   -45
            ScaleHeight     =   4425
            ScaleWidth      =   5730
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   30
            Width           =   5790
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   0
               Left            =   3420
               TabIndex        =   49
               Top             =   45
               Width           =   465
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   4860
               TabIndex        =   41
               Top             =   0
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2355
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   -30
               Visible         =   0   'False
               Width           =   1965
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
               Left            =   2640
               TabIndex        =   42
               Top             =   0
               Width           =   3060
            End
            Begin VB.Label txtprompt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Account Holder"
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   0
               Left            =   30
               TabIndex        =   84
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   4785
            Left            =   5730
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   6465
         Width           =   1410
      End
   End
   Begin VB.Frame fraProps 
      Height          =   6900
      Left            =   270
      TabIndex        =   50
      Top             =   630
      Width           =   7800
      Begin VB.TextBox txtIntPayable 
         Height          =   315
         Left            =   2220
         TabIndex        =   108
         Top             =   4380
         Width           =   1365
      End
      Begin VB.CommandButton cmdIntPayable 
         Caption         =   "Update Interest Payable"
         Height          =   405
         Left            =   210
         TabIndex        =   107
         Top             =   4830
         Width           =   3255
      End
      Begin VB.CommandButton cmdUndoIntPayble 
         Caption         =   "Undo Interest Payable"
         Height          =   405
         Left            =   4290
         TabIndex        =   106
         Top             =   4860
         Width           =   3255
      End
      Begin VB.Frame fraInterest 
         Caption         =   "Interest rates (%)"
         Height          =   4245
         Left            =   180
         TabIndex        =   89
         Top             =   60
         Width           =   7425
         Begin VB.TextBox txtIntDate 
            Height          =   345
            Left            =   1740
            TabIndex        =   24
            Top             =   3244
            Width           =   1365
         End
         Begin VB.TextBox txtSenInt 
            Height          =   315
            Left            =   2310
            TabIndex        =   105
            Top             =   2420
            Width           =   795
         End
         Begin VB.TextBox txtEmpInt 
            Height          =   315
            Left            =   2310
            TabIndex        =   101
            Top             =   2008
            Width           =   795
         End
         Begin VB.TextBox txtGenInt 
            Height          =   315
            Left            =   2310
            TabIndex        =   100
            Top             =   1596
            Width           =   795
         End
         Begin VB.ComboBox cmbTo 
            Height          =   315
            Left            =   1770
            TabIndex        =   99
            Top             =   1184
            Width           =   1335
         End
         Begin VB.ComboBox cmbFrom 
            Height          =   315
            Left            =   150
            TabIndex        =   98
            Top             =   1184
            Width           =   1335
         End
         Begin VB.OptionButton optDays 
            Caption         =   "Days"
            Height          =   300
            Left            =   210
            TabIndex        =   97
            Top             =   390
            Width           =   1335
         End
         Begin VB.OptionButton optMon 
            Caption         =   "Month"
            Height          =   300
            Left            =   1770
            TabIndex        =   96
            Top             =   390
            Width           =   1335
         End
         Begin VB.TextBox txtLoanInt 
            Height          =   315
            Left            =   2310
            TabIndex        =   91
            Text            =   "+"
            Top             =   2832
            Width           =   795
         End
         Begin VB.CommandButton cmdIntApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   400
            Left            =   1890
            TabIndex        =   90
            Top             =   3690
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid grdInt 
            Height          =   3795
            Left            =   3210
            TabIndex        =   93
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6694
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            AllowUserResizing=   3
         End
         Begin VB.Label lblIntApply 
            Caption         =   "Int apply date"
            Height          =   555
            Left            =   120
            TabIndex        =   27
            Top             =   3244
            Width           =   1455
         End
         Begin VB.Label lblSenInt 
            Caption         =   "Senior Citizen"
            Height          =   300
            Left            =   120
            TabIndex        =   104
            Top             =   2420
            Width           =   1905
         End
         Begin VB.Label lblEmpInt 
            Caption         =   "Emplyees Interest Rate"
            Height          =   300
            Left            =   120
            TabIndex        =   103
            Top             =   2008
            Width           =   1965
         End
         Begin VB.Label lblGenlInt 
            Caption         =   "General Interest"
            Height          =   300
            Left            =   120
            TabIndex        =   102
            Top             =   1596
            Width           =   1995
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   300
            Left            =   1770
            TabIndex        =   95
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label lblFrom 
            Caption         =   "from"
            Height          =   300
            Left            =   210
            TabIndex        =   94
            Top             =   787
            Width           =   1035
         End
         Begin VB.Label lblLoanInt 
            Caption         =   "Max loan percent:"
            Height          =   300
            Left            =   120
            TabIndex        =   92
            Top             =   2832
            Width           =   1965
         End
      End
      Begin ComctlLib.ProgressBar prg 
         Height          =   315
         Left            =   180
         TabIndex        =   109
         Top             =   5910
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label txtLastIntDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   6210
         TabIndex        =   110
         Top             =   4410
         Width           =   1365
      End
      Begin VB.Label lblIntPayable 
         Caption         =   "Interest Payable Date :"
         Height          =   315
         Left            =   150
         TabIndex        =   79
         Top             =   4410
         Width           =   2205
      End
      Begin VB.Label lblLastIntDate 
         Caption         =   "Last Interest Updated on :"
         Height          =   315
         Left            =   4020
         TabIndex        =   80
         Top             =   4440
         Width           =   2235
      End
      Begin VB.Label lblStatus 
         Caption         =   "x"
         Height          =   375
         Left            =   210
         TabIndex        =   112
         Top             =   5430
         Width           =   6705
      End
      Begin VB.Label txtFailAccIds 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   180
         TabIndex        =   111
         Top             =   6270
         Width           =   7275
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7545
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   13309
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transactions"
            Key             =   ""
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
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Properties"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraReports 
      Height          =   6900
      Left            =   270
      TabIndex        =   83
      Top             =   630
      Width           =   7800
      Begin VB.Frame fraDateRange 
         Caption         =   "Sepcify a date range"
         Height          =   1800
         Left            =   150
         TabIndex        =   68
         Top             =   4560
         Width           =   7545
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   400
            Left            =   6120
            TabIndex        =   77
            Top             =   1290
            Width           =   1215
         End
         Begin VB.OptionButton optName 
            Caption         =   "Name "
            Height          =   315
            Left            =   3870
            TabIndex        =   70
            Top             =   390
            Width           =   1590
         End
         Begin VB.OptionButton optAccID 
            Caption         =   "Account No"
            Height          =   315
            Left            =   285
            TabIndex        =   69
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.CommandButton cmdDate2 
            Caption         =   "..."
            Height          =   315
            Left            =   7050
            TabIndex        =   75
            Top             =   780
            Width           =   315
         End
         Begin VB.CommandButton cmdDate1 
            Caption         =   "..."
            Height          =   315
            Left            =   3030
            TabIndex        =   72
            Top             =   900
            Width           =   315
         End
         Begin VB.TextBox txtToDate 
            Height          =   315
            Left            =   5400
            TabIndex        =   76
            Top             =   900
            Width           =   1245
         End
         Begin VB.TextBox txtFromDate 
            Height          =   315
            Left            =   1740
            TabIndex        =   73
            Top             =   900
            Width           =   1245
         End
         Begin VB.Line Line4 
            X1              =   7020
            X2              =   60
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Label lblDate2 
            Caption         =   "but before (dd/mm/yyyy)"
            Height          =   315
            Left            =   3480
            TabIndex        =   74
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   1485
         End
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report"
         Height          =   4395
         Left            =   150
         TabIndex        =   51
         Top             =   240
         Width           =   7545
         Begin VB.OptionButton optMFD 
            Caption         =   "List of Matured  FD"
            Height          =   315
            Left            =   3840
            TabIndex        =   60
            Top             =   315
            Width           =   3135
         End
         Begin VB.OptionButton optMFdCashBook 
            Caption         =   "Mat FD sub Cash Book"
            Height          =   315
            Left            =   3840
            TabIndex        =   62
            Top             =   1329
            Width           =   3105
         End
         Begin VB.OptionButton optDepCashBook 
            Caption         =   "Show Cash Book"
            Height          =   315
            Left            =   300
            TabIndex        =   54
            Top             =   1329
            Width           =   3345
         End
         Begin VB.OptionButton optMature 
            Caption         =   "Deposit that mature"
            Height          =   315
            Left            =   300
            TabIndex        =   59
            Top             =   3870
            Width           =   3225
         End
         Begin VB.OptionButton optMFDTrans 
            Caption         =   "Matured Deposit Transaction Made"
            Height          =   315
            Left            =   3840
            TabIndex        =   64
            Top             =   2343
            Width           =   3135
         End
         Begin VB.OptionButton optDepTrans 
            Caption         =   "Deposit Trans Made"
            Height          =   315
            Left            =   300
            TabIndex        =   56
            Top             =   2343
            Width           =   3225
         End
         Begin VB.OptionButton optMatDepDtCr 
            Caption         =   "Matured Deposit sub day book"
            Height          =   315
            Left            =   3840
            TabIndex        =   61
            Top             =   822
            Width           =   3135
         End
         Begin VB.OptionButton optMatDepGLedger 
            Caption         =   "Mat Depoist General Ledger"
            Height          =   315
            Left            =   3840
            TabIndex        =   63
            Top             =   1836
            Width           =   3135
         End
         Begin VB.OptionButton optJoint 
            Caption         =   "Joint Accounts "
            Height          =   315
            Left            =   3840
            TabIndex        =   65
            Top             =   2850
            Width           =   3135
         End
         Begin VB.OptionButton optMonthBal 
            Caption         =   "Deposits Monthly Balance"
            Height          =   315
            Left            =   3840
            TabIndex        =   67
            Top             =   3870
            Width           =   3135
         End
         Begin VB.OptionButton optDepositBalance 
            Caption         =   "List Of Deposits "
            Height          =   315
            Left            =   300
            TabIndex        =   52
            Top             =   315
            Value           =   -1  'True
            Width           =   3225
         End
         Begin VB.OptionButton optDepGLedger 
            Caption         =   "Deposit General Ledger"
            Height          =   315
            Left            =   300
            TabIndex        =   55
            Top             =   1836
            Width           =   3225
         End
         Begin VB.OptionButton optDepDtCr 
            Caption         =   "Sub Day book"
            Height          =   315
            Left            =   300
            TabIndex        =   53
            Top             =   822
            Width           =   3225
         End
         Begin VB.OptionButton optLiabilities 
            Caption         =   "Liabilities"
            Height          =   315
            Left            =   300
            TabIndex        =   57
            Top             =   2850
            Width           =   3225
         End
         Begin VB.OptionButton optOpened 
            Caption         =   "Deposits opened"
            Height          =   315
            Left            =   300
            TabIndex        =   58
            Top             =   3357
            Width           =   3225
         End
         Begin VB.OptionButton optClosed 
            Caption         =   "Deposits closed"
            Height          =   315
            Left            =   3840
            TabIndex        =   66
            Top             =   3357
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   400
         Left            =   6480
         TabIndex        =   78
         Top             =   6390
         Width           =   1215
      End
   End
   Begin VB.Label lblFDName 
      Alignment       =   2  'Center
      Caption         =   "Fixed Deposit"
      Height          =   475
      Left            =   240
      TabIndex        =   120
      Top             =   7800
      Width           =   6615
   End
End
Attribute VB_Name = "frmFDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_AccHeadId As Long

Private m_clsRepOption As clsRepOption
Attribute m_clsRepOption.VB_VarHelpID = -1

Private m_AccID As Long
Private m_AccNum As String
Private m_CustReg As clsCustReg
Private m_LastDepsoitNo As Integer
Private m_FirstDepsoitNo As Integer
Private M_setUp As New clsSetup
Private m_accUpdatemode As Integer
Private m_JointCustID(3) As Long
Private m_frmFDJoint As frmJoint
Private m_Transferred As Boolean
Private m_Cumulative As Boolean

Private m_DepositName As String
Private m_DepositNameEnglish As String
Private m_DepositType As Integer
Private M_ModuleID As wisModules
Private m_Notes As New clsNotes

Private m_IntHeadID As Long

Private Const CTL_MARGIN = 15

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1

Public Event SelectDepositType(ByRef DepositType As Integer)
Public Event AccountChanged(ByVal AccId As Long)
Public Event AccountTransaction(transType As wisTransactionTypes)
Public Event ShowReport(ReportType As wis_FDReports, ReportOrder As wis_ReportOrder, _
             fromDate As String, toDate As String, _
             RepOptClass As clsRepOption)

Public Event PayInterest(ByVal AccId As Long, Cancel As Integer)
Public Event FDClose(ByVal AccId As Long, Cancel As Integer)
Public Event WindowClosed()
Public Event FDRenew(ByVal AccId As Long, Cancel As Integer)
Public Event AddInterestPayable(ByVal AsOnDate As Date, ByVal TransDate As Date)

Private Function AcceptTransaction() As Boolean
Dim TransDate As String
TransDate = GetSysFormatDate(txtDepositDate.Text)

Dim EffectiveDate As String
EffectiveDate = GetSysFormatDate(txtEffective.Text)

Dim MatureDate As String
MatureDate = GetSysFormatDate(txtMatureDate.Text)

Dim Interest As Double
Interest = Val(txtInterest.Text)

Dim DepositAmount As Currency
DepositAmount = CCur(Trim$(txtDepositAmount.Text))

Dim MatureAmount As Currency
MatureAmount = CCur(Trim$(txtMatureAmount.Text))
'CERTIFICATE NO
'Get the Particulars(Certificate no)
Dim Particulars As String * 49
Particulars = Trim(txtCertificate.Text)
'    If Len(Particulars) > 49 Then Particulars = Left(Particulars, 49)
'Get the Deposit Number
Dim NewDeposit As Boolean
Dim rst As ADODB.Recordset
gDbTrans.SqlStmt = "Select AccID,CustomerID from FDMaster Where " & _
        " AccNum = " & AddQuotes(m_AccNum, True) & _
        " AND FDMaster.DepositType = " & m_DepositType
Dim AccId As Long
Dim customerID As Long
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
AccId = Val(FormatField(rst("AccId")))
customerID = Val(FormatField(rst("CustomerId")))

'If any transaction has done  on this deposit then
'we have to create new deposit else update the existing one
gDbTrans.SqlStmt = "SELECT TransID From FDTrans WHERE AccID = " & AccId
NewDeposit = IIf(gDbTrans.Fetch(rst, adOpenStatic) > 0, True, False)

'Perform the Transaction to the Database
   'We are treating one depoist as one account
   'So First We have to insert the details into FDMAster
   'IF it is first deposit then we have to update the FD MAster
   'ELSE we have insert new record in to the FDMAster
   'if it is first depoist then update the FDMaster,
   ' because the Depoisit one has created in FDMaster
   ' while creating Account Details
If Not NewDeposit Then
    gDbTrans.SqlStmt = "Update FDMaster Set " & _
        " CreateDate =  #" & TransDate & "# ," & _
        " EffectiveDate =  #" & TransDate & "# ," & _
        " MaturityDate = #" & MatureDate & "# ," & _
        " DepositAmount = " & DepositAmount & ", " & _
        " MaturityAmount = " & MatureAmount & " ," & _
        " CertificateNo = " & AddQuotes(Particulars, True) & ", " & _
        " RateOfInterest = " & Interest & ", " & _
        " ClosedDate = NULL " & _
        " Where AccId = " & AccId

Else
    
    'Now Check For The Joint Account deatial
    gDbTrans.SqlStmt = "Select * From FDJoint " & _
        " Where AccId = " & AccId
    Dim rstJoint As Recordset
    If gDbTrans.Fetch(rstJoint, adOpenForwardOnly) <= 1 Then Set rstJoint = Nothing
    
    gDbTrans.SqlStmt = "Select Max(AccId) From FDMaster"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    AccId = Val(FormatField(rst(0))) + 1
    gDbTrans.SqlStmt = "Insert into FDMaster (AccNum,AccID, " & _
        " CustomerID, CreateDate, EffectiveDate,CertificateNo,MaturityDate, " & _
        " DepositAmount,MaturityAmount, RateOfInterest, DepositType,AccGroupID,UserID) " & _
        " values ( " & _
        AddQuotes(m_AccNum, True) & ", " & _
        AccId & "," & _
        customerID & ", " & _
        "#" & TransDate & "#," & _
        "#" & EffectiveDate & "#," & AddQuotes(Particulars, True) & ", " & _
        "#" & MatureDate & "#," & _
        DepositAmount & ", " & MatureAmount & "," & Interest & "," & _
        m_DepositType & "," & GetAccGroupID & "," & gUserID & " )"
End If
    
 'Fire the query
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If
'If this account is of joint account
'then insert the details
If Not rstJoint Is Nothing Then
    While Not rstJoint.EOF
        gDbTrans.SqlStmt = "Insert INTO FDJoint " & _
            " (AccID,CustomerID, CustomerNum) VALUES (" & _
            AccId & ", " & _
            rstJoint("CustomerID") & ", " & _
            rstJoint("CustomerNum") & ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
        rstJoint.MoveNext
    Wend
End If
'Now insert the Depoist Amount & othe details in to the FDTrans Table
'It Is Deposit So Loan is False
'Set the Trans Type
'(a value of 1 if amount goes into account)
Dim Trans As wisTransactionTypes
Trans = wDeposit
'Get the new Transaction ID
'This is a new deposit, and will have a transaction ID of 1
Dim TransID As Long

TransID = 1
'Get the Depoist type
gDbTrans.SqlStmt = "Insert into FDTrans (AccID, " & _
        " TransID, TransType, " & _
        " TransDate, Amount, Balance, " & _
        " Particulars,UserID) values ( " & _
        AccId & "," & _
        TransID & "," & _
        Trans & "," & _
        "#" & TransDate & "#," & _
        DepositAmount & "," & DepositAmount & "," & _
        AddQuotes(Particulars, True) & _
        "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

If m_AccHeadId = 0 Then
    MsgBox "unable make the transaction", vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
    Exit Function
End If

Call bankClass.UpdateCashDeposits(m_AccHeadId, DepositAmount, TransDate)

gDbTrans.CommitTrans

Set bankClass = Nothing
AcceptTransaction = True

End Function

Private Function AccountDelete() As Boolean

Dim SqlStr As String
Dim SqlAcc As String
Dim rst As ADODB.Recordset
Dim AccId As Long

'Prelim Checks
    If m_AccID <= 0 Then
        'MsgBox "Deposit not loaded !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(523), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

SqlAcc = "SELECT AccID From FDMAster " & _
        " Where AccNum = " & AddQuotes(m_AccNum, True) & _
        " AND FDMaster.DepositType = " & m_DepositType

'You can delete a deposit only if you do not have transactions more than one
SqlStr = "Select * from FDTrans where " & _
            " AccID IN (" & SqlAcc & ") Order by TransID desc "
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        'MsgBox "You cannot delete a deposit with transactions.", _
        vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(530) & vbCrLf & _
            GetResourceString(553), vbExclamation, gAppName & " - Error"
        Exit Function
    End If


'Now Check transactions In INtrest account
    SqlStr = "Select * from FDIntTrans where " & _
            " AccID = " & m_AccID & " Order by TransID DESC "
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Set rst = Nothing
        Exit Function
    End If
'Now Check transactions In IntrestPayable account
    SqlStr = "Select * from FDIntPayable where " & _
            " AccID = " & m_AccID & " Order by TransID desc "
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Set rst = Nothing
        Exit Function
    End If
        
'Beging DB Transactions
Dim InTrans As Boolean
InTrans = True

    gDbTrans.BeginTrans
    
'Fire SQL Stmt
    'First Delate the Related record in FDMaster
    gDbTrans.SqlStmt = "Delete * from FDMaster " & _
              "where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
    
    'Then Delete the Deposited Amount On from FDTrans
    gDbTrans.SqlStmt = "Delete from FDTrans where AccID = " & m_AccID
    
    If Not gDbTrans.SQLExecute Then GoTo ExitLine

'Then Delete the Amount On from FD Interest
    gDbTrans.SqlStmt = "Delete from FDIntTrans where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
'Then Delete the Deposited Amount On from Payable
    gDbTrans.SqlStmt = "Delete from FDIntPayable where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine


'Close DB Operations
gDbTrans.CommitTrans

AccountDelete = True

ExitLine:

    If InTrans Then
        gDbTrans.RollBack
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
    End If

End Function
Private Function CreateFDLedgerView() As Boolean
    
gDbTrans.SqlStmt = "SELECT 'PRINCIPAL',TransID, TransDate, Amount, TransType " & _
    "FROM FDTrans " & _
    "WHERE AccID = " & m_AccID & _
    " UNION ALL SELECT 'INTEREST',TransID, TransDate, Amount, TransType+10 " & _
    "FROM FDIntTrans " & _
    "WHERE AccID = " & m_AccID & _
    " UNION ALL SELECT 'PAYABLE',TransID, TransDate, Amount, TransType+20 FROM FDIntPayable " & _
    "WHERE TransType = " & wWithdraw & _
    " AND AccID = " & m_AccID

If gDbTrans.CreateView("QryFDLedger") Then CreateFDLedgerView = True

End Function

Public Property Let DepositType(NewValue As Integer)
    
    Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "SELECT * FROM DepositName WHERE " & _
        " DepositID = " & NewValue
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Property
    
    m_DepositType = NewValue
    m_Cumulative = FormatField(rst("Cumulative"))
    m_DepositName = FormatField(rst("DepositName"))
    m_DepositNameEnglish = FormatField(rst("DepositNameEnglish"))
    M_ModuleID = wis_Deposits + (m_DepositType)
    If M_ModuleID >= wis_Deposits * 2 Then M_ModuleID = M_ModuleID - wis_Deposits
    
    cmdSelect.Caption = m_DepositName
    lblFDName.Caption = m_DepositName
    lblFDName.ToolTipText = m_DepositNameEnglish
    
    Call LoadProperties
    
    Call LoadInterestRates
    
    'Get the Account Head ID
    Dim ClsBank As New clsBankAcc
    gDbTrans.BeginTrans
    m_AccHeadId = ClsBank.GetHeadIDCreated(m_DepositName, m_DepositNameEnglish, parMemberDeposit, 0, M_ModuleID)
    gDbTrans.CommitTrans
    Set ClsBank = Nothing

End Property

Private Sub EnableInterestApply()
cmdIntApply.Enabled = False

If cmbFrom.ListIndex < 0 Then Exit Sub
If cmbTo.ListIndex < 0 Then Exit Sub
If Val(txtGenInt) = 0 Then Exit Sub
If Len(txtEmpInt) And Val(txtEmpInt) = 0 Then Exit Sub
If Len(txtSenInt) And Val(txtSenInt) = 0 Then Exit Sub

Dim Perms As wis_Permissions
Perms = gCurrUser.UserPermissions
'cmdIntApply.Enabled = True
cmdIntApply.Enabled = CBool((Perms And perBankAdmin) Or (Perms And perOnlyWaves))

End Sub

Private Sub InitLedgerGrid()

With grdLedger
    .AllowUserResizing = flexResizeBoth
    .Clear: .Rows = 2: .FixedRows = 1: .FixedCols = 1: .Cols = 6
    .Row = 0
    .Col = 0: .Text = GetResourceString(33): .ColWidth(0) = 400       '"SLnO"
    .Col = 1: .Text = GetResourceString(37): .ColWidth(1) = 1250       ' "Date"
    .Col = 2: .Text = GetResourceString(39): .ColWidth(2) = 1400       '"Particulars"
    .Col = 3: .Text = GetResourceString(276): .ColWidth(3) = 1050       '"Debit"
    .Col = 4: .Text = GetResourceString(277): .ColWidth(4) = 1050       '"Credit"
    .Col = 5: .Text = GetResourceString(42): .ColWidth(5) = 1050       '"Balance"
End With

End Sub

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
        retstr = SetUp.ReadSetupValue("DEPOSIT" & m_DepositType, strKey, "")
'        If retstr = "" Then Exit For
        If retstr <> "" Then
            
            strTo = cmbTo.List(I)
            If Prevstr <> retstr Then
                If Prevstr = "" Then StrFrom = cmbFrom.List(I)
                If .Rows = .Row + 1 Then .Rows = .Rows + 1
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
            Prevstr = retstr
        End If
'        Prevstr = retstr
    Next
    
    optMon.Value = True
    MaxI = cmbFrom.ListCount - 1
    'strFrom = cmbFrom.List(0)
    For I = 0 To MaxI
        strKey = "YEAR" & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
        'strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
        retstr = SetUp.ReadSetupValue("DEPOSIT" & m_DepositType, strKey, "")
        If retstr <> "" Then
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
        End If
    Next
    
End With

End Sub

Private Sub LoadProperties()

' Set the label & text properties wherever required.
    txtDepositDate.Text = gStrDate
    txtEffective = txtDepositDate
    txtRTF.Text = ""
    
    Dim rst As ADODB.Recordset
    'Get The Last Interest  Updated date
    gDbTrans.SqlStmt = "Select Max(StartDate) as LastIntDate " & _
          "From InterestTab Where EndDate is Null " & _
          " And Moduleid = " & M_ModuleID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
              txtLastIntDate = FormatField(rst("LastIntDate"))
    Set rst = Nothing
    cmdIntApply.Enabled = False

'Disable UI
    ResetUserInterface
'Set the report Form
optDepositBalance.Value = True
Call optDepositBalance_Click
Screen.MousePointer = vbDefault

Dim ClsInt As clsInterest
Dim TransDate As Date

TransDate = CDate(gStrDate)
 'Load properties
  With M_setUp
      'txtLoanPercent.Text = .ReadSetupValue("FDAcc" & CStr(m_DepositType), _
        "MaxLoanPercent", Val(txtLoanPercent.Text))
  End With

fraAccStmt.ZOrder 0
'fraAccStmt.Visible = True
'fraLedger.Visible = False
InitLedgerGrid
LoadDeposits

cmdIntApply.Enabled = False

End Sub

Private Sub ShowLedger()
Dim rst As ADODB.Recordset
Dim SlNo As Long
Dim transType As wisTransactionTypes
Dim Balance As Long
Dim Particulars As String

If Not CreateFDLedgerView Then Exit Sub
gDbTrans.SqlStmt = "SELECT * FROM QryFDLedger ORDER BY TransID"
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

InitLedgerGrid

grdLedger.Row = 1
SlNo = 1
Dim Payable As Boolean
Dim Amount As Currency
Dim PrevAmount As Currency

Do While Not rst.EOF
With grdLedger
    
    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    
    transType = FormatField(rst("TransType"))
    Amount = FormatField(rst("Amount"))
    
    Select Case transType
        Case wDeposit, wContraDeposit
            Particulars = "Amount deposited"
            Balance = Balance + Amount
            .Col = 3
        Case wWithdraw, wContraWithdraw
            Balance = 0 'Balance + Amount
            Particulars = "Account Closed"
            .Col = 4
        
        'Detail Of Interset Account
        Case wDeposit + 10, wContraDeposit + 10
            Particulars = "Interest recovered"
            .Col = 3: .CellForeColor = vbRed
        Case wWithdraw + 10, wContraWithdraw + 10
            Particulars = "Interest Paid"
            .Col = 4: '.CellForeColor = vbRed
            If Payable Then
                .Row = .Row - 1: SlNo = SlNo - 1
                Payable = False
                Particulars = "Interest to Payable A/c"
                .CellForeColor = vbBlue
                
            End If
        
        'Details of Interest Payable account
        Case wDeposit + 20, wContraDeposit + 20
            Particulars = "Interest to Payable"
            .Col = 3
            Payable = True
            If Amount = PrevAmount Then
                .Row = .Row - 1: SlNo = SlNo - 1
                Payable = False
                Particulars = "Interest to Payable A/c"
                .CellForeColor = vbBlue
            End If
        Case wWithdraw + 20, wContraWithdraw + 20
            Particulars = "Interest From Payable"
            .Col = 4: .CellForeColor = vbRed
            
            
    End Select
    
    
    .Text = FormatCurrency(Amount): .CellAlignment = 7
    
    .Col = 0: .Text = SlNo
    .Col = 1: .Text = FormatField(rst("TransDate"))
    .Col = 2: .Text = Particulars
    .Col = 5: .Text = FormatCurrency(Balance): .CellAlignment = 7
    rst.MoveNext
    SlNo = SlNo + 1
    .Row = .Row + 1
    PrevAmount = Amount
    
End With
Loop
    

'fraLedger.ZOrder 0

End Sub

Private Function UndoInterestPayableOfFD(OnIndianDate As String) As Boolean
Dim rstPayble As ADODB.Recordset
Dim rstInt As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim DimPos As Integer
Dim OnDate As Date
Dim transType As wisTransactionTypes
Dim strParticulars As String

lblStatus = ""

If m_DepositType = 0 Then
    MsgBox "SELECT THE DEPOSIT", vbInformation
    'ActivateTextBox me.txtSelect
    Exit Function
End If

OnDate = GetSysFormatDate(OnIndianDate)
strParticulars = IIf(m_Cumulative, "'Interest Added To Account'", "'Interest Payable'")

DimPos = InStr(1, OnIndianDate, "31/3/", vbTextCompare)
If DimPos = 0 Then DimPos = InStr(1, OnIndianDate, "31/03/", vbTextCompare)
If DimPos = 0 Then
    'MsgBox "Invalid date,Do u want to continuew, vbInformation, wis_MESSAGE_TITLE
    If MsgBox(GetResourceString(501, 541), vbInformation + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If

OnDate = GetSysFormatDate(OnIndianDate)

'Before undoing check whether he has already added the interestpayble amount or not
transType = wContraWithdraw
gDbTrans.SqlStmt = "Select * from FDIntTrans Where " & _
    " TransDate = #" & OnDate & "# " & _
    " And Particulars = " & strParticulars

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    UndoInterestPayableOfFD = True
    Exit Function
End If
  
Me.MousePointer = vbHourglass
On Error GoTo ErrLine
'declare the variables necessary

'Get the Payble Amount
If m_Cumulative Then
    gDbTrans.SqlStmt = "SELECT SUM(A.Amount) From FdIntTrans A" & _
        " WHERE A.TransID = " & _
            "(SELECT TransID FROM FDIntTrans C WHERE" & _
            " Particulars = " & strParticulars & " AND TransDate = #" & OnDate & "#" & _
            " AND C.AccID = A.AccID) AND TransDate = #" & OnDate & "#" & _
        " AND A.TransID = (SELECT Max(TransID) FROM FDTrans E WHERE " & _
            " A.AccID = E.AccID)"
Else
    gDbTrans.SqlStmt = "SELECT SUM(A.Amount) From FdIntPayable A" & _
        " WHERE A.TransID = " & _
            "(SELECT TransID FROM FDIntTrans C WHERE" & _
            " Particulars = " & strParticulars & " AND TransDate = #" & OnDate & "#" & _
            " AND C.AccID = A.AccID) AND TransDate = #" & OnDate & "#" & _
        " AND A.TransID > (SELECT Max(TransID) FROM FDTrans E WHERE " & _
            " A.AccID = E.AccID)"
End If
'Dim Rst As Recordset
Dim Amount As Currency

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ErrLine
Amount = FormatField(rst(0))
If Amount Then
    If MsgBox("You are withdrawing the Rs." & Amount & " from " & m_DepositName & _
            IIf(m_Cumulative, "", GetResourceString(375, 47)) & _
        vbCrLf & GetResourceString(541), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
Else
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    UndoInterestPayableOfFD = True
    Exit Function
End If

If m_Cumulative Then
    gDbTrans.SqlStmt = "DELETE A.*, B.* From FdTrans A," & _
        " FDIntTrans B WHERE A.AccID = B.AccID AND A.TransID = " & _
            "(SELECT TransID FROM FDIntTrans C WHERE" & _
            " Particulars = " & strParticulars & " AND TransDate = #" & OnDate & "#" & _
            " AND C.AccID = A.AccID) ANd B.TransID = A.TransID " & _
        " AND A.TransID = (SELECT Max(TransID) FROM FDTrans E WHERE " & _
            " A.AccID = E.AccID)"
Else
    gDbTrans.SqlStmt = "DELETE A.*, B.* From FdIntPayable A," & _
        " FDIntTrans B WHERE A.AccID = B.AccID AND A.TransID = " & _
            "(SELECT TransID FROM FDIntTrans C WHERE" & _
            " Particulars = " & strParticulars & " AND TransDate = #" & OnDate & "#" & _
            " AND C.AccID = A.AccID) ANd B.TransID = A.TransID " & _
        " AND A.TransID > (SELECT Max(TransID) FROM FDTrans E WHERE " & _
            " A.AccID = E.AccID)"
End If

gDbTrans.BeginTrans
gDbTrans.SQLExecute

Dim bankClass As clsBankAcc
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim headName As String
Dim headNameEnglish As String

'Now Get the ledger head id of the payble/interest and Depsoit
headName = m_DepositName & " " & GetResourceString(450)
headNameEnglish = m_DepositNameEnglish & " " & LoadResString(450)
Set bankClass = New clsBankAcc
FromHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemDepIntPaid, 0, wis_Deposits)
If m_Cumulative Then
    ToHeadID = bankClass.GetHeadIDCreated(m_DepositName, m_DepositNameEnglish, parMemberDeposit, 0, wis_Deposits)
Else
    headName = m_DepositName & " " & GetResourceString(375, 47)
    headNameEnglish = m_DepositNameEnglish & " " & LoadResourceStringS(375, 47)
    ToHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parDepositIntProv, 0, wis_Deposits)
End If

'Now Make the transaction to the ledger heads
If Not bankClass.UndoContraTrans(FromHeadID, ToHeadID, Amount, OnDate) Then
    gDbTrans.RollBack
    GoTo ExitLine
End If

gDbTrans.CommitTrans

UndoInterestPayableOfFD = True

'now Check If Any  Account are unable to the undo
gDbTrans.SqlStmt = "Select AccNum from FDMaster A,FDIntTrans B Where " & _
    " A.AccId = B.accID And TransDate = #" & OnDate & "# " & _
    " And B.Particulars = " & strParticulars

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo ExitLine

While Not rst.EOF
    txtFailAccIDs = txtFailAccIDs & "," & rst("AccNum")
    rst.MoveNext
Wend

txtFailAccIDs = Mid(txtFailAccIDs, 2)
Set rst = Nothing

GoTo ExitLine

ErrLine:
Set bankClass = Nothing
MsgBox "Error In FDAccount -- Remove Interest payble", vbCritical, wis_MESSAGE_TITLE
Me.MousePointer = vbDefault

ExitLine:
'Resume
'Set BankClass = Nothing

Set rst = Nothing
Set rstInt = Nothing
Set rstPayble = Nothing

End Function
Private Function AccountExists(ByVal AccId As Long, Optional ClosedON As String) As Boolean
Dim ret As Integer
Dim rst As ADODB.Recordset

'Query Database

    gDbTrans.SqlStmt = "Select AccID from FDMaster where " & _
                        " AccID = " & AccId
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function

    If ret < 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    Set rst = Nothing
    
AccountExists = True

End Function

Public Function AccountLoad(ByVal AccountId As Long) As Boolean

Dim strField As String
Dim rstMaster As ADODB.Recordset
Dim rstDeposits As ADODB.Recordset
Dim rstJoint As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim JointHolders() As String
Dim I As Integer
Dim IntroId As Long
Dim NomineeId As Long
Dim AccId As Long
Dim ClosedDate As String

Call ResetUserInterface

gDbTrans.SqlStmt = "SELECT * FROM FDMaster WHERE AccID = " & AccountId

If gDbTrans.Fetch(rstMaster, adOpenForwardOnly) < 1 Then Exit Function
'object
m_AccID = rstMaster("ACCID")
m_AccNum = FormatField(rstMaster("ACCNUM"))
AccId = rstMaster("ACCID")
ClosedDate = FormatField(rstMaster("ClosedDate"))

'm_CustId = RstMaster("CustomerID")

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
'Load the Name details
If Not m_CustReg.LoadCustomerInfo(FormatField(rstMaster("CustomerID"))) Then
    'MsgBox "Unable to load customer information !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
    Call ResetUserInterface
    Exit Function
End If

'Now Fetch the details of the Nominee
NomineeId = FormatField(rstMaster("NomineeID"))
IntroId = FormatField(rstMaster("Introduced"))

'Now Fetch The Details of the Joint AccountHolder
gDbTrans.SqlStmt = "Select * from FDJoint where AccId = " & m_AccID
    'AddQuotes(m_AccNum, True)
If gDbTrans.Fetch(rstJoint, adOpenForwardOnly) > 0 Then
    If m_frmFDJoint Is Nothing Then Set m_frmFDJoint = New frmJoint
    m_frmFDJoint.ModuleID = wis_Deposits
End If

'Update TAB 2
m_accUpdatemode = wis_UPDATE
'Enable controls on the UI
    'Fill Name
With cmbNames
    .Enabled = True: .BackColor = vbWhite: .Clear
    .AddItem m_CustReg.FullName
    
    If Not rstJoint Is Nothing Then
        I = 0
        While Not rstJoint.EOF
            .AddItem m_CustReg.CustomerName(rstJoint("CustomerID"))
            m_JointCustID(I) = FormatField(rstJoint("CustomerID"))
            m_frmFDJoint.JointCustId(I) = m_JointCustID(I)
            rstJoint.MoveNext: I = I + 1
        Wend
    End If
    .ListIndex = 0
End With
With txtDepositDate
    .BackColor = vbWhite: .Enabled = True
End With
cmdDepositDate.Enabled = True
If gOnLine Then cmdDepositDate.Enabled = False

With txtMatureDate
    .BackColor = vbWhite: .Enabled = True
End With
cmdMatureDate.Enabled = True
With txtEffective
    .BackColor = vbWhite: .Enabled = True
End With
With cmdEffective
    .Enabled = True
End With
With txtCertificate
    .BackColor = vbWhite: .Enabled = True
End With
With txtDays
    .BackColor = vbWhite: .Enabled = True
End With
With txtInterest
    .BackColor = vbWhite: .Enabled = True
End With
With txtDepositAmount
    .BackColor = vbWhite: .Enabled = True
End With
With txtMatureAmount
    .BackColor = vbWhite: .Enabled = True
End With

lblOperation.Caption = GetResourceString(56) ' "Operation Mode : <UPDATE>"

For I = 0 To txtPrompt.count - 1
    ' Read the bound field of this control.
    On Error Resume Next
    strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If strField <> "" Then
        With txtData(I)
          Select Case UCase$(strField)
            Case "ACCID"
                .Text = rstMaster("AccNum")
                .Locked = True
            Case "ACCNAME"
                .Text = m_CustReg.FullName
            Case "NOMINEEID"
                .Text = IIf(NomineeId, NomineeId, "")
            Case "NOMINEENAME"
                .Text = IIf(NomineeId, m_CustReg.CustomerName(NomineeId), FormatField(rstMaster("NomineeName")))
            Case "NOMINEERELATION"
                .Text = FormatField(rstMaster("NomineeRelation"))
            Case "JOINTHOLDER"
                .Text = ""
                If Not m_frmFDJoint Is Nothing Then _
                    .Text = m_frmFDJoint.JointHolders
            Case "CERTIFICATENO"
                .Text = FormatField(rstMaster("CertificateNo"))
            Case "DEPAMOUNT"
                .Text = FormatField(rstMaster("DepositAmount"))
            Case "MATAMOUNT"
                .Text = FormatField(rstMaster("MaturityAmount"))
            Case "MATAMOUNT"
                .Text = FormatField(rstMaster("MaturityAmount"))
            Case "INTRODUCERID"
                .Text = IIf(IntroId > 0, IntroId, "")
            Case "INTRODUCERNAME"
                .Text = IIf(IntroId, m_CustReg.CustomerName(IntroId), "")
            Case "LEDGERNO"
                .Text = rstMaster("LedgerNo")
            Case "FOLIONO"
                .Text = rstMaster("FolioNO")
            Case "CREATEDATE"
                .Text = FormatField(rstMaster("CreateDate"))
            Case "MATDATE"
                .Text = FormatField(rstMaster("maturityDate"))
            Case "EFFECTDATE"
                .Text = FormatField(rstMaster("EffectiveDate"))
            Case "INTRATE"
                .Text = FormatField(rstMaster("RateOfInterest"))
            Case "ACCGROUP"
                gDbTrans.SqlStmt = "SELECT GroupName FROM AccountGroup WHERE " & _
                        "AccGroupID = " & FormatField(rstMaster("AccGroupId"))
                If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                .Text = FormatField(rst("GroupName"))
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
'Disable the Reset button (for auto acc no generation)
cmdAccept.Enabled = True
cmdDelete.Enabled = gCurrUser.IsAdmin

'Now Set the caption Delete Button
'If there is any loan on this deposit disbale the Delete button

Dim LoanID As Long
Dim Loan As Boolean
LoanID = FormatField(rstMaster("LoanId"))
 
gDbTrans.SqlStmt = "SELECT Top 1 Balance From DepositLoanTrans " & _
    " Where LoanID = " & LoanID & " ORDER BY TransID Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    Loan = FormatField(rst("Balance"))

Set rst = Nothing
Set rstMaster = Nothing
Set rstJoint = Nothing
Set rstDeposits = Nothing

'If loan is there on this deposit then dissble the delete Button
cmdDelete.Enabled = (Not Loan) And gCurrUser.IsAdmin
cmdDelete.Caption = GetResourceString(14)
   
'Display The last deposit details
Do
    IntroId = DepositIDNext
    If IntroId = 0 Then Exit Do
    m_AccID = IntroId
Loop

Call DepositIDDisplay

'Now Get the Total Balance OF the Customer at this account
gDbTrans.SqlStmt = "Select Sum(balance) from FDTrans A" & _
        " Where AccID In (Select Distinct AccID From FDMAster " & _
                " where AccNum = " & AddQuotes(m_AccNum, True) & _
                " AND DepositType = " & m_DepositType & ")" & _
        " And A.TransID = (Select MAx(TransID) From FDTrans B" & _
                " Where B.AccID = A.AccID )"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                txtBalance.Caption = FormatField(rst(0))
    
    With Me.rtfNote
        .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
        .Enabled = IIf(ClosedDate = "", True, False)
         Call m_Notes.LoadNotes(wis_Deposits + m_DepositType, AccId)
    End With
    cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
    Me.TabStrip1.Tabs(IIf(m_Notes.NoteCount, 1, 2)).Selected = True

RaiseEvent AccountChanged(m_AccID)
AccountLoad = True

End Function
Private Function AccountSave()
Dim AccId As Long
Dim JointNo As Integer
Dim count As Integer
Dim txtIndex As Integer
Dim rst As ADODB.Recordset
Dim CreationDate As String
Dim AccNum As String
Dim DepAmt As Currency

'Now Check For the duplicate entry of the Accids
If m_accUpdatemode = wis_INSERT Then
    txtIndex = GetIndex("AccID")
    With txtData(txtIndex)
        gDbTrans.SqlStmt = "SELECT * FROM FDMaster WHERE AccNum = " & _
            AddQuotes(.Text, True) & " AND DepositType = " & m_DepositType
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
            'MsgBox "This account already exists", vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(545), vbInformation, wis_MESSAGE_TITLE
            Set rst = Nothing
            Exit Function
        End If
        AccNum = .Text
    End With
    'Get the New Account ID for the new deposit
    gDbTrans.SqlStmt = "SELECT Max(AccID) FROM FDMaster"
    AccId = 1
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        AccId = FormatField(rst(0)) + 1
Else
    AccId = m_AccID
    DepAmt = Val(GetVal("DepAmount"))
    gDbTrans.SqlStmt = "SELECT Amount FROM FDTrans" & _
            " Where AccID= " & AccId & " AND TransID = 1"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then DepAmt = FormatField(rst(0))
End If

'Start Transactions to Data base
gDbTrans.BeginTrans
m_CustReg.ModuleID = wis_Deposits
If Not m_CustReg.SaveCustomer Then
    'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
    gDbTrans.RollBack
    Exit Function
End If
If Not m_frmFDJoint Is Nothing Then
    If Not m_frmFDJoint.SaveJointCustomers Then
        'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
End If

CreationDate = GetSysFormatDate(GetVal("CreateDate"))
Debug.Print GetVal("CertificateNo")
' Insert/update to database.
If m_accUpdatemode = wis_INSERT Then
    'Build the SQL insert statement.
    gDbTrans.SqlStmt = "Insert into FDMaster (AccID,AccNum,CustomerID," & _
        "CreateDate, EffectiveDate,MaturityDate,DepositAmount,MaturityAmount, " & _
        "CertificateNo,NomineeID,NomineeName,NomineeRelation, Introduced,RateOfInterest, " & _
        "LedgerNo, FolioNo,LastIntDate,DepositType,LastPrintId,AccGroupID,UserID) " & _
        " VALUES (" & _
        AccId & ", " & _
        AddQuotes(GetVal("AccID"), True) & "," & _
        m_CustReg.customerID & "," & _
        "#" & CreationDate & "#," & _
        "#" & GetSysFormatDate(GetVal("EffectDate")) & "#," & _
        "#" & GetSysFormatDate(GetVal("MatDate")) & "#," & _
        Val(GetVal("DepAmount")) & "," & _
        Val(GetVal("MatAmount")) & "," & _
        AddQuotes(GetVal("CertificateNo"), True) & "," & _
        Val(GetVal("NomineeID")) & "," & _
        AddQuotes(GetVal("NomineeName"), True) & "," & _
        AddQuotes(GetVal("NomineeRelation"), True) & "," & _
        Val(GetVal("IntroducerID")) & ", 1," & _
        AddQuotes(GetVal("LedgerNo"), True) & "," & _
        AddQuotes(GetVal("FolioNo"), True) & "," & _
        "#" & CreationDate & "#," & _
        m_DepositType & "," & _
        " 1, " & GetAccGroupID & "," & gUserID & " )"
    'Insert/update the data.
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If

    ' Insert the data to the joint table
    JointNo = 0
    If Not m_frmFDJoint Is Nothing Then
    For count = 0 To m_frmFDJoint.txtName.count - 1
        JointNo = JointNo + 1
        m_JointCustID(count) = m_frmFDJoint.JointCustId(count)
        If m_JointCustID(count) = 0 Then Exit For
        gDbTrans.SqlStmt = "Insert into FDJoint (AccId,CustomerID, CustomerNum) " _
            & "values (" & AccId & "," & _
            m_JointCustID(count) & "," & _
            JointNo & _
            ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
    Next
    End If
    gDbTrans.CommitTrans

ElseIf m_accUpdatemode = wis_UPDATE Then
    ' The user has selected updation.
    ' Confirm before proceeding with updation.
    ' Build the SQL update statement.
    ' SHASHI As the Create Date is TransDate/Create
    ' Date Of First Depostit. All this Update
    ' has to take effect on all depoists
    ' So Conidtion is only ACCID & notalso deposit Id
    gDbTrans.SqlStmt = "Update FDMaster set " & _
        " NomineeID = " & Val(GetVal("NomineeID")) & ", " & _
        " NomineeName = " & AddQuotes(GetVal("NomineeName"), True) & ", " & _
        " NomineeRelation = " & AddQuotes(GetVal("NomineeRelation"), True) & ", " & _
        " Introduced = " & Val(GetVal("IntroducerID")) & "," & _
        " LedgerNo = " & AddQuotes(GetVal("LedgerNo"), True) & "," & _
        " CreateDate = #" & CreationDate & "#," & _
        " MaturityDate = #" & GetSysFormatDate(GetVal("MatDate")) & "#," & _
        " EffectiveDate = #" & GetSysFormatDate(GetVal("EffectDate")) & "#," & _
        " CertificateNo = " & AddQuotes(GetVal("CertificateNo"), True) & "," & _
        " DepositAmount = " & DepAmt & "," & _
        " MaturityAmount = " & Val(GetVal("MatAmount")) & "," & _
        " RateOfINterest = " & Val(GetVal("IntRate")) & "," & _
        " FolioNo = " & AddQuotes(GetVal("FolioNo"), True) & ", " & _
        " AccGroupID = " & GetAccGroupID & _
        " where AccID = " & m_AccID
    'Insert/update the data.
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    JointNo = 0
    'm_JointCustID(0) = m_CustReg.CustomerId
    If Not m_frmFDJoint Is Nothing Then
      For count = 0 To 3
        JointNo = JointNo + 1
        If m_JointCustID(count) <> m_frmFDJoint.JointCustId(count) Then
            'If MsgBox("You have changed the joint account holder" & _
                vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, _
                wis_MESSAGE_TITLE) = vbNo Then Exit Function
            If MsgBox(GetResourceString(675) & vbCrLf & _
                GetResourceString(541), vbQuestion + vbYesNo, _
                wis_MESSAGE_TITLE) = vbNo Then Exit Function
            If m_frmFDJoint.JointCustId(count) = 0 Then 'Delte the Joint account details
                gDbTrans.SqlStmt = "DELETE * FROM FDJoint WHERE AccID = " & AccId & _
                    " AND CustomerNum >= " & JointNo
                'gDBTrans.SQLStmt = SqlStr
                If Not gDbTrans.SQLExecute Then
                    gDbTrans.RollBack
                    Exit Function
                End If
                m_JointCustID(count) = 0
            End If
        End If
        If m_frmFDJoint.JointCustId(count) = 0 Then Exit For
        If m_JointCustID(count) = 0 Then 'Insert the new record
            gDbTrans.SqlStmt = "Insert into FDJoint (AccID,CustomerID, CustomerNum) " _
                    & "values (" & AccId & "," & _
                    m_frmFDJoint.JointCustId(count) & "," & _
                    JointNo & _
                    ")"
        Else 'Update the existing record
            gDbTrans.SqlStmt = "UPDATE FDJoint Set CustomerID = " & m_frmFDJoint.JointCustId(count) & _
                " WHERE Accid = " & AccId & " AND CustomerNum = " & JointNo
        End If
        'gDBTrans.SQLStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
      Next
    End If
    gDbTrans.CommitTrans
End If

'MsgBox "Saved the account details.", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE

Set rst = Nothing
AccountSave = True

End Function

Private Sub DateChange(Ctrl As Control)

Dim boolEffDate As Boolean
Dim boolMatDate As Boolean
Dim boolDay As Boolean
Dim fromDate As Date
Dim toDate As Date

Static Insub As Boolean

If Insub Then Exit Sub
Insub = True

On Error GoTo ErrLine

If DateValidate(txtEffective.Text, "/", True) Then
    boolEffDate = True
    fromDate = GetSysFormatDate(txtEffective)
End If

If DateValidate(txtMatureDate.Text, "/", True) Then
    boolMatDate = True
    toDate = GetSysFormatDate(txtMatureDate)
End If

If Not boolEffDate And Not boolMatDate Then GoTo ErrLine

If Ctrl.name = txtDays.name Then
    If boolEffDate Then
        If Val(txtDays) > 0 And Val(txtDays) < 99999 Then
            toDate = DateAdd("d", Val(txtDays), fromDate)
            txtMatureDate = GetIndianDate(toDate)
            GoTo ExeLine
        End If
    End If
    If boolMatDate Then
        fromDate = DateAdd("d", Val(txtDays) * -1, toDate)
        txtEffective = GetIndianDate(fromDate)
        GoTo ExeLine
    End If
    GoTo ErrLine
End If

If Ctrl.name = txtEffective.name Then
    If Not boolEffDate Then GoTo ErrLine
    If boolMatDate And boolEffDate Then
        txtDays.Text = DateDiff("D", fromDate, toDate)
        GoTo ExeLine
    End If
    If Val(txtDays.Text) > 0 And Val(txtDays) < 99999 Then
        toDate = DateAdd("D", Val(txtDays), fromDate)
        txtMatureDate = GetIndianDate(toDate)
        GoTo ExeLine
    End If
    GoTo ErrLine
End If

If Ctrl.name = txtMatureDate.name Then
    If Not boolMatDate Then GoTo ErrLine
    If boolEffDate Then
        txtDays.Text = DateDiff("D", fromDate, toDate)
        GoTo ExeLine
    End If
    If Val(txtDays) > 0 And Val(txtDays) < 99999 Then
        fromDate = DateAdd("d", Val(txtDays) * -1, toDate)
        txtEffective = GetIndianDate(fromDate)
        GoTo ExeLine
    End If
    GoTo ErrLine
End If

ExeLine:

txtInterest.Text = GetDepositInterestRate(m_DepositType, fromDate, toDate)
'txtMatureAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + _
          ComputeFDInterest(Val(txtDepositAmount.Text), txtEffective.Text, txtMatureDate.Text, m_DepositType, CSng(Val(txtInterest.Text))))
txtMatureAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + _
          (Val(txtDepositAmount) * CDbl(DateDiff("D", fromDate, toDate)) * (Val(txtInterest) / 100)))

ErrLine:
    Insub = False

End Sub

Private Function DepositDelete() As Boolean
Dim rst As ADODB.Recordset
Dim TransDate As Date

'Prelim Checks
If m_AccID <= 0 Then
    'MsgBox "Deposit not loaded !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(523), vbExclamation, gAppName & " - Error"
    Exit Function
End If

'You can delete a deposit only if you do not have transactions more than one
Dim strRemark As String
Dim Amount As Currency
Dim transType As wisTransactionTypes
    
    gDbTrans.SqlStmt = "Select * from FDTrans where " & _
            " AccID = " & m_AccID & " AND Amount <> 0 " & _
            " Order by TransID desc "
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 1 Then
        'MsgBox "You cannot delete a deposit with transactions." & _
            "You cannot delete a deposit with transactions.", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(530) & vbCrLf & _
            GetResourceString(553), vbExclamation, gAppName & " - Error"
        Set rst = Nothing
        Exit Function
    End If
    strRemark = FormatField(rst("Particulars"))
    Amount = rst("Amount")
    transType = rst("Transtype")
    TransDate = rst("TransDate")
'Now Check transactions In INtrest account
    gDbTrans.SqlStmt = "Select * from FDIntTrans where " & _
            " AccID = " & m_AccID & " AND Amount <> 0 " & _
            " Order by TransID desc "
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Set rst = Nothing
        Exit Function
    End If
'Now Check transactions In IntrestPayable account
    gDbTrans.SqlStmt = "Select * from FDIntPayable where " & _
            " AccID = " & m_AccID & " AND Amount <> 0 " & _
            " Order by TransID desc "
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Set rst = Nothing
        Exit Function
    End If
    
'Beging DB Transactions
gDbTrans.BeginTrans
    
    If transType = wContraDeposit Or transType = wContraWithdraw Then
        gDbTrans.SqlStmt = "Select * from ContraTrans " & _
            " where AccID = " & m_AccID & " And AccHeadID = " & m_AccHeadId
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            Dim ContraClass As clsContra
            Set ContraClass = New clsContra
            If Not ContraClass.UndoTransaction(rst("ContraID"), TransDate) Then
                'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
                gDbTrans.RollBack
                Set ContraClass = New clsContra
                Exit Function
            End If
            Set ContraClass = New clsContra
        End If
    End If

    If InStr(1, strRemark, "From MatFD of ") Then
        If MsgBox("This deposit has transferred from the Matured Fd " & vbCrLf & _
            " This will undo the Matured Fd also" & " Do you want to continue?", _
            vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
       
        Dim MatAcc As Long
        MatAcc = Val(Mid(strRemark, Len("From MatFD of ")))
        gDbTrans.SqlStmt = "Delete from MatFDTrans where AccID = " & MatAcc & _
            " And TransID = (Select Max(TransID) FROM MatFDTrans " & _
                " where AccID = " & MatAcc & ")"
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
            gDbTrans.RollBack
            Exit Function
        End If
    End If
        
'Fire SQL Stmt
    
    'First Delete the Deposited Amount On from FDTrans
    gDbTrans.SqlStmt = "Delete from FDTrans where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
'Then Delete the Amount On from FD Interest
    gDbTrans.SqlStmt = "Delete from FDIntTrans where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If

'Then Delete the Deposited Amount On from Payable
    gDbTrans.SqlStmt = "Delete from FDIntPayable where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If

'Then Delete the joint account information
    'Check whether this accoutn is having any othe Deposit
    
    gDbTrans.SqlStmt = "Delete FROM FDJoint WHERE AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If

'Finally Delate the Related record in FDMaste
    gDbTrans.SqlStmt = "Delete from FDMaster where AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

'Now undo the amount from the Ledger headd
If transType = wDeposit Then _
    If Not bankClass.UndoCashDeposits(m_AccHeadId, Amount, TransDate) Then GoTo ExitLine

Set bankClass = Nothing
'Close DB Operations
    gDbTrans.CommitTrans

DepositDelete = True
Exit Function

ExitLine:
    'MsgBox "Unable to delete account !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(532), vbExclamation, gAppName & " - Error"
    gDbTrans.RollBack

End Function

Private Function DepositExists(Optional IndianClosedDate As String) As Boolean
Dim rst As ADODB.Recordset

'Prelim checks
    If m_AccID <= 0 Then
        'MsgBox "Account not loaded !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(523), vbExclamation, gAppName & " - Error"
        Exit Function
    End If

'Check DB the status
    gDbTrans.SqlStmt = "Select * from FDMaster where " & _
        " AccID = " & m_AccID
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        DepositExists = True
        If Not IsMissing(IndianClosedDate) Then
            IndianClosedDate = FormatField(rst("ClosedDate"))
        End If
        Set rst = Nothing
    End If

End Function

Private Function DepositIDNext() As Integer
Dim rst As ADODB.Recordset
'Given the current deposittype this function will return the next deposit ID

'Prelim checks
'    If m_CustId <= 0 Then GoTo Errline
    
'Set SQL
    gDbTrans.SqlStmt = "Select TOP 1 * from FDMAster where AccID > " & m_AccID & _
        " AND AccNum = " & AddQuotes(m_AccNum, True) & _
        " AND DepositType = " & m_DepositType & " ORDER BY ACCId Asc"
        
    DepositIDNext = 0
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then DepositIDNext = rst("AccID")
    Set rst = Nothing

Exit Function
ErrLine:
    DepositIDNext = 0

End Function

Private Function DepositIDPrevious() As Integer
'Given the current depositID this function will return the next deposit ID
Dim rst As ADODB.Recordset
'Prelim checks
'    If m_CustId <= 0 Then GoTo Errline
    
'Set SQL
    gDbTrans.SqlStmt = "Select TOP 1 * from FDMAster where AccID < " & m_AccID & _
        " AND AccNum = " & AddQuotes(m_AccNum, True) & _
        " AND DepositType = " & m_DepositType & " ORDER BY ACCId Desc"
    
    DepositIDPrevious = 0
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        DepositIDPrevious = FormatField(rst("AccID"))
    Set rst = Nothing
    
Exit Function
ErrLine:
    DepositIDPrevious = 0
    Exit Function
End Function

Private Function DepositReopen() As Boolean
Dim TransID As Long
Dim ClosedDate As String
Dim transType As wisTransactionTypes
Dim rst As ADODB.Recordset
Dim headName As String

Dim AsOnDate As Date

On Error GoTo ErrLine
    'Check whether deposit has transferred to the Matured FD
    gDbTrans.SqlStmt = "Select * FROM FDMaster where " & _
                    "AccID = " & m_AccID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine
    If FormatField(rst("MaturedOn")) <> "" Then
        'The closed account has transferred to the Matured FD
        'So check whether that account is still in the MFD
        'or it has closedFrom MFD also
        m_Transferred = True
        gDbTrans.SqlStmt = "SELECT Top 1 Balance From MatFDTrans " & _
                    " WHERE AccID = " & m_AccID & " ORDER BY TransID Desc"
        Call gDbTrans.Fetch(rst, adOpenForwardOnly) '< 1 'Then GoTo ErrLine
        If FormatField(rst("Balance")) = 0 Then
            'the account has been Closed from the MAtured Fd Account
            'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
            Exit Function
        End If
    End If
    
    'Now get the Last Transction Id Of this Deposit
    TransID = GetFDMaxTransID(m_AccID)

    'Get the TransDate and Transid on which this deposit was closed
    gDbTrans.SqlStmt = "Select TransType,TransID,TransDate" & _
                        " From FDTrans where AccID = " & m_AccID & _
                        " And TransId = " & TransID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
        Exit Function
    Else
        TransID = FormatField(rst("TransID"))
        ClosedDate = FormatField(rst("TransDate"))
        AsOnDate = rst("TransDate")
    End If
    transType = FormatField(rst("TransType"))
    'if last transaction is not with drwan then do not reopen it
    If transType <> wWithdraw And transType <> wContraWithdraw Then GoTo ErrLine
    
    If transType = wContraWithdraw Then
        Dim ContraClass As clsContra
        
        gDbTrans.SqlStmt = "Select * From ContraTrans Where " & _
                " AccheadId = " & m_AccHeadId & _
                " And AccID = " & m_AccID & _
                " And TransType = " & wContraWithdraw
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            Set ContraClass = New clsContra
            If ContraClass.UndoTransaction(rst("ContraId"), GetSysFormatDate(ClosedDate)) <> Success Then GoTo ErrLine
            Set ContraClass = Nothing
        End If
    End If

'Start DB Operations

Dim DepAmount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim MiscAmount As Currency

'Now Get the Amount removed from the Ledger Heads
'Deposit Amount
gDbTrans.SqlStmt = "Select * from FDTrans where " & _
        " AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then DepAmount = FormatField(rst("Amount"))
'Now Interest Amount
gDbTrans.SqlStmt = "Select * From FDIntTrans where " & _
        " AccID = " & m_AccID & " And Transid = " & TransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    transType = FormatField(rst("TransType"))
    IntAmount = FormatField(rst("Amount"))
    If transType = wContraDeposit Or transType = wDeposit Then IntAmount = IntAmount * -1
End If

gDbTrans.SqlStmt = "Select Amount from FDIntPayable where " & _
        " AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then PayableAmount = FormatField(rst("Amount"))

Dim PayableHeadID As Long
Dim IntHeadID As Long

'Interest HeadID
headName = GetDepositTypeText(m_DepositType) & " " & GetResourceString(487)
IntHeadID = GetIndexHeadID(headName)
'Payable Headid
headName = GetDepositTypeText(m_DepositType) & " " & _
        GetResourceString(375) & " " & GetResourceString(47)
PayableHeadID = GetIndexHeadID(headName)

gDbTrans.BeginTrans
    
    'Delete entry where you would have withdrawn
    'interest from Interest payble acccont
    gDbTrans.SqlStmt = "Delete from FDIntPayable where " & _
        " AccID = " & m_AccID & " And Transid = " & TransID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Error in Undo Operation !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(561), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
     
     'Delete next entry where you would have given
    'interest to the deposit
    gDbTrans.SqlStmt = "Delete from FDIntTrans where " & _
        " AccID = " & m_AccID & " And TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Error in Undo Operation !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(561), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
        
    'Delete the last entry where you would have returned his deposit amount
    gDbTrans.SqlStmt = "Delete from FDTrans where " & _
        " AccID = " & m_AccID & " And Transid >= " & TransID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Error in Undo Operation !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(561), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
    
    If m_Transferred Then
        'Then delete the information of the transferrd account
        gDbTrans.SqlStmt = "DELETE * FROM MatFDTRans Where " & _
            " AccID = " & m_AccID
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Error in Undo Operation !", vbCritical, gAppName & " - Error"
            MsgBox GetResourceString(561), vbCritical, gAppName & " - Error"
            gDbTrans.RollBack
            Exit Function
        End If
    End If

'Update the Record where TransID = 1
    gDbTrans.SqlStmt = "Update FDMaster set ClosedDate = NULL, " & _
            " MaturedON = NULL WHERE AccID = " & m_AccID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Error in Undo Operation !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(561), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If

'Now Undo the Transation From the ledger Heads
Dim bankClass As New clsBankAcc
If Not bankClass.UndoCashWithdrawls(m_AccHeadId, DepAmount, AsOnDate) Then GoTo ErrLine
If IntAmount > 0 Then If Not bankClass.UndoCashWithdrawls(IntHeadID, IntAmount, AsOnDate) Then GoTo ErrLine
If IntAmount < 0 Then If Not bankClass.UndoCashDeposits(IntHeadID, Abs(IntAmount), AsOnDate) Then GoTo ErrLine
If PayableAmount > 0 Then If Not bankClass.UndoCashWithdrawls(IntHeadID, PayableAmount, AsOnDate) Then GoTo ErrLine

'Close DB Operations
gDbTrans.CommitTrans

Set bankClass = Nothing
Set rst = Nothing
DepositReopen = True

Exit Function

ErrLine:
    gDbTrans.RollBack
    If Err Then
        MsgBox "Error in deposit reopen", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
'   Resume
    'MsgBox "Error in undo operation !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(561), vbExclamation, gAppName & " - Error"
        

End Function

Public Property Get Nominee() As String
' The Nominee string consists of
' Nominee_name;Nominee_age;Nominee_Relation.

Nominee = GetVal("Nomineename") & ";" _
        & GetVal("NomineeAge") & ";" _
        & GetVal("NomineeRelation")

End Property

Private Sub ResetUserInterface()

If m_AccID = 0 And m_CustReg.customerID = 0 Then Exit Sub


'Disable controls on tab 1
    txtBalance = ""
    With cmbNames
        .BackColor = wisGray: .Enabled = False: .Clear
    End With
    With txtDepositDate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With cmdDepositDate
        .Enabled = False
    End With
    With txtEffective
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With txtCertificate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With txtMatureDate
        .BackColor = wisGray: .Enabled = False
    End With
    With cmdMatureDate
        .Enabled = False
    End With
    With txtDays
        .BackColor = wisGray: .Enabled = False
    End With
    With txtInterest
        .BackColor = wisGray: .Enabled = False
    End With
    With cmdEffective
        .Enabled = False
    End With
    With txtCertificate
        .BackColor = wisGray: .Enabled = False
    End With
    With txtDepositAmount
        .BackColor = wisGray: .Enabled = False
    End With
    With txtMatureAmount
        .BackColor = wisGray: .Enabled = False
    End With
    With txtRTF
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    
    cmdAccept.Enabled = False
    cmdDelete.Enabled = False
    cmdClose.Enabled = False
    cmdInterest.Enabled = False
    
    grdLedger.Clear
'Now the TAB 2
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
    lblOperation.Caption = GetResourceString(54) '"Operation Mode : <INSERT>"
    txtIndex = GetIndex("AccID")
    txtData(txtIndex).Text = GetLastAccountNumber
    txtData(txtIndex).Locked = False
    
'The form level variables
    m_accUpdatemode = wis_INSERT
    Set m_CustReg = Nothing
    Set m_CustReg = New clsCustReg
    m_CustReg.ModuleID = wis_Deposits
    m_AccID = 0
    m_AccNum = ""
    RaiseEvent AccountChanged(0)
    
    On Error Resume Next
'    Unload m_frmFDJoint
    Set m_frmFDJoint = Nothing
    m_JointCustID(0) = 0
    m_JointCustID(1) = 0
    m_JointCustID(2) = 0
    m_JointCustID(3) = 0
    Err.Clear
End Sub

Private Function DepositIDDisplay() As Boolean

'Dim TransRst As Recordset
Dim MasterRst As ADODB.Recordset
Dim Rate As Double
Dim DepDt As Date, MatDt As Date, EffectDate As Date, CreateDate As Date
Dim MatAmt As Currency, LnAmt As Currency, DepAmt As Currency
Dim Days As Long, PADLEN As Integer
Dim ClosedDate As String
Dim transType As wisTransactionTypes
Dim IntAmount As Currency
Dim CertificateNo  As String * 19
Dim MaturedDate As String
Dim Balance As Currency
Dim rst As ADODB.Recordset
Dim LoanID As Long
Dim strRtf As String

If m_AccID <= 0 Then GoTo ExitLine

'Fill Tab1e i.e. Transaction Frame
'Table
'Clear All the Text Boxes
m_Transferred = False
txtDepositDate.Text = gStrDate
txtEffective.Text = gStrDate
txtMatureDate.Text = ""
txtInterest.Text = ""
txtDepositAmount = 0

'Build and Fire the SQL Query
gDbTrans.SqlStmt = "Select * from FDMaster where AccID = " & m_AccID '& _
        " AND CustomerID = " & m_CustReg.CustomerId

If gDbTrans.Fetch(MasterRst, adOpenForwardOnly) < 0 Then GoTo ExitLine
ClosedDate = FormatField(MasterRst("ClosedDate"))
m_Transferred = IIf(FormatField(MasterRst("MaturedOn")) <> "", True, False)
txtData(GetIndex("IntRate")) = MasterRst("RateOfInterest")

'Build and Fire the SQL Query
'set  the conditions
transType = wDeposit
  gDbTrans.SqlStmt = "Select Top 1 Balance from FDTrans where AccID = " & _
        m_AccID & " Order by TransID desc "
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ExitLine
Balance = FormatField(rst("Balance"))

transType = wWithdraw
gDbTrans.SqlStmt = "Select SUM(Amount)" & _
      " From FDIntTrans where AccID = " & m_AccID & _
      " ANd (TransType = " & wWithdraw & " Or TransType = " & wContraWithdraw & ")"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                        IntAmount = FormatField(rst(0))
If ClosedDate <> "" Then Balance = MasterRst("DepositAmount")

'Assign to local variables
PADLEN = 30
Rate = Format(FormatField(MasterRst("RateOfInterest")), "#.00")
CreateDate = MasterRst("CreateDate")
EffectDate = MasterRst("EffectiveDate")
MatDt = MasterRst("MaturityDate")
DepDt = MasterRst("CreateDate")
Days = DateDiff("D", DepDt, MatDt)
ClosedDate = FormatField(MasterRst("ClosedDate"))
CertificateNo = FormatField(MasterRst("CertificateNo"))

DepAmt = FormatField(MasterRst("DepositAmount"))
If ClosedDate <> "" Then
    MatAmt = DepAmt + IntAmount
Else
    MatAmt = FormatField(MasterRst("MAturityAmount"))
    If MatAmt = 0 Then MatAmt = DepAmt + _
            ComputeFDInterest(DepAmt, DepDt, MatDt, CSng(Rate), False)
End If
    
'Get total loan Balance if any
LoanID = FormatField(MasterRst("LoanID"))

gDbTrans.SqlStmt = "Select Balance From DepositLoanTrans where " & _
    " LoanID = " & LoanID & " ORDER BY TransID Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then LnAmt = Val(FormatField(rst(0)))

'Fill the text box with the data collected
' If Language is english Then Set Font as Arial

    'If Not gLangOffSet Then txtRTF.Font.Name = "Courier"
    txtRTF.Font.name = IIf(gLangOffSet = 0, "Courier", gFontName)

    strRtf = RPad(" Certificate No :", " ", PADLEN) & CertificateNo & vbCrLf
    strRtf = strRtf & RPad(" Deposit Date :", " ", PADLEN) & GetIndianDate(DepDt) & vbCrLf
    If CreateDate <> EffectDate Then
        strRtf = strRtf & RPad(" Effective Date :", " ", PADLEN) & _
                    GetIndianDate(EffectDate) & vbCrLf
    End If
    strRtf = strRtf & RPad(" Deposit Amount :", " ", PADLEN) & "Rs." & _
        FormatCurrency(DepAmt) & vbCrLf
    If DepAmt <> Balance Then _
        strRtf = strRtf & RPad(" Deposit Balance:", " ", PADLEN) & _
        "Rs." & FormatCurrency(Balance) & vbCrLf
    strRtf = strRtf & RPad(" Maturity Date :", " ", PADLEN) & GetIndianDate(MatDt) & vbCrLf
    strRtf = strRtf & RPad(" Maturiy Amount :", " ", PADLEN) & "Rs." & _
        FormatCurrency(MatAmt) & vbCrLf
    strRtf = strRtf & RPad(" Rate of Interest :", " ", PADLEN) & Rate & " %" & vbCrLf   '& _
    '
    
    'Reset UI parts
    MaturedDate = FormatField(MasterRst("MaturedOn"))
    If ClosedDate <> "" Then
        cmdDeleteAcc.Enabled = False
        cmdDelete.Caption = GetResourceString(313) '"&Reopen"
        cmdClose.Enabled = True
        If FormatField(MasterRst("MaturedOn")) <> "" Then
            cmdClose.Enabled = False
            cmdClose.Caption = GetResourceString(266) '"&Renew"
            gDbTrans.SqlStmt = "SELECT Top 1 Balance From MatFDTrans WHERE " & _
                " accID = " & m_AccID & " ORDER BY TransID Desc"
            If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                cmdClose.Enabled = FormatField(rst("Balance"))
        End If
        cmdDelete.Enabled = gCurrUser.IsAdmin
        cmdInterest.Enabled = False
        strRtf = strRtf & RPad(" Interest Paid :", " ", PADLEN) & "Rs." & FormatCurrency(IntAmount) & vbCrLf
        If Not m_Transferred Then
            cmdClose.Enabled = False
            strRtf = strRtf & RPad(" Closed On :", " ", PADLEN) & ClosedDate & vbCrLf
        Else
            strRtf = strRtf & RPad(" Tranferred to MFD :", " ", PADLEN) & ClosedDate & vbCrLf
        End If
    Else
        txtRTF.Enabled = True
        cmdClose.Caption = GetResourceString(11) '"&Close"
        cmdDelete.Caption = GetResourceString(14)  'DELETE
        gDbTrans.SqlStmt = "SELECT MAX(TransDate) From FDIntTrans WHERE " & _
            " AccID = " & m_AccID
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 And FormatField(rst(0)) <> "" Then _
            cmdDelete.Caption = GetResourceString(5)  'UnDo Last
    
        cmdClose.Enabled = True
        cmdDelete.Enabled = gCurrUser.IsAdmin
        cmdInterest.Enabled = True
        If IntAmount <> 0 Then
            strRtf = strRtf & RPad(" Interest Paid :", " ", PADLEN) & _
                "Rs." & FormatCurrency(IntAmount) & vbCrLf
        End If
        cmdInterest.Enabled = IIf(MaturedDate = "", True, False)
        'If m_Cumulative Then cmdInterest.Enabled = False
        
        strRtf = strRtf & RPad(" Loans drawn :", " ", PADLEN) & "Rs." & FormatCurrency(LnAmt)
    End If
''''///KANNADA
If gLangOffSet > 0 Then
    strRtf = RPad(" " + GetResourceString(337) + " " + GetResourceString(60) + " :", " ", PADLEN) & CertificateNo & vbCrLf
    strRtf = strRtf & RPad(" " & GetResourceString(38, 37) & " :", " ", PADLEN) & GetIndianDate(DepDt) & vbCrLf
    If CreateDate <> EffectDate Then
        strRtf = strRtf & RPad(" " & GetResourceString(43, 37) & " :", " ", PADLEN) & _
                    GetIndianDate(EffectDate) & vbCrLf
    End If
    strRtf = strRtf & RPad(" " & GetResourceString(43, 40) & " :", " ", PADLEN) & "Rs." & _
        FormatCurrency(DepAmt) & vbCrLf
    If DepAmt <> Balance Then _
        strRtf = strRtf & RPad(" " & GetResourceString(43, 67) & " :", " ", PADLEN) & _
        "Rs." & FormatCurrency(Balance) & vbCrLf
    strRtf = strRtf & RPad(" " & GetResourceString(48, 37) & " :", " ", PADLEN) & GetIndianDate(MatDt) & vbCrLf
    strRtf = strRtf & RPad(" " + GetResourceString(48) + GetResourceString(40) + " :", " ", PADLEN) & "Rs." & _
        FormatCurrency(MatAmt) & vbCrLf
    strRtf = strRtf & RPad(" " & GetResourceString(186) & " :", " ", PADLEN) & Rate & " %" & vbCrLf    '& _
    '
    If ClosedDate <> "" Then
        strRtf = strRtf & RPad(" " & GetResourceString(487) & " :", " ", PADLEN) & "Rs." & FormatCurrency(IntAmount) & vbCrLf
        If Not m_Transferred Then
            strRtf = strRtf & RPad(" " & GetResourceString(282) & " :", " ", PADLEN) & ClosedDate & vbCrLf
        Else
            strRtf = strRtf & RPad(" " & GetResourceString(307, 37) & " :", " ", PADLEN) & ClosedDate & vbCrLf
        End If
    Else
        If IntAmount <> 0 Then
            strRtf = strRtf & RPad(" " & GetResourceString(487) & " :", " ", PADLEN) & _
                "Rs." & FormatCurrency(IntAmount) & vbCrLf
        End If
        cmdInterest.Enabled = IIf(MaturedDate = "", True, False)
        'If m_Cumulative Then cmdInterest.Enabled = False
        
        strRtf = strRtf & RPad(" " & GetResourceString(43, 58) & " :", " ", PADLEN) & "Rs." & FormatCurrency(LnAmt)
    End If
End If
''''///KANNADA

txtRTF.Text = strRtf

txtRTF.BackColor = IIf(ClosedDate = "", vbWhite, wisGray)

'Check if next deposit is present
cmdNext.Enabled = DepositIDNext()
cmdPrevious.Enabled = DepositIDPrevious()

fraAccStmt.ZOrder 0
If ClosedDate <> "" Then Balance = 0
'txtBalance.Caption = FormatCurrency(Balance)
DepositIDDisplay = True

'Now Diplay the Detial of transaction of this account
Call ShowLedger

Exit Function

ExitLine:
    txtRTF.Text = "": txtRTF.BackColor = vbWhite
    cmdPrevious.Enabled = False: cmdNext.Enabled = False
    cmdClose.Enabled = False
    cmdDelete.Enabled = gCurrUser.IsAdmin: cmdInterest.Enabled = False
    Set MasterRst = Nothing
    Set rst = Nothing
    
Exit Function

DisplayError:
    'MsgBox "Error displaying deposit details"
    MsgBox GetResourceString(531)
    
End Function

'****************************************************
'Returns a new account number
'Author: Girish
'Date : 29th Dec, 1999
'****************************************************
Private Function GetLastAccountNumber() As String
    Dim NewAccNo As String
    Dim rst As ADODB.Recordset
    NewAccNo = 0
    gDbTrans.SqlStmt = "Select AccNum from FDMaster " & _
        " WHERE DepositType = " & m_DepositType & _
        " And AccId=(Select Max(AccID) from FDMaster " & _
        " WHERE DepositType = " & m_DepositType & ")" '& _

    If gDbTrans.Fetch(rst, adOpenForwardOnly) = 1 Then _
        NewAccNo = FormatField(rst(0))
        
    If IsNumeric(NewAccNo) Then NewAccNo = Val(NewAccNo) + 1
    GetLastAccountNumber = NewAccNo
End Function

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

Private Function ValidateSave() As Boolean
Dim txtIndex As Integer

'Check for required details
'Check if account number was specified
txtIndex = GetIndex("AccID")
With txtData(txtIndex)
    If Trim$(.Text) = "" Then
        'MsgBox "No Account number specified!", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(523), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If
End With

'Check for account holder name.
txtIndex = GetIndex("AccName")
With txtData(txtIndex)
    If Trim$(.Text) = "" Then
        'MsgBox "Account holder name not specified!", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(529), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If
End With

' Check for nominee name.
Dim NomineeSpecified As Boolean
txtIndex = GetIndex("NomineeName")
NomineeSpecified = True
With txtData(txtIndex)
    If Trim$(.Text) = "" Then
        'MsgBox "Nominee name not specified!", vbExclamation, wis_MESSAGE_TITLE
        If MsgBox(GetResourceString(558) & vbCrLf & GetResourceString(541), vbInformation + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If
        NomineeSpecified = False
    End If
End With

' Check for nominee relationship.
txtIndex = GetIndex("NomineeRelation")
With txtData(txtIndex)
    If NomineeSpecified And Trim$(.Text) = "" Then
        'MsgBox "Specify nominee relationship.", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(559), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If
End With

Dim rst As ADODB.Recordset
txtIndex = GetIndex("IntroducerID")
With txtData(txtIndex)
    ' Check if an introducerID has been specified.
    If Trim$(.Text) = "" Then
        'If MsgBox("No introducer has been specified!" & vbCrLf &
            '"Add this Account anyway?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox(GetResourceString(560) & vbCrLf & _
                GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If
    Else
        ' Check if the introducer exists.
        gDbTrans.SqlStmt = "SELECT * FROM NameTab WHERE CustomerID = " & Val(.Text)
        If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
            'MsgBox "Invalid Introducer specified !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(508), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If
        Set rst = Nothing
    End If
End With

' Validate Date
txtIndex = GetIndex("CreateDate")
With txtData(txtIndex)
    If Not DateValidate(.Text, "/", True) Then
        MsgBox GetResourceString(501), _
                    vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If
End With
'Check for the Account Group
If GetAccGroupID = 0 Then
    'MsgBox "You have not selected the Account Group", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(749), vbInformation, wis_MESSAGE_TITLE
    txtIndex = GetIndex("AccGroup")
    ActivateTextBox txtData(txtIndex)
    Exit Function
End If


ValidateSave = True

End Function

Private Function ValidateTrans() As Boolean

'Validate the date and assign to variable
If Not DateValidate(Trim$(txtDepositDate.Text), "/", True) Then
    'MsgBox "Please specify deposit date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDepositDate
    Exit Function
End If

If Not DateValidate(Trim$(txtEffective.Text), "/", True) Then
    'MsgBox "Please specify deposit date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtEffective
    Exit Function
End If

'Validate the Maturity date
If Not DateValidate(Trim$(txtMatureDate.Text), "/", True) Then
    'MsgBox "Please specify maturity date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtMatureDate
    Exit Function
End If

If WisDateDiff(txtEffective, txtMatureDate) <= 2 Then
    'MsgBox "Invalid date specified!", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtMatureDate
    Exit Function
End If

If WisDateDiff(txtDepositDate, txtMatureDate) <= 2 Then
    'MsgBox "Invalid date specified!", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtMatureDate
    Exit Function
End If

'Validate the rate of interest
If Val(Trim$(txtInterest.Text)) <= 0 Then
    'MsgBox "Invalid rate of interest specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(505), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtInterest
    Exit Function
End If

'Validate the deposit Amount
If Not CurrencyValidate(txtDepositAmount.Text, False) Then
    'MsgBox "Invalid deposit amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDepositAmount
    Exit Function
End If

'Validate the maturity Amount
If Not CurrencyValidate(txtMatureAmount.Text, False) Then
    'MsgBox "Invalid maturity amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtMatureAmount
    Exit Function
End If

If Trim(txtCertificate.Text) = "" Then
    'If MsgBox("Certificate no not speicfied " & vbCrLf & _
        " Do yoy want to continue", vbInformation, "wis_message_title") = vbNo Then
    If MsgBox(GetResourceString(337, 60, 296) & _
            GetResourceString(541), vbInformation + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
        ActivateTextBox txtCertificate
        Exit Function
    End If

ElseIf Val(txtCertificate.Text) = 0 Then
    'If MsgBox("Invalid Certificate no speicfied " & vbCrLf & _
        " Do yoy want to continue", vbInformation, "wis_message_title") = vbNo Then
    If MsgBox(GetResourceString(337, 60, 296) & _
            GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
        ActivateTextBox txtCertificate
        Exit Function
    End If
End If

ValidateTrans = True

End Function

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub
'
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
'
Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" Then _
    txtData(Val(txtIndex)).Text = cmb(Index).Text

End Sub


Private Sub cmbFrom_Click()
Call EnableInterestApply
End Sub

Private Sub cmbTo_Click()
Call EnableInterestApply
End Sub

Private Sub cmd_Click(Index As Integer)

Dim txtIndex As String
'Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
Select Case UCase$(strField)
    Case "ACCID"
        If m_accUpdatemode = wis_INSERT Then _
            txtData(txtIndex).Text = GetLastAccountNumber

    Case "ACCNAME"
    
        m_CustReg.ShowDialog
        If m_CustReg.customerID <> 0 Then _
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
    
    Case "INTRODUCERNAME", "NOMINEENAME"
        
        ' Build a query for getting introducer details.
        ' If an account number specified, exclude it from the list.
        Dim strSearch As String
        Dim rst As ADODB.Recordset
        strSearch = InputBox("Enter some letters of the customer whomU are serching", wis_MESSAGE_TITLE)
        
        gDbTrans.SqlStmt = "SELECT CustomerID as [Cust No], " _
                    & "Title + FirstName + Space(1) + Middlename " _
                    & "+ space(1) + LastName as CustName " _
                    & " FROM NameTab "
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " WHERE " & _
                 " CustomerID <> " & m_CustReg.customerID
        strSearch = Trim(strSearch)
        If strSearch <> "" Then
            gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND (" & _
                " FirstName like '" & strSearch & "%' OR " & _
                " MiddleName like '" & strSearch & "%' OR " & _
                " LastName like '" & strSearch & "%' )"
        End If
        
        Dim Lret As Long
        Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
        If Lret <= 0 Then
            'MsgBox "No accounts present !", vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(525), vbExclamation, wis_MESSAGE_TITLE
            Exit Sub
        End If
        'Fill the details to report dialog and display it.
        If m_frmLookUp Is Nothing Then _
            Set m_frmLookUp = New frmLookUp
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
            If .lvwReport.ColumnHeaders.count > 2 Then
            .lvwReport.ColumnHeaders(3).Width = 3750
            End If
            .Title = "Select Introducer..."
            .Show vbModal, Me
            txtData(txtIndex - 1).Text = .lvwReport.SelectedItem.Text
            txtData(txtIndex).Text = .lvwReport.SelectedItem.SubItems(1)
        End With
        Set rst = Nothing
    Case "JOINTHOLDER"
        If m_frmFDJoint Is Nothing Then Set m_frmFDJoint = New frmJoint
        m_frmFDJoint.ModuleID = wis_Deposits + m_DepositType
        With m_frmFDJoint
            .Left = Me.Left + picViewport.Left + _
                txtData(txtIndex).Left + fraNew.Left + CTL_MARGIN
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fraNew.Top + 300
            Screen.MousePointer = vbDefault
            .Show vbModal
            If .Status = "OK" Then txtData(txtIndex).Text = .JointHolders
        End With
End Select

End Sub

Private Sub cmdAccept_Click()

If Not ValidateTrans() Then Exit Sub

If Not AcceptTransaction() Then Exit Sub

txtMatureAmount = 0
If Not AccountLoad(m_AccID) Then Exit Sub

End Sub


Private Sub LoadDeposits()
Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "SELECT * FROM DepositName"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
'        cmbDeposits.Clear
        While Not rst.EOF
'            cmbDeposits.AddItem Rst("DepositName")
'            cmbDeposits.ItemData(cmbDeposits.NewIndex) = Rst("DepositID")
            rst.MoveNext
        Wend
        Set rst = Nothing
    End If
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

Private Sub cmdClear_Click()
' Clear the account holder details.
Call ResetUserInterface
End Sub
Private Sub cmdClose_Click()
Dim TransID As Long

Dim rst As ADODB.Recordset
'Prelim checks
 If m_AccID = 0 Then Exit Sub
 
'Check if already closed
 gDbTrans.SqlStmt = "Select * from FDMaster where " & _
     "AccID = " & m_AccID
                     
 If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

 If FormatField(rst("ClosedDate")) <> "" Then
     'MsgBox "Deposit already closed on " & FormatField(gDBTrans.Rst("ClosedDate")), vbExclamation, gAppName & " - Error"
     MsgBox GetResourceString(524) & FormatField(rst("ClosedDate")), vbExclamation, gAppName & " - Error"
     Exit Sub
 End If
 Set rst = Nothing

Dim Cancel As Integer
'If Account has transferred then
If m_Transferred Then
    RaiseEvent FDRenew(m_AccID, Cancel)
Else 'Close the account
    RaiseEvent FDClose(m_AccID, Cancel)
End If
    
'Display the depositdetails after the relevant operation
Call DepositIDDisplay

End Sub

Private Sub cmdDate1_Click()

With Calendar
    .selDate = gStrDate
    .Left = Me.Left + Me.fraReports.Left + Me.fraDateRange.Left + Me.ActiveControl.Left - .Width / 2
    .Top = Me.Top + Me.fraReports.Top + Me.fraDateRange.Top + Me.ActiveControl.Top
    .selDate = txtFromDate.Text
    .Show vbModal
    txtFromDate.Text = .selDate
End With

End Sub


Private Sub cmdDate2_Click()
With Calendar
    .Left = Me.Left + Me.fraReports.Left + Me.fraDateRange.Left + Me.ActiveControl.Left - .Width / 2
    .Top = Me.Top + Me.fraReports.Top + Me.fraDateRange.Top + Me.ActiveControl.Top
    .selDate = txtToDate.Text
    .Show vbModal
    txtToDate.Text = .selDate
End With

End Sub

Private Sub cmdDelete_Click()
'Check if deposit exists
    If Not DepositExists() Then
        'MsgBox "Deposit does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Call ResetUserInterface
        Exit Sub
    End If

    If cmdDelete.Caption = GetResourceString(5) Then
        'Delte the Last Update d transaction
        Call UndoLastTransaction
    
    ElseIf InStr(1, cmdDelete.Caption, GetResourceString(313), vbTextCompare) > 0 Then  'Do undo close
        'If account Has already Closed then reopen the account
        'If MsgBox("Are you sure you want to reopen this deposit ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
        If MsgBox(GetResourceString(538), _
            vbQuestion + vbYesNo + vbDefaultButton2, gAppName & " - Error") = vbNo Then _
            Exit Sub
        Call DepositReopen
    Else
        'Check for he has add any deposit
      'Confirm
        'If MsgBox("Are you sure you want to delete this account ?", vbYesNo + vbQuestion, gAppName & " - Confirmation") = vbNo Then
        If MsgBox(GetResourceString(539), _
            vbYesNo + vbQuestion, gAppName & " - Confirmation") = vbNo Then _
            Exit Sub
          'Delete the deposit
        If Not DepositDelete() Then Exit Sub
        m_AccID = DepositIDNext()
        'MsgBox "Deposit for this account deleted successfully !", vbInformation, gAppName & " - Message"
        MsgBox GetResourceString(546), vbInformation, gAppName & " - Message"
    End If
    
    'Refresh the display
    Call DepositIDDisplay

End Sub

Private Sub cmdDeleteAcc_Click()
    Call AccountDelete
End Sub

Private Sub cmdDepositDate_Click()
With Calendar
    .Left = Me.Left + Me.fraDeposits.Left + Me.ActiveControl.Left - .Width / 2
    .Top = Me.Top + Me.fraDeposits.Top + Me.ActiveControl.Top
    .selDate = txtDepositDate.Text
    .Show vbModal
    txtDepositDate.Text = .selDate
End With
End Sub

Private Sub cmdEffective_Click()
With Calendar
    .Left = Me.Left + Me.fraDeposits.Left + Me.ActiveControl.Left - .Width / 2
    .Top = Me.Top + Me.fraDeposits.Top + Me.ActiveControl.Top
    .selDate = txtEffective.Text
    .Show vbModal
    txtEffective.Text = .selDate
End With
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

FromIndex = cmbFrom.ListIndex
ToIndex = cmbTo.ListIndex
If ToIndex < FromIndex Then
    MsgBox "Select proper period for interest", vbOKOnly, wis_MESSAGE_TITLE
    Exit Sub
End If

Dim SetUp As New clsSetup
Dim strModule As String
Dim strValue As String
Dim strDef As String

strModule = "DEPOSIT" & m_DepositType
strDef = IIf(optDays, "DAYS", "YEAR")

'First check whether he has enter the previous slab interest rates or not
'if he has not entered the previous slab interest rates
'then enter the same rate for thse slabs

For I = 0 To FromIndex - 1
    strKey = strDef & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
    'First get any previous interest on this slab
    'IF user wants to enter the no interest rate for 0 to 30 days or 0 to 60 days then
    'this case fails 'soo need not to enter this value in this for the early values
    'i.e. when earlySlab is true
    strValue = GetInterestRateOnDate(M_ModuleID, strKey, TransDate)
    If Len(strValue) > 0 And optDays Then
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
    strValue = GetInterestRateOnDate(M_ModuleID, strKey, TransDate)
    
    If Len(strValue) > 0 Then
        strKey = strDef & cmbFrom.ItemData(I) & "-" & cmbTo.ItemData(I)
        'strValue = Val(txtGenInt) & "," & Val(txtEmpInt) & "," & Val(txtSenInt)
        Call SetUp.WriteSetupValue(strModule, strKey, strValue)
    End If
Next

Call LoadInterestRates
cmdIntApply.Enabled = False

End Sub

Private Sub cmdInterest_Click()
    Dim Cancel As Integer
    RaiseEvent PayInterest(m_AccID, Cancel)
    DepositIDDisplay
End Sub

Private Sub cmdIntPayable_Click()
If Not DateValidate(txtIntPayable.Text, "/", True) Then
    'MsgBox "Invalid Date Format Specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntPayable
    Exit Sub
End If
Me.Refresh
RaiseEvent AddInterestPayable(GetSysFormatDate(txtIntPayable), GetSysFormatDate(gStrDate))

End Sub

Private Sub cmdLoad_Click()
Dim rst As ADODB.Recordset

'Initilise the account Id
m_AccID = 0

If m_DepositType = 0 Then
    MsgBox "CREATE THE DEPOSIT AND FILL THE PROPERTIES", vbInformation
'    ActivateTextBox txtSelect
    Exit Sub
End If

'First Get the AccountId Of the for this account
gDbTrans.SqlStmt = "SELECT * FROM FDMaster WHERE " & _
    " AccNum = " & AddQuotes(txtAccNo.Text, True) & " AND DepositType = " & _
    m_DepositType

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    txtFromDate.Text = ""
'    txtToAmt.Text = ""
    txtBalance = ""
    txtCertificate.Text = ""
    txtFromDate.Text = ""
    txtEffective.Text = ""
    grdLedger.Clear
    Exit Sub
End If

m_AccNum = txtAccNo

If Not AccountLoad(rst("AccID")) Then
    Set rst = Nothing
    ActivateTextBox txtAccNo
    Exit Sub
End If
Set rst = Nothing

End Sub

Private Sub cmdMatureDate_Click()
With Calendar
    .Left = Me.Left + Me.fraDeposits.Left + Me.ActiveControl.Left - .Width / 2
    .Top = Me.Top + Me.fraDeposits.Top + Me.ActiveControl.Top
    .selDate = txtMatureDate.Text
    .Show vbModal
    txtMatureDate.Text = .selDate
End With

End Sub


Private Sub cmdNext_Click()
    m_AccID = DepositIDNext
    RaiseEvent AccountChanged(m_AccID)
    Call DepositIDDisplay
End Sub

Private Sub cmdNext_LostFocus()
'txtAccNo.SetFocus
End Sub

Private Sub cmdOk_Click()
Set frmFDAcc = Nothing
Set m_Notes = Nothing
'Ask the user before Closing
'If MsgBox(GetResourceString(750), vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
'    Exit Sub
'End If

gWindowHandle = 0
Unload Me
End Sub

Private Sub cmdPhoto_Click()
    If Not m_CustReg Is Nothing Then
        frmPhoto.setAccNo (m_CustReg.customerID)
        If (m_CustReg.customerID > 0) Then _
            frmPhoto.Show vbModal
    End If
End Sub
Private Sub cmdPrevious_Click()
    m_AccID = DepositIDPrevious
    RaiseEvent AccountChanged(m_AccID)
    Call DepositIDDisplay
End Sub
Private Sub cmdPrint_Click()
    If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    m_frmPrintTrans.optCert.Visible = True
    m_frmPrintTrans.lblLast.Visible = False
    m_frmPrintTrans.cmbRecords.Visible = False
    m_frmPrintTrans.Tag = Me.txtAccNo.Text
    m_frmPrintTrans.Show vbModal
End Sub
Private Sub cmdSave_Click()

'validate input
If Not ValidateSave Then Exit Sub


'SaveAccount
If Not AccountSave Then Exit Sub

txtAccNo.Text = GetVal("Accid")
Call cmdLoad_Click
ActivateTextBox txtAccNo

End Sub

Private Sub cmdSelect_Click()
    
    Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "SELECT * FROM DepositName"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        Set rst = Nothing
        Exit Sub
    End If
    Dim Deptype As Integer
    RaiseEvent SelectDepositType(Deptype)
    
    DepositType = Deptype Mod 100
        
End Sub

Private Sub cmdUndoIntPayble_Click()
    If Not DateValidate(txtIntPayable.Text, "/", True) Then
    '''    MsgBox "Invalid Date Format Specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtIntPayable
        Exit Sub
    End If
    Call UndoInterestPayableOfFD(txtIntPayable.Text)
    Screen.MousePointer = vbDefault
    MousePointer = vbDefault
End Sub

Private Sub cmdView_Click()

If m_DepositType = 0 Then
    MsgBox "SELECT DEPOSIT", vbInformation
    Exit Sub
End If

'First check the dates specified
Dim fromDate As String
Dim toDate As String
With txtFromDate
    If .Enabled Then
        If Not DateValidate(.Text, "/", True) Then
            'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtFromDate
            Exit Sub
        End If
        fromDate = .Text
    If WisDateDiff(.Text, txtToDate.Text) < 0 Then
        'MsgBox "TO date is earlier that the specified FROM date!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtFromDate
        Exit Sub
    End If
    End If
    
End With

With txtToDate
    If .Enabled Then
        If Not DateValidate(.Text, "/", True) Then
            'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtToDate
            Exit Sub
        End If
        toDate = .Text
    End If
End With

   
If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    

Dim ReportType As wis_FDReports

If optOpened Then ReportType = repFDAccOpen
If optClosed Then ReportType = repFDAccClose
If optDepGLedger Then ReportType = repFDLedger
If optDepositBalance Then ReportType = repFDBalance
If optDepDtCr Then ReportType = repFDDayBook
If optJoint Then ReportType = repFDJoint
If optMonthBal Then ReportType = repFDMonbal
If optMature Then ReportType = repFDMat
If optLiabilities Then ReportType = repFDLaib
If optMFD Then ReportType = repMFDBalance
If optMatDepGLedger Then ReportType = repMFDLedger
If optMatDepDtCr Then ReportType = repMFDDayBook
If optDepTrans Then ReportType = repFDTrans
If optMFDTrans Then ReportType = repMFDTrans
If optDepCashBook Then ReportType = repFDCashBook
If optMFdCashBook Then ReportType = repMFDCashBook

RaiseEvent ShowReport(ReportType, IIf(optAccId, wisByAccountNo, wisByName), _
            fromDate, toDate, m_clsRepOption)
            

ExitLine:
   On Error Resume Next
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


Private Sub m_frmPrintTrans_pageClick()
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim TransID As Long
Dim AccType As Long
Dim AccId As Long
Dim metaRst As ADODB.Recordset
Set clsPrint = New clsTransPrint

SqlStr = "SELECT  LastPrintID From SBMaster WHERE AccId = " & m_AccID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

  'While Not Rst.EOF
  clsPrint.printFDPage (m_AccID)

'If bNewPage Then
 
 
   '.printFDPage
      'clsPrint.printFDPage wis_Deposits, m_AccID
'End If
End Sub

Private Sub optClosed_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
    
    If optClosed.Value = True Then
        cmdDate1.Enabled = True
        txtFromDate.Enabled = True
        txtFromDate.BackColor = vbWhite
    End If
End Sub

Private Sub optDays_Click()


Dim Str As String
Str = " " & GetResourceString(44)

With cmbFrom
    .Clear
    .AddItem "0" & Str
    .ItemData(.newIndex) = 0
    .AddItem "15" & Str
    .ItemData(.newIndex) = 15
    .AddItem "30" & Str
    .ItemData(.newIndex) = 30
    .AddItem "45" & Str
    .ItemData(.newIndex) = 45
    .AddItem "60" & Str
    .ItemData(.newIndex) = 60
    .AddItem "90" & Str
    .ItemData(.newIndex) = 90
    .AddItem "120" & Str
    .ItemData(.newIndex) = 120
    .AddItem "180" & Str
    .ItemData(.newIndex) = 180
End With

With cmbTo
    .Clear
    .AddItem "15" & Str
    .ItemData(.newIndex) = 15
    .AddItem "30" & Str
    .ItemData(.newIndex) = 30
    .AddItem "45" & Str
    .ItemData(.newIndex) = 45
    .AddItem "60" & Str
    .ItemData(.newIndex) = 60
    .AddItem "90" & Str
    .ItemData(.newIndex) = 90
    .AddItem "120" & Str
    .ItemData(.newIndex) = 120
    .AddItem "180" & Str
    .ItemData(.newIndex) = 180
    .AddItem "1" & GetResourceString(208)
    .ItemData(.newIndex) = 365
End With

End Sub


Private Sub optDepCashBook_Click()
Call optDepDtCr_Click
End Sub

Private Sub optDepDtCr_Click()
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
    
    cmdDate1.Enabled = True
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
End Sub

Private Sub optDepGLedger_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(False, False)
    
    cmdDate1.Enabled = True
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
End Sub
Private Sub optDepositBalance_Click()
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, True)
        
    cmdDate1.Enabled = False
    txtFromDate.Enabled = False
    txtFromDate.BackColor = wisGray
End Sub

Private Sub optDepTrans_Click()
Call optDepGLedger_Click
End Sub

Private Sub optJoint_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(False, True)
    
    cmdDate1.Enabled = False
    txtFromDate.Enabled = False
    txtFromDate.BackColor = wisGray
End Sub

Private Sub optLiabilities_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
        
    If optLiabilities.Value = True Then
        cmdDate1.Enabled = False
        txtFromDate.Enabled = False
        txtFromDate.BackColor = wisGray
    End If

End Sub

Private Sub optMatDepDtCr_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
        
    cmdDate1.Enabled = True
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite

End Sub

Private Sub optMatDepGLedger_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(False, False)
        
    cmdDate1.Enabled = True
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite

End Sub

Private Sub optMature_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
        
    If optMature.Value = True Then
        cmdDate1.Enabled = True
        txtFromDate.Enabled = True
        txtFromDate.BackColor = vbWhite
    End If

End Sub

Private Sub optMFD_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, True)
    
    cmdDate1.Enabled = False
    txtFromDate.Enabled = False
    txtFromDate.BackColor = wisGray
End Sub

Private Sub optMFdCashBook_Click()
Call optDepDtCr_Click
End Sub

Private Sub optMFDTrans_Click()
Call optDepGLedger_Click
End Sub

Private Sub optMon_Click()
Dim Str As String

Str = " " & GetResourceString(208)

With cmbFrom
    .Clear
    .AddItem "1" & Str
    .ItemData(.newIndex) = 1
    .AddItem "2" & Str
    .ItemData(.newIndex) = 2
    .AddItem "3" & Str
    .ItemData(.newIndex) = 3
    .AddItem "4" & Str
    .ItemData(.newIndex) = 4
    .AddItem "5" & Str
    .ItemData(.newIndex) = 5
    .AddItem "6" & Str
    .ItemData(.newIndex) = 6
    .AddItem "7" & Str
    .ItemData(.newIndex) = 7
    .AddItem "8" & Str
    .ItemData(.newIndex) = 8
    .AddItem "9" & Str
    .ItemData(.newIndex) = 9
End With

With cmbTo
    .Clear
    .AddItem "2" & Str
    .ItemData(.newIndex) = 2
    .AddItem "3" & Str
    .ItemData(.newIndex) = 3
    .AddItem "4" & Str
    .ItemData(.newIndex) = 4
    .AddItem "5" & Str
    .ItemData(.newIndex) = 5
    .AddItem "6" & Str
    .ItemData(.newIndex) = 6
    .AddItem "7" & Str
    .ItemData(.newIndex) = 7
    .AddItem "8" & Str
    .ItemData(.newIndex) = 8
    .AddItem "9" & Str
    .ItemData(.newIndex) = 9
    .AddItem "10" & Str
    .ItemData(.newIndex) = 10

End With


End Sub

Private Sub optMonthBal_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(False, False)
        
    If optMonthBal.Value = True Then
        cmdDate1.Enabled = True
        txtFromDate.Enabled = True
        txtFromDate.BackColor = vbWhite
    End If

End Sub

Private Sub optOpened_Click()
    
    'Enable/Disable the Place,Caste, Amount Range
    Call SetOptionDialogControls(True, False)
        
    If optOpened.Value = True Then
        cmdDate1.Enabled = True
        txtFromDate.Enabled = True
        txtFromDate.BackColor = vbWhite
    End If

End Sub

Private Sub Form_Load()
cmdUndoIntPayble.Enabled = gCurrUser.IsAdmin
cmdDelete.Enabled = gCurrUser.IsAdmin

Screen.MousePointer = vbHourglass

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
cmdPrevious.Picture = LoadResPicture(101, vbResIcon)
cmdNext.Picture = LoadResPicture(102, vbResIcon)

'set Module Id
'm_ModuleID = wis_FDAcc + m_Deposittype
' Center the winodw.
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

 Dim count As Integer
Call SetKannadaCaption
lblFDName.FONTSIZE = 18
 
 'Load the property Sheet
Call LoadPropSheet

'Load Caste .Place & Gender
'Call LoadCastes(cmbCastes)
'Call LoadPlaces(cmbPlaces)
'Call LoadGender(cmbGender)
'Call LoadAccountGroups(cmbAccGroup)

'Now Load the Account Groups
Dim cmbIndex As Byte
cmbIndex = GetIndex("AccGroup")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
Call LoadAccountGroups(cmb(cmbIndex))
txtData(Val(GetIndex("AccGroup"))).Text = cmb(cmbIndex).Text

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg


' Hide the frames that are not visible currently.
fraDeposits.ZOrder (0)
fraDeposits.Visible = True

fraProps.Visible = False
fraReports.Visible = False

txtToDate = gStrDate

If gOnLine Then
    txtDepositDate.Locked = True
    cmdDepositDate.Enabled = False
End If

optDays.Value = True
Call LoadInterestRates

Screen.MousePointer = vbDefault

cmdPhoto.Enabled = Len(gImagePath)

End Sub


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



Private Function LoadPropSheet() As Boolean
Const CTL_MARGIN = 15

lblDesc.BorderStyle = 0
lblHeading.BorderStyle = 0
lblOperation.Caption = GetResourceString(54)    '"Operation Mode : <INSERT>"

'
' Read the data from SBSetup.ini and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\FDAcc_" & gLangOffSet & ".prp"
If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet Then
        PropFile = App.Path & "\FDAccKan.prp"
    Else
        PropFile = App.Path & "\FDAcc.prp"
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
            MsgBox GetResourceString(603) _
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
    Dim CmdLoaded As Boolean, ChkLoaded As Boolean
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
                    .Caption = GetResourceString(294) '"Reset"
                    .Width = 1000
                ElseIf I = 2 Then
                    .Caption = GetResourceString(295)   ' "Details..."
                    .Width = 1000
                Else
                    .Caption = "..."
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
txtData(txtIndex).Text = GetLastAccountNumber

' Show the current date wherever necessary.
txtIndex = GetIndex("CreateDate")
txtData(txtIndex).Text = gStrDate

' Set the default updation mode.
m_accUpdatemode = wis_INSERT

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

If NeedsScrollbar Then VScroll1.Visible = True

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


Private Sub Form_Unload(Cancel As Integer)
gWindowHandle = 0
RaiseEvent WindowClosed
End Sub

Private Sub TabStrip_Click()
    If m_DepositType = 0 And Not TabStrip.SelectedItem.Index = 4 Then
      '  MsgBox "CREATE THE DEPOSIT AND FILL THE PROPERTIES", vbInformation
'        ActivateTextBox txtSelect
        Exit Sub
    End If
    fraDeposits.Visible = False
    fraNew.Visible = False
    fraProps.Visible = False
    fraReports.Visible = False

If TabStrip.SelectedItem.Index = 1 Then
    fraDeposits.Visible = True
    fraDeposits.ZOrder 0
    cmdAccept.Default = True
    
ElseIf TabStrip.SelectedItem.Index = 2 Then
    fraNew.Visible = True
    fraNew.ZOrder 0
    cmdSave.Default = True
    
ElseIf TabStrip.SelectedItem.Index = 3 Then
    fraReports.Visible = True
    fraReports.ZOrder 0
    cmdView.Default = True
    
ElseIf TabStrip.SelectedItem.Index = 4 Then
    fraProps.Visible = True
    fraProps.ZOrder 0
    cmdIntApply.Default = True
    
End If

End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Index = 1 Then
    fraInstructions.Visible = True
    fraInstructions.ZOrder 0
    fraLedger.Visible = False
    fraAccStmt.Visible = False
ElseIf TabStrip1.SelectedItem.Index = 2 Then
    fraAccStmt.ZOrder 0
    fraLedger.Visible = False
    fraAccStmt.Visible = True
    fraInstructions.Visible = False
Else 'If TabStrip1.SelectedItem.Index = 2 Then
    fraLedger.ZOrder 0
    fraLedger.Visible = True
    fraAccStmt.Visible = False
    fraInstructions.Visible = False
End If

End Sub

Private Sub txtAccNo_Change()
If Me.ActiveControl.name <> txtAccNo.name Then Exit Sub

If m_DepositType > 0 Then
    If Len(txtAccNo.Text) > 0 Then
        cmdLoad.Enabled = True
    Else
        cmdLoad.Enabled = False
    End If
    Call ResetUserInterface
End If
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
ElseIf StrComp(strDispType, "Boolean", vbTextCompare) = 0 Then
    ' Get the checkbox index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    If TextIndex <> "" Then
        chk(Val(TextIndex)).SetFocus
    End If
ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
    ' Get the combo index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
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
    If I <> Val(TextIndex) Or TextIndex = "" Then cmd(I).Visible = False
Next

' Hide all other combo boxes.
For I = 0 To cmb.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then cmb(I).Visible = False
Next

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


Private Sub SetDescription(Ctl As Control)
    ' Extract the description title.
    lblHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
    lblDesc.Caption = ExtractToken(Ctl.Tag, "Description")
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
If StrComp(strDatSrc, "IntroducerID", vbTextCompare) = 0 Then
    ' Check if any data is found in this text.
    If Trim$(txtData(Index).Text) <> "" Then
        'Check if valid account number is given
        If Val(txtData(Index).Text) <= 0 Then
            'MsgBox "Invalid account number specified !", vbExclamation, gAppName & " - error"
            MsgBox GetResourceString(501), vbExclamation, gAppName & " - error"
            ActivateTextBox txtData(Index)
            Exit Sub
        End If
        gDbTrans.SqlStmt = "SELECT AccID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM FDMaster, " _
                & "NameTab WHERE FDMaster.AccID = " & Val(txtData(Index).Text) _
                & " AND FDMaster.CustomerID = NameTab.CustomerID" _
                & " AND FDMaster.DepositType = " & m_DepositType
        Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
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

Private Sub txtDays_Change()

If Me.ActiveControl.name = txtDays.name Then Call DateChange(txtDays)

Exit Sub

On Error GoTo ExitLine
If Me.ActiveControl.name <> txtDays.name Then Exit Sub
Dim Days As Long

If Val(txtDays.Text) > 99999 Then Exit Sub
If Val(txtDays.Text) <= 0 Or Not IsNumeric(txtDays.Text) Then
    txtInterest.Text = "0.00"
    txtMatureDate.Text = txtEffective.Text
    Exit Sub
Else
    Days = Val(txtDays.Text)
End If


If Not DateValidate(txtEffective, "/", True) Then Exit Sub

If Trim(txtDays.Text) = "" Then
    txtMatureDate.Text = txtDepositDate.Text
    txtEffective.Text = txtDepositDate
    Exit Sub
End If

Dim DateDep As String
Dim DateMature As String
Dim SchemeName As String

DateDep = GetSysFormatDate(txtEffective.Text)  'get the american date
DateMature = CStr(DateAdd("d", Days, CDate(DateDep)))
txtMatureDate.Text = GetIndianDate(CDate(DateMature))

If Days < 45 Then
   SchemeName = "0_1.5_"
ElseIf Days >= 45 And Days < 91 Then
   SchemeName = "1.5_3_"
ElseIf Days >= 91 And Days < 181 Then
   SchemeName = "3_6_"
ElseIf Days >= 181 And Days < 366 Then
   SchemeName = "6_12_"
ElseIf Days >= 366 And Days < (366 * 2) - 1 Then
   SchemeName = "12_24_"
ElseIf Days >= (366 * 2) - 1 And Days < 365 * 3 Then
   SchemeName = "24_36_"
ElseIf Days >= 365 * 3 And Days < 365 * 4 Then
   SchemeName = "36_48_"
ElseIf Days >= 365 * 4 And Days < 365 * 5 Then
   SchemeName = "48_60_"
ElseIf Days >= 365 * 5 Then
   SchemeName = "Above60_"
End If

SchemeName = SchemeName & "Deposit"

Dim ClsInt As New clsInterest
Me.txtInterest.Text = ClsInt.InterestRate(wis_Deposits + m_DepositType, SchemeName, txtEffective.Text, txtMatureDate.Text)
txtMatureAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + _
            CCur(ComputeFDInterest(Val(txtDepositAmount.Text), txtEffective.Text, Me.txtMatureDate.Text, _
            CSng(txtInterest.Text), False)))

ExitLine:
End Sub

Private Sub txtDepositAmount_Change()
        
If Not DateValidate(txtMatureDate.Text, "/", True) Then Exit Sub
If Not DateValidate(txtDepositDate.Text, "/", True) Then Exit Sub
Dim Days As Integer
Dim MatAmount As Currency
    
    Days = WisDateDiff(txtDepositDate.Text, txtMatureDate.Text)
    
    MatAmount = txtDepositAmount.Text
    MatAmount = MatAmount + _
          (txtDepositAmount * Val(txtInterest.Text) / 100 * (Days / 365))
    txtMatureAmount = MatAmount \ 1
    
End Sub

Private Sub txtDepositAmount_GotFocus()
With txtDepositAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtDepositAmount_LostFocus()

If Not m_Cumulative Then Exit Sub

If Not DateValidate(txtMatureDate.Text, "/", True) Then Exit Sub
If Not DateValidate(txtDepositDate.Text, "/", True) Then Exit Sub

Dim Days As Integer
Dim CalDays As Integer
Dim actDays As Integer
Dim MatAmount As Currency
Days = WisDateDiff(txtDepositDate.Text, txtMatureDate.Text)
    
    'Get the whether cumulative or not
    
    MatAmount = txtDepositAmount.Text
    Do
        actDays = 365
        If CalDays + 365 >= Days Then actDays = Days - (Days \ 365) * 365
      MatAmount = MatAmount + _
          (MatAmount * Val(txtInterest.Text) / 100 * (actDays / 365))
      CalDays = CalDays + 365
      If CalDays >= Days Then Exit Do
   Loop
   txtMatureAmount = MatAmount \ 1

End Sub

Private Sub txtEffective_Change()
On Error Resume Next
    If Me.ActiveControl.name = txtEffective.name Or _
        Me.ActiveControl.name = cmdEffective.name Then _
        Call DateChange(txtEffective)
Err.Clear
End Sub



Private Sub txtFromDate_GotFocus()
On Error Resume Next
With txtFromDate
    .SelStart = 0
    .SelLength = InStr(1, .Text, "/") - 1
End With

End Sub


Private Sub txtGenInt_Change()
Call EnableInterestApply
txtSenInt = txtGenInt
txtEmpInt = txtGenInt
End Sub

Private Sub txtIntPayable_GotFocus()
With txtIntPayable
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub




Private Sub txtMatureAmount_GotFocus()
With txtMatureAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtMatureDate_Change()

If Me.ActiveControl.name <> txtMatureDate.name And Me.ActiveControl.name <> cmdMatureDate.name Then Exit Sub
Call DateChange(txtMatureDate)

End Sub

Private Sub txtToDate_GotFocus()
On Error Resume Next
With txtToDate
    .SelStart = 0
    .SelLength = InStr(1, .Text, "/") - 1
End With

End Sub


Private Sub VScroll1_Change()
' Move the picSlider.
picSlider.Top = -VScroll1.Value

End Sub


Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

If gLangOffSet = wis_KannadaOffset Then
    cmdSelect.Left = 5287
    cmdSelect.Top = 120
    cmdSelect.Height = 432
    cmdSelect.Width = 2595
Else
    cmdSelect.Left = 4880
    cmdSelect.Top = 120
    cmdSelect.Height = 335
    cmdSelect.Width = 2595
End If

cmdSelect.Left = TabStrip.Left + TabStrip.Width - cmdSelect.Width
cmdSelect.Top = TabStrip.Top
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
lblDepositDate.Caption = GetResourceString(38, 37) 'Transaction date
lblEffective.Caption = GetResourceString(43, 37) 'Deposit Date
lblDays.Caption = GetResourceString(44) & GetResourceString(92)
lblDepositAmount.Caption = GetResourceString(43, 40)
lblMatureDate.Caption = GetResourceString(48, 37) '"Matures On"
lblInterest.Caption = GetResourceString(47)
lblMatureAmount.Caption = GetResourceString(48) + GetResourceString(40)   '"MaturityAmount (Rs)"
lblCertificate.Caption = GetResourceString(337, 60) ' Certificate No

TabStrip1.Tabs(1).Caption = GetResourceString(219) 'Instructions
TabStrip1.Tabs(2).Caption = GetResourceString(36, 295)
TabStrip1.Tabs(3).Caption = GetResourceString(36, 38)

cmdAccept.Caption = GetResourceString(4)
cmdDelete.Caption = GetResourceString(14)
cmdClose.Caption = GetResourceString(11)    '
'cmdLoans.Caption = GetResourceString(18)
'---------------frame1 Complete-------------------

'Now Change the Font of New/modify Account Frame
cmdSave.Caption = GetResourceString(7)    '"
cmdClear.Caption = GetResourceString(8)    '
lblOperation.Caption = GetResourceString(54)
cmdDeleteAcc.Caption = GetResourceString(14)
cmdInterest.Caption = GetResourceString(47)
cmdPhoto.Caption = GetResourceString(415)
'-----------------frame2 complete---------------

'Now Change The Caption of Report Frame 4th frame
fraChooseReport.Caption = GetResourceString(288)
optDepositBalance.Caption = GetResourceString(70)
optDepDtCr.Caption = GetResourceString(390, 63) ' Sub Day book
optMature.Caption = GetResourceString(72)   '"Deposits That Mature"
optDepGLedger.Caption = GetResourceString(43, 93) '"general ledger
optOpened.Caption = GetResourceString(64)
optClosed.Caption = GetResourceString(78)    '
optLiabilities.Caption = GetResourceString(77)
optJoint.Caption = GetResourceString(265, 43) '"Joint Deposit
optMonthBal.Caption = GetResourceString(43, 463, 42) '"Deposit
optMFD.Caption = GetResourceString(46, 43) '"Matured Deposit
optMatDepDtCr.Caption = GetResourceString(220, 390, 63)
optMatDepGLedger.Caption = GetResourceString(220, 93) '"Deposit Daily Ledger
optDepTrans.Caption = GetResourceString(43, 28, 295) '"Deposit Transaction Details
optMFDTrans.Caption = GetResourceString(220, 28, 295) '"MAt Deposit Transaction MAde
optDepCashBook.Caption = GetResourceString(43, 390, 85) '"Sub Cash Book
optMFdCashBook.Caption = GetResourceString(220, 390, 85) '"MAt Deposit sub cash Book

'fraOrder.Caption = GetResourceString(287)
optAccId.Caption = GetResourceString(36, 60)
optName.Caption = GetResourceString(35)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
cmdView.Caption = GetResourceString(3)
fraDateRange.Caption = GetResourceString(106)  '"Specify a Date range"
cmdView.Caption = GetResourceString(13)

'now Change the Captions Of Properites frame"
fraInterest.Caption = GetResourceString(186)
lblFrom.Caption = GetResourceString(109)
lblTo.Caption = GetResourceString(110)
lblIntPayable.Caption = GetResourceString(110)
'lblMaxLoanPercent.Caption = GetResourceString(194)
Me.lblLastIntDate.Caption = GetResourceString(228) '"Last Interest Updated on :"
lblIntPayable.Caption = GetResourceString(450, 37)
cmdIntPayable.Caption = GetResourceString(450, 171)
cmdUndoIntPayble.Caption = GetResourceString(5, 450)

lblStatus.Caption = "" 'GetResourceString(190)
optDays.Caption = GetResourceString(44) & GetResourceString(92)
optMon.Caption = GetResourceString(192) & GetResourceString(92)
lblGenlInt = GetResourceString(344)
lblEmpInt = GetResourceString(155, 47) & GetResourceString(305)

cmdIntApply.Caption = GetResourceString(6)
lblFDBalance = GetResourceString(42)

cmdAdvance.Caption = GetResourceString(491)    'Options
End Sub

Private Function UndoLastTransaction() As Boolean

Dim TransID As Long
Dim IsIntPayableTrans As Boolean
Dim rst As ADODB.Recordset

Dim PayableHeadID As Long
Dim IntHeadID As Long
Dim MatFdHeadId As Long
Dim DepAmount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim TransDate As Date

Dim headName As String
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes
Dim payableTransType As wisTransactionTypes


If m_AccHeadId = 0 Then GoTo ErrLine

'Get the last transaction ID
TransID = GetFDMaxTransID(m_AccID)

gDbTrans.SqlStmt = "Select * from FDTrans " & _
            " where AccID = " & m_AccID & " And TransID = " & TransID

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    TransDate = rst("TransDate")
    DepAmount = FormatField(rst("Amount"))
    transType = rst("TransType")
End If

gDbTrans.SqlStmt = "Select top 1 * from FDIntTrans " & _
        " where AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    IntAmount = FormatField(rst("Amount"))
    IntTransType = rst("TransType")
    TransDate = rst("TransDate")
    headName = GetDepositTypeText(m_DepositType) & " " & GetResourceString(487)
    IntHeadID = GetIndexHeadID(headName)
End If

If TransID = 1 Then
    'MsgBox "No transaction has been performaed on this account!", vbInformation, gAppName & " - Message"
    MsgBox GetResourceString(645), vbInformation, gAppName & " - Message"
    Exit Function
End If

'Now Check whether last transaction is Interest PAyble
gDbTrans.SqlStmt = "SELECT top 1 * FROM FDIntPayable " & _
                " WHERE AccID =" & m_AccID & " And TransID = " & TransID

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    IsIntPayableTrans = True
    PayableAmount = FormatField(rst("Amount"))
    PayableAmount = FormatField(rst("Amount"))
    TransDate = rst("TransDate")
    payableTransType = rst("TransType")
    'The TransCtion is the Interest payble so Confirm once again
    If MsgBox("Last Transaction is the interest payble" & vbCrLf & _
                GetResourceString(541), vbInformation + vbYesNo + _
                vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    headName = GetDepositTypeText(m_DepositType) & " " & _
            GetResourceString(375, 47)
    PayableHeadID = GetIndexHeadID(headName)
Else
    'Confirm about transaction
    'If MsgBox("Are you sure you want to undo a previous transaction ?", _
        vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
    If MsgBox(GetResourceString(583), vbQuestion + vbYesNo, _
                        wis_MESSAGE_TITLE) = vbNo Then Exit Function

End If

If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
                " WHERE AccHeadID = " & m_AccHeadId & _
                " And AccID = " & m_AccID & " And TransID = " & TransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
                        UndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If

Dim InTrans As Boolean
InTrans = True
gDbTrans.BeginTrans
 
 'Remove the last transaction From The FDTrans
  Debug.Assert TransID = 1
 'Stopped to Cheque the code which is combination of below 3(three)
 gDbTrans.SqlStmt = "Delete from FDTrans where " & _
            " AccID = " & m_AccID & " And TransID = " & TransID
 If Not gDbTrans.SQLExecute Then GoTo ErrLine
 
 'Remove the last transaction From The FDIntTrans
 gDbTrans.SqlStmt = "Delete from FDIntTrans where " & _
            " AccID = " & m_AccID & " and TransID = " & TransID
If Not gDbTrans.SQLExecute Then GoTo ErrLine

 'Remove the last transaction From The FDIntTrans
gDbTrans.SqlStmt = "Delete from FDIntPayable where " & _
            " AccID = " & m_AccID & " and TransID = " & TransID
'if amount is withdrawn from the Interes Payble account then
If Not gDbTrans.SQLExecute Then GoTo ErrLine

gDbTrans.SqlStmt = "Delete * from FDTrans A,FdIntTrans B, FdIntPayable C " & _
            " Where A.AccID = " & m_AccID & " And A.TransID = " & TransID & _
            " And B.AccID = A.AccID And B.TransID = A.TransID" & _
            " And C.AccID = A.AccID And C.TransID = A.TransID;"
'This sql will not work and this has done for testing
'If Not gDbTrans.SQLExecute Then GoTo ErrLine

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

'Undo Principal amount
If transType = wDeposit Or transType = wContraDeposit Then
    If transType = wContraDeposit And DepAmount = IntAmount Then
        If Not bankClass.UndoContraTrans(IntHeadID, m_AccHeadId, _
                                DepAmount, TransDate) Then GoTo ErrLine
        IntAmount = 0
    Else
        If Not bankClass.UndoCashDeposits(m_AccHeadId, _
                                DepAmount, TransDate) Then GoTo ErrLine
        If Not bankClass.UndoCashDeposits(IntHeadID, IntAmount, TransDate) Then GoTo ErrLine
    End If
Else
    If Not bankClass.UndoCashWithdrawls(m_AccHeadId, DepAmount, TransDate) Then GoTo ErrLine
    'If Not BankClass.UndoCashDeposits(IntHeadId, IntAmount, Transdate) Then GoTo ErrLine
End If

'Undo of interest
If IntTransType = wDeposit Or IntTransType = wContraDeposit Then
    
    If Not bankClass.UndoCashDeposits(IntHeadID, IntAmount, TransDate) Then GoTo ErrLine

ElseIf transType = wContraWithdraw Or IntTransType = wWithdraw Then
    
    If transType = wContraWithdraw And IntAmount = PayableAmount Then
        If Not bankClass.UndoContraTrans(IntHeadID, PayableHeadID, _
                    IntAmount, TransDate) Then GoTo ErrLine
        IntAmount = 0
    Else
        If Not bankClass.UndoCashWithdrawls(IntHeadID, _
                    IntAmount, TransDate) Then GoTo ErrLine
        PayableAmount = 0
    End If
End If

If payableTransType = wDeposit Or payableTransType = wContraDeposit Then
    If Not bankClass.UndoCashDeposits(PayableHeadID, PayableAmount, TransDate) Then GoTo ErrLine
ElseIf PayableHeadID Then
    If Not bankClass.UndoCashWithdrawls(PayableHeadID, PayableAmount, TransDate) Then GoTo ErrLine
End If

gDbTrans.CommitTrans
InTrans = False
UndoLastTransaction = True
   
Exit Function

ErrLine:

If Err Then
    MsgBox GetResourceString(561), vbCritical, wis_MESSAGE_TITLE
Else
    'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
     MsgBox GetResourceString(609), vbCritical, gAppName & " - Critical Error"
End If

If InTrans Then gDbTrans.RollBack

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
    If StrComp(strVisible, "True") = 0 Then VisibleCount = VisibleCount + 1
Next
Err_line:
End Function

Private Sub SetOptionDialogControls(EnablePlaceCaste As Boolean, EnableAmount As Boolean)
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = EnablePlaceCaste
        .EnableAmountRange = EnableAmount
    End With
    
End Sub
