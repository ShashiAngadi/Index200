VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmPDAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INDEX-2000   -  Pigmy Deposit  Account Wizard"
   ClientHeight    =   8505
   ClientLeft      =   1020
   ClientTop       =   1065
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   430
      Left            =   6960
      TabIndex        =   23
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame fraAgent 
      ClipControls    =   0   'False
      Height          =   7000
      Left            =   360
      TabIndex        =   24
      Top             =   660
      Width           =   7650
      Begin VB.CommandButton cmdAgentTransactDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3060
         TabIndex        =   28
         Top             =   1027
         Width           =   315
      End
      Begin VB.TextBox txtAgentCheque 
         Height          =   315
         Left            =   5880
         TabIndex        =   34
         Top             =   1650
         Width           =   1380
      End
      Begin VB.ComboBox cmbAgentTrans 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1650
         Width           =   1935
      End
      Begin VB.CommandButton cmdAgentAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   450
         Left            =   6300
         TabIndex        =   37
         Top             =   6435
         Width           =   1215
      End
      Begin VB.TextBox txtAgentDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   29
         Top             =   1027
         Width           =   1545
      End
      Begin VB.CommandButton cmdAgentUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   450
         Left            =   4410
         TabIndex        =   38
         Top             =   6435
         Width           =   1575
      End
      Begin VB.ComboBox cmbAgentParticulars 
         Height          =   315
         Left            =   1455
         TabIndex        =   36
         Top             =   2280
         Width           =   5850
      End
      Begin VB.ComboBox cmbAgentList 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PDacc.frx":0000
         Left            =   1335
         List            =   "PDacc.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   270
         Width           =   5850
      End
      Begin WIS_Currency_Text_Box.CurrText txtAgentAmount 
         Height          =   345
         Left            =   5880
         TabIndex        =   31
         Top             =   1020
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Frame fraAgentPassbook 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2415
         Left            =   270
         TabIndex        =   95
         Top             =   3660
         Width           =   7005
         Begin VB.CommandButton cmdAgentPrint 
            Height          =   375
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   1860
            Width           =   435
         End
         Begin VB.CommandButton cmdAgentNextTrans 
            Height          =   375
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   600
            Width           =   435
         End
         Begin VB.CommandButton cmdAgentPrevTrans 
            Height          =   375
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   180
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grdAgent 
            Height          =   2475
            Left            =   90
            TabIndex        =   99
            Top             =   30
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   4366
            _Version        =   393216
            Rows            =   5
            Cols            =   3
            WordWrap        =   -1  'True
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame fraAgentInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2475
         Left            =   300
         TabIndex        =   100
         Top             =   3555
         Width           =   6975
         Begin RichTextLib.RichTextBox rtfAgent 
            Height          =   2235
            Left            =   60
            TabIndex        =   102
            Top             =   150
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   3942
            _Version        =   393217
            TextRTF         =   $"PDacc.frx":0004
         End
         Begin VB.CommandButton cmdAgentNote 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6450
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   150
            Width           =   405
         End
      End
      Begin ComctlLib.TabStrip TabAgentStrip2 
         Height          =   3195
         Left            =   150
         TabIndex        =   103
         Top             =   3075
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5636
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
      Begin VB.Line Line4 
         BorderWidth     =   2
         DrawMode        =   2  'Blackness
         X1              =   90
         X2              =   7140
         Y1              =   806
         Y2              =   806
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   150
         X2              =   7200
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Label lblAgentTrans 
         Caption         =   "Transaction : "
         Height          =   300
         Left            =   150
         TabIndex        =   39
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblAgentAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   300
         Left            =   3870
         TabIndex        =   30
         Top             =   1140
         Width           =   1605
      End
      Begin VB.Label lblAgentInstrNo 
         Caption         =   "Instument no:"
         Height          =   300
         Left            =   3870
         TabIndex        =   33
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Label lblAgentParticular 
         Caption         =   "Particulars : "
         Height          =   300
         Left            =   180
         TabIndex        =   35
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label lblAgentDate 
         Caption         =   "Date : "
         Height          =   300
         Left            =   180
         TabIndex        =   27
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label lblAgent 
         Caption         =   "Agent:"
         Height          =   300
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame fraNew 
      Height          =   7000
      Left            =   360
      TabIndex        =   49
      Top             =   660
      Width           =   7650
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   450
         Left            =   6210
         TabIndex        =   133
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "&Terminate"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6210
         TabIndex        =   48
         Top             =   5190
         Width           =   1200
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   5280
         Left            =   150
         ScaleHeight     =   5220
         ScaleWidth      =   5895
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1275
         Width           =   5955
         Begin VB.VScrollBar VScroll1 
            Height          =   4785
            Left            =   5640
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picSlider 
            Height          =   3645
            Left            =   -45
            ScaleHeight     =   3585
            ScaleWidth      =   5610
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   15
            Width           =   5670
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Caption         =   "Check1"
               Height          =   255
               Index           =   0
               Left            =   4470
               TabIndex        =   82
               Top             =   0
               Width           =   315
            End
            Begin VB.ComboBox cmb 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   80
               Top             =   -30
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   240
               Index           =   0
               Left            =   4860
               TabIndex        =   79
               Top             =   0
               Visible         =   0   'False
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
               TabIndex        =   50
               TabStop         =   0   'False
               Text            =   "Account Holder"
               Top             =   0
               Width           =   2475
            End
            Begin VB.TextBox txtData 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   41
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
         ScaleWidth      =   5865
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   255
         Width           =   5925
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
            Height          =   525
            Left            =   990
            TabIndex        =   78
            Top             =   360
            Width           =   4860
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
            TabIndex        =   77
            Top             =   45
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Clear"
         Height          =   450
         Left            =   6210
         TabIndex        =   43
         Top             =   6195
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   450
         Left            =   6210
         TabIndex        =   42
         Top             =   5685
         Width           =   1200
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
         Height          =   195
         Left            =   135
         TabIndex        =   81
         Top             =   6540
         Width           =   1230
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   7695
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   13573
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Agent Transactions"
            Key             =   "AgentTrans"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transactions"
            Key             =   "Transactions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New / Modify Account"
            Key             =   "AddModify"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
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
      Height          =   7000
      Left            =   360
      TabIndex        =   83
      Top             =   660
      Width           =   7650
      Begin VB.TextBox txtIntPayable 
         Height          =   345
         Left            =   2340
         TabIndex        =   129
         Top             =   4950
         Width           =   1215
      End
      Begin VB.CommandButton cmdIntPayable 
         Caption         =   "Update Interest Payable"
         Height          =   450
         Left            =   270
         TabIndex        =   128
         Top             =   5520
         Width           =   2565
      End
      Begin VB.CommandButton cmdUndoPayable 
         Caption         =   "Undo Interest payble"
         Height          =   450
         Left            =   4890
         TabIndex        =   127
         Top             =   5520
         Width           =   2565
      End
      Begin VB.TextBox txtFailAccIDs 
         Height          =   345
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   6510
         Width           =   7185
      End
      Begin VB.Frame fraInterest 
         Caption         =   "Interest rates (%)"
         Height          =   4755
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   7650
         Begin VB.CommandButton cmdIntApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   450
            Left            =   2130
            TabIndex        =   118
            Top             =   3510
            Width           =   1215
         End
         Begin VB.TextBox txtLoanInt 
            Height          =   315
            Left            =   2550
            TabIndex        =   117
            Text            =   "+"
            Top             =   2727
            Width           =   795
         End
         Begin VB.OptionButton optMon 
            Caption         =   "Month"
            Height          =   300
            Left            =   1770
            TabIndex        =   116
            Top             =   390
            Width           =   1335
         End
         Begin VB.OptionButton optDays 
            Caption         =   "Days"
            Height          =   300
            Left            =   90
            TabIndex        =   115
            Top             =   390
            Width           =   1335
         End
         Begin VB.ComboBox cmbFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   114
            Top             =   1174
            Width           =   1335
         End
         Begin VB.ComboBox cmbTo 
            Height          =   315
            Left            =   2010
            TabIndex        =   113
            Top             =   1159
            Width           =   1335
         End
         Begin VB.TextBox txtGenInt 
            Height          =   315
            Left            =   2550
            TabIndex        =   112
            Top             =   1551
            Width           =   795
         End
         Begin VB.TextBox txtEmpInt 
            Height          =   315
            Left            =   2550
            TabIndex        =   111
            Top             =   1943
            Width           =   795
         End
         Begin VB.TextBox txtSenInt 
            Height          =   315
            Left            =   2550
            TabIndex        =   110
            Top             =   2335
            Width           =   795
         End
         Begin VB.TextBox txtIntDate 
            Height          =   345
            Left            =   1980
            TabIndex        =   109
            Top             =   3120
            Width           =   1365
         End
         Begin VB.TextBox txtPigmyCommission 
            Height          =   315
            Left            =   6150
            TabIndex        =   106
            Top             =   4200
            Width           =   945
         End
         Begin MSFlexGridLib.MSFlexGrid grdInt 
            Height          =   3885
            Left            =   3420
            TabIndex        =   108
            Top             =   150
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6853
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            AllowUserResizing=   3
         End
         Begin VB.Label lblLoanInt 
            Caption         =   "Max loan percent:"
            Height          =   300
            Left            =   90
            TabIndex        =   125
            Top             =   2757
            Width           =   1965
         End
         Begin VB.Label lblFrom 
            Caption         =   "from"
            Height          =   300
            Left            =   180
            TabIndex        =   124
            Top             =   782
            Width           =   1035
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   300
            Left            =   1770
            TabIndex        =   123
            Top             =   782
            Width           =   1095
         End
         Begin VB.Label lblGenlInt 
            Caption         =   "General Interest"
            Height          =   300
            Left            =   90
            TabIndex        =   122
            Top             =   1581
            Width           =   1995
         End
         Begin VB.Label lblEmpInt 
            Caption         =   "Emplyees Interest Rate"
            Height          =   300
            Left            =   90
            TabIndex        =   121
            Top             =   1973
            Width           =   1965
         End
         Begin VB.Label lblSenInt 
            Caption         =   "Senior Citizen"
            Height          =   300
            Left            =   90
            TabIndex        =   120
            Top             =   2365
            Width           =   1905
         End
         Begin VB.Label lblIntApply 
            Caption         =   "Int apply date"
            Height          =   300
            Left            =   90
            TabIndex        =   119
            Top             =   3150
            Width           =   1455
         End
         Begin VB.Label lblPigmycommission 
            Caption         =   "Pigmy commission"
            Height          =   300
            Left            =   4170
            TabIndex        =   107
            Top             =   4230
            Width           =   1845
         End
         Begin VB.Label lblIntInstr 
            Caption         =   " --ve Rate of interst represents Deduction Percent in Maturty Amount :"
            Height          =   555
            Left            =   120
            TabIndex        =   85
            Top             =   4110
            Width           =   3885
         End
      End
      Begin ComctlLib.ProgressBar prg 
         Height          =   345
         Left            =   300
         TabIndex        =   130
         Top             =   6510
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   609
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblLastIntDate 
         Caption         =   "Last Interest Updated on :"
         Height          =   285
         Left            =   3780
         TabIndex        =   15
         Top             =   5010
         Width           =   2115
      End
      Begin VB.Label txtLastIntDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   6090
         TabIndex        =   32
         Top             =   4950
         Width           =   1365
      End
      Begin VB.Label lblIntPayableDate 
         Caption         =   "Interest Payable Date"
         Height          =   255
         Left            =   300
         TabIndex        =   70
         Top             =   4980
         Width           =   1845
      End
      Begin VB.Label lblStatus 
         Caption         =   "x"
         Height          =   255
         Left            =   390
         TabIndex        =   71
         Top             =   6060
         Width           =   6915
      End
   End
   Begin VB.Frame fraTransact 
      Height          =   7125
      Left            =   360
      TabIndex        =   88
      Top             =   660
      Width           =   7650
      Begin VB.CheckBox chkPigmyComission 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Pigmy commission"
         Height          =   300
         Left            =   4560
         TabIndex        =   94
         Top             =   3420
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton cmdTransactDate 
         Caption         =   "..."
         Height          =   285
         Left            =   3270
         TabIndex        =   10
         Top             =   1885
         Width           =   315
      End
      Begin VB.TextBox txtCheque 
         Height          =   345
         Left            =   5790
         TabIndex        =   17
         Top             =   2345
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   1740
         MaxLength       =   9
         TabIndex        =   4
         Top             =   790
         Width           =   1395
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1250
         Width           =   5730
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2345
         Width           =   1845
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Enabled         =   0   'False
         Height          =   450
         Left            =   6330
         TabIndex        =   20
         Top             =   6540
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1740
         TabIndex        =   9
         Top             =   1885
         Width           =   1485
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   450
         Left            =   4710
         TabIndex        =   21
         Top             =   6540
         Width           =   1425
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3270
         TabIndex        =   5
         Top             =   750
         Width           =   1080
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   2745
         Width           =   5640
      End
      Begin VB.ComboBox cmbAgents 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PDacc.frx":0086
         Left            =   1740
         List            =   "PDacc.frx":0088
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   5700
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Height          =   450
         Left            =   3300
         TabIndex        =   22
         Top             =   6540
         Width           =   1215
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   5790
         TabIndex        =   14
         Top             =   1885
         Width           =   1635
         _ExtentX        =   2884
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
         Left            =   270
         TabIndex        =   91
         Top             =   3960
         Width           =   7125
         Begin VB.CommandButton cmdPrint 
            Height          =   375
            Left            =   6690
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   1740
            Width           =   435
         End
         Begin VB.CommandButton cmdNextTrans 
            Height          =   435
            Left            =   6690
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   600
            Width           =   435
         End
         Begin VB.CommandButton cmdPrevTrans 
            Height          =   435
            Left            =   6690
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   135
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   2235
            Left            =   90
            TabIndex        =   92
            Top             =   150
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   3942
            _Version        =   393216
            Rows            =   5
            Cols            =   3
            WordWrap        =   -1  'True
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2445
         Left            =   300
         TabIndex        =   89
         Top             =   3975
         Width           =   7095
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
            TabIndex        =   45
            Top             =   90
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   2235
            Left            =   60
            TabIndex        =   90
            Top             =   120
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   3942
            _Version        =   393217
            TextRTF         =   $"PDacc.frx":008A
         End
      End
      Begin ComctlLib.TabStrip TabStrip2 
         Height          =   2985
         Left            =   150
         TabIndex        =   44
         Top             =   3495
         Width           =   7395
         _ExtentX        =   13044
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
      Begin VB.Label lblBalance 
         Caption         =   "Label1"
         Height          =   255
         Left            =   4800
         TabIndex        =   135
         Top             =   840
         Width           =   2535
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   150
         X2              =   7140
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   180
         X2              =   7140
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No. : "
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   790
         Width           =   1125
      End
      Begin VB.Label lblName 
         Caption         =   "Name(s) : "
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   1250
         Width           =   1245
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   300
         Left            =   150
         TabIndex        =   11
         Top             =   2345
         Width           =   1035
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   300
         Left            =   3810
         TabIndex        =   13
         Top             =   1885
         Width           =   1215
      End
      Begin VB.Label lblInstrNo 
         Caption         =   "Instument no:"
         Height          =   300
         Left            =   3810
         TabIndex        =   16
         Top             =   2345
         Width           =   1035
      End
      Begin VB.Label lblParticular 
         Caption         =   "Particulars : "
         Height          =   300
         Left            =   180
         TabIndex        =   18
         Top             =   2805
         Width           =   1185
      End
      Begin VB.Label lblDate 
         Caption         =   "Date : "
         Height          =   300
         Left            =   210
         TabIndex        =   8
         Top             =   1885
         Width           =   1095
      End
      Begin VB.Label lblAgents 
         Caption         =   "Agent[s] :"
         Height          =   300
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Width           =   1425
      End
   End
   Begin VB.Frame fraReports 
      ForeColor       =   &H00000005&
      Height          =   7000
      Left            =   360
      TabIndex        =   86
      Top             =   660
      Width           =   7650
      Begin VB.CheckBox chkAgentName 
         Caption         =   "Show Agent Name"
         Height          =   300
         Left            =   330
         TabIndex        =   68
         Top             =   5640
         Width           =   4215
      End
      Begin VB.Frame fraOrder 
         Caption         =   "Order By"
         Height          =   2220
         Left            =   135
         TabIndex        =   104
         Top             =   4050
         Width           =   7245
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   450
            Left            =   5670
            TabIndex        =   69
            Top             =   1545
            Width           =   1215
         End
         Begin VB.TextBox txtToDate 
            Height          =   350
            Left            =   5295
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFromDate 
            Height          =   350
            Left            =   1560
            TabIndex        =   64
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdFromDate 
            Caption         =   "..."
            Height          =   315
            Left            =   2850
            TabIndex        =   63
            Top             =   1080
            Width           =   315
         End
         Begin VB.CommandButton cmdToDate 
            Caption         =   "..."
            Height          =   315
            Left            =   6570
            TabIndex        =   66
            Top             =   1080
            Width           =   315
         End
         Begin VB.OptionButton optName 
            Caption         =   "Name "
            Height          =   300
            Left            =   3870
            TabIndex        =   61
            Top             =   360
            Width           =   2310
         End
         Begin VB.OptionButton optAccID 
            Caption         =   "Account No"
            Height          =   300
            Left            =   210
            TabIndex        =   60
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label lblDate2 
            Caption         =   "but before (dd/mm/yyyy)"
            Height          =   300
            Left            =   3300
            TabIndex        =   65
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   300
            Left            =   60
            TabIndex        =   62
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Line Line6 
            X1              =   7050
            X2              =   90
            Y1              =   915
            Y2              =   915
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   525
         Left            =   6000
         TabIndex        =   72
         Top             =   6390
         Width           =   1335
      End
      Begin VB.Frame fraChooseReports 
         Caption         =   "Choose a report"
         Height          =   3900
         Left            =   135
         TabIndex        =   87
         Top             =   210
         Width           =   7245
         Begin VB.OptionButton optDailyTransaction 
            Caption         =   "Month Daily Transaction"
            Height          =   300
            Left            =   240
            TabIndex        =   134
            Top             =   3360
            Width           =   2655
         End
         Begin VB.ComboBox cmbRepAgent 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "PDacc.frx":010C
            Left            =   1755
            List            =   "PDacc.frx":010E
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   300
            Width           =   4380
         End
         Begin VB.OptionButton optSubCashBook 
            Caption         =   "Sub Cash book"
            Height          =   300
            Left            =   240
            TabIndex        =   105
            Top             =   1884
            Width           =   3375
         End
         Begin VB.OptionButton optMonthlyBalance 
            Caption         =   "Monthly Balance"
            Height          =   300
            Left            =   3870
            TabIndex        =   55
            Top             =   2868
            Width           =   2595
         End
         Begin VB.OptionButton optMonthly 
            Caption         =   "Monthly transction"
            Height          =   300
            Left            =   3870
            TabIndex        =   57
            Top             =   1392
            Width           =   2715
         End
         Begin VB.OptionButton optClosed 
            Caption         =   "Deposits closed"
            Height          =   300
            Left            =   3870
            TabIndex        =   59
            Top             =   2376
            Width           =   2745
         End
         Begin VB.OptionButton optOpened 
            Caption         =   "Deposits opened"
            Height          =   300
            Left            =   3870
            TabIndex        =   54
            Top             =   1884
            Width           =   2655
         End
         Begin VB.OptionButton optSubDayBook 
            Caption         =   "Sub day book"
            Height          =   300
            Left            =   240
            TabIndex        =   52
            Top             =   1392
            Width           =   2775
         End
         Begin VB.OptionButton optMature 
            Caption         =   "Deposits that mature"
            Height          =   300
            Left            =   240
            TabIndex        =   53
            Top             =   2868
            Width           =   2775
         End
         Begin VB.OptionButton optDepGLedger 
            Caption         =   "Deposit General Ledger"
            Height          =   300
            Left            =   240
            TabIndex        =   58
            Top             =   2376
            Width           =   2775
         End
         Begin VB.OptionButton optDepositBalance 
            Caption         =   "Deposits Where Balances"
            Height          =   300
            Left            =   240
            TabIndex        =   51
            Top             =   900
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optAgentTrans 
            Caption         =   "Agent TransaCtions"
            Height          =   300
            Left            =   3870
            TabIndex        =   56
            Top             =   900
            Width           =   2595
         End
         Begin VB.Label lblRepAgent 
            Caption         =   "Agent:"
            Height          =   300
            Left            =   300
            TabIndex        =   132
            Top             =   360
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmPDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_AccID As Long

Private m_PDHeadId As Long
Private M_ModuleID As wisModules

Private m_AccClosed As Boolean
Private m_rstAgent As ADODB.Recordset
Private m_rstPassBook As ADODB.Recordset
Private m_CustReg As New clsCustReg
Private m_Notes As New clsNotes
Private m_AgentNotes As New clsNotes
Private M_setUp As New clsSetup

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_clsRepOption As clsRepOption
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private M_UserPermission As wis_Permissions

Const CTL_MARGIN = 15
Private m_accUpdatemode As Integer


Public Event WindowClosed()
Public Event AccountChanged(ByVal AccId As Long)
Public Event AgentChanged(ByVal AgentID As Integer)
Public Event ShowReport(ShowAgent As Boolean, ReportType As wis_PDReports, ReportOrder As wis_ReportOrder, _
            fromDate As String, toDate As String, RepOption As clsRepOption, AgentID As Integer)
            
Private Sub AgentGridInitialize()
    With grdAgent
        .Clear: .Cols = 5
        .Rows = 12: .FixedRows = 1: .FixedCols = 0
        .Row = 0:
        grdAgent.Col = 0: grdAgent.Text = GetResourceString(37): .ColWidth(0) = 1200   ' "Date"
        grdAgent.Col = 1: grdAgent.Text = GetResourceString(39): .ColWidth(1) = 950    '"Particulars"
        grdAgent.Col = 2: grdAgent.Text = GetResourceString(328): .ColWidth(2) = 1400  '"Pigmy Commission"
        grdAgent.Col = 3: grdAgent.Text = GetResourceString(271): .ColWidth(3) = 900  '"Debit"
        grdAgent.Col = 4: grdAgent.Text = GetResourceString(42): .ColWidth(4) = 1000    '"Balance"
    End With

End Sub

Private Function AgentTransaction() As Boolean

Dim TransDate As Date
Dim Amount As Currency
Dim PrevAmount As Currency
Dim LastDate As Date
Dim Trans As wisTransactionTypes
Dim AgentID As Long

''Validate the controls
'Check whether Agent has selected or not
If cmbAgentList.ListIndex < 0 Then
    'MsgBox "you hace not specified the agnet name", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(590), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

With cmbAgentList
    AgentID = .ItemData(.ListIndex)
End With
'Check the date of transaction
If Not DateValidate(txtAgentDate, "/", True) Then
    'MsgBox "Invalid date speicfied", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtAgentDate
    Exit Function
End If

TransDate = GetSysFormatDate(txtAgentDate)
Amount = txtAgentAmount

'Get Th LAst TransCtion Date
Dim Balance As Currency
Dim IntBalance As Currency
Dim TransID As Long
Dim rst As Recordset

'Get the Balance and new transid
LastDate = "1/1/100"

gDbTrans.SqlStmt = "Select TOP 1 * from AgentTrans " & _
            " Where AgentID = " & AgentID & " order by TransID desc"

With cmbAgentTrans
    If cmbAgentTrans.ListIndex = -1 Then
        'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(588), vbExclamation, gAppName & " - Error"
        cmbAgentTrans.SetFocus
        Exit Function
    Else
        If .ListIndex = 0 Then Trans = wDeposit
    End If
End With

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'Check The Transaction date w.r.t to Account CreateDate
    TransID = 100
    LastDate = "1/1/100"
    Balance = Val(InputBox("Please enter a balance to start with as this account has not transaction performed", "Initial Balance", "0.00"))
    If Balance < 0 Then
        'MsgBox "Invalid initial balance specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(517), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
Else
    Balance = FormatField(rst("Balance"))
    TransID = FormatField(rst("TransID")) + 1
    LastDate = rst.Fields("TransDate")
    
    'See if the date is earlier than last date of transaction
    If DateDiff("D", TransDate, LastDate) > 0 Then
        'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(568), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtAgentDate
        Exit Function
    ElseIf DateDiff("D", TransDate, LastDate) = 0 Then
        PrevAmount = FormatField(rst.Fields("Amount"))
        If MsgBox("Transaction of " & Me.cmbAgentList.Text & _
                " On " & txtAgentDate & " already made " & vbCrLf & _
                " Do you want to update this transaction", vbQuestion + vbYesNo, _
                wis_MESSAGE_TITLE) = vbNo Then GoTo SecondTransaction
    End If
End If


'Get the Particulars
    Dim Particulars As String
    Particulars = Trim$(cmbAgentParticulars.Text)
    If Particulars = "" Then Particulars = " "
    Balance = Balance + Amount
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Insert into AgentTrans (AgentID, TransID, " & _
            " TransDate, Amount,Balance, Particulars," & _
            " TransType, VoucherNo) values ( " & _
            AgentID & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            Balance & "," & _
            AddQuotes(Particulars, True) & "," & _
            Trans & "," & _
            AddQuotes(Trim$(txtAgentCheque.Text), True) & ")"
    
    If DateDiff("D", TransDate, LastDate) = 0 Then
        Balance = rst("Balance") - PrevAmount + Amount
        TransID = rst("TransID")
        gDbTrans.SqlStmt = "UPDATE AgentTrans " & _
            " SET Amount = " & Amount & "" & _
            " WHERE AgentID = " & AgentID & _
            " AND TransID = " & TransID & _
            " AND #" & TransDate & "#"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
        
        gDbTrans.SqlStmt = "UPDATE AgentTrans " & _
            " SET Balance = Balance + " & "(" & Amount - PrevAmount & ")" & _
            " WHERE AgentID = " & AgentID & "" & _
            " AND TransID >= " & TransID & _
            " AND TransDate >= #" & TransDate & "#"
    End If
    
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
'Perform the tranaction in the Bank Head
If Not bankClass.UpdateCashDeposits(m_PDHeadId, Amount - PrevAmount, TransDate) Then
    gDbTrans.RollBack
    Set bankClass = Nothing
    Exit Function
End If

Set bankClass = Nothing
    
    gDbTrans.CommitTrans
    
    cmbAgentTrans.ListIndex = -1
    txtAgentAmount.Text = ""

Exit Function

SecondTransaction:

    Particulars = Trim$(cmbAgentParticulars.Text)
    If Particulars = "" Then Particulars = " "
    Balance = Balance + Amount
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Insert into AgentTrans (AgentID, TransID, " & _
            " TransDate, Amount,Balance, Particulars," & _
            " TransType, VoucherNo) values ( " & _
            AgentID & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            Balance & "," & _
            AddQuotes(Particulars, True) & "," & _
            Trans & "," & _
            AddQuotes(Trim$(txtAgentCheque.Text), True) & ")"
    
   
    
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If

'Dim BankClass As clsBankAcc
Set bankClass = New clsBankAcc
'Perform the tranaction in the Bank Head
If Not bankClass.UpdateCashDeposits(m_PDHeadId, Amount, TransDate) Then
    gDbTrans.RollBack
    Set bankClass = Nothing
    Exit Function
End If

Set bankClass = Nothing
    
    gDbTrans.CommitTrans
    
    cmbAgentTrans.ListIndex = -1
    txtAgentAmount.Text = ""

End Function

Private Function GetPigmyAmount() As Currency
 
'Setup an err Handler
On Error GoTo Err_line
 
If cmbTrans.ListIndex = -1 Then Exit Function
If m_AccID <= 0 Then Exit Function

'Get the amount
    Dim txtIndex As Byte
    Dim rst As Recordset
    Dim Amount As Currency
    Dim lstIndex As Byte
    
    'Get the index of the deposit only
    lstIndex = cmbTrans.ListIndex
        
    'Get the Amount here
    gDbTrans.SqlStmt = "SELECT * from PDMaster where AccId=" & m_AccID
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo Err_line
    
    Amount = FormatCurrency(rst("PigmyAmount"))
    
    If lstIndex = 0 Then GetPigmyAmount = Amount
    
  Exit Function
  
Err_line:
    
    MsgBox "GetPigmyAmount() : " & Chr(13) + Chr(10) & Err.Description, vbInformation _
                                , wis_MESSAGE_TITLE & " - Error "
  
  
End Function
 
Private Function UndoAgentLastTrans() As Boolean

Dim TransID As Long
Dim TransDate As Date
Dim AgentID As Long
Dim Amount  As Currency
Dim AgentDate As Date
Dim rst As ADODB.Recordset

With cmbAgentList
    If .ListIndex < 0 Then Exit Function
    AgentID = .ItemData(.ListIndex)
End With

gDbTrans.SqlStmt = "SELECT TOP 1 * From AgentTrans " & _
    " WHERE AgentID = " & AgentID & _
    " ORDER BY TransId Desc "

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(645), vbInformation, gAppName & " - Error"
    Exit Function
End If

Amount = rst("Amount")
TransID = rst("TransID")
TransDate = rst("TransDate")

'Confirm Deletion
'If MsgBox("Are you sure you want to undo the last transaction ?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(583), vbYesNo + vbQuestion, _
        gAppName & " - Error") = vbNo Then Exit Function

'Delete the last Transaction made by him
gDbTrans.BeginTrans
gDbTrans.SqlStmt = "DELETE * FROM AgentTrans WHERE AgentID = " & AgentID & _
    " AND TransID = " & TransID

If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc
If Not bankClass.UndoCashDeposits(m_PDHeadId, Amount, TransDate) Then
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.CommitTrans

UndoAgentLastTrans = True

End Function

Private Function UndoInterestPayableOfPD(OnIndianDate As String) As Boolean
lblStatus = ""

Dim DimPos As Integer
Dim OnDate As Date
OnDate = GetSysFormatDate(OnIndianDate)
Dim rst As ADODB.Recordset

DimPos = InStr(1, OnIndianDate, "31/3/", vbTextCompare)
If DimPos = 0 Then DimPos = InStr(1, OnIndianDate, "31/03/", vbTextCompare)
If DimPos = 0 Then
    'MsgBox "Unable to perform the transactions", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

OnDate = GetSysFormatDate(OnIndianDate)

'Before undoing check whether he has already added the interestpayble amount or not
gDbTrans.SqlStmt = "Select *  from PDIntTrans Where " & _
    " TransDate = #" & OnDate & "# " & _
    " And Particulars ='Interest Payable'"

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    'MsgBox "No interests were deposited on the specified date !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(623), vbExclamation, gAppName & " - Error"
    UndoInterestPayableOfPD = True
    Exit Function
End If
  
Screen.MousePointer = vbHourglass
  On Error GoTo ErrLine
  'declare the variables necessary

'nwo Get amount he is deleting
'Get the Payble Amount
gDbTrans.SqlStmt = "SELECT SUM(A.Amount) From PdIntPayable A" & _
    " WHERE A.TransID = " & _
        "(SELECT TransID FROM PDIntTrans C WHERE" & _
        " Particulars = 'Interest Payable' AND TransDate = #" & OnDate & "#" & _
        " AND C.AccID = A.AccID) AND TransDate = #" & OnDate & "#" & _
    " AND A.TransID > (SELECT Max(TransID) FROM PDTrans E WHERE " & _
        " A.AccID = E.AccID)"

'Dim Rst As Recordset
Dim Amount As Currency

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ErrLine
Amount = FormatField(rst(0))


Dim SqlStr As String

'DELETE THE TRANSCTION FROM Interest payable account _
'and respective transaction in PD Interest account
SqlStr = "DELETE A.*, B.* From PDIntPayable A," & _
    " PDIntTrans B WHERE A.AccID = B.AccID " & _
    " AND B.Particulars = 'Interest Payable' "

'Where The Interest payable Transction Should be the last transaction
SqlStr = SqlStr & " AND A.TransID = (SELECT Max(TransID) FROM " & _
    " PDIntTrans C WHERE TransDate = #" & OnDate & "# AND C.AccID = A.AccID)"

'And The Interest paid Transction Should also be the last transaction
SqlStr = SqlStr & " AND B.TransID = (SELECT Max(TransID) FROM " & _
    " PDIntPayable D WHERE TransDate = #" & OnDate & "# AND D.AccID = A.AccID)"

'And Transid's of bOthe Intrest payble interest accounte should be same
'After this Transction No Transacion should have taken place in the PD TRans
SqlStr = SqlStr & " AND B.TransID = A.TransID " & _
 " AND (B.TransID > (Select Max(TransID) From PDTrans E Where E.AccId = A.AccId)) "

gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
gDbTrans.SQLExecute

'Now remove the Amount From The Ledger heads
Dim bankClass As clsBankAcc
Dim PayableHeadID As Long
Dim IntHeadID As Long
Set bankClass = New clsBankAcc

Dim headName As String
'Noew ge the Ledger head id of the Pigmy deposit payble
headName = GetResourceString(425, 450)        'PIgmy INterest provision
PayableHeadID = bankClass.GetHeadIDCreated(headName, LoadResString(425) & " " & LoadResString(450), parDepositIntProv, 0, wis_PDAcc)

headName = GetResourceString(425, 375, 47)       'PIgmy Payble INterest
IntHeadID = bankClass.GetHeadIDCreated(headName, LoadResString(425) & " " & LoadResString(375) & " " & LoadResString(47), parMemDepIntPaid, 0, wis_PDAcc)

'Now Make the same transaction to the ledger heads
If Not bankClass.UndoContraTrans(IntHeadID, PayableHeadID, Amount, OnDate) Then _
    gDbTrans.RollBack: GoTo ExitLine

gDbTrans.CommitTrans
Set bankClass = Nothing
UndoInterestPayableOfPD = True

'now Check If Any  Account are unable to the undo
gDbTrans.SqlStmt = "Select AccNum from PDMaster A,PDIntTrans B Where " & _
    " A.AccId = B.accID And TransDate = #" & OnDate & "# " & _
    " And B.Particulars ='Interest Payable'"

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo ExitLine

While Not rst.EOF
    txtFailAccIDs = txtFailAccIDs & "," & rst("AccNum")
    rst.MoveNext
Wend

txtFailAccIDs.Visible = True
txtFailAccIDs = Mid(txtFailAccIDs, 2)

Set rst = Nothing

GoTo ExitLine

ErrLine:
    MsgBox "Error In PDAccount -- Remove Interest payble", vbCritical, wis_MESSAGE_TITLE
    'Resume

ExitLine:

Set bankClass = Nothing
Screen.MousePointer = vbDefault

End Function

Public Function AgentLoad(ByVal AgentID As Integer) As Boolean
Dim rst As ADODB.Recordset
Dim Found As Boolean
'Else Fetch the Agent Name
gDbTrans.SqlStmt = "SELECT * From UserTab Where UserId = " & AgentID
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    'msgbox "Invalid Account number specified", vbInformation
    MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
    cmbAgentList.SetFocus
    Exit Function
End If

Dim count  As Integer
'Get the TotalAgents form the List.

For count = 0 To cmbAgentList.ListCount - 1
    If AgentID = cmbAgentList.ItemData(count) Then
        cmbAgentList.ListIndex = count
        Found = True
        Exit For
    End If
Next

If Not Found Then Exit Function
Call ActiveAgentDetails

gDbTrans.SqlStmt = "Select *  From AgentTrans " & _
    " Where AgentId = " & AgentID & " ORDER By TransID"
Set m_rstAgent = Nothing
Dim BalanceAmount As Currency
'if new agents and accounts are created(out of scope)
If gDbTrans.Fetch(m_rstAgent, adOpenDynamic) > 0 Then
    m_rstAgent.MoveLast
    BalanceAmount = m_rstAgent("Balance")

    'Position to first record of last page
    With m_rstAgent
        '.MoveFirst
        .Move -(.RecordCount Mod 10)
        If .AbsolutePosition < 0 Then .MoveFirst
    End With
    
    cmdUndo.Enabled = gCurrUser.IsAdmin
    AgentGridInitialize
    cmdUndo.Enabled = False
    
End If

With Me.rtfAgent
    .BackColor = IIf(AgentID, vbWhite, wisGray)
    Call m_AgentNotes.LoadNotes(wis_Users, AgentID)
End With
Call m_AgentNotes.DisplayNote(rtfAgent)

Call InitPassBook(m_rstAgent, 10, cmdAgentPrevTrans)
Call AgentBookShow
cmdAgentNextTrans.Enabled = True
AgentLoad = True
End Function

Private Sub ChkAgentNameValue(OptButton As OptionButton)
If OptButton.name = optDepGLedger.name Then
    chkAgentName.Value = 0
    chkAgentName.Enabled = False
Else
    chkAgentName.Enabled = True
End If
If OptButton.name = optAgentTrans.name Then
    chkAgentName.Value = 1
    chkAgentName.Enabled = False
End If
End Sub

Private Function AccountClose() As Boolean
    gDbTrans.SqlStmt = "Update PDMaster Set Closeddate = #" & GetSysFormatDate(txtDate.Text) & "#" & _
                    " Where And AccId = " & m_AccID
     If gDbTrans.SQLExecute Then
        AccountClose = True
    End If
End Function

Private Function AccountName(AccId As Long) As String

Dim Lret As Long
Dim rst As Recordset
'Prelim checks
    If AccId <= 0 Then Exit Function

'Query DB
        gDbTrans.SqlStmt = "SELECT AccID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM PDMaster, " _
                & "NameTab WHERE PDMaster.AccID = " & AccId _
                & " AND PDMaster.CustomerID = NameTab.CustomerID"
        Lret = gDbTrans.Fetch(rst, adOpenStatic)
        If Lret = 1 Then
            AccountName = FormatField(rst.Fields("Name"))
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
Dim AgentID As Integer
Dim AccId As Long
Dim AccNum As String

Dim rst As Recordset
'Check for valid Agent Name & Id 'Code By shashi 21/2/2000
    txtIndex = GetIndex("AgentName")
    With txtData(txtIndex)
        AgentID = GetAgentID("AgentName")
        If Trim$(.Text) = "" Then
            'MsgBox "Agent name not specified !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(590), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtData(txtIndex)
            Exit Function
        End If
    End With
    
' Check for a valid Account number.
    AccIndex = GetIndex("AccID")
    With txtData(AccIndex)
        'See if acc no has been specified
        If Trim$(.Text) = "" Then
            'MsgBox "No Account number specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            
            MsgBox GetResourceString(500), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    
        'See if account already exists if it is new
        If m_accUpdatemode = wis_INSERT Then
            gDbTrans.SqlStmt = "Select AccNum from PDMaster where AccNum = " & _
                AddQuotes(Trim$(.Text), True) & " AND AgentID = " & AgentID
        Else
            gDbTrans.SqlStmt = "Select AccID from PDMaster where AccNum = " & _
                    AddQuotes(Trim$(.Text), True) & " and AccId <> " & m_AccID & " AND AgentID = " & AgentID
        End If
        
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
            'MsgBox "Account number " & .Text & "already exists." & vbCrLf & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(545) & vbCrLf & GetResourceString(641), vbExclamation, gAppName & " - Error"
            ActivateTextBox txtData(txtIndex)
        End If
        AccNum = GetVal("AccID")
    End With

    ' Check for account holder name.
    txtIndex = GetIndex("AccName")
    With txtData(txtIndex)
        'If he has not selected the custiomer then
        If m_CustReg.CustomerID = 0 Then .Text = ""
        If Trim$(.Text) = "" Then
            'MsgBox "Account holder name not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(529), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With
    
    'Get the New Account Id
    If m_accUpdatemode = wis_INSERT Then
        AccId = 1
        gDbTrans.SqlStmt = "SELECT MAX(AccID) From PDMaster "
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then AccId = FormatField(rst(0)) + 1
    End If
    
    ' Check for nominee name.
    Dim NomineeSpecified As Boolean
    txtIndex = GetIndex("NomineeName")
    With txtData(txtIndex)
        If Trim$(.Text) = "" Then
            'MsgBox "Nominee name not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            If MsgBox(GetResourceString(558) & vbCrLf & GetResourceString(541), _
                    vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
            End If
            NomineeSpecified = False
        End If
    End With
    
    ' Check for nominee age.
    txtIndex = GetIndex("NomineeAge")
    With txtData(txtIndex)
        If Trim$(.Text) = "" And NomineeSpecified Then
            'MsgBox "Nominee age not specified!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(507), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        If Not IsNumeric(Trim$(.Text)) And NomineeSpecified Then
            'MsgBox "Invalid nominee age specified!", vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(507), vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        If NomineeSpecified And Val(Trim$(.Text)) <= 0 Or Val(Trim$(.Text)) >= 100 Then
            'MsgBox "Invalid nominee age specified!", vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(507), vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With
    ' Check for nominee relationship.
    txtIndex = GetIndex("NomineeRelation")
    With txtData(txtIndex)
        If Trim$(.Text) = "" And NomineeSpecified Then
            'MsgBox "Specify nominee relationship.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(559), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With
    'Check For Pigmy Type
    txtIndex = GetIndex("PigmyType")
    With txtData(txtIndex)
        If Trim$(.Text) = "" Then
            'MsgBox "Invalid pigmy Type specified !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(512), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
    End With
    ' Check For Installment Amount
    txtIndex = GetIndex("PigmyAmount")
    With txtData(txtIndex)
        If Trim$(.Text) = "" Then
            'MsgBox "Invalid pigmy amount specified !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(506), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtData(txtIndex)
            GoTo Exit_Line
        End If
        If Not IsNumeric(Trim$(.Text)) Then
            'MsgBox "Invalid pigmy amount specified !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(506), _
                    vbExclamation, wis_MESSAGE_TITLE
                ActivateTextBox txtData(txtIndex)
                GoTo Exit_Line
        End If
    End With
        
    'Check For Maturity Date
    If Not DateValidate(GetVal("MaturityDate"), "/", True) Then
        'MsgBox "Invalid create date specified !" & vbCrLf & "Please specify in DD/MM/YYYY format!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & GetResourceString(573), vbExclamation, gAppName & " - Error"
        txtIndex = GetIndex("MaturityDate")
        ActivateTextBox txtData(txtIndex)
        Exit Function
    End If

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
    End With
    
    txtIndex = GetIndex("IntroducerID")
    With txtData(txtIndex)
        ' Check if an introducerID has been specified.
        If Trim$(.Text) = "" Then
            'If MsgBox("No introducer has been specified!" _
                & vbCrLf & "Add this Account anyway?", vbQuestion + vbYesNo) = vbNo Then
            If MsgBox(GetResourceString(560) _
                & vbCrLf & GetResourceString(663), vbQuestion + vbYesNo) = vbNo Then
                GoTo Exit_Line
            End If
        Else
            ' Check if the introducer exists.
            gDbTrans.SqlStmt = "SELECT AccID FROM PDMaster WHERE AccID = " & Val(.Text)
            If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
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
'Compare Createdate with Maturity date
If DateDiff("d", GetSysFormatDate(GetVal("MaturityDate")), GetSysFormatDate(GetVal("CreateDate"))) > 0 Then
    'MsgBox "Date of maturity is earlier than the date of creation", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(580), vbExclamation, wis_MESSAGE_TITLE
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
        If MsgBox(GetResourceString(520) & "  " & GetVal("AccID") _
                & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
            GoTo Exit_Line
        End If
    ElseIf m_accUpdatemode = wis_INSERT Then
        'If MsgBox("This will create a new account with an account number " & GetVal("AccID") _
                & "." & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox(GetResourceString(540) & " " & GetVal("AccID") _
                & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo) = vbNo Then
            GoTo Exit_Line
        End If
    End If


'Start Transactions to Data base
    m_CustReg.ModuleId = wis_PDAcc
    gDbTrans.BeginTrans
    
    'Save the customer (or Update the customer to set to current reference)
    If Not m_CustReg.SaveCustomer Then
        'MsgBox "Unable to register customer details !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Function
    End If
    
    'Begin the Transaction First here
    gDbTrans.BeginTrans
    
    ' Insert/update to database.
    If m_accUpdatemode = wis_INSERT Then
        'nRet = MsgBox("Add this account to database?", vbQuestion + vbYesNo)
        'If nRet = vbNo Then GoTo exit_line
        ' Build the SQL insert statement.
        'Modified by shashi 21/2/2000
        gDbTrans.SqlStmt = "Insert into PDMaster (AccID,AgentId, AccNum, CustomerID, " _
                & "CreateDate, MaturityDate, PigmyType, JointHolder, Nominee, " & _
                " LastPrintID,Introduced,NomineeID, " & _
                " LedgerNo,FolioNo, PigmyAmount, RateOfInterest,AccGroupID,UserID)" & _
                "values (" & AccId & "," & _
                GetAgentID("AgentName") & ", " & _
                AddQuotes(AccNum, True) & ", " & _
                m_CustReg.CustomerID & ", " & _
                "#" & GetSysFormatDate(GetVal("CreateDate")) & "#, " & _
                "#" & GetSysFormatDate(GetVal("MaturityDate")) & "#, " & _
                AddQuotes(GetVal("PigmyType"), True) & ", " & _
                AddQuotes(GetVal("JointHolder"), True) & ", " & _
                AddQuotes(Me.Nominee, True) & ", 1," & _
                Val(GetVal("IntroducerID")) & ", 1," & _
                AddQuotes(GetVal("LedgerNo"), True) & ", " & _
                AddQuotes(GetVal("FolioNo"), True) & ", " & _
                CCur(GetVal("PigmyAmount")) & ", " & _
                GetVal("RateOfInterest") & ", " & _
                GetAccGroupID & "," & gUserID & " )"
    ElseIf m_accUpdatemode = wis_UPDATE Then
        ' The user has selected updation.
        ' Build the SQL update statement.
        gDbTrans.SqlStmt = "Update PDMaster set " & _
                " AccNum = " & AddQuotes(AccNum, True) & "," & _
                " Nominee = " & AddQuotes(Me.Nominee, True) & ", " & _
                " Introduced = " & Val(GetVal("IntroducerID")) & "," & _
                " JointHolder = " & AddQuotes(GetVal("JointHolder"), True) & ", " & _
                " MaturityDate = #" & GetSysFormatDate(GetVal("MaturityDate")) & "#, " & _
                " LedgerNo = " & AddQuotes(GetVal("LedgerNo"), True) & "," & _
                " FolioNo = " & AddQuotes(GetVal("FolioNo"), True) & ", " & _
                " CreateDate = #" & GetSysFormatDate(GetVal("CreateDate")) & "#," & _
                " PigmyAmount = " & CCur(GetVal("PigmyAmount")) & ", " & _
                " RateOfInterest = " & GetVal("RateOfInterest") & ", " & _
                " AccGroupID = " & GetAccGroupID & ", " & _
                " PigmyType = " & AddQuotes(GetVal("PigmyType"), True) & " " & _
                " WHERE AccID = " & m_AccID & _
                " AND AgentID = " & AgentID
    End If
    
    ' Insert/update the data.
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo Exit_Line
    End If
    
    'MsgBox "Saved the account details.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
    AccountSave = True
    gDbTrans.CommitTrans
    
    
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

Public Function AccountUndoLastTransaction() As Boolean

Dim ClosedON As String
Dim TransDate As Date
Dim Amount As Currency
Dim IntAmount As Currency

'Dim Ret As Integer
Dim TransID As Long
Dim SqlStmt As String
Dim LoanBalance As Currency
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes
Dim AgentID As Integer

'Prelim check
If m_AccID <= 0 Then
    'MsgBox "Account not loaded !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
    cmdUndo.Enabled = False
    Exit Function
End If

'Check if account exists
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
    Else  'Reopen the account first
        If Not AccountReopen(m_AccID) Then
            'MsgBox "Unable to reopen the account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(536), vbExclamation, gAppName & " - Error"
        End If
        'Account reopen WillUN do the Last Transaction So
        'So Exit The Function
        Exit Function
    End If
End If
        
'Get last transaction record
TransID = GetPigmyMaxTransID(m_AccID)
If TransID = 0 Then
    'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(645), vbInformation, gAppName & " - Error"
    Exit Function
End If

Dim rst As Recordset
gDbTrans.SqlStmt = "Select TOP 1 * from PDTrans where " & _
        " AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
    Amount = FormatField(rst.Fields("Amount"))
    transType = FormatField(rst("TransType"))
    TransDate = rst.Fields("TransDate")
End If
        
'Check for the Interest transaction
'Dim Rst As Recordset
gDbTrans.SqlStmt = "Select TOP 1 * from PDintTrans where " & _
            " AccID = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
    IntAmount = FormatField(rst.Fields("Amount"))
    IntTransType = FormatField(rst("TransType"))
    TransDate = rst.Fields("TransDate")
End If
        
Dim ConType As String
        
'Confirm the Which account u are about Delete....
If transType = wDeposit Or transType = wContraDeposit Then ConType = GetResourceString(271) '"Deposit"
If transType = wWithdraw Or transType = wContraWithdraw Then ConType = GetResourceString(272)  '"WithDraw"
        
'Confirm UNDO
'If MsgBox("Are you sure you want to undo the last transaction ?", vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(583) & vbCrLf & " (" & ConType & ")", _
            vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then Exit Function

If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & m_PDHeadId & _
            " And AccId = " & m_AccID & " And TransID = " & TransID
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
Dim bankClass As clsBankAcc

gDbTrans.SqlStmt = "Delete from PDTrans where " & _
            " AccID = " & m_AccID & " and TransID = " & TransID
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.SqlStmt = "Delete from PDintTrans where " & _
    " AccID = " & m_AccID & " and TransID = " & TransID
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

Dim IntHeadID As Long
Set bankClass = New clsBankAcc
IntHeadID = bankClass.GetHeadIDCreated(GetResourceString(425, 487), _
                    LoadResString(425) & " " & LoadResString(487), parMemDepIntPaid, 0, wis_PDAcc)
    
If transType = wWithdraw Then
    If Not bankClass.UndoCashWithdrawls(m_PDHeadId, Amount, TransDate) Then
        gDbTrans.RollBack
        Exit Function
    End If
    If Not bankClass.UndoCashWithdrawls(IntHeadID, IntAmount, TransDate) Then
        gDbTrans.RollBack
        Exit Function
    End If
End If

If transType = wContraDeposit Or transType = wContraWithdraw Then
    If transType = wContraDeposit Then Call bankClass.UndoContraTrans(IntHeadID, m_PDHeadId, Amount, TransDate)
    'Perform the transaction in the Bank interest Head
    If transType = wContraWithdraw Then Call bankClass.UndoContraTrans(m_PDHeadId, IntHeadID, Amount, TransDate)
End If

If GetPigmyMaxTransID(m_AccID) = 0 Then
    'If there are no more transaction
    If MsgBox(GetResourceString(539) & "Do you Want To Delete This " & _
            "Account Permanently ?", vbInformation + vbYesNo + vbDefaultButton2, _
            "Undo Last") = vbYes Then
        gDbTrans.SqlStmt = "Delete from PDMaster where AccID = " & m_AccID
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
    End If
End If

gDbTrans.CommitTrans

Set bankClass = Nothing

'If Commission Added then undo the pigmy commission also
'Call BankClass.UndoPigmyCommission(Transdate, "PD Comm " & _
                m_agentID & "-" & m_AccID & "-" & TransID)

AccountUndoLastTransaction = True

ErrLine:

End Function
Public Sub ResetAgentDetails()

cmbAgentList.ListIndex = -1
cmbAgentTrans.Enabled = False
cmbAgentTrans.BackColor = wisGray
txtAgentAmount.Enabled = False
txtAgentAmount.BackColor = wisGray
cmbAgentParticulars.Enabled = False
cmbAgentParticulars.BackColor = wisGray
txtAgentCheque.Enabled = False
txtAgentCheque.BackColor = wisGray
txtAgentDate.Enabled = False
txtAgentDate.BackColor = wisGray
grdAgent.Clear

End Sub


Private Sub SetKannadaCaption()

'Set the Kannada captions for all the controls.
Call SetFontToControlsSkipGrd(Me)

'Now Assign The Names to the Controls
'The Below Code load From The the resource file
TabStrip.Tabs(1).Caption = GetResourceString(330) & _
                           " " & GetResourceString(38)     'Agent transaction
TabStrip.Tabs(2).Caption = GetResourceString(38)
TabStrip.Tabs(3).Caption = GetResourceString(211)
TabStrip.Tabs(4).Caption = GetResourceString(283) & GetResourceString(92)
TabStrip.Tabs(5).Caption = GetResourceString(213)


cmdAccept.Caption = GetResourceString(4)   '
cmdOk.Caption = GetResourceString(1)    '"
cmdLoad.Caption = GetResourceString(3)


' TransCtion Frame
lblAgents.Caption = GetResourceString(330)
lblAccNo.Caption = GetResourceString(36) + " " + GetResourceString(60)
lblName.Caption = GetResourceString(35)
lblDate.Caption = GetResourceString(37)
lblTrans.Caption = GetResourceString(38)
lblParticular.Caption = GetResourceString(39)
lblAmount.Caption = GetResourceString(40)
'lblBalance.Caption = GetResourceString(42)
lblInstrNo.Caption = GetResourceString(41)
cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)
cmdClose.Caption = GetResourceString(11)

'chkBackLog.Caption = GetResourceString(164)
chkPigmyComission.Caption = GetResourceString(331)

TabStrip2.Tabs(1).Caption = GetResourceString(219)
TabStrip2.Tabs(2).Caption = GetResourceString(218)

'Now Change the Font of New Account Frame
cmdTerminate.Caption = GetResourceString(14)
cmdSave.Caption = GetResourceString(7)
cmdReset.Caption = GetResourceString(8)
lblOperation.Caption = GetResourceString(54)
cmdPhoto.Caption = GetResourceString(415)

'Now Change The Caption of Report Frame
fraChooseReports.Caption = GetResourceString(288)
lblRepAgent.Caption = GetResourceString(330)

optDepositBalance.Caption = GetResourceString(70)
optSubDayBook.Caption = GetResourceString(390, 63) 'sub day book
optSubCashBook.Caption = GetResourceString(390, 85) 'Sub Cash book
optDepGLedger.Caption = GetResourceString(43, 93) '"Deposit GeneralLegder

optMature.Caption = GetResourceString(72)   '"Deposits That Mature"
optOpened.Caption = GetResourceString(64)    '"
optClosed.Caption = GetResourceString(78)    '"
optAgentTrans.Caption = GetResourceString(330, 38)
optMonthly.Caption = GetResourceString(463, 283) & GetResourceString(92)
optMonthlyBalance.Caption = GetResourceString(463, 42)
optDailyTransaction.Caption = GetResourceString(410, 28)
fraOrder.Caption = GetResourceString(287)
optAccId.Caption = GetResourceString(36)
optName.Caption = GetResourceString(35)
Me.chkAgentName.Caption = GetResourceString(330, 35)

lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
lblFrom.Caption = GetResourceString(109)
lblTo.Caption = GetResourceString(110)

cmdView.Caption = GetResourceString(13) '
fraOrder.Caption = GetResourceString(106)  '"Specify a Date range"
cmdView.Caption = GetResourceString(13)


'`now Change the Captions Of Properites frame"
fraInterest.Caption = GetResourceString(191)
lblIntInstr.Caption = GetResourceString(206)
lblPigmycommission.Caption = GetResourceString(328)
lblIntPayableDate.Caption = GetResourceString(450, 37)
cmdIntPayable.Caption = GetResourceString(450, 171)
cmdUndoPayable.Caption = GetResourceString(188)

'Now assign the caption for agent trans
'Fill up particulars with default values from PDAgent.INI
    Dim Particulars As String
    Dim I As Integer
    Do
        Particulars = ReadFromIniFile("Particulars", _
                "Key" & I, gAppPath & "\PDAgent.INI")
        If Trim$(Particulars) <> "" Then
            cmbAgentParticulars.AddItem Particulars
        End If
        I = I + 1
    Loop Until Trim$(Particulars) = ""

'se the captions of Agent Tab
TabAgentStrip2.Tabs(1).Caption = GetResourceString(219)
TabAgentStrip2.Tabs(2).Caption = GetResourceString(218)

'Adjust the Grid for agent transaction
With grdAgent
    .Clear
    .Rows = 11
    .Cols = 5
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 700 ' "Date"
    .Col = 1: .Text = GetResourceString(39): .ColWidth(1) = 800 '"Particulars"
    .Col = 2: .Text = GetResourceString(328): .ColWidth(2) = 1000 '"Pigmy commission"
    .Col = 3: .Text = GetResourceString(276): .ColWidth(3) = 800 '"Debit"
    .Col = 4: .Text = GetResourceString(42): .ColWidth(4) = 900
End With
 
'Load Agent Name
    Call LoadAgentNames(cmbAgentList)
    Call LoadAgentNames(cmbRepAgent)
    ''Now add the All agents
    cmbRepAgent.AddItem GetResourceString(338, 330), 0
    cmbRepAgent.ItemData(cmbRepAgent.newIndex) = 0
    
lblAgent.Caption = GetResourceString(330)
lblAgentAmount = GetResourceString(40)
lblAgentDate = GetResourceString(37)
lblAgentTrans.Caption = GetResourceString(38)
lblAgentParticular = GetResourceString(39)
lblAgentInstrNo = GetResourceString(41)
cmdAgentAccept.Caption = GetResourceString(4)
cmdAgentUndo.Caption = GetResourceString(19)

lblStatus.Caption = "" 'GetResourceString(190)
optDays.Caption = GetResourceString(44) & GetResourceString(92)
optMon.Caption = GetResourceString(192) & GetResourceString(92)
lblGenlInt = GetResourceString(344)
lblEmpInt = GetResourceString(155, 47) & GetResourceString(305)

cmdIntApply.Caption = GetResourceString(6)
cmdAdvance.Caption = GetResourceString(491)    'Options

End Sub
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
Private Function AccountDelete(AccId As Long) As Boolean
Dim rst As Recordset
'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & AccId
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    If FormatField(rst.Fields("ClosedDate")) <> "" Then
        'MsgBox "This account has already been closed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(524), vbExclamation, gAppName
        Exit Function
    End If

'Check if transactions are there
    gDbTrans.SqlStmt = "Select TOP 1 * from PDTrans where AccID = " & AccId
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        'MsgBox "You cannot delete an account having transactions !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(553), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Delete account from DB
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Delete from PDMaster where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans
AccountDelete = True
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

Private Function GetAgentID(Optional strSource As String) As Integer
Dim count As Integer
Dim txtIndex As Integer

GetAgentID = -1
On Error Resume Next
If Trim$(strSource) <> "" Then
    txtIndex = GetIndex("AgentName")
    For count = 0 To cmb.count - 1
        If CStr(txtIndex) = ExtractToken(cmb(count).Tag, _
                "TextIndex") Then Exit For
    Next count
    If count < cmb.count Then
        GetAgentID = cmb(count).ItemData(cmb(count).ListIndex)
        Exit Function
    End If
Else
    If cmbAgents.ListIndex >= 0 Then
        GetAgentID = cmbAgents.ItemData(cmbAgents.ListIndex)
    End If
End If


End Function

' Returns the index of the control bound to "strDatasrc".
Private Function GetIndex(strDataSrc As String) As Integer
GetIndex = -1
Dim strTmp As String
Dim I As Integer
For I = 0 To txtPrompt.count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(txtPrompt(I).Tag, "DataSource")
'    Debug.Assert i <> 6
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
Private Function GetNewAccountNumber(AgentID) As String
'Generate new account number for this agent here
    
    Dim NewAccNum As String
    Dim rst As Recordset
    
        gDbTrans.SqlStmt = "Select AccNum from PDMaster where " & _
            " AgentID = " & AgentID & " order by val(AccNum) desc"
        NewAccNum = "1"
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewAccNum = Val(FormatField(rst(0))) + 1
        
    GetNewAccountNumber = NewAccNum

End Function
Private Function AccountTransaction() As Boolean

On Error GoTo ErrLine

Dim AccountCloseFlag As Boolean
Dim TransTypes As wisTransactionTypes
Dim TransDate As Date

TransDate = GetSysFormatDate(txtDate)
'Check if the date of transaction is earlier than account opening date itself
Dim ret As Integer
Dim rst As Recordset
gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & m_AccID
ret = gDbTrans.Fetch(rst, adOpenStatic)
If ret <> 1 Then
    'MsgBox "DB error !", vbCritical, gAppName & " - ERROR"
    MsgBox GetResourceString(601), vbCritical, gAppName & " - ERROR"
    Exit Function
End If
    
'Check whether pigmy commission has to debit to loss account or not
If Me.chkPigmyComission.Value = vbChecked Then ' Check whehter pigmy commission has specified
    Dim PigmyCommission As Double
    Dim SetupClass As New clsSetup
    
    PigmyCommission = Val(SetupClass.ReadSetupValue("PDAcc", "PigmyCommission", "00"))
    Set SetupClass = Nothing
    If PigmyCommission = 0 Then
        MsgBox "Please specify the pigmy commission " & vbCrLf & _
            " Then Continue the transaction", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If

'Get the Transaction Type
    Dim Trans As wisTransactionTypes
    Dim PLTrans As wisTransactionTypes
    Dim lstIndex As Integer
    With cmbTrans
        If .ListIndex = 0 Then Trans = wDeposit
        If .ListIndex = 1 Then Trans = wWithdraw
        If .ListIndex = 2 Then Trans = wContraWithdraw: PLTrans = wContraDeposit
        If .ListIndex = 3 Then Trans = wContraDeposit: PLTrans = wContraWithdraw
        lstIndex = .ListIndex
    End With

'Validate the Amount
    Dim Amount As Currency
    Amount = CCur(Trim$(txtAmount.Text))

'Validate the Cheque No
    Dim ChequeNo As Long
    
'Get the Particulars
    Dim Particulars As String
    Particulars = Trim$(cmbParticulars.Text)
    If Particulars = "" Then Particulars = " "

'Get the Balance and new transid
    Dim Balance As Currency
    Dim IntBalance As Currency
    Dim TransID As Long
    
    gDbTrans.SqlStmt = "Select TOP 1 * from PDTrans " & _
                    " where AccID = " & m_AccID & _
                    " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        Balance = Val(InputBox("Please enter a balance to start with as this account has not transaction performed", "Initial Balance", "0.00"))
        If Balance < 0 Then
            'MsgBox "Invalid initial balance specified !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(517), vbExclamation, gAppName & " - Error"
            Exit Function
        End If
        If Trans = wWithdraw Then
            If Balance <= 0 Then
                MsgBox "You are trying to withdraw the amount where there is no deposit", vbInformation, wis_MESSAGE_TITLE
                Exit Function
            End If
        End If
    Else
        Balance = FormatField(rst.Fields("Balance"))
    End If

'Calculate new balance
'nOw Get the Transaction
TransID = GetPigmyMaxTransID(m_AccID) + 1

If Trans = wWithdraw Or Trans = wContraWithdraw Then
    Balance = Balance - Amount
Else
    Balance = Balance + Amount
End If
'Perform the Transaction to the Database
gDbTrans.BeginTrans

gDbTrans.SqlStmt = "Insert into PDTrans (AccID, TransID, " & _
        " TransDate, Amount,Balance, Particulars," & _
        " TransType, VoucherNo,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        "#" & TransDate & "#," & _
        Amount & "," & _
        Balance & "," & "'" & Particulars & "'," & _
        Trans & "," & _
        AddQuotes(Trim$(txtCheque.Text), True) & "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
If lstIndex = 2 Or lstIndex = 3 Then
    'This is the case we are doing Transction which affects the
    'Profit Or Loss
    gDbTrans.SqlStmt = "Insert into PDIntTrans (AccID, TransID, " & _
            " TransDate, Amount,Balance, Particulars," & _
            " TransType, VoucherNo,UserID) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            IntBalance & "," & _
            AddQuotes(Particulars, True) & "," & _
            PLTrans & "," & _
            AddQuotes(Trim$(txtCheque.Text), True) & "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
End If

'Update the BankDetails
Dim ClsBank As clsBankAcc

Set ClsBank = New clsBankAcc

'Perform the tranaction in the Bank Head
If lstIndex = 1 Then Call ClsBank.UpdateCashWithDrawls(m_PDHeadId, Amount, TransDate)
'Perform the tranaction in the Bank  interest Head
If lstIndex > 1 Then
    Dim IntHeadID As Long
    IntHeadID = ClsBank.GetHeadIDCreated(GetResourceString(425, 487), _
                LoadResString(425) & " " & LoadResString(487), parMemDepIntPaid, 0, wis_PDAcc)
    If lstIndex = 2 Then _
        If Not ClsBank.UpdateContraTrans(m_PDHeadId, IntHeadID, Amount, TransDate) Then gDbTrans.RollBacknRaiseError
    'Perform the tranaction in the Bank interest Head
    If lstIndex = 3 Then _
        If Not ClsBank.UpdateContraTrans(IntHeadID, m_PDHeadId, Amount, TransDate) Then gDbTrans.RollBacknRaiseError
End If

'If transaction is cash withdraw & there is cashier window
'then transfer the While Amount cashier window
If Trans = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(m_PDHeadId, m_AccID, _
            TransDate, TransID, Amount) < 1 Then gDbTrans.RollBacknRaiseError
End If

gDbTrans.CommitTrans

'Update the Particulars combo
    'Read to part array
    Dim ParticularsArr() As String
    ReDim ParticularsArr(20)
    Dim I As Integer
    'Read elements of combo to array
    For I = 0 To cmbParticulars.ListCount - 1
        ParticularsArr(I) = cmbParticulars.List(I)
    Next I
    
    'Update last accessed elements
    Call UpdateLastAccessedElements(Trim$(cmbParticulars.Text), ParticularsArr, True)
    
    'Write to
    cmbParticulars.Clear
    For I = 0 To UBound(ParticularsArr)
        If Trim$(ParticularsArr(I)) <> "" Then
            Call WriteToIniFile("Particulars", "Key" & I, ParticularsArr(I), App.Path & "\PDAcc.ini")
            cmbParticulars.AddItem ParticularsArr(I)
        End If
    Next I

    If Not gOnLine Then txtDate.Text = GetIndianDate(TransDate)
    
    AccountTransaction = True

ErrLine:

    Set ClsBank = Nothing
    Set Cashclass = Nothing

End Function

' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetVal(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
For I = 0 To txtData.count - 1
    strTxt = ExtractToken(txtPrompt(I).Tag, "DataSource")
'    Debug.Assert i <> 7
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        GetVal = txtData(I).Text
        Exit For
    End If
Next
End Function

Private Sub LoadAgentNames(cmbAgents As ComboBox)
Dim I As Integer
Dim Perms As wis_Permissions
Dim rst As Recordset

    cmbAgents.Clear
    'perms=
    Perms = perPigmyAgent
    gDbTrans.SqlStmt = "Select * from UserTab WHERE (DELETED = FALSE or DELETED is NULL) "
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
     
    Dim custReg As clsCustReg
    Set custReg = New clsCustReg
    
    For I = 1 To rst.RecordCount
        If Val(rst("Permissions")) And Perms Then
            'CustReg.LoadCustomerInfo (Val(Rst("CustomerID")))
            cmbAgents.AddItem custReg.CustomerName(Val(rst("CustomerId")))
            cmbAgents.ItemData(cmbAgents.newIndex) = Val(rst("UserID"))
        End If
        rst.MoveNext
    Next I
    Set custReg = Nothing
End Sub

Private Function PassBookPageInitialize()
    
    With grd
        .Clear: .Cols = 5
        .Rows = 12: .FixedRows = 1: .FixedCols = 0
        .Row = 0:
        grd.Col = 0: grd.Text = GetResourceString(37): .ColWidth(0) = 1000   ' "Date"
        grd.Col = 1: grd.Text = GetResourceString(39): .ColWidth(1) = 950    '"Particulars"
        grd.Col = 2: grd.Text = GetResourceString(272): .ColWidth(2) = 1000  '"Credit"
        grd.Col = 3: grd.Text = GetResourceString(271): .ColWidth(3) = 1000  '"Debit"
        grd.Col = 4: grd.Text = GetResourceString(42): .ColWidth(4) = 1000    '"Balance"
    End With
    
End Function

Private Function LoadPropSheet() As Boolean
Me.TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True
lblDesc.BorderStyle = 0
lblHeading.BorderStyle = 0
lblOperation.Caption = GetResourceString(54)    '"Operation Mode : <INSERT>"

' Read the data from CAAcc.PRP and load the relevant data.
'
' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\PDAcc_" & gLangOffSet & ".PRP"
If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet <> wis_NoLangOffset Then
        PropFile = App.Path & "\PDAccKan.PRP"
    Else
        PropFile = App.Path & "\PDAcc.PRP"
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
            'MsgBox "Unknown Property type encountered " & "in Property file!", vbCritical
            MsgBox GetResourceString(603) & "in Property file!", vbCritical
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
        .Left = txtPrompt(CtlIndex).Left + txtPrompt(CtlIndex).Width + CTL_MARGIN
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
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, "TextIndex", CStr(cmb.count - 1))
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
                    .Caption = GetResourceString(294)  '"Reset"
                    .Width = 1000
                ElseIf I = 2 Then
                    .Caption = GetResourceString(295)   '"Details..."
                    .Width = 1000
                Else
                    .Caption = "..." '"Specify..."
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
                .BackColor = vbWhite
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
                'txtData(i).Visible = False
            End With

    End Select

    ' Increment the loop count.
    I = I + 1
Loop

ArrangePropSheet

' Get a new account number and display it to accno textbox.
Dim txtIndex As Integer
txtIndex = GetIndex("AccID")
If cmb(0).ListCount > 0 Then
    cmb(0).ListIndex = 0
    txtData(txtIndex).Text = GetNewAccountNumber(cmb(0).ItemData(cmb(0).ListIndex))
End If

' Show the current date wherever necessary.
txtIndex = GetIndex("CreateDate")
txtData(txtIndex).Text = gStrDate

' Set the default updation mode.
m_accUpdatemode = wis_INSERT

'
' Fill up the combobox bound to agent names.
'
Dim cmbIndex As Integer
    ' Find out the textbox bound to agentname.
    txtIndex = GetIndex("AgentName")
    ' Get the combobox index for this text.
    cmbIndex = ExtractToken(txtPrompt(txtIndex).Tag, "TextIndex")
    LoadAgentNames cmb(cmbIndex)
End Function
Private Sub ResetUserInterface()

'Get the Pigmy headId
'get HeadID in the HeadsAccTrans Table(PigmyHeadID)
Dim ClsBank As clsBankAcc

'Get the Pigmy HeadID
If m_PDHeadId = 0 Then
    Set ClsBank = New clsBankAcc
    gDbTrans.BeginTrans
    m_PDHeadId = ClsBank.GetHeadIDCreated(GetResourceString(425), LoadResString(425), parMemberDeposit, 0, wis_PDAcc)
    gDbTrans.CommitTrans
    M_ModuleID = wis_Deposits + wisDeposit_PD
    Set ClsBank = Nothing
    
End If

If m_AccID = 0 And m_CustReg.CustomerID = 0 Then Exit Sub
RaiseEvent AccountChanged(0)
'First the TAB 1
    'Disable the UI if you are unable to load the specified account number
    lblBalance.Caption = ""
    With cmbAccNames
        .BackColor = wisGray: .Enabled = False: .Clear
    End With
    With txtDate
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With txtCheque
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    
    With cmdTransactDate
        .Enabled = False
    End With
    With txtAmount
        .BackColor = wisGray: .Enabled = False: .Text = ""
    End With
    With cmbTrans
        .BackColor = wisGray: .Enabled = False
    End With
    With cmbParticulars
        .BackColor = wisGray: .Enabled = False
    End With
    With Me.rtfNote
        .BackColor = wisGray: .Enabled = False: .Text = GetResourceString(259)   '"< No notes defined >"
        If gLangOffSet Then
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
    With cmdClose
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
    
    'Enable the reset (auto acc no generator button)
    cmd(0).Enabled = True
    'Enable the combo Boxes Modified by shashi 22/2/2000
    For I = 0 To cmb.count - 1
        cmb(I).Enabled = True
        cmb(I).Locked = False
    Next I
        
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
    txtData(txtIndex).Text = "" 'GetNewAccountNumber
    txtData(txtIndex).Locked = False
    cmdTerminate.Enabled = False
    txtIndex = GetIndex("AgentName")
    txtData(txtIndex).Text = ""
    txtData(txtIndex).Locked = False
'The form level variables
    m_accUpdatemode = wis_INSERT
    m_CustReg.NewCustomer
    m_AccID = 0

txtFailAccIDs.Visible = False
lblStatus = ""
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
'
Private Sub SetDescription(Ctl As Control)
' Extract the description title.
lblHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblDesc.Caption = ExtractToken(Ctl.Tag, "Description")
End Sub
'
Private Sub PassBookPageShow()
Dim I As Integer
Dim transType As Byte
'Check if Rec Set has been set
    If m_rstPassBook Is Nothing Then Exit Sub

'Show 10 records or till eof of the page being pointed to
Call PassBookPageInitialize

With grd
    If m_rstPassBook.BOF = True And m_rstPassBook.EOF = True Then
        MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
       Exit Sub
    End If
    
    .Visible = False
    .Row = 1
    .Col = 1: .Text = "Brought Fwd"
    .Col = 4
    'm_rstPassBook.MoveNext
    transType = m_rstPassBook("TransType")
    .Text = FormatCurrency(m_rstPassBook("Balance") - _
        IIf(transType = wWithdraw Or transType = wContraWithdraw, 0, m_rstPassBook("Amount")))
    I = 1
    Do
        I = I + 1
        If .Rows <= I + 1 Then .Rows = I + 1
        transType = m_rstPassBook("TransType")
        .Row = I
        .Col = 0: .Text = FormatField(m_rstPassBook("TransDate"))
        .Col = 1: .Text = FormatField(m_rstPassBook("Particulars"))
        .Col = IIf(transType = wWithdraw Or transType = wContraWithdraw, 2, 3)
        .Text = FormatField(m_rstPassBook("Amount"))
        .Col = 4: .Text = FormatField(m_rstPassBook("Balance"))
        If I <= 10 Then m_rstPassBook.MoveNext Else Exit Do
        If m_rstPassBook.EOF Then Exit Do
    Loop
    .Visible = True
    .Row = 1
End With

If m_rstPassBook.EOF = False And m_rstPassBook.RecordCount = m_rstPassBook.AbsolutePosition Then m_rstPassBook.MoveNext
cmdNextTrans.Enabled = IIf(m_rstPassBook.EOF, False, True)
If m_rstPassBook.EOF And m_rstPassBook.RecordCount > 10 Then
    cmdPrevTrans.Enabled = True
Else
    cmdPrevTrans.Enabled = m_rstPassBook.AbsolutePosition > 10
End If
    
End Sub

Private Sub AgentBookShow()
Dim I As Integer
Dim PigmyComm As Currency
Dim transType As Byte
Dim AgentID As Long
'Check if Recordset has been set
    If m_rstAgent Is Nothing Then Exit Sub
    AgentID = cmbAgentList.ItemData(cmbAgentList.ListIndex)
'Show 10 records or till eof of the page being pointed to
    With grdAgent
        Call AgentGridInitialize
        .Visible = False
        .Row = 1
        .Col = 1: .Text = "Brought Fwd"
        .Col = 4
        
        If m_rstAgent.EOF = True Or m_rstAgent.BOF = True Then .Visible = True: Exit Sub
        'm_rstAgent.MoveNext
        
        transType = m_rstAgent("TransType")
        .Text = FormatCurrency(m_rstAgent("Balance") - _
            IIf(m_rstAgent("TransType") < 0, 0, m_rstAgent("Amount")))
         .Col = 2
        I = 1
        Do
            I = I + 1
            .Row = I
            .Col = 0: .Text = FormatField(m_rstAgent("TransDate"))
            .Col = 1: .Text = FormatField(m_rstAgent("Particulars"))
            
            'Calculate the Pigmy Comission for the Agent make
            PigmyComm = m_rstAgent("amount") * Val(txtPigmyCommission) / 100
            .Col = 2
            .Text = Val(PigmyComm) \ 1
            
            .Col = IIf(FormatField(m_rstAgent("TransType")) < 0, 2, 3)
            .Text = FormatField(m_rstAgent("Amount"))
            .Col = 4: .Text = FormatField(m_rstAgent("Balance"))
            If I < 10 Then m_rstAgent.MoveNext Else Exit Do
            If m_rstAgent.EOF Then Exit Do
        Loop
        .Visible = True
        .Row = 1
    End With
End Sub

Private Function ValidControls() As Boolean
'Prelim check
Dim rst As Recordset
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
If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
    'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If
Dim TransDate As Date
TransDate = GetSysFormatDate(txtDate.Text)
'Check For the Last Transction Date
If DateDiff("d", TransDate, GetPigmyLastTransDate(m_AccID)) > 0 Then
    'MsgBox "transaction date is earlier!", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Function
End If

If cmbTrans.ListIndex = -1 Then
    'MsgBox "Transaction type not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(588), vbExclamation, gAppName & " - Error"
    cmbTrans.SetFocus
    Exit Function
End If

If Not CurrencyValidate(txtAmount.Text, True) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtAmount
    Exit Function
End If
'Check The Transaction date w.r.t to Account CreateDate
gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & m_AccID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If DateDiff("D", rst.Fields("CreateDate"), TransDate) < 0 Then
        'MsgBox "You have specified a transaction date that is earlier than the account creation date !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(568), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    If DateDiff("d", rst.Fields("MaturityDate"), TransDate) > 0 Then
        'If MsgBox("You have specified a transaction date that is later than the maturity date of account !" & vbCrLf & _
            "Do you want to continue ", vbQuestion + vbYesNo + vbDefaultButton2, gAppName & " - Information") = vbNo Then
        If MsgBox(GetResourceString(578) & vbCrLf & _
                GetResourceString(541), vbQuestion + vbYesNo + _
                vbDefaultButton2, gAppName & " - Information") = vbNo Then Exit Function
        
    End If
End If

ValidControls = True

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

Private Sub chkBackLog_Click()
    txtCheque.Enabled = True
    txtCheque.BackColor = vbWhite
    cmdUndo.Caption = GetResourceString(5)  '"&Undo first"
End Sub

Private Sub cmb_Click(Index As Integer)
    
    If Index = 0 Then
        If cmb(0).ListIndex >= 0 Then
            txtData(3).Text = GetNewAccountNumber(cmb(0).ItemData(cmb(0).ListIndex))
        End If
    End If

End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
Dim txtIndex As String
txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" Then
    txtData(Val(txtIndex)).Text = cmb(Index).Text
    txtData(Val(txtIndex)).Visible = True
    cmb(Index).Visible = False
End If

'Generate new account number for this agent here
If Index = 0 Then
    If cmb(0).ListIndex >= 0 Then
        txtData(3).Text = GetNewAccountNumber(cmb(0).ItemData(cmb(0).ListIndex))
    End If
End If


End Sub

Private Sub cmbAgentList_Change()

cmbAgentList.ListIndex = -1
cmbAgentTrans.Enabled = False
cmbAgentTrans.BackColor = wisGray
txtAgentAmount.Enabled = False
txtAgentAmount.BackColor = wisGray
cmbAgentParticulars.Enabled = False
cmbAgentParticulars.BackColor = wisGray
txtAgentCheque.Enabled = True
txtAgentCheque.ForeColor = wisGray
txtAgentDate.Enabled = False
txtAgentDate.BackColor = wisGray

End Sub

Private Sub cmbAgentList_Click()
Dim Index As Integer
Dim rst As Recordset

'if Agent is not selected
If cmbAgentList.ListIndex < 0 Then Exit Sub

Dim AgentID As Integer
Dim ret As Integer

'Initilize the Agent PassBook
Call AgentGridInitialize

'Get the AgentId
With cmbAgentList
    If .ListIndex < 0 Then Exit Sub
    AgentID = .ItemData(.ListIndex)
End With
cmdAgentAccept.Enabled = True
cmdAgentUndo.Enabled = True

'Check for the Existing
gDbTrans.SqlStmt = "SELECT TOP 1 * From AgentTrans " & _
    " WHERE AgentID = " & AgentID & _
    " ORDER BY TransId Desc "

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    Call clearAgentControls
    Exit Sub
End If

AgentID = FormatField(rst(0))
txtAgentDate = gStrDate

RaiseEvent AgentChanged(AgentID)
cmbAgentList.Tag = AgentID

'Load the agent details
If Not AgentLoad(AgentID) Then ResetAgentDetails

End Sub

Private Sub clearAgentControls()

If Not gOnLine Then txtAgentDate.Text = ""
txtAgentAmount.Text = ""

End Sub

Private Sub cmbAgentParticulars_GotFocus()
cmbAgentParticulars.AddItem "By Cash"
End Sub

Private Sub cmbAgents_Click()

'Disable accno if agent not selected
If Val(cmbAgents.Tag) <> GetAgentID Then Call txtAccNo_Change
If cmbAgents.ListIndex < 0 Then txtAccNo.Text = ""

End Sub


Private Sub cmbTrans_Click()
'Disable the Pigmy Commission
chkPigmyComission.Enabled = False

If cmbTrans.ListCount = 0 Then
    'MsgBox "Initialization Error"
    MsgBox GetResourceString(608)
    Exit Sub
End If

If cmbTrans.ListIndex = 0 Then  'A case of deposit
    txtAmount.Text = 0
    chkPigmyComission.Enabled = True
    txtCheque.Enabled = True
    txtCheque.BackColor = vbWhite
    Exit Sub
End If

If cmbTrans.ListIndex = 1 Then  'A case of withdraw
    'txtCheque.Enabled = False
    'txtCheque.BackColor = wisGray
    txtAmount.Text = 0
    Exit Sub
End If

End Sub

'
Private Sub cmbTrans_GotFocus()
 
End Sub

Private Sub cmbTrans_LostFocus()
   
If cmbTrans.ListIndex Then Exit Sub

txtAmount = GetPigmyAmount
       
End Sub

Private Sub cmd_Click(Index As Integer)
Dim txtIndex As String

' Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")

Select Case UCase$(strField)
    'Code Adde By Shashi 20/2/2000
    Case "UserId"
        
    Case "AGENTNAME"
        
    Case "ACCID"
        If m_accUpdatemode = wis_INSERT Then
            'txtData(txtindex).Text = GetNewAccountNumber
            txtIndex = GetIndex("AccID")
            If cmb(0).ListCount > 0 Then
                cmb(0).ListIndex = 0
                txtData(txtIndex).Text = GetNewAccountNumber(cmb(0).ItemData(cmb(0).ListIndex))
            End If
        End If

    Case "ACCNAME"
        m_CustReg.ModuleId = wis_PDAcc
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
    
    Case "MATURITYDATE"
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
        gDbTrans.SqlStmt = "SELECT PDMaster.AccID as [Acc No], " _
                    & "Title + FirstName + Space(1) + Middlename " _
                    & "+ space(1) + LastName as Name, HomeAddress, " _
                    & "OfficeAddress FROM PDMaster, NameTab WHERE " _
                    & "PDMaster.CustomerID = NameTab.CustomerID"
        Dim intIndex As Integer
        intIndex = GetIndex("AccID")
        If txtData(intIndex).Text <> "" And _
                IsNumeric(txtData(intIndex).Text) Then
            gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND " _
                & "AccID <> " & txtData(intIndex).Text
        End If
        Dim Lret As Long
        Dim rst As Recordset
        Lret = gDbTrans.Fetch(rst, adOpenStatic)
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
            MsgBox GetResourceString(562), _
                    vbCritical, wis_MESSAGE_TITLE
            Exit Sub
        End If
        
    Case "NOMINEENAME"
       Set m_frmLookUp = New frmLookUp
        With m_frmLookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
         '   .lvwReport.ColumnHeaders(2).Width = 3750
          '  .lvwReport.ColumnHeaders(3).Width = 3750
            m_CustReg.ModuleId = wis_PDAcc
            m_CustReg.ShowDialog
            txtData(txtIndex).Text = m_CustReg.FullName
            
            .Title = "Select Introducer..."
            .m_SelItem = ""
          '  .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem <> "" Then
                txtData(txtIndex).Text = .lvwReport.SelectedItem.Text
                txtData(txtIndex + 1).Text = .lvwReport.SelectedItem.SubItems(1)
            End If
        End With
End Select
End Sub
'
Private Sub cmdAccept_Click()
If Not ValidControls Then Exit Sub
If Not AccountTransaction() Then Exit Sub

'Reload the account
If Not AccountLoad(m_AccID) Then
    Me.TabStrip2.Tabs(2).Selected = True
    Exit Sub
End If

If txtDate.Enabled Then Call ActivateTextBox(txtDate)
    
TabStrip2.Tabs(2).Selected = True

End Sub
'
Private Sub cmdAddNote_Click()
If m_Notes.ModuleId = 0 Then
    Exit Sub
End If


Call m_Notes.Show
Call m_Notes.DisplayNote(rtfNote)


End Sub

Private Sub cmdAgents_Click()
''frmAgents.Show vbModal
'If cmbAgents.ListIndex <= 0 Or cmbAgents.ListIndex = cmbAgents.ListCount - 1 Then Exit Sub
'    'm_CustReg.CustomerID = cmbAgents.ItemData(cmbAgents.ListIndex)
'    m_CustReg.LoadCustomerInfo (cmbAgents.ItemData(cmbAgents.ListIndex))
'    m_CustReg.ShowDialog
'    If m_CustReg.Modified Then
'        If MsgBox("Do you want Keep Changes Made In Agent Details", _
'                vbYesNo + vbInformation, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
'        m_CustReg.SaveCustomer
'
'    End If
'Call LoadAgentNames

gCurrUser.ShowUserDialog
Call LoadAgentNames(cmbAgents)

' Load the agent names to the combox in property sheet also.
Dim cmbIndex As Integer, txtIndex As Integer
    ' Find out the textbox bound to agentname.
    txtIndex = GetIndex("AgentName")
    ' Get the combobox index for this text.
    cmbIndex = ExtractToken(txtPrompt(txtIndex).Tag, "TextIndex")
    LoadAgentNames cmb(cmbIndex)


End Sub

Private Sub cmdAdvance_Click()
    
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    m_clsRepOption.ShowDialog
End Sub

Private Sub cmdAgentAccept_Click()

Call AgentTransaction

With cmbAgentList
    If .ListIndex < 0 Then Exit Sub
    If Not AgentLoad(.ItemData(.ListIndex)) Then Exit Sub
    cmbAgentTrans.ListIndex = 0
    ActivateDateTextBox txtAgentDate
End With
End Sub

Private Sub ActiveAgentDetails()

With cmbAgentTrans
    .Enabled = True
    .BackColor = vbWhite
    .ListIndex = 0
End With
With txtAgentAmount
    .Enabled = True
    .BackColor = vbWhite
End With
With cmbAgentParticulars
    .Enabled = True
    .BackColor = vbWhite
End With
With txtAgentCheque
    .Enabled = True
    .BackColor = vbWhite
End With
With txtAgentDate
    .Enabled = True
    .BackColor = vbWhite
End With
cmdAgentAccept.Enabled = True
cmdAgentUndo.Enabled = True
ActivateDateTextBox txtAgentDate

End Sub

Private Sub cmdAgentNextTrans_Click()
cmdAgentPrevTrans.Enabled = True
Call AgentBookShow
Exit Sub
'Call GetPassBookNextButton(m_rstAgent, 10, cmdAgentPrevTrans, cmdAgentNextTrans)

If m_rstAgent Is Nothing Then Exit Sub

If m_rstAgent.EOF And m_rstAgent.BOF Then Exit Sub

Dim CurPos As Integer

'Position cursor to start of next page
    If m_rstAgent.EOF Then m_rstAgent.MoveLast
    CurPos = m_rstAgent.AbsolutePosition
    CurPos = 10 - (CurPos Mod 10)
    If m_rstAgent.AbsolutePosition + CurPos >= m_rstAgent.RecordCount Then
        Beep
        Exit Sub
    Else
       ' m_rstPassBook.Move CurPos
    End If

Call AgentBookShow

End Sub

Private Sub cmdAgentNote_Click()

If m_AgentNotes.ModuleId = 0 Then Exit Sub

Call m_AgentNotes.Show
Call m_AgentNotes.DisplayNote(rtfNote)


End Sub

Private Sub cmdAgentPrevTrans_Click()
 

'Drag the Agent details
If m_rstAgent Is Nothing Then
    Exit Sub
End If

 cmdAgentNextTrans.Enabled = True
 Call PassBookPrevButtonClicked(m_rstAgent, 10, cmdAgentPrevTrans)
 Call AgentBookShow
 
Exit Sub
'Call GetPassBookPrevButton(m_rstAgent, 10, cmdAgentPrevTrans, cmdAgentNextTrans)

Dim CurPos As Integer
If m_rstAgent.EOF = True Or m_rstAgent.BOF = True Then Exit Sub
'Position cursor to previous page
    If m_rstAgent.EOF Then
        'm_rstPassBook.MoveFirst
        m_rstAgent.MoveLast
        'm_rstPassBook.MovePrevious
    End If
    
    CurPos = m_rstAgent.AbsolutePosition
    
    CurPos = CurPos - CurPos Mod 10 - 10
    If CurPos < 0 Then
        Beep
        Exit Sub
    Else
        m_rstAgent.MoveFirst
        m_rstAgent.Move (CurPos)
    End If
    
    Call AgentBookShow
End Sub

Private Sub cmdAgentTransactDate_Click()
With Calendar
    .Left = Me.Left + fraAgent.Left + cmdAgentTransactDate.Left
    .Top = Me.Top + fraAgent.Top + cmdAgentTransactDate.Top - .Height / 2
    .selDate = txtAgentDate
    .Show 1
    txtAgentDate = .selDate
End With

End Sub

Private Sub cmdAgentUndo_Click()

Dim AgentID As Integer

With cmbAgentList
    If .ListIndex < 0 Then Exit Sub
    AgentID = .ItemData(.ListIndex)
End With

If Not UndoAgentLastTrans Then Exit Sub

'Undo the last Transaction Made
If Not AgentLoad(AgentID) Then Exit Sub

End Sub

Private Sub cmdAmount_Click()

End Sub

Private Sub cmdApply_Click()
Dim ret As Boolean
ret = True

If Not ret Then
    'MsgBox "Unable to save settings !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(533), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

'cmdApply.Enabled = False
'MsgBox "Settings will only take effect only when you restart this module !", vbInformation, gAppName & " - Message"
MsgBox GetResourceString(537), vbInformation, gAppName & " - Message"
'If Not GetPDInterestChanged(GetAppFormatDate(gStrDate)) Then
   'MsgBox " Unable To Add Interest ", vbCritical, wis_MESSAGE_TITLE
   Exit Sub
'End If


End Sub

Private Sub cmdClose_Click()
Dim ClosedON As String
    
    If Not AccountExists(m_AccID, ClosedON) Then
        'MsgBox "This Account number does not exists", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(525), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    If ClosedON = "" Then
        frmPDClose.AccountId = m_AccID
        
        frmPDClose.Show vbModal
    Else
        Call AccountReopen(m_AccID)
        cmdClose.Caption = GetResourceString(11) '"&Close"
    End If
    Call cmdLoad_Click
    
End Sub

Private Sub cmdFromDate_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtFromDate.Text, "/", True) Then .selDate = txtFromDate.Text
    .Left = Me.Left + Me.fraReports.Left + cmdFromDate.Left - .Width / 2
    .Top = Me.Top + Me.fraReports.Top + cmdFromDate.Top + 2800
    .Show vbModal
    If .selDate <> "" Then txtFromDate.Text = .selDate
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

Dim SetUp As New clsSetup
Dim strModule As String
Dim strValue As String
Dim strDef As String

'strModule = "DEPOSIT" & m_DepositType
strModule = "DEPOSIT" & wisDeposit_PD
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
        'retstr = SetUp.ReadSetupValue("DEPOSIT" & m_DepositType, strKey, "")
        retstr = SetUp.ReadSetupValue("DEPOSIT" & wisDeposit_PD, strKey, "")
        If retstr = "" Then Exit For
        strTo = cmbTo.List(I)
        If Val(Prevstr) <> Val(retstr) Then
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


Private Sub cmdIntPayable_Click()
If Not DateValidate(txtIntPayable.Text, "/", True) Then
'''    MsgBox "Invalid Date Format Specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntPayable
    Exit Sub
End If
Call AddInterestPayableOfPD(txtIntPayable.Text)
End Sub

Private Function AddInterestPayableOfPD(OnIndianDate As String) As Boolean

Dim DimPos As Integer
DimPos = InStr(1, OnIndianDate, "31/3/", vbTextCompare)
If DimPos = 0 Then DimPos = InStr(1, OnIndianDate, "31/03/", vbTextCompare)
    If DimPos = 0 Then
'''    MsgBox "Unable to perform the transactions", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If

   On Error GoTo ErrLine
  'declare the variables necessary
  
Dim transType As wisTransactionTypes
Dim rstMain As Recordset

Dim UserCount As Integer
Dim count As Integer

'Dim BankClass As New clsBankAcc
Dim TransDate As Date
Dim rst As Recordset
TransDate = GetSysFormatDate(OnIndianDate)


'Before Adding check whether he has already added the amount
gDbTrans.SqlStmt = "select *  from PDTrans " & _
            " Where TransDate = #" & TransDate & "#" & _
            " And Particulars ='Interest Payable'"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    MsgBox "Interest Payble already added to the Accounts", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Build The Querry
Screen.MousePointer = vbHourglass

gDbTrans.SqlStmt = " SELECT AccNum,Title +' '+ FirstName +' '+ MiddleName" & _
        " + ' '+LastName As CustName, Balance," & _
        " B.TransID, A.AccId, A.AgentID, CreateDate, MaturityDate," & _
        " TransDate, TransType, RateOfInterest " & _
        " From PDMaster A, PDTrans B, NameTab C Where " & _
        " A.AccId = B.AccId And A.CustomerID = C.CustomerID " & _
        " And (ClosedDate is NULL OR ClosedDate = #1/1/100# )" & _
        " And CreateDate < #" & TransDate & "# And TransID =  " & _
            " (Select Max(TransID) From PDTrans D Where D.AccId = A.AccId)" & _
        " AND Balance <> 0 " & _
        " Order By a.AgentID, val(A.AccNum),A.AccId"

count = gDbTrans.Fetch(rstMain, adOpenStatic)
If count < 1 Then GoTo ExitLine


Dim InterestRate As Currency
Dim LastIntDate As Date
Dim CreateDate As Date
Dim MatDate As Date
Dim Duration As Long
Dim IntAmount As Currency

Dim Balance  As Currency
Dim TransID As Long
Dim AccId As Long
Dim TotalInt As Currency
Dim IntPayable As Currency
Dim TotalIntPayable As Currency

Dim rstPayable As Recordset
Dim rstInt As Recordset

'nOW GET THE interest payble alrady added in the prevous years
gDbTrans.SqlStmt = "Select Balance As Payable, " & _
            " AccId,TransID From PDIntPayable A " & _
            " Where Transid = (SELECT Max(TransID) From PDIntPayable B " & _
                " Where B.AccID = A.AccID )" & _
            " ORDER BY Accid"

If gDbTrans.Fetch(rstPayable, adOpenForwardOnly) < 1 Then Set rstPayable = Nothing
    
gDbTrans.SqlStmt = "Select Balance As Payable," & _
            " TransDate,TransID,AccId From PDIntTrans A " & _
            " Where Transid = (SELECT Max(TransID) From PDIntTrans B" & _
                    " Where B.AccID = A.AccID )" & _
            " ORDER BY Accid"
Set rstInt = Nothing

If gDbTrans.Fetch(rstInt, adOpenForwardOnly) < 1 Then Set rstInt = Nothing

'lblStatus.Caption = "Computing Interests for
lblStatus.Caption = GetResourceString(906) & "  ............"

'Now get the No of pigmy Agent
gDbTrans.SqlStmt = "Select Distinct AgentID From PDMAster"
UserCount = gDbTrans.Fetch(rst, adOpenStatic)


Dim tmpTransID As Long
Dim AccTransID As Long
Dim AgentID As Integer

txtFailAccIDs = ""
Unload frmIntPayble
Load frmIntPayble

With frmIntPayble
    Call .LoadContorls(UserCount + count + 1, 20)
    .lblTitle.Caption = GetResourceString(425) & " " & _
                    GetResourceString(375, 47)
    .PutTotal = True
    .Title(0) = GetResourceString(36)
    .Title(1) = GetResourceString(35)
    .Title(2) = GetResourceString(250, 450)
    .Title(3) = GetResourceString(450)
    .Title(4) = GetResourceString(52, 450)
End With


prg.Min = 0: prg.Max = UserCount + count + 1
UserCount = 0: count = 1

While Not rstMain.EOF
    If Val(FormatField(rstMain("AgentId"))) <> AgentID Then
        If count > 1 Then UserCount = UserCount + 1
        AgentID = FormatField(rstMain("AgentId"))
        gDbTrans.SqlStmt = "SELECT Title +' '+ FirstName +' '+ MiddleName" & _
            " +' '+ LastName as AgentName From NameTab " & _
                " Where CustomerID = (SELECT CustomerID FROM UserTab " & _
                " WHERE UserID = " & AgentID & ")"
        If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine
        With frmIntPayble
            .CustName(count) = FormatField(rst("AgentName"))
            .KeyData(count) = -1
        End With
        count = count + 1
    End If
    AccId = Val(FormatField(rstMain("AccId")))
    AccTransID = Val(FormatField(rstMain("TransID")))
    Balance = CCur(FormatField(rstMain("Balance")))
    CreateDate = rstMain("CreateDate")
    InterestRate = CCur(FormatField(rstMain("RateofInterest")))
    MatDate = rstMain("MaturityDate")
    If DateDiff("d", rstMain("TransDate"), TransDate) < 0 Then AccTransID = 0
    
    If Balance = 0 Then TransID = 0
    
    LastIntDate = FormatField(rstMain("CreateDate"))
    If Not rstInt Is Nothing Then
        rstInt.MoveFirst
        rstInt.Find "AccId = " & AccId
        If Not rstInt.EOF Then
            tmpTransID = rstInt("TransID")
            LastIntDate = rstInt("TransDate")
            If DateDiff("D", LastIntDate, TransDate) < 0 Then AccTransID = 0
            If AccTransID Then _
                AccTransID = IIf(AccTransID > tmpTransID, AccTransID, tmpTransID)
        End If
    End If
    
    Balance = 0
    If Not rstPayable Is Nothing Then
        rstPayable.MoveFirst
        rstPayable.Find "AccId = " & AccId
        If Not rstPayable.EOF Then
            Balance = FormatField(rstPayable("Balance"))
            tmpTransID = rstPayable("TransID")
            If DateDiff("D", rstPayable("TransDate"), TransDate) < 0 Then AccTransID = 0
            If AccTransID Then _
                AccTransID = IIf(AccTransID > tmpTransID, AccTransID, tmpTransID)
        End If
    End If
    
    MatDate = DateAdd("yyyy", -1, TransDate)
    If DateDiff("d", MatDate, LastIntDate) <= 1 Then _
            LastIntDate = DateAdd("d", 1, MatDate)
    
    'Now Get The Date Difference
    Duration = DateDiff("D", LastIntDate, TransDate)
    
    If AccTransID = 0 Then
        Duration = 0
        TransID = 0
    End If
    
    'If InterestRate = 0 Then _
        InterestRate = CCur(GetPDDepositInterest(CInt(Duration), OnIndianDate))
    If InterestRate = 0 Then _
        InterestRate = CCur(GetDepositInterestRate(wis_PDAcc, rstMain("CreateDate"), rstMain("MaturityDate")))

    If InterestRate <= 0 Then InterestRate = 4
    
    IntAmount = (((InterestRate / 100) * rstMain("Balance") * 1) / 12) \ 1
    
    IntAmount = IntAmount \ 1
    'If IntAmount = 0 Then GoTo NextDeposit
    'Check for the prevously added interest payble of this account
    IntPayable = Balance
    If AccTransID Then TransID = AccTransID + 1
    With frmIntPayble
        .Balance(count) = IntPayable
        .AccNum(count) = rstMain("AccNum")
        .CustName(count) = FormatField(rstMain("CustName"))
        .Amount(count) = IntAmount
        .Total(count) = IntPayable + IntAmount
        .KeyData(count) = TransID
        TotalIntPayable = TotalIntPayable + IntPayable
        TotalInt = TotalInt + IntAmount
    End With
Repeat:
    
NextDeposit:
    rstMain.MoveNext: count = count + 1
    prg.Value = count
     Debug.Assert count <> 215
Wend

With frmIntPayble
    count = .grd.Rows - 1
    .CustName(count) = GetResourceString(274) & " " & _
        GetResourceString(450, 346) 'Total Interest Payble
    .Balance(count) = TotalIntPayable
    TotalIntPayable = TotalIntPayable + IntPayable
    .Amount(count) = TotalInt
End With

prg.Value = 0
Screen.MousePointer = vbDefault

Me.Refresh
frmIntPayble.ShowForm
Me.Refresh

If frmIntPayble.grd.Rows < 3 Then GoTo ExitLine


Screen.MousePointer = vbHourglass

Dim MaxCount As Integer
Dim CurrUserID As Long


MaxCount = frmIntPayble.grd.Rows - 1

lblStatus.Caption = GetResourceString(907)

TotalIntPayable = 0 'TotalIntPayable + IntPayable
gDbTrans.BeginTrans

TotalInt = 0
rstMain.MoveFirst
For count = 1 To MaxCount
    
    AccId = rstMain("accId")
    With frmIntPayble
        IntAmount = Val(.Amount(count))
        TransID = Val(.KeyData(count))
        Balance = Val(.Total(count))
        'If .CustName(Count) = "" Then AccId = 0
        If .KeyData(count) = -1 Then AccId = 0
        
    End With
    
    TotalInt = TotalInt + IntAmount
    If TransID > 0 And IntAmount > 0 Then
        TotalIntPayable = TotalIntPayable + IntAmount
        'With draw the Amount from yhe Interest Account
        transType = wContraWithdraw
        gDbTrans.SqlStmt = "INSERT INTO PDIntTrans (AccID, TransID," & _
            " TransDate,Amount, TransType," & _
            " Balance, Particulars,UserID ) VALUES " & _
            " (" & AccId & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            transType & "," & _
            Balance + IntAmount & "," & _
            "'Interest Payable' ," & _
            CurrUserID & ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            GoTo ErrLine
        End If
        
        'Deposit the Amount to the Interest payable Account
        transType = wContraDeposit
        gDbTrans.SqlStmt = "INSERT INTO PDIntPayable (AccID, TransID," & _
            " TransDate,Amount, TransType," & _
            " Balance, Particulars,UserID ) VALUES " & _
            " (" & AccId & "," & _
            TransID & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            transType & "," & _
            Balance + IntAmount & "," & _
            "'Interest Payable' ," & _
            CurrUserID & ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            GoTo ErrLine
        End If
    ElseIf TransID = 0 And AccId Then
        txtFailAccIDs = txtFailAccIDs & AccId & ", "
    End If
    prg.Value = count
   
    If AccId Then rstMain.MoveNext
    If rstMain.EOF Then Exit For
Next count


Dim bankClass As clsBankAcc
Dim PayableHeadID As Long
Dim IntHeadID As Long
Set bankClass = New clsBankAcc

Dim headName As String
'Noew ge the Ledger head id of the Pigmy deposit payble
headName = GetResourceString(425, 450)        'PIgmy INterest provision
PayableHeadID = bankClass.GetHeadIDCreated(headName, LoadResString(425) & " " & LoadResString(450), parDepositIntProv, 0, wis_PDAcc)

headName = GetResourceString(425, 375, 47)       'PIgmy Payble INterest
IntHeadID = bankClass.GetHeadIDCreated(headName, _
        LoadResString(425) & " " & LoadResString(375) & " " & LoadResString(47), parMemDepIntPaid, 0, wis_PDAcc)

'Now Make the same transaction to the ledger heads
Call bankClass.UpdateContraTrans(IntHeadID, PayableHeadID, TotalIntPayable, TransDate)

gDbTrans.CommitTrans
Set bankClass = Nothing


lblStatus = ""
If Len(txtFailAccIDs) > 0 Then
    lblStatus = GetResourceString(544) & " " & _
    GetResourceString(36) & GetResourceString(92)
    txtFailAccIDs.Visible = True
End If

'MsgBox " Interest payble  added success fully", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(274, 450) & " " & _
    GetResourceString(637), vbInformation, wis_MESSAGE_TITLE

AddInterestPayableOfPD = True
GoTo ExitLine

ErrLine:
MsgBox "Error In PDAccount --Interest payble", vbCritical, wis_MESSAGE_TITLE
'Resume

ExitLine:

Screen.MousePointer = vbDefault
Set bankClass = Nothing


End Function

Private Sub cmdLoad_Click()
Dim AgentID As Integer

With cmbAgents
    If .ListIndex < 0 Then
        MsgBox "Please Choose the Agent Name", vbInformation, wis_MESSAGE_TITLE
        cmbAgents.SetFocus
        Exit Sub
    Else
        AgentID = .ItemData(.ListIndex)
    End If
    
End With


Dim rst As Recordset
Dim LoanID As Long
Dim ret As Integer

'First get the Account Id from the Date base
gDbTrans.SqlStmt = "SELECT AccNum,ACCID,AgentID,loanid " & _
    " From PDMASTER WHERE " & _
    " AccNum = " & AddQuotes(Trim$(txtAccNo.Text), True) & _
    " And AgentID = " & AgentID
    
ret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If ret < 1 Then
    'MsgBox "THis Account number does not exists !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

Dim AccId As Long
AccId = FormatField(rst("AccId"))
LoanID = FormatField(rst("LoanId"))


If Not AccountLoad(AccId) Then
    ActivateTextBox txtAccNo
    Exit Sub
End If

Me.TabStrip2.Tabs(IIf(m_Notes.NoteCount > 0, 1, 2)).Selected = True

End Sub

Private Sub cmdNewAgent_Click()
gCurrUser.ShowUserDialog
End Sub

Private Sub cmdNextTrans_Click()

If m_rstPassBook Is Nothing Then
    Exit Sub
End If

Dim CurPos As Integer

'Position cursor to start of next page
    If m_rstPassBook.EOF Then m_rstPassBook.MoveLast
    
    CurPos = m_rstPassBook.AbsolutePosition
    CurPos = 10 - (CurPos Mod 10)
    If CurPos = 10 Then CurPos = 0
    If m_rstPassBook.AbsolutePosition + CurPos >= m_rstPassBook.RecordCount Then
        Beep
        Exit Sub
    Else
       ' m_rstPassBook.Move CurPos
    End If

m_rstPassBook.MoveNext
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

Unload Me
End Sub

Private Sub cmdPhoto_Click()
If Not m_CustReg Is Nothing Then
    frmPhoto.setAccNo (m_CustReg.CustomerID)
        If (m_CustReg.CustomerID > 0) Then frmPhoto.Show vbModal
End If
End Sub

Private Sub cmdPrevTrans_Click()

If m_rstPassBook Is Nothing Then
    Exit Sub
End If

Dim CurPos As Integer
cmdNextTrans.Enabled = True
Call PassBookPrevButtonClicked(m_rstPassBook, 10, cmdPrevTrans)
Call PassBookPageShow
Exit Sub
'Position cursor to previous page
    If m_rstPassBook.EOF Then
        'm_rstPassBook.MoveFirst
        m_rstPassBook.MoveLast
        'm_rstPassBook.MovePrevious
        CurPos = m_rstPassBook.AbsolutePosition
        CurPos = CurPos - CurPos Mod 10 - 10
    Else
        CurPos = m_rstPassBook.AbsolutePosition - 10 - 10
    End If
    
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
    If Not AccountSave Then
        Exit Sub
    End If

'Reload the account details once saved

    Dim AccNo As Long
    Dim AgentID As Integer
    Dim AccNum As String
    AgentID = GetIndex("AgentName")
    'Get AgetnId
    For AccNo = 0 To cmb.count - 1
        If AgentID = CInt(ExtractToken(cmb(AccNo).Tag, "TextIndex")) Then
            AgentID = cmb(AccNo).ItemData(cmb(AccNo).ListIndex)
            Exit For
        End If
    Next AccNo
    
    Dim rst As Recordset
    
    AccNum = Trim$(GetVal("AccID"))
    txtAccNo.Text = AccNum
    'First get the Account Id from the Date base
    gDbTrans.SqlStmt = "SELECT AccNum,ACCID,AgentID From PDMASTER WHERE " & _
        " AccNum = " & AddQuotes(Trim$(txtAccNo.Text), True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number for this agent does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(550), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

    If Not AccountLoad(FormatField(rst("AccID"))) Then
        'MsgBox "Error loading account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(526), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

End Sub

Private Sub cmdTerminate_Click()
Dim I As Integer
Dim strField As String
Dim ret As Integer
Dim rst As Recordset
'Prelim check
    If m_AccID = 0 Then
        'MsgBox "No account loaded !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(523), vbCritical, gAppName & " - Error"
        Exit Sub
    End If

'Check if account number exists in data base
    gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & m_AccID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
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
            MsgBox GetResourceString(664), vbExclamation, gAppName & " - Error"
            Exit Sub
        End If

        Exit Sub
    Else
        'Check if there are any transactions
        gDbTrans.SqlStmt = "Select TOP 1 * from PDTrans where " & _
                    " AccID = " & m_AccID & " order by TransID desc"
        ret = gDbTrans.Fetch(rst, adOpenStatic)
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
                'MsgBox "This account has a balance of Rs. " & FormatField(rst("Balance")) & " and thus cannot be closed !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(549) & FormatField(rst.Fields("Balance")) & GetResourceString(655), vbExclamation, gAppName & " - Error"
                Exit Sub
            End If
            'If MsgBox("Are you sure you want to close this account ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
            If MsgBox(GetResourceString(656), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
                Exit Sub
            End If
        End If
        
        'Close this account now
        If Not AccountClose() Then
            Exit Sub
        End If
        'MsgBox "Account closed successfully !", vbInformation, gAppName & " - Message"
        MsgBox GetResourceString(657), vbInformation, gAppName & " - Message"
        'Reload the account
        If Not AccountLoad(m_AccID) Then
            'MsgBox "Unable to reload the account !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(664), vbExclamation, gAppName & " - Error"
            Exit Sub
        End If
        Exit Sub
    End If
    
     
End Sub

Private Function AccountReopen(AccId As Long) As Boolean

Dim rst As Recordset
Dim ClosedDate As String

'Check if account number exists in data base
gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "Specified account number does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If

ClosedDate = FormatField(rst.Fields("ClosedDate"))

'Opening of this account will undo the depoist Refunded
'First
'If MsgBox("This will undo the Amount refunded and Charges/Interest " & vbCrLf & vbCrLf & _
      "Do You want to continue ?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
If MsgBox(GetResourceString(592) & vbCrLf & vbCrLf & _
      GetResourceString(541), vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
    Exit Function
End If

gDbTrans.SqlStmt = "Select Top 1 TransId,TransDate," & _
            " Amount,TransType From PDTrans " & _
            " Where AccId = " & m_AccID & _
            " Order By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

Dim TransID As Integer
Dim transType As wisTransactionTypes
Dim TransDate As Date

Dim Amount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency

TransID = rst("TransID")
TransDate = rst("TransDate")
Amount = FormatField(rst("Amount"))
transType = FormatField(rst("TransType"))

If transType = wContraDeposit Or transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & m_PDHeadId & _
            " And AccId = " & m_AccID & " And TransID = " & TransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
                AccountReopen = True
        Set ContraClass = Nothing
        
        gDbTrans.BeginTrans
        'Now make the necessary changes in PDMaster
        gDbTrans.SqlStmt = "UpDate PDMaster set ClosedDate = NULL " & _
                " where AccId = " & m_AccID
        Call gDbTrans.SQLExecute
        gDbTrans.CommitTrans
        
        Exit Function
    End If
End If

'Now Get the Interest Amount
gDbTrans.SqlStmt = "Select Top 1 TransId,TransDate," & _
            " Amount From PDIntTrans " & _
            " Where AccId = " & m_AccID & _
            " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    Dim IntHeadID As Long
    Dim headName As String
    headName = GetResourceString(425, 487)
    IntHeadID = GetIndexHeadID(headName)
    If IntHeadID = 0 Then IntHeadID = GetHeadID(headName, parMemDepIntPaid)
    IntAmount = FormatField(rst("Amount"))
End If
'Now Get the payable Amount if any
gDbTrans.SqlStmt = "Select Top 1 TransId,TransDate,AMount " & _
            " From PDIntPayable " & _
            " Where AccId = " & m_AccID & " And TransID = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    Dim PayableHeadID As Long
    'HeadName = GetResourceString(425,450)
    headName = GetResourceString(425, 375, 47)
    PayableHeadID = GetHeadID(headName, parDepositIntProv)
    PayableAmount = FormatField(rst("Amount"))
End If


'First Remove the Amount Refunded
Dim InTrans As Boolean
gDbTrans.BeginTrans
InTrans = True

'Now make the necessary changes in PDMaster
gDbTrans.SqlStmt = "UpDate PDMaster set ClosedDate = NULL " & _
        " where AccId = " & m_AccID

If Not gDbTrans.SQLExecute Then
    'MsgBox "unable to Reopen the Account", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(536), vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
End If

gDbTrans.CommitTrans
AccountReopen = True
Exit Function

' If While Closing this A/c if any Misc Amount collected that has to return(Undo)
Dim bankClass As clsBankAcc

If transType = wContraDeposit Or transType = wContraWithdraw Then
    MsgBox "Unable to reopen this account", vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
    Exit Function
End If

Set bankClass = New clsBankAcc
If PayableAmount Then _
    If Not bankClass.UndoCashWithdrawls(PayableHeadID, PayableAmount, _
        TransDate) Then GoTo ExitLine
If IntAmount Then _
    If Not bankClass.UndoCashWithdrawls(IntHeadID, IntAmount, _
         TransDate) Then GoTo ExitLine
If Amount Then _
    If Not bankClass.UndoCashWithdrawls(m_PDHeadId, Amount, _
        TransDate) Then GoTo ExitLine

Set bankClass = Nothing

gDbTrans.CommitTrans
AccountReopen = True

ExitLine:
If InTrans Then gDbTrans.RollBack

End Function

Private Sub cmdToDate_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtToDate.Text, "/", True) Then .selDate = txtToDate.Text
    .Left = Me.Left + Me.fraReports.Left + cmdToDate.Left - .Width / 2
    .Top = Me.Top + Me.fraReports.Top + cmdToDate.Top + 2800
    .Show vbModal
    If .selDate <> "" Then txtToDate.Text = .selDate
End With

End Sub

Private Sub cmdTransactDate_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtDate.Text, "/", True) Then .selDate = txtDate.Text
    .Left = Me.Left + Me.fraTransact.Left + cmdTransactDate.Left - .Width / 2
    .Top = Me.Top + Me.fraTransact.Top + cmdTransactDate.Top - 100
    .Show vbModal
    If .selDate <> "" Then txtDate.Text = .selDate
End With

End Sub

Private Sub cmdUndo_Click()

If Not AccountUndoLastTransaction() Then
    Call cmdLoad_Click
    Exit Sub
End If

If Not AccountLoad(m_AccID) Then
    'MsgBox "Unable to undo transaction !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(609), vbCritical, gAppName & " - Error"
    Exit Sub
End If
Me.TabStrip2.Tabs(2).Selected = True
End Sub


Private Sub cmdUndoPayable_Click()

If Not DateValidate(txtIntPayable.Text, "/", True) Then
    MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If UndoInterestPayableOfPD(txtIntPayable.Text) Then
    MsgBox "Interest payble removed", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

End Sub

Private Sub cmdView_Click()
MousePointer = vbHourglass
'First check the dates specified
If txtFromDate.Enabled And Not DateValidate(txtFromDate.Text, "/", True) Then
    'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtFromDate
    Me.MousePointer = vbDefault
    Exit Sub
End If
If txtToDate.Enabled Then
    If Not DateValidate(txtToDate.Text, "/", True) Then
        'MsgBox "Please specify from date in DD/mm/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtToDate
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    If txtFromDate.Enabled Then
    If WisDateDiff(txtFromDate.Text, txtToDate.Text) < 0 Then
        'MsgBox "TO date is earlier that the specified FROM date!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtToDate
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    End If
End If

Dim ReportType As wis_PDReports

If optAgentTrans Then ReportType = repPDAgentTrans
If optClosed Then ReportType = repPDAccClose
If optDepGLedger Then ReportType = repPDLedger
If optDepositBalance Then ReportType = repPDBalance
If optMature Then ReportType = repPDMat
If optSubDayBook Then ReportType = repPDDayBook
If optSubCashBook Then ReportType = repPDCashBook
If optMonthly Then ReportType = repPDMonTrans
If optOpened Then ReportType = repPDAccOpen
If optMonthlyBalance Then ReportType = repPDMonBal
If optDailyTransaction Then ReportType = repPDDailyTrans

Dim ShowAgentName As Boolean

ShowAgentName = False
If chkAgentName.Enabled And chkAgentName = vbChecked Then ShowAgentName = True

If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
If cmbRepAgent.ListIndex < 0 Then cmbRepAgent.ListIndex = 0

RaiseEvent ShowReport(ShowAgentName, ReportType, IIf(optName, wisByName, wisByAccountNo), _
             IIf(txtFromDate.Enabled, txtFromDate, ""), IIf(txtToDate.Enabled, txtToDate, ""), _
             m_clsRepOption, cmbRepAgent.ItemData(cmbRepAgent.ListIndex))
            
MousePointer = vbDefault
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

'Copied from  sb Account
'Modified By shashi to Pd Account on 21/2/2000
Public Function AccountExists(AccId As Long, Optional ClosedON As String) As Boolean

Dim ret As Integer
Dim rst As Recordset

'Query Database
gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & AccId
ret = gDbTrans.Fetch(rst, adOpenStatic)
If ret <= 0 Then Exit Function

If ret > 1 Then  'Screwed case
    'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
    Exit Function
End If
    
'Check the closed status
If Not IsMissing(ClosedON) Then ClosedON = FormatField(rst.Fields("ClosedDate"))
    

AccountExists = True

End Function

Public Function AccountLoad(ByVal AccId As Long) As Boolean

Dim rstMaster As Recordset
Dim rstTemp As Recordset

Dim ClosedDate As String
Dim ret As Integer
Dim JointHolders() As String
Dim I As Integer
Dim AgentID As Long

Call ResetUserInterface

'Check if account number is valid
If AccId <= 0 Then GoTo DisableUserInterface
   
'Check if account number exists
    If Not AccountExists(AccId) Then
        'MsgBox "Account number for this agent does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(550), vbExclamation, gAppName & " - Error"
        GoTo DisableUserInterface
    End If

'Query data base
    Dim strAgents As String
    
    gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & AccId
        
    If gDbTrans.Fetch(rstMaster, adOpenForwardOnly) < 1 Then GoTo DisableUserInterface
    AgentID = FormatField(rstMaster("AgentID"))
    
'Load the Name details
    If Not m_CustReg.LoadCustomerInfo(FormatField(rstMaster("CustomerID"))) Then
        'MsgBox "Unable to load customer information !", vbCritical, gAppName & " - Error"
        MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
        GoTo DisableUserInterface
    End If
    
'Get the transaction details of this account holder

    gDbTrans.SqlStmt = "Select * from PDTrans where AccID = " & AccId & _
            " ORDER BY TransDate,TransID"
    ret = gDbTrans.Fetch(m_rstPassBook, adOpenStatic)
    
    If ret < 0 Then
        GoTo DisableUserInterface
    ElseIf ret > 0 Then
        Dim BalanceAmount As Currency
        m_rstPassBook.MoveLast
        BalanceAmount = m_rstPassBook("Balance")
        Call InitPassBook(m_rstPassBook, 10, cmdPrevTrans)
        With lblBalance
            .FontBold = True
            .ForeColor = vbBlue
            .Caption = GetResourceString(42) & " : " & FormatCurrency(BalanceAmount) '"Balance:  Rs. "
            If ClosedDate <> "" Then .Caption = GetResourceString(282) & ": " & ClosedDate: .ForeColor = vbRed '.Caption & " - Account Closed":
            .Alignment = vbRightJustify
        End With
        'Position to first record of last page
        With m_rstPassBook
            Dim transid1 As Long
            transid1 = m_rstPassBook("TransID")
            If .RecordCount > 10 Then
                If transid1 Mod 10 = 0 Then transid1 = transid1 - 10
                .Find "TransiD > " & transid1 - transid1 Mod 10
                'If .RecordCount = 10 Then .Move -1 Else .Move -1 * ((.AbsolutePosition Mod 10) - 1)
                'If .RecordCount = 11 Then .Move -1 Else .Move -1 * ((.AbsolutePosition Mod 10) - 1)
            Else
                .MoveFirst
            End If
        End With
        cmdUndo.Enabled = gCurrUser.IsAdmin
    Else
        Set m_rstPassBook = Nothing
        PassBookPageInitialize
        cmdUndo.Enabled = False
    End If
    
'Assign to some module level variables
    m_AccID = AccId
    m_accUpdatemode = wis_UPDATE
    m_AccClosed = IIf(FormatField(rstMaster("ClosedDate")) <> "", True, False)

'Load account to the User Interface
    'TAB 1
    ClosedDate = FormatField(rstMaster("ClosedDate"))

    With Me
        
        With .cmbAccNames
            .Enabled = True: .BackColor = vbWhite: .Clear
            .AddItem m_CustReg.FullName
            Call GetStringArray(FormatField(rstMaster("JointHolder")), JointHolders, ";")

            For I = 0 To UBound(JointHolders) - 1
                .AddItem JointHolders(I)
            Next I
            .ListIndex = 0
        End With

        'Set the agent Name
        With cmbAgents
            For I = 0 To .ListCount - 1
                If AgentID = .ItemData(I) Then
                    .ListIndex = I
                    Exit For
                End If
            Next
        End With
        

        With .txtDate
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            If .Text = "" Then .Text = gStrDate
        End With

        With cmdTransactDate
            .Enabled = True
            If gOnLine Then .Enabled = False
        End With

        With .cmbTrans
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .ListIndex = -1 'IIf(ClosedDate = "", 0, -1)
        End With
        
        With .cmbParticulars
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .ListIndex = IIf(ClosedDate = "" And .ListCount, 0, -1)
        End With
        
        With .txtAmount
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Text = FormatField(rstMaster("PigmyAmount"))
        End With
        
        With .txtCheque
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            .Text = ""
        End With
        
        cmdPrevTrans.Enabled = IIf(ClosedDate = "", True, False)
        cmdNextTrans.Enabled = IIf(ClosedDate = "", True, False)
        
        With .rtfNote
            .BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
            .Enabled = IIf(ClosedDate = "", True, False)
            Call m_Notes.LoadNotes(wis_PDAcc, AccId)
        End With
        
        Call m_Notes.DisplayNote(.rtfNote)
        
        .cmdAddNote.Enabled = IIf(ClosedDate = "", True, False)
        .cmdAccept.Enabled = IIf(ClosedDate = "", True, False)
        If ClosedDate = "" Then
            .cmdClose.Enabled = True
            .cmdClose.Caption = GetResourceString(11) '"&Close"
        Else
            .cmdClose.Enabled = True
            .cmdClose.Caption = GetResourceString(313) '"Re&open"
        End If

    End With
    
    Call PassBookPageShow
       
    'TAB 2
    'Update labels and other buttons
    With Me
        lblOperation.Caption = GetResourceString(56) '"Operation Mode : <UPDATE>"
        cmdTerminate.Caption = IIf(ClosedDate = "", "&Terminate", "&Reopen")
        cmdTerminate.Enabled = True
    'mallikpatil@usa.net
        Dim NomineeInfo() As String
        If Not IsNull(rstMaster("Nominee")) Then
            GetStringArray FormatField(rstMaster("Nominee")), NomineeInfo(), ";"
        End If
        If UBound(NomineeInfo) < 2 Then
            ReDim NomineeInfo(2)
            NomineeInfo(0) = " "
            NomineeInfo(1) = " "
            NomineeInfo(2) = " "
        End If
        Dim strField As String
        For I = 0 To txtPrompt.count - 1
            ' Read the bound field of this control.
           ' On Error Resume Next
            strField = ExtractToken(txtPrompt(I).Tag, "DataSource")
            If strField <> "" Then
                With txtData(I)
                    Select Case UCase$(strField)
                        Case "AGENTNAME"
                            Dim cmbCount As Integer, count As Integer
                    'Load the Agent Name & Select the Listindex for Agents Combo box
                            For cmbCount = 0 To cmb.count - 1
                                If I = Val(ExtractToken(cmb(cmbCount).Tag, "TextIndex")) Then
                                  For count = 0 To cmb(cmbCount).ListCount - 1
                                    If AgentID = cmb(cmbCount).ItemData(count) Then
                                        .Text = cmb(cmbCount).List(count)
                                        .Locked = True
                                        cmb(cmbCount).ListIndex = count
                                        cmb(cmbCount).Locked = False
                                    End If
                                  Next count
                                End If
                            Next cmbCount
                        Case "ACCID"
                            .Text = rstMaster("AccNum")
                            .Locked = True
                        Case "ACCNAME"
                            .Text = m_CustReg.FullName
                        Case "NOMINEENAME"
                            .Text = NomineeInfo(0)
                        Case "NOMINEEAGE"
                            .Text = NomineeInfo(1)
                        Case "NOMINEERELATION"
                            .Text = NomineeInfo(2)
                        Case "JOINTHOLDER"
                            .Text = FormatField(rstMaster("JointHolder"))
                        Case "INTRODUCERID"
                            .Text = IIf(FormatField(rstMaster("Introduced")) = "0", "", FormatField(rstMaster("Introduced")))
                        Case "INTRODUCERNAME"
                            .Text = AccountName(Val(FormatField(rstMaster("Introduced"))))
                        Case "LEDGERNO"
                            .Text = FormatField(rstMaster("LedgerNo"))
                        Case "FOLIONO"
                            .Text = FormatField(rstMaster("FolioNO"))
                        Case "CREATEDATE"
                            .Text = FormatField(rstMaster("CreateDate"))
                        Case "PIGMYTYPE"
                            .Text = FormatField(rstMaster("PigmyType"))
                        Case "PIGMYAMOUNT"
                            .Text = FormatField(rstMaster("PigmyAmount"))
                             txtAmount.Text = FormatField(rstMaster("PigmyAmount"))
                        Case "MATURITYDATE"
                            .Text = FormatField(rstMaster("MaturityDate"))
                        Case "ACCGROUP"
                            gDbTrans.SqlStmt = "SELECT GroupName FROM AccountGroup WHERE " & _
                                    "AccGroupID = " & FormatField(rstMaster("AccGroupId"))
                            If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
                            .Text = FormatField(rstTemp("GroupName"))
                        Case "NOTIFYONMATURITY"
                            .Text = FormatField(rstMaster("NotifyOnMaturity"))
                        Case "RATEOFINTEREST"
                            .Text = FormatField(rstMaster("RateOfInterest"))
                        Case "NOTIFY"
                            .Text = FormatField(rstMaster("NotifyOnMaturity"))
                            
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
        .cmd(1).Enabled = False
    End With

AccountLoad = True
cmdPhoto.Enabled = Len(gImagePath)
RaiseEvent AccountChanged(m_AccID)
'cmbAgents.Locked = True
Exit Function

DisableUserInterface:
    Call ResetUserInterface
    
Exit Function
    
ErrLine:
'MsgBox "Account Load:" & vbCrLf & "     Error Loading account", vbCritical, gAppName & " - Error"
MsgBox GetResourceString(521) & vbCrLf & GetResourceString(551), vbCritical, gAppName & " - Error"
End Function

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
'Centre the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

cmdAgentPrint.Picture = LoadResPicture(120, vbResBitmap)

'set kannada caption
    Call SetKannadaCaption
 
'Intialize the custreg Calss
    m_CustReg.ModuleId = wis_PDAcc

'Fill up transaction Types
    With cmbTrans
        .Clear
        .AddItem GetResourceString(271)
        If Not gCashier Or (gCashier And (M_UserPermission = perBankAdmin)) Then _
            .AddItem GetResourceString(272)
        If M_UserPermission = perBankAdmin Then
            .AddItem GetResourceString(273)
            .AddItem GetResourceString(274)
        End If
    End With
     
     cmbAgentTrans.AddItem GetResourceString(271) 'Deposit
         
     
'Fill up particulars with default values from PDAcc.INI
    Dim Particulars As String
    Dim I As Integer
    Do
        Particulars = ReadFromIniFile("Particulars", _
                "Key" & I, gAppPath & "\PDAcc.INI")
        If Trim$(Particulars) <> "" Then
            cmbParticulars.AddItem Particulars
        End If
        I = I + 1
    Loop Until Trim$(Particulars) = ""

'Load ICONS
    cmdAddNote.Picture = LoadResPicture(103, vbResBitmap)

'Adjust the Grid for Pass book
    With grd
        .Rows = 11
        .Cols = 5
        .FixedCols = 1
        .Row = 0
        
        .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1000 ' "Date"
        .Col = 1: .Text = GetResourceString(39): .ColWidth(1) = 1000 '"Particulars"
        .Col = 2: .Text = GetResourceString(275): .ColWidth(2) = 1000 '"Cheque"
        .Col = 3: .Text = GetResourceString(276): .ColWidth(3) = 1000 '"Debit"
        .Col = 4: .Text = GetResourceString(42): .ColWidth(4) = 1000 '"Balance"
    End With
    Me.txtCheque.Visible = True

Call LoadPropSheet


'Load the Setup values
    Dim SetUp As New clsSetup

'Reset the User Interface
    Call ResetUserInterface

'Load properties
    With M_setUp
        'txtLoanPercent.Text = .ReadSetupValue("PDAcc", "MaxLoanPercent", Val(txtLoanPercent.Text))
        txtPigmyCommission.Text = .ReadSetupValue("PDAcc", "PigmyCommission", "03")
    End With
    Call LoadInterestRates
    Set M_setUp = Nothing
    

'Load Agent Name
Call LoadAgentNames(cmbAgents)
Dim cmbIndex As Byte
cmbIndex = GetIndex("AccGroup")
cmbIndex = ExtractToken(txtPrompt(cmbIndex).Tag, "TextIndex")
Call LoadAccountGroups(cmb(cmbIndex))


'Set Report Frame
optDepositBalance.Value = True
Call optDepositBalance_Click

Me.TabStrip2.Tabs(1).Selected = True

Me.fraInstructions.ZOrder 0
Screen.MousePointer = vbDefault
txtToDate = gStrDate
txtDate = txtToDate
txtAgentDate = txtToDate
If gOnLine Then
    txtAgentDate.Locked = True
    cmdAgentTransactDate.Enabled = False
    txtDate.Locked = True
    cmdTransactDate.Enabled = False
End If

cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)
cmdAgentPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdAgentNextTrans.Picture = LoadResPicture(102, vbResIcon)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
gWindowHandle = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

' Report form.
If Not m_frmLookUp Is Nothing Then
    Unload m_frmLookUp
    Set m_frmLookUp = Nothing
End If

' Notes object.
Set m_Notes = Nothing

' Customer Registration object.
Set m_CustReg = Nothing
'""(Me.hwnd, False)
gWindowHandle = 0
RaiseEvent WindowClosed
End Sub

Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)

Call PrintBetweenDates(wis_PDAcc, m_AccID, StartIndiandate, EndIndianDate)
Exit Sub

Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer
'1. Fetch last print row from PDmaster table.
'First get the last printed txnID From the PDMaster
SqlStr = "SELECT  LastPrintRow,LastPrintID From PDMaster WHERE AccId = " & m_AccID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
'lastPrintRow = IIf(IsNull(metaRst("LastPrintrow")), 0, metaRst("LastPrintrow"))
lastPrintRow = 0

'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From PDTrans WHERE AccId = " & m_AccID
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

If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE_A4 - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If


'3. Getting matching records for passbook printing
    SqlStr = "SELECT * From PDTrans WHERE AccId = " & m_AccID & _
        " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
        " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#"
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If


'Printer.PaperSize = 9
Printer.Font = gFontName
'Printer.Font.Size = 12 'gFontSize
'Printer.Font = "Courier New"
Printer.FONTSIZE = 12
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
            .ColText(3) = " "
        Else
            .ColText(2) = " "
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
SqlStr = "UPDATE PDMaster set LastPrintRow = " & curPrintRow - 1 & _
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
'First get the last printed transaID From the PDMaster

SqlStr = "SELECT  LastPrintID, LastPrintRow From PDMaster WHERE AccId = " & m_AccID
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(metaRst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
'get position of the last Trans Printed in
lastPrintId = IIf(IsNull(metaRst("LastPrintID")), 0, metaRst("LastPrintId"))
'get position of the last print point
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If IsNull(metaRst("LastPrintRow")) And lastPrintId = 1 Then lastPrintId = 0


' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From PDTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


 'Print the first page of passbook, if newPassbook option is chosen.
If bNewPassbook Then
    clsPrint.printPassbookPage wis_PDAcc, m_AccID
    SqlStr = "UPDATE PDMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
            ", LastPrintRow = 0 Where Accid = " & m_AccID
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
    Else
        gDbTrans.CommitTrans
    End If
End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
If (rst(0) = 0) Then
    
    Dim iRet As Integer
    iRet = MsgBox("There are no transactions available for printing." & vbCrLf & _
    "Do you want to reset the printing?", vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE)
    If (iRet = vbYes) Then
            'Now Update the Last Print Id to the master
        SqlStr = "UPDATE PDMaster set LastPrintId = " & TransID & _
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

'If (rst(0) > 15) Then
'    MsgBox "plz insert new page."
'End If

' Fetch records for txns that have been created after lasttxnId.
SqlStr = "SELECT * From PDTrans WHERE AccId = " & m_AccID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If


'Printer.PaperSize = 9
Printer.Font = gFontName
'Printer.Font.Size = 12 'gFontSize
'Printer.Font = "Courier New"
Printer.FONTSIZE = 12
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 4
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .COlHeader(1) = GetResourceString(275) 'Particulars
    '.ColWidth(2) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(2) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(3) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(4) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    
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
        '.ColWidth(5) = 15

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
            .ColText(3) = ""
        Else
            .ColText(2) = ""
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
SqlStr = "UPDATE PDMaster set LastPrintId = " & TransID & _
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

Private Sub optAgentTrans_Click()
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(False, False)
    
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True
    Call ChkAgentNameValue(optAgentTrans)
    
End Sub

Private Sub optClosed_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(True, True)
        
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True

    Call ChkAgentNameValue(optClosed)

End Sub

Private Sub optDailyTransaction_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(False, False)
    
    txtFromDate.Enabled = True
    txtFromDate.BackColor = wisWhite
    cmdFromDate.Enabled = True
        
    Call ChkAgentNameValue(optDailyTransaction)
    chkAgentName.Enabled = True
End Sub

Private Sub optDays_Click()
    'Load Cmb From
    cmbFrom.Clear
    cmbFrom.AddItem ("0")
    cmbFrom.AddItem ("30")
    cmbFrom.AddItem ("90")
    cmbFrom.AddItem ("180")
    cmbFrom.AddItem ("365")
    cmbFrom.AddItem ("365")
    cmbFrom.AddItem ("730")
    cmbFrom.AddItem ("1090")
    
    cmbTo.Clear
    cmbTo.AddItem ("30")
    cmbTo.AddItem ("90")
    cmbTo.AddItem ("180")
    cmbTo.AddItem ("365")
    cmbTo.AddItem ("365")
    cmbTo.AddItem ("730")
    cmbTo.AddItem ("Above 1090")
    

End Sub

Private Sub optDepGLedger_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(False, False)
        
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True
    chkAgentName.Value = 0
    chkAgentName.Enabled = False
    Call ChkAgentNameValue(optDepGLedger)
    
End Sub

Private Sub optDepositBalance_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(True, True)
        
    txtFromDate.Enabled = False
    txtFromDate.BackColor = wisGray
    cmdFromDate.Enabled = False
        
    Call ChkAgentNameValue(optDepositBalance)

End Sub

Private Sub optSubCashBook_Click()

Call optSUbDayBook_Click

End Sub

Private Sub optSUbDayBook_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(True, True)
        
    With txtFromDate
        .Enabled = True
        .BackColor = wisWhite
    End With
    cmdFromDate.Enabled = True
   
    Call ChkAgentNameValue(optSubDayBook)
    
End Sub

Private Sub optMature_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(True, True)
        
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True
    
    Call ChkAgentNameValue(optMature)
    
End Sub

Private Sub optMonthly_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(False, False)
        
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True
    chkAgentName.Value = 0
    chkAgentName.Enabled = False
    Call ChkAgentNameValue(optMonthly)
    

End Sub

Private Sub optMonthlyBalance_Click()
    
    'Enable Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(False, False)
    
    txtFromDate.Enabled = True
    txtFromDate.BackColor = wisWhite
    cmdFromDate.Enabled = True
        
    Call ChkAgentNameValue(optMonthlyBalance)
    chkAgentName.Enabled = False
End Sub

Private Sub optOpened_Click()
    
    'Eanble Place,Caste Group And Amount Range Controls
    Call SetCastePlaceAmountRange(True, False)
    
    txtFromDate.Enabled = True
    txtFromDate.BackColor = vbWhite
    cmdFromDate.Enabled = True
    
    Call ChkAgentNameValue(optOpened)
    
End Sub

Private Sub TabAgentStrip2_Click()
    
    If TabAgentStrip2.SelectedItem.Index = 1 Then
        fraAgentInstructions.Visible = True
        fraAgentInstructions.ZOrder 0
        fraAgentPassbook.Visible = False
    Else
    'End If
    'If TabAgentStrip2.SelectedItem.Index = 2 Then
        fraAgentInstructions.Visible = False
        fraAgentPassbook.Visible = True
        fraAgentPassbook.ZOrder 0
    End If

End Sub

Private Sub TabStrip_Click()

Dim strKey As String
strKey = TabStrip.SelectedItem.Key

fraAgent.Visible = False
fraNew.Visible = False
fraProps.Visible = False
fraReports.Visible = False
fraTransact.Visible = False

Select Case UCase(strKey)
    Case "AGENTTRANS"
        fraAgent.Visible = True
        fraAgent.ZOrder 0
        cmdAgentAccept.Default = True
        
    Case "ADDMODIFY"
        fraNew.Visible = True
        fraNew.ZOrder 0
        cmdSave.Default = True
        
    Case "PROPERTIES"
        fraProps.Visible = True
        fraProps.ZOrder 0
        'txtData(1).SetFocus
        cmdSave.Default = True
        
    Case "TRANSACTIONS"
        fraTransact.Visible = True
        fraTransact.ZOrder 0
        cmdAccept.Default = True
        
    Case "REPORTS"
        fraReports.Visible = True
        fraReports.ZOrder 0
        cmdView.Default = True
        
End Select

End Sub

Private Sub TabStrip2_Click()
    If TabStrip2.SelectedItem.Index = 1 Then
        fraInstructions.Visible = True
        fraInstructions.ZOrder 0
        fraPassBook.Visible = False
    Else
        fraInstructions.Visible = False
        fraPassBook.Visible = True
        fraPassBook.ZOrder
    End If
End Sub

Private Sub txtAccNo_Change()
cmdLoad.Enabled = IIf(Trim$(txtAccNo.Text) <> "", True, False)

If m_AccID Then Call ResetUserInterface

End Sub

Private Sub txtAgentAmount_GotFocus()
With txtAgentAmount
    .SelStart = 1
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtAgentAmount_LostFocus()
txtAgentCheque.Text = 333
End Sub


Private Sub txtAgentCheque_LostFocus()
cmbAgentParticulars.Text = "By Cash"
End Sub


Private Sub txtAgentDate_GotFocus()
TabAgentStrip2.Tabs(2).Selected = True
End Sub


Private Sub txtAmount_GotFocus()
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub txtAmount_LostFocus()
txtCheque.Text = 222
End Sub

Private Sub txtCheque_LostFocus()
cmbParticulars.Text = "bycash"
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
ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    ' Get the cmdbutton index.
    On Error Resume Next
    If TextIndex <> "" Then
        If cmb(Val(TextIndex)).Visible Then Exit Sub
        cmb(Val(TextIndex)).Visible = True: cmb(Val(TextIndex)).ZOrder 0
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
Dim rst As Recordset

' If the item is IntroducerID, validate the
' ID and name.
strDatSrc = ExtractToken(txtPrompt(Index).Tag, "DataSource")
If StrComp(strDatSrc, "IntroducerID", vbTextCompare) = 0 Then
    ' Check if any data is found in this text.
    If Trim$(txtData(Index).Text) <> "" Then
        gDbTrans.SqlStmt = "SELECT AccID, Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM PDMaster, " _
                & "NameTab WHERE PDMaster.AccID = " & Val(txtData(Index).Text) _
                & " AND PDMaster.CustomerID = NameTab.CustomerID"
        Lret = gDbTrans.Fetch(rst, adOpenStatic)
        txtIndex = GetIndex("IntroducerName")
        If Lret > 0 Then
            txtData(txtIndex).Text = FormatField(rst("Name"))
        Else
            txtData(txtIndex).Text = ""
        End If
    Else
        txtIndex = GetIndex("IntroducerName")
        txtData(txtIndex).Text = ""
    End If
End If

'Set The Rate Of Interest
If StrComp(strDatSrc, "MaturityDate", vbTextCompare) = 0 Then
Dim Days As Long
Dim Dt1 As String
Dim InterestRate As Single

    'Check For ValidDate
  If DateValidate(txtData(Index).Text, "/", True) And DateValidate(txtData(GetIndex("CreateDate")).Text, "/", True) Then
    txtIndex = GetIndex("CreateDate")
    Dt1 = txtData(txtIndex).Text
    On Error Resume Next
    Days = WisDateDiff(Dt1, Trim$(txtData(Index).Text))
    InterestRate = GetPDDepositInterest(Days, txtData(txtIndex).Text)
    txtIndex = GetIndex("RateOfInterest")
    txtData(txtIndex).Text = InterestRate
    On Error GoTo 0
    Exit Sub
  End If
End If
End Sub

Private Sub txtDate_GotFocus()
TabStrip2.Tabs(2).Selected = True
End Sub

Private Sub txtIntPayable_GotFocus()
With txtIntPayable
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtPigmyCommission_GotFocus()
With txtPigmyCommission
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtPrompt_GotFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlue
End Sub

Private Sub txtPrompt_LostFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlack
End Sub

Private Sub VScroll1_Change()
' Move the picSlider.
picSlider.Top = -VScroll1.Value
End Sub

Public Property Get Nominee() As String
' The Nominee string consists of
' Nominee_name;Nominee_age;Nominee_Relation.

Nominee = GetVal("Nomineename") & ";" _
        & GetVal("NomineeAge") & ";" _
        & GetVal("NomineeRelation")

End Property

Private Sub SetCastePlaceAmountRange(EnablePlaceCaste As Boolean, EnableAmountRange As Boolean)
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = EnablePlaceCaste
        .EnableAmountRange = EnableAmountRange
    End With
    
End Sub
