VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCustReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer details form"
   ClientHeight    =   7290
   ClientLeft      =   2175
   ClientTop       =   1740
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pic 
      Height          =   5900
      Index           =   4
      Left            =   195
      ScaleHeight     =   5835
      ScaleWidth      =   5640
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   500
      Width           =   5700
      Begin VB.TextBox txtKycAccNum2 
         Height          =   345
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   109
         Top             =   5310
         Width           =   1665
      End
      Begin VB.TextBox txtKycIfsc2 
         Height          =   345
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   108
         Top             =   5310
         Width           =   1425
      End
      Begin VB.TextBox txtKycBankName2 
         Height          =   345
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   107
         Top             =   4800
         Width           =   3345
      End
      Begin VB.TextBox txtKycAccNum1 
         Height          =   345
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   106
         Top             =   4320
         Width           =   1665
      End
      Begin VB.TextBox txtKycId2 
         Height          =   345
         Left            =   2160
         TabIndex        =   105
         Top             =   3255
         Width           =   3345
      End
      Begin VB.ComboBox cmbKycID2 
         Height          =   315
         Left            =   2160
         TabIndex        =   104
         Text            =   "Combo1"
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtKYCIfsc1 
         Height          =   345
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   103
         Top             =   4322
         Width           =   1425
      End
      Begin VB.TextBox txtKycBankName1 
         Height          =   345
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   102
         Top             =   3816
         Width           =   3345
      End
      Begin VB.TextBox txtKYCPhone 
         Height          =   345
         Left            =   2160
         TabIndex        =   101
         Top             =   1286
         Width           =   3345
      End
      Begin VB.TextBox txtKycId1 
         Height          =   345
         Left            =   2160
         TabIndex        =   100
         Top             =   2298
         Width           =   3345
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   60
         TabIndex        =   99
         Top             =   630
         Width           =   5415
      End
      Begin VB.TextBox txtKYCCustname 
         Height          =   345
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   780
         Width           =   3345
      End
      Begin VB.ComboBox cmbKycId1 
         Height          =   315
         Left            =   2160
         TabIndex        =   97
         Text            =   "Combo1"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblKycAccNum2 
         Caption         =   "IFSC and Account #"
         Height          =   315
         Left            =   90
         TabIndex        =   120
         Top             =   5370
         Width           =   1860
      End
      Begin VB.Label lblKycBank2 
         Caption         =   "Ext Bank Name"
         Height          =   315
         Left            =   90
         TabIndex        =   119
         Top             =   4860
         Width           =   1740
      End
      Begin VB.Label lblKycID2 
         Caption         =   "Identification Number"
         Height          =   315
         Left            =   90
         TabIndex        =   118
         Top             =   3315
         WhatsThisHelpID =   10
         Width           =   1740
      End
      Begin VB.Label lblKycIdType2 
         Caption         =   "Identification Type"
         Height          =   315
         Left            =   90
         TabIndex        =   117
         Top             =   2805
         WhatsThisHelpID =   10
         Width           =   1740
      End
      Begin VB.Label lblKycAccNum1 
         Caption         =   "IFSC and Account #"
         Height          =   315
         Left            =   90
         TabIndex        =   116
         Top             =   4380
         Width           =   1980
      End
      Begin VB.Label lblKycBank1 
         Caption         =   "Ext Bank Name"
         Height          =   315
         Left            =   90
         TabIndex        =   115
         Top             =   3870
         Width           =   1980
      End
      Begin VB.Label lblKycPhoneType1 
         Caption         =   "Secondary Phone"
         Height          =   315
         Left            =   90
         TabIndex        =   114
         Top             =   1350
         WhatsThisHelpID =   10
         Width           =   1860
      End
      Begin VB.Label lblKycID1 
         Caption         =   "Identification Number"
         Height          =   315
         Left            =   90
         TabIndex        =   113
         Top             =   2355
         WhatsThisHelpID =   10
         Width           =   1740
      End
      Begin VB.Label lblKycIdType1 
         Caption         =   "Identification Type"
         Height          =   315
         Left            =   90
         TabIndex        =   112
         Top             =   1845
         WhatsThisHelpID =   10
         Width           =   1740
      End
      Begin VB.Label lblKycTitle 
         AutoSize        =   -1  'True
         Caption         =   "Enter the KYC Details  for <customer>"
         Height          =   195
         Left            =   870
         TabIndex        =   111
         Top             =   300
         Width           =   2670
      End
      Begin VB.Image Image5 
         Height          =   465
         Left            =   210
         Picture         =   "custReg.frx":0000
         Stretch         =   -1  'True
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lblKycPhone 
         Caption         =   "Customer Name :"
         Height          =   315
         Left            =   90
         TabIndex        =   110
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   1860
      End
   End
   Begin VB.PictureBox pic 
      Height          =   6015
      Index           =   5
      Left            =   120
      ScaleHeight     =   5955
      ScaleWidth      =   5595
      TabIndex        =   121
      Top             =   480
      Width           =   5655
      Begin VB.Frame fraPhoto 
         Height          =   5085
         Left            =   0
         TabIndex        =   122
         Top             =   480
         Width           =   5685
         Begin VB.CommandButton cmdImgDel 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1275
            TabIndex        =   130
            Top             =   3600
            Width           =   420
         End
         Begin VB.CommandButton cmdImgPrev 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   129
            Top             =   3600
            Width           =   375
         End
         Begin VB.CommandButton cmdImgAdd 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   720
            TabIndex        =   128
            Top             =   3600
            Width           =   420
         End
         Begin VB.CommandButton cmdImgNext 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   127
            Top             =   3600
            Width           =   375
         End
         Begin VB.CommandButton cmdSgnDel 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4230
            TabIndex        =   126
            Top             =   3600
            Width           =   420
         End
         Begin VB.CommandButton cmdSgnNext 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4920
            TabIndex        =   125
            Top             =   3600
            Width           =   375
         End
         Begin VB.CommandButton cmdSgnAdd 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3690
            TabIndex        =   124
            Top             =   3600
            Width           =   420
         End
         Begin VB.CommandButton cmdSgnPrev 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3000
            TabIndex        =   123
            Top             =   3600
            Width           =   375
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   2040
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image picphoto 
            Height          =   2800
            Left            =   120
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2400
         End
         Begin VB.Image picSign 
            Height          =   2805
            Left            =   3000
            Stretch         =   -1  'True
            Top             =   600
            Width           =   2400
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00404040&
            BorderWidth     =   2
            X1              =   2750
            X2              =   2750
            Y1              =   120
            Y2              =   4020
         End
         Begin VB.Label lblImgDate 
            Caption         =   "Date: 2/4/2013"
            Height          =   315
            Left            =   255
            TabIndex        =   134
            Top             =   4185
            Width           =   1440
         End
         Begin VB.Label lblImgCount 
            Alignment       =   2  'Center
            Caption         =   "Photo: 0/0"
            Height          =   315
            Left            =   480
            TabIndex        =   133
            Top             =   4440
            Width           =   1290
         End
         Begin VB.Label lblSgnDate 
            Caption         =   "Date: 13/06/2013"
            Height          =   345
            Left            =   2850
            TabIndex        =   132
            Top             =   4170
            Width           =   1560
         End
         Begin VB.Label lblSgnCount 
            Alignment       =   2  'Center
            Caption         =   "Signature: 0/0"
            Height          =   345
            Left            =   2880
            TabIndex        =   131
            Top             =   4440
            Width           =   1410
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   400
      Left            =   150
      TabIndex        =   61
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "&Lookup..."
      Height          =   400
      Left            =   1710
      TabIndex        =   0
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   3270
      TabIndex        =   49
      Tag             =   "40"
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4830
      TabIndex        =   60
      Top             =   6780
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Height          =   5900
      Index           =   2
      Left            =   195
      ScaleHeight     =   5835
      ScaleWidth      =   5640
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   5700
      Begin VB.TextBox txtHomeTaluka 
         Height          =   345
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   36
         Top             =   2319
         Width           =   3100
      End
      Begin VB.TextBox txtEnglishName 
         Height          =   345
         Left            =   1995
         TabIndex        =   92
         Top             =   5280
         Width           =   3100
      End
      Begin VB.CommandButton cmdPlace 
         Caption         =   "..."
         Height          =   300
         Left            =   4740
         TabIndex        =   33
         Top             =   1800
         Width           =   315
      End
      Begin VB.TextBox txtEMail 
         Height          =   345
         Left            =   1995
         MaxLength       =   30
         TabIndex        =   47
         Top             =   4784
         Width           =   3100
      End
      Begin VB.TextBox txtHomePhone 
         Height          =   345
         Left            =   1995
         MaxLength       =   30
         TabIndex        =   45
         Top             =   4291
         Width           =   3100
      End
      Begin VB.TextBox txtHomeCity 
         Height          =   345
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   40
         Top             =   1800
         Width           =   2505
      End
      Begin VB.TextBox txtHomePin 
         Height          =   345
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   43
         Top             =   3798
         Width           =   3100
      End
      Begin VB.TextBox txtHomeStreet 
         Height          =   345
         Left            =   1995
         TabIndex        =   31
         Top             =   1363
         Width           =   3100
      End
      Begin VB.TextBox txtHomeNo 
         Height          =   345
         Left            =   1995
         TabIndex        =   29
         Top             =   870
         Width           =   3100
      End
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   120
         TabIndex        =   63
         Top             =   630
         Width           =   4995
      End
      Begin VB.TextBox txtHomeDistrict 
         Height          =   345
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   38
         Top             =   2812
         Width           =   3100
      End
      Begin VB.TextBox txtHomeState 
         Height          =   345
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   41
         Top             =   3305
         Width           =   3100
      End
      Begin VB.ComboBox cmbHomeCity 
         Height          =   315
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1856
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblTaluka 
         Caption         =   "Taluka : "
         Height          =   315
         Left            =   210
         TabIndex        =   35
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblEnglishName 
         AutoSize        =   -1  'True
         Caption         =   "Name in English"
         Height          =   315
         Left            =   210
         TabIndex        =   93
         Top             =   5355
         Width           =   1260
      End
      Begin VB.Label lblHomeEmail 
         Caption         =   "eMail :"
         Height          =   315
         Left            =   210
         TabIndex        =   46
         Top             =   4860
         Width           =   1395
      End
      Begin VB.Label lblHomePhone 
         Caption         =   "Phone :"
         Height          =   315
         Left            =   210
         TabIndex        =   44
         Top             =   4365
         Width           =   1215
      End
      Begin VB.Label lblDistrict 
         Caption         =   "District : "
         Height          =   315
         Left            =   210
         TabIndex        =   37
         Top             =   2895
         Width           =   1335
      End
      Begin VB.Label lblState 
         Caption         =   "State : "
         Height          =   315
         Left            =   210
         TabIndex        =   39
         Top             =   3390
         Width           =   1065
      End
      Begin VB.Label lblCity 
         Caption         =   "City / Village :"
         Height          =   315
         Left            =   210
         TabIndex        =   32
         Top             =   1905
         Width           =   1365
      End
      Begin VB.Label lblPINCode 
         Caption         =   "PIN Code :"
         Height          =   315
         Left            =   210
         TabIndex        =   42
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label lblStreetAddr 
         Caption         =   "Street address :"
         Height          =   315
         Left            =   210
         TabIndex        =   30
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label lblHouseNo 
         Caption         =   "House Number :"
         Height          =   315
         Left            =   210
         TabIndex        =   28
         Top             =   930
         WhatsThisHelpID =   10
         Width           =   1305
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   210
         Picture         =   "custReg.frx":2A42
         Stretch         =   -1  'True
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lblHomeTitle 
         AutoSize        =   -1  'True
         Caption         =   "Enter home related information of the client here."
         Height          =   195
         Left            =   1110
         TabIndex        =   62
         Top             =   300
         Width           =   3900
      End
   End
   Begin VB.PictureBox pic 
      Height          =   5900
      Index           =   1
      Left            =   195
      ScaleHeight     =   5835
      ScaleWidth      =   5640
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   600
      Width           =   5700
      Begin VB.Frame fraCust 
         Height          =   5900
         Index           =   0
         Left            =   60
         TabIndex        =   82
         Top             =   0
         Width           =   5700
         Begin VB.ComboBox cmbFarmerType 
            Height          =   315
            Left            =   1995
            TabIndex        =   24
            Top             =   5400
            Width           =   3045
         End
         Begin VB.CommandButton cmdFarmerType 
            Caption         =   "..."
            Height          =   330
            Left            =   5145
            TabIndex        =   25
            Top             =   5430
            Width           =   315
         End
         Begin VB.CommandButton cmdCaste 
            Caption         =   "..."
            Height          =   300
            Left            =   5190
            TabIndex        =   22
            Top             =   5010
            Width           =   315
         End
         Begin VB.OptionButton optUI 
            Caption         =   "Change UI"
            Height          =   315
            Index           =   0
            Left            =   3300
            TabIndex        =   89
            Top             =   780
            Width           =   1485
         End
         Begin VB.TextBox txtGuardian 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   345
            Left            =   1995
            TabIndex        =   12
            Top             =   3510
            Width           =   2955
         End
         Begin VB.TextBox txtFirstName 
            Height          =   345
            Left            =   1995
            TabIndex        =   4
            Top             =   1305
            Width           =   2955
         End
         Begin VB.TextBox txtMiddleName 
            Height          =   345
            Left            =   1995
            TabIndex        =   6
            Top             =   1740
            Width           =   2955
         End
         Begin VB.TextBox txtLastName 
            Height          =   345
            Left            =   1995
            TabIndex        =   8
            Top             =   2160
            Width           =   2955
         End
         Begin VB.ComboBox cmbTitle 
            Height          =   315
            ItemData        =   "custReg.frx":333B
            Left            =   1995
            List            =   "custReg.frx":3351
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   780
            Width           =   915
         End
         Begin VB.TextBox txtProfession 
            Height          =   330
            Left            =   1995
            TabIndex        =   18
            Top             =   4500
            Width           =   2955
         End
         Begin VB.TextBox txtCaste 
            Height          =   345
            Left            =   1995
            TabIndex        =   20
            Top             =   4965
            Width           =   3015
         End
         Begin VB.TextBox txtDOB 
            Height          =   345
            Left            =   1995
            TabIndex        =   10
            Top             =   2595
            Width           =   1110
         End
         Begin VB.ComboBox cmbGender 
            Height          =   315
            ItemData        =   "custReg.frx":337D
            Left            =   1995
            List            =   "custReg.frx":338A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3105
            Width           =   1215
         End
         Begin VB.ComboBox cmbMaritalStatus 
            Height          =   315
            ItemData        =   "custReg.frx":33A9
            Left            =   1995
            List            =   "custReg.frx":33AB
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3990
            Width           =   2955
         End
         Begin VB.ComboBox cmbCaste 
            Height          =   315
            Left            =   1995
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   5040
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.Label lblLastName 
            AutoSize        =   -1  'True
            Caption         =   "Last :"
            Height          =   315
            Left            =   195
            TabIndex        =   7
            Top             =   2250
            Width           =   600
         End
         Begin VB.Label lblFarmerType 
            AutoSize        =   -1  'True
            Caption         =   "Cust Type :"
            Height          =   315
            Left            =   195
            TabIndex        =   23
            Top             =   5415
            Width           =   810
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "(DD/MM/YYYY)"
            Height          =   195
            Left            =   3210
            TabIndex        =   87
            Top             =   2670
            Width           =   1170
         End
         Begin VB.Label lblGuardian 
            AutoSize        =   -1  'True
            Caption         =   "Guardian :"
            Enabled         =   0   'False
            Height          =   315
            Left            =   195
            TabIndex        =   11
            Top             =   3570
            Width           =   735
         End
         Begin VB.Label lblFirstName 
            AutoSize        =   -1  'True
            Caption         =   "First :"
            Height          =   315
            Left            =   195
            TabIndex        =   3
            Top             =   1330
            Width           =   600
         End
         Begin VB.Label lblMiddleName 
            AutoSize        =   -1  'True
            Caption         =   "Middle :"
            Height          =   195
            Left            =   195
            TabIndex        =   5
            Top             =   1785
            Width           =   555
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Title :"
            Height          =   315
            Left            =   195
            TabIndex        =   1
            Top             =   870
            Width           =   450
         End
         Begin VB.Label lblProfession 
            AutoSize        =   -1  'True
            Caption         =   "Profession :"
            Height          =   315
            Left            =   195
            TabIndex        =   17
            Top             =   4490
            Width           =   1155
         End
         Begin VB.Label lblCaste 
            AutoSize        =   -1  'True
            Caption         =   "Caste :"
            Height          =   315
            Left            =   195
            TabIndex        =   19
            Top             =   4950
            Width           =   1035
         End
         Begin VB.Label lbldoB 
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth :"
            Height          =   315
            Left            =   195
            TabIndex        =   9
            Top             =   2710
            Width           =   975
         End
         Begin VB.Label lblGender 
            AutoSize        =   -1  'True
            Caption         =   "Gender :"
            Height          =   255
            Left            =   165
            TabIndex        =   13
            Top             =   3170
            Width           =   975
         End
         Begin VB.Label lblMarital 
            Caption         =   "Marital Status :"
            Height          =   315
            Left            =   195
            TabIndex        =   15
            Top             =   4030
            Width           =   1140
         End
         Begin VB.Image Image1 
            Height          =   405
            Left            =   360
            Picture         =   "custReg.frx":33AD
            Stretch         =   -1  'True
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblPersonTitle 
            AutoSize        =   -1  'True
            Caption         =   "Enter personal information about the client here."
            Height          =   195
            Left            =   1230
            TabIndex        =   86
            Top             =   330
            Width           =   3375
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   300
            X2              =   4530
            Y1              =   690
            Y2              =   690
         End
      End
      Begin VB.Frame fraCust 
         Height          =   5900
         Index           =   1
         Left            =   30
         TabIndex        =   91
         Top             =   0
         Width           =   5700
         Begin VB.ComboBox cmbCustType 
            Height          =   315
            Left            =   2100
            TabIndex        =   85
            Top             =   3630
            Width           =   3195
         End
         Begin VB.OptionButton optUI 
            Caption         =   "Change UI"
            Height          =   315
            Index           =   1
            Left            =   4020
            TabIndex        =   90
            Top             =   900
            Width           =   1485
         End
         Begin VB.TextBox txtEstd 
            Height          =   345
            Left            =   2100
            TabIndex        =   26
            Top             =   3060
            Width           =   1500
         End
         Begin VB.ComboBox cmbInstTitle 
            Height          =   315
            ItemData        =   "custReg.frx":354F
            Left            =   2100
            List            =   "custReg.frx":3565
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   930
            Width           =   915
         End
         Begin VB.TextBox txtInstHead 
            Height          =   345
            Left            =   2100
            TabIndex        =   81
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox txtInstName 
            Height          =   945
            Left            =   2100
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   1470
            Width           =   3285
         End
         Begin VB.Label lblCustType 
            Caption         =   "Customer Type"
            Height          =   285
            Left            =   150
            TabIndex        =   83
            Top             =   3660
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   210
            X2              =   5400
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblEstDate 
            AutoSize        =   -1  'True
            Caption         =   "Date of Estd:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   3090
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Title :"
            Height          =   195
            Left            =   150
            TabIndex        =   73
            Top             =   1020
            Width           =   750
         End
         Begin VB.Label lblInstHead 
            Caption         =   "Head of the institutions"
            Height          =   225
            Left            =   90
            TabIndex        =   79
            Top             =   2610
            Width           =   1395
         End
         Begin VB.Label lblInstName 
            Caption         =   "Name Of the Initution"
            Height          =   525
            Left            =   150
            TabIndex        =   76
            Top             =   1500
            Width           =   1845
         End
         Begin VB.Image Image4 
            Height          =   405
            Left            =   180
            Picture         =   "custReg.frx":3591
            Stretch         =   -1  'True
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblInstitution 
            AutoSize        =   -1  'True
            Caption         =   "Enter information about the institution here."
            Height          =   350
            Left            =   1050
            TabIndex        =   88
            Top             =   360
            Width           =   3015
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6585
      Left            =   120
      TabIndex        =   50
      Top             =   90
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11615
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Personal"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Home"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Office"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "KYC Docs"
            Key             =   "KYC"
            Object.Tag             =   ""
            Object.ToolTipText     =   "KYC documents"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Photo"
            Key             =   "photo"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Photo and signature of Customer"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic 
      Height          =   5900
      Index           =   3
      Left            =   195
      ScaleHeight     =   5835
      ScaleWidth      =   5640
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   570
      Width           =   5700
      Begin VB.TextBox txtOffTaluka 
         Height          =   345
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   94
         Top             =   3310
         Width           =   3465
      End
      Begin VB.TextBox txtOffState 
         Height          =   345
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   56
         Top             =   4322
         Width           =   3465
      End
      Begin VB.TextBox txtOffDistrict 
         Height          =   345
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   59
         Top             =   3816
         Width           =   3465
      End
      Begin VB.TextBox txtOffPin 
         Height          =   345
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   57
         Top             =   4828
         Width           =   3465
      End
      Begin VB.TextBox txtOffCity 
         Height          =   345
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   55
         Top             =   2804
         Width           =   3465
      End
      Begin VB.TextBox txtOffPhone 
         Height          =   345
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   58
         Top             =   5340
         Width           =   3465
      End
      Begin VB.TextBox txtJobTitle 
         Height          =   345
         Left            =   1920
         TabIndex        =   52
         Top             =   1286
         Width           =   3465
      End
      Begin VB.TextBox txtOffStreet 
         Height          =   345
         Left            =   1920
         TabIndex        =   54
         Top             =   2298
         Width           =   3465
      End
      Begin VB.TextBox txtOfficeNo 
         Height          =   345
         Left            =   1920
         TabIndex        =   53
         Top             =   1792
         Width           =   3465
      End
      Begin VB.Frame Frame5 
         Height          =   45
         Left            =   60
         TabIndex        =   65
         Top             =   630
         Width           =   4515
      End
      Begin VB.TextBox txtCompanyName 
         Height          =   345
         Left            =   1920
         TabIndex        =   51
         Top             =   780
         Width           =   3465
      End
      Begin VB.Label lblOffTaluka 
         Caption         =   "Taluka : "
         Height          =   315
         Left            =   210
         TabIndex        =   95
         Top             =   3370
         Width           =   1500
      End
      Begin VB.Label lblOffPIN 
         Caption         =   "PIN Code :"
         Height          =   315
         Left            =   210
         TabIndex        =   75
         Top             =   4888
         Width           =   1500
      End
      Begin VB.Label lblOffCity 
         Caption         =   "City : "
         Height          =   315
         Left            =   210
         TabIndex        =   70
         Top             =   2864
         Width           =   1500
      End
      Begin VB.Label lblOffState 
         Caption         =   "State : "
         Height          =   315
         Left            =   210
         TabIndex        =   72
         Top             =   4382
         Width           =   1500
      End
      Begin VB.Label lblOffDist 
         Caption         =   "District : "
         Height          =   315
         Left            =   210
         TabIndex        =   71
         Top             =   3876
         Width           =   1500
      End
      Begin VB.Label lblOffPhone 
         Caption         =   "Phone :"
         Height          =   315
         Left            =   210
         TabIndex        =   77
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label lblJob 
         Caption         =   "Job title :"
         Height          =   315
         Left            =   210
         TabIndex        =   67
         Top             =   1346
         WhatsThisHelpID =   10
         Width           =   1500
      End
      Begin VB.Label lblOffAddr 
         Caption         =   "Street Address :"
         Height          =   315
         Left            =   210
         TabIndex        =   69
         Top             =   2358
         WhatsThisHelpID =   10
         Width           =   1500
      End
      Begin VB.Label lblOffNo 
         Caption         =   "Office Number :"
         Height          =   315
         Left            =   210
         TabIndex        =   68
         Top             =   1852
         WhatsThisHelpID =   10
         Width           =   1500
      End
      Begin VB.Label lblOffTitle 
         AutoSize        =   -1  'True
         Caption         =   "Enter office information of the client here."
         Height          =   195
         Left            =   870
         TabIndex        =   64
         Top             =   300
         Width           =   3675
      End
      Begin VB.Image Image3 
         Height          =   465
         Left            =   210
         Picture         =   "custReg.frx":39D3
         Stretch         =   -1  'True
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lblCompanyName 
         Caption         =   "Company Name :"
         Height          =   315
         Left            =   210
         TabIndex        =   66
         Top             =   840
         WhatsThisHelpID =   10
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmCustReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PrevTab As Integer
Public Event OKClick()
Public Event CancelClick()
Public Event LookUpClick(strSearch As String)
Public Event ClearClick()
Public Event WindowClosed()

'Status variable to track if an action is complete.
Public Status As Integer

Private m_StrName As String
Private m_Cancel As Boolean
Private m_Remove As Boolean
Private WithEvents m_AddGroup As clsAddGroup
Attribute m_AddGroup.VB_VarHelpID = -1
Private m_NewCustomer As Boolean
''Photo/Sig start
Dim m_CustID As String

Dim fCurImgPhoto As String
Dim fCurPhotoNum As Integer
Dim fCurPhotoDate As String

Dim fCurImgSignature As String
Dim fCurSignatureNum As Integer
Dim fCurSignatureDate As String

'Form level variables
Private signatures() As String
Private photos() As String
Private m_PhotTabIndex As Byte
Private m_kycDocTabIndex As Byte

''photo/Sing end

Public Property Let NewCustomer(NewValue As Boolean)
    m_NewCustomer = NewValue
    m_CustID = ""
End Property



Private Function ValidData() As Boolean
On Error GoTo Exit_Line
    'Check the UI Type
    Dim UIType As Byte
    UIType = Val(cmbCustType.Tag)
    If Val(fraCust(UIType).Tag) = 0 Then
        Debug.Print "Kannada"
        MsgBox "Select the " & IIf(UIType, "Institution", "individual") & " UI " _
            , vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
' Verify the data entered by user...
    ' Check the name.
    Dim StrName As String
    If UIType = 0 Then
        StrName = txtFirstName & txtMiddleName & txtLastName
    Else
        StrName = txtInstName
    End If
    
    If Trim$(StrName) = "" Then
        'MsgBox "Enter the customer name to register.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(565), _
                vbInformation, wis_MESSAGE_TITLE
        TabStrip1.Tabs(1).Selected = True
        txtFirstName.SetFocus
        GoTo Exit_Line
    End If

    ' Check for invalid names.
    With txtFirstName
        If IsNumeric(Left$(.Text, 1)) Then
            'MsgBox "Name cannot begin with a digit.", _
                    vbExclamation, wis_MESSAGE_TITLE
             MsgBox GetResourceString(516), _
                    vbExclamation, wis_MESSAGE_TITLE
            TabStrip1.Tabs(1).Selected = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End If
    End With
    
    ' Date of Birth.
    If UIType = 0 Then
        ' Check, if age is entered.
        If Trim$(txtDOB.Text) = "" Then
            'If he is updating then need not to ask about the DOB
            'Do Nothing
        ' Check, if the date is valid.
        ElseIf Not DateValidate(txtDOB.Text, "/", True) Then
            'MsgBox "Enter a valid date.", _
                vbInformation, wis_MESSAGE_TITLE
             MsgBox GetResourceString(501), _
                vbInformation, wis_MESSAGE_TITLE
            TabStrip1.Tabs(1).Selected = True
            ActivateTextBox txtDOB
            GoTo Exit_Line
        Else
            'If it is a minor, check if the guardian is specified.
            If DateDiff("yyyy", GetSysFormatDate(txtDOB.Text), GetSysFormatDate(gStrDate)) < 18 Then
                With txtGuardian
                    If Trim$(.Text) = "" Then
                        'MsgBox "Specify the guardian for minor.", vbInformation
                        MsgBox GetResourceString(542), vbInformation
                        TabStrip1.Tabs(1).Selected = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .SetFocus
                        GoTo Exit_Line
                    End If
                End With
            End If
        End If
    Else
      ' Check, if the date is valid.
      If txtEstd <> "" Then
        If Not DateValidate(txtEstd.Text, "/", True) Then
            'MsgBox "Enter a valid date.", _
                vbInformation, wis_MESSAGE_TITLE
             MsgBox GetResourceString(501), _
                vbInformation, wis_MESSAGE_TITLE
            TabStrip1.Tabs(1).Selected = True
            ActivateTextBox txtEstd
            GoTo Exit_Line
        End If
      End If
    End If
    
    ' Gender.
    If cmbGender.ListIndex = -1 Then cmbGender.ListIndex = 0
    
    ' Marital status.
    If cmbMaritalStatus.ListIndex = -1 Then cmbMaritalStatus.ListIndex = 0
    
    ' Check if address is specified.
    If Trim$(HomeAddress) = "" And _
            Trim$(OfficeAddress) = "" Then
        'MsgBox "Specify either home or office address " _
            & "of the customer.", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(543), vbInformation, wis_MESSAGE_TITLE
        TabStrip1.Tabs(2).Selected = True
        txtHomeNo.SetFocus
        GoTo Exit_Line
    End If

ValidData = True
Exit_Line:
    Exit Function

End Function

Private Sub setPhotosLabel()
    If (UBound(photos) > 0) Then
        lblImgCount.Caption = GetResourceString(415) & ":" & fCurPhotoNum & "/" & UBound(photos)
    End If
    If (UBound(photos) > 0) Then
        lblImgDate.Caption = GetResourceString(37) & ":" & FileDateTime(gImagePath & photos(fCurPhotoNum))
    End If
    If (UBound(signatures) > 0) Then
        lblSgnCount.Caption = GetResourceString(416) & ":" & fCurSignatureNum & "/" & UBound(signatures)
    End If
    If (UBound(signatures) > 0) Then
        lblSgnDate.Caption = GetResourceString(37) & ":" & FileDateTime(gImagePath & signatures(fCurSignatureNum))
    End If
    
cmdImgNext.Enabled = (UBound(photos) > fCurPhotoNum) And UBound(photos) > 0
cmdSgnNext.Enabled = (UBound(signatures) > fCurSignatureNum) And UBound(signatures) > 0

cmdImgPrev.Enabled = fCurPhotoNum > 1
cmdSgnPrev.Enabled = fCurSignatureNum > 1
   
End Sub


Private Sub cmbCaste_LostFocus()
    Me.txtCaste.Text = cmbCaste.Text
    txtCaste.Visible = True
    cmbCaste.Visible = False
    
    'Get the defalult Socity Place fromInstall
    'TabStrip1.Tabs(2).Selected = True

End Sub


Private Sub cmbCustType_Click()
Dim UIType As Byte
If cmbCustType.ListIndex < 0 Then Exit Sub
Dim rst As Recordset
With cmbCustType
    gDbTrans.SqlStmt = "SELECT UITYPE FROM CustomerType WHERE CustType = " & .ItemData(.ListIndex)
End With
UIType = 0
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    UIType = FormatField(rst("UIType"))
    cmbCustType.Tag = UIType
    Set rst = Nothing
End If

fraCust(UIType).Tag = 1
fraCust(1 - UIType).Tag = 0


End Sub


Private Sub cmbCustType_LostFocus()
fraCust(Val(cmbCustType.Tag)).ZOrder 0

optUI(Val(cmbCustType.Tag)).Value = False
End Sub


Private Sub cmbFarmerType_LostFocus()
    TabStrip1.Tabs(2).Selected = True
End Sub


Private Sub cmbGender_GotFocus()
cmdLookUp.Default = False
End Sub


Private Sub cmbHomeCity_Change()
    txtHomeCity.Text = cmbHomeCity.List(cmbHomeCity.ListIndex)
End Sub

Private Sub cmbHomeCity_Click()
With cmbHomeCity
    txtHomeCity.Text = .List(.ListIndex)
    If .ListIndex >= 0 Then cmdOk.Default = True
End With
End Sub


Private Sub cmbHomeCity_LostFocus()
txtHomeCity.Text = cmbHomeCity.Text
txtHomeCity.Visible = True
cmbHomeCity.Visible = False
End Sub

Private Sub cmbMaritalStatus_GotFocus()
    cmdLookUp.Default = False
End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClick
Me.Hide
End Sub

Private Sub cmdCaste_Click()
Set m_AddGroup = New clsAddGroup
'm_AddGroup.ShowAddGroup (grpCaste)
m_AddGroup.ShowAddGroup grpCaste
Set m_AddGroup = Nothing

Exit Sub

'Dim Rst As adodb.Recordset
'Set m_frmPlace = New frmPlaceCaste '8th may 00
'With m_frmPlace
'   'Load m_frmPlace
'   .Caption = GetResourceString(100)
'   .lblPlace.Caption = GetResourceString(100)
'   .lblPlaceList.Caption = GetResourceString(100)
'   .Show vbModal, Me
'End With
'
'If m_Cancel Then Exit Sub
'If m_Remove Then
'   gDbTrans.BeginTrans
'   gDbTrans.SQLStmt = "Delete * from CasteTab  where Caste= " & AddQuotes(m_StrName, True)
'Else
'    ' Search for the existing Place
'    gDbTrans.SQLStmt = "Select * from CasteTab Where Caste = " & AddQuotes(m_StrName, True)
'     If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then Exit Sub ' If this place has already enterd
'     gDbTrans.BeginTrans
'
'     gDbTrans.SQLStmt = "Insert Into CasteTAB Values (" & AddQuotes(m_StrName, True) & ")"
'End If
'
'If Not gDbTrans.SQLExecute Then
'gDbTrans.RollBack
''msgbox "Unable To "
'End If
'gDbTrans.CommitTrans
'
'Call LoadCastes
End Sub

Private Sub cmdCaste_GotFocus()
cmdLookUp.Default = False
End Sub


Private Sub cmdFarmerType_Click()
    Set m_AddGroup = New clsAddGroup
    'm_AddGroup.ShowAddGroup (grpCaste)
    m_AddGroup.ShowAddGroup (grpFarmer)
    Set m_AddGroup = Nothing


End Sub

Private Sub cmdLookup_Click()
If InStr(1, cmdLookUp.Caption, GetResourceString(8), vbTextCompare) <> 0 Then     ' Clear
    RaiseEvent ClearClick
    cmdLookUp.Caption = GetResourceString(17)     ' Set as lookup

ElseIf InStr(1, cmdLookUp.Caption, GetResourceString(17), vbTextCompare) <> 0 Then    ' Lookup
    Screen.MousePointer = vbHourglass
    Dim strSearch As String
    strSearch = Trim$(txtFirstName.Text)
    If strSearch = "" Then strSearch = Trim$(txtMiddleName.Text)
    If strSearch = "" Then strSearch = Trim$(txtLastName.Text)
    If fraCust(1).Tag = 1 Then strSearch = Trim$(txtInstName.Text)
    RaiseEvent LookUpClick(Trim(strSearch))
    Screen.MousePointer = vbDefault
    
End If
cmdLookUp.Enabled = True

End Sub

Private Sub cmdNewCust_Click()

End Sub

Private Sub cmdOk_Click()
If Not ValidData Then Exit Sub
RaiseEvent OKClick

Me.Hide
End Sub

Private Sub cmdPlace_Click()
Set m_AddGroup = New clsAddGroup
m_AddGroup.ShowAddGroup (grpPlace)

Set m_AddGroup = Nothing

Exit Sub

'
'Dim Rst As adodb.Recordset
'Set m_frmPlace = New frmPlaceCaste '8th may 00

'With m_frmPlace
'    Load m_frmPlace
'    .lblPlace = LoadResString(glangoffset + 270)
'    .lblPlaceList = LoadResString(glangoffset + 270)
'    .Show vbModal, Me
'End With
'If m_Cancel Then Exit Sub
'If m_Remove Then
'   gDbTrans.BeginTrans
'    gDbTrans.SQLStmt = "Delete * from placeTab  where place= " & AddQuotes(m_StrName, True)
'Else
'    ' Search for the existing Place
'    gDbTrans.SQLStmt = "Select * from PlaceTab Where Place = " & AddQuotes(m_StrName, True)
'     If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then Exit Sub
'     gDbTrans.BeginTrans
'     gDbTrans.SQLStmt = "Insert Into PlaceTAB Values (" & AddQuotes(m_StrName, True) & ")"
'End If
'
'If Not gDbTrans.SQLExecute Then
'    gDbTrans.RollBack
'End If
'gDbTrans.CommitTrans
'Call LoadPlaces
'   TabStrip1.Tabs(2).Selected = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
    Dim I As Byte
    With TabStrip1
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
End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(";") Then KeyAscii = 0
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim count As Integer
'count=iif(shift
count = TabStrip1.Tabs.count
'If KeyCode = 9 Then
'    If Shift = 2 Then
'        If TabStrip1.Tabs(Count) = Count Then
'            Count = 1
'        End If
'    ElseIf Shift = 3 Then
''        If TabStrip1.Tabs(Count) = 1 Then
''
''        End If
'    End If
'End If
End Sub


Private Sub Form_Load()
' Center the form.
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
'Set The Kannada fonts & caption
Call SetKannadaCaption
Dim I As Integer
' Remove the border for frames.
pic(1).BorderStyle = 0
pic(1).Visible = False
For I = pic.LBound + 1 To pic.UBound
    pic(I).BorderStyle = 0
    pic(I).Visible = False
    pic(I).Left = pic(1).Left
    pic(I).Top = pic(1).Top
Next
fraCust(0).ZOrder 0
fraCust(0).Tag = 1
fraCust(1).Tag = 0

'Marital Status Combo
    With cmbMaritalStatus
        .Clear
        .AddItem GetResourceString(297)    '"Single"
        .AddItem GetResourceString(298)   '"Maried"
        .AddItem GetResourceString(299)    '"Divorced"
        .AddItem GetResourceString(300)    '"Widower"
        .AddItem GetResourceString(301)    '"Widow"
        .AddItem GetResourceString(237)    '"Unspecified"
    End With
' Set the display order for tabctl.
TabStrip1.ZOrder 1
TabStrip1_Click

    
'I have added this because iam loading this in run time insted of design time - lingappa
With cmbTitle
    .Clear
    .AddItem GetResourceString(321)
    .AddItem GetResourceString(322)
    .AddItem GetResourceString(323)
    .AddItem GetResourceString(324)
    .AddItem GetResourceString(325)
    .AddItem GetResourceString(326)
    .AddItem "M/S"
End With
With cmbGender
    .Clear
    .AddItem GetResourceString(385)
    .ItemData(.newIndex) = wisMale
    .AddItem GetResourceString(386)
    .ItemData(.newIndex) = wisFemale
    .AddItem GetResourceString(237)
    .ItemData(.newIndex) = wisNoGender
End With

'Set the Customer Name UI
fraCust(0).Left = 0
fraCust(1).Left = 0
fraCust(0).Top = 0
fraCust(1).Top = 0
fraCust(0).BorderStyle = 0
fraCust(1).BorderStyle = 0

'Load places for Places tab
    Call LoadPlaces(cmbHomeCity)
''Load Caste From Table
    Call LoadCastes(cmbCaste)
'' Load City
'load Farmer Types
    Call LoadFarmerTypes(cmbFarmerType)

Dim SetUp As clsSetup
Set SetUp = New clsSetup
txtHomeTaluka = SetUp.ReadSetupValue("Customer", "Taluka", "")
txtOffTaluka = SetUp.ReadSetupValue("Customer", "Taluka", "")
txtHomeDistrict = SetUp.ReadSetupValue("Customer", "District", "")
txtOffDistrict = SetUp.ReadSetupValue("Customer", "District", "")
txtHomePin = SetUp.ReadSetupValue("Customer", "PINCODE", "")
txtOffPin = SetUp.ReadSetupValue("Customer", "PINCODE", "")
txtHomeState = SetUp.ReadSetupValue("Customer", "State", "")
txtOffState = SetUp.ReadSetupValue("Customer", "State", "")
Set SetUp = Nothing

'Load Customer Type
Call LoadCustomerTypes(cmbCustType)
'Load The Id type for KYC
Call LoadKYCIDTypes(cmbKycId1)
Call LoadKYCIDTypes(cmbKycID2)


Me.TabStrip1.Tabs(1).Selected = True
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
   'set icon for the form caption
   Me.Icon = LoadResPicture(161, vbResIcon)
   
If Len(gImagePath) Then
    PhotoInitialize
Else
    TabStrip1.Tabs.Remove (TabStrip1.Tabs.count)
End If
''Show or Not to Show the FarmerType combo
Dim Show As Boolean
Show = CBool(GetConfigValue("FarmerType", "True"))
lblFarmerType.Visible = Show
cmbFarmerType.Visible = Show
cmdFarmerType.Visible = Show


End Sub

' Returns the address string by concatenating
' the relevant address elements.
Public Property Get HomeAddress() As String

' Make sure that atleast one of the
' address fields is specified.
If Trim(txtHomeNo.Text) = "" And _
        Trim$(txtHomeStreet) = "" And _
        Trim$(txtHomeCity.Text) = "" And _
        Trim$(txtHomeState.Text) = "" And _
        Trim$(txtHomePin.Text) = "" Then
    Exit Property
End If


Dim strAddress As String
HomeAddress = txtHomeNo.Text & ";" _
    & txtHomeStreet.Text & ";" _
    & txtHomeCity.Text & ";" _
    & txtHomeDistrict.Text & ";" _
    & txtHomeState.Text & ";" _
    & txtHomePin.Text & ";" _
    & txtHomeTaluka.Text

End Property

Public Property Let HomeAddress(ByVal vNewValue As String)
On Error Resume Next
Dim strTmp() As String

' Break up the given string into
' array elements for processing.
GetStringArray vNewValue, strTmp(), ";"
ReDim Preserve strTmp(6)

'Reset the text boxes
txtHomeNo.Text = ""
txtHomeStreet.Text = ""
txtHomeDistrict.Text = ""
txtHomeCity.Text = ""
txtHomeState.Text = ""
txtHomePin.Text = ""

Dim SetUp As New clsSetup
' Fill the form controls.
txtHomeNo.Text = strTmp(0)
txtHomeStreet.Text = strTmp(1)
txtHomeCity.Text = strTmp(2)
txtHomeTaluka.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "Taluka", ""), strTmp(6))
txtHomeDistrict.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "District", ""), strTmp(3))
txtHomeState.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "State", ""), strTmp(4))
txtHomePin.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "Pincode", ""), strTmp(5))


Set SetUp = Nothing
End Property
' Returns the address string by concatenating
' the relevant address elements.
Public Property Get OfficeAddress() As String

' Make sure that atleast one of the
' address fields is specified.
If Trim$(txtCompanyName.Text) = "" And _
    Trim$(txtJobTitle.Text) = "" And _
    Trim(txtOfficeNo.Text) = "" And _
        Trim$(txtOffStreet) = "" And _
        Trim$(txtOffCity.Text) = "" And _
        Trim$(txtOffState.Text) = "" And _
        Trim$(txtOffPin.Text) = "" Then
    Exit Property
End If

OfficeAddress = txtCompanyName.Text & ";" _
    & txtJobTitle.Text & ";" _
    & txtOfficeNo.Text & ";" _
    & txtOffStreet.Text & ";" _
    & txtOffCity.Text & ";" _
    & txtOffDistrict.Text & ";" _
    & txtOffState.Text & ";" _
    & txtOffPin.Text & ";" _
    & txtOffTaluka.Text

End Property
Public Property Let OfficeAddress(ByVal vNewValue As String)
On Error Resume Next
Dim strTmp() As String

Dim SetUp As New clsSetup

' Break up the given string into
' array elements for processing.
GetStringArray vNewValue, strTmp(), ";"
ReDim Preserve strTmp(9)
'Clear text boxes
txtCompanyName.Text = ""
txtJobTitle.Text = ""
txtOfficeNo.Text = ""
txtOffStreet.Text = ""
txtOffCity.Text = ""
txtOffDistrict.Text = ""
txtOffState.Text = ""
txtOffPin.Text = ""


' Fill the form controls.
txtCompanyName.Text = strTmp(0)
txtJobTitle.Text = strTmp(1)
txtOfficeNo.Text = strTmp(2)
txtOffStreet.Text = strTmp(3)
txtOffCity.Text = strTmp(4)
txtOffTaluka.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "Taluka", ""), strTmp(8))
txtOffDistrict.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "District", ""), strTmp(5))
txtOffState.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "State", ""), strTmp(6))
txtOffPin.Text = IIf(vNewValue = "", SetUp.ReadSetupValue("Customer", "PinCode", ""), strTmp(7))

Set SetUp = Nothing
End Property

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    cancel = True
    Me.Hide
End If
End Sub

Private Sub Form_Unload(cancel As Integer)
'   ""(Me.hwnd, False)
RaiseEvent WindowClosed
End Sub

Private Sub m_AddGroup_ItemAdded(strAddItem As String, NewID As Long)

If m_AddGroup.GroupType = grpPlace Then
    Me.cmbHomeCity.AddItem strAddItem
ElseIf m_AddGroup.GroupType = grpCaste Then
    Me.cmbCaste.AddItem strAddItem
ElseIf m_AddGroup.GroupType = grpCustomer Then
    With cmbCustType
        .AddItem strAddItem
        .ItemData(.newIndex) = NewID
    End With
ElseIf m_AddGroup.GroupType = grpFarmer Then
    With cmbFarmerType
        .AddItem strAddItem
        .ItemData(.newIndex) = NewID
    End With
End If

End Sub

Private Sub m_AddGroup_ItemDeleting(strDelItem As String, cancel As Integer)

Dim cmb As ComboBox
Dim rst As Recordset

    Set rst = Nothing
    Set cmb = cmbCustType

Dim count As Integer

If m_AddGroup.GroupType = grpPlace Then Set cmb = cmbHomeCity
If m_AddGroup.GroupType = grpCaste Then Set cmb = cmbCaste


With cmb
    For count = 0 To .ListCount - 1
        If StrComp(.List(count), strDelItem, vbTextCompare) = 0 Then
            .RemoveItem count
            Exit For
        End If
    Next
End With


End Sub


Private Sub m_frmPlace_AddClick(StrName As String)
m_StrName = StrName
m_Remove = False
m_Cancel = False
End Sub


Private Sub m_frmPlace_CancelClick(cancel As Boolean)
m_Cancel = True
End Sub

Private Sub m_frmPlace_RemoveClick(StrName As String)
m_StrName = StrName
m_Remove = True
m_Cancel = False
End Sub

Private Sub optIndividual_Click()

End Sub

Private Sub optUI_Click(Index As Integer)

    fraCust(IIf(Index, 0, 1)).ZOrder 0

    Dim TabIndex As Integer
    Dim Ctrl As Control
    TabIndex = 2
    
    Call SetControlsTabIndex(IIf(Index, 70, 1), lblTitle, cmbTitle, lblFirstName, txtFirstName, lblMiddleName, txtMiddleName, _
    lblLastName, txtLastName, lbldoB, txtDOB, lblGuardian, txtGuardian, lblGender, cmbGender, _
    lblMarital, cmbMaritalStatus, lblProfession, txtProfession, lblCaste, cmbCaste, cmdCaste)
    
    Call SetControlsTabIndex(IIf(Index, 1, 70), Label1, cmbInstTitle, lblCompanyName, txtCompanyName, _
    lblInstHead, txtInstHead, lblEstDate, txtEstd, lblCustType, cmbCustType)

    optUI(0).TabIndex = 90
    optUI(1).TabIndex = 91

fraCust(IIf(Index, 0, 1)).Tag = 1
fraCust(IIf(Index, 1, 0)).Tag = 0

optUI(IIf(Index, 0, 1)).Value = False
End Sub


Private Sub TabStrip1_Click()
On Error Resume Next

Dim curTab As Integer
curTab = TabStrip1.SelectedItem.Index
If curTab = m_PrevTab Then Exit Sub

' Hide the previous frame.
If m_PrevTab <> 0 Then pic(m_PrevTab).Visible = False
pic(curTab).Visible = True
pic(curTab).ZOrder (0)
Debug.Print pic(curTab).Parent.name
'fraPhoto.Visible = IIf(curTab = 5, True, False)
fraPhoto.Visible = True
#If COMMENTED Then
    If curTab = 1 Then
        cmbTitle.SetFocus
    ElseIf curTab = 2 Then
        txtHomeNo.SetFocus
    ElseIf curTab = 3 Then
        txtCompanyName.SetFocus
    End If
#End If
' Reset the previous tab.
m_PrevTab = curTab


End Sub

Private Sub txtCaste_GotFocus()
    cmbCaste.Visible = True
    txtCaste.Visible = False
    cmbCaste.SetFocus
End Sub


Private Sub txtDOB_LostFocus()
'
' Check if the customer is minor.
' If minor, enable the guardian textbox.
'

' First, check if it is a valid date.
If DateValidate(txtDOB.Text, "/", True) Then
    ' Compare it with current date, to find
    ' if the client is a minor.
    If DateDiff("yyyy", GetSysFormatDate(txtDOB.Text), GetSysFormatDate(gStrDate)) < 18 Then
        With txtGuardian
            .Enabled = True
            .BackColor = vbWhite
        End With
        lblGuardian.Enabled = True
    Else
        With txtGuardian
            .Enabled = False
            .BackColor = .Parent.BackColor
        End With
        lblGuardian.Enabled = False
    End If
Else
    With txtGuardian
        .Enabled = False
        .BackColor = .Parent.BackColor
    End With
    lblGuardian.Enabled = False
End If
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
 
 'Common to the Form
    
    Me.TabStrip1.Tabs(1).Caption = GetResourceString(116)
    TabStrip1.Tabs(2).Caption = GetResourceString(117) '
    TabStrip1.Tabs(3).Caption = GetResourceString(118)
    TabStrip1.Tabs(4).Caption = "KYC DOCs"
    TabStrip1.Tabs(5).Caption = GetResourceString(415)
    cmdHelp.Caption = GetResourceString(16) '
    Me.cmdLookUp.Caption = GetResourceString(17)
    cmdOk.Caption = GetResourceString(1)
    cmdCancel.Caption = GetResourceString(2)
    
'Tabstrip of Office
    lblOffTitle.Caption = GetResourceString(137)
    lblCompanyName.Caption = GetResourceString(138)
    lblJob = GetResourceString(139) '
    lblOffNo = GetResourceString(60)
    lblOffAddr = GetResourceString(130)
    lblOffCity = GetResourceString(131) '
    lblOffDist = GetResourceString(132) '
    Me.lblOffState = GetResourceString(133)
    lblOffPIN = GetResourceString(134) '
    lblOffPhone = GetResourceString(135)  '"
  'TabStrip of Home Address
    lblHomeTitle.Caption = GetResourceString(129)
    lblHouseNo.Caption = GetResourceString(60)
    lblStreetAddr = GetResourceString(130)
    lblCity = GetResourceString(131)
    lblDistrict = GetResourceString(132)
    Me.lblState = GetResourceString(133)
    lblPinCode = GetResourceString(134)
    lblHomePhone = GetResourceString(135)
    lblHomeEmail = GetResourceString(136)
  'TabStrip of  Personnel
    lblPersonTitle = GetResourceString(114)
    lblTitle = GetResourceString(119)
    lblFirstName = GetResourceString(120)
    lblMiddleName = GetResourceString(121)
    lblLastName = GetResourceString(122)   '
    lbldoB = GetResourceString(123)   '
    lblGuardian = GetResourceString(124)
    lblGender = GetResourceString(125)
    lblMarital = GetResourceString(126)
    lblProfession = GetResourceString(127)
    lblFarmerType = GetConfigValue("FarmerTypeName", GetResourceString(378)) & " " & GetResourceString(253)
    lblCaste = GetResourceString(111)
'Tab strip of KYC
    lblKycTitle.Caption = "Pleae Enter KYC Details"
    lblKycPhone.Caption = GetResourceString(205, 35)
    lblKycPhoneType1.Caption = GetResourceString(239, 60)
    lblKycIdType1.Caption = GetResourceString(447, 253)
    lblKycIdType2.Caption = GetResourceString(447, 253)
    lblKycID1.Caption = GetResourceString(447, 60)
    lblKycID2.Caption = GetResourceString(447, 60)
    lblKycBank1.Caption = GetResourceString(446, 35)
    lblKycBank2.Caption = GetResourceString(237, 418, 35)
    lblKycAccNum1.Caption = GetResourceString(446, 60, 482)
    lblKycAccNum2.Caption = GetResourceString(237, 418, 60, 482)
    
'''2nd frame
    'cmdNewCust.Caption = GetResourceString(260,205,253)
    lblCustType.Caption = GetResourceString(205, 253)
    optUI(0).Caption = GetResourceString(254) & " UI"
    optUI(1).Caption = GetResourceString(254) & " UI"
    
   
'Now Change the font of English only
'lblEnglishName.Font = "MS Sans Serif"
txtEnglishName.Font = "MS Sans Serif"
lblEnglishName.Caption = GetResourceString(35, 468)

End Sub
Private Sub txtEnglishName_GotFocus()
    Call ToggleWindowsKey(winScrlLock, False)
    If Len(Trim$(txtEnglishName.Text)) < 1 And gLangOffSet = wis_KannadaSamhitaOffset Then
        ''Translate teh first name and last name
        txtEnglishName.Text = Trim$(ConvertToEnglish(txtFirstName.Text) & " " & ConvertToEnglish(txtLastName.Text))
    
    End If
End Sub

Private Sub txtEnglishName_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub

Private Sub txtHomeCity_GotFocus()
'cmbHomeCity.Text = txtHomeCity.Text
cmbHomeCity.Visible = True
txtHomeCity.Visible = False
cmbHomeCity.SetFocus

End Sub

Private Sub cmdImgAdd_Click()

    CommonDialog1.Filter = "Pictures (*.bmp;*.jpg;*.tif)|*.jpg;|*.tif;"
    CommonDialog1.ShowOpen
    Dim imgfile As String

    If (CommonDialog1.Filename = "") Then Exit Sub

    imgfile = AddImageFile(m_CustID, INDEX2000_PHOTO, CommonDialog1.Filename)
    Dim pos As Integer
    pos = InstrRev(imgfile, "\")
    If pos Then imgfile = Mid$(imgfile, pos + 1)

    Dim newIndex As Integer
    newIndex = UBound(photos) + 1
    ReDim Preserve photos(newIndex)
    
    photos(newIndex) = imgfile
    fCurPhotoNum = fCurPhotoNum + 1
    picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))

    setPhotosLabel
End Sub
Private Sub cmdImgDel_Click()

If MsgBox(GetResourceString(830), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub

On Error GoTo Hell
 
'Now delete the image from file system.
Kill gImagePath & photos(fCurPhotoNum)
picphoto.Picture = LoadPicture()

PhotoInitialize

Hell:

setPhotosLabel
End Sub

Private Sub cmdImgNext_Click()
Dim imgfile As String

'photo display
If (fCurPhotoNum < UBound(photos)) Then
    fCurPhotoNum = fCurPhotoNum + 1
    picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
End If

    
    setPhotosLabel
End Sub
Private Sub cmdImgPrev_Click()
 
    fCurPhotoNum = fCurPhotoNum - 1
    picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))

    
    setPhotosLabel

End Sub

Private Sub cmdSgnAdd_Click()

On Error GoTo Hell

CommonDialog1.Filter = "Pictures (*.bmp;*.jpg;*.tif)|*.jpg;|*.tif;"
CommonDialog1.ShowOpen

Dim imgfile As String

If (CommonDialog1.Filename = "") Then Exit Sub
    
    imgfile = AddImageFile(m_CustID, INDEX2000_SIGN, CommonDialog1.Filename)
    Dim pos As Integer
    pos = InstrRev(imgfile, "\")
    
    If pos Then imgfile = Mid$(imgfile, pos + 1)

    Dim newIndex As Integer
    newIndex = UBound(signatures) + 1
    ReDim Preserve signatures(newIndex)
    signatures(newIndex) = imgfile
    fCurSignatureNum = fCurSignatureNum + 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))

setPhotosLabel
Exit Sub

Hell:
picSign.Picture = LoadPicture()
On Error GoTo 0

End Sub
Private Sub cmdSgnDel_Click()

    If MsgBox(GetResourceString(831), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
 
    On Error GoTo Hell

    ' Now delete the image from file system.
    Kill gImagePath & signatures(fCurSignatureNum)
    picSign.Picture = LoadPicture()
    PhotoInitialize

Hell:
  
    
setPhotosLabel
End Sub

Private Sub cmdSgnNext_Click()
On Error GoTo Hell

Dim imgfile As String
'signature display

If (fCurSignatureNum < UBound(signatures)) Then
    fCurSignatureNum = fCurSignatureNum + 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
End If

setPhotosLabel

Exit Sub

Hell:
picSign.Picture = LoadPicture()
On Error GoTo 0
End Sub
Private Sub cmdSgnPrev_Click()
On Error GoTo Hell
    fCurSignatureNum = fCurSignatureNum - 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))

    setPhotosLabel

Exit Sub

Hell:
picSign.Picture = LoadPicture()
On Error GoTo 0
    
    
End Sub
Public Sub PhotoInitialize()

On Error GoTo Hell
    
    ' Read the configuration file for finding out the path of image files.
    Dim ret As Long
    Dim strRet As String
    cmdImgNext.Enabled = False
    cmdSgnNext.Enabled = False
    cmdImgPrev.Enabled = False
    cmdSgnPrev.Enabled = False
    cmdImgDel.Enabled = False
    cmdSgnDel.Enabled = False
    
    'Load the image names (photos and signatures) into array
    If (loadImageFiles(m_CustID, INDEX2000_PHOTO, photos) = False) Then
        MsgBox "Error while loading photos.", vbOKOnly
    End If
    
    If (loadImageFiles(m_CustID, INDEX2000_SIGN, signatures) = False) Then
        MsgBox "Error while loading signatures.", vbOKOnly
    End If
    fCurPhotoNum = 0
    ' Display the first picture if there exists one.
    If (UBound(photos) > 0) Then
        fCurPhotoNum = 1
        picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
        cmdImgDel.Enabled = True
    ElseIf Len(Dir(gImagePath & "image_default.jpg")) > 0 Then
        picphoto.Picture = VB.LoadPicture(gImagePath & "image_default.jpg")
        fCurPhotoNum = 0
    End If
    
    fCurSignatureNum = 0
    If (UBound(signatures) > 0) Then
        fCurSignatureNum = 1
        picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
        cmdSgnDel.Enabled = True
    ElseIf Len(Dir(gImagePath & "sign_default.jpg")) Then
        picSign.Picture = VB.LoadPicture(gImagePath & "sign_default.jpg")
        fCurSignatureNum = 0
        cmdSgnDel.Enabled = False
    End If
    
    ' Since the first img will be loaded, there is no prev image to display,
    ' disable the prev button.
    
    'aswell as signature count label
    setPhotosLabel

Hell:

End Sub

Public Function setAccNo(CustomerID As String)
    m_CustID = CustomerID
End Function


