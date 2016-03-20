VERSION 5.00
Begin VB.Form frmComponent 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   3255
   ClientLeft      =   1455
   ClientTop       =   2010
   ClientWidth     =   6375
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6375
   Begin VB.CheckBox chkLanguage 
      Caption         =   "Kannada Langauge enabled"
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   660
      Width           =   5475
   End
   Begin VB.CheckBox chkDemo 
      Caption         =   "Install 30 days Demo Version "
      Height          =   345
      Left            =   360
      TabIndex        =   6
      Top             =   210
      Width           =   5475
   End
   Begin VB.OptionButton optServer 
      Caption         =   "Install INDEX 2000 - Server Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1260
      TabIndex        =   5
      Top             =   1320
      Value           =   -1  'True
      Width           =   4365
   End
   Begin VB.OptionButton optClient 
      Caption         =   "Install INDEX 2000 - Client Component"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      TabIndex        =   4
      Top             =   2040
      Width           =   4395
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4710
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   2670
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "#"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3150
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Line Line1 
      X1              =   300
      X2              =   6090
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Image imgWelcome 
      Height          =   480
      Left            =   630
      Picture         =   "COMPNENT.frx":0000
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1305
      TabIndex        =   2
      Top             =   1350
      Visible         =   0   'False
      Width           =   4800
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRunning 
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   3
      Top             =   1755
      Visible         =   0   'False
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpWelcome 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   2490
      Left            =   150
      Top             =   75
      Width           =   6015
   End
End
Attribute VB_Name = "frmComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub chkDemo_Click()
    If chkDemo.Value = vbChecked Then
        gDemoVersion = True
    ElseIf chkDemo.Value = vbUnchecked Then
        gDemoVersion = False
    End If
End Sub


Private Sub chkLanguage_Click()
gstrLangName = IIf(chkLanguage = vbChecked, "Kannada", "English")
End Sub


'frmWelcome
Private Sub cmdExit_Click()
    ExitSetup Me, gintRET_EXIT
End Sub

Private Sub cmdOK_Click()
    If optServer.Value Then
        gboolServer = True
    Else
        gboolServer = False
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    gstrLangName = "English"
    Dim intWidth As Integer

    SetFormFont Me
    cmdExit.Caption = ResolveResString(resBTNEXIT)
    cmdOK.Caption = ResolveResString(resBTNOK)
    'lblRunning.Caption = ResolveResString(resLBLRUNNING)
    
    Caption = gstrTitle
    intWidth = TextWidth(Caption) + cmdOK.Width * 2
    If intWidth > Width Then
        Width = intWidth
    End If

    lblWelcome.Caption = ResolveResString(resWELCOME, "|1", gstrAppName)

    shpWelcome.Move (ScaleWidth - shpWelcome.Width) \ 2
    cmdOK.Left = (ScaleWidth - cmdOK.Width * 1.5 - cmdExit.Width) \ 2
    cmdExit.Left = cmdOK.Left + cmdOK.Width * 1.5

    EtchedLine Me, shpWelcome.Left - 50, cmdOK.Top - cmdOK.Height \ 2, shpWelcome.Width + 100

    CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        ExitSetup Me, gintRET_EXIT
        Cancel = 1
    End If
End Sub

