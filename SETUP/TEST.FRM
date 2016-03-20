VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   1740
   ClientTop       =   2745
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   4650
      TabIndex        =   9
      Top             =   2280
      Width           =   1485
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   345
      Left            =   2730
      TabIndex        =   8
      Top             =   2280
      Width           =   1485
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   2340
      TabIndex        =   7
      Top             =   1710
      Width           =   3225
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   2340
      TabIndex        =   6
      Top             =   1380
      Width           =   3225
   End
   Begin VB.TextBox txtModule 
      Height          =   285
      Left            =   2340
      TabIndex        =   5
      Top             =   1050
      Width           =   3225
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   345
      Left            =   900
      TabIndex        =   0
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Setup Class Module"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   540
      TabIndex        =   4
      Top             =   300
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Key name:"
      Height          =   225
      Left            =   420
      TabIndex        =   3
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Value:"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1770
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Module name:"
      Height          =   225
      Left            =   420
      TabIndex        =   1
      Top             =   1110
      Width           =   1905
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Setup As New clsSetup

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRead_Click()
txtValue.Text = m_Setup.ReadSetupValue(txtModule.Text, txtKey.Text, "Not Found")
End Sub


Private Sub cmdWrite_Click()
If Not m_Setup.WriteSetupValue(txtModule.Text, txtKey.Text, txtValue.Text) Then
    MsgBox "Unable to write Setup property !"
End If
txtValue.Text = ""
End Sub


Private Sub Form_Load()
Call Initialize



End Sub


