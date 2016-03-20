VERSION 5.00
Begin VB.Form frmDateChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Date Change"
   ClientHeight    =   1275
   ClientLeft      =   2445
   ClientTop       =   3075
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   990
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2280
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
   Begin VB.TextBox txtTransDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   3510
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label lblStDate 
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Top             =   150
      Width           =   1740
   End
End
Attribute VB_Name = "frmDateChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DateChanged(strIndianDate As String)
Public Event CancelClicked()

Private Sub SetKannadaCaption()
    
    'Set the FOnt name & size
    SetFontToControls Me
    
    lblStDate = GetResourceString(38, 37)
    cmdOk.Caption = GetResourceString(1)
    cmdCancel.Caption = GetResourceString(2)
End Sub


Private Sub cmdCancel_Click()
RaiseEvent DateChanged("")
Unload Me
End Sub

Private Sub cmdOk_Click()

If Not TextBoxDateValidate(txtTransDate, "/", True, True) Then Exit Sub

RaiseEvent DateChanged(txtTransDate)

Unload Me

End Sub


Private Sub Form_Load()
'Center the form
CenterMe Me

Call SetKannadaCaption
'Set the icon
Me.Icon = LoadResPicture(147, vbResIcon)

txtTransDate.Text = DayBeginDate

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFinYear = Nothing
End Sub




Private Sub txtTransDate_GotFocus()
With txtTransDate
    .SetFocus
    .SelStart = 0
    .SelLength = 2
End With

End Sub


