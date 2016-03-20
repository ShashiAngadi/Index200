VERSION 5.00
Begin VB.Form frmPrintOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WIS   -   Print Options..."
   ClientHeight    =   2280
   ClientLeft      =   3135
   ClientTop       =   2565
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optExcel 
      Caption         =   "Print to MSExcel"
      Height          =   250
      Left            =   90
      TabIndex        =   6
      Top             =   1350
      Width           =   3945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   4140
      TabIndex        =   5
      Top             =   690
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   350
      Left            =   4140
      TabIndex        =   4
      Top             =   270
      Width           =   1200
   End
   Begin VB.CheckBox chkPause 
      Caption         =   "Pause between pages."
      Height          =   250
      Left            =   120
      TabIndex        =   3
      Top             =   1770
      Width           =   3405
   End
   Begin VB.OptionButton optPrintCur 
      Caption         =   "Print only the current page"
      Height          =   250
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   3945
   End
   Begin VB.OptionButton optPrintAllCur 
      Caption         =   "Print all pages from the current page"
      Height          =   250
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   3945
   End
   Begin VB.OptionButton optPrintAllBegin 
      Caption         =   "Print all pages from the beginning"
      Height          =   250
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   3945
   End
End
Attribute VB_Name = "frmPrintOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As Integer

Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
   Ctrl.FontName = gFontName
   If Not TypeOf Ctrl Is ComboBox Then
      Ctrl.FontSize = gFontSize
   End If
Next
cmdPrint.Caption = LoadResString(gLangOffSet + 23)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
optPrintAllBegin.Caption = LoadResString(gLangOffSet + 474)
optPrintAllCur.Caption = LoadResString(gLangOffSet + 475)
optPrintCur.Caption = LoadResString(gLangOffSet + 476)
chkPause.Caption = LoadResString(gLangOffSet + 477)
End Sub


Private Sub cmdCancel_Click()
Me.Status = wis_CANCEL
Me.Hide
End Sub

Private Sub cmdPrint_Click()
Me.Status = wis_OK
Me.Hide
End Sub
Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Status = wis_CANCEL
Me.Caption = "INDEX 2000  - Print Options"
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
' To set the kannada fonts
Call SetKannadaCaption

End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)

End Sub


