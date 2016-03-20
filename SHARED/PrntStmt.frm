VERSION 5.00
Begin VB.Form frmPrintTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Tranascation"
   ClientHeight    =   4530
   ClientLeft      =   7890
   ClientTop       =   2670
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optNewPassbook 
      Caption         =   "New Passbook"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.Frame fraDate 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   4575
         Begin VB.CommandButton cmdEndDate 
            Caption         =   ".."
            Height          =   315
            Left            =   3360
            TabIndex        =   14
            Top             =   720
            Width           =   315
         End
         Begin VB.CommandButton cmdStartDate 
            Caption         =   ".."
            Height          =   315
            Left            =   3360
            TabIndex        =   13
            Top             =   240
            Width           =   345
         End
         Begin VB.TextBox txtEndDate 
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   720
            Width           =   1485
         End
         Begin VB.TextBox txtStDate 
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "End Date :"
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblStDate 
            AutoSize        =   -1  'True
            Caption         =   "&Start Date :"
            Height          =   300
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1590
         End
      End
      Begin VB.ComboBox cmbRecords 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PrntStmt.frx":0000
         Left            =   1320
         List            =   "PrntStmt.frx":0013
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optCert 
         Caption         =   "Print FD Certificate"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CheckBox chkNewPassbook 
         Caption         =   "This is a new passbook"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.OptionButton optLastPrint 
         Caption         =   "Transaction from Last Stament to till date"
         Height          =   465
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   4515
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Transaction between Two dates"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4395
      End
      Begin VB.Label lblLast 
         Caption         =   "Print previous                        records"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   2400
         Y2              =   2400
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DateClick(StartIndiandate As String, EndIndianDate As String)

' Mrudula: 26/June/2014
' Mrudula: 20/June/2014
' Modified the signature of this event to include a boolean parameter.
' when this parameter is set as true, it means, printing is to be
' done on a new passbook.
'Public Event TransClick()
Public Event TransClick(bNewPassbook As Boolean)
Public Event PageClick()

Public Event CancelClick()

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
lblStDate.Caption = GetResourceString(109)
lblEndDate.Caption = GetResourceString(110)

End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClick
Me.Hide
End Sub

Private Sub cmdEndDate_Click()
With Calendar
    .Left = Screen.Width / 2
    .Top = Screen.Height / 2
    .selDate = gStrDate
    .Show vbModal
    Me.txtEndDate.Text = .selDate
End With

End Sub


Private Sub cmdOk_Click()

If optLastPrint.Value Then
    'RaiseEvent TransClick  // changed signature as below.
    RaiseEvent TransClick(False)
    GoTo Last_Line
ElseIf optCert.Value = True Then
    RaiseEvent PageClick
    GoTo Last_Line
ElseIf optNewPassbook.Value = True Then
    RaiseEvent TransClick(True)
    GoTo Last_Line
End If

'Check for the Date
If txtStDate.Text = "" Then
   MsgBox "Please enter date"
    Exit Sub
End If

If txtEndDate.Text = "" Then
   MsgBox "Please enter date"
    Exit Sub
End If


'Check For Validate of Dates
If Not DateValidate(txtStDate.Text, "/", True) Then
    'Err.Raise 10012, "Invalid Date"
    MsgBox GetResourceString(501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtStDate
    Exit Sub
End If

If Not DateValidate(txtEndDate.Text, "/", True) Then
    'Err.Raise 10012, "Invalid Date"
    MsgBox GetResourceString(501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtEndDate
    Exit Sub
End If

If WisDateDiff(txtStDate.Text, txtEndDate.Text) < 1 Then
    'Err.Raise 10013, "Invalid date difference"
    MsgBox GetResourceString(501), vbCritical, wis_MESSAGE_TITLE
    ActivateTextBox txtEndDate
    Exit Sub
End If


Me.Hide
Screen.MousePointer = vbHourglass
RaiseEvent DateClick(txtStDate.Text, txtEndDate.Text)

'Dim SBAcc As clsSBAcc
'Set SBAcc = New clsSBAcc
'SBAcc.printTxnDetails (Me)


Screen.MousePointer = vbDefault

Last_Line:
Me.Hide
End Sub



Private Sub cmdStartDate_Click()
With Calendar
    .Left = Screen.Width / 2
    .Top = Screen.Height / 2
    .selDate = "1/4/2000"
    .Show vbModal
    Me.txtStDate.Text = .selDate
End With
End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
' to set kannada fonts
Call SetKannadaCaption
optCert.Top = lblLast.Top
cmbRecords.ListIndex = 1
End Sub



Private Sub Option1_Click()

End Sub

Private Sub optDate_Click()
optLastPrint.Value = False
optNewPassbook.Value = False
cmbRecords.Enabled = False
lblLast.Enabled = False
fraDate.Enabled = True
End Sub

Private Sub optLastPrint_Click()
optNewPassbook.Value = False
optDate.Value = False
cmbRecords.Enabled = False
lblLast.Enabled = False
fraDate.Enabled = False
End Sub

Private Sub optNewPassbook_Click()
optDate.Value = False
optLastPrint.Value = False
cmbRecords.Enabled = True
lblLast.Enabled = True
fraDate.Enabled = False
End Sub
