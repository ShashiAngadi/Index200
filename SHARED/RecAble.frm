VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmReceivable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount Receivable"
   ClientHeight    =   3750
   ClientLeft      =   3150
   ClientTop       =   2715
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5025
   Begin VB.CheckBox chkAdd 
      Caption         =   "Add to the account"
      Height          =   300
      Left            =   150
      TabIndex        =   8
      Top             =   1050
      Width           =   3345
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   2430
      TabIndex        =   6
      Top             =   3330
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3750
      TabIndex        =   7
      Top             =   3330
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   400
      Left            =   3720
      TabIndex        =   4
      Top             =   1020
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   1755
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   7
      Cols            =   3
   End
   Begin WIS_Currency_Text_Box.CurrText txtAmount 
      Height          =   345
      Left            =   3570
      TabIndex        =   3
      Top             =   570
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.ComboBox cmbHead 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   3405
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount"
      Height          =   315
      Left            =   3570
      TabIndex        =   2
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblHeadName 
      Caption         =   "Head Name"
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3315
   End
End
Attribute VB_Name = "frmReceivable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CancelClicked()
Public Event AcceptClicked()
Public Event UpdateClicked()
Public Event OkClicked()
Public Event GridClicked()
Public Event WindowClosed()


Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

lblHeadName.Caption = GetResourceString(36, 35)
lblAmount.Caption = GetResourceString(40)

chkAdd.Caption = GetResourceString(36, 271) 'Depositto aaccount
cmdAccept.Caption = GetResourceString(4)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)


End Sub

Private Sub cmdAccept_Click()
'First Validate the controls
If cmbHead.ListIndex < 0 Then
    MsgBox "Select the account head", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If txtAmount.Value = 0 And Not cmbHead.Locked Then
    'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Check whether we are updating the amount or
'entering new head & amount
'if it is updating then the cobo will be locked
If cmbHead.Locked Then
    RaiseEvent UpdateClicked
Else
    RaiseEvent AcceptClicked
End If


End Sub

Private Sub cmdCancel_Click()
RaiseEvent CancelClicked

End Sub

Private Sub cmdOk_Click()
RaiseEvent OkClicked
Me.Hide
End Sub


Private Sub Form_Load()

Call SetKannadaCaption

With grd
    .Clear
    .Rows = 1
    .Rows = 7
    .Cols = 3
    .FixedCols = 1: .FixedRows = 1
    .ColWidth(0) = .Width / 8
    .ColWidth(1) = .Width / 2
    .ColWidth(2) = .Width / 4
    .Row = 0
    .Col = 0: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(33)
    .Col = 1: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(36) & GetResourceString(35)
    .Col = 2: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(40)
End With

txtAmount.Height = cmbHead.Height
txtAmount.Top = cmbHead.Top

'Now Load the Heads
cmbHead.Clear
Dim I As Integer

'load Bank Income Heads
cmbHead.AddItem GetResourceString(366)
Call LoadLedgersToCombo(cmbHead, parBankIncome, False)
If cmbHead.ListCount = 1 Then cmbHead.Clear

'Load Bank Receivable Heads
cmbHead.AddItem GetResourceString(364)
I = cmbHead.ListCount
Call LoadLedgersToCombo(cmbHead, parReceivable, False)
'If the are no heads in the Receivable account
If I = cmbHead.ListCount Then cmbHead.RemoveItem I - 1

'Load Bank Payable Heads
cmbHead.AddItem GetResourceString(357)
I = cmbHead.ListCount
Call LoadLedgersToCombo(cmbHead, parPayAble, False)
'If the are no heads in the Payable account
If I = cmbHead.ListCount Then cmbHead.RemoveItem I - 1

'Load The Bank Expense Heads
cmbHead.AddItem GetResourceString(22)
I = cmbHead.ListCount
Call LoadLedgersToCombo(cmbHead, parBankExpense, False)
'If the are no heads in the Expense account
If I = cmbHead.ListCount Then cmbHead.RemoveItem I - 1

If cmbHead.ListCount = 0 Then cmbHead.AddItem ""

End Sub

Private Sub grd_Click()
With grd
    If .Row = 0 Then Exit Sub
    If .RowData(.Row) = 0 Then Exit Sub
End With

RaiseEvent GridClicked

End Sub


Private Sub txtAmount_GotFocus()

On Error Resume Next

With txtAmount
    .SelStart = Len(.CurrencySymbol)
    .SelLength = Len(.Text) - .SelStart
End With

End Sub

