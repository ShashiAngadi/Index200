VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmCashTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash transaction"
   ClientHeight    =   5430
   ClientLeft      =   990
   ClientTop       =   1695
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDate 
      Caption         =   ".."
      Height          =   285
      Left            =   3420
      TabIndex        =   2
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtTransDate 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Clos&e"
      Height          =   400
      Left            =   6930
      TabIndex        =   17
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Payment Details"
      Height          =   3855
      Index           =   1
      Left            =   270
      TabIndex        =   34
      Top             =   900
      Width           =   7875
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   400
         Index           =   1
         Left            =   4920
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   400
         Index           =   1
         Left            =   6210
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtScroll 
         Height          =   345
         Index           =   1
         Left            =   2010
         TabIndex        =   36
         Top             =   330
         Width           =   1545
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Index           =   1
         Left            =   2010
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   840
         Width           =   5385
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Index           =   1
         Left            =   2010
         TabIndex        =   41
         Top             =   1380
         Width           =   1065
      End
      Begin VB.CommandButton cmdCustName 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   7440
         TabIndex        =   44
         Top             =   1380
         Width           =   285
      End
      Begin VB.TextBox txtCustName 
         Height          =   345
         Index           =   1
         Left            =   4800
         TabIndex        =   43
         Top             =   1380
         Width           =   2565
      End
      Begin VB.CommandButton cmdPay 
         Caption         =   ".."
         Height          =   315
         Left            =   3690
         TabIndex        =   37
         Top             =   330
         Width           =   405
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   375
         Index           =   1
         Left            =   2010
         TabIndex        =   46
         Top             =   1890
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   5880
         TabIndex        =   29
         Top             =   1950
         Width           =   1485
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance"
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   47
         Top             =   2010
         Width           =   1125
      End
      Begin VB.Label lblName 
         Caption         =   "Cust Name :"
         Height          =   225
         Index           =   1
         Left            =   3300
         TabIndex        =   42
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblScroll 
         Caption         =   "Scroll No:"
         Height          =   165
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblAccType 
         Caption         =   "Account type"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   38
         Top             =   870
         Width           =   1755
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account NO"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   45
         Top             =   2040
         Width           =   1635
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4515
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   7964
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Receipt"
            Key             =   "Receipt"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Payment"
            Key             =   "Payment"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "Cash receipt Details"
      Height          =   3855
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   900
      Width           =   7875
      Begin VB.CheckBox chkNew 
         Alignment       =   1  'Right Justify
         Caption         =   "New Account"
         Height          =   345
         Left            =   5310
         TabIndex        =   9
         Top             =   270
         Width           =   2385
      End
      Begin VB.Frame fraInterest 
         BorderStyle     =   0  'None
         Caption         =   "fraInt"
         Height          =   555
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   7515
         Begin WIS_Currency_Text_Box.CurrText txtMisc 
            Height          =   315
            Left            =   6420
            TabIndex        =   28
            Top             =   90
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtRegInt 
            Height          =   345
            Left            =   1950
            TabIndex        =   22
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtPenal 
            Height          =   345
            Left            =   4170
            TabIndex        =   25
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin VB.Label lblMisc 
            Caption         =   "Misceleneous"
            Height          =   255
            Left            =   5280
            TabIndex        =   27
            Top             =   150
            Width           =   975
         End
         Begin VB.Label lblPenalInt 
            Caption         =   "Penal Interest"
            Height          =   255
            Left            =   3060
            TabIndex        =   24
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label lblRegInt 
            Caption         =   "Regular interest"
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   400
         Index           =   0
         Left            =   6570
         TabIndex        =   33
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optAccount 
         Caption         =   "To Account"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3060
         Value           =   -1  'True
         Width           =   2505
      End
      Begin VB.TextBox txtScroll 
         Height          =   345
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txtCustName 
         Height          =   345
         Index           =   0
         Left            =   4380
         TabIndex        =   13
         Top             =   1350
         Width           =   2985
      End
      Begin VB.CommandButton cmdCustName 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   7440
         TabIndex        =   14
         Top             =   1350
         Width           =   285
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   810
         Width           =   5715
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   5250
         TabIndex        =   32
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optPassing 
         Caption         =   "To Passing Officer"
         Height          =   315
         Left            =   3060
         TabIndex        =   31
         Top             =   3090
         Width           =   2235
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   0
         Left            =   5940
         TabIndex        =   19
         Top             =   1890
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance"
         Height          =   285
         Index           =   0
         Left            =   4380
         TabIndex        =   18
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblName 
         Caption         =   "Cust Name :"
         Height          =   225
         Index           =   0
         Left            =   3180
         TabIndex        =   12
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblScroll 
         Caption         =   "Scroll No:"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account NO"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label lblAccType 
         Caption         =   "Account type"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   810
         Width           =   1755
      End
   End
   Begin VB.Label lblTransDate 
      Caption         =   "Trans Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmCashTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents m_frmLookUp As frmLookUp
'Private WithEvents m_frmCashIndex As frmCashIndex

Private m_retVar As Variant
'Private m_CashIndex As Boolean
'Private m_rstPayment As Recordset
'Const iReceipt = 0
'Const iPayment = 1
'Private m_VoucherType As Integer

Public Event AcceptClicked(VoucherType As Wis_VoucherTypes, Cancel As Integer)
Public Event CancelClicked(VoucherType As Wis_VoucherTypes)
Public Event LookUpClick(VoucherType As Wis_VoucherTypes)
Public Event SelectPaymentVoucher()

Private Sub EnableAccept()

Dim I As Byte

I = TabStrip1.SelectedItem.Index - 1

If cmbAccType(I).ListIndex < 0 Then Exit Sub
If Trim(txtScroll(I)) = "" Then Exit Sub
If Trim(txtAccNo(I)) = "" Then Exit Sub
If txtAmount(I) = 0 Then Exit Sub

cmdAccept(I).Enabled = True


End Sub

Private Sub chkNew_Click()
Dim Bool As Boolean

Bool = IIf(chkNew.Value = vbChecked, False, True)

lblAccNo(0).Enabled = Bool
txtAccNo(0).Enabled = Bool
lblName(0).Enabled = Bool
txtCustName(0).Enabled = Bool

End Sub

Private Sub cmbAccType_Click(Index As Integer)

'Only For Receipt
If Index Then fraInterest.Enabled = False: Exit Sub

If cmbAccType(Index).ListIndex < 0 Then Exit Sub

If Not optAccount Then fraInterest.Enabled = False: Exit Sub

Dim AccHeadID As Long
Dim ModuleId As wisModules
AccHeadID = cmbAccType(Index).ItemData(cmbAccType(Index).ListIndex)

ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

fraInterest.Enabled = Val(fraInterest.Tag)
lblRegInt = GetResourceString(344)
lblAccNo(0).Caption = GetResourceString(36, 60)

If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True
ElseIf ModuleId = wis_DepositLoans Then
    txtPenal.Enabled = False
    txtMisc.Enabled = False
ElseIf ModuleId = wis_Loans Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True

ElseIf ModuleId = wis_PDAcc Then
    lblAccNo(0).Caption = GetResourceString(330, 60)
ElseIf ModuleId = wis_Members Then
    
    lblRegInt = GetResourceString(53, 191) 'Share Fee
    txtPenal.Enabled = False
    txtMisc.Enabled = False
    chkNew.Enabled = chkNew.Tag
Else
    fraInterest.Enabled = False
    chkNew.Enabled = chkNew.Tag
End If

lblPenalInt.Enabled = txtPenal.Enabled
lblMisc.Enabled = txtMisc.Enabled

fraInterest.Visible = fraInterest.Enabled

End Sub
Private Sub cmdAccept_Click(Index As Integer)

If Not DateValidate(txtTransDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Sub
End If

If Trim$(txtScroll(Index)) = "" Then
    'MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    MsgBox "Please specify the Scroll NO", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtScroll(Index)
     Exit Sub
End If

Dim Cancel As Integer

RaiseEvent AcceptClicked(IIf(Index, payment, Receipt), Cancel)

'if cancel then
End Sub

Private Sub cmdCancel_Click(Index As Integer)

cmbAccType(Index).ListIndex = -1
txtAccNo(Index).Text = ""
txtCustName(Index).Text = ""
txtAmount(Index) = 0
txtBalance(Index) = ""

txtRegInt = 0
txtPenal = 0
txtMisc = 0
fraInterest.Visible = False

'If Index Then Call LoadPaymentVoucher

End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCustName_Click(Index As Integer)

RaiseEvent LookUpClick(IIf(Index, payment, Receipt))

Call txtAccNo_LostFocus(Index)

End Sub

Private Sub cmdDate_Click()
With Calendar
    .Left = cmdDate.Left + cmdDate.Width
    .Top = cmdDate.Top - .Height / 2
    .selDate = IIf(DateValidate(txtTransDate, "/", True), txtTransDate, gStrDate)
    .Show 1
    txtTransDate = .selDate
End With

End Sub

Private Sub cmdPay_Click()
txtAccNo(1).Tag = "0"

RaiseEvent SelectPaymentVoucher
If Val(txtAccNo(1).Tag) = 0 Then Exit Sub

If Val(txtAccNo(1).Tag) = 0 Then Exit Sub
'call get
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
If Not CtrlDown Then Exit Sub
If KeyCode <> vbKeyTab Then Exit Sub
        Dim I As Byte
With TabStrip1
    I = .SelectedItem.Index
    I = I - 1
    If I = 0 Then I = .Tabs.count
    .Tabs(I).Selected = True
End With


End Sub

Private Sub LoadAccountBalance()

Dim I As Byte
I = TabStrip1.SelectedItem.Index - 1

If cmbAccType(I).ListIndex < 0 Then Exit Sub

Dim ret As Integer
Dim BalanceAmount As Currency
Dim AccHeadID As Long
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim ModuleId As wisModules

'Check for the Selected AccountType if Account type is Sb then
'get the SB balance and if it is CA then Get the CA Balance-siddu
Dim LstInd As Long

With cmbAccType(I)
    AccHeadID = .ItemData(.ListIndex)
End With

ModuleId = GetModuleIDFromHeadID(AccHeadID)

'Check for the Account tyeps
'If it is Current or savings Account then get the Balance
If ModuleId = wis_CAAcc Then
    gDbTrans.SqlStmt = "SELECT AccID FROM CAMaster " & _
        "WHERE AccNum = " & AddQuotes(Trim$(txtAccNo(I).Text), True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        'MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        If ActiveControl.name = cmdAccept(0).name Or ActiveControl.name <> cmdClose.name Then Exit Sub
        ActivateTextBox txtAccNo(I)
        Exit Sub
    End If
   
    AccId = FormatField(rst("AccID"))
    
    'get teh TotalBalance Form the Particular Account
    gDbTrans.SqlStmt = "Select TOP 1 Balance from CATrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox "No Records"
    Else
        BalanceAmount = FormatField(rst(0))
        txtBalance(I) = BalanceAmount
        Set rst = Nothing
    End If
End If

'if the Account is SB Get the Sb Account balance Given from the AccountId
If ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    'AccId=
    gDbTrans.SqlStmt = "Select * from SBMaster" & _
                " where AccNum = " & AddQuotes(txtAccNo(I), True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        txtAmount(I).Text = ""
        txtAccNo(I).SetFocus
        Exit Sub
    End If
    
    AccId = FormatField(rst("AccId"))
    gDbTrans.SqlStmt = "Select TOP 1 Balance from SBTrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        BalanceAmount = FormatField(rst(0))
        txtBalance(I) = BalanceAmount
        Set rst = Nothing
    End If
    Exit Sub
End If

If I Then Exit Sub

If Not DateValidate(txtTransDate, "/", True) Then Exit Sub

Set rst = GetAccRecordSet(AccHeadID, txtAccNo(I))
If rst Is Nothing Then Exit Sub

AccId = FormatField(rst("ID"))

Dim TransDate As Date
Dim ClsObject As Object

TransDate = GetSysFormatDate(txtTransDate)

If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)
If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    Set ClsObject = New clsBkcc
    txtRegInt = ClsObject.RegularInterest(AccId, TransDate)
    txtPenal = ClsObject.PenalInterest(AccId, TransDate)
End If
If ModuleId = wis_DepositLoans Then
    Set ClsObject = New clsDepLoan
    txtRegInt = ClsObject.RegularInterest(AccId, TransDate)
End If
If ModuleId = wis_Loans Then
    Set ClsObject = New clsLoan
    txtRegInt = ClsObject.RegularInterest(AccId, , TransDate)
    txtPenal = ClsObject.PenalInterest(AccId, , TransDate)
End If

Set ClsObject = Nothing

End Sub

Private Sub Form_Load()

'First cetre the form
Call CenterMe(Me)

'Now Set the kannad action to all controls
Call SetKannadaCaption

'Now Load the Account Type
Call LoadAccountType

'Check Whether cas index is required
Dim SetUp As New clsSetup
'm_CashIndex = IIf(UCase(SetUp.ReadSetupValue("General", "CashIndex", "False")) = "FALSE", False, True)
optPassing.Enabled = IIf(UCase(SetUp.ReadSetupValue("General", "Passing", "False")) = "FALSE", False, True)
optAccount = True
Set SetUp = Nothing

txtTransDate = gStrDate

txtTransDate.Locked = gOnLine
cmdDate.Enabled = Not gOnLine

fraInterest.Enabled = False
fraInterest.Visible = False

'Refresh The Payment voucher

End Sub

Private Sub LoadAccountType()

Dim rstDeposit As Recordset
Dim rstLoan As Recordset

gDbTrans.SqlStmt = "Select * FROM DepositName"
Call gDbTrans.Fetch(rstDeposit, adOpenStatic)
gDbTrans.SqlStmt = "Select * FROM LoanScheme"
Call gDbTrans.Fetch(rstLoan, adOpenStatic)

'Dim ModuleId As wisModules
Dim AccHeadID As Long
Dim ClsBank As clsBankAcc
Dim AccName As String

Set ClsBank = New clsBankAcc

With cmbAccType(0)
    
    'Withdrawing of the amount will be only from the
    'account type where cheque facility is given
    'Such type of account are only two(three) types
    'They are
    '1) Saving Bank Account
    '2) Current Accounts
    '3) Od Account 'Presetly no such accounts are in our s/w

    .Clear
    
    AccName = GetResourceString(53, 36) 'Share Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(421) '"SB Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    AccName = GetResourceString(422) '"CA AccOUnt
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    AccName = GetResourceString(424)  '"RD Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(425) '"Pigmy Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    'Now Load All othe Deposits
    'ModuleId = wis_Deposits
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID 'ModuleId + FormatField(RstDeposit("DepositID"))
            rstDeposit.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 43) '"BKCC
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    'Now Load All Loans
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        While Not rstLoan.EOF
            AccName = FormatField(rstLoan("SchemeName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID
            rstLoan.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 58) '"BKCC Loan
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    'ModuleId = wis_DepositLoans
    AccName = GetResourceString(43, 58)
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    '.AddItem AccName
    '.ItemData(.NewIndex) = AccHeadID 'ModuleId
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName")) & " " & GetResourceString(58)
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID
            rstDeposit.MoveNext
        Wend
    End If
    
End With


With cmbAccType(1)
    .Clear
    
    AccName = GetResourceString(53, 36) 'Share Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(421) '"SB Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(422) '"CA AccOUnt
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(424)  '"RD Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    AccName = GetResourceString(425) '"Pigmy Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    'Now Load All othe Deposits
    'ModuleId = wis_Deposits
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID 'ModuleId + FormatField(RstDeposit("DepositID"))
            rstDeposit.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 43) '"BKCC
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    'Now Load All Loans
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        While Not rstLoan.EOF
            AccName = FormatField(rstLoan("SchemeName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID
            rstLoan.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 58) '"BKCC Loan
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    'ModuleId = wis_DepositLoans
    AccName = GetResourceString(43, 58)
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    '.AddItem AccName
    '.ItemData(.NewIndex) = AccHeadID 'ModuleId
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName")) & " " & GetResourceString(58)
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID
            rstDeposit.MoveNext
        Wend
    End If
    
End With


End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
   
'TransCtion Frame
lblTransDate = GetResourceString(38, 37)
TabStrip1.Tabs(1).Caption = GetResourceString(196)
TabStrip1.Tabs(2).Caption = GetResourceString(197)

lblScroll(0).Caption = GetResourceString(41)
lblAccType(0).Caption = GetResourceString(36) & " " & _
                                GetResourceString(35)
lblAccNo(0).Caption = GetResourceString(36) & " " & _
                            GetResourceString(60)
lblAmount(0).Caption = GetResourceString(40)
lblRegInt = GetResourceString(344)
lblPenalInt = GetResourceString(345)
lblMisc = GetResourceString(327)
chkNew.Caption = GetResourceString(260, 36)
cmdAccept(0).Caption = GetResourceString(4)
cmdCancel(0).Caption = GetResourceString(2)

'fra(1).Caption = GetResourceString(108)
lblScroll(1).Caption = GetResourceString(41)
lblAccType(1).Caption = GetResourceString(36) & " " & _
                GetResourceString(35)
lblAccNo(1).Caption = GetResourceString(36) & " " & _
                GetResourceString(60)
lblAmount(1).Caption = GetResourceString(40)
cmdAccept(1).Caption = GetResourceString(4)
cmdCancel(1).Caption = GetResourceString(2)




End Sub




Private Sub m_frmCashIndex_CancelClicked()
m_retVar = "0"
End Sub

Private Sub m_frmCashIndex_OKClicked()
m_retVar = "1"
End Sub

Private Sub optAccount_Click()
With fraInterest
   If .Enabled Then .Visible = True
End With
End Sub

Private Sub optPassing_Click()
    fraInterest.Visible = False
End Sub


Private Sub TabStrip1_Click()

Dim I As Byte

I = TabStrip1.SelectedItem.Index - 1

fra(I).Visible = True
fra(I).ZOrder 0

fra(Abs(I - 1)).Visible = False

End Sub

Private Sub txtAccNo_LostFocus(Index As Integer)

'Check for the Account Number
If txtAccNo(Index).Text = "" Then Exit Sub
'Now Get the Account No
Dim rst As Recordset
Dim AccHeadID As Long
If cmbAccType(Index).ListIndex < 0 Then Exit Sub

AccHeadID = cmbAccType(Index).ItemData(cmbAccType(Index).ListIndex)
txtCustName(Index) = ""
Dim ModuleId As wisModules
ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId = wis_PDAcc And Index = 0 Then
    gDbTrans.SqlStmt = "Select A.UserID, Permissions,Title +' '+FirstName +' '+" & _
        "MiddleName +' '+ LastName As CustName " & _
        " From UserTab A,NameTab B" & _
        " Where A.CustomerID = B.CustomerID And A.UserID = " & Val(txtAccNo(Index))
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Set rst = Nothing: Exit Sub
    If rst("Permissions") And perPigmyAgent = 0 Then Set rst = Nothing
    txtCustName(Index) = FormatField(rst("CustName"))
    Exit Sub
Else
    Set rst = GetAccRecordSet(AccHeadID, Trim(txtAccNo(Index)))
End If

If rst Is Nothing Then Exit Sub
txtCustName(Index) = FormatField(rst("CustName"))

Call LoadAccountBalance

Call EnableAccept

End Sub


Private Sub txtAmount_Change(Index As Integer)
Call EnableAccept
End Sub

