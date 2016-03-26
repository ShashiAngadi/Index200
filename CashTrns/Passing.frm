VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmPassing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Passing Officer Wibndow"
   ClientHeight    =   4905
   ClientLeft      =   540
   ClientTop       =   1875
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTransDate 
      Height          =   315
      Left            =   2010
      TabIndex        =   27
      Top             =   30
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   6960
      TabIndex        =   11
      Top             =   4470
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Payment Details"
      Height          =   3945
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   8145
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Left            =   1980
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   750
         Width           =   6015
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   1980
         TabIndex        =   14
         Top             =   1230
         Width           =   1215
      End
      Begin VB.CommandButton cmdPay 
         Caption         =   ".."
         Height          =   315
         Left            =   3630
         TabIndex        =   3
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtScroll 
         Height          =   345
         Left            =   1980
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
      Begin VB.Frame fraInterest 
         BorderStyle     =   0  'None
         Caption         =   "fraInt"
         Height          =   1035
         Left            =   60
         TabIndex        =   12
         Top             =   2790
         Width           =   8025
         Begin WIS_Currency_Text_Box.CurrText txtPenal 
            Height          =   345
            Left            =   1920
            TabIndex        =   19
            Top             =   600
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtMisc 
            Height          =   345
            Left            =   6270
            TabIndex        =   17
            Top             =   570
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtRegInt 
            Height          =   345
            Left            =   6270
            TabIndex        =   21
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtPrinAmount 
            Height          =   345
            Left            =   1920
            TabIndex        =   26
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin VB.Label lblPrinAmount 
            Caption         =   "Principal  Amount "
            Height          =   255
            Left            =   180
            TabIndex        =   25
            Top             =   150
            Width           =   1725
         End
         Begin VB.Label lblRegInt 
            Caption         =   "Regular interest"
            Height          =   285
            Left            =   4350
            TabIndex        =   20
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label lblPenalInt 
            Caption         =   "Penal Interest"
            Height          =   255
            Left            =   150
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblMisc 
            Caption         =   "Misceleneous"
            Height          =   285
            Left            =   4320
            TabIndex        =   16
            Top             =   630
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdAccNo 
         Caption         =   ".."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3300
         TabIndex        =   2
         Top             =   1230
         Width           =   315
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   1980
         TabIndex        =   23
         Top             =   2250
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtCustNAme 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1980
         TabIndex        =   15
         Top             =   1740
         Width           =   5985
      End
      Begin VB.Label txtTrans 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   345
         Left            =   6030
         TabIndex        =   24
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lblTrans 
         Caption         =   "Label3"
         Height          =   255
         Left            =   4380
         TabIndex        =   29
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   285
         Left            =   210
         TabIndex        =   22
         Top             =   2280
         Width           =   1725
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account NO"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblAccType 
         Caption         =   "Account type"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lblScroll 
         Caption         =   "Scroll No:"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label lblName 
         Caption         =   "Cust Name :"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance"
         Height          =   285
         Left            =   4380
         TabIndex        =   5
         Top             =   2370
         Width           =   1725
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6270
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   400
      Left            =   5670
      TabIndex        =   10
      Top             =   4470
      Width           =   1215
   End
   Begin VB.Label lblTransDate 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmPassing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private m_retVar As Variant
Private m_Rst As Recordset

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

With cmbAccType
    
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

End Sub


Private Sub cmbAccType_Click()

If m_Rst Is Nothing Then Exit Sub
If m_Rst.EOF Or m_Rst.BOF Then Exit Sub
If cmbAccType.ListIndex < 0 Then Exit Sub

Dim transType As wisTransactionTypes

transType = FormatField(m_Rst("TransType"))


Dim AccHeadID As Long
Dim ModuleId As wisModules
AccHeadID = cmbAccType.ItemData(cmbAccType.ListIndex)

ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

fraInterest.Enabled = True
lblRegInt = GetResourceString(344)
lblAccNo.Caption = GetResourceString(36, 60)

If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True
ElseIf ModuleId = wis_DepositLoans Then
    txtPenal.Enabled = False
    txtMisc.Enabled = False
ElseIf ModuleId = wis_Loans Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True

ElseIf ModuleId = wis_PDAcc And transType = wDeposit Then
    lblAccNo.Caption = GetResourceString(330, 60)
ElseIf ModuleId = wis_Members Then
    lblRegInt = GetResourceString(53, 191) 'Share Fee
    txtPenal.Enabled = False
    txtMisc.Enabled = False
Else
    fraInterest.Enabled = False
End If

lblPenalInt.Enabled = txtPenal.Enabled
lblMisc.Enabled = txtMisc.Enabled

fraInterest.Visible = fraInterest.Enabled

If transType = wWithdraw Then
    fraInterest.Enabled = False
    fraInterest.Visible = False
End If


End Sub

Private Sub cmdAccept_Click()

If m_Rst Is Nothing Then Exit Sub
If m_Rst.EOF Then Exit Sub

Dim transType As wisTransactionTypes
transType = m_Rst("TransType")
If transType = wWithdraw Then
    Call CashPaid
Else
    Call CashReceived
End If


End Sub

Private Sub cmdCustName_Click()

If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

Dim I As Byte


Dim headID As wisModules
With cmbAccType
    If .ListIndex < 0 Then Exit Sub
    headID = .ItemData(.ListIndex)
End With

On Error GoTo Exit_Line

Dim SqlStr As String

Dim RstCust As Recordset


If GetModuleIDFromHeadID(headID) = wis_PDAcc Then
    SqlStr = "Select UserID as AgentNum, UserId as ID From UserTab A,NameTab B " & _
        " Where A.CustomerID = B.CustomerID"
    Call gDbTrans.Fetch(RstCust, adOpenDynamic)
Else
    Set RstCust = GetAccRecordSet(headID, , IIf(Trim(txtAccNo) = "", txtCustName, ""))
End If
    

If RstCust Is Nothing Then
    MsgBox "There are no customers in the " & GetHeadName(headID), _
            vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

Screen.MousePointer = vbHourglass
If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

m_retVar = ""
With m_LookUp
    Call FillView(.lvwReport, RstCust, True)
    .lvwReport.ColumnHeaders(2).Width = 0
    .Show 1
End With

MousePointer = vbDefault

txtAccNo = m_retVar
Screen.MousePointer = vbDefault
    
    
Exit_Line:

    MousePointer = vbDefault

End Sub

Private Sub cmdAccNo_Click()

If cmbAccType.ListIndex < 0 Then Exit Sub

Dim ModuleId As wisModules
Dim SqlStr As String

ModuleId = GetModuleIDFromHeadID(cmbAccType.ItemData(cmbAccType.ListIndex))
If ModuleId = wis_None Then Exit Sub

If ModuleId = wis_Members Then
    SqlStr = "Select AccNum, " & _
        "Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
        "From MemMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
            " ANd AccID Not IN (Select Distinct AccID From MemTrans C)"
    
ElseIf ModuleId = wis_CAAcc Then 'Current Account
    SqlStr = "Select A.AccID as ID,AccNum, " & _
        "Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
        "From CAMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
            " ANd AccID Not IN (Select Distinct AccID From CATrans C)"

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    Dim subDepType As Integer
    subDepType = ModuleId Mod 100 'wis_Deposits
    
    'If subDepType >= wisDeposit_RD And subDepType < wisDeposit_RD + 10 Then
        SqlStr = "Select AccNum, " & _
            ", Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            "From RDMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
                " ANd AccID Not IN (Select Distinct AccID From FDTrans C)"
        If subDepType > wisDeposit_RD Then _
            SqlStr = SqlStr & " ANd A.DepositType = " & subDepType - 10
    
    'ElseIf subDepType >= wisDeposit_PD And subDepType < wisDeposit_PD + 10 Then
        SqlStr = "Select AccNum, " & _
            ", Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            "From PDMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
                " ANd AccID Not IN (Select Distinct AccID From FDTrans C)"
        If subDepType > wisDeposit_PD Then _
            SqlStr = SqlStr & " ANd A.DepositType = " & subDepType - wis_Deposits
   'Else
   
   
        SqlStr = "Select AccNum, " & _
            ", Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            "From FDMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
                " ANd AccID Not IN (Select Distinct AccID From FDTrans C)"
        If ModuleId > wis_Deposits Then _
            SqlStr = SqlStr & " ANd A.DepositType = " & ModuleId - wis_Deposits
    'End If
    

ElseIf ModuleId >= wis_PDAcc And ModuleId < wis_PDAcc + 10 Then
    'Pigmy Accounts
    SqlStr = "Select AccNum, AccId as ID From PDMaster A"

ElseIf ModuleId >= wis_RDAcc And ModuleId < wis_RDAcc + 10 Then  'Recurring Accounts
    SqlStr = "Select AccNum, " & _
        " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
        "From RDMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
            " ANd AccID Not IN (Select Distinct AccID From RDTrans )"

ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 10 Then
    SqlStr = "Select AccNum, " & _
        "Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
        "From SBMaster A,NameTab B Where A.CustomerID = B.CustomerID " & _
            " ANd AccID Not IN (Select Distinct AccID From SBTrans )"

End If

gDbTrans.SqlStmt = SqlStr & " ORDER By FirstName"

Dim rst As Recordset

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    MsgBox "There are no New accounts in the " & _
        cmbAccType.Text, vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp
Call FillView(m_LookUp.lvwReport, rst)

m_retVar = ""
m_LookUp.Show 1

If Len(m_retVar) > 0 Then
    txtAccNo = m_retVar
    'call txta
End If

End Sub

Private Sub cmdPay_Click()
Call GetNewRecordset(m_Rst)
If m_Rst Is Nothing Then Exit Sub

If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp
If Not FillViewNew(m_LookUp.lvwReport, m_Rst, "CashTransID") Then Exit Sub

m_LookUp.Show 1
If m_retVar = "" Then Exit Sub

m_Rst.MoveFirst
m_Rst.Find ("CashTransID = " & m_retVar)
If m_Rst.EOF Then Exit Sub

Call FillDetails

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
'optPassing.Enabled = IIf(UCase(SetUp.ReadSetupValue("General", "Passing", "False")) = "FALSE", False, True)
'optAccount = True
Set SetUp = Nothing

txtTransDate = gStrDate

txtTransDate.Locked = gOnLine

fraInterest.Enabled = False
fraInterest.Visible = False

'Refresh The Payment voucher
Call GetNewRecordset(m_Rst)
If Not m_Rst Is Nothing Then Call FillDetails

End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
   
'TransCtion Frame
lblTransDate = GetResourceString(38, 37)
fra.Caption = GetResourceString(38, 295)

lblScroll.Caption = GetResourceString(41)
lblAccType.Caption = GetResourceString(36) & " " & _
                                GetResourceString(35)
lblAccNo.Caption = GetResourceString(36) & " " & _
                            GetResourceString(60)
lblAmount.Caption = GetResourceString(40)
lblRegInt = GetResourceString(344)
lblPenalInt = GetResourceString(345)
lblMisc = GetResourceString(327)
'chkNew.Caption = GetResourceString(260,36)

cmdAccept.Caption = GetResourceString(4)
cmdCancel.Caption = GetResourceString(2)
lblTrans = GetResourceString(38)
txtTrans = ""

End Sub


Private Function CashPaid() As Boolean

Dim VoucherNo As String
Dim AccHeadID As Long
Dim AccNum As String
Dim ScrollNo As String

Dim I As Integer

On Error GoTo Exit_Line

Dim TransDate As Date

    TransDate = GetSysFormatDate(txtTransDate)
    AccHeadID = cmbAccType.ItemData(cmbAccType.ListIndex)
    AccNum = txtAccNo
    ScrollNo = txtScroll
    
    Dim rst As Recordset
    Dim AccId As Long
    
    Set rst = GetAccRecordSet(AccHeadID, AccNum)
    If rst Is Nothing Then Exit Function
    AccId = FormatField(rst("Id"))

Dim transType As wisTransactionTypes
Dim CashTransID As Long
Dim TransID As Long
Dim RecordPos As Long
Dim TransUserID As Long

CashTransID = FormatField(m_Rst("CashTransID"))
TransID = FormatField(m_Rst("TransID"))
TransUserID = FormatField(m_Rst("TransUserID"))
RecordPos = m_Rst.AbsolutePosition
transType = FormatField(m_Rst("TransType"))

'Get the Balance of the tranacting account
'Balance = Val(txtBalance(I))

gDbTrans.BeginTrans

'Now Update the CashTrans as this amount is paid

gDbTrans.SqlStmt = "UPDate CashTrans Set PassUSerID = " & gUserID & "," & _
        " Where AccHeadID = " & AccHeadID & " AND AccID = " & AccId

If transType = wWithdraw Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND TransID = " & TransID
Else
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND TransUserID = " & TransUserID
End If

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

CashPaid = True
 
Exit_Line:

End Function


Private Function CashReceived() As Boolean

Dim VoucherNo As String
Dim AccHeadID As Long
Dim AccNum As String

Dim TransDate As Date
Dim ScrollNo As String

On Error GoTo Exit_Line

TransDate = GetSysFormatDate(txtTransDate)
AccHeadID = cmbAccType.ItemData(cmbAccType.ListIndex)

AccNum = txtAccNo
ScrollNo = txtScroll
AccHeadID = cmbAccType.ItemData(cmbAccType.ListIndex)

Dim rst As Recordset
Dim AccId As Long

AccId = FormatField(rst("Id"))

If AccId < 0 Then  ' Id Account Id smaller than one then
                    'If indicates that the it is new account transaction '
                    'So Select the New Account
    If txtAccNo = "" Then Call cmdAccNo_Click
    If txtAccNo = "" Then
        MsgBox "Select the new account no for this transaction ", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    Set rst = GetAccRecordSet(AccHeadID, AccNum)
    If rst Is Nothing Then Exit Function
    AccId = FormatField(rst("Id"))
End If


Dim transType As wisTransactionTypes

transType = wDeposit

Dim CashTransID As Long
'Get the Max TransCtion ID
gDbTrans.SqlStmt = "Select Max(CashTransID) From CashTrans " & _
        " Where TransDate >= #" & FinUSFromDate & "#"

CashTransID = 1
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then CashTransID = FormatField(rst(0)) + 1

Dim TransID As Long
Dim CounterId As Integer
Dim Balance As Currency
Dim ModuleId As wisModules

'Get the Balance of the tranacting account
'Balance = Val(txtBalance(I))


'Dim BankClass As clsBankAcc
'Set BankClass = New clsBankAcc
Dim ClsObject As Object


gDbTrans.BeginTrans

Dim IntAmount As Currency
Dim PenalAmount As Currency
Dim PrincAmount As Currency
Dim MiscAmount As Currency

PrincAmount = txtPrinAmount
If fraInterest.Visible Then
    IntAmount = txtRegInt
    PenalAmount = txtPenal
    MiscAmount = txtMisc
End If

ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

If ModuleId = wis_PDAcc Then

    'Do nothing
    If ModuleId = wis_PDAcc Then Set ClsObject = New clsPDAcc
    TransID = ClsObject.DepositAgentAmount(AccId, PrincAmount, _
                    "From Cash Counter " & CounterId, TransDate, VoucherNo, True)
    
ElseIf ModuleId = wis_CAAcc Or ModuleId = wis_SBAcc Or _
    ModuleId = wis_Deposits Or ModuleId = wis_RDAcc Then

    If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then Set ClsObject = New clsBkcc
    If ModuleId = wis_SBAcc Then Set ClsObject = New clsSBAcc
    If ModuleId = wis_Deposits Then Set ClsObject = New clsFDAcc
    If ModuleId = wis_PDAcc Then Set ClsObject = New clsPDAcc
    If ModuleId = wis_CAAcc Then Set ClsObject = New ClsCAAcc
    If ModuleId = wis_RDAcc Then Set ClsObject = New clsRDAcc

    TransID = ClsObject.DepositAmount(AccId, PrincAmount, _
                     "From Cash Counter " & CounterId, TransDate, VoucherNo, True)
    If TransID <= 1 Then GoTo Exit_Line
    Set ClsObject = Nothing

ElseIf ModuleId = wis_DepositLoans Or ModuleId = wis_BKCC Or _
    ModuleId = wis_BKCCLoan Or ModuleId = wis_Loans Or ModuleId = wis_Members Then
    'Now search whether he is Paying any Interest Amount
    'on this loan account
    
    If ModuleId = wis_DepositLoans Then
        Set ClsObject = New clsDepLoan
        TransID = ClsObject.DepositAmount(AccId, PrincAmount, _
            IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
        
    ElseIf ModuleId = wis_Loans Or ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
        If ModuleId = wis_Loans Then Set ClsObject = New clsLoan
        If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then Set ClsObject = New clsBkcc
        
        TransID = ClsObject.DepositAmount(CLng(AccId), PrincAmount, _
            IntAmount, PenalAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
       
    ElseIf ModuleId = wis_Members Then
        Set ClsObject = New clsMMAcc
        TransID = ClsObject.DepositAmount(CLng(AccId), PrincAmount, _
                    IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
        
    End If
    
    Set ClsObject = Nothing
End If

gDbTrans.SqlStmt = "UPDate CashTrans Set PassUserID = " & gUserID & _
    " Where CashTransID = " & m_Rst("CashTransID") & _
    " And AccHeadID = " & m_Rst("AccHeadID") & _
    " And AccID = " & AccId & _
    " And TransType = " & transType


If Not gDbTrans.SQLExecute Then
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.CommitTrans

CashReceived = True

Exit_Line:

End Function

Private Sub FillDetails()

If m_Rst Is Nothing Then Exit Sub
If m_Rst.EOF Then Exit Sub

Dim AccHeadID As Long
Dim AccId As Long
Dim transType As wisTransactionTypes

txtScroll = FormatField(m_Rst("ScrollNo"))
txtAmount = FormatField(m_Rst("Amount"))
transType = FormatField(m_Rst("TransType"))

txtTrans.Caption = GetResourceString(IIf(transType = wDeposit, 196, 197))

AccHeadID = FormatField(m_Rst("AccHeadID"))
AccId = FormatField(m_Rst("AccID"))
txtAccNo.Tag = FormatField(m_Rst("AccID"))

Dim I As Integer
Dim MaxI As Integer

fraInterest.Visible = CBool(transType = wDeposit)
fraInterest.Enabled = CBool(transType = wDeposit)

With cmbAccType
    MaxI = .ListCount - 1
    For I = 0 To MaxI
        If .ItemData(I) = AccHeadID Then
            .ListIndex = I
            Exit For
        End If
    Next
End With

If AccId > 0 Then
    'Now Get the AccountNo And Customer Name
    Call GetSetCustomerName(AccHeadID, AccId)
    cmdAccNo.Enabled = False
Else
    cmdAccNo.Enabled = True
End If

End Sub

Private Function GetSetCustomerName(AccHeadID As Long, ByVal AccId As Long) As Recordset

txtCustName = ""
txtAccNo = ""

On Error Resume Next

Dim rstReturn As Recordset
Dim SqlStr As String
Dim pos As Long
Dim sqlClause As String

Set rstReturn = Nothing
On Error GoTo Exit_Line

Dim ModuleId As wisModules

ModuleId = GetModuleIDFromHeadID(AccHeadID)

Set rstReturn = Nothing

'Members
If ModuleId = wis_Members And ModuleId < wis_Members + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID, " & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From MemMaster A , NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And AccId = " & AccId
    
    'If ModuleId > wis_Members Then _
     '   sqlClause = " ANd A.MemberType = " & ModuleId - wis_Members

'BKcc Account
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    SqlStr = "Select AccNum, A.LoanID as ID, " & ""
    SqlStr = SqlStr & " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From BKCCMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    
'Current Account
ElseIf ModuleId = wis_CAAcc Then
    SqlStr = "Select AccNum, AccId as Id ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From CAMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

'Deposit Loans
ElseIf ModuleId >= wis_DepositLoans And ModuleId < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID "
    SqlStr = SqlStr & " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From DepositLoanMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    
    If ModuleId > wis_DepositLoans Then _
        sqlClause = " ANd A.DepositType = " & ModuleId - wis_DepositLoans

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From FDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId
    'If ModuleID > wis_Deposits Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_Deposits
'Loan Accounts
ElseIf ModuleId = wis_Loans Then
    SqlStr = "Select AccNum, LoanId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From LoanMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    If ModuleId > wis_Loans Then _
        sqlClause = " AND A.SchemeID = " & ModuleId - wis_Loans

'Pigmy Accounts
ElseIf ModuleId >= wis_PDAcc And ModuleId < wis_PDAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From PDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

'Recurring Accounts
ElseIf ModuleId >= wis_RDAcc And ModuleId < wis_RDAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From RDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From SBMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId
    
End If
    
    gDbTrans.SqlStmt = SqlStr & " " & sqlClause & " ORDER By FirstName"


If gDbTrans.Fetch(rstReturn, adOpenStatic) < 1 Then Exit Function

txtCustName = FormatField(rstReturn("custName"))
txtAccNo = FormatField(rstReturn("AccNUm"))

Exit_Line:


End Function

Private Sub EnableAccept()

cmdAccept.Enabled = False

If cmbAccType.ListIndex < 0 Then Exit Sub
If Trim(txtScroll) = "" Then Exit Sub
If Trim(txtAccNo) = "" Then Exit Sub
If txtAmount = 0 Then Exit Sub

cmdAccept.Enabled = True


End Sub
Private Sub LoadAccountBalance()

If cmbAccType.ListIndex < 0 Then Exit Sub

Dim ret As Integer
Dim BalanceAmount As Currency
Dim AccHeadID As Long
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim ModuleId As wisModules

'Check for the Selected AccountType if Account type is Sb then
'get the SB balance and if it is CA then Get the CA Balance-siddu
Dim LstInd As Long

With cmbAccType
    AccHeadID = .ItemData(.ListIndex)
End With

ModuleId = GetModuleIDFromHeadID(AccHeadID)

'Check for the Account tyeps
'If it is Current or savings Account then get the Balance
If ModuleId = wis_CAAcc Then
    gDbTrans.SqlStmt = "SELECT AccID FROM CAMaster " & _
        "WHERE AccNum = " & AddQuotes(Trim$(txtAccNo.Text), True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        'MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        If ActiveControl.name = cmdAccept.name Then Exit Sub
        ActivateTextBox txtAccNo
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
        txtBalance = BalanceAmount
        Set rst = Nothing
    End If
End If

'if the Account is SB Get the Sb Account balance Given from the AccountId
If ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    'AccId=
    gDbTrans.SqlStmt = "Select * from SBMaster" & _
                " where AccNum = " & AddQuotes(txtAccNo, True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        txtAmount.Text = ""
        txtAccNo.SetFocus
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
        txtBalance = BalanceAmount
        Set rst = Nothing
    End If
    Exit Sub
End If

If m_Rst("TransType") = wWithdraw Then Exit Sub

If Not DateValidate(txtTransDate, "/", True) Then Exit Sub

Set rst = GetAccRecordSet(AccHeadID, txtAccNo)

If rst Is Nothing Then Exit Sub

AccId = FormatField(rst("ID"))

Dim TransDate As Date
Dim ClsObject As Object

TransDate = GetSysFormatDate(txtTransDate.Text)

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

Private Sub LoadPaymentVoucher(Optional nextRecord As Boolean = True)

If m_Rst Is Nothing Then Exit Sub
If m_Rst.EOF Then Exit Sub

If nextRecord Then m_Rst.MoveNext

If m_Rst.EOF Then Exit Sub

Dim AccHeadID As Long
Dim AccId As Long
Dim ModuleId As wisModules
Dim AccNum As String
Dim TableName As String

AccHeadID = FormatField(m_Rst("AccheadID"))
AccId = FormatField(m_Rst("AccID"))

ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

If ModuleId = wis_CAAcc Then TableName = "CAMaster"
If ModuleId = wis_SBAcc Then TableName = "SBMaster"
If ModuleId = wis_Deposits Then TableName = "FDMaster"
If ModuleId = wis_PDAcc Then TableName = "PDMaster"
If ModuleId = wis_RDAcc Then TableName = "RDMaster"
If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then TableName = "BKCCMaster"
If ModuleId = wis_Members Then TableName = "MemMaster"

If ModuleId = wis_DepositLoans Then TableName = "DepositLoanMaster"
If ModuleId = wis_Loans Then TableName = "LoanMaster"

Debug.Assert TableName <> ""
If TableName = "" Then Exit Sub

gDbTrans.SqlStmt = "Select AccNum, " & _
    " Title +' '+FirstNAme+' '+MiddleName+' '+LastName As CustName" & _
    " From " & TableName & " A, NameTab B " & _
    " Where B.CustomerID = A.CustomerID "

If ModuleId = wis_DepositLoans Or ModuleId = wis_Loans Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
        " AND A.LoanID = " & m_Rst("AccID")
Else
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
        " AND A.AccID = " & m_Rst("AccID")
End If

Dim rstTemp As Recordset

If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Sub
Dim I As Integer
I = 1

txtAccNo = FormatField(rstTemp("AccNum"))
txtCustName = FormatField(rstTemp("CustName"))
txtAmount = FormatField(m_Rst("Amount"))
txtScroll = FormatField(m_Rst("ScrollNo"))

With cmbAccType
    I = 0
    Do
        If I = .ListCount Then Exit Do
        If AccHeadID = .ItemData(I) Then .ListIndex = I: Exit Do
        I = I + 1
    Loop
End With


End Sub

Private Sub GetNewRecordset(rst As Recordset, Optional transType As wisTransactionTypes)

Dim SqlStr As String

SqlStr = "Select CashTransID,AccID,HeadName,TransType, " & _
        " Amount,ScrollNo,AccHeadID From CashTrans A,Heads B Where " & _
        " B.HeadID = A.AccHeadid "

If transType = wWithdraw Or transType = wContraWithdraw Then
    SqlStr = SqlStr & " And TransType = " & transType & " And TransID > 0   " & _
        " And (TransUserID is NULL or TransUserID = 0) " & _
        " And (PassUserID is NULL or PassUserId = 0)" & _
        " order by TransDate Desc,ScrollNo"
        
ElseIf transType = wDeposit Or transType = wContraDeposit Then
    SqlStr = SqlStr & " And TransType = " & transType & " And TransUserID > 0 " & _
        " And (TransID is NULL or TransID = 0) " & _
        " And (PassUserID is NULL or PassUserId = 0)" & _
        " order by TransDate Desc,ScrollNo"

Else

    SqlStr = SqlStr & _
        " And (PassUserID is NULL or PassUserId = 0)" & _
        " order by TransDate Desc,ScrollNo"
        
End If


gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Set rst = Nothing


End Sub

Private Sub m_LoopUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub


