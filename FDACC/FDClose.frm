VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmFdClose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FDClose"
   ClientHeight    =   5685
   ClientLeft      =   1770
   ClientTop       =   1590
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   400
      Left            =   5850
      TabIndex        =   36
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   7260
      TabIndex        =   37
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fraDeposit 
      Caption         =   "Deposit Details"
      Height          =   3675
      Left            =   90
      TabIndex        =   10
      Top             =   420
      Width           =   4185
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3810
         TabIndex        =   2
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInterest 
         Height          =   315
         Left            =   2550
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1170
         Width           =   1185
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2550
         TabIndex        =   3
         Top             =   270
         Width           =   1185
      End
      Begin WIS_Currency_Text_Box.CurrText txtPayable 
         Height          =   345
         Left            =   2550
         TabIndex        =   9
         Top             =   1620
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepInterest 
         Height          =   345
         Left            =   2550
         TabIndex        =   12
         Top             =   2070
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotalInterest 
         Height          =   345
         Left            =   2550
         TabIndex        =   16
         Top             =   2520
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Line Line1 
         X1              =   4050
         X2              =   120
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Label lblNetAmount 
         Caption         =   "Net payable amount:"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   3120
         Width           =   2235
      End
      Begin VB.Label txtPayableAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2550
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label txtDepositAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2550
         TabIndex        =   5
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblIntOnDep 
         Caption         =   "Total Interest :"
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblDepInt 
         Caption         =   "Rate of Interest:"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1170
         Width           =   1935
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposited amount:"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lblPayable 
         Caption         =   "Payable :"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest :"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   2070
         Width           =   1935
      End
   End
   Begin VB.Frame fraCharges 
      Caption         =   "Charges"
      Height          =   1545
      Left            =   4350
      TabIndex        =   31
      Top             =   2550
      Width           =   4245
      Begin WIS_Currency_Text_Box.CurrText txtCharges 
         Height          =   345
         Left            =   3150
         TabIndex        =   26
         Top             =   450
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTax 
         Height          =   345
         Left            =   3150
         TabIndex        =   29
         Top             =   1050
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblOthers 
         Caption         =   "Other charges (Tax, etc.)"
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   1170
         Width           =   2235
      End
      Begin VB.Label lblPreClose 
         Caption         =   "Premature closure charges: "
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   570
         Width           =   2355
      End
   End
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   90
      TabIndex        =   23
      Top             =   4110
      Width           =   8505
      Begin VB.OptionButton optTransfer 
         Caption         =   "Transfer"
         Height          =   315
         Left            =   5850
         TabIndex        =   35
         Top             =   240
         Width           =   1995
      End
      Begin VB.OptionButton optMature 
         Caption         =   "Mature"
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   270
         Width           =   2355
      End
      Begin VB.OptionButton optClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   3255
         TabIndex        =   32
         Top             =   270
         Width           =   2265
      End
      Begin VB.CommandButton cmdTfr 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   34
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame fraLoanDet 
      Caption         =   "Loan Details"
      Height          =   2205
      Left            =   4350
      TabIndex        =   27
      Top             =   420
      Width           =   4245
      Begin VB.CheckBox chkDeductLoan 
         Alignment       =   1  'Right Justify
         Caption         =   "Deduct Loan Amount :"
         Height          =   285
         Left            =   210
         TabIndex        =   24
         Top             =   1740
         Width           =   3795
      End
      Begin VB.TextBox txtLoanRate 
         Height          =   315
         Left            =   2790
         TabIndex        =   20
         Top             =   810
         Width           =   1185
      End
      Begin WIS_Currency_Text_Box.CurrText txtLoanInterest 
         Height          =   345
         Left            =   2790
         TabIndex        =   22
         Top             =   1230
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtLoanAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2790
         TabIndex        =   18
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label lblLoanInt 
         Caption         =   "Loan Interest Rate"
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   810
         Width           =   2085
      End
      Begin VB.Label lblIntOnLoan 
         Caption         =   "Interest on loans: "
         Height          =   225
         Left            =   210
         TabIndex        =   21
         Top             =   1230
         Width           =   2175
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Total loan amount:"
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   390
         Width           =   2085
      End
   End
   Begin VB.Label lblName 
      Caption         =   "Customer Name"
      Height          =   345
      Left            =   120
      TabIndex        =   30
      Top             =   30
      Width           =   8205
   End
   Begin VB.Label lblLoanDet 
      Caption         =   "Loan details:"
      Height          =   225
      Left            =   -30
      TabIndex        =   0
      Top             =   1800
      Width           =   2115
   End
End
Attribute VB_Name = "frmFdClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_AccID As Long
Private m_Loaded As Boolean
'Private M_setUp As New clsSetup
Private m_AccHeadId As Long

Private m_DepositType As Integer
Private m_DepositName As String
Private m_DepositNameEnglish As String
Private m_Cumulative As Boolean

Private m_ContraClass As clsContra
Private m_AccType As wis_AmountType


Public Event FDClose()
Public Event WindowClosed()
Public Property Let AccountId(NewValue As Long)

    m_AccID = NewValue
    
If m_AccID <= 0 Then Exit Property
Dim rstTemp As Recordset

gDbTrans.SqlStmt = "Select A.DepositName,A.Cumulative,b.DepositType,A.DepositNameEnglish " & _
    " From DepositName A, FDMaster  B " & _
    " WHERE A.DepositId = B.DepositType " & _
    " And B.AccId = " & m_AccID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then m_AccID = 0: Exit Property
m_DepositType = FormatField(rstTemp("DepositType"))
m_Cumulative = FormatField(rstTemp("Cumulative"))
m_DepositName = FormatField(rstTemp("DepositName"))
m_DepositNameEnglish = FormatField(rstTemp("DepositNameEnglish"))
'Deposit Head
m_AccHeadId = GetIndexHeadID(m_DepositName)
If m_AccHeadId = 0 Then m_AccHeadId = GetHeadID(m_DepositName, parMemberDeposit)

If m_Loaded Then
    Call UpdateDetails
    If m_Cumulative Then _
        lblPayable.Caption = GetResourceString(47, 267)
        'Interest Paid
    txtPayable.Enabled = IIf(m_Cumulative, False, True)
End If
End Property

Private Function FDClose() As Boolean

Dim bankClass  As clsBankAcc
'Perform the transaction with closed flag and send the guy home
'Check date
Dim TransDate As Date
Dim TransID As Long
Dim Amount As Currency
Dim InTrans As Boolean

Dim rst As ADODB.Recordset

Dim PayableAmount As Currency
Dim PayableBalance As Currency
Dim transType As wisTransactionTypes
Dim InterestAmount As Currency
Dim TotalIntAmount As Currency

'Get the Transaction date
TransDate = GetSysFormatDate(txtDate.Text)

'Now Get the Interest Amount deposited in InterestPayble Account
PayableAmount = IIf(m_Cumulative, 0, txtPayable)
'PayableAmount = txtPayable
InterestAmount = txtDepInterest
TotalIntAmount = txtTotalInterest

gDbTrans.SqlStmt = "Select Balance From FDIntPayable " & _
                " Where Accid = " & m_AccID & _
                " ORDER By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                    PayableBalance = FormatField(rst(0))
Set rst = Nothing

'AMOUNT TO BE DEDUCTED FROM INTEREST PAYBLE ACCOUNT IS PayableAmount
PayableBalance = PayableBalance - PayableAmount
If PayableBalance < 0 Then
    'Witthdrawing more amount from payableaccount
    If MsgBox(GetResourceString(577) & vbCrLf & GetResourceString(541) & "?", _
        vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then GoTo ExitLine
    PayableBalance = 0
End If

'so the interest from the P & L account is
'AMOUNT TO BE DEDUCTED FROM PROFIT & LOSS ACCOUNT IS InterestAmount
'total interest - payble interest i.e. InterestAmount
If m_Cumulative Then
'    InterestAmount = txtTotalInterest - txtPayable - txtCharges
    Amount = Val(txtDepositAmount) + txtPayable
Else
'    InterestAmount = txtTotalInterest - PayableAmount - txtCharges
    Amount = Val(txtDepositAmount)
End If

'Get the Next TransCtionId
TransID = GetFDMaxTransID(m_AccID) + 1

'First Check whehter this transction is Contra Or Cash
'If Any Amount is transferring to Loan then also this is Contra
Dim boolContra As Boolean
If optTransfer Then boolContra = True

Set bankClass = New clsBankAcc
gDbTrans.BeginTrans: InTrans = True

'in the cumalative deposiut the interest already creditted to the account
'so there will be no transaction from payable account
If m_Cumulative Then PayableAmount = 0

'Transction of the interest Amount
'first Make the Transction To The Payable Account
If PayableAmount Then
    'Now withdraw the amount from Interest Paybale account
    'if he is repaying the loan amount or transferring to other account
    'then the transction will be contra
    transType = wWithdraw
    If boolContra Then transType = wContraWithdraw
    gDbTrans.SqlStmt = "Insert into FDIntPayable(AccID,TransID, " & _
                " TransType,TransDate,Amount,Balance, " & _
                " Particulars,UserID) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                transType & "," & _
                "#" & TransDate & "#," & _
                PayableAmount & "," & _
                PayableBalance & ", " & _
                "'Closing FD " & m_AccID & "' " & _
                "," & gUserID & " ) "
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    Dim PayableHeadID As Long
    Dim engHeadName As String
    If Len(m_DepositNameEnglish) > 0 Then _
            engHeadName = m_DepositNameEnglish & " " & LoadResourceStringS(375, 47)
    
    PayableHeadID = bankClass.GetHeadIDCreated(m_DepositName & " " & GetResourceString(375, 47), _
            engHeadName, parDepositIntProv, 0, wis_Deposits + m_DepositType)
    
    If Not boolContra Then
        If Not bankClass.UpdateCashWithDrawls(PayableHeadID, _
                        PayableAmount, TransDate) Then GoTo ErrLine
    End If
End If

'Now make the transction to the
'Interest Table
If InterestAmount Then
    transType = wWithdraw
    If boolContra Then transType = wContraWithdraw
    'If Calculated interest is smaller than the Interest Payable
    'Then take the Difference Amount as Profit
    If InterestAmount < 0 Then _
            transType = IIf(boolContra, wContraDeposit, wDeposit)
    
    gDbTrans.SqlStmt = "Insert into FDIntTrans (AccID," & _
                " TransID, TransType,TransDate, Amount, " & _
                " Balance, Particulars,UserID) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                transType & "," & _
                "#" & TransDate & "#," & _
                Abs(InterestAmount) & ", 0," & _
                "'Account Closing Interest'" & _
                "," & gUserID & " )"
    
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    Dim IntHeadID As Long
    'Dim engHeadName As String
    If Len(m_DepositNameEnglish) > 0 Then engHeadName = m_DepositNameEnglish & " " & LoadResString(487)
    IntHeadID = bankClass.GetHeadIDCreated(m_DepositName & " " & GetResourceString(487), _
            engHeadName, parMemDepIntPaid, 0, wis_Deposits)
    
    'do transactions t the respective account heads here it self
    If Not boolContra Then
     'First Case Cumulative Deposit
      If transType = wDeposit Then
        'If it is cash transaction then affect this amount to the cash
        Call bankClass.UpdateCashDeposits(IntHeadID, Abs(InterestAmount), TransDate)
      ElseIf transType = wWithdraw Then
        Call bankClass.UpdateCashWithDrawls(IntHeadID, Abs(InterestAmount), TransDate)
      ElseIf transType = wContraWithdraw And m_Cumulative Then
        Call bankClass.UpdateContraTrans(IntHeadID, m_AccHeadId, Abs(InterestAmount), TransDate)
      ElseIf transType = wContraDeposit And m_Cumulative Then
        Call bankClass.UpdateContraTrans(m_AccHeadId, IntHeadID, Abs(InterestAmount), TransDate)
      End If
    End If
End If

'Now withdraw the amount from FD account
'if he is repaying the loan amount or transferring to other account
'then the transction will be contra
transType = wWithdraw
If boolContra Then transType = wContraWithdraw

gDbTrans.SqlStmt = "Insert into FDTrans (AccID,TransID, " & _
            " TransType,TransDate, Amount, Balance, " & _
            " Particulars,UserID) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            transType & "," & _
            "#" & TransDate & "#," & _
            Amount & "," & _
            " 0 , " & _
            "'Deposit Closed " & m_AccID & "' " & _
            "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

If Not boolContra Then _
    Call bankClass.UpdateCashWithDrawls(m_AccHeadId, Amount, TransDate)

'Update the first transaction with close = 1
gDbTrans.SqlStmt = "UPdate FDmaster Set ClosedDate = #" & TransDate & "#" & _
                " where AccID = " & m_AccID
If Not gDbTrans.SQLExecute Then GoTo ErrLine

'if he is repaying the loan amount or transferring to other account
'then the transction will be contra
If boolContra Then
    Dim ContraID As Long
    Dim UserID As Integer
    Dim VoucherNo As String
    
    'Get the Contra ID
    'put withdrawal transction details int to contra Table
    ContraID = GetMaxContraTransID + 1
    Set rst = Nothing
    transType = wContraWithdraw
    If PayableAmount > 0 Then
        gDbTrans.SqlStmt = "Insert into ContraTrans " & _
                "(ContraId,AccId,AccHeadID," & _
                "TransType,TransId,Amount,VoucherNo,UserID)" & _
                " Values (" & _
                ContraID & "," & _
                m_AccID & "," & _
                PayableHeadID & "," & _
                transType & ", " & _
                TransID & "," & _
                PayableAmount & "," & _
                AddQuotes(VoucherNo, True) & _
                "," & gUserID & " )"
        
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
    
    If InterestAmount Then
        'interest transction
        gDbTrans.SqlStmt = "Insert into ContraTrans " & _
               "(ContraId,AccId,AccHeadID," & _
               "TransType,TransId,Amount,VoucherNo,UserID)" & _
               " Values (" & _
               ContraID & "," & _
               m_AccID & "," & _
               IntHeadID & "," & _
               wContraWithdraw & ", " & _
               TransID & "," & _
               InterestAmount & "," & _
               AddQuotes(VoucherNo, True) & _
               "," & gUserID & " )"
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
    
    If Amount Then
        gDbTrans.SqlStmt = "Insert into ContraTrans " & _
               "(ContraId,AccId,AccHeadID," & _
               "TransType,TransId,Amount,VoucherNo,UserID)" & _
               " Values (" & _
               ContraID & "," & _
               m_AccID & "," & _
               m_AccHeadId & "," & _
               transType & ", " & _
               TransID & "," & _
               Amount & "," & _
               AddQuotes(VoucherNo, True) & _
               "," & gUserID & " )"
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
    
End If

'''NOW ENTER THE DETAILS IF AMOUNT TRANSFERRED TO
'OTHER THAN LOAN ACCOUNT AND SUSPENCE ACCOUNT
'LIKE SB ACCOUNT 'RD ACCOUNT SHARE ACCOUNT

If optTransfer Then
    If m_ContraClass Is Nothing Then
        Set m_ContraClass = New clsContra
        m_ContraClass.TransDate = TransDate
        Call cmdTfr_Click
    End If
    
    If m_ContraClass.Status <> Success Then
        MsgBox "Unable to transfer the details", vbInformation, wis_MESSAGE_TITLE
        GoTo ExitLine
    End If
    
    '''Now Effect to the transaction to respective heads
    m_ContraClass.TransDate = TransDate
    If Not m_ContraClass.SaveDetails() Then GoTo ErrLine
    
End If


'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(m_AccHeadId, _
            m_AccID, TransDate, TransID, Val(txtPayableAmount)) < 1 Then
        gDbTrans.RollBack
        Exit Function
    End If
    Set Cashclass = Nothing
End If

gDbTrans.CommitTrans: InTrans = False

FDClose = True

ExitLine:

If InTrans Then gDbTrans.RollBack: InTrans = False

Set m_ContraClass = Nothing
Set bankClass = Nothing

Exit Function

ErrLine:
    If InTrans Then gDbTrans.RollBack: InTrans = False
    If Err Then
        MsgBox Err.Number & " : " & vbCrLf & Err.Description, vbExclamation, gAppName & " - Error"
    Else
        'MsgBox "Unable to close account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(534), vbExclamation, gAppName & " - Error"
    End If
    
    Set m_ContraClass = Nothing
    Set bankClass = Nothing

End Function

Private Function TransferToLoanAccount() As Boolean
On Error GoTo ErrLine

Dim LoanID As Long
Dim custId As Long
Dim bankClass  As clsBankAcc
'Perform the transaction with closed flag and send the guy home

'Check date
Dim TransDate As Date
Dim TransID As Long
Dim Amount As Currency
Dim InTrans As Boolean
Dim ContraID As Long
Dim UserID As Integer
Dim VoucherNo As String

Dim rst As ADODB.Recordset

Dim PayableAmount As Currency
Dim PayableBalance As Currency

Dim transType As wisTransactionTypes
Dim InterestAmount As Currency
Dim TotalIntAmount As Currency

'Now Get the Interest Amount deposited in InterestPayble Account
PayableAmount = txtPayable
InterestAmount = txtDepInterest - txtCharges
TotalIntAmount = txtTotalInterest

'Get the Transaction date
TransDate = GetSysFormatDate(txtDate.Text)

gDbTrans.SqlStmt = "Select * From FDMaster Where " & _
            " Accid = " & m_AccID
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
LoanID = FormatField(rst("LoanID"))
custId = FormatField(rst("CustomerID"))
UserID = gCurrUser.UserID
Set rst = Nothing

gDbTrans.SqlStmt = "Select Balance From FDIntPayable " & _
                " Where Accid = " & m_AccID & _
                " ORDER By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                    PayableBalance = FormatField(rst(0))
Set rst = Nothing

'AMOUNT TO BE DEDUCTED FROM INTEREST PAYBLE ACCOUNT IS PayableAmount
PayableBalance = PayableBalance - PayableAmount
If PayableBalance < 0 Then
    If MsgBox(GetResourceString(577) & vbCrLf & GetResourceString(541) & "?", _
        vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then GoTo ExitLine
    PayableBalance = 0
End If

'so the interest from the P & L account is
'AMOUNT TO BE DEDUCTED FROM PROFIT & LOSS ACCOUNT IS InterestAmount
'total interest - payble interest i.e. InterestAmount
InterestAmount = txtTotalInterest - PayableAmount - txtCharges
'InterestAmount = txtTotalInterest - PayableAmount

Amount = Val(txtDepositAmount)

'Get the Next TransCtionId
TransID = GetFDMaxTransID(m_AccID) + 1

'Get the Contra ID
'put withdrawal transction details into to contra Table
ContraID = GetMaxContraTransID + 1

Set bankClass = New clsBankAcc
gDbTrans.BeginTrans: InTrans = True

'Transction of the interest Amount
'first Make the Transction To The Payable Account
Dim PayableHeadID As Long
Dim engHeadName As String
If Len(m_DepositNameEnglish) > 0 Then engHeadName = m_DepositNameEnglish & " " & LoadResourceStringS(375, 47)
    PayableHeadID = bankClass.GetHeadIDCreated(m_DepositName & " " & GetResourceString(375, 47), _
            engHeadName, parDepositIntProv, 0, wis_Deposits + m_DepositType)

If PayableAmount Then
    'Now withdraw the amount from Interest Paybale account
    'if he is repaying the loan amount or transferring to other account
    'then the transction will be contra
    transType = wContraWithdraw
    gDbTrans.SqlStmt = "Insert into FDIntPayable(AccID,TransID, " & _
                    " TransType,TransDate,Amount,Balance, " & _
                    " Particulars,UserID) values ( " & _
                    m_AccID & "," & _
                    TransID & "," & _
                    transType & "," & _
                    "#" & TransDate & "#," & _
                    PayableAmount & "," & _
                    PayableBalance & ", " & _
                    "'Closing FD " & m_AccID & "' " & _
                    "," & gUserID & " ) "
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
        MsgBox GetResourceString(535), vbCritical, gAppName & " - Critical Error"
        GoTo ExitLine
    End If
    
    gDbTrans.SqlStmt = "Insert into ContraTrans " & _
                    "(ContraId,AccId,AccHeadID," & _
                    "TransType,TransId,Amount,VoucherNo,UserID)" & _
                    " Values (" & _
                    ContraID & "," & _
                    m_AccID & "," & _
                    PayableHeadID & "," & _
                    transType & ", " & _
                    TransID & "," & _
                    PayableAmount & "," & _
                    AddQuotes(VoucherNo, True) & _
                    "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

End If

'Now make the transction to the
'Interest Table
Dim IntHeadID As Long
engHeadName = ""

If Len(m_DepositNameEnglish) > 0 Then engHeadName = m_DepositNameEnglish & " " & LoadResString(487)
IntHeadID = bankClass.GetHeadIDCreated(m_DepositName & " " & GetResourceString(487), _
            engHeadName, parMemDepIntPaid, 0, wis_Deposits + wis_Deposits)
    
If InterestAmount Then
    'if he is repaying the loan amount or transferring to other account
    'then the transaction will be contra
    transType = wContraWithdraw
    
    'If Calculated interest is smaller than the Interest Payable
    'Then take the Difference Amount as Profit
    If InterestAmount < 0 Then transType = wContraDeposit
    
    gDbTrans.SqlStmt = "Insert into FDIntTrans (AccID," & _
                    " TransID, TransType,TransDate, Amount, " & _
                    " Balance, Particulars,UserID) values ( " & _
                    m_AccID & "," & _
                    TransID & "," & _
                    transType & "," & _
                    "#" & TransDate & "#," & _
                    Abs(InterestAmount) & ", 0," & _
                    "'Account Closing Interest'" & _
                    "," & gUserID & " )"
    
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    
    'contra Transaction
    transType = wContraWithdraw
    gDbTrans.SqlStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," & _
            "TransType,TransId,Amount,VoucherNo,UserID)" & _
            " Values (" & _
            ContraID & "," & _
            m_AccID & "," & _
            IntHeadID & "," & _
            transType & ", " & _
            TransID & "," & _
            InterestAmount & "," & _
            AddQuotes(VoucherNo, True) & _
            "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

End If

'Now withdraw the amount from FD account
'if he is repaying the loan amount or transferring to other account
'then the transction will be contra

transType = wContraWithdraw
gDbTrans.SqlStmt = "Insert into FDTrans (AccID,TransID, " & _
        " TransType,TransDate, Amount, Balance, " & _
        " Particulars,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        transType & "," & _
        "#" & TransDate & "#," & _
        Amount & "," & _
        " 0 , " & _
        "'Deposit Closed " & m_AccID & "' " & _
        "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Update the first transaction with close = 1
gDbTrans.SqlStmt = "UPdate FDmaster Set ClosedDate = #" & TransDate & "#" & _
                " Where AccID = " & m_AccID
If Not gDbTrans.SQLExecute Then GoTo ErrLine
     
gDbTrans.SqlStmt = "Insert into ContraTrans " & _
                "(ContraId,AccId,AccHeadID," & _
                "TransType,TransId,Amount,VoucherNo,UserId)" & _
                " Values (" & _
                ContraID & "," & _
                m_AccID & "," & _
                m_AccHeadId & "," & _
                transType & ", " & _
                TransID & "," & _
                Amount & "," & _
                AddQuotes(VoucherNo, True) & _
                "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
'if he is transfering the Amount to the loan account,
'transfer it to loan account
Dim LoanAmount As Currency
Dim LoanIntAmount As Currency
Dim MatAmount As Currency
Dim SuspAmount As Currency

MatAmount = txtDepositAmount + txtPayable + txtDepInterest - txtCharges
LoanAmount = Val(txtLoanAmount)
LoanIntAmount = txtLoanInterest

Dim DepLOanClass As clsDepLoan
Set DepLOanClass = New clsDepLoan

'if loan amount is more than the matured amount
'then first take the interest then remaining amount as principal
'loanclass
SuspAmount = MatAmount - LoanIntAmount
LoanAmount = IIf(LoanAmount < SuspAmount, LoanAmount, SuspAmount)
SuspAmount = MatAmount - LoanAmount - LoanIntAmount

'Deposit The AMount into Deposit Lons
TransferToLoanAccount = DepLOanClass.DepositAmount(LoanID, _
            LoanAmount, LoanIntAmount, "FD Closed", TransDate, VoucherNo)

Dim LoanHeadID As Long
Dim LoanIntHeadID As Long
Dim AccHeadName As String
'Load Head
AccHeadName = m_DepositName & " " & GetResourceString(58)
LoanHeadID = bankClass.GetHeadIDCreated(AccHeadName)
'Loan INterest head
AccHeadName = m_DepositName & " " & GetResourceString(58, 483)
LoanIntHeadID = bankClass.GetHeadIDCreated(AccHeadName)

'if matured amount is more than the loan loan amount
'then transfer tht remaining amount to the suspence account
If SuspAmount Then
    Dim SuspHeadID As Long
    SuspHeadID = GetIndexHeadID(GetResourceString(365))
    Debug.Assert SuspAmount = 0
    SuspHeadID = bankClass.GetHeadIDCreated(GetResourceString(365), LoadResString(365), _
                                    parSuspAcc, 0, wis_SuspAcc)
    
    Dim SuspClass As New clsSuspAcc
    If SuspClass.DepositAmount(m_AccHeadId, m_AccID, custId, "", _
                        TransDate, SuspAmount, TransID, VoucherNo) < 1 Then GoTo ExitLine

End If

'Update the transactin to the ledger Heads
'Transfer all the Amount to the to DepositAccountHead
'then Transfer remaining amount
If PayableAmount > 0 Then
    If Not bankClass.UpdateContraTrans(PayableHeadID, LoanHeadID, _
        PayableAmount + IIf(InterestAmount > 0, 0, InterestAmount), TransDate) Then GoTo ErrLine
End If

'Transaction Interest amount
If InterestAmount > 0 Then
    If Not bankClass.UpdateContraTrans(IntHeadID, LoanHeadID, _
                        InterestAmount, TransDate) Then GoTo ErrLine
Else

    'withdraw all the from Payble head id
    'then undo the difference amount from the Payble headid & DepositHeadID
    If Not bankClass.UpdateContraTrans(PayableHeadID, IntHeadID, _
                        Abs(InterestAmount), TransDate) Then GoTo ErrLine
End If

'new code
If Not bankClass.UpdateContraTrans(m_AccHeadId, LoanHeadID, _
                    LoanAmount, TransDate) Then GoTo ErrLine
 
If Not bankClass.UpdateContraTrans(m_AccHeadId, LoanIntHeadID, _
                    LoanIntAmount, TransDate) Then GoTo ErrLine

'now transfer the remaining amount ot the suspence head
If SuspAmount Then
    If Not bankClass.UpdateContraTrans(m_AccHeadId, SuspHeadID, _
                                    SuspAmount, TransDate) Then GoTo ErrLine
End If

'Now Compansate the Difference Amount from Payble Head ID & DepositHead
'The DiffAmount is
Amount = PayableAmount + InterestAmount 'IIf(InterestAmount < 0, InterestAmount, 0)
If Not bankClass.UndoContraTrans(m_AccHeadId, _
            LoanHeadID, Amount, TransDate) Then GoTo ErrLine


gDbTrans.CommitTrans: InTrans = False

TransferToLoanAccount = True

ExitLine:

If InTrans Then gDbTrans.RollBack: InTrans = False

Set bankClass = Nothing
Exit Function

ErrLine:
    If InTrans Then gDbTrans.RollBack: InTrans = False
If Err Then
    MsgBox "ERROR in TransfertoLoanaccount" & Err.Number & vbCrLf & _
        Err.Description, vbInformation, wis_MESSAGE_TITLE
Else
    'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
    MsgBox GetResourceString(535), vbCritical, gAppName & " - Critical Error"
End If
Set bankClass = Nothing

End Function

Private Function TransferTOMatureFD() As Boolean

TransferTOMatureFD = False

Dim bankClass  As clsBankAcc
Dim TransDate As Date
Dim TransID As Long
Dim Amount As Currency
Dim InTrans As Boolean
Dim rst As ADODB.Recordset
Dim PayableAmount As Currency
Dim PayableBalance As Currency
Dim transType As wisTransactionTypes
Dim InterestAmount As Currency
Dim TotalIntAmount As Currency


'Get the Transaction date
TransDate = GetSysFormatDate(txtDate.Text)

'Now Get the Interest Amount deposited in InterestPayble Account
PayableAmount = txtPayable
InterestAmount = txtDepInterest
TotalIntAmount = txtTotalInterest

gDbTrans.SqlStmt = "Select Balance From FDIntPayable Where " & _
            " Accid = " & m_AccID & " ORDER By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then PayableBalance = FormatField(rst(0))
Set rst = Nothing

'AMOUNT TO BE DEDUCTED FROM INTEREST PAYBLE ACCOUNT IS PayableAmount
PayableBalance = PayableBalance - PayableAmount
If PayableBalance < 0 Then
    If MsgBox(GetResourceString(577) & vbCrLf & GetResourceString(541) & "?", _
        vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    PayableBalance = 0
End If

'so the interest from the P & L account is
'AMOUNT TO BE DEDUCTED FROM PROFIT & LOSS ACCOUNT IS InterestAmount
'total interest - payble interest i.e. InterestAmount
InterestAmount = txtTotalInterest - PayableAmount
Amount = Val(txtDepositAmount)

Dim ContraID As Long
Dim UserID As Integer
Dim VoucherNo As String

'Get the Transction
TransID = GetFDMaxTransID(m_AccID) + 1

'Get the Contra ID
ContraID = GetMaxContraTransID + 1

Set rst = Nothing

Set bankClass = New clsBankAcc
gDbTrans.BeginTrans: InTrans = True

Dim AccHeadName As String
Dim AccHeadNameEnglish As String
Dim MatHeadId As Long
Dim IntHeadID As Long
Dim PayableHeadID As Long

'Matured deposit
AccHeadName = GetResourceString(46) & " " & m_DepositName
If Len(m_DepositNameEnglish) > 0 Then AccHeadNameEnglish = LoadResString(46) & " " & m_DepositNameEnglish
MatHeadId = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemberDeposit, _
                    0, wis_Deposits + m_DepositType)
'interest paid
AccHeadName = m_DepositName & " " & GetResourceString(487)
If Len(m_DepositNameEnglish) > 0 Then AccHeadNameEnglish = m_DepositNameEnglish & " " & LoadResString(487)
IntHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepIntPaid, _
                    0, wis_Deposits + m_DepositType)
'Interesr Payble
AccHeadName = m_DepositName & " " & GetResourceString(375, 47)
If Len(m_DepositNameEnglish) > 0 Then AccHeadNameEnglish = m_DepositNameEnglish & " " & LoadResourceStringS(375, 47)
PayableHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parDepositIntProv, _
                    0, wis_Deposits + m_DepositType)

'Transction of the interest Amount
'first Make the Transction To The Payable Account
If PayableAmount Then
    'Now withdraw the amount from Interest Paybale account
    transType = wContraWithdraw
    gDbTrans.SqlStmt = "Insert into FDIntPayable(AccID,TransID, " & _
        " TransType,TransDate,Amount,Balance, " & _
        " Particulars,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        transType & "," & _
        "#" & TransDate & "#," & _
        PayableAmount & "," & _
        PayableBalance & ", " & _
        "'Closing FD " & m_AccID & "' " & _
        "," & gUserID & " ) "
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    gDbTrans.SqlStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," _
             & " TransType,TransId,Amount,VoucherNo,UserID)" & _
            " Values (" _
             & ContraID & "," & _
             m_AccID & "," & _
             PayableHeadID & "," & _
             transType & ", " & TransID & "," & _
             PayableAmount & "," & _
             AddQuotes(VoucherNo, True) & _
             "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

'Now make the transction to the
'Interest Table
If InterestAmount Then
    transType = wContraWithdraw
    If InterestAmount < 0 Then transType = wContraDeposit
    gDbTrans.SqlStmt = "Insert into FDIntTrans (AccID," & _
        " TransID, TransType,TransDate, Amount, " & _
        " Balance, Particulars,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        transType & "," & _
        "#" & TransDate & "#," & _
        Abs(InterestAmount) & ", 0," & _
        "'Account Closing Interest'" & _
        "," & gUserID & " )"
    
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    gDbTrans.SqlStmt = "Insert into ContraTrans " & _
            "(ContraId,AccId,AccHeadID," _
             & " TransType,TransId,Amount,VoucherNo,UserID)" & _
            " Values (" _
             & ContraID & "," & _
             m_AccID & "," & _
             IntHeadID & "," & _
             transType & ", " & TransID & "," & _
             Abs(InterestAmount) & "," & _
             AddQuotes(VoucherNo, True) & _
             "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

'Now withdraw the amount from FD account
transType = wContraWithdraw
gDbTrans.SqlStmt = "Insert Into FDTrans (AccID,TransID, " & _
        " TransType,TransDate, Amount, Balance, " & _
        " Particulars,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        transType & "," & _
        "#" & TransDate & "#," & _
        Amount & "," & _
        " 0 , " & _
        "'Deposit Closed " & m_AccID & "' " & _
        "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Update the first transaction with close = 1
gDbTrans.SqlStmt = "UPdate FDmaster Set ClosedDate = #" & TransDate & "#" & _
                " where AccID = " & m_AccID
If Not gDbTrans.SQLExecute Then GoTo ErrLine

gDbTrans.SqlStmt = "Insert into ContraTrans " & _
        "(ContraId,AccId,AccHeadID," _
         & " TransType,TransId,Amount,VoucherNo,UserID)" & _
        " Values (" _
         & ContraID & "," & _
         m_AccID & "," & _
         m_AccHeadId & "," & _
         transType & ", " & TransID & "," & _
         Amount & "," & _
         AddQuotes(VoucherNo, True) & _
         "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine

'First Transfer the Interest Amount Then TransFer the Deposit Amount
 transType = wContraDeposit
 TransID = 1
 TransDate = GetSysFormatDate(txtDate.Text)
 gDbTrans.SqlStmt = "Insert into MatFDTrans (AccID, TransID, " & _
         " TransType,TransDate, Amount, Balance, " & _
         " Particulars,UserID) values ( " & _
         m_AccID & "," & _
         TransID & "," & _
         transType & "," & _
         "#" & TransDate & "#," & _
         PayableAmount + InterestAmount & "," & _
         PayableAmount + InterestAmount & "  , " & _
         "'From Deposit Interest', " & _
         gUserID & " )"
 
If Not gDbTrans.SQLExecute Then GoTo ErrLine

gDbTrans.SqlStmt = "Insert into ContraTrans " & _
        "(ContraId,AccId,AccHeadID," _
         & " TransType,TransId,Amount,VoucherNo,UserID)" & _
        " Values (" _
         & ContraID & "," & _
         m_AccID & "," & _
         MatHeadId & "," & _
         transType & ", " & TransID & "," & _
         PayableAmount + InterestAmount & "," & _
         AddQuotes(VoucherNo, True) & _
         "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine

TransID = 2
gDbTrans.SqlStmt = "Insert into MatFDTrans (AccID,TransID, " & _
        " TransType,TransDate, Amount, Balance, " & _
        " Particulars,UserID) values ( " & _
        m_AccID & "," & _
        TransID & "," & _
        transType & "," & _
        "#" & TransDate & "#," & _
        Amount & "," & _
        Amount + PayableAmount + InterestAmount & "  , " & _
        "'From Deposit Principle', " & _
        gUserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
gDbTrans.SqlStmt = "Insert into ContraTrans " & _
        "(ContraId,AccId,AccHeadID," _
         & " TransType,TransId,Amount,VoucherNo,UserID)" & _
        " Values (" _
         & ContraID & "," & _
         m_AccID & "," & _
         MatHeadId & "," & _
         transType & ", " & TransID & "," & _
         Amount & "," & _
         AddQuotes(VoucherNo, True) & _
         "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then GoTo ErrLine

gDbTrans.SqlStmt = "UPDATE FDMaster SET MaturedON = #" & TransDate & "#" & _
    " WHERE AccID = " & m_AccID
If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
'Now Update the transction to ledger heads
    
If PayableAmount Then
    If Not bankClass.UpdateContraTrans(PayableHeadID, IntHeadID, _
            PayableAmount, TransDate) Then GoTo ErrLine
    InterestAmount = InterestAmount + PayableAmount
End If
If InterestAmount Then
    If Not bankClass.UpdateContraTrans(IntHeadID, MatHeadId, _
            InterestAmount, TransDate) Then GoTo ErrLine
    InterestAmount = 0
End If
'Now Fd Amount
If Not bankClass.UpdateContraTrans(m_AccHeadId, MatHeadId _
        , Amount, TransDate) Then GoTo ErrLine
    

TransferTOMatureFD = True

gDbTrans.CommitTrans
InTrans = False
Set bankClass = Nothing

Exit Function

ErrLine:
    If Err Then
        MsgBox "ERROR Trfer MAt FD" & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Else
        'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
        MsgBox GetResourceString(535), vbCritical, gAppName & " - Critical Error"
    End If

End Function

Private Sub UpdateDetails()

ChkDeductLoan.Enabled = False
optMature.Enabled = False

'Todays date
If txtDate.Text = "" Then txtDate.Text = gStrDate

Dim custId As Long
Dim rst As ADODB.Recordset
Dim rstMaster As Recordset

gDbTrans.SqlStmt = "Select * from FDMaster Where " & _
                    " AccId = " & m_AccID
If gDbTrans.Fetch(rstMaster, adOpenDynamic) <= 0 Then
    'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

custId = FormatField(rstMaster("CustomerID"))
gDbTrans.SqlStmt = "Select Title +' '+ FirstName +' '+ " & _
            " MiddleName +' '+ LastName as CustName from NameTab " & _
            " Where CustomerID = " & custId
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then GoTo ErrLine

lblName.Caption = Trim$(FormatField(rst("CustName")))

Set rst = Nothing

'AS THE CALCULATION IS SLOW IN INDIAN FORMAT DATE
'SO IN THIS WE HAVE TAKING THE DATE IN FORMAT OF "MM/DD/YYYY"
'ALL DATE VARIABLE ARE DECLARED AS DATE INSTEAD OF STRING

Dim LoanID As Long
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes

Dim Payable As Double
Dim DepAmt As Double
Dim DepBalance As Double

Dim DepDate As Date
Dim TransDate As Date
Dim LastIntDate As Date

Dim ClsBank As clsBankAcc
Dim rstTemp As Recordset

'Check for valid date
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
If m_AccID = 0 Then GoTo ErrLine

cmdAccept.Enabled = True

'Get the Transction Date In "mm/dd/yyyy" format
TransDate = GetSysFormatDate(txtDate.Text)
If DateDiff("d", GetFDMaxTransDate(m_AccID), TransDate) < 0 Then
    'MsgBox "Date of transaction should be later that the ones already transacted !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If

DepDate = rstMaster("EffectiveDate")
DepAmt = rstMaster("DepositAmount")

'Now Get the Deposit Type
gDbTrans.SqlStmt = "SELECT * From DepositName Where " & _
                " DepositID = " & m_DepositType Mod wis_Deposits
Call gDbTrans.Fetch(rstTemp, adOpenForwardOnly)
m_DepositName = FormatField(rstTemp("DepositName"))
m_DepositNameEnglish = FormatField(rstTemp("DepositNameEnglish"))
Set rstTemp = Nothing

'Get the deposit amount
transType = wDeposit
ContraTransType = wContraDeposit

'GET THE DEPOSIT Balance & DEPOSIT AMOUNT
gDbTrans.SqlStmt = "Select top 1 TransDate,TransID,Balance FROM FDTrans " & _
        " WHERE AccID = " & m_AccID & " ORDER BY TransID Desc"
Call gDbTrans.Fetch(rstTemp, adOpenForwardOnly)
DepBalance = FormatField(rstTemp("Balance"))

LastIntDate = DepDate
'If DepBalance <> DepAmt Then
    'so get the last interest paid date
    gDbTrans.SqlStmt = "Select top 1 TransDate FROM FDIntTrans " & _
            " WHERE AccID = " & m_AccID & " ORDER BY TransID Desc"
    If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) Then _
                            LastIntDate = rstTemp("TransDate")
'End If

If m_Cumulative Then
    Payable = DepBalance - DepAmt
Else
    gDbTrans.SqlStmt = "Select Top 1 TransID,Balance " & _
                        " From FDIntPayable Where " & _
                        " Accid = " & m_AccID & _
                        " Order By TransID Desc "
    If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
                            Payable = FormatField(rstTemp("Balance"))
    Set rstTemp = Nothing
End If

txtPayable.Value = Payable

'Now Get the Interest Amount deposited in InterestPayble Account
'If m_Cumulative And Payable = 0 Then txtPayable.Enabled = False
Dim MaturedIntAmount As Currency
Dim IntPayable As Currency
Dim MatDate As Date
Dim RateOfInt As Single

MatDate = rstMaster("MaturityDate")
txtDate.Tag = MatDate
'RateOfInt = FormatField(rstMaster("RateOfInterest"))
LoanID = FormatField(rstMaster("LoanId"))
If DateDiff("D", MatDate, TransDate) >= 0 Then optMature.Enabled = True

RateOfInt = FormatField(rstMaster("RateOfInterest"))
txtDepositAmount = FormatCurrency(DepAmt)
optClose.Value = True
optMature.Enabled = True
optMature.Tag = 1

Dim FdClass As clsFDAcc
Dim IntAmount As Currency

Set FdClass = New clsFDAcc
'IntAmount = FdClass.InterestAmountTillDate(m_AccID, TransDate)
'If IntAmount < 0 Then IntAmount = 0

If DateDiff("d", MatDate, TransDate) < 0 Then
    'Deposit is closing before the maturity date
    'So reduce the Interest rate by 2%
    RateOfInt = GetDepositInterestRate(m_DepositType, DepDate, TransDate)
    If RateOfInt <= 0 Then RateOfInt = FormatField(rstMaster("RateOfInterest"))
    RateOfInt = RateOfInt - 2
    'No Possibality of Transferring it to MFD
    optMature.Enabled = False
    optMature.Tag = 0
    optClose.Value = True
    
    IntAmount = FdClass.InterestAmountTillDate(m_AccID, TransDate)
    If IntAmount < 0 Then IntAmount = 0
Else
    IntAmount = FdClass.InterestAmount(m_AccID, LastIntDate, MatDate)
    IntAmount = IntAmount + Payable
End If

Set FdClass = Nothing

txtInterest.Text = RateOfInt
IntAmount = IntAmount \ 1
IntAmount = IntAmount - Payable
'txtDepInterest = IIf(IntAmount > 0, IntAmount, 0)
'txtCharges = IIf(IntAmount < 0, Abs(IntAmount), 0)
txtDepInterest = IntAmount
'txtCharges = IIf(IntAmount < 0, Abs(IntAmount), 0)
 
If IntAmount < 0 Then
    lblInterest = GetResourceString(47, 398)
    With txtCharges
'        .Value = Abs(IntAmount)
'        .Enabled = False
    End With
Else
    'txtCharges.Enabled = True
    lblInterest = GetResourceString(47)
End If
txtTotalInterest = txtPayable + txtDepInterest

'Get total loan amount
Dim LoanAmount As Currency
Dim LoanInt As Currency
Dim LoanDate As Date
Dim LoanBalance As Currency
        
'Get Loan Balance
txtLoanAmount = "0.00"
txtLoanInterest = 0

fraLoanDet.Enabled = False
If LoanID > 0 Then
    fraLoanDet.Enabled = True
    ChkDeductLoan.Enabled = True
    ChkDeductLoan.Value = 1
    optMature.Value = False
    optMature.Enabled = False
    gDbTrans.SqlStmt = "Select * from DepositLoanMaster where " & _
                " LoanID = " & LoanID
    If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then
        optClose.Value = True
        LoanAmount = CCur(FormatField(rstTemp("LoanAmount")))
        LoanDate = rstTemp("LoanIssueDate")
    End If
    gDbTrans.SqlStmt = "Select top 1 Balance,TransDate from DepositLoanTrans where " & _
                " LoanID = " & LoanID & " Order by TransId Desc "
    If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then
        LoanBalance = CCur(FormatField(rstTemp(0)))
        If LoanBalance <= 0 Then ChkDeductLoan.Value = vbUnchecked
        'LastIntDate = rstTemp("TransDate")
        txtLoanAmount = FormatCurrency(LoanAmount)
    End If
    
    'Calculate the RegularLoanInterest Paid.
    LoanInt = ComputeDepLoanRegularInterest(TransDate, LoanID)
    txtLoanInterest = LoanInt
    
    txtLoanAmount = FormatCurrency(LoanBalance)
End If

Set rstTemp = Nothing

Exit Sub

ErrLine:
    
    'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
    Exit Sub

End Sub

Private Sub ChkDeductLoan_Click()
If ChkDeductLoan.Value = vbChecked Then
    optTransfer.Enabled = False
    optMature.Enabled = False
    optClose = True
Else
    optTransfer.Enabled = True
    optMature.Enabled = Val(optMature.Tag)
End If
Call txtDepositAmount_Change
End Sub


Private Sub cmdAccept_Click()

If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date not in dd/mm/yyyy format...! ", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If

'Do not allow if this deposit has loans
If Val(txtLoanAmount) > 0 And ChkDeductLoan.Value = vbUnchecked Then
   'MsgBox "The deposit you are trying to close has loans against it" & vbCrLf & _
            "You must first get the loan repayment and then close this deposit", _
            "Do you want to continue?",vbquestion+vbInformation, gAppName & " - Message"
   If MsgBox(GetResourceString(574) & vbCrLf & vbCrLf & _
        GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, _
         gAppName & " - Confirmation") = vbNo Then Exit Sub
End If

'Warn for premature closure
gDbTrans.SqlStmt = "Select MaturityDate from FDMaster where AccID = " & m_AccID

Dim Days As Integer
Dim TransDate As Date
Dim rst As Recordset

TransDate = GetSysFormatDate(txtDate)
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
Days = DateDiff("d", rst("MaturityDate"), TransDate)
If Days < 0 Then
    'If MsgBox("You are attempting to close this deposit prematurely !" & vbCrLf & vbCrLf & "Are you sure you want to continue this operation ?", vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
    If MsgBox(GetResourceString(576) & vbCrLf & vbCrLf & _
        GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, _
        gAppName & " - Confirmation") = vbNo Then Exit Sub
End If

If ChkDeductLoan.Value = vbChecked And (optTransfer Or optMature) Then
    If MsgBox("After paying the Loan amount " & vbCrLf & _
        "The remaining Deposit amount will be transferred to the Suspence account" & _
        vbCrLf & "DO You want to continue?", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) _
                = vbYes Then Exit Sub
        optClose = True
End If

If optMature Then
    If Not TransferTOMatureFD Then Exit Sub
Else
    If ChkDeductLoan = vbChecked Then
        If Not TransferToLoanAccount Then Exit Sub
    Else
        If Not FDClose Then Exit Sub
    End If
End If

Unload Me

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDate_Click()
Dim strDate As String
With Calendar
    .Left = Me.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + cmdDate.Top
    .selDate = txtDate.Text
    strDate = .selDate
    .Show vbModal
    txtDate.Text = .selDate
    If .selDate = strDate Then Exit Sub
End With

Call txtDate_LostFocus

End Sub

Private Sub cmdTfr_Click()

Dim AccNum As String
Dim rst As Recordset

Dim PayableAmount As Currency
Dim Amount As Currency
Dim IntAmount As Currency

Dim PayableHeadID As Long
Dim IntHeadID As Long
Dim BankCls As clsBankAcc

Set BankCls = New clsBankAcc

gDbTrans.SqlStmt = "SELECT AccNum From FDMaster Where AccId = " & m_AccID
If gDbTrans.Fetch(rst, adOpenDynamic) Then AccNum = FormatField(rst(0))

PayableAmount = IIf(m_Cumulative, 0, txtPayable)
IntAmount = txtDepInterest.Value - txtCharges
'If IntAmount = 0 Then IntAmount = txtCharges * -1

Amount = Val(txtDepositAmount) + IIf(m_Cumulative, txtPayable, 0)
'Amount = txtPayableAmount

If Not m_ContraClass Is Nothing Then Set m_ContraClass = Nothing
Set m_ContraClass = New clsContra
On Error Resume Next
m_ContraClass.TransDate = GetSysFormatDate(txtDate)

On Error GoTo Tfr_error

With m_ContraClass
    If .Transfer(GetSysFormatDate(txtDate), "12", m_AccHeadId, AccNum, Amount) <> Success Then GoTo Tfr_error
    If PayableAmount Then
        Debug.Assert PayableAmount = 0
        'parDepositIntProv
        'DOUBT IN aMOUNT TYPE
        PayableHeadID = GetIndexHeadID(m_DepositName & " " & GetResourceString(375, 47))
        If .Transfer(GetSysFormatDate(txtDate), "12", PayableHeadID, AccNum, PayableAmount, wisPayable) <> Success Then GoTo Tfr_error
    End If
    If IntAmount > 0 Then
        IntHeadID = GetIndexHeadID(m_DepositName & " " & GetResourceString(487))
        If .Transfer(GetSysFormatDate(txtDate), "12", IntHeadID, AccNum, IntAmount, wisRegularInt) <> Success Then GoTo Tfr_error
    End If
    'Below Commented Line moved above 'If PayableAmount Then' line
    'If .Transfer(GetSysFormatDate(txtDate), "12", m_AccHeadId, AccNum, Amount) <> Success Then GoTo Tfr_error
    
    If IntAmount < 0 Then
        IntHeadID = GetIndexHeadID(m_DepositName & " " & GetResourceString(487))
        If .Transfer(GetSysFormatDate(txtDate), "12", IntHeadID, AccNum, IntAmount, wisRegularInt) <> Success Then GoTo Tfr_error
    End If
    .Show
End With

Exit Sub

Tfr_error:
    MsgBox "Please close & reload this form"
    Exit Sub


End Sub

Private Sub Form_Load()
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

' Set Kannada Fonts
Call SetKannadaCaption
m_Loaded = True

If m_AccID Then
    Call UpdateDetails
    If m_Cumulative Then lblPayable.Caption = GetResourceString(47, 267)
    txtPayable.Enabled = IIf(m_Cumulative, False, True)
End If
If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If


End Sub
Private Sub Form_Unload(Cancel As Integer)
m_Loaded = False
If Not m_ContraClass Is Nothing Then Set m_ContraClass = Nothing
'Set frmFdClose = Nothing
   
End Sub


Private Sub optClose_Click()
cmdTfr.Enabled = optTransfer.Value
End Sub

Private Sub optMature_Click()
cmdTfr.Enabled = optTransfer.Value
End Sub


Private Sub optTransfer_Click()
cmdTfr.Enabled = optTransfer.Value
End Sub

Private Sub txtCharges_Change()
Call txtDepositAmount_Change
End Sub

Private Sub txtCharges_GotFocus()
With txtCharges
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtDate_GotFocus()
On Error Resume Next
With txtDate
    .SelStart = 0
    .SelLength = InStr(1, .Text, "/") - 1
End With

End Sub

Private Sub txtDate_LostFocus()

'If Me.ActiveControl.Name = cmdAccept.Name Then TXTdEPOSITaMOUNT.SetFocus
If Not DateValidate(txtDate, "/", True) Then Exit Sub

If Not gOnLine Then Call UpdateDetails

End Sub


Private Sub txtDepInterest_Change()
    txtTotalInterest = txtDepInterest + txtPayable
End Sub

Private Sub txtDepInterest_GotFocus()
With txtDepInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtDepositAmount_Change()
If ChkDeductLoan Then
    txtPayableAmount = Val(txtDepositAmount) + _
            txtTotalInterest - Val(txtLoanAmount) - txtLoanInterest - txtCharges - txtTax
Else
    txtPayableAmount = Val(txtDepositAmount) + _
            txtTotalInterest - txtCharges - txtTax
End If
End Sub

Private Sub txtLoanInterest_GotFocus()
With txtLoanInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtPayable_GotFocus()
With txtPayable
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtTax_GotFocus()
With txtTax
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtTotalInterest_Change()
Call txtDepositAmount_Change
End Sub

Private Sub txtLoanAmount_Change()
Call txtDepositAmount_Change
End Sub


Private Sub txtLoanInterest_Change()
Call txtDepositAmount_Change
End Sub


Private Sub txtPayable_Change()
txtTotalInterest = txtDepInterest + txtPayable - txtCharges

End Sub

Private Sub txtTax_Change()
Call txtDepositAmount_Change
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

fraDeposit.Caption = GetResourceString(43, 295)
fraLoanDet = GetResourceString(58, 295)
fraCharges = GetResourceString(237, 273)

'Set the Kannada caption to the Command buttons
cmdAccept.Caption = GetResourceString(11)
cmdCancel.Caption = GetResourceString(2)

lblDate = GetResourceString(37)
lblDepAmount = GetResourceString(43, 42)
lblDepInt = GetResourceString(186)

lblLoanDet = GetResourceString(58, 295)
lblLoanAmount = GetResourceString(80, 91)
lblLoanInt = GetResourceString(186)
lblIntOnLoan = GetResourceString(80, 47)

lblPreClose = GetResourceString(238)
lblOthers = GetResourceString(237, 273)
lblNetAmount = GetResourceString(240)
lblIntOnDep.Caption = GetResourceString(52, 47) & " :"
lblInterest.Caption = GetResourceString(47) & " :"
lblPayable.Caption = GetResourceString(450) & " :"

optClose.Caption = GetResourceString(11)  'Close
optMature.Caption = GetResourceString(46, 45) 'Matures FD

End Sub



Private Sub txtTotalInterest_GotFocus()
With txtTotalInterest
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


