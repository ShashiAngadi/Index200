VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSBAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements IAccounts
'Implements IHTMLTable

Public Event UpdateStatus(strMsg As String)
Private m_rstCustBalance As Recordset
Private m_AccID As Long
Private m_CustomerID As Long
Private m_ParentHandle As Long
Private m_TransDate As Date

Private WithEvents m_frmSBAcc As frmSBAcc
Attribute m_frmSBAcc.VB_VarHelpID = -1
Private WithEvents m_frmSBReport As frmSBReport
Attribute m_frmSBReport.VB_VarHelpID = -1
Public Property Let customerID(NewValue As Long)

Dim rst As Recordset
gDbTrans.SqlStmt = "Select AccID From SBMaster Where CustomerID = " & NewValue
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Property

m_AccID = FormatField(rst("AccId"))
Set rst = Nothing

End Property

Public Property Let AccountId(NewValue As Long)
    m_AccID = NewValue
End Property
'This Function will Depist the Amount to The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
                ByVal Description As String, ByVal TransDate As Date, _
                Optional VoucherNo As String = "", _
                Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

'Check Whether Account Exist or Closed
Dim ClosedDate As String
If Not SBAccountExists(CLng(AccId), ClosedDate) Then GoTo ErrLine
Dim rst As ADODB.Recordset

'If it is already closed then
If Trim$(ClosedDate) <> "" Then GoTo ErrLine

'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
gDbTrans.SqlStmt = "Select TransId,TransDate,Balance From SBTrans " & _
    " Where AccId = " & AccId & " And TransId = " & _
    "(SELECT MAX(TransID) From SBTrans WHERE Accid = " & AccId & ")"

Dim TransID As Long
Dim Balance As Currency
Dim Particulars As String
Dim transType As wisTransactionTypes
TransID = 1

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    'Assign The Necessary Values
    TransID = rst("TransID")
    Balance = FormatField(rst("Balance"))
End If
Particulars = Description

'Query To Put The Acquired Values into Corresponding Table
TransID = TransID + 1
Balance = Balance + Amount

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

Dim UserID As Long
UserID = gCurrUser.UserID

gDbTrans.SqlStmt = "INSERT INTO SBTrans (Accid ,TransId," & _
         " Amount ,Balance, TransType, " & _
         " Particulars,TransDate,VoucherNo,UserId) " & _
         " VALUES (" & AccId & "," & TransID & "," & _
         Amount & "," & Balance & "," & transType & "," & _
         AddQuotes(Particulars, True) & "," & _
         " #" & TransDate & "#," & _
         AddQuotes(VoucherNo, True) & "," & UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'update it to te Contra Trans table
Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(GetResourceString(421))

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
    'insert into contra trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans" & _
        "(ContraID,AccHeadID,AccId," & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

DepositAmount = TransID

ErrLine:

End Function
'This Function will remove the last contra deposit to The Speicfied Account No
Public Function UndoContraDepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
                ByVal TransDate As Date) As Boolean

UndoContraDepositAmount = False
Dim TransID As Long
If Not IsContraTransactionRemovable(AccId, Amount, TransDate, TransID) Then Exit Function

gDbTrans.SqlStmt = "Delete * From SBTrans " & _
    " Where AccId = " & AccId & " And TransId = " & _
    "(SELECT MAX(TransID) From SBTrans WHERE Accid = " & AccId & ")"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'update it to te Contra Trans table
Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(GetResourceString(421))

    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID
    'insert into contra trans table
    gDbTrans.SqlStmt = "Delete * from ContraTrans " & _
        " where ContraID = (Select distinct contraID from ContraTrans " & _
        " Where AccHeadID = " & AccHeadID & " And AccID = " & AccId & _
        "  And TransID = " & TransID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

UndoContraDepositAmount = True

ErrLine:

End Function
Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean
    
    'Get last transaction record
    Dim Amount As Currency
    Dim lastTransID As Long
    Dim TransDate As Date
    Dim ChequeNo As Long
    Dim transType As wisTransactionTypes
    Dim rst As ADODB.Recordset
     
    gDbTrans.SqlStmt = "Select TOP 1 * from SBTrans where " & _
                " AccID = " & AccId & " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    Amount = FormatField(rst("Amount"))
    lastTransID = FormatField(rst("TransID"))
    TransDate = rst("TransDate")
    ChequeNo = FormatField(rst("ChequeNo"))
    transType = FormatField(rst("TransType"))
    
    If TransID <> lastTransID Then
        MsgBox GetResourceString(421, 530), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
'Delete record from Data base
    'If it has transaction in the PL Trans w.r.t this transaction
    'dletete that transaction also
    gDbTrans.SqlStmt = "Delete from SBPLTrans where AccID = " & AccId & _
            " AND TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then Exit Function
    
    'Delete From the Sb Trans Table
    gDbTrans.SqlStmt = "Delete from SBTrans where AccID = " & AccId & _
            " AND TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then Exit Function
    
    'Prepare the cheque list
    If transType = wContraWithdraw Or transType = wWithdraw Then
        If ChequeNo > 0 Then
            gDbTrans.SqlStmt = "UPDATE ChequeMaster Set Trans = " & wischqIssue & _
                " WHERE ChequeNo = " & ChequeNo & " AND AccID = " & AccId & _
                " AND AccHeadID = " & GetIndexHeadID(GetResourceString(421))
            If Not gDbTrans.SQLExecute Then Exit Function
        End If
    End If

    UndoTransaction = True

End Function

'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can be used in Standing Instuctions Presently this is Using by Clearing Class
Public Function WithdrawAmount(ByVal AccId As Long, ByVal Amount As Currency, _
                    Description As String, TransDate As Date, _
                    Optional VoucherNo As String = " ", _
                    Optional CashTransaction As Boolean, Optional ContraTransID As Long = 0) As Integer

If VoucherNo = "" Then VoucherNo = " "
'Check Whether Account Exist or Closed
Dim ClosedDate As String
If Not SBAccountExists(AccId, ClosedDate) Then GoTo ErrLine

'If it is already closed then
If Trim$(ClosedDate) <> "" Then GoTo ErrLine

Dim rst As ADODB.Recordset
Dim SqlStmt As String
'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
SqlStmt = "SELECT MAX(TransID) From SBTrans WHERE Accid = " & AccId & _
                    " And TransDate <= #" & TransDate & "# "
gDbTrans.SqlStmt = SqlStmt

Call gDbTrans.Fetch(rst, adOpenForwardOnly)

If rst Is Nothing Then GoTo ErrLine

Dim TransID As Long
TransID = FormatField(rst(0))

'Now Get the  Max TransID of the Account
gDbTrans.SqlStmt = "SELECT Top 1 TransID,Balance From SBTrans " & _
        " WHERE Accid = " & AccId & " Order By TransId Desc"

Call gDbTrans.Fetch(rst, adOpenForwardOnly)
If rst Is Nothing Then GoTo ErrLine

If TransID <> FormatField(rst(0)) Then GoTo ErrLine

Dim Balance As Currency

'Assign The Necessary Values
TransID = rst("TransID") + 1
Balance = Val(FormatField(rst("Balance"))) - Amount

If Balance < 0 Then
    'MsgBox "Insufficient balance in account for this transaction!", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(547), vbExclamation, gAppName & " - Error"
    GoTo ErrLine
End If

Dim transType As wisTransactionTypes
Dim UserID As Long

UserID = gCurrUser.UserID
transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SqlStmt = "INSERT INTO SBTrans (Accid, TransId, Amount," _
         & "Balance, TransType, Particulars,TransDate,VoucherNo,UserID) " _
         & " VALUES(" & AccId & "," & TransID & "," _
         & Amount & "," & Balance & "," & transType & "," _
         & AddQuotes(Description, True) & "," _
         & "#" & TransDate & "#," _
         & AddQuotes(VoucherNo, True) & "," & UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Get the Headid
Dim ClsBank As clsBankAcc
Dim AccHeadID As Long

    
If Not CashTransaction Then
    Set ClsBank = New clsBankAcc
    AccHeadID = ClsBank.GetHeadIDCreated(GetResourceString(421))
    Set ClsBank = Nothing
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID + 1
    'update it to te Contra Trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,UserId,VoucherNo)" & _
            " Values (" & ContraID & "," & AccHeadID & "," & _
            AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            UserID & "," & AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

WithdrawAmount = TransID

Exit Function

ErrLine:

End Function
 'This funtion returns the customer balance on given date ,His SB account No
'Arguments CustomerId, Id of the customet whose balance is needed
Public Function CustomerBalance(customerID As Long, retAccNum As String, AsOnDate As Date) As Currency
'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
'Sql stmt to get Accid
gDbTrans.SqlStmt = "Select A.AccID,AccNum,Balance from " & _
    " SBMaster A Inner Join SBTrans B On A.AccID = B.AccId " & _
    " where Customerid=" & customerID & " and TransID IN " & _
        "(Select Max(transID) From SBtrans C WHERE " & _
        " TransDate <= #" & AsOnDate & "# AND C.AccID =A.AccID)" & _
    " And (ClosedDate >= #" & AsOnDate & "# OR ClosedDate is NULL )"
    

If gDbTrans.Fetch(m_rstCustBalance, adOpenForwardOnly) < 1 Then Exit Function

retAccNum = Val(FormatField(m_rstCustBalance("AccNum")))
CustomerBalance = FormatField(m_rstCustBalance("Balance"))

ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If

End Function

Public Property Get ObjectName()
    ObjectName = "SBClass"
End Property


'THis Function Retuns the Next Account Balance of the customewr
'Before calling this function, we should call the CustomerBlance function
Public Function NextBalance(retAccNum As String) As Currency

'To trap an error
On Error GoTo ErrLine

NextBalance = 0

If m_rstCustBalance Is Nothing Then GoTo ExitLine
If m_rstCustBalance.EOF Then GoTo ExitLine
m_rstCustBalance.MoveNext
If m_rstCustBalance.EOF Then GoTo ExitLine

retAccNum = Val(FormatField(m_rstCustBalance("AccNum")))
NextBalance = FormatField(m_rstCustBalance("Balance"))

ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If

End Function

'
Private Function PLTransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency

PLTransactionDetails = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = " Select Sum(Amount) as Deposits from SBPLTrans where " & _
                   " TransType = " & transType & _
                   " And TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

If Not rst Is Nothing Then _
    PLTransactionDetails = Val(FormatField(rst("Deposits")))

End Function

Private Function TransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency

TransactionDetails = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = " Select Sum(Amount) as Deposits from SBTrans where " & _
                   " TransType = " & transType & _
                   " And TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

If Not rst Is Nothing Then _
    TransactionDetails = Val(FormatField(rst("Deposits")))

End Function
Public Sub TotalDepositTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from SBTrans where " & _
                   " TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#" & _
                   " Group By TransType"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim TotalAmount As Currency
While rst.EOF = False
    TotalAmount = Val(FormatField(rst("TotalAmount")))
    If rst("TransType") = wDeposit Then CashDeposit = TotalAmount
    If rst("TransType") = wWithdraw Then cashWithdraw = TotalAmount
    If rst("TransType") = wContraDeposit Then contraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then contraWithdraw = TotalAmount
    rst.MoveNext
Wend

Set rst = Nothing
End Sub
Public Sub TotalDepositIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from SBPLTrans where " & _
                   " TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#" & _
                   " Group By TransType"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim TotalAmount As Currency
While rst.EOF = False
    TotalAmount = Val(FormatField(rst("TotalAmount")))
    If rst("TransType") = wDeposit Then CashDeposit = TotalAmount
    If rst("TransType") = wWithdraw Then cashWithdraw = TotalAmount
    If rst("TransType") = wContraDeposit Then contraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then contraWithdraw = TotalAmount
    
    rst.MoveNext
Wend

Set rst = Nothing
End Sub

Public Sub Show()

Dim Perms As wis_Permissions
Dim StartSB As Boolean

StartSB = False

If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers: Exit Sub

'Now Check the permission
'whether logged inperson is having permission of sb account or not
Perms = gCurrUser.UserPermissions
If Perms And perBankAdmin Then StartSB = True
If Perms And perCreateAccount Then StartSB = True
If Perms And perCashier Then StartSB = True
If Perms And perClerk Then StartSB = True
If Perms And perReadOnly Then StartSB = True

If StartSB Then
    If m_frmSBAcc Is Nothing Then Set m_frmSBAcc = New frmSBAcc
    Load m_frmSBAcc
    If gWindowHandle Then m_ParentHandle = gWindowHandle
    gWindowHandle = m_frmSBAcc.hwnd
    m_frmSBAcc.Show
    If m_AccID > 0 Then
        With m_frmSBAcc
            .AccountLoad (m_AccID)
            .txtAccNo.Locked = True
            .cmdReset.Enabled = False
        End With
    End If
Else
    MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
End If

End Sub

'
Public Function Deposits(fromDate As Date, toDate As Date) As Currency
'Deposit inflow to the bank between 2  given dates
    
    Dim transType As wisTransactionTypes
    Deposits = 0
    transType = wDeposit
    Deposits = TransactionDetails(fromDate, toDate, transType)

End Function

'
Public Function Loss(fromDate As Date, toDate As Date) As Currency
'this is the amount withdrawn in the sb Pl  trans table
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes

Dim Amount As Currency
Amount = 0

'AMOUNT DEPOSITED WITH CONTRA
transType = wWithdraw
ContraTransType = wContraWithdraw

Amount = PLTransactionDetails(fromDate, toDate, transType)
Amount = Amount + PLTransactionDetails(fromDate, toDate, ContraTransType)

Loss = Amount
End Function

'
Public Function ContraWithdrawls(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons
ContraWithdrawls = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

'Contra Withdrwals
transType = wContraWithdraw
Amount = TransactionDetails(fromDate, toDate, transType)
ContraWithdrawls = Amount
End Function

'
Public Function ContraDeposits(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons

ContraDeposits = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency
'Cash Withdrwals
transType = wContraDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
ContraDeposits = Amount

End Function

Public Function CashDeposits(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons

CashDeposits = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency
transType = wDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
CashDeposits = Amount
End Function


'
Public Function CashWithdrawls(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons
    
    CashWithdrawls = 0
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    'Cash Withdrwals
    transType = wWithdraw
    Amount = TransactionDetails(fromDate, toDate, transType)
    CashWithdrawls = Amount

End Function


'
Public Function WithDrawls(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons

    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    'Contra Withdrawls
    transType = wWithdraw
    Amount = TransactionDetails(fromDate, toDate, transType)
    
    'Cash Withdrwals
    transType = wContraWithdraw
    Amount = Amount + TransactionDetails(fromDate, toDate, transType)
    WithDrawls = Amount
End Function

'
Public Function Profit(fromDate As Date, toDate As Date) As Currency
    ' Profit to the bank due to  deposits
    'THIS THE AMOUNT DEPOSITED IN SBPLTRANS TABLE
    
    Dim transType As wisTransactionTypes
    Dim ContraTransType As wisTransactionTypes
    
    Dim Amount As Currency
    Amount = 0
    
    'AMOUNT DEPOSITED WITH CONTRA
    transType = wDeposit
    ContraTransType = wContraDeposit
    
    Amount = PLTransactionDetails(fromDate, toDate, transType)
    Amount = Amount + PLTransactionDetails(fromDate, toDate, ContraTransType)
    
    Profit = Amount

End Function

Public Function GetAccountID(AccNum As String) As Integer

GetAccountID = 0

Dim ret As Integer
Dim rst As ADODB.Recordset

'Query Database
    gDbTrans.SqlStmt = "Select * from SBMaster where AccNum = " & AddQuotes(AccNum, True)
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    
GetAccountID = FormatField(rst("AccID"))
End Function



'
Public Function Balance(AsOnDate As Date) As Currency
    Balance = 0
    Balance = ComputeTotalSBLiability(AsOnDate)
End Function

Private Sub Class_Initialize()

m_TransDate = IIf(Len(gStrDate), gStrDate, Format(Now, "mm/dd/yyyy"))

End Sub

Private Sub Class_Terminate()

Set m_frmSBReport = Nothing
Set m_frmSBAcc = Nothing

End Sub


Private Sub m_frmSBAcc_AccountChanged(AccId As Long)

If m_AccID = AccId Then Exit Sub
m_AccID = AccId

End Sub

Private Sub m_frmSBAcc_AccountTransaction(transType As wisTransactionTypes, Cancel As Integer)

Cancel = True
If AccountTransaction(transType) Then Cancel = False

End Sub

Private Function AccountTransaction(transType As wisTransactionTypes) As Boolean

Dim AccountCloseFlag As Boolean
Dim ClosedON As String
Dim TransDate As Date
Dim ret As Integer
Dim Amount As Currency
Dim ChequeNo As Long
Dim VoucherNo As String

Dim Particulars As String
Dim Balance As Currency
Dim TransID As Long
Dim MinBalance As Currency
Dim SetUp As New clsSetup
Dim ParticularsArr() As String
Dim I As Integer
Dim ChequeArr() As Long
Dim rst As ADODB.Recordset

'Check if account exists
If Not SBAccountExists(m_AccID, ClosedON) Then
    'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If

If ClosedON <> "" Then
    'MsgBox "This account has been closed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(524), vbExclamation, gAppName & " - Error"
    Exit Function
End If

TransDate = m_TransDate 'FormatDate(m_frmSBAcc.txtDate.Text)

'Check if the date of transaction is earlier than account opening date itself
gDbTrans.SqlStmt = "Select * from SBMaster where AccID = " & m_AccID
ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If ret <> 1 Then
    'MsgBox "DB error !", vbCritical, gAppName & " - ERROR"
    MsgBox GetResourceString(601), vbCritical, gAppName & " - ERROR"
    Exit Function
End If

If DateDiff("d", TransDate, rst("CreateDate")) > 0 Then
    'MsgBox "Date of transaction is earlier than the date of account creation itself !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(568), vbExclamation, gAppName & " - Error"
    Exit Function
End If


With m_frmSBAcc
    Amount = .txtAmount
    TransDate = GetSysFormatDate(.txtDate.Text)
    'Validate the Cheque No
    Dim IntBalance As Currency
    'Get the Minimum Balance
    If .cmbCheque.ListCount <= 1 Then  'No Cheque issued
        MinBalance = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithoutChequeBook", "0.00")
    Else
        MinBalance = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithChequeBook", "0.00")
    End If
    If transType = wWithdraw Or transType = wContraWithdraw Then
        With .cmbCheque
            If .ListIndex < 0 Then
                ChequeNo = 0
            Else
                'ChequeNo = .List(.ListIndex)
                ChequeNo = .ItemData(.ListIndex)
            End If
        End With
    Else 'If Trans = wDeposit Or Trans = wContraDeposit Then
        ChequeNo = 0
    End If
    
    VoucherNo = Trim$(.txtVoucherNo.Text)
    Particulars = Trim$(.cmbParticulars.Text)
    Dim lstIndex As Integer
    lstIndex = .cmbTrans.ListIndex
End With

'Get the Balance and new transid
gDbTrans.SqlStmt = "Select TOP 1 * from SBTrans " & _
            " where AccID = " & m_AccID & _
            " order by TransID DESC"

Balance = 0
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then Balance = FormatField(rst("Balance"))

'Opening balance or Depositamount should be specified
If Balance = 0 And Amount = 0 Then
    MsgBox "Either DepositAmount or opening balance sholud be specified ", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'now Get the Tranaction ID
TransID = GetSBMaxTransID(m_AccID) + 1
'See if the date is earlier than last date of transaction
If DateDiff("D", TransDate, GetSBLastTransDate(m_AccID)) > 0 Then
    'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    Exit Function
End If


'Check if Transaction is possible (Check any other constraints here)
If transType = wWithdraw Or transType = wContraWithdraw Then
    If Balance - Amount < 0 Then
        'MsgBox "Insufficient balance in account for this transaction!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(547), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    If Balance - Amount < MinBalance Then
        'If MsgBox("This transaction will cause the balance to go below the minimum balance limit" & _
                vbCrLf & vbCrLf & "Do you want to continue with this transaction ?", vbYesNo + vbQuestion, gAppName & " - Confirmation") = vbNo Then
        If MsgBox(GetResourceString(630) & _
            vbCrLf & vbCrLf & GetResourceString(541), _
            vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    End If
End If

'Calculate new balance
If transType = wDeposit Or transType = wContraDeposit Then
    Balance = Balance + Amount
Else
    Balance = Balance - Amount
End If

'Perform the Transaction to the Database
Dim InTrans As Boolean
Dim bankClass As clsBankAcc

gDbTrans.BeginTrans
InTrans = True

Set bankClass = New clsBankAcc
Dim SbHeadID As Long

If lstIndex = 2 Or lstIndex = 3 Then  'Interest or charges transaction
    
    'If transaction is affecting the Profit or loss of the bank then
    'Insert the record into SBPLTrans
    'Get the Contra Transaction ID
    Dim SbIntHeadID As Long
    Dim headName As String
    Dim SqlInt As String

    headName = GetResourceString(421, 487)
    SbIntHeadID = bankClass.GetHeadIDCreated(headName, LoadResourceStringS(421, 487), parMemDepIntPaid, 0, wis_SBAcc)
    Dim Trans As Byte
    Trans = IIf(transType = wContraDeposit, wContraWithdraw, wContraDeposit)
    
    'Here the transaction will be reverse of the SbTrans
    IntBalance = IntBalance + (Amount * IIf(Trans = wContraDeposit, 1, -1))
    
    SqlInt = "Insert into SBPLTrans (AccID, TransID, TransDate, Amount, " & _
                " Balance, Particulars, TransType,VoucherNO,UserID) values ( " & _
                m_AccID & "," & _
                TransID & "," & _
                "#" & TransDate & "#," & _
                Amount & "," & IntBalance & "," & _
                AddQuotes(Particulars, True) & "," & _
                Trans & "," & _
                AddQuotes(VoucherNo, True) & "," & _
                gCurrUser.UserID & ")"

End If


headName = GetResourceString(421)
SbHeadID = bankClass.GetHeadIDCreated(headName, LoadResString(421), parMemberDeposit, 0, wis_SBAcc)

'Perform the tranaction in the Sb Head
If lstIndex = 0 Then
    If DepositAmount(m_AccID, Amount, Particulars, _
                TransDate, VoucherNo, True) < 1 Then GoTo ExitLine
    If Not bankClass.UpdateCashDeposits(SbHeadID, Amount, TransDate) Then GoTo ExitLine

'Perform the tranaction in the Sb Head
ElseIf lstIndex = 1 Then
    If WithdrawAmount(m_AccID, Amount, Particulars, _
            TransDate, VoucherNo, True) < 1 Then GoTo ExitLine
    If Not bankClass.UpdateCashWithDrawls(SbHeadID, Amount, TransDate) Then GoTo ExitLine

ElseIf lstIndex = 3 Then

    'Deposit the Amount to SB
    If DepositAmount(m_AccID, Amount, Particulars, _
                TransDate, VoucherNo, False) < 1 Then GoTo ExitLine
    
    'Now Insert Into the P&L
    gDbTrans.SqlStmt = SqlInt
    
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    If Not bankClass.UpdateContraTrans(SbIntHeadID, _
                SbHeadID, Amount, TransDate) Then GoTo ExitLine

ElseIf lstIndex = 2 Then
    If WithdrawAmount(m_AccID, Amount, Particulars, _
                TransDate, VoucherNo, False) < 1 Then GoTo ExitLine
    
    gDbTrans.SqlStmt = SqlInt
    'Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    'Now to the ledger head
    If Not bankClass.UpdateContraTrans(SbHeadID, _
                SbIntHeadID, Amount, TransDate) Then GoTo ExitLine
    
End If

If lstIndex >= 2 Then
    'Transaction Is of the Interest head
    'Therefore Remove the Transaction int Contra Table
    gDbTrans.SqlStmt = "Delete * From ContraTrans " & _
                " Where ContraID = " & GetMaxContraTransID
    
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
End If


Set bankClass = Nothing

'Update Cheque table
If ChequeNo <> 0 And lstIndex = 1 Then
    gDbTrans.SqlStmt = "Update ChequeMaster set Trans = " & wischqPay & "," & _
        "TransDate = #" & TransDate & "#, Amount = " & Amount & "," & _
        "Particulars = " & AddQuotes(Particulars, True) & _
        " Where AccID = " & m_AccID & " AND AccHeadID = " & SbHeadID & _
        " AND ChequeNo = " & ChequeNo
                        
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
End If
 
'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(SbHeadID, m_AccID, _
                TransDate, TransID, Amount) < 1 Then GoTo ExitLine
    Set Cashclass = Nothing
End If


'Now you have to see if you want to close the account
'We shall close the account if no balance is maintained Like how they do in
'Alnavar Bank, Dharwar
If Balance = 0 Then
    'Prompt for closing the account
    'If MsgBox("This transaction will result in a balance of Rs. 0 " & _
                " and hence the account will be closed !" & vbCrLf & _
                " Do you want to continue with the transaction ?", vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
    If MsgBox(GetResourceString(630) & _
        GetResourceString(631) & vbCrLf & _
        GetResourceString(541), vbQuestion + vbYesNo, gAppName & _
        " - Confirmation") = vbYes Then AccountCloseFlag = True
End If

If AccountCloseFlag = True Then
    If Not AccountClose() Then
       ' MsgBox "Unable to close the account !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(534), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
End If

gDbTrans.CommitTrans
InTrans = False

AccountTransaction = True

Exit Function

ExitLine:

If InTrans Then gDbTrans.RollBack: InTrans = False

Set Cashclass = Nothing
Set bankClass = Nothing

End Function

Private Function AccountClose() As Boolean
Dim ret As Integer
Dim AccNo As Long

'Prelim checks
AccNo = m_AccID
If AccNo <= 0 Then Exit Function

'Check if account exists
If Not SBAccountExists(m_AccID) Then Exit Function
    

gDbTrans.SqlStmt = "Update SBMaster set ClosedDate = #" & m_TransDate & "#" & _
        " Where AccID = " & AccNo

If Not gDbTrans.SQLExecute Then Exit Function

AccountClose = True

End Function


Private Sub m_frmSBAcc_DateChanged(IndianDate As String)
If Not DateValidate(IndianDate, "/", True) Then Exit Sub
m_TransDate = GetSysFormatDate(IndianDate)
End Sub

Private Sub m_frmSBAcc_ShowReport(ReportType As wis_SBReports, ReportOrder As wis_ReportOrder, _
            fromDate As String, toDate As String, RepOption As clsRepOption)
        
        If Not m_frmSBReport Is Nothing Then Unload m_frmSBReport
        Set m_frmSBReport = New frmSBReport
        
        With m_frmSBReport
            .ReportOrder = ReportOrder
            .ReportType = ReportType
            .FromIndianDate = fromDate
            .ToIndianDate = toDate
            
            .FromAmount = RepOption.FromAmount
            .ToAmount = RepOption.ToAmount
            .Place = RepOption.Place
            .Caste = RepOption.Caste
            .Gender = RepOption.Gender
            .AccountGroup = RepOption.AccountGroupID
        End With
        
        gCancel = 0
        frmCancel.Show
        frmCancel.Refresh
        Load m_frmSBReport
        Unload frmCancel
        Set frmCancel = Nothing
        If gCancel = 2 Then
            Unload m_frmSBReport
            Set m_frmSBReport = Nothing
            Exit Sub
        End If
        
        If Not gCancel Then m_frmSBReport.Show 'vbModal
        Set frmCancel = Nothing

End Sub

Private Sub m_frmSBAcc_WindowClosed()

If m_ParentHandle Then gWindowHandle = m_ParentHandle
m_ParentHandle = 0

Set m_frmSBAcc = Nothing
End Sub


Public Property Get ObjectHandle() As Long

    If Not m_frmSBAcc Is Nothing Then _
            ObjectHandle = m_frmSBAcc.hwnd: Exit Property
    If Not m_frmSBReport Is Nothing Then _
        Unload m_frmSBReport: Set m_frmSBReport = Nothing

End Property

Private Sub m_frmSBReport_Initialise(Min As Long, Max As Long)
gCancel = 0
frmCancel.Show
    If Max <> 0 Then
        With frmCancel
            UpdateStatus frmCancel.PicStatus, 0, True
'            .prg.Visible = True
'            .prg.Min = Min
'            If Max > 32500 Then Max = 32500
'            .prg.Max = Max
        End With
    End If
End Sub


Private Sub m_frmSBReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS :" & vbCrLf & strMessage
    UpdateStatus frmCancel.PicStatus, Ratio
'    If Ratio > 0 Then
'        If Ratio > 1 Then Ratio = Ratio / .prg.Max
'        .prg.Value = .prg.Max * Ratio
'    End If
End With

End Sub
''This Function Checks whehter the Transaction ss matches with any given amount and Date and contra Deposit
Public Function IsContraTransactionRemovable(AccId As Long, Amount As Currency, TransDate As Date, Optional TransID As Long) As Boolean

    
    IsContraTransactionRemovable = False
    
    Dim ClosedDate As String
    If Not SBAccountExists(CLng(AccId), ClosedDate) Then Exit Function
    If Trim$(ClosedDate) <> "" Then Exit Function
    
    gDbTrans.SqlStmt = "Select Amount,TransDate,TransID, TransType from SBTrans " & _
            " Where AccID = " & AccId & " order by transID desc"
                
    Dim rstFetch As Recordset
    If gDbTrans.Fetch(rstFetch, adOpenForwardOnly) > 0 Then
        
        If Not IsMissing(TransID) Then TransID = FormatField(rstFetch("TransID"))

        If FormatField(rstFetch("Amount")) <> Amount Then Exit Function
        If FormatField(rstFetch("TransDate")) <> TransDate Then Exit Function
        
        Dim sbTransType As wisTransactionTypes
        sbTransType = FormatField(rstFetch("TransType"))
        If sbTransType = wDeposit Then Exit Function
        If sbTransType = wWithdraw Then Exit Function
        
        IsContraTransactionRemovable = True
    End If
    
End Function

