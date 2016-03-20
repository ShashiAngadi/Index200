VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_rstAcc As ADODB.Recordset

Private WithEvents m_frmRDAcc As frmRDAcc
Attribute m_frmRDAcc.VB_VarHelpID = -1
Private WithEvents m_frmRDReport As frmRDReport
Attribute m_frmRDReport.VB_VarHelpID = -1
Private m_AccID As Long
Private m_ParentHandle As Long

Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset

gDbTrans.SqlStmt = "Select AccID From RDMAster " & _
    "Where CustomerID = " & NewValue
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Property

m_AccID = FormatField(rst("AccID"))
Set rst = Nothing


End Property


Public Property Get ObjectName()
    ObjectName = "RDClass"
End Property


Public Property Let AccountId(NewValue As Long)
m_AccID = NewValue

End Property



Public Function CashDeposits(fromDate As Date, toDate As Date) As Currency
' Total recurring Deposit deposited where physical cash handled
Dim transType As wisTransactionTypes
CashDeposits = 0

transType = wDeposit
CashDeposits = TransactionDetails(fromDate, toDate, transType)

End Function


Public Function CashWithdrawls(fromDate As Date, toDate As Date) As Currency
'Recurring Deposit Given back to the customer in the form of cash
CashWithdrawls = 0

Dim transType As wisTransactionTypes

transType = wWithdraw
CashWithdrawls = TransactionDetails(fromDate, toDate, transType)

End Function

Public Function ContraDeposits(fromDate As Date, toDate As Date) As Currency
' Total Fixed Deposit flown in the bank between 2 given dates

Dim transType As wisTransactionTypes
ContraDeposits = 0

transType = wContraDeposit
ContraDeposits = TransactionDetails(fromDate, toDate, transType)

End Function


Public Function ContraWithdrawls(fromDate As Date, toDate As Date) As Currency
'Recurring Deposit transferred to other accounts
'Where cash ha not gone out of the bank

ContraWithdrawls = 0

Dim transType As wisTransactionTypes
transType = wContraWithdraw
ContraWithdrawls = TransactionDetails(fromDate, toDate, transType)

End Function

Public Function InterestOnDeposits(AsOnDate As Date) As Currency
   On Error GoTo ErrLine
  'declare the variables necessary
  
Dim Retval As Long
Dim rstMain As Recordset
Dim AccId As Long
Dim InterestRate As Currency
Dim LastIntDate As Date
Dim CreateDate As Date
Dim MatDate As Date
Dim Duration As Long
Dim IntAmount As Currency
Dim Balance As Currency

Dim TotalAmount As Currency


'Build The Querry
gDbTrans.SqlStmt = " SELECT Balance, A.AccId,  CreateDate, MaturityDate, TransDate, RateOfInterest" & _
         " From RDMaster A, RDTrans B Where A.AccId = B.AccId " & _
    "  And (ClosedDate is NULL or ClosedDate >= #" & AsOnDate & "#)" & _
    "  And CreateDate < #" & AsOnDate & "# And TransID =  " & _
        " ( Select max(TransID) From RDTrans C Where A.AccId = C.AccId  " & _
        " and transDate <= #" & AsOnDate & "# ) " & _
    " Order By A.AccId "
 
Retval = gDbTrans.Fetch(rstMain, adOpenForwardOnly)
If Retval < 1 Then GoTo ExitLine

While Not rstMain.EOF
    AccId = Val(FormatField(rstMain("AccId")))
    Balance = CCur(FormatField(rstMain("Balance")))
    LastIntDate = rstMain("TransDate")
    CreateDate = rstMain("CreateDate")
    InterestRate = CCur(FormatField(rstMain("RateofInterest")))
    MatDate = rstMain("MaturityDate")
    
    'Now Get The Date Difference
    ' On that Difference Get the InterestRate
    Duration = DateDiff("d", LastIntDate, AsOnDate)
    If Duration <= 0 Then GoTo NextDeposit
    
    If InterestRate = 0 Then
       MsgBox "Function for RD INterest"
       'InterestRate = GetRDDepositInterest(CreateDate)
       InterestRate = 7
    End If
    
    Duration = DateDiff("d", LastIntDate, AsOnDate)
    If Duration <= 0 Then GoTo NextDeposit
    IntAmount = ((InterestRate / 100) * Balance * Duration) / 365
    TotalAmount = CCur(TotalAmount + IntAmount)
    
NextDeposit:
    rstMain.MoveNext
Wend
InterestOnDeposits = TotalAmount
Set rstMain = Nothing

GoTo ExitLine

ErrLine:
    MsgBox "Error In RDClass--Interest On Deposits", vbCritical, wis_MESSAGE_TITLE
    Err.Clear
    
ExitLine:
End Function

Public Function InterestAmount(ByVal AccountId As Long, ByVal AsOnDate As Date)

InterestAmount = ComputeRDDepositInterestAmount(AccountId, AsOnDate)

End Function

Public Function InterestOnDepositsAdded(AsOnDate As Date) As Currency
   On Error GoTo ErrLine
  'declare the variables necessary
  
Dim Retval As Long
Dim rstMain As ADODB.Recordset
Dim TotalAmount As Currency

gDbTrans.SqlStmt = "SELECT Sum(Amount) AS TOTAL " & _
    "From AccTrans WHERE AccID = 13003 And TransType =  " & wContraWithdraw & _
    "And TransDate <=  #" & AsOnDate & "#"
 
Retval = gDbTrans.Fetch(rstMain, adOpenForwardOnly)
If Retval < 1 Then GoTo ExitLine
TotalAmount = FormatField(rstMain("TOTAL"))

gDbTrans.SqlStmt = "SELECT Sum(Amount) AS TOTAL " & _
        "From AccTrans WHERE AccID = 13003 And TransType =  " & wContraDeposit & _
        "And TransDate <=  #" & AsOnDate & "#"
Retval = gDbTrans.Fetch(rstMain, adOpenForwardOnly)
If Retval < 1 Then GoTo ExitLine

TotalAmount = TotalAmount - FormatField(rstMain("TOTAL"))
Set rstMain = Nothing

InterestOnDepositsAdded = TotalAmount
GoTo ExitLine

ErrLine:
    MsgBox "Error In RDClass--Interest On Deposits", vbCritical, wis_MESSAGE_TITLE
    Err.Clear
    
ExitLine:
End Function

'This Function wil Deposit the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
                        ByVal Description As String, ByVal TransDate As Date, _
                        Optional VoucherNo As String = " ", _
                        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

Dim rst As ADODB.Recordset

'Check Whether Account Exist or Closed
gDbTrans.SqlStmt = "Select AccID From RDMAster WHERE Accid = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine

gDbTrans.SqlStmt = "Select TransId,Balance,TransDate From RDTrans where " & _
    " AccId = " & AccId & " And TransID = (SELECT MAX(TransID) " & _
    " From RdTrans WHERE Accid = " & AccId & ")"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine

'Check For the ttransaction date
If DateDiff("D", TransDate, rst("Transdate")) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency

'Assign The Necessary Values
TransID = Val(FormatField(rst("TransId"))) + 1
Balance = Val(FormatField(rst("Balance")))

Dim transType As wisTransactionTypes

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

Balance = Balance + Amount

Dim UserID As Long
UserID = gCurrUser.UserID

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SqlStmt = "INSERT INTO RDTrans " & _
        " (Accid ,TransId ,Amount ,Balance," & _
        " TransType, Particulars,TransDate,UserID,VoucherNo) " & _
        " VALUES(" & AccId & "," & TransID & "," & _
        Amount & "," & Balance & "," & _
        transType & "," & AddQuotes(Description, True) & "," & _
        "#" & TransDate & "#," & _
        UserID & "," & AddQuotes(VoucherNo, True) & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(GetResourceString(424)) '"Recurrring Deposit

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
    
    'update it to te Contra Trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,userid,VoucherNo)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & UserID & "," & _
        AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

DepositAmount = TransID

Exit Function

ErrLine:

End Function

Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean

'Prelim check
Dim Trans As Byte
Dim TransDate As Date


'Get last transaction record
Dim Amount As Currency
Dim ret As Integer
Dim lastTransID As Long
Dim rst As ADODB.Recordset
    
gDbTrans.SqlStmt = "Select TOP 1 * from RDTrans where AccID = " & AccId & _
                " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
Amount = FormatField(rst("Amount"))
lastTransID = FormatField(rst("TransID"))
Trans = FormatField(rst("TransType"))
TransDate = rst("TransDate")

gDbTrans.SqlStmt = "Select TOP 1 * from RDIntTrans where AccID = " & AccId & _
                " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
If lastTransID < rst("TransID") Then
    Amount = FormatField(rst("Amount"))
    lastTransID = rst("TransID")
ElseIf lastTransID = rst("TransID") Then
    Amount = FormatField(rst("Amount"))
    lastTransID = rst("TransID")
End If

gDbTrans.SqlStmt = "Select TOP 1 * from RDIntPayable where AccID = " & AccId & _
                " order by TransID desc"
ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If ret >= 1 Then
    If lastTransID < rst("TransID") Then
        Amount = FormatField(rst("Amount"))
        lastTransID = rst("TransID")
    ElseIf lastTransID = rst("TransID") Then
        Amount = FormatField(rst("Amount"))
        lastTransID = rst("TransID")
    End If
End If

If TransID <> lastTransID Then
    MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Delete record from Data base
gDbTrans.SqlStmt = "Delete from RDTrans where " & _
        " AccID = " & AccId & " and TransID = " & lastTransID
If Not gDbTrans.SQLExecute Then Exit Function

'Un do the interest paid/received
gDbTrans.SqlStmt = "Delete from RDIntTrans where " & _
            " AccID = " & AccId & " and TransID = " & lastTransID
If Not gDbTrans.SQLExecute Then Exit Function
'Un do the interest Payable Withdraw or deposit if any
gDbTrans.SqlStmt = "Delete from RDIntPayable where " & _
            " AccID = " & AccId & " and TransID = " & lastTransID
If Not gDbTrans.SQLExecute Then Exit Function

Set rst = Nothing
UndoTransaction = True
    
End Function

'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function WithdrawAmount(ByVal AccId As Long, Amount As Currency, _
                Description As String, TransDate As Date, Optional VoucherNo As String, _
                Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

Dim rst As ADODB.Recordset

'Check Whether Account Exist or Closed
gDbTrans.SqlStmt = "Select TransId,Balance,TransDate From RDTrans " & _
                " where AccId = " & AccId & " And TransID = " & _
                "(SELECT MAX(TransID)From RdTrans" & _
                    " WHERE Accid = " & AccId & ")"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ErrLine

'Check For the ttransaction date
If DateDiff("D", TransDate, rst("Transdate")) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency
Dim Particulars As String
Dim transType As wisTransactionTypes

'Assign The Necessary Values
TransID = Val(FormatField(rst("TransId"))) + 1
Balance = Val(FormatField(rst("Balance")))
Set rst = Nothing
Particulars = Description

transType = IIf(CashTransaction, wDeposit, wContraDeposit)
If Balance < Amount Then GoTo ErrLine

Dim UserID As Long
UserID = gCurrUser.UserID

'Query To Put The Acquired Values into Corresponding Table
transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)
Balance = Balance - Amount

gDbTrans.SqlStmt = "INSERT INTO RDTrans " & _
            " (Accid ,TransId ,Amount ,Balance," & _
            " TransType, Particulars,TransDate,UserID ) " & _
            " VALUES (" & AccId & "," & TransID & "," & _
            Amount & "," & Balance & "," & transType & "," & _
            AddQuotes(Particulars) & "," & _
            " #" & TransDate & "# ," & UserID & ")"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

Dim ClsBank As clsBankAcc
Dim AccHeadID As Long

Set ClsBank = New clsBankAcc
AccHeadID = ClsBank.GetHeadIDCreated(GetResourceString(424))
Set ClsBank = Nothing
If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID + 1
    
    Set rst = Nothing
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        "Values (" & ContraID & "," & AccHeadID & "," & _
        AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

WithdrawAmount = TransID

Exit Function

ErrLine:

End Function

Public Function NextBalance(retAccNum As String) As Currency
'To trap an error
On Error GoTo ErrLine

NextBalance = 0

'Sql stmt to get Accid
If m_rstAcc Is Nothing Then GoTo ExitLine
m_rstAcc.MoveNext
If m_rstAcc.EOF Then GoTo ExitLine

retAccNum = FormatField(m_rstAcc("AccNum"))
NextBalance = FormatField(m_rstAcc("Balance"))
 
ExitLine:
    Exit Function
    
ErrLine:
    If Err Then
        MsgBox "RDAccount NextBalance: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitLine
    End If

End Function

Public Function CustomerBalance(CustomerID As Long, retAccNum As String, AsOnDate As Date) As Currency
'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
'Sql stmt to get Accid
gDbTrans.SqlStmt = "Select A.Accid,AccNum,Balance from RDMaster A Inner Join RDTrans B " & _
        " On A.AccID = B.AccID Where " & _
        " Customerid = " & CustomerID & _
        " AND TransID = (select MAx(transID) From RdTrans C Where " & _
            " C.AccId= A.Accid And TransDate <= #" & AsOnDate & "# )" & _
        " AnD Balance > 0"
        
If gDbTrans.Fetch(m_rstAcc, adOpenStatic) < 1 Then
    retAccNum = ""
    Exit Function
End If

retAccNum = FormatField(m_rstAcc("AccNum"))
CustomerBalance = FormatField(m_rstAcc("Balance"))
 
ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "RDAccountBalance: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitLine
    End If

End Function

Public Function Liability(AsOnDate As Date) As Currency
' Computes the Liability (Deposit plus interest on Deposit)
Dim SetUp As New clsSetup
Dim LiabBeforeInterest As Currency
Dim InterestOnDeposit As Double

Liability = 0

LiabBeforeInterest = ComputeTotalRDLiability(AsOnDate)

InterestOnDeposit = CDbl(SetUp.ReadSetupValue("RDacc", "Interest On RDDeposit", "12"))

Liability = ComputeRDInterest(LiabBeforeInterest, InterestOnDeposit)

End Function

Private Function ProfitLossDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency
ProfitLossDetails = 0

Dim SqlStr As String
Dim rst As ADODB.Recordset

SqlStr = "Select Sum(Amount) as Deposits from RDIntTrans where " & _
                " TransType = " & transType & _
                " And TransDate >= #" & fromDate & "#" & _
                " And TransDate <=#" & toDate & "#"
                
gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

    ProfitLossDetails = Val(FormatField(rst("Deposits")))

End Function

Private Function TransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency
TransactionDetails = 0

Dim SqlStr As String
Dim rst As ADODB.Recordset

SqlStr = "Select Sum(Amount) as Deposits from RDTrans where " & _
                " TransType = " & transType & _
                " And TransDate >= #" & fromDate & "#" & _
                " And TransDate <=#" & toDate & "#"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    TransactionDetails = Val(FormatField(rst("Deposits")))

End Function

Public Sub Show()
Dim Perms As wis_Permissions
Dim StartRD As Boolean

StartRD = False

If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers
    
    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Perms = gCurrUser.UserPermissions
    If Perms And perBankAdmin Then StartRD = True
    If Perms And perCreateAccount Then StartRD = True
    If Perms And perClerk Then StartRD = True
    If Perms And perCashier Then StartRD = True
    If Perms And perReadOnly Then StartRD = True
    
    If StartRD Then
        If m_frmRDAcc Is Nothing Then Set m_frmRDAcc = New frmRDAcc
        Load m_frmRDAcc
        If gWindowHandle Then m_ParentHandle = gWindowHandle
        gWindowHandle = m_frmRDAcc.hwnd
        m_frmRDAcc.Show 'vbModal
        If m_AccID Then m_frmRDAcc.AccountLoad (m_AccID)
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
    End If
    
End Sub

Public Function Deposits(fromDate As Date, toDate As Date) As Currency
' Total Fixed Deposit flown in the bank between 2 given dates

Dim transType As wisTransactionTypes
Dim DepAmount As Currency
Deposits = 0


transType = wDeposit
DepAmount = TransactionDetails(fromDate, toDate, transType)
transType = wContraDeposit
DepAmount = DepAmount + TransactionDetails(fromDate, toDate, transType)
Deposits = DepAmount
End Function

Public Function Loss(fromDate As Date, toDate As Date) As Currency

Loss = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wWithdraw
Amount = ProfitLossDetails(fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + ProfitLossDetails(fromDate, toDate, transType)

Loss = Amount

End Function

Public Function WithDrawls(fromDate As Date, toDate As Date) As Currency

'Recurring Deposit Given back to the customer
WithDrawls = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wWithdraw
Amount = TransactionDetails(fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + TransactionDetails(fromDate, toDate, transType)

WithDrawls = Amount

End Function

Public Function Profit(fromDate As Date, toDate As Date) As Currency

' profit to bank thru deposits

    Profit = 0
    
    Dim transType As wisTransactionTypes
    Dim ProfitAmount As Currency
    
    transType = wContraDeposit
    ProfitAmount = ProfitLossDetails(fromDate, toDate, transType)
    transType = wDeposit
    ProfitAmount = ProfitAmount + ProfitLossDetails(fromDate, toDate, transType)
    
    Profit = ProfitAmount
    
End Function

Public Function Balance(AsOnDate As Date) As Currency
' fixed deposits in the bank as on the given date
Balance = 0
Balance = ComputeTotalRDLiability(AsOnDate)
End Function
Public Sub TotalDepositTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from RDTrans where " & _
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
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from RDIntTrans where " & _
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

Public Sub TotalLoanTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)
    Dim depLoan As New clsDepLoan
    Call depLoan.TotalLoanTransactions(fromDate, toDate, CashDeposit, contraDeposit, cashWithdraw, contraWithdraw, wis_RDAcc)
    Set depLoan = Nothing
End Sub
Public Sub TotalLoanIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)
    Dim depLoan As New clsDepLoan
    Call depLoan.TotalLoanIntTransactions(fromDate, toDate, CashDeposit, contraDeposit, cashWithdraw, contraWithdraw, wis_RDAcc)
    Set depLoan = Nothing
End Sub


Private Sub Class_Terminate()
On Error Resume Next

'If Not m_frmRDACC Is Nothing Then
'    Unload m_frmRDACC
'    Set m_frmRDACC = Nothing
'End If

End Sub


Private Sub m_frmRDAcc_AccountChanged(ByVal AccId As Long)
m_AccID = AccId
End Sub

Private Sub m_frmRDAcc_ShowReport(ReportType As wis_RDReports, ReportOrder As wis_ReportOrder, _
            fromDate As String, toDate As String, _
            RepOption As clsRepOption)
    gCancel = 0
    frmCancel.Show
    frmCancel.Refresh
    
    If Not m_frmRDReport Is Nothing Then Unload m_frmRDReport
    Set m_frmRDReport = New frmRDReport
    
    With m_frmRDReport
        .ReportType = ReportType
        .ReportOrder = ReportOrder
        .FromIndianDate = fromDate
        .ToIndianDate = toDate
        
        .FromAmount = RepOption.FromAmount
        .ToAmount = RepOption.ToAmount
        .Caste = RepOption.Caste
        .Place = RepOption.Place
        .Gender = RepOption.Gender
        .AccountGroup = RepOption.AccountGroupID
    End With
    
    Load m_frmRDReport
    If Not gCancel Then
        Unload frmCancel
        m_frmRDReport.Show
    Else
        Unload frmCancel
    End If
    
    Set frmCancel = Nothing

End Sub

Private Sub m_frmRDAcc_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle
m_ParentHandle = 0
Set m_frmRDAcc = Nothing

End Sub


Private Sub m_frmRDReport_Initialising(Min As Long, Max As Long)
    If Max <> 0 And Max > Min Then
        With frmCancel
            .PicStatus.Visible = True
            UpdateStatus .PicStatus, 0
        End With
    End If
End Sub


Private Sub m_FrmRDReport_Processing(strMessages As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS :" & vbCrLf & strMessages
    UpdateStatus .PicStatus, Ratio
End With
End Sub


Private Sub m_frmRDReport_WindowClosed()
Set m_frmRDReport = Nothing

End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmRDAcc Is Nothing Then _
            ObjectHandle = m_frmRDAcc.hwnd: Exit Property
    If Not m_frmRDReport Is Nothing Then _
            Unload m_frmRDReport: Set m_frmRDReport = Nothing
    
End Property

