VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCAAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CAAcc As New ClsCAAcc
Private m_AccID As Long
Private m_rstCustBalance As ADODB.Recordset
Private m_ParentHandle As Long
Private WithEvents m_frmCAAcc As frmCAAcc
Attribute m_frmCAAcc.VB_VarHelpID = -1
Public Event UpdateStatus(strMsg As String)
Private WithEvents m_frmCAReport As frmCAReport
Attribute m_frmCAReport.VB_VarHelpID = -1


Public Property Get ObjectHandle() As Long
    If Not m_frmCAAcc Is Nothing Then _
             ObjectHandle = m_frmCAAcc.hwnd: Exit Property
    If Not m_frmCAReport Is Nothing Then _
                Set m_frmCAReport = Nothing
End Property

Public Property Get ObjectName()
    ObjectName = "CAClass"
End Property

Public Property Let AccountId(NewValue As Long)
    
    m_AccID = IIf(CAAccountExists(NewValue), NewValue, 0)
    
End Property


Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset
gDbTrans.SqlStmt = "Select AccID From CAMAster " & _
    " WHERE CustomerID = " & NewValue

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then m_AccID = rst("AccID")

Set rst = Nothing

End Property


Private Function PLTransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency
Dim StrSql As String
Dim rst As ADODB.Recordset

PLTransactionDetails = 0

gDbTrans.SqlStmt = "Select Sum(Amount) as Deposits from CAPLTrans where" & _
        " TransType = " & transType & _
        " And TransDate >= #" & fromDate & "#" & _
        " And TransDate <= #" & toDate & "#"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
PLTransactionDetails = Val(FormatField(rst(0)))
Set rst = Nothing

End Function

Private Function TransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency
Dim StrSql As String
Dim rst As ADODB.Recordset

TransactionDetails = 0

StrSql = "Select Sum(Amount) as Deposits from CATrans where " & _
         "TransType = " & transType & " And  TransDate Between #" & _
         fromDate & "# And #" & toDate & "#"

gDbTrans.SqlStmt = StrSql
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
TransactionDetails = Val(FormatField(rst(0)))
Set rst = Nothing

End Function

Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean


'Check if account exists
Dim ClosedON As String
If Not CAAccountExists(AccId, ClosedON) Then
    'MsgBox "Specified account does not exist !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
    Exit Function
End If
    
'Get last transaction record
Dim lastTransID As Long
Dim rst As ADODB.Recordset
    
    gDbTrans.SqlStmt = "Select TOP 1 * from CATrans where " & _
        "AccID = " & AccId & " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
             'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
        MsgBox GetResourceString(582), vbInformation, gAppName & " - Error"
        Exit Function
    End If
    lastTransID = FormatField(rst("TransID"))
    If lastTransID <> TransID Then
        MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
    
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim ChequeNo As Long
    
    Amount = FormatField(rst("Amount"))
    
    'If you had a cheque book here...
    transType = FormatField(rst("TransType"))
    ChequeNo = FormatField(rst("ChequeNo"))
    Set rst = Nothing

'Delete record from Data base
    If transType = wContraWithdraw Then
        If ChequeNo > 0 Then
            gDbTrans.SqlStmt = "UPDATE ChequeMaster Set Trans = " & wischqIssue & _
                " WHERE ChequeNo = " & ChequeNo & " AND AccID = " & AccId & _
                " AND AccHeadId = " & GetIndexHeadID(GetResourceString(422))
            If Not gDbTrans.SQLExecute Then Exit Function
            
        End If
    End If
     
    'Remove the Related records first
    gDbTrans.SqlStmt = "Delete from CAPLtrans where AccID = " & AccId & _
                      " and TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then Exit Function
    
    'next remove the records
    gDbTrans.SqlStmt = "Delete from CAtrans where AccID = " & AccId & _
                       " and TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then Exit Function
    
    gDbTrans.SqlStmt = "UPDATE CaMaster Set ClosedDate = NULL " & _
        " WHERE AccID = " & AccId
    If Not gDbTrans.SQLExecute Then Exit Function

UndoTransaction = True

End Function


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
'Cash Withdrwals
transType = wDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
CashDeposits = Amount

End Function


Public Function CashWithdrawls(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons

CashWithdrawls = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

'Cash Withdrwals
transType = wWithdraw
Amount = Amount + TransactionDetails(fromDate, toDate, transType)
CashWithdrawls = Amount

End Function

Public Function CustomerBalance(CustomerID As Long, retAccNum As String, AsOnDate As Date) As Currency
'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
'Sql stmt to get Accid

gDbTrans.SqlStmt = "Select A.Accid,AccNum,Balance from " & _
    " CAMaster A INNER Join CATrans B On A.AccID = B.AccID " & _
    " where A.Customerid = " & CustomerID & " AND TransID = " & _
        "(Select MAx(TransID) From CATrans C Where C.AccId = A.AccID " & _
        " AND TransDate <= #" & AsOnDate & "# )" & _
    " AND Balance > 0"


If gDbTrans.Fetch(m_rstCustBalance, adOpenForwardOnly) <= 0 Then Exit Function

NextAccount:
retAccNum = CInt(FormatField(m_rstCustBalance("AccNum")))
CustomerBalance = FormatField(m_rstCustBalance("Balance"))

If CustomerBalance = 0 And m_rstCustBalance.RecordCount > 1 And m_rstCustBalance.EOF = False Then
    m_rstCustBalance.MoveNext
    GoTo NextAccount
End If
ErrLine:
    If Err Then _
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical

End Function
'Public f_frmSBAcc As frmSBAcc
Public Sub Show()
Dim StartCA As Boolean

If gCurrUser Is Nothing Then
    Set gCurrUser = New clsUsers
Else
    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Dim Perms As wis_Permissions
    Perms = gCurrUser.UserPermissions
    If Perms And perBankAdmin Then StartCA = True
    If Perms And perCreateAccount Then StartCA = True
    If Perms And perClerk Then StartCA = True
    If Perms And perCashier Then StartCA = True
    If Perms And perReadOnly Then StartCA = True
    If StartCA Then
        If m_frmCAAcc Is Nothing Then Set m_frmCAAcc = New frmCAAcc
        Load m_frmCAAcc
        If gWindowHandle Then m_ParentHandle = gWindowHandle
        gWindowHandle = m_frmCAAcc.hwnd
        m_frmCAAcc.Show
        If m_AccID < 0 Then m_frmCAAcc.AccountLoad (m_AccID)
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
End If

End Sub
Public Function Deposits(fromDate As Date, toDate As Date) As Currency
'Deposit inflow to the bank between 2  given dates

Dim transType As wisTransactionTypes
Deposits = 0

transType = wDeposit
Deposits = TransactionDetails(fromDate, toDate, transType)

End Function

Public Function Loss(fromDate As Date, toDate As Date) As Currency
Dim Amount As Currency
Dim transType As wisTransactionTypes

Amount = 0

'first get the profit direct cash
transType = wWithdraw
Amount = PLTransactionDetails(fromDate, toDate, transType)

';Now calculate the contra transaction
transType = wContraWithdraw
Amount = Amount + PLTransactionDetails(fromDate, toDate, transType)


Loss = Amount
End Function

'THis Function Retuns the Next Account Blance of the customewr
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


'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function WithdrawAmount(AccId As Long, Amount As Currency, _
    Description As String, TransDate As Date, Optional VoucherNo As String = "", _
    Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

'Check Whether Account Exist or Closed
Dim ClosedDate As String
Dim rst As ADODB.Recordset

If Not CAAccountExists(AccId, ClosedDate) Then Exit Function

'If it is already closed then
If Trim$(ClosedDate) <> "" Then Exit Function

'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
gDbTrans.SqlStmt = "SELECT MAX(TransID) From " & _
            " CATrans WHERE Accid = " & AccId & _
            " And TransDate <= #" & TransDate & "# "

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
Dim TransID As Long

TransID = Val(FormatField(rst(0)))

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID + 1
End If
Set rst = Nothing

'Now Get the  Max TransID of the Account
gDbTrans.SqlStmt = "SELECT MAX(TransID) From CATrans WHERE Accid = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If TransID <> Val(FormatField(rst(0))) Then
        Set rst = Nothing
       'if There is a transaction After This Date
       'then this  TransCtion Canonot Be done
       Exit Function
    End If
End If

'First get The Latest Balance For This Transaction
gDbTrans.SqlStmt = "SELECT Transid, Balance FROM CATrans " & _
    " WHERE Accid = " & AccId & " And TransId = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

Dim Balance As Currency
Dim Particulars As String

'Assign The Necessary Values
TransID = Val(FormatField(rst("TransID"))) + 1
Balance = Val(FormatField(rst("Balance"))) - Amount

Set rst = Nothing

Particulars = Description

If Balance < 0 Then Exit Function

Dim transType As wisTransactionTypes
Dim UserID As Long

UserID = gCurrUser.UserID
transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Query To Put The Acquired Values into Corresponding Table

gDbTrans.SqlStmt = "INSERT INTO CATrans (Accid ,TransId ,Amount," _
         & "Balance,TransType,Particulars,TransDate,VoucherNo,UserID ) " _
         & "VALUES(" & AccId & "," & TransID & "," _
         & Amount & "," & Balance & "," & transType & "," _
         & AddQuotes(Particulars, True) & "," _
         & "#" & TransDate & "#, " _
         & AddQuotes(VoucherNo, True) & "," & UserID & ")"

If Not gDbTrans.SQLExecute Then GoTo Exit_Line


'update it to te Contra Trans table
'Get the Headid
Dim ClsBank As clsBankAcc
Dim AccHeadID As Long

AccHeadID = GetIndexHeadID(GetResourceString(422))

If Not CashTransaction Then
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

WithdrawAmount = TransID

Exit_Line:
End Function

'This Function will Deposit the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
                ByVal Description As String, ByVal TransDate As Date, _
                Optional VoucherNo As String = "", _
                Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

'Check Whether Account Exist or Closed
Dim ClosedDate As String
Dim rst As ADODB.Recordset

If Not CAAccountExists(AccId, ClosedDate) Then Exit Function

'If it is already closed then
If Trim$(ClosedDate) <> "" Then Exit Function

If Not CashTransaction Then
'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
End If

'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
gDbTrans.SqlStmt = "SELECT MAX(TransID) From CATrans" & _
        " WHERE Accid = " & AccId & _
        " And TransDate <= #" & TransDate & "#"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
Dim TransID As Long
TransID = Val(FormatField(rst(0)))

'Now Get the  Max TransID of the Account
gDbTrans.SqlStmt = "SELECT MAX(TransID) From CATrans WHERE Accid = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If TransID <> Val(FormatField(rst(0))) Then
       'if There is a transaction After This Date then this  TransCtion Canon Be done
       Exit Function
    End If
End If

'First get The Latest Balance For This Transaction
gDbTrans.SqlStmt = " SELECT Transid, Balance FROM CATrans WHERE Accid = " & AccId & _
   " And TransId = " & TransID
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

Dim Balance As Currency
Dim Particulars As String

'Assign The Necessary Values
TransID = Val(FormatField(rst("TransID"))) + 1
Balance = CCur(Val(FormatField(rst("Balance"))) + Amount)

Set rst = Nothing

Particulars = Description

Dim transType As wisTransactionTypes
Dim UserID As Long

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SqlStmt = "INSERT INTO CATrans (Accid ,TransId ,Amount ," _
    & "Balance,TransType, Particulars,TransDate,VoucherNo,UserId) " _
    & " VALUES(" & AccId & "," & TransID & "," _
    & Amount & "," & Balance & "," & transType & "," _
    & AddQuotes(Particulars, True) & "," _
    & "#" & TransDate & "#," _
    & AddQuotes(VoucherNo, True) & "," & UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(GetResourceString(422))  'Current Account

If Not CashTransaction Then
    'update it to te Contra Trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans" & _
        " (ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,VoucherNo,UserID) " & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

End If

DepositAmount = TransID

Exit_Line:

End Function


Public Function WithDrawls(fromDate As Date, toDate As Date) As Currency
'Deposits of customers given back to the respective clients due a/c closure,
'matured or due to any other reasons

WithDrawls = 0
Dim transType As wisTransactionTypes

transType = wWithdraw


WithDrawls = TransactionDetails(fromDate, toDate, transType)

End Function

Public Function Profit(fromDate As Date, toDate As Date) As Currency
Dim Amount As Currency
Dim transType As wisTransactionTypes

Amount = 0

'first get the profit direct cash
transType = wWithdraw
Amount = PLTransactionDetails(fromDate, toDate, transType)

';Now calculate the contra transaction
transType = wContraWithdraw
Amount = Amount + PLTransactionDetails(fromDate, toDate, transType)

Profit = Amount
End Function

Public Function GetAccountID(AccNum As String) As Integer

GetAccountID = 0

Dim ret As Integer
Dim rst As ADODB.Recordset

'Query Database
    gDbTrans.SqlStmt = "Select * from CAMaster where AccNum = " & AddQuotes(AccNum, True)
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
    
GetAccountID = FormatField(rst("AccID"))
Set rst = Nothing
End Function


Public Function Balance(AsOnDate As Date) As Currency
'Flow of deposit into the bank till date
Balance = 0
Balance = ComputeTotalCALiability(AsOnDate)

End Function

Private Sub m_frmCAAcc_AccountChanged(ByVal AccId As Long)
m_AccID = AccId
End Sub

Public Sub TotalDepositTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from CATrans where " & _
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
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from CAPLTrans where " & _
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



Private Sub m_frmCAAcc_ShowReport(ReportType As wis_CAReports, ReportOrder As wis_ReportOrder, _
                fromDate As String, toDate As String, RepOption As clsRepOption)

If Not m_frmCAReport Is Nothing Then Set m_frmCAReport = Nothing
Set m_frmCAReport = New frmCAReport

With m_frmCAReport
    .FromIndianDate = fromDate
    .ToIndianDate = toDate
    .FromAmount = RepOption.FromAmount
    .ToAmount = RepOption.ToAmount
    .Place = RepOption.Place
    .Caste = RepOption.Caste
    .Gender = RepOption.Gender
    .ReportOrder = ReportOrder
    .AccountGroup = RepOption.AccountGroupID
    .ReportType = ReportType
End With

gCancel = 0
Load m_frmCAReport
If gCancel Then Exit Sub
Unload frmCancel
m_frmCAReport.Show 'vbModal

'Unload m_frmCAReport

End Sub

Private Sub m_frmCAAcc_WindowClosed()

If m_ParentHandle Then gWindowHandle = m_ParentHandle
Set m_frmCAAcc = Nothing

End Sub


Private Sub m_frmCAReport_Initialize(Min As Integer, Max As Integer)
    If Max <> 0 Then
        With frmCancel
            '.prg.Visible = True
            '.prg.Min = Min
            If Max > 32500 Then Max = 32500
            '.prg.Max = Max
        End With
    End If

End Sub


Private Sub m_frmCAReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
gCancel = 0
frmCancel.Show
With frmCancel
    .lblMessage = "PROCESS :" & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
    If Ratio > 0 Then
        'If Ratio > 1 Then Ratio = Ratio / .prg.Max
        '.prg.Value = .prg.Max * Ratio
    End If
End With

End Sub


