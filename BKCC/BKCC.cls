VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBkcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ConsiderDeposit As Boolean
Private WithEvents LoanDlg As frmBKCCAcc
Attribute LoanDlg.VB_VarHelpID = -1
Private m_rstCustBalance As Recordset
Private WithEvents m_frmBKCCReport As frmBkccReport
Attribute m_frmBKCCReport.VB_VarHelpID = -1
Private m_LoanID As Long
Private m_ParentHandle As Long

Public Property Let AccountId(NewValue As Long)
    m_LoanID = NewValue
End Property

Public Function CustomerBalance(ByVal customerID As Long, _
        retAccNum As String, ByVal AsOnDate As Date) As Currency

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Balance " & _
    " FROM BKCCMAster A,BKCCTrans B " & _
    " WHERE A.CustomerID = " & customerID & _
    " AND A.LOanId = B.LoanID And TransID = " & _
        "(SELECT Max(TransID) From BKCCTrans C Where " & _
        " C.LoanID = A.LoanID AND TransDate <= #" & AsOnDate & "# )"

retAccNum = ""
CustomerBalance = 0
If gDbTrans.Fetch(m_rstCustBalance, adOpenDynamic) > 0 Then
    CustomerBalance = FormatField(m_rstCustBalance("Balance"))
    retAccNum = FormatField(m_rstCustBalance("AccNum"))
End If

End Function

Public Property Let customerID(NewValue As Long)

Dim rst As Recordset

gDbTrans.SqlStmt = "Select LoanID From BKCCMAster " & _
        " WHERE CustomerID = " & NewValue
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then m_LoanID = rst("LoanId")

Set rst = Nothing

End Property

Public Function DepositAmount(ByVal LoanID As Long, ByVal Amount As Currency, _
                ByVal IntAmount As Currency, ByVal PenalAmount As Currency, _
                ByVal strDesc As String, ByVal TransDate As Date, _
                Optional VoucherNo As String = " ", _
                Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

On Error GoTo Exit_Line

Dim TransID As Long
Dim transType As wisTransactionTypes
Dim LastTransDate As Date
Dim Balance As Currency
Dim Deposit As Boolean
Dim IntBalance As Currency
Dim PenalIntBalance As Currency
Dim rst As Recordset

Dim DepHeadID As Long
Dim LoanHeadID As Long
Dim RegIntHeadID As Long
Dim PenalHeadID As Long
Dim DepIntHeadId As Long

Dim AccHeadName As String

'Now Get the Head ID's From The BankHeadIS
Dim ClsBank As clsBankAcc
Set ClsBank = New clsBankAcc
'Loand Head
AccHeadName = GetResourceString(229, 58)
LoanHeadID = ClsBank.GetHeadIDCreated(AccHeadName, LoadResString(229) & " " & LoadResString(58), parMemberLoan, 0, wis_BKCCLoan)

'Loan Reg Interest
AccHeadName = GetResourceString(229, 58, 344)
RegIntHeadID = ClsBank.GetHeadIDCreated(AccHeadName, _
    LoadResString(229) & " " & LoadResString(58) & " " & LoadResString(344), parMemLoanIntReceived, 0, wis_BKCCLoan)
'Penal int Interest
AccHeadName = GetResourceString(229, 58, 345)
PenalHeadID = ClsBank.GetHeadIDCreated(AccHeadName, _
    LoadResString(229) & " " & LoadResString(58) & " " & LoadResString(345), parMemLoanPenalInt, 0, wis_BKCCLoan)

'Deposit Head
DepHeadID = LoanHeadID
DepIntHeadId = RegIntHeadID
If m_ConsiderDeposit Then
    AccHeadName = GetResourceString(229, 43)
    DepHeadID = ClsBank.GetHeadIDCreated(AccHeadName, LoadResString(229) & " " & LoadResString(43), parMemberDeposit, 0, wis_BKCC)
    'Deposit Interest Head
    AccHeadName = GetResourceString(229, 43, 487)
    DepIntHeadId = ClsBank.GetHeadIDCreated(AccHeadName, _
        LoadResString(229) & " " & LoadResString(43) & " " & LoadResString(487), parMemDepIntPaid, 0, wis_BKCC)
End If

Set ClsBank = Nothing

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
End If

'Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Balance,TransID,TransDate FROM " & _
    " BKCCTrans WHERE loanid = " & LoanID & " And TransID = " & _
        " (SELECT MAX(TransID) FROM BKCCTrans WHERE loanid = " & LoanID & ")"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    
Balance = Val(FormatField(rst("Balance")))
TransID = Val(FormatField(rst("TransID")))
LastTransDate = rst("TransDate")

Deposit = IIf(Balance < 0, True, False)
'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, LastTransDate) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Get the Transaction Id From The Interest table
gDbTrans.SqlStmt = "SELECT top 1 * FROM BKCCIntTrans " & _
    " WHERE loanid = " & LoanID & " " & _
        " ORDER By TransID Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    LastTransDate = rst("TransDate")
    'Now Compare the Last Date Of Transaction with transction date
    If DateDiff("d", TransDate, LastTransDate) > 0 Then
        'Date Trasnaction should be later
        MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    If rst("TransID") > TransID Then TransID = rst("TransID")
    IntBalance = FormatField(rst("IntBalance"))
    PenalIntBalance = FormatField(rst("PenalIntBalance"))
End If

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

'Update the Regular & Penal interest amount.
TransID = TransID + 1

Dim NewBalance As Currency
Dim PrincAmount As Currency
Dim UserID As Long

UserID = gCurrUser.UserID
PrincAmount = Amount
NewBalance = Balance - Amount

''''''''''''''''''''''
Dim Cal_RegInt As Currency
Dim Cal_PenalInt As Currency
Dim IntBalance_pr As Currency

Cal_RegInt = BKCCRegularInterest(TransDate, LoanID) \ 1
Cal_PenalInt = BKCCPenalInterest(TransDate, LoanID) \ 1
IntBalance_pr = Cal_RegInt + Cal_PenalInt
'now get the Interest Balance difference

Deposit = IIf(Balance < 0, True, False)

    'If he is repaying the amount then check whether the interest
    'he is paying is equal to the actual interest accrued
    'Or he is collecting the more/less interest
    'If he collecting more/less interest then keep it as balance
    
If IntAmount > 0 Or PenalAmount > 0 Then
    Dim RetMsg As VbMsgBoxResult
    If (IntAmount > (Cal_RegInt + IntBalance)) Or _
        (PenalAmount > (Cal_PenalInt + PenalIntBalance)) Then
       'If MsgBox("The amount he is paying is more than his Interest till date " & _
            vbCrLf & " Do you want to keep this Extra amount as advance interest", vbYesNo + vbQuestion + _
            vbDefaultButton1, wis_MESSAGE_TITLE) = vbNo Then
        RetMsg = MsgBox(GetResourceString(666) & _
            vbCrLf & GetResourceString(667), vbYesNo + vbQuestion + _
            vbDefaultButton2, wis_MESSAGE_TITLE)
        'IntBalance = 0: PenalIntBalance = 0
        If RetMsg = vbCancel Then GoTo Exit_Line
        If RetMsg = vbYes Then
            IntBalance = IntBalance + Cal_RegInt - IntAmount
            PenalIntBalance = PenalIntBalance + Cal_PenalInt - PenalAmount
        End If
    ElseIf (IntAmount < Cal_RegInt Or PenalAmount < Cal_PenalInt) Then
       'If MsgBox("The amount he is paying is less than his Interest till date " & _
            vbCrLf & " Do you want to keep this difference amount as interest balance ", vbYesNo + vbQuestion + _
            vbDefaultButton1, wis_MESSAGE_TITLE) = vbNo Then
        RetMsg = MsgBox(GetResourceString(668) & _
            vbCrLf & GetResourceString(669), vbYesNoCancel + vbQuestion + _
            vbDefaultButton2, wis_MESSAGE_TITLE)
        If RetMsg = vbCancel Then GoTo Exit_Line
        If RetMsg = vbNo Then
            IntBalance = IntBalance + Cal_RegInt - IntAmount
            PenalIntBalance = PenalIntBalance + Cal_PenalInt - PenalAmount
        End If
    End If
End If

'''''''''''''''''''''

If (Balance > 0 And NewBalance < 0) Or (Balance < 0 And NewBalance > 0) Then
    'suppose we have to do two transaction
    'Then do the first transaction
    PrincAmount = Abs(Balance)
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID, " _
        & "TransType, Amount, TransDate, " _
        & "Balance,Deposit,Particulars,UserID) " _
        & "VALUES (" & LoanID & ", " & TransID & ", " _
        & transType & ", " & PrincAmount & "," _
        & "#" & TransDate & "#, " _
        & " 0 , " & Deposit & ", " _
        & AddQuotes(strDesc, True) _
        & UserID & ")"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    PrincAmount = Abs(NewBalance)
    Deposit = Not Deposit
    
    If Not CashTransaction Then
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TrasnType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & _
            IIf(Balance > 0, LoanHeadID, DepHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & PrincAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    'TransID = TransID + 1
End If

If NewBalance < 0 Then Deposit = True
If NewBalance > 0 Then Deposit = False
    
'The Deposit Interest amount alrady added to the the deposit
'So if the difference is more than the deposit interest
'then only do this transction

If IntAmount > 0 Or PenalAmount > 0 Then
    'If Deposit Then TransType = wContraWithdraw
    'Deposit = False
    gDbTrans.SqlStmt = "INSERT INTO BKCCIntTrans (LoanID, TransID," _
            & " TransDate, TransType, IntAmount, PenalIntAmount,  " _
            & "MiscAmount, IntBalance,Deposit,UserId) " _
            & "VALUES (" & LoanID & ", " & TransID & ", " _
            & " #" & TransDate & "#," & transType & ", " _
            & IntAmount & ", " & PenalAmount & "," _
            & "0 , " & IntBalance & "," & Deposit & "," _
            & UserID & ")"
            
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    gDbTrans.SqlStmt = "UPDATE BKCCmaster SET [LastIntDate] = " & _
                "#" & TransDate & "# WHERE loanid = " & LoanID
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,userID)" & _
            " Values (" & ContraID & "," & _
            IIf(Deposit, DepIntHeadId, RegIntHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & IntAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If RegIntHeadID > 0 Then _
            If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID ," & _
            "TrasnType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & _
            IIf(Deposit, DepIntHeadId, PenalHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & PenalAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If PenalAmount > 0 Then _
            If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
End If

If PrincAmount > 0 Then
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID, " _
        & "TransType, Amount, TransDate," _
        & "Balance, Deposit, Particulars,VoucherNo,UserID) " _
        & "VALUES (" & LoanID & ", " & TransID & ", " _
        & transType & ", " & PrincAmount & ", " _
        & "#" & TransDate & "#, " _
        & NewBalance & "," & Deposit & "," _
        & AddQuotes(strDesc, True) & "," _
        & AddQuotes(VoucherNo, True) & "," & UserID & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & _
            IIf(Deposit, DepHeadID, LoanHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & PrincAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
End If
    
' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.
If NewBalance = 0 And PrincAmount > 0 Then
    gDbTrans.SqlStmt = "UPDATE Bkccmaster SET loanclosed = 1, " & _
        " InterestBalance = 0 Where LoanID =  " & LoanID
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If
   
DepositAmount = TransID

Exit_Line:

If Err Then
    MsgBox "Error In AmountDeposit - BKCC", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

End Function

Public Function UndoTransaction(ByVal LoanID As Long, ByVal TransID As Long) As Boolean

' Variables of the procedure...
Dim lastTransID As Long
Dim inTransaction As Boolean
Dim transType As wisTransactionTypes
Dim Lret As Long
Dim rst As Recordset

Dim RegInt As Currency
Dim PenalInt As Currency
Dim LoanAmount As Currency
Dim DepAmount As Currency
Dim MiscAmount As Currency
Dim DepInt As Currency
Dim boolDeposit As Boolean

Dim TransDate As Date

' Setup the error handler.
On Error GoTo Err_line

'Get the last transaction from the loan transaction table.
gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans Where Loanid = " & LoanID & _
    " And TransID = (SELECT Max(Transid) FROM BKCCTrans " & _
        " WHERE Loanid = " & LoanID & " )"
    
Lret = gDbTrans.Fetch(rst, adOpenDynamic)

If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "There are no transaction to undo.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(661), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

lastTransID = FormatField(rst("TransID"))
TransDate = rst("TransDate")
transType = FormatField(rst("TransType"))
boolDeposit = FormatField(rst("Deposit"))
If boolDeposit Then
    DepAmount = FormatField(rst("Amount"))
Else
    LoanAmount = FormatField(rst("Amount"))
End If

If rst.RecordCount > 1 Then
    rst.MoveNext
    boolDeposit = FormatField(rst("Deposit"))
    If boolDeposit Then
        DepAmount = FormatField(rst("Amount"))
    Else
        LoanAmount = FormatField(rst("Amount"))
    End If
End If

gDbTrans.SqlStmt = "SELECT *  FROM BKCCIntTrans " & _
        " WHERE loanid = " & LoanID & _
        " AND TransID = (SELECT Max(TransID) FROM BKCCIntTrans " & _
            " WHERE loanid = " & LoanID & ")"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    If lastTransID < FormatField(rst("TransID")) Then DepAmount = 0: LoanAmount = 0
    If lastTransID <= FormatField(rst("TransID")) Then
        lastTransID = FormatField(rst("TransID"))
        TransDate = rst("TransDate")
        boolDeposit = FormatField(rst("Deposit"))
        If boolDeposit Then
            DepInt = FormatField(rst("IntAmount"))
        Else
            RegInt = FormatField(rst("IntAmount"))
            PenalInt = FormatField(rst("PenalIntAmount"))
        End If
        MiscAmount = FormatField(rst("MiscAmount"))
    End If
End If

If lastTransID <> TransID Then
    'MsgBox "This account has transaction", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


' Delete the entry of principle amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM BKCCTrans WHERE transid = " & lastTransID _
            & " AND loanid = " & LoanID

If Not gDbTrans.SQLExecute Then
    MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Delete the entry of Interest amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM BKCCIntTrans WHERE Transid = " & lastTransID _
        & " AND loanid = " & LoanID
If Not gDbTrans.SQLExecute Then
    MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'UpDate The LoanMAster
gDbTrans.SqlStmt = "UPDATE BkCCMaster SET LoanClosed = 0 ,ClosedDate = NULL " & _
            " Where Loanid =  " & LoanID
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
If DepInt Or RegInt Or PenalInt Then
    gDbTrans.SqlStmt = "Select Max(TransDate) From BKCCIntTrans Where Loanid =  " & LoanID
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    If Not IsNull(rst(0)) Then
    gDbTrans.SqlStmt = "UPDATE BkccMaster SET LastIntDate = " & _
        "(Select Max(TransDate) From BKCCIntTrans Where Loanid =  " & LoanID & ") " & _
        " Where Loanid =  " & LoanID
    Else
        gDbTrans.SqlStmt = "UPDATE BkccMaster SET LastIntDate = NULL"
    End If
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
End If

UndoTransaction = True

Exit_Line:
    Exit Function

Err_line:
    
    If Err Then
        MsgBox "BKCCClassUndoTransaction: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(731) & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    UndoTransaction = False
    GoTo Exit_Line
    
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


Public Function WithdrawAmount(LoanID As Long, Amount As Currency, _
        strDesc As String, TransDate As Date, _
        Optional VoucherNo As String, _
        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

On Error GoTo Exit_Line

Dim TransID As Long
Dim transType As wisTransactionTypes
Dim LastTransDate As Date
Dim Balance As Currency
Dim Deposit As Boolean
Dim rst As Recordset

' Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Balance,TransID,TransDate FROM " & _
    " BKCCTrans WHERE loanid = " & LoanID & " ORDER BY TransId Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then GoTo ExitLine

    Balance = Val(FormatField(rst("Balance")))
    TransID = Val(FormatField(rst("TransID")))
    LastTransDate = rst("TransDate")

Deposit = IIf(Balance < 0, True, False)

'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, LastTransDate) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Get the Transaction Id From The Interest table
gDbTrans.SqlStmt = "SELECT TransID,TransDate FROM BKCCIntTrans " & _
    " WHERE loanid = " & LoanID & " And TransID =  (SELECT MAX(TransID) " & _
        " FROM BKCCIntTrans WHERE loanid = " & LoanID & ")"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    LastTransDate = rst("TransDate")
    'Now Compare the Last Date Of Transaction with transction date
    If DateDiff("d", TransDate, LastTransDate) > 0 Then
        'Date Trasnaction should be later
        MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    If rst("TransID") > TransID Then TransID = rst("TransID")
End If

transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Update the Regular & Penal interest amount.
TransID = TransID + 1

Dim NewBalance As Currency
Dim PrincAmount As Currency
Dim DepHeadID As Long
Dim LoanHeadID As Long

Dim AccHeadName As String

'Now Get the Head ID's From The BankHeadIS
Dim ClsBank As clsBankAcc
Set ClsBank = New clsBankAcc
'Loand Head
AccHeadName = GetResourceString(229, 58)
LoanHeadID = ClsBank.GetHeadIDCreated(AccHeadName)
'Deposit Head
AccHeadName = GetResourceString(229, 43)
DepHeadID = ClsBank.GetHeadIDCreated(AccHeadName)

Set ClsBank = Nothing

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID + 1
End If

PrincAmount = Amount
NewBalance = Balance + Amount

If (Balance > 0 And NewBalance < 0) Or (Balance < 0 And NewBalance > 0) Then
    'suppose we have to do two transaction
    'Then do the first traNASCTION HERE
    PrincAmount = Abs(Balance)
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID, " _
        & "TransType, Amount, TransDate, Balance,Deposit,Particulars,UserID) " _
        & "VALUES (" & LoanID & ", " & TransID & ", " _
        & transType & ", " & PrincAmount & "," _
        & "#" & TransDate & "#, " _
        & " 0 , " & Deposit & ", " _
        & AddQuotes(strDesc, True) _
        & "," & gUserID & ")"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'Update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccId," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & _
            IIf(Deposit, DepHeadID, LoanHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
    TransID = TransID + 1
    PrincAmount = Abs(NewBalance)
    Deposit = Not Deposit
End If

If NewBalance < 0 Then Deposit = True
If NewBalance > 0 Then Deposit = False
    
'The Deposit Interest amount alrady added to the the deposit
'So if the difference is more than the deposit interest
'then only do this transction

If PrincAmount > 0 Then
    gDbTrans.SqlStmt = "INSERT INTO BKCCTrans (LoanID, TransID, " _
        & "TransType, Amount, TransDate, Balance, Deposit, Particulars,UserID) " _
        & "VALUES (" & LoanID & ", " & TransID & ", " _
        & transType & ", " & PrincAmount & ", " _
        & "#" & TransDate & "#, " _
        & NewBalance & ", " & Deposit & "," _
        & AddQuotes(strDesc, True) _
        & "," & gUserID & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'Update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & _
            IIf(Deposit, DepHeadID, LoanHeadID) & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gCurrUser.UserID & ")"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
End If
    
' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.
If NewBalance = 0 Then
    gDbTrans.SqlStmt = "UPDATE Bkccmaster SET loanclosed = 1, " & _
            " InterestBalance = 0 Where LoanID =  " & LoanID
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
End If
   
WithdrawAmount = TransID

Exit_Line:

If Err Then
    MsgBox "Error In AmountDeposit - BKCC", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

ExitLine:

End Function

'Shashi Anagadi
'Created
'Modified On 27/7/02
'This Function return the Amount deposit
'in the BKCC
Public Function BalanceDeposit(AsonIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

Dim AsOnDate As Date
AsOnDate = GetSysFormatDate(AsonIndianDate)

' Setup error handler.
Err.Clear
On Error GoTo Exit_Line

' Declare variables...
Dim Lret As Long
Dim rst As Recordset
Dim Balance As Currency


'If any account has Deposit Balance
'Then Blance has to be less than zero for deposit

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM BKCCTrans A WHERE " & _
    " A.TransId = (Select max(TransID) from BKCCTrans B" & _
    " where B.LoanId = A.LoanID And TransDate <= #" & AsOnDate & "#)" & _
    " And Balance < 0"
Call gDbTrans.Fetch(rst, adOpenDynamic)
BalanceDeposit = Abs(FormatField(rst(0)))
Set rst = Nothing

Exit_Line:
    If Err Then
        MsgBox "BKCCLoanBalance: " & vbCrLf _
            & Err.Description, vbCritical
        Err.Clear
    End If
End Function

'Shashi Angadi
'This function returns Loan BAlance of the Bkcc Accounts
'Modified ON 9TH AUGUST 2002
Public Function BalanceLoan(AsonIndianDate As String) As Currency

Dim AsOnDate As Date
AsOnDate = GetSysFormatDate(AsonIndianDate)

' Setup error handler.
On Error GoTo Err_line

'This Accounts may also have deposits
'and balance for the deposit will be negative
'so consider only positive balance

gDbTrans.SqlStmt = "SELECT SUM(Balance) AS TotalBalance " & _
                " FROM BkccTrans A,BkccMaster B WHERE A.LoanID= B.LoanID " & _
                " And A.TransId = (Select Max(TransID) from BKCCTrans C " & _
                        " where C.LoanId = A.LoanID And TransDate <= #" & AsOnDate & "#)"

If m_ConsiderDeposit Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND Balance > 0 "


Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line

BalanceLoan = Val(FormatField(rst("TotalBalance")))

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "LoanBalance: " & vbCrLf _
            & Err.Description, vbCritical
        Err.Clear
    End If

    BalanceLoan = 0
    GoTo Exit_Line

End Function

'Shashi Angadi
'this function returns the amount deposited as deposit
'in the bkcc accounts
'Modified 9th August 2002
Public Function CashDepositDeposits(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    On Error GoTo Err_line
       
If Not m_ConsiderDeposit Then Exit Function

    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    Amount = 0
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    'GET THE CASH dEPOSITS
    transType = wDeposit
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    
    CashDepositDeposits = Amount
    
Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Loss occured BKCC Deposit
'during the specified period
'Argument FromDate And to date in Indian Format
'Returns amount
Public Function CashDepositLoss(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Dim Amount As Currency
Dim transType As wisTransactionTypes

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
    CashDepositLoss = Amount

End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Deposit
'throught cash during the specified period
Public Function CashDepositProfit(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Dim Amount As Currency
Dim transType As wisTransactionTypes

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, True)
    
    CashDepositProfit = Amount
End Function

'Shashi Angadi
'this function returns the amount withdraw as loan from
'the bkcc accounts
'Modified 9th Augus 2002
Public Function CashDepositWithDrawls(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    
    CashDepositWithDrawls = Amount

Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function

'Shashi Angadi
'this function returns the amount withdraw as loan
'from the bkcc accounts
'Modified 9th AugusT 2002
Public Function CashLoanDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
    
    CashLoanDeposits = Amount
    
End Function


'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Loan
'during the specified period By cash Only
'Argument FromDate And to date in Indian Format
'Returns amount Profit Throuth Cash
Public Function CashLoanProfit(FromIndianDate As String, ToIndianDate As String) As Currency
On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
Err_line:
    
    CashLoanProfit = Amount
End Function

'Shashi Angadi
'this function returns the amount withdraw  as loan
'from the bkcc accounts
'Modified 9th AugusT 2002
Public Function CashLoanWithDrawls(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
        
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
    
    CashLoanWithDrawls = Amount
    
End Function

'Shashi Angadi
'this function returns the amount withdraw as loan from
'the bkcc accounts
'Modified 9th Augus 2002
Public Function ContraDepositDeposits(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    Amount = 0
    
    
    
    'GET THE contra dEPOSITS
    transType = wContraDeposit
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    
    ContraDepositDeposits = Amount
    
Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Loss occured BY BKCC Deposit
'during the specified period where cash is not handled
'Argument FromDate And to date in Indian Format
'Returns amount
Public Function ContraDepositLoss(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, True)
    
Err_line:

    ContraDepositLoss = Amount
    
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Deposit
'during the specified period
Public Function ContraDepositProfit(FromIndianDate As String, ToIndianDate As String) As Currency
    
If Not m_ConsiderDeposit Then Exit Function

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Dim Amount As Currency
Dim transType As wisTransactionTypes


Amount = 0
    'Now take the amount Where cash IS NOTH HANDLED
    transType = wContraDeposit
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, True)
    
    
    ContraDepositProfit = Amount
    
End Function

'Shashi Angadi
'this function returns the amount withdraw  as loan
'and transferred from the bkcc accounts
'Modified 9th AugusT 2002
Public Function ContraLoanDeposits(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wContraWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
        
    ContraLoanDeposits = Amount
    
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Loass coccurder throgh BKCC loan
'during the specified period
'Argument FromDate And to date in Indian Format
'Returns amount (Profit Throuth Loan
Public Function ContraLoanLoss(FromIndianDate As String, ToIndianDate As String) As Currency

On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
Err_line:
    
    ContraLoanLoss = Amount
End Function


'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Deposit
'during the specified period and cash is not handled
'Argument FromDate And to date in Indian Format
'Returns amount
Public Function ContraLoanProfit(FromIndianDate As String, ToIndianDate As String) As Currency
On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraDeposit
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
Err_line:
    
    ContraLoanProfit = Amount
End Function


'Shashi Angadi
'this function returns the amount withdraw  as loan
'and transferred from the bkcc accounts
'Modified 9th AugusT 2002
Public Function ContraLoanWithDrawls(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wContraWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
    
    
    ContraLoanWithDrawls = Amount
    
End Function

'Shashi Angadi
'this function returns the amount with draw as loan form
'the bkcc accounts
'Modified 9th August 2002

Public Function DepositDeposits(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    
If Not m_ConsiderDeposit Then Exit Function

    On Error GoTo Err_line
    
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    Amount = 0
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    'GET THE CASH dEPOSITS
    transType = wDeposit
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    
    'GET THE contra dEPOSITS
    transType = wContraDeposit
    Amount = Amount + GetTransactedAmount(fromDate, toDate, transType, True)
    
    DepositDeposits = Amount
    
Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function



'Shashi Angadi
'9th August 2002
'This Function Returns the Loss occured BKCC Deposit
'during the specified period
'Argument FromDate And to date in Indian Format
'Returns amount
Public Function DepositLoss(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraWithdraw
    Amount = Amount + GetIntTransactedAmount(fromDate, toDate, transType, True)
    
Err_line:
    
    DepositLoss = Amount
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Deposit
'during the specified period
Public Function DepositProfit(FromIndianDate As String, ToIndianDate As String) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, True)
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraDeposit
    Amount = Amount + GetIntTransactedAmount(fromDate, toDate, transType, True)
    
Err_line:
    
    DepositProfit = Amount
End Function


'Shashi Angadi
'tthis function returns the amount withdraw as loan from
'the bkcc accounts
'Modified 9th Augus 2002
Public Function DepositWithDrawls(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    transType = wContraWithdraw
    Amount = Amount + GetTransactedAmount(fromDate, toDate, transType, True)
    
    DepositWithDrawls = Amount

Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function


'Shashi Angadi
'9/8/2002
'This function REturns the profit/loss amount in the BKCC
'Arguments FromDate,ToDate to identify the transaction period
'   TransaType to identiify the Withdraw/Deposit
'   Deposit which optional identified thewhther deposit or loan transaction
'   default is loan
'Returns Proit/loss of deposit/loan as currency
Private Function GetIntTransactedAmount(fromDate As Date, toDate As Date, _
        transType As wisTransactionTypes, Optional Deposit As Boolean = False) As Currency

Dim SqlStr As String
    SqlStr = "SELECT SUM(amount) as TotalAmount From BKCCIntTrans" & _
        " WHERE FromDate >= #" & fromDate & "#" & _
        " AND TransDate <= #" & toDate & "# " & _
        " AND TransType = " & transType
        
If m_ConsiderDeposit Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND Deposit = " & Deposit
        
    gDbTrans.SqlStmt = SqlStr
Dim rst As Recordset
    GetIntTransactedAmount = 0
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        GetIntTransactedAmount = FormatField(rst("TotalAmount"))
        
End Function

Public Sub TotalDepositTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from BKCCTrans where Deposit = " & True & _
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
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from BKCCIntTrans where Deposit = " & True & _
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

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from BKCCTrans where Deposit = " & False & _
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
Public Sub TotalLoanIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from BKCCIntTrans where Deposit = " & False & _
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

'Shashi Angadi
'9/8/2002
'this function calculates the transcted amount in the BKCC
'The Amount may be withdraw,Deposit, ContraWithdraw,Contra Deposit
'It also saparate the whether transacted amount is of deposit or loan
'Arguments FromDate,ToDate to identify the transaction period
'   TransaType to identiify the Withdraw/Deposit
'   Deposit which optional identified thewhther deposit or loan transaction
'   default is loan
Private Function GetTransactedAmount(fromDate As Date, toDate As Date, transType As wisTransactionTypes, Optional Deposit As Boolean = False) As Currency

Dim SqlStr As String
    SqlStr = "SELECT SUM(amount) as TotalAmount From BKCCTrans" & _
        " WHERE FromDate >= #" & fromDate & "#" & _
        " AND TransDate <= #" & toDate & "# " & _
        " AND TransType = " & transType
        
If m_ConsiderDeposit Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND Deposit = " & Deposit
        
    gDbTrans.SqlStmt = SqlStr
    GetTransactedAmount = 0
    Dim rst As Recordset
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
            GetTransactedAmount = FormatField(rst("TotalAmount"))
        
End Function

'this function returns the amount of loan Suppose to recovered
'Between the speicified date
'Irrespeictive of the amount Recovered
'If he has paid the the amount then also that amount lies in a period
Public Function LoanDemand(LoanID As Long, fromDate As Date, toDate As Date) As Currency

LoanDemand = 0

Dim SqlStr As String
Dim rst As Recordset
Dim DemAmount As Currency
'Dim Balance As Currency

'Dim Duration As Long
'Dim InstType As wisInstallmentTypes

'Dim DueDate As Date

Err.Clear
On Error GoTo ExitLine
DemAmount = 0
    

'Get The Amount Withdrawn during the same period in last year
SqlStr = "SELECT Sum(Amount) as Demand FROM BKCCTrans Where " & _
    " LoanId = " & LoanID & " AND Deposit = False " & _
    " AND TransDate >= #" & DateAdd("yyyy", -1, fromDate) & "#" & _
    " AND TransDate <= #" & DateAdd("yyyy", -1, toDate) & "# " & _
    " And (TransType = " & wWithdraw & " OR TransType = " & wContraWithdraw & ")"
    
SqlStr = "SELECT Max(Balance) as Demand FROM BKCCTrans Where " & _
    " LoanId = " & LoanID & " AND Deposit = False " & _
    " AND TransDate <= #" & toDate & "# " & _
    " AND TransDate >= #" & fromDate & "# " '"Order By TransID Desc"

'SqlStr = "SELECT top 1 Balance as Demand FROM BKCCTrans Where " & _
    " LoanId = " & LoanID & " AND Deposit = False " & _
    " AND TransDate <= #" & ToDate & "# " & _
    " Order By TransID Desc"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    DemAmount = FormatField(rst("Demand"))
    If DemAmount = 0 Then
        SqlStr = "SELECT top 1 Balance as Demand FROM BKCCTrans Where " & _
            " LoanId = " & LoanID & " AND Deposit = False " & _
            " AND TransDate <= #" & fromDate & "# " & _
            " Order By TransID Desc"
        
        gDbTrans.SqlStmt = SqlStr
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                        DemAmount = FormatField(rst("Demand"))
    End If
End If

ExitLine:
    If Err Then
        MsgBox "ERROR In Loan demand" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
        Exit Function
    End If
    LoanDemand = DemAmount

End Function


'Shashi Angadi
'this function returns the amount with draw as loan form
'the bkcc accounts
'Modified 9th August 2002

Public Function LoanDeposits(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    On Error GoTo Err_line
    
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    
    Amount = 0
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    'GET THE CASH dEPOSITS
    transType = wDeposit
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
    
    'GET THE contra dEPOSITS
    transType = wContraDeposit
    Amount = Amount + GetTransactedAmount(fromDate, toDate, transType, False)
    
    LoanDeposits = Amount
    
Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function



'Shashi Angadi
'20 August 2002
'THis funcion returns the amount collected between
'specified dates as regular interest for loans
'Arguments date in "dd/mm/yyyy"
Public Function LoanInterest_Penal(FromIndianDate As String, ToIndianDate As String) As Currency
Dim fromDate  As Date
Dim toDate As Date

fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Dim SqlStr As String

    SqlStr = "SELECT SUM(PenalIntAmount) as TotalAmount From BKCCIntTrans" & _
        " WHERE FromDate >= #" & fromDate & "#" & _
        " AND TransDate <= #" & toDate & "# " & _
        " AND (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & " )" & _
        " AND Deposit = False "
        
    gDbTrans.SqlStmt = SqlStr
Dim rst As Recordset
    If gDbTrans.Fetch(rst, adOpenDynamic) Then _
        LoanInterest_Penal = FormatField(rst("TotalAmount"))

End Function

'Shashi Angadi
'20 August 2002
'THis funcion returns the amount collected between
'specified dates as regular interest for loans
'Arguments date in "dd/mm/yyyy"
Public Function LoanInterest_Regular(FromIndianDate As String, ToIndianDate As String)
Dim fromDate  As Date
Dim toDate As Date

fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)


Dim SqlStr As String

    SqlStr = "SELECT SUM(IntAmount) as TotalAmount From BKCCIntTrans" & _
        " WHERE FromDate >= #" & fromDate & "#" & _
        " AND TransDate <= #" & toDate & "# " & _
        " AND (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & " )" & _
        " AND Deposit = False "
        
    gDbTrans.SqlStmt = SqlStr
Dim rst As Recordset
    If gDbTrans.Fetch(rst, adOpenDynamic) Then _
        LoanInterest_Regular = FormatField(rst("TotalAmount"))

End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Loass coccurder throgh BKCC loan
'during the specified period
'Argument FromDate And to date in Indian Format
'Returns amount (Profit Throuth Loan
Public Function LoanLoss(FromIndianDate As String, ToIndianDate As String) As Currency
On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
    'Now take the amount Where SACH IS NOTH HANDLED
    transType = wContraWithdraw
    Amount = Amount + GetIntTransactedAmount(fromDate, toDate, transType, False)
    
Err_line:
    
    LoanLoss = Amount
End Function

'this function returns the Over due amount of loan as on given date
'If fromdate is also given then
'it returns the overdue amount during the specified period
Private Function LoanOverDue(LoanID As Long, AsOnDate As Date) As Currency

Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim rstTrans As Recordset
 
Dim ODAmount As Currency

On Error GoTo Exit_Line

'Now Get the Balance Of All loans One Year Before
gDbTrans.SqlStmt = "SELECT Loanid, Balance, TransDate, TransID FROM " _
        & " BKCCTrans WHERE LoanID= " & LoanID _
        & " AND TransID = (SELECT MAx(TransID) FROM BKCCTrans " _
            & " Where LoanID = " & LoanID & " And Deposit = False " _
            & " And TransDate < #" & DateAdd("yyyy", -1, AsOnDate) & "# )"

If gDbTrans.Fetch(rst, adOpenStatic) = 0 Then GoTo Exit_Line

'Now Fetch The Amount Repaid in a year
gDbTrans.SqlStmt = "SELECT sum(Amount) as RepaidAmount " _
        & " FROM BKCCTrans WHERE LoanID = " & LoanID _
        & " AND (Transtype = " & wDeposit & " OR Transtype = " & wContraDeposit & ")" _
        & " And TransDate > #" & DateAdd("yyyy", -1, AsOnDate) & "# " _
        & " And TransDate <= #" & AsOnDate & "# And Deposit = False "

If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then GoTo Exit_Line

Dim Balance As Currency
Dim PaidAmount As Currency

Balance = FormatField(rst("Balance"))
PaidAmount = 0
If Not rstTrans Is Nothing Then PaidAmount = FormatField(rstTrans("RepaidAmount"))

ODAmount = Balance - PaidAmount
If ODAmount < 0 Then ODAmount = 0

Err.Clear
On Error GoTo Exit_Line
ODAmount = 0

Exit_Line:
    If Err Then
        MsgBox "ERROR In over duw", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
        Exit Function
    End If
    LoanOverDue = ODAmount

End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Profit Made by the BKCC Deposit
'during the specified period
'Argument FromDate And to date in Indian Format
'Returns amount (Profit Throuth Loan
Public Function LoanProfit(FromIndianDate As String, ToIndianDate As String) As Currency
On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
    'Now take the amount Where CASH IS NOT HANDLED
    transType = wContraDeposit
    Amount = Amount + GetIntTransactedAmount(fromDate, toDate, transType, False)
    
Err_line:
    
    LoanProfit = Amount
End Function



'Shashi Angadi
'tthis function returns the amount withdraw as loan from
'the bkcc accounts
'Modified 9th Augus 2002
Public Function LoanWithDrawls(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, False)
    transType = wContraWithdraw
    Amount = Amount + GetTransactedAmount(fromDate, toDate, transType, False)
    
    LoanWithDrawls = Amount
    
End Function


' Calculates the penal interest for defaulted repayments.
Function PenalInterest(LoanID As Long, TillDate As Date) As Currency

' Setup error handler
Err.Clear
On Error GoTo Exit_Line
' Variables...
'Dim LoanDate As Date
Dim LoanAmount As Currency
Dim RepaidAmount As Currency
Dim PenaltyRate As Single
Dim transType As wisTransactionTypes
Dim LastDate As Date
Dim TransID As Long
Dim ODAmount As Currency
Dim ODIntAmount As Currency
Dim ODDate As Date
Dim BalanceAmount As Currency

Dim rstMaster As Recordset
Dim rstTrans As Recordset

'Load the Repayment date to a local variable
If LoanID <= 0 Then Exit Function

gDbTrans.SqlStmt = "SELECT * FROM BKCCMaster" & _
                        " WHERE loanID = " & LoanID
If gDbTrans.Fetch(rstMaster, adOpenStatic) <= 0 Then GoTo Exit_Line

If FormatField(rstMaster("LoanClosed")) Then Exit Function
PenaltyRate = FormatField(rstMaster("PenalIntRate"))
If PenaltyRate <= 0 Then Exit Function

'For the Kcc Account The Repayment period is of one year
'from the date when the particular amount is issued
'Here in this loan Penal INterest is like installment Amount

'First Check the Amount whether is there any due Amout
'so get the Balance before one year
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT Balance,TransDate,Amount,TransType,TransID " & _
        " From BKCCTrans WHERE LoanId = " & LoanID & _
        " AND TransID = (Select Max(TransID) From BKCCTrans Where LoanId = " & _
            LoanID & " ANd TransDate < #" & DateAdd("yyyy", -1, TillDate) & "#)"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function
BalanceAmount = FormatField(rst("Balance"))
LoanAmount = FormatField(rst("Amount"))
transType = FormatField(rst("TransType"))
TransID = FormatField(rst("TransID"))
LastDate = rst("TransDate")

If BalanceAmount <= 0 Then Exit Function

'Now Get the Repaid Amount After That date till today
gDbTrans.SqlStmt = "SELECT Sum(Amount) as RepaidAmount From BKCCTrans " & _
    " WHere LoanId = " & LoanID & " AND Deposit = False " & _
    " AND TransDate > #" & LastDate & "# AND TransDate <= #" & TillDate & "# " & _
    " AND (TransType = " & wDeposit & " OR  TransType = " & wContraDeposit & ")"

RepaidAmount = 0
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
            RepaidAmount = FormatField(rst("RepaidAmount"))

ODAmount = BalanceAmount - RepaidAmount
'If there is no over due amount then no penal interest
If ODAmount <= 0 Then Exit Function

'Get the Balance before this tranasction
If transType = wContraDeposit Or transType = wDeposit Then
    RepaidAmount = RepaidAmount + LoanAmount
    BalanceAmount = BalanceAmount + LoanAmount
Else
    BalanceAmount = BalanceAmount - LoanAmount
End If

'Get the Over due amount if any as on last transaction before 1 year
'ODAmount = BalanceAmount - RepaidAmount

If BalanceAmount - RepaidAmount <= 0 Then
    'if there is no OVER dUE amount as on last date then
    ODDate = DateAdd("yyyy", 1, LastDate)
    ODIntAmount = ODAmount * (DateDiff("D", ODDate, TillDate) / 365) * (PenaltyRate / 100)
    PenalInterest = ODIntAmount \ 1
    GoTo Exit_Line
End If

gDbTrans.SqlStmt = "SELECT Balance,TransDate,Amount,TransType,TransID" & _
        " FROM BKCCTrans Where LoanId = " & LoanID & _
        " AND TransID < " & TransID & " ORDER By TransId Desc"

'Get the previous balance which is over due amount
    'so TO get the previous balance
    'loop through previous transaction
    'and get on which loan amount the od False
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
Do While BalanceAmount - RepaidAmount > 0
    If rst.EOF Then Exit Do
    gDbTrans.SqlStmt = "SELECT Top 1 Balance,TransDate,Amount,TransType,TransID" & _
        " FROM BKCCTrans Where LoanId = " & LoanID & _
        " AND TransID <= " & TransID - 1 & " ORDER By TransId Desc"
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Exit Do
    LoanAmount = FormatField(rst("Amount"))
    transType = FormatField(rst("TransType"))
 '   LastDate = Rst("TransDate")
'    TransID = FormatField(Rst("TransID"))
'    BalanceAmount = FormatField(Rst("Balance"))
    
    'Get the Loan Balance Before This Transction
    If transType = wContraDeposit Or transType = wDeposit Then
        RepaidAmount = RepaidAmount + LoanAmount
        BalanceAmount = BalanceAmount + LoanAmount
    Else
        'Get the Loan Blance Before This Transction
        BalanceAmount = BalanceAmount - LoanAmount
    End If
    rst.MoveNext
Loop

    BalanceAmount = FormatField(rst("Balance"))
'    Debug.Assert 1 = 0
End If

gDbTrans.SqlStmt = "SELECT * From BKCCTrans " & _
        " WHERE LoanId = " & LoanID & " AND TransDate >= #" & LastDate & "# " & _
        " AND TransDate <= #" & TillDate & "# AND Deposit = False ORDER BY TransID"
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)

'The Amount with drwan After last date is over due
'Calculate the Over Due
ODDate = DateAdd("yyyy", 1, LastDate)
Do
    If rstTrans.EOF Then Exit Do
    BalanceAmount = FormatField(rstTrans("Balance"))
    LastDate = rstTrans("TransDate")
    ODDate = DateAdd("yyyy", 1, LastDate)
    If ODDate > TillDate Then Exit Do
    gDbTrans.SqlStmt = "SELECT Sum(Amount) as RepaidAmount From BKCCTrans " & _
        " WHere LoanId = " & LoanID & " AND Deposit = False " & _
        " AND TransDate > #" & LastDate & "# AND TransDate <= #" & ODDate & "# " & _
        " AND (TransType = " & wDeposit & " OR  TransType = " & wContraDeposit & ")"
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    ODAmount = BalanceAmount - FormatField(rst("RepaidAmount"))
    If ODAmount > 0 Then _
        ODIntAmount = ODIntAmount + _
            (ODAmount * (DateDiff("D", ODDate, TillDate) / 365) * (PenaltyRate / 100))
    rstTrans.MoveNext
Loop

PenalInterest = ODIntAmount \ 1

Exit Function

Exit_Line:
    If Err Then
        MsgBox "ComputePenalInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
    End If
End Function
'This Function returns the Over due amount
'of loan/Scheme on Speicified date
'
Public Function OverDueAmount(Optional LoanID As Long = 0, Optional AsOnDate As Date) As Currency

OverDueAmount = 0

'Dim AsOndate As Date
Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim Duration As Long
Dim InstType As wisInstallmentTypes


Err.Clear
On Error GoTo ExitLine
ODAmount = 0

'Get The over due Amount of all loan scheme
'Get The Loan over due of the Particular loan scheme
If LoanID = 0 Then
    ODAmount = SchemeOverDue(AsOnDate)
Else
    ODAmount = LoanOverDue(LoanID, AsOnDate)
End If


ExitLine:
    If Err Then
        MsgBox "ERROR In Over Due amount ", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    OverDueAmount = ODAmount \ 1

End Function

Public Function LoanDueDate(LoanID As Long) As Date

Dim rst As Recordset
Dim LastDate As Date

LastDate = DateAdd("yyyy", 1, CDate(gStrDate))

Err.Clear
On Error GoTo ExitLine

'Get the Transactin ID & Date whose balance has not paid yet
gDbTrans.SqlStmt = " Select Top 1 TransDate,TransID From BKCCTrans A" & _
        " Where LoanID = " & LoanID & " AND Balance >=" & _
            " (Select sum(Amount) From " & _
            " BkccTrans B Where B.LoanID = " & LoanID & " And" & _
            " (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
            " And B.TransID > A.TransID)" & _
        " Order by TransID"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
LastDate = DateAdd("yyyy", 1, rst("TransDate"))


ExitLine:
    If Err Then _
        MsgBox "ERROR In Penal Interest", vbInformation, wis_MESSAGE_TITLE: Err.Clear
    
LoanDueDate = LastDate

End Function


Public Function OverDueInterest(LoanID As Long, AsOnDate As Date) As Currency

OverDueInterest = 0

'Dim AsonDate As Date
Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim ODInt As Currency
Dim IntRate As Single
Dim LastDate As Date

Err.Clear
On Error GoTo ExitLine
ODInt = 0

ODInt = PenalInterest(LoanID, AsOnDate)
OverDueInterest = ODInt \ 1
Exit Function

'Get the Loan DueDate & Intreatrate
SqlStr = "SELECT * FROM BKCCMaster Where " & _
    " LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
IntRate = FormatField(rst("IntRate"))

'Interest Over due is Interest on OverDue Amount
'So Get the Interest Amount
'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Top 1 * FROM BKCCIntTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate <=  #" & AsOnDate & "#" & _
    " ORDER BY TransID Desc"

gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
'Set Rst = gDBTrans.Rst.Clone
ODInt = FormatField(rst("IntBalance"))
LastDate = rst("TransDate")

If DateDiff("d", DueDate, AsOnDate) > 0 Then
    ODInt = FormatField(rst("IntBalance"))
    GoTo ExitLine
End If

ODInt = ODInt + PenalInterest(LoanID, AsOnDate)
ODAmount = OverDueAmount(LoanID, AsOnDate)
If ODAmount = 0 Then GoTo ExitLine
ODInt = ODInt '+ Balance(LoanID , AsOnDate) * (DateDiff("d", LastDate, AsOnDate) / 365) * IntRate

ExitLine:
    If Err Then
        MsgBox "ERROR In Penal Interest", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    OverDueInterest = ODInt \ 1
End Function



'
Function OverDueSince(NoOfYear As Integer, _
    Optional LoanID As Long = 0, Optional AsOnDate As Date) As Currency

Err.Clear

On Error GoTo ExitLine

Dim ODAmount As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim rstTrans As Recordset
Dim DueDate As Date

DueDate = DateAdd("yyyy", -NoOfYear, AsOnDate)

'First Get the Balance on ONE YEAR BEFORE Due date
'Then find the amount repaid during the this Period
SqlStr = "SELECT A.Loanid, Balance, TransDate, TransID FROM " _
    & " BkccMaster A, BKCCTrans B WHERE " _
    & " A.LoanID = " & IIf(LoanID, LoanID, "A.LoanID") _
    & " AND TransID = (SELECT MAx(TransID) FROM BKCCTrans D " _
        & " Where D.LoanID = " & IIf(LoanID, LoanID, "A.LoanID") _
        & " And TransDate < #" & DateAdd("yyyy", -1, DueDate) & "# )"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo ExitLine
If FormatField(rst("Balance")) < 0 Then GoTo ExitLine

'Now Fetch The Amount Repaid in a year
SqlStr = "SELECT Loanid, sum(Amount) as RepaidAmount " _
        & " FROm BKCCTrans B WHERE " _
        & " (Transtype = " & wDeposit & " OR Transtype = " & wContraDeposit & ")" _
        & " And TransDate > #" & DateAdd("yyyy", -1, DueDate) & "# " _
        & " And TransDate <= #" & AsOnDate & "# "
SqlStr = SqlStr & IIf(LoanID, " AND LoanId = " & LoanID, "") & " Group BY LoanID "

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Set rstTrans = Nothing
SqlStr = ""
    
Dim Balance As Currency
Dim PaidAmount As Currency
Dim TotalOD As Currency

While Not rst.EOF
    Balance = FormatField(rst("Balance"))
    PaidAmount = 0
    If Not rstTrans Is Nothing Then
        rstTrans.MoveFirst
        rstTrans.Find "LoanID = " & rst("LoaniD")
        If Not rstTrans.EOF Then PaidAmount = FormatField(rstTrans("RepaidAmount"))
    End If
    ODAmount = Balance - PaidAmount
    If ODAmount > 0 Then TotalOD = TotalOD + ODAmount
    rst.MoveNext
Wend

ExitLine:
    If Err Then
        MsgBox "ERROR in OverDue since" & vbCrLf & Err.Description _
        , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
    OverDueSince = TotalOD \ 1

End Function



Function RegularInterest(Optional LoanID As Long, Optional TillDate As Date) As Currency
' Setup error handler.
'If Not DateValidate(TillIndianDate, "/", True) Then Exit Function
On Error GoTo Err_line

' Define varaibles.
Dim Lret As Long
Dim InterestAmt As Currency
Dim InterestRate As Single
Dim SlabIntRate As Single
Dim Balance As Long
Dim Duration As Long
Dim ClsInt As New clsInterest
Dim IntRate As Single

Dim NextDate As String
Dim LastIntDate As Date
Dim LastTransDate As Date

Dim rstMaster As Recordset
Dim rstTrans As Recordset

If LoanID <= 0 Then Exit Function

'As this function is called while showing reports
'the addational check og the Account existance will
'hit the performance
gDbTrans.SqlStmt = "SELECT * FROM BKccMaster WHERE loanID = " & LoanID
Lret = gDbTrans.Fetch(rstMaster, adOpenDynamic)
If Lret <= 0 Then GoTo Exit_Line

    gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans WHERE " _
                & "loanid = " & LoanID & " ORDER BY transID"
    Lret = gDbTrans.Fetch(rstTrans, adOpenDynamic)
    If Lret <= 0 Then GoTo Exit_Line

If FormatField(rstMaster("LoanClosed")) Then Exit Function

' Get the rate of interest.
InterestRate = Val(FormatField(rstMaster("IntRate")))

'Get the balance amount or/and  date till interest paid for this loan account.
'There is Only one option whehter He has to Consider the Date till int paid
'Or the Interest Balance
Dim rstTemp As Recordset

gDbTrans.SqlStmt = "SELECT TOP 1 Amount,balance,transdate," & _
                " TransID,Deposit FROM BKCCTrans " & _
                " WHERE loanid = " & LoanID & " And TransDate <= #" & TillDate & "#" & _
                " ORDER BY Transid DESC"
Lret = gDbTrans.Fetch(rstTemp, adOpenDynamic)
If Lret = 0 Then Exit Function
If Lret < 0 Then
    MsgBox "Error retrieving loan details from database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
    
Balance = CCur(FormatField(rstTemp("Balance")))
If Balance < 0 Then Exit Function

LastTransDate = rstTemp("TransDate")
If FormatField(rstMaster("LastIntDate")) <> "" Then
    LastIntDate = rstMaster("LastIntDate")
    If DateDiff("d", LastIntDate, LastTransDate) < 1 Then LastIntDate = LastTransDate
Else
    LastIntDate = vbNull
End If

' Get the duration in days from the last transaction date
' until this day.
' Duration = DateDiff("d", LastIntDate, TillDate)

'If Duration < 0 Then GoTo Exit_Line
On Error GoTo Err_line

If LastIntDate = vbNull Then
    gDbTrans.SqlStmt = "SELECT Top 1 * From BKCCIntTrans WHERE " & _
        " LoanID = " & LoanID & " AND (IntAmount > 0 OR PenalIntAmount > 0) " & _
        " AND TransDate <= #" & TillDate & "# ORDER BY TransID Desc "
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        LastIntDate = rstTrans("TransDate")
    Else
        LastIntDate = rstMaster("IssueDate")
    End If
End If

'Now Get the Transaction after the LastIntPaid date
gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans Where LoanID = " & LoanID & _
                    " AND TransDate >= #" & LastIntDate & "#" & _
                    " AND TransDate <= #" & TillDate & "# ORDER BY TransID"
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then GoTo Exit_Line

Dim IntAmount As Currency
Dim AsOnDate As Date
Dim PrevBalance As Currency
Dim SlabInt As Boolean

SlabIntRate = GetKCCSlabIntRate(Balance)
If SlabIntRate > 0 Then SlabInt = True: InterestRate = SlabIntRate

IntAmount = 0
Balance = rstTrans("Balance")

Do
    If Balance < 0 Then Exit Do 'InterestRate = Val(FormatField(RstMaster("DepIntRate")))
'    If Balance > 0 Then InterestRate = Val(FormatField(RstMaster("IntRate")))
    
    'Now Get the Slb interest rate
    If Balance <> PrevBalance Then
        If SlabInt Then InterestRate = GetKCCSlabIntRate(Balance)
    End If
    If rstTrans.EOF Then
        Duration = DateDiff("d", CDate(LastIntDate), TillDate)
        'If Balance = 0 Then IntAmount = 0
        If Balance > 0 Then
            IntAmount = IntAmount + ((Duration / 365) * (InterestRate / 100) * Balance)
        ElseIf Duration >= 30 Then
            IntAmount = IntAmount + ((Duration / 365) * (InterestRate / 100) * Balance)
        End If
        Exit Do
    End If
    AsOnDate = rstTrans("TransDate")
    Duration = DateDiff("d", LastIntDate, AsOnDate)
    If Balance > 0 Then
        IntAmount = IntAmount + ((Duration / 365) * (InterestRate / 100) * Balance)
    ElseIf Duration > 30 Then
        IntAmount = IntAmount + ((Duration / 365) * (InterestRate / 100) * Balance)
    End If

    LastIntDate = AsOnDate
    'Now Store the the calulated balance
    PrevBalance = Balance
    Balance = FormatField(rstTrans("BALANCE"))
    rstTrans.MoveNext

Loop
    
GoTo Exit_Line

DepositInterest:

InterestRate = Val(FormatField(rstMaster("DepIntRate")))
    
    
'''m_InterestBalance
Exit_Line:
    RegularInterest = IntAmount \ 1

    Exit Function

Err_line:
    If Err Then
        MsgBox "RegularInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
     End If

End Function

Public Property Get ObjectName()
    ObjectName = "KCCClass"
End Property
'This Function Returns the over due amount as on date
'If SchmeID is missing it returns the over due of All
'Type of loans , If Date is missing it returns OD
'as on the login date
Private Function SchemeOverDue(Optional AsOnDate As Date) As Currency

Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim rstTrans As Recordset
 
Dim ODAmount As Currency
Dim TotalOD As Currency

'Now Get the Balance Of All loans One Year Before
gDbTrans.SqlStmt = "SELECT A.Loanid, Balance, TransDate, TransID FROM " _
        & " BkccMaster A, BKCCTrans B WHERE " _
        & " TransID = (SELECT MAx(TransID) FROM BKCCTrans D " _
            & " Where D.LoanID = A.LoanID " _
            & " And TransDate < #" & DateAdd("yyyy", -1, AsOnDate) & "# )" _
        & " And A.LoanID= B.LoanID "

If gDbTrans.Fetch(rst, adOpenStatic) = 0 Then GoTo Exit_Line

'Now Fetch The Amount Repaid in a year
gDbTrans.SqlStmt = "SELECT Loanid, sum(Amount) as RepaidAmount " _
        & " FROm BKCCTrans B WHERE " _
        & " (Transtype = " & wDeposit & " OR Transtype = " & wContraDeposit & ")" _
        & " And TransDate > #" & DateAdd("yyyy", -1, AsOnDate) & "# " _
        & " And TransDate <= #" & AsOnDate & "# )" _
        & " Group BY LoanID "
Call gDbTrans.Fetch(rstTrans, adOpenStatic)

Dim Balance As Currency
Dim PaidAmount As Currency

While Not rst.EOF
    Balance = FormatField(rst("Balance"))
    PaidAmount = 0
    If Not rstTrans Is Nothing Then
        rstTrans.MoveFirst
        rstTrans.Find "LoanID = " & rst("LoaniD")
        If Not rstTrans.EOF Then PaidAmount = FormatField(rstTrans("RepaidAmount"))
    End If
    ODAmount = Balance - PaidAmount
    If ODAmount > 0 Then TotalOD = TotalOD + ODAmount
    rst.MoveNext
Wend

Err.Clear
On Error GoTo Exit_Line
ODAmount = 0

'The sum Over due Of all the Account is over due of the scheme
'Find The Over Due Of Each Loan

Exit_Line:
    If Err Then
        MsgBox "ERROR In Scheme over due", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    SchemeOverDue = TotalOD \ 1

End Function


'Shashi Angadi
'this function returns the amount withdraw as Deposit from
'From BKCC accountand transferred to other account
'Modified 9th Augus 2002
Public Function ContraDepositWithDrawls(FromIndianDate As String, ToIndianDate As String, SchemeID As Long) As Currency

If Not m_ConsiderDeposit Then Exit Function

On Error GoTo Err_line
    
    Dim fromDate As Date
    Dim toDate As Date
    
    fromDate = GetSysFormatDate(FromIndianDate)
    toDate = GetSysFormatDate(ToIndianDate)
    
    Dim Amount As Currency
    Dim transType As wisTransactionTypes
    
    transType = wWithdraw
    Amount = GetTransactedAmount(fromDate, toDate, transType, True)
    
    ContraDepositWithDrawls = Amount

Err_line:
    If Err Then
        MsgBox "Error In Bkcc Class Module", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function

'Shashi Angadi
'9th August 2002
'This Function Returns the Loass coccurder throgh BKCC loan
'during the specified period trough
'Argument FromDate And to date in Indian Format
'Returns amount (Profit Throuth Loan
Public Function CashLoanLoss(FromIndianDate As String, ToIndianDate As String) As Currency
On Error GoTo Err_line

Dim Amount As Currency
Dim transType As wisTransactionTypes

Dim fromDate As Date
Dim toDate As Date
fromDate = GetSysFormatDate(FromIndianDate)
toDate = GetSysFormatDate(ToIndianDate)

Amount = 0
    'First Take the Amount where cash is Handled
    transType = wWithdraw
    Amount = GetIntTransactedAmount(fromDate, toDate, transType, False)
    
    
Err_line:
    
    CashLoanLoss = Amount
End Function

Public Sub Show()
    If LoanDlg Is Nothing Then Set LoanDlg = New frmBKCCAcc
    Load LoanDlg
    gWindowHandle = LoanDlg.hwnd
    LoanDlg.Show 'vbModal
    If m_LoanID Then LoanDlg.LoanLoad (m_LoanID)
    
End Sub

Public Property Get ObjectHandle() As Long
    If Not m_frmBKCCReport Is Nothing Then _
            ObjectHandle = m_frmBKCCReport.hwnd: Exit Property
    If Not LoanDlg Is Nothing Then _
            ObjectHandle = LoanDlg.hwnd: Exit Property
    If Not m_frmBKCCReport Is Nothing Then _
            Set m_frmBKCCReport = Nothing

End Property

Private Sub Class_Initialize()
Dim SetUp As New clsSetup
m_ConsiderDeposit = False
If UCase(SetUp.ReadSetupValue("General", "KCCDeposit", "TRUE")) = "TRUE" Then _
        m_ConsiderDeposit = True

Set SetUp = Nothing
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set LoanDlg = Nothing
Set m_frmBKCCReport = Nothing
End Sub

Private Sub LoanDlg_AccountChanged(ByVal LoanID As Long)
m_LoanID = LoanID
End Sub

Private Sub LoanDlg_ShowReport(ReportType As wis_BKCCReports, _
        ReportOrder As wis_ReportOrder, FarmerType As wisFarmerClassification, _
        fromDate As String, toDate As String, RepOptClass As clsRepOption)

Set m_frmBKCCReport = New frmBkccReport

Screen.MousePointer = vbHourglass

With m_frmBKCCReport

    .FromIndianDate = fromDate
    .ToIndianDate = toDate
    .FromAmount = RepOptClass.FromAmount
    .ToAmount = RepOptClass.ToAmount
    .Place = RepOptClass.Place
    .Caste = RepOptClass.Caste
    .ReportOrder = ReportOrder
    
    .FarmerType = FarmerType
    .Gender = RepOptClass.Gender
    .ReportType = ReportType
    
    Load m_frmBKCCReport
    Screen.MousePointer = vbDefault
    If gCancel = 2 Then GoTo ExitLine
    Unload frmCancel
    .Show 'vbModal
End With

ExitLine:
   On Error Resume Next
   Set frmCancel = Nothing
   Set m_frmBKCCReport = Nothing
   gCancel = 0

Exit_Line:
Screen.MousePointer = vbDefault

End Sub


Private Sub LoanDlg_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set LoanDlg = Nothing
End Sub

Private Sub m_frmBKCCReport_Initialise(Min As Long, Max As Long)
gCancel = 0
frmCancel.Show
UpdateStatus frmCancel.PicStatus, 0
    If Max <> 0 Then
        With frmCancel
            .PicStatus.Visible = True
        End With
    End If

End Sub


Private Sub m_frmBKCCReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS :" & vbCrLf & strMessage
    UpdateStatus frmCancel.PicStatus, Ratio
End With

End Sub


