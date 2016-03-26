VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBankAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_BankAccToCustomerAcc As Boolean
Public Property Let BankAccountToCustomerAccount(ByVal vData As Boolean)
    m_BankAccToCustomerAcc = vData
End Property


Public Property Get ObjectName()
    ObjectName = "BankClass"
End Property

Public Function Balance(ByVal AsonIndianDate As String, ByVal headID As Long) As Currency

'Trap an Error
On Error GoTo ErrLine

Dim AccTransClass As clsAccTrans

Set AccTransClass = New clsAccTrans

Balance = AccTransClass.GetOpBalance(headID, GetSysFormatDate(AsonIndianDate))

Set AccTransClass = Nothing

Exit Function

ErrLine:
    
    MsgBox "Balance: " & vbCrLf & Err.Description, vbCritical
        
End Function



Private Function InsertIntoIndexIds(ByVal headID As Long, _
                ByVal headName As String, AccType As wisModules) As Boolean

InsertIntoIndexIds = False

If headID = 0 Then Exit Function

gDbTrans.SqlStmt = "INSERT INTO BankHeadIds " & _
                "(HeadID,HeadName,AccType,AliasName) " & _
                 " VALUES ( " & _
                 headID & "," & _
                 AddQuotes(headName, True) & ", " & _
                 AccType & "," & _
                 AddQuotes(headName, True) & ")"

If Not gDbTrans.SQLExecute Then Exit Function

InsertIntoIndexIds = True

End Function

Public Function Loss(ByVal FromIndianDate As String, ByVal ToIndianDate As String, ByVal headID As Long) As Currency

' To Trap an Error
On Error GoTo ErrLine

'Declare variables
Dim AccountType As wis_AccountType
Dim rstLoss As ADODB.Recordset

Loss = 0

' Validations
If headID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

' set the accounttype
AccountType = Loss

' This is the sql to fetch the profits

gDbTrans.SqlStmt = " SELECT b.Credit as LossReceived, b.Debit as LossGiven, e.AccountType " & _
                   " FROM qryAccTransmerge AS d, Heads AS c, ParentHeads AS e " & _
                   " WHERE a.HeadId = " & headID & _
                   " AND b.HeadID<>" & headID & _
                   " AND a.TransDate BETWEEN #" & GetSysFormatDate(FromIndianDate) & "#" & _
                   " AND #" & GetSysFormatDate(ToIndianDate) & "#" & _
                   " AND b.HeadID=c.HeadID " & _
                   " AND c.ParentID=e.ParentID " & _
                   " AND e.AccountType=" & AccountType & _
                   " ORDER BY a.TransDate, a.TransID "
                       
' this will create a view
Call gDbTrans.CreateView("QryHeadIDLoss")

' this will fetch the details

gDbTrans.SqlStmt = " SELECT (SUM(LossGiven)-SUM(LossReceived)) as Loss" & _
                   " FROM qryHeadIDLoss"
                   
Call gDbTrans.Fetch(rstLoss, adOpenForwardOnly)


Loss = rstLoss.Fields("Loss")

Exit Function

ErrLine:
    MsgBox "Loss: " & vbCrLf & Err.Description, vbCritical
        
End Function

Public Sub RepairPigmyDeposits()
''This function created to repir the any Pigmy Deposit which have
'invalid entries in the Acctrans table
'Shashidhar Angadi  25/12/2004 '
'This function is made to rectify the error
'which is occuring in Huvin Hipparagi
Dim rstPD As New Recordset
Dim PDHeadID As Long

PDHeadID = GetHeadIDCreated(GetResourceString(425), parMemberDeposit, 0, wis_PDAcc)
    'First check the receipt side
gDbTrans.SqlStmt = "Select * From AccTrans Where TransID IN " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID =" & wis_CashHeadID & _
            " AND A.VoucherType = " & Receipt & ")"
If gDbTrans.Fetch(rstPD, adOpenDynamic) < 1 Then GoTo NextLine





NextLine:


End Sub

Public Function UndoContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean


'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If TransDate < FinUSFromDate Then
    UndoContraTrans = True
    Exit Function
End If

UndoContraTrans = False

If gOnLine Then UndoContraTrans = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then UndoContraTrans = True: Exit Function

If CrHeadID = DrHeadID Then
    'There is chance of undo the transactionfrom the same Head
    If Not UndoCashDeposits(CrHeadID, Amount, TransDate) Then Exit Function
    If Not UndoCashWithdrawls(DrHeadID, Amount, TransDate) Then Exit Function
    UndoContraTrans = True
    Exit Function
End If

VoucherType = Journal

If UndoTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
UndoContraTrans = True

Exit Function

Hell:
    
    UndoContraTrans = False
    
End Function
Public Function UpdateContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If TransDate < FinUSFromDate Then
    UpdateContraTrans = True
    Exit Function
End If


UpdateContraTrans = False

If gOnLine Then UpdateContraTrans = True: Exit Function

On Error GoTo Hell:

'Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function

If Amount = 0 Then UpdateContraTrans = True: Exit Function

If CrHeadID = DrHeadID Then
    'ther is chance of withdrawing accounthead and depositing accounthead
    'might be same In such case the transaction will not allowed
    'so withdraw of amount will be with draw n from the cash
    'and will be deposited to cash head
    If Not UpdateCashDeposits(CrHeadID, Amount, TransDate) Then Exit Function
    If Not UpdateCashWithDrawls(DrHeadID, Amount, TransDate) Then Exit Function
    UpdateContraTrans = True
    Exit Function
End If


VoucherType = Journal

If UpdateTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
UpdateContraTrans = True

Exit Function

Hell:
    
    UpdateContraTrans = False
    
End Function

Public Function AddContraTrans(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraTrans = False

If Not gOnLine Then AddContraTrans = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then AddContraTrans = True: Exit Function

'If Not DateValidate(TransDate, "/", False) Then Exit Function
Dim USTransDate As Date
VoucherType = Journal

Set AccTransClass = New clsAccTrans
                
If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.SaveVouchers(VoucherType, TransDate) <> Success Then Exit Function

Set AccTransClass = Nothing

AddContraTrans = True

Exit Function

Hell:
    
    AddContraTrans = False
    
End Function

Public Function UpdateCashDeposits(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If TransDate < FinUSFromDate Then
    UpdateCashDeposits = True
    Exit Function
End If

UpdateCashDeposits = False

If gOnLine Then UpdateCashDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If headID = 0 Then Exit Function
If Amount = 0 Then UpdateCashDeposits = True: Exit Function

CrHeadID = headID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If UpdateTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    UpdateCashDeposits = True

Exit Function

Hell:
    
    UpdateCashDeposits = False
    
End Function

Public Function AddCashDeposits(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashDeposits = False

If Not gOnLine Then AddCashDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If headID = 0 Then Exit Function
If Amount = 0 Then AddCashDeposits = True: Exit Function

CrHeadID = headID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddCashDeposits = True

Exit Function

Hell:
    
    AddCashDeposits = False
    
End Function

Public Function AddContraDeposits(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraDeposits = False

If Not gOnLine Then AddContraDeposits = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If headID = 0 Then Exit Function
If Amount = 0 Then AddContraDeposits = True: Exit Function

VoucherType = Receipt

CrHeadID = headID
DrHeadID = 1

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddContraDeposits = True

Exit Function

Hell:
    
    AddContraDeposits = False
    
End Function


Public Function AddCashWithDrawls(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddCashWithDrawls = False

If Not gOnLine Then AddCashWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If headID = 0 Then Exit Function
If Amount = 0 Then AddCashWithDrawls = True: Exit Function


CrHeadID = wis_CashHeadID
DrHeadID = headID

VoucherType = payment

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddCashWithDrawls = True

Exit Function

Hell:
    
    AddCashWithDrawls = False
    
End Function

Public Function AddContraWithDrawls(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean

AddContraWithDrawls = False

If Not gOnLine Then AddContraWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If headID = 0 Then Exit Function
If Amount = 0 Then AddContraWithDrawls = True: Exit Function


CrHeadID = 1
DrHeadID = headID

VoucherType = payment

If AddTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    AddContraWithDrawls = True

Exit Function

Hell:
    
    AddContraWithDrawls = False
    
End Function

Public Function UndoCashWithdrawls(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date) As Boolean


'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If USTransDate > FinUSEndDate Or USTransDate < FinUSFromDate Then
    UndoCashWithdrawls = True
    Exit Function
End If


UndoCashWithdrawls = False

If gOnLine Then UndoCashWithdrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If headID = 0 Then Exit Function
If Amount = 0 Then UndoCashWithdrawls = True: Exit Function

DrHeadID = headID
CrHeadID = wis_CashHeadID

VoucherType = payment

If UndoTransaction(CrHeadID, DrHeadID, Amount, USTransDate, VoucherType) = Success Then _
UndoCashWithdrawls = True

Exit Function

Hell:
    
    UndoCashWithdrawls = False
    
End Function


Public Function UndoCashDeposits(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date) As Boolean

'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If USTransDate > FinUSEndDate Or USTransDate < FinUSFromDate Then
    UndoCashDeposits = True
    Exit Function
End If



UndoCashDeposits = False

If gOnLine Then UndoCashDeposits = False: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long


If headID = 0 Then Exit Function
If Amount = 0 Then UndoCashDeposits = True: Exit Function


CrHeadID = headID
DrHeadID = wis_CashHeadID

VoucherType = Receipt

If UndoTransaction(CrHeadID, DrHeadID, Amount, USTransDate, VoucherType) = Success Then _
UndoCashDeposits = True

Exit Function

Hell:
    
    UndoCashDeposits = False
    
End Function

Public Function UpdateCashWithDrawls(ByVal headID As Long, _
                                ByVal Amount As Currency, _
                                ByVal TransDate As Date) As Boolean


'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If TransDate < FinUSFromDate Then
    UpdateCashWithDrawls = True
    Exit Function
End If

UpdateCashWithDrawls = False

If gOnLine Then UpdateCashWithDrawls = True: Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans
Dim VoucherType As Wis_VoucherTypes
Dim DrHeadID As Long
Dim CrHeadID As Long

If headID = 0 Then Exit Function
If Amount = 0 Then UpdateCashWithDrawls = True: Exit Function

DrHeadID = headID
CrHeadID = wis_CashHeadID

VoucherType = payment

If UpdateTransaction(CrHeadID, DrHeadID, Amount, TransDate, VoucherType) = Success Then _
    UpdateCashWithDrawls = True

Exit Function

Hell:
    
    UpdateCashWithDrawls = False
    
End Function

Public Function UndoBalances(ByVal CrHeadID As Long, _
                                ByVal DrHeadID As Long, _
                                ByVal Amount As Currency, _
                                ByVal USTransDate As Date, _
                                ByVal VoucherType As Wis_VoucherTypes) As Boolean



'if transaction is not in the current financialyear then
'do not delete the Transction in Acctranstable
If USTransDate > FinUSEndDate Or USTransDate < FinUSFromDate Then
    UndoBalances = True
    Exit Function
End If

UndoBalances = False

If gOnLine Then Exit Function

On Error GoTo Hell:

Dim AccTransClass As clsAccTrans

If CrHeadID = 0 Then Exit Function
If DrHeadID = 0 Then Exit Function
If Amount = 0 Then UndoBalances = True: Exit Function

Set AccTransClass = New clsAccTrans

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UndoVouchers(VoucherType, USTransDate) <> Success Then Exit Function

UndoBalances = True

Exit Function

Hell:
    UndoBalances = False
    
End Function
'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function AddTransaction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

If Not gOnLine Then AddTransaction = Success: Exit Function

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim headID As Long
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean
Dim Debit As Currency
Dim Credit As Currency

headID = IIf(VoucherType = payment, CrHeadID, DrHeadID)

Set AccTransClass = New clsAccTrans
gDbTrans.SqlStmt = "Select * From AccTrans Where TransID = " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID = " & headID & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")"

If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Set rstTrans = Nothing
Set AccTransClass = New clsAccTrans

If Not rstTrans Is Nothing Then
    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        headID = FormatField(rstTrans.Fields("HeadID"))
        If headID = CrHeadID Then CrHeadExists = True
        If headID = DrHeadID Then DrHeadExists = True
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
        
        'Update the Collections
        If AccTransClass.AllTransHeadsAdd(headID, Debit, Credit) <> Success Then Exit Function
        rstTrans.MoveNext
    Loop
End If

If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)

If AccTransClass.AllTransHeadsAdd(DrHeadID, Amount, 0) <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, 0, Amount) <> Success Then Exit Function
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

AddTransaction = Success
Hell:

Exit Function

End Function

'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function UpdateTransaction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

If gOnLine Then UpdateTransaction = Success: Exit Function

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim headID As Long
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean
Dim CRDebit As Currency
Dim DRDebit As Currency

Dim CRCredit As Currency
Dim DRCredit As Currency

Dim sqlSupport As String
Dim strTrans As String

headID = IIf(VoucherType = payment, CrHeadID, DrHeadID)

If VoucherType = Journal Then
    headID = CrHeadID
    CrHeadID = DrHeadID
    DrHeadID = headID
End If

strTrans = ""

If headID = wis_CashHeadID Then
    sqlSupport = " = " & wis_CashHeadID & " "
    
    'The Query(View) QryTransAcc called  used in the following code
    'will be created later in this pricedure as of now this query is not existing
    strTrans = "Select Distinct transid from QryTransAcc " & _
        " where vouchertype=" & VoucherType & _
        " and headid =100001 And " & IIf(VoucherType = payment, "Debit", "Credit") & " > 0"
Else
    sqlSupport = " IN (Select Distinct HeadId From BankHeadIds) "
End If
If VoucherType = Journal Then strTrans = ""

Set AccTransClass = New clsAccTrans

'AccTransClass.ForcedTransaction = True

gDbTrans.SqlStmt = "Select * From AccTrans Where TransID IN " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID " & sqlSupport & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")" & _
            " Order By TransID"

gDbTrans.CreateView ("qryTransAcc")
''first Check for the No Of TransActions U get
Dim transCount As Integer

gDbTrans.SqlStmt = "Select Distinct TransID From qryTransAcc"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) >= 0 Then
    While Not rstTrans.EOF
        transCount = transCount + 1
        rstTrans.MoveNext
    Wend
End If


gDbTrans.SqlStmt = "Select * From qryTransAcc"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Set rstTrans = Nothing
Set AccTransClass = New clsAccTrans

ReCheck:


If Not rstTrans Is Nothing Then
    Dim Debit As Currency
    Dim Credit As Currency
    Dim ParentID As Long
    Dim ExlcudeIdFound As Boolean

    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        ExlcudeIdFound = False
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        headID = FormatField(rstTrans.Fields("HeadID"))
        
        ParentID = GetParentID(headID)
        ''Updates the Amount Of the Existing Transaction.
        ''If no transaction created new Transaction
        ''IN case of Bank Accounts We need sapreate transaction,
            ''because each transaction is one transaction as its is not parent head for group of Accounts
        If (ParentID = parBankAccount Or ParentID = parBankLoanAccount) And m_BankAccToCustomerAcc Then
            
            ExlcudeIdFound = True
            'GoTo NextRecord
        End If
        
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
        
        If headID = CrHeadID Then
            CrHeadExists = True
            CRDebit = Debit
            CRCredit = Credit
        ElseIf headID = DrHeadID Then
            DrHeadExists = True
            DRDebit = Debit
            DRCredit = Credit
        End If
        'Update the Collections
        If AccTransClass.AllTransHeadsAdd(headID, Debit, Credit) _
                    <> Success Then Exit Function

nextRecord:

        rstTrans.MoveNext
    Loop
End If


If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)


If AccTransClass.AllTransHeadsAdd(DrHeadID, DRDebit + Amount, DRCredit) _
        <> Success Then Exit Function
If AccTransClass.AllTransHeadsAdd(CrHeadID, CRDebit, CRCredit + Amount) _
        <> Success Then Exit Function
        
If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function

Dim rst As Recordset

UpdateTransaction = Success
Set AccTransClass = Nothing

Hell:

Exit Function

End Function

'This Function Update the transction to the Acctrans Table
'Here All the HeadId's tobe transferred
Private Function UndoTransaction(ByVal CrHeadID As Long, ByVal DrHeadID As Long, _
        ByVal Amount As Currency, ByVal TransDate As Date, _
        ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

On Error GoTo Hell

Dim AccTransClass As clsAccTrans
Dim rstTrans As Recordset
Dim TransID As Long
Dim headID As Long
Dim ObjectCount As Integer
Dim CrHeadExists As Boolean
Dim DrHeadExists As Boolean

Dim CRDebit As Currency
Dim DRDebit As Currency
Dim CRCredit As Currency
Dim DRCredit As Currency
Dim Debit As Currency
Dim Credit As Currency

Dim sqlSupport As String


headID = IIf(VoucherType = payment, CrHeadID, DrHeadID)

If VoucherType = Journal Then
    headID = CrHeadID
    CrHeadID = DrHeadID
    DrHeadID = headID
End If

If headID = wis_CashHeadID Then
    sqlSupport = " = " & wis_CashHeadID & " "
Else
    sqlSupport = " IN (Select Distinct HeadId From BankHeadIds) "
End If

gDbTrans.SqlStmt = "Select * From AccTrans Where TransID IN " & _
        " (Select Distinct TransID From QryAccBankTrans " & _
            " Where B.HeadID " & sqlSupport & _
            " AND TransDate = #" & TransDate & "#" & _
            " ANd A.VoucherType = " & VoucherType & ")" & " Order By TransID"

If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then
    Set rstTrans = Nothing
    'There is possibility Of the the Other HeadID (Non Index Head)
    'has done some transaction in this concept e.g (misceleneous credit during laon transaction)
    'Check for that condition
    If headID = wis_CashHeadID Then
        gDbTrans.SqlStmt = "Select TransID From AccTrans " & _
                " Where TransDate = #" & TransDate & "# " & _
                " And HeadID= " & wis_CashHeadID & " And VoucherType = " & VoucherType & _
                " AND " & IIf(VoucherType = Receipt, "Debit", "Credit") & " = " & Amount
        gDbTrans.SqlStmt = "Select Sum(Debit),TransID From AccTrans" & _
                        " Where TransID = (" & gDbTrans.SqlStmt & ") Group By TransID "
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
          If rstTrans(0) = Amount Then
            gDbTrans.SqlStmt = "DELETE * FROM AccTrans Where TransID = " & rstTrans("TransID")
            If Not gDbTrans.SQLExecute Then GoTo Hell
            UndoTransaction = Success
            Exit Function
          End If
        End If
    End If
End If
Set AccTransClass = New clsAccTrans
'AccTransClass.ForcedTransaction = True

If Not rstTrans Is Nothing Then
    ' Start the Loop
    TransID = rstTrans.Fields("TransID")
    AccTransClass.TransID = TransID
    Do While Not rstTrans.EOF
        If TransID <> rstTrans.Fields("TransID") Then Exit Do
        ' Feed data to Controls
        headID = FormatField(rstTrans.Fields("HeadID"))
        Debit = FormatField(rstTrans.Fields("Debit"))
        Credit = FormatField(rstTrans.Fields("Credit"))
            
        If headID = CrHeadID Then
            CrHeadExists = True
            CRDebit = Debit
            CRCredit = Credit
        ElseIf headID = DrHeadID Then
            DrHeadExists = True
            DRDebit = Debit
            DRCredit = Credit
        End If
        'Update the Collections
        If AccTransClass.AllTransHeadsAdd(headID, Debit, Credit) _
                <> Success Then Exit Function
        ObjectCount = ObjectCount + 1
        rstTrans.MoveNext
    Loop
End If

''Information for the Developer

Debug.Assert DrHeadExists And CrHeadExists

If DrHeadExists Then AccTransClass.AllTransHeadIDRemove (DrHeadID)
If CrHeadExists Then AccTransClass.AllTransHeadIDRemove (CrHeadID)

Debit = DRDebit - Amount: Credit = DRCredit
DRDebit = DRDebit - Amount
If Debit Or Credit Then
    If AccTransClass.AllTransHeadsAdd(DrHeadID, Debit, Credit) <> Success Then Exit Function
Else
    'Delete the this transaction
    ObjectCount = ObjectCount - 1
End If

Debit = CRDebit: Credit = CRCredit - Amount
CRCredit = CRCredit - Amount
If Debit Or Credit Then
    If AccTransClass.AllTransHeadsAdd(CrHeadID, Debit, Credit) <> Success Then Exit Function
Else
    ObjectCount = ObjectCount - 1
End If

If DRDebit Or CRCredit Then
    If AccTransClass.UpdateVouchers(VoucherType, TransDate) <> Success Then Exit Function
ElseIf ObjectCount = 0 Then
    'there are not other transaction so delete this transacion id
    gDbTrans.SqlStmt = "DELETE * FROM AccTrans Where TransID = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo Hell
End If
UndoTransaction = Success

Hell:

Exit Function

End Function


Public Function WithDrawls(ByVal FromIndianDate As String, ByVal ToIndianDate As String, ByVal headID As Long) As Currency

On Error GoTo Err_line:

Dim rstWithdrawal As ADODB.Recordset

WithDrawls = 0

If headID = 0 Then Exit Function
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

gDbTrans.SqlStmt = " SELECT Sum(Credit) as WithDrawal " & _
                   " FROM AccTrans " & _
                   " WHERE HeadID=" & headID & _
                   " AND TransDate BETWEEN " & "#" & GetSysFormatDate(FromIndianDate) & "#" & _
                   " AND #" & GetSysFormatDate(ToIndianDate) & "#"
                  
Call gDbTrans.Fetch(rstWithdrawal, adOpenForwardOnly)
  
WithDrawls = rstWithdrawal.Fields("WithDrawal")
  
Exit Function

Err_line:

    MsgBox "WithDrawls : " & Err.Description
        
End Function
Public Function Profit(ByVal fromDate As Date, ByVal toDate As Date, ByVal headID As Long) As Currency

' To trap an error
On Error GoTo ErrLine

'Declare variables
Dim AccountType As wis_AccountType
Dim rstProfit As ADODB.Recordset

Profit = 0

' Validations
If headID = 0 Then Exit Function

' set the accounttype
AccountType = Profit

' This is the sql to fetch the profits

gDbTrans.SqlStmt = " SELECT b.Credit as ProfitReceived, b.Debit as ProfitGiven, e.AccountType " & _
                   " FROM qryAccTransmerge AS d, Heads AS c, ParentHeads AS e " & _
                   " WHERE a.HeadId = " & headID & _
                   " AND b.HeadID<>" & headID & _
                   " AND a.TransDate BETWEEN #" & fromDate & "#" & _
                   " AND #" & toDate & "#" & _
                   " AND b.HeadID=c.HeadID " & _
                   " AND c.ParentID=e.ParentID " & _
                   " AND e.AccountType=" & AccountType & _
                   " ORDER BY a.TransDate, a.TransID "
                       
Call gDbTrans.CreateView("QryHeadIDProfit")

gDbTrans.SqlStmt = " SELECT (SUM(ProfitReceived)-SUM(ProfitGiven)) as Profits" & _
                   " FROM qryHeadIDProfit"
                   
Call gDbTrans.Fetch(rstProfit, adOpenForwardOnly)

Profit = rstProfit.Fields("Profits")

ErrLine:
    
    MsgBox "Profit: " & vbCrLf & Err.Description, vbCritical
    
End Function


Public Function Deposits(fromDate As Date, toDate As Date, headID As Long) As Currency

Dim rstDeposits As ADODB.Recordset

Deposits = 0

If headID = 0 Then Exit Function

gDbTrans.SqlStmt = " SELECT Sum(Debit) as Deposits " & _
                   " FROM AccTrans " & _
                   " WHERE HeadID=" & headID & _
                   " AND TransDate BETWEEN " & "#" & fromDate & "#" & _
                   " AND #" & toDate & "#"
                 
Call gDbTrans.Fetch(rstDeposits, adOpenForwardOnly)
  
Deposits = rstDeposits.Fields("Deposits")

End Function
Public Function MainHeads(HeadsName() As String, headID() As Long) As Integer

On Error GoTo Hell:

Dim RetLng As Long
Dim rst As ADODB.Recordset

ReDim headNames(0)
ReDim HeadIds(0)
    
gDbTrans.SqlStmt = " SELECT ParentName,ParentID " & _
                   " FROM ParentHeads " & _
                   " WHERE ParentID mod " & HEAD_OFFSET & "=0" & _
                   " ORDER BY ParentName "

RetLng = gDbTrans.Fetch(rst, adOpenStatic)

If RetLng < 1 Then Exit Function
    
ReDim headNames(RetLng - 1)
ReDim HeadIds(RetLng - 1)

RetLng = 0

Do While rst.EOF
    headNames(RetLng) = FormatField(rst("ParentName"))
    HeadIds(RetLng) = FormatField(rst("ParentID"))
    RetLng = RetLng + 1
    rst.MoveNext
Loop

MainHeads = RetLng
Exit Function

Hell:
    
    MsgBox "MainHeads :" & vbCrLf & Err.Description
    
End Function
Public Function GetHeadNames(retHeadNames() As String, _
            retHeadIds() As Long, ByVal ParentID As Long) As Integer

Dim RetLng As Long
Dim rst As ADODB.Recordset

If ParentID = 0 Then Exit Function

ReDim retHeadNames(0)
ReDim retHeadIds(0)
    
gDbTrans.SqlStmt = " SELECT HeadName,HeadID " & _
                   " FROM Heads " & _
                   " WHERE ParentID=" & ParentID
                   
RetLng = gDbTrans.Fetch(rst, adOpenDynamic)
If RetLng < 1 Then Exit Function
    
ReDim retHeadNames(RetLng - 1)
ReDim retHeadIds(RetLng - 1)

RetLng = 0
Do While Not rst.EOF
    retHeadNames(RetLng) = FormatField(rst("HeadName"))
    retHeadIds(RetLng) = FormatField(rst("HeadId"))
    RetLng = RetLng + 1
    rst.MoveNext
Loop

GetHeadNames = RetLng
    
End Function

' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it will create the
' Head with given HeadID & parentID
' Inputs :
'           HeadName as String
'           ParentID as long
' OutPut :
Public Function GetHeadIDCreated(ByVal headName As String, Optional headNameEnglish As String, _
            Optional ParentID As Long, Optional opBalance As Currency, _
            Optional AccType As wisModules) As Long

headName = Trim$(headName)
headNameEnglish = Trim$(headNameEnglish)

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim headID As Long

'initialise the function
GetHeadIDCreated = 0

'Validate the inputs
If headName = "" Then Exit Function

'Check the given Heads in the database
gDbTrans.SqlStmt = "SELECT HeadID,HeadName,HeadNameEnglish FROM BankHeadIds " & _
                        " WHERE HeadName = " & AddQuotes(headName, True)
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    headID = FormatField(rstHeads.Fields(0))
    GetHeadIDCreated = headID
    If Len(FormatField(rstHeads.Fields("HeadNameEnglish"))) > 0 Or Len(Trim$(headNameEnglish)) = 0 Then Exit Function
End If

If ParentID = 0 Then Exit Function

If headID > 0 Then
    'Head is created Just Update the Head name in English
    gDbTrans.SqlStmt = "UPDATE Heads Set[HeadNameEnglish] = " & _
            AddQuotes(headNameEnglish, True) & " WHERE HEADID = " & headID
    
    Call gDbTrans.SQLExecute
    Exit Function
End If

'Get the Maximum Head From the database
gDbTrans.SqlStmt = " SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)
headID = FormatField(rstHeads.Fields(0)) + 1

Set rstHeads = Nothing
    
gDbTrans.SqlStmt = "SELECT headID from BankHeadIDs where HeadID= " & headID
Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)
    
While rstHeads.recordCount <> 0
    headID = headID + 1
    gDbTrans.SqlStmt = "SELECT headID from BankHeadIDs where HeadID= " & headID
    Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

Wend
If headID < ParentID Then headID = headID + ParentID

' Else create the head with given headID & parentiD
'Insert the heads into the database
gDbTrans.SqlStmt = " INSERT INTO Heads (HeadID,HeadName,HeadNameEnglish,ParentID) " & _
                  " VALUES ( " & headID & "," & _
                  AddQuotes(headName, True) & "," & _
                  AddQuotes(headNameEnglish, True) & "," & _
                  ParentID & ")"

If Not gDbTrans.SQLExecute Then Exit Function


'Insert the opening balance
gDbTrans.SqlStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 headID & "," & _
                 "#" & FinUSFromDate & "#," & _
                 opBalance & ")"
If Not gDbTrans.SQLExecute Then Exit Function

If Not InsertIntoIndexIds(headID, headName, AccType) Then Exit Function

GetHeadIDCreated = headID


Exit Function

ErrLine:
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreated = -1
    
    Set rstHeads = Nothing
    
    Exit Function
End Function

