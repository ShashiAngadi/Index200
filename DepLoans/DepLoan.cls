VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDepLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents LoanDlg As frmDepLoan
Attribute LoanDlg.VB_VarHelpID = -1
Private m_rstCustBalance As ADODB.Recordset
Private WithEvents m_frmLoanReport As frmDepLoanReport
Attribute m_frmLoanReport.VB_VarHelpID = -1
Private m_LoanID As Long
Private m_ParentHandle As Long

Public Property Get ObjectName()
    ObjectName = "DepLoanClass"
End Property


Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset
gDbTrans.SqlStmt = "Select LoanID From DepositLoanMAster " & _
    "Where CustomerID = " & NewValue
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        m_LoanID = rst("LoanID")
Set rst = Nothing
End Property

Public Property Let LoanID(NewValue As Long)
    LoanID = NewValue
End Property

Private Sub LoanDlg_AccountChanged(ByVal AccId As Long)
m_LoanID = AccId
End Sub

Private Sub LoanDlg_ShowReport(ReportType As wis_DepLoanReports, ReportOrder As wis_ReportOrder, _
        DepositType As Integer, fromDate As String, toDate As String, _
        RepOption As clsRepOption)

If Not m_frmLoanReport Is Nothing Then Set m_frmLoanReport = Nothing
Set m_frmLoanReport = New frmDepLoanReport

With m_frmLoanReport
    .ReportOrder = ReportOrder
    .ReportType = ReportType
    .FromAmount = RepOption.FromAmount
    .FromIndainDate = fromDate
    .ToAmount = RepOption.ToAmount
    .ToIndainDate = toDate
    .DepositType = DepositType
    .Gender = RepOption.Gender
    .Place = RepOption.Place
    .Caste = RepOption.Caste
End With

Load m_frmLoanReport

m_frmLoanReport.Refresh
Unload frmCancel
m_frmLoanReport.Show 'vbModal
'Unload m_frmLoanReport
'Set m_frmLoanReport = Nothing

End Sub

Public Property Get ObjectHandle() As Long
    If Not LoanDlg Is Nothing Then _
                ObjectHandle = LoanDlg.hwnd: Exit Property
    If Not m_frmLoanReport Is Nothing Then _
                Set m_frmLoanReport = Nothing
End Property

Private Sub LoanDlg_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle
Set LoanDlg = Nothing

End Sub

Private Sub m_frmLoanReport_Initialise(Min As Long, Max As Long)

frmCancel.Show
On Error Resume Next
With frmCancel
    If Max <> 0 Then
        .PicStatus.Visible = True
        UpdateStatus .PicStatus, 0
        .Refresh
    End If
End With

End Sub


'This Function will Deposit the Amount given to The Speicfied loan account
'Interest till date
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal LoanID As Long, ByVal PrincAmount As Currency, _
                        ByVal IntAmount As Currency, ByVal strDesc As String, _
                        ByVal TransDate As Date, Optional ByVal VoucherNo As String = "", _
                        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

Dim Balance As Currency
Dim IntBalance As Currency
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim transType As wisTransactionTypes
Dim Msg As Boolean

If Not CashTransaction Then
'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
End If

Dim ModuleId As wisModules
Dim DepositType As wis_DepositType

gDbTrans.SqlStmt = "Select DepositType From " & _
    " DepositLoanMaster Where LoanID = " & LoanID
Call gDbTrans.Fetch(rst, adOpenDynamic)
DepositType = FormatField(rst("DepositType"))

'Now Get the Head ID's From The BankHeadIS
Dim ClsBank As clsBankAcc
Dim LoanHeadID As Long
Dim IntHeadID As Long
Dim AccHeadName As String
Dim IntHeadName As String

Set ClsBank = New clsBankAcc
Dim depName As String
Dim depNameEnglish As String

If DepositType >= wisDeposit_PD And DepositType < wisDeposit_PD + 10 Then
    'Loand Head
    depName = GetDepositName(wis_PDAcc, DepositType Mod 10, depNameEnglish)
    AccHeadName = depName & " " & GetResourceString(58)
    'AccHeadName = GetResourceString(425, 58)
    'Deposit Head
    IntHeadName = AccHeadName & " " & GetResourceString(483)
ElseIf DepositType >= wisDeposit_RD And DepositType < wisDeposit_RD + 10 Then
    'Loand Head
    depName = GetDepositName(wis_RDAcc, DepositType Mod 10, depNameEnglish)
    AccHeadName = depName & " " & GetResourceString(58)
    'AccHeadName = GetResourceString(424, 58)
    'Deposit Head
    IntHeadName = AccHeadName & " " & GetResourceString(483)
ElseIf DepositType >= wisDeposit_FD And DepositType < wisDeposit_FD + 10 Then
    gDbTrans.SqlStmt = "Select DepositName From DepositNAme " & _
            " WHERE DepositID = " & DepositType
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    'Loan Head
    AccHeadName = FormatField(rst("DepositName")) & " " & GetResourceString(58)
    'Deposit Head
    IntHeadName = AccHeadName & " " & GetResourceString(483)
End If

LoanHeadID = ClsBank.GetHeadIDCreated(AccHeadName)
IntHeadID = ClsBank.GetHeadIDCreated(IntHeadName)

Set ClsBank = Nothing

gDbTrans.SqlStmt = "Select LoanID, Balance, TransDate, TransID  " & _
    " From DepositLoanTrans Where Balance <> 0 And loanid = " & LoanID & _
    " And TransId = (Select Max(TransId) From DepositLoanTrans  " & _
        " Where LoanID = " & LoanID & ")"

If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo ErrLine
'If Transaction has done later than the given date
If DateDiff("d", TransDate, rst("TransDate")) > 0 Then GoTo ErrLine
'Assign The Necessary Values
TransID = rst("TransID")
Balance = Val(FormatField(rst("Balance")))
IntBalance = 0

'Check for the lAst Date of transCation
gDbTrans.SqlStmt = "SELECT Top 1 * From DepositLoanIntTrans " & _
                "WHERE LoanID = " & LoanID & " ORDER By TransID Desc"
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
    TransID = IIf(TransID >= rst("TransID"), TransID, rst("TransID"))
    IntBalance = FormatField(rst("amount"))
    'If Transaction has done later than the given date
    If DateDiff("d", TransDate, rst("TransDate")) > 0 Then GoTo ErrLine
End If
TransID = TransID + 1

If Balance = PrincAmount Then Msg = False
If IntAmount = 0 Then Msg = False

transType = IIf(CashTransaction, wDeposit, wContraDeposit)
'check for the Pending Balance
If Balance < PrincAmount Then GoTo ErrLine

'before Making TransCtion Calculate the Interest till today
'So to Caluclate the Interest assignt frmloanacc form to a variable
Dim C_IntAmount As Currency

'First load the Loan
'PenalIntAmount = ComputeDepLoanPenalInterest(TransDate) \ 1
C_IntAmount = ComputeDepLoanRegularInterest(TransDate, LoanID) \ 1

If Msg Then
    Dim RetMsg As VbMsgBoxResult
    'The Amount He is payinG less than the Balance interest
    If IntAmount < C_IntAmount + IntBalance Then
        'RetMsg = MsgBox("The amount he is paying is lesser than his interest till date" & _
            "Do you want keep the difference as balance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(668) & _
            GetResourceString(669), vbYesNo + vbQuestion, wis_MESSAGE_TITLE)
        If RetMsg = vbNo Then C_IntAmount = IntAmount - IntBalance
    ElseIf IntAmount > C_IntAmount + IntBalance Then
        'RetMsg = MsgBox("The amount he is paying is more than his interest till date" & _
            "Do you want keep the difference as blance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(666) & _
            GetResourceString(667), vbYesNo + vbQuestion, wis_MESSAGE_TITLE)
        If RetMsg = vbNo Then C_IntAmount = IntAmount - IntBalance
    End If
End If

'Query To Put The Acquired Values into Corresponding Table
'First insert the INterest Amount into loan then Principle
Dim UserID As Long

UserID = gCurrUser.UserID

If IntAmount > 0 Then
    IntBalance = IntBalance + C_IntAmount - IntAmount
    transType = IIf(CashTransaction, wDeposit, wContraDeposit)
    
    gDbTrans.SqlStmt = "INSERT INTO DepositloanIntTrans " & _
        " (Loanid ,TransId ,Amount ,Balance" & _
        ", TransType,TransDate,UserID)  " & _
        " VALUES(" & LoanID & "," & TransID & "," & _
        IntAmount & "," & IntBalance & "," & _
        transType & ", # " & TransDate & "#," & UserID & " )"

    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & IntHeadID & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & IntAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
End If

'Next insert the  principle Amount
transType = IIf(CashTransaction, wDeposit, wContraDeposit)

If PrincAmount > 0 Then
    Balance = Balance - PrincAmount
    gDbTrans.SqlStmt = "INSERT INTO DepositloanTrans " & _
        " (Loanid ,TransId ,Amount ,Balance," & _
        " TransType, Particulars,TransDate,UserId) " & _
        " VALUES(" & LoanID & "," & TransID & "," & _
        PrincAmount & "," & Balance & "," & transType & "," & _
        AddQuotes(strDesc, True) & "," & _
        " #" & TransDate & "#," & UserID & " )"
    
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans" & _
            " (ContraID,AccHeadID,AccId," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & LoanHeadID & "," & _
            LoanID & ", " & transType & ", " & TransID & "," & PrincAmount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
    
End If

DepositAmount = TransID

Exit Function

ErrLine:

End Function

'This Function Returns Interest Balance
Public Function InterestBalance(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntBalance As Currency

IntBalance = 0

Err.Clear
On Error GoTo ExitLine
SqlStr = " SELECT TOp 1 * FROM DepositLoanIntTrans " & _
    " WHERE LoanID = " & LoanID & " AND TransDate <= #" & AsOnDate & "# " & _
    " ORDER BY TransId Desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then _
    IntBalance = FormatField(rst("IntBalance"))

ExitLine:
    If Err Then
        MsgBox "ERROR InterestBalance", vbInformation, wis_MESSAGE_TITLE
        'resume
        Err.Clear
    End If
    If IntBalance < 0 Then IntBalance = 0
    InterestBalance = IntBalance

End Function

Public Function RegularInterest(LoanID As Long, AsOnDate As Date) As Currency
' Setup error handler.
'If Not DateValidate(AsOnIndianDate, "/", True) Then Exit Function
'Dim AsOnDate As Date
'AsOnDate = FormatDate(AsOnIndianDate)

On Error GoTo Err_line

' Define varaibles.
Dim Lret As Long
Dim IntDiff As Single
Dim ActualIntRate As Single
Dim Balance As Long
Dim LastintPaidDate As Date
Dim Duration As Long
Dim IntRate As Single
Dim IntAmount As Currency
Dim rstLoanMast As Recordset
Dim RstLoanTrans As Recordset
Dim Days As Integer
Dim rst As ADODB.Recordset

gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE LoanID = " & LoanID
Lret = gDbTrans.Fetch(rstLoanMast, adOpenStatic)
If Lret <= 0 Then GoTo Exit_Line
' Save the resultset for future references.

IntRate = Val(FormatField(rstLoanMast("InterestRate")))

If FormatField(rstLoanMast("LastIntdate")) = "" Then
    LastintPaidDate = Now
    gDbTrans.SqlStmt = "SELECT * From DepositLoanIntTrans WHERE " & _
            " LoanID = " & LoanID & " AND TransDate <= #" & AsOnDate & "#" & _
            " ORDER BY TransID Desc "
    If gDbTrans.Fetch(RstLoanTrans, adOpenStatic) Then
        LastintPaidDate = RstLoanTrans("TransDate")
    Else
        gDbTrans.SqlStmt = "SELECT * From DepositLoanTrans WHERE " & _
            " LoanID = " & LoanID & " ORDER BY TransID "
        If gDbTrans.Fetch(rst, adOpenForwardOnly) Then LastintPaidDate = rst("TransDate")
    End If
Else
    LastintPaidDate = rstLoanMast("LastIntdate")
End If

    
'Now Get the TransCtion Of the LastIntPaid date
gDbTrans.SqlStmt = "SELECT * FROM DepositLoanTrans Where LoanID = " & LoanID & _
    " AND TransDate >= #" & LastintPaidDate & "# ORDER BY TransID"
If gDbTrans.Fetch(RstLoanTrans, adOpenStatic) < 1 Then GoTo Exit_Line
IntAmount = 0
Balance = RstLoanTrans("Balance")
Dim TillDate As Date

Do
    If RstLoanTrans.EOF Then
        Duration = DateDiff("d", LastintPaidDate, AsOnDate)
        IntAmount = IntAmount + ((Duration / 365) * (IntRate / 100) * Balance)
        Exit Do
    End If
    TillDate = RstLoanTrans("TransDate")
    Duration = DateDiff("d", LastintPaidDate, TillDate)
    IntAmount = IntAmount + ((Duration / 365) * (IntRate / 100) * Balance)
    LastintPaidDate = TillDate
    Balance = FormatField(RstLoanTrans("BALANCE"))
    RstLoanTrans.MoveNext
Loop

    RegularInterest = IntAmount \ 1

    
Exit_Line:
    Set rst = Nothing
    Set rstLoanMast = Nothing
    Set RstLoanTrans = Nothing
    RegularInterest = IntAmount \ 1
    Exit Function

Err_line:
    If Err Then
        MsgBox "ComputeRegularInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
     End If
End Function

' This Function returns the next Loan Balance of
'the customer whose first account details already fetched
'if the same customer already gor the any account His Details will fetched
Public Function NextBalance(retDepositType As Long, retLoanNum As String) As Currency
    
    NextBalance = 0
    retDepositType = 0
    retLoanNum = ""
    
    If m_rstCustBalance Is Nothing Then Exit Function
    If m_rstCustBalance.EOF Then Exit Function
    'NOw move the cursor to the next record Record
    m_rstCustBalance.MoveNext
    If m_rstCustBalance.EOF Then Exit Function
    
    retLoanNum = FormatField(m_rstCustBalance("AccNum"))
    retDepositType = FormatField(m_rstCustBalance("DepositType"))
    NextBalance = FormatField(m_rstCustBalance("Balance"))
    
End Function
' This Function returns the Loan Balance  , the customer Having if
' It also returns the Loan Acc No
Public Function CustomerBalance(ByVal CustomerID As Long, ByVal AsOnDate As Date, _
    retDepositType As Long, retLoanNum As String) As Currency
    
    CustomerBalance = 0
    Set m_rstCustBalance = Nothing
    
    gDbTrans.SqlStmt = "Select A.LoanID,AccNum,DepositType,Balance from " & _
        " DepositLoanMaster A, DepositLoanTrans B Where A.LoanID = B.LoanID " & _
        " And CustomerId = " & CustomerID & " AND TransID = (SELECT Max(TransID) " & _
            " From DepositLoanTrans C WHERE C.LoanID = B.LoanID AND " & _
            " TransDate <= #" & AsOnDate & "#) " & _
        " AND Balance > 0 "
    
    If gDbTrans.Fetch(m_rstCustBalance, adOpenForwardOnly) < 1 Then Exit Function
    
    retLoanNum = FormatField(m_rstCustBalance("AccNum"))
    retDepositType = FormatField(m_rstCustBalance("DepositType"))
    CustomerBalance = FormatField(m_rstCustBalance("Balance"))
    
End Function

Private Function LoanTransAmount(ByVal fromDate As Date, ByVal toDate As Date, _
        ByVal DepositType As Integer, ByVal transType As wisTransactionTypes) As Currency
Dim SqlStr As String
Dim rst As Recordset

LoanTransAmount = 0

SqlStr = "Select LoanId From DepositLoanMaster " & _
    " Where DepositType = " & DepositType
gDbTrans.SqlStmt = SqlStr: SqlStr = ""
If Not gDbTrans.CreateView("QryTemp") Then Exit Function

SqlStr = "Select Sum(Amount) From DepositLoanTrans A, QryTemp B " & _
    " Where A.LoanID = B.LoanID" & _
    " And TransDate >= #" & fromDate & "# " & _
    " And Transdate <= #" & toDate & "# " & _
    " And TransType = " & transType

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
If Not rst Is Nothing Then LoanTransAmount = Val(FormatField(rst(0)))

End Function
Public Sub TotalLoanTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef ContraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef ContraWithDraw As Currency, Optional DepositType As Integer)

CashDeposit = 0
cashWithdraw = 0
ContraDeposit = 0
ContraWithDraw = 0
Dim subDepType As Integer


Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from DepositLoanTrans where "

If DepositType > 0 Then
    SqlStr = SqlStr & " LoanID In " & _
        "(select distinct(LoanID) From DepositLoanMaster where DepositType = " & DepositType & ") And "
End If

SqlStr = SqlStr & " TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#" & _
                   " Group By TransType"


gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim TotalAmount As Currency
While rst.EOF = False
    TotalAmount = Val(FormatField(rst("TotalAmount")))
    If rst("TransType") = wDeposit Then CashDeposit = TotalAmount
    If rst("TransType") = wWithdraw Then cashWithdraw = TotalAmount
    If rst("TransType") = wContraDeposit Then ContraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then ContraWithDraw = TotalAmount
    
    rst.MoveNext
Wend

Set rst = Nothing
End Sub
Public Sub TotalLoanIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef ContraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef ContraWithDraw As Currency, Optional DepositType As Integer = 0)

CashDeposit = 0
cashWithdraw = 0
ContraDeposit = 0
ContraWithDraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from DepositLoanTrans where "

If DepositType > 0 Then
    SqlStr = SqlStr & " LoanID In " & _
        "(select distinct(loanID) From DepositLoanMaster where DepositType = " & DepositType & ") And "
End If

SqlStr = SqlStr & " TransDate >= #" & fromDate & "#" & _
                   " And TransDate <= #" & toDate & "#" & _
                   " Group By TransType"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim TotalAmount As Currency

While rst.EOF = False
    TotalAmount = Val(FormatField(rst("TotalAmount")))
    If rst("TransType") = wDeposit Then CashDeposit = TotalAmount
    If rst("TransType") = wWithdraw Then cashWithdraw = TotalAmount
    If rst("TransType") = wContraDeposit Then ContraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then ContraWithDraw = TotalAmount
    
    rst.MoveNext
Wend

Set rst = Nothing
End Sub

'
Private Function ProfitLossTransAmount(ByVal fromDate As Date, ByVal toDate As Date, _
        ByVal DepositType As Integer, ByVal transType As wisTransactionTypes) As Currency
Dim SqlStr As String
Dim rst As Recordset

ProfitLossTransAmount = 0

SqlStr = "Select LoanId From DepositLoanMaster " & _
    " Where DepositType = " & DepositType
If Not gDbTrans.CreateView("QryTemp") Then Exit Function

SqlStr = "Select Sum(Amount) From DepositLoanIntTrans A, QryTemp B " & _
    " Where A.LoanID = B.LoanID" & _
    " And TransDate >= #" & fromDate & "# " & _
    " And Transdate <= #" & toDate & "# " & _
    " And TransType = " & transType

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
If Not rst Is Nothing Then ProfitLossTransAmount = Val(FormatField(rst(0)))

End Function

Public Function LoanBalance(ByVal AsonIndianDate As String, ByVal SchemeID As Long) As Currency
' Setup error handler.
On Error GoTo Err_line

' Declare variables...
Dim Lret As Long
Dim rst As Recordset
Dim Balance As Currency
' Prepare the SQL query for getting loan balance of the given scheme.
gDbTrans.SqlStmt = "SELECT loanid FROM loanmaster WHERE schemeid = " & SchemeID

gDbTrans.SqlStmt = "SELECT A.LoanID Balance FROM LoanTrans A, LoanMaster B WHERE B.SchemeId = " & SchemeID & _
    " And A.LoanId in (Select LoanId From LoanMAster where  SchemeId = " & SchemeID & ")" & _
    " And A.TransId = (Select max(TransID) from LoanTrans C " & _
    " where C.LoanId = A.LoanID And TransDate <= #" & AsonIndianDate & "#)"

gDbTrans.SqlStmt = "SELECT A.LoanID, A.Balance FROM LoanTrans A WHERE Balance >= 1 " & _
    " And LoanId in (Select LoanId From LoanMAster where  SchemeId = " & SchemeID & ")" & _
    " And A.TransId = (Select max(TransID) from LoanTrans C " & _
    " where C.LoanId = A.LoanID And TransDate <= #" & AsonIndianDate & "#)"


Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
DoEvents

' For each of the above loanid's get the balance amt as on date.
    Balance = 0
    Do While Not rst.EOF
        ' Get the balance for each  loanid.
        Balance = Balance + Val(FormatField(rst("Balance")))
        'Move to next rec.
        rst.MoveNext
    Loop

LoanBalance = Balance
Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "LoanBalance: " & vbCrLf _
            & Err.Description, vbCritical
    End If

    LoanBalance = 0
    GoTo Exit_Line

End Function
Public Function TotalBalance(ByVal AsOnDate As Date) As Currency

' Setup error handler.
Err.Clear
On Error GoTo Exit_Line

' Declare variables...
Dim Lret As Long
Dim rst As Recordset
Dim Balance As Currency

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM DepositLoanTrans A WHERE Balance < 0" & _
    " And A.TransId = (Select max(TransID) from DepositLoanTrans B" & _
    " Where B.LoanId = A.LoanID And TransDate <= #" & AsOnDate & "#" & _
    " And Balance < 0)"
    
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
TotalBalance = FormatField(rst(0))
Set rst = Nothing

Exit_Line:
    If Err Then
        MsgBox "DepLoanBalance: " & vbCrLf _
            & Err.Description, vbCritical
    End If
End Function

'Reverts the last transaction in loan transactions table.
Public Function UndoTransaction(LoanID As Long, TransID As Long) As Boolean

' Variables of the procedure...
Dim lastTransID As Long
Dim rst As ADODB.Recordset
Dim IntAmount As Currency
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes

' Setup the error handler.
On Error GoTo Err_line

' Get the last transaction from the loan transaction table.
' Fetch the last Record.
gDbTrans.SqlStmt = "SELECT Top 1 TransID,TransType,Amount,TransDate,Balance " & _
    " FROM DepositLoanTrans WHERE loanid = " & LoanID & _
    " ORDER BY TransID DESC"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "There are no transaction to undo.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(661), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If


' Check the transaction types of the records to undo.
lastTransID = Val(FormatField(rst("TransID")))
transType = Val(FormatField(rst("TransType")))
TransDate = rst("TransDate")
Amount = FormatField(rst("Amount"))

'Now Check For THe LastTransaction in Intreste Trans
' Fetch the last Record.
gDbTrans.SqlStmt = "SELECT Top 1 TransID,Transtype,Amount, " & _
    " TransDate FROM DepositLoanIntTrans WHERE loanid = " & LoanID & _
    " ORDER BY TransID DESC"

Dim rstTemp As ADODB.Recordset
If gDbTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then
    If rstTemp("TransID") >= lastTransID Then
        If rstTemp("TransID") > lastTransID Then Amount = 0
        lastTransID = rstTemp("TransID")
        IntAmount = FormatField(rstTemp("amount"))
        TransDate = rstTemp("TransDate")
    End If
End If

If lastTransID <> TransID Then
    MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'First Delete the interest transaction then
gDbTrans.SqlStmt = "DELETE FROM DepositLoanIntTrans WHERE TransID = " & lastTransID _
        & " AND loanid = " & LoanID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

' Delete the entry of principle amount from the transaction table.
gDbTrans.SqlStmt = "DELETE FROM DepositLoanTrans WHERE TransID = " & lastTransID _
        & " AND loanid = " & LoanID
If Not gDbTrans.SQLExecute Then
    MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'UpDate The LoanMAster
gDbTrans.SqlStmt = "UPDATE Depositloanmaster SET " & _
    " LoanClosed = 1 Where Loanid =  " & LoanID
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

gDbTrans.SqlStmt = "Select Max(TransDate) From depositLoanIntTrans " & _
            " Where Loanid =  " & LoanID
Call gDbTrans.Fetch(rstTemp, adOpenDynamic)
        
gDbTrans.SqlStmt = "UPDATE Depositloanmaster SET " & _
    " LastIntDate = " & IIf(IsNull(rstTemp(0)), "NULL", "#" & rstTemp(0) & "#") & _
    " Where Loanid =  " & LoanID
        
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

UndoTransaction = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "DepLoanClassUndoTransaction: " & vbCrLf _
                '& Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(731) & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    Else
        MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function

Public Sub Show()
    
    If LoanDlg Is Nothing Then Set LoanDlg = New frmDepLoan
    If gWindowHandle Then m_ParentHandle = gWindowHandle
    Load LoanDlg
    gWindowHandle = LoanDlg.hwnd
    LoanDlg.Show 'vbModal
    If m_LoanID Then LoanDlg.LoanLoad (m_LoanID)
    
End Sub

Public Function WithdrawAmount(LoanID As Long, Amount As Currency, _
        strDesc As String, TransDate As Date, _
        Optional VoucherNo As String, _
        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

On Error GoTo Exit_Line

Dim TransID As Long
Dim transType As wisTransactionTypes
Dim LastTransDate As Date
Dim Balance As Currency
Dim rst As Recordset

Dim DepositType As wis_DepositType
gDbTrans.SqlStmt = "Select DepositType From " & _
    " DepositLoanMaster Where LoanID = " & LoanID
Call gDbTrans.Fetch(rst, adOpenDynamic)
DepositType = FormatField(rst("DepositType"))

' Get a new transactionID.
gDbTrans.SqlStmt = "SELECT top 1 Balance,TransID,TransDate FROM " & _
    " DepositLoanTrans WHERE loanid = " & LoanID & " ORDER BY TransId Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then GoTo ExitLine

    Balance = Val(FormatField(rst("Balance")))
    TransID = Val(FormatField(rst("TransID")))
    LastTransDate = rst("TransDate")

'Now Compare the Last Date Of Transaction with transction date
If DateDiff("d", TransDate, LastTransDate) > 0 Then
    'Date Trasnaction should be later
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Get the Transaction Id From The Interest table
gDbTrans.SqlStmt = "SELECT TransID,TransDate FROM DepositLoanIntTrans " & _
    " WHERE loanid = " & LoanID & " ORDER BY TransId Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    LastTransDate = rst("TransDate")
    'Now Compare the Last Date Of Transaction with Transction date
    If DateDiff("d", TransDate, LastTransDate) > 0 Then
        'Date Trasnaction should be later
        MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
End If

transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Update the Regular & Penal interest amount.
TransID = TransID + 1

Dim NewBalance As Currency
Dim LoanHeadID As Long

If Not CashTransaction Then
    Dim AccHeadName As String
    Dim depName As String
    'Loand Head
    If DepositType >= wisDeposit_PD And DepositType < wisDeposit_PD + 10 Then
        depName = GetDepositName(wis_PDAcc, DepositType Mod 10)
        AccHeadName = depName & " " & GetResourceString(58) ' GetResourceString(425, 58)
    ElseIf DepositType >= wisDeposit_RD And DepositType < wisDeposit_RD + 10 Then
        depName = GetDepositName(wis_RDAcc, DepositType Mod 10)
        AccHeadName = depName & " " & GetResourceString(58) ' AccHeadName = GetResourceString(424, 58)
    ElseIf DepositType >= wisDeposit_FD Then
        gDbTrans.SqlStmt = "Select DepositName From DepositNAme " & _
                " WHERE DepositID = " & DepositType
        Call gDbTrans.Fetch(rst, adOpenDynamic)
        'Loan Head
        AccHeadName = FormatField(rst("DepositName")) & " " & GetResourceString(58)
    End If

    ''Get the head ID for the Ledger entry
    AccHeadName = FormatField(rst("SchemeName"))
    LoanHeadID = GetIndexHeadID(AccHeadName)

    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraTransID = 0 Then ContraID = GetMaxContraTransID '+ 1
End If

'PrincAmount = Amount
NewBalance = Balance + Amount

gDbTrans.SqlStmt = "INSERT INTO DepositLoanTrans (LoanID, TransID, " _
    & "TransType, Amount, TransDate, Balance, Particulars,UserID) " _
    & "VALUES (" & LoanID & ", " & TransID & ", " _
    & transType & ", " & Amount & ", " _
    & "#" & TransDate & "#, " _
    & NewBalance & "," _
    & AddQuotes(strDesc, True) _
    & "," & gUserID & ")"
' Execute the updation.
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

If Not CashTransaction Then
    'Update it to the Contra Trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccID," & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & _
        LoanHeadID & "," & _
        LoanID & ", " & transType & ", " & TransID & "," & Amount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gCurrUser.UserID & ")"
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


Public Function WithDrawls(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType As wisTransactionTypes
    transType = wWithdraw
    WithDrawls = LoanTransAmount(fromDate, toDate, DepositType, transType)
    
End Function

'This function returns the of the Amount withdrawn from the
'specified depositLoan between the specified period
'this includes only the internal transfer
Public Function ContraWithdrawls(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType  As wisTransactionTypes
    transType = wContraWithdraw
    
    ContraWithdrawls = LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)

End Function


'This function returns the of the Amount withdrawn from the
'specified depositLoan between the specified period
'this includes only cash transaction
Public Function CashWithdrawls(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType  As wisTransactionTypes
    transType = wWithdraw
    CashWithdrawls = LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)

End Function


'This function returns the of the Amoutn deposited to the
'specified depositLoan between the specified period
'this includes only cash transaction
Public Function CashDeposits(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType  As wisTransactionTypes
    Dim Amount As Currency
    transType = wDeposit
    Amount = LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)
    
    CashDeposits = Amount

End Function


'This function returns the of the Amoutn deposited to the
'specified depositLoan between the specified period
'this amount returns only interenal transfer
Public Function ContraDeposits(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType  As wisTransactionTypes
    Dim Amount As Currency
    
    transType = wContraDeposit
    Amount = LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)
    
    ContraDeposits = Amount

End Function

'This function returns the of the Amoutn deposited to the
'specified depositLoan between the specified period
'this amount includes tha cash and any interenal transfer
Public Function Deposits(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
    
    Dim transType  As wisTransactionTypes
    Dim Amount As Currency
    transType = wDeposit
    Amount = LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)
    transType = wContraDeposit
    Amount = Amount + LoanTransAmount(fromDate, toDate, CInt(DepositType), transType)
    
    Deposits = Amount

End Function

'This function returns the of the Amount earned by bank /society
'from specified depositLoan between the specified period
'this amount includes only cash transaction
Public Function CashProfit(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wDeposit
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        CashProfit = Amount
End Function

'This function returns the of the Amount earned by bank /society
'from specified depositLoan between the specified period
'this amount includes only interenal transfer
Public Function ContraProfit(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wContraDeposit
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        ContraProfit = Amount
End Function


'This function returns the of the Amount earned by bank /society
'from specified depositLoan between the specified period
'this amount includes tha cash and any interenal transfer
Public Function Profit(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wDeposit
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        transType = wContraDeposit
        Amount = Amount + ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        Profit = Amount
End Function

'This function returns the Amount given as loss from bank /society
'from specified depositLoan between the specified period
'this amount includes tha cash transaction
Public Function CashLoss(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wWithdraw
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        CashLoss = Amount
End Function


'This function returns the Amount given as loss from bank /society
'from specified depositLoan between the specified period
'this amount includes tha only interenal transfer
Public Function ContraLoss(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wContraWithdraw
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        ContraLoss = Amount
End Function


'This function returns the Amount given as loss from bank /society
'from specified depositLoan between the specified period
'this amount includes tha cash and any interenal transfer
Public Function Loss(fromDate As Date, toDate As Date, DepositType As Integer) As Currency
        Dim transType As wisTransactionTypes
        Dim Amount As Currency

        transType = wWithdraw
        Amount = ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        transType = wContraWithdraw
        Amount = Amount + ProfitLossTransAmount(fromDate, toDate, DepositType, transType)
        Loss = Amount
End Function


Private Sub Class_Terminate()
On Error Resume Next
If Not LoanDlg Is Nothing Then
    Unload LoanDlg
    Set LoanDlg = Nothing
End If
    
    
End Sub

Private Sub m_frmLoanReport_Processing(strMessage As String, Ratio As Single)
If Ratio > 1 Then Ratio = 1
With frmCancel
    .lblMessage = strMessage
    UpdateStatus .PicStatus, Ratio
    
End With
End Sub

