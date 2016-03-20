VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_rstCustBalance As Recordset

Private WithEvents m_frmSelectReport As frmLoanReport
Attribute m_frmSelectReport.VB_VarHelpID = -1
Private WithEvents m_frmRegReport As frmRegReport
Attribute m_frmRegReport.VB_VarHelpID = -1
Private WithEvents m_frmLoanReport As frmLoanView
Attribute m_frmLoanReport.VB_VarHelpID = -1

Private WithEvents m_frmScheme As frmLoanScheme
Attribute m_frmScheme.VB_VarHelpID = -1
Private WithEvents m_frmLoanMaster As frmLoanMaster
Attribute m_frmLoanMaster.VB_VarHelpID = -1
Private WithEvents m_frmLoanTrans As frmLoanTrans
Attribute m_frmLoanTrans.VB_VarHelpID = -1
Private WithEvents m_frmLoanInst As frmLoanInst
Attribute m_frmLoanInst.VB_VarHelpID = -1
Private WithEvents m_frmLoanPay As frmLoanPay
Attribute m_frmLoanPay.VB_VarHelpID = -1

Private m_ParentHandle As Long

Private m_DetailHandle As Long
Private m_TransHandle As Long
Private m_InstHandle As Long
Private m_ReportHandle As Long

Private m_LoanID As Long
Private m_SchemeId As Integer
Public Event CurrentEvent(Message As String)

'
Public Function CustomerBalance(ByVal CustomerID As Long, ByVal AsOnDate As Date, _
        retLoanNum As String, retLoanName As String) As Currency

' This Function returns the Loan Balance  , the customer Having if
' It also returns the Loan Acc No
    Dim LoanID As Long

    CustomerBalance = 0
    gDbTrans.SqlStmt = "Select A.LoanID,AccNum,Balance,SchemeName " & _
        " From LoanMaster A, LoanTrans B, LoanScheme C Where " & _
        " CustomerId = " & CustomerID & " And A.LoanId = B.LoanId " & _
        " AND TransID = (Select Max(transID) from LoanTrans D Where " & _
            " D.LoanId = A.LoanID And Transdate <= #" & AsOnDate & "#)" & _
        " ANd C.SchemeId = A.SchemeId AND Balance > 0 "
    
    If gDbTrans.Fetch(m_rstCustBalance, adOpenDynamic) < 1 Then Exit Function
        
    CustomerBalance = FormatField(m_rstCustBalance("Balance"))
    retLoanName = FormatField(m_rstCustBalance("SchemeName"))
    retLoanNum = FormatField(m_rstCustBalance("AccNum"))

End Function


Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset
gDbTrans.SqlStmt = "select * From LoanMaster " & _
    "Where CustomerId = " & NewValue
If m_SchemeId Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ANd SchemeID = " & m_SchemeId

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order by LoanId desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    If m_LoanID Then
        rst.Find "LoanID = " & m_LoanID
        If rst.EOF Then  'no match found then
            rst.MoveFirst
            m_LoanID = rst("LoanId")
        End If
    Else
        m_LoanID = FormatField(rst("LoanID"))
    End If
End If


Set rst = Nothing

End Property

Public Property Let LoanID(NewValue As Long)
    m_LoanID = NewValue
End Property

'This function returns the name of the Loan scheme
'Arguments SchemeID
Public Function LoanName(ByVal SchemeID As Integer) As String
Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
        " WHERE SchemeID = " & SchemeID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then LoanName = FormatField(rst(0))
End Function

Public Function NextBalance(retLoanNum As String, retLoanName As String)
    If m_rstCustBalance Is Nothing Then Exit Function
    If m_rstCustBalance.EOF Then Exit Function
    m_rstCustBalance.MoveNext
    If m_rstCustBalance.EOF Then Exit Function
    
    NextBalance = FormatField(m_rstCustBalance("Balance"))
    retLoanName = FormatField(m_rstCustBalance("SchemeName"))
    retLoanNum = FormatField(m_rstCustBalance("AccNum"))

End Function


'
'This Function Returns Interest Balance
Public Function RegInterestBalance(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntBalance As Currency

IntBalance = 0

Err.Clear
On Error GoTo ExitLine
SqlStr = "SELECT IntBalance FROM LoanIntTrans WHERE LoanID = " & LoanID & _
        " AND TransID =(SELECT MAX(TransID) FROM LoanIntTrans " & _
            " WHERE LoanID = " & LoanID & _
            " AND TransDate <= #" & AsOnDate & "#)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then IntBalance = FormatField(rst("IntBalance"))

ExitLine:
    If Err Then
        MsgBox "ERROR InterestBalance", vbInformation, wis_MESSAGE_TITLE
        'Resume
        Debug.Assert Err.Number
        Err.Clear
    End If
    If IntBalance < 0 Then IntBalance = 0
    
    RegInterestBalance = IntBalance
End Function

'
'This Function Returns Interest Balance
Public Function PenalInterestBalance(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntBalance As Currency

IntBalance = 0

Err.Clear
On Error GoTo ExitLine
SqlStr = "SELECT PenalIntBalance FROM LoanIntTrans WHERE LoanID = " & LoanID & _
        " AND TransID =(SELECT MAX(TransID) FROM LoanIntTrans " & _
            " WHERE LoanID = " & LoanID & _
            " AND TransDate <= #" & AsOnDate & "#)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    IntBalance = FormatField(rst("PenalIntBalance"))
End If

ExitLine:
    If Err Then
        MsgBox "ERROR InterestBalance", vbInformation, wis_MESSAGE_TITLE
        'resume
        Err.Clear
    End If
    If IntBalance < 0 Then IntBalance = 0
    
    PenalInterestBalance = IntBalance

End Function

'
'This Function Returns Interest Balance
Public Function InterestBalance(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntBalance As Currency

IntBalance = 0

Err.Clear
On Error GoTo ExitLine
SqlStr = "SELECT IntBalance,PenalIntBalance FROM LoanIntTrans WHERE LoanID = " & LoanID & _
        " AND TransID =(SELECT MAX(TransID) FROM LoanIntTrans " & _
            " WHERE LoanID = " & LoanID & _
            " AND TransDate <= #" & AsOnDate & "#)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    IntBalance = Val(FormatField(rst("IntBalance"))) + Val(FormatField(rst("PenalIntBalance")))
End If

ExitLine:
    If Err Then
        MsgBox "ERROR InterestBalance", vbInformation, wis_MESSAGE_TITLE
        'resume
        Err.Clear
    End If
    If IntBalance < 0 Then IntBalance = 0
    
    InterestBalance = IntBalance

End Function

'This function returns the Amount of loan Suppose to be recovered
'Between the speicified date
'Irrespeictive of the Recovered amount
'If he has paid the the amount then also that amount lies in a period
Public Function SchemeLoanDemand(ByVal SchemeID As Integer, _
            ByVal fromDate As Date, ByVal toDate As Date) As Currency
MsgBox "Incomplete"
Exit Function

SchemeLoanDemand = 0

Dim SqlStr As String
Dim rst As Recordset
Dim DemAmount As Currency
Dim opBalance As Currency

Dim Duration As Long
Dim InstType As wisInstallmentTypes

Dim DueDate As Date

Err.Clear
On Error GoTo ExitLine
DemAmount = 0
    

'Get the deatail of the particular loan
'SqlStr = "SELECT * FROM LoanMaster" & _
                " Where LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
InstType = FormatField(rst("InstMode"))
DueDate = rst("LoanDueDate")

'Get The Loan opening balance as on fromdate of the Particular loan account
'SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate < #" & FromDate & "# ORDER BY TransID Desc"
    
SqlStr = "SELECT Sum(A.Balance) as OpBalance " & _
    " FROM LoanTrans A, LoanTrans B,Where " & _
    " A.LoanID = B.LoanID" & _
    " A.LoanID In (Select LoanID From LoanMaster Where " & _
            " LoanDueDate is Between #" & fromDate & "# AND #" & toDate & "# )" & _
    " A.TransId = (SELECT Max(TransID) FROM LoanTrans Where " & _
            " LoanId = A.LoanID AND TransDate < #" & fromDate & "#)"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    opBalance = 0
Else
    opBalance = FormatField(rst("opBalance"))
End If

'Get The Last Balance and TransDate of the loan
'SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate <= #" & ToDate & "# ORDER BY TransID Desc"

'SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate <= #" & ToDate & "# ORDER BY TransID Desc"

gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine

If DueDate >= fromDate And DueDate <= toDate Then
    DemAmount = opBalance
    GoTo ExitLine
End If

'Check for The Instalment type
If InstType = Inst_No Then GoTo ExitLine

If InstType <> Inst_No Then
    'Get The Deatils form Instalment table
    'SqlStr = "SELECT SUM(InstAmount) FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstDate >= #" & FromDate & "#" & _
        " AND InstDate <= #" & ToDate & "#"
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
    DemAmount = 0
    DemAmount = Val(FormatField(rst(0)))
End If

ExitLine:
    If Err Then
        MsgBox "ERROR In Loan demand" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    SchemeLoanDemand = DemAmount

End Function

'this function returns the amount of loan Suppose to recovered
'Between the speicified date
'Irrespeictive of the Recored amount
'If he has paid the the amount then also that amount lies in a period
Public Function LoanDemand(ByVal LoanID As Long, _
            ByVal fromDate As Date, ByVal toDate As Date) As Currency
LoanDemand = 0

Dim SqlStr As String
Dim rst As Recordset
Dim DemAmount As Currency
Dim opBalance As Currency

Dim Duration As Long
Dim InstType As wisInstallmentTypes

Dim DueDate As Date

Err.Clear
On Error GoTo ExitLine
DemAmount = 0
    

'Get the deatail of the particular loan
SqlStr = "SELECT * FROM LoanMaster" & _
                " Where LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
InstType = FormatField(rst("InstMode"))
DueDate = rst("LoanDueDate")

'Get The Loan opening balance as on fromdate of the Particular loan account
SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate < #" & fromDate & "# ORDER BY TransID Desc"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    opBalance = 0
Else
    opBalance = FormatField(rst("Balance"))
End If

'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
    " LoanId = " & LoanID & _
    " AND TransDate <= #" & toDate & "# ORDER BY TransID Desc"
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine

If DueDate >= fromDate And DueDate <= toDate Then
    DemAmount = opBalance
    GoTo ExitLine
End If

'Check for The Instalment type
If InstType = Inst_No Then GoTo ExitLine

If InstType <> Inst_No Then
    'Get The Deatils form Instalment table
    SqlStr = "SELECT SUM(InstAmount) FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstDate >= #" & fromDate & "#" & _
        " AND InstDate <= #" & toDate & "#"
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
    DemAmount = 0
    DemAmount = Val(FormatField(rst(0)))
End If

ExitLine:
    If Err Then
        MsgBox "ERROR In Loan demand" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    LoanDemand = DemAmount

End Function

'this function returns the Over due amount of loan as on given date
'If fromdate is also given then
'it returns the overdue amount during the specified period
Private Function LoanOverDue(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
LoanOverDue = 0
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim Duration As Long
Dim InstType As wisInstallmentTypes

Dim DueDate As Date

Dim Balance As Currency

Err.Clear
On Error GoTo Exit_Line
ODAmount = 0

'Get The Loan over due of the Particular loan account
SqlStr = "SELECT * FROM LoanMaster Where LoanId = " & LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
InstType = FormatField(rst("InstMode"))
DueDate = rst("LoanDueDate")

'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
        " LoanId = " & LoanID & " AND TransDate <= #" & AsOnDate & "#" & _
        " ORDER BY TransID Desc"
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
'Set rst = gDBTrans.rst.Clone
'The Loan Balance as on Date
Balance = FormatField(rst("Balance"))

Duration = DateDiff("d", DueDate, AsOnDate)
If Duration > 0 Then
    ODAmount = Balance
    GoTo Exit_Line
End If

'Check for The Instalment type
If InstType = Inst_No Then GoTo Exit_Line

If InstType <> Inst_No Then
    
    'Get the Loan Advanced and Loan repaid upto this date
    Dim curAdvance As Currency
    Dim curRepay As Currency
    SqlStr = "SELECT SUM(AMOUNT),TransType FROM LoanTrans WHERE" & _
        " LoanId = " & LoanID & _
        " AND TransDate <= #" & AsOnDate & "# GROUP By TransType"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    'Set rst = gDBTrans.rst.Clone
    While Not rst.EOF
        If FormatField(rst("TransType")) = wContraDeposit Or FormatField(rst("TransType")) = wDeposit Then
            curRepay = curRepay + FormatField(rst(0))
        ElseIf FormatField(rst("TransType")) = wContraWithdraw Or FormatField(rst("TransType")) = wWithdraw Then
            curAdvance = curAdvance + FormatField(rst(0))
        End If
        rst.MoveNext
    Wend
    Debug.Assert Balance = curAdvance - curRepay  'Inform To Shashi Angadi Immedietly
    Balance = 0
    Balance = curAdvance - curRepay
    'Get The total Instalment amount supposed
    'to be paid till on specified date
    Dim NetInstAmount As Currency
    SqlStr = "SELECT SUM(InstAmount) FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstDate < #" & AsOnDate & "#"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    ODAmount = 0
    NetInstAmount = Val(FormatField(rst(0)))
    
    'The positive Difference between the netInstAmount and amount repaid
    'is the over due as on this date
    ODAmount = NetInstAmount - curRepay
    'If he has paid some installments in advance
    'Then it shows the negative Over due 'So
    If ODAmount < 0 Then ODAmount = 0
End If

Exit_Line:
    If Err Then
        MsgBox "ERROR In Loan over due" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    LoanOverDue = ODAmount \ 1
End Function
Public Function RepayLoan(ByVal LoanID As Long, ByVal TransDate As Date, _
        ByVal PrincAmount As Currency, ByVal Particulars As String, _
        Optional ByVal IntAmount As Currency, _
        Optional ByVal PenalAmount As Currency, Optional IntBalance As Currency, _
        Optional ByVal IntPaidUpToDate As Date, Optional ByVal MiscAmount As Double, _
        Optional ByVal CONTRA As Boolean = False) As Boolean
Dim SqlStr As String
Dim rst As Recordset
Dim RstInst As Recordset
Dim Balance As Double
Dim LastTransDate As Date

Dim transType As wisTransactionTypes
Dim TransID As Long
'TransDate = FormatDate(IndianTransDate)
SqlStr = "SELECT Top 1 * FROM LoanTrans WHERE LoanID = " & LoanID & _
    " ORDER BY TransId Desc"
gDbTrans.SqlStmt = SqlStr

TransID = gDbTrans.Fetch(rst, adOpenDynamic)
If TransID < 0 Then Exit Function
If TransID > 0 Then
    LastTransDate = rst("TransDate")
    If DateDiff("d", LastTransDate, TransDate) < 0 Then
        Err.Raise 5102, "Loan Class", "Invalid Date"
        Exit Function
    End If
    Balance = FormatField(rst("Balance"))
    TransID = FormatField(rst("TransID"))
    'Now get the Max Trans Id form Interest Transaction
    SqlStr = "SELECT Top 1 * FROM LoanIntTrans WHERE LoanID = " & LoanID & _
            " ORDER BY TransId Desc"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then _
        TransID = IIf(TransID < rst("TransID"), rst("TransID"), TransID)
End If

If Balance < PrincAmount Then
    'MsgBox "Total amount is greater than LoanAmount.", vbInformation
    MsgBox GetResourceString(585), vbInformation, wis_MESSAGE_TITLE
    Exit Function
    'Err.Raise 5103, "LoanRepay", "Amount not matching"
End If

Balance = Balance - PrincAmount
TransID = TransID + 1
transType = IIf(CONTRA, wContraDeposit, wDeposit)
SqlStr = "SELECT * From LoanInst Where InstBalance > 0 And LOanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
Set RstInst = Nothing
Call gDbTrans.Fetch(RstInst, adOpenDynamic)

'Now Insert Into the Loan Trans
Dim UserID As Integer
UserID = gCurrUser.UserID

SqlStr = "Insert INTO LoanTrans(LoanID,TransDate,TransID," & _
    " TransType,Amount,Particulars,Balance,UserID) " & _
    " VALUES (" & LoanID & ", #" & TransDate & "#," & _
    TransID & ", " & transType & "," & PrincAmount & "," & _
    AddQuotes(Particulars, True) & "," & Balance & "," & UserID & ")"

gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    'MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If Not IsMissing(IntPaidUpToDate) Then
    LastTransDate = IntPaidUpToDate
Else
    LastTransDate = TransDate
End If

'Insert in to the loan int trans
SqlStr = "INSERT INTO LoanIntTrans (LoanID,TransDate," & _
    " TransID,TransType,IntAmount, PenalIntAmount," & _
    " IntBalance,MiscAmount,IntUptoDate,UserID) " & _
    "Values (" & LoanID & ", #" & TransDate & "#, " & TransID & "," & _
    transType & "," & IntAmount & "," & PenalAmount & "," & _
    IntBalance & "," & MiscAmount & ",#" & LastTransDate & "#," & UserID & " )"

gDbTrans.SqlStmt = SqlStr
'If he is paying ny interest then only do this transaction
If IntAmount > 0 Or PenalAmount > 0 Then
    If Not gDbTrans.SQLExecute Then
        'gDBTrans.RollBack
        MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If

If Balance = 0 Then
    SqlStr = "UPDATE LoanMaster SET LoanClosed = 1 " & _
                " WHERE LoanID = " & LoanID
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        'gDBTrans.RollBack
        'MsgBox "Unable to perform the operation", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If

If Not RstInst Is Nothing Then
    Dim Amount As Currency
    Dim InstAmount As Currency
    Dim InstBalance As Currency
    Dim InstNo As Integer
    Amount = PrincAmount
    Do
        If RstInst.EOF Then Exit Do
        If Amount <= 0 Then Exit Do
        InstAmount = FormatField(RstInst("InstBalance"))
        InstNo = FormatField(RstInst("InstNo"))
        If InstAmount >= Amount Then
            InstBalance = InstAmount - Amount
            Amount = 0 'Amount - Instp
        Else
            InstBalance = 0
            Amount = Amount - InstAmount
        End If
        SqlStr = "UPDATE LoanInst  Set InstBalance = " & InstBalance & _
            ", PaidDate = #" & TransDate & "#" & _
            " WHERE LoanID = " & LoanID & _
            " AND InstNo = " & InstNo
        gDbTrans.SqlStmt = SqlStr
        
        If Not gDbTrans.SQLExecute Then Exit Function
        
        RstInst.MoveNext
    Loop
End If

RepayLoan = True

Exit_Line:

        
End Function

Function OverDueSince(NoOfYear As Integer, Optional ByVal LoanID As Long = 0, _
        Optional ByVal SchemeID As Integer = 0, _
        Optional ByVal AsOnDate As Date) As Currency

Err.Clear

On Error GoTo ExitLine
Dim ODAmount As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim DueDate As Date

If IsMissing(AsOnDate) Then AsOnDate = CDate(gStrDate)

DueDate = DateAdd("yyyy", -NoOfYear, AsOnDate)
'First Get The Over due OF Non Installment Type
Dim InstType As wisInstallmentTypes

InstType = Inst_No
If LoanID Then
    SqlStr = "SELECT Balance FROM LoanTrans A,LoanMaster B WHERE B.LoanID = " & LoanID & _
        " AND LoanDueDate <= #" & DueDate & "# AND A.LoanID = B.LoanID" & _
        " AND InstMode = " & InstType & " " & _
        " AND TransID = (SELECT MAX(TransID) FROM " & _
            " LoanTrans C WHERE A.LoanID = C.LoanID " & _
            " AND C.TransDate <= #" & AsOnDate & "#)"
ElseIf SchemeID Then
    SqlStr = "SELECT SUM(Balance) as OverDue FROM LoanTrans A WHERE LoanID In " & _
        "(SELECT LoanID From LoanMaster WHERE LoanDueDate <= #" & DueDate & "# " & _
        " AND InstMode = " & InstType & " AND SchemeID = " & SchemeID & ") " & _
      "AND TransID = (SELECT MAX(TransID) FROM LoanTrans B WHERE A.LoanID = B.LoanID " & _
        " AND B.TransDate <= #" & AsOnDate & "#) "
Else
    SqlStr = "SELECT SUM(Balance) as OverDue FROM LoanTrans A WHERE LoanID In " & _
        "(SELECT LoanID From LoanMaster WHERE LoanDueDate <= #" & DueDate & "# " & _
        " AND InstMode = " & InstType & ") " & _
      "AND TransID = (SELECT MAX(TransID) FROM LoanTrans B WHERE A.LoanID = B.LoanID " & _
        "AND B.TransDate <= #" & AsOnDate & "#) "
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        ODAmount = FormatField(rst(0))
    End If

'Now Get The Over due Installment loans
If LoanID Then
    ODAmount = 0
    SqlStr = "SELECT SUM(InstBalance) as OverDue FROM LoanInst WHERE " & _
        " LoanId = " & LoanID & _
        " AND InstBalance > 0 AND InstDate <= #" & DueDate & "# "
ElseIf SchemeID Then
    SqlStr = "SELECT SUM(InstBalance) as OverDue FROM LoanInst WHERE LoanId In " & _
        "(SELECT LoanId FROM LOanMaster WHERE Schemeid = " & SchemeID & _
        " AND InstMode <> " & InstType & ") " & _
        "AND InstBalance > 0 AND InstDate <= #" & DueDate & "# "
Else
    SqlStr = "SELECT SUM(InstBalance) as OverDue FROM LoanInst WHERE " & _
        "InstBalance > 0 AND InstDate <= #" & DueDate & "# "
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        ODAmount = ODAmount + FormatField(rst(0))
    End If

ExitLine:
    If Err Then
        MsgBox "ERROR in OverDue since" & vbCrLf & Err.Description _
        , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
    OverDueSince = ODAmount \ 1

End Function

'THis function Returns the Balance Of Particular scheme
' OR Of the All Loan scheme, if schemeID is NULL
'This Private Function if it is necessary for  public use
'It will be called thru the Balance Function
Private Function SchemeBalance(Optional SchemeID As Integer = 0, Optional AsOnDate As Date) As Currency
Dim rst As Recordset

' Setup error handler.
Err.Clear
On Error GoTo Err_line

If AsOnDate = vbNull Then AsOnDate = gStrDate

' Declare variables...
Dim SqlStr As String

'Prepare the SQL query for getting loan balance of the given scheme.
SqlStr = "SELECT SUM(Balance) As TotalBalance FROM LoanTrans A WHERE " & _
    " TransID = (SELECT MAX(TransID) FROM LoanTrans B WHERE " & _
        "A.LoanId = B.LoanID AND TransDate <= #" & AsOnDate & "#)"

If SchemeID Then
    SqlStr = SqlStr & " AND A.LoanID In (SELECT LoanID From LoanMaster " & _
        " WHERE SchemeID = " & SchemeID & ")"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    SchemeBalance = FormatField(rst(0))
End If

Err_line:
    If Err Then
        MsgBox "LoanSchemeBalance: " & vbCrLf _
            & Err.Description, vbCritical
        Err.Clear
    End If

End Function
'This function Returns the Balance Of Particular Loan id loan id is specified,
'If LoanId is Null Then of Paricular scheme's Total Loan Balance
'OR Of the All Loan scheme, if schemeID is NULL
Public Function Balance(Optional ByVal LoanID As Long = 0, _
            Optional ByVal SchemeID As Integer = 0, Optional ByVal AsOnDate As Date) As Currency

'Setup error handler.
Err.Clear
On Error GoTo Err_line
Balance = 0


If IsMissing(AsOnDate) Then AsOnDate = CDate(gStrDate)
Dim rst As Recordset

If LoanID = 0 Then
    Balance = SchemeBalance(SchemeID, AsOnDate)
    Exit Function
End If


' Declare variables...
Dim SqlStr As String

'Prepare the SQL query for getting loan balance of the given scheme.
SqlStr = "SELECT TOP 1 Balance FROM LoanTrans A Where LoanID = " & LoanID & _
                " AND TransDate <= #" & AsOnDate & "# ORDER BY TransId Desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Balance = FormatField(rst(0))


Err_line:
    If Err Then
        MsgBox "LoanBalance: " & vbCrLf _
            & Err.Description, vbCritical
        'Resume
        Err.Clear
    End If
End Function

'This Function returns the Over due amount
'of loan/Scheme on Speicified date
'
Public Function OverDueAmount(Optional ByVal LoanID As Long = 0, _
        Optional ByVal SchemeID As Integer = 0, Optional ByVal AsOnDate As Date) As Currency

OverDueAmount = 0

Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim Duration As Long
Dim InstType As wisInstallmentTypes


Err.Clear
On Error GoTo ExitLine
ODAmount = 0

If IsMissing(AsOnDate) Then AsOnDate = CDate(gStrDate)

'Get The over due Amount of all loan scheme
'Get The Loan over due of the Particular loan scheme
If LoanID = 0 Then
    ODAmount = SchemeOverDue(SchemeID, AsOnDate)
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

Public Function OverDueInterest(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency
OverDueInterest = 0

Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim ODInt As Currency
Dim InstType As wisInstallmentTypes
Dim EMI As Boolean
Dim IntRate As Single
Dim LastDate As Date

Err.Clear
On Error GoTo ExitLine
ODInt = 0


'Get the Loan DueDate & Intreatrate
SqlStr = "SELECT * FROM LoanMaster Where " & _
    " LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
IntRate = FormatField(rst("IntRate"))
InstType = FormatField(rst("InstMode"))
DueDate = rst("LoanDueDate")
EMI = FormatField(rst("EMI"))

If EMI Then GoTo ExitLine  'Because interest _
will be included in the Principal

'Interest Over due is Interest on OverDue Amount
'So Get the Interest Amount

'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Top 1 * FROM LoanIntTrans Where " & _
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

ODInt = ODInt + PenalInterest(LoanID, , AsOnDate)
ODAmount = OverDueAmount(LoanID, , AsOnDate)
If ODAmount = 0 Then GoTo ExitLine
ODInt = ODInt '+ Balance(LoanID, , AsOnIndianDate) * (DateDiff("d", LastDate, AsOnDate) / 365) * IntRate


GoTo ExitLine





'Check for The Installmemet type
If InstType = Inst_No Then GoTo ExitLine
If InstType <> Inst_No Then
    'Get The Details form Instalment table
    Dim RstInst As Recordset
    SqlStr = "SELECT * FROM LoanInst WHERE LOanID = " & LoanID & _
        " AND InstDate < #" & AsOnDate & "# "
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
    'Set RstInst = gDBTrans.Rst.Clone
    'Now Calculate the PenalInterest Amount on each int
    ODAmount = 0
    While Not RstInst.EOF
        ODAmount = ODAmount + RstInst("InstBalance")
        RstInst.MoveNext
    Wend
End If

OverDueInterest = ODInt \ 1

ExitLine:
    
    If Err Then
        MsgBox "ERROR In Penal Interest", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    OverDueInterest = ODInt \ 1
End Function

Public Function DueInstallments(ByVal LoanID As Long, ByVal AsOnDate As Date) As Integer
DueInstallments = 0

Dim SqlStr As String
Dim rst As Recordset
Dim InstType As wisInstallmentTypes
Dim InstAmount As Currency

Err.Clear
On Error GoTo ExitLine


'Get the Loan DueDate & Interestrate
SqlStr = "SELECT * FROM LoanMaster Where " & _
    " LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
InstType = FormatField(rst("InstMode"))
InstAmount = FormatField(rst("InstAmount"))

'Check for The Installmemet type
If InstType = Inst_No Then GoTo ExitLine

If InstType <> Inst_No Then
    'Get The Details form Instalment table
    SqlStr = "SELECT * FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstDate < #" & AsOnDate & "#" & _
        " AND InstBalance = " & InstAmount
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
    'DueInstallments = gDBTrans.Records
    DueInstallments = rst.RecordCount
End If

ExitLine:
    If Err Then
        MsgBox "ERROR in Penal Interest", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    
End Function


'
'This Function Calculates the Penal interest of
'given loanid as on given date
'It Considers the Installments also
Public Function PenalInterest(Optional ByVal LoanID As Long = 0, _
            Optional ByVal SchemeID As Integer = 0, _
            Optional ByVal AsOnDate As Date) As Currency

PenalInterest = 0

Dim IntAmount As Currency

Err.Clear
On Error GoTo Exit_Line
IntAmount = 0

If IsMissing(AsOnDate) Then AsOnDate = gStrDate

If LoanID = 0 Then
    IntAmount = 0
    GoTo Exit_Line
End If


Dim SqlStr As String
Dim rst As Recordset
Dim PrinAmount As Currency
Dim LastTransDate As Date
Dim LastintPaidDate As Date

Dim Duration As Long
Dim PenalIntRate As Single
Dim DueDate As Date
Dim InstType As wisInstallmentTypes


'Get the LoanDueDate & Interest rate
SqlStr = "SELECT * FROM LoanMaster Where " & _
            " LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line

PenalIntRate = FormatField(rst("PenalIntRate"))
InstType = FormatField(rst("InstMode"))
DueDate = rst("LoanDueDate")

LastintPaidDate = vbNull
If FormatField(rst("LastIntDate")) <> "" Then LastintPaidDate = rst("LastIntDate")

If DateDiff("d", LastintPaidDate, AsOnDate) < 0 Then LastintPaidDate = vbNull
If DateDiff("d", DueDate, LastintPaidDate) > 0 Then DueDate = LastintPaidDate

If PenalIntRate = 0 Then GoTo Exit_Line

'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Top 1 * FROM LoanTrans Where " & _
        " LoanId = " & LoanID & " AND TransDate <= " & _
        " #" & AsOnDate & "# ORDER BY TransID Desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line

PrinAmount = FormatField(rst("Balance"))
If PrinAmount = 0 Then GoTo Exit_Line

'If Transcted date is as on the speicfied date then exit
LastTransDate = rst("TransDate")
Duration = DateDiff("d", LastTransDate, AsOnDate)

If Duration <= 0 Then GoTo Exit_Line

'Chek given date with Duedate
Duration = DateDiff("d", DueDate, AsOnDate)
If Duration > 0 Then
    IntAmount = PrinAmount * (Duration / 365) * (PenalIntRate / 100)
End If

'If It has installment the calculate the Penal Interest on installment
If InstType <> Inst_No Then
    'Get The Deatils from Instalment table
    Dim RstInst As Recordset
    SqlStr = "SELECT * FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstDate <= #" & AsOnDate & "# "
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(RstInst, adOpenDynamic) < 1 Then GoTo Exit_Line
    
    'Now Calculate the PenalInterest Amount on each inst
    Dim PenalAmount As Currency
    'If it has not crosed the ovedue date then penalAmount should be 0
    If DateDiff("d", DueDate, AsOnDate) <= 0 Then GoTo PenalCalc
    
    'If IT HasCrossed the Overdue date then
    If DateDiff("d", DueDate, AsOnDate) > 0 Then AsOnDate = DueDate
    
    While Not RstInst.EOF
       
       'instDate is lastTransDate if he fetches instdate it calaculates form the
       'last TransDate so Penal intrest will be wrong.so below line commented...
             
        'dueDate = RstInst("instDate")
        PrinAmount = RstInst("InstBalance")
        If PrinAmount > 0 Then
            Duration = DateDiff("d", DueDate, AsOnDate)
            PenalAmount = PrinAmount * Duration / 365 * PenalIntRate / 100
            IntAmount = IntAmount + PenalAmount
        End If
        RstInst.MoveNext
    Wend
End If

PenalCalc:
PenalInterest = IntAmount \ 1

Exit_Line:
    If Err Then
        MsgBox "ERROR In Penal Interest", vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
        Exit Function
    End If
    PenalInterest = IntAmount \ 1

End Function

'Shashi Angadi
'Modified on 20 the August 2002
'This Function Returns the Amount Collected as Regular INterest From Loan
'If Loan is provided it returns the Regular Interest of Particualr Loan Account
'If Scheme Id provided it returns the Regular Interest of Particualr type of Loan
'If Loanid And ScheemID both are missing then it returns the total Regular Interest collected
Public Function Proft_RegularInterest(Optional ByVal LoanID As Long = 0, _
    Optional ByVal SchemeID As Integer, Optional ByVal fromDate As Date, _
    Optional ByVal toDate As Date) As Currency

Dim Amount As Currency
On Error GoTo ExitLine

Dim SqlStr As String
Dim rst As Recordset

If LoanID Then   ' Repaid Amount Of Particular LoanId
    SqlStr = "Select Sum(IntAmount) From LoanIntTrans " & _
        " Where LoanID = " & LoanID & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
ElseIf SchemeID Then ' Repaid Amount Of Particular Loan
    SqlStr = "Select Sum(IntAmount) From LoanIntTrans " & _
        " Where LoanID IN " & _
            " (Select LoanId From LoanMaster Where Schemeid = " & SchemeID & ")" & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
Else   ' Repaid Amount Of All Loan
    SqlStr = "Select Sum(IntAmount) From LoanIntTrans Where " & _
        " TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Proft_RegularInterest = Val(FormatField(rst(0)))
    End If

ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description _
            , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    Exit Function

End Function

'Shashi Angadi
'Modifieson 20 the August 2002
'This Function Returns the Amount Collected as Penal INterest From Loan
'If Loan is provided it returns the PenalInterest f Particualr Loan Account
'If Scheme Id provided it returns the PenalInterest f Particualr type of Loan
'If Loan is provided it returns the PenalInterest f Particualr Loan Account
'If Loanid And ScheemID both are missing then it returns the total PenalInterest collected
Public Function Proft_PenalInterest(Optional ByVal LoanID As Long = 0, _
        Optional ByVal SchemeID As Integer, Optional ByVal fromDate As Date, _
        Optional ByVal toDate As Date) As Currency
        
Dim Amount As Currency
Dim rst As Recordset

On Error GoTo ExitLine
    
Dim SqlStr As String

If LoanID Then   ' Repaid Amount Of Particular Loan
    SqlStr = "Select Sum(PenalIntAmount) From LoanIntTrans " & _
        " Where LoanID = " & LoanID & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
ElseIf SchemeID Then ' Amount Of Particular Loan type
    SqlStr = "Select Sum(PenalIntAmount) From LoanIntTrans " & _
        " Where LoanID IN " & _
            " (Select LoanId From LoanMaster Where Schemeid = " & SchemeID & ")" & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
Else   ' Repaid Amount Of All Loans
    SqlStr = "Select Sum(PenalIntAmount) From LoanIntTrans Where " & _
        " TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        Proft_PenalInterest = Val(FormatField(rst(0)))
    
ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description _
            , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    Exit Function

End Function


Public Property Let SchemeID(NewValue As Integer)
    m_SchemeId = NewValue
End Property

'This Function Returns the over due amount as on date
'If SchmeID is missing it returns the over due of All
'Type of loans , If Date is missing it returns OD
'as on the login date
Private Function SchemeOverDue(Optional SchemeID As Integer = 0, Optional AsOnDate As Date) As Currency

Dim DueDate As Date
Dim SqlStr As String
Dim rst As Recordset
Dim ODAmount As Currency
Dim Duration As Long
Dim InstType As wisInstallmentTypes


Err.Clear
On Error GoTo Exit_Line
ODAmount = 0

'Get The over due Amount of all loan scheme
'Get The Loan over due of the Particular loan scheme
Dim SQLLoan As String

'sql to get the LoanIds
InstType = Inst_No
SQLLoan = "SELECT LoanID FROM LoanMaster WHERE LoanDueDate <= " & _
    " #" & AsOnDate & "# AND InstMode = " & InstType

If SchemeID Then
   SQLLoan = SQLLoan & " AND SchemeID = " & SchemeID
End If
  
'Prepare the SqlStatement for non instalment Loans
SqlStr = "SELECT SUM(Balance) as OverDue From LoanTrans A " & _
   " WHERE LoanID IN (" & SQLLoan & ") " & _
   " AND TransID =(SELECT MAX(TransID) FROM LoanTrans B WHERE " & _
      " B.LoanID = A.LoanID AND TransDate <= #" & AsOnDate & "# )"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'Set Rst = gDBTrans.Rst.Clone
    ODAmount = FormatField(rst(0))
End If

'Now get the over due amount Of Installment loans
'Get the Loan Advanced and Loan repaid upto this date
Dim curAdvance As Currency
Dim curRepay As Currency
    InstType = Inst_No
    SQLLoan = "SELECT LoanID From LoanMaster WHERE InstMode <> " & InstType
    If SchemeID Then _
            SQLLoan = SQLLoan & " And SchemeID = " & SchemeID
            
    SqlStr = "SELECT SUM(AMOUNT),TransType FROM LoanTrans WHERE" & _
        " LoanId IN (" & SQLLoan & ") " & _
        " AND TransDate < #" & AsOnDate & "# Group By TransType"
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    'Set Rst = gDBTrans.Rst.Clone
    While Not rst.EOF
        If FormatField(rst("TransType")) = wContraDeposit Or FormatField(rst("TransType")) = wDeposit Then
            curRepay = curRepay + FormatField(rst(0))
        ElseIf FormatField(rst("TransType")) = wContraWithdraw Or FormatField(rst("TransType")) = wWithdraw Then
            curAdvance = curAdvance + FormatField(rst(0))
        End If
        rst.MoveNext
    Wend
    
    Debug.Assert Balance = curAdvance - curRepay  'Inform To Shashi Angadi Immedietly
    
    'Get The total Instalment amount supposed
    'to be paid till on specified date
    Dim NetInstAmount As Currency
    
    SqlStr = "SELECT SUM(InstBalance) FROM LoanInst WHERE " & _
        " LoanID IN (" & SQLLoan & ") AND InstDate < #" & AsOnDate & "#"
    gDbTrans.SqlStmt = SqlStr
    
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    ODAmount = 0
    NetInstAmount = Val(FormatField(rst(0)))
    
    'The positive Difference between the netInstAmount and amount repaid
    'is the over due as on this date
    ODAmount = NetInstAmount - curRepay
    'If he has paid some installments in advance
    'Then it shows the negative Over due 'So
    If ODAmount < 0 Then ODAmount = 0

Exit_Line:
    If Err Then
        MsgBox "ERROR In Scheme over duw", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    SchemeOverDue = ODAmount \ 1

End Function

Public Property Get ObjectName()
    ObjectName = "LoanClass"
End Property


Private Function SchemePenalInterest(Optional SchemeID As Integer = 0, Optional AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntAmount As Currency

IntAmount = 0

SqlStr = "SELECT Distinct LoanID FROM LoanTrans "
If SchemeID Then
    SqlStr = SqlStr & " AND SchemeID = " & SchemeID
End If
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

'Set Rst = gDBTrans.Rst.Clone
Dim LoanID As Long

While Not rst.EOF
    LoanID = FormatField(rst(0))
    If LoanID Then IntAmount = IntAmount + PenalInterest(LoanID, , AsOnDate)
    rst.MoveNext
Wend

SchemePenalInterest = IntAmount \ 1

End Function

Private Function SchemeRegInterest(Optional SchemeID As Integer = 0, Optional AsOnDate As Date) As Currency
Dim SqlStr As String
Dim rst As Recordset
Dim IntAmount As Currency

IntAmount = 0

SqlStr = "SELECT Distinct LoanID FROM LoanTrans "
If SchemeID Then
    SqlStr = SqlStr & " AND SchemeID = " & SchemeID
End If
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

'Set Rst = gDBTrans.Rst.Clone
Dim LoanID As Long

While Not rst.EOF
    LoanID = FormatField(rst(0))
    If LoanID Then IntAmount = IntAmount + RegularInterest(LoanID, , AsOnDate)
    rst.MoveNext
Wend

SchemeRegInterest = IntAmount \ 1

End Function


Public Sub ShowLoanAccountDetail()
    If m_frmLoanTrans Is Nothing Then _
        Set m_frmLoanTrans = New frmLoanTrans
    Load m_frmLoanTrans
    If gWindowHandle Then m_ParentHandle = gWindowHandle
    gWindowHandle = m_frmLoanTrans.hwnd
    m_frmLoanTrans.Show
    
End Sub

Public Sub ShowCreateLoanAccount(Optional ByVal SchemeID As Integer)
    
If m_frmLoanMaster Is Nothing Then _
                Set m_frmLoanMaster = New frmLoanMaster
        
Load m_frmLoanMaster
If gWindowHandle Then m_ParentHandle = gWindowHandle
gWindowHandle = m_frmLoanMaster.hwnd
m_frmLoanMaster.Show

If m_LoanID Then m_frmLoanMaster.LoadLoan (m_LoanID)
    
End Sub

Public Sub ShowLoanReport()
    
    If m_frmSelectReport Is Nothing Then _
            Set m_frmSelectReport = New frmLoanReport
        
    Load m_frmSelectReport
    If gWindowHandle Then m_ParentHandle = gWindowHandle
    
    gWindowHandle = m_frmSelectReport.hwnd
    m_frmSelectReport.Show
    
End Sub
Public Sub ShowLoanSchemes(Optional ByVal SchemeID As Integer)
    If m_frmScheme Is Nothing Then _
        Set m_frmScheme = New frmLoanScheme
        
    Load m_frmScheme
    If gWindowHandle Then m_ParentHandle = gWindowHandle
    
    gWindowHandle = m_frmScheme.hwnd
    m_frmScheme.Show
    
End Sub

'This function Returns amount withdrawn in form of cash From a Particular loan
'If loanid is missing then It returns the Withdrawl amount
'of All Loans or particular loan type
Public Function CashWithdrawls(ByVal fromDate As Date, ByVal toDate As Date, _
    Optional ByVal LoanID As Long, Optional ByVal SchemeID As Integer = 0) As Currency
On Error GoTo ExitLine
    
    
Dim transType  As wisTransactionTypes

Dim Amount As Currency
transType = wWithdraw
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

CashWithdrawls = Amount
ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description & _
           " ,vbInformation ,wis_MESSAGE_TITLE"
        Err.Clear
    End If
    Exit Function
            
End Function

'This function Returns amount withdrawn and Physical cash is not handled
'From a Particular loan
'If loanid is missing then It returns the Withdrawl amount
'of All Loans or particular loan type
Public Function ContraWithdrawls(ByVal fromDate As Date, ByVal toDate As Date, _
        Optional ByVal LoanID As Long, Optional ByVal SchemeID As Integer = 0) As Currency
On Error GoTo ExitLine
    
    
Dim transType  As wisTransactionTypes

Dim Amount As Currency
transType = wContraWithdraw
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

ContraWithdrawls = Amount
ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description & _
           " ,vbInformation ,wis_MESSAGE_TITLE"
        Err.Clear
    End If
    Exit Function
            
End Function


'This function Returns amount withdrawn From a Particular loan
'If loanid is missing then It returns the Withdrawl amount
'of All Loans or particular loan type
Public Function WithDrawls(ByVal fromDate As Date, ByVal toDate As Date, _
    Optional ByVal LoanID As Long, Optional ByVal SchemeID As Integer = 0) As Currency
On Error GoTo ExitLine


Dim transType  As wisTransactionTypes
Dim Amount As Currency
transType = wWithdraw
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

WithDrawls = Amount
ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description & _
           " ,vbInformation ,wis_MESSAGE_TITLE"
        Err.Clear
    End If
    Exit Function
            
End Function


'This function Returns amount deposits THROUGH to a Particular loan
'If loanid is missing then It returns the repayment of All Loans
Public Function ContraDeposits(ByVal fromDate As Date, ByVal toDate As Date, _
    Optional ByVal LoanID As Long, _
    Optional ByVal SchemeID As Integer = 0) As Currency

Dim transType  As wisTransactionTypes
Dim Amount As Currency
transType = wContraDeposit
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)


ContraDeposits = Amount
End Function

'This function Returns amount deposits where physical cash handled
'If loanid is missing then It returns the repayment of All Loans
Public Function CashDeposits(fromDate As Date, toDate As Date, _
    Optional LoanID As Long, Optional SchemeID As Integer = 0) As Currency

Dim transType  As wisTransactionTypes
Dim Amount As Currency
transType = wDeposit
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

CashDeposits = Amount
End Function

'This function Returns amount deposits to a Particular loan
'If loanid is missing then It returns the repayment of All Loans
Public Function Deposits(fromDate As Date, toDate As Date, _
    Optional LoanID As Long, Optional SchemeID As Integer = 0) As Currency

Dim transType  As wisTransactionTypes
Dim Amount As Currency
transType = wDeposit
Amount = TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)
transType = wContraDeposit
Amount = Amount + TransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)


Deposits = Amount
End Function
Public Sub TotalLoanTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional SchemeID As Integer = 0)
CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from LoanIntTrans where "

If SchemeID > 0 Then
    SqlStr = SqlStr & " AccID In " & _
        "(select distinct(LoanID) From LoanMaster where SchemeID = " & SchemeID & ") And "
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
    If rst("TransType") = wContraDeposit Then contraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then contraWithdraw = TotalAmount
Wend

Set rst = Nothing


End Sub
Public Sub TotalLoanIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional SchemeID As Integer = 0)
CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(IntAmount) as TotalAmount,TransType from LoanIntTrans where "

If SchemeID > 0 Then
    SqlStr = SqlStr & " AccID In " & _
        "(select distinct(LoanID) From LoanMaster where SchemeID = " & SchemeID & ") And "
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
    If rst("TransType") = wContraDeposit Then contraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then contraWithdraw = TotalAmount
Wend

Set rst = Nothing

End Sub
Public Sub TotalLoanPenalIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional SchemeID As Integer = 0)
CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(PenalIntAmount) as TotalAmount,TransType from LoanIntTrans where "

If SchemeID > 0 Then
    SqlStr = SqlStr & " AccID In " & _
        "(select distinct(LoanID) From LoanMaster where SchemeID = " & SchemeID & ") And "
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
    If rst("TransType") = wContraDeposit Then contraDeposit = TotalAmount
    If rst("TransType") = wContraWithdraw Then contraWithdraw = TotalAmount
Wend

Set rst = Nothing


End Sub

Private Function TransactionDetails(Optional ByVal LoanID As Long, _
        Optional ByVal SchemeID As Integer = 0, Optional ByVal fromDate As Date, _
        Optional ByVal toDate As Date, Optional ByVal transType As wisTransactionTypes) As Currency

On Error GoTo ExitLine
    
Dim SqlStr As String
Dim rst As Recordset

If LoanID Then   ' Repaid Amount Of Particular LoanId
    SqlStr = "Select Sum(Amount) From LoanTrans Where LoanID = " & LoanID & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
ElseIf SchemeID Then ' Repaid Amount Of Particular Loan
    SqlStr = "Select Sum(Amount) From LoanTrans Where LoanID IN " & _
            " (Select LoanId From LoanMaster Where Schemeid = " & SchemeID & ")" & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
Else   ' Repaid Amount Of All Loan
    SqlStr = "Select Sum(Amount) From LoanTrans Where " & _
        " TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then TransactionDetails = Val(FormatField(rst(0)))

ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description _
            , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    Exit Function
            
End Function

Public Function UndoTransaction(LoanID As Long, TransID As Long) As Boolean

Dim SqlStr As String
Dim rst As Recordset
Dim lastTransID  As Integer
Dim Amount As Currency
Dim IntAmount As Currency
Dim PenalIntAmount As Currency
Dim MiscAmount As Currency

Dim InstAmount As Currency
Dim PaidInst As Currency
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes

Dim TransDate As Date

Dim InstBalance() As Currency
Dim InstNo() As Integer
Dim count As Integer
Dim rstMaster As Recordset
gDbTrans.SqlStmt = "Select * From LoanMaster Where LoanID = " & LoanID
Call gDbTrans.Fetch(rstMaster, adOpenDynamic)

InstAmount = FormatField(rstMaster("InstAmount"))

SqlStr = "SELECT Top 1 * FROM LoanTrans WHERE LoanID = " & LoanID & _
    " ORDER BY TransID Desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

lastTransID = FormatField(rst("TransID"))
Amount = FormatField(rst("Amount"))
transType = rst("TransType")
TransDate = rst("TransDate")

SqlStr = "SELECT Top 1 * FROM LoanIntTrans WHERE LoanID = " & LoanID & _
            " ORDER BY TransID Desc"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
  If rst("TransID") >= lastTransID Then
    If lastTransID < rst("TransID") Then Amount = 0
    IntTransType = rst("TransType")
    lastTransID = rst("TransID")
    TransDate = rst("TransDate")
    IntAmount = FormatField(rst("IntAmount"))
    PenalIntAmount = FormatField(rst("PenalIntAmount"))
    MiscAmount = FormatField(rst("MiscAmount"))
  End If
End If

If lastTransID <> TransID Then
    MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Dim TotAmount As Currency
If rstMaster("EMI") Then
    TotAmount = PenalIntAmount + IntAmount + Amount
Else
    TotAmount = IntAmount + Amount
End If

'Fetch the details of Installment table
If FormatField(rstMaster("InstMode")) <> 0 And Amount > 0 Then
    SqlStr = "SELECT * FROM LoanInst WHERE LoanID = " & LoanID & _
        " AND InstBalance < " & InstAmount & _
        " ORDER BY InstDate Desc"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
        count = 0
        Do
            If rst.EOF Then Exit Do
            
            ReDim Preserve InstBalance(count)
            ReDim Preserve InstNo(count)
            InstBalance(count) = FormatField(rst("InstBalance"))
            InstNo(count) = FormatField(rst("InstNo"))
            count = count + 1
            rst.MoveNext
        Loop
        Set rst = Nothing
    End If
End If


'Now Delete the Transcation in the loan trans
SqlStr = "DELETE * FROM LoanTrans WHERE LoanID = " & LoanID & _
        " AND TransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

'Now Delete the Transcation of interest
SqlStr = "DELETE * FROM LoanIntTrans WHERE LoanID = " & LoanID & _
        " AND TransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

SqlStr = "UPDATE LoanMaster Set LoanClosed = 0 WHERE LoanID = " & LoanID

If IntTransType = wContraDeposit Then
    gDbTrans.SqlStmt = "Select Max(TransDate) From LoanIntTrans " & _
            " WHERE LoanID = " & LoanID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        SqlStr = "UPDATE LoanMaster Set LoanClosed = 0, " & _
            "LastIntDate = #" & rst(0) & "# WHERE LoanID = " & LoanID
End If

gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

'NOW UPDate the Installment detailas
If transType = wContraWithdraw And lastTransID <> 1 Then

Else
  If count > 0 Then
    count = 0
    If FormatField(rstMaster("EMI")) Then Amount = TotAmount
    TotAmount = Amount
    Do
        If count > UBound(InstBalance) Then Exit Do
        If TotAmount <= 0 Then Exit Do
        PaidInst = InstAmount - InstBalance(count)
        If TotAmount <= PaidInst Then
            InstBalance(count) = InstBalance(count) + TotAmount
            TotAmount = 0
        Else
          'If his last payment has affected more than one
          'installment amount then 'Consider it
            InstBalance(count) = InstAmount
            TotAmount = TotAmount - PaidInst
        End If
        If lastTransID = 1 Then InstBalance(count) = InstAmount
        SqlStr = "UPDATE LoanInst Set InstBalance = " & InstBalance(count) & _
            " WHERE LoanID = " & LoanID & " AND InstNo = " & InstNo(count)
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        count = count + 1
    Loop
  End If
End If

UndoTransaction = True

ExitLine:


End Function

Public Function ContraLoss(fromDate As Date, toDate As Date, _
    Optional LoanID As Long = 0, Optional SchemeID As Integer) As Currency

Dim Amount As Currency
Dim transType As wisTransactionTypes

transType = wContraWithdraw
Amount = IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

ContraLoss = Amount

End Function

Public Function CashLoss(fromDate As Date, toDate As Date, _
    Optional LoanID As Long = 0, Optional SchemeID As Integer) As Currency

Dim Amount As Currency
Dim transType As wisTransactionTypes

transType = wWithdraw
Amount = IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

CashLoss = Amount

End Function

Public Function Loss(Optional LoanID As Long = 0, Optional SchemeID As Integer, Optional fromDate As Date, Optional toDate As Date) As Currency

Dim Amount As Currency
Dim transType As wisTransactionTypes

transType = wWithdraw
Amount = IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

Loss = Amount

End Function

Public Function CashProfit(fromDate As Date, toDate As Date, _
    Optional LoanID As Long = 0, Optional SchemeID As Integer) As Currency

Dim Amount As Currency
Dim transType As wisTransactionTypes

transType = wDeposit
Amount = IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

CashProfit = Amount

End Function


Public Function ContraProfit(fromDate As Date, toDate As Date, _
    Optional LoanID As Long = 0, Optional SchemeID As Integer) As Currency

Dim Amount As Currency
Dim transType As wisTransactionTypes

transType = wContraDeposit
Amount = IntTransactionDetails(LoanID, SchemeID, fromDate, toDate, transType)

ContraProfit = Amount

End Function

Private Function IntTransactionDetails(Optional ByVal LoanID As Long = 0, _
        Optional ByVal SchemeID As Integer = 0, Optional ByVal fromDate As Date, _
        Optional ByVal toDate As Date, Optional ByVal transType As wisTransactionTypes) As Currency
On Error GoTo ExitLine
    

Dim SqlStr As String
Dim rst As Recordset

If LoanID Then   ' Repaid Amount Of Particular LoanId
    SqlStr = "Select Sum(IntAmount+PenalIntAmount) From LoanIntTrans " & _
        " Where LoanID = " & LoanID & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
ElseIf SchemeID Then ' Repaid Amount Of Particular Loan
    SqlStr = "Select Sum(IntAmount+PenalIntAmount) From LoanIntTrans " & _
        " Where LoanID IN " & _
            " (Select LoanId From LoanMaster Where Schemeid = " & SchemeID & ")" & _
        " And TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
Else   ' Repaid Amount Of All Loan
    SqlStr = "Select Sum(IntAmount+PenalIntAmount) From LoanIntTrans Where " & _
        " TransDate >= #" & fromDate & "# " & _
        " And Transdate <= #" & toDate & "# " & _
        " And TransType = " & transType
End If
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        IntTransactionDetails = Val(FormatField(rst(0)))
    End If

ExitLine:
    If Err Then
        MsgBox "ERROR Loan_Deposits" & vbCrLf & Err.Description _
            , vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If
    Exit Function
End Function

Public Function LoanList(LoanNames() As String, SchemeID() As Long, _
                    Optional ByVal LoanCategory As wisLoanCategories, _
                    Optional ByVal LoanTerm As wisLoanTerm) As Integer

' Initialize the return value.
LoanList = 0

' Declare variables.
Dim Lret As Long
Dim count As Integer
Dim rst As Recordset

' Setup error handler.
On Error GoTo Err_line

' Initialize the loan names and scheme id arrays.
ReDim LoanNames(0)
ReDim SchemeID(0)
Dim SqlStmt As String

' Query the loan master table.
SqlStmt = "Select * From LoanScheme"

If LoanCategory <> 0 Then
    SqlStmt = SqlStmt & " Where Category = " & LoanCategory
End If

If LoanTerm <> 0 Then
    If InStr(1, SqlStmt, "Where", vbTextCompare) Then
        SqlStmt = SqlStmt & " And "
    Else
        SqlStmt = SqlStmt & " Where "
    End If
    SqlStmt = SqlStmt & " LoanScheme.TermType = " & LoanTerm
End If

gDbTrans.SqlStmt = SqlStmt

Lret = gDbTrans.Fetch(rst, adOpenDynamic)
If Lret < 1 Then Exit Function
'Set Rst = gDBTrans.Rst.Clone
 While Not rst.EOF
    LoanNames(count) = FormatField(rst("SchemeName"))
    SchemeID(count) = FormatField(rst("SchemeId"))
    
    ReDim Preserve LoanNames(UBound(LoanNames) + 1)
    ReDim Preserve SchemeID(UBound(SchemeID) + 1)
    count = count + 1
    rst.MoveNext
 Wend
 ReDim Preserve LoanNames(UBound(LoanNames) - 1)
 ReDim Preserve SchemeID(UBound(SchemeID) - 1)
    
' Return the value for function
LoanList = rst.RecordCount

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "LoanList: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
    End If

    GoTo Exit_Line

End Function


'This Function Calculates the regular interest of
'given loan id as on given date
Public Function RegularInterest(Optional ByVal LoanID As Long = 0, _
                Optional ByVal SchemeID As Integer = 0, _
                Optional ByVal AsOnDate As Date) As Currency

RegularInterest = 0

Err.Clear
On Error GoTo ExitLine


If IsMissing(AsOnDate) Then AsOnDate = gStrDate

'If Interest is of particular Scheme or of all balance then
If LoanID = 0 Then
    RegularInterest = SchemeRegInterest(SchemeID, AsOnDate)
    Exit Function
End If

Dim SqlStr As String
Dim rst As Recordset
Dim IntAmount As Currency
Dim IntRate As Single
Dim LastintPaidDate As Date
Dim Balance As Currency

'Get the LoanDueDate & Intreatrate
SqlStr = "SELECT * FROM LoanMaster Where LoanId = " & LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
IntRate = FormatField(rst("IntRate"))
If IntRate = 0 Then GoTo ExitLine

'Get The Last Balance and TransDate of the loan
SqlStr = "SELECT Balance,TransDate,SchemeID,IntRate,IssueDate," & _
    " B.LastIntDate FROM LoanTrans A,LoanMaster B " & _
    " Where B.LoanID = A.LoanID And A.LoanId = " & LoanID & _
    " AND TransDate <= #" & AsOnDate & "# ORDER BY TransID Desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine

'Set Rst = gDBTrans.Rst.Clone
IntRate = rst("IntRate")
LastintPaidDate = IIf(FormatField(rst("LastIntdate")) = "", vbNull, rst("LastIntdate"))
If DateDiff("d", LastintPaidDate, AsOnDate) < 0 Then LastintPaidDate = vbNull

'If FormatField(rst("LastIntdate")) = "" Then
If LastintPaidDate = vbNull Then
    gDbTrans.SqlStmt = "SELECT * From LoanIntTrans WHERE " & _
            " LoanID = " & LoanID & " AND TransDate <= #" & AsOnDate & "#" & _
            " ORDER BY TransID Desc "
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        If rst.BOF Then GoTo nextRecord:
        LastintPaidDate = rst("TransDate")
    Else
        gDbTrans.SqlStmt = "SELECT * From LoanTrans WHERE " & _
            " LoanID = " & LoanID & " ORDER BY TransID"
        If gDbTrans.Fetch(rst, adOpenDynamic) Then LastintPaidDate = rst("TransDate")
    End If
Else
    LastintPaidDate = rst("LastIntdate")
    'calculate the intrestAmount form lastintpaid Date.
    gDbTrans.SqlStmt = "SELECT * From LoanIntTrans WHERE " & _
            " LoanID = " & LoanID & " AND TransDate <= #" & AsOnDate & "#" & _
            " ORDER BY TransID Desc "
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        'If DateDiff("d", LastintPaidDate, rst("TransDate")) > 1 Then LastintPaidDate = rst("TransDate")
    Else
        gDbTrans.SqlStmt = "SELECT * From LoanTrans WHERE " & _
            " LoanID = " & LoanID & " ORDER BY TransID"
    End If
End If

nextRecord:
'Now Get the Transaction Of the LastIntPaid date
gDbTrans.SqlStmt = "SELECT * FROM LoanTrans Where LoanID = " & LoanID & _
    " AND TransDate >= #" & LastintPaidDate & "# ORDER BY TransID"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    'Get the Balance as on lastIntpaid date
    gDbTrans.SqlStmt = "SELECT Top 1 * FROM LoanTrans Where LoanID = " & LoanID & _
        " ORDER BY TransID desc "
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
    Balance = rst("Balance")
    rst.MoveNext
Else
    Balance = rst("Balance")
End If

'Set Rst = gDBTrans.Rst.Clone
IntAmount = 0

Dim TillDate As Date
Dim Duration As Long

Do
    If rst.EOF Then
        Duration = DateDiff("d", LastintPaidDate, AsOnDate)
        IntAmount = IntAmount + ((Duration / 365) * (IntRate / 100) * Balance)
        Exit Do
    End If
    TillDate = rst("TransDate")
    Duration = DateDiff("d", LastintPaidDate, TillDate)
    IntAmount = IntAmount + ((Duration / 365) * (IntRate / 100) * Balance)
    LastintPaidDate = TillDate
    Balance = FormatField(rst("BALANCE"))
    rst.MoveNext
Loop
        
    RegularInterest = IntAmount \ 1

ExitLine:
    If Err Then
        MsgBox "ERROR In RegularInterest", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
    End If
    
RegularInterest = IntAmount \ 1
    
End Function


'This Function Calculates the regular interest of
'given loan id as on given date
Public Function PrincipalAmountAsOn(ByVal LoanID As Long, ByVal AsOnDate As Date) As Currency

PrincipalAmountAsOn = 0

Err.Clear
On Error GoTo ExitLine

Dim SqlStr As String
Dim rst As Recordset
Dim PrincAmount As Currency

'Get the LoanDueDate & Intreatrate
SqlStr = "SELECT * FROM LoanMaster Where LoanId = " & LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine


SqlStr = "SELECT * FROM LoanInst WHERE LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine
''Get the Principal till now
While rst.EOF = False
    If DateDiff("D", rst("InstDate"), AsOnDate) < 0 Then
        rst.MoveLast
    Else
        PrincAmount = PrincAmount + rst("InstAmount")
    End If
    rst.MoveNext
Wend

''Get the Total Principal paid till this date
SqlStr = "SELECT sum(amount) FROM LoanTrans WHERE LoanId = " & LoanID & " and (TransType = " & wisTransactionTypes.wDeposit & " or TransType = " & wisTransactionTypes.wContraDeposit & " ) and TransDate <= #" & AsOnDate & "#"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    PrincipalAmountAsOn = PrincAmount - FormatField(rst(0))
    If PrincipalAmountAsOn < 0 Then PrincipalAmountAsOn = 0
Else
    PrincipalAmountAsOn = PrincAmount
End If


ExitLine:
    If Err Then
        MsgBox "ERROR In RegularInterest", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
        Exit Function
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
gDbTrans.SqlStmt = "SELECT top 1 Balance,TransID,TransDate FROM " & _
    " LoanTrans WHERE loanid = " & LoanID & " ORDER BY TransId Desc"

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
gDbTrans.SqlStmt = "SELECT TransID,TransDate FROM LoanIntTrans " & _
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
    gDbTrans.SqlStmt = "Select SchemeName From LoanScheme " & _
            " WHERE SchemeID =(Select SchemeID From LoanMaster " & _
                " Where LOanID = " & LoanID & ")"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
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

gDbTrans.SqlStmt = "INSERT INTO LoanTrans (LoanID, TransID, " _
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


Public Function DepositAmount(ByVal LoanID As Long, ByVal PrincAmount As Currency, _
        Optional ByVal IntAmount As Currency, Optional ByVal PenalAmount As Currency, _
        Optional ByVal Particulars As String, Optional ByVal TransDate As Date, _
        Optional ByVal VoucherNo As String = " ", _
        Optional CashTransaction As Boolean, Optional ContraTransID As Long = 0) As Integer

Dim SqlStr As String
Dim rst As Recordset
Dim RstInst As Recordset
Dim Balance As Double
Dim LastTransDate As Date

Dim c_RegInt As Currency
Dim c_PenalInt As Currency
Dim IntBalance As Currency
Dim PenalBalance As Currency


c_RegInt = RegularInterest(LoanID, , TransDate)
c_PenalInt = PenalInterest(LoanID, , TransDate)

IntBalance = RegInterestBalance(LoanID, TransDate)
PenalBalance = PenalInterestBalance(LoanID, TransDate)

''''''''
'First Case If He is paying the interest amount then
Dim RetMsg As VbMsgBoxResult

'First Case If He is paying the interest amount then
If IntAmount + PenalAmount > 0 Then
    'The Amount He is payin less than the Blance interest
    If IntAmount + IntAmount < c_RegInt + c_PenalInt + IntBalance + PenalBalance Then
        'RetMsg = MsgBox("The amount he is paying is lesser than his interest till date" & _
            "Do you want keep the difference as balance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(668) & _
            GetResourceString(669), vbYesNo + vbQuestion, wis_MESSAGE_TITLE)
        'If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then
            c_RegInt = IntAmount - IntBalance
            c_PenalInt = PenalAmount - PenalBalance
        End If
    ElseIf IntAmount + PenalAmount > c_RegInt + c_PenalInt + IntBalance + PenalBalance Then
        'RetMsg = MsgBox("The amount he is paying is more than his interest till date" & _
            "Do you want keep the difference as balance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(666) & _
            GetResourceString(667), vbYesNo + vbQuestion, wis_MESSAGE_TITLE)
        'If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then
            c_RegInt = IntAmount - IntBalance
            c_PenalInt = PenalAmount - PenalBalance
        End If
    End If
End If

IntBalance = IntBalance + c_RegInt - IntAmount
PenalBalance = PenalBalance + c_PenalInt - PenalAmount

Dim transType As wisTransactionTypes
Dim TransID As Long

If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
End If

SqlStr = "SELECT Top 1 * FROM LoanTrans WHERE LoanID = " & LoanID & _
    " ORDER BY TransId Desc"
gDbTrans.SqlStmt = SqlStr

TransID = gDbTrans.Fetch(rst, adOpenDynamic)
If TransID < 0 Then Exit Function

If TransID > 0 Then
    LastTransDate = rst("TransDate")
    If DateDiff("d", LastTransDate, TransDate) < 0 Then
        Err.Raise 5102, "Loan Class", "Invalid Date"
        Exit Function
    End If
    Balance = FormatField(rst("Balance"))
    TransID = FormatField(rst("TransID"))
    'Now get the Max Trans Id form Interest Transaction
    SqlStr = "SELECT Top 1 * FROM LoanIntTrans WHERE LoanID = " & LoanID & _
            " ORDER BY TransId Desc"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then _
        TransID = IIf(TransID < rst("TransID"), rst("TransID"), TransID)
End If

If Balance < PrincAmount Then
    'MsgBox "Total amount is greater than LoanAmount.", vbInformation
    MsgBox GetResourceString(585), vbInformation, wis_MESSAGE_TITLE
    Exit Function
    'Err.Raise 5103, "LoanRepay", "Amount not matching"
End If

Balance = Balance - PrincAmount
TransID = TransID + 1
transType = IIf(CashTransaction, wDeposit, wContraDeposit)

SqlStr = "SELECT * From LoanInst Where InstBalance > 0 And LOanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
Set RstInst = Nothing
Call gDbTrans.Fetch(RstInst, adOpenDynamic)

'Now Insert Into the Loan Trans
Dim UserID As Integer
UserID = gCurrUser.UserID

If PrincAmount > 0 Then
    SqlStr = "Insert INTO LoanTrans(LoanID,TransDate,TransID," & _
        " TransType,Amount,Particulars,Balance,VoucherNo,UserID) " & _
        " VALUES (" & LoanID & ", #" & TransDate & "#," & _
        TransID & ", " & transType & "," & PrincAmount & "," & _
        AddQuotes(Particulars, True) & "," & Balance & "," & _
        AddQuotes(VoucherNo, True) & "," & UserID & " )"
        
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
End If

'Insert in to the loan int trans
SqlStr = "INSERT INTO LoanIntTrans (LoanID,TransDate," & _
    " TransID,TransType,IntAmount, PenalIntAmount," & _
    " IntBalance,MiscAmount,IntUptoDate,VoucherNo,UserID) " & _
    "Values (" & LoanID & ", #" & TransDate & "#, " & TransID & "," & _
    transType & "," & IntAmount & "," & PenalAmount & "," & _
    IntBalance & ",0  ,#" & LastTransDate & "#," & _
    AddQuotes(VoucherNo, True) & "," & UserID & " )"

gDbTrans.SqlStmt = SqlStr
'If he is paying ny interest then only do this transaction
If IntAmount > 0 Or PenalAmount > 0 Then
    If Not gDbTrans.SQLExecute Then
        'gDBTrans.RollBack
        MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If

'update it to the Contra Trans table

Dim AccHeadID As Long
Dim headName As String

gDbTrans.SqlStmt = "Select SchemeName From LoanScheme " & _
        " WHERE SchemeID =(Select SchemeID From LoanMaster " & _
            " Where LOanID = " & LoanID & ")"
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
headName = FormatField(rst("SchemeName"))

If PrincAmount <> 0 And Not CashTransaction Then
    AccHeadID = GetIndexHeadID(headName) 'loadresstring("
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        " (ContraID,AccHeadID,AccID, " & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        LoanID & ", " & transType & ", " & TransID & "," & _
        PrincAmount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

If IntAmount <> 0 And Not CashTransaction Then
    AccHeadID = GetIndexHeadID(headName & " " & GetResourceString(344))
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        " (ContraID,AccHeadID,AccID, " & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        LoanID & ", " & transType & ", " & TransID & "," & _
         IntAmount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

If PenalAmount <> 0 And Not CashTransaction Then
    AccHeadID = GetIndexHeadID(headName & " " & GetResourceString(345))
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        " (ContraID,AccHeadID,AccID, " & _
        "TransType, TransID,Amount,VoucherNo,UserID)" & _
        " Values (" & ContraID & "," & AccHeadID & "," & _
        LoanID & ", " & transType & ", " & TransID & "," & _
        PenalAmount & ", " & _
        AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

If Balance = 0 Then
    SqlStr = "UPDATE LoanMaster SET LoanClosed = 1 " & _
                " WHERE LoanID = " & LoanID
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        'gDBTrans.RollBack
        'MsgBox "Unable to perform the operation", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If

If Not RstInst Is Nothing Then
    Dim Amount As Currency
    Dim InstAmount As Currency
    Dim InstBalance As Currency
    Dim InstNo As Integer
    Amount = PrincAmount
    Do
        If RstInst.EOF Then Exit Do
        If Amount <= 0 Then Exit Do
        InstAmount = FormatField(RstInst("InstBalance"))
        InstNo = FormatField(RstInst("InstNo"))
        If InstAmount >= Amount Then
            InstBalance = InstAmount - Amount
            Amount = 0 'Amount - Instp
        Else
            InstBalance = 0
            Amount = Amount - InstAmount
        End If
        SqlStr = "UPDATE LoanInst  Set InstBalance = " & InstBalance & _
            ", PaidDate = #" & TransDate & "#" & _
            " WHERE LoanID = " & LoanID & _
            " AND InstNo = " & InstNo
        gDbTrans.SqlStmt = SqlStr
        
        If Not gDbTrans.SQLExecute Then Exit Function
        
        RstInst.MoveNext
    Loop
End If

DepositAmount = TransID

Exit_Line:

End Function

Private Sub Class_Terminate()
    
On Error Resume Next
If Not m_frmLoanInst Is Nothing Then
    Unload m_frmLoanInst
    Set m_frmLoanInst = Nothing
End If
If Not m_frmLoanMaster Is Nothing Then
    Unload m_frmLoanMaster
    Set m_frmLoanMaster = Nothing
End If
If Not m_frmLoanPay Is Nothing Then
    Unload m_frmLoanPay
    Set m_frmLoanPay = Nothing
End If
If Not m_frmLoanReport Is Nothing Then
    Unload m_frmLoanReport
    Set m_frmLoanReport = Nothing
End If
If Not m_frmLoanTrans Is Nothing Then
    Unload m_frmLoanTrans
    Set m_frmLoanTrans = Nothing
End If
Set m_frmRegReport = Nothing
If Not m_frmScheme Is Nothing Then
    Unload m_frmScheme
    Set m_frmScheme = Nothing
End If
If Not m_frmSelectReport Is Nothing Then
    Unload m_frmSelectReport
    Set m_frmSelectReport = Nothing
End If

End Sub

Private Sub m_frmLoanInst_WindowClosed()

If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmLoanInst = Nothing
End Sub


Private Sub m_frmLoanMaster_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmLoanMaster = Nothing
End Sub


Private Sub m_frmLoanPay_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmLoanPay = Nothing
End Sub


Private Sub m_frmLoanReport_WindowClosed()
    Set m_frmLoanReport = Nothing
End Sub

Private Sub m_frmLoanTrans_WindowClosed()
    If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
    Set m_frmLoanTrans = Nothing
End Sub


Private Sub m_frmRegReport_Initialise(Min As Long, Max As Long)
On Error Resume Next
With frmCancel
    If Max <> 0 Then
        .PicStatus.Visible = True
        .Refresh
    End If
    UpdateStatus .PicStatus, 0
End With

End Sub

Private Sub m_frmRegReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS: " & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
End With
End Sub


Private Sub m_frmScheme_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmScheme = Nothing
End Sub


Private Sub m_frmSelectReport_ShowReport11(SchemeID As Integer, _
                ReportType As wis_LoanReports, ReportOrder As wis_ReportOrder, _
                fromDate As String, toDate As String, _
                FromAmount As Currency, ToAmount As Currency, _
                Gender As wis_Gender, Place As String, _
                Caste As String, AccountGroup As Integer)
#If OldVersion Then
                FromAmount As Currency, ToAmount As Currency, _
                Gender As wis_Gender, Place As String, _
                Caste As String, AccountGroup As Integer
gCancel = 0
frmCancel.Show
frmCancel.Refresh
Screen.MousePointer = vbHourglass
Set m_frmRegReport = Nothing

If ReportType = repMonthlyRegister Or ReportType = repMonthlyRegisterAll Or _
    ReportType = repShedule_1 Or ReportType = repShedule_2 Or ReportType = repShedule_3 Or _
    ReportType = repShedule_4A Or ReportType = repShedule_4B Or ReportType = repShedule_4C Or _
    ReportType = repShedule_5 Or ReportType = repShedule_6 Then
    
    Set m_frmRegReport = New frmRegReport
    With m_frmRegReport
        .ReportType = ReportType
        .FromIndianDate = fromDate
        .ToIndianDate = toDate
        .FromAmount = FromAmount
        .ToAmount = ToAmount
        .Place = Place
        .Caste = Caste
        .ReportOrder = ReportOrder
        .LoanSchemeType = SchemeID
    End With
    Load m_frmRegReport
    Screen.MousePointer = vbDefault
    Unload frmCancel
    If Not gCancel Then m_frmRegReport.Show 'vbmodal
'    Unload m_frmRegReport
'    Set m_frmRegReport = Nothing
Else
    Set m_frmLoanReport = Nothing
    Set m_frmLoanReport = New frmLoanView
    With m_frmLoanReport
        .ReportType = ReportType
        .FromIndianDate = fromDate
        .ToIndianDate = toDate
        .FromAmount = FromAmount
        .ToAmount = ToAmount
        .Place = Place
        .Caste = Caste
        .Gender = Gender
        .ReportOrder = ReportOrder
        .LoanSchemeType = SchemeID
        .AccountGroup = AccountGroup
    End With
    Load m_frmLoanReport
    
    Screen.MousePointer = vbDefault
    Unload frmCancel
    If gCancel = 2 Then
        Unload m_frmLoanReport
        Set m_frmLoanReport = Nothing
        Exit Sub
    End If
    
    m_frmLoanReport.Show 'vbmodal
    'Set m_frmLoanReport = Nothing
End If

#End If

End Sub


Private Sub m_frmLoanReport_Initialise(Min As Long, Max As Long)

On Error Resume Next
With frmCancel
    If Max <> 0 Then
        .PicStatus.Visible = True
        .Refresh
    End If
End With

End Sub

Private Sub m_frmLoanReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS: " & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
End With
End Sub


Private Sub m_frmSelectReport_ShowReport(SchemeID As Integer, _
        ReportType As wis_LoanReports, ReportOrder As wis_ReportOrder, _
        fromDate As String, toDate As String, LoanPurpose As String, frmRepOption As clsRepOption)

Dim FromAmount As Currency, ToAmount As Currency
Dim Gender As wis_Gender, Place As String
Dim Caste As String, AccountGroup As Integer

With frmRepOption
    FromAmount = .FromAmount
    ToAmount = .ToAmount
    Gender = .Gender
    Place = .Place
    Caste = .Caste
    AccountGroup = .AccountGroupID
    
End With



gCancel = 0
frmCancel.Show
frmCancel.Refresh
Screen.MousePointer = vbHourglass
Set m_frmRegReport = Nothing

If ReportType = repMonthlyRegister Or ReportType = repMonthlyRegisterAll Or _
    ReportType = repShedule_1 Or ReportType = repShedule_2 Or ReportType = repShedule_3 Or _
    ReportType = repShedule_4A Or ReportType = repShedule_4B Or ReportType = repShedule_4C Or _
    ReportType = repShedule_5 Or ReportType = repShedule_6 Then
    
    Set m_frmRegReport = New frmRegReport
    With m_frmRegReport
        .ReportType = ReportType
        .FromIndianDate = fromDate
        .ToIndianDate = toDate
        .FromAmount = FromAmount
        .ToAmount = ToAmount
        .Place = Place
        .Caste = Caste
        .ReportOrder = ReportOrder
        .LoanSchemeType = SchemeID
    End With
    Load m_frmRegReport
    Screen.MousePointer = vbDefault
    Unload frmCancel
    If Not gCancel Then m_frmRegReport.Show 'vbmodal
'    Unload m_frmRegReport
'    Set m_frmRegReport = Nothing
Else
    Set m_frmLoanReport = Nothing
    Set m_frmLoanReport = New frmLoanView
    With m_frmLoanReport
        .ReportType = ReportType
        .FromIndianDate = fromDate
        .ToIndianDate = toDate
        .FromAmount = FromAmount
        .ToAmount = ToAmount
        .Place = Place
        .Caste = Caste
        .Gender = Gender
        .ReportOrder = ReportOrder
        .LoanSchemeType = SchemeID
        .AccountGroup = AccountGroup
        .LoanPurpose = LoanPurpose
    End With
    Load m_frmLoanReport
    
    Screen.MousePointer = vbDefault
    Unload frmCancel
    If gCancel = 2 Then
        Unload m_frmLoanReport
        Set m_frmLoanReport = Nothing
        Exit Sub
    End If
    
    m_frmLoanReport.Show 'vbmodal
    'Set m_frmLoanReport = Nothing
End If

End Sub

Private Sub m_frmSelectReport_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle: m_ParentHandle = 0
Set m_frmSelectReport = Nothing
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmLoanInst Is Nothing Then _
                ObjectHandle = m_frmLoanInst.hwnd: Exit Property
    If Not m_frmLoanMaster Is Nothing Then _
                ObjectHandle = m_frmLoanMaster.hwnd: Exit Property
    If Not m_frmLoanPay Is Nothing Then _
                ObjectHandle = m_frmLoanPay.hwnd: Exit Property
    If Not m_frmLoanReport Is Nothing Then _
                ObjectHandle = m_frmLoanReport.hwnd: Exit Property
    If Not m_frmLoanTrans Is Nothing Then _
                ObjectHandle = m_frmLoanTrans.hwnd: Exit Property
    If Not m_frmScheme Is Nothing Then _
                ObjectHandle = m_frmScheme.hwnd: Exit Property
    If Not m_frmSelectReport Is Nothing Then _
                ObjectHandle = m_frmSelectReport.hwnd: Exit Property
    If Not m_frmRegReport Is Nothing Then _
                Set m_frmRegReport = Nothing
    If Not m_frmLoanReport Is Nothing Then _
                Set m_frmLoanReport = Nothing

End Property


