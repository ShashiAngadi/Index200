Attribute VB_Name = "basBKCC"
Option Explicit

Public Enum wis_BKCCReports
    '''Regular Reportr
    repBKCCLoanBalance = 1
    repBkccLoanHolder
    repBKCCLoanDayBook
    repBKCCLoanIssued
    repBKCCLoanReturned
    repBKCCLoanIntCol
    repBKCCLoanDailyCash
    repBKCCLoanGLedger
    repBkccOD
    repBkccGuarantor
    repBkccLoanMonBal
    
    
    repBkccDepBalance
    repBkccDepHolder
    repBKCCDepDayBook
    repBkccDepIssued
    repBkccDepIntPaid
    repBkccDepDailyCash
    repBkccDepGLedger
    repBkccDepRepMade
    repBkccDepMonBal
    
    repBkccMonTrans
    
    repBKCCMonthlyRegister
    repBKCCShedule_1
    repBKCCShedule_2
    repBKCCShedule_3
    repBKCCShedule_4
    repBKCCShedule_5
    repBKCCShedule_6
    
    repBKCCMemberTrans
    repBkccDepMonTrans
    repBKCCReceivable
    
    repBKCCLoanClaimBill
    repBKCCLoanClaimBill_Yearly
    repBKCCLoanClaimBill_PrevYearly
End Enum


'This Function Returns the Last Transaction Date and Transaction ID
'of the Loan Transaction of the particular account
Private Sub GetLastTransDate(ByVal LoanID As Integer, _
                Optional TransID As Long, Optional TransDate As Date)

Dim rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Integer
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM BKCCTrans " & _
                    " where LoanId = " & LoanID & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM BKCCIntTrans " & _
                    " where LoanId = " & LoanID & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Check int the amount receivable also
Dim AccHeadID As Long
AccHeadID = GetHeadID(GetResourceString(229) & _
                " " & GetResourceString(58), parMemberLoan)
gDbTrans.SqlStmt = "Select * FROM AmountReceivable " & _
                    " where AccHeadID = " & AccHeadID & _
                    " And AccId = " & LoanID & _
                    " Order By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("AccTransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
    If TransDate < rst("TransDate") Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

ErrLine:


End Sub

'This Function Returns the Last Transction Date of The Fd
'of the given KCCloan account Id
' In case there is no transaction it reurns Default date
Public Function GetKCCLastTransDate(ByVal LoanID As Integer) As Date
Dim TransDate As Date
Call GetLastTransDate(LoanID, , TransDate)
GetKCCLastTransDate = TransDate

End Function
'This Function Returns the Max Transction ID of
'the given KCCLOan account Id
'In case there is no transaction it reurns 0
Public Function GetKCCMaxTransID(ByVal LoanID As Integer) As Long
On Error GoTo ExitLine
Dim TransID As Long
Call GetLastTransDate(LoanID, TransID)
GetKCCMaxTransID = TransID
ExitLine:
End Function


Function BKCCRegularInterest(TillDate As Date, ByVal LoanID As Long) As Currency

' Setup error handler.
On Error GoTo Err_line

' Define varaibles.
Dim Lret As Long
Dim InterestAmt As Currency
Dim InterestRate As Single
Dim SlabIntRate As Single

Dim Balance As Long
Dim IntAmount As Currency
Dim Duration As Long
Dim ClsInt As New clsInterest
Dim IntRate As Single

Dim NextDate As String
Dim LastIntDate As Date
Dim LastTransDate As Date

Dim rstMaster As Recordset
Dim rstTrans As Recordset

If LoanID > 0 Then
    gDbTrans.SqlStmt = "SELECT * FROM BKccMaster WHERE loanID = " & LoanID
    Lret = gDbTrans.Fetch(rstMaster, adOpenDynamic)
    If Lret <= 0 Then GoTo Exit_Line
    'Save the resultset for future references.
Else
    Exit Function
End If

gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans WHERE " _
            & "Loanid = " & LoanID & " ORDER BY transID"
Lret = gDbTrans.Fetch(rstTrans, adOpenDynamic)
If Lret <= 0 Then GoTo Exit_Line

If FormatField(rstMaster("LoanClosed")) Then Exit Function

' Get the rate of interest.
InterestRate = Val(FormatField(rstMaster("IntRate")))

'Get the balance amount or/and  date till interest paid for this loan account.
'There is Only one option whehter He has to Consider the Date till int paid
'Or the Interest Balance
Dim rstTemp As Recordset

gDbTrans.SqlStmt = "SELECT TOP 1 Amount,Balance,transdate" & _
                " FROM BKCCTrans WHERE Loanid = " & LoanID & _
                " And TransDate <= #" & TillDate & "# ORDER BY transid DESC"
                
Lret = gDbTrans.Fetch(rstTemp, adOpenDynamic)
If Lret = 0 Then Exit Function
If Lret < 0 Then
    MsgBox "Error retrieving loan details from database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
    
Balance = CCur(FormatField(rstTemp("Balance")))
'If Balance < 0 Then InterestRate = Val(FormatField(RstMaster("DepIntRate")))
If Balance < 0 Then Exit Function

LastTransDate = rstTemp("TransDate")
If FormatField(rstMaster("LastIntDate")) <> "" Then
    LastIntDate = rstMaster("LastIntDate")
Else
    LastIntDate = vbNull
End If

'Get the duration in days from the last transaction date
'until this day.
On Error GoTo Err_line

If LastIntDate = vbNull Then
    gDbTrans.SqlStmt = "SELECT Top 1 * From BKCCIntTrans WHERE LoanID = " & LoanID & _
                    " AND (IntAmount > 0 OR PenalIntAmount > 0)" & _
                    " AND TransDate <= #" & TillDate & "# ORDER BY TransID Desc"
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        LastIntDate = rstTrans("TransDate")
    Else
        LastIntDate = rstMaster("IssueDate")
    End If
End If

'Now Get the Transaction Of the LastIntPaid date
gDbTrans.SqlStmt = "SELECT * FROM BKCCTrans Where LoanID = " & LoanID & _
                " AND TransDate >= #" & LastIntDate & "#" & _
                " AND TransDate <= #" & TillDate & "# ORDER BY TransID"
                
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then
    Duration = DateDiff("d", CDate(LastIntDate), TillDate)
    If Balance > 0 Then
        IntAmount = ((Duration / 365) * (InterestRate / 100) * Balance)
    ElseIf Duration >= 30 Then
        IntAmount = ((Duration / 365) * (InterestRate / 100) * Balance)
    End If
    GoTo Exit_Line
End If
'Set rstTrans = gDBTrans.Rst.Clone


Dim AsOnDate As Date
Dim SlabInt As Boolean
SlabIntRate = GetKCCSlabIntRate(Balance)
If SlabIntRate > 0 Then SlabInt = True

IntAmount = 0
Balance = rstTrans("Balance")
 Do
    
    'If Balance < 0 Then Exit Do 'InterestRate = Val(FormatField(RstMaster("DepIntRate")))
    'If balance move it to the then next positive balance
    'Now Get the interestRate for this amount
    If SlabInt Then SlabIntRate = GetKCCSlabIntRate(Balance)
    InterestRate = IIf(SlabIntRate > 0, SlabIntRate, InterestRate)
    
    If rstTrans.EOF Then
        Duration = DateDiff("d", CDate(LastIntDate), TillDate)
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
    Balance = FormatField(rstTrans("BALANCE"))
    rstTrans.MoveNext
    If Balance < 0 Then
        'IntAmount = 0
        Balance = 0
        rstTrans.MovePrevious
        rstTrans.Find "Balance > 0", , adSearchForward
    End If
Loop
    
GoTo Exit_Line

DepositInterest:

InterestRate = Val(FormatField(rstMaster("DepIntRate")))
    
    
'''m_InterestBalance
Exit_Line:
    BKCCRegularInterest = IntAmount \ 1

    Exit Function

Err_line:
    If Err Then
        MsgBox "ComputeRegularInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
     End If

End Function

'Ths function returns the rate of interest for the amount givet
'In KCc the interest Rate varies from Amount to amount
'So this will get interate the particular Amount
'Created by Shaidhar Angadi  on 11/02/2004,
Public Function GetKCCSlabIntRate(ByVal Amount As Currency) As Single

GetKCCSlabIntRate = 0

'Now Get the SchemNAme For this Amount range
Dim SetupClass As New clsSetup

Dim strSchemeName As String
'Dim strKeyName As String
Dim strSchemeValue As String
Dim MinLimit As Currency
Dim MaxLimit As Currency

Dim I As Integer
Dim pos As Integer

'Error Rtrapping
On Error GoTo Exit_Line

Do
    'strSchemeName =
    strSchemeValue = SetupClass.ReadSetupValue("KCCLoanInt", "Slab" & I, "0")
    
    If strSchemeValue = "0" Then I = 100: Exit Do
    
    'Now Get the Range of this slab
    pos = InStr(1, strSchemeValue, "-", vbBinaryCompare)
    If pos = 0 Then Exit Do
    
    MinLimit = Val(Left(strSchemeValue, pos - 1))
    MaxLimit = Val(Mid(strSchemeValue, pos + 1))
    If MaxLimit = 0 Then MaxLimit = Amount + 1
    
    'Now Check Whether Given amount falls int this Range
    If Amount >= MinLimit And Amount <= MaxLimit Then
        'If it falls in this range then Get the Iinterest Rate for  this
        Exit Do
    End If
    
    I = I + 1
    strSchemeName = ""
Loop

'If ther is no Range exit function
If strSchemeName = "0" Then Exit Function

Dim IntClass As New clsInterest
'Now Get the Interest Rate for this Loan
strSchemeValue = IntClass.InterestRate(wis_BKCCLoan, "Slab" & I, FinUSFromDate)

GetKCCSlabIntRate = Val(strSchemeValue)

Exit Function

Exit_Line:

End Function


Public Function BKCCDepositInterest(LoanID As Long, TransDate As Date) As Currency
Dim rst_PM As Recordset
Dim rst_1TO10 As Recordset
Dim rst_11TO31 As Recordset
Dim DateLimit As String
Dim YearLimit As Integer
Dim Balance As Currency
Dim Balance1TO10 As Currency
Dim Balance11TO31 As Currency
Dim TransID As Long
Dim Dy As Integer
Dim Mon As Integer
Dim MonLimit As Integer
Dim Yr As Integer
Dim Interest As Currency
Dim IntRate As Single
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT * FROM BKCCMaster WHERE LoanID = " & LoanID
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function
IntRate = FormatField(rst("DepIntRate"))

gDbTrans.SqlStmt = "SELECT TOP 1 TransDate, Balance FROM BKCCTrans WHERE" & _
    " LoanID = " & LoanID & "  ORDER BY TransID ASC"
Call gDbTrans.Fetch(rst, adOpenDynamic)

If rst.EOF Then Exit Function
DateLimit = rst("TransDate")

gDbTrans.SqlStmt = "SELECT Top 1 IntAmount, TransDate FROM BKCCIntTrans WHERE" & _
    " LoanID = " & LoanID & " And Deposit = True " & _
    " And TransType < 0 ORDER BY TransID DESC"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Interest = FormatField(rst("IntAmount"))

If Interest > 0 Then DateLimit = rst("TransDate")
'Call gDBTrans.Rst.Close

If DateDiff("d", DateLimit, TransDate) < 30 Then Exit Function

Mon = Month(DateLimit)
MonLimit = Month(TransDate)
Yr = Year(DateLimit)
YearLimit = Year(TransDate)
Dy = Day(TransDate)


If Dy < 30 Then MonLimit = MonLimit - 1

gDbTrans.SqlStmt = "SELECT Max(TransID) AS MaxTransID, Month(transdate) AS Months" & _
    " From BKCCTrans" & _
    " WHERE TransDate BETWEEN #" & DateLimit & "# " & _
    " And #" & TransDate & "# And LoanID = " & LoanID & _
    " GROUP BY Month(TransDate);"
    
Call gDbTrans.Fetch(rst_PM, adOpenDynamic)


'Balance before 10th of this month
gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID, Month(TransDate) as Months" & _
    " from BKCCTrans WHERE Day(TransDate) < 11 and Transdate between" & _
    " #" & DateLimit & "# And #" & TransDate & "# and LoanID = " & LoanID & _
    " And Balance < 0 GROUP BY Month(transdate);"
Call gDbTrans.Fetch(rst_1TO10, adOpenDynamic) ' > 0 Then Set rst_1TO10 = gDBTrans.Rst.Clone

'Minimum Balance between 11th and last Day of this month
gDbTrans.SqlStmt = "SELECT Min(Balance) as MinBalance, Month(TransDate) as Months" & _
    " from BKCCTrans WHERE Day(TransDate) >= 11 and transdate between" & _
    " #" & DateLimit & "# And #" & TransDate & "# and LoanID = " & LoanID & _
    " And Balance < 0 GROUP BY Month(transdate);"
Call gDbTrans.Fetch(rst_11TO31, adOpenDynamic) ' > 0 Then Set rst_11TO31 = gDBTrans.Rst.Clone

Interest = 0
While (Mon >= MonLimit And Yr < YearLimit) Or (Mon <= MonLimit And Yr = YearLimit)

    Balance1TO10 = 0: Balance11TO31 = 0: TransID = 0
     
    rst_PM.MoveFirst
    rst_PM.Find "Months = " & Mon - 1
    If Not rst_PM.EOF Then TransID = rst_PM.Fields("MaxTransID")
    If TransID > 0 Then
        gDbTrans.SqlStmt = "SELECT Balance FROM BKCCTrans WHERE LoanID = " & _
            LoanID & " And TransID = " & TransID
        Call gDbTrans.Fetch(rst, adOpenDynamic)
        Balance = Abs(rst(0))
        TransID = 0
    End If
    
    If Not rst_1TO10 Is Nothing Then
        'rst_1TO10.MoveFirst
        rst_1TO10.Find "Months=" & Mon
        If Not rst_1TO10.EOF Then TransID = rst_1TO10.Fields("MaxTransID")
        If TransID > 0 Then
            gDbTrans.SqlStmt = "SELECT Balance FROM BKCCTrans WHERE LoanID = " & _
                LoanID & " And TransID = " & TransID
            Call gDbTrans.Fetch(rst, adOpenDynamic)
            Balance1TO10 = Abs(rst(0))
            Balance = Balance1TO10
        End If
    End If
    
    If Not rst_11TO31 Is Nothing Then
        'rst_11TO31.MoveFirst
        rst_11TO31.Find "Months=" & Mon
        If Not rst_11TO31.EOF Then Balance11TO31 = Abs(rst_11TO31.Fields("MinBalance"))
    End If
    
    If Balance1TO10 > 0 Then
        If Balance11TO31 > 0 Then Balance = Balance11TO31
    ElseIf Balance11TO31 > 0 Then
        Balance = IIf(Balance <= Balance11TO31, Balance, Balance11TO31)
    End If

    Interest = Interest + (Balance * 1 / 12 * IntRate / 100)
    Mon = Mon + 1
    If Mon > 12 Then Yr = Yr + 1: Mon = 1
    
Wend

Interest = Interest \ 1
BKCCDepositInterest = Interest

End Function

' Calculates the penal interest for defaulted repayments.
Function BKCCPenalInterest(TillDate As Date, ByVal LoanIDNo As Long) As Currency
'If Not DateValidate(TillIndianDate, "/", True) Then Exit Function
' Setup error handler
Err.Clear
On Error GoTo Exit_Line
' Variables...
Dim LoanID As Long
Dim LoanDate As Date
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

'TillDate = FormatDate(TillIndianDate)

'Load the Repayment date to a local variable
If LoanIDNo > 0 Then
    LoanID = LoanIDNo
    gDbTrans.SqlStmt = "SELECT * FROM BKCCMaster WHERE loanID = " & LoanIDNo
    If gDbTrans.Fetch(rstMaster, adOpenDynamic) <= 0 Then GoTo Exit_Line
Else
    Exit Function
End If

If FormatField(rstMaster("LoanClosed")) Then Exit Function
PenaltyRate = FormatField(rstMaster("PenalIntrate"))
If PenaltyRate <= 0 Then Exit Function

'For the BKcc Account The Repaymnet period is of one year
'from the date when the particular amount is issued
'Here in this loan Penal INterest is like installment Amount

'First Check the Amount whether is there any due Amout
'so get the Balance before one year
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT Balance,TransDate,Amount,TransType,TransID " & _
        " From BKCCTrans WHere LoanId = " & LoanID & _
        " AND TransID = (Select Max(transID) From BKCCTrans Where LoanId = " & _
            LoanID & " ANd TransDate < #" & DateAdd("yyyy", -1, TillDate) & "#)"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function
BalanceAmount = FormatField(rst("Balance"))
LoanAmount = FormatField(rst("Amount"))
transType = FormatField(rst("TransType"))
TransID = FormatField(rst("TransID"))
LastDate = rst("TransDate")

If BalanceAmount <= 0 Then Exit Function

'Now Get the Repaid Amount Afetr That date till today
gDbTrans.SqlStmt = "SELECT Sum(amount) as RepaidAmount From BKCCTrans " & _
    " WHere LoanId = " & LoanID & " AND Deposit = False " & _
    " AND TransDate > #" & LastDate & "# AND TransDate <= #" & TillDate & "# " & _
    " AND (TransType = " & wDeposit & " OR  TransType = " & wContraDeposit & ")"

RepaidAmount = 0
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        RepaidAmount = FormatField(rst("RepaidAmount"))
If BalanceAmount - RepaidAmount <= 0 Then Exit Function

ODAmount = BalanceAmount - RepaidAmount

'If the Last transaction is Repayment then
'consider that amount as als loan balance
'and amount as the Repaid Amount
If transType = wContraDeposit Or transType = wDeposit Then
    RepaidAmount = RepaidAmount + LoanAmount
    BalanceAmount = BalanceAmount + LoanAmount
Else
    BalanceAmount = BalanceAmount - LoanAmount
End If

If BalanceAmount - RepaidAmount <= 0 Then
    'OVER DUE IS OF THE LOAN THEN
    ODDate = DateAdd("yyyy", 1, LastDate)
    ODIntAmount = ODAmount * (DateDiff("D", ODDate, TillDate) / 365) * (PenaltyRate / 100)
    BKCCPenalInterest = ODIntAmount \ 1
    GoTo Exit_Line
End If

gDbTrans.SqlStmt = "SELECT Balance,TransDate,Amount,TransType,TransID" & _
    " FROM BKCCTrans WHere LoanId = " & LoanID & _
    " AND TransID <= " & TransID & " ORDER By TransId Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    Do While BalanceAmount - RepaidAmount > 0
        'gDbTrans.SQLStmt = "SELECT Top 1 Balance,TransDate,Amount,TransType,TransID" & _
            " FROM BKCCTrans WHere LoanId = " & LoanID & _
            " AND TransID <= " & TransID - 1 & " ORDER By TransId Desc"
        'If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Do
        If rst.EOF Then Exit Do
        
        LoanAmount = FormatField(rst("Amount"))
        transType = FormatField(rst("TransType"))
        LastDate = rst("TransDate")
        TransID = FormatField(rst("TransID"))
        BalanceAmount = FormatField(rst("Balance"))
        
        If transType = wContraDeposit Or transType = wDeposit Then
            RepaidAmount = RepaidAmount + LoanAmount
            'Get the Lan Blance Before This Transction
            BalanceAmount = BalanceAmount + LoanAmount
        Else
            'Get the Lan Blance Before This Transction
            BalanceAmount = BalanceAmount - LoanAmount
        End If
        rst.MoveNext
    Loop
End If

BalanceAmount = FormatField(rst("Balance"))
LastDate = LastDate
gDbTrans.SqlStmt = "SELECT * From BKCCTrans " & _
    " WHERE LoanId = " & LoanID & " AND TransDate >= #" & LastDate & "# " & _
    " AND TransDate <= #" & TillDate & "# AND Deposit = False ORDER BY TransID"
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)

'The Amount with drwan After last date  is over due
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

BKCCPenalInterest = ODIntAmount \ 1
Exit_Line:
    If Err Then
        MsgBox "ComputePenalInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        'Resume
    End If
End Function

Public Function GetReceivAbleAmount(ByVal AccHeadID As Long, ByVal AccId As Long) As Long
    
Dim TransID As Long
Dim rst As Recordset

'Now Get the MAxTransID & Balance
gDbTrans.SqlStmt = "Select Balance From AmountReceivAble" & _
            " WHERE AccHeadID = " & AccHeadID & _
            " And AccID = " & AccId & _
            " Order By TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
'If It HAs Got the Balance Then Return the
    'TransID 'Other Wise Not
    GetReceivAbleAmount = FormatField(rst("Balance"))
End If

End Function


Public Function GetReceivAbleAmountID(ByVal AccHeadID As Long, ByVal AccId As Long) As Long
    
Dim TransID As Long
Dim rst As Recordset

'Now Get the MAxTransID & Balance
gDbTrans.SqlStmt = "Select * From AmountReceivAble" & _
            " WHERE AccHeadID = " & AccHeadID & _
            " And AccID = " & AccId & _
            " Order By TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'If It HAs Got the Balance Then Return the
    'TransID 'Other Wise Not
    If FormatField(rst("Balance")) > 0 Then GetReceivAbleAmountID = rst("TransID")
    
End If

End Function



Public Function GetMaxReceivAbleID(ByVal AccHeadID As Long, AccId As Long) As Long
    
Dim TransID As Long
Dim rst As Recordset

'Now Get the MAxTransID & Balance
gDbTrans.SqlStmt = "Select * From AmountReceivAble" & _
            " WHERE AccHeadID = " & AccHeadID & _
            " And AccID = " & AccId & _
            " Order Bt TransID Desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
    GetMaxReceivAbleID = FormatField(rst("TransID"))
    

End Function


Public Function LoadReceivableAmounts(AccHeadID As Long, AccId As Long) As clsReceive
    Dim rstTemp As Recordset
    Dim TransID As Long
    Dim clsTemp As clsReceive
    'Get the MAx Transaction Id When the last transaction has become null
    gDbTrans.SqlStmt = "select MAx(TransID) From AmountReceivAble" & _
                " WHERE AccHeadID = " & AccHeadID & _
                " AND AccID = " & AccId & _
                " And Balance = 0 "
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                                TransID = FormatField(rstTemp(0))
    gDbTrans.SqlStmt = "Select * From AmountReceivable" & _
                " WHERE AccHeadID = " & AccHeadID & _
                " AND AccID = " & AccId & _
                " And TransID > " & TransID
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        Set clsTemp = New clsReceive
        While Not rstTemp.EOF
            Call clsTemp.AddHeadAndAmount(rstTemp("DueHeadID"), rstTemp("Amount"))
            rstTemp.MoveNext
        Wend
    End If
    
    Set LoadReceivableAmounts = clsTemp
    
End Function


