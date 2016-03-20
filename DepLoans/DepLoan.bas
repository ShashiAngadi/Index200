Attribute VB_Name = "basDepLoans"
Option Explicit

Public Enum wis_DepLoanReports
    repDepLnBalance = 1
    repDepLnTransaction = 2
    repDepLnDetail = 3
    repDepLnOverDue = 4
    repDepLnAccOpen = 5
    repDepLnAccClose = 6
    repDepLnGenLedger = 7
    repDepLnCashBook = 8
    repDepLnMonthlyBalance
    repDepSubDayBook
End Enum

' Gets the Deposit Type .
'Arguments DepositName
Public Function GetDepositType(DepName As String) As Integer
Dim rstDeposit As Recordset
Dim Retval As Integer
'
On Error GoTo ErrLine
Retval = 0
Select Case DepName
    Case ""
        Retval = 0
    Case GetResourceString(424) '"Recurring Deposit"
        Retval = wisDeposit_RD
    Case GetResourceString(425) '"Pigmy Deposit
        Retval = wisDeposit_PD ' 3
    Case Else
        gDbTrans.SqlStmt = "SELECT * From DepositName Where " & _
            " DepositName = " & AddQuotes(DepName, True)
        If gDbTrans.Fetch(rstDeposit, adOpenDynamic) > 0 Then
            While Not rstDeposit.EOF
                If UCase(DepName) = UCase(FormatField(rstDeposit("DepositName"))) Then
                    Retval = FormatField(rstDeposit("DepositID"))
                    'after theis we have to exit from this loop
                    'To exit from this loop we are
                    'moving the recordset to last position
                    rstDeposit.MoveLast
                End If
                rstDeposit.MoveNext
            Wend
        End If
End Select

'If Retval Then Retval = Retval + wis_Deposits

GetDepositType = Retval
Exit Function

ErrLine:
MsgBox Err.Number & vbCrLf & Err.Description
End Function

Public Function GetDepositTypeText(DepositType As Integer) As String
    Dim DepositNameEnglish As String
    GetDepositTypeText = GetDepositTypeTextEnglish(DepositType, DepositNameEnglish)
End Function
Public Function GetDepositTypeTextEnglish(DepositType As Integer, DepositNameEnglish) As String

If DepositType > wis_Deposits Then DepositType = DepositType - wis_Deposits

On Error GoTo ErrLine
Select Case DepositType
    Case 0
        GetDepositTypeTextEnglish = GetResourceString(43)
        DepositNameEnglish = LoadResString(43)
    Case wisDeposit_RD
        GetDepositTypeTextEnglish = GetResourceString(424)
        DepositNameEnglish = LoadResString(424)
    Case wisDeposit_PD
        GetDepositTypeTextEnglish = GetResourceString(425)
        DepositNameEnglish = LoadResString(425)
        
    Case Else
        GetDepositTypeTextEnglish = " "
        Dim rstDeposit As Recordset
        If DepositType > wis_Deposits Then DepositType = DepositType - wis_Deposits
        gDbTrans.SqlStmt = "SELECT * From DepositName Where " & _
            " DepositID = " & DepositType
        If gDbTrans.Fetch(rstDeposit, adOpenDynamic) > 0 Then
            GetDepositTypeTextEnglish = FormatField(rstDeposit("DepositName"))
            DepositNameEnglish = FormatField(rstDeposit("DepositName"))
        End If

End Select
Exit Function

ErrLine:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function ComputeDepLoanInterstBalance(LoanID As Long, AsOnDate As Date) As Currency

On Error GoTo Err_line

Dim IntBalance As Currency
Dim RstLoanTrans As Recordset

If LoanID = 0 Then GoTo Exit_line

gDbTrans.SqlStmt = "SELECT * From LoanTrans WHERE " & _
    " TransDate <= #" & AsOnDate & "# ORDER By TransID Desc"
    
If gDbTrans.Fetch(RstLoanTrans, adOpenDynamic) < 1 Then GoTo Exit_line

If RstLoanTrans Is Nothing Then GoTo Exit_line
RstLoanTrans.MoveFirst

Do
    If RstLoanTrans.BOF Or RstLoanTrans.EOF Then Exit Do
    IntBalance = FormatField(RstLoanTrans("Balance"))
    If RstLoanTrans("Amount") > 0 Then Exit Do
    RstLoanTrans.MoveNext
Loop

Exit_line:
ComputeDepLoanInterstBalance = IntBalance
Set RstLoanTrans = Nothing

Exit Function
Err_line:
    If Err.Number Then MsgBox "Error in ComputeInterestBalance :" & Err.Description & _
         vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Err.Clear
    IntBalance = 0
    GoTo Exit_line
End Function


Public Function ComputeDepLoanRegularInterest(AsOnDate As Date, LoanID As Long) As Currency
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
Dim Rst As ADODB.Recordset

gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE " _
            & "loanID = " & LoanID
Lret = gDbTrans.Fetch(rstLoanMast, adOpenStatic)
If Lret <= 0 Then GoTo Exit_line

' Save the resultset for future references.

IntRate = Val(FormatField(rstLoanMast("InterestRate")))

If FormatField(rstLoanMast("LastIntdate")) = "" Then
ReCheck:
    LastintPaidDate = Now
    gDbTrans.SqlStmt = "SELECT * From DepositLoanIntTrans WHERE " & _
            " LoanID = " & LoanID & " AND TransDate <= #" & AsOnDate & "#" & _
            " ORDER BY TransID Desc "
    If gDbTrans.Fetch(RstLoanTrans, adOpenStatic) Then
        LastintPaidDate = RstLoanTrans("TransDate")
    Else
        gDbTrans.SqlStmt = "SELECT * From DepositLoanTrans WHERE " & _
            " LoanID = " & LoanID & " ORDER BY TransID "
        If gDbTrans.Fetch(Rst, adOpenForwardOnly) Then LastintPaidDate = Rst("TransDate")
    End If
Else
    LastintPaidDate = rstLoanMast("LastIntdate")
End If

    
'Now Get the Transaction Of the LastIntPaid date
gDbTrans.SqlStmt = "SELECT * FROM DepositLoanTrans Where LoanID = " & LoanID & _
    " AND TransDate >= #" & LastintPaidDate & "# ORDER BY TransID"
If gDbTrans.Fetch(RstLoanTrans, adOpenStatic) < 1 Then
    gDbTrans.SqlStmt = "SELECT * FROM DepositLoanTrans Where LoanID = " & LoanID & " ORDER BY TransID Desc"
    Call gDbTrans.Fetch(RstLoanTrans, adOpenStatic)
    Balance = RstLoanTrans("Balance")
    Duration = DateDiff("d", LastintPaidDate, AsOnDate)
    IntAmount = IntAmount + ((Duration / 365) * (IntRate / 100) * Balance)
    GoTo Exit_line
End If
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

    ComputeDepLoanRegularInterest = IntAmount \ 1

    
Exit_line:
    Set Rst = Nothing
    Set rstLoanMast = Nothing
    Set RstLoanTrans = Nothing
    ComputeDepLoanRegularInterest = IntAmount \ 1
    Exit Function

Err_line:
    If Err Then
        MsgBox "ComputeRegularInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
     End If
End Function


' Calculates the penal interest for defaulted repayments.
Public Function ComputeDepLoanPenalInterest(AsOnDate As Date, Optional LoanIDNo As Long) As Currency

'If Not DateValidate(AsOnIndianDate, "/", True) Then Exit Function
'Dim AsOnDate As Date
'AsOnDate = FormatDate(AsOnIndianDate)

' Setup error handler
Err.Clear
On Error GoTo Exit_line
' Variables...
Dim LoanID As Long
Dim LastDate As Date
Dim LoanDueDate As Date
Dim Lret As Long
Dim PenaltyRate As Single
Dim PenalAmount As Currency
Dim BalanceAmount As Currency
Dim rstLoanMast As Recordset
Dim rstTrans As Recordset
Dim Days As Integer


'Load the Repayment date to a local variable
    gDbTrans.SqlStmt = "SELECT * FROM DepositLoanMaster WHERE " _
        & "loanID = " & LoanIDNo
    Lret = gDbTrans.Fetch(rstLoanMast, adOpenStatic)
    If Lret <= 0 Then GoTo Exit_line

'Get Loan Date & LoanDueDate
LoanDueDate = rstLoanMast("LoanDueDate")
'm_DepositType = FormatField(rstLoanMast("DepositType"))

Days = DateDiff("d", LoanDueDate, AsOnDate)
If Days <= 0 Then Exit Function

' Save the loanid to a local variable.
LoanID = Val(FormatField(rstLoanMast("LoanID")))
If LoanID = 0 Then LoanID = LoanIDNo
' Get the rate of penalty.
PenaltyRate = Val(FormatField(rstLoanMast("PenalInterestRate")))

gDbTrans.SqlStmt = "SELECT Top 1 * FROM DepositLoanTrans WHERE " _
    & "loanID = " & LoanID & " AND TransID = (SELECT Max(TransID) " & _
        " From DepositLoanTrans Where LoanID = " & LoanID & _
        " And TransDate <= #" & AsOnDate & "# )"
    
Lret = gDbTrans.Fetch(rstTrans, adOpenForwardOnly)
If Lret <= 0 Then GoTo Exit_line

'Get total Loan Amount
BalanceAmount = CCur(FormatField(rstTrans("Balance")))
If BalanceAmount = 0 Then GoTo Exit_line
LastDate = rstTrans("TransDate")

PenalAmount = 0
PenalAmount = BalanceAmount * (PenaltyRate / 100) * (Days / 365)

ComputeDepLoanPenalInterest = IIf(PenalAmount < 0, 0, PenalAmount \ 1)

Exit_line:
    Set rstLoanMast = Nothing
    Set rstTrans = Nothing
    If Err Then
        MsgBox "ComputePenalInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
    
End Function


