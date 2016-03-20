Attribute VB_Name = "LoanTransfer"
Option Explicit
'Public gLangOffset As Integer

'This Bas file is made to transfer the Index 2000
'data base to the this existing loansdatabase

Function TransferLoan(OldDBName As String, NewDBName As String) As Boolean

Dim OldTrans As New clsTransact
Dim NewTrans As New clsTransact


If Not OldTrans.OpenDB(OldDBName, "PRAGMANS") Then
    MsgBox " No Index Db"
    Exit Function
End If

If Not NewTrans.OpenDB(NewDBName, "WIS!@#") Then
    OldTrans.CloseDB
    MsgBox " No Index Db"
    Exit Function
End If

Dim dbIndex As clsTransact

If Not SchemeTransfer(OldTrans, NewTrans) Then Exit Function
If Not LoanMasterTransfer(OldTrans, NewTrans) Then Exit Function
If Not LoanTransTransfer(OldTrans, NewTrans) Then Exit Function
    
TransferLoan = True

End Function



Private Function LoanMasterTransfer(oldLoanTrans As clsTransact, NewLOanTrans As clsTransact) As Boolean
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerID As Long
Dim MemID As Long
Dim AccNum As String
Dim BankID As Long
Dim InstMode As Integer
Dim InstAmount As Currency
Dim NoOfINstall As Integer
Dim RecNo As Integer

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Database of loans

SqlStr = "SELECT A.*,B.CustomerID FROM LoanMaster A, MMMaster B " & _
    " WHERE B.AccID=A.MemberID AND " & _
        " A.SchemeID In (SELECT SchemeID FRom LoanTypes Where BKCC = False )" & _
    " ORDER BY SchemeID,LoanID"
oldLoanTrans.SQLStmt = SqlStr
If oldLoanTrans.SQLFetch < 1 Then Exit Function
Set RstIndex = oldLoanTrans.Rst.Clone

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

NewLOanTrans.BeginTrans
While Not RstIndex.EOF
    
    'first get the mem Id of the old db
     MemID = FormatField(RstIndex("MemberId"))
     Debug.Assert MemID <> 927
     
    'now get the customer Of this Member
    CustomerID = 0
    InstMode = 0: NoOfINstall = 0: InstAmount = 0
    InstMode = FormatField(RstIndex("InstalmentMode"))
    'Set oldLoanTrans.Rst = Nothing
    'SqlStr = "SELECT CustomerID from MMMaster where AccID = " & MemId
    'oldLoanTrans.SqlStmt = SqlStr
    'If oldLoanTrans.SQLFetch > 0 Then
    CustomerID = FormatField(RstIndex("CustomerId")) 'FormatField(oldLoanTrans.Rst(0))
    'Else
        'GoTo NextAccount
    'End If
    'RstIndex.AbsolutePosition = RecNo
    AccNum = FormatField(RstIndex("LoanAccNo"))
    If Trim(AccNum) = "" Then _
        AccNum = FormatField(RstIndex("SchemeID")) & "_" & FormatField(RstIndex("LoanId"))
    On Error Resume Next
    
    If InstMode > 0 Then
        If FormatField(RstIndex("InstalmentAmt")) < 10 Then
            
        Else
            InstAmount = FormatField(RstIndex("InstalmentAmt"))
            If InstAmount = 0 Then
                NoOfINstall = 0
            Else
                NoOfINstall = FormatField(RstIndex("LoanAmt")) / InstAmount
            End If
            If NoOfINstall > 2000 Then NoOfINstall = 0
            If NoOfINstall = 1 Then NoOfINstall = 0: InstAmount = 0: InstMode = 0
        End If
    End If
    On Error GoTo ErrLine
    SqlStr = "INSERT INTO LoanMaster (" & _
            " LoanId,SchemeID, MemID,CustomerID," & _
            " AccNUm,Issuedate,LoanDueDate," & _
            " PledgeItem,Pledgevalue,LoanAmount," & _
            " InstMode,InstAmount, NooFInstall, " & _
            " EMI, Intrate,PenalIntRate, " & _
            " Guarantor1,Guarantor2,LoanClosed, Remarks) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            RstIndex("SchemeId") & "," & _
            MemID & ", " & CustomerID & "," & _
            AddQuotes(AccNum, True) & "," & _
            "#" & RstIndex("IssueDate") & "#," & _
            "#" & RstIndex("LoanDueDate") & "#," & _
            AddQuotes(FormatField(RstIndex("PledgeDescription")), True) & "," & _
            FormatField(RstIndex("PledgeValue")) & "," & _
            FormatField(RstIndex("LoanAmt")) & "," & _
            InstMode & ", " & InstAmount & "," & _
            NoOfINstall & ", False, " & _
            FormatField(RstIndex("InterestRate")) & "," & _
            FormatField(RstIndex("PenalInterestrate")) & "," & _
            FormatField(RstIndex("GuarantorId1")) & "," & _
            FormatField(RstIndex("GuarantorId2")) & "," & _
            FormatField(RstIndex("LoanClosed")) & "," & _
            AddQuotes(FormatField(RstIndex("Remarks")), True) & ")"

    NewLOanTrans.SQLStmt = SqlStr
    If Not NewLOanTrans.SQLExecute Then    'THere aRE MORE FIELDS
        NewLOanTrans.RollBack              'IN lOANS THAN iNDEX
        Exit Function
    End If
    'If loan has got the installment then insert the installment details
    If InstMode > 0 Then
        'so we have consder daily loan installment in the new loans
        'and that is not the case with index 2000 so increase insttype by one
        NewLOanTrans.CommitTrans
        InstMode = InstMode + 1
        If Not SaveInstallmentDetails(NewLOanTrans, RstIndex("LoanID"), InstMode, NoOfINstall, FormatField(RstIndex("LoanAmt")), _
                    InstAmount, FormatField(RstIndex("IssueDate"))) Then
            MsgBox "Unable to save the installment details of the " & RstIndex("LoanID")
            Exit Function
        End If
        NewLOanTrans.BeginTrans
    End If
NextAccount:
    RecNo = RecNo + 1
    RstIndex.MoveNext
Wend

NewLOanTrans.CommitTrans

LoanMasterTransfer = True
Exit Function

ErrLine:
Debug.Assert Err.Number = 0
Resume
End Function



Private Function SaveInstallmentDetails(NewLOanTrans As clsTransact, LoanID As Long, InstMode As Integer, NoOfInst As Integer, LoanAmount As Currency, InstAmount As Currency, IssueIndianDate As String) As Boolean

Dim InstNo As Integer
Dim lpCount As Integer
Dim SqlStr As String
Dim Rst As Recordset
Dim NextDate As Date
Dim FortNight As Boolean
Dim TotalInstAmount As Currency

NextDate = FormatDate(IssueIndianDate)
NewLOanTrans.BeginTrans
InstNo = 1
Do
     If InstNo > NoOfInst Then Exit Do
     If TotalInstAmount >= LoanAmount Then Exit Do
     'Get The Next INstallment date
     If InstMode = Inst_Daily Then NextDate = DateAdd("d", 1, NextDate)
     If InstMode = Inst_Weekly Then NextDate = DateAdd("WW", 1, NextDate)
     If InstMode = Inst_FortNightly Then
         If FortNight Then
             FortNight = False
             NextDate = DateAdd("d", 15, NextDate)
         Else
             FortNight = True
             NextDate = DateAdd("d", -15, NextDate)
             NextDate = DateAdd("m", 1, NextDate)
         End If
     End If
     If InstMode = Inst_Monthly Then NextDate = DateAdd("M", 1, NextDate)
     If InstMode = Inst_BiMonthly Then NextDate = DateAdd("m", 2, NextDate)
     If InstMode = Inst_Quartery Then NextDate = DateAdd("q", 1, NextDate)
     If InstMode = Inst_HalfYearly Then
         If FortNight Then
             FortNight = False
             NextDate = DateAdd("M", 6, NextDate)
         Else
             FortNight = True
             NextDate = DateAdd("M", -6, NextDate)
             NextDate = DateAdd("YYYY", 1, NextDate)
         End If
     End If
     If InstMode = Inst_Yearly Then NextDate = DateAdd("YYYY", 1, NextDate)
     
     'WRITE Into the databsae
    SqlStr = "INSERT INTO LoanInst (LoanID,InstNo," & _
            " InstDate,InstAmount,InstBalance )" & _
         " Values ( " & _
         LoanID & "," & _
         InstNo & "," & _
         " #" & NextDate & "#," & _
         InstAmount & "," & _
         InstAmount & " ) "
    NewLOanTrans.SQLStmt = SqlStr
    If Not NewLOanTrans.SQLExecute Then
       NewLOanTrans.RollBack
       Exit Function
    End If
     TotalInstAmount = TotalInstAmount + InstAmount
     InstNo = InstNo + 1
Loop
NewLOanTrans.CommitTrans

SaveInstallmentDetails = True
End Function


Private Function LoanTransTransfer(oldLoanTrans As clsTransact, NewLOanTrans As clsTransact) As Boolean
'Dim
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerID As Long
Dim MemID As Long
Dim BankID As Long
Dim TransID As Long
Dim TransType As Integer
Dim Particualrs As String
Dim ItIsIntTrans As Boolean
Dim LoanID As Long
Dim OldLoanId As Long
Dim RegInt As Double
Dim PenalInt As Double
Dim InstType As Integer
Dim RstInst As Recordset
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency

Dim BKCC As Boolean

'BankID = Val(InputBox("Enter the Id for Bank in numeric", "Bank Identification"))


'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM LoanTrans ORDER By LoanID, TransID"
oldLoanTrans.SQLStmt = SqlStr

If oldLoanTrans.SQLFetch < 1 Then
    BKCC = True
    SqlStr = "SELECT * FROM BKCCTrans ORDER By LoanID, TransID"
    oldLoanTrans.SQLStmt = SqlStr
    If oldLoanTrans.SQLFetch < 1 Then GoTo ExitLine
End If
    
Set RstIndex = oldLoanTrans.Rst.Clone

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id
    
Particualrs = "Penal interest"  ' Extarcted from Data BAse Differnet for Kannada
If UCase(Trim(InputBox("Enter language either ENGLISH or KANNADA", "Loan Transfer", "KANNADA"))) = "KANNADA" Then
    Particualrs = LoadResString(5345)
Else
    Particualrs = LoadResString(345)
End If

Dim InTrans As Boolean

TransactionTransfer:

While Not RstIndex.EOF
    ItIsIntTrans = False
    TransType = FormatField(RstIndex("TransType"))
    OldLoanId = RstIndex("LoanID")
    Debug.Assert RstIndex("LoanID") <> 607
    If LoanID <> FormatField(RstIndex("LoanID")) Then
        Set RstInst = Nothing
        TransID = 1
        LoanID = RstIndex("LoanID")
        'Get the insttalmente type
        
        NewLOanTrans.SQLStmt = "SELECT InstMode From LoanMaster Where LoanID = " & LoanID
        If NewLOanTrans.SQLFetch < 1 Then
            MsgBox "Insufficient information of Instalment = " & LoanID
            RstIndex.FindLast "LoanID = " & LoanID
            GoTo NextAccount
            
            'Exit Function
            InstType = 0
        Else
            InstType = FormatField(NewLOanTrans.Rst("InstMode"))
        End If
    End If
    
    If InstType > 0 Then
        Set RstInst = Nothing
        SqlStr = "SELECT * FROM LoanInst Where LoanID = " & LoanID & _
            " AND InstBalance > 0 ORDER BY InstDate"
        NewLOanTrans.SQLStmt = SqlStr
        If NewLOanTrans.SQLFetch < 0 Then
            MsgBox "Error in loan installment of " & LoanID
            Exit Function
        End If
        Set RstInst = NewLOanTrans.Rst.Clone
    End If
    
    'Begin the transaction
    NewLOanTrans.BeginTrans
    InTrans = True
    TransID = TransID + 1
    
    If TransType = -2 Or TransType = 2 Then
        ItIsIntTrans = True
        RegInt = FormatField(RstIndex("Amount"))
        If InStr(1, FormatField(RstIndex("Particulars")), Particualrs, vbTextCompare) Then
            PenalInt = FormatField(RstIndex("Amount"))
            RstIndex.MoveNext
            RegInt = FormatField(RstIndex("Amount"))
            If RstIndex("TransType") = 1 Or RstIndex("TransType") = -1 Then
                RstIndex.MovePrevious: RegInt = 0
            End If
        End If
        TransType = IIf(TransType < 0, wDeposit, wWithDraw)
        SqlStr = "INSERT INTO LoanIntTrans (" & _
                " LoanId,TransDate," & _
                " TransID,TransType," & _
                " IntAmount,PenalIntAmount, IntBalance)"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("LoanId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & Abs(TransType) / TransType & "," & _
                RegInt & "," & _
                PenalInt & "," & _
                FormatField(RstIndex("Balance")) & _
                ")"
        RegInt = 0: PenalInt = 0
        NewLOanTrans.SQLStmt = SqlStr
        If Not NewLOanTrans.SQLExecute Then
            NewLOanTrans.RollBack
            InTrans = False
            Exit Function
        End If
        RstIndex.MoveNext
        If RstIndex.EOF Then
            NewLOanTrans.CommitTrans
            InTrans = False
            GoTo NextAccount
        End If
    End If
    
    TransType = FormatField(RstIndex("TransType"))
    If Not (TransType = wWithDraw Or TransType = wDeposit) Then
        If TransType = 7 Or TransType = -7 Then
            Amount = Amount * -1
        Else
            NewLOanTrans.CommitTrans
            InTrans = False
            RstIndex.MovePrevious
            GoTo NextAccount
            
        End If
    End If
    Amount = FormatField(RstIndex("Amount"))
    
    SqlStr = "INSERT INTO LoanTrans (" & _
            " LoanId,TransDate," & _
            " TransID,TransType," & _
            " Amount, Balance, Particulars)"
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            "#" & RstIndex("TransDate") & "#," & _
            TransID & "," & TransType & "," & _
            Amount & "," & _
            FormatField(RstIndex("Balance")) & "," & _
            AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
            If TransID = 1 Then TransID = 2
            
    NewLOanTrans.SQLStmt = SqlStr
    
    If Not OldLoanId = RstIndex("LoanID") Then SqlStr = ""
    If Not Amount <> 0 Then SqlStr = ""
    
    If SqlStr <> "" Then
        If Not NewLOanTrans.SQLExecute Then
            NewLOanTrans.RollBack
            InTrans = False
            Exit Function
        End If
    End If
    
    If Not RstInst Is Nothing And (TransType = wContraDeposit Or TransType = wDeposit) Then
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
                ", PaidDate = #" & RstIndex("TransDate") & "#" & _
                " WHERE LoanID = " & LoanID & _
                " AND InstNo = " & InstNo
            NewLOanTrans.SQLStmt = SqlStr
            If Not NewLOanTrans.SQLExecute Then
                NewLOanTrans.RollBack
                Exit Function
            End If
            RstInst.MoveNext
        Loop
    End If
    
    NewLOanTrans.CommitTrans
    InTrans = False

NextAccount:
    Debug.Assert InTrans = False
    RstIndex.MoveNext
Wend

BKCCTransfer:
If Not BKCC Then
    BKCC = True
    SqlStr = "SELECT * FROM BKCCTrans ORDER By LoanID, TransID"
    oldLoanTrans.SQLStmt = SqlStr
    If oldLoanTrans.SQLFetch < 1 Then GoTo ExitLine
    
    Set RstIndex = oldLoanTrans.Rst.Clone
 '   GoTo TransactionTransfer
End If


ExitLine:
Debug.Print Now
LoanTransTransfer = True
Exit Function
ErrLine:
Debug.Assert Err.Number = 0
'Resume
End Function


Private Function SchemeTransfer(oldLoanTrans As clsTransact, NewLOanTrans As clsTransact) As Boolean
'Now transfer the Loans Schemes
Dim RstIndex As Recordset
'Dim DbIndex As clsTransact
oldLoanTrans.SQLStmt = "SELECT  * From LoanTypes"
If oldLoanTrans.SQLFetch < 1 Then
    MsgBox "No Loan Schemes"
    Exit Function
End If
Set RstIndex = oldLoanTrans.Rst.Clone

Dim SqlStr As String
Dim SchemeId As Integer
Dim SchemeName As String
Dim LoanCategary As wisLoanCategories
Dim TermType As Integer
Dim LoanType As Integer
LoanType = 1
Dim Monthduration As Integer
Dim DayDuration As Integer
Dim MaxRepayments As Integer
Dim InterestRate As Single
Dim PenalInterestRaste As Single
Dim InsurenceFee As Currency
Dim LegalFee As Currency
Dim Description  As String
Dim CreateDate As Date

CreateDate = Now

NewLOanTrans.BeginTrans
While Not RstIndex.EOF
    On Error Resume Next
    SchemeId = RstIndex("SchemeID")
    SchemeName = RstIndex("SchemeName")
    LoanType = 1
    LoanCategary = RstIndex("Category")
    TermType = FormatField(RstIndex("TermType"))
    CreateDate = RstIndex("Createdate")
    Monthduration = FormatField(RstIndex("MaxRepaymentTime")) * 12
    DayDuration = 0
    
    If Not Monthduration Then Monthduration = 4
    If CreateDate = Null Then CreateDate = Now
    
    On Error GoTo ErrLine
    
    SqlStr = "INSERT INTO LoanScheme (SchemeID, SchemeName," & _
            " Category,TermType,LoanType, MonthDuration," & _
            " DayDuration,Intrate, PenalIntrate," & _
            " LOanPurpose, InsuranceFee,LegalFee," & _
            " Description,Createdate ) "
    SqlStr = SqlStr & " Values (" & _
            SchemeId & "," & AddQuotes(SchemeName, True) & "," & _
            LoanCategary & ", " & TermType & "," & _
            LoanType & "," & Monthduration & "," & _
            DayDuration & "," & FormatField(RstIndex("InterestRate")) & _
            "," & FormatField(RstIndex("PenalInterestRate")) & "," & _
            " 'Individual'," & FormatField(RstIndex("InsuranceFee")) & "," & _
            FormatField(RstIndex("LegalFee")) & ",'Loan Name'," & _
            "#" & CreateDate & "#)"
    NewLOanTrans.SQLStmt = SqlStr
    If Not NewLOanTrans.SQLExecute Then
        NewLOanTrans.RollBack
        Exit Function
    End If
    
    RstIndex.MoveNext
Wend
NewLOanTrans.CommitTrans
SchemeTransfer = True
ErrLine:
Debug.Assert Err.Number = 0
End Function


