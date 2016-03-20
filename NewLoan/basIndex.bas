Attribute VB_Name = "basIndex"
Option Explicit
'Public gLangOffset As Integer

'This Bas file is made to transfer the Index 2000
'data base to the this existing loansdatabase

Function IndexTransferred(dbFile As String) As Boolean
Dim DbIndex As New clsTransact

If Not DbIndex.OpenDB(dbFile, "PRAGMANS") Then
    MsgBox " No Index Db"
    Exit Function
End If

If Not SchemeTransfer(DbIndex) Then Exit Function
If Not LoanMasterTransfer(DbIndex) Then Exit Function
If Not LoanTransTransfer(DbIndex) Then Exit Function
    
IndexTransferred = True

End Function



Private Function LoanMasterTransfer(DbIndex As clsTransact) As Boolean
'Dim DbIndex As clsTransact
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim LoanAccNo As String
Dim BankID As Long
Dim InstMode As Integer
Dim InstAmount As Currency
Dim NoOfINstall As Integer
Dim RecNo As Integer

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM LoanMaster ORDER BY SchemeID,LoanID"
DbIndex.SqlStmt = SqlStr
If DbIndex.SQLFetch < 1 Then Exit Function
Set RstIndex = DbIndex.Rst.Clone

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

gDbTrans.BeginTrans
While Not RstIndex.EOF
    
    'first get the mem Id of the old db
    MemID = FormatField(RstIndex("memberId"))
    'now get the customer Of this Member
    CustomerId = 0
    InstMode = 0: NoOfINstall = 0: InstAmount = 0
    InstMode = FormatField(RstIndex("InstalmentMode"))
    Set DbIndex.Rst = Nothing
    SqlStr = "SELECT CustomerID from MMMaster where AccID = " & MemID
    DbIndex.SqlStmt = SqlStr
    If DbIndex.SQLFetch > 0 Then CustomerId = FormatField(DbIndex.Rst(0))
    RstIndex.AbsolutePosition = RecNo
    LoanAccNo = FormatField(RstIndex("SchemeID")) & "_" & FormatField(RstIndex("LoanId"))
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
    SqlStr = "INSERT INTO LoanMAster (" & _
            " LoanId,SchemeID, CustomerID," & _
            " LoanAccNo,Issuedate,LoanDueDate," & _
            " PledgeItem,Pledgevalue,LoanAmount," & _
            " InstMode,InstAmount, NooFInstall, " & _
            " EMI, Intrate,PenalIntRate, " & _
            " Guarantor1,Guarantor2,LoanClosed, Remarks) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            RstIndex("SchemeId") & "," & _
            CustomerId & "," & _
            AddQuotes(LoanAccNo, True) & "," & _
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

    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then    'THere aRE MORE FIELDS
        gDbTrans.RollBack              'IN lOANS THAN iNDEX
        Exit Function
    End If
    'If loan has got the installment then insert the installment details
    If InstMode > 0 Then
        'so we have consder daily loan installment in the new loans
        'and that is not the case with index 2000 so increase insttype by one
        gDbTrans.CommitTrans
        InstMode = InstMode + 1
        If Not SaveInstallmentDetails(FormatField(RstIndex("LoanID")), InstMode, NoOfINstall, FormatField(RstIndex("LoanAmt")), _
                    InstAmount, FormatField(RstIndex("IssueDate"))) Then
            MsgBox "Unable to save the installment details of the " & FormatField(RstIndex("LoanID"))
            Exit Function
        End If
        gDbTrans.BeginTrans
    End If
    RecNo = RecNo + 1
    RstIndex.MoveNext
Wend

gDbTrans.CommitTrans

LoanMasterTransfer = True
Exit Function

ErrLine:
Debug.Assert Err.Number = 0
Resume
End Function



Private Function SaveInstallmentDetails(LoanID As Long, InstMode As Integer, NoOfInst As Integer, LoanAmount As Currency, InstAmount As Currency, IssueIndianDate As String) As Boolean

Dim InstNo As Integer
Dim lpCount As Integer
Dim SqlStr As String
Dim Rst As Recordset
Dim NextDate As Date
Dim FortNight As Boolean
Dim TotalInstAmount As Currency

NextDate = FormatDate(IssueIndianDate)
gDbTrans.BeginTrans
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
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
       gDbTrans.RollBack
       Exit Function
    End If
     TotalInstAmount = TotalInstAmount + InstAmount
     InstNo = InstNo + 1
Loop
gDbTrans.CommitTrans

SaveInstallmentDetails = True
End Function


Private Function LoanTransTransfer(DbIndex As clsTransact) As Boolean
'Dim
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim BankID As Long
Dim TransID As Long
Dim TransType As Integer
Dim Particualrs As String
Dim ItIsIntTrans As Boolean
Dim LoanID As Long
Dim RegInt As Double
Dim PenalInt As Double
Dim InstType As Integer
Dim RstInst As Recordset
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency

'BankID = Val(InputBox("Enter the Id for Bank in numeric", "Bank Identification"))


'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM LoanTrans ORDER By LoanID, TransID"
DbIndex.SqlStmt = SqlStr

If DbIndex.SQLFetch < 1 Then Exit Function
Set RstIndex = DbIndex.Rst.Clone
'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id
    
Particualrs = "Penal interest"  ' Extarcted from Data BAse Differnet for Kannada
'Particualrs = LoadResString(gLangOffset + 100)

While Not RstIndex.EOF
    ItIsIntTrans = False
    TransType = FormatField(RstIndex("TransType"))
    If LoanID <> FormatField(RstIndex("LoanID")) Then
        TransID = 1
        LoanID = FormatField(RstIndex("LoanID"))
        'Get the insttalmente type
        gDbTrans.SqlStmt = "SELECT InstMode From LoanMaster Where LoanID = " & LoanID
        If gDbTrans.SQLFetch < 1 Then
            MsgBox "Error in loanid = " & LoanID
            Exit Function
        End If
        InstType = FormatField(gDbTrans.Rst("InstMode"))
    End If
    If TransType = wWithdraw Then
        'TransId = TransId + 1
        
    ElseIf TransType = wDeposit Then
        'TransId = TransId + 1
        
    Else
        ItIsIntTrans = True
        TransID = TransID + 1
        RegInt = FormatField(RstIndex("Amount"))
        If InStr(1, FormatField(RstIndex("Particulars")), Particualrs, vbTextCompare) Then
            PenalInt = FormatField(RstIndex("Amount"))
            RstIndex.MoveNext
            RegInt = FormatField(RstIndex("Amount"))
        End If
    End If
    Amount = 0
    If ItIsIntTrans Then
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
        'TransID = TransID + 1
    Else
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
    End If
        gDbTrans.BeginTrans
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Function
        End If
        gDbTrans.CommitTrans
        
    If InstType > 0 And (TransType = wContraDeposit Or TransType = wDeposit) Then
        SqlStr = "SELECT * FROM LoanInst Where LoanID = " & LoanID & _
            " AND InstBalance > 0 ORDER BY InstDate"
        gDbTrans.SqlStmt = SqlStr
        If gDbTrans.SQLFetch < 0 Then
            MsgBox "Error in loan installment of " & LoanID
            Exit Function
        End If
        Set RstInst = gDbTrans.Rst.Clone
        gDbTrans.BeginTrans
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
            gDbTrans.SqlStmt = SqlStr
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                Exit Function
            End If
            RstInst.MoveNext
        Loop
        gDbTrans.CommitTrans
    End If
    RstIndex.MoveNext
Wend
Debug.Print Now
LoanTransTransfer = True
Exit Function
ErrLine:
Debug.Assert Err.Number = 0
'Resume
End Function


Private Function SchemeTransfer(DbIndex As clsTransact) As Boolean
'Now transfer the Loans Schemes
Dim RstIndex As Recordset
'Dim DbIndex As clsTransact
DbIndex.SqlStmt = "SELECT  * From LoanTypes"
If DbIndex.SQLFetch < 1 Then
    MsgBox "No Loan Schemes"
    Exit Function
End If
Set RstIndex = DbIndex.Rst.Clone

Dim SqlStr As String
Dim SchemeID As Integer
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

gDbTrans.BeginTrans
While Not RstIndex.EOF
    On Error Resume Next
    SchemeID = RstIndex("SchemeID")
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
            SchemeID & "," & AddQuotes(SchemeName, True) & "," & _
            LoanCategary & ", " & TermType & "," & _
            LoanType & "," & Monthduration & "," & _
            DayDuration & "," & FormatField(RstIndex("InterestRate")) & _
            "," & FormatField(RstIndex("PenalInterestRate")) & "," & _
            " 'Individual'," & FormatField(RstIndex("InsuranceFee")) & "," & _
            FormatField(RstIndex("LegalFee")) & ",'Loan Name'," & _
            "#" & CreateDate & "#)"
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    RstIndex.MoveNext
Wend
gDbTrans.CommitTrans
SchemeTransfer = True
ErrLine:
Debug.Assert Err.Number = 0
End Function


