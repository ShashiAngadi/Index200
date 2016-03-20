Attribute VB_Name = "BKCCTransfer"
Option Explicit
'Public gLangOffset As Integer

'This Bas file is made to transfer the Index 2000
'data base to the this existing loansdatabase

Function TransferBKCC(OldDBName As String, NewDBName As String) As Boolean

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

'If Not SchemeTransfer(OldTrans, NewTrans) Then Exit Function
If Not BKCCMasterTransfer(OldTrans, NewTrans) Then Exit Function
If Not BKCCTransTransfer(OldTrans, NewTrans) Then Exit Function
    
TransferBKCC = True

End Function



Private Function BKCCMasterTransfer(oldLoanTrans As clsTransact, NewLOanTrans As clsDBUtils) As Boolean
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim AccNum As String
Dim BankID As Long
Dim RecNo As Integer

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Database of loans


SqlStr = "SELECT A.*,B.CustomerID FROM LoanMaster A, MMMaster B " & _
    " WHERE B.AccID=A.MemberID AND LoanId IN (SELECT LoanID " & _
        " From LOanMaster Where SchemeID In (SELECT SchemeID FRom LoanTypes Where BKCC = TRUE ))" & _
    " ORDER BY LoanID"

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
    'now get the customer Of this Member
    CustomerId = 0
    CustomerId = FormatField(RstIndex("CustomerId"))
    AccNum = FormatField(RstIndex("LoanAccNo"))
    If Trim(AccNum) = "" Then _
        AccNum = FormatField(RstIndex("SchemeID")) & "_" & FormatField(RstIndex("LoanId"))
    On Error Resume Next
    
    On Error GoTo Errline
    SqlStr = "INSERT INTO BKCCMaster (" & _
            " LoanId,MemID,CustomerID," & _
            " AccNum,Issuedate," & _
            " SanctionAmount," & _
            " Intrate,PenalIntRate,DepIntRate, " & _
            " Guarantor1,Guarantor2,LoanClosed, Remarks) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            MemID & ", " & CustomerId & "," & _
            AddQuotes(AccNum, True) & "," & _
            "#" & RstIndex("IssueDate") & "#," & _
            FormatField(RstIndex("LoanAmt")) & "," & _
            FormatField(RstIndex("InterestRate")) & "," & _
            FormatField(RstIndex("PenalInterestrate")) & ", 10 ," & _
            FormatField(RstIndex("GuarantorId1")) & "," & _
            FormatField(RstIndex("GuarantorId2")) & "," & _
            FormatField(RstIndex("LoanClosed")) & "," & _
            AddQuotes(FormatField(RstIndex("Remarks")), True) & ")"

    NewLOanTrans.SQLStmt = SqlStr
    If Not NewLOanTrans.SQLExecute Then
        NewLOanTrans.RollBack
        Exit Function
    End If
    
NextAccount:
    RecNo = RecNo + 1
    RstIndex.MoveNext
Wend

NewLOanTrans.CommitTrans

BKCCMasterTransfer = True
Exit Function

Errline:
Debug.Assert Err.Number = 0

End Function



Private Function BKCCTransTransfer(oldLoanTrans As clsTransact, NewLOanTrans As clsTransact) As Boolean
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
Dim DepTrans As Boolean

'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM BKCCTrans ORDER By LoanID, TransID"
oldLoanTrans.SQLStmt = SqlStr
If oldLoanTrans.SQLFetch < 1 Then GoTo ExitLine
    
Set RstIndex = oldLoanTrans.Rst.Clone

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id
    
Particualrs = "Penal Interest"  ' Extarcted from Data BAse Differnet for Kannada
If UCase(Trim(InputBox("Enter language either ENGLISH or KANNADA", "Loan Transfer", "KANNADA"))) = "KANNADA" Then
    Particualrs = LoadResString(5345)
Else
    Particualrs = LoadResString(345)
End If

Dim InTrans As Boolean

DepTrans = False

While Not RstIndex.EOF
    ItIsIntTrans = False
    TransType = FormatField(RstIndex("TransType"))
'    Debug.Assert LoanID <> 740
    If LoanID <> FormatField(RstIndex("LoanID")) Then
        Set RstInst = Nothing
        TransID = 1: DepTrans = False
        LoanID = FormatField(RstIndex("LoanID"))
    End If
'    Debug.Assert LoanID <> 1629
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
        End If
        TransType = IIf(TransType < 0, wDeposit, wWithDraw)
        SqlStr = "INSERT INTO BKccIntTrans (" & _
                " LoanId,TransDate," & _
                " TransID,TransType," & _
                " IntAmount,PenalIntAmount, " & _
                " IntBalance,Deposit,Particulars)"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("LoanId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & Abs(TransType) / TransType & "," & _
                RegInt & "," & _
                PenalInt & "," & _
                FormatField(RstIndex("Balance")) & ", " & _
                DepTrans & "," & _
                AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
        If RegInt Or PenalInt Then
            NewLOanTrans.SQLStmt = SqlStr
            If Not NewLOanTrans.SQLExecute Then
                NewLOanTrans.RollBack
                InTrans = False
                Exit Function
            End If
            NewLOanTrans.SQLStmt = "Update BKCCMaster Set LastIntDate = " & _
                "#" & RstIndex("transDate") & "# WHERE LoanID = " & RstIndex("LoanId")
            If Not NewLOanTrans.SQLExecute Then
                NewLOanTrans.RollBack
                InTrans = False
                Exit Function
            End If
            RegInt = 0: PenalInt = 0
        End If
        RstIndex.MoveNext
        If RstIndex.EOF Then
            NewLOanTrans.CommitTrans
            InTrans = False
            GoTo NextAccount
        End If
    End If
    
    TransType = FormatField(RstIndex("TransType"))
    If RstIndex("Balance") < 0 Then DepTrans = True
    If RstIndex("Balance") > 0 Then DepTrans = False
    
    If Not (TransType = wWithDraw Or TransType = wDeposit) Then
        If TransType = 7 Then
            DepTrans = True
            TransType = wDeposit
        ElseIf TransType = -7 Then
            DepTrans = True
            TransType = wWithDraw
        Else
            NewLOanTrans.CommitTrans
            InTrans = False
            RstIndex.MovePrevious
            GoTo NextAccount
        End If
    End If
    
    Amount = FormatField(RstIndex("Amount"))
    SqlStr = "INSERT INTO BKCCTrans (" & _
            " LoanId,TransDate," & _
            " TransID,TransType," & _
            " Amount, Balance, Deposit,Particulars)"
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("LOanId") & "," & _
            "#" & RstIndex("TransDate") & "#," & _
            TransID & "," & TransType & "," & _
            Amount & "," & _
            FormatField(RstIndex("Balance")) & "," & _
            DepTrans & ", " & _
            AddQuotes(FormatField(RstIndex("Particulars")), True) & ")"
            If TransID = 1 Then TransID = 2
            
    NewLOanTrans.SQLStmt = SqlStr
    If Not NewLOanTrans.SQLExecute Then
        NewLOanTrans.RollBack
        InTrans = False
        Exit Function
    End If

    NewLOanTrans.CommitTrans
    InTrans = False
    
NextAccount:
    RstIndex.MoveNext
Wend


ExitLine:
Debug.Print Now
BKCCTransTransfer = True
Exit Function
Errline:
Debug.Assert Err.Number = 0
'Resume
End Function


