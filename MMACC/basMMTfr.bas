Attribute VB_Name = "MMTransfer"
Option Explicit
'Public gLangOffset As Integer


Private Function MemMasterTransfer(oldMMTrans As clsTransact, NewMMTrans As clsTransact) As Boolean
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim AccNum As String
Dim BankID As Long
Dim InstMode As Integer
Dim InstAmount As Currency
Dim NoOfINstall As Integer
Dim RecNo As Integer




On Error GoTo ErrLine

'Before Fetching Update the Values
'where It can be Null with default value
'Then Fetch the records

'Update Modify date
SqlStr = "UPDATE MMMASTER SET ModifiedDate = #1/1/100# WHERE ModifiedDate = NULL"
oldMMTrans.SQLStmt = SqlStr
oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
oldMMTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE MMMASTER Set ClosedDate = #1/1/100# WHEre ClosedDate = NULL"
oldMMTrans.SQLStmt = SqlStr
oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
oldMMTrans.CommitTrans

'Update Nominee
Dim SngSpace As String
SngSpace = ""
SqlStr = "UPDATE MMMASTER set Nominee = '" & SngSpace & "' WHERE Nominee = NULL"
oldMMTrans.SQLStmt = SqlStr
oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
oldMMTrans.CommitTrans


'This function Tranfers all the data from MMMaster of Index 200
'to MMMaster Of New data Base
'Fetch the data from Old Database

SqlStr = "SELECT * FROM MMMaster ORDER BY AccId"
oldMMTrans.SQLStmt = SqlStr
If oldMMTrans.SQLFetch < 1 Then Exit Function
Set RstIndex = oldMMTrans.Rst.Clone

'In the Loan Daabase of index 2000 we are using Member Id
'and In New Loan we are using customer id
'We have to get the customerId from the respective member id

NewMMTrans.BeginTrans
While Not RstIndex.EOF
    
    'first get the mem Id of the old db
     MemID = FormatField(RstIndex("AccId"))
    'now get the customer Of this Member
    CustomerId = 0
    InstMode = 0: NoOfINstall = 0: InstAmount = 0
    'Set oldLoanTrans.Rst = Nothing
    'SqlStr = "SELECT CustomerID from MMMaster where AccID = " & MemId
    'oldLoanTrans.SqlStmt = SqlStr
    CustomerId = FormatField(RstIndex("CustomerId")) 'FormatField(oldLoanTrans.Rst(0))
    If CustomerId = 0 Then GoTo NextAccount
    On Error GoTo ErrLine
    SqlStr = "INSERT INTO MemMaster (" & _
            " AccID,AccNum,CustomerID," & _
            " CreateDate,ModifiedDate,ClosedDate," & _
            " NomineeID,IntroducerID,LedgerNo," & _
            " FolioNo,MemberType) "
    
    SqlStr = SqlStr & " Values ( " & _
            RstIndex("AccID") & "," & _
            "'" & Format(RstIndex("AccID"), "000") & "' ," & _
            RstIndex("CustomerID") & "," & _
            "#" & RstIndex("CreateDate") & "#," & _
            "#" & RstIndex("ModifiedDate") & "#," & _
            "#" & RstIndex("ClosedDate") & "#," & _
            "0 ," & RstIndex("Introduced") & ", " & _
            "'" & FormatField(RstIndex("LedgerNo")) & "'," & _
            "'" & FormatField(RstIndex("FolioNo")) & "'," & _
            RstIndex("MemberType") + 1 & ") "
    
    NewMMTrans.SQLStmt = SqlStr
    If Not NewMMTrans.SQLExecute Then    'THere aRE MORE FIELDS
        'gDBTrans.CommitTrans
        
        NewMMTrans.RollBack              'IN lOANS THAN iNDEX
        Exit Function
    End If
    'If loan has got the installment then insert the installment details
    
NextAccount:
    RecNo = RecNo + 1
    RstIndex.MoveNext
Wend

NewMMTrans.CommitTrans


'Update Modify date
SqlStr = "UPDATE MMMASTER SET ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
oldMMTrans.SQLStmt = SqlStr
oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
'oldMMTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE MMMASTER Set ClosedDate = NULL WHEre ClosedDate = #1/1/100#"
'oldMMTrans.SQLStmt = SqlStr
'oldMMTrans.BeginTrans
If Not oldMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
oldMMTrans.CommitTrans

SqlStr = "UPDATE MemMASTER SET ModifiedDate = NULL WHERE ModifiedDate = #1/1/100# "
NewMMTrans.SQLStmt = SqlStr
NewMMTrans.BeginTrans
If Not NewMMTrans.SQLExecute Then
    oldMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewMMTrans.CommitTrans

'Update Closed date
SqlStr = "UPDATE MemMASTER Set ClosedDate = NULL Where ClosedDate = #1/1/100#"
NewMMTrans.SQLStmt = SqlStr
NewMMTrans.BeginTrans
If Not NewMMTrans.SQLExecute Then
    NewMMTrans.RollBack
    MsgBox "UPdate the Data Base Instead Of transfer"
    Exit Function
End If
NewMMTrans.CommitTrans

MemMasterTransfer = True
Exit Function

ErrLine:
Debug.Assert Err.Number = 0
Resume

End Function



Private Function MemTransTransfer(oldMemTrans As clsTransact, NewMemTrans As clsTransact) As Boolean
'Dim
Dim RstIndex As Recordset

Dim SqlStr As String
Dim SqlSupport As String
Dim CustomerId As Long
Dim MemID As Long
Dim BankID As Long
Dim TransID As Long
Dim TransType As Integer
Dim ItIsIntTrans As Boolean
Dim Amount As Currency
Dim InstAmount As Currency
Dim InstNo As Integer
Dim InstBalance As Currency


'BankID = Val(InputBox("Enter the Id for Bank in numeric", "Bank Identification"))


'This function Tranfers all the data from LoanMster of Index 200
'to LoanMaster Of Loan data Base
'Fetch the data from Old Index Dbof loans

SqlStr = "SELECT * FROM MMTrans ORDER By AccID, TransID"
oldMemTrans.SQLStmt = SqlStr

If oldMemTrans.SQLFetch < 1 Then GoTo ExitLine
    
Set RstIndex = oldMemTrans.Rst.Clone

NewMemTrans.BeginTrans
    
While Not RstIndex.EOF
    ItIsIntTrans = False
    TransType = FormatField(RstIndex("TransType"))
    If MemID <> RstIndex("accID") Then TransID = 1
    MemID = RstIndex("accID")
    
    'Begin the transaction
    If TransType = -2 Or TransType = 2 Then
        
        ItIsIntTrans = True
        'TransType = wDeposit
        Amount = FormatField(RstIndex("Amount"))
        SqlStr = "INSERT INTO MemIntTrans (" & _
                " AccId,TransDate," & _
                " TransID,TransType," & _
                " Amount,Balance, " & _
                " Particulars)"
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("AccId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                RstIndex("TransID") & "," & TransType / Abs(TransType) * -1 & "," & _
                Amount & ", 0, " & _
                "'" & IIf(RstIndex("TransID") = 1, "Membership Fee", "Share Fee") & "' )"
    Else
        TransID = TransID + 1
        TransType = FormatField(RstIndex("TransType"))
        Amount = FormatField(RstIndex("Amount"))
        SqlStr = "INSERT INTO MemTrans (" & _
                " AccId,TransDate," & _
                " TransID,Leaves,TransType," & _
                " Amount, Balance)"
        
        SqlStr = SqlStr & " Values ( " & _
                RstIndex("AccId") & "," & _
                "#" & RstIndex("TransDate") & "#," & _
                TransID & "," & _
                RstIndex("Leaves") & "," & _
                TransType / Abs(TransType) & "," & _
                Amount & "," & _
                RstIndex("Balance") & "  )"
    End If
    
    NewMemTrans.SQLStmt = SqlStr
    If Not NewMemTrans.SQLExecute Then
        NewMemTrans.RollBack
        Exit Function
    End If

NextAccount:
    RstIndex.MoveNext
Wend

    NewMemTrans.CommitTrans


ExitLine:
Debug.Print Now
MemTransTransfer = True
Exit Function
ErrLine:
Debug.Assert Err.Number = 0
'Resume
End Function


Private Function ShareTransfer(oldMemTrans As clsTransact, NewMemTrans As clsTransact) As Boolean
'Now transfer the Loans Schemes

Dim RstIndex As Recordset

oldMemTrans.SQLStmt = "SELECT  * From ShareLeaves"
If oldMemTrans.SQLFetch < 1 Then
    MsgBox "No Loan Schemes"
    Exit Function
End If
Set RstIndex = oldMemTrans.Rst.Clone

Dim SqlStr As String
Dim ACCID As Integer
Dim SchemeName As String
Dim Description  As String
Dim CreateDate As Date

CreateDate = Now

NewMemTrans.BeginTrans
While Not RstIndex.EOF
    
    SqlStr = "INSERT INTO ShareTrans (AccID, SaleTransID," & _
            " ReturnTransID, CertNo,CertID," & _
            " FaceValue ) "
    SqlStr = SqlStr & " Values (" & _
            RstIndex("AccID") & "," & _
            RstIndex("SaleTransID") & "," & _
            FormatField(RstIndex("ReturnTransID")) & "," & _
            "'" & RstIndex("CertNo") & "'," & _
            RstIndex("CertNo") & "," & _
            RstIndex("FaceValue") & ")"
    NewMemTrans.SQLStmt = SqlStr
    If Not NewMemTrans.SQLExecute Then
        NewMemTrans.RollBack
        Exit Function
    End If
    
    RstIndex.MoveNext
Wend
NewMemTrans.CommitTrans
ShareTransfer = True
ErrLine:
Debug.Assert Err.Number = 0
End Function


Public Function MemberTransfer(OldDBName As String, NewDBName As String) As Boolean
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

Screen.MousePointer = vbHourglass
If Not MemMasterTransfer(OldTrans, NewTrans) Then GoTo ExitLine
If Not MemTransTransfer(OldTrans, NewTrans) Then GoTo ExitLine
If Not ShareTransfer(OldTrans, NewTrans) Then GoTo ExitLine
    
MemberTransfer = True


ExitLine:

Screen.MousePointer = vbDefault

End Function


