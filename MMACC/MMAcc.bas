Attribute VB_Name = "basMMAcc"
Option Explicit

Public Enum wis_MemReports
    '''Regular Reportr
    repMemBalance = 1
    repMembers = 2
    repMemDayBook = 3
    repShareIssued = 4
    repFeeCol = 5
    repMemCashBook = 6
    repMemLedger = 7
    repShareReturn = 8
    repMemOpen = 9
    repMemClose = 10
    repMemShareCert = 11
    repMonthlyBalance = 12
    repMemSubCashBook
    repMemGenLedger
    repMemLoanMembers
    repMemNonLoanMembers
End Enum


Public Enum wis_MemberType
    memRegular = 1
    memAssociate = 2
    memNominee = 3
End Enum

'This Functionm Returns the Last Transaction Date of the
'Memeber Transaction of the particular account
Private Sub GetLastTransDate(ByVal AccountId As Long, _
                Optional TransID As Long, Optional TransDate As Date)

Dim rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Long
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM MemTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM MemIntTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SqlStmt = "Select TransID,TransDate FROM MemIntPayable " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = FormatField(rst("TransDate"))
End If

ErrLine:

End Sub

'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetMemberLastTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetMemberLastTransDate = TransDate

End Function


'This Function Returns the Max Transction ID of
'the given Member share account Id
'In case there is no transaction it reurns 0
Public Function GetMemberMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetMemberMaxTransID = TransID

End Function


Public Function ComputeTotalMMLiability(AsOnDate As Date, Optional memberTYpe As Integer) As Currency

Dim ret As Long
Dim rst As Recordset

ComputeTotalMMLiability = 0

Debug.Assert ComputeTotalMMLiability <> 0

gDbTrans.SqlStmt = "SELECT MaxTransID,Y.AccID, Balance From MemTrans X " & _
    " Inner join (SELECT Max(transID) as MaxTransID,AccID From MemTrans  Where TransDate " & _
    " <= #" & AsOnDate & "#  Group by AccId) AS Y on  X.AccId = Y.AccId"
gDbTrans.CreateView ("qryMemMaxTransIDBalance")

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM qryMemMaxTransIDBalance"
If memberTYpe > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " WHERE AccID in " & _
        " (Select distinct AccId from memMaster Where MemberType = " & memberTYpe & ")"

If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then ComputeTotalMMLiability = FormatField(rst(0))


Dim SqlStr As String
SqlStr = "SELECT Max(TransID) As MaxTransID,AccID " & _
    " FROM MemTrans WHERE TransDate <= #" & AsOnDate & "# GROUP BY AccID"
gDbTrans.SqlStmt = SqlStr
gDbTrans.CreateView ("qryTemp")

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM MemTrans A, qryTemp B " & _
    " WHERE A.AccID = B.AccID And A.TransID = B.MaxTransID"

'Dim Rst As Recordset

If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then ComputeTotalMMLiability = FormatField(rst(0))

Exit Function

End Function
Public Function GetMemberNameNumberByCustID(ByVal CustomerID As Long, ByRef MemberNum As String, ByRef memberTYpe As Integer) As String
    GetMemberNameNumberByCustID = ""
    If CustomerID < 1 Then Exit Function
    
    Dim rstName As Recordset
    gDbTrans.SqlStmt = "SELECT * FROM QryMemName WHERE CustomerID = " & CustomerID
    'If memberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & memberType
    
    If gDbTrans.Fetch(rstName, adOpenStatic) > 0 Then
        GetMemberNameNumberByCustID = Trim(FormatField(rstName("Name")))
        MemberNum = FormatField(rstName("MemberNum"))
        memberTYpe = FormatField(rstName("MemberType"))
    End If
End Function
Public Function IsMemberExistsForCustomerID(CustomerID As Long, Optional memberTYpe As Integer) As Boolean
    IsMemberExistsForCustomerID = False
    If CustomerID < 1 Then Exit Function
    Dim rstName As Recordset
    gDbTrans.SqlStmt = "SELECT * FROM QryName WHERE CustomerID = " & CustomerID
    gDbTrans.SqlStmt = "SELECT A.CustomerID,B.AccNum as MemberNum,B.ACCID as MemID, Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,ClosedDate From NameTab A " & _
        " Inner Join MemMaster B on B.CustomerID = A.CustomerID" & _
        " WHERE A.CustomerId = " & CustomerID & " AND (ClosedDate IS NULL Or ClosedDate > #" & FinUSEndDate & "#)"
    
    If memberTYpe > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.MemberType = " & memberTYpe
            
    If gDbTrans.Fetch(rstName, adOpenStatic) > 0 Then IsMemberExistsForCustomerID = True
        
End Function

Public Function GetMemberNameByMemberNum(memNum As String, Optional memberTYpe As Integer) As String
    GetMemberNameByMemberNum = ""
    If Len(Trim$(memNum)) < 1 Then Exit Function
    Dim rstName As Recordset
    'gDbTrans.SQLStmt = "SELECT * FROM QryName WHERE CustomerID = " & CustomerID
    gDbTrans.SqlStmt = "Select * From QryMemName Where MemberNum = " & Trim(memNum)
    If memberTYpe > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & memberTYpe
    
    If gDbTrans.Fetch(rstName, adOpenStatic) > 0 Then _
        GetMemberNameByMemberNum = Trim(FormatField(rstName("Name")))
        
End Function

Public Function GetMemberNameCustIDByMemberNum(memNum As String, ByRef CustomerID As Long, Optional memberTYpe As Integer) As String
    GetMemberNameCustIDByMemberNum = ""
    CustomerID = 0
    If Len(Trim$(memNum)) < 1 Then Exit Function
    Dim rstName As Recordset
    gDbTrans.SqlStmt = "Select * From QryMemName Where MemberNum = " & AddQuotes(Trim(memNum))
    If memberTYpe > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & memberTYpe
                
    Dim recCount As Integer
    recCount = gDbTrans.Fetch(rstName, adOpenStatic)
    If recCount > 0 Then
        If recCount = 1 Then
            GetMemberNameCustIDByMemberNum = Trim(FormatField(rstName("Name")))
            CustomerID = FormatField(rstName("CustomerID"))
        Else
            Dim clsMem As New clsMMAcc
            GetMemberNameCustIDByMemberNum = clsMem.GetMemberNameCustIDByMemberNum(memNum, CustomerID)
            Set clsMem = Nothing
        End If
    End If
End Function

Public Sub LoadMemberTypes(cmbMember As ComboBox)
    Dim recCount As Integer
    Dim rst As Recordset
    gDbTrans.SqlStmt = "Select * From MemberTypeTab order by membertype"
    
    recCount = gDbTrans.Fetch(rst, adOpenDynamic)
    'cmbMember.Clear
    If recCount > 1 Then
        With cmbMember
            While rst.EOF = False
                .AddItem FormatField(rst("MemberTypeName"))
                .ItemData(.newIndex) = FormatField(rst("MemberType"))
                'Move to next record
                rst.MoveNext
            Wend
        End With
    Else
        cmbMember.AddItem GetResourceString(338, 49)
        cmbMember.ItemData(cmbMember.newIndex) = 0
        cmbMember.ListIndex = 0
        'cmb(cmbIndex).Locked = True
    End If

End Sub
