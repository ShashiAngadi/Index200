VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMMAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_rstAccDet As ADODB.Recordset
Private m_AccID As Long
Private m_ParentHandle As Long
Private m_MemberType As Integer
Private m_selValues(5) As String


Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmSelectMemType As frmSelectFD
Attribute m_frmSelectMemType.VB_VarHelpID = -1
Private WithEvents m_FrmMemAcc As frmMMAcc
Attribute m_FrmMemAcc.VB_VarHelpID = -1
Private WithEvents m_frmMMShare As frmMMShare
Attribute m_frmMMShare.VB_VarHelpID = -1
Private WithEvents m_frmMMReport As frmMMReport
Attribute m_frmMMReport.VB_VarHelpID = -1
Private m_Share() As String
Public Function GetMemberNameCustIDByMemberNum(memNum As String, ByRef CustomerID As Long, Optional MemberType As Integer) As String
    GetMemberNameCustIDByMemberNum = ""
    CustomerID = 0
    If Len(Trim$(memNum)) < 1 Then Exit Function
    Dim rstName As Recordset
    gDbTrans.SqlStmt = "Select MemberNum + ' - ' +  MemberTypeName as MemberNum,Name, CustomerID From QryMemName A " & _
        " Left Join MemberTypeTab B On A.MemberType = B.MemberType  Where MemberNum = " & AddQuotes(Trim(memNum))
    If MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And MemberType = " & MemberType
                
    Dim recCount As Integer
    
    recCount = gDbTrans.Fetch(rstName, adOpenStatic)
    If recCount > 0 Then
        If recCount > 1 Then
            'ReDim m_selValues(10)
            m_selValues(0) = ""
            m_selValues(1) = ""
            If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
            Call FillView(m_frmLookUp.lvwReport, rstName, True)
            With m_frmLookUp
                .lvwReport.ColumnHeaders(3).Width = 0
                .Show 1
            End With
            Set m_frmLookUp = Nothing
            GetMemberNameCustIDByMemberNum = Trim(m_selValues(1))
            On Error Resume Next
            CustomerID = CLng(m_selValues(2))
        Else
            GetMemberNameCustIDByMemberNum = Trim(FormatField(rstName("Name")))
            CustomerID = FormatField(rstName("CustomerID"))
        End If
    End If
End Function



Public Property Get ObjectHandle() As Long
    If Not m_FrmMemAcc Is Nothing Then _
            ObjectHandle = m_FrmMemAcc.hwnd: Exit Property
    
    If Not m_frmMMReport Is Nothing Then _
        Unload m_frmMMReport: Set m_frmMMReport = Nothing
End Property

Private Property Let AccountId(NewValue As Long)
m_AccID = NewValue
End Property

Public Function NextBalance(retAccNum As String) As Currency
'To trap an error
On Error GoTo ErrLine

NextBalance = 0

'Sql stmt to get Accid
If m_rstAccDet Is Nothing Then GoTo ExitLine
m_rstAccDet.MoveNext
If m_rstAccDet.EOF Then GoTo ExitLine

retAccNum = FormatField(m_rstAccDet("AccNum"))
NextBalance = FormatField(m_rstAccDet("Balance"))
 
ExitLine:
    Exit Function
    
ErrLine:
    If Err Then
        MsgBox "Mamber Account NextBalance: " & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitLine
    End If

End Function

Public Function CustomerBalance(ByVal CustomerID As Long, retAccNum As String, Optional AsOnDate As Date) As Currency

Dim MemberTypeName As String

If IsMissing(AsOnDate) Then AsOnDate = CDate(gStrDate)

CustomerBalance = CustomerBalanceNew(CustomerID, retAccNum, MemberTypeName, AsOnDate)

End Function


Public Function CustomerBalanceNew(ByVal CustomerID As Long, retAccNum As String, ByRef AccountTypeName As String, Optional AsOnDate As Date) As Currency

'To trap an error
On Error GoTo ErrLine
CustomerBalanceNew = 0

If IsMissing(AsOnDate) Then AsOnDate = CDate(gStrDate)
retAccNum = ""
  
gDbTrans.SqlStmt = "SELECT X.AccID, MaxTransID,Balance From MemTrans X" & _
    " Inner join (SELECT Max(transID) as MaxTransID,AccID From MemTrans " & _
    " Where TransDate <= #" & AsOnDate & "#  Group by AccId) AS Y " & _
    " on  X.AccId = Y.AccId  and X.TransID = Y.MaxTransID"
gDbTrans.CreateView ("qryMemTransIDBalance")


' OLD SqlStmt
gDbTrans.SqlStmt = "Select AccNum, A.AccId, Balance from MemMaster A, MemTrans B " & _
    " WHERE Customerid = " & CustomerID & " AND B.AccID = A.AccID " & _
    " AND TransID = (SELECT Max(transID) From MemTrans C Where C.AccID = B.AccID " & _
        " AND TransDate <= #" & AsOnDate & "#)" & _
    " And (ClosedDate >= #" & AsOnDate & "# OR ClosedDate is NULL )" & _
    " ORDER BY A.AccID "

gDbTrans.SqlStmt = "Select AccNum, A.AccId, B.Balance,C.MemberTypeName from " & _
    " (MemMaster A  Left Join MemberTypeTab  C on A.MemberType = C.MemberType) " & _
    " Left Join qryMemTransIDBalance B On A.AccId = B.AccId" & _
    " WHERE A.Customerid = " & CustomerID & _
    " AND (ClosedDate >= #" & AsOnDate & "# OR ClosedDate is NULL )" & _
    " ORDER BY A.AccID "
    
AccountTypeName = ""
If gDbTrans.Fetch(m_rstAccDet, adOpenDynamic) < 1 Then
    'There are some nominal members they will not have any share transaction
    'Fetch such customers here
    gDbTrans.SqlStmt = "Select AccNum, AccId, 0 as Balance from MemMaster" & _
        " WHERE Customerid = " & CustomerID & _
        " ORDER BY AccID "
    If gDbTrans.Fetch(m_rstAccDet, adOpenDynamic) < 1 Then Exit Function
End If
retAccNum = FormatField(m_rstAccDet("AccNum"))
CustomerBalanceNew = FormatField(m_rstAccDet("Balance"))
AccountTypeName = FormatField(m_rstAccDet("MemberTypeName"))

'gDbTrans.SqlStmt = "Select * From MemberTypeTab"
'recCount = gDbTrans.Fetch(rstMemType, adOpenDynamic)
        

ExitLine:
    Exit Function

ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If

End Function

Public Property Get ObjectName()
    ObjectName = "MemberClass"
End Property



Public Function Balance(AsOnDate As Date, Optional memberType As Integer) As Currency
    Balance = 0
    If IsMissing(memberType) Then memberType = 0
    Balance = ComputeTotalMMLiability(AsOnDate, memberType)
End Function
Public Function LedgerOpBalance(AsOnDate As Date, Optional memberType As Integer) As Currency
    Dim headID As Long
    Dim AccTransClass As clsAccTrans
    Dim memTypeName As String
    Dim rst As Recordset
    
    LedgerOpBalance = 0
    memTypeName = ""
    If IsMissing(memberType) Then memberType = 0
        
    If memberType > 0 Then
        gDbTrans.SqlStmt = "SeLECT * from MemberTypeTab where membertype = " & memberType
        Call gDbTrans.Fetch(rst, adOpenDynamic)
        memTypeName = FormatField(rst("MemberTypeName"))
        Set rst = Nothing
    End If
    
    headID = GetHeadID(memTypeName & " " & GetResourceString(53, 36), parMemberShare)
    
    Set AccTransClass = New clsAccTrans
    LedgerOpBalance = AccTransClass.GetOpBalance(headID, AsOnDate)
    Set AccTransClass = Nothing
    
End Function

Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset

gDbTrans.SqlStmt = "Select AccID From MemMaster " & _
    "Where CustomerID = " & NewValue
If gDbTrans.Fetch(rst, adOpenDynamic) Then m_AccID = rst("AccID")

Set rst = Nothing


End Property

Public Function CustomerNextBalance(retAccNum As String) As Currency
    Dim MemberType As String
    
    CustomerNextBalance = CustomerNextBalanceNew(retAccNum, MemberType)
    
End Function
Public Function CustomerNextBalanceNew(ByRef retAccNum As String, ByRef MemberTypeName) As Currency
    CustomerNextBalanceNew = 0
    retAccNum = ""
    If m_rstAccDet Is Nothing Then Exit Function
    If m_rstAccDet.EOF Then Exit Function
    m_rstAccDet.MoveNext
    If m_rstAccDet.EOF Then Exit Function
    
    retAccNum = FormatField(m_rstAccDet("AccNum"))
    CustomerNextBalanceNew = FormatField(m_rstAccDet("Balance"))
    MemberTypeName = FormatField(m_rstAccDet("memberTypeName"))
    
End Function

Public Function DepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
            ByVal ShareFee As Currency, ByVal strDesc As String, _
            ByVal TransDate As Date, Optional VoucherNo As String = " ", _
            Optional CashTransaction As Boolean = False, Optional ContraID As Long) As Integer

Dim I As Long
Dim MaxI As Integer
Dim FaceValue As Currency
Dim rst As ADODB.Recordset
Dim UserID As Integer
Dim AccHeadID As Long

Dim MemberType As Integer
Dim MemberTypeName As String
Dim MemberTypeNameEnglish As String

'Get the MemberTYpe for the AccID
MemberType = GetMemberTypeByAccID(AccId, MemberTypeName, MemberTypeNameEnglish)
'Get the Transaction Type
Dim transType As wisTransactionTypes
transType = IIf(CashTransaction, wDeposit, wContraDeposit)

'Get the Share Leaves No
    'Display share form
    Set m_frmMMShare = New frmMMShare
    Load m_frmMMShare
    
    ReDim m_Share(0)
    m_Share(0) = ""
    
    Dim SetupClass As New clsSetup
    Dim ShareValue As Currency
    
    ShareValue = Val(SetupClass.ReadSetupValue("MMAcc", "ShareValue", "10"))
    ShareValue = Val(SetupClass.ReadSetupValue("MMAcc" & MemberType, "ShareValue", CStr(ShareValue)))
    
    'Temporay message if he has not having any share value
    If ShareValue = 0 Then
        MsgBox "There is no share value", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
    
    With m_frmMMShare
        .Caption = "Share Details"
        .fraPurchase.Caption = GetResourceString(314) '"Issue Of Shares"
        .fraPurchase.Visible = True
        .cmdSelect.Visible = False
        .cmdAdd.Visible = True: .cmdAdd.Default = True
        .cmdInvert.Visible = False
        .fraSale.Visible = False
        .txtLeaves = Amount / ShareValue
        .Show vbModal
    End With
    If UBound(m_Share) = 0 And m_Share(0) = "" Then GoTo Exit_Line
    MaxI = UBound(m_Share)
    'Calculate the total amount
    ShareValue = ShareValue * (MaxI + 1)
    Set SetupClass = Nothing

''''''''''''''''' Share Leaves
'The value of the shares needs to be determined previously
If ShareValue = 0 Then
    'MsgBox "Unable to compute total share worth !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(764), vbExclamation, gAppName & " - Error"
    FaceValue = Amount / MaxI
    ShareValue = Amount
End If

If ShareValue <> Amount Then
    MsgBox "The share value and amount depositing are not matching. ", vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Get the last TransID,TransDate and Balance
    Dim TransID As Long
    Dim LastTransDate As Date
    Dim Balance As Long
    'Get the Contra Transaction ID
    If ContraID = 0 Then ContraID = GetMaxContraTransID
    
    TransID = 1
    Balance = Amount
    gDbTrans.SqlStmt = "Select TOP 1 TransID, TransDate, " & _
        " Balance from MemTrans where Accid = " & AccId & _
        " order by TransID desc"
    LastTransDate = "1/1/100"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        TransID = FormatField(rst("TransID"))
        LastTransDate = rst("TransDate")
        Balance = FormatField(rst("Balance"))
        Balance = Balance + Amount
        'Check for the last transaction date
        If DateDiff("D", TransDate, LastTransDate) > 0 Then
            'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    End If
    
    gDbTrans.SqlStmt = "Select TOP 1 TransID, TransDate, " & _
        " Balance from MemIntTrans where Accid = " & AccId & _
        " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        TransID = IIf(TransID > rst("TransID"), TransID, rst("TransID"))
        LastTransDate = rst("TransDate")
        'Check for the last transaction date
        If DateDiff("d", TransDate, LastTransDate) > 0 Then
            'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    End If
    TransID = TransID + 1
    
'Obtain the Face Value from current settings
    Dim SetUp As New clsSetup
    FaceValue = Val(SetUp.ReadSetupValue("MMAcc", "ShareValue", "10.00"))
    FaceValue = Val(SetUp.ReadSetupValue("MMAcc" & MemberType, "ShareValue", CStr(FaceValue)))


UserID = gCurrUser.UserID

'Perform DB Transactions
    'First perform the transaction for Share fee(charges)
    'Note that this will not have any effect on the Balance.
    'IT is here only to determine the profit made by the bank
    gDbTrans.SqlStmt = "INSERT into MemIntTrans (AccID, TransID, TransDate, " & _
        " Amount, TransType, Balance,UserID) values ( " & _
        AccId & "," & _
        TransID & ", " & _
        "#" & TransDate & "#, " & _
        ShareFee & ", " & _
        transType & ", " & _
        Balance & "," & UserID & ")"
    
    If ShareFee > 0 Then
        AccHeadID = GetIndexHeadID(MemberTypeName & " " & GetResourceString(53, 191)) 'Share Fee
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
        
        If Not CashTransaction Then
            'update it to te Contra Trans table
            gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    "(ContraID,AccHeadID,AccID," & _
                    "TransType, TransID,Amount,VoucherNo,UserID)" & _
                    " Values (" & ContraID & "," & AccHeadID & "," & _
                    AccId & ", " & transType & ", " & TransID & "," & ShareFee & ", " & _
                    AddQuotes(VoucherNo, True) & "," & gUserID & " )"
            ' Execute the updation.
            If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        End If
    End If
    
    'Insert / Update the the Share leaves table depending upon the transaction type
    MaxI = UBound(m_Share)
    For I = 0 To MaxI
        'Insert the share numbers into the ShareLeaves tab
        gDbTrans.SqlStmt = "Insert into ShareTrans (AccID, SaleTransID, " & _
                " CertNo, FaceValue) values (" & _
                AccId & ", " & _
                TransID & ", " & _
                AddQuotes(m_Share(I), True) & ", " & _
                FaceValue & ")"
    
        'Fire the query just built
        gDbTrans.BeginTrans
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to perform transactions !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    Next I
    
    'Insert into MemTrans
    gDbTrans.SqlStmt = "INSERT INTO MemTrans (AccID, TransID, TransDate, " & _
            " Leaves, Amount, TransType, Balance,UserID,VoucherNo ) values ( " & _
            AccId & "," & _
            TransID & ", " & _
            "#" & TransDate & "#, " & _
            MaxI + 1 & ", " & _
            Amount & ", " & _
            transType & ", " & _
            Balance & "," & _
            UserID & "," & _
            AddQuotes(VoucherNo, True) & ")"
            
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        GoTo Exit_Line
    End If
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        AccHeadID = GetIndexHeadID(MemberTypeName & " " & GetResourceString(53, 36))
        gDbTrans.SqlStmt = "Insert INTO ContraTrans" & _
            " (ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserId)" & _
            " Values (" & ContraID & "," & AccHeadID & "," & _
            AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
DepositAmount = TransID

Exit_Line:

If Err Then
    MsgBox "Error In depositAmount - MMClass", vbInformation, wis_MESSAGE_TITLE
    'resume
    Err.Clear
End If

End Function

Public Function MemberName(AccNum As String, Optional MemberType As Integer) As String

If IsMissing(MemberType) Then MemberType = 0

Dim Lret As Long
gDbTrans.SqlStmt = "Select Title + Space(1) + FirstName + " & _
        "Space(1) + MiddleName + Space(1) + LastName as Name " & _
        " from NameTab, MemMaster where MemMaster.AccNum  = " & AddQuotes(AccNum, True) & _
        " And  NameTab.CustomerID =  MemMaster.CustomerId"

If MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And memberType = " & MemberType

Dim rst As ADODB.Recordset
Lret = gDbTrans.Fetch(rst, adOpenStatic)

If Lret <= 0 Then
    MemberName = "NO Match Found"
Else
    MemberName = CStr(FormatField(rst(0)))
    Set rst = Nothing
End If
End Function

Public Function ContraProfit(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim RetAmount As Currency

ContraProfit = 0
transType = wContraDeposit
RetAmount = IntTransactionDetails(fromDate, toDate, transType)
ContraProfit = RetAmount

End Function

Public Function CashProfit(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim RetAmount As Currency

CashProfit = 0
transType = wDeposit
RetAmount = IntTransactionDetails(fromDate, toDate, transType)

CashProfit = RetAmount

End Function

Public Function Profit(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim RetAmount As Currency

Profit = 0
transType = wDeposit
RetAmount = IntTransactionDetails(fromDate, toDate, transType)
transType = wContraDeposit
RetAmount = RetAmount + IntTransactionDetails(fromDate, toDate, transType)
Profit = RetAmount

End Function
    
Public Function ContraLoss(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim Amount As Currency
ContraLoss = 0

transType = wContraWithdraw
Amount = IntTransactionDetails(fromDate, toDate, transType)

ContraLoss = Amount

End Function
Public Function CashLoss(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim Amount As Currency
CashLoss = 0

transType = wWithdraw
Amount = IntTransactionDetails(fromDate, toDate, transType)

CashLoss = Amount

End Function

Public Function Loss(fromDate As Date, toDate As Date) As Currency

Dim transType As wisTransactionTypes
Dim Amount As Currency
Loss = 0

transType = wWithdraw
Amount = IntTransactionDetails(fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + IntTransactionDetails(fromDate, toDate, transType)

Loss = Amount

End Function


Public Function CashWithdrawls(fromDate As Date, toDate As Date) As Currency

CashWithdrawls = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wWithdraw
Amount = TransactionDetails(fromDate, toDate, transType)

CashWithdrawls = Amount
End Function

Public Function ContraWithdrawls(fromDate As Date, toDate As Date) As Currency

ContraWithdrawls = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wContraWithdraw
Amount = TransactionDetails(fromDate, toDate, transType)

ContraWithdrawls = Amount
End Function

Public Function WithdrawAmount(AccId As Long, Amount As Currency, _
            CancellationFee As Currency, strDesc As String, _
            TransDate As Date, Optional VoucherNo As String, _
            Optional CashTransaction As Boolean = False, Optional ContraID As Long) As Integer

MsgBox "Not Tested", vbInformation
Exit Function

Dim I As Long
Dim FaceValue As Currency
Dim ShareAmount As Currency
Dim rst As ADODB.Recordset

Dim MemberType As Integer
Dim MemberTypeName As String
Dim MemberTypeNameEnglish As String

'Get the MemberTYpe for the AccID
MemberType = GetMemberTypeByAccID(AccId, MemberTypeName, MemberTypeNameEnglish)

'Get the Transaction Type
Dim transType As wisTransactionTypes

'Get the Headid
Dim ClsBank As clsBankAcc
Dim AccHeadID As Long
Dim FeeHeadId As Long


Set ClsBank = New clsBankAcc
AccHeadID = ClsBank.GetHeadIDCreated(MemberTypeName & " " & GetResourceString(53, 36)) 'ClsBank.GetHeadIDCreated(GetResourceString(49, 53))

FeeHeadId = ClsBank.GetHeadIDCreated(MemberTypeName & " " & GetResourceString(53, 191)) 'GetResourceString(49, 53, 191))  ' SHARE Fee

Set ClsBank = Nothing
'Get the Contra Transaction ID
If Not CashTransaction And ContraID = 0 Then ContraID = GetMaxContraTransID + 1

transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Check if leaves have been specified
'Display share form
Set m_frmMMShare = New frmMMShare
ReDim m_Share(0)
m_Share(0) = ""
With m_frmMMShare
    .Caption = "Share Details"
    .fraPurchase.Visible = False
    .fraSale.Caption = GetResourceString(315) '"Current Share Nos"
    .fraSale.Visible = True
    .cmdAdd.Visible = False
    .cmdSelect.Visible = True: .cmdSelect.Default = True
    .cmdInvert.Visible = True
    'Fill up the leaves that this guy has got
    gDbTrans.SqlStmt = "Select * from ShareTrans where AccID = " & AccId & _
                " order by CertNo"
    If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then GoTo Exit_Line
    
    'FIll up the list box in the form
    .lstCheque.Clear
    For I = 1 To rst.recordCount
        If FormatField(rst("ReturnTransID")) = 0 Then _
        .lstCheque.AddItem FormatField(rst("CertNo"))
        rst.MoveNext
    Next I
End With

m_frmMMShare.Show vbModal

If m_Share(0) = "" Then GoTo Exit_Line
    
'''''''''

'The value of the shares needs to be determined previously
    'YOu need to compute this using the database
Dim MaxI As Integer

MaxI = UBound(m_Share)
For I = 0 To MaxI
    gDbTrans.SqlStmt = "Select FaceValue from ShareTrans " & _
        " where CertNo = " & AddQuotes(m_Share(I), True)
    If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then
        'MsgBox "Unable to locate the Certificate Number : " & cmbLeaves.List(i)
        MsgBox GetResourceString(587) & " " & m_Share(I)
        Exit Function
    Else
        ShareAmount = ShareAmount + FormatField(rst("FaceValue"))
        If ShareAmount >= Amount Then MaxI = I
    End If
Next I


'Get the last TransID,TransDate and Balance
    Dim TransID As Long
    Dim LastTransDate As Date
    Dim Balance As Long
    'Dim Rst As ADODB.Recordset
    TransID = 1
    
    Balance = Amount
    gDbTrans.SqlStmt = "Select TOP 1 TransID, TransDate, " & _
        " Balance from MemTrans where Accid = " & AccId & _
        " order by TransID desc"
    LastTransDate = "1/1/100"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        TransID = FormatField(rst("TransID"))
        LastTransDate = rst("TransDate")
        Balance = FormatField(rst("Balance"))
        Balance = Balance + IIf(transType = wDeposit, Amount, (-1) * Amount)
        'Check for the last transaction date
        If DateDiff("D", TransDate, LastTransDate) > 0 Then
            'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    End If
    
    gDbTrans.SqlStmt = "Select TOP 1 TransID, TransDate, " & _
        " Balance from MemIntTrans where Accid = " & AccId & _
        " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        TransID = IIf(TransID > rst("TransID"), TransID, rst("TransID"))
        LastTransDate = rst("TransDate")
        'Check for the last transaction date
        If DateDiff("d", TransDate, LastTransDate) > 0 Then
            'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    End If
    TransID = TransID + 1
    
'Obtain the Face Value from current settings
    Dim SetUp As New clsSetup
    FaceValue = Val(SetUp.ReadSetupValue("MMAcc", "ShareValue", "10.00"))
    FaceValue = Val(SetUp.ReadSetupValue("MMAcc" & m_MemberType, "ShareValue", CStr(FaceValue)))
Dim UserID As Integer
UserID = gCurrUser.UserID

'Perform DB Transactions
    'First perform the transaction for Share fee(charges)
    'Note that this will not have any effect on the Balance.
    'IT is here only to determine the profit made by the bank
    gDbTrans.SqlStmt = "INSERT into MemIntTrans " & _
            " (AccID, TransID, TransDate, " & _
            " Amount, TransType, Balance,UserID) " & _
            " values ( " & AccId & "," & _
            TransID & ", " & _
            "#" & TransDate & "#, " & _
            CancellationFee & ", " & _
            wDeposit & ", " & _
            Balance & "," & UserID & ")"
    
    If CancellationFee > 0 Then
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
            gDbTrans.RollBack
            Exit Function
        End If
    End If
    'Insert / Update the the Share leaves table depending upon the transaction type
    For I = 0 To MaxI
        'Update the PurchaseTransID of the ShareLeaves
        gDbTrans.SqlStmt = "Update ShareTrans set ReturnTransID = " & _
            TransID & " where CertNo = " & AddQuotes(m_Share(I), True)
        'Fire the query just built
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to perform transactions !", vbExclamation, gAppName & " - Error"
            MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
            GoTo Exit_Line
        End If
    Next I
    
    'Insert into MemTrans Tab
    gDbTrans.SqlStmt = "INSERT INTO MemTrans (AccID, TransID, TransDate, " & _
            " Leaves, Amount, TransType, Balance,UserID ) values ( " & _
            AccId & "," & _
            TransID & ", " & _
            "#" & TransDate & "#, " & _
            MaxI + 1 & ", " & _
            Amount & ", " & _
            transType & ", " & _
            Balance & "," & UserID & ")"
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        GoTo Exit_Line
    End If
    
    
'IF The Balance Falls Below Zero Ask For account Closure
If Balance <= 0 Then
    'If MsgBox(" This Transaction will Result In Negative Balance " & "Do You Wish To Close The Account ", _
                            vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then
     If MsgBox(GetResourceString(547, 541), _
               vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbYes Then

    End If
End If


'Now Insert Into Ciontra Trans
'update it to te Contra Trans table
If Not CashTransaction Then
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            " (ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & AccHeadID & "," & _
            AccId & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                "(ContraID,AccHeadID, " & _
                "AccID, TransID,Amount,VoucherNo,userID) " & _
                " Values (" & ContraID & "," & FeeHeadId & ", " & _
                AccId & ", " & transType & ", " & TransID & "," & CancellationFee & ", " & _
                AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If CancellationFee > 0 Then _
            If Not gDbTrans.SQLExecute Then GoTo Exit_Line
End If

WithdrawAmount = TransID


Exit_Line:
If Err Then
    MsgBox "ERROR in WithdrawAmount - MMClss", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

End Function

Public Function WithDrawls(fromDate As Date, toDate As Date) As Currency

WithDrawls = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wWithdraw
Amount = TransactionDetails(fromDate, toDate, transType)
transType = wContraWithdraw
Amount = Amount + TransactionDetails(fromDate, toDate, transType)

WithDrawls = Amount
End Function

Public Function ShareFee(fromDate As Date, toDate As Date) As Currency
ShareFee = 0

Dim TransID As Long
Dim Lret As Long

TransID = 1
gDbTrans.SqlStmt = "Select sum(Amount) as deposit from MemIntTrans " & _
    " WHERE TransId > " & TransID & " And TransDate >= #" & fromDate & "#" & _
    " And TransDate <= #" & toDate & "# "
Dim rst As ADODB.Recordset

Lret = gDbTrans.Fetch(rst, adOpenStatic)
If Lret <= 0 Then Exit Function
ShareFee = Val(FormatField(rst("Deposit")))

End Function

Public Function MembershipFee(fromDate As Date, toDate As Date, Optional MemberType As Integer) As Currency
MembershipFee = 0

Dim TransID As Long
Dim Lret As Long

TransID = 1

TransID = 1
gDbTrans.SqlStmt = "Select sum(Amount) as deposit from MemIntTrans " & _
    " WHERE TransId = " & TransID & " And TransDate >= #" & fromDate & "#" & _
    " And TransDate <= #" & toDate & "# "
    
If MemberType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
    " And AccId in (Select Distinct AccId from MemMaster where memberType = " & MemberType & ")"
    
Dim rst As ADODB.Recordset
Lret = gDbTrans.Fetch(rst, adOpenDynamic)
If Lret <= 0 Then Exit Function
MembershipFee = Val(FormatField(rst("Deposit")))

End Function

Public Sub Show()
'gWindowHandle = frmMMAcc.hWnd
Dim Perms As wis_Permissions
Dim StartMem As Boolean
Dim rst As Recordset

StartMem = False

If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers: Exit Sub

    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Perms = gCurrUser.UserPermissions
    If Perms And perBankAdmin Then StartMem = True
    If Perms And perCreateAccount Then StartMem = True
    If Perms And perCashier Then StartMem = True
    If Perms And perClerk Then StartMem = True
    If Perms And perReadOnly Then StartMem = True
    If StartMem Then
        Dim recCount As Integer
        Dim rstMemType As Recordset
        
        gDbTrans.SqlStmt = "Select * From MemberTypeTab"
        recCount = gDbTrans.Fetch(rstMemType, adOpenDynamic)
        
        'TEMPCODE
        If recCount = 0 Then  ' THERE IS NO MEMBER TYPES
            gDbTrans.SqlStmt = "Select count(*) From MemMaster"
            Call gDbTrans.Fetch(rst, adOpenDynamic)
            If FormatField(rst(0)) > 0 Then
                gDbTrans.BeginTrans
                gDbTrans.SqlStmt = "Insert into MemberTypeTab " & _
                    " (MemberType, MemberTypeName,MemberTypeNameEnglish)" & _
                    " Values (" & 0 & "," & _
                    AddQuotes(" ") & "," & _
                    AddQuotes(" ") & ")"
                ''Insert the Record if there are no records
                If recCount < 1 Then Call gDbTrans.SQLExecute
                gDbTrans.CommitTrans
            Else
                MsgBox "Please create the member type before proceeding", vbInformation + vbOKOnly, wis_MESSAGE_TITLE
                Exit Sub
            End If
                    
        End If
        'TEMPCODE ENDS
                        
        If m_FrmMemAcc Is Nothing Then Set m_FrmMemAcc = New frmMMAcc
        Load m_FrmMemAcc
        If gWindowHandle Then m_ParentHandle = gWindowHandle
        gWindowHandle = m_FrmMemAcc.hwnd
        m_FrmMemAcc.Show
                
        If m_AccID Then
            gDbTrans.SqlStmt = "Select A.*,B.DepositTypeName From MemMaster A Inner Join memberTypeTab B" & _
                " on A.DeposityType = B.DepositTYpe Where AccID = " & m_AccID
            
            If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then m_MemberType = rst("MemberType")
            
            With m_FrmMemAcc
                .MemberType = m_MemberType
                .AccountLoad (m_AccID)
                .cmdLoad.Enabled = False
                .txtAccNo.Locked = True
                .cmdSave.Enabled = False
                .cmdAccept.Enabled = False
                .cmdView.Enabled = False
            End With
        Else
            'Set Default Value to the memeber
           
            If recCount > 1 Then
                Dim cancel As Boolean
                Dim memType As Integer
                memType = SelectMemberType 'selectFD.SelectDeposit(wis_Members, grpMember, multiDeposit, Cancel)
                'Set selectFD = Nothing
                If memType < 0 Then
                    Unload m_FrmMemAcc
                    gWindowHandle = 0
                    Exit Sub
                End If
                m_MemberType = memType
                'If m_frmSelectMemType Is Nothing Then Set m_frmSelectMemType = New frmSelectFD
                'm_frmSelectMemType.groupType = grpMember
                'm_frmSelectMemType.Show 1
                'Unload m_frmSelectMemType
            End If
            
            'Set m_frmSelectMemType = Nothing
            
            m_FrmMemAcc.MemberType = m_MemberType
        End If

        m_FrmMemAcc.SingleMemberModule = recCount < 2
        Set rst = Nothing
        Set rstMemType = Nothing
        
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
    End If


End Sub

Private Function SelectMemberType() As Integer
    If m_FrmMemAcc Is Nothing Then Set m_FrmMemAcc = New frmMMAcc
    
    Dim selectMem As New clsSelectDeposit
    Dim multiDeposit As Boolean
    Dim cancel As Boolean
    Dim memType As Integer
    memType = selectMem.SelectDeposit(wis_Members, grpMember, multiDeposit, cancel)
    Set selectMem = Nothing
    
    If memType > -1 Then
        m_FrmMemAcc.MemberType = memType
        m_FrmMemAcc.SingleMemberModule = Not multiDeposit
        If m_MemberType <> memType Then m_FrmMemAcc.txtAccNo = ""
    End If
    m_MemberType = memType
    SelectMemberType = memType
    
End Function

Public Function ContraDeposits(fromDate As Date, toDate As Date) As Currency
ContraDeposits = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wContraDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
ContraDeposits = Amount
End Function

Public Function CashDeposits(fromDate As Date, toDate As Date) As Currency
CashDeposits = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
CashDeposits = Amount
End Function

Public Function Deposits(fromDate As Date, toDate As Date) As Currency
Deposits = 0

Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wDeposit
Amount = TransactionDetails(fromDate, toDate, transType)
transType = wContraDeposit
Amount = Amount + TransactionDetails(fromDate, toDate, transType)
Deposits = Amount
End Function

Private Function IntTransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency

IntTransactionDetails = 0

gDbTrans.SqlStmt = " Select Sum(Amount) as Deposits from MemIntTrans where " & _
    " TransType = " & transType & " And TransDate >= " & _
    "#" & fromDate & "# And TransDate <= #" & toDate & "#"

Dim rst As ADODB.Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

IntTransactionDetails = Val(FormatField(rst("Deposits")))

End Function

Private Function TransactionDetails(fromDate As Date, toDate As Date, transType As wisTransactionTypes) As Currency

TransactionDetails = 0

gDbTrans.SqlStmt = " Select Sum(Amount) as Deposits from MemTrans where " & _
        " TransType = " & transType & " And TransDate >= " & _
        "#" & fromDate & "# And TransDate <= #" & toDate & "#"
Dim rst As ADODB.Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

TransactionDetails = Val(FormatField(rst("Deposits")))

End Function

Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean
    
'Get last transaction record
    Dim Amount As Currency
    Dim ShareFee As Currency
    Dim ret As Integer
    Dim lastTransID As Long
    Dim TransDate As Date
    Dim transType As wisTransactionTypes
    Dim rst As Recordset
    
    gDbTrans.SqlStmt = "Select TOP 1 * from MemTrans where " & _
        " AccID = " & AccId & " order by TransID desc"
    
    If gDbTrans.Fetch(rst, adOpenDynamic) >= 1 Then
        Amount = FormatField(rst("Amount"))
        lastTransID = FormatField(rst("TransID"))
        transType = rst("TransType")
        TransDate = rst("TransDate")
    ElseIf ret = 0 Then
        'MsgBox "No transaction have been performed on this account !", vbInformation, gAppName & " - Error"
        MsgBox GetResourceString(645), vbInformation, gAppName & " - Error"
        Exit Function
    End If
    
    gDbTrans.SqlStmt = "Select TOP 1 * from MemIntTrans where " & _
        " AccID = " & AccId & " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        If lastTransID <= rst("TransID") Then
            If lastTransID < rst("TransID") Then Amount = 0
            transType = rst("TransType")
            lastTransID = rst("TransID")
            ShareFee = FormatField(rst("amount"))
        End If
    End If
    
    If lastTransID <> TransID Then
        MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If

    'Delete the last two records(one will be share purchase/sale and the other will be FEES)
    gDbTrans.SqlStmt = "Delete from MemTrans where AccID = " & AccId & _
            " and TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    gDbTrans.SqlStmt = "Delete from MemIntTrans where AccID = " & AccId & _
            " and TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line

    'Build SQL
    If transType = wContraWithdraw Then
        gDbTrans.SqlStmt = "Update ShareTrans set ReturnTransID = Null " & _
            " where ReturnTransID = " & lastTransID & ";"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    ElseIf transType = wContraDeposit Then
        gDbTrans.SqlStmt = "Delete from ShareTrans where Accid = " & AccId & _
                           " And SaleTransID = " & lastTransID
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
    UndoTransaction = True
    
    Exit Function
    
Exit_Line:
    
End Function


Private Sub Class_Terminate()
On Error Resume Next
If Not m_FrmMemAcc Is Nothing Then
    Unload m_FrmMemAcc
    Set m_FrmMemAcc = Nothing
End If
If Not m_frmMMReport Is Nothing Then
    Unload m_frmMMReport
    Set m_frmMMReport = Nothing
End If
If Not m_frmMMShare Is Nothing Then
    Unload m_frmMMShare
    Set m_frmMMShare = Nothing
End If

End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
    m_selValues(0) = strSelection
End Sub

Private Sub m_frmLookUp_SubItems(strSubItem() As String)
    Dim I As Integer
    For I = LBound(strSubItem) To UBound(strSubItem)
        If I <= UBound(m_selValues) Then m_selValues(I) = strSubItem(I)
    Next I
End Sub

Private Sub m_FrmMemAcc_AccountChanged(ByVal AccId As Long)

m_AccID = AccId

End Sub

Private Sub m_FrmMemAcc_SelectMemberType(memeberType As Integer, cancel As Boolean)
     Call SelectMemberType
End Sub


Private Sub m_FrmMemAcc_ShowReport(ReportType As wis_MemReports, _
                ReportOrder As wis_ReportOrder, MemberType As wis_MemberType, _
                fromDate As String, toDate As String, _
                RepOption As clsRepOption)
gCancel = 0
frmCancel.Show
frmCancel.Refresh
If Not m_frmMMReport Is Nothing Then Unload frmMMReport
Set m_frmMMReport = New frmMMReport

With m_frmMMReport
    .FromIndianDate = fromDate
    .ToIndianDate = toDate
    .FromAmount = RepOption.FromAmount
    .ToAmount = RepOption.ToAmount
    
    .ReportOrder = ReportOrder
    .ReportType = ReportType
    .MemberType = MemberType
    .Place = RepOption.Place
    .Caste = RepOption.Caste
    .Gender = RepOption.Gender
    .AccountGroup = RepOption.AccountGroupID
    
    Load m_frmMMReport
    
    If gCancel = 2 Then GoTo ExitLine
    Unload frmCancel
    .Show 'vbModal
End With


ExitLine:
    On Error Resume Next
    Set frmCancel = Nothing
    Set m_frmMMReport = Nothing
    Exit Sub
End Sub

Private Sub m_FrmMemAcc_WindowClosed()

If m_ParentHandle Then gWindowHandle = m_ParentHandle
m_ParentHandle = 0
Set m_FrmMemAcc = Nothing
End Sub

Private Sub m_frmMMReport_Initialise(Min As Long, Max As Long)
On Error Resume Next
frmCancel.Show

With frmCancel
    If Max <> 0 And Max > Min Then
        '.prg.Visible = True
        '.prg.Min = Min
        'If Max > 32500 Then Max = 32500
        '.prg.Max = Max
    End If
End With
End Sub

Private Sub m_frmMMReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS:" & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
'    If Ratio > 0 Then
'        If Ratio > 1 Then Ratio = Ratio / .prg.Max
'        .prg.Value = Ratio * .prg.Max
'    End If
End With
End Sub

Private Sub m_frmMMShare_ShareIssued(ShareNos() As String, cancel As Boolean)
Dim I As Long
Dim MaxI As Integer
Dim rst As Recordset

MaxI = UBound(ShareNos)
'Check for duplicate certificate numbers
For I = 0 To MaxI
    gDbTrans.SqlStmt = "Select * from ShareTrans where " & _
        " CertNo = " & AddQuotes(ShareNos(I), True)
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        'MsgBox "Some of the Share Certificate numbers you have specified have already been issued !", vbExclamation, gAppName & " - Error"
        gDbTrans.SqlStmt = "Select AccNum From MemMaster Where AcciD = " & rst("AccID")
        Call gDbTrans.Fetch(rst, adOpenDynamic)
        MsgBox GetResourceString(589) & _
            "The Shares are already given to Member No. : " & _
            FormatField(rst("AccNum")), vbExclamation, gAppName & " - Error"
        cancel = True
        Exit Sub
    End If
Next I

m_Share = ShareNos
End Sub


Private Sub m_frmMMShare_ShareReturned(Leaves() As String)
m_Share = Leaves
End Sub


Private Sub m_frmSelectMemType_OKClick(ByVal DepositType As Integer)
     m_MemberType = DepositType
End Sub


