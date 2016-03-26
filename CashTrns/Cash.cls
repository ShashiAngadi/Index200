VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_retVar As Variant
Private m_ShowDenomination As Boolean
Private m_IsTherePassing As Boolean

Private WithEvents m_CashTrans As frmCashTrans
Attribute m_CashTrans.VB_VarHelpID = -1
Private WithEvents m_CashIndex As frmCashIndex
Attribute m_CashIndex.VB_VarHelpID = -1
Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1

'Private m_DbTrans As clsTransact
Private m_DbTrans As Object

Private m_rstPayment As Recordset

Public Property Get ObjectName()
    ObjectName = "CashClass"
End Property


Private Function CashPaid() As Boolean

Dim VoucherNo As String
Dim AccHeadID As Long
Dim AccNum As String
Dim ScrollNo As String

Dim I As Integer

On Error GoTo Exit_Line

Dim TransDate As Date

I = 1

With m_CashTrans
    TransDate = GetSysFormatDate(.txtTransDate)
    AccHeadID = .cmbAccType(I).ItemData(.cmbAccType(I).ListIndex)
    AccNum = .txtAccNo(I)
    ScrollNo = .txtScroll(I)
End With
    
Dim rst As Recordset
Dim AccId As Long

Set rst = GetAccRecordSet(AccHeadID, AccNum)
If rst Is Nothing Then Exit Function
AccId = FormatField(rst("Id"))

Dim transType As wisTransactionTypes
Dim CashTransID As Long
Dim TransID As Long
Dim RecordPos As Long
Dim CounterId As Integer
'Dim ModuleId As wisModules
Dim Amount As Currency

CashTransID = FormatField(m_rstPayment("CashTransID"))
TransID = FormatField(m_rstPayment("TransID"))
Amount = FormatField(m_rstPayment("Amount"))
RecordPos = m_rstPayment.AbsolutePosition
transType = wWithdraw

'Get the Balance of the tranacting account
'Balance = Val(txtBalance(I))
'ModuleId = GetModuleIDFromHeadID(AccHeadID)

If m_ShowDenomination Then
    m_retVar = False
    If m_CashIndex Is Nothing Then Set m_CashIndex = New frmCashIndex
    m_CashIndex.ExpectedCash = Amount * -1
    m_CashIndex.Show 1
    If m_retVar = False Then Exit Function
End If

gDbTrans.BeginTrans

'Now Update the CashTrans as this amount is paid
gDbTrans.SqlStmt = "UPDate CashTrans Set TransUSerID = " & gCurrUser.UserID & "," & _
    " ScrollNo = " & AddQuotes(ScrollNo, True) & ",CounterID = " & CounterId & _
    " Where AccHeadID = " & AccHeadID & " AND AccID = " & AccId & _
    " AND TransID = " & TransID

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

CashPaid = True
 
Exit_Line:

End Function

Private Sub LoadPaymentVoucher(Optional nextRecord As Boolean = True)

If m_rstPayment Is Nothing Then Exit Sub
If m_rstPayment.EOF Then Exit Sub

If nextRecord Then m_rstPayment.MoveNext

If m_rstPayment.EOF Then Exit Sub

Dim AccHeadID As Long
Dim AccId As Long
Dim ModuleId As wisModules
Dim AccNum As String
Dim TableName As String

AccHeadID = FormatField(m_rstPayment("AccheadID"))
AccId = FormatField(m_rstPayment("AccID"))

ModuleId = GetModuleIDFromHeadID(AccHeadID)
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

If ModuleId = wis_CAAcc Then TableName = "CAMaster"
If ModuleId = wis_SBAcc Then TableName = "SBMaster"
If ModuleId = wis_Deposits Then TableName = "FDMaster"
If ModuleId = wis_PDAcc Then TableName = "PDMaster"
If ModuleId = wis_RDAcc Then TableName = "RDMaster"
If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then TableName = "BKCCMaster"
If ModuleId = wis_Members Then TableName = "MemMaster"

If ModuleId = wis_DepositLoans Then TableName = "DepositLoanMaster"
If ModuleId = wis_Loans Then TableName = "LoanMaster"

Debug.Assert TableName <> ""
If TableName = "" Then Exit Sub

gDbTrans.SqlStmt = "Select AccNum, " & _
    " Title +' '+FirstNAme+' '+MiddleName+' '+LastName As CustName" & _
    " From " & TableName & " A, NameTab B " & _
    " Where B.CustomerID = A.CustomerID "

If ModuleId = wis_DepositLoans Or ModuleId = wis_Loans Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
        " AND A.LoanID = " & m_rstPayment("AccID")
Else
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
        " AND A.AccID = " & m_rstPayment("AccID")
End If

Dim rstTemp As Recordset

If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Sub
Dim I As Integer
I = 1

With m_CashTrans
    .txtAccNo(I) = FormatField(rstTemp("AccNum"))
    .txtCustName(I) = FormatField(rstTemp("CustName"))
    .txtAmount(I) = FormatField(m_rstPayment("Amount"))
    .txtScroll(I) = GetNextScroll(wWithdraw)

    With .cmbAccType(1)
        I = 0
        Do
            If I = .ListCount Then Exit Do
            If AccHeadID = .ItemData(I) Then .ListIndex = I: Exit Do
            I = I + 1
        Loop
    End With
End With


End Sub

Private Function GetNextScroll(VoucherType As Wis_VoucherTypes) As String
 Dim transType As wisTransactionTypes
 
transType = IIf(VoucherType = payment, wWithdraw, wDeposit)

gDbTrans.SqlStmt = "Select ScrollNo From CashTrans " & _
    " Where CashTransID = (Select Max(CashTransID) From CashTrans " & _
    " Where TransDate >= #" & FinUSFromDate & "# And TransType = " & transType & " )" & _
    " And TransDate >= #" & FinUSFromDate & "#"

Dim rst As Recordset
Dim ScrollNo As String

ScrollNo = "1"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo ExitLine

ScrollNo = FormatField(rst(0))

If ScrollNo = "" Then GoTo ExitLine

If IsNumeric(ScrollNo) Then ScrollNo = CStr(Val(ScrollNo)): GoTo ExitLine

Dim I As Integer
Dim strAlpha As String

I = 1
If IsNumeric(Left(ScrollNo, 1)) Then
    Do
        If Not IsNumeric(Left(ScrollNo, I)) Then Exit Do
        I = I + 1
    Loop
    strAlpha = Mid(ScrollNo, I)
    
    ScrollNo = Left(ScrollNo, I - 1)
    ScrollNo = CStr(Val(ScrollNo) + 1)
    ScrollNo = strAlpha & ScrollNo
    GoTo ExitLine
End If
If IsNumeric(Right(ScrollNo, 1)) Then
    Do
        If Not IsNumeric(Right(ScrollNo, I)) Then Exit Do
        I = I + 1
    Loop
    strAlpha = Mid(ScrollNo, 1, Len(ScrollNo) - I - 1)
    ScrollNo = Right(ScrollNo, I - 1)
    ScrollNo = CStr(Val(ScrollNo) + 1)
    ScrollNo = ScrollNo & strAlpha
End If


ExitLine:
    GetNextScroll = ScrollNo

End Function


Private Sub RefreshPaymentRecordset()

Dim transType As wisTransactionTypes
transType = wWithdraw

gDbTrans.SqlStmt = "Select CashTransID,AccID,HeadName, TransID," & _
        " Amount,ScrollNo,AccHeadID From CashTrans A,Heads B Where " & _
        " B.HeadID = A.AccHeadid And TransType = " & transType & _
        " And TransID > 0 " & _
        " and (TransUserID is Null or TransUserId =0 )"

If m_IsTherePassing Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And (PassUserId  is NULL or PassUserID = 0)"
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " order by transDate Desc,ScrollNo"
        
Call gDbTrans.Fetch(m_rstPayment, adOpenDynamic)
If m_rstPayment.EOF And m_rstPayment.BOF Then Set m_rstPayment = Nothing
End Sub

Private Sub FillPaymentVoucher()

If m_rstPayment Is Nothing Then Exit Sub
If m_rstPayment.EOF Then Exit Sub

Dim AccHeadID As Long
Dim AccId As Long

With m_CashTrans
    .txtScroll(1) = FormatField(m_rstPayment("ScrollNo"))
    .txtAmount(1) = FormatField(m_rstPayment("Amount"))
End With

AccHeadID = FormatField(m_rstPayment("AccHeadID"))
AccId = FormatField(m_rstPayment("AccID"))

Dim I As Integer
Dim MaxI As Integer

With m_CashTrans.cmbAccType(1)
    MaxI = .ListCount - 1
    .ListIndex = -1
    For I = 0 To MaxI
        If .ItemData(I) = AccHeadID Then
            .ListIndex = I
            Exit For
        End If
    Next
End With

'Now Get the AccountNo And Customer Name
Call GetSetCustomerName(AccHeadID, AccId)


End Sub

Private Function CashReceived() As Boolean

Dim VoucherNo As String
Dim AccHeadID As Long
Dim AccNum As String

Dim TransDate As Date
Dim ScrollNo As String
Dim IsNewAccount As Boolean
Dim PassUserID As Integer

On Error GoTo Exit_Line

With m_CashTrans
    TransDate = GetSysFormatDate(m_CashTrans.txtTransDate)
    AccHeadID = .cmbAccType(0).ItemData(.cmbAccType(0).ListIndex)
    
    AccNum = .txtAccNo(0)
    ScrollNo = .txtScroll(0)
    IsNewAccount = IIf(.chkNew = vbChecked, True, False)
    AccHeadID = .cmbAccType(0).ItemData(.cmbAccType(0).ListIndex)
End With

Dim rst As Recordset
Dim AccId As Long

If Not IsNewAccount Then
    Set rst = GetAccRecordSet(AccHeadID, AccNum)
    If rst Is Nothing Then Exit Function
    AccId = FormatField(rst("Id"))
Else
    'Get the Min AccId From the Cash Trans
    gDbTrans.SqlStmt = "Select Min(AccID) as MinAccID From CashTrans " & _
            " Where AccHeadID = " & AccHeadID & " AND TransDate = #" & TransDate & "#"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then AccId = FormatField(rst("MinAccId"))
    If AccId = 0 Then AccId = -1
End If

Dim transType As wisTransactionTypes

transType = wDeposit

Dim CashTransID As Long
'Get the Max TransCtion ID
gDbTrans.SqlStmt = "Select Max(CashTransID) From CashTrans " & _
        " Where TransDate >= #" & FinUSFromDate & "#"

CashTransID = 1
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then CashTransID = FormatField(rst(0)) + 1

Dim TransID As Long
Dim CounterId As Integer
Dim Balance As Currency
Dim ModuleId As wisModules

'Get the Balance of the tranacting account
'Balance = Val(txtBalance(I))

ModuleId = GetModuleIDFromHeadID(AccHeadID)

'Dim BankClass As clsBankAcc
'Set BankClass = New clsBankAcc
Dim ClsObject As Object

If m_CashTrans.optAccount And AccId < 0 Then
    MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


gDbTrans.BeginTrans

If Not m_CashTrans.optAccount Then GoTo CashLine
If m_IsTherePassing Then GoTo CashLine

Dim IntAmount As Currency
Dim PenalAmount As Currency
Dim PrincAmount As Currency
Dim MiscAmount As Currency

With m_CashTrans
    PrincAmount = .txtAmount(0)
    If .fraInterest.Visible Then
        IntAmount = .txtRegInt
        PenalAmount = .txtPenal
        MiscAmount = .txtMisc
    End If
End With
    
If ModuleId > 100 Then ModuleId = ModuleId - (ModuleId Mod 100)

If ModuleId = wis_PDAcc Then
    Set ClsObject = New clsPDAcc
    TransID = ClsObject.DepositAgentAmount(AccId, PrincAmount, _
                "From Cash Counter " & CounterId, TransDate, VoucherNo, True)
    
ElseIf ModuleId = wis_CAAcc Or ModuleId = wis_SBAcc Or _
    ModuleId = wis_Deposits Or ModuleId = wis_RDAcc Then

    If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then Set ClsObject = New clsBkcc
    If ModuleId = wis_SBAcc Then Set ClsObject = New clsSBAcc
    If ModuleId = wis_Deposits Then Set ClsObject = New clsFDAcc
    If ModuleId = wis_PDAcc Then Set ClsObject = New clsPDAcc
    If ModuleId = wis_CAAcc Then Set ClsObject = New ClsCAAcc
    If ModuleId = wis_RDAcc Then Set ClsObject = New clsRDAcc

    TransID = ClsObject.DepositAmount(AccId, PrincAmount, _
                     "From Cash Counter " & CounterId, TransDate, VoucherNo, True)
    If TransID <= 1 Then GoTo Exit_Line
    Set ClsObject = Nothing

ElseIf ModuleId = wis_DepositLoans Or ModuleId = wis_BKCC Or _
    ModuleId = wis_Loans Or ModuleId = wis_Members Then
    'Now search whether he is Paying any Interest Amount
    'on this loan account
    
    If ModuleId = wis_DepositLoans Then
        Set ClsObject = New clsDepLoan
        TransID = ClsObject.DepositAmount(AccId, PrincAmount, _
            IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
        
    ElseIf ModuleId = wis_Loans Or ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
        If ModuleId = wis_Loans Then Set ClsObject = New clsLoan
        If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then Set ClsObject = New clsBkcc
        
        TransID = ClsObject.DepositAmount(CLng(AccId), PrincAmount, _
            IntAmount, PenalAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
       
    ElseIf ModuleId = wis_Members Then
        Set ClsObject = New clsMMAcc
        TransID = ClsObject.DepositAmount(CLng(AccId), PrincAmount, _
                    IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_Line
        
    End If
    Set ClsObject = Nothing
End If

PassUserID = gUserID

CashLine:

If m_ShowDenomination Then
    m_retVar = False
    If m_CashIndex Is Nothing Then Set m_CashIndex = New frmCashIndex
    m_CashIndex.ExpectedCash = PrincAmount + IntAmount + PenalAmount + MiscAmount
    m_CashIndex.Show 1
    If m_retVar = False Then gDbTrans.RollBacknRaiseError
End If


gDbTrans.SqlStmt = "Insert Into CashTrans " & _
    "(TransDate,CashTransID,AccHeadID," & _
    " AccID,TransType,TransId,Amount," & _
    " ScrollNo,TransUserID,PassUserID,CounterID)" & _
  " VALUES (#" & TransDate & "#, " & CashTransID & "," & AccHeadID & _
    " ," & AccId & "," & transType & "," & IIf(TransID, TransID, "Null") & "," & _
    PrincAmount + IntAmount + PenalAmount & "," & _
    AddQuotes(ScrollNo, True) & "," & _
    gUserID & "," & IIf(PassUserID, PassUserID, "Null") & "," & CounterId & " )"

If Not gDbTrans.SQLExecute Then
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.CommitTrans

CashReceived = True

Exit_Line:

End Function

Public Sub Show()
    
    If m_CashTrans Is Nothing Then Set m_CashTrans = New frmCashTrans
    Load m_CashTrans
    
    m_CashTrans.chkNew.Enabled = m_IsTherePassing
    m_CashTrans.chkNew.Tag = m_IsTherePassing
    
    Call RefreshPaymentRecordset
    Call FillPaymentVoucher
    'error nnn
        'here it not show the cash transaction form
    
    m_CashTrans.Show vbModal
    Set m_CashTrans = Nothing
   
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_CashTrans Is Nothing Then _
            ObjectHandle = m_CashTrans.hwnd
End Property

Private Function GetSetCustomerName(AccHeadID As Long, ByVal AccId As Long) As Recordset

With m_CashTrans
    .txtCustName(1) = ""
    .txtAccNo(1) = ""
End With

On Error Resume Next

Dim rstReturn As Recordset
Dim SqlStr As String
Dim pos As Long
Dim sqlClause As String

Set rstReturn = Nothing
On Error GoTo Exit_Line

Dim ModuleId As wisModules

ModuleId = GetModuleIDFromHeadID(AccHeadID)

Set rstReturn = Nothing

'Members
If ModuleId >= wis_Members And ModuleId < wis_Members + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID, " & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From MemMaster A , NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And AccId = " & AccId

'BKcc Account
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    SqlStr = "Select AccNum, A.LoanID as ID, " & ""
    SqlStr = SqlStr & " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From BKCCMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    
'Current Account
ElseIf ModuleId = wis_CAAcc Then
    SqlStr = "Select AccNum, AccId as Id ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From CAMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

'Deposit Loans
ElseIf ModuleId >= wis_DepositLoans And ModuleId < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID "
    SqlStr = SqlStr & " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From DepositLoanMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    
    If ModuleId > wis_DepositLoans Then _
        sqlClause = " ANd A.DepositType = " & ModuleId - wis_DepositLoans

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From FDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId
    'If ModuleID > wis_Deposits Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_Deposits
'Loan Accounts
ElseIf ModuleId >= wis_Loans And ModuleId < wis_Loans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From LoanMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And LoanId = " & AccId
    If ModuleId > wis_Loans Then _
        sqlClause = " AND A.SchemeID = " & ModuleId - wis_Loans

'Pigmy Accounts
ElseIf ModuleId >= wis_PDAcc And ModuleId < wis_PDAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From PDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

'Recurring Accounts
ElseIf ModuleId >= wis_RDAcc And ModuleId < wis_RDAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From RDMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId

ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID ," & _
            " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
            " From SBMaster A, NameTab B WHERE B.CustomerID = A.CustomerID " & _
            " And Accid = " & AccId
    
End If
    
    gDbTrans.SqlStmt = SqlStr & " " & sqlClause & " ORDER By FirstName"


If gDbTrans.Fetch(rstReturn, adOpenStatic) < 1 Then Exit Function

With m_CashTrans
    .txtCustName(1) = FormatField(rstReturn("custName"))
    .txtAccNo(1) = FormatField(rstReturn("AccNUm"))
End With

Exit_Line:

End Function

Public Function TransferToCashier(ByVal AccHeadID As Long, _
        ByVal AccId As Long, ByVal TransDate As Date, _
        ByVal TransID As Integer, ByVal Amount As Currency) As Long

Dim rst As Recordset

'Now get the max transid
Dim CashTransID As Long

Debug.Assert TransID <> 0
Debug.Assert Amount <> 0

gDbTrans.SqlStmt = "Select Max(CashTransID) From CashTrans " & _
        " Where TransDate >= #" & FinUSFromDate & "#"
CashTransID = 1

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then CashTransID = FormatField(rst(0)) + 1

Dim transType As wisTransactionTypes

transType = wWithdraw

gDbTrans.SqlStmt = "Insert Into CashTrans " & _
    "(TransDate,CashTransID,AccHeadID," & _
    " AccID,TransType,TransId, " & _
    " Amount, TransUserID)" & _
    " VALUES (#" & TransDate & "#, " & CashTransID & "," & AccHeadID & _
    " ," & AccId & "," & transType & "," & TransID & "," & _
    Amount & "," & gUserID & " )"

If Not gDbTrans.SQLExecute Then Exit Function
TransferToCashier = CashTransID

End Function

Private Sub Class_Initialize()

Dim SetUp As New clsSetup

m_ShowDenomination = IIf(UCase(SetUp.ReadSetupValue("General", "CashIndex", "False")) = "FALSE", False, True)
m_IsTherePassing = IIf(UCase(SetUp.ReadSetupValue("General", "Passing", "False")) = "FALSE", False, True)

Set SetUp = Nothing

End Sub

Private Sub m_CashIndex_CancelClicked()

m_retVar = False

End Sub

Private Sub m_CashIndex_OKClicked()
m_retVar = True

End Sub

Private Sub m_CashTrans_AcceptClicked(VoucherType As Wis_VoucherTypes, Cancel As Integer)

If VoucherType = payment Then
    If m_rstPayment Is Nothing Then Exit Sub
    If m_rstPayment.EOF Or m_rstPayment.BOF Then Exit Sub
    If Not CashPaid Then Exit Sub
    Call FillPaymentVoucher
Else
    If Not CashReceived Then Exit Sub
End If

End Sub


Private Sub m_CashTrans_LookUpClick(VoucherType As Wis_VoucherTypes)

If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

Dim I As Byte

I = IIf(VoucherType = Receipt, 0, 1)

Dim headID As wisModules
With m_CashTrans.cmbAccType(I)
    If .ListIndex < 0 Then Exit Sub
    headID = .ItemData(.ListIndex)
End With

On Error GoTo Exit_Line

Dim SqlStr As String

Dim RstCust As Recordset

Dim ModuleId As wisModules
ModuleId = GetModuleIDFromHeadID(headID)
If ModuleId >= wis_PDAcc And ModuleId < wis_PDAcc + 100 Then
    SqlStr = "Select A.UserID as AgentNum, A.UserId as ID, " & _
        "Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
        " From UserTab A,NameTab B " & _
        " Where A.CustomerID = B.CustomerID"
    gDbTrans.SqlStmt = SqlStr
    Call gDbTrans.Fetch(RstCust, adOpenDynamic)
Else
    Set RstCust = GetAccRecordSet(headID, , IIf(Trim(m_CashTrans.txtAccNo(I)) = "", m_CashTrans.txtCustName(I), ""))
End If
    

If RstCust Is Nothing Then
    MsgBox "There are no customers in the " & GetHeadName(headID), _
            vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

Screen.MousePointer = vbHourglass
If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

m_retVar = ""
With m_LookUp
    Call FillView(.lvwReport, RstCust, True)
    .lvwReport.ColumnHeaders(2).Width = 0
    Screen.MousePointer = vbDefault
    .Show 1
End With

Screen.MousePointer = vbDefault

m_CashTrans.txtAccNo(I) = m_retVar
'm_CashTrans.txtCustName (I) = m_retVar
Screen.MousePointer = vbDefault

Exit_Line:

    Screen.MousePointer = vbDefault

End Sub

Private Sub m_CashTrans_SelectPaymentVoucher()

Call RefreshPaymentRecordset
If m_rstPayment Is Nothing Then Exit Sub

If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp
If Not FillViewNew(m_LookUp.lvwReport, m_rstPayment, "CashTransID") Then Exit Sub

m_LookUp.Show 1
If m_retVar = "" Then Exit Sub

m_rstPayment.MoveFirst
m_rstPayment.Find ("CashTransID = " & m_retVar)
If m_rstPayment.EOF Then Exit Sub

Call FillPaymentVoucher

End Sub


Private Sub m_LookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub


