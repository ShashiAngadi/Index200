VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_rstCustBalance As Recordset
Private m_AccID As Long
Private m_DepositType As Integer

Private m_ParentHandle As Long

Private WithEvents m_frmFDAcc As frmFDAcc
Attribute m_frmFDAcc.VB_VarHelpID = -1
Private WithEvents m_frmFDReport As frmFDReport
Attribute m_frmFDReport.VB_VarHelpID = -1
Private WithEvents m_frmSelectFD As frmSelectFD
Attribute m_frmSelectFD.VB_VarHelpID = -1
Private WithEvents m_frmFDClose As frmFdClose
Attribute m_frmFDClose.VB_VarHelpID = -1


Public Property Get ObjectName()
    ObjectName = "FDClass"
End Property


Public Property Get ObjectHandle() As Long
    If Not m_frmFDAcc Is Nothing Then _
                ObjectHandle = m_frmFDAcc.hwnd: Exit Property
    If Not m_frmFDReport Is Nothing Then _
        Unload m_frmFDReport: Set m_frmFDReport = Nothing
End Property

Private Function AddInterestPayableOfFD(AsOnDate As Date, TransDate As Date) As Boolean

Dim rst As ADODB.Recordset
Dim rstTransID As Recordset
Dim CumulativeDeposit As Byte
Dim DimPos As Integer
Dim DepositName As String
Dim DepositNameEnglish As String
Dim MsgStr As String
Dim strParticulars As String
Dim TableName As String

Screen.MousePointer = vbHourglass

If m_DepositType = 0 Then
    MsgStr = "SELECT THE DEPOSIT"
    GoTo ErrLine
End If

If Month(AsOnDate) <> 3 Or Day(AsOnDate) <> 31 Then
    MsgStr = GetResourceString(501)
    GoTo ErrLine
End If

On Error GoTo ErrLine
'declare the variables necessary
'Get theDeposit typem whether cumulative or not
gDbTrans.SqlStmt = "Select * From DepositName WHere DepositID = " & m_DepositType

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    MsgStr = "SELECT THE DEPOSIT"
    GoTo ErrLine
End If

CumulativeDeposit = FormatField(rst("Cumulative"))
DepositName = FormatField(rst("DepositNAme"))
DepositNameEnglish = FormatField(rst("DepositNameEnglish"))

strParticulars = IIf(CumulativeDeposit, "'Interest Added To Account'", "'Interest Payable'")
     TableName = IIf(CumulativeDeposit, " FdMaster ", " FdIntPayable ")

'First Check the Any transaction Has Taken Place After this Date
gDbTrans.SqlStmt = "SELECT B.AccId,TransDate,TransID " & _
                " FROM FDTrans A, FDMaster B " & _
                " WHERE A.AccID = B.AccID AND TransDate >= #" & AsOnDate & "#" & _
                " AND B.DepositType = " & m_DepositType & _
                " ORDER By val(B.AccNum)"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    'Date Of Transaction
    MsgStr = GetResourceString(572)
    GoTo ErrLine
End If

gDbTrans.SqlStmt = "SELECT B.AccId,TransDate,TransID " & _
                " FROM FDIntTrans A, FDMaster B" & _
                " WHERE A.AccID = B.AccID AND TransDate >= #" & AsOnDate & "#" & _
                " AND B.Deposittype = " & m_DepositType & _
                " Order BY val(B.AccNum)"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    'Date Of Transaction
    MsgStr = GetResourceString(572)
    GoTo ErrLine
End If

Dim transType As wisTransactionTypes
Dim rstMain As ADODB.Recordset
Dim rstPayable As ADODB.Recordset
Dim rstInt As ADODB.Recordset
Dim AccId As Long
Dim TransID As Long
Dim tmpTransID As Long

Dim Balance As Currency
Dim PayBalance As Currency

Dim InterestRate As Currency
Dim CreateDate As String
Dim Duration As Long
Dim IntAmount As Currency
Dim TotalInt As Currency

Dim IntBalance As Currency
Dim TotalBalance As Currency

Dim count As Integer

'initialize some variables
gDbTrans.SqlStmt = "SELECT C.Title + ' '+ C.FirstName " & _
                    " +' '+ C.MiddleName+' '+ C.LastName As CustName, " & _
                " Balance, B.TransID, TransType,A.AccNum, A.AccID,RateOfInterest, " & _
                " CreateDate, MaturityDate,EffectiveDate,LastIntDate, TransDate " & _
                " From FDMaster A, FDTrans B, NameTab C Where " & _
                " A.AccId = B.AccId And A.CustomerID = C.CustomerID" & _
                " And ClosedDate Is NULL AND MaturedON Is NULL " & _
                " And CreateDate < #" & AsOnDate & "#" & _
                " And TransID = (Select Max(TransID) From FDTrans D " & _
                " Where D.AccId = A.AccId AND A.DepositType = " & m_DepositType & ")" & _
                " AND Balance > 0 AND A.DepositType = " & m_DepositType & _
                " Order By val(A.AccNum), A.AccID"

If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then
    MsgStr = "There are no accounts"
    GoTo ErrLine
End If

gDbTrans.SqlStmt = "SELECT Balance, B.TransID, TransType," & _
                " A.AccNum, A.AccID,TransDate " & _
                " From FDMaster A, FDIntTrans B Where " & _
                " A.AccId = B.AccId And TransID =  " & _
                "(Select Max(TransID) From FDIntTrans C Where C.AccId = A.AccId " & _
                "AND A.DepositType = " & m_DepositType & ")" & _
                " AND A.DepositType = " & m_DepositType & _
                " Order By A.AccNum, A.AccID"
If gDbTrans.Fetch(rstInt, adOpenStatic) < 1 Then Set rstInt = Nothing

'Now get the interest payable already added in the prevous years
Set rstPayable = Nothing
If Not CumulativeDeposit Then
    gDbTrans.SqlStmt = "SELECT Balance as Payable, B.TransID," & _
                " TransType, A.AccNum, A.AccID, TransDate " & _
                " From FDMaster A, FDIntPayable B Where " & _
                " A.AccId = B.AccId And TransID =  " & _
                "(Select Max(TransID) From FDIntPayable C " & _
                " Where C.AccId = A.AccId AND A.DepositType = " & m_DepositType & ")" & _
                " AND A.Deposittype = " & m_DepositType & _
                " Order By A.AccNum, A.AccID"
    If gDbTrans.Fetch(rstPayable, adOpenStatic) < 1 Then Set rstPayable = Nothing
End If

'Before Adding check whether he has already added the amount
transType = wContraWithdraw
gDbTrans.SqlStmt = "Select * from FDIntTrans Where " & _
                " AccID in (Select AccID From FDMAster " & _
                    " Where DepositType = " & m_DepositType & ")" & _
                " AND TransDate = #" & AsOnDate & "# " & _
                " And TransType = " & transType

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    MsgStr = "Interest " & IIf(CumulativeDeposit, "", "payable") & " already added to the Accounts"
    GoTo ErrLine
End If

'Dim TransDate As Date
Dim LastIntDate As Date
Dim MatDate As Date

'Unload frmIntPayble
Load frmIntPayble
With frmIntPayble
    Call .LoadContorls(rstMain.RecordCount + 1, 20)
    .lblTitle = GetDepositTypeText(m_DepositType) & " " & _
                GetResourceString(375, 47)
    .PutTotal = True
'    .TotalColoumn = m_Cumulative
'    .BalanceColoumn = m_Cumulative
    .Title(0) = GetResourceString(36)
    .Title(1) = GetResourceString(35)
    If CumulativeDeposit Then
        .Title(2) = GetResourceString(43, 42)
        .Title(3) = GetResourceString(47)
        .Title(4) = GetResourceString(52, 43, 42)
    Else
        .Title(2) = GetResourceString(250, 450)
        .Title(3) = GetResourceString(374, 450)
        .Title(4) = GetResourceString(52, 450)
    End If
End With


'lblStatus.Caption = "Computing Interests for
With m_frmFDAcc
    .lblStatus.Caption = GetResourceString(906) & "  .........."
    .txtFailAccIDs.Visible = False
    .prg.Min = 0: .prg.Max = rstMain.RecordCount + 2
End With

If Not frmIntPayble.LoadContorls(rstMain.RecordCount + 1, 15) Then Exit Function


count = 1
Dim DepositNo As Integer
Dim AccNum As String

While Not rstMain.EOF
    AccId = FormatField(rstMain("AccId"))
    Balance = CCur(FormatField(rstMain("Balance")))
    TransID = rstMain("TransID")
    tmpTransID = TransID
    DepositNo = IIf(AccNum = rstMain("AccNum"), DepositNo + 1, 1)
    AccNum = rstMain("AccNum")
    If Balance = 0 Then TransID = 0
    'Optimise the code while Convertin to ADO
    If IsNull(rstMain("LastIntDate")) Then
        LastIntDate = rstMain("EffectiveDate")
    Else
        LastIntDate = rstMain("LastIntDate")
    End If
    
    MatDate = rstMain("MaturityDate")
    
    CreateDate = FormatField(rstMain("CreateDate"))
    InterestRate = FormatField(rstMain("RateofInterest"))
    Debug.Assert AccId <> 1178
    
    'checking interest rate 'case of error occurance
    'Next time write the description of error till then below line will be commented
    If DateDiff("d", rstMain("TransDate"), LastIntDate) < 0 Then _
                        LastIntDate = DateAdd("d", 1, rstMain("TransDate"))
    
    'Now Get The Date Difference
    'On that Difference Get the InterestRate
    'get the Last Interest paid date
    Duration = DateDiff("d", LastIntDate, AsOnDate)
    If Not rstInt Is Nothing Then
        rstInt.MoveFirst
        rstInt.Find "AccID = " & AccId, , adSearchForward
        If Not rstInt.EOF Then
            LastIntDate = rstInt("TransDate")
            If tmpTransID < rstInt("TransID") Then tmpTransID = rstInt("TransID")
            'Check the Last Transction Date
            If DateDiff("d", AsOnDate, LastIntDate) > 0 Then
                'If last transaction date is later than this date
                'then duration will be 0
                'and need not to allow the transaction
                Duration = 0
                TransID = 0
            Else
                Duration = DateDiff("d", LastIntDate, AsOnDate)
            End If
        End If
    End If
    'Compare with Maturauty date
    If DateDiff("d", AsOnDate, MatDate) < 0 Then _
        Duration = DateDiff("d", LastIntDate, MatDate)
    
    If Duration <= 0 Then Duration = 0: TransID = 0
    
    If InterestRate = 0 Then _
        InterestRate = CCur(GetDepositInterestRate(m_DepositType, rstMain("EffectiveDate"), rstMain("MaturityDate")))

    IntAmount = (((InterestRate / 100) * Balance * Duration) / 365) \ 1
    IntAmount = IntAmount \ 1
    
    'Check for the previously added interest payble of this account
    PayBalance = 0
    If Not rstPayable Is Nothing Then
        rstPayable.MoveFirst
        rstPayable.Find "AccID = " & AccId, , adSearchForward
        If Not rstPayable.EOF Then
            PayBalance = FormatField(rstPayable("Payable"))
            If PayBalance < 0 Then PayBalance = 0
            If tmpTransID < rstPayable("TransID") Then tmpTransID = rstPayable("TransID")
        End If
    End If
    
    IntBalance = IntBalance + PayBalance
    TotalBalance = TotalBalance + Balance
    TotalInt = TotalInt + IntAmount
    If TransID Then TransID = tmpTransID + 1
    
    With frmIntPayble
        .AccNum(count) = rstMain("AccNum") & IIf(DepositNo = 1, "", "-" & DepositNo)
        .CustName(count) = FormatField(rstMain("CustName"))
        .Amount(count) = IntAmount
        .KeyData(count) = TransID
        .Balance(count) = IIf(CumulativeDeposit, Balance, PayBalance)
        .Total(count) = IIf(CumulativeDeposit, Balance, PayBalance) + IntAmount
    End With
    
    rstMain.MoveNext
    count = count + 1
    m_frmFDAcc.prg.Value = count
Wend

With frmIntPayble
    count = .grd.Rows - 1
    .CustName(count) = GetResourceString(450, 52, 42) 'Total Interest Payble
    .Amount(count) = TotalInt
    .Balance(count) = IIf(CumulativeDeposit, TotalBalance, IntBalance)
    .Total(count) = TotalInt + IIf(CumulativeDeposit, TotalBalance, IntBalance)
End With

With m_frmFDAcc
    .prg.Value = 0
    .lblStatus.Caption = ""
    .Refresh
End With
 
Screen.MousePointer = vbDefault
frmIntPayble.ShowForm
m_frmFDAcc.Refresh
Screen.MousePointer = vbHourglass

'Now check whehter user has clicked 'OK' or 'CANCEL'
'cancel if he has clicked cancel then exit the function
'if user clicked the ok then it would be hiden
'else it was unloaded
If frmIntPayble.grd.Rows < 3 Then GoTo ExitLine

transType = wContraWithdraw
gDbTrans.BeginTrans

'lblStatus.Caption = "Adding Interest to the accounts..."
m_frmFDAcc.lblStatus.Caption = GetResourceString(907)

'Now Insert the Interest Payble to the Fd accounts
TotalInt = 0
With m_frmFDAcc
    .txtFailAccIDs = 0
    .prg.Min = 0
    .prg.Max = rstMain.RecordCount + 3
End With

rstMain.MoveFirst
'For Count = 0 To UBound(TransID)
count = 1
While Not rstMain.EOF
    AccId = rstMain("AccID")
    DepositNo = IIf(AccNum = rstMain("AccNum"), DepositNo + 1, 1)
    AccNum = rstMain("AccNum")
    With frmIntPayble
        TransID = Val(.KeyData(count))
        IntAmount = Val(.Amount(count))
        Balance = rstMain("Balance")
        PayBalance = 0
        If Not rstPayable Is Nothing Then
            rstPayable.MoveFirst
            rstPayable.Find "AccID = " & rstMain("AccID"), , adSearchForward
            If Not rstPayable.EOF Then PayBalance = rstPayable("Payable")
        End If
    End With
    TotalInt = TotalInt + IntAmount
    If TransID > 0 And IntAmount > 0 Then
        'If Not CumulativeDeposit Then
            transType = wContraWithdraw
            gDbTrans.SqlStmt = "INSERT INTO FDIntTrans (AccID,TransType,TransID," & _
                " TransDate,Amount,Balance, Particulars,UserID)" & _
                " VALUES (" & AccId & "," & transType & "," & _
                TransID + 1 & "," & "#" & TransDate & "#," & _
                IntAmount & "," & 0 & "," & strParticulars & "," & gUserID & " )"
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                GoTo ErrLine
            End If
        'End If
        transType = wContraDeposit
        If CumulativeDeposit Then
            gDbTrans.SqlStmt = "INSERT INTO FDTrans (AccID,TransType,TransID," & _
                " TransDate,Amount,Balance, Particulars,UserID)" & _
                " VALUES (" & AccId & "," & transType & "," & _
                TransID + 1 & ", #" & TransDate & "#," & _
                IntAmount & "," & Balance + IntAmount & "," & strParticulars & "," & gUserID & " )"
        Else
            gDbTrans.SqlStmt = "INSERT INTO FDIntPayable (AccID,TransType,TransID," & _
                " TransDate,Amount,Balance, Particulars,UserID)" & _
                " VALUES (" & AccId & "," & transType & "," & _
                TransID + 1 & ", #" & TransDate & "#," & _
                IntAmount & "," & PayBalance + IntAmount & "," & strParticulars & "," & gUserID & " )"
        End If
'        gDbTrans.SQLStmt = "INSERT INTO " & TableName & "(AccID,TransType,TransID," & _
                    " TransDate,Amount,Balance, Particulars)" & _
                    " VALUES (" & AccId & "," & TransType & "," & _
                    TransID + 1 & ", #" & TransDate & "#," & _
                    IntAmount & "," & Balance + IntAmount & "," & strParticulars & ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            GoTo ErrLine
        End If
    ElseIf TransID = 0 Then
        m_frmFDAcc.txtFailAccIDs = rstMain("AccNum") & IIf(DepositNo = 1, "", ":" & DepositNo) & ","
    End If
    m_frmFDAcc.prg.Value = count
    count = count + 1
    rstMain.MoveNext
Wend

Dim bankClass As clsBankAcc
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim headName As String
Dim headNameEnglish As String

'Now Get the ledger head id of the payble/interest and Depsoit
headName = DepositName & " " & GetResourceString(450)
If Len(DepositNameEnglish) > 0 Then headNameEnglish = DepositNameEnglish & " " & LoadResString(450)

Set bankClass = New clsBankAcc
FromHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemDepIntPaid, 0, wis_Deposits)
If CumulativeDeposit Then
    ToHeadID = bankClass.GetHeadIDCreated(DepositName, DepositNameEnglish, parMemberDeposit, 0, wis_Deposits)
Else
    headName = DepositName & " " & GetResourceString(375, 47)
    If Len(DepositNameEnglish) > 0 Then headNameEnglish = DepositNameEnglish & " " & LoadResourceStringS(375, 47)
    ToHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parDepositIntProv, 0, wis_Deposits)
End If

'Now Make the transaction to the ledger heads
Call bankClass.UpdateContraTrans(FromHeadID, ToHeadID, TotalInt, TransDate)

gDbTrans.CommitTrans

Set bankClass = Nothing

AddInterestPayableOfFD = True

With m_frmFDAcc
    'lblFailAccIDs = ""
    .lblStatus = ""
    If Len(.txtFailAccIDs) > 0 Then .lblStatus = GetResourceString(36) & _
        GetResourceString(92, 544)
    'MsgBox "Interest payble  added success fully", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(450, 637), vbInformation, wis_MESSAGE_TITLE
    
    .txtFailAccIDs.Visible = True
    .prg.Visible = False
End With

ExitLine:

Set rst = Nothing
Set rstInt = Nothing
Set rstMain = Nothing
Set rstPayable = Nothing

Unload frmIntPayble
Set frmIntPayble = Nothing
Screen.MousePointer = vbDefault

Exit Function
With m_frmFDAcc
    .prg.Value = 0
    .lblStatus.Caption = ""
End With
 
ErrLine:

Screen.MousePointer = vbDefault
If Err Then
    MsgBox "Error In FDAccount --Interest payble", vbCritical, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
ElseIf Len(MsgStr) Then
    MsgBox MsgStr, vbInformation, wis_MESSAGE_TITLE
End If

Unload frmIntPayble
Set frmIntPayble = Nothing


End Function


Public Property Let DepositType(NewValue As Integer)
    
    Dim rst As Recordset
    gDbTrans.SqlStmt = "Select * FROM DepositName WHere DepositID = " & NewValue
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then m_DepositType = NewValue
    Set rst = Nothing
    
End Property


Public Function InterestAmountTillDate(ByVal AccountId As Long, _
            ByVal TillDate As Date) As Currency

On Error GoTo ErrLine

'Get the Information abount the Account
Dim rstFD As Recordset
Dim fromDate As Date
Dim rstTemp As Recordset

Dim DepAmount As Currency
Dim Cumulative As wisInstallmentTypes

gDbTrans.SqlStmt = "Select a.*,b.Balance From FDMaster A,FDTrans B Where " & _
                    " A.AccId = " & AccountId & " AND B.AccId = " & AccountId & _
                    " And TransID =(select MAx(transID) From " & _
                        " FdTrans Where AccID = " & AccountId & ")"
                        
If gDbTrans.Fetch(rstFD, adOpenDynamic) < 1 Then Exit Function
DepAmount = rstFD("DepositAmount")

gDbTrans.SqlStmt = "select Cumulative From DepositName " & _
                    " Where DepositID = " & rstFD("DepositType")
Call gDbTrans.Fetch(rstTemp, adOpenDynamic)
Cumulative = FormatField(rstTemp("Cumulative"))

'Get effective date
fromDate = rstFD("EffectiveDate")

Dim IntRate As Single
IntRate = FormatField(rstFD("RateOfInterest"))

'if deposit is closing then only worry about the
'lesser interest rate 'other wise pay it as it is
If DateDiff("D", TillDate, rstFD("MaturityDate")) > 0 Then
    ' If maturity date is later then the closing date
    IntRate = GetDepositInterestRate(rstFD("DepositType"), fromDate, TillDate)
    If IntRate = 0 Then IntRate = FormatField(rstFD("RateOfInterest"))
    If Cumulative Then
        gDbTrans.SqlStmt = "Select * From DepositName Where Cumulative = 0"
'        If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
            IntRate = GetDepositInterestRate(rstTemp("Depositid"), FromDate, TillDate)
    End If
    If IntRate = 0 Then IntRate = FormatField(rstFD("RateOfInterest"))
    IntRate = IntRate - 2
End If

InterestAmountTillDate = DepAmount * _
        CDbl(DateDiff("d", fromDate, TillDate) / 365) * CDbl(IntRate / 100)

Exit Function

If Cumulative <= Inst_No Then Exit Function
fromDate = rstFD("EffectivDate")

Dim FirstDate As Date
Dim Mon As Byte

FirstDate = rstFD("DepositDate")
Mon = Month(FirstDate)
#If junk Then
If Inst = Inst_Monthly Then
    FirstDate = GetSysFirstDate(FirstDate)
    
ElseIf Inst = Inst_BiMonthly Then
    Mon = IIf(Mon Mod 2, Mon + 1, Mon)
    If Mon < 13 Then
        FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
    Else
        FirstDate = "1/1/" & Year(FirstDate) + 1
    End If
ElseIf Inst = Inst_Quartery Then
    Mon = IIf(Mon < 4, 3, IIf(Mon < 7, 6, IIf(Mon < 10, 9, 12)))
    FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
ElseIf Inst = Inst_HalfYearly Then
    Mon = IIf(Mon > 3 And Mon < 10, 9, 3)
    FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
ElseIf Inst = Inst_Yearly Then
    Mon = IIf(Mon > 3 And Mon < 10, 9, 3)
    FirstDate = GetSysFormatDate("1/4/" & Year(FirstDate) + IIf(Mon > 3, 1, 0))
End If
FirstDate = DateAdd("d", -1, FirstDate)
#End If

'Now Calculate the interest amount
Dim IntAmount As Currency
'Get the Rate of interest from Fixed Deposit
gDbTrans.SqlStmt = "Select DepositID From DepositName Where Cumulative = 0"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    IntAmount = GetDepositInterestRate(rstTemp("DepositID"), fromDate, TillDate) - 2
    IntRate = IIf(IntAmount, IntAmount, IntRate)
End If


ErrLine:

Exit Function




End Function

Public Function InterestAmount(ByVal AccountId As Long, _
                Optional ByVal fromDate As Date = "1/1/1000", _
                Optional ByVal toDate As Date = "1/1/1000") As Currency

On Error GoTo ErrLine

'Get the Information abount the Account
Dim rstFD As Recordset
Dim rstTemp As Recordset

Dim DepAmount As Currency
Dim Cumulative As wisInstallmentTypes

gDbTrans.SqlStmt = "Select a.*,b.Balance From FDMaster A,FDTrans B Where " & _
                    " A.AccId = " & AccountId & " AND B.AccId = " & AccountId & _
                    " And TransID =(select MAx(transID) From " & _
                        " FdTrans Where AccID = " & AccountId & _
                        " And TransDate <= #" & toDate & "# )"
                        
If gDbTrans.Fetch(rstFD, adOpenDynamic) < 1 Then Exit Function
DepAmount = rstFD("DepositAmount")
DepAmount = rstFD("Balance")

gDbTrans.SqlStmt = "select Cumulative From DepositName " & _
                    " Where DepositID = " & rstFD("DepositType")
Call gDbTrans.Fetch(rstTemp, adOpenDynamic)
Cumulative = FormatField(rstTemp("Cumulative"))

'if he has not send the from date then the
'last interest paid date or effective will be the will be the from date
If fromDate = "1/1/1000" Then _
    fromDate = IIf(IsNull(rstFD("LastIntDate")), rstFD("EffectiveDate"), rstFD("LastIntDate"))

Dim IntRate As Single
IntRate = FormatField(rstFD("RateOfInterest"))

'if he has not send the TOdate then the
'today or maturity date will be the will be the Todate
If toDate = "1/1/1000" Then toDate = rstFD("MaturityDate")

If DateDiff("D", toDate, rstFD("MaturityDate")) < 0 Then toDate = rstFD("MaturityDate")
If IntRate = 0 Then IntRate = GetDepositInterestRate(rstFD("DepositType"), rstFD("EffectiveDate"), rstFD("MaturityDate"))
InterestAmount = DepAmount * _
        CDbl(DateDiff("d", fromDate, toDate) / 365) * CDbl(IntRate / 100)

If InterestAmount < 0 Then InterestAmount = 0
Exit Function


#If junk Then

    If Cumulative <= Inst_No Then Exit Function
    If ClosingDeposit Then fromDate = rstFD("EffectivDate")
    
    Dim FirstDate As Date
    Dim Mon As Byte
    
    FirstDate = rstFD("DepositDate")
    Mon = Month(FirstDate)
    If Inst = Inst_Monthly Then
        FirstDate = GetSysFirstDate(DateAdd("m", 1, FirstDate))
        
    ElseIf Inst = Inst_BiMonthly Then
        Mon = IIf(Mon Mod 2, Mon + 1, Mon)
        If Mon < 13 Then
            FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
        Else
            FirstDate = "1/1/" & Year(FirstDate) + 1
        End If
    ElseIf Inst = Inst_Quartery Then
        Mon = IIf(Mon < 4, 3, IIf(Mon < 7, 6, IIf(Mon < 10, 9, 12)))
        FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
    ElseIf Inst = Inst_HalfYearly Then
        Mon = IIf(Mon > 3 And Mon < 10, 9, 3)
        FirstDate = GetSysFormatDate("1/" & Mon & "/" & Year(FirstDate))
    ElseIf Inst = Inst_Yearly Then
        Mon = IIf(Mon > 3 And Mon < 10, 9, 3)
        FirstDate = GetSysFormatDate("1/4/" & Year(FirstDate) + IIf(Mon > 3, 1, 0))
    End If
    FirstDate = DateAdd("d", -1, FirstDate)
    
    'Now Calculate the interest amount
    Dim IntAmount As Currency
    'Get the Rate of interest from Fixed Deposit
    gDbTrans.SqlStmt = "Select DepositID From DepositName Where Cumulative = 0"
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        IntAmount = GetDepositInterestRate(rstTemp("DepositID"), fromDate, toDate) - 2
        IntRate = IIf(IntAmount, IntAmount, IntRate)
    End If
#End If


ErrLine:
    Debug.Assert Err.Number
    Err.Clear


End Function


Private Sub Class_Terminate()
On Error Resume Next
If Not m_frmFDAcc Is Nothing Then
    Unload m_frmFDAcc
    Set m_frmFDAcc = Nothing
End If
If Not m_frmFDReport Is Nothing Then
    Unload m_frmFDReport
    Set m_frmFDReport = Nothing
End If

If Not m_frmSelectFD Is Nothing Then
    Unload m_frmSelectFD
    Set m_frmSelectFD = Nothing
End If

End Sub


Private Sub m_frmFDAcc_AccountChanged(ByVal AccId As Long)
m_AccID = AccId
End Sub



Private Sub m_frmFDAcc_AddInterestPayable(ByVal AsOnDate As Date, ByVal TransDate As Date)
Call AddInterestPayableOfFD(AsOnDate, TransDate)
End Sub


Private Sub m_frmFDAcc_FDClose(ByVal AccId As Long, Cancel As Integer)
    
    'Load the for for closure (and leave the rest to the form)
    'now Load the Varaible to the Closing form
    'If account is premature then show FD Close
    If m_frmFDClose Is Nothing Then _
            Set m_frmFDClose = New frmFdClose
    
    m_frmFDClose.AccountId = m_AccID
    Load m_frmFDClose
    m_frmFDClose.Show vbModal
    
End Sub

Private Sub m_frmFDAcc_FDRenew(ByVal AccId As Long, Cancel As Integer)
    Dim rst As Recordset
    'Check Whther This Deposit is there in the Matured FD account
    gDbTrans.SqlStmt = "Select Top 1 * from MatFDTrans where " & _
                "AccID = " & m_AccID & " ORDER BY transID Desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        MsgBox "This account is not present in Matured FD account", _
                vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    If FormatField(rst("Balance")) = 0 Then
        MsgBox "This account has been closed from the Matured FD account"
        Exit Sub
    End If
    
    Set rst = Nothing
    frmFDRenew.AccountId = m_AccID
    frmFDRenew.Show vbModal

End Sub


Private Sub m_frmFDAcc_PayInterest(ByVal AccId As Long, Cancel As Integer)
    
    frmFDInterest.AccountId = m_AccID
    frmFDInterest.Show vbModal
    Set frmFDInterest = Nothing
End Sub


Private Sub m_frmFDAcc_SelectDepositType(ByRef retDepType As Integer)

If m_frmSelectFD Is Nothing Then _
    Set m_frmSelectFD = New frmSelectFD
        
m_frmSelectFD.Show 1
Unload m_frmSelectFD

DepositType = m_DepositType
retDepType = m_DepositType

End Sub

Private Sub m_frmFDAcc_ShowReport(ReportType As wis_FDReports, ReportOrder As wis_ReportOrder, _
                fromDate As String, toDate As String, _
                RepOptClass As clsRepOption)

If Not m_frmFDReport Is Nothing Then Set m_frmFDReport = Nothing
Set m_frmFDReport = New frmFDReport
        
With m_frmFDReport
    
    .DepositType = m_DepositType
    .FromIndianDate = fromDate
    .ToIndianDate = toDate
    .ReportType = ReportType
    .ReportOrder = ReportOrder
    .AccountGroup = RepOptClass.AccountGroupID
    .Place = RepOptClass.Place
    .Caste = RepOptClass.Caste
    .Gender = RepOptClass.Gender
    .FromAmount = RepOptClass.FromAmount
    .ToAmount = RepOptClass.ToAmount

End With

Screen.MousePointer = vbHourglass

Load m_frmFDReport
Screen.MousePointer = vbDefault

Unload frmCancel
If gCancel = 2 Then Exit Sub
m_frmFDReport.Show 'vbModal


End Sub


Private Sub m_frmFDAcc_WindowClosed()
If m_ParentHandle Then gWindowHandle = m_ParentHandle
m_ParentHandle = 0
Set m_frmFDAcc = Nothing
End Sub

Private Sub m_frmFDReport_Initialise(Min As Long, Max As Long)

With frmCancel
    .Show
    .PicStatus.Visible = True
End With
End Sub

Private Sub m_frmFDReport_Processing(strMessage As String, Ratio As Single)

On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS " & vbCrLf & strMessage
    UpdateStatus .PicStatus, Ratio
    .Refresh
    
End With
End Sub

Public Property Let AccountId(NewValue As Long)
    Dim rst As Recordset
    gDbTrans.SqlStmt = "Select * From FDMaster Where AccID = " & NewValue
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Property
    
    m_DepositType = FormatField(rst("DepositType"))
    m_AccID = NewValue
    
End Property



'This function returns the customer balance if the given customer is having any account
'if that customer his having  more than on account
'we can get thos balance in NextBlance function
Public Function CustomerBalance(CustomerID As Long, AsOnDate As Date, retAccNum As String, retDepositName) As Currency

'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
'Sql stmt to get Accid
gDbTrans.SqlStmt = "Select A.Accid,AccNum,A.DepositType,DepositName,Balance " & _
    " from FDMaster A, FDTrans B,DepositName C where A.Customerid=" & CustomerID & _
    " AND A.AccId= B.AccID ANd TransID = (Select Max(TransID) From " & _
        " FDTrans C Where C.AccID = A.AccID " & _
        " And TransDate <= #" & AsOnDate & "#)" & _
    " ANd Balance > 0 And C.DepositID = A.DepositType "

If gDbTrans.Fetch(m_rstCustBalance, adOpenDynamic) < 1 Then Exit Function

CustomerBalance = FormatField(m_rstCustBalance("Balance"))
retAccNum = FormatField(m_rstCustBalance("AccNum"))
retDepositName = FormatField(m_rstCustBalance("DepositName"))
ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If
Resume
End Function

Public Property Let CustomerID(NewValue As Long)
Dim rst As Recordset

gDbTrans.SqlStmt = "Select AccID,DepositType From FDMAster Where " & _
    "CustomerID = " & NewValue

If m_DepositType Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND DepositType = " & m_DepositType

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Property
    
    m_AccID = rst("accId")
    m_DepositType = rst("DepositType")
    
Set rst = Nothing

End Property



'This function is used to Transfer the Amount
'to fixed Deposit from other accounts types is SB,CA
'To do this transfer the account must have created
'But no Deposits has created (means no transaction transwith tht account nO)
Public Function DepositAmount(ByVal AccId As Long, ByVal Amount As Currency, _
        ByVal strDesc As String, ByVal TransDate As Date, _
        Optional VoucherNo As String = "", _
        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

On Error GoTo Exit_Line

'Check For Accoujht Existance
Dim SqlStr As String
Dim rst As Recordset
Dim rstTemp As Recordset

gDbTrans.SqlStmt = "Select DepositName From DepositNAme " & _
        " WHERE DepositId = (Select DepositType From " & _
            " FDMaster Where AccId = " & AccId & " )"

Call gDbTrans.Fetch(rst, adOpenDynamic)
'Deposit Head
Dim AccHeadName As String
AccHeadName = FormatField(rst("DepositName")) & " " & GetResourceString(58)
Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(AccHeadName)

'interest Head
'Dim IntHeadID As Long
'AccHeadName = AccHeadName & " " & GetResourceString(483)
'IntHeadID = GetIndexHeadID(AccHeadName)

Dim ContraID As Long
ContraID = ContraTransID
'Get the Contra Transaction ID
If Not CashTransaction And ContraID = 0 Then ContraID = GetMaxContraTransID

SqlStr = "SELECT * from FDMaster Where AccID = " & AccId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo Exit_Line

'Now Check The TransCion Of that account
'If TransCtion are there in this accoun then

SqlStr = "SELECT * from FDTrans Where AccID = " & AccId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 And CashTransaction Then GoTo Exit_Line


Dim transType As wisTransactionTypes
Dim TransID As Long
Dim Balance As Currency
Dim UserID As Integer

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

TransID = 1
Balance = Amount
UserID = gCurrUser.UserID
If VoucherNo = "" Then VoucherNo = " "

'get the TransCtion Id & check the Whther Alread a Deposit Is Thre are Not
SqlStr = "SELECT Max(TransID) as MaxID from FDTrans Where AccID = " & AccId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then _
            TransID = FormatField(rst("MaxID")) + 1

If TransID = 1 Then
    'Account Created but amount not deposited
    SqlStr = "INSERT INTO FDTrans (AccId,TransType,TransID,Amount," & _
        " Balance,TransDate,VoucherNo, UserId,Particulars )" & _
        " VALUES (" & AccId & ", " & transType & "," & TransID & "," & Amount & ", " & _
        Balance & ", #" & TransDate & "#, " & AddQuotes(VoucherNo, True) & "," & _
        UserID & ", " & AddQuotes(strDesc, True) & ")"
    
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    'Now Update the Deposit Amount in FD MAster
    SqlStr = "UPDATE FDMAster Set DepositAmount = " & Amount & "," & _
            " CreateDate = #" & TransDate & "#," & _
            " EffectiveDate = #" & TransDate & "# " & _
            " WHERE AccID = " & AccId
            
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'update it to te Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & AccHeadID & "," & _
            AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
Else
    'if it is not first depoist so insert the values into FDMaster
    'Get CutomeId of this Account From Fiirst Deposit
    Debug.Assert AccId = 0
    gDbTrans.SqlStmt = "Select * From FDMaster where " & _
        " AccId = " & AccId
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    Dim MatDate As String
    Do
        MatDate = InputBox("Enter the MaturityDate of the deposit", "Maturity Date")
        
        If DateValidate(MatDate, "/", True) Then
            MatDate = GetSysFormatDate(MatDate)
            If DateDiff("d", TransDate, MatDate) > 0 Then Exit Do
        End If
    Loop
    
    gDbTrans.SqlStmt = "Select Max(AccId) From FDMaster"
    Call gDbTrans.Fetch(rstTemp, adOpenForwardOnly)
    AccId = Val(FormatField(rstTemp(0))) + 1
    Set rstTemp = Nothing
    
    SqlStr = "Insert into FDMaster (AccNum,AccID, " & _
        " CustomerID, CreateDate, EffectiveDate,CertificateNo,MaturityDate, " & _
        " DepositAmount,MaturityAmount, RateOfInterest, DepositType,UserID) values ( " & _
        AddQuotes(FormatField(rst("ACCNum")), True) & ", " & _
        AccId & "," & _
        rst("CustomerId") & ", " & _
        "#" & TransDate & "#," & _
        "#" & TransDate & "#," & _
        AddQuotes(FormatField(rst("CertificateNo")), True) & ", " & _
        "#" & MatDate & "#," & _
        Amount & ", " & _
        (Amount * DateDiff("d", TransDate, CDate(MatDate)) / 365 * rst("RateOfInterest") / 100) & "," & _
        rst("RateOfInterest") & "," & _
        rst("DepositType") & "," & gUserID & " )"
    
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    SqlStr = "INSERT INTO FDTrans (AccId,TransType,TransID,Amount," & _
        " Balance,TransDate,VoucherNo, UserId,Particulars )" & _
        " VALUES (" & AccId & ", " & transType & "," & _
        TransID & "," & Amount & ", " & _
        Balance & ", #" & TransDate & "#, " & _
        AddQuotes(VoucherNo, True) & "," & _
        UserID & ", " & _
        AddQuotes(strDesc, True) & ")"

    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If Not CashTransaction Then
        'update it to the Contra Trans table
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            " (ContraID,AccHeadID,AccID, " & _
            "TransType, TransID,Amount,VoucherNo,UserID)" & _
            " Values (" & ContraID & "," & AccHeadID & "," & _
            AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            AddQuotes(VoucherNo, True) & "," & gUserID & " )"
        'Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If
    
End If

DepositAmount = TransID

Exit_Line:

If Err Then
    MsgBox "Error in DepositAmount - FD", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

End Function
Public Function DepositList(DepositIds() As Long, DepositNames()) As Integer
Dim rst As ADODB.Recordset
gDbTrans.SqlStmt = "Select * From DepositName"
DepositList = 0
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

Dim count As Integer
count = 0
While Not rst.EOF
    ReDim Preserve DepositIds(count)
    ReDim Preserve DepositNames(count)
    DepositIds(count) = FormatField(rst("DepositId"))
    DepositNames(count) = FormatField(rst("DepositName"))
    count = count + 1
    rst.MoveNext
Wend
DepositList = count
End Function


'this function returns the next account's balance
'f the given customer (i.e in customerBalance)
Public Function NextBalance(retAccNum As String, retDepositName As String) As Currency

retAccNum = ""
NextBalance = 0
If m_rstCustBalance Is Nothing Then Exit Function
If m_rstCustBalance.EOF Then Exit Function
m_rstCustBalance.MoveNext
If m_rstCustBalance.EOF Then Exit Function

NextBalance = FormatField(m_rstCustBalance("Balance"))
retAccNum = FormatField(m_rstCustBalance("AccNum"))
retDepositName = FormatField(m_rstCustBalance("DepositName"))

End Function

Private Function TransactedAmount_Matured(fromDate As Date, toDate As Date, _
    transType As wisTransactionTypes, Optional DepositType As Integer = 0) As Currency

Dim rst As ADODB.Recordset
Dim SqlStr As String

SqlStr = " Select Sum(Amount) as TotalAmount from MatFDTrans where " & _
    " TransType = " & transType & _
    " And TransDate >= #" & fromDate & "# " & _
    " And TransDate <= #" & toDate & "#"

If DepositType Then
    SqlStr = SqlStr & " AND LoanID In (SELECT LoanId " & _
        " From FDMaster WHERE DepositType = " & DepositType & ")"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

TransactedAmount_Matured = Val(FormatField(rst("TotalAmount")))

End Function

Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean

Dim lastTransID As Long
Dim rst As ADODB.Recordset
Dim PayableHeadID As Long
Dim IntHeadID As Long
Dim MatFdHeadId As Long

Dim headName As String

'Get the last transaction ID
gDbTrans.SqlStmt = "Select Max(TransID) from FDTrans " & _
        " where AccID = " & AccId

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo ExitLine
lastTransID = FormatField(rst(0))
gDbTrans.SqlStmt = "Select Max(TransID) from FDIntTrans " & _
        " where AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    If FormatField(rst(0)) > lastTransID Then lastTransID = rst(0)

gDbTrans.SqlStmt = "Select Max(TransID) from MatFDTrans " & _
        " where AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    If FormatField(rst(0)) > lastTransID Then lastTransID = rst(0)

If lastTransID <> TransID Then
    MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

 'Remove the last transaction From The FDTrans
 gDbTrans.SqlStmt = "Delete from FDTrans where AccID = " & AccId & _
         " and TransID = " & lastTransID
 If Not gDbTrans.SQLExecute Then Exit Function
 
 'Remove the last transaction From The FDIntTrans
gDbTrans.SqlStmt = "Delete from FDIntTrans where AccID = " & AccId & _
         " and TransID = " & lastTransID
If Not gDbTrans.SQLExecute Then Exit Function
 
'Remove the last transaction From The FDIntTrans
gDbTrans.SqlStmt = "Delete from FDIntPayable where AccID = " & AccId & _
      " and TransID = " & lastTransID
 'if amount is withdrawn from the Interes Payble account then
If Not gDbTrans.SQLExecute Then Exit Function
    
'Remove the last transaction From The Matured FD table
gDbTrans.SqlStmt = "Delete from MatFDTrans where AccID = " & AccId & _
      " and TransID = " & lastTransID
 'if amount is withdrawn from the Interes Payble account then
If Not gDbTrans.SQLExecute Then Exit Function
    
gDbTrans.SqlStmt = "Update FDMaster Set ClosedDate = NULL,LastIntDate = Null,MaturedOn = Null where AccID = " & AccId
If Not gDbTrans.SQLExecute Then Exit Function

gDbTrans.SqlStmt = "Select Max(TransDate) From FdIntTrans Where AccID =" & AccId
Call gDbTrans.Fetch(rst, adOpenDynamic)
If Not IsNull(rst(0)) Then
    gDbTrans.SqlStmt = "Update FDMaster Set LastIntDate = " & rst(0) & " where AccID = " & AccId
    If Not gDbTrans.SQLExecute Then Exit Function
End If

UndoTransaction = True

ExitLine:

   Exit Function
ErrLine:

'MsgBox "Error in undo"
MsgBox GetResourceString(561), vbCritical, wis_MESSAGE_TITLE

End Function

Private Function TransactedAmount(fromDate As Date, toDate As Date, _
    transType As wisTransactionTypes, Optional DepositType As Integer = 0) As Currency

Dim rst As ADODB.Recordset
Dim SqlStr As String

SqlStr = " Select Sum(Amount) as TotalAmount from FDTrans where " & _
    " TransType = " & transType & _
    " And TransDate >= #" & fromDate & "# " & _
    " And TransDate <= #" & toDate & "#"

If DepositType Then
    SqlStr = SqlStr & " AND AccID In (SELECT AccId " & _
        " From FDMaster WHERE DepositType = " & DepositType & ")"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

TransactedAmount = Val(FormatField(rst("TotalAmount")))

End Function


Private Function TransactedAmount_Interest(fromDate As Date, toDate As Date, _
    transType As wisTransactionTypes, Optional DepositType As Integer = 0) As Currency

Dim rst As ADODB.Recordset
Dim SqlStr As String

SqlStr = " Select Sum(Amount) as TotalAmount from FDIntTrans where " & _
    " TransType = " & transType & _
    " And TransDate >= #" & fromDate & "# " & _
    " And TransDate <= #" & toDate & "#"

If DepositType Then
    SqlStr = SqlStr & " AND LoanID In (SELECT LoanId " & _
        " From FDMaster WHERE DepositType = " & DepositType & ")"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

TransactedAmount_Interest = Val(FormatField(rst("TotalAmount")))

End Function
Public Sub Show()
If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers: Exit Sub

'Now Check the permission
'whether logged inperson is having permission of sb account or not
Dim Perms As wis_Permissions
Perms = gCurrUser.UserPermissions
If Perms And perBankAdmin Then GoTo ExecutLine
If Perms And perCreateAccount Then GoTo ExecutLine
If Perms And perClerk Then GoTo ExecutLine
If Perms And perCashier Then GoTo ExecutLine
If Perms And perReadOnly Then GoTo ExecutLine

MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
Exit Sub

ExecutLine:

Dim Amount As Currency


If m_frmFDAcc Is Nothing Then Set m_frmFDAcc = New frmFDAcc

Load m_frmFDAcc

If gWindowHandle Then m_ParentHandle = gWindowHandle
gWindowHandle = m_frmFDAcc.hwnd

m_frmFDAcc.Show

If m_AccID = 0 Then
    If m_frmSelectFD Is Nothing Then Set m_frmSelectFD = New frmSelectFD
    m_frmSelectFD.Show 1
    Unload m_frmSelectFD
    m_frmFDAcc.DepositType = m_DepositType
    Exit Sub
End If
Dim rst As Recordset
gDbTrans.SqlStmt = "Select * From FDMaster Where AccID = " & m_AccID
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub
With m_frmFDAcc
    .DepositType = rst("DepositType") Mod 100
    Call .AccountLoad(m_AccID)
    Set rst = Nothing
End With


End Sub

'This function returns total amount earned as profit
'by Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
'this return the anount deposited through cash only
Public Function CashProfit(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wDeposit
    Amount = TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    CashProfit = Amount
End Function
'This function returns total amount deposited to the
'Deposit Loans if DepositId is specified
'Then particular deposit type detals will be returned
'this return the anount deposited through cash only
Public Function CashDeposits(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wDeposit
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    
    CashDeposits = Amount
End Function

'This function returns total amount deposited as profit to the
'Deposit Loans if DepositId is specified
'Then particular deposit type detals will be returned
'This Returns the amount which deposited through interena transfer only
Public Function ContraProfit(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    
    transType = wContraDeposit
    Amount = TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    ContraProfit = Amount
End Function

'This function returns total amount deposited to the
'Deposit Loans if DepositId is specified
'Then particular deposit type detals will be returned
'This Returns the amount which deposited through interena transfer only
Public Function ContraDeposits(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    
    transType = wContraDeposit
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    ContraDeposits = Amount
End Function

'This function returns total amount withdraw as loss from the
'Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
'This Retuns the amount withdrawn in form cash
Public Function CashLoss(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wWithdraw
    Amount = TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    
    CashLoss = Amount
End Function

'This function returns total amount withdrawn from the
'Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
'This Retuns the amount withdrawn in form cash
Public Function CashWithdrawls(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wWithdraw
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    
    CashWithdrawls = Amount
End Function

'This function returns total amount withdrawn as loss from the
'Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
'This functin returns tha amount transferrd to other accounts inside the bank
Public Function ContraLoss(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wContraWithdraw
    Amount = TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    ContraLoss = Amount
End Function

'This function returns total amount withdrawn from the
'Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
'This functin returns tha amount transferrd to other accounts inside the bank
Public Function ContraWithdrawls(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wContraWithdraw
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    ContraWithdrawls = Amount
End Function


'This function returns total amount withdrawn from the
'Deposit Loans As loan if DepositId is specified
'Then particular deposit type details will be returned
Public Function Loss(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wWithdraw
    Amount = TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    
    transType = wContraWithdraw
    Amount = Amount + TransactedAmount_Interest(fromDate, toDate, transType, DepositType)
    Loss = Amount
End Function

'This function returns total amount withdrawn from the
'Deposit Loans if DepositId is specified
'Then particular deposit type details will be returned
Public Function WithDrawls(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wWithdraw
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    
    transType = wContraWithdraw
    Amount = Amount + TransactedAmount(fromDate, toDate, transType, DepositType)
    WithDrawls = Amount
End Function


'This function returns total amount deposited as profit to the
'Deposit Loans if DepositId is specified
'Then particular deposit type detals will be returned
Public Function Profit(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Amount = 0
    transType = wDeposit
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    
    transType = wContraDeposit
    Amount = Amount + TransactedAmount(fromDate, toDate, transType, DepositType)
    Profit = Amount
End Function

'This function returns total amount deposited to the
'Deposit Loans if DepositId is specified
'Then particular deposit type detals will be returned
Public Function Deposits(fromDate As Date, toDate As Date, _
                Optional DepositType As Integer = 0) As Currency
    Dim transType As wisTransactionTypes
    Dim Amount As Currency
    Deposits = 0
    transType = wDeposit
    Amount = TransactedAmount(fromDate, toDate, transType, DepositType)
    
    transType = wContraDeposit
    Amount = Amount + TransactedAmount(fromDate, toDate, transType, DepositType)
    Deposits = Amount
End Function


'this function returns the amount deposited to Matured Deposits
'the amount deposited to this account will be contra all the time
Public Function ContraDepositMaturedFD(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency

'Matured FD trnasferred to matured FD account
ContraDepositMaturedFD = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wContraDeposit
Amount = TransactedAmount_Matured(fromDate, toDate, transType, DepositType)

ContraDepositMaturedFD = Amount
End Function
'this function returns the amount deposited to Matured Deposits
'the amount deposited to this account will be contra all the time
Public Function CashDepositMaturedFD(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency

'Matured FD trnasferred to matured FD account
CashDepositMaturedFD = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wDeposit
Amount = TransactedAmount_Matured(fromDate, toDate, transType, DepositType)

CashDepositMaturedFD = Amount

End Function


'this function returns the amount deposited to Matured Deposits
'the amount deposited to this account will be contra all the time
Public Function DepositMaturedFD(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency

'Matured FD trnasferred to matured FD account
DepositMaturedFD = 0
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wDeposit
Amount = TransactedAmount_Matured(fromDate, toDate, transType, DepositType)
transType = wContraDeposit
Amount = Amount + TransactedAmount_Matured(fromDate, toDate, transType, DepositType)

DepositMaturedFD = Amount

End Function

'This function returns then amount withdrawn
'from the Matured account account
Public Function WithDrawlMaturedFD(fromDate As Date, toDate As Date, Optional DepositType As Integer = 0) As Currency
'Matured Fixed Deposit Given back to the customer or renewd as Fd
Dim transType As wisTransactionTypes
Dim Amount As Currency

transType = wWithdraw
Amount = TransactedAmount_Matured(fromDate, toDate, transType, DepositType)
transType = wContraWithdraw
Amount = Amount + TransactedAmount_Matured(fromDate, toDate, transType, DepositType)

WithDrawlMaturedFD = Amount

End Function

Public Function Balance(AsOnDate As Date, Optional DepositType As Integer = 0) As Currency
' fixed deposits in the bank as on the given date
Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim TotalBalance As Currency

SqlStmt = "Select SUM(Balance) From FDTrans A " & _
    " Where TransID = (Select Max(TransID) from FDTrans C " & _
        " where A.Accid = C.Accid " & _
        " and TransDate <= #" & AsOnDate & "# )"
If DepositType Then
    SqlStmt = SqlStmt & " ANd AccId In (Select AccId " & _
        " From FDMaster Where DepositType = " & DepositType & ")"
End If

gDbTrans.SqlStmt = SqlStmt
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    Balance = FormatField(rst(0))

End Function

Public Function BalanceMaturedFD(AsOnDate As Date, Optional DepositType As Integer = 0) As Currency

Dim rst As ADODB.Recordset

gDbTrans.SqlStmt = "Select SUM(Balance) From MatFDTrans A " & _
    " Where TransID = (Select Max(TransID) from MatFDTrans C " & _
        " where C.Accid = A.Accid " & _
        " AND TransDate <= #" & AsOnDate & "#)"
If DepositType Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AccID In " & _
        "(select AccID From FdMaster where DepositType = " & DepositType & ")"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

BalanceMaturedFD = FormatField(rst(0))

End Function

Public Sub TotalDepositTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional DepositType As Integer)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from FDTrans where "

If DepositType > 0 Then
    SqlStr = SqlStr & " AccID In " & _
        "(select distinct(AccID) From FDMaster where DepositType = " & DepositType & ") And "
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
    
    rst.MoveNext
Wend

Set rst = Nothing
End Sub
Public Sub TotalDepositIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional DepositType As Integer = 0)

CashDeposit = 0
cashWithdraw = 0
contraDeposit = 0
contraWithdraw = 0

Dim rst As ADODB.Recordset
Dim SqlStr As String
SqlStr = "Select Sum(Amount) as TotalAmount,TransType from FDTrans where "

If DepositType > 0 Then
    SqlStr = SqlStr & " AccID In " & _
        "(select distinct(AccID) From FDMaster where DepositType = " & DepositType & ") And "
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
    
    rst.MoveNext
Wend

Set rst = Nothing
End Sub
Public Sub TotalLoanTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional DepositType As Integer)
    Dim depLoan As New clsDepLoan
    Call depLoan.TotalLoanTransactions(fromDate, toDate, CashDeposit, contraDeposit, cashWithdraw, contraWithdraw, DepositType)
    Set depLoan = Nothing
End Sub
Public Sub TotalLoanIntTransactions(fromDate As Date, toDate As Date, ByRef CashDeposit As Currency, ByRef contraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef contraWithdraw As Currency, Optional DepositType As Integer)
    Dim depLoan As New clsDepLoan
    Call depLoan.TotalLoanIntTransactions(fromDate, toDate, CashDeposit, contraDeposit, cashWithdraw, contraWithdraw, DepositType)
    Set depLoan = Nothing
End Sub

Private Sub m_frmSelectFD_OKClick(ByVal DepositType As Integer)
    m_DepositType = DepositType Mod 100
    
End Sub

