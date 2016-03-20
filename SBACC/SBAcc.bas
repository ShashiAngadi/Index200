Attribute VB_Name = "basSBAcc"
Option Explicit

Public Enum wis_SBReports
    repSBBalance = 1
    repSBDayBook = 2
    repSBLedger = 3
    repSBAccOpen = 4
    repSBAccClose = 5
    repSBJoint = 6
    repSBProduct = 7
    repSBCheque = 8
    repSbMonthlyBalance = 9
    repSBSubCashBook = 10
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
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM SBTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM SBPLTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If


ErrLine:

End Sub

'This Function Returns the Max Transction ID of
'the given Member share account Id
'In case there is no transaction it reurns 0
Public Function GetSBMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetSBMaxTransID = TransID

End Function



'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetSBLastTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetSBLastTransDate = TransDate

End Function



Public Function SBAccountExists(ByVal AccId As Long, Optional ClosedON As String) As Boolean
Dim ret As Integer
Dim rst As ADODB.Recordset

'Query Database
    gDbTrans.SqlStmt = "Select * from SBMaster where " & _
                        " AccID = " & AccId
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Check the closed status
    If Not IsMissing(ClosedON) Then
        ClosedON = FormatField(rst("ClosedDate"))
    End If

SBAccountExists = True
End Function


Public Function ComputeSBINT(fromDate As Date, toDate As Date) As Currency

Dim rstMain As ADODB.Recordset
Dim rstMaster As ADODB.Recordset
Dim FromDatePM As String
Dim TransID As Long
Dim TempBalance As Currency

FromDatePM = DateAdd("d", "-1", fromDate)

gDbTrans.SqlStmt = "SELECT MAX(TransID) AS MaxTransID, AccID, " & _
    "Month(TransDate) AS transmonth " & _
    "From SBTrans " & _
    "WHERE TransDate Between #" & FromDatePM & "# And #" & toDate & "# " & _
    "GROUP BY accid, Month(transdate);"
If Not gDbTrans.CreateView("QryPMID") Then Exit Function

gDbTrans.SqlStmt = "SELECT a.AccID AS AccIDPM, a.Balance AS BalancePM, " & _
    "Month(TransDate) AS TransMonthPM " & _
    "FROM SBTrans AS a, qrysbpmid AS b " & _
    "Where (((a.ACCID) = b.ACCID) " & _
    "And ((Month(TransDate)) = b.transmonth) " & _
    "And ((a.TransID) = b.maxtransid)) " & _
    "ORDER BY a.AccID;"
If Not gDbTrans.CreateView("QryPM") Then Exit Function

gDbTrans.SqlStmt = "SELECT MAX([TransID]) AS MaxTransID, [AccID], " & _
    "Month([TransDate]) AS TransMonth " & _
    "From SBTrans " & _
    "WHERE TransDate Between #" & fromDate & "# And #" & toDate & "# " & _
    "And day(transdate) Between 1 And 10 " & _
    "GROUP BY [accid], month([transdate]);"
If Not gDbTrans.CreateView("QrySB1TO10ID") Then Exit Function

gDbTrans.SqlStmt = "SELECT a.AccID AS AccID1to10, Balance AS Balance1to10, " & _
    "Month(TransDate) AS TransMonth1to10 " & _
    "FROM SBTrans AS a, qrysb1to10ID AS b " & _
    "Where a.TransID = b.maxtransid " & _
    "And a.ACCID = b.ACCID " & _
    "And Month(a.TransDate) = b.transmonth " & _
    "ORDER BY a.accid;"
If Not gDbTrans.CreateView("QrySB1TO10") Then Exit Function

gDbTrans.SqlStmt = "SELECT min(Balance) AS Balance11to30, " & _
    "AccID, Month(TransDate) AS TransMonth11to30 " & _
    "From SBTrans " & _
    "WHERE TransDate Between #" & fromDate & "# And #" & toDate & "#" & _
    "And day(transdate) Between 11 And 31 " & _
    "GROUP BY accid, month(transdate);"
If Not gDbTrans.CreateView("QrySB11TO31") Then Exit Function

gDbTrans.SqlStmt = "SELECT AccIDPM, TransMonthPM, BalancePM, " & _
    "Balance1TO10, Balance11TO30 " & _
    "FROM (qrySBPM LEFT JOIN qrySB1to10 ON " & _
    "(qrySBPM.AccIDPM=qrySB1to10.AccID1to10) " & _
    "AND (qrySBPM.TransMonthPM=qrySB1to10.TransMonth1to10)) " & _
    "LEFT JOIN qrySB11to30 ON " & _
    "(qrySB1to10.AccID1to10=qrySB11to30.AccID) " & _
    "AND (qrySB1to10.TransMonth1to10=qrySB11to30.TransMonth11to30);"
If Not gDbTrans.CreateView("QrySBProducts") Then Exit Function

gDbTrans.SqlStmt = "SELECT * FROM QrySBProducts"
Call gDbTrans.Fetch(rstMain, adOpenStatic)

gDbTrans.SqlStmt = "Select AccID FROM SBtrans ORDER BY AccID"
Call gDbTrans.Fetch(rstMaster, adOpenForwardOnly)

Dim StartDate As Date
Dim StartMonth As Byte
Dim EndMonth As Byte
Dim I As Byte

StartMonth = Month(fromDate)
EndMonth = Month(toDate)

Do While Not rstMaster.EOF
    StartDate = fromDate
    Do While (DateDiff("m", StartDate, toDate) < 0)
        StartMonth = Month(StartDate)
        
    Loop
    rstMaster.MoveNext
Loop

End Function

Public Function GetSBInterestChanged(fromDate As Date, InterestRate As Single) As Boolean
'This Function Talks With ClsInterest To Dump The Values Into Interest Tab
'It Is Necessary To Get The ModuleID ,SchemeName,FromIndianDate ,To Indian Date

Dim ClsInt As New clsInterest
Dim SBModule As Integer
Dim SchemeName As String

         '1) Get The ModuleID
         SBModule = wis_SBAcc
         
         '2) Get The SchemeName
         SchemeName = "SBAcc Interest"
         
         '3) Get The Dates Validated
         'If Not DateValidate(FromDate, "/", False) Then GoTo ErrLine
         
         '4) Pass The NEcessary Values To ClsInt.saveInterest
         If Not ClsInt.SaveInterest(SBModule, SchemeName, InterestRate, , , fromDate) Then GoTo ErrLine
         
         GetSBInterestChanged = True
         
ErrLine:

Set ClsInt = Nothing

End Function


Public Function ComputeSBInterest(ByVal Product As Currency, Rate As Double) As Currency
    ComputeSBInterest = (Product * 1 * Rate) / (100 * 12)
End Function

Public Function ComputeSBProducts_New(AccIDArr() As Long, Mon As Integer, _
    Yr As Integer, ByRef Products() As Currency, MaxTrans As Integer, NoInterestOnMinBal As Boolean) As Currency
Dim I As Long
Dim rst_Main As ADODB.Recordset
Dim rst_PM As ADODB.Recordset
Dim rst_1_10 As ADODB.Recordset
Dim rst_11_30 As ADODB.Recordset
Dim TransDate1 As String
Dim TransDate2 As String
Dim TransID As Long
Dim TempBalance As Currency
Dim SqlStr As String
Dim MinBalance As Currency
'Validate month and year
    If Mon < 1 Or Mon > 12 Then Exit Function

'First get rec set of all the accounts in SB
    Dim L_frmcancel As New frmCancel
    Set rst_Main = Nothing
    gDbTrans.SqlStmt = "Select A.AccID, ClosedDate from SBMaster A " & _
        " where AccID in (Select Distinct AccID From SBTRANS) order BY A.AccID"
    Call gDbTrans.Fetch(rst_Main, adOpenForwardOnly)
    'L_frmcancel.prg.Max = rst_Main.RecordCount
    
    Dim SetUp As New clsSetup
    MinBalance = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithChequeBook", "0.00")
    Set SetUp = Nothing
    If MinBalance < 10 Then NoInterestOnMinBal = False
    
On Error Resume Next
L_frmcancel.Show
L_frmcancel.Refresh
L_frmcancel.PicStatus.Visible = True
'L_frmcancel.prg.Min = 0
'L_frmcancel.prg.Max = 20
L_frmcancel.Refresh
L_frmcancel.lblMessage.Caption = "Computing Sb Products"
    
'Get Balance upto Last Day of PrevMonth
'This query gets the maximum transaction performed on or before the
'last day of the previous month specified for every account
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    TransDate1 = DateAdd("d", -1, TransDate1)
    
    gDbTrans.SqlStmt = "Select B.AccID, MAX(TransID) As MaxTransID " & _
                " FROM SBTrans A, SBMaster B " & _
                " WHERE A.AccID = B.AccID And TransDate <= #" & TransDate1 & "#" & _
                " AND (ClosedDate is NULL OR CLosedDate <= #" & TransDate1 & "#)" & _
                " GROUP BY B.AccID"
    
    If Not gDbTrans.CreateView("QrySBPMIDs") Then Exit Function
    
    Set rst_PM = Nothing
    gDbTrans.SqlStmt = "Select B.AccID, Balance from SBTrans A, QrySBPMIDs B " & _
                    " WHERE A.TransID = B.MaxTransID " & _
                    " AND A.AccID = B.AccID " & _
                    " ORDER BY B.AccID"
    Call gDbTrans.Fetch(rst_PM, adOpenForwardOnly)

'Get max Balance between 1 to 10
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    TransDate2 = GetSysFormatDate("10/" & Mon & "/" & Yr)
    gDbTrans.SqlStmt = "Select A.AccID, MAX(TransID) AS MaxTransID " & _
                " FROM SBMaster A, SBTrans B " & _
                " WHERE A.AccID = B.AccID and TransDate between #" & _
                TransDate1 & "# and #" & TransDate2 & "# GROUP BY A.AccID" & _
                " Order By A.AccID"
    
    If Not gDbTrans.CreateView("QrySB1TO10IDs") Then Exit Function
'Build the Query for max trans between 1 and 10
    Set rst_1_10 = Nothing
    gDbTrans.SqlStmt = "Select B.AccID, Balance from SBTrans A, QrySB1TO10IDs B " & _
                    " WHERE A.Transid = B.MaxTransID " & _
                    " AND A.AccID = B.AccID ORDER by B.AccID"
    Call gDbTrans.Fetch(rst_1_10, adOpenForwardOnly)

'Get balance upto LAST DAY of current month
'Set the TransCction dates according to month
    TransDate1 = GetSysFormatDate("11/" & Mon & "/" & Yr)
    TransDate2 = GetSysLastDate(TransDate1)                 'Last Day oo the current month
    
    Set rst_11_30 = Nothing
    gDbTrans.SqlStmt = "Select MIN(Balance) as MinBalance,AccID from SBTrans  " & _
                " where TransDate between #" & _
                TransDate1 & "# and #" & TransDate2 & "# " & _
                "Group By AccID ORDER by AccID"
    Call gDbTrans.Fetch(rst_11_30, adOpenForwardOnly)

'**************************************************************
If MaxTrans > 0 Then
    'Get 1 and Last Day of current month
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    TransDate2 = GetSysLastDate(TransDate1)
    
    Dim transType As wisTransactionTypes
    Dim rst_Trans As ADODB.Recordset
    Dim rst_Trans_Count As ADODB.Recordset
    
    transType = wWithdraw
    
    'Get transactions for all the accounts made during that month
    gDbTrans.SqlStmt = "Select AccID, TransDate from SBTrans where " & _
                    "TransDate >= #" & TransDate1 & "# and " & _
                    "TransDate <= #" & TransDate2 & "# and " & _
                    "TransType = " & transType & _
                    " order by AccID, TransDate, TransID"
    If gDbTrans.Fetch(rst_Trans, adOpenForwardOnly) <= 0 Then Set rst_Trans = Nothing
    
    'Get Count of transactions for all the accounts made during that month
    gDbTrans.SqlStmt = "Select Count(*) as TotalTrans, AccID from SBTrans where " & _
                    "TransDate >= #" & TransDate1 & "# and " & _
                    "TransDate <= #" & TransDate2 & "# and " & _
                    "TransType = " & transType & " group by AccID" & _
                    " Order By AccID"
    If gDbTrans.Fetch(rst_Trans_Count, adOpenForwardOnly) <= 0 Then _
                            Set rst_Trans_Count = Nothing

End If
'**************************************************************
Dim Balance As Currency
Dim AccId As Long
Dim ClosedDate As String
Dim Balance1TO10 As Currency
Dim Balance11TO31 As Currency
Dim rst As ADODB.Recordset
'***********************************************************
Dim Day7 As Date, Day14 As Date, Day21 As Date, Day30 As Date
If MaxTrans > 0 Then
    Day7 = GetSysFormatDate("7/" & Mon & "/" & Yr)
    Day14 = GetSysFormatDate("14/" & Mon & "/" & Yr)
    Day21 = GetSysFormatDate("21/" & Mon & "/" & Yr)
    Day30 = GetSysLastDate(Day21)                   'Last day of the Month
End If
Dim Count7 As Integer, Count14 As Integer, Count21 As Integer, Count30 As Integer
'***********************************************************

'Loop through all the accounts to calculate the products
    For I = 1 To rst_Main.RecordCount

        'L_frmcancel.prg.Value = rst_Main.AbsolutePosition
        UpdateStatus L_frmcancel.PicStatus, I / rst_Main.RecordCount
        Balance = 0: Balance1TO10 = 0: Balance11TO31 = 0
        AccId = FormatField(rst_Main("AccID"))
        If Not rst_PM Is Nothing Then
            If Not rst_PM.EOF Then
                If AccId = FormatField(rst_PM("AccID")) Then
                    Balance = FormatField(rst_PM("Balance"))
                    rst_PM.MoveNext
                End If
            End If
        End If
        
        If Not rst_1_10 Is Nothing Then
            If Not rst_1_10.EOF Then
                If AccId = FormatField(rst_1_10("AccID")) Then
                    Balance1TO10 = FormatField(rst_1_10("Balance"))
                    rst_1_10.MoveNext
                End If
            End If
        End If
        
        If Balance1TO10 > 0 Then Balance = Balance1TO10
        
        If Not rst_11_30 Is Nothing Then
            If Not rst_11_30.EOF Then
                If AccId = FormatField(rst_11_30("AccID")) Then
                    Balance11TO31 = FormatField(rst_11_30("MinBalance"))
                    rst_11_30.MoveNext
                End If
            End If
        End If

        If Balance11TO31 > 0 Then _
            Balance = IIf(Balance < Balance11TO31, Balance, Balance11TO31)

'*****************************************************************
    Dim TotalTrans As Integer
    Dim CheckOut As Boolean
If MaxTrans > 0 Then
    CheckOut = False
    TotalTrans = 0

    Count7 = 0: Count14 = 0: Count21 = 0: Count30 = 0
    If Not rst_Trans_Count Is Nothing Then
        If Not rst_Trans_Count.EOF Then
            If AccId = FormatField(rst_Trans_Count("AccID")) Then
                TotalTrans = FormatField(rst_Trans_Count("TotalTrans"))
                If TotalTrans > MaxTrans * 4 Then
                ElseIf TotalTrans <= MaxTrans Then
                    'Do nothing
                Else
                    CheckOut = True
                End If
                rst_Trans_Count.MoveNext
            End If
        End If
    End If
    
    If CheckOut Then 'Traverse thro the rec set
        If Not rst_Trans Is Nothing Then
            Do
                If rst_Trans.EOF Then Exit Do
                
                If FormatField(rst_Trans("AccID")) <> AccId Then Exit Do
                
                If DateDiff("d", Day7, rst_Trans("TransDate")) <= 0 Then
                    Count7 = Count7 + 1
                ElseIf DateDiff("d", Day14, rst_Trans("TransDate")) <= 0 Then
                    Count14 = Count14 + 1
                ElseIf DateDiff("d", Day21, rst_Trans("TransDate")) <= 0 Then
                    Count21 = Count21 + 1
                Else
                    Count30 = Count30 + 1
                End If
                
                rst_Trans.MoveNext
            Loop
        End If
    Else
        'Reach the current account number
        If Not rst_Trans Is Nothing Then
            If Not rst_Trans.EOF Then
                If AccId = FormatField(rst_Trans("AccID")) Then
                    rst_Trans.Move TotalTrans
                End If
            End If
        End If
    End If
    
    If MaxTrans > 0 Then
        If Count7 > MaxTrans Or Count14 > MaxTrans Or Count21 > MaxTrans Or Count30 > MaxTrans Then
            Balance = 0
        End If
    End If
End If
'*****************************************************************
        'Check for closure
        If Not IsNull(rst_Main("ClosedDate")) Then  'Account has been closed
            ClosedDate = rst_Main("ClosedDate")
            If (Yr = Year(ClosedDate) And Mon >= Month(ClosedDate)) Or Yr > Year(ClosedDate) Then
                Balance = -1
            End If
        End If

        AccIDArr(UBound(AccIDArr)) = AccId
        Products(UBound(Products)) = Balance
        ReDim Preserve AccIDArr(UBound(AccIDArr) + 1)
        ReDim Preserve Products(UBound(Products) + 1)
     
        rst_Main.MoveNext
        
        DoEvents
        If gCancel Then Exit Function
        
    Next I
    Unload L_frmcancel
    Set L_frmcancel = Nothing
End Function

Public Function ComputeSBProducts_Daily(AccIDArr() As Long, ByRef IntAmounts() As Currency, fromDate As Date, toDate As Date, Rate As Double, NoInterestOnMinBal As Boolean) As Currency

Dim I As Long
Dim rst_Main As ADODB.Recordset
Dim rst_PM As ADODB.Recordset
Dim rst_Daily As ADODB.Recordset
Dim TransID As Long
Dim SqlStr As String
Dim MinBalance As Currency
Dim MinBalance_1 As Currency

Dim SetUp As New clsSetup
   MinBalance = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithoutChequeBook", "0.00")
   MinBalance_1 = SetUp.ReadSetupValue("SBAcc", "MinBalanceWithChequeBook", CStr(MinBalance))
   
Set SetUp = Nothing
If MinBalance < 10 Then NoInterestOnMinBal = False

'Validate month and year
    
'First get rec set of all the accounts in SB
    Dim L_frmcancel As New frmCancel
    Set rst_Main = Nothing
    gDbTrans.SqlStmt = "Select A.AccID, ClosedDate from SBMaster A " & _
        " where AccID in (Select Distinct AccID From SBTRANS)" & _
        " AND (ClosedDate is NULL OR CLosedDate > #" & toDate & "#) order BY A.AccID"
    Call gDbTrans.Fetch(rst_Main, adOpenForwardOnly)
    
On Error Resume Next
L_frmcancel.Show
L_frmcancel.Refresh
L_frmcancel.PicStatus.Visible = True
L_frmcancel.Refresh
L_frmcancel.lblMessage.Caption = "Computing Sb Interest"
    
    ''Get teh Transcation ID for each account before the FromDate
    gDbTrans.SqlStmt = "Select B.AccID, MAX(TransID) As MaxTransID " & _
                " FROM SBTrans A, SBMaster B " & _
                " WHERE A.AccID = B.AccID And TransDate < #" & fromDate & "#" & _
                " AND (ClosedDate is NULL OR CLosedDate <= #" & fromDate & "#)" & _
                " GROUP BY B.AccID"
    If Not gDbTrans.CreateView("QrySBPMIDs") Then Exit Function
    
    Set rst_PM = Nothing
    'Get Account Balance  on before day of the Interest
    gDbTrans.SqlStmt = "Select B.AccID, Balance from SBTrans A, QrySBPMIDs B " & _
                    " WHERE A.TransID = B.MaxTransID " & _
                    " AND A.AccID = B.AccID " & _
                    " ORDER BY B.AccID"
    Call gDbTrans.Fetch(rst_PM, adOpenForwardOnly)


    'Get the Last Transaction ID for Each Account On each Day
    gDbTrans.SqlStmt = "Select A.AccID, MAX(TransID) AS MaxTransID, TransDate " & _
                " FROM SBMaster A, SBTrans B " & _
                " WHERE A.AccID = B.AccID " & _
                " and TransDate between #" & fromDate & "# and #" & toDate & "#" & _
                " GROUP BY A.AccID,TransDate"
    If Not gDbTrans.CreateView("QrysbDayBalance") Then Exit Function
    
    'Now Get the Balance for each account on end of each day
    Set rst_Daily = Nothing
    gDbTrans.SqlStmt = "Select B.AccID, Balance,B.TransDate from SBTrans A, QrysbDayBalance B " & _
                    " WHERE A.Transid = B.MaxTransID " & _
                    " AND A.AccID = B.AccID " & _
                    " Order By B.AccID, B.TransDate"
    Call gDbTrans.Fetch(rst_Daily, adOpenForwardOnly)


'**************************************************************
Dim PrevBalance As Currency
Dim Balance As Currency
Dim Interest As Currency
Dim TotalInterest As Currency
Dim DaysForInterest As Integer
Dim PrevDate As Date
Dim CurrentDate As Date
Dim AccId As Long
Dim ClosedDate As String
Dim Balance1TO10 As Currency
Dim Balance11TO31 As Currency
Dim rst As ADODB.Recordset
'***********************************************************
'***********************************************************

'Loop through all the accounts to calculate the IntAmounts
    For I = 1 To rst_Main.RecordCount
         'Debug.Assert rst_Main.Fields("AccID") <> 1400
        'L_frmcancel.prg.Value = rst_Main.AbsolutePosition
        UpdateStatus L_frmcancel.PicStatus, I / rst_Main.RecordCount
        Balance = 0
        PrevBalance = 0
        AccId = FormatField(rst_Main("AccID"))
        If Not rst_PM Is Nothing Then
            If Not rst_PM.EOF Then
                If AccId = FormatField(rst_PM("AccID")) Then
                    PrevBalance = FormatField(rst_PM("Balance"))
                    rst_PM.MoveNext
                End If
            End If
        End If
        'Now Get the Balance for each closing DAy and calculate the interest
        'Filter for each Account
        'rst_Daily.Filter =
        rst_Daily.Filter = adFilterNone
        rst_Daily.Filter = "AccID=" & rst_Main.Fields("AccID")
        'Debug.Assert rst_Main.Fields("AccID") <> 1400
        PrevDate = FinUSFromDate
        Interest = 0
        TotalInterest = 0
        If rst_Daily.RecordCount > 0 Then
            Do While Not rst_Daily.EOF
                CurrentDate = rst_Daily("TransDate")
                DaysForInterest = DateDiff("d", PrevDate, CurrentDate)
                
                Interest = PrevBalance * (DaysForInterest / 365) * (Rate / 100) '(PrevBalance * DaysForInterest * Rate) / (100 * 365)
                
                ''Check whether the balance is more than the Minimum and interest has to be issued on Minimum balance
                If NoInterestOnMinBal And PrevBalance < MinBalance Then Interest = 0
                
                TotalInterest = TotalInterest + Interest
                
                PrevBalance = FormatField(rst_Daily("Balance"))
                PrevDate = CurrentDate
                
                rst_Daily.MoveNext
            Loop
        Else
            CurrentDate = FinUSEndDate
        End If
        CurrentDate = FinUSEndDate
        DaysForInterest = DateDiff("d", PrevDate, CurrentDate)
        Interest = PrevBalance * (DaysForInterest / 365) * (Rate / 100)
        ''Check whether the balance is more than the Minimum and interest has to be issued on Minimum balance
        If NoInterestOnMinBal And PrevBalance < MinBalance Then Interest = 0
        TotalInterest = TotalInterest + Interest
        
        AccIDArr(UBound(AccIDArr)) = AccId
        IntAmounts(UBound(IntAmounts)) = TotalInterest
        ReDim Preserve AccIDArr(UBound(AccIDArr) + 1)
        ReDim Preserve IntAmounts(UBound(IntAmounts) + 1)
     
        rst_Main.MoveNext
        
        DoEvents
        If gCancel Then Exit Function
        
    Next I
    Unload L_frmcancel
    Set L_frmcancel = Nothing
End Function
'
Public Function ComputeTotalSBLiability(AsOnDate As Date) As Currency
Dim rst As ADODB.Recordset
Dim SqlStr As String

ComputeTotalSBLiability = 0

SqlStr = "SELECT AccID, Max(TransID) As MaxTransID " & _
    "FROM SBTrans WHERE TransDate <= #" & AsOnDate & _
    "# GROUP BY AccID"
gDbTrans.SqlStmt = SqlStr
gDbTrans.CreateView ("SBTemp")
gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM SBTrans A, SBTemp B " & _
    " WHERE A.AccID=B.AccID " & _
    " And A.TransID = B.MaxTransID"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    ComputeTotalSBLiability = FormatField(rst(0))

Set rst = Nothing
Exit Function

DoEvents
End Function

