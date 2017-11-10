Attribute VB_Name = "basPDAcc"
Option Explicit
Dim M_setUp As clsSetup

Public Enum wis_PDReports
    repPDBalance = 1
    repPDDayBook = 2
    repPDLedger = 3
    repPDAccOpen = 4
    repPDAccClose = 5
    repPDJoint = 6
    repPDMonTrans = 7
    repPDMat = 8
    repPDLaib = 9
    repPDAgentTrans = 10
    repPDMonBal
    repPDCashBook
    repPDDailyTrans
End Enum

'This Functionm Returns the Last Transaction Date of the
'Pigmy Transaction of the particular account
Private Sub GetLastTransDate(ByVal AccountId As Long, _
                Optional ByRef TransID As Long, Optional ByRef TransDate As Date)

Dim rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Long
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM PDTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM PDIntTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SqlStmt = "Select TransID,TransDate FROM PDIntPayable " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

ErrLine:

End Sub

'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetPigmyLastTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetPigmyLastTransDate = TransDate

End Function

'This Function Returns the Max Transction ID of
'the given FD account Id
'In case there is no transaction it reurns 0
Public Function GetPigmyMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetPigmyMaxTransID = TransID

End Function


Public Function PDInterest(AccId As Long) As Currency
Dim FirstDate As String
Dim LastDate As Date
Dim NextDate As Date
Dim rst As Recordset
Dim TotalAmount As Currency
Dim Product As Currency

TotalAmount = 0
Product = 0

Dim ROI As Single
gDbTrans.SqlStmt = "SELECT * FROM PDMaster WHERE AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
ROI = FormatField(rst.Fields("RateOFInterest"))

gDbTrans.SqlStmt = "SELECT TransDate FROM PDTrans WHERE AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
FirstDate = rst.Fields(0)

gDbTrans.SqlStmt = "SELECT TOP 1 TransDate FROM PDTrans WHERE AccID = " & AccId & _
    " ORDER BY TransDate  DESC"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
LastDate = rst.Fields(0)
NextDate = DateAdd("m", 1, FirstDate)

Do
    If DateDiff("M", FirstDate, LastDate) <= 0 Then Exit Do
    gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount From PdTrans" & _
        " where AccID = " & AccId & _
        " AND Transdate >= #" & FirstDate & "# And TransDate < #" & NextDate & "# " & _
        " AND (TransType = " & wDeposit & " or TransType = " & wContraDeposit & ")"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then GoTo NextMonth
    TotalAmount = TotalAmount + FormatField(rst(0))
    If FormatField(rst(0)) > 0 Then Product = Product + TotalAmount
    
NextMonth:
    FirstDate = NextDate
    NextDate = DateAdd("m", 1, NextDate)
Loop


PDInterest = ((Product * 30 * ROI) / 36500) \ 1

End Function


'
Public Function GetAgentName(AgentID As Long) As String

Dim rst As Recordset
    GetAgentName = ""
    gDbTrans.SqlStmt = "Select CustomerId From UserTab Where UserId = " & AgentID
    Dim CustClass As New clsCustReg
    GetAgentName = " "
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        GetAgentName = CustClass.CustomerName(Val(FormatField(rst(0))))
    End If
    Set CustClass = Nothing
End Function
Public Function GetPDInterestChanged1(fromDate As Date) As Boolean
'This Function Talks With ClsInterest To Dump The Values Into Interest Tab
'It Is Necessary To Get The ModuleID ,SchemeName,FromIndianDate ,To Indian Date

Dim ClsInt As New clsInterest
Dim PDModule  As wisModules
Dim SchemeName As String
Dim InterestRate As Single

'1) Get The ModuleID
PDModule = wis_PDAcc

'2) Get The SchemeName For Each Interest Label
Dim j As Integer
       'Go For Deposits
'For j = frmPDAcc.txtInterestRates.LBound To frmPDAcc.txtInterestRates.UBound
         SchemeName = "Deposit Interest For PD" & CStr(j)
         
         '3) Get The Dates Validated
         
         'If Not DateValidate(FromDate, "/", True) Then GoTo Errline
'         MsgBox "check for the Date Passed Argu.. is in DateFormat"
                     
         'InterestRate = CSng(frmPDAcc.txtInterestRates(j).Text)
         
         '4) Pass The Necessary Values To ClsInt.saveInterest
         If Not ClsInt.SaveInterest(wis_PDAcc, SchemeName, InterestRate, , , fromDate) Then GoTo ErrLine
'Next j

GetPDInterestChanged1 = True

ErrLine:
Set ClsInt = Nothing

End Function

'
Public Function GetPDDepositInterest(Days As Long, AsonIndianDate As String) As Single

'Why We Should we Read From Setup ,When There Is Interest Tab Which Keep Tracks Of Interest Changed
Dim Key As String
Dim IntRate As Double
        
        If Days > 0 And Days <= 30 Then
            Key = "0_1_Deposit"
        ElseIf Days > 30 And Days <= 90 Then
            Key = "1_3_Deposit"
        ElseIf Days > 90 And Days <= 180 Then
            Key = "3_6_Deposit"
        ElseIf Days > 180 And Days < 365 Then
            Key = "6_12_Deposit"
        ElseIf Days > 365 And Days < 730 Then
            Key = "12_24_Deposit"
        ElseIf Days > 730 And Days < 1090 Then
            Key = "24_36_Deposit"
        Else
            Key = "Above36_Deposit"
        End If
        
Dim SetupClass As New clsSetup

Dim strIntRate As String
strIntRate = SetupClass.ReadSetupValue("PDAcc", Key, "")
If Len(strIntRate) > 0 Then IntRate = CDbl(strIntRate)

Set SetupClass = Nothing

        If Days > 0 And Days <= 30 Then
            Key = "Deposit Interest For PD0"
        ElseIf Days > 30 And Days <= 90 Then
            Key = "Deposit Interest For PD1"
        ElseIf Days > 90 And Days <= 180 Then
            Key = "Deposit Interest For PD2"
        ElseIf Days > 180 And Days < 365 Then
            Key = "Deposit Interest For PD3"
        ElseIf Days > 365 And Days < 730 Then
            Key = "Deposit Interest For PD4"
        ElseIf Days > 730 And Days < 1090 Then
            Key = "Deposit Interest For PD5"
        Else
            Key = "Deposit Interest For PD6"
        End If
'Key = Key & "Deposit"

'GetPDDepositInterest = ClsInt.InterestRate(wis_PDAcc, Key, AsOnIndiandate)
GetPDDepositInterest = IntRate
End Function


'
Public Function GetPDLoanInterest(Days As Long, AsOnDate As Date) As Single
Dim IntRate As String
Dim Key As String
        If Days > 0 And Days <= 30 Then
            Key = "0_1_Loan"
        ElseIf Days > 30 And Days <= 90 Then
            Key = "1_3_Loan"
        ElseIf Days > 90 And Days <= 180 Then
            Key = "3_6_Loan"
        ElseIf Days > 180 And Days < 365 Then
            Key = "6_12_Loan"
        ElseIf Days > 365 And Days < 730 Then
            Key = "12_24_Loan"
        ElseIf Days > 730 And Days < 1090 Then
            Key = "24_36_Loan"
        Else
            Key = "Above36_Loan"
        End If
    If M_setUp Is Nothing Then
        Set M_setUp = New clsSetup
    End If

IntRate = M_setUp.ReadSetupValue("PDAcc", Key, "15")
GetPDLoanInterest = IntRate

'Dim Key As String
        If Days > 0 And Days <= 30 Then
            Key = "Loan Interest For PD0"
        ElseIf Days > 30 And Days <= 90 Then
            Key = "Loan Interest For PD1"
        ElseIf Days > 90 And Days <= 180 Then
            Key = "Loan Interest For PD2"
        ElseIf Days > 180 And Days <= 365 Then
            Key = "Loan Interest For PD3"
        ElseIf Days > 365 And Days <= 730 Then
            Key = "Loan Interest For PD4"
        ElseIf Days > 730 And Days <= 1090 Then
            Key = "Loan Interest For PD5"
        Else
            Key = "Loan Interest For PD6"
        End If
'Key = Key & "Loan"
Dim ClsInt As clsInterest
Set ClsInt = New clsInterest

MsgBox "Passed date Argumet is in Dateformat"

'If Val(ClsInt.InterestRate(wis_PDAcc, Key, AsOnDate)) > 0 Then
'    GetPDLoanInterest = ClsInt.InterestRate(wis_PDAcc, Key, AsOnDate)
'Else
'    GetPDLoanInterest = IntRate
'End If
Set ClsInt = Nothing

End Function
'
Public Function ComputePDInterest(Amount As Currency, Rate As Double) As Currency
    ComputePDInterest = (Amount * 1 * Rate) / (100 * 12)
End Function

'
Public Function ComputeTotalPDLiability(AsonIndianDate As String) As Currency
Dim AsOnDate As Date
AsOnDate = GetSysFormatDate(AsonIndianDate)
Dim rst As Recordset
Dim SqlStr As String

SqlStr = "SELECT AccID, Max(TransID) As MaxTransID " & _
    " FROM PDTrans WHERE TransDate <= #" & AsOnDate & "#" & _
    " GROUP BY AccID "

gDbTrans.SqlStmt = SqlStr
'gDBTrans.CreateQueryDef (SqlStr)
gDbTrans.CreateView ("QryTemp")

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM PDTrans A, qryTEMP B " & _
    " WHERE A.AccID=B.AccID And A.TransID = B.MaxTransID "
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then ComputeTotalPDLiability = FormatField(rst(0))
DoEvents

End Function


'Craeted on 1/3/2000
'This Function Will Returns the Pigmy Deposit Balnace at a give date
Public Function GetPDBalance(AsonIndianDate As String) As Currency
    GetPDBalance = ComputeTotalPDLiability(AsonIndianDate)
End Function


Public Function ComputePDInterestAmount(AccId As Long, _
    AsOnDate As Date, Optional ConsiderPremature As Boolean = False) As Currency

Dim transType As wisTransactionTypes
Dim rstTrans As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim MatDate As Date
Dim DepDate As Date
Dim IntRate As Single
Dim IntAmount As Currency

Dim LastTransDate As Date
Dim TransDate As Date
    
gDbTrans.SqlStmt = "Select * from PDMaster where AccID = " & AccId
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
    'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
    Exit Function
End If
IntRate = FormatField(rst("RateOfinterest"))
MatDate = rst("MaturityDate")
DepDate = rst("CreateDate")

gDbTrans.SqlStmt = "Select * from PDTrans where AccID = " & AccId
If gDbTrans.Fetch(rstTrans, adOpenStatic) <= 0 Then
    'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
    Exit Function
End If

'Extract the rate of interest from Setup values
'Dim SetUp As New clsSetup
'If IntRate <= 0 Then _
    IntRate = SetUp.ReadSetupValue("PDAcc", "Interest on PDDeposit", "7")
'Set SetUp = Nothing
If IntRate <= 0 Then _
    IntRate = GetDepositInterestRate(wis_PDAcc, rst("CreateDate"), AsOnDate)

If ConsiderPremature Then _
    If DateDiff("d", AsOnDate, MatDate) < 0 Then IntRate = IntRate - 2

    
'Now check for the valid date
Dim Days As Integer

    'Calculate the number of days
    Days = DateDiff("D", AsOnDate, DepDate)
    If Days > 0 Then  'Account being closed prematurely
        'If deposit is not a year old then do not pay some interest
        If Days < 365 Then GoTo ExitLine
    End If
   
   'Now Calulate the total product
   Dim Product As Currency
   Dim NoOfMonths As Integer
   Dim ContraTrans As wisTransactionTypes
   
'   rstTrans.MoveFirst
   TransDate = rstTrans("TransDate")
   LastTransDate = TransDate
   
    rstTrans.MoveLast
    LastTransDate = GetSysFirstDate(TransDate)
    transType = wDeposit: ContraTrans = wContraDeposit
    
    
    Do
        TransDate = DateAdd("m", 1, CDate(LastTransDate))
        If DateDiff("d", rstTrans("TransDate"), LastTransDate) > 0 Then Exit Do
        gDbTrans.SqlStmt = "Select sum( Amount * Transtype /abs(TransType)) " & _
                    " AS TotalAmount From PDTrans Where AccId = " & AccId & _
                    " AND TransDate >= #" & LastTransDate & "#" & _
                    " And Transdate < #" & TransDate & "#" & _
                    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")"
        
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                    Product = Product + Val(FormatField(rst("TotalAmount")))
        LastTransDate = TransDate
        NoOfMonths = NoOfMonths + 1
        IntAmount = IntAmount + ((Product * CDbl(IntRate / 100)) / 12)
   Loop
   
   'If NoOfMonths > 0 Then IntAmount = Product * CDbl(NoOfMonths / 12) * CDbl(IntRate / 100)

ExitLine:

ComputePDInterestAmount = IntAmount

End Function

