Attribute VB_Name = "basRDAcc"
Option Explicit

Public Enum wis_RDReports
    repRDBalance = 1
    repRDDayBook = 2
    repRDLedger = 3
    repRDAccOpen = 4
    repRDAccClose = 5
    repRDJoint = 6
    repRDMonbal = 7
    repRDMat = 8
    repRDLaib = 9
    repRDCashBook
End Enum

'This Functionm Returns the Last Transaction Date of the
'Pigmy Transaction of the particular account
Private Sub GetLastTransDate(ByVal AccountId As Long, _
                Optional ByRef TransID As Long, Optional ByRef TransDate As Date)

Dim Rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Long
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM RDTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(Rst("TransID")): TransDate = Rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM RDIntTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(Rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = Rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SqlStmt = "Select TransID,TransDate FROM RDIntPayable " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(Rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = Rst("TransDate")
End If

ErrLine:

End Sub

'This Function Returns the Max Transction ID of
'the given RD account Id
'In case there is no transaction it reurns 0
Public Function GetRDMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetRDMaxTransID = TransID

End Function



'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns vb deafault date
Public Function GetRDLastTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetRDLastTransDate = TransDate

End Function


Public Function ComputeRDDepositInterestAmount(AccId As Long, _
    AsOnDate As Date, Optional ConsiderPremature As Boolean = False) As Currency

Dim transType As wisTransactionTypes
Dim rstTrans As ADODB.Recordset
Dim Rst As ADODB.Recordset
Dim MatDate As Date
Dim IntRate As Single
Dim IntAmount As Currency

Dim LastTransDate As Date
Dim TransDate As Date
    
    
    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & AccId
    If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    IntRate = FormatField(rstTrans("RateOfinterest"))
    MatDate = rstTrans("MaturityDate")

    gDbTrans.SqlStmt = "Select * from RDTrans where AccID = " & AccId
    If gDbTrans.Fetch(rstTrans, adOpenStatic) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Function
    End If
    
'Extract the rate of interest from Setup values
    Dim SetUp As New clsSetup
    If IntRate <= 0 Then _
        IntRate = SetUp.ReadSetupValue("RDAcc", "Interest on RDDeposit", "7")
If ConsiderPremature Then _
    If DateDiff("d", AsOnDate, MatDate) < 0 Then IntRate = IntRate - 2

Set SetUp = Nothing
    
'Now check for the valid date
Dim Days As Integer

    'Calculate the number of days
    Days = DateDiff("D", AsOnDate, MatDate)
    If Days > 0 Then  'Account being closed prematurely
        'If deposit is not a year old then do not pay some interest
        'If Days < 365 Then GoTo ExitLine
   End If
   
   'Now Calulate the total product
   Dim Product As Currency
   Dim NoOfMonths As Integer
   Dim ContraTrans As wisTransactionTypes
   
'   rstTrans.MoveFirst
   TransDate = rstTrans("TransDate")
   LastTransDate = TransDate
   
    rstTrans.MoveLast
    LastTransDate = GetSysFirstDate(LastTransDate)
    transType = wDeposit: ContraTrans = wContraDeposit
    
    Do
        TransDate = DateAdd("m", 1, CDate(LastTransDate))
        If DateDiff("d", rstTrans("TransDate"), LastTransDate) > 0 Then Exit Do
        gDbTrans.SqlStmt = "Select sum( Amount * Transtype /abs(TransType)) " & _
                    " AS TotalAmount From RDTrans Where AccId = " & AccId & _
                    " AND TransDate >= #" & LastTransDate & "#" & _
                    " And Transdate < #" & TransDate & "#" & _
                    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")"
        If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then _
                    Product = Product + Val(FormatField(Rst("TotalAmount")))
        
        LastTransDate = TransDate
        NoOfMonths = NoOfMonths + 1
   Loop
   
   If NoOfMonths > 0 Then IntAmount = Product * CDbl(NoOfMonths / 12) * CDbl(IntRate / 100)

ExitLine:

ComputeRDDepositInterestAmount = IntAmount

End Function

Public Function ComputeRDInterest(Amount As Currency, Rate As Double) As Currency
    ComputeRDInterest = (Amount * 1 * Rate) / (100 * 12)
End Function
 
 
Public Function ComputeTotalRDLiability(AsOnDate As Date) As Currency

Dim SqlStmt As String
Dim Rst As ADODB.Recordset
Dim TotalBalance As Currency
    
ComputeTotalRDLiability = 0

''''Changerd By Shashi as on 23/11/20001
gDbTrans.SqlStmt = "Select SUM(Balance) From RDTrans A, RDMaster B Where " & _
   " (ClosedDate > #" & AsOnDate & "# OR ClosedDate is NULL)" & _
   " AND B.AccID = A.AccID And TransId = (Select Max(TransID) " & _
   " From RDTrans C Where C.AccId = A.AccID " & _
   " And TransDate <= #" & AsOnDate & "#)"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Function
ComputeTotalRDLiability = FormatField(Rst(0))
Set Rst = Nothing

Exit Function

DoEvents
End Function

'
''Author Shashi
'Craeted on 1/3/2000
'This Function Will Returns the Pigmy Deposit Balnace at a give date
Public Function GetRDBalance(AsOnDate As Date) As Currency
    GetRDBalance = ComputeTotalRDLiability(AsOnDate)
End Function

