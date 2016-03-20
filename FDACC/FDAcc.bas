Attribute VB_Name = "basFDAcc"
Option Explicit

Public Enum wis_FDReports
    repFDBalance = 1
    repFDDayBook = 2
    repFDLedger = 3
    repFDAccOpen = 4
    repFDAccClose = 5
    repFDJoint = 6
    repFDMonbal = 7
    repFDMat
    repFDLaib
    repFDTrans
    repMFDBalance
    repMFDDayBook
    repMFDLedger
    repMFDMat
    repMFDTrans
    repFDCashBook
    repMFDCashBook

End Enum


'This Function Returns the Max Transction ID of
'the given FD account Id
'In case there is no transaction it reurns 0
Public Function GetFDMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetFDMaxTransID = TransID

End Function

'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetFDMaxTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetFDMaxTransDate = TransDate

End Function

'This Functionm Returns the Last Transaction Date of the
'FD Transaction of the particular account
Private Sub GetLastTransDate(ByVal AccountId As Long, _
                Optional TransID As Long, Optional TransDate As Date)

Dim Rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Long
'Now Assume deposit date as the last int paid amount
gDbTrans.SQLStmt = "Select Top 1 TransID,TransDate FROM FDTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(Rst("TransID")): TransDate = Rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SQLStmt = "Select TransID,TransDate FROM FDIntTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(Rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = Rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SQLStmt = "Select TransID,TransDate FROM FDIntPayable " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(Rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = Rst("TransDate")
End If

ErrLine:

End Sub

Public Function GetDepositInterestRateByDays(DepositType As Integer, _
            ByVal DiffDays As Long, Optional IsEmployee As Boolean, _
            Optional isSenior As Boolean) As Single

On Error GoTo ErrLine
'Dim DiffDays As Integer

If DiffDays < 0 Then DiffDays = DiffDays * -1

Dim SetUp As New clsSetup
Dim strRet As String
Dim IntRate() As String
Dim strKey As String

'If DepositType = wis_PDAcc Then GoTo PIGMY
'If DepositType = wis_RDAcc Then GoTo RECURRING
If DepositType < wis_Deposits Then DepositType = wis_Deposits + DepositType

If DiffDays > 365 Then
    Dim DiffYear As Integer
    DiffYear = DiffDays \ 365 'DateDiff("YYYY", FromDate, ToDate)
    'If DateDiff("D", DateAdd("yyyy", DiffYear, FromDate), _
                    ToDate) < 0 Then DiffYear = DiffYear - 1
    
    If DiffYear > 10 Then DiffYear = 9
    strKey = "YEAR" & DiffYear & "-" & (DiffYear + 1)
Else
    strKey = "DAYS0-15"
    If DiffDays > 15 Then strKey = "DAYS15-30"
    If DiffDays > 30 Then strKey = "DAYS30-45"
    If DiffDays > 45 Then strKey = "DAYS45-60"
    If DiffDays > 60 Then strKey = "DAYS60-90"
    If DiffDays > 90 Then strKey = "DAYS90-120"
    If DiffDays > 120 Then strKey = "DAYS120-180"
    If DiffDays > 180 Then strKey = "DAYS180-365"
End If

DepositType = DepositType Mod 100
strRet = SetUp.ReadSetupValue("DEPOSIT" & DepositType, strKey, "")
If Len(Trim$(strRet)) = 0 Then Exit Function

IntRate = Split(strRet, ",")

If IsEmployee Then IntRate(0) = IntRate(1)
If isSenior Then IntRate(0) = IntRate(2)

If IsEmployee Then
    GetDepositInterestRateByDays = IntRate(1)
ElseIf isSenior Then
    GetDepositInterestRateByDays = IntRate(2)
Else
    GetDepositInterestRateByDays = IntRate(0)
End If
Exit Function

PIGMY:
'GetDepositInterestRate = GetPDDepositInterest(DateDiff("D", FromDate, ToDate), _
                          GetAppFormatDate(ToDate))

Exit Function

RECURRING:
'GetDepositInterestRate = SetUp.ReadSetupValue("RDAcc", "Interest on RDDeposit", "7")

Exit Function

ErrLine:

    Set SetUp = Nothing
End Function


Public Function GetDepositInterestRate(DepositType As Integer, _
        ByVal FromDate As Date, ByVal ToDate As Date, _
        Optional IsEmployee As Boolean, Optional isSenior As Boolean) As Single
        
On Error GoTo ErrLine
Dim DiffDays As Integer

DiffDays = DateDiff("D", FromDate, ToDate)
If DiffDays < 0 Then DiffDays = DiffDays * -1

Dim SetUp As New clsSetup
Dim strRet As String
Dim IntRate() As String
Dim strKey As String
'If DepositType = wis_PDAcc Then GoTo PIGMY
'If DepositType = wis_RDAcc Then GoTo RECURRING
If DepositType < wis_Deposits Then DepositType = wis_Deposits + DepositType

If DiffDays > 365 Then
    Dim DiffYear As Integer
    DiffYear = DateDiff("YYYY", FromDate, ToDate)
    If DateDiff("D", DateAdd("yyyy", DiffYear, FromDate), _
                    ToDate) < 0 Then DiffYear = DiffYear - 1
    
    If DiffYear > 10 Then DiffYear = 9
    strKey = "YEAR" & DiffYear & "-" & (DiffYear + 1)
Else
    strKey = "DAYS0-15"
    If DiffDays > 15 Then strKey = "DAYS15-30"
    If DiffDays > 30 Then strKey = "DAYS30-45"
    If DiffDays > 45 Then strKey = "DAYS45-60"
    If DiffDays > 60 Then strKey = "DAYS60-90"
    If DiffDays > 90 Then strKey = "DAYS90-120"
    If DiffDays > 120 Then strKey = "DAYS120-180"
    If DiffDays > 180 Then strKey = "DAYS180-365"
End If

DepositType = DepositType Mod 100
strRet = SetUp.ReadSetupValue("DEPOSIT" & DepositType, strKey, "")
If Len(Trim$(strRet)) = 0 Then Exit Function

IntRate = Split(strRet, ",")

If IsEmployee Then IntRate(0) = IntRate(1)
If isSenior Then IntRate(0) = IntRate(2)

If IsEmployee Then
    GetDepositInterestRate = IntRate(1)
ElseIf isSenior Then
    GetDepositInterestRate = IntRate(2)
Else
    GetDepositInterestRate = IntRate(0)
End If
Exit Function

PIGMY:
GetDepositInterestRate = GetPDDepositInterest(DateDiff("D", FromDate, ToDate), _
                           GetIndianDate(ToDate))

Exit Function
RECURRING:

GetDepositInterestRate = SetUp.ReadSetupValue("RDAcc", "Interest on RDDeposit", "7")

Exit Function
ErrLine:

    Set SetUp = Nothing
End Function

Public Function ComputeFDInterest(Principle As Currency, _
    FromDate As Date, ToDate As Date, DepositType As Integer, _
    Optional InterestRate As Single) As Currency
Dim IntAmount As Currency
Dim IntRate As Single
Dim IntDiff As Single
Dim Days As Integer

Days = DateDiff("d", FromDate, ToDate)

If InterestRate = 0 Then InterestRate = GetDepositInterestRate(DepositType, FromDate, ToDate)
If Days > 0 Then _
    IntAmount = CDbl(Principle) * CDbl(Days / 365) * CDbl(InterestRate / 100)

ComputeFDInterest = (IntAmount \ 1)

End Function

Public Function GetFDDepositInterest1(Days As Integer, _
    AsOnDate As Date, DepositType As Integer) As Single

Dim SchemeName As String

If Days < 45 Then
   SchemeName = "0_1.5_"
ElseIf Days >= 45 And Days < 91 Then
   SchemeName = "1.5_3_"
ElseIf Days >= 91 And Days < 181 Then
   SchemeName = "3_6_"
ElseIf Days >= 181 And Days < 366 Then
   SchemeName = "6_12_"
ElseIf Days >= 366 And Days < (366 * 2) - 1 Then
   SchemeName = "12_24_"
ElseIf Days >= (366 * 2) - 1 And Days < 365 * 3 Then
   SchemeName = "24_36_"
ElseIf Days >= 365 * 3 And Days < 365 * 4 Then
   SchemeName = "36_48_"
ElseIf Days >= 365 * 4 And Days < 365 * 5 Then
   SchemeName = "48_60_"
ElseIf Days >= 365 * 5 Then
   SchemeName = "Above60_"
End If

SchemeName = SchemeName & "Deposit"

Dim ClsInt As New clsInterest
GetFDDepositInterest1 = ClsInt.InterestRate(wis_Deposits + DepositType, SchemeName, AsOnDate)
Set ClsInt = Nothing

End Function

Public Function GetFDLoanInterest(Days As Integer, AsOnDate As Date) As Single

Dim SchemeName As String

If Days < 15 Then
   SchemeName = "0.5_1_"
ElseIf Days >= 15 And Days < 31 Then
   SchemeName = "0.5_1_"
ElseIf Days >= 31 And Days < 45 Then
   SchemeName = "1_1.5_"
ElseIf Days >= 45 And Days < 90 Then
   SchemeName = "1.5_3_"
ElseIf Days >= 90 And Days < 180 Then
   SchemeName = "3_6_"
ElseIf Days >= 180 And Days < 365 Then
   SchemeName = "6_12_"
ElseIf Days >= 365 And Days < 730 Then
   SchemeName = "12_24_"
ElseIf Days >= 730 And Days < 365 * 3 Then
   SchemeName = "24_36_"
Else 'Days >= 365 * 3 Then
   SchemeName = "36_Above_"
End If
SchemeName = SchemeName & "Loan"
Dim ClsInt As New clsInterest
GetFDLoanInterest = ClsInt.InterestRate(wis_Deposits, SchemeName, AsOnDate)
Set ClsInt = Nothing
   
End Function


