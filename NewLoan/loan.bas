Attribute VB_Name = "basLoan"
Option Explicit

Enum wisSeason
    wisNoSeason = 0
    wisKhariff = 1
    wisRabi = 2
    wisT_Belt = 3
    wisAnnual = 4
    wisOtherSeason = 5
End Enum

Enum wisFarmerClassification
    SmallFarmer = 1
    BigFarmer = 2
    MarginalFarmer = 3
    OtherFarmer = 4
    NoFarmer = 0
End Enum

Public Enum wisInstallmentTypes
   Inst_No = 0
   Inst_Daily = 1
   Inst_Weekly = 2
   Inst_FortNightly = 3
   Inst_Monthly = 4
   Inst_BiMonthly = 5
   Inst_Quartery = 6
   Inst_HalfYearly = 7
   Inst_Yearly = 8
End Enum

Enum wis_LoanType
    wisCashCreditLoan = 1
    wisVehicleloan = 2
    wisCropLoan = 4
    wisIndividualLoan = 8
    wisBKCC = 16
    wisSHGLoan = 32
End Enum

Public Enum wis_LoanReports
    '''Regular Reportr
    repMonthlyRegister = 1
    repMonthlyRegisterAll = 2
    repShedule_1
    repShedule_2
    repShedule_3
    repShedule_4A
    repShedule_4B
    repShedule_4C
    repShedule_5
    repShedule_6
  ''Reports
    repLoanBalance = 21
    repLoanHolder
    repLoanCashBook
    repLoanIssued
    repLoanInstOD
    repLoanIntCol
    repLoanIntReceivable
    repLoanIntReceivableTill
    repLoanDailyCash
    repLoanGLedger
    repLoanRepMade
    repLoanOD
    repLoanSanction
    repLoanGuarantor
    repLoanCustRP
    repLoanReceivable
    
    repConsBalance = 40
    repConsInstOD = 41
    repConsOD = 42
   
   
End Enum


'This Function Returns the Last Transaction Date and Transaction ID
'of the Loan Transaction of the particular account
Private Sub GetLastTransDate(ByVal LoanID As Integer, _
                Optional TransID As Long, Optional TransDate As Date)

Dim rst As Recordset
TransID = 0
TransDate = vbNull

On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Integer
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM LoanTrans " & _
                    " where LoanId = " & LoanID & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM LoanIntTrans " & _
                    " where LoanId = " & LoanID & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SqlStmt = "Select TransID,TransDate FROM LoanIntReceivable " & _
                    " where LoanId = " & LoanID & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Check the same for amountreceivable also
Dim AccHeadID As Long

gDbTrans.SqlStmt = "Select SchemeName FROM LoanScheme " & _
                    " Where SchemeID = (Select schemeID " & _
                        " From LoanMaster where LoanID = " & LoanID & ")"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

AccHeadID = GetHeadID(FormatField(rst("SchemeName")), parMemberLoan)

gDbTrans.SqlStmt = "Select * FROM AmountReceivable " & _
                    " where AccHeadID = " & AccHeadID & _
                    " And AccId = " & LoanID & _
                    " Order By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("AccTransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
    If TransDate < rst("TransDate") Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

ErrLine:

End Sub

'This Function Returns the Last Transction Date of The Fd
'of the given loan account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetLoanLastTransDate(ByVal LoanID As Integer) As Date
Dim TransDate As Date
Call GetLastTransDate(LoanID, , TransDate)
GetLoanLastTransDate = TransDate

End Function

Public Function HasOverdueLoans(customerID As Long) As Boolean
'HasOverdueLoans=
End Function


'This Function Returns the Max Transction ID of
'the given LOan account Id
'In case there is no transaction it reurns 0
Public Function GetLoanMaxTransID(ByVal LoanID As Integer) As Long
Dim TransID As Long
Call GetLastTransDate(LoanID, TransID)
GetLoanMaxTransID = TransID

End Function


'This Functionm is used to Save the LOan Purpose
'If Scheme id given it stores this for particular loan scheme
Public Sub SaveLoanPurpose(cmbPurpose As ComboBox, Optional SchemeID As Integer = 0)

Dim SqlStr As String
Dim PurposeID As Long
Dim strPurpose As String
Dim rst As Recordset

If cmbPurpose.ListIndex = -1 Then strPurpose = cmbPurpose.Text

If strPurpose = "" Then Exit Sub

'Check for the existing Loanpurpose
SqlStr = "SELECT * FROM LoanPurpose WHERE Purpose = " & AddQuotes(strPurpose)

If SchemeID <> 0 Then SqlStr = SqlStr & " AND SchemeID = " & SchemeID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Exit Sub

'Now get the max(PurposeID) from Purpose table
SqlStr = "SELECT MAX(PurposeID) from LoanPurpose "
gDbTrans.SqlStmt = SqlStr

PurposeID = 1
If gDbTrans.Fetch(rst, adOpenDynamic) = 1 Then
   PurposeID = FormatField(rst(0)) + 1
Else
   Exit Sub
End If
strPurpose = Trim$(cmbPurpose.Text)
'Now Insert this into Database
If SchemeID = 0 Then
    SqlStr = "INSERT INTO LoanPurpose (PurposeId,Purpose) Values ( " & _
      PurposeID & "," & _
      AddQuotes(strPurpose, True) & " ) "
Else
    SqlStr = "INSERT INTO LoanPurpose" & _
            " (PurposeId,Purpose,SchemeID)" & _
            " Values ( " & _
            PurposeID & "," & _
            AddQuotes(strPurpose, True) & ", " & _
            SchemeID & " ) "
End If
gDbTrans.SqlStmt = SqlStr
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
   gDbTrans.RollBack
   Exit Sub
End If

gDbTrans.CommitTrans

End Sub
'This function will load the loan Purposes to a given combo box
' into the cmbobox
Public Sub LoadLoanPurposes(cmbBox As ComboBox, Optional SchemeID As Integer)
Dim SqlStr As String
Dim rst As Recordset

SqlStr = "SELECT * FROM LoanPurpose WHERE SchemeID = 0 or SchemeID is Null"
If SchemeID Then SqlStr = SqlStr & " or SchemeID = " & SchemeID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

With cmbBox
    .Clear
    While Not rst.EOF
        .AddItem FormatField(rst("Purpose"))
        .ItemData(.newIndex) = FormatField(rst("PurposeID"))
        rst.MoveNext
    Wend
End With

End Sub



'This function will load the loan schemes
' into the cmbobox
Public Sub LoadLoanSchemes(cmbBox As ComboBox)
Dim SqlStr As String
Dim rst As Recordset
'Dim obj As Object
'Set obj = New clsTransact


SqlStr = "SELECT SchemeID,SchemeName from LoanScheme Order by SchemeName "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'Set Rst = gDBTrans.Rst.Clone
Else
    Exit Sub
End If
cmbBox.Clear

While Not rst.EOF
    cmbBox.AddItem FormatField(rst("SchemeName"))
    cmbBox.ItemData(cmbBox.newIndex) = FormatField(rst("SchemeID"))
    rst.MoveNext
Wend

End Sub



