VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private NewDbClass As clsTransact
'Private NewDBClass As Object
Private ConnectionString As String
Private DbFile As String
Private StrPwd As String

Private WithEvents m_frmFinYear As frmFinYear
Attribute m_frmFinYear.VB_VarHelpID = -1


'THis FUnction Transafer the Customer From Old TO NEw Customer Table
'
Private Function ExportCustomer() As Boolean
' Handle error
On Error GoTo Hell:

' Declarations
Dim RstCust As ADODB.Recordset

'ExportProducts = False

'Fetch Customer Data
'Make some Change to the data base Temporary To aake tha Data Transafer faset
With gDbTrans
    '.SQLStmt = "UPDATE NAMETAB SET "

    .SqlStmt = "SELECT * FROM NAMETAB"

    Call .Fetch(RstCust, adOpenForwardOnly)
End With

' Start Loop
Do While Not RstCust.EOF
    
    NewDbClass.SqlStmt = "INSERT INTO NAMETAB(CustomerID,Title,FirstName,MiddleName," & _
                        "LastName,Gender,Profession,Caste,Place,DOB,MaritalStatus," & _
                        "HomeAddress,OfficeAddress,HomePhone,OfficePhone,eMail," & _
                        "Reference,IsciName,UserId,CustType)" & _
                         "VALUES(" & _
                        RstCust("CustomerID") & "," & AddQuotes(RstCust("Title")) & "," & _
                        AddQuotes(RstCust("FirstName")) & "," & AddQuotes(RstCust("MiddleName")) & "," & _
                        AddQuotes(RstCust("LastName")) & "," & RstCust("Gender") & "," & _
                        AddQuotes(RstCust("Profession")) & "," & AddQuotes(RstCust("Caste")) & _
                        AddQuotes(RstCust("Place")) & ",#" & RstCust("DOB") & "#," & _
                        RstCust("MaritalStatus") & "," & AddQuotes(RstCust("HomeAddress")) & "," & _
                        AddQuotes(RstCust("OfficeAddress")) & "," & AddQuotes(RstCust("HomePhone")) & "," & _
                        AddQuotes(RstCust("OfficePhone")) & "," & AddQuotes(RstCust("eMail")) & "," & _
                        RstCust("Reference") & "," & AddQuotes(RstCust("IsciName")) & "," & _
                        RstCust("UserId") & "," & RstCust("CustType") & " )"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    RstCust.MoveNext
Loop

ExportCustomer = True

' Exit

Exit Function

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Function


'This Function checks the Sb accounts with the last 10 transction
Private Sub CheckSBCAAccounts()
' Handle error
On Error GoTo Hell:

'ExportProducts = False
'Fetch Customer Data
'Make some Change to the data base Temporary To aake tha Data Transafer faset
Dim RstCust As Recordset
With NewDbClass
    'Delete the Account which are closed before the new Financial Year
    .SqlStmt = "Delete * from SBTRANS " & _
            " where ACCID IN (SELECT ACCID FROM SBMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianFromDate) & "#)"
    Call .SQLExecute
    'Delete PLTrans
    .SqlStmt = "Delete * from SBPLTrans" & _
            " where ACCID IN (SELECT ACCID FROM SBMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianFromDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * from SBJOINT" & _
            " where ACCID IN (SELECT ACCID FROM SBMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    
    
    .SqlStmt = "Delete * From SbPLTrans where AccID IN " & _
            " ( Select Distinct AccID FROM SBMASTER Where " & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "# )"
    Call .SQLExecute
    .SqlStmt = "Delete * From SbTrans where AccID IN " & _
            "(Select Distinct AccID FROM SBMASTER Where " & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "# )"
    Call .SQLExecute
    
    .SqlStmt = "Delete * FROM SBMASTER Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
    
    'Now Delete the CA Accounta which are closed before the new Financial Year
    .SqlStmt = "Delete * from CATRANS " & _
            " where ACCID IN (SELECT ACCID FROM CAMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete PLTrans
    .SqlStmt = "Delete * from CAPLTrans" & _
            " where ACCID IN (SELECT ACCID FROM CAMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianFromDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * from CAJOINT" & _
            " where ACCID IN (SELECT ACCID FROM CAMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianFromDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * FROM CAMASTER Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
    
    'Delete the Previous Last Years transaction
    'Now Delete the Master
    .SqlStmt = "Delete A.* FROM SBTrans A WHERE " & _
                " TransID <= ((SELECT MAX(TranSID) FROM SBTrans B " & _
                    " Where A.AccID = B.AccID AND B.TransDate < #" & FinIndianFromDate & "# ) -10)"
    'Call .SQLExecute
    .SqlStmt = "Delete A.* FROM CATrans A WHERE " & _
                " TransID <= ((SELECT MAX(TranSID) FROM CATrans B " & _
                    " Where A.AccID = B.AccID) -10)"
    'Call .SQLExecute
    
End With

' Exit

Exit Sub

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Sub



'This Function checks the RD & Pigmy accounts with the last 10 transction
Private Sub CheckRDPDAccounts()
' Handle error
On Error GoTo Hell:

'ExportProducts = False
'Fetch Customer Data
'Make some Change to the data base Temporary To aake tha Data Transafer faset
Dim RstCust As Recordset
With NewDbClass
    'Delete the Account which are closed befor the new Financial Year
    .SqlStmt = "Delete * from RDTRANS " & _
            " where ACCID IN (SELECT ACCID FROM RDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete Int Trans
    .SqlStmt = "Delete * from RDIntTrans" & _
            " where ACCID IN (SELECT ACCID FROM RDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete Int Trans
    .SqlStmt = "Delete * from RDIntPayable" & _
            " where ACCID IN (SELECT ACCID FROM RDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * from RDJOINT" & _
            " where ACCNUM IN (SELECT DISTINCT(AccNum) FROM RDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    
    .SqlStmt = "Delete * FROM RDMASTER Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
    
    'Now Delete the Pigmy Accounts which are closed before the new Financial Year
    .SqlStmt = "Delete * from PDTRANS " & _
            " where ACCID IN (SELECT ACCID FROM PDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete Paybale Trans
    .SqlStmt = "Delete * from PDIntPayable" & _
            " where ACCID IN (SELECT ACCID FROM PDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete PLTrans
    .SqlStmt = "Delete * from PDIntTrans" & _
            " where ACCID IN (SELECT ACCID FROM PDMASTER " & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * FROM PDMASTER Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
End With

Exit Sub

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Sub



'This Function checks the Loan accounts with the last 10 transction
Private Sub CheckLoanAccounts()
' Handle error
On Error GoTo Hell:

'ExportProducts = False
'Fetch Customer Data
'Make some Change to the data base Temporary To take tha Data Transafer faset
Dim RstCust As Recordset
With NewDbClass
    'Delete the Account which are closed befor the new Financial Year
    .SqlStmt = "Delete * from LoanABNEP " & _
            " where LoanID IN (SELECT LoanID FROM LoanMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Delete Int Trans
    .SqlStmt = "Delete * from LoanIntTrans" & _
            " where LoanID IN (SELECT LoanID FROM LoanMASTER " & _
                " Where ClosedDate < #" & FinUSEndDate & "#)"
    Call .SQLExecute
    'Delete Int Trans
    .SqlStmt = "Delete * from LoanIntReceivAble" & _
            " where LoanID IN (SELECT LoanID FROM RDMASTER " & _
                " Where ClosedDate < #" & FinUSEndDate & "#)"
    Call .SQLExecute
    'Delete Inst Trans
    .SqlStmt = "Delete * from LoanInstTrans" & _
            " where LoanID IN (SELECT LoanID FROM LoanMASTER " & _
                " Where ClosedDate < #" & FinUSEndDate & "#)"
    Call .SQLExecute
    'Delete Int Trans
    .SqlStmt = "Delete * from LoanInst" & _
            " where LoanID IN (SELECT LoanID FROM LoanMASTER " & _
                " Where ClosedDate < #" & FinUSEndDate & "#)"
    Call .SQLExecute
    
    .SqlStmt = "Delete * from LoanInst where LoanID NOT IN (SELECT DISTINCT(LoanID) FROM LoanTrans)"
    Call .SQLExecute
    'now Delete the MAster detial
    .SqlStmt = "Delete * from LoanMaster" & _
            " where LoanID NOT IN (SELECT DISTINCT(LoanID) FROM LoanTrans)"
    Call .SQLExecute
    'Delete Loan Definitions
    .SqlStmt = "Delete * from LoanScheme" & _
            " where SchemeID NOT IN (SELECT DISTINCT(SchemeID) FROM LoanMaster)"
    Call .SQLExecute
    
    'Now Delete the Pigmy Accounts which are closed before the new Financial Year
    .SqlStmt = "Delete * from DepositLoanTRANS " & _
            " where LoanID IN (SELECT LoanID FROM DepositLoanMASTER " & _
                " Where LoanClosedDate < #" & FinUSFromDate & "#)"
    Call .SQLExecute
    'Delete Paybale Trans
    .SqlStmt = "Delete * from DepositLoanIntTrans" & _
            " where LoanID IN (SELECT LoanID FROM DepositLOanMASTER " & _
                " Where LoanClosedDate < #" & FinUSFromDate & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * FROM DepositLoanMASTER Where" & _
                " LoanClosedDate < #" & FinUSFromDate & "#"
    Call .SQLExecute
End With

Exit Sub

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Sub


'This Function checks the Sb accounts with the last 10 transction
Private Sub CheckMemberAccounts()

'Handle error
On Error GoTo Hell:

'ExportProducts = False
'Fetch Customer Data
'Make some Change to the data base Temporary To aake tha Data Transafer faset
Dim RstCust As Recordset
With NewDbClass
    'Delete the Account which are closed befor the new Financial Year
    'First delete the Interest
    .SqlStmt = "Delete * from MemIntTrans" & _
            " where ACCID IN (SELECT ACCID FROM MemMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Trans
    .SqlStmt = "Delete * from MemIntPayable" & _
            " where ACCID IN (SELECT ACCID FROM MemMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    
    'Now Delete the Trans
    .SqlStmt = "Delete * from MemTrans" & _
            " where ACCID IN (SELECT ACCID FROM MemMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * FROM MemMaster Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
End With

' Exit

Exit Sub

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Sub


'This Function checks the Sb accounts with the last 10 transction
Private Sub CheckDepositAccounts()

'Handle error
On Error GoTo Hell:

'ExportProducts = False
'Fetch Customer Data
'Make some Change to the data base Temporary To aake tha Data Transafer faset
Dim RstCust As Recordset
With NewDbClass
    'Delete the Account which are closed befor the new Financial Year
    'First delete the Interest
    .SqlStmt = "Delete * from FDIntTrans" & _
            " where ACCID IN (Select ACCID FROM FDMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Trans
    .SqlStmt = "Delete * from FDIntPayable" & _
            " where ACCID IN (SELECT ACCID FROM FDMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    
    'Now Delete the Trans
    .SqlStmt = "Delete * from FDTrans" & _
            " where ACCID IN (SELECT ACCID FROM FDMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Joint
    .SqlStmt = "Delete * from FDJOINT" & _
            " where ACCID IN (SELECT Distinct(ACCID) FROM FDMaster" & _
                " Where ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#)"
    Call .SQLExecute
    'Now Delete the Master
    .SqlStmt = "Delete * FROM FDMaster Where" & _
                " ClosedDate < #" & GetSysFormatDate(FinIndianEndDate) & "#"
    Call .SQLExecute
End With


' Exit

Exit Sub

Hell:
    MsgBox "Exportcustomer: " & Err.Description

End Sub

Public Property Get ObjectName()
    ObjectName = "FinYearChangeClass"
End Property


'After istalling the software
'first time database creation
Private Sub CreateDBFirstTime()

Dim UtilClass As clsDBUtilities

Call ShowFinYear

'If m_frmFinYear Is Nothing Then Exit Sub

Call SaveFinYear(App.Path & "\" & constFINYEARFILE, True)

Set UtilClass = New clsDBUtilities

'copy the default shipped database to the respective directory

Call UtilClass.CopyDefaultDatabase(App.Path & "\BlankDataBase\" & constDBName, DbFile)

Set UtilClass = Nothing

End Sub

'private withevents
'This will create the new mdb for the next finacial year

Private Function CreateNewDb() As Boolean

On Error GoTo Hell

Dim UtilClass As clsDBUtilities

CreateNewDb = False

'Set UtilClass = New clsDBUtilities
'If Not UtilClass.CreateDBFromDB(DbFile, StrPwd) Then Exit Function
'Set UtilClass = nothing

'Instead of CreatingNew database copy the existing data base
'the new fin year folder
'get the existing DB Name
Dim strSource As String
strSource = gDbTrans.DataBaseName
'Close the source Data BAse
gDbTrans.CloseDB

'Now Copy the data BAse
FileCopy strSource, DbFile

'Now Open the Data Source Database
If Not gDbTrans.OpenDB(strSource, StrPwd) Then
    MsgBox "Unable to open the date base"
    End
End If

Set UtilClass = Nothing

CreateNewDb = True

Exit Function

Hell:
    If Err.Number > 0 Then MsgBox "CreateNewDB: " & Err.Description
    Set UtilClass = Nothing
    
End Function

Private Function ExportFunctions() As Boolean
    ' Export Heads excluding revenue heads and Opening Balances
    If Not ExportHeadBalances Then Exit Function
    
    
    If Not TransferStock Then Exit Function
    'Export Product Group
    'If Not ExportProductGroup Then Exit Function
        
    'Export Products
    'If Not ExportProducts Then Exit Function
        
    'Export Products
'    If Not ExportTransPortMode Then Exit Function
    
    ' Export Stock & Relations
'    If Not ExportStocknRelations Then Exit Function
    
    Call CheckMemberAccounts
    Call CheckSBCAAccounts
    Call CheckRDPDAccounts
    Call CheckDepositAccounts
    Call CheckLoanAccounts
    Call UpdateHeadTrans
        
    ExportFunctions = True

End Function

'This will Export the Relations and Stock to New mdb
Private Function ExportStocknRelations() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstRelation As ADODB.Recordset
Dim rstNewRelation As ADODB.Recordset

' Fetch Data
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM RelationMaster " & _
                   " ORDER By RelationID "

Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)


' Start Loop
Do While Not rstRelation.EOF
    
    NewDbClass.SqlStmt = "INSERT INTO RelationMaster (RelationID,GodownID," & _
                         "GroupID,ProductID,UnitID,TradingPrice,MRP,SalesPrice,PriceChanged)" & _
                         " VALUES ( " & _
                         rstRelation.Fields("RelationID") & "," & _
                         rstRelation.Fields("GodownID") & "," & _
                         rstRelation.Fields("GroupID") & "," & _
                         rstRelation.Fields("ProductID") & "," & _
                         rstRelation.Fields("UnitID") & "," & rstRelation.Fields("TradingPrice") & "," & rstRelation.Fields("MRP") & "," & rstRelation.Fields("SalesPrice") & "," & rstRelation.Fields("PriceChanged") & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    If Not ExportStock(rstRelation.Fields("RelationID"), rstRelation.Fields("GodownID")) Then Exit Function
    
    rstRelation.MoveNext
    
Loop

ExportStocknRelations = True
' Exit

Exit Function

Hell:
    MsgBox "ExportStocknRelations: " & Err.Description
    
End Function


Private Function CreateStockMergeQuery() As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

CreateStockMergeQuery = False

USFromDate = FinUSFromDate 'FormatDate(FinFromDate)
USAsOnDate = FinUSEndDate  'FormatDate(FinEndDate)    'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,B.RelationID" & _
                   " FROM Stock A,RelationMaster B,NewRelation C" & _
                   " WHERE A.GodownID =B.GodownID" & _
                   " AND A.ItemID=C.ItemID" & _
                   " AND C.RelationID=B.RelationID" & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & FreeRejectionsIN & _
                   " OR VoucherType= " & eRejectionVoucher & " )" & _
                   " GROUP BY A.RelationID"

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,B.RelationID" & _
                   " FROM Stock A,RelationMaster B,NewRelation C" & _
                   " WHERE A.GodownID =B.GodownID" & _
                   " AND A.ItemID=C.ItemID" & _
                   " AND C.RelationID=B.RelationID" & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & FreeRejectionsOUT & _
                   " OR VoucherType= " & eRejectionVoucher & " )" & _
                   " GROUP BY B.RelationID"

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,iif( B.StockSales is null,0,B.StockSales)AS StockSales" & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B" & _
                   " ON A.ItemID= B.ItemID"

Call gDbTrans.CreateView("QryStockMerge")

gDbTrans.SqlStmt = " SELECT RelationID" & _
                   " FROM QryStockMerge" & _
                   " WHERE (StockPurchase-StockSales) > 0"
                   
Call gDbTrans.CreateView("QryTotalStock")
                   
'CreateStockMergeQuery = True
End Function


Private Sub ShowFinYear()

If m_frmFinYear Is Nothing Then Set m_frmFinYear = New frmFinYear

m_frmFinYear.Show vbModal


End Sub

Public Function GetDBNameWithPath(ByVal strFinYearFile As String, ByVal YearID As Integer) As String

'Declare the constansts
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"

'Declare the variables
Dim encrKey As String
Dim encrSection As String
Dim strRet As String
Dim strRootPath As String

encrKey = strYear & YearID 'EncryptData(strYear & YearID)
encrSection = strFinYearSection ' EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Function

strRet = strRet 'DecryptData(strRet)

GetDBNameWithPath = ExportExtractToken(strRet, strDBPath, , ",")
    
strRootPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ServerName")
If Len(strRootPath) Then
    Dim strShareName As String
    strShareName = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ShareName")
    strRootPath = "\\" & strRootPath & "\" & strShareName
Else
    strRootPath = App.Path
End If

GetDBNameWithPath = strRootPath & "\" & GetDBNameWithPath & "\" & constDBName

Exit Function

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description
    
End Function

Public Function GetFinYearData(ByVal strFinYearFile As String, ByVal cmbFinYear As ComboBox) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare Variables
Dim I As Integer

Dim FromYear As String
Dim ToYear As String
Dim DBPath As String
Dim strRet As String

Dim encrKey As String
Dim encrSection As String

Const strFromYear = "FromYear#"
Const strToYear = "ToYear#"
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"

GetFinYearData = False

encrKey = strYear & 1 'EncryptData(strYear & 1)
encrSection = strFinYearSection 'EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then CreateDBFirstTime
    
'Call SaveFinYear(strFinYearFile, True)

cmbFinYear.Clear

I = 1
Do
   
    ' Read the dbname from datafile.
    encrKey = strYear & I 'EncryptData(strYear & i)
    encrSection = strFinYearSection 'EncryptData(strFinYearSection)
    
    strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)
    
    If strRet = "" Then Exit Do
    
    GetFinYearData = True
    
    FromYear = ExportExtractToken(strRet, strFromYear, , ",")
    DBPath = ExportExtractToken(strRet, strDBPath, , ",")
    
    ToYear = GetSysFormatDate("31/3/" & CStr(Year(FromYear) + 1))
    cmbFinYear.AddItem "April " & Year(CDate(FromYear)) & " TO March " & Year(CDate(ToYear))
    cmbFinYear.ItemData(cmbFinYear.newIndex) = I
    
    I = I + 1
Loop

Exit Function

ErrLine:
    MsgBox "GetFinYearData()" & vbCrLf & Err.Description
    'Resume
End Function


 
Public Sub LoadFinYearData(ByVal strFinYearFile As String, ByVal YearID As Integer)
'Declare the constansts
Const strFromYear = "FromYear#"
Const strToYear = "ToYear#"
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"


'Declare the variables
Dim encrKey As String
Dim encrSection As String
Dim strRet As String

encrKey = strYear & YearID ' EncryptData(strYear & YearID)
encrSection = strFinYearSection 'EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Sub

'strRet = DecryptData(strRet)

FinIndianFromDate = ExportExtractToken(strRet, strFromYear, , ",")
FinIndianEndDate = ExportExtractToken(strRet, strToYear, , ",")
DayBeginDate = GetLastTransactionDate

If DayBeginDate = "" Then DayBeginDate = FinIndianFromDate
'now Check the date of transaction is in the selected finnancail year
'If DateDiff("d", FinUSFromDate, DayBeginUSDate) < 0 _
    Or DateDiff("d", FinUSEndDate, DayBeginUSDate) > 0 Then
    'select the date which lies in financial Year
'    DayBeginDate = FinIndianEndDate
'End If
If DateDiff("d", FinUSFromDate, DayBeginUSDate) < 0 Then DayBeginDate = FinIndianFromDate
If DateDiff("d", FinUSEndDate, DayBeginUSDate) > 0 And Month(DayBeginUSDate) <> 4 Then
    DayBeginDate = FinIndianEndDate
End If

'InsertPreRequiredHeads

'If DayBeginDate <> "" Then
'    DayBeginDate = FormatDate(DateAdd("d", 1, CDate(DayBeginUSDate)))
'    DayBeginDate = FormatDate(DayBeginDate)
'Else
'    DayBeginDate = FinIndianFromDate
'End If

gStrDate = DayBeginDate ' DayBeginUSDate
Exit Sub

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description

End Sub

Private Function GetLastTransactionDate() As String
'Declare the Variables
Dim rstLastTransDate As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

GetLastTransactionDate = FinIndianFromDate 'FinFromDate
'Fetch the data from the database
gDbTrans.SqlStmt = "SELECT MAX(TransDate)" & _
                                " FROM AccTrans where transdate >= #" & FinUSFromDate & "#"
If gDbTrans.Fetch(rstLastTransDate, adOpenForwardOnly) < 1 Then Exit Function

'return the data
GetLastTransactionDate = FormatField(rstLastTransDate.Fields(0))

Exit Function

ErrLine:
    MsgBox "GetLastTransactionDate: " & Err.Description, vbCritical
    

End Function


Private Function SaveFinYear(ByVal strFinYearFile As String, IsCurFinYear As Boolean) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim I As Long
Dim strFinYear As String
Dim USFromDate As String
Dim USToDate As String
Dim strYearKey As String
Dim strRet As String

Dim encrKey As String
Dim encrSection As String
Dim encrStrFinYear As String

Const strFromYear = "FromYear#"
Const strToYear = "ToYear#"
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"
If Not IsCurFinYear Then
    USFromDate = DateAdd("YYYY", 1, FinUSFromDate)
    USToDate = DateAdd("YYYY", 1, FinUSEndDate)
Else
    USFromDate = FinUSFromDate
    USToDate = FinUSEndDate
End If
strYearKey = "WISDB"
strYearKey = strYearKey & Right$(Str(Year(USFromDate)), 2)

strYearKey = strYearKey & Right$(Str(Year(USToDate)), 2)

'strFinYear = strFromYear & USFromDate & "," & strToYear & USToDate & "," & strDBPath & strYearKey
strFinYear = strFromYear & GetIndianDate(CDate(USFromDate)) & "," & strToYear & GetIndianDate(CDate(USToDate)) & "," & strDBPath & strYearKey

encrStrFinYear = strFinYear 'EncryptData(strFinYear)
I = 1

Do
   
   ' Read the dbname from datafile.
   encrKey = strYear & I 'EncryptData(strYear & I)
   'Do not encrypt the Data
   encrSection = strFinYearSection 'EncryptData(strFinYearSection)
   
   strRet = ReadFromIniFile(encrSection, encrKey, App.Path & "\" & constFINYEARFILE)
   
   If strRet = "" Then Exit Do
   If strRet = encrStrFinYear Then
        MsgBox "You have already Exported the Data !!", vbExclamation
        Exit Function
   End If
   
   I = I + 1
Loop

encrKey = strYear & I 'EncryptData(strYear & i)

''Set the Registry path
#If WIN_2000 Then
    Call CreateRegistryKey(HKEY_LOCAL_MACHINE, INI_FILE_MAPPING_PATH & "\" & constFINYEARFILE)
        
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, INI_FILE_MAPPING_PATH & "\" & constFINYEARFILE, encrSection, "USR:" & constAPPLICATION_NAME & "\" & encrSection)
    
    Call CreateRegistryKey(HKEY_CURRENT_USER, constAPPLICATION_NAME)
#End If
    
Call WriteToIniFile(encrSection, encrKey, encrStrFinYear, App.Path & "\" & constFINYEARFILE)

DbFile = App.Path & "\" & strYearKey & "\" & constDBName
StrPwd = constDBPWD

SaveFinYear = True

Exit Function

ErrLine:
    MsgBox "SaveFinYear()" & vbCrLf & Err.Description
    

End Function

' Retrieves the value for a specified token
' in a given source string.
' The source should be of type :
'       name1=value1,name2=value2,...,name(n)=value(n)
'   similar to DSN strings maintained by ODBC manager.
Public Function ExportExtractToken(src As String, TokenName As String, _
        Optional ByVal TokenDelim As String, Optional ByVal SepDelim As String) As String

' If the src is empty, exit.
If Len(src) = 0 Or _
    Len(TokenName) = 0 Then Exit Function

' Search for the token name.
Dim token_pos As Integer
Dim strSearch As String
Dim Delim_pos As Integer

strSearch = Trim$(TokenName & TokenDelim)

' Search for the token_name in the src string.
 token_pos = InStr(1, src, strSearch, vbTextCompare)
Do
    ' The character before the token_name
    ' should be "," or, it should be the first word.
    ' Else, search for the next occurance of the token.
    If token_pos = 0 Then
        If token_pos = 0 Then
            'Try ignoring the white space
            strSearch = TokenName & " ="
            token_pos = InStr(src, strSearch)
            If token_pos = 0 Then Exit Function
        End If
    ElseIf token_pos = 1 Then
        Exit Do
    ElseIf Mid$(src, token_pos - 1, 1) = "," Then
        Exit Do
    Else
        'Get next occurance.
        token_pos = InStr(token_pos + 1, src, TokenName, vbTextCompare)
    End If
Loop

token_pos = token_pos + Len(strSearch)

' Search for the delimiter ",", after the token_pos.
Delim_pos = InStr(token_pos, src, SepDelim)

If Delim_pos = 0 Then Delim_pos = Len(src) + 1

'Return the token_value.
ExportExtractToken = Mid$(src, token_pos, Delim_pos - token_pos)

End Function




Private Function ExportStock(ByVal RelationID As Long, ByVal GodownID As Byte) As Boolean

Dim MaterialClass As clsMaterial
Dim UsOpDate As String
Dim AsOnIndainDate As String
Dim opBalance As Double
Dim UnitPrice As Currency

ExportStock = False

If RelationID = 0 Then Exit Function

Set MaterialClass = New clsMaterial

' This will get the new fin year's first date
UsOpDate = DateAdd("YYYY", 1, FinUSFromDate)

' This is current year's last date
AsOnIndainDate = DayBeginDate

' Get the Opening Balance
opBalance = MaterialClass.GetItemOnDateClosingStock(RelationID, AsOnIndainDate, GodownID)
UnitPrice = MaterialClass.GetPurchasePrice(RelationID, GodownID, FinIndianFromDate, DayBeginDate)

NewDbClass.SqlStmt = "INSERT INTO Stock (RelationID,GodownID,TransID,Quantity," & _
                " UnitPrice,VoucherType,PurORSaleID,TransDate)" & _
                " VALUES ( " & _
                 RelationID & "," & _
                 GodownID & "," & _
                 1 & "," & _
                 opBalance & "," & _
                 UnitPrice & "," & _
                 Purchase & "," & _
                 0 & "," & _
                 "#" & UsOpDate & "# )"
                 

If NewDbClass.SQLExecute Then ExportStock = True

Set MaterialClass = Nothing

End Function

'This will fetch the Heads from current db and will post them to new db
'Also will export the headid's opening balance
Private Function ExportHeadBalances() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstHeads As ADODB.Recordset

Dim ProfitTypeEnum As wis_AccountType
Dim LossTypeEnum As wis_AccountType
Dim AsOnDate As Date
Dim UsOpDate As Date
Dim headID As Long

Dim AccTransClass As clsAccTrans

ExportHeadBalances = False

ProfitTypeEnum = Profit
LossTypeEnum = Loss

' Fetch Data
gDbTrans.SqlStmt = " SELECT a.* " & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType<> " & ProfitTypeEnum & _
                   " AND b.AccountType<> " & LossTypeEnum

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

AsOnDate = DateAdd("d", 1, FinUSEndDate)  'DayBeginUSDate

' This will get the new fin year's first date
'UsOpDate = DateAdd("YYYY", 1, GetAppFormatDate(FinIndianFromDate))
UsOpDate = DateAdd("YYYY", 1, FinUSFromDate)

Dim opBalance As Currency

Set AccTransClass = New clsAccTrans
    
' Start Loop
Do While Not rstHeads.EOF
    
    'This will Export the HeadID's OpBalance
    'If Not ExportHeadIDBalance(rstHeads.Fields("HeadID")) Then Exit Function
    'This is current year's last date
    
    headID = rstHeads.Fields("HeadID")
    Debug.Assert headID <> 10101
    Debug.Assert headID <> 110001
    Debug.Assert headID <> 30001
    Debug.Assert headID <> 130101
    Debug.Assert headID <> 10101
    Debug.Assert headID <> 10101
    
    ' Get the Opening Balance
    opBalance = AccTransClass.GetOpBalance(headID, AsOnDate)
    
    'Keep only One Entry in the new database that is also for the new year
    NewDbClass.SqlStmt = "Delete * From OpBalance" & _
                    " Where HeadID = " & headID '& _
                    " And OpDate <> #" & DateAdd("yyyy", -1, UsOpDate) & "#"
    If Not NewDbClass.SQLExecute Then NewDbClass.RollBacknRaiseError
    
    'Now Update the New Opening balance
    'NewDbClass.SQLStmt = "UpDate OpBalance Set" & _
                     " OpDate  = #" & UsOpDate & "#," & _
                     " opAmount= " & opBalance & _
                     " WHERE HeadID = " & HeadID
    NewDbClass.SqlStmt = "Insert into OpBalance (OpDate, opAmount,HeadID)" & _
                     " Values (#" & UsOpDate & "#," & _
                     opBalance & ", " & headID & ")"
        
    If Not NewDbClass.SQLExecute Then NewDbClass.RollBacknRaiseError

    'Now Delete all other Entries in the New Data base
    'NewDbClass.SQLStmt = "Delete * From AccTrans " & _
                    " Where HeadID = " & HeadID & _
                    " And OpDate < #" & DateAdd("D", -1, UsOpDate) & "#"

    'If NewDbClass.SQLExecute Then Exit Function
    rstHeads.MoveNext
    
Loop


ExportHeadBalances = True
' Exit
Exit Function

Hell:
    MsgBox "ExportHeads: " & Err.Description
    
End Function

'This will fetch the Heads from current db and will post them to new db
'Also will export the headid's opening balance
Private Function ExportHeads() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstHeads As ADODB.Recordset

Dim ProfitTypeEnum As wis_AccountType
Dim LossTypeEnum As wis_AccountType
Dim AsOnDate As Date
Dim UsOpDate As Date
Dim headID As Long

Dim AccTransClass As clsAccTrans

ExportHeads = False

ProfitTypeEnum = Profit
LossTypeEnum = Loss

' Fetch Data
gDbTrans.SqlStmt = " SELECT a.* " & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType<> " & ProfitTypeEnum & _
                   " AND b.AccountType<> " & LossTypeEnum

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

AsOnDate = DayBeginUSDate

' This will get the new fin year's first date
UsOpDate = DateAdd("YYYY", 1, GetSysFormatDate(FinIndianFromDate))
Dim opBalance As Currency

' Start Loop
Do While Not rstHeads.EOF
    
    'This will Export the HeadID's OpBalance
    'If Not ExportHeadIDBalance(rstHeads.Fields("HeadID")) Then Exit Function
    'This is current year's last date
    
    headID = rstHeads.Fields("HeadID")
    ' Get the Opening Balance
    opBalance = AccTransClass.GetOpBalance(headID, AsOnDate)
    
    'Keep only One Entry in the new database that is also for the OB
    NewDbClass.SqlStmt = "Delete * From OpBalance " & _
                    " Where HeadID = " & headID & _
                    " And OpDate <> #" & DateAdd("yyyy", -1, UsOpDate) & "#"
    If NewDbClass.SQLExecute Then NewDbClass.RollBacknRaiseError
    
    'Now Update the New Opening balance
    NewDbClass.SqlStmt = "UpDate OpBalance Set" & _
                     " OpDate  = #" & UsOpDate & "#," & _
                     " opAmount= " & opBalance & _
                     " WHERE HeadID = " & headID
        
    If NewDbClass.SQLExecute Then Exit Function

    'Now Delete all other Entries in the New Data base
    NewDbClass.SqlStmt = "Delete * From AccTrans " & _
                    " Where HeadID = " & headID & _
                    " And OpDate < #" & DateAdd("D", -1, UsOpDate) & "#"

    If NewDbClass.SQLExecute Then Exit Function
    rstHeads.MoveNext
    
Loop


ExportHeads = True
' Exit
Exit Function

Hell:
    MsgBox "ExportHeads: " & Err.Description
    
End Function


'This will fetch the parents from current db and will post them to new db
Private Function ExportProductGroup() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstProductGrp As ADODB.Recordset

ExportProductGroup = False

' Fetch Data
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM ProductGroup"
Call gDbTrans.Fetch(rstProductGrp, adOpenForwardOnly)

'Before Inserting into stock delete all the existing Product
NewDbClass.SqlStmt = "Delete * FROM ProductGroup"
If Not NewDbClass.SQLExecute Then Exit Function

''''
'First get the Product List Which are in Product
Dim Stockcls As New clsStockRep
Dim rstStock As Recordset
Set rstStock = Stockcls.GetYearEndStockDetails(1)
If rstStock Is Nothing Then Exit Function

'''''

' Start Loop
Do While Not rstProductGrp.EOF
        
    NewDbClass.SqlStmt = "INSERT INTO ProductGroup(GroupID,GroupName)" & _
                         "VALUES(" & _
                         rstProductGrp.Fields("GroupID") & "," & _
                         AddQuotes(rstProductGrp.Fields("GroupName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstProductGrp.MoveNext

nextProduct:

Loop

ExportProductGroup = True
' Exit

Exit Function

Hell:
    MsgBox "ExportParentHeads: " & Err.Description
    
End Function

'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function TransferStock() As Boolean

Dim GodownID As Byte

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset
Dim rstOpStock As ADODB.Recordset
Dim rstGodown As ADODB.Recordset

Dim RelationID As Long

Dim StockBalance As Double

Dim ValueBalance As Currency
Dim TotalValue As Currency
Dim ClosingStock As Currency
Dim PurchasePrice As Currency
Dim OpStockValue As Currency
Dim OpStock As Double

'Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
'Dim fldStockPurchase As ADODB.Field
'Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldUnitName As ADODB.Field

'''
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes
'''

'Iniitialse the function
TransferStock = False

gDbTrans.SqlStmt = "Select Distinct GodownID from GodownDet"
If gDbTrans.Fetch(rstGodown, adOpenForwardOnly) < 1 Then
    TransferStock = True
    Exit Function
End If
'''Now Loop for the Godwon
While Not rstGodown.EOF
    GodownID = FormatField(rstGodown("GodownID"))
    'Prepare the recordset
    'Call CreateOpStockQuery(GodownID, m_FromRpDate, "qryOpStock")
    ''''TEEEEEEEEE
        
    eVoucherType = Purchase
    eFreeVoucherType = FreePurchase
    eTransferVoucher = StockIn
    eRejectionVoucher = RejectionsIn
    
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS StockPurchase,A.RelationID" & _
                       " FROM Stock A,RelationMaster C" & _
                       " WHERE A.GodownID = " & GodownID & _
                       " AND A.RelationID=C.RelationID" & _
                       " AND TransDate BETWEEN #" & FinUSFromDate & "#" & _
                       " AND #" & FinUSEndDate & "#" & _
                       " AND ( VoucherType = " & eVoucherType & _
                       " OR VoucherType = " & eFreeVoucherType & _
                       " OR VoucherType = " & eTransferVoucher & _
                       " OR VoucherType = " & FreeRejectionsIN & _
                       " OR VoucherType= " & eRejectionVoucher & " )" & _
                       " GROUP BY A.RelationID"
    
    Call gDbTrans.CreateView("QryStockPurchase")
    
    eVoucherType = Sales
    eFreeVoucherType = FreeSales
    eTransferVoucher = StockOut
    eSootVoucher = StockSoot
    eRejectionVoucher = RejectionsOut
    
      gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,A.RelationID" & _
                       " FROM Stock A,RelationMaster C" & _
                       " WHERE A.GodownID = " & GodownID & _
                       " AND A.RelationID=C.RelationID" & _
                       " AND TransDate BETWEEN #" & FinUSFromDate & "#" & _
                       " AND #" & FinUSEndDate & "# " & _
                       " AND ( VoucherType = " & eVoucherType & _
                       " OR VoucherType = " & eFreeVoucherType & _
                       " OR VoucherType = " & eSootVoucher & _
                       " OR VoucherType = " & eTransferVoucher & _
                       " OR VoucherType = " & FreeRejectionsOUT & _
                       " OR VoucherType= " & eRejectionVoucher & " )" & _
                       " GROUP BY A.RelationID"
        
    Call gDbTrans.CreateView("QryStockSales")
    
    gDbTrans.SqlStmt = " SELECT A.*,iif( B.StockSales is null,0,B.StockSales)AS StockSales" & _
                       " FROM QryStockPurchase A LEFT JOIN QryStockSales B" & _
                       " ON A.RelationID = B.RelationID"
    
    Call gDbTrans.CreateView("QryStockMerge")
    
    gDbTrans.SqlStmt = "SELECT (StockPurchase-StockSales) AS OpStock,RelationID " & _
                       " FROM QryStockMerge " & _
                       " WHERE (StockPurchase - StockSales) > 0"
    gDbTrans.CreateView ("qryClStock")
    
    ''''TEEEEEEEEE
    'If Not PrepareRecordsetForTotalStock(rstStock, GodownID) Then Exit Function
    '''''SSSSSS
    gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, " & _
                        " F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY C.GroupName,D.ProductName "

    If gDbTrans.Fetch(rstStock, adOpenStatic) <= 0 Then Exit Function
    '''''SSSSSS
    
    gDbTrans.SqlStmt = "SELECT OpStock,RelationID FROM QryCLStock"
    Call gDbTrans.Fetch(rstOpStock, adOpenForwardOnly)
    
    'Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
    ''''''TTTTTTTTTTTTT
    Dim VoucherType As Wis_VoucherTypes
    
    VoucherType = Purchase
    eVoucherType = StockIn
    
    gDbTrans.SqlStmt = " SELECT Max(TransID) AS maxTransID, RelationID" & _
                       " FROM Stock " & _
                       " WHERE (VoucherType = " & VoucherType & _
                       " OR VoucherType = " & eVoucherType & " ) " & _
                       " AND TransDate BETWEEN #" & FinUSFromDate & "#" & _
                       " AND #" & FinUSEndDate & "#" & _
                       " GROUP BY RelationID"
                       
    Call gDbTrans.CreateView("QryMaxTransID")
    
    'Fire the Query
    gDbTrans.SqlStmt = " SELECT UnitPrice,a.RelationID " & _
                       " FROM Stock a,qryMaxTransID b " & _
                       " WHERE a.TransID=b.maxTransID " & _
                       " AND a.RelationID=b.RelationID"
    
    gDbTrans.CreateView ("qryPrice")
    
    gDbTrans.SqlStmt = " SELECT UnitPrice as UnitPrice1,A.RelationID " & _
                       " FROM qryPrice A " & _
                       " WHERE UnitPrice > 0"
    gDbTrans.CreateView ("qryPrice1")
    
    'There are Some Items which will have purchase price will be 0
    'Like Container whih will come free with some materails
    'and User Sales and Gets some amount From It
    'So For Such Items Get the Item Price from Sales details
    'Some thng Like Max Price
    
    'So get the the Free items from getting whose purchase price is 0
    gDbTrans.SqlStmt = " SELECT Max(UnitPrice) As UnitPrice1,RelationID " & _
                       " FROM Stock A " & _
                       " WHERE RelationID IN (SELECT RelationID " & _
                       " FROM qryPrice WHERE UnitPrice = 0) " & _
                       " Group BY RelationID"
    gDbTrans.CreateView ("qryPrice2")
    
    gDbTrans.SqlStmt = " Select RelationID,UnitPrice1 as UnitPrice" & _
                       " From qryPrice1 " & _
                       " UNION " & _
                       " Select RelationID,UnitPrice1 as UnitPrice" & _
                       " From qryPrice2 "
    
    gDbTrans.CreateView ("qryItemPrice")
    ''''''TTTTTTTTTTTTT
    
    gDbTrans.SqlStmt = "Select * from qryItemPrice"
    Call gDbTrans.Fetch(rstPrice, adOpenDynamic)
    
    'Set fldStockPurchase = rstStock.Fields("StockPurchase")
    'Set fldStockSales = rstStock.Fields("StockSales")
    Set fldRelationID = rstStock.Fields("RelationID")
    Set fldGroupID = rstStock.Fields("GroupID")
    Set fldGroupName = rstStock.Fields("GroupName")
    Set fldProductName = rstStock.Fields("ProductName")
    Set fldUnitName = rstStock.Fields("UnitName")
    
    
    ClosingStock = 0
    'Set MaterialClass = New clsMaterial
    
    Do
        If rstStock.EOF Then Exit Do
        
        OpStock = 0: OpStockValue = 0
        PurchasePrice = 0
        
        StockBalance = FormatField(rstStock("StockPurchase")) - FormatField(rstStock("StockSales"))
        RelationID = fldRelationID.Value
        
        If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then _
                            PurchasePrice = rstPrice.Fields("UnitPrice")
        ValueBalance = StockBalance * PurchasePrice
        
        
        If gDbTrans.FindRecord(rstOpStock, "RelationID=" & RelationID) Then OpStock = rstOpStock.Fields("opStock")
        OpStockValue = OpStock * PurchasePrice
'        TotalValue = TotalValue + ValueBalance + OpStockValue
        
        OpStock = OpStock + StockBalance
        OpStockValue = OpStockValue + ValueBalance
        'Total Closing Stock
        'ClosingStock = ClosingStock + ValueBalance + OpStockValue
        
        ''Now Insert closing stock as Opening Stock for Next Fin Year
        ''i.e. Update the Stock
        NewDbClass.SqlStmt = "Update Stock Set " & _
                " TransDate =#" & DateAdd("d", 1, FinUSEndDate) & "#," & _
                " Quantity = " & StockBalance & "," & _
                " UnitPrice = " & PurchasePrice & "" & _
                " WHERE GodownID = " & GodownID & _
                " AND PurOrSaleID = 0 AND RelationID = " & RelationID & _
                " ANd TransID = 1 AND VoucherType = " & Purchase & _
                " AND TransDate =#" & FinUSFromDate & "#"
        If OpStock > 0 Then If Not NewDbClass.SQLExecute Then Exit Function
        
        'Move the recordset to Next Position
        rstStock.MoveNext
    Loop
    
    rstGodown.MoveNext
    ''Now end Loop for Godown
Wend

'Now Delete all other details of the Stock Table
'i.e. Delete the transaction of Previous Year
NewDbClass.SqlStmt = "Delete From Stock " & _
                " WHERE TransDate <=#" & FinUSEndDate & "#"

If Not NewDbClass.SQLExecute Then Exit Function

'Now Delete Realtion Table Whic Are not hvaing any stock
NewDbClass.SqlStmt = "Delete From RelationMaster " & _
                " WHERE RelationID not In (Select Distinct RelationID From Stock)"

If Not NewDbClass.SQLExecute Then Exit Function


TransferStock = True


Exit Function

ErrLine:
    If Err Then
        'Resume
        MsgBox "TransferStock(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function


'This will fetch the Products from current db and will post them to new db
Private Function ExportProducts() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstProducts As ADODB.Recordset

ExportProducts = False

' Fetch Data
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM Products"

Call gDbTrans.Fetch(rstProducts, adOpenForwardOnly)

' Start Loop
Do While Not rstProducts.EOF
    
    NewDbClass.SqlStmt = "INSERT INTO Products(AliasID,GroupID,ProductID,ProductName)" & _
                         "VALUES(" & _
                         rstProducts.Fields("AliasID") & "," & _
                         rstProducts.Fields("GroupID") & "," & _
                         rstProducts.Fields("ProductID") & "," & _
                         AddQuotes(rstProducts.Fields("ProductName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstProducts.MoveNext
Loop

ExportProducts = True

' Exit

Exit Function

Hell:
    MsgBox "ExportProducts: " & Err.Description
    
End Function
'This will fetch the Units from current db and will post them to new db
Private Function ExportTransPortMode() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstTransport As ADODB.Recordset


' Fetch Data
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM TransportMode"

Call gDbTrans.Fetch(rstTransport, adOpenForwardOnly)

' Start Loop
Do While Not rstTransport.EOF
    
    NewDbClass.SqlStmt = "INSERT INTO TransportMode (TransModeID,TransModeName)" & _
                         "VALUES(" & _
                         rstTransport.Fields("TransModeID") & "," & _
                         AddQuotes(rstTransport.Fields("{TransModeName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstTransport.MoveNext
    
Loop

ExportTransPortMode = True

' Exit

Exit Function

Hell:
    MsgBox "ExportTransPortMode " & Err.Description
    
End Function

'THis FUnction Delete the All the Transaction of the
'previous year & update the TransID
Private Function UpdateHeadTrans() As Boolean
    
    Dim MinTransId As Long
    
    'Get the mIN tRANSCTION OF the Acctrans table
    Dim rstNew As Recordset
    NewDbClass.SqlStmt = "Select MIN(TransID) from AccTrans " & _
            " Where TransDate > #" & FinUSEndDate & "#"
    If NewDbClass.Fetch(rstNew, adOpenDynamic) Then _
                        MinTransId = FormatField(rstNew(0)) + 1
    If MinTransId <= 1 Then MinTransId = 0
    
    'Now Delete the All transaction of the previus year
    NewDbClass.SqlStmt = "DELETE * FROM ACCTRANS WHERE " & _
                " TransDate < #" & FinUSEndDate & "#"
    
    'NewDbClass.BeginTrans
    If NewDbClass.SQLExecute Then
                
        'NOw Fetch the Details And Reduce the TransID
        MinTransId = MinTransId - 5
        If MinTransId < 0 Then MinTransId = 0
        'Delete the Trans Particulares
        NewDbClass.SqlStmt = "Delete * from TransParticulars WHERE TransID >= " & MinTransId
        Call NewDbClass.SQLExecute
        
        'Fetch the Details
       
        NewDbClass.SqlStmt = "Select Distinct TransID FROM AccTrans Order by TransID "
        If NewDbClass.Fetch(rstNew, adOpenForwardOnly) > 0 And MinTransId Then
            While Not rstNew.EOF
                NewDbClass.SqlStmt = "UPDATE ACCTRANS SET TRANSID = TransID - " & MinTransId & _
                                    " WHERE TransID = " & FormatField(rstNew("TransID"))
                Call NewDbClass.SQLExecute
                rstNew.MoveNext
            Wend
        End If
    End If
    
    UpdateHeadTrans = True

End Function

'This will convert current financial year and will post the
'data to new mdb with series of functions
Public Function YearEndFunctions() As Boolean

On Error GoTo Hell:

If Not SaveFinYear(App.Path & "\" & constFINYEARFILE, False) Then Exit Function

Dim DbUtils As New clsDBUtilities
Dim srcFile As String
srcFile = GetDataBaseName
gDbTrans.CloseDB

'Create DB
'If Not DbUtils.CopyDefaultDatabase(srcFile, DbFile) Then Exit Function
'copy the the existing data base to the ext Year
Dim pos As Integer
Do
    If InStr(pos + 1, DbFile, "\") = 0 Then Exit Do
    pos = InStr(pos + 1, DbFile, "\")
Loop

'make directrory
If pos Then
    Call MakeDirectories(Left(DbFile, pos - 1))
End If
FileCopy srcFile, DbFile

If Not gDbTrans.OpenDB(srcFile, StrPwd) Then Exit Function
If Not NewDbClass.OpenDB(DbFile, StrPwd) Then Exit Function

' Export ParentHeads
NewDbClass.BeginTrans

If ExportFunctions Then
    'commit the transaction
    NewDbClass.CommitTrans
Else
    'rollback the transaction
    NewDbClass.RollBack
End If

' CLose the new db
NewDbClass.CloseDB

MsgBox "Database Exported", vbInformation

Call CreateStockMergeQuery
Exit Function

Hell:

End Function

Private Sub Class_Initialize()

Set NewDbClass = New clsTransact
'Set NewDBClass = CreateObject("Transaction.Transact")
End Sub


Private Sub Class_Terminate()

On Error Resume Next

Set NewDbClass = Nothing

End Sub



Private Sub m_frmFinYear_CancelClicked()
    Unload m_frmFinYear
    Set m_frmFinYear = Nothing
    End
End Sub

Private Sub m_frmFinYear_OkClicked(strFinFromDate As String)
    Dim strUSFinToDate As Date
    
    FinIndianFromDate = strFinFromDate
    
    strUSFinToDate = GetSysFormatDate(strFinFromDate)
    strUSFinToDate = DateAdd("YYYY", 1, CDate(strUSFinToDate))
    
    strUSFinToDate = DateAdd("d", -1, strUSFinToDate)
    FinIndianEndDate = GetIndianDate(strUSFinToDate)

End Sub


