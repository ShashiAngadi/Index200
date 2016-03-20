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
Private ConnectionString As String
Private DbFile As String
Private StrPwd As String

Private WithEvents m_frmFinYear As frmFinYear
Attribute m_frmFinYear.VB_VarHelpID = -1


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
'

End Sub

'private withevents
'This will create the new mdb for the next finacial year

Private Function CreateNewDb() As Boolean

On Error GoTo Hell:

Dim UtilClass As clsDBUtilities

CreateNewDb = False

Set UtilClass = New clsDBUtilities


'If Not UtilClass.CreateDBFromDB(DbFile, StrPwd) Then Exit Function
'Set UtilClass = Nothing

'copy the db file from existing data base
'FileCopy SourceFile, TargetFile
'Creete the Direcory
Dim strSource As String

strSource = gDbTrans.DatabaseName
'Close the Data base
gDbTrans.CloseDB
'FileCopy App.Path & "\BlankDataBase\" & constDBName, DbFile
Kill DbFile
FileCopy strSource, DbFile

'Now open the Data base
If Not gDbTrans.OpenDB(strSource, StrPwd) Then
    MsgBox "Unable to open the Source data base", vbInformation, wis_MESSAGE_TITLE
    End
End If



CreateNewDb = True

Exit Function

Hell:
    If Err.Number > 0 Then MsgBox "CreateNewDB: " & Err.Description
    Set UtilClass = Nothing
    
End Function


Private Function ExportFunctions() As Boolean

'Export the parent heads
'If Not ExportParentHeads Then Exit Function

' Export Heads excluding revenue heads and Opening Balances
'If Not ExportHeads Then Exit Function
    
' Export Company Details
'If Not ExportCompanyDet Then Exit Function
    
'Export Company Details
'If Not ExportCompany Then Exit Function

'Export Company Code
'If Not ExportCompanyCode Then Exit Function

'Export Godown Details
'If Not ExportGodownDetails Then Exit Function

'Export Install
'If Not ExportInstall Then Exit Function

' Export Products
'If Not ExportPlaces Then Exit Function

' Export Product Group
'If Not ExportProductGroup Then Exit Function
    
' Export Products
'If Not ExportProducts Then Exit Function
    
'Export Install
'If Not ExportSetup Then Exit Function

' Export Products
'If Not ExportTransPortMode Then Exit Function

' Export Units
'If Not ExportUnits Then Exit Function
    
' Export Users
'If Not ExportUsers Then Exit Function

' Export Stock & Relations
If Not ExportStocknRelations Then Exit Function
    
ExportFunctions = True

End Function

Private Function ExportGodownDetails() As Boolean
' Handle error
On Error GoTo Hell:

' Declarations
Dim rstGodownDet As ADODB.Recordset

ExportGodownDetails = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM GodownDet"

Call gDbTrans.Fetch(rstGodownDet, adOpenForwardOnly)

' Start Loop
Do While Not rstGodownDet.EOF
    
    NewDbClass.SQLStmt = " INSERT INTO GodownDet (GodownID,GodownName,Address," & _
                         "PhoneNo,ContactPerson,MobileNo,Email) " & _
               " VALUES ( " & _
               rstGodownDet.Fields("GodownID") & "," & _
               AddQuotes(rstGodownDet.Fields("GodownName")) & "," & _
               AddQuotes(rstGodownDet.Fields("Address")) & "," & _
               AddQuotes(rstGodownDet.Fields("PhoneNO")) & "," & _
               AddQuotes(rstGodownDet.Fields("ContactPerson")) & "," & _
               AddQuotes(rstGodownDet.Fields("MobileNo")) & "," & _
               AddQuotes(rstGodownDet.Fields("Email")) & " ) "
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstGodownDet.MoveNext
Loop

ExportGodownDetails = True
' Exit

Exit Function

Hell:
    MsgBox "ExportGodownDetails: " & Err.Description
End Function


'This will export the given HeadID's Balance
Private Function ExportHeadIDBalance(ByVal HeadID As Long) As Boolean

Dim AccTransClass As clsAccTrans
Dim USOpDate As String
Dim AsOnDate As String
Dim OpBalance As Currency

ExportHeadIDBalance = False

If HeadID = 0 Then Exit Function

Set AccTransClass = New clsAccTrans

' This will get the new fin year's first date
USOpDate = DateAdd("YYYY", 1, FormatDate(FinFromDate))

' This is current year's last date
AsOnDate = DayBeginDate

' Get the Opening Balance
OpBalance = AccTransClass.GetOpBalance(HeadID, AsOnDate)

NewDbClass.SQLStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 HeadID & "," & _
                 "#" & USOpDate & "#," & _
                 OpBalance & ")"
                 
If NewDbClass.SQLExecute Then ExportHeadIDBalance = True

Set AccTransClass = Nothing

End Function

'This will fetch the parents from current db and will post them to new db
Private Function ExportParentHeads() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstParents As ADODB.Recordset
Dim rstTemp As ADODB.Recordset
Dim i As Integer

ExportParentHeads = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT *" & _
                   " FROM ParentHeads"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

'THere may be some heads already exists like Share capital
'There fore it may rise some error soneglect the error
' Start Loop
Do While Not rstParents.EOF
    NewDbClass.SQLStmt = "SELECT * FROM PARENTHEADS WHERE PARENTID = " & rstParents.Fields("ParentID")
    
    i = i + 1
    If NewDbClass.Fetch(rstTemp, adOpenForwardOnly) < 1 Then
    NewDbClass.SQLStmt = "INSERT INTO ParentHeads(ParentID,ParentName,AccountType,PrintDetailed,PrintOrder,UserCreated)" & _
                         "VALUES(" & _
                         rstParents.Fields("ParentID") & "," & _
                         AddQuotes(rstParents.Fields("ParentName")) & "," & _
                         rstParents.Fields("AccountType") & "," & _
                         1 & "," & _
                         i & "," & _
                         rstParents.Fields("UserCreated") & ")"
                             
    
    If Not NewDbClass.SQLExecute Then Exit Function
    End If
    rstParents.MoveNext
    
Loop

ExportParentHeads = True

Exit Function

Hell:
    MsgBox "ExportParentHeads: " & Err.Description
    
End Function
'This will Export the Relations and Stock to New mdb
Private Function ExportStocknRelations() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstRelation As ADODB.Recordset
Dim rstNewRelation As ADODB.Recordset
Dim RelnId As Long
Dim PrevRelnId As Long

Dim MaterialClass As clsMaterial
Dim USOpDate As String
Dim AsOnDate As String
Dim OpBalance As Double
Dim UnitPrice As Currency
Dim ItemID As Long, GodownID As Integer
Set MaterialClass = New clsMaterial

' This will get the new fin year's first date
USOpDate = DateAdd("YYYY", 1, FormatDate(FinFromDate))

' This is current year's last date
AsOnDate = DayBeginDate

'Make Some changes to the database
'Remove the error making nulls
'PUT 1/1/100 Whereever NUll date are there
gDbTrans.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " MfgDate = #1/1/1000#" & _
                    " WHERE MfgDate is NULL"
gDbTrans.BeginTrans
gDbTrans.SQLExecute

gDbTrans.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " ExpDate = #1/1/1000#" & _
                    " WHERE ExpDate is NULL"
gDbTrans.SQLExecute
gDbTrans.CommitTrans

' Fetch Data
gDbTrans.SQLStmt = " SELECT a.*,b.*" & _
                   " FROM RelationMaster A,NewRelation B" & _
                   " WHERE A.RelationID=B.RelationID" & _
                   " ORDER BY A.RelationID"

Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)

PrevRelnId = 0
' Start Loop
Do While Not rstRelation.EOF

    RelnId = rstRelation.Fields("A.RelationID")
    
    If PrevRelnId <> RelnId Then
    
        NewDbClass.SQLStmt = "INSERT INTO RelationMaster (RelationID,GodownID,HeadID," & _
                             "GroupID,ProductID,UnitID,ReOrderLevel )" & _
                             " VALUES ( " & _
                             rstRelation.Fields("A.RelationID") & "," & _
                             rstRelation.Fields("GodownID") & "," & _
                             rstRelation.Fields("HeadID") & "," & _
                             rstRelation.Fields("GroupID") & "," & _
                             rstRelation.Fields("ProductID") & "," & _
                             rstRelation.Fields("UnitID") & "," & _
                             IIf(IsNull(rstRelation.Fields("ReOrderLevel")), "NULL", rstRelation.Fields("ReOrderLevel")) & ")"
        
        If Not NewDbClass.SQLExecute Then Exit Function
        
        PrevRelnId = RelnId
    End If
    
    ItemID = rstRelation.Fields("ItemID")
    GodownID = rstRelation.Fields("GodownID")
    ' Get the Opening Balance
    OpBalance = MaterialClass.GetItemOnDateClosingStock(ItemID, AsOnDate, GodownID)
    'If There s No Opening balance then Need noto to Insert the Details
    If OpBalance > 0 Then
        NewDbClass.SQLStmt = " INSERT INTO NewRelation (ItemID,RelationID,NewRelationID," & _
                        "TradingPrice,MRP,SalesPrice,CST,KST,BatchNo,MfgDate,ExpDate," & _
                        "FirstSales )" & _
                       " VALUES ( " & _
                       rstRelation.Fields("ItemID") & "," & _
                       rstRelation.Fields("B.RelationID") & "," & _
                       rstRelation.Fields("NewRelationID") & "," & _
                       rstRelation.Fields("TradingPrice") & "," & _
                       rstRelation.Fields("MRP") & "," & _
                       rstRelation.Fields("SalesPrice") & "," & _
                       rstRelation.Fields("CST") & "," & _
                       rstRelation.Fields("KST") & ",'" & _
                       rstRelation.Fields("BatchNo") & "'," & _
                       "#" & rstRelation.Fields("MfgDate") & "#," & _
                       "#" & rstRelation.Fields("ExpDate") & "#," & _
                       rstRelation.Fields("FirstSales") & ")"
    
        If Not NewDbClass.SQLExecute Then Exit Function
        
        UnitPrice = MaterialClass.GetPurchasePrice(ItemID, GodownID, FinFromDate, DayBeginDate)
        NewDbClass.SQLStmt = " INSERT INTO Stock (ItemID,TransID,GodownID,Quantity," & _
                " UnitPrice,VoucherType,PurORSaleID,TransDate)" & _
                " VALUES ( " & _
                 ItemID & "," & _
                 1 & "," & _
                 GodownID & "," & _
                 OpBalance & "," & _
                 UnitPrice & "," & _
                 Purchase & "," & _
                 0 & "," & _
                 "#" & USOpDate & "# )"
                 

        If Not NewDbClass.SQLExecute Then Exit Function
        
    End If

    rstRelation.MoveNext
    
Loop

'UN do the Temporary Changes
gDbTrans.BeginTrans
gDbTrans.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " MfgDate = NULL " & _
                    " WHERE MfgDate = #1/1/1000#"
gDbTrans.SQLExecute

gDbTrans.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " ExpDate = NULL " & _
                    " WHERE ExpDate = #1/1/1000#"
gDbTrans.SQLExecute
gDbTrans.CommitTrans

'UN do the Temporary Changes
NewDbClass.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " MfgDate = NULL " & _
                    " WHERE MfgDate = #1/1/1000#"
NewDbClass.SQLExecute

NewDbClass.SQLStmt = "UPDATE NEWRELATION SET" & _
                    " ExpDate = NULL " & _
                    " WHERE ExpDate = #1/1/1000#"
NewDbClass.SQLExecute



ExportStocknRelations = True
' Exit

Exit Function

Hell:
    MsgBox "ExportStocknRelations: " & Err.Description
    ExportStocknRelations = False
    
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

USFromDate = FormatDate(FinFromDate)
USAsOnDate = FormatDate(FinEndDate)    'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SQLStmt = " SELECT SUM(Quantity) as StockPurchase,B.RelationID" & _
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


gDbTrans.SQLStmt = " SELECT SUM(Quantity) as StockSales,B.RelationID" & _
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


gDbTrans.SQLStmt = " SELECT A.*,iif( B.StockSales is null,0,B.StockSales)AS StockSales" & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B" & _
                   " ON A.ItemID= B.ItemID"

Call gDbTrans.CreateView("QryStockMerge")

gDbTrans.SQLStmt = " SELECT RelationID" & _
                   " FROM QryStockMerge" & _
                   " WHERE (StockPurchase-StockSales) > 0"
                   
Call gDbTrans.CreateView("QryTotalStock")
                   
CreateStockMergeQuery = True
End Function


Private Function InsertPreRequiredHeads() As Boolean
'Pre Required Heads are
'Cash in Hand HeadID=40001
'Tax Received Account=80001
'Discount Received Account=80101
'Tax Paid Account =90001
'Discount Given Account=90101

'Declare the variables
Dim Rst As ADODB.Recordset
Dim StartDateUS As String

'Setup an error handler...
On Error GoTo ErrLine

If FinFromDate = "" Then Exit Function

StartDateUS = FormatDate(FinFromDate)

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_CashHeadID
                  
gDbTrans.BeginTrans

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_CashHeadID & "," & _
                       "'Cash In Hand'" & "," & _
                       wis_CashParentID & " ) "
    gDbTrans.BeginTrans
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_CashHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_GeneralCashSalesCustomer

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_GeneralCashSalesCustomer & "," & _
                       "'Cash Sales'" & "," & _
                       wis_DebitorsParentID & " ) "
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_GeneralCashSalesCustomer & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO CompanyCreation ( HeadID,CompanyName,CompanyType,OtherType )" & _
                       " VALUES ( " & _
                       wis_GeneralCashSalesCustomer & "," & _
                       "'Cash Sales'" & "," & _
                       2 & "," & _
                       2 & " ) "
                       
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If


gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_TaxReceivedHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_TaxReceivedHeadID & "," & _
                       "'Tax Received A/C'" & "," & _
                       wis_IncomeParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_TaxReceivedHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If


gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_DiscountReceivedHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_DiscountReceivedHeadID & "," & _
                       "'Discount Received A/C'" & "," & _
                       wis_TradingIncomeParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_DiscountReceivedHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If


gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_TaxPaidHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_TaxPaidHeadID & "," & _
                       "'Tax Paid A/C'" & "," & _
                       wis_ExpenseParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_TaxPaidHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_DiscountPaidHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_DiscountPaidHeadID & "," & _
                       "'Discount Paid A/C'" & "," & _
                       wis_TradingExpenseParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_DiscountPaidHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_SalesHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_SalesHeadID & "," & _
                       "'Sales A/C'" & "," & _
                       wis_SalesParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_SalesHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_PurchaseHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_PurchaseHeadID & "," & _
                       "'Purchase A/C'" & "," & _
                       wis_PurchaseParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_PurchaseHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.SQLStmt = " SELECT HeadID" & _
                  " FROM Heads" & _
                  " WHERE HeadID=" & wis_AccumulatedPLHeadID
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    gDbTrans.SQLStmt = " INSERT INTO Heads ( HeadID,HeadName,ParentID )" & _
                       " VALUES ( " & _
                       wis_AccumulatedPLHeadID & "," & _
                       "'Accumulated Profit Or Loss'" & "," & _
                       wis_AccumulatedPLParentID & " ) "
    
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SQLStmt = " INSERT INTO OpBalance ( HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       wis_AccumulatedPLHeadID & "," & _
                       "#" & StartDateUS & "#," & _
                       0 & " ) "
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.CommitTrans

InsertPreRequiredHeads = True

Exit Function

ErrLine:
    MsgBox "InsertPreRequiredHeads() " & Err.Description
    
End Function

Private Sub ShowFinYear()

If m_frmFinYear Is Nothing Then Set m_frmFinYear = New frmFinYear

m_frmFinYear.Show vbModal


End Sub

'
Public Function GetDBNameWithPath(ByVal strFinYearFile As String, ByVal YearID As Integer) As String
'Declare the constansts
Const strDBPath = "DBPath#"
Const strYear = "Year"
Const strFinYearSection = "FinYear"

'Declare the variables
Dim encrKey As String
Dim encrSection As String
Dim strRet As String


encrKey = EncryptData(strYear & YearID)
encrSection = EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Function

 strRet = DecryptData(strRet)

GetDBNameWithPath = ExportExtractToken(strRet, strDBPath, , ",")
    
GetDBNameWithPath = App.Path & "\" & GetDBNameWithPath & "\" & constDBName

Exit Function

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description
    
End Function

Public Function GetFinYearData(ByVal strFinYearFile As String, ByVal cmbFinYear As ComboBox) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare Variables
Dim i As Integer

Dim USFromYear As String
Dim USToYear As String
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

encrKey = EncryptData(strYear & 1)
encrSection = EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)
strRet = "òš ¤Ì´¼š8  &Ö Ì´¼š8  &öúÞ¼–®8ä¼–"
If strRet = "" Then CreateDBFirstTime
    

Call SaveFinYear(strFinYearFile, True)

cmbFinYear.Clear

i = 1
Do
   
    ' Read the dbname from datafile.
    'encrKey = EncryptData(strYear & i)
    'encrSection = EncryptData(strFinYearSection)
    
    'strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)
    strRet = ReadFromIniFile(strFinYearSection, strYear & i, strFinYearFile)
        
    Debug.Assert 1 = 0
    If strRet = "" Then Exit Do
    
    GetFinYearData = True
    
    strRet = DecryptData(strRet)
    USFromYear = ExportExtractToken(strRet, strFromYear, , ",")
    DBPath = ExportExtractToken(strRet, strDBPath, , ",")
    
    USToYear = "3/31/" & CStr(Year(USFromYear) + 1)
    
    cmbFinYear.AddItem "April " & Year(CDate(USFromYear)) & " TO March " & Year(CDate(USToYear))
    cmbFinYear.ItemData(cmbFinYear.NewIndex) = i
    
    i = i + 1
Loop

Exit Function

ErrLine:
    MsgBox "GetFinYearData()" & vbCrLf & Err.Description
    
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


encrKey = EncryptData(strYear & YearID)
encrSection = EncryptData(strFinYearSection)

strRet = ReadFromIniFile(encrSection, encrKey, strFinYearFile)

If strRet = "" Then Exit Sub

strRet = DecryptData(strRet)

FinFromDate = FormatDate(ExportExtractToken(strRet, strFromYear, , ","))
FinEndDate = FormatDate(ExportExtractToken(strRet, strToYear, , ","))
DayBeginDate = GetLastTransactionDate

InsertPreRequiredHeads


If DayBeginDate <> "" Then
    DayBeginDate = CStr(DateAdd("d", 1, CDate(FormatDate(DayBeginDate))))
    DayBeginDate = FormatDate(DayBeginDate)
Else
    DayBeginDate = FinFromDate
End If

Exit Sub

ErrLine:
    MsgBox "GetDBNameWithPath()" & vbCrLf & Err.Description

End Sub

Private Function GetLastTransactionDate() As String
'Declare the Variables
Dim rstLastTransDate As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

GetLastTransactionDate = FinFromDate
'Fetch the data from the database
gDbTrans.SQLStmt = " SELECT MAX(TransDate)" & _
                                " FROM AccTrans"
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
Dim i As Long
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
    USFromDate = DateAdd("YYYY", 1, FormatDate(FinFromDate))
    USToDate = DateAdd("YYYY", 1, FormatDate(FinEndDate))
Else
    USFromDate = FormatDate(FinFromDate)
    USToDate = FormatDate(FinEndDate)
End If
strYearKey = "Mat"
strYearKey = strYearKey & Right$(Str(Year(USFromDate)), 2)

strYearKey = strYearKey & Right$(Str(Year(USToDate)), 2)

strFinYear = strFromYear & USFromDate & "," & strToYear & USToDate & "," & strDBPath & strYearKey

encrStrFinYear = EncryptData(strFinYear)

' Now get the key for the Year and also check the existing year
i = 1

Do
   
   ' Read the dbname from datafile.
   encrKey = EncryptData(strYear & i)
   encrSection = EncryptData(strFinYearSection)
   
   strRet = ReadFromIniFile(encrSection, encrKey, App.Path & "\" & constFINYEARFILE)
   
   If strRet = "" Then Exit Do
   
   If strRet = encrStrFinYear Then
        MsgBox "You have already Exported the Data !!", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
   End If
   
   i = i + 1
Loop

encrKey = EncryptData(strYear & i)

''Set the Registry path
#If WIN_2000 Then
    Call CreateRegistryKey(HKEY_LOCAL_MACHINE, INI_FILE_MAPPING_PATH & "\" & constFINYEARFILE)
        
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, INI_FILE_MAPPING_PATH & "\" & constFINYEARFILE, encrSection, "USR:" & constAPPLICATION_NAME & "\" & encrSection)
    
    Call CreateRegistryKey(HKEY_CURRENT_USER, constAPPLICATION_NAME)
#End If
    
Call WriteToIniFile(encrSection, encrKey, encrStrFinYear, App.Path & "\" & constFINYEARFILE)

DbFile = App.Path & "\" & strYearKey & "\" & constDBName
'Create the dirctory
On Error Resume Next
Call MkDir(App.Path & "\" & strYearKey)
If Err.Number = 71 Then Exit Function

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
Public Function ExportExtractToken(src As String, TokenName As String, Optional ByVal TokenDelim As String, Optional ByVal SepDelim As String) As String

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

Private Function ExportStock(ByVal ItemID As Long, ByVal GodownID As Byte) As Double

Dim MaterialClass As clsMaterial
Dim USOpDate As String
Dim AsOnDate As String
Dim OpBalance As Double
Dim UnitPrice As Currency

ExportStock = False

If ItemID = 0 Then Exit Function

Set MaterialClass = New clsMaterial

' This will get the new fin year's first date
USOpDate = DateAdd("YYYY", 1, FormatDate(FinFromDate))

' This is current year's last date
AsOnDate = DayBeginDate

' Get the Opening Balance
OpBalance = MaterialClass.GetItemOnDateClosingStock(ItemID, AsOnDate, GodownID)
'If There s No OPening balance then Need noto to Insert the Details
If OpBalance > 0 Then
    UnitPrice = MaterialClass.GetPurchasePrice(ItemID, GodownID, FinFromDate, DayBeginDate)
    NewDbClass.SQLStmt = " INSERT INTO Stock (ItemID,TransID,GodownID,Quantity," & _
                " UnitPrice,VoucherType,PurORSaleID,TransDate)" & _
                " VALUES ( " & _
                 ItemID & "," & _
                 1 & "," & _
                 GodownID & "," & _
                 OpBalance & "," & _
                 UnitPrice & "," & _
                 Purchase & "," & _
                 0 & "," & _
                 "#" & USOpDate & "# )"
                 

    If NewDbClass.SQLExecute Then ExportStock = OpBalance
Else
    ExportStock = 0
End If
Set MaterialClass = Nothing

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

ExportHeads = False

ProfitTypeEnum = Profit
LossTypeEnum = Loss

' Fetch Data
gDbTrans.SQLStmt = " SELECT a.* " & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND b.AccountType<> " & ProfitTypeEnum & _
                   " AND b.AccountType<> " & LossTypeEnum

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

' Start Loop
Do While Not rstHeads.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO Heads(HeadID,HeadName,ParentID)" & _
                         "VALUES(" & _
                         rstHeads.Fields("HeadID") & "," & _
                         AddQuotes(rstHeads.Fields("HeadName")) & "," & _
                         rstHeads.Fields("ParentID") & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    ' This will Export the HeadID's OpBalance
    If Not ExportHeadIDBalance(rstHeads.Fields("HeadID")) Then Exit Function
    
    rstHeads.MoveNext
    
Loop


ExportHeads = True
' Exit
Exit Function

Hell:
    MsgBox "ExportHeads: " & Err.Description
    
End Function

'This will fetch the Company from current db and will post them to new db
Private Function ExportCompanyCode() As Boolean
' Declarations
Dim rstCompany As ADODB.Recordset

' Handle error
On Error GoTo Hell:

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM CompanyCode"

Call gDbTrans.Fetch(rstCompany, adOpenForwardOnly)


' Start Loop
Do While Not rstCompany.EOF
    
NewDbClass.SQLStmt = " INSERT INTO CompanyCode (HeadID,ShortName,SameState,DisAmount)" & _
                " VALUES ( " & _
               rstCompany.Fields("HeadID") & "," & _
               AddQuotes(rstCompany.Fields("ShortName")) & "," & _
               FormatField(rstCompany.Fields("SameState")) & "," & _
               FormatField(rstCompany.Fields("DisAmount")) & " ) "
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstCompany.MoveNext
Loop

ExportCompanyCode = True
' Exit

Exit Function

Hell:
    MsgBox "ExportCompanyCode " & Err.Description
    
End Function
'This will fetch the Company from current db and will post them to new db
Private Function ExportCompanyDet() As Boolean
' Declarations
Dim rstCompany As ADODB.Recordset
Dim KSTNo As String
Dim CSTNo As String
Dim Address As String
Dim PhoneNo As String
Dim ContactPerson As String
Dim MobileNo As String
Dim Email As String
Dim DLNo As String


' Handle error
On Error GoTo Hell:

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM CompanyDet"

Call gDbTrans.Fetch(rstCompany, adOpenForwardOnly)


' Start Loop
Do While Not rstCompany.EOF
    

    KSTNo = IIf(IsNull(rstCompany.Fields("VATNo")), "", rstCompany.Fields("VATNo"))
    CSTNo = IIf(IsNull(rstCompany.Fields("CSTNo")), "", rstCompany.Fields("CSTNo"))
    Address = IIf(IsNull(rstCompany.Fields("Address")), "", rstCompany.Fields("Address"))
    PhoneNo = IIf(IsNull(rstCompany.Fields("PhoneNo")), "", rstCompany.Fields("PhoneNo"))
    ContactPerson = IIf(IsNull(rstCompany.Fields("ContactPerson")), "Mr. X", rstCompany.Fields("ContactPerson"))
    MobileNo = IIf(IsNull(rstCompany.Fields("MobileNo")), "", rstCompany.Fields("MobileNo"))
    Email = IIf(IsNull(rstCompany.Fields("Email")), "", rstCompany.Fields("Email"))
    DLNo = IIf(IsNull(rstCompany.Fields("DlNo")), "", rstCompany.Fields("DlNo"))
    
NewDbClass.SQLStmt = " INSERT INTO CompanyDet (CompanyID,CompanyName,VATNo," & _
                    "CSTNo,DOB,Address,PhoneNo,ContactPerson,MobileNo,Email,DLNo) " & _
               " VALUES ( " & _
               rstCompany.Fields("CompanyID") & "," & _
               AddQuotes(rstCompany.Fields("CompanyName")) & "," & _
               AddQuotes(KSTNo) & "," & _
               AddQuotes(CSTNo) & "," & _
               "#" & IIf(FormatField(rstCompany.Fields("DOB")) = "", "1/1/75", rstCompany.Fields("DOB")) & "#," & _
               AddQuotes(Address) & "," & _
               AddQuotes(PhoneNo) & "," & _
               AddQuotes(ContactPerson) & "," & _
               AddQuotes(MobileNo) & "," & _
               AddQuotes(Email) & "," & _
               AddQuotes(DLNo) & "  ) "
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstCompany.MoveNext
Loop

ExportCompanyDet = True
' Exit

Exit Function

Hell:
    MsgBox "ExportParentHeads: " & Err.Description
    
End Function

'This will fetch the Company from current db and will post them to new db
Private Function ExportCompany() As Boolean
' Declarations
Dim rstCompany As ADODB.Recordset
Dim KSTNo As String
Dim CSTNo As String
Dim Address As String
Dim PhoneNo As String
Dim ContactPerson As String
Dim MobileNo As String
Dim Email As String
Dim DLNo As String


' Handle error
On Error GoTo Hell:

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM CompanyCreation"

Call gDbTrans.Fetch(rstCompany, adOpenForwardOnly)


' Start Loop
Do While Not rstCompany.EOF
    

    KSTNo = IIf(IsNull(rstCompany.Fields("KST")), "", rstCompany.Fields("KST"))
    CSTNo = IIf(IsNull(rstCompany.Fields("CST")), "", rstCompany.Fields("CST"))
    Address = IIf(IsNull(rstCompany.Fields("Address")), "", rstCompany.Fields("Address"))
    PhoneNo = IIf(IsNull(rstCompany.Fields("PhoneNo")), "", rstCompany.Fields("PhoneNo"))
    ContactPerson = IIf(IsNull(rstCompany.Fields("ContactPerson")), "Mr. X", rstCompany.Fields("ContactPerson"))
    MobileNo = IIf(IsNull(rstCompany.Fields("MobileNo")), "", rstCompany.Fields("MobileNo"))
    Email = IIf(IsNull(rstCompany.Fields("Email")), "", rstCompany.Fields("Email"))
    DLNo = IIf(IsNull(rstCompany.Fields("DlNo")), "", rstCompany.Fields("DlNo"))
    
NewDbClass.SQLStmt = " INSERT INTO CompanyCreation (HeadID,CompanyName,CompanyType, " & _
               " KST,CST,Address,PhoneNo,ContactPerson,MobileNo,Email,SameState,DLNo,PlaceID ) " & _
               " VALUES ( " & _
               rstCompany.Fields("HeadID") & "," & _
               AddQuotes(rstCompany.Fields("CompanyName")) & "," & _
               rstCompany.Fields("CompanyType") & "," & _
               AddQuotes(KSTNo) & "," & _
               AddQuotes(CSTNo) & "," & _
               AddQuotes(Address) & "," & _
               AddQuotes(PhoneNo) & "," & _
               AddQuotes(ContactPerson) & "," & _
               AddQuotes(MobileNo) & "," & _
               AddQuotes(Email) & "," & _
               FormatField(rstCompany.Fields("SameState")) & "," & _
               AddQuotes(DLNo) & "," & _
               FormatField(rstCompany.Fields("PlaceID")) & " ) "
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstCompany.MoveNext
Loop

ExportCompany = True
' Exit

Exit Function

Hell:
    MsgBox "ExportParentHeads: " & Err.Description
    
End Function


'This will fetch the Users from current db and will post them to new db
Private Function ExportUsers() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstUsers As ADODB.Recordset

ExportUsers = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Users"

Call gDbTrans.Fetch(rstUsers, adOpenForwardOnly)

' Start Loop
Do While Not rstUsers.EOF
    
NewDbClass.SQLStmt = " INSERT INTO Users (UserID,LoginName,LoginPassword,Permissions) " & _
               " VALUES ( " & _
               rstUsers.Fields("UserID") & "," & _
               AddQuotes(rstUsers.Fields("LoginName")) & "," & _
               AddQuotes(rstUsers.Fields("LoginPassword")) & "," & _
               rstUsers.Fields("Permissions") & " ) "
    
    'If Not NewDbClass.SQLExecute Then Exit Function
    NewDbClass.SQLExecute
    rstUsers.MoveNext
Loop

ExportUsers = True
' Exit

Exit Function

Hell:
    MsgBox "ExportUsers: " & Err.Description
    
End Function

'This will fetch the parents from current db and will post them to new db
Private Function ExportProductGroup() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstProductGrp As ADODB.Recordset

ExportProductGroup = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM ProductGroup"

Call gDbTrans.Fetch(rstProductGrp, adOpenForwardOnly)


' Start Loop
Do While Not rstProductGrp.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO ProductGroup(GroupID,GroupName)" & _
                         "VALUES(" & _
                         rstProductGrp.Fields("GroupID") & "," & _
                         AddQuotes(rstProductGrp.Fields("GroupName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstProductGrp.MoveNext
    
Loop

ExportProductGroup = True
' Exit

Exit Function

Hell:
    MsgBox "ExportParentHeads: " & Err.Description
    
End Function
'This will fetch the Products from current db and will post them to new db
Private Function ExportProducts() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstProducts As ADODB.Recordset

ExportProducts = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Products"

Call gDbTrans.Fetch(rstProducts, adOpenForwardOnly)

' Start Loop
Do While Not rstProducts.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO Products(AliasID,GroupID,ProductID,ProductName)" & _
                         "VALUES(" & _
                         IIf(IsNull(rstProducts.Fields("AliasID")), "NULL", rstProducts.Fields("AliasID")) & "," & _
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
'This will fetch the Products from current db and will post them to new db
Private Function ExportSetup() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstSetup As ADODB.Recordset


' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Setup"

Call gDbTrans.Fetch(rstSetup, adOpenForwardOnly)

' Start Loop
Do While Not rstSetup.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO SetUp(KeyData,TransDate,ValueData)" & _
                         "VALUES(" & _
                        AddQuotes(rstSetup.Fields("KeyData")) & "," & _
                        "#" & rstSetup.Fields("TransDate") & "#," & _
                        AddQuotes(rstSetup.Fields("ValueData")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstSetup.MoveNext
Loop

ExportSetup = True

' Exit

Exit Function

Hell:
    MsgBox "ExportSetup " & Err.Description
    
End Function
'This will fetch the Products from current db and will post them to new db
Private Function ExportInstall() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstInstall As ADODB.Recordset

ExportInstall = False

' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Install"

Call gDbTrans.Fetch(rstInstall, adOpenForwardOnly)

' Start Loop
Do While Not rstInstall.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO Install(KeyData,ValueData)" & _
                         "VALUES(" & _
                        AddQuotes(rstInstall.Fields("KeyData")) & "," & _
                        AddQuotes(rstInstall.Fields("ValueData")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstInstall.MoveNext
Loop

ExportInstall = True

' Exit

Exit Function

Hell:
    MsgBox "ExportProducts: " & Err.Description
    
End Function


'This will fetch the Units from current db and will post them to new db
Private Function ExportPlaces() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstPlaces As ADODB.Recordset


' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Places"

Call gDbTrans.Fetch(rstPlaces, adOpenForwardOnly)

' Start Loop
Do While Not rstPlaces.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO Places (PlaceID,PlaceName)" & _
                         "VALUES(" & _
                         rstPlaces.Fields("PlaceID") & "," & _
                         AddQuotes(rstPlaces.Fields("PlaceName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstPlaces.MoveNext
    
Loop

ExportPlaces = True

' Exit

Exit Function

Hell:
    MsgBox "ExportPlaces " & Err.Description
    
End Function
'This will fetch the Units from current db and will post them to new db
Private Function ExportTransPortMode() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstTransport As ADODB.Recordset


' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM TransportMode"

Call gDbTrans.Fetch(rstTransport, adOpenForwardOnly)

' Start Loop
Do While Not rstTransport.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO TransportMode (TransModeID,TransModeName)" & _
                         "VALUES(" & _
                         rstTransport.Fields("TransModeID") & "," & _
                         AddQuotes(rstTransport.Fields("TransModeName")) & ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    rstTransport.MoveNext
    
Loop

ExportTransPortMode = True

' Exit

Exit Function

Hell:
    MsgBox "ExportTransPortMode " & Err.Description
    
End Function

'This will fetch the Units from current db and will post them to new db
Private Function ExportUnits() As Boolean

' Handle error
On Error GoTo Hell:

' Declarations
Dim rstUnits As ADODB.Recordset

ExportUnits = False
' Fetch Data
gDbTrans.SQLStmt = " SELECT * " & _
                   " FROM Units"

Call gDbTrans.Fetch(rstUnits, adOpenForwardOnly)

' Start Loop
Do While Not rstUnits.EOF
    
    NewDbClass.SQLStmt = "INSERT INTO Units(UnitID,UnitName)" & _
                         "VALUES(" & _
                         rstUnits.Fields("UnitID") & "," & _
                         AddQuotes(rstUnits.Fields("UnitName")) & _
                          ")"
    
    If Not NewDbClass.SQLExecute Then Exit Function
    
    
    rstUnits.MoveNext
    
Loop

ExportUnits = True

' Exit

Exit Function

Hell:
    MsgBox "ExportUnits: " & Err.Description
    
End Function


'This function will validate the yearend function
Private Function ValidateExportData(ByVal encrStrKey As String) As Boolean
' Handle error
On Error GoTo Hell:

ValidateExportData = False
' Declarations

' Validate


' Exit

ValidateExportData = True
Exit Function

Hell:
    MsgBox "ValidateExportData: " & Err.Description
    
End Function

'This will convert current financial year and will post the
' data to new mdb with series of functions

Public Function YearEndFunctions() As Boolean

On Error GoTo Hell:

If Not SaveFinYear(App.Path & "\" & constFINYEARFILE, False) Then Exit Function

' Create DB
If Not CreateNewDb Then Exit Function

If Not NewDbClass.OpenDB(DbFile, StrPwd) Then Exit Function

' Export ParentHeads
NewDbClass.BeginTrans

If ExportFunctions Then
    'commit the transaction
    NewDbClass.CommitTrans
    MsgBox "Database Exported", vbInformation
Else
    'rollback the transaction
    NewDbClass.RollBack
    MsgBox "Database Not NOT NOT Exported", vbCritical
End If

' CLose the new db
NewDbClass.CloseDB



Call CreateStockMergeQuery

Exit Function

Hell:

End Function


Private Sub Class_Initialize()

Set NewDbClass = New clsTransact

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
Dim strUSFinToDate As String

FinFromDate = strFinFromDate

strUSFinToDate = FormatDate(strFinFromDate)
strUSFinToDate = DateAdd("YYYY", 1, CDate(strUSFinToDate))

strUSFinToDate = DateAdd("d", -1, CDate(strUSFinToDate))
FinEndDate = FormatDate(strUSFinToDate)

End Sub


