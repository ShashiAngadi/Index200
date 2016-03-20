VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Class Name  : clsStockReport
'Author      : Lingappa Sindhanur , Waves Information Systems
'Date        : 10th April 2002
'Description : Gives the details of the reports pertaining to the stock or
'               Inventory, Such as total stock , Item wise stock , total purchase, sales etc
'
'Revisions
Option Explicit

Private WithEvents m_frmRptDt As frmRptDt
Attribute m_frmRptDt.VB_VarHelpID = -1
Private WithEvents m_WebDoc As HTMLDocument
Attribute m_WebDoc.VB_VarHelpID = -1

Private m_frmWeb As frmWeb

Private m_WebTable As HTMLTable
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell


' To Handle Report Dates
Private m_FromRpDate As String
Private m_ToRpDate As String
Private m_FromUSDate As Date
Private m_ToUSDate As Date


Private m_TheReportType As Wis_ReportType

' This function gives the purchase price of the item from the database
' Inputs : RelationID as long
' Output : Gives the Purchase Price of the Item
Public Sub LoadPurchasePrice(ByVal fromDate As String, _
                                 ByVal toDate As String, _
                                 rstPrice As ADODB.Recordset)
'Trap an error
On Error GoTo ErrLine
'declare Variables

Dim VoucherType As Wis_VoucherTypes
Dim eVoucherType As Wis_VoucherTypes

VoucherType = Purchase
eVoucherType = StockIn

gDbTrans.SqlStmt = " SELECT Max(TransID) AS maxTransID, RelationID" & _
                   " FROM Stock " & _
                   " WHERE ( VoucherType = " & VoucherType & _
                   " OR VoucherType = " & eVoucherType & " ) " & _
                   " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                   " AND #" & GetSysFormatDate(toDate) & "#" & _
                   " GROUP BY RelationID"
                   
Call gDbTrans.CreateView("QryMaxTransID")


'Fire the Query
gDbTrans.SqlStmt = " SELECT UnitPrice,a.RelationID " & _
                   " FROM Stock a,qryMaxTransID b " & _
                   " WHERE  a.TransID=b.maxTransID " & _
                   " AND a.RelationID=b.RelationID"
                   
If gDbTrans.Fetch(rstPrice, adOpenDynamic) < 1 Then Exit Sub

Exit Sub

ErrLine:
    
End Sub

'Declare for Font Size Setting
Private Sub PauseApplication()
Dim Start

Start = Timer
Do
    'it will be in the loop
    If Timer >= Start + 1 Then Exit Do
Loop



End Sub

Private Function PrepareRstForSalesInvoices(ByRef rstInvoices As ADODB.Recordset, _
                                        ByVal headID As Long, ByVal GodownID As Byte) As Boolean

PrepareRstForSalesInvoices = False

gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,InvoiceAmount " & _
                   " FROM Sales " & _
                   " WHERE HeadID = " & headID & _
                   " AND GodownID = " & GodownID & _
                   " AND InvoiceDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " ORDER BY InvoiceDate"
 
If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

PrepareRstForSalesInvoices = True

End Function
Private Function PrepareRstForPurchasesInvoices(ByRef rstInvoices As ADODB.Recordset, _
                                        ByVal headID As Long) As Boolean

PrepareRstForPurchasesInvoices = False

gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,TransDate,InvoiceAmount " & _
                   " FROM Purchase " & _
                   " WHERE HeadID = " & headID & _
                   " AND TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " ORDER BY InvoiceDate,TransDate "

If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

PrepareRstForPurchasesInvoices = True

End Function

'This Function Prepare the Recodset Total Stock for group and product
Private Function PrepareRstOfStockForGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, _
                                                        ByVal intGroupID As Integer, _
                                                        ByVal lngProductID As Long, ByVal GodownID As Byte) As Boolean

'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

PrepareRstOfStockForGroupAndProduct = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.PRoductID = " & lngProductID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY E.CompanyName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockForGroupAndProduct = True

End Function

'This function will prepeare the recordset for the Stock Group
' Inputs
'       byRef rstTotalStock  as recordset
'       byval intGroupID as integer
' OutPuts
'       rstTotalStock as Recordset
Private Function PrepareRstOfStockForGroupIncBranch(ByRef rstTotalStock As ADODB.Recordset, _
                                          ByVal intGroupID As Integer) As Boolean
'Declare the Variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eStockSoot As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes


PrepareRstOfStockForGroupIncBranch = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,GodownID,RelationID FROM Stock " & _
                   " WHERE TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY GodownID,RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eStockSoot = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,GodownID,RelationID FROM Stock " & _
                   " WHERE TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eStockSoot & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY GodownID,RelationID "

Call gDbTrans.CreateView("QryStockSales")

gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID" & _
                   " AND A.GodownID=B.GodownID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName,A.GodownID,G.GodownName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F,GodownDet G " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID=G.GodownID " & _
                   " ORDER BY E.HeadID,ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function


PrepareRstOfStockForGroupIncBranch = True

End Function

'This function will prepeare the recordset for the Stock Group
' Inputs
'       byRef rstTotalStock  as recordset
'       byval intGroupID as integer
' OutPuts
'       rstTotalStock as Recordset
Private Function PrepareRstOfStockForGroup(ByRef rstTotalStock As ADODB.Recordset, _
                                          ByVal intGroupID As Integer, ByVal GodownID As Byte) As Boolean
'Declare the Variables
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eStockSoot As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes


PrepareRstOfStockForGroup = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eStockSoot = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eStockSoot & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")

gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY E.HeadID "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function


PrepareRstOfStockForGroup = True

End Function



'
Private Function PrepareRstOfStockForManuGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, _
                 ByVal lngHeadID As Long, ByVal intGroupID As Integer, _
                 ByVal lngProductID As Long, ByVal GodownID As Byte) As Boolean

'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

PrepareRstOfStockForManuGroupAndProduct = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID" & _
                   " FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.RelationID," & _
                        " D.ProductName," & _
                        " F.UnitName " & _
                   " FROM QryStockMerge AS A,RelationMaster AS B," & _
                        " ProductGroup AS C,Products AS D, " & _
                        " Units AS F" & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.HeadID = " & lngHeadID & _
                   " AND B.PRoductID = " & lngProductID & _
                   " AND B.GroupID = C.GroupID" & _
                   " AND B.ProductID = D.ProductID" & _
                   " AND B.UnitID =F.UnitID" & _
                   " ORDER BY D.ProductName"

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockForManuGroupAndProduct = True

End Function


Private Function PrepareRecordsetForTotalStock(ByRef rstTotalStock As ADODB.Recordset, ByVal GodownID As Byte) As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

PrepareRecordsetForTotalStock = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY E.CompanyName,C.GroupName,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForTotalStock = True

End Function

Private Function PrepareRecordsetForStockManufacturer(ByRef rstTotalStock As ADODB.Recordset, ByVal lngHeadID As Long, ByVal GodownID As Byte) As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

PrepareRecordsetForStockManufacturer = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")

gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.HeadID = " & lngHeadID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY E.CompanyName,C.GroupName,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForStockManufacturer = True
End Function

Private Function PrepareRstForStockManuAndGroup(ByRef rstTotalStock As ADODB.Recordset, _
                                                                ByVal lngHeadID As Long, _
                                                                ByVal intGroupID As Integer, ByVal GodownID As Byte) As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

PrepareRstForStockManuAndGroup = False

USFromDate = FinUSFromDate
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.HeadID = " & lngHeadID & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY E.CompanyName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForStockManuAndGroup = True

End Function


Private Function PrepareRstTransactionsOfManufacturer(ByRef rstTotalStock As ADODB.Recordset, _
                                                    ByVal lngHeadID As Long) As Boolean
'Decalre the variables

PrepareRstTransactionsOfManufacturer = False



'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D," & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND B.HeadID = " & lngHeadID & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstTransactionsOfManufacturer = True

End Function

Private Function PrepareRstTransactionsOfManuAndGroup(ByRef rstTotalStock As ADODB.Recordset, _
                                                            ByVal lngHeadID As Long, _
                                                            ByVal intGroupID As Integer) As Boolean


PrepareRstTransactionsOfManuAndGroup = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " C.ProductName,D.UnitID,D.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C,Units AS D " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.UnitID =D.UnitID " & _
                   " AND B.HeadID = " & lngHeadID & _
                   " AND B.GroupID = " & intGroupID & _
                   " GROUP BY C.ProductName,D.UnitID, D.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
                        
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstTransactionsOfManuAndGroup = True

End Function


Private Function PrepareRstForTransactionsOfGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, ByVal intGroupID As Integer, _
                                                       ByVal lngProductID As Long) As Boolean


PrepareRstForTransactionsOfGroupAndProduct = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " C.ProductName,D.CompanyName,D.HeadID,E.UnitID,E.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C, CompanyCreation AS D," & _
                        " Units AS E " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.HEadID = D.HeadID " & _
                   " AND B.UnitID =E.UnitID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.ProductID = " & lngProductID & _
                   " GROUP BY D.CompanyName, D.HeadID, " & _
                        " C.ProductName,E.UnitID, E.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForTransactionsOfGroupAndProduct = True

End Function
Private Function PrepareRstForTransactionsOfManuGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, _
                                                    ByVal lngHeadID As Long, _
                                                    ByVal intGroupID As Integer, _
                                                    ByVal lngProductID As Long) As Boolean

PrepareRstForTransactionsOfManuGroupAndProduct = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " C.ProductName,E.UnitID,E.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C,Units AS E " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.UnitID =E.UnitID " & _
                   " AND B.HEadID = " & lngHeadID & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.ProductID = " & lngProductID & _
                   " GROUP BY C.ProductName,E.UnitID, E.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
                        
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForTransactionsOfManuGroupAndProduct = True

End Function


Private Function PrepareRstForTransactionsOfGroup(ByRef rstTotalStock As ADODB.Recordset, _
                                                    ByVal intGroupID As Integer) As Boolean

PrepareRstForTransactionsOfGroup = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " C.ProductName,D.CompanyName,D.HeadID,E.UnitID,E.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C, CompanyCreation AS D," & _
                        " Units AS E " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.HEadID = D.HeadID " & _
                   " AND B.UnitID =E.UnitID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " GROUP BY D.CompanyName, D.HeadID, " & _
                        " C.ProductName,E.UnitID, E.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForTransactionsOfGroup = True

End Function


Private Function PrepareRecordsetForTransactions(ByRef rstTotalStock As ADODB.Recordset) As Boolean

PrepareRecordsetForTransactions = False

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalQuantity,A.RelationID," & _
                        " B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,VoucherType" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E," & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID, C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,VoucherType "
If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForTransactions = True

End Function



'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRstOfPurchaseForGroup(ByRef rstTotalPurchase As ADODB.Recordset, _
                                                    ByVal intGroupID As Integer) As Boolean

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRstOfPurchaseForGroup = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " C.ProductName,D.CompanyName,D.HeadID,E.UnitID,E.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C, CompanyCreation AS D, Units AS E " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.HEadID = D.HeadID " & _
                   " AND B.UnitID =E.UnitID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " GROUP BY D.CompanyName, D.HeadID, " & _
                        " C.ProductName,E.UnitID,E.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfPurchaseForGroup = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRstOfPurchaseForGroupAndProduct(ByRef rstTotalPurchase As ADODB.Recordset, _
                                                    ByVal intGroupID As Integer, _
                                                    ByVal lngProductID As Long) As Boolean


Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRstOfPurchaseForGroupAndProduct = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " C.ProductName,D.CompanyName,D.HeadID,E.UnitID,E.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS C, CompanyCreation AS D, Units AS E " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = C.ProductID " & _
                   " AND B.HEadID = D.HeadID " & _
                   " AND B.UnitID =E.UnitID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.ProductID = " & lngProductID & _
                   " GROUP BY D.CompanyName, D.HeadID, " & _
                        " C.ProductName,E.UnitID,E.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfPurchaseForGroupAndProduct = True

End Function
'This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTCompanyNamesForPurchase(ByVal CompanyName As String)

'Increment Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)

'Print the Company Name
With m_WebRow.cells(1)
    .innerText = CompanyName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

'Increment one More Row
'Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

End Sub

''This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTCompanyNames(ByVal CompanyName As String)
'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(1)
    .innerText = CompanyName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

'Increment the Row
'Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 3)

End Sub

''This function will print the Group Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTGroupName(ByVal GroupName As String)

'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(1)
    .innerText = "  " & GroupName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
End Sub

''This function will print the Group Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTGroupNameForTransactions(ByVal GroupName As String)

'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(1)
    .innerText = "  " & GroupName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
End Sub

''This function will print the Group Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written'
Public Sub PRINTGroupNameForPurchase(ByVal GroupName As String)
'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)
        

'Print the group Name
With m_WebRow.cells(1)
    .innerText = "  " & GroupName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
End Sub


'This will print the PrductName, StockBalance,ValueBalance of the
'''This function will print the product Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTProductsForTotalPurchase(ByVal TransDate As String, ByVal ProductName As String, _
                        ByVal UnitPrice As Currency, ByVal UnitName As String, _
                        ByVal StockQuantity As Double, ByVal StockValue As Currency, _
                        ByVal FreeQuantity As Double, ByVal FreeValue As Currency)
    
'ByVal FreeQuantity As Double, ByVal FreeValue As Currency
'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)

With m_WebRow.cells(0)
    .innerText = TransDate
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(1)
    .innerText = "  " & ProductName & " (" & UnitName & ")"
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(2)
    .innerText = FormatCurrency(UnitPrice)
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

'With m_WebRow.cells(3)
'    .innerText = OpStock
'    .Align = "Right"
'    .Style.fontWeight = "Normal"
'    .Style.FontSize = 13
'    If gLangOffSet <> 0 Then .Style.FontSize = 15
'End With
With m_WebRow.cells(3)
    .innerText = StockQuantity
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(4)
    .innerText = FormatCurrency((StockValue))
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(5)
    .innerText = FreeQuantity
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(6)
    .innerText = FreeValue
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
End Sub

'This will print the PrductName, StockBalance,ValueBalance of the
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTProductsForTransactions(ByVal ProductName As String, ByVal UnitPrice As Currency, _
                        ByVal Quantity As Double, TransDate As String, _
                        ByVal UnitName As String, ByVal VoucherType As Wis_VoucherTypes, _
                        ByVal IsSameRow As Boolean)


Dim colno As Integer

Select Case VoucherType
    Case Purchase
        colno = 3
    Case Sales
        colno = 4
End Select
If Not IsSameRow Then
    ''Inrcrement the row
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

    'Print the data
    With m_WebRow.cells(0)
        .innerText = TransDate
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
    With m_WebRow.cells(1)
        .innerText = "  " & ProductName & " (" & UnitName & ")"
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
    With m_WebRow.cells(2)
        .innerText = FormatCurrency(UnitPrice)
        .Align = "Right"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
End If

With m_WebRow.cells(colno)
    .innerText = Quantity
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

End Sub

'This will print the PrductName, StockBalance,ValueBalance of the
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTPurchaseInvoiceNos(ByVal TransDate As String, _
                                       ByVal InvoiceDate As String, _
                                       ByVal InvoiceNo As String, _
                                       ByVal InvoiceAmount As Currency)

'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
    
'Print the data
With m_WebRow.cells(0)
    .innerText = InvoiceDate
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(1)
    .innerText = TransDate
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With m_WebRow.cells(2)
    .innerText = InvoiceNo
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(3)
    .innerText = FormatCurrency(InvoiceAmount)
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

End Sub


'This will print the PrductName, StockBalance,ValueBalance of the
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTSalesInvoiceNos(ByVal InvoiceDate As String, _
                                       ByVal InvoiceNo As String, _
                                       ByVal InvoiceAmount As Currency)

'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 3)
    
'Print the data
With m_WebRow.cells(0)
    .innerText = InvoiceDate
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With m_WebRow.cells(1)
    .innerText = InvoiceNo
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(2)
    .innerText = FormatCurrency(InvoiceAmount)
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

End Sub



'This will print the PrductName, StockBalance,ValueBalance of the
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTProductsForTotalStock(ByVal ProductName As String, _
                                       ByVal StockBalance As Double, _
                                       ByVal ValueBalance As Currency, _
                                       ByVal UnitName As String, ByVal PurchasePrice, ByVal SlNo As Integer)

'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
    

'Print the data
With m_WebRow.cells(0)
    .innerText = SlNo
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(1)
    .innerText = "  " & ProductName & "(" & UnitName & ")"
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(2)
    .innerText = PurchasePrice
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With m_WebRow.cells(3)
    .innerText = StockBalance
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With
With m_WebRow.cells(4)
    .innerText = FormatCurrency((ValueBalance))
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

End Sub



'''This function will print the Total  into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTTotalForPurchase(ByVal TotalValue As Currency, ByVal FreeValue As Currency)
                                    
'ByVal FreeStock As Double,ByVal FreeValue As Currency)


'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)

'Print Total Text
With m_WebRow.cells(1)
    .innerText = "Total"
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

''Print the total
'With m_WebRow.cells(3)
'    .innerText = TotalStock
'    .Align = "Right"
'    .Style.fontWeight = "Bold"
'    .Style.FontSize = 15
'    If gLangOffSet <> 0 Then .Style.FontSize = 17
'End With
With m_WebRow.cells(4)
    .innerText = FormatCurrency((TotalValue))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
'With m_WebRow.cells(5)
'    .innerText = FreeStock
'    .Align = "Right"
'    .Style.fontWeight = "Bold"
'    '.Style.FontSize = "x-small"
'End With
With m_WebRow.cells(6)
    .innerText = FormatCurrency((FreeValue))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub

'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTClosingStock(ByVal ClosingStock As Currency, ByVal ColCount As Integer, Optional ByVal GrandTotal As String)

'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, ColCount)
            
'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, ColCount)

'Print the Total
If GrandTotal = "" Then
    With m_WebRow.cells(1)
        .innerText = "Closing Stock Value"
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
Else
    With m_WebRow.cells(2)
        .innerText = GrandTotal
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If
With m_WebRow.cells(ColCount - 1)
    .innerText = FormatCurrency((ClosingStock))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub


'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTTotal(ByVal TotalValue As Currency)

'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
            
'Print the Total
With m_WebRow.cells(1)
    .innerText = "Total"
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(4)
    .innerText = FormatCurrency((TotalValue))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub
'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTTotalForTransactions(ByVal TotalPurchase As Double, _
                                      ByVal TotalSales As Double)

'ByVal FreePurchase As Currency,, ByVal FreeSales As Double
'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
            
'Print the Total
With m_WebRow.cells(1)
    .innerText = "Total"
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(3)
    .innerText = TotalPurchase
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(4)
    .innerText = TotalSales
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub



Public Sub SetWebDocument()

'This will set the  oncotextmenu  to return false
' so that if user clicks the right button popup menu will not be displayed
'm_frmWeb.web.Document.write "<Body oncontextmenu = " & AddQuotes("return false", True) & ">"
'content="text/html; charset=x-user-defined" http-equiv=Content-Type

'm_frmWeb.web.Document.write "<Body oncontextmenu =" & "'return false'" & ">"

Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName


End Sub

Public Sub SetCompanyTitle(Optional ByVal GodownID As Byte)
'Declare the variables
Dim rstCompany As ADODB.Recordset
Dim rstBranch As ADODB.Recordset

Dim eCompanyType As wis_CompanyType

eCompanyType = Enum_Self

gDbTrans.SqlStmt = " SELECT CompanyName,Address " & _
                   " FROM CompanyCreation " & _
                   " WHERE CompanyType = " & eCompanyType

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 1 Then Exit Sub


If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    
    Call StartHtmlTable("HeaderTbl", m_WebDoc.body, 1)
    
    m_WebTable.border = 0
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("CompanyName").Value
        .Style.FONTSIZE = "19"
        .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("Address").Value
        .Style.FONTSIZE = "13"
        .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
End With

'Get the Branch
If GodownID <= 1 Then Exit Sub

gDbTrans.SqlStmt = " SELECT GodownName FROM GodownDet" & _
                   " WHERE GodownID = " & GodownID
If gDbTrans.Fetch(rstBranch, adOpenForwardOnly) < 1 Then Exit Sub

With m_WebDoc
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    With m_WebRow.cells(0)
      .innerText = rstBranch.Fields("GodownName")
      .Style.FONTSIZE = "16"
      .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
End With





End Sub

'This function will list the Purchase Invoice Numbers of the
' Specified Company
'Inputs     HeadID  as long This is the CompanyID
Public Function ShowSalesInvoices(ByVal headID As Long, ByVal GodownID As Byte) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the varaibles
Dim rstInvoices As ADODB.Recordset

Dim InvoiceDate As String
Dim InvoiceNo As String
Dim InvoiceAmount As Currency

Dim CompanyName As String

Dim MaterialClass As clsMaterial

'Declare Fild variables to Store Field Values
Dim fldInvoiceNo As ADODB.Field
Dim fldInvoiceDate As ADODB.Field
Dim fldInvoiceAmount As ADODB.Field


ShowSalesInvoices = False

If Not PrepareRstForSalesInvoices(rstInvoices, headID, GodownID) Then Exit Function


Set fldInvoiceNo = rstInvoices.Fields("InvoiceNo")
Set fldInvoiceDate = rstInvoices.Fields("InvoiceDate")
Set fldInvoiceAmount = rstInvoices.Fields("InvoiceAmount")


Set MaterialClass = New clsMaterial

CompanyName = MaterialClass.GetCompanyName(headID)

Call SetCompanyTitle(1)

Call SetSalesInvoicesHeader(CompanyName)

Do While Not rstInvoices.EOF
    InvoiceDate = FormatField(rstInvoices.Fields("InvoiceDate"))
    InvoiceNo = rstInvoices.Fields("InvoiceNo").Value
    InvoiceAmount = rstInvoices.Fields("InvoiceAmount").Value
    
    Call PRINTSalesInvoiceNos(InvoiceDate, InvoiceNo, InvoiceAmount)
    
        
    rstInvoices.MoveNext
Loop

Set MaterialClass = Nothing

ShowSalesInvoices = True

Exit Function
ErrLine:
    If Err Then
        MsgBox "ShowSalesInvoices(): " & vbCrLf & Err.Description
        Exit Function
    End If

End Function

'This function will list the Purchase Invoice Numbers of the
' Specified Company
'Inputs     HeadID  as long This is the CompanyID
Public Function ShowPurchaseInvoices(ByVal headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the varaibles
Dim rstInvoices As ADODB.Recordset

Dim InvoiceDate As String
Dim TransDate As String
Dim InvoiceNo As String
Dim InvoiceAmount As Currency

Dim CompanyName As String

Dim MaterialClass As clsMaterial

'Declare Fild variables to Store Field Values
Dim fldInvoiceNo As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldInvoiceDate As ADODB.Field
Dim fldInvoiceAmount As ADODB.Field


ShowPurchaseInvoices = False

If Not PrepareRstForPurchasesInvoices(rstInvoices, headID) Then Exit Function


Set fldInvoiceNo = rstInvoices.Fields("InvoiceNo")
Set fldTransDate = rstInvoices.Fields("TransDate")
Set fldInvoiceDate = rstInvoices.Fields("InvoiceDate")
Set fldInvoiceAmount = rstInvoices.Fields("InvoiceAmount")


Set MaterialClass = New clsMaterial

CompanyName = MaterialClass.GetCompanyName(headID)

Call SetCompanyTitle(1)

Call SetPurchaseInvoicesHeader(CompanyName)

Do While Not rstInvoices.EOF
    InvoiceDate = FormatField(rstInvoices.Fields("InvoiceDate"))
    TransDate = FormatField(rstInvoices.Fields("InvoiceDate"))
    InvoiceNo = rstInvoices.Fields("InvoiceNo").Value
    InvoiceAmount = rstInvoices.Fields("InvoiceAmount").Value
    
    Call PRINTPurchaseInvoiceNos(TransDate, InvoiceDate, InvoiceNo, InvoiceAmount)
    
        
    rstInvoices.MoveNext
Loop

Set MaterialClass = Nothing

ShowPurchaseInvoices = True

Exit Function
ErrLine:
    If Err Then
        MsgBox "Show Purchase Invoices(): " & vbCrLf & Err.Description
        Exit Function
    End If

End Function


Private Sub StartHtmlTable(ByVal TableID As String, ObjectToAppend, ByVal ColsCount As Integer)
Dim DOMElement

With m_WebDoc

    Set m_WebTable = .createElement("Table")
    
    With m_WebTable
        .Id = TableID
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FONTSIZE = 10
        .Style.fontFamily = gFontName
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = ColsCount
    End With
    
    Set DOMElement = .createElement("TBody")
    
    m_WebTable.appendChild DOMElement
    
    ObjectToAppend.appendChild m_WebTable
    
    m_WebTable.Style.fontFamily = gFontName
    m_WebTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
End With

End Sub





Private Function InsertRowAndCellSToTable(Table As HTMLTable, Optional CellsCount As Integer) As HTMLTableRow

Dim Cell As Integer
Dim Row As HTMLTableRow


With m_WebDoc

    'Set Row = Table.tBodies(0).insertRow(-1)
    
    Set Row = .createElement("TR")
    Row.Height = 18
    Table.tBodies(0).appendChild Row
    
    ' if user has not specified the cols then get the columns from the table
    If CellsCount = 0 Then CellsCount = Table.Cols
    
    For Cell = 0 To CellsCount - 1
    
        Set m_WebCell = .createElement("TD")
        Row.appendChild m_WebCell
        
        'm_WebCell.Style.borderLeftWidth = "0"
        ''m_WebCell.Style.borderRight = "none"
        'm_WebCell.Style.borderTopWidth = "0"
        'm_WebCell.Style.borderBottomWidth = "0"
        
    Next Cell
    
    m_WebCell.innerText = " "

End With

Set InsertRowAndCellSToTable = Row

End Function

''This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTGodownNames(ByVal GodownName As String)
'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 1)

m_WebRow.cells(0).colSpan = 5

'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 1)


With m_WebRow.cells(0)
    .innerText = GodownName
    .Align = "Center"
    .colSpan = 5
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 16
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
End With

'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 1)
m_WebRow.cells(0).colSpan = 5

End Sub







'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTBranchTotal(ByVal TotalValue As Currency)

'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
            
'Print the Total
With m_WebRow.cells(1)
    .innerText = "Branch Total"
    If gLangOffSet <> 0 Then _
        .innerText = GetResourceString(227, 52)
    
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(4)
    .innerText = FormatCurrency((TotalValue))   ' )
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub



'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowStockForGroup(ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim headID As Long 'It is ManufactuererID
Dim RelationID As Long
Dim SlNo As Integer

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String

Dim StockPurchase As Double
Dim StockSales As Double

Dim StockBalance As Double
Dim ValueBalance As Currency

Dim TotalValue As Currency

Dim ClosingStock As Currency
Dim PurchasePrice As Currency

Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldHeadID As ADODB.Field
'Dim fldUnitID As ADODB.Field
Dim fldUnitName As ADODB.Field


'Iniitialse the function
ShowStockForGroup = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForGroup(rstStock, GroupID, GodownID) Then Exit Function

Call LoadPurchasePrice(m_FromRpDate, m_ToRpDate, rstPrice)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGroupID = rstStock.Fields("GroupID")
Set fldGroupName = rstStock.Fields("GroupName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldCompanyName = rstStock.Fields("CompanyName")
Set fldHeadID = rstStock.Fields("HeadID")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0


Do
    If rstStock.EOF Then Exit Do
    
    'If company Changes then
    If headID <> fldHeadID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        
        headID = fldHeadID.Value
        headName = fldCompanyName.Value
        
        Call PRINTCompanyNames(headName)
    End If
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then _
                    PurchasePrice = rstPrice.Fields("UnitPrice")
                    
    ValueBalance = StockBalance * PurchasePrice
    
'    If UnitID <> fldUnitID.Value Then
'        UnitID = fldUnitID.Value
'        If TotalStock > 0 Then
'            'Print the total
'            Call PRINTTotal(TotalStock, TotalValue)
'
'            TotalStock = 0: TotalValue = 0
'        End If
'    End If
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Total Closing Stock
    ClosingStock = ClosingStock + ValueBalance
    
    'Move the recordset to Next Position
    rstStock.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotal(TotalValue)
    
    TotalValue = 0
End If
    
'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer


ShowStockForGroup = True
Exit Function

ErrLine:
    If Err Then
        MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function

'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowStockForGroupIncBranches(ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim headID As Long
Dim RelationID As Long
Dim SlNo As Integer
Dim GodownID As Byte

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String

Dim StockPurchase As Double
Dim StockSales As Double

Dim StockBalance As Double

Dim ValueBalance As Currency
Dim TotalValue As Currency
Dim ClosingStock As Currency
Dim PurchasePrice As Currency
Dim BranchTotal As Currency

Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldUnitName As ADODB.Field


'Iniitialse the function
ShowStockForGroupIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForGroupIncBranch(rstStock, GroupID) Then Exit Function

Call LoadPurchasePrice(m_FromRpDate, m_ToRpDate, rstPrice)

'Set the Header for the report
Call SetCompanyTitle

'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldCompanyName = rstStock.Fields("CompanyName")
Set fldHeadID = rstStock.Fields("HeadID")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0


Do
    If rstStock.EOF Then Exit Do
        
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        If BranchTotal > 0 Then
            'Print the total
            Call PRINTBranchTotal(BranchTotal)
            BranchTotal = 0
        End If
        
        GodownID = fldGodownID.Value
        GodownName = fldGodownName.Value
        
        Call PRINTGodownNames(GodownName)
    End If
    'If company Changes then
    If headID <> fldHeadID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        
        headID = fldHeadID.Value
        headName = fldCompanyName.Value
        
        Call PRINTCompanyNames(headName)
    End If
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then _
                    PurchasePrice = rstPrice.Fields("UnitPrice")
                    
    ValueBalance = StockBalance * PurchasePrice
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Total Closing Stock
    ClosingStock = ClosingStock + ValueBalance
    
    'Move the recordset to Next Position
    rstStock.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotal(TotalValue)
    
    TotalValue = 0
End If
    
'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer


ShowStockForGroupIncBranches = True
Exit Function

ErrLine:
    If Err Then
        MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function



'This Sub Routine will Set the Header Transactions
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetTransactionsHeader(Optional ByVal CompanyName As String, _
                                    Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim DOMElement

'Set the Header Name

With m_WebDoc
    Set DOMElement = .createElement("H3")
    .body.appendChild DOMElement
    
    strHeader = "Stock Tranasactions Between " & m_FromRpDate & " AND " & m_ToRpDate
    
    With DOMElement
        .innerText = strHeader
        .Align = "Center"
        .Style.FONTSIZE = 17
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set DOMElement = .createElement("H4")
        .body.appendChild DOMElement
        With DOMElement
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    'If group Name  exists then
    If GroupName <> "" Then
        Set DOMElement = .createElement("H5")
        .body.appendChild DOMElement
        With DOMElement
            .innerText = GroupName
            .Align = "Center"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    Call StartHtmlTable("StockSales", m_WebDoc.body, 5)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(0)
    .innerText = "Date"
    .Align = "Left"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "27.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "UnitPrice"
    .Align = "Right"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = "Purchase"
    .Align = "Right"
    .Width = "12.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
'With m_WebRow.cells(4)
'    .innerText = "FreePurchase"
'    .Align = "Right"
'    .Width = "12.5%"
'    .Style.fontWeight = "Bold"
'    .Style.FontSize = "x-Small"
'End With
With m_WebRow.cells(4)
    .innerText = "Sales"
    .Align = "Right"
    .Width = "12.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
'With m_WebRow.cells(6)
'    .innerText = "FreeSale"
'    .Align = "Right"
'    .Width = "12.5%"
'    .Style.fontWeight = "Bold"
'    '.Style.FontSize = "x-Small"
'End With
End Sub
''This Sub Routine will Set the Header For Purchases
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetPurchaseHeader(Optional ByVal CompanyName As String, _
                                Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow

With m_WebDoc

    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    
    strHeader = "Stock Purchase Between " & m_FromRpDate & " AND " & m_ToRpDate
    
    With CurrentRow.cells(0)
        .innerText = strHeader
        .Align = "Center"
        .Style.FONTSIZE = 17
        .Style.fontWeight = "Bold"
        .Style.fontFamily = gFontName
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        
        With CurrentRow.cells(0)
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            .Style.fontFamily = gFontName
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    'If group Name  exists then
    If GroupName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        
        With CurrentRow.cells(0)
            .innerText = GroupName
            .Align = "Center"
            .Style.FONTSIZE = 15
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    Call StartHtmlTable("StockPurchase", m_WebDoc.body, 7)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)

With m_WebRow.cells(0)
    .innerText = "Date"
    .Align = "Left"
    .Width = "7%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "36.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "Price"
    .Align = "Right"
    .Width = "4%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = "Qty Purchase"
    .Align = "Right"
    .Width = "10.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(4)
    .innerText = "P Value"
    .Align = "Right"
    .Width = "17.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(5)
    .innerText = "Free"
    .Align = "Right"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(6)
    .innerText = "Value"
    .Align = "Right"
    .Width = "15.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub
''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetPurchaseInvoicesHeader(Optional ByVal CompanyName As String)
'Declare the variables
Dim strHeader As String
Dim DOMElement

With m_WebDoc
    Set DOMElement = .createElement("H3")
    .body.appendChild DOMElement
    
    strHeader = "Invoice Numbers Between " & m_FromRpDate & " AND " & m_ToRpDate
    With DOMElement
        .innerText = strHeader
        .Align = "Center"
        .Style.FONTSIZE = 17
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set DOMElement = .createElement("H4")
        .body.appendChild DOMElement
        With DOMElement
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    Call StartHtmlTable("PurchaseInvoiceNos", m_WebDoc.body, 4)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebTable.Width = 500

With m_WebRow.cells(0)
    .innerText = "InvoiceDate"
    .Align = "Left"
    .Width = "22.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "TransDate"
    .Align = "Left"
    .Width = "22.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "Invoice No"
    .Align = "Left"
    .Width = "25%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = "Invoice Amount"
    .Align = "Right"
    .Width = "30%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub
''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetSalesInvoicesHeader(Optional ByVal CompanyName As String)
'Declare the variables
Dim strHeader As String
Dim DOMElement

With m_WebDoc
    Set DOMElement = .createElement("H3")
    .body.appendChild DOMElement
    
    strHeader = "Invoice Numbers Between " & m_FromRpDate & " AND " & m_ToRpDate
    With DOMElement
        .innerText = strHeader
        .Align = "Center"
        .Style.FONTSIZE = 17
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set DOMElement = .createElement("H4")
        .body.appendChild DOMElement
        With DOMElement
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    Call StartHtmlTable("PurchaseInvoiceNos", m_WebDoc.body, 3)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 3)

m_WebTable.Width = 500

With m_WebRow.cells(0)
    .innerText = "InvoiceDate"
    .Align = "Left"
    .Width = "22.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "Invoice No"
    .Align = "Left"
    .Width = "25%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "Invoice Amount"
    .Align = "Right"
    .Width = "30%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub

''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetSalesHeader(Optional ByVal CompanyName As String, _
                                    Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow


With m_WebDoc
    
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    
    strHeader = "Stock Sales Between " & m_FromRpDate & " AND " & m_ToRpDate
    
    With CurrentRow.cells(0)
        .innerText = strHeader
        .Align = "Center"
        .Style.FONTSIZE = 17
        .Style.fontWeight = "Bold"
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        
        With CurrentRow.cells(0)
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    'If group Name  exists then
    If GroupName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        
        With CurrentRow.cells(0)
            .innerText = GroupName
            .Align = "Center"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    Call StartHtmlTable("StockSales", m_WebDoc.body, 7)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 7)

With m_WebRow.cells(0)
    .innerText = "Date"
    .Align = "Left"
    .Width = "7%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "36.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "Price"
    .Align = "Right"
    .Width = "4%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = "Qty Sales"
    .Align = "Right"
    .Width = "10.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(4)
    .innerText = "S Value"
    .Align = "Right"
    .Width = "17.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(5)
    .innerText = "Free"
    .Align = "Right"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(6)
    .innerText = "Value"
    .Align = "Right"
    .Width = "15.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
End Sub

''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Public Sub SetStockBalanceHeader(Optional ByVal CompanyName As String, _
                                  Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim MetElement As HTMLMetaElement
Dim CurrentRow As HTMLTableRow

With m_WebDoc
    strHeader = "Stock As On  " & m_ToRpDate 'GetResourceString(175) & "    " & m_ToRpDate    '
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    
    With CurrentRow.cells(0)
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = "17"
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    
    'If Company  Name  exists then
    If CompanyName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        With CurrentRow.cells(0)
            .innerText = CompanyName
            .Align = "Center"
            .Style.FONTSIZE = 16
            .Style.fontWeight = "Bold"
            .Style.fontFamily = gFontName
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
        End With
    End If
    
    'If group Name  exists then
    If GroupName <> "" Then
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
        With CurrentRow.cells(0)
            .innerText = GroupName
            .Align = "Center"
            .Style.FONTSIZE = 15
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    Call StartHtmlTable("StockBalance", m_WebDoc.body, 5)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(0)
    .innerText = "SlNo"
    .Align = "Left"
    .Width = "6%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(1)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "41.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "UnitPrice"
    .Align = "Right"
    .Width = "9.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(3)
    .innerText = "Stock Balance"
    .Align = "Right"
    .Width = "16%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(4)
    .innerText = "Value Balance"
    .Align = "Right"
    .Width = "22%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub
Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property

Public Property Let FromRpDate(ByVal NewValue As String)
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
    m_FromUSDate = GetSysFormatDate(NewValue)
    'm_FromRpDate = GetIndianDate(m_FromUSDate)
End Property

'This will Show the Web Browser form
Public Sub ShowCurrentReport()

'Set m_frmWeb.m_HtmlToNavigate = m_WebDoc

'm_frmWeb.web.Navigate m_frmWeb.m_HtmlToNavigate

m_frmWeb.Show vbModal

End Sub


'This function will show the report date form
Public Function ShowReportDate(ByVal StDateLoacked As Boolean) As Boolean

If m_frmRptDt Is Nothing Then Set m_frmRptDt = New frmRptDt
With m_frmRptDt
    .txtStDate = FinIndianFromDate
    .txtEndDate = FinIndianEndDate
    
    .txtStDate.Locked = False
    If StDateLoacked Then .txtStDate.Locked = True
    
    .Show vbModal
    
End With
End Function

Public Property Get ObjectName()
    ObjectName = "StockReportClass"
End Property


Public Property Get ToRpDate() As String
    ToRpDate = m_ToRpDate
End Property


Public Property Let ToRpDate(ByVal NewValue As String)
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_ToRpDate = NewValue
    m_ToUSDate = GetSysFormatDate(NewValue)
    'm_ToRpDate = GetIndianDate(m_ToUSDate)
End Property


Private Sub m_frmRptDt_CancelClick()
Unload m_frmRptDt

Set m_frmRptDt = Nothing

End Sub

Private Sub m_frmRptDt_OKClick(StDate As String, EndDate As String)
m_FromRpDate = StDate
m_ToRpDate = EndDate

Unload m_frmRptDt
Set m_frmRptDt = Nothing

End Sub






