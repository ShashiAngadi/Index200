VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStockRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_frmRepSlect As frmRepSelect

Private m_FromRpDate As String
Private m_ToRpDate As String
Private m_FromUSDate As Date
Private m_ToUSDate As Date


Private WithEvents m_frmRptDt As frmRptDt
Attribute m_frmRptDt.VB_VarHelpID = -1
Private WithEvents m_WebDoc As HTMLDocument
Attribute m_WebDoc.VB_VarHelpID = -1

Private m_frmWeb As frmWeb

Private m_WebTable As HTMLTable
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell

'This function returns the Record set
'which includes the Product List & and their stock
'as on the end date of FinYear
Public Function GetYearEndStockDetails(GodownID As Byte) As Recordset
    
    FromRpDate = FinIndianFromDate
    ToRpDate = FinIndianEndDate
    
    
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset
Dim rstOpStock As ADODB.Recordset

Dim USAsOnDate As String
Dim headID As Long
Dim GroupID As Long
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
Dim OpStockValue As Currency
Dim OpStock As Double

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
Dim fldUnitName As ADODB.Field


'Iniitialse the function
'ShowStockAsOnDate = False

'Call StartTimer

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Prepare the recordset
If Not PrepareRecordsetForTotalStock(rstStock, GodownID) Then Exit Function

Call CreateOpStockQuery(GodownID, m_FromRpDate, "qryOpStock")
gDbTrans.SqlStmt = "SELECT OpStock,RelationID FROM QryOpStock"
Call gDbTrans.Fetch(rstOpStock, adOpenForwardOnly)

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * from qryItemPrice"

If gDbTrans.Fetch(rstPrice, adOpenDynamic) > 0 Then
    Set GetYearEndStockDetails = rstPrice
Else
    Set GetYearEndStockDetails = Nothing
End If

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGroupID = rstStock.Fields("GroupID")
Set fldGroupName = rstStock.Fields("GroupName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0
Set MaterialClass = New clsMaterial

Do
    If rstStock.EOF Then Exit Do
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    OpStock = 0
    OpStockValue = OpStock * PurchasePrice
    
    SlNo = SlNo + 1
    TotalValue = TotalValue + ValueBalance + OpStockValue
    
    'Total Closing Stock
    ClosingStock = ClosingStock + ValueBalance + OpStockValue
    
    'Move the recordset to Next Position
    rstStock.MoveNext
Loop


Exit Function

ErrLine:
    If Err Then
        'Resume
        MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If
End Function

''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetStockSummaryHeader(Optional ByVal CompanyName As String, _
                                  Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow

With m_WebDoc
    strHeader = "Stock Summary From " & m_FromRpDate & " To " & m_ToRpDate   'GetResourceString(137) & "    " & m_ToRpDate    '
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
    
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
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
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
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
        With CurrentRow.cells(0)
            .innerText = GroupName
            .Align = "Center"
            .Style.FONTSIZE = 15
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    Call StartHtmlTable("StocSummary", m_WebDoc.body, 15)
         
    m_WebTable.Width = 950
    
    
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 3)

m_WebRow.cells(0).colSpan = 3
m_WebRow.cells(1).colSpan = 6
m_WebRow.cells(2).colSpan = 6

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(1)
    .innerText = "PHYSICAL"
    .Align = "Center"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(2)
    .innerText = "FINANCIAL"
    .Align = "Center"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)


With m_WebRow.cells(0)
    .innerText = "Particulars"
    .Align = "Left"
    .Width = "18%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(1)
    .innerText = "Pack"
    .Align = "Left"
    .Width = "5.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(2)
    .innerText = "Rate"
    .Align = "Right"
    .Width = "4%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(3)
    .innerText = "O_Stk"
    .Align = "Right"
    .Width = "6%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
With m_WebRow.cells(4)
    .innerText = "Pur  "
    .Align = "Right"
    .Width = "4%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
With m_WebRow.cells(5)
    .innerText = "Free"
    .Align = "Right"
    .Width = "3%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
With m_WebRow.cells(6)
    .innerText = "Sales"
    .Align = "Right"
    .Width = "4%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
With m_WebRow.cells(7)
    .innerText = "Free"
    .Align = "Right"
    .Width = "3%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(8)
    .innerText = "C_Stk"
    .Align = "Right"
    .Width = "6%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(9)
    .innerText = "OpStock"
    .Align = "Right"
    .Width = "8%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
With m_WebRow.cells(10)
    .innerText = "Pur  "
    .Align = "Right"
    .Width = "8%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(11)
    .innerText = "Free"
    .Align = "Right"
    .Width = "6%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(12)
    .innerText = "Sales"
    .Align = "Right"
    .Width = "8%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(13)
    .innerText = "Free"
    .Align = "Right"
    .Width = "6%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

With m_WebRow.cells(14)
    .innerText = "OnHand"
    .Align = "Right"
    .Width = "8%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With

End Sub



Private Function CreatePurchaseQuery(rstPurchase As ADODB.Recordset, ByVal GodownID As Byte, ByVal headID As Long, _
                                ByVal FromDateUS As String, ByVal ToDateUS As String) As Boolean


'Purchase = 3,RejectionsIn = 8,StockIn = 11
'FreePurchase = 5,FreeRejectionsIN = 16
'Sales = 4,RejectionsOut = 9,StockOut = 12,StockSoot = 13
'FreeSales = 10,FreeRejectionsOUT = 17

'gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS PurchaseQty,SUM(Quantity*UnitPrice) AS PurValue, B.RelationID, SUM(A.TaxAmount) as PurTaxAmount"
gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS PurchaseQty,Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) AS PurValue, B.RelationID, SUM(A.TaxAmount) as PurTaxAmount" & _
                   " FROM Stock A,RelationMaster B" & _
                   " WHERE A.GodownID=" & GodownID & _
                   " AND B.RelationID=A.RelationID" & _
                   " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                   " AND #" & ToDateUS & "#" & _
                   " AND PurORSaleID >0 AND ( VoucherType =" & Purchase & _
                         " OR VoucherType =" & RejectionsIn & _
                         " OR VoucherType =" & StockIn & " )" & _
                   " GROUP BY B.RelationID"

If headID > 0 Then _
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS PurchaseQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  AS PurValue, SUM(A.TaxAmount) as PurTaxAmount, B.RelationID" & _
                       " FROM Stock A,RelationMaster B" & _
                       " WHERE A.GodownID=" & GodownID & _
                       " AND B.RelationID=A.RelationID" & _
                       " AND B.HeadID=" & headID & _
                       " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                       " AND #" & ToDateUS & "#" & _
                       " AND PurORSaleID >0 AND ( VoucherType =" & Purchase & _
                             " OR VoucherType =" & RejectionsIn & _
                             " OR VoucherType =" & StockIn & " )" & _
                       " GROUP BY B.RelationID"

gDbTrans.CreateView ("QryPurchase")


gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS FreePurchaseQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  AS FreePurValue, SUM(A.TaxAmount) as FreeTaxAmount, B.RelationID" & _
                   " FROM Stock A,RelationMaster B" & _
                   " WHERE A.GodownID=" & GodownID & _
                   " AND B.RelationID=A.RelationID" & _
                   " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                   " AND #" & ToDateUS & "#" & _
                   " AND ( VoucherType =" & FreePurchase & _
                         " OR VoucherType =" & FreeRejectionsIN & " )" & _
                   " GROUP BY B.RelationID"

If headID > 0 Then _
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS FreePurchaseQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  AS FreePurValue, SUM(A.TaxAmount) as FreeTaxAmount,B.RelationID" & _
                       " FROM Stock A,RelationMaster B" & _
                       " WHERE A.GodownID=" & GodownID & _
                       " AND B.RelationID=A.RelationID" & _
                       " AND B.HeadID=" & headID & _
                       " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                       " AND #" & ToDateUS & "#" & _
                       " AND ( VoucherType =" & FreePurchase & _
                             " OR VoucherType =" & FreeRejectionsIN & " )" & _
                       " GROUP BY B.RelationID"

gDbTrans.CreateView ("QryFreePurchase")

gDbTrans.SqlStmt = " SELECT A.*,iif( B.FreePurchaseQty is null,0,B.FreePurchaseQty) AS FreePurchase," & _
                        "iif(B.FreePurValue IS NULL,0,B.FreePurValue) AS FreePurValue," & _
                        "iif(B.FreeTaxAmount IS NULL,0,B.FreeTaxAmount) AS FreeTaxAmount" & _
                   " FROM QryPurchase A LEFT JOIN QryFreePurchase B" & _
                   " ON A.RelationID= B.RelationID"

Call gDbTrans.Fetch(rstPurchase, adOpenForwardOnly)


CreatePurchaseQuery = True


End Function

Private Function CreateSalesQuery(rstSales As ADODB.Recordset, ByVal GodownID As Byte, ByVal headID As Long, _
                                ByVal FromDateUS As String, ByVal ToDateUS As String, _
                                ByVal qryName As String, ByVal FreeQryName As String) As Boolean


'Purchase = 3,RejectionsIn = 8,StockIn = 11
'FreePurchase = 5,FreeRejectionsIN = 16
'Sales = 4,RejectionsOut = 9,StockOut = 12,StockSoot = 13
'FreeSales = 10,FreeRejectionsOUT = 17

gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS SalesQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) AS SaleValue,B.RelationID" & _
                   " FROM Stock A,RelationMaster B" & _
                   " WHERE A.GodownID=" & GodownID & _
                   " AND B.RelationID=A.RelationID" & _
                   " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                   " AND #" & ToDateUS & "#" & _
                   " AND ( VoucherType =" & Sales & _
                         " OR VoucherType =" & RejectionsOut & _
                         " OR VoucherType =" & StockSoot & _
                         " OR VoucherType =" & StockOut & " )" & _
                   " GROUP BY B.RelationID"


If headID > 0 Then _
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS SalesQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  AS SaleValue,B.RelationID" & _
                       " FROM Stock A,RelationMaster B" & _
                       " WHERE A.GodownID=" & GodownID & _
                       " AND B.RelationID=A.RelationID" & _
                       " AND B.HeadID=" & headID & _
                       " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                       " AND #" & ToDateUS & "#" & _
                       " AND ( VoucherType =" & Sales & _
                         " OR VoucherType =" & RejectionsOut & _
                         " OR VoucherType =" & StockSoot & _
                         " OR VoucherType =" & StockOut & " )" & _
                       " GROUP BY B.RelationID"

gDbTrans.CreateView (qryName)


gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS FreeSalesQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) AS FreeSaleValue,B.RelationID" & _
                   " FROM Stock A,RelationMaster B" & _
                   " WHERE A.GodownID=" & GodownID & _
                   " AND B.RelationID=A.RelationID" & _
                   " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                   " AND #" & ToDateUS & "#" & _
                   " AND ( VoucherType =" & FreeSales & _
                         " OR VoucherType =" & FreeRejectionsOUT & " )" & _
                   " GROUP BY B.RelationID"

If headID > 0 Then _
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS FreeSalesQty,SUM (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  AS FreeSaleValue,B.RelationID" & _
                       " FROM Stock A,RelationMaster B" & _
                       " WHERE A.GodownID=" & GodownID & _
                       " AND B.RelationID=A.RelationID" & _
                       " AND B.HeadID=" & headID & _
                       " AND TransDate BETWEEN #" & FromDateUS & "#" & _
                       " AND #" & ToDateUS & "#" & _
                       " AND ( VoucherType =" & FreeSales & _
                             " OR VoucherType =" & FreeRejectionsOUT & " )" & _
                       " GROUP BY B.RelationID"

gDbTrans.CreateView (FreeQryName)

gDbTrans.SqlStmt = " SELECT A.*,iif( B.FreeSalesQty is null,0,B.FreeSalesQty)AS FreeSales" & _
                        ",iif( B.FreeSaleValue is null,0,B.FreeSaleValue)AS FreeSaleValue" & _
                   " FROM " & qryName & " A LEFT JOIN " & FreeQryName & " B" & _
                   " ON A.RelationID= B.RelationID"
                   
Call gDbTrans.Fetch(rstSales, adOpenForwardOnly)

CreateSalesQuery = True

End Function


' This function gives the purchase price of the item from the database
' Inputs : ItemID as long
' Output : Gives the Purchase Price of the Item
Private Sub LoadPurchasePriceForRelationID(ByVal fromDate As String, _
                                 ByVal toDate As String, _
                                 rstPrice As ADODB.Recordset)
'Trap an error
On Error GoTo ErrLine
'declare Variables

Dim VoucherType As Wis_VoucherTypes
Dim eVoucherType As Wis_VoucherTypes

VoucherType = Purchase
eVoucherType = StockIn

gDbTrans.SqlStmt = " SELECT MAX(TradingPrice) as UnitPrice,RelationID" & _
                 " FROM RelationMaster" & _
                 " GROUP BY RelationID"
                 
If gDbTrans.Fetch(rstPrice, adOpenDynamic) < 0 Then Exit Sub

Exit Sub

ErrLine:
    
End Sub

Public Property Get ObjectName()
    ObjectName = "StockRepClass"
End Property



Private Function CreateOpStockQuery(ByVal GodownID As Byte, _
            ByVal USAsOnDate As String, ByVal qryName As String) As Boolean
'Declare the variables
Dim USFromDate As String

Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

CreateOpStockQuery = False

USFromDate = FinUSFromDate 'FormatDate(FinFromDate)

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

If USAsOnDate = USFromDate Then
    gDbTrans.SqlStmt = " SELECT SUM(Quantity) as OpStock, A.RelationID," & _
                       " Sum (IIf(IsNull(A.Amount), (A.UnitPrice * A.Quantity), A.Amount)) as opStockAmount" & _
                       " FROM Stock A,RelationMaster C" & _
                       " WHERE A.GodownID = " & GodownID & _
                       " AND A.RelationID=C.RelationID" & _
                       " AND TransDate =#" & USAsOnDate & "#" & _
                       " AND PurOrSaleID=0" & _
                       " GROUP BY A.RelationID"
    
    gDbTrans.CreateView (qryName)
    
    Exit Function
    
Else
    USAsOnDate = DateAdd("d", -1, CDate(USAsOnDate))

End If

gDbTrans.SqlStmt = " SELECT SUM(Quantity) AS StockPurchase,A.RelationId," & _
                   "Sum(IIf(IsNull(A.Amount), (A.UnitPrice * A.Quantity), A.Amount)) as StockPurchaseAmount" & _
                   " FROM Stock A,RelationMaster C" & _
                   " WHERE A.GodownID = " & GodownID & _
                   " AND A.RelationID=C.RelationID" & _
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


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales, A.RelationID," & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) as StockSalesAmount" & _
                   " FROM Stock A,RelationMaster C" & _
                   " WHERE A.GodownID = " & GodownID & _
                   " AND A.RelationID=C.RelationID" & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & FreeRejectionsOUT & _
                   " OR VoucherType= " & eRejectionVoucher & " )" & _
                   " GROUP BY A.RelationID"

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,iif( B.StockSales is null,0,B.StockSales) AS StockSales," & _
                   " iif(B.StockSalesAmount is null,0,B.StockSalesAmount) as StockSalesAmount  " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B" & _
                   " ON A.RelationID = B.RelationID"

Call gDbTrans.CreateView("QryStockMerge")

gDbTrans.SqlStmt = "SELECT (StockPurchase-StockSales) AS OpStock, " & _
                   " (StockPurchaseAmount-StockSalesAmount) AS OpStockAmount,RelationID" & _
                   " FROM QryStockMerge" & _
                   " WHERE (StockPurchase-StockSales) > 0"
                   
Call gDbTrans.CreateView(qryName)
                   
CreateOpStockQuery = True

End Function

Public Function ShowStockSummaryReports() As Boolean
'Trap an error
On Error GoTo ErrLine

'Delclare the variables
Dim GodownID As Byte
Dim headID As Long
Dim headName As String

ShowStockSummaryReports = False

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Set the Web Doc
Call SetWebDocument

With m_frmRepSlect
    If .cmbGodown.ListIndex = -1 Then Exit Function
    GodownID = .cmbGodown.ItemData(.cmbGodown.ListIndex)
    
'    If .cmbManufacturer.ListIndex = -1 Then Exit Function
'    HeadID = .cmbManufacturer.ItemData(.cmbManufacturer.ListIndex)
'    HeadName = .cmbManufacturer.Text
End With

If GodownID = 0 Then Exit Function
    
If Not ShowStockSummary(GodownID) Then Exit Function


Screen.MousePointer = vbDefault

Call ShowCurrentReport

ShowStockSummaryReports = True

Exit Function

ErrLine:
    MsgBox "ShowStockSummaryReports " & vbCrLf & Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Function


''This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTHeadNames(ByVal headName As String)
'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 1)

With m_WebRow.cells(0)
    .innerText = headName
    .Align = "Center"
    m_WebRow.cells(0).colSpan = 15
    
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 16
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
End With

'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 1)

'm_WebRow.cells(0).colSpan = 15


End Sub

Private Function ShowStockSummary(ByVal GodownID As Byte) As Boolean

Dim rstItemNames As ADODB.Recordset
Dim rstPrice As ADODB.Recordset
Dim rstOpStock As ADODB.Recordset
Dim rstPurchase As ADODB.Recordset
Dim rstSales As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function


'frmCancel.Show
Call CreateOpStockQuery(GodownID, m_FromUSDate, "QryOpStock")

gDbTrans.SqlStmt = "SELECT OpStock,RelationID,opStockAmount FROM QryOpStock"

Call gDbTrans.Fetch(rstOpStock, adOpenForwardOnly)

Call CreatePurchaseQuery(rstPurchase, GodownID, 0, m_FromUSDate, m_ToUSDate)

Call CreateSalesQuery(rstSales, GodownID, 0, m_FromUSDate, m_ToUSDate, "QrySales", "QryFreeSales")

Call PrepareItemNameRst(rstItemNames, GodownID)

Call LoadPurchasePriceForRelationID(m_FromRpDate, m_ToRpDate, rstPrice)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

'Set stock balance Header
Call SetStockSummaryHeader

'Declare the variables for Displaying on Web Browser
Dim headID As Long
Dim RelationID As Long

Dim headName As String
Dim ItemName As String
Dim UnitName As String

Dim PurPrice As Currency

Dim ClValue As Currency
Dim OpStockValue As Currency
Dim PurValue As Currency
Dim FreePurValue As Currency
Dim SaleValue As Currency
Dim FreeSaleValue As Currency

Dim OpStock As Double

Dim TotalValue As Currency
Dim TotalOpStockValue
Dim TotalPurValue As Currency
Dim TotalFreePurValue As Currency
Dim TotalSaleValue As Currency
Dim TotalFreeSaleValue As Currency


Dim Total As Double
Dim TotalOpStock As Double
Dim TotalPur As Double
Dim TotalFreePur As Double
Dim TotalSale As Double
Dim TotalFreeSale As Double

Dim VendorTotalValue As Currency
Dim VendorTotalOpStockValue
Dim VendorTotalPurValue As Currency
Dim VendorTotalFreePurValue As Currency
Dim VendorTotalSaleValue As Currency
Dim VendorTotalFreeSaleValue As Currency

Dim VendorTotalOpStock As Double
Dim VendorTotalPur As Double
Dim VendorTotalFreePur As Double
Dim VendorTotalSale As Double
Dim VendorTotalFreeSale As Double
Dim VendorTotal As Double



'Dim OpStock As Double
Dim ClStock As Double
Dim PurQty As Single
Dim FreePurQty As Single
Dim SaleQty As Single
Dim FreeSaleQty As Single
    DoEvents
    gCancel = 0
    With frmCancel
        .PicStatus.Visible = True
        .lblMessage = "Preparing the Report"
    End With
    

Do While Not rstItemNames.EOF
    RelationID = rstItemNames.Fields("RelationID")
    ItemName = rstItemNames.Fields("ProductName")
    UnitName = rstItemNames.Fields("UnitName")
    
    OpStock = 0: PurPrice = 0: OpStockValue = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID =" & RelationID) Then PurPrice = FormatField(rstPrice.Fields("UnitPrice"))
    If gDbTrans.FindRecord(rstOpStock, "RelationID =" & RelationID) Then
        OpStock = rstOpStock.Fields("OpStock")
        OpStockValue = FormatField(rstOpStock("opStockAmount"))
        If OpStockValue = 0 Then OpStockValue = OpStock * PurPrice
    End If
    
    PurQty = 0: FreePurQty = 0: PurValue = 0: FreePurValue = 0
    
    If gDbTrans.FindRecord(rstPurchase, "RelationID=" & RelationID) Then
        PurQty = rstPurchase.Fields("PurchaseQty")
        FreePurQty = rstPurchase.Fields("FreePurchase")
        PurValue = rstPurchase.Fields("PurValue")
        FreePurValue = rstPurchase.Fields("FreePurValue")
    End If
    
    SaleQty = 0: FreeSaleQty = 0: SaleValue = 0: FreeSaleValue = 0
    If gDbTrans.FindRecord(rstSales, "RelationID=" & RelationID) Then
        SaleQty = rstSales.Fields("SalesQty")
        FreeSaleQty = rstSales.Fields("FreeSales")
        SaleValue = rstSales.Fields("SaleValue")
        FreeSaleValue = rstSales.Fields("FreeSaleValue")
    End If
    
    ClStock = Format(OpStock + PurQty + FreePurQty - SaleQty - FreeSaleQty, "#.00")
    ClValue = ClStock * PurPrice
    
    VendorTotalValue = VendorTotalValue + ClValue
    VendorTotalOpStockValue = VendorTotalOpStockValue + OpStockValue
    VendorTotalPurValue = VendorTotalPurValue + PurValue
    VendorTotalFreePurValue = VendorTotalFreePurValue + FreePurValue
    VendorTotalSaleValue = VendorTotalSaleValue + SaleValue
    VendorTotalFreeSaleValue = VendorTotalFreeSaleValue + FreeSaleValue:
    VendorTotal = VendorTotal + ClStock
    VendorTotalOpStock = VendorTotalOpStock + OpStock
    VendorTotalPur = VendorTotalPur + PurQty
    VendorTotalFreePur = VendorTotalFreePur + FreePurQty
    VendorTotalSale = VendorTotalSale + SaleQty
    VendorTotalFreeSale = VendorTotalFreeSale + FreeSaleQty
    
    TotalValue = TotalValue + ClValue
    TotalOpStockValue = TotalOpStockValue + OpStockValue
    TotalPurValue = TotalPurValue + PurValue
    TotalFreePurValue = TotalFreePurValue + FreePurValue
    TotalSaleValue = TotalSaleValue + SaleValue
    TotalFreeSaleValue = TotalFreeSaleValue + FreeSaleValue:
    Total = Total + ClStock
    TotalOpStock = TotalOpStock + OpStock
    TotalPur = TotalPur + PurQty
    TotalFreePur = TotalFreePur + FreePurQty
    TotalSale = TotalSale + SaleQty
    TotalFreeSale = TotalFreeSale + FreeSaleQty
    
    
    Call PRINTStockSummary(ItemName, UnitName, PurPrice, OpStock, PurQty, FreePurQty, SaleQty, FreeSaleQty, ClStock, OpStockValue, PurValue, FreePurValue, SaleValue, FreeSaleValue, ClValue)
    
    DoEvents
    If gCancel Then Exit Do
    With frmCancel
        UpdateStatus .PicStatus, rstItemNames.AbsolutePosition / rstItemNames.RecordCount
        .lblMessage = "Fetching record"
    End With
    'Move the recordset
    rstItemNames.MoveNext
Loop


UpdateStatus frmCancel.PicStatus, 1, True

If VendorTotalValue > 0 Then Call PRINTStockSummaryTotal( _
            VendorTotalOpStockValue, VendorTotalPurValue, VendorTotalFreePurValue, _
            VendorTotalSaleValue, VendorTotalFreeSaleValue, VendorTotalValue, _
            VendorTotalOpStock, VendorTotalPur, VendorTotalFreePur, _
            VendorTotalSale, VendorTotalFreeSale, VendorTotal)

Call PRINTStockSummaryTotal(TotalOpStockValue, TotalPurValue, TotalFreePurValue, _
        TotalSaleValue, TotalFreeSaleValue, TotalValue, _
        TotalOpStock, TotalPur, TotalFreePur, _
        TotalSale, TotalFreeSale, Total)

Unload frmCancel
ShowStockSummary = True

Exit Function

ErrLine:
        MsgBox "ShowStockSummary " & Err.Description
        'Resume

End Function


''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub PRINTStockSummary(ByVal ItemName As String, ByVal UnitName As String, _
                             ByVal UnitPrice As Currency, ByVal OpStock As Single, _
                             ByVal PurQty As Single, ByVal FreePurQty As Single, _
                             ByVal SalesQty As Single, ByVal FreeSalesQty As Single, _
                             ByVal ClStock As Double, ByVal OpStockValue As Currency, _
                             ByVal PurValue As Currency, ByVal FreePur As Currency, _
                             ByVal SaleValue As Currency, ByVal FreeSale As Currency, _
                             ByVal OnHand As Currency)




Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)

With m_WebRow.cells(0)
    .innerText = ItemName
    .Align = "Left"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(1)
    .innerText = UnitName
    .Align = "Left"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(2)
    .innerText = FormatCurrency(UnitPrice)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(3)
    .innerText = OpStock
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(4)
    .innerText = PurQty
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(5)
    .innerText = FreePurQty
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(6)
    .innerText = SalesQty
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(7)
    .innerText = FreeSalesQty
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(8)
    .innerText = ClStock
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(9)
    .innerText = FormatCurrency(OpStockValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(10)
    .innerText = FormatCurrency(PurValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(11)
    .innerText = FormatCurrency(FreePur)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(12)
    .innerText = FormatCurrency(SaleValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(13)
    .innerText = FormatCurrency(FreeSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(14)
    .innerText = FormatCurrency(OnHand)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
End Sub

''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub PRINTStockSummaryTotal(ByVal TotalOpStockValue As Currency, _
                             ByVal TotalPurValue As Currency, ByVal TotalFreePurValue As Currency, _
                             ByVal TotalSaleValue As Currency, ByVal TotalFreeSaleValue As Currency, _
                             ByVal OnHandValue As Currency, ByVal TotalOpStock As Double, _
                             ByVal TotalPur As Double, ByVal TotalFreePur As Double, _
                             ByVal TotalSale As Double, ByVal TotalFreeSale As Double, _
                             ByVal OnHand As Double)


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = "Total"
    .Align = "Left"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(3)
    .innerText = FormatCurrency(TotalOpStock)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(4)
    .innerText = FormatCurrency(TotalPur)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(5)
    .innerText = FormatCurrency(TotalFreePur)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(6)
    .innerText = FormatCurrency(TotalSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(7)
    .innerText = FormatCurrency(TotalFreeSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(8)
    .innerText = FormatCurrency(OnHand)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With



With m_WebRow.cells(9)
    .innerText = FormatCurrency(TotalOpStockValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(10)
    .innerText = FormatCurrency(TotalPurValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(11)
    .innerText = FormatCurrency(TotalFreePurValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(12)
    .innerText = FormatCurrency(TotalSaleValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(13)
    .innerText = FormatCurrency(TotalFreeSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(14)
    .innerText = FormatCurrency(OnHandValue)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

End Sub

''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub PRINTStockSummaryTotalQuantity(ByVal TotalOpStock As Double, _
                             ByVal TotalPur As Double, ByVal TotalFreePur As Double, _
                             ByVal TotalSale As Double, ByVal TotalFreeSale As Double, _
                             ByVal OnHand As Double)




Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 15)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = "Total"
    .Align = "Left"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(9)
    .innerText = FormatCurrency(TotalOpStock)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(10)
    .innerText = FormatCurrency(TotalPur)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(11)
    .innerText = FormatCurrency(TotalFreePur)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(12)
    .innerText = FormatCurrency(TotalSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
With m_WebRow.cells(13)
    .innerText = FormatCurrency(TotalFreeSale)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With

With m_WebRow.cells(14)
    .innerText = FormatCurrency(OnHand)
    .Align = "Right"
    .Style.FONTSIZE = 12
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 14
End With
End Sub



Private Function PrepareItemNameRst(rstItemNames As ADODB.Recordset, ByVal GodownID As Integer) As Boolean

gDbTrans.SqlStmt = " SELECT DISTINCT ProductName,UnitName,A.RelationID" & _
                  " FROM RelationMaster A,Products D,Units E" & _
                  " WHERE " & _
                  " A.ProductID=D.ProductID" & _
                  " AND A.GroupID=D.GroupID" & _
                  " AND A.UnitID=E.UnitID"
If GodownID > 0 Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
                    " AND A.GodownID = " & GodownID

End If
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY ProductName"


If gDbTrans.Fetch(rstItemNames, adOpenForwardOnly) < 1 Then Exit Function


PrepareItemNameRst = True

End Function



Private Sub LoadFrmWeb()

If Not m_frmWeb Is Nothing Then Exit Sub

Set m_frmWeb = New frmWeb
Load m_frmWeb

End Sub
'This function will list the Purchase Invoice Numbers of the
' Specified Company
'Inputs     HeadID  as long This is the CompanyID
Public Function ShowSalesInvoices(ByVal GodownID As Byte) As Boolean
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


'Show the report Date to be selected
'Call ShowReportDate(True)

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Set the Web Doc
Call SetWebDocument


ShowSalesInvoices = False

If Not PrepareRstForSalesInvoices(rstInvoices, GodownID) Then Exit Function

Set fldInvoiceNo = rstInvoices.Fields("InvoiceNo")
Set fldInvoiceDate = rstInvoices.Fields("InvoiceDate")
Set fldInvoiceAmount = rstInvoices.Fields("InvoiceAmount")


Set MaterialClass = New clsMaterial

'CompanyName = MaterialClass.GetCompanyName(HeadID)

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

Screen.MousePointer = vbDefault

Call ShowCurrentReport

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesInvoices(): " & vbCrLf & Err.Description
        

End Function
'This function will list the Purchase Invoice Numbers of the
' Specified Company
'Inputs     HeadID  as long This is the CompanyID
Public Function ShowPurchaseInvoices() As Boolean
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

'Show the report Date to be selected
'Call ShowReportDate( True)

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Set the Web Doc
Call SetWebDocument

ShowPurchaseInvoices = False

If Not PrepareRstForPurchasesInvoices(rstInvoices) Then Exit Function


Set fldInvoiceNo = rstInvoices.Fields("InvoiceNo")
Set fldTransDate = rstInvoices.Fields("TransDate")
Set fldInvoiceDate = rstInvoices.Fields("InvoiceDate")
Set fldInvoiceAmount = rstInvoices.Fields("InvoiceAmount")


Set MaterialClass = New clsMaterial

'CompanyName = MaterialClass.GetCompanyName(HeadID)

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

Screen.MousePointer = vbDefault

Call ShowCurrentReport

Exit Function

ErrLine:
    If Err Then MsgBox "Show Purchase Invoices(): " & vbCrLf & Err.Description
        

End Function

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



''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetPurchaseInvoicesHeader(Optional ByVal CompanyName As String)
'Declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow

With m_WebDoc
    strHeader = "Invoice Numbers Between " & m_FromRpDate & " AND " & m_ToRpDate
    
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    
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
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 18
            .Style.fontWeight = "Bold"
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

Private Function PrepareRstForPurchasesInvoices( _
        ByRef rstInvoices As ADODB.Recordset) As Boolean


PrepareRstForPurchasesInvoices = False

gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,TransDate,InvoiceAmount " & _
                   " FROM Purchase " & _
                   " WHERE TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " ORDER BY InvoiceDate,TransDate "

If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

PrepareRstForPurchasesInvoices = True

End Function



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




''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetSalesInvoicesHeader(Optional ByVal CompanyName As String)
'Declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow


With m_WebDoc
    
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl"))
    strHeader = "Invoice Numbers Between " & m_FromRpDate & " AND " & m_ToRpDate
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


Private Function PrepareRstForSalesInvoices(ByRef rstInvoices As ADODB.Recordset, _
                                        ByVal GodownID As Byte) As Boolean

PrepareRstForSalesInvoices = False

gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,InvoiceAmount " & _
                   " FROM Sales " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND InvoiceDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " ORDER BY InvoiceDate"
 
If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

PrepareRstForSalesInvoices = True

End Function

Private Function CreateStockMergeQuery(ByVal GodownID As Byte) As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

CreateStockMergeQuery = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

Dim sqlClause As String
sqlClause = ""
If GodownID > 0 Then sqlClause = " AND GodownID = " & GodownID

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase, " & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) as StockPurchaseAmount," & _
                   " RelationID FROM Stock " & _
                   " WHERE TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND GodownID = " & GodownID & _
                   " AND Quantity > 0 AND ( VoucherType = " & eVoucherType & _
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

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales," & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) as StockSalesAmount," & _
                   " RelationID FROM Stock " & _
                   " WHERE GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND Quantity > 0 AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales,B.StockSalesAmount " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B " & _
                   " ON A.RelationID = B.RelationID "

Call gDbTrans.CreateView("QryStockMerge")




CreateStockMergeQuery = True
End Function

Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property

Public Property Let FromRpDate(ByVal NewValue As String)
    'If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
    'Assign the FromDate value in system date format also
    If NewValue = "" Then Exit Property
    m_FromUSDate = GetSysFormatDate(NewValue)
    'm_FromRpDate = GetAppFormatDate(m_FromUSDate)
End Property


Public Sub SetStockRepSelectForm(ByRef RepSelectForm As frmRepSelect)
Set m_frmRepSlect = RepSelectForm
End Sub
Private Function CreateStockMergeQueryIncBranches() As Boolean
'Declare the variables
Dim USFromDate As String
Dim USAsOnDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

CreateStockMergeQueryIncBranches = False

USFromDate = GetSysFormatDate(FinIndianFromDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase,GodownID,RelationID," & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) as StockPurchaseAmount" & _
                   " FROM Stock " & _
                   " WHERE TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " AND Quantity > 0 " & _
                   " GROUP BY GodownID,RelationID"

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,GodownID,RelationID," & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount)) as StockSalesAmount" & _
                   " FROM Stock" & _
                   " WHERE TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " AND Quantity > 0 " & _
                   " GROUP BY GodownID,RelationID"

Call gDbTrans.CreateView("QryStockSales")


gDbTrans.SqlStmt = " SELECT A.*,B.StockSales, B.StockSalesAmount " & _
                   " FROM QryStockPurchase A LEFT JOIN QryStockSales B" & _
                   " ON A.GodownID=B.GodownID" & _
                   " AND A.RelationID = B.RelationID"

Call gDbTrans.CreateView("QryStockMerge")

CreateStockMergeQueryIncBranches = True

End Function

Public Property Get ToRpDate() As String
    ToRpDate = m_ToRpDate
End Property


Public Property Let ToRpDate(ByVal NewValue As String)
    'If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_ToRpDate = NewValue
    'Assign the TOdate value in system date format also
    If NewValue = "" Then Exit Property
    m_ToUSDate = GetSysFormatDate(NewValue)
    'm_ToRpDate = GetAppFormatDate(m_ToUSDate)
End Property
Private Function GetReportType(ByVal GodownID As Byte, _
            ByVal GroupID As Integer, ByVal ProductID As Long) As Wis_ReportType
                            
Select Case GodownID
    Case 0
        If GroupID = 0 Then If ProductID = 0 Then GetReportType = StockIncludingBranches
        If GroupID > 0 Then
            GetReportType = StockOfGroupIncBranches
            If ProductID > 0 Then GetReportType = StockOfGroupAndProductIncBranches
        End If
            
    Case Else
        If GroupID = 0 Then If ProductID = 0 Then GetReportType = StockAsOn
        If GroupID > 0 Then
            GetReportType = StockOfGroup
            If ProductID > 0 Then GetReportType = StockOfGroupAndProduct
        End If
End Select

End Function

Private Sub ProcessRepDetails(ByRef GodownID As Byte, _
                    ByRef GroupID As Integer, ByRef ProductID As Long)
                            
          


With m_frmRepSlect

    If .cmbGodown.ListIndex = -1 Then Exit Sub
    GodownID = .cmbGodown.ItemData(.cmbGodown.ListIndex)
    
    If .cmbGroup.ListCount > 0 Then If .cmbGroup.ListIndex = -1 Then .cmbGroup.ListIndex = 0
    
    If .cmbGroup.ListIndex >= 0 Then GroupID = .cmbGroup.ItemData(.cmbGroup.ListIndex)
    
    If .cmbProductName.ListCount = 0 Then
        .cmbProductName.AddItem ""
        .cmbProductName.ItemData(.cmbProductName.newIndex) = 0
         ProductID = 0
    Else
        If .cmbProductName.ListIndex >= 0 Then ProductID = .cmbProductName.ItemData(.cmbProductName.ListIndex)
    End If
    
End With
                            
End Sub

Public Function ShowStockReports() As Boolean
'Trap an error
On Error GoTo ErrLine

'Delclare the variables
Dim GodownID As Byte
Dim headID As Long
Dim GroupID As Integer
Dim ProductID As Long

Dim RepType As Wis_ReportType

ShowStockReports = False

'Show the report Date to be selected
'Call ShowReportDate(True)

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call LoadFrmWeb

'Set the Web Doc
Call SetWebDocument



'Get the GodownID,HeadID,GroupID,ProductID
Call ProcessRepDetails(GodownID, GroupID, ProductID)

'Get the Proper Reort Type according the specified Variables
RepType = GetReportType(GodownID, GroupID, ProductID)

If RepType < 1 Then Exit Function

Select Case RepType
    Case StockIncludingBranches
        If Not ShowStockAsOnDateIncludingBranches Then Exit Function
        
        
    Case StockOfGroupIncBranches
        If Not ShowStockForGroupIncBranches(GroupID) Then Exit Function
        
        
    Case StockOfGroupAndProductIncBranches
        If Not ShowStockOfGroupAndProductIncBranches(GroupID, ProductID) Then Exit Function
        
        
    Case StockAsOn
        If Not ShowStockAsOnDate(GodownID) Then Exit Function
        
    Case StockOfGroup
        If Not ShowStockForGroup(GodownID, GroupID) Then Exit Function
        
    Case StockOfGroupAndProduct
        If Not ShowStockOfGroupAndProduct(GodownID, GroupID, ProductID) Then Exit Function
        
        
End Select


Screen.MousePointer = vbDefault

Call ShowCurrentReport
ShowStockReports = True
Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockReports():" & vbCrLf & Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Function

'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfManuGroupAndProduct(ByVal GodownID As Byte, _
                            ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Integer) As Boolean

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String


Dim RelationID As Long
Dim SlNo As Integer


Dim StockPurchase As Double
Dim StockSales As Double
Dim StockBalance As Double

Dim ValueBalance As Currency
Dim TotalValue As Currency
Dim ClosingStock As Currency
Dim PurchasePrice As Currency

Dim MaterialClass As clsMaterial

'Declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field

'Iniitialse the function
ShowStockOfManuGroupAndProduct = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForManuGroupAndProduct(rstStock, headID, GroupID, ProductID, GodownID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * from qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

Set MaterialClass = Nothing

'Set stock balance Header
Call SetStockBalanceHeader(headName, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0


Do
    If rstStock.EOF Then Exit Do
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
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


ShowStockOfManuGroupAndProduct = True
Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfManuGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
'
Private Function PrepareRstOfStockForManuGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, _
                 ByVal lngHeadID As Long, ByVal intGroupID As Integer, _
                 ByVal lngProductID As Long, ByVal GodownID As Byte) As Boolean


If Not CreateStockMergeQuery(GodownID) Then Exit Function

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.RelationID," & _
                        " D.ProductName," & _
                        " F.UnitName, A.StockPurchaseAmount, A.StockSalesAmount " & _
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
'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfGroupAndProduct(ByVal GodownID As Byte, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim headID As Long
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
Dim fldProductName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldUnitName As ADODB.Field

'Iniitialse the function
ShowStockOfGroupAndProduct = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForGroupAndProduct(rstStock, GroupID, GodownID, ProductID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
'Set fldCompanyName = rstStock.Fields("CompanyName")
'Set fldHeadID = rstStock.Fields("HeadID")
Set fldUnitName = rstStock.Fields("UnitName")

ClosingStock = 0

Do
    If rstStock.EOF Then Exit Do
    
    'If company Changes then

    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00") 'Val(StockPurchase - StockSales)
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    
    
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
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


ShowStockOfGroupAndProduct = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function

'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockForGroup(ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean
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
Dim fldStockAmount As ADODB.Field
Dim fldUnitName As ADODB.Field


'Iniitialse the function
ShowStockForGroup = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForGroup(rstStock, GroupID, GodownID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

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
Set fldStockAmount = rstStock.Fields("StockPurchaseAmount")
'Set fldHeadID = rstStock.Fields("HeadID")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0


Do
    If rstStock.EOF Then Exit Do
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00") 'Val(StockPurchase - StockSales)
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
                    
    ValueBalance = 0
    If StockSales = 0 Then
        ValueBalance = FormatField(fldStockAmount)
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
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
    'Resume
    If Err Then
        MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function


'This function will prepeare the recordset for the Stock Group
' Inputs
'       byRef rstTotalStock  as recordset
'       byval intGroupID as integer
' OutPuts
'       rstTotalStock as Recordset
Private Function PrepareRstOfStockForGroup(ByRef rstTotalStock As ADODB.Recordset, _
                                          ByVal intGroupID As Integer, ByVal GodownID As Byte) As Boolean

PrepareRstOfStockForGroup = False

If Not CreateStockMergeQuery(GodownID) Then Exit Function

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, " & _
                        " F.UnitName,A.StockPurchaseAmount,A.StockSalesAmount " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function


PrepareRstOfStockForGroup = True

End Function
'This function will prepeare the recordset for the Stock Group
' Inputs
'       byRef rstTotalStock  as recordset
'       byval intGroupID as integer
' OutPuts
'       rstTotalStock as Recordset
Private Function PrepareRstOfStockForGroupAndProduct(ByRef rstTotalStock As ADODB.Recordset, _
                                          ByVal intGroupID As Integer, ByVal GodownID As Byte, ByVal ProductID As Long) As Boolean

PrepareRstOfStockForGroupAndProduct = False

If Not CreateStockMergeQuery(GodownID) Then Exit Function

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, " & _
                        " F.UnitName,A.StockPurchaseAmount,A.StockSalesAmount " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function


PrepareRstOfStockForGroupAndProduct = True

End Function


'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfManuAndGroup(ByVal GodownID As Byte, _
                            ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String

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
Dim fldUnitName As ADODB.Field

'Iniitialse the function
ShowStockOfManuAndGroup = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstForStockManuAndGroup(rstStock, headID, GroupID, GodownID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(headName, GroupName)

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
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    If StockSales = 0 Then
        FormatField (rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
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


ShowStockOfManuAndGroup = True
Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfManuAndGroup(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function


Private Function PrepareRstForStockManuAndGroup(ByRef rstTotalStock As ADODB.Recordset, _
                        ByVal lngHeadID As Long, _
                        ByVal intGroupID As Integer, ByVal GodownID As Byte) As Boolean

If Not CreateStockMergeQuery(GodownID) Then Exit Function

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName,A.StockPurchaseAmount,A.StockSalesAmount " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.HeadID = " & lngHeadID & _
                   " AND B.GroupID = " & intGroupID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY B.HeadID "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForStockManuAndGroup = True

End Function








'
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockForManufacturer(ByVal GodownID As Byte, ByVal headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String

Dim GroupID As Long
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
Dim fldUnitName As ADODB.Field


'Iniitialse the function
ShowStockForManufacturer = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForManufacutrer(rstStock, GodownID, headID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
headName = MaterialClass.GetCompanyName(headID)

'Set stock balance Header
Call SetStockBalanceHeader(headName)

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
    
    'If Group Changes
    If GroupID <> fldGroupID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        
        'Print the GroupName
        Call PRINTGroupName(GroupName)
    End If
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
                    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
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


ShowStockForManufacturer = True
Exit Function

ErrLine:
    If Err Then
        MsgBox "ShowStockForManufacturer(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function

Private Function PrepareRstOfStockForManufacutrer(ByRef rstTotalStock As ADODB.Recordset, _
                                    ByVal GodownID As Byte, ByVal headID As Long) As Boolean

PrepareRstOfStockForManufacutrer = False

If Not CreateStockMergeQuery(GodownID) Then Exit Function

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, E.CompanyName," & _
                        " E.HeadID,F.UnitName,A.StockPurchaseAmount,A.StockSalesAmount " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND B.HeadID = " & headID & _
                   " ORDER BY C.GroupName,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockForManufacutrer = True

End Function


'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockAsOnDate(ByVal GodownID As Byte) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset
Dim rstOpStock As ADODB.Recordset

Dim USAsOnDate As String
Dim headID As Long
Dim GroupID As Long
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
Dim OpStockValue As Currency
Dim OpStock As Double

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
Dim fldUnitName As ADODB.Field


'Iniitialse the function
ShowStockAsOnDate = False

Call StartTimer

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Prepare the recordset
If Not PrepareRecordsetForTotalStock(rstStock, GodownID) Then Exit Function

'COMMENTED BELOW LINE BECAUSE NOT USED IN THE CODE
'Call CreateOpStockQuery(GodownID, m_FromRpDate, "qryOpStock")
'gDbTrans.SqlStmt = "SELECT OpStock,RelationID FROM QryOpStock"
'Call gDbTrans.Fetch(rstOpStock, adOpenForwardOnly)


'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * from qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle(GodownID)

'Set stock balance Header
Call SetStockBalanceHeader

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGroupID = rstStock.Fields("GroupID")
Set fldGroupName = rstStock.Fields("GroupName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")


ClosingStock = 0
Set MaterialClass = New clsMaterial

Do
    If rstStock.EOF Then Exit Do
    
  
    'If Group Changes
    If GroupID <> fldGroupID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        
        'Print the GroupName
        Call PRINTGroupName(GroupName)
    End If
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    ValueBalance = 0
    
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    OpStock = 0
    OpStockValue = OpStock * PurchasePrice
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, OpStock + StockBalance, OpStockValue + ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance + opstock
    TotalValue = TotalValue + ValueBalance + OpStockValue
    
    'Total Closing Stock
    ClosingStock = ClosingStock + ValueBalance + OpStockValue
    
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


ShowStockAsOnDate = True

Exit Function

ErrLine:
    If Err Then
        'Resume
        MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Function
    End If

End Function

Public Function PrepareRecordsetForTotalStock(ByRef rstTotalStock As ADODB.Recordset, ByVal GodownID As Byte) As Boolean


PrepareRecordsetForTotalStock = False

If Not CreateStockMergeQuery(GodownID) Then Exit Function

gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales, A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, " & _
                        " F.UnitName,A.StockPurchaseAmount,A.StockSalesAmount  " & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F " & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " ORDER BY C.GroupName,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForTotalStock = True

End Function

'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfManuAndGroupIncBranches(ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim GodownID As Byte
Dim RelationID As Long
Dim SlNo As Integer

Dim USAsOnDate As String
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

'Declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockOfManuAndGroupIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockOfManuAndGroupIncBranches(rstStock, headID, GroupID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)

GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(headName, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")
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
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
                    
    
    'PurchasePrice = MaterialClass.GetPurchasePrice(RelationID, GodownID, m_FromRpDate, m_ToRpDate)
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
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

If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    
    BranchTotal = 0
End If

'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer


ShowStockOfManuAndGroupIncBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfManuAndGroupIncBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockForManufacturerIncBranches(ByVal headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim GodownID As Byte
Dim GroupID As Integer
Dim RelationID As Long
Dim SlNo As Integer

Dim USAsOnDate As String
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

'Declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockForManufacturerIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForManuIncBranches(rstStock, headID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)

'Set stock balance Header
Call SetStockBalanceHeader(headName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGroupID = rstStock.Fields("GroupID")
Set fldGroupName = rstStock.Fields("GroupName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldHeadID = rstStock.Fields("HeadID")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")
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
    
    'If Group Changes
    If GroupID <> fldGroupID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        
        'Print the GroupName
        Call PRINTGroupName(GroupName)
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
                    
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
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

If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    
    BranchTotal = 0
End If

'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer


ShowStockForManufacturerIncBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockForManufacturerIncBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function

'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockForGroupIncBranches(ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim GodownID As Byte
Dim headID As Integer
Dim RelationID As Long
Dim SlNo As Integer

Dim USAsOnDate As String
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

'Declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockForGroupIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Call StartTimer

'Prepare the recordset
If Not PrepareRstOfStockForGroupIncBranches(rstStock, GroupID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")
Set fldHeadID = rstStock.Fields("HeadID")
Set fldHeadName = rstStock.Fields("CompanyName")


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
    
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
    RelationID = fldRelationID.Value
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
        
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
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

If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    
    BranchTotal = 0
End If

'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer


ShowStockForGroupIncBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockForGroupIncBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
Private Function PrepareRstOfStockOfManuAndGroupIncBranches(ByRef rstTotalStock As ADODB.Recordset, _
                                    ByVal headID As Long, ByVal GroupID As Integer) As Boolean


PrepareRstOfStockOfManuAndGroupIncBranches = False

If Not CreateStockMergeQueryIncBranches Then Exit Function

'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " D.ProductName," & _
                        " F.UnitName,G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C,Products AS D," & _
                        " Units AS F,GodownDet AS G" & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID = C.GroupID" & _
                   " AND B.ProductID = D.ProductID" & _
                   " AND B.UnitID =F.UnitID" & _
                   " AND B.HeadID = " & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND A.GodownID= G.GodownID" & _
                   " ORDER BY A.GodownID,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockOfManuAndGroupIncBranches = True

End Function
Private Function PrepareRstOfStockForManuIncBranches(ByRef rstTotalStock As ADODB.Recordset, _
                                    ByVal headID As Long) As Boolean

PrepareRstOfStockForManuIncBranches = False
If Not CreateStockMergeQueryIncBranches Then Exit Function
'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " B.GroupID,C.GroupName,D.ProductName," & _
                        " E.HeadID,F.UnitName,G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C,Products AS D," & _
                        " Units AS F,GodownDet AS G" & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID = C.GroupID" & _
                   " AND B.ProductID = D.ProductID" & _
                   " AND B.UnitID =F.UnitID" & _
                   " AND B.HeadID = " & headID & _
                   " AND A.GodownID= G.GodownID" & _
                   " ORDER BY A.GodownID,C.GroupName,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockForManuIncBranches = True

End Function

Private Function PrepareRstOfStockForGroupIncBranches(ByRef rstTotalStock As ADODB.Recordset, _
                                    ByVal GroupID As Long) As Boolean

PrepareRstOfStockForGroupIncBranches = False
If Not CreateStockMergeQueryIncBranches Then Exit Function
'''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " D.ProductName, E.HeadID,E.CompanyName,F.UnitName, " & _
                        " G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C,Products AS D," & _
                        " CompanyCreation E,Units AS F,GodownDet AS G" & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID = C.GroupID" & _
                   " AND B.ProductID = D.ProductID" & _
                   " AND B.UnitID =F.UnitID AND E.HeadID = G.GodownID" & _
                   " AND B.GroupID = " & GroupID & _
                   " ORDER BY A.GodownID,D.ProductName "

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstOfStockForGroupIncBranches = True

End Function


'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockAsOnDateIncludingBranches() As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim GroupID As Long
Dim GodownID As Byte
Dim RelationID As Long
Dim SlNo As Integer

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String

Dim StockPurchase As Double
Dim StockSales As Double

Dim StockBalance As Double
Dim StockBalanceNew As Double
Dim ValueBalance As Currency
Dim ValueBalanceNew As Currency
Dim TotalValue As Currency
Dim BranchTotal As Currency

'Dim TotalStock As Double

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
Dim fldUnitName As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockAsOnDateIncludingBranches = False

Screen.MousePointer = vbHourglass

Call StartTimer

'Prepare the recordset
If Not PrepareRstForTotalStockInclBranches(rstStock) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, m_ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

'Set stock balance Header
Call SetStockBalanceHeader

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldGroupID = rstStock.Fields("GroupID")
Set fldGroupName = rstStock.Fields("GroupName")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")
ClosingStock = 0
Set MaterialClass = New clsMaterial

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
    
    'If Group Changes
    If GroupID <> fldGroupID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotal(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        
        'Print the GroupName
        Call PRINTGroupName(GroupName)
    End If
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    
    StockBalance = Format(StockPurchase - StockSales, "#.00")
        
'    Debug.Assert StockBalance > 0
    RelationID = fldRelationID.Value
    
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    'PurchasePrice = MaterialClass.GetPurchasePrice(RelationID, 1, m_FromRpDate, m_ToRpDate)
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
    'Total Closing Stock
    ClosingStock = ClosingStock + ValueBalance
    Debug.Print Format(ValueBalance, "#.00")
    'Move the recordset to Next Position
    rstStock.MoveNext
Loop


'Print the total
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
'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer

ShowStockAsOnDateIncludingBranches = True

Screen.MousePointer = vbDefault



Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockAsOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Screen.MousePointer = vbDefault
        

End Function
'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfGroupAndProductIncBranches(ByVal GroupID As Integer, ByVal ProductID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim GodownID As Byte
Dim RelationID As Long
Dim SlNo As Integer

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
Dim BranchTotal As Currency

Dim ClosingStock As Currency
Dim PurchasePrice As Currency

Dim MaterialClass As clsMaterial


'declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockOfGroupAndProductIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Screen.MousePointer = vbHourglass

Call StartTimer

'Prepare the recordset
If Not PrepareRstGroupAndProductInclBranches(rstStock, GroupID, ProductID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, m_ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")

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
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
'    Debug.Assert StockBalance > 0
    RelationID = fldRelationID.Value
    
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
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
    
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    
    BranchTotal = 0
End If
'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer

ShowStockOfGroupAndProductIncBranches = True

Screen.MousePointer = vbDefault

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfGroupAndProductIncBranches" & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Screen.MousePointer = vbDefault
        

End Function

'''This function will show the Total Stock  As on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowStockOfManuGroupAndProductIncBranches(ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStock As ADODB.Recordset
Dim rstPrice As ADODB.Recordset

Dim USAsOnDate As String
Dim GodownID As Byte
Dim RelationID As Long
Dim SlNo As Integer

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
Dim BranchTotal As Currency

Dim ClosingStock As Currency
Dim PurchasePrice As Currency

Dim MaterialClass As clsMaterial


'declare the variables to Store the Field DAta
Dim fldStockPurchase As ADODB.Field
Dim fldStockSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowStockOfManuGroupAndProductIncBranches = False

'If the dates are nothing then  exit
If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

Screen.MousePointer = vbHourglass

Call StartTimer

'Prepare the recordset
If Not PrepareRstManuGroupAndProductInclBranches(rstStock, headID, GroupID, ProductID) Then Exit Function

'Call LoadPurchasePrice(m_FromRpDate, m_ToRpDate, rstPrice)
Call CreateItemPriceQuery(m_FromRpDate, ToRpDate, "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

'Set the Header for the report
Call SetCompanyTitle

Set MaterialClass = New clsMaterial
headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

'Set stock balance Header
Call SetStockBalanceHeader(, GroupName)

Set fldStockPurchase = rstStock.Fields("StockPurchase")
Set fldStockSales = rstStock.Fields("StockSales")
Set fldRelationID = rstStock.Fields("RelationID")
Set fldProductName = rstStock.Fields("ProductName")
Set fldUnitName = rstStock.Fields("UnitName")
Set fldGodownID = rstStock.Fields("GodownID")
Set fldGodownName = rstStock.Fields("GodownName")

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
    
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    
    StockPurchase = FormatField(fldStockPurchase)
    StockSales = FormatField(fldStockSales)
    StockBalance = Format(StockPurchase - StockSales, "#.00")
    
'    Debug.Assert StockBalance > 0
    RelationID = fldRelationID.Value
    
    PurchasePrice = 0
    If gDbTrans.FindRecord(rstPrice, "RelationID=" & RelationID) Then PurchasePrice = rstPrice.Fields("UnitPrice")
    
    If StockSales = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    SlNo = SlNo + 1
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalStock(ProductName, StockBalance, ValueBalance, UnitName, PurchasePrice, SlNo)
    
    'TotalStock = TotalStock + StockBalance
    TotalValue = TotalValue + ValueBalance
    
    'Branch Total
    BranchTotal = BranchTotal + ValueBalance
    
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
    
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    
    BranchTotal = 0
End If
'Insert Closing Stock
Call PRINTClosingStock(ClosingStock, 5)

Call StopTimer

ShowStockOfManuGroupAndProductIncBranches = True

Screen.MousePointer = vbDefault

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockOfManuGroupAndProductIncBranches()" & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    Screen.MousePointer = vbDefault
        

End Function


'This will Show the Web Browser form
Private Sub ShowCurrentReport()

'Set m_frmWeb.m_HtmlToNavigate = m_WebDoc

'm_frmWeb.web.Navigate m_frmWeb.m_HtmlToNavigate

m_frmWeb.Show vbModal

End Sub


'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTClosingStock(ByVal ClosingStock As Currency, ByVal ColCount As Integer, Optional ByVal GrandTotal As String)

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
Private Function InsertRowAndCellSToTable(Table As HTMLTable, Optional CellsCount As Integer) As HTMLTableRow

Dim Cell As Integer
Dim Row As HTMLTableRow


With m_WebDoc

    Set Row = .createElement("TR")
    
    Row.Height = 18
    
    Table.tBodies(0).appendChild Row
    
    ' if user has not specified the cols then get the columns from the table
    If CellsCount = 0 Then CellsCount = Table.Cols
    
    For Cell = 0 To CellsCount - 1
        
        Set m_WebCell = .createElement("TD")
        
        Row.appendChild m_WebCell
    Next Cell
    
    m_WebCell.innerText = " "

End With

Set InsertRowAndCellSToTable = Row

End Function




'This will print the PrductName, StockBalance,ValueBalance of the
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTProductsForTotalStock(ByVal ProductName As String, _
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




''This function will print the Group Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTGroupName(ByVal GroupName As String)

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


''This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTGodownNames(ByVal GodownName As String)
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
Private Sub PRINTBranchTotal(ByVal TotalValue As Currency)

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




'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTTotal(ByVal TotalValue As Currency)

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

''This Sub Routine will Set the Header Stock Balance
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetStockBalanceHeader(Optional ByVal CompanyName As String, _
                                  Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow

With m_WebDoc
    strHeader = GetResourceString(175) & " " & GetFromDateString(m_ToRpDate)      '
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

Private Sub SetCompanyTitle(Optional ByVal GodownID As Byte)
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
        .Style.fontWeight = "Bold"
        .Align = "Center"
    End With
    
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    If Len(FormatField(rstCompany.Fields("Address"))) Then
        
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("Address").Value
        .Style.FONTSIZE = "13"
        .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    End If
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
      .Style.fontWeight = "Bold"
      .Align = "Center"
    End With
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
End With
End Sub


Private Sub SetWebDocument()

'This will set the  oncotextmenu  to return false
' so that if user clicks the right button popup menu will not be displayed
'm_frmWeb.web.Document.write "<Body oncontextmenu = " & AddQuotes("return false", True) & ">"
'content="text/html; charset=x-user-defined" http-equiv=Content-Type

'm_frmWeb.web.Document.write "<Body oncontextmenu =" & "'return false'" & ">"

Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName


End Sub

'This function will show the report date form
Private Function ShowReportDate(ByVal StDateLoacked As Boolean) As Boolean

If m_frmRptDt Is Nothing Then Set m_frmRptDt = New frmRptDt
With m_frmRptDt
    .txtStDate = FinIndianFromDate
    .txtEndDate = FinIndianEndDate
    
    .txtStDate.Locked = False
    If StDateLoacked Then .txtStDate.Locked = True
    
    .Show vbModal
    
End With
End Function
' This function gives the purchase price of the item from the database
' Inputs : RelationID as long
' Output : Gives the Purchase Price of the Item
Private Sub LoadPurchasePrice1(ByVal fromDate As String, _
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
                   " WHERE (VoucherType = " & VoucherType & _
                   " OR VoucherType = " & eVoucherType & " ) " & _
                   " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                   " AND #" & GetSysFormatDate(toDate) & "#" & _
                   " GROUP BY RelationID"
                   
Call gDbTrans.CreateView("QryMaxTransID")

'Fire the Query
gDbTrans.SqlStmt = " SELECT UnitPrice,a.RelationID " & _
                   " FROM Stock a,qryMaxTransID b " & _
                   " WHERE a.TransID=b.maxTransID " & _
                   " AND a.RelationID=b.RelationID"

gDbTrans.CreateView ("qryPrice")

gDbTrans.SqlStmt = " SELECT UnitPrice as PurchasePrice,A.RelationID " & _
                   " FROM qryPrice A " & _
                   " WHERE UnitPrice > 0"
gDbTrans.CreateView ("qryPrice1")

'There are Some Items which will have purchase price will be 0
'Like Container whih will come free with some materails
'and User Sales and Gets some amount From It
'So For Such Items Get the Item Price from Sales details
'Some thng Like Max Price

'So get the the Free items from getting whose purchase price is 0
gDbTrans.SqlStmt = " SELECT Max(UnitPrice) As PurchasePrice,RelationID " & _
                   " FROM Stock A " & _
                   " WHERE RelationID IN (SELECT RelationID " & _
                   " FROM qryPrice WHERE UnitPrice = 0) " & _
                   " Group BY RelationID"
gDbTrans.CreateView ("qryPrice2")

gDbTrans.SqlStmt = " Select RelationID,PurchasePrice as UnitPrice" & _
                   " From qryPrice1 " & _
                   " UNION " & _
                   " Select RelationID,PurchasePrice as UnitPrice" & _
                   " From qryPrice2 "


If gDbTrans.Fetch(rstPrice, adOpenDynamic) < 1 Then Exit Sub

Exit Sub

ErrLine:
    
End Sub

Public Function PrepareRstForTotalStockInclBranches(ByRef rstTotalStock As ADODB.Recordset) As Boolean
'Declare the variables
PrepareRstForTotalStockInclBranches = False

If Not CreateStockMergeQueryIncBranches Then Exit Function
''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " B.GroupID, C.GroupName, D.ProductName, " & _
                        " F.UnitName,G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F,GodownDet as G" & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID= G.GodownID " & _
                   " ORDER BY A.GodownID,C.GroupName,D.ProductName "


If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstForTotalStockInclBranches = True

End Function
Private Function PrepareRstManuGroupAndProductInclBranches(ByRef rstTotalStock As ADODB.Recordset, _
                                        ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Declare the variables
PrepareRstManuGroupAndProductInclBranches = False

If Not CreateStockMergeQueryIncBranches Then Exit Function

''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " D.ProductName," & _
                        " F.UnitName,G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F,GodownDet as G" & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.HeadId = " & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID= G.GodownID " & _
                   " ORDER BY A.GodownID,D.ProductName"


If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstManuGroupAndProductInclBranches = True

End Function
Private Function PrepareRstGroupAndProductInclBranches(ByRef rstTotalStock As ADODB.Recordset, _
                                        ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Declare the variables
PrepareRstGroupAndProductInclBranches = False

If Not CreateStockMergeQueryIncBranches Then Exit Function

''Get the SQL Query
gDbTrans.SqlStmt = " SELECT A.StockPurchase,A.StockSales,A.GodownID,A.RelationID," & _
                        " D.ProductName,E.CompanyName," & _
                        " E.HeadID,F.UnitName,G.GodownName,A.StockPurchaseAmount,A.StockSalesAmount" & _
                   " FROM QryStockMerge A,RelationMaster AS B," & _
                        " ProductGroup AS C, Products AS D, " & _
                        " CompanyCreation AS E, Units AS F,GodownDet as G" & _
                   " WHERE A.RelationID = B.RelationID " & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND G.GodownID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID= G.GodownID " & _
                   " ORDER BY A.GodownID,E.CompanyName,D.ProductName"
'

If gDbTrans.Fetch(rstTotalStock, adOpenStatic) <= 0 Then Exit Function

PrepareRstGroupAndProductInclBranches = True

End Function

Private Sub Class_Initialize()
'Load the web form
LoadFrmWeb

End Sub

Private Sub Class_Terminate()
Set m_frmRepSlect = Nothing

End Sub


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


