VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_HeadID As Long

'local variable(s) to hold property value(s)
Private m_Quantity As Double 'local copy
Private m_UnitPrice As Currency 'local copy
Private m_TradingPrice As Currency 'local copy
Private m_RelationID As Long 'local copy
Private m_Amount As Currency 'local copy
Private m_TaxAmount As Currency 'local copy
Private m_MRP As Currency 'local copy
Private m_Manufacturer As String  'local copy
Private m_ProductName As String 'local copy
Private m_ProductGroupName As String 'local copy
Private m_Soot As Double 'local copy
Private m_ParentID As Long



Public Function GetGodownName(ByVal GodownID As Byte) As String
Dim rstGodown As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT GodownName FROM GodownDet" & _
                   " WHERE GodownID = " & GodownID
If gDbTrans.Fetch(rstGodown, adOpenForwardOnly) < 1 Then Exit Function

GetGodownName = rstGodown.Fields("GodownName")

End Function
'This function Gets the Stockbalance,StockValue of the particluar RealtionID
'Inputs
'           RelationID as long
'           AsOndate  as string
'Outputs
' Returns the ClosingStock
Public Function GetItemClStockBetweenDates(ByVal RelationID As Long, ByVal FromDateUS As String, ByVal ToDateUS As String, ByVal OpStock As Double) As Double

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStockBalance As ADODB.Recordset

Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes


'Initialise the function
GetItemClStockBetweenDates = 0

'validate the inputs
If RelationID = 0 Then Exit Function


'Get the Last Transaction of the realtionid
eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase FROM Stock" & _
                   " WHERE RelationID = " & RelationID & _
                   " AND TransDate > #" & FromDateUS & "#" & _
                   " AND TransDate <= #" & ToDateUS & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" '& _
                   " GROUP BY TransDate"

Call gDbTrans.CreateView("QryItemStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales FROM Stock" & _
                   " WHERE RelationID = " & RelationID & _
                   " AND TransDate > #" & FromDateUS & "#" & _
                   " AND TransDate <= #" & ToDateUS & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )" '& _
                   " GROUP BY TransDate"

Call gDbTrans.CreateView("QryItemStockSales")

gDbTrans.SqlStmt = " SELECT StockPurchase,StockSales " & _
                  " FROM QryItemStockPurchase, QryItemStockSales "
                  
If gDbTrans.Fetch(rstStockBalance, adOpenForwardOnly) < 1 Then Exit Function

GetItemClStockBetweenDates = OpStock + Val(FormatField(rstStockBalance("StockPurchase")) - FormatField(rstStockBalance("StockSales")))

ErrLine:

End Function





Public Sub LoadProductsToCombo(ByVal TheCombo As ComboBox, _
                               ByVal GroupID As Integer, ByVal UnitID As Long, ByVal GodownID As Byte)

Dim rstRelation As ADODB.Recordset
Dim ProductName As ADODB.Field
Dim RelationID As ADODB.Field

gDbTrans.SqlStmt = " SELECT A.ProductID,B.ProductName,RelationID " & _
                " FROM RelationMaster A, Products B" & _
                " WHERE A.GodownID  = " & GodownID & _
                " AND A.GroupID = " & GroupID & _
                " AND UnitID = " & UnitID & _
                " AND A.ProductID = B.ProductID "

If gDbTrans.Fetch(rstRelation, adOpenForwardOnly) < 1 Then Exit Sub

Set ProductName = rstRelation.Fields("ProductName")
Set RelationID = rstRelation.Fields("RelationID")

'Clear the Combo box
TheCombo.Clear

'Load the data to the combo box
Do While Not rstRelation.EOF
   
   TheCombo.AddItem ProductName.Value
   TheCombo.ItemData(TheCombo.newIndex) = RelationID.Value
   
   'Move the recordset
   rstRelation.MoveNext
Loop

End Sub
Public Sub LoadProductGroups(ByVal TheCombo As ComboBox)
'Declare the variables
Dim rstGroups As ADODB.Recordset
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field

gDbTrans.SqlStmt = " SELECT GroupID,GroupName FROM ProductGroup " & _
                 " ORDER BY GroupID "
        
Call gDbTrans.Fetch(rstGroups, adOpenForwardOnly)

Set fldGroupID = rstGroups.Fields("GroupID")
Set fldGroupName = rstGroups.Fields("GroupName")

'Clear the combo
TheCombo.Clear
Do
   If rstGroups.EOF Then Exit Sub
   TheCombo.AddItem fldGroupName.Value
   TheCombo.ItemData(TheCombo.newIndex) = fldGroupID.Value
   
   'Move to next record
   rstGroups.MoveNext
Loop

End Sub

'Private m_GodownID As Byte 'local copy
Public Sub LoadBranches(ByVal TheCombo As ComboBox, ByVal IsHO As Boolean)
Dim rstBranches As ADODB.Recordset
Dim GodownID As ADODB.Field
Dim GodownName As ADODB.Field

If IsHO Then
    gDbTrans.SqlStmt = " SELECT GodownID,GodownName FROM GodownDet " & _
                       " WHERE GoDownID = " & 1 & _
                       " ORDER BY GodownID "
Else
    gDbTrans.SqlStmt = " SELECT GodownID,GodownName FROM GodownDet " & _
                       " WHERE GoDownID <> " & 1 & _
                       " ORDER BY GodownID "
End If
        
Call gDbTrans.Fetch(rstBranches, adOpenStatic)

Set GodownID = rstBranches.Fields("GodownID")
Set GodownName = rstBranches.Fields("GodownName")
TheCombo.Clear

Do While Not rstBranches.EOF
   
   TheCombo.AddItem GodownName.Value
   TheCombo.ItemData(TheCombo.newIndex) = GodownID.Value
   
   'Move to next record
   rstBranches.MoveNext
Loop

End Sub
'Private m_GodownID As Byte 'local copy
Public Sub LoadBranchesExceptSpecified(ByVal TheCombo As ComboBox, ByVal GodownID As Byte)
Dim rstBranches As ADODB.Recordset
Dim fldGodownName As ADODB.Field
Dim fldGodownID As ADODB.Field


gDbTrans.SqlStmt = " SELECT GodownID,GodownName FROM GodownDet " & _
                   " WHERE GoDownID <> " & GodownID & _
                   " ORDER BY GodownID "

        
Call gDbTrans.Fetch(rstBranches, adOpenForwardOnly)

Set fldGodownID = rstBranches.Fields("GodownID")
Set fldGodownName = rstBranches.Fields("GodownName")
TheCombo.Clear

Do While Not rstBranches.EOF
   
   TheCombo.AddItem fldGodownName.Value
   TheCombo.ItemData(TheCombo.newIndex) = fldGodownID.Value
   
   'Move to next record
   rstBranches.MoveNext
Loop

TheCombo.ListIndex = -1
If TheCombo.ListCount = 1 Then TheCombo.ListIndex = 0

End Sub

'Private m_GodownID As Byte 'local copy
Public Sub LoadAllBranches(ByVal TheCombo As ComboBox)
Dim rstBranches As ADODB.Recordset
Dim GodownID As ADODB.Field
Dim GodownName As ADODB.Field

gDbTrans.SqlStmt = " SELECT GodownID,GodownName" & _
                   " FROM GodownDet " & _
                   " ORDER BY GodownID "
        
Call gDbTrans.Fetch(rstBranches, adOpenForwardOnly)

Set GodownID = rstBranches.Fields("GodownID")
Set GodownName = rstBranches.Fields("GodownName")

TheCombo.Clear
Do While Not rstBranches.EOF
   TheCombo.AddItem GodownName.Value
   TheCombo.ItemData(TheCombo.newIndex) = GodownID.Value
   
   'Move to next record
   rstBranches.MoveNext
Loop

If TheCombo.ListCount = 1 Then TheCombo.ListIndex = 0

End Sub

' This function gives the purchase price of the item from the database
' Inputs : RelationID as long
' Output : Gives the Purchase Price of the Item
Public Function GetSalesPrice(ByVal lngRelationID As Long, ByVal GodownID As Byte) As Currency
'Trap an error
On Error GoTo ErrLine
'declare Variables
Dim rst As ADODB.Recordset
Dim VoucherType As Wis_VoucherTypes
Dim eVoucherType As Wis_VoucherTypes
'Validat eh inputs
If lngRelationID = 0 Then Exit Function

VoucherType = Sales
eVoucherType = StockOut

GetSalesPrice = 0

gDbTrans.SqlStmt = " SELECT TOP 1 UnitPrice FROM Stock " & _
                   " WHERE RelationID = " & lngRelationID & _
                   " AND GodownID = " & GodownID & _
                   " AND ( VoucherType = " & VoucherType & _
                   " OR VoucherType = " & eVoucherType & " ) " & _
                   " ORDER BY TransID DESC "

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function


GetSalesPrice = FormatField(rst("UnitPrice"))


Exit Function

ErrLine:
    
End Function

' This function gives the purchase price of the item from the database
' Inputs : RelationID as long
' Output : Gives the Purchase Price of the Item
Public Function GetPurchasePrice(ByVal lngRelationID As Long, ByVal GodownID As Byte, _
                                                            ByVal fromDate As String, ByVal toDate As String) As Currency
'Trap an error
On Error GoTo ErrLine

'declare Variables
Dim rst As ADODB.Recordset
Dim VoucherType As Wis_VoucherTypes
Dim eVoucherType As Wis_VoucherTypes

If lngRelationID = 0 Then Exit Function

VoucherType = Purchase
eVoucherType = StockIn

gDbTrans.SqlStmt = " SELECT TOP 1 UnitPrice FROM Stock " & _
                   " WHERE RelationID = " & lngRelationID & _
                   " AND ( VoucherType = " & VoucherType & _
                   " OR VoucherType = " & eVoucherType & " ) " & _
                   " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                   " AND #" & GetSysFormatDate(toDate) & "#" & _
                  " ORDER BY TransID DESC "

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

GetPurchasePrice = FormatField(rst("UnitPrice"))

Exit Function

ErrLine:
    
End Function

'Gives the Company the given VendorID belongs to
'Input : VendorID as Long
'Returns : Company Type
Public Function GetCompanyType(VendorID As Long) As wis_CompanyType
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim rst As ADODB.Recordset

'initialise the function
GetCompanyType = 0

If VendorID = 0 Then Exit Function

'Get the query which retreives the CompanyType
gDbTrans.SqlStmt = " SELECT CompanyType FROM CompanyCreation " & _
                  " WHERE HeadID =" & VendorID
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

'Return the output
GetCompanyType = FormatField(rst("CompanyType"))

Exit Function

ErrLine:
    GetCompanyType = 0
End Function
' This function Gets the Stockbalance,StockValue of the particluar RealtionID
' Inputs
'           RelationID as long
'           AsOndate  as string
' Outputs
' Returns the MaxTransID
Public Function GetMaxTransID() As Long
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstTransID As ADODB.Recordset

'Initialise the function
GetMaxTransID = 0

'Get the Last Transaction of the realtionid
gDbTrans.SqlStmt = " SELECT MAX(TransID) AS MaxTransID FROM Stock "
                
If gDbTrans.Fetch(rstTransID, adOpenForwardOnly) < 1 Then Exit Function


'Get teh Max TransID
GetMaxTransID = FormatField(rstTransID("MaxTransID"))

ErrLine:

End Function

Public Property Get ObjectName()
    ObjectName = "MaterialClass"
End Property


'This function Gets the opening Stock value
'Outputs    TotalStockValue as on date as currency
' Returns the OpeningStockvalue
Public Function GetOpeningStockValue(Optional ByVal GodownID As Byte) As Currency
'Trap an error
On Error GoTo ErrLine


'In case he is not mantaining the Material Stock Details & salse & Prchase
'As in case Old INdex 2000 (i.e. Ver 2.0)
'In such case Fetch the Direct Close Stock Value
gDbTrans.SqlStmt = "Select * From ParentHeads " & _
        " Where ParentID = " & parSales
Dim rstTemp As Recordset
Dim StockValue As Currency

If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    If rstTemp("UserCreated") > 2 Then GoTo NextLine
    Set rstTemp = Nothing
    gDbTrans.SqlStmt = "Select Sum(OpAmount) as OpeningStockValue " & _
        " From OpBalance A Where A.HeadID IN " & _
            " (Select HeadID From Heads B Where " & _
            " ParentID = " & parPurchase & ")"

    If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Function
    
    While Not rstTemp.EOF
        StockValue = StockValue + FormatField(rstTemp("OpeningStockValue"))
        rstTemp.MoveNext
    Wend
    GetOpeningStockValue = StockValue
    Exit Function
End If


NextLine:
'Declare the variables
Dim rstStockBalance As ADODB.Recordset
Dim USFromDate As String

'Initialise the function
GetOpeningStockValue = 0
 
USFromDate = GetSysFormatDate(FinIndianFromDate)
'select top 10  (unitprice*quantity) as salesvalue, Amount , iif(isnull(Amount), (unitprice*quantity),Amount) as MatAmount from stock order by transdate desc
gDbTrans.SqlStmt = " SELECT SUM(iif(isnull(Amount), (Quantity*UnitPrice) ,Amount)) as StockValue" & _
                   " FROM Stock " & _
                   " WHERE  TransDate =#" & USFromDate & "#" & _
                   " AND PurORSaleID = 0 "

If GodownID > 0 Then
    gDbTrans.SqlStmt = " SELECT SUM(iif (isnull(Amount), (Quantity*UnitPrice) ,Amount)) as StockValue" & _
                   " FROM Stock " & _
                   " WHERE  TransDate =#" & USFromDate & "#" & _
                   " AND PurORSaleID =0 " & _
                   " AND GodownID = " & GodownID
End If
If gDbTrans.Fetch(rstStockBalance, adOpenForwardOnly) < 1 Then Exit Function

GetOpeningStockValue = FormatField(rstStockBalance.Fields("StockValue"))

ErrLine:
    

End Function
'This function Gets the Stock value Balance on specified date
'Inputs
'AsOndate  as string
'Outputs    TotalStockValue as on date as currency
'Returns the ClosingStockvalue
Public Function GetOnDateClosingStockValue(ByVal AsOnDate As Date) As Currency
'Trap an error
On Error GoTo ErrLine
'''''Test Temp Code on 15/1/2005
Dim rstStock As Recordset
Dim rstPrice As Recordset
Dim stockReportClass As New clsStockRep
stockReportClass.FromRpDate = FinIndianEndDate
stockReportClass.ToRpDate = GetIndianDate(AsOnDate)
Call stockReportClass.PrepareRstForTotalStockInclBranches(rstStock)

Call CreateItemPriceQuery(FinIndianFromDate, GetIndianDate(AsOnDate), "qryItemPrice")
gDbTrans.SqlStmt = "Select * From qryItemPrice"
Call gDbTrans.Fetch(rstPrice, adOpenDynamic)

Dim StockBalance As Double
Dim PurchasePrice As Double
Dim ValueBalance As Currency
Dim ClosingStock As Currency
Dim productRelationId As Long
Do
    If rstStock.EOF Then Exit Do
 
    StockBalance = Format(FormatField(rstStock("StockPurchase")) - FormatField(rstStock("StockSales")), "#.00")
        
    PurchasePrice = 0
        
    ValueBalance = 0
    If FormatField(rstStock("StockSales")) = 0 Then
        ValueBalance = FormatField(rstStock("StockPurchaseAmount"))
    Else
        productRelationId = FormatField(rstStock("RelationID"))
        If gDbTrans.FindRecord(rstPrice, "RelationID=" & productRelationId) Then PurchasePrice = rstPrice.Fields("UnitPrice")
        ValueBalance = StockBalance * PurchasePrice
    End If
    
    ClosingStock = ClosingStock + ValueBalance
    
    'Move the recordset to Next Position
    rstStock.MoveNext
Loop

GetOnDateClosingStockValue = ClosingStock
If ClosingStock > 0 Then Exit Function

'///test ends


'In case he is not mantaining the Material Stock Details & salse & Prchase
'As in case Old INdex 2000 (i.e. Ver 2.0)
'In such case Fetch the Direct Close Stock Value
gDbTrans.SqlStmt = "Select * From ParentHeads " & _
        " Where ParentID = " & parSales
Dim rstTemp As Recordset
Dim StockValue As Currency

If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    If rstTemp("UserCreated") > 2 Then GoTo NextLine
    Set rstTemp = Nothing
    gDbTrans.SqlStmt = "Select Sum(Stockvalue) as CloseStockValue,A.HeadID " & _
        " From CloseStock A Where A.TransDate = " & _
            " (Select Max(TransDate) From CloseStock B Where " & _
            " A.HeadID = B.HeadID And TransDate >= #" & FinUSFromDate & "#" & _
        " And TransDate <= #" & AsOnDate & "#) Group BY A.HeadID"
'    gDbTrans.SQLStmt = "Select Sum(Stockvalue) as CloseStockValue" & _
        " From CloseStock A Where TransDate = " & _
            " (Select Max(transDate) From CloseStock B Where " & _
            " A.HeadID = B.HeadID And TransDate >= #" & FinUSFromDate & "#)"
        
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then
        GetOnDateClosingStockValue = GetOpeningStockValue
        Exit Function
    End If
    While Not rstTemp.EOF
        StockValue = StockValue + FormatField(rstTemp("CloseStockValue"))
        
        rstTemp.MoveNext
    Wend
    GetOnDateClosingStockValue = StockValue
    Exit Function
End If


NextLine:

'Declare the variables
Dim rstStockBalance As ADODB.Recordset
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

'Initialise the function
GetOnDateClosingStockValue = 0
 
'validate the inputs
'If Not DateValidate(AsOnDate, "/", False) Then Exit Function

'USAsOnDate = FormatDate(AsOnDate)
'Get the Last Transaction of the realtionid

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

''''''Considering the Price of last transaction
'' Get the Unit Price
''Create a Query For to get the Unit price
Call CreateItemPriceQuery(FinIndianFromDate, GetIndianDate(AsOnDate), "qryItemPrice")

'gDbTrans.SQLStmt = " SELECT SUM(Quantity) as StockPurchase, " & _
                   " MAX(UnitPrice) as PurchasePrice, RelationID " & _
                   " FROM Stock " & _
                   " WHERE TransDate BETWEEN #" & FinUSFromDate & "#" & _
                   " AND #" & AsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & FreeRejectionsIN & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "
''New Code
gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase, A.RelationID, " & _
                   " Sum (IIf(IsNull(A.Amount), (UnitPrice * Quantity), A.Amount))  as StockPurchaseAmount " & _
                   " FROM STOCK A WHERE A.TransDate BETWEEN #" & FinUSFromDate & "#" & _
                   " AND #" & AsOnDate & "#" & _
                   " AND ( A.VoucherType = " & eVoucherType & _
                   " OR A.VoucherType = " & eFreeVoucherType & _
                   " OR A.VoucherType = " & eTransferVoucher & _
                   " OR A.VoucherType = " & FreeRejectionsIN & _
                   " OR A.VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY A.RelationID "
Call gDbTrans.CreateView("qryItemQuantity")

gDbTrans.SqlStmt = "Select X.RelationID,X.StockPurchase,Y.UnitPrice as PurchasePrice,X.StockPurchaseAmount" & _
                    " FROM qryItemQuantity X Inner Join qryItemPrice Y" & _
                    " ON X.RelationID = Y.RelationID "
'''New code

Call gDbTrans.CreateView("QryStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales,RelationID, " & _
                   " Sum (IIf(IsNull(Amount), (UnitPrice * Quantity), Amount))  as StockSalesAmount " & _
                   " FROM Stock A WHERE TransDate BETWEEN #" & FinUSFromDate & "#" & _
                   " AND #" & AsOnDate & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & FreeRejectionsOUT & _
                   " OR VoucherType = " & eRejectionVoucher & " )" & _
                   " GROUP BY RelationID "

Call gDbTrans.CreateView("QryStockSales")

gDbTrans.SqlStmt = " SELECT M.*,iif( N.StockSales is null,0,N.StockSales)AS StockSales,N.StockSalesAmount" & _
                   " FROM QryStockPurchase M LEFT JOIN QryStockSales N" & _
                   " ON M.RelationID=N.RelationID"

Call gDbTrans.CreateView("QryStockMerge")

gDbTrans.SqlStmt = " SELECT format(StockPurchase*PurchasePrice,'#.00') as TotalPurchase,StockPurchaseAmount, " & _
                   " StockSales*PurchasePrice as TotalSales" & _
                   " FROM QryStockMerge "
Call gDbTrans.CreateView("QryStockTotal")

gDbTrans.SqlStmt = " SELECT SUM(TotalPurchase),SUM(TotalSales),SUM(StockPurchaseAmount) " & _
                   " FROM QryStockTotal "

If gDbTrans.Fetch(rstStockBalance, adOpenForwardOnly) < 1 Then Exit Function

GetOnDateClosingStockValue = Val(FormatField(rstStockBalance(0)) - FormatField(rstStockBalance(1)))
GetOnDateClosingStockValue = Val(FormatField(rstStockBalance(2)) - FormatField(rstStockBalance(1)))

ErrLine:

End Function




'This function Gets the Stockbalance,StockValue of the particluar RealtionID
'Inputs
'           RelationID as long
'           AsOndate  as string
'Outputs
' Returns the ClosingStock
Public Function GetItemOnDateClosingStock(ByVal RelationID As Long, _
                                        ByVal AsonIndianDate As String, ByVal GodownID As Byte) As Double

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStockBalance As ADODB.Recordset
Dim USFromDate As String
Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes

'Initialise the function
GetItemOnDateClosingStock = 0

'validate the inputs
If RelationID = 0 Then Exit Function
If GodownID = 0 Then Exit Function

USFromDate = GetSysFormatDate(FinIndianFromDate)

'Get the Last Transaction of the realtionid

eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase FROM Stock " & _
                   " WHERE RelationID = " & RelationID & _
                   " AND GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & GetSysFormatDate(AsonIndianDate) & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )"
                   

Call gDbTrans.CreateView("QryItemStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut


gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales FROM Stock " & _
                   " WHERE RelationID = " & RelationID & _
                   " AND GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & GetSysFormatDate(AsonIndianDate) & "#" & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )"
                   
Call gDbTrans.CreateView("QryItemStockSales")


gDbTrans.SqlStmt = " SELECT StockPurchase,StockSales" & _
                  " FROM QryItemStockPurchase,QryItemStockSales"
                  
If gDbTrans.Fetch(rstStockBalance, adOpenForwardOnly) < 1 Then Exit Function

'Get teh Max TransID
GetItemOnDateClosingStock = Val(FormatField(rstStockBalance("StockPurchase")) - FormatField(rstStockBalance("StockSales")))

ErrLine:

End Function

'This function Gets the Stockbalance,StockValue of the particluar RealtionID
'Inputs
'           RelationID as long
'           AsOndate  as string
'Outputs
' Returns the ClosingStock
Public Function GetItemClosingStock(ByVal RelationID As Long, ByVal GodownID As Byte) As Double

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstStockBalance As ADODB.Recordset
Dim USFromDate As String
Dim USToDate As String

Dim eVoucherType As Wis_VoucherTypes
Dim eFreeVoucherType As Wis_VoucherTypes
Dim eTransferVoucher As Wis_VoucherTypes
Dim eSootVoucher As Wis_VoucherTypes
Dim eRejectionVoucher As Wis_VoucherTypes


'Initialise the function
GetItemClosingStock = 0

'validate the inputs
If RelationID = 0 Then Exit Function

If GodownID = 0 Then Exit Function

USFromDate = GetSysFormatDate(FinIndianFromDate)
USToDate = GetSysFormatDate(FinIndianEndDate)

'Get the Last Transaction of the realtionid
eVoucherType = Purchase
eFreeVoucherType = FreePurchase
eTransferVoucher = StockIn
eRejectionVoucher = RejectionsIn

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockPurchase FROM Stock " & _
                   " WHERE RelationID = " & RelationID & _
                   " AND GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USToDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )"

Call gDbTrans.CreateView("QryItemStockPurchase")

eVoucherType = Sales
eFreeVoucherType = FreeSales
eTransferVoucher = StockOut
eSootVoucher = StockSoot
eRejectionVoucher = RejectionsOut

gDbTrans.SqlStmt = " SELECT SUM(Quantity) as StockSales FROM Stock " & _
                   " WHERE RelationID = " & RelationID & _
                   " AND GodownID = " & GodownID & _
                   " AND TransDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USToDate & "# " & _
                   " AND ( VoucherType = " & eVoucherType & _
                   " OR VoucherType = " & eFreeVoucherType & _
                   " OR VoucherType = " & eSootVoucher & _
                   " OR VoucherType = " & eTransferVoucher & _
                   " OR VoucherType = " & eRejectionVoucher & " )"

Call gDbTrans.CreateView("QryItemStockSales")

gDbTrans.SqlStmt = " SELECT StockPurchase,StockSales " & _
                  " FROM QryItemStockPurchase, QryItemStockSales "
                  
If gDbTrans.Fetch(rstStockBalance, adOpenForwardOnly) < 1 Then Exit Function

GetItemClosingStock = Val(FormatField(rstStockBalance("StockPurchase")) - FormatField(rstStockBalance("StockSales")))

ErrLine:

End Function

Public Property Let Amount(ByVal vData As Currency)
Attribute Amount.VB_Description = "Amount for the Quantity of purchase."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    m_Amount = vData
End Property

Public Property Get Amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    Amount = m_Amount
End Property
Public Property Let TaxAmount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    m_TaxAmount = vData
End Property

Public Property Get TaxAmount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    TaxAmount = m_TaxAmount
End Property

Public Sub FixDailyAccTransHeads(AsOnDate As Date)

Dim CashDeposit As Currency
Dim cashWithdraw As Currency
Dim contraDeposit As Currency
Dim contraWithdraw As Currency

Dim Deposit As Currency
Dim WithDraw As Currency
Dim isCOntra As Boolean
Dim PrevTransID As Long
Dim firstTransId As Long
Dim rstTrans As Recordset
Dim headID As Long

'Debit Cash
    gDbTrans.SqlStmt = "Select * from AccTrans " & _
        " Where TransDate = #" & AsOnDate & "#" & _
        " And VoucherType = 1 Order By TransID"
    
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        PrevTransID = 0
        firstTransId = rstTrans("TransID")
        
        While Not rstTrans.EOF
            If PrevTransID <> rstTrans("TransID") Then
                If PrevTransID > 0 Then
                    If isCOntra Then
                        contraDeposit = contraDeposit + Deposit
                        contraWithdraw = contraWithdraw + WithDraw
                    Else
                        CashDeposit = CashDeposit + Deposit
                        cashWithdraw = cashWithdraw + WithDraw
                    End If
                End If
                Deposit = 0
                WithDraw = 0
                isCOntra = True
            End If
            If rstTrans("HeadID") = wis_CashHeadID Then isCOntra = False
            If rstTrans("HeadID") = headID Then
                Deposit = Deposit + FormatField(rstTrans("Debit"))
                WithDraw = WithDraw + FormatField(rstTrans("Credit"))
            End If
        rstTrans.MoveNext
        Wend
    End If

End Sub
Private Sub FixTransIdOfAccTrans(headID As Long, AsOnDate As Date)
        
Dim PrevTransID As Long
Dim firstTransId As Long
Dim rstTrans As Recordset
Dim existHeadIDs As String
Dim isMerge As Boolean
Dim isCOntra As Boolean

Dim rstTransID As Recordset
Dim ParentID As Long
Dim VoucherType As Wis_VoucherTypes

Dim AccHeadID As Long
Dim debitCash As Currency
Dim creditCash As Currency
Dim Debit As Currency
Dim Credit As Currency
Dim transCount As Integer
Dim innerSql As String

Dim intLoop As Integer
For intLoop = 0 To 2
    If intLoop = 0 Then VoucherType = payment
    If intLoop = 1 Then VoucherType = Receipt
    If intLoop = 2 Then VoucherType = Journal  'Contra


    innerSql = "Select distinct TransID from AccTrans A" & _
        " Inner Join Heads B On A.HeadID = B.HeadID " & _
        " Where TransDate = #" & AsOnDate & "#" & _
        " And HeadID = " & headID & _
        " ANd B.ParentID in ( " & parProfitORLoss & "," & parIncome & "," & parBankAccount & "," & parBankLoanAccount & _
                 "," & parFixedAsset & "," & parShareCapital & "," & parGovtLoanSubsidy & ", " & parOtherLoans & _
                 "," & parPayAble & " ," & parLoanIntProv & " ," & parSuspAcc & _
                 "," & parInvestment & "," & parLoanAdvanceAsset & ", " & parSalaryAdvance & _
                 "," & parSalaryExpense & "," & parReceivable & "," & parReserveFunds & _
                 "," & parBankIncome & "," & parExpense & "," & parTradingExpense & "," & parTradingIncome & _
                 "," & parBankExpense & "," & parSales & "," & parPurchase & _
                ")"
    
    
    gDbTrans.SqlStmt = "Select Distinct TransID from AccTrans where TransDate = #" & AsOnDate & "#" & _
        " And Transid Not in (" & innerSql & ")"
        
    If gDbTrans.Fetch(rstTransID, adOpenDynamic) > 0 Then
    End If
    
    gDbTrans.SqlStmt = "Select A.*,B.ParentID from AccTrans,Heads B " & _
        " Where A.HeadID=B.HeadID and TransDate = #" & AsOnDate & "#" & _
        " And HeadID = #" & headID & _
        " And VoucherType = " & VoucherType & " Order By TransID"
    
    
    existHeadIDs = ""
    PrevTransID = 0
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        firstTransId = rstTrans("TransID")
    End If
    
    While Not rstTrans.EOF
        AccHeadID = FormatField(rstTrans("HeadID"))
        ParentID = FormatField(rstTrans("ParentID"))
        
        If PrevTransID <> rstTrans("TransID") Then
            If firstTransId <> rstTrans("TransID") Then
                'The TransCtion has Change, check Can we merg this
                'Transacton into Prev Transctions
             If isMerge Then
                
                
                
             End If
            End If
            If PrevTransID > 0 Then
                If isCOntra Then
                    'Debit = contraDeposit + Deposit
                    'Credit = contraWithdraw + WithDraw
                Else
                    'CashDeposit = CashDeposit + Deposit
                    'cashWithdraw = cashWithdraw + WithDraw
                End If
            End If
            
            
            Debit = 0
            Credit = 0
            
            debitCash = 0
            creditCash = 0
            transCount = 0
            isMerge = True
            existHeadIDs = ""
        End If
        
        If (ParentID = parBankAccount Or ParentID = parBankLoanAccount _
             Or ParentID = parFixedAsset Or parGovtLoanSubsidy Or ParentID = parOtherLoans _
             Or ParentID = parPayAble Or ParentID = parLoanIntProv Or ParentID = parSuspAcc _
             Or ParentID = parInvestment Or ParentID = parLoanAdvanceAsset Or ParentID = parSalaryAdvance _
             Or ParentID = parSalaryExpense Or ParentID = parReceivable Or ParentID = parReserveFunds _
             Or ParentID = parBankIncome Or ParentID = parExpense Or ParentID = parTradingExpense Or parTradingIncome _
             Or ParentID = parBankExpense Or ParentID = parSales Or ParentID = parPurchase _
            ) Then
             
                isMerge = False
                If rstTrans("TransID") = firstTransId Then rstTrans.MoveLast
        End If
        
        If firstTransId = rstTrans("TransID") Then
            If InStr(0, existHeadIDs, CStr(headID), vbTextCompare) < 1 Then _
                existHeadIDs = existHeadIDs + CStr(headID) + ","
            'End If
        Else
        
        End If
        
        transCount = transCount + 1
        If rstTrans("HeadID") = wis_CashHeadID Then
            isCOntra = False
            debitCash = debitCash + FormatField(rstTrans("Debit"))
            creditCash = creditCash + FormatField(rstTrans("Credit"))
        End If
        If rstTrans("HeadID") = headID Then
            Debit = Debit + FormatField(rstTrans("Debit"))
            Credit = Credit + FormatField(rstTrans("Credit"))
        End If
        
        rstTrans.MoveNext
    Wend
    
Next
End Sub
Public Function GetHeadID() As Long
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim StockistType As wis_CompanyType
Dim rst As ADODB.Recordset
Dim Retval As Integer

'Initialise the function
GetHeadID = 0

StockistType = Enum_Stockist

'Dim ManufacturerType As wis_CompanyType
'ManufacturerType = Enum_Manufacturer

'fire the query
gDbTrans.SqlStmt = " SELECT HeadID as VendorID,CompanyName,Address " & _
                   " FROM CompanyCreation " & _
                   " WHERE CompanyType = " & StockistType & _
                   " ORDER BY CompanyName "
            
Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)

If Retval < 1 Then Exit Function
If Retval < 1 Then
    'Crate ONe head For Cash TransaCtion
    gDbTrans.BeginTrans
    Dim bankClass As New clsBankAcc
    GetHeadID = bankClass.GetHeadIDCreated( _
        GetResourceString(269, 176), LoadResString(269) & " " & (176), parPayAble, 0)
    gDbTrans.CommitTrans
    Set bankClass = Nothing
    Exit Function
End If

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

'if there is more than one Vendor then Need not to select the vendor
'If Retval > 1 Then
'If there is one vendor select him
If Retval > 0 Then
    'If Not FillView( m_frmLookUp.lvwReport, Rst, "VendorID", True) Then Exit Function
    If Not FillView(m_frmLookUp.lvwReport, rst, True) Then Exit Function
    
    m_HeadID = 0
    
    m_frmLookUp.Show vbModal
Else
    m_HeadID = FormatField(rst(0))
End If
    
    GetHeadID = m_HeadID

Exit Function
ErrLine:
    GetHeadID = 0

End Function


Public Function isCompanyWithinState(headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine
'Declare the Variables
Dim rst As ADODB.Recordset
Dim StateValue As Byte

'Initialise the function
isCompanyWithinState = False

'Fire the query
gDbTrans.SqlStmt = " SELECT SameState FROM CompanyCreation " & _
                   " WHERE HeadID = " & headID
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

StateValue = FormatField(rst("SameState"))

If StateValue = 1 Then isCompanyWithinState = True

Exit Function

ErrLine:
    
End Function
Public Sub LoadCompaniesToCombo(CompanyType As wis_CompanyType, TheCmb As ComboBox)
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim rst As ADODB.Recordset
Dim fldHeadID As ADODB.Field
Dim fldCompanyName As ADODB.Field


'Fire the query
gDbTrans.SqlStmt = " SELECT HeadID,CompanyName FROM CompanyCreation " & _
                 " WHERE CompanyType= " & CompanyType & _
                 " ORDER BY HeadID"
                 
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub


Set fldHeadID = rst.Fields("HeadID")
Set fldCompanyName = rst.Fields("CompanyName")

'Clear the combobox
TheCmb.Clear

'Load the  Comapnies to the combo box
Do While Not rst.EOF
    TheCmb.AddItem fldCompanyName.Value
    TheCmb.ItemData(TheCmb.newIndex) = fldHeadID.Value
    
    rst.MoveNext
Loop

Exit Sub

ErrLine:
    
End Sub

Public Sub LoadUnits(ByVal UnitCombo As ComboBox)
Dim rstUnits As ADODB.Recordset
Dim UnitID As ADODB.Field
Dim UnitName As ADODB.Field

'
gDbTrans.SqlStmt = " SELECT UnitID,UnitName,UnitName FROM Units" & _
                   " ORDER BY UnitID "
        
Call gDbTrans.Fetch(rstUnits, adOpenForwardOnly)

Set UnitID = rstUnits.Fields("UnitID")
Set UnitName = rstUnits.Fields("UnitName")

Do
   If rstUnits.EOF Then Exit Sub
   UnitCombo.AddItem UnitName.Value
   UnitCombo.ItemData(UnitCombo.newIndex) = UnitID.Value
   
   'Move to next record
   rstUnits.MoveNext
Loop

End Sub

Public Property Let RelationID(ByVal vData As Long)
Attribute RelationID.VB_Description = "Unique Recognition of  the product."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_RelationID = vData
End Property


Public Property Get RelationID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    RelationID = m_RelationID
End Property
Public Property Let Manufacturer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_Manufacturer = vData
End Property


Public Property Get Manufacturer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    Manufacturer = m_Manufacturer
End Property
Public Property Let ProductName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_ProductName = vData
End Property


Public Property Get ProductName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    ProductName = m_ProductName
End Property
Public Property Let ProductGroupName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_ProductGroupName = vData
End Property


Public Property Get ProductGroupName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    ProductGroupName = m_ProductGroupName
End Property

Public Property Let UnitPrice(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitPrice = 5
    m_UnitPrice = vData
End Property


Public Property Get UnitPrice() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitPrice
    UnitPrice = m_UnitPrice
End Property


Public Property Let Quantity(ByVal vData As Double)
Attribute Quantity.VB_Description = "This is the Quantity of Purchase"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    m_Quantity = vData
End Property


Public Property Get Quantity() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Quantity = m_Quantity
End Property

Public Function GetCompanyName(lngHeadID As Long) As String

Dim rst As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT CompanyName " & _
                  " FROM CompanyCreation " & _
                  " WHERE HeadID = " & lngHeadID
                  
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

GetCompanyName = FormatField(rst.Fields("CompanyName"))

Set rst = Nothing

End Function
Public Function GetProductName(ByVal intGroupID As Integer, ByVal lngProductID As Long) As String

Dim rst As ADODB.Recordset

GetProductName = ""

gDbTrans.SqlStmt = " SELECT ProductName " & _
                  " FROM Products" & _
                  " WHERE GroupID = " & intGroupID & _
                  " AND ProductID = " & lngProductID
                  
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

GetProductName = FormatField(rst.Fields("ProductName"))

Set rst = Nothing
End Function

Public Function GetGroupName(intGroupID As Integer) As String

Dim rst As ADODB.Recordset

GetGroupName = ""

gDbTrans.SqlStmt = " SELECT GroupName " & _
                  " FROM ProductGroup " & _
                  " WHERE GroupID = " & intGroupID
                  
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

GetGroupName = FormatField(rst.Fields("GroupName"))

Set rst = Nothing

End Function

'geives the headID selected by the user
Public Function GetHeadIDFromHeadsList(enumCompanyType As wis_CompanyType) As Long
Dim rstVendors As ADODB.Recordset
Dim Retval As Integer

GetHeadIDFromHeadsList = 0
gDbTrans.SqlStmt = " SELECT HeadID as VendorID,CompanyName,Address " & _
                   " FROM CompanyCreation " & _
                   " WHERE CompanyType = " & enumCompanyType & _
                   " ORDER BY CompanyName "
Retval = gDbTrans.Fetch(rstVendors, adOpenStatic)
If Retval < 1 Then Exit Function

If Retval > 1 Then
    If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
    
    'If Not FillView(m_frmLookUp.lvwReport, rstVendors, "VendorID", True) Then Exit Function
    If Not FillView(m_frmLookUp.lvwReport, rstVendors, True) Then Exit Function
    
    
    m_HeadID = 0
    
    m_frmLookUp.Show vbModal
Else
    m_HeadID = FormatField(rstVendors("VendorID"))
End If

GetHeadIDFromHeadsList = m_HeadID
Set rstVendors = Nothing
End Function

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_HeadID = CLng(strSelection)
End Sub


Public Property Let TradingPrice(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TradingPrice = 5
    m_TradingPrice = vData
End Property


Public Property Get TradingPrice() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TradingPrice
    TradingPrice = m_TradingPrice
End Property

Public Property Let MRP(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MRP = 5
    m_MRP = vData
End Property


Public Property Get MRP() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TradingPrice
    MRP = m_MRP
End Property

Public Property Let Soot(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantity = 5
    m_Soot = vData
End Property


Public Property Get Soot() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantity
    Soot = m_Soot
End Property

Public Property Let ParentID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_ParentID = vData
End Property


Public Property Get ParentID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    ParentID = m_ParentID
End Property
