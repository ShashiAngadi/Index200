VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalesRep"
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

Private m_SootReport As Boolean '(shashi 25/12/2004)
Public Property Get ObjectName()
    ObjectName = "SalesReportClass"
End Property


Public Function ShowStockSalesReports() As Boolean
'Trap an error
On Error GoTo ErrLine

'Delclare the variables
Dim GodownID As Byte
Dim GroupID As Integer
Dim ProductID As Long

Dim RepType As Wis_ReportType

ShowStockSalesReports = False

'Show the report Date to be selected
'Call ShowReportDate(True)

If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Set the Web Doc
Call SetWebDocument

'Get the GodownID,HeadID,GroupID,ProductID
Call ProcessRepDetails(GodownID, GroupID, ProductID)

'Get the Proper Reort Type according the specified Variables
RepType = GetReportType(GodownID, GroupID, ProductID)

If RepType < 1 Then Exit Function

Select Case RepType
        Case SalesIncludingBranches
            If Not ShowSalesOnDateInclBranches Then Exit Function
            
            
        Case SalesOfGroupIncBranches
            If Not ShowSalesOfGroupInclBranches(GroupID) Then Exit Function
            
            
        Case SalesOfGroupAndProductIncBranches
            If Not ShowSalesOfGroupProductInclBranches(GroupID, ProductID) Then Exit Function
            
            
        Case SalesOfBranches
            If Not ShowSalesOnDate(GodownID) Then Exit Function
            
            
        Case SalesOfGroup
            If Not ShowSalesOfGroup(GodownID, GroupID) Then Exit Function
            
        Case SalesOfGroupAndProduct
            If Not ShowSalesOfGroupAndProduct(GodownID, GroupID, ProductID) Then Exit Function
        
End Select


Screen.MousePointer = vbDefault

Call ShowCurrentReport

ShowStockSalesReports = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockReports():" & vbCrLf & Err.Description, vbCritical
    Screen.MousePointer = vbDefault
    
End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManu(ByRef rstTotalPurchase As ADODB.Recordset, ByVal GodownID As Byte, ByVal headID As Long) As Boolean

'Declare the variables
Dim PurchaseVoucher As Wis_VoucherTypes

PrepareSalesRstOfManu = False

PurchaseVoucher = Sales
If m_SootReport Then PurchaseVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice" & _
                   " FROM Stock AS A,RelationMaster AS B, " & _
                        " ProductGroup AS C,Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND B.HeadID= " & headID & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID,D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManu = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManu(ByRef rstTotalPurchase As ADODB.Recordset, ByVal GodownID As Byte, ByVal headID As Long) As Boolean

'Declare the variables
Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManu = False


PurchaseVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice" & _
                   " FROM Stock AS A,RelationMaster AS B, " & _
                        " ProductGroup AS C,Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND B.HeadID= " & headID & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID,D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManu = True

End Function

'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOnDateInclBranches() As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim GodownID As Byte
Dim GroupID As Long
Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double


'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field


'Iniitialse the function
ShowSalesOnDateInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareRstForSalesInclBranches(rstSales) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

'Set the Header for the report
Call SetSalesHeader

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldGroupID = rstSales.Fields("GroupID")
Set fldGroupName = rstSales.Fields("GroupName")
Set fldProductName = rstSales.Fields("ProductName")
Set fldCompanyName = rstSales.Fields("CompanyName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")
GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
            Call PRINTTotalForPurchase(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        
        'Print the GroupName
        Call PRINTGroupNameForPurchase(GroupName)
    End If
        
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    BranchTotal = BranchTotal + QtyValue
    TotalValue = TotalValue + QtyValue
    GrandTotal = GrandTotal + QtyValue
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If

If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If

Call PRINTGrandTotal(GrandTotal)

    

ShowSalesOnDateInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOnDateInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function
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





'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfGroupInclBranches(ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim GodownID As Byte
Dim RelationID As Long
Dim UnitID As Long

Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double

Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field
'Iniitialse the function
ShowSalesOfGroupInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesOfGroupInclBranches(rstSales, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
Set MaterialClass = Nothing

'Set the Header for the report
Call SetSalesHeader(, GroupName)

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, _
                                UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    BranchTotal = BranchTotal + QtyValue
    GrandTotal = GrandTotal + QtyValue
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfGroupInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfGroupInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function
'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfGroupProductInclBranches( _
            ByVal GroupID As Integer, ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long
Dim GodownID As Byte

Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String

Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowSalesOfGroupProductInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesOfGroupProductInclBranches(rstSales, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
Set MaterialClass = Nothing

'Set the Header for the report
Call SetSalesHeader(, GroupName)

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    BranchTotal = BranchTotal + QtyValue
    TotalValue = TotalValue + QtyValue
    GrandTotal = GrandTotal + QtyValue
            
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If
Call PRINTGrandTotal(GrandTotal)

ShowSalesOfGroupProductInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfGroupProductInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function

'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfManuInclBranches(ByVal headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim GodownID As Byte
Dim GroupID As Long
Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim GodownName As String

Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowSalesOfManuInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfManuInclBranches(rstSales, headID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)

'Set the Header for the report
Call SetSalesHeader(headName)

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldGroupID = rstSales.Fields("GroupID")
Set fldGroupName = rstSales.Fields("GroupName")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
            Call PRINTTotalForPurchase(TotalValue)
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        'Print the GroupName
        Call PRINTGroupNameForPurchase(GroupName)
    End If
        
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    BranchTotal = BranchTotal + QtyValue
    GrandTotal = GrandTotal + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfManuInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOnDateInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function

'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfManuGroupInclBranches(ByVal headID As Long, GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim GodownID As Byte
Dim RelationID As Long
Dim UnitID As Long

Dim GodownName As String
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field
'Iniitialse the function
ShowSalesOfManuGroupInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfManuGroupInclBranches(rstSales, headID, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

'Set the Header for the report
Call SetSalesHeader(headName, GroupName)

Set MaterialClass = Nothing

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")
GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    BranchTotal = BranchTotal + QtyValue
    GrandTotal = GrandTotal + QtyValue
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If
Call PRINTGrandTotal(GrandTotal)

    
ShowSalesOfManuGroupInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOnDateInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function

'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfManuGroupProductInclBranches(ByVal headID As Long, GroupID As Integer, ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long
Dim GodownID As Byte

Dim GodownName As String
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim BranchTotal As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'declare the variables to Store the Field DAta
Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldGodownID As ADODB.Field
Dim fldGodownName As ADODB.Field

'Iniitialse the function
ShowSalesOfManuGroupProductInclBranches = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfManuGroupProductInclBranches(rstSales, headID, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

'Set the Header for the report
Call SetSalesHeader(headName, GroupName)

Set MaterialClass = Nothing

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldGodownID = rstSales.Fields("GodownID")
Set fldGodownName = rstSales.Fields("GodownName")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If godown Changes
    If GodownID <> fldGodownID.Value Then
        If TotalValue > 0 Then
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
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
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    BranchTotal = BranchTotal + QtyValue
    GrandTotal = GrandTotal + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    TotalValue = 0
End If
If BranchTotal > 0 Then
    'Print the total
    Call PRINTBranchTotal(BranchTotal)
    BranchTotal = 0
End If
Call PRINTGrandTotal(GrandTotal)

    
ShowSalesOfManuGroupProductInclBranches = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfManuGroupProductInclBranches(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    

End Function



'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRstForSalesInclBranches(ByRef rstSales As ADODB.Recordset) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareRstForSalesInclBranches = False

SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID," & _
                        " A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,G.GodownName " & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, " & _
                        " Units AS F,GodownDet AS G" & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID= G.GodownID " & _
                   " GROUP BY A.GodownID,C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,G.GodownName"

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareRstForSalesInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesOfGroupInclBranches(ByRef rstSales As ADODB.Recordset, ByVal GroupID As Integer) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesOfGroupInclBranches = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,GodownName" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products D, Units F,GodownDet G " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID =G.GodownID" & _
                   " GROUP BY A.GodownID,A.RelationID," & _
                        " D.ProductName, F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,GodownName" & _
                   " ORDER BY A.GodownID, D.ProductName,F.UnitID," & _
                        " F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice"

If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesOfGroupInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesOfGroupProductInclBranches(ByRef rstSales As ADODB.Recordset, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesOfGroupProductInclBranches = False

SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,GodownName" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, Units AS F,GodownDet G " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND B.ProductID=D.ProductID " & _
                   " AND A.GodownID=G.GodownID" & _
                   " GROUP BY A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,GodownName" & _
                   " ORDER BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"
                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesOfGroupProductInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesOfGroupProductInclBranches(ByRef rstSales As ADODB.Recordset, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesOfGroupProductInclBranches = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID,D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesOfGroupProductInclBranches = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesOfGroupInclBranches(ByRef rstFreeSales As ADODB.Recordset, ByVal GroupID As Integer) As Boolean
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesOfGroupInclBranches = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID,D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreeSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesOfGroupInclBranches = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManuInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfManuInclBranches = False

SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,G.GodownName" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D,Units AS F,GodownDet AS G " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HEadID =" & headID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID= G.GodownID " & _
                   " GROUP BY A.GodownID,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName,TransDate,UnitPrice,G.GodownName" & _
                   " ORDER BY A.GodownID,C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManuInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManuGroupInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfManuGroupInclBranches = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,G.GodownName" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F,GodownDet G" & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HEadID =" & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID = G.GodownID" & _
                   " GROUP BY A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,G.GodownName" & _
                   " ORDER BY D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice"

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManuGroupInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManuGroupProductInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long, _
                    ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfManuGroupProductInclBranches = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.GodownID,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,G.GodownName" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F,GodownDet G " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HeadID =" & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.GodownID=G.GodownID" & _
                   " AND B.ProductID=D.ProductID" & _
                   " GROUP BY A.GodownID,A.RelationID,D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,G.GodownName" & _
                   " ORDER BY D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice"

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManuGroupProductInclBranches = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManuGroupProductInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long, _
                    ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManuGroupProductInclBranches = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HeadID =" & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice"

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManuGroupProductInclBranches = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManuGroupInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManuGroupInclBranches = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HEadID =" & headID & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice"

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManuGroupInclBranches = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManuInclBranches(ByRef rstSales As ADODB.Recordset, ByVal headID As Long) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManuInclBranches = False


SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.HEadID =" & headID & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManuInclBranches = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRstForFreeSalesInclBranches(ByRef rstFreeSales As ADODB.Recordset) As Boolean

'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareRstForFreeSalesInclBranches = False


SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID, C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreeSales, adOpenDynamic) <= 0 Then Exit Function

PrepareRstForFreeSalesInclBranches = True

End Function


'''This function will print the Total into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Public Sub PRINTGrandTotal(ByVal GrandTotal As Currency)

'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
            
'Increment the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow.cells(1)
    .innerText = "GrandTotal"
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(4)
    .innerText = FormatCurrency((GrandTotal))
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


'This will print the PrductName, StockBalance,ValueBalance of the
'''This function will print the product Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTProductsForTotalPurchase(ByVal transdate As String, ByVal ProductName As String, _
                        ByVal UnitPrice As Currency, ByVal UnitName As String, _
                        ByVal StockQuantity As Double, ByVal StockValue As Currency)
    
'ByVal FreeQuantity As Double, ByVal FreeValue As Currency
'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(0)
    .innerText = transdate
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
    .innerText = UnitPrice
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
'With m_WebRow.cells(5)
'    .innerText = FreeQuantity
'    .Align = "Right"
'    .Style.fontWeight = "Normal"
'.Style.FontSize = 13
'    If gLangOffSet <> 0 Then .Style.FontSize = 15
'End With
'With m_WebRow.cells(6)
'    .innerText = FreeValue
'    .Align = "Right"
'    .Style.fontWeight = "Normal"
'    '.Style.FontSize = "x-small"
'End With
End Sub


''This function will print the Group Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written'
Private Sub PRINTGroupNameForPurchase(ByVal GroupName As String)
'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
        

'Print the group Name
With m_WebRow.cells(1)
    .innerText = "  " & GroupName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 14
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
End With
End Sub



'This function will print the company Name into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTCompanyNamesForPurchase(ByVal CompanyName As String)

'Increment Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

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


'''This function will print the Total  into the web browser
'This function is Tight Coupled,Just for the sake of readability of the main function
' This functions is written
Private Sub PRINTTotalForPurchase(ByVal TotalValue As Currency)
                                    
'ByVal FreeStock As Double,ByVal FreeValue As Currency)


'Increment the Row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

'Print Total Text
With m_WebRow.cells(1)
    .innerText = "Total"
    If gLangOffSet <> 0 Then .innerText = GetResourceString(52)
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
'With m_WebRow.cells(6)
'    .innerText = FreeValue
'    .Align = "Right"
'    .Style.fontWeight = "Bold"
'    '.Style.FontSize = "x-small"
'End With
End Sub


'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRecordsetForPurchase(ByRef rstTotalPurchase As ADODB.Recordset) As Boolean

'Declare the variables
Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRecordsetForPurchase = False


PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY E.CompanyName, E.HeadID, C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForPurchase = True

End Function


Private Sub SetWebDocument()

'This will set the  oncotextmenu  to return false
' so that if user clicks the right button popup menu will not be displayed
'm_frmWeb.web.Document.write "<Body oncontextmenu = " & AddQuotes("return false", True) & ">"
'content="text/html; charset=x-user-defined" http-equiv=Content-Type

'm_frmWeb.web.Document.write "<Body oncontextmenu =" & "'return false'" & ">"

Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName


End Sub


'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRecordsetForFreePurchase(ByRef rstFreePurchase As ADODB.Recordset) As Boolean

'Declare the variables
Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRecordsetForFreePurchase = False

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY E.CompanyName, E.HeadID, C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreePurchase, adOpenDynamic) <= 0 Then Exit Function

PrepareRecordsetForFreePurchase = True

End Function



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


Public Property Get ToRpDate() As String
    ToRpDate = m_ToRpDate
End Property


Public Property Let ToRpDate(ByVal NewValue As String)
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_ToRpDate = NewValue
    m_ToUSDate = GetSysFormatDate(NewValue)
    'm_ToRpDate = GetAppFormatDate(m_ToUSDate)
End Property
Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property

Public Property Let FromRpDate(ByVal NewValue As String)
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
    m_FromUSDate = GetSysFormatDate(NewValue)
    'm_FromRpDate = GetAppFormatDate(m_FromUSDate)
End Property



Private Sub ProcessRepDetails(ByRef GodownID As Byte, _
                            ByRef GroupID As Integer, ByRef ProductID As Long)
                            
With m_frmRepSlect

    If .cmbGodown.ListIndex = -1 Then Exit Sub
    GodownID = .cmbGodown.ItemData(.cmbGodown.ListIndex)
    
    'If .cmbManufacturer.ListCount > 0 Then _
            If .cmbManufacturer.ListIndex = -1 Then .cmbManufacturer.ListIndex = 0
    
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
''This Sub Routine will Set the Header for Sales
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetSalesHeader(Optional ByVal CompanyName As String, _
                                    Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow

With m_WebDoc
    
    strHeader = "Stock Sales Between " & m_FromRpDate & " AND " & m_ToRpDate
    If gLangOffSet <> 0 Then _
        strHeader = GetResourceString(175) & " " & _
        GetResourceString(180) & " " & GetFromDateString(m_FromRpDate, m_ToRpDate)
    
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
    
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
        Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
        With CurrentRow.cells(0)
            .innerText = GroupName
            .Align = "Center"
            .Style.FONTSIZE = 15
            .Style.fontWeight = "Bold"
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
    'Set m_WebTable = .createElement("TABLE")
    'Set DOMELement = .createElement("TBODY")
    
    Call StartHtmlTable("StockSales", m_WebDoc.body, 5)
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

With m_WebRow.cells(0)
    .innerText = "Date"
    If gLangOffSet <> 0 Then .innerText = GetResourceString(37)
    .Align = "Left"
    .Width = "7.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = "Particulars"
    If gLangOffSet <> 0 Then .innerText = GetResourceString(39)
    .Align = "Left"
    .Width = "39%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = "UnitPrice"
    If gLangOffSet <> 0 Then .innerText = GetResourceString(305)
    .Align = "Right"
    .Width = "7%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

'With m_WebRow.cells(3)
'    .innerText = "Op Stock"
'    .Align = "Right"
'    .Width = "13.5%"
'    .Style.fontWeight = "Bold"
'    .Style.FontSize = 15
'     If gLangOffSet <> 0 Then .Style.FontSize = 17
'End With

With m_WebRow.cells(3)
    .innerText = "Qty Sales"
    If gLangOffSet <> 0 Then _
            .innerText = GetResourceString(175, 180)
    .Align = "Right"
    .Width = "13.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(4)
    .innerText = "S Value"
    If gLangOffSet <> 0 Then _
            .innerText = GetResourceString(180, 393)
    .Align = "Right"
    .Width = "19.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
'With m_WebRow.cells(5)
'    .innerText = "Free"
'    .Align = "Right"
'    .Width = "12.5%"
'    .Style.fontWeight = "Bold"
'    '.Style.FontSize = "x-Small"
'End With
'With m_WebRow.cells(6)
'    .innerText = "Value"
'    .Align = "Right"
'    .Width = "12.5%"
'    .Style.fontWeight = "Bold"
'    '.Style.FontSize = "x-Small"
'End With
End Sub


'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOnDate(ByVal GodownID As Byte) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim GroupID As Long
Dim RelationID As Long
Dim UnitID As Long

Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double

Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field


'Iniitialse the function
ShowSalesOnDate = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareRecordsetForSales(rstSales, GodownID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

'Set the Header for the report
Call SetSalesHeader

Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldGroupID = rstSales.Fields("GroupID")
Set fldGroupName = rstSales.Fields("GroupName")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    'If Group Changes
    If GroupID <> fldGroupID.Value Then
        If TotalValue > 0 Then
            
            'Print the total
            Call PRINTTotalForPurchase(TotalValue)
            GrandTotal = GrandTotal + TotalValue
            TotalValue = 0
        End If
        GroupID = fldGroupID.Value
        GroupName = fldGroupName.Value
        'Print the GroupName
        Call PRINTGroupNameForPurchase(GroupName)
    End If
        
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    TotalValue = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOnDate = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfGroup(ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long

Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field

'Iniitialse the function
ShowSalesOfGroup = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfGroup(rstSales, GodownID, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)
'Set the Header for the report
Call SetSalesHeader(, GroupName)

Set MaterialClass = Nothing


Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, _
                    UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    TotalValue = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfGroup = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfGroup(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function

'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfManuGroup(ByVal GodownID As Byte, ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field

'Iniitialse the function
ShowSalesOfManuGroup = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfManuGroup(rstSales, GodownID, headID, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)
headName = MaterialClass.GetCompanyName(headID)

'Set the Header for the report
Call SetSalesHeader(headName, GroupName)

Set MaterialClass = Nothing


Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    TotalValue = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfManuGroup = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfManuGroup(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfManuGroupAndProduct(ByVal GodownID As Byte, ByVal headID As Long, _
                    ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field

'Iniitialse the function
ShowSalesOfManuGroupAndProduct = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfManuGroupAndProduct(rstSales, GodownID, headID, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)
headName = MaterialClass.GetCompanyName(headID)

'Set the Header for the report
Call SetSalesHeader(headName, GroupName)

Set MaterialClass = Nothing


Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    TotalValue = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfManuGroupAndProduct = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfManuGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function


'This function will show the Total Sales as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Public Function ShowSalesOfGroupAndProduct(ByVal GodownID As Byte, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstSales As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long

Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim transdate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

Dim fldTotalSales As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field

'Iniitialse the function
ShowSalesOfGroupAndProduct = False

'Prepare the recordset
'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareSalesRstOfGroupProduct(rstSales, GodownID, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial
GroupName = MaterialClass.GetGroupName(GroupID)
'Set the Header for the report
Call SetSalesHeader(, GroupName)

Set MaterialClass = Nothing


Set fldTotalSales = rstSales.Fields("TotalSales")
Set fldRelationID = rstSales.Fields("RelationID")
Set fldProductName = rstSales.Fields("ProductName")
Set fldUnitName = rstSales.Fields("UnitName")
Set fldTransDate = rstSales.Fields("TransDate")
Set fldUnitPrice = rstSales.Fields("UnitPrice")

GrandTotal = 0

Do
    If rstSales.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalSales.Value
    UnitPrice = fldUnitPrice.Value
    transdate = FormatField(fldTransDate)
    
    QtyValue = Quantity * UnitPrice
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(transdate, ProductName, _
                                UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstSales.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    TotalValue = 0
End If

Call PRINTGrandTotal(GrandTotal)

ShowSalesOfGroupAndProduct = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowSalesOfGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function



'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRecordsetForSales(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte) As Boolean

Dim SalesVoucher As Wis_VoucherTypes

PrepareRecordsetForSales = False

SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY  C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForSales = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfGroup(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean

Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfGroup = False

SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName, F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfGroup = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManuGroup(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal headID As Long, ByVal GroupID As Integer) As Boolean

Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfManuGroup = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.HEadID = " & headID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManuGroup = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfManuGroupAndProduct(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal headID As Long, _
                            ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfManuGroupAndProduct = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.HEadID = " & headID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY A.relationID,D.ProductName,F.UnitID,F.UnitName,TransDate,UnitPrice" & _
                   " ORDER BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfManuGroupAndProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManuGroupAndProduct(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal headID As Long, _
                            ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManuGroupAndProduct = False


SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.HEadID = " & headID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManuGroupAndProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfManuGroup(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfManuGroup = False


SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.HEadID = " & headID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfManuGroup = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareSalesRstOfGroupProduct(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareSalesRstOfGroupProduct = False


SalesVoucher = Sales
If m_SootReport Then SalesVoucher = StockSoot

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalSales,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.ProductID = D.ProductID" & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice"


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareSalesRstOfGroupProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfGroupProduct(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfGroupProduct = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = " & ProductID & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID,D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfGroupProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreeSalesRstOfGroup(ByRef rstSales As ADODB.Recordset, _
                            ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareFreeSalesRstOfGroup = False


SalesVoucher = FreeSales


'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = " & GroupID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID,D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "


If gDbTrans.Fetch(rstSales, adOpenStatic) <= 0 Then Exit Function

PrepareFreeSalesRstOfGroup = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareRecordsetForFreeSales(ByRef rstFreeSales As ADODB.Recordset, _
                            ByVal GodownID As Byte) As Boolean
'declare the variables
Dim SalesVoucher As Wis_VoucherTypes

PrepareRecordsetForFreeSales = False

SalesVoucher = FreeSales

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreeSales,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " AND A.GodownID = " & GodownID & _
                   " AND VoucherType = " & SalesVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " GROUP BY E.CompanyName, E.HeadID, C.GroupName," & _
                        " B.GroupID, D.ProductName,F.UnitID, F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "


If gDbTrans.Fetch(rstFreeSales, adOpenDynamic) <= 0 Then Exit Function

PrepareRecordsetForFreeSales = True

End Function


Private Function GetReportType(ByVal GodownID As Byte, ByVal GroupID As Integer, _
                             ByVal ProductID As Long) As Wis_ReportType
                            
Select Case GodownID
    Case 0
        If GroupID = 0 Then If ProductID = 0 Then GetReportType = SalesIncludingBranches
        If GroupID > 0 Then
            GetReportType = SalesOfGroupIncBranches
            If ProductID > 0 Then GetReportType = SalesOfGroupAndProductIncBranches
        End If
            
    Case Else
        If GroupID = 0 Then If ProductID = 0 Then GetReportType = SalesOfBranches
        If GroupID > 0 Then
            GetReportType = SalesOfGroup
            If ProductID > 0 Then GetReportType = SalesOfGroupAndProduct
        End If
End Select
End Function

'This will Show the Web Browser form
Private Sub ShowCurrentReport()

'Set m_frmWeb.m_HtmlToNavigate = m_WebDoc

'm_frmWeb.web.Navigate m_frmWeb.m_HtmlToNavigate

m_frmWeb.Show vbModal

End Sub


Public Sub SetStockRepSelectForm(ByRef RepSelectForm As frmRepSelect)
Set m_frmRepSlect = RepSelectForm
End Sub


Private Sub Class_Initialize()
Set m_frmWeb = New frmWeb

'Load the form
Load m_frmWeb

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


Public Property Get SootReport() As Boolean
    SootReport = m_SootReport
End Property

Public Property Let SootReport(ByVal NewValue As Boolean)
    m_SootReport = NewValue
End Property
