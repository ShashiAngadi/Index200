VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchaseRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_frmRepSlect As frmRepSelect

Private m_FromRpDate As String
Private m_ToRpDate As String

Private WithEvents m_frmRptDt As frmRptDt
Attribute m_frmRptDt.VB_VarHelpID = -1
Private WithEvents m_WebDoc As HTMLDocument
Attribute m_WebDoc.VB_VarHelpID = -1

Private m_frmWeb As frmWeb

Private m_WebTable As HTMLTable
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell

Public Property Get ObjectName()
    ObjectName = "PurchaseReportClass"
End Property


Public Function ShowStockPurchaseReports() As Boolean
'Trap an error




On Error GoTo ErrLine

'Delclare the variables
Dim GodownID As Byte
Dim headID As Long
Dim GroupID As Integer
Dim ProductID As Long

Dim RepType As Wis_ReportType

ShowStockPurchaseReports = False

'Show the report Date to be selected
'Call ShowReportDate(True)


If m_FromRpDate = "" Then Exit Function
If m_ToRpDate = "" Then Exit Function

'Set the Web Doc
Call SetWebDocument


'Get the GodownID,HeadID,GroupID,ProductID
Call ProcessRepDetails(GodownID, headID, GroupID, ProductID)

'Get the Proper Reort Type according the specified Variables
If GodownID = 0 Then Exit Function

RepType = GetReportType(headID, GroupID, ProductID)

If RepType < 1 Then Exit Function

Select Case RepType
    Case PurchaseOfBranches
        If Not ShowPurchasesOnDate(GodownID) Then Exit Function
        
    
    Case PurchaseOfGroup
        If Not ShowPurchasesOfGroup(GodownID, GroupID) Then Exit Function
                
    Case PurchaseOfGroupAndProduct
        If Not ShowPurchasesOfGroupAndProduct(GodownID, GroupID, ProductID) Then Exit Function
        
        
        
        
End Select


Screen.MousePointer = vbDefault

Call ShowCurrentReport
ShowStockPurchaseReports = True
Exit Function

ErrLine:
    If Err Then MsgBox "ShowStockReports():" & vbCrLf & Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Function

'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOfGroup(ByVal GodownID As Byte, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim headID As Long
Dim RelationID As Long
Dim UnitID As Long
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency

Dim UnitPrice As Currency

Dim Quantity As Double
Dim FreeQty As Double

Dim MaterialClass As clsMaterial

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOfGroup = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PreparePurchaseRstOfGroup(rstPurchase, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
'Set the Header for the report
Call SetPurchaseHeader(, GroupName)

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")


GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOfGroup = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOfGroup(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function
'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOfManuAndGroup(ByVal GodownID As Byte, ByVal headID As Long, ByVal GroupID As Integer) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency

Dim UnitPrice As Currency
Dim Quantity As Double
Dim MaterialClass As clsMaterial

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOfManuAndGroup = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PreparePurchaseRstManuAndOfGroup(rstPurchase, headID, GroupID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)
GroupName = MaterialClass.GetGroupName(GroupID)

'Set the Header for the report
Call SetPurchaseHeader(headName, GroupName)

Set MaterialClass = Nothing

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")


GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOfManuAndGroup = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOfManuAndGroup(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function

'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOfGroupAndProduct(ByVal GodownID As Byte, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim headID As Long
Dim RelationID As Long
Dim UnitID As Long
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency
Dim Quantity As Double
Dim MaterialClass As clsMaterial

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldCompanyName As ADODB.Field
Dim fldHeadID As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOfGroupAndProduct = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PreparePurchaseRstOfGroupAndProduct(rstPurchase, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
'Set the Header for the report
Call SetPurchaseHeader(, GroupName)

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")

GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOfGroupAndProduct = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOfGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function

'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOfManuGroupAndProduct(ByVal GodownID As Byte, ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim RelationID As Long
Dim UnitID As Long
Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency

Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOfManuGroupAndProduct = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PreparePurchaseRstOfManuGroupAndProduct(rstPurchase, headID, GroupID, ProductID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

GroupName = MaterialClass.GetGroupName(GroupID)
headName = MaterialClass.GetCompanyName(headID)

Set MaterialClass = Nothing

'Set the Header for the report
Call SetPurchaseHeader(headName, GroupName)

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")

GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
    RelationID = fldRelationID.Value
    ProductName = fldProductName.Value
    UnitName = fldUnitName.Value
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOfManuGroupAndProduct = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOfManuGroupAndProduct(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

End Function


'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOnDate(ByVal GodownID As Byte) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim headID As Long
Dim GroupID As Long
Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOnDate = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PrepareRecordsetForPurchase(rstPurchase) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

'Set the Header for the report
SetPurchaseHeader

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldGroupID = rstPurchase.Fields("GroupID")
Set fldGroupName = rstPurchase.Fields("GroupName")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")

GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
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
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    'TotalValue = IIf(IsNull(fldAmount.Value), TotalValue + QtyValue, FormatField(fldAmount))
    TotalValue = TotalValue + QtyValue
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOnDate = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOnDate(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    'Resume

End Function

'This function will show the Total Purchases as on Date
' This is main function , this will call few Titht coupled sub routines
' this is done for the sake of readablity of the code
Private Function ShowPurchasesOfManufacturer(ByVal GodownID As Byte, ByVal headID As Long) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstPurchase As ADODB.Recordset

Dim GroupID As Long
Dim RelationID As Long
Dim UnitID As Long

Dim headName As String
Dim GroupName As String
Dim UnitName As String
Dim ProductName As String
Dim TransDate As String

Dim QtyValue As Currency
Dim TotalValue As Currency
Dim GrandTotal As Currency
Dim UnitPrice As Currency

Dim Quantity As Double
Dim MaterialClass As clsMaterial

'Declare the flds to store the values
Dim fldTotalPurchase As ADODB.Field
Dim fldRelationID As ADODB.Field
Dim fldGroupID As ADODB.Field
Dim fldGroupName As ADODB.Field
Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldTransDate As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldAmount As ADODB.Field

'Iniitialse the function
ShowPurchasesOfManufacturer = False

'Get the Recordset of TotalPurchases between the specified dates
If Not PreparePurchaseRstOfManu(rstPurchase, headID) Then Exit Function

'Set the Header Name
Call SetCompanyTitle(GodownID)

Set MaterialClass = New clsMaterial

headName = MaterialClass.GetCompanyName(headID)
'Set the Header for the report
Call SetPurchaseHeader(headName)

Set MaterialClass = Nothing

Set fldTotalPurchase = rstPurchase.Fields("TotalPurchase")
Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldGroupID = rstPurchase.Fields("GroupID")
Set fldGroupName = rstPurchase.Fields("GroupName")
Set fldProductName = rstPurchase.Fields("ProductName")
Set fldUnitName = rstPurchase.Fields("UnitName")
Set fldTransDate = rstPurchase.Fields("TransDate")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldAmount = rstPurchase.Fields("Amount")

GrandTotal = 0

Do
    If rstPurchase.EOF Then Exit Do
    
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
    Quantity = fldTotalPurchase.Value
    UnitPrice = fldUnitPrice.Value
    TransDate = FormatField(fldTransDate)
    
    QtyValue = IIf(IsNull(fldAmount.Value), Quantity * UnitPrice, FormatField(fldAmount))
    
    'Print the Productname,Stock,on the Web Browser
    Call PRINTProductsForTotalPurchase(TransDate, ProductName, UnitPrice, UnitName, Quantity, QtyValue)
    
    TotalValue = TotalValue + QtyValue
    
    'Move the recordset to Next Position
    rstPurchase.MoveNext
Loop

'Print the total
If TotalValue > 0 Then
    'Print the total
    Call PRINTTotalForPurchase(TotalValue)
    GrandTotal = GrandTotal + TotalValue
    
End If
    
Call PRINTGrandTotal(GrandTotal)
 
ShowPurchasesOfManufacturer = True

Exit Function

ErrLine:
    If Err Then MsgBox "ShowPurchasesOfManufacturer(): " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        

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
    .innerText = FormatCurrency(GrandTotal)
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
Private Sub PRINTProductsForTotalPurchase(ByVal TransDate As String, ByVal ProductName As String, _
                        ByVal UnitPrice As Currency, ByVal UnitName As String, _
                        ByVal StockQuantity As Double, ByVal StockValue As Currency)
    
'ByVal FreeQuantity As Double, ByVal FreeValue As Currency
'Inrcrement the row
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)

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
    .innerText = FormatCurrency(StockValue)
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
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRecordsetForPurchase = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase," & _
             " A.RelationID,B.GroupID,C.GroupName," & _
             " D.ProductName,F.UnitID,F.UnitName," & _
             " TransDate, UnitPrice,A.Amount,A.TaxAmount" & _
        " FROM Stock AS A, RelationMaster AS B, " & _
             " ProductGroup AS C, Products AS D, Units AS F " & _
        " WHERE A.RelationID = B.RelationID" & _
        " AND TransDate  BETWEEN #" & USFromDate & "#" & _
        " AND #" & USAsOnDate & "#" & _
        " AND A.GodownID = " & 1 & _
        " AND VoucherType = " & PurchaseVoucher & _
        " AND B.GroupID = C.GroupID " & _
        " AND B.ProductID = D.ProductID " & _
        " AND B.UnitID =F.UnitID " & _
        "  AND A.PurORSaleID <> " & 0 & _
        " GROUP BY C.GroupName," & _
             " B.GroupID, D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
             " TransDate,UnitPrice,A.Amount,A.TaxAmount "
                  
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareRecordsetForPurchase = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PreparePurchaseRstOfGroup(ByRef rstTotalPurchase As ADODB.Recordset, ByVal GroupID As Integer) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PreparePurchaseRstOfGroup = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,A.Amount,A.TaxAmount" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PreparePurchaseRstOfGroup = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PreparePurchaseRstManuAndOfGroup(ByRef rstTotalPurchase As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PreparePurchaseRstManuAndOfGroup = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice,A.Amoount,A.TaxAmount" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.HeadID = " & headID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PreparePurchaseRstManuAndOfGroup = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreePurchaseRstManuAndOfGroup(ByRef rstTotalPurchase As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreePurchaseRstManuAndOfGroup = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.HeadID = " & headID & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreePurchaseRstManuAndOfGroup = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PreparePurchaseRstOfGroupAndProduct(ByRef rstTotalPurchase As ADODB.Recordset, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PreparePurchaseRstOfGroupAndProduct = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND B.ProductID =" & ProductID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY " & _
                        " D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PreparePurchaseRstOfGroupAndProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PreparePurchaseRstOfManuGroupAndProduct(ByRef rstTotalPurchase As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PreparePurchaseRstOfManuGroupAndProduct = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.HeadID = " & headID & _
                   " AND B.GroupID =" & GroupID & _
                   " AND B.ProductID =" & ProductID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PreparePurchaseRstOfManuGroupAndProduct = True

End Function
'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreePurchaseRstOfManuGroupAndProduct(ByRef rstTotalPurchase As ADODB.Recordset, ByVal headID As Long, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreePurchaseRstOfManuGroupAndProduct = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.HeadID = " & headID & _
                   " AND B.GroupID =" & GroupID & _
                   " AND B.ProductID =" & ProductID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreePurchaseRstOfManuGroupAndProduct = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreePurchaseRstOfGroupAndProduct(ByRef rstFreePurchase As ADODB.Recordset, ByVal GroupID As Integer, ByVal ProductID As Long) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreePurchaseRstOfGroupAndProduct = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND B.ProductID =" & ProductID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY E.CompanyName,E.HeadID," & _
                        " D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreePurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreePurchaseRstOfGroupAndProduct = True

End Function

'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreePurchaseRstOfGroup(ByRef rstFreePurchase As ADODB.Recordset, ByVal GroupID As Integer) As Boolean

'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreePurchaseRstOfGroup = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND B.GroupID =" & GroupID & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.HEadID = E.HeadID " & _
                   " AND B.UnitID =F.UnitID " & _
                   " AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY E.CompanyName, E.HeadID," & _
                        " D.ProductName,F.UnitID,F.UnitName, A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreePurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreePurchaseRstOfGroup = True

End Function


'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PrepareFreePurchaseRstOfManu(ByRef rstFreePurchase As ADODB.Recordset, ByVal headID As Long) As Boolean
'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareFreePurchaseRstOfManu = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice" & _
                   " FROM Stock AS A,RelationMaster AS B, " & _
                        " ProductGroup AS C,Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND B.HeadID= " & headID & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID,D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice "

                   
If gDbTrans.Fetch(rstFreePurchase, adOpenStatic) <= 0 Then Exit Function

PrepareFreePurchaseRstOfManu = True

End Function


'This function gives th recordset for total Quantity of Purchases between specifed dates
Private Function PreparePurchaseRstOfManu(ByRef rstTotalPurchase As ADODB.Recordset, ByVal headID As Long) As Boolean
'Declare the variables
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PreparePurchaseRstOfManu = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = Purchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as TotalPurchase,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,F.UnitID,F.UnitName," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount" & _
                   " FROM Stock AS A,RelationMaster AS B, " & _
                        " ProductGroup AS C,Products AS D,Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
                   " AND A.GodownID = " & 1 & _
                   " AND B.HeadID= " & headID & _
                   " AND VoucherType = " & PurchaseVoucher & _
                   " AND B.GroupID = C.GroupID " & _
                   " AND B.ProductID = D.ProductID " & _
                   " AND B.UnitID =F.UnitID " & _
                   "  AND A.PurORSaleID <> " & 0 & _
                   " GROUP BY C.GroupName," & _
                        " B.GroupID,D.ProductName,F.UnitID,F.UnitName,A.RelationID," & _
                        " TransDate,UnitPrice,A.Amount,A.TaxAmount "

                   
If gDbTrans.Fetch(rstTotalPurchase, adOpenStatic) <= 0 Then Exit Function

PreparePurchaseRstOfManu = True

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
Dim USAsOnDate As String
Dim USFromDate As String

Dim PurchaseVoucher As Wis_VoucherTypes

PrepareRecordsetForFreePurchase = False

USFromDate = GetSysFormatDate(m_FromRpDate)
USAsOnDate = GetSysFormatDate(m_ToRpDate) 'Stock As on Date

PurchaseVoucher = FreePurchase

'Get the SQL Query
gDbTrans.SqlStmt = " SELECT SUM(Quantity)as FreePurchase,A.RelationID,B.GroupID,C.GroupName," & _
                        " D.ProductName,E.CompanyName,E.HeadID,F.UnitID,F.UnitName," & _
                        " TransDate, UnitPrice" & _
                   " FROM Stock AS A, RelationMaster AS B, " & _
                        " ProductGroup AS C, Products AS D, CompanyCreation AS E, Units AS F " & _
                   " WHERE A.RelationID = B.RelationID" & _
                   " AND TransDate  BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USAsOnDate & "#" & _
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
End Property
Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property

Public Property Let FromRpDate(ByVal NewValue As String)
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
End Property



''This Sub Routine will Set the Header For Purchases
' to Print to  web browser Document Object mOdel Concept is followed
Private Sub SetPurchaseHeader(Optional ByVal CompanyName As String, _
                                Optional ByVal GroupName As String)

'declare the variables
Dim strHeader As String
Dim CurrentRow As HTMLTableRow



With m_WebDoc
    strHeader = "Stock Purchase Between " & m_FromRpDate & " AND " & m_ToRpDate
    If gLangOffSet <> 0 Then _
        strHeader = GetResourceString(175, 176) & " " & GetFromDateString(m_FromRpDate, m_ToRpDate)
        
    Set CurrentRow = InsertRowAndCellSToTable(.All("HeaderTbl", 0))
    
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
        'Set DOMElement = .createElement("H4")
        '.body.appendChild DOMElement
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
        'Set DOMElement = .createElement("H5")
        '.body.appendChild DOMElement
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
    
    Call StartHtmlTable("StockPurchase", m_WebDoc.body, 5)
    
    
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

With m_WebRow.cells(3)
    .innerText = "Qty Purchase"
    If gLangOffSet <> 0 Then _
            .innerText = GetResourceString(175, 176)
    .Align = "Right"
    .Width = "13.5%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(4)
    .innerText = "P.Value"
    If gLangOffSet <> 0 Then _
            .innerText = GetResourceString(176, 393)
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

Private Sub ProcessRepDetails(ByRef GodownID As Byte, ByRef headID As Long, _
                            ByRef GroupID As Integer, ByRef ProductID As Long)
                            
          
With m_frmRepSlect

    If .cmbGodown.ListIndex = -1 Then Exit Sub
    GodownID = .cmbGodown.ItemData(.cmbGodown.ListIndex)
    
    'If .cmbManufacturer.ListCount > 0 Then If .cmbManufacturer.ListIndex = -1 Then .cmbManufacturer.ListIndex = 0
    
    'If .cmbManufacturer.ListIndex >= 0 Then HeadID = .cmbManufacturer.ItemData(.cmbManufacturer.ListIndex)
    
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

Private Function GetReportType(ByVal headID As Long, _
                            ByVal GroupID As Integer, ByVal ProductID As Long) As Wis_ReportType
                            

GetReportType = 0

If headID = 0 Then
    If GroupID = 0 Then If ProductID = 0 Then GetReportType = PurchaseOfBranches
    If GroupID > 0 Then
        GetReportType = PurchaseOfGroup
        If ProductID > 0 Then GetReportType = PurchaseOfGroupAndProduct
    End If
End If

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

