VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_frmPurchase As frmPurchase
Attribute m_frmPurchase.VB_VarHelpID = -1
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_dbOperation As wis_DBOperation
Private m_colMaterials As clsMaterials
Private m_GrdFunctions As clsGrdFunctions

Private m_VendorID As Long
Private m_RelationID As Long
Private m_PurTransID As Long
Private m_ReDirectHeadID As Long
Private m_InvoiceAmount As Currency
Private m_UpdateStatus As Boolean
Private m_TransIDString As String
Private m_InvoiceNo As String
Private m_TransDate As String
Private m_InvoiceDate As String
Private m_InvoiceType As Byte

Private m_isUpdatingGrid As Integer

Private Function AddPLHeadsToCollection(ByVal AccType As wis_AccountType) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim HeadAmount As Currency
Dim headName As String
Dim headID As Long
Dim ParentID As Long

AddPLHeadsToCollection = False
'Assign the values
With m_frmPurchase
    If AccType = Profit Then
        If .cmbIncomeHead.ListIndex = -1 Then Exit Function
        headID = .cmbIncomeHead.ItemData(.cmbIncomeHead.ListIndex)
        HeadAmount = Val(.txtIncome.Text)
        headName = .cmbIncomeHead.Text
    ElseIf AccType = Loss Then
        If .cmbExpenseHead.ListIndex = -1 Then Exit Function
        headID = .cmbExpenseHead.ItemData(.cmbExpenseHead.ListIndex)
        HeadAmount = Val(.txtExpense.Text)
        headName = .cmbExpenseHead.Text
    End If
End With
    
ParentID = GetParentID(headID)

'check whther this head is already  added or not
If m_colMaterials.Exists("MAT" & headID) Then
    MsgBox GetResourceString(752), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


Call m_colMaterials.Add(0, 0, 0, 0, headID, "", HeadAmount, 0, headName, "", ParentID, 0, "MAT" & headID)

AddPLHeadsToCollection = True

Exit Function

ErrLine:
    If Err.Number = 457 Then MsgBox "Head is already Exists in Collection !"

End Function


Private Function DeleteDetails() As Boolean

If MsgBox("Are you sure you want to permanently delete the data", vbInformation + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function

If m_UpdateStatus Then
    gDbTrans.BeginTrans
    If Not DeleteTransaction Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans
End If

MsgBox "Data Deleted", vbInformation

Call ClearPurchaseForm

End Function

Private Function InsertToInvoiceDetails() As Boolean

InsertToInvoiceDetails = False



InsertToInvoiceDetails = True


End Function

Public Property Let isUpdatingGrid(NewValue As Integer)
    m_isUpdatingGrid = NewValue
    
If m_frmPurchase Is Nothing Then Exit Property

With m_frmPurchase
    .cmbGroup.Locked = NewValue
    .cmbProductName.Locked = NewValue
    .cmbUnit.Locked = NewValue
End With
End Property
Public Property Get isUpdatingGrid() As Integer
    isUpdatingGrid = m_isUpdatingGrid
    
End Property

Private Function RemoveDataFromCollection() As Boolean
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim SelectedRow As Integer
Dim lngRelationID As Long
'Initialise the function
RemoveDataFromCollection = False

'Get the Selected row
With m_frmPurchase.grd
    SelectedRow = .Row
    .Col = 2
    'Get the RelationID from the selected row
    lngRelationID = .RowData(SelectedRow)
End With


'remove the item from the collection
Call m_colMaterials.Remove("MAT" & lngRelationID)


RemoveDataFromCollection = True

Exit Function

ErrLine:
    

End Function


Private Function AddToGridFromCollection() As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim MaxCount As Integer
Dim ItemCount As Integer

Dim lngRelationID As Long
Dim ParentID As Long

Dim strManufacturer As String
Dim ProductName As String

Dim Quantity As Double
'Dim FreeQuantity As Double

'Dim FreeAmount As Currency
Dim UnitPrice As Currency
Dim curMRP As Currency
Dim TotalAmount As Currency
Dim Amount As Currency
Dim TotalTaxAmount As Currency
Dim TaxAmount As Currency
Dim TRP As Currency
Dim InvoiceAmount  As Currency

Dim lKeys As Variant

Dim MaterialClass As clsMaterial

'Initlaise the function
AddToGridFromCollection = Failure

If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions

Set m_GrdFunctions.fGrd = m_frmPurchase.grd

'Get the data from collection
MaxCount = m_colMaterials.count

lKeys = m_colMaterials.GetKeys

TotalAmount = 0
'Add data to the grid
For ItemCount = 0 To MaxCount - 1
    
    Set MaterialClass = m_colMaterials.Item(lKeys(ItemCount))
    
    With m_frmPurchase.grd
        m_GrdFunctions.fMoreRows (2)
        .Row = ItemCount + 1
        .Col = 0: .Text = ItemCount + 1
        
'        strManufacturer = MaterialClass.Manufacturer
'        .Col = 1: .Text = ""
'        If strManufacturer <> "" Then .Text = strManufacturer
        
        lngRelationID = MaterialClass.RelationID
        .Col = 1: .Text = MaterialClass.ProductName: .RowData(ItemCount + 1) = lngRelationID: .CellFontBold = False
        
        Quantity = MaterialClass.Quantity
        .Col = 2: .Text = ""
        If Quantity > 0 Then .Text = Quantity
        
        UnitPrice = MaterialClass.UnitPrice
        .Col = 3: .Text = ""
        If UnitPrice > 0 Then .Text = UnitPrice
        
        curMRP = MaterialClass.MRP
        .Col = 4: .Text = FormatCurrency(MaterialClass.Amount)
        'If curMRP > 0 Then .Col = 6: .Text = curMRP
        
        TRP = MaterialClass.TradingPrice
        .Col = 5: .Text = ""
        If TRP > 0 Then .Text = TRP
        
        Amount = MaterialClass.Amount
        TaxAmount = MaterialClass.TaxAmount
        .Col = 6: .Text = TaxAmount: .CellFontBold = False
        
        .Col = 7: .Text = FormatCurrency(Amount + TaxAmount): .CellFontBold = False
         
        ParentID = MaterialClass.ParentID
        
        'Get the total amount
        Select Case ParentID
            Case 0 ' will be given zero for stock transactions
                TotalAmount = TotalAmount + Amount
                TotalTaxAmount = TotalTaxAmount + TaxAmount
            Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
                TotalAmount = TotalAmount - Amount
            Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
                TotalAmount = TotalAmount + Amount
        End Select
        
    End With
Next ItemCount
  
'Print the total
With m_frmPurchase.grd

    m_GrdFunctions.fMoreRows (2)
    .Row = .Row + 1
                
    .Col = 2: .Text = GetResourceString(286): .CellFontBold = True
    '.Col = 3: .Text = "Total Amount": .CellFontBold = True
    '.MergeCells = flexMergeFree
    .Col = 7: .Text = FormatCurrency(TotalAmount + TotalTaxAmount): .CellFontBold = True
    .Col = 6: .Text = TotalTaxAmount: .CellFontBold = True
    .Col = 4: .Text = TotalAmount: .CellFontBold = True
End With

With m_frmPurchase
    .TallyAmount = TotalAmount
    .lblTallyFigure.Caption = FormatCurrency(TotalAmount + TotalTaxAmount)
    .txtInvoiceDate.Locked = True
    .txtTransDate.Locked = True
    
    If TotalAmount = 0 Then
        If Val(.txtInvoiceAmount.Text) <> 0 Then
            .InitGrid
            .txtInvoiceDate.Locked = False
            .txtTransDate.Locked = True
        End If
    End If
End With


AddToGridFromCollection = Success

Exit Function

ErrLine:
    'Resume
    MsgBox "AddToGRidFromCollection " & vbCrLf & Err.Description
    AddToGridFromCollection = FatalError
    
End Function

Private Function LoadDataForUpdate() As Boolean
If isUpdatingGrid Then Exit Function

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim lngRelationID As Long
Dim strManufacter As String
Dim strProductName As String

Dim curTotal As Currency
Dim curAmount As Currency
Dim curPurchasePrice As Currency
Dim dblQuantity As Double
Dim rst As ADODB.Recordset

Dim GroupID As Integer
Dim UnitID As Integer
Dim MaxI As Integer
Dim I As Integer

LoadDataForUpdate = False

Dim MatClass As clsMaterial

'Get the required data from the purchase form
With m_frmPurchase
    lngRelationID = .grd.RowData(.grd.Row)
    If lngRelationID <= 0 Then GoTo Exit_Line
    
    gDbTrans.SqlStmt = "Select * From RelationMaster " & _
                " where RelationId = " & lngRelationID
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
    
    Set MatClass = m_colMaterials.Item("MAT" & lngRelationID)
    
    GroupID = FormatField(rst("GroupId"))
    UnitID = FormatField(rst("UnitId"))
    
    'Now Select The Group
    With .cmbGroup
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .ItemData(I) = GroupID Then .ListIndex = I: Exit For
        Next
    End With
    With .cmbUnit
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .ItemData(I) = UnitID Then .ListIndex = I: Exit For
        Next
    End With
    With .cmbProductName
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .ItemData(I) = lngRelationID Then Exit For
        Next
        If I = MaxI + 1 Then
            .AddItem MatClass.ProductName
            .ItemData(.newIndex) = lngRelationID
        End If
        .ListIndex = I
    End With
        
    .txtPrice.Text = MatClass.UnitPrice
    .txtQuantity.Text = MatClass.Quantity
    .txtAmount.Text = MatClass.Amount
    .txtTaxAmount = MatClass.TaxAmount
    isUpdatingGrid = .grd.Row
     
End With

Set MatClass = Nothing

LoadDataForUpdate = True

Exit_Line:
Exit Function

ErrLine:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"

End Function

Private Function AddDataToCollection() As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim lngRelationID As Long
Dim strManufacter As String
Dim strProductName As String
Dim strGroupName As String

Dim curTotal As Currency
Dim curAmount As Currency
Dim curTaxAmount As Currency
Dim curPurchasePrice As Currency
Dim dblQuantity As Double
Dim rst As ADODB.Recordset

AddDataToCollection = False

'Get the required data from the purchase form
With m_frmPurchase
    If .cmbGroup.ListIndex = -1 Then Exit Function
    If .cmbProductName.ListIndex = -1 Then Exit Function
    
    strManufacter = .lblCompanyName.Caption
    
    lngRelationID = .cmbProductName.ItemData(.cmbProductName.ListIndex)
    
    strProductName = .cmbProductName.Text
    strGroupName = .cmbGroup.Text
        
    curPurchasePrice = Val(.txtPrice.Text)
    dblQuantity = Val(.txtQuantity.Text)
    curAmount = Val(.txtAmount.Text)
    curTaxAmount = Val(.txtTaxAmount)
    
End With

If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials


If m_colMaterials.Exists("MAT" & lngRelationID) And isUpdatingGrid = 0 Then
    'Item already exists
    MsgBox GetResourceString(752), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Add to the collection
If isUpdatingGrid Then m_colMaterials.Remove ("MAT" & lngRelationID): isUpdatingGrid = 0
Call m_colMaterials.Add(dblQuantity, curPurchasePrice, 0, 0, lngRelationID, strManufacter, curAmount, curTaxAmount, strProductName, strGroupName, 0, 0, "MAT" & lngRelationID)

AddDataToCollection = True
Exit Function

ErrLine:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"

End Function


Private Function CreateProductProperty() As Long
' To be completed
On Error GoTo ErrLine

'Declare the Variables
Dim btGodownID As Byte
Dim intGroupID As Integer
Dim lngProductID As Long
Dim lngUnitID As Long
Dim lngRelationID As Long
Dim lngTransID As Long

Dim rstRelation As ADODB.Recordset

Dim curTradingprice As Currency
Dim curMRP As Currency
Dim curSalesPrice As Currency
'Dim curPurchasePrice As Currency
Dim OpStockValue As Currency

Dim dblOpStockBalance As Double
Dim isPriceChanged As Byte
Dim USTransDate As String
Dim oldRelnId As Long

Dim opVoucherType As Wis_VoucherTypes

'Initialise the function
CreateProductProperty = 0
'Exit the function if comboboxes are empty
With m_frmPurchase
    If .cmbGroup.ListIndex = -1 Then Exit Function
    If .cmbProductName.ListIndex = -1 Then Exit Function
    If .cmbUnit.ListIndex = -1 Then Exit Function

    With .cmbUnit
        If .ListIndex < 0 Then Exit Function
        lngRelationID = .ItemData(.ListIndex)
    End With
    oldRelnId = .cmbProductName.ItemData(.cmbProductName.ListIndex)

End With


gDbTrans.SqlStmt = "SELECT * FROM RelationMaster " & _
                    " Where RelationId = " & oldRelnId
Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)

isPriceChanged = 0
On Error GoTo ErrLine

USTransDate = GetSysFormatDate(FinIndianFromDate)
btGodownID = 1 'cmbBranch.itemData(cmbBranch.ListIndex)
With m_frmPurchase.cmbGroup
intGroupID = .ItemData(.ListIndex)
End With

lngUnitID = FormatField(rstRelation("UnitID"))

curTradingprice = Val(Trim$(m_frmPurchase.txtPrice.Text))
curMRP = FormatField(rstRelation("MRP"))
curSalesPrice = FormatField(rstRelation("SalesPrice"))
lngProductID = FormatField(rstRelation("ProductID"))
dblOpStockBalance = 0

lngTransID = 1

opVoucherType = Purchase
OpStockValue = 0 'curTradingprice * dblOpStockBalance

gDbTrans.SqlStmt = "SELECT MAX(RelationID) FROM RelationMaster "
         
Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)
lngRelationID = FormatField(rstRelation(0)) + 1

Set rstRelation = Nothing
   
gDbTrans.SqlStmt = "Update RelationMaster Set PriceChanged = " & 1 & _
            " Where RelationId = " & oldRelnId

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.SqlStmt = "INSERT INTO RelationMaster " & _
            "(RelationID,GodownID," & _
            " GroupID,ProductID,UnitID,TradingPrice," & _
            "MRP,SalesPrice, PriceChanged ) " & _
            " VALUES ( " & _
            lngRelationID & "," & _
            btGodownID & "," & _
            intGroupID & "," & _
            lngProductID & "," & _
            lngUnitID & "," & _
            curTradingprice & "," & _
            curMRP & "," & _
            curSalesPrice & "," & _
            isPriceChanged & " ) "

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

'If txtOpBalance.Enabled = True Then
    'Insert this OpBalance into stock table as its firstentry
    gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID,Quantity," & _
                       " UnitPrice,VoucherType,PurORSaleID,TransDate ) " & _
                       " VALUES ( " & _
                       lngRelationID & "," & _
                       lngTransID & "," & _
                       btGodownID & "," & _
                       dblOpStockBalance & "," & _
                       curTradingprice & "," & _
                       opVoucherType & "," & _
                       0 & "," & _
                       "#" & USTransDate & "# ) "
    
    'If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
       
'End If

gDbTrans.CommitTrans

CreateProductProperty = lngRelationID

'MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE

Exit Function

ErrLine:
   MsgBox "SaveProductPropertyDetails: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
   
End Function


Private Sub ClearPurchaseForm()
With m_frmPurchase
    .txtInvoiceDate.Locked = False
    .txtTransDate.Locked = False
    .txtInvoiceAmount.Text = ""
    .txtInvoiceNumber.Text = ""
    .txtInvoiceDate = gStrDate
    .txtTransDate = gStrDate
    .lblTallyFigure.Caption = ""
    .txtIncome.Text = ""
    .txtExpense.Text = ""
    .chkRedirect.Value = vbUnchecked
    .txtTax = ""
    .txtTaxAmount = ""
    .InitGrid
    
    ActivateTextBox .txtInvoiceDate
    .cmdAccept.Caption = GetResourceString(4) ' "&Accept"
    .DBOperation = Insert
    .cmbExpenseHead.ListIndex = -1
    .cmbIncomeHead.ListIndex = -1
    .cmdDelete.Enabled = False
    
End With

m_UpdateStatus = False
Set m_colMaterials = Nothing

End Sub

Private Function InsertToPurchase(ByVal RegCashTrans As Boolean) As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim VoucherType As Wis_VoucherTypes
Dim eInvoiceType As Wis_InvoiceType
Dim eRedirect As Wis_RedirectType

Dim RedirectHeadID As Long
Dim strTransTransID As String
Dim Trans_TransID As Long
Dim InvoiceNo As String
Dim InvoiceAmount As Currency
Dim USInvoiceDate As String
Dim USTransDate As String
Dim DiscountAmount As Currency
Dim TaxAmount  As Currency

Dim loopCount As Integer

Dim AccTransClass As clsAccTrans
Dim rst As ADODB.Recordset

'Initialise the function
InsertToPurchase = Failure

Set AccTransClass = New clsAccTrans
Dim headID As Long
headID = IIf(RegCashTrans, wis_CashHeadID, m_VendorID)
Trans_TransID = AccTransClass.GetMaxTransID(headID)

strTransTransID = Trans_TransID

If RegCashTrans Then
    strTransTransID = strTransTransID & "," & (Trans_TransID - 1)
Else
    If m_frmPurchase.optCash Then
        strTransTransID = AccTransClass.GetMaxTransID(wis_CashHeadID)
    End If
End If
If Val(m_frmPurchase.txtInvoiceAmount) = 0 Then strTransTransID = 0

Set AccTransClass = Nothing

VoucherType = Purchase

'Get the Max TransID
gDbTrans.SqlStmt = "SELECT MAX(TransID) FROM Purchase"

Call gDbTrans.Fetch(rst, adOpenForwardOnly)

m_PurTransID = FormatField(rst(0)) + 1

With m_frmPurchase
    InvoiceNo = Trim$(.txtInvoiceNumber)
    USInvoiceDate = GetSysFormatDate(.txtInvoiceDate)
    USTransDate = GetSysFormatDate(.txtTransDate)
    InvoiceAmount = Val(.txtInvoiceAmount)
    
    If .optROno.Value Then
        eInvoiceType = RONumber
    ElseIf .optSTANo.Value Then
        eInvoiceType = STANumber
    Else
        eInvoiceType = InvoiceNumber
    End If
    
    eRedirect = NonReDirected
    RedirectHeadID = 0
    If .chkRedirect.Value = vbChecked Then
        RedirectHeadID = .cmbRedirect.ItemData(.cmbRedirect.ListIndex)
        eRedirect = ReDirected
    End If
    
End With
'Insert these values into Purchase, and Stock tables
gDbTrans.SqlStmt = "INSERT INTO Purchase ( HeadID,TransID," & _
                  " InvoiceType,InvoiceNo,InvoiceDate,TransDate,InvoiceAmount," & _
                  " VoucherType,Trans_TransID,Redirected,RedirectHeadID) " & _
                  " VALUES ( " & _
                  m_VendorID & "," & _
                  m_PurTransID & "," & _
                  eInvoiceType & "," & _
                  "'" & InvoiceNo & "'," & _
                  "#" & USInvoiceDate & "#," & _
                  "#" & USTransDate & "#," & _
                  InvoiceAmount & "," & _
                  VoucherType & "," & _
                  AddQuotes(strTransTransID) & "," & _
                  eRedirect & "," & _
                  RedirectHeadID & " ) "

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

'Insert into InvoiceDetails
gDbTrans.SqlStmt = "INSERT INTO InvoiceDetails ( HeadID,PurchaseID ,InvoiceType, " & _
                " InvoiceNo,InvoiceDate,InvoiceAmount )" & _
                " VALUES ( " & _
                m_VendorID & "," & _
                m_PurTransID & "," & _
                eInvoiceType & "," & _
                "'" & InvoiceNo & "'," & _
                "#" & USInvoiceDate & "#," & _
                InvoiceAmount & " ) "

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

InsertToPurchase = Success

Exit Function

ErrLine:
    MsgBox "InsertToPurchase" & vbCrLf & Err.Description, vbCritical
    
End Function

Public Property Get ObjectName()
    ObjectName = "PurchaseClass"
End Property


Private Function InsertToStock() As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rst As ADODB.Recordset
Dim lpCount As Integer
Dim MaxCount As Integer
Dim RelationID As Long
Dim ParentID As Long
Dim TransID As Long
Dim Quantity As Double 'local copy
Dim Amount As Currency 'local copy
Dim TaxAmount As Currency 'local copy
Dim PurchasePrice As Currency
Dim USTransDate As String
Dim VoucherType As Wis_VoucherTypes
Dim FreeVoucherType As Wis_VoucherTypes

Dim lKeys As Variant

Dim MaterialClass As clsMaterial
Dim colMaterials As clsMaterials

'Initalise the function
InsertToStock = Failure

USTransDate = GetSysFormatDate(m_frmPurchase.txtTransDate)

MaxCount = m_colMaterials.count

Set MaterialClass = New clsMaterial

lKeys = m_colMaterials.GetKeys
VoucherType = Purchase
lpCount = 1

For lpCount = 0 To MaxCount - 1
    Set MaterialClass = m_colMaterials.Item(lKeys(lpCount))
    
    RelationID = MaterialClass.RelationID
    ParentID = MaterialClass.ParentID
    If ParentID = 0 Then
        TransID = 0
        
        TransID = MaterialClass.GetMaxTransID()
        
        Quantity = MaterialClass.Quantity
        
        Amount = MaterialClass.Amount
        TaxAmount = MaterialClass.TaxAmount
        
        TransID = TransID + 1
        
        PurchasePrice = MaterialClass.UnitPrice
                        
        gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID, " & _
                      " Quantity,UnitPrice,Amount,TaxAmount,TotalAmount,VoucherType," & _
                      " PurOrSaleID,TransDate ) " & _
                      " VALUES ( " & _
                      RelationID & "," & _
                      TransID & "," & _
                      1 & "," & _
                      Quantity & "," & _
                      PurchasePrice & "," & _
                      Amount & "," & TaxAmount & "," & (Amount + TaxAmount) & "," & _
                      VoucherType & "," & _
                      m_PurTransID & "," & _
                      "#" & USTransDate & "# ) "
        
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
          
    End If
Next lpCount

InsertToStock = Success

Exit Function

ErrLine:
    MsgBox "InsertToStock" & vbCrLf & Err.Description, vbCritical
    
End Function


Public Sub Show()
Dim MaterialClass As clsMaterial

If m_frmPurchase Is Nothing Then Set m_frmPurchase = New frmPurchase
If MaterialClass Is Nothing Then Set MaterialClass = New clsMaterial

m_frmPurchase.VendorID = m_VendorID
m_frmPurchase.lblCompanyName = MaterialClass.GetCompanyName(m_VendorID)
gWindowHandle = m_frmPurchase.hwnd
m_frmPurchase.Show 'vbModal

End Sub

Private Function Validated() As Boolean
'Declare the variables
Dim InvoiceAmount As Currency
Dim TalliedAmount As Currency

Validated = False

If m_frmPurchase Is Nothing Then Set m_frmPurchase = New frmPurchase

With m_frmPurchase
    If Not DateValidate(.txtInvoiceDate.Text, "/", True) Then
        'MsgBox "Invalid date specified!", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If

    If Not DateValidate(.txtTransDate.Text, "/", True) Then
       'MsgBox "Invalid date specified!", vbInformation, wis_MESSAGE_TITLE
       MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
       Exit Function
    End If
    If Not m_UpdateStatus Then
        If InvoiceNoExists(Trim$(.txtInvoiceNumber.Text)) Then
          ' MsgBox "Invoice Number already exists", vbInformation, wis_MESSAGE_TITLE
           MsgBox GetResourceString(679), vbInformation, wis_MESSAGE_TITLE
           ActivateTextBox .txtInvoiceNumber
           Exit Function
        End If
    End If
    
    If .txtInvoiceNumber.Text = "" Then Exit Function
    If Not CurrencyValidate(.txtInvoiceAmount.Text, True) Then Exit Function
    

    InvoiceAmount = Val(.txtInvoiceAmount.Text)
    TalliedAmount = Val(.lblTallyFigure.Caption)
    
    If Not CurrencyValidate(.txtExpense.Text, True) Then Exit Function
End With
'Check for the all the items in the invoice are added
If InvoiceAmount = 0 Then
    If MsgBox("Invoice Amount is Zero. " & vbCrLf & _
            "Do you want to enter zero amount Invoice?", vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
            
End If

If (InvoiceAmount \ 1) <> (TalliedAmount \ 1) Then Exit Function

Validated = True

End Function

'
Private Function InvoiceNoExists(ByVal InvoiceNo As String) As Boolean
'Declare the variables
Dim rst As Recordset

InvoiceNoExists = False
'Check the invoice no in the query
gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,TransDate FROM Purchase " & _
                   " WHERE InvoiceNO = '" & InvoiceNo & "'"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then InvoiceNoExists = True

End Function







Private Sub ClearControls()
With m_frmPurchase
    .cmbProductName.ListIndex = -1
    .txtPrice.Text = ""
    .txtQuantity.Text = ""
    .txtAmount.Text = ""
    
    .cmdUndo.Enabled = True
    .cmdAccept.Enabled = True
    .cmdDelete.Enabled = False
    
    On Error Resume Next
    .cmbGroup.SetFocus
    isUpdatingGrid = 0
     
End With
End Sub
    
    
Private Sub Class_Initialize()
'Initialise the objects
Set m_colMaterials = New clsMaterials

End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
If strSelection = "" Then Exit Sub
m_PurTransID = CLng(strSelection)
End Sub

Private Sub m_frmLookUp_SubItems(strSubItem() As String)
m_InvoiceNo = strSubItem(0)
m_TransDate = strSubItem(1)
m_InvoiceDate = strSubItem(2)
m_InvoiceAmount = strSubItem(3)
m_ReDirectHeadID = strSubItem(4)
m_InvoiceType = strSubItem(5)
m_TransIDString = strSubItem(UBound(strSubItem) - 1)
End Sub


Private Sub m_frmPurchase_AddClicked()
'Validation are done at the Command Add CLickevent  no need to validate
'Add the data to the collection then add to grid
If AddDataToCollection Then If AddToGridFromCollection <> Success Then Exit Sub
Call ClearControls

End Sub

Private Sub m_frmPurchase_ClearClicked()
Call ClearPurchaseForm
End Sub

Private Sub m_frmPurchase_DeleteClicked()
Call DeleteDetails

End Sub

Private Sub m_frmPurchase_GridClicked()
Call LoadDataForUpdate
End Sub

Private Sub m_frmPurchase_InvoiceClicked()
Call ClearPurchaseForm
Call LoadDetailsToUpdate
End Sub

Private Function LoadDetailsToUpdate() As Boolean
If Not GetPurchaseTableDetails Then Exit Function
If Not GetStockDetailsAddedToCollection Then Exit Function
If Not GetDetailsFromAccTrans Then Exit Function
Call GetCreditORCashFromAccTrans
If AddToGridFromCollection <> Success Then Exit Function
Call ClearControls

With m_frmPurchase
    .cmdAccept.Caption = GetResourceString(171) ' "&Update"
    .cmdDelete.Enabled = True
End With

m_UpdateStatus = True

End Function
Private Function GetDetailsFromAccTrans() As Boolean
'Declare the variables
Dim arrTransID As Variant
Dim loopCount As Integer

Dim rstAccTrans As ADODB.Recordset

Dim StrSql As String
Dim headID As Long
Dim ParentID As Long
Dim Amount As Currency
Dim headName As String

'Initialise the function
GetDetailsFromAccTrans = False

arrTransID = Split(m_TransIDString, ",")


For loopCount = LBound(arrTransID) To UBound(arrTransID) - 1
    StrSql = StrSql & "TransID =" & arrTransID(loopCount) & " OR "
Next loopCount

StrSql = StrSql & "TransID =" & arrTransID(loopCount)

gDbTrans.SqlStmt = " SELECT A.HeadID,A.Credit,A.Debit,B.HeadName,B.ParentID,AccountType" & _
                  " FROM AccTrans A,Heads B, ParentHeads C" & _
                  " WHERE ( " & StrSql & " )" & _
                  " AND B.ParentID=C.ParentID" & _
                  " AND A.HeadId=B.HeadID" & _
                  " AND ( C.AccountType = " & Loss & " OR C.AccountType = " & Profit & " )"

Call gDbTrans.Fetch(rstAccTrans, adOpenStatic)

Do While Not rstAccTrans.EOF

    Select Case rstAccTrans.Fields("AccountType")
        Case Profit
            Amount = FormatField(rstAccTrans.Fields("Credit"))
            headID = rstAccTrans.Fields("HeadID")
            headName = rstAccTrans.Fields("HeadName")
            ParentID = rstAccTrans.Fields("ParentID")
            Call m_colMaterials.Add(0, 0, 0, 0, headID, "", Amount, 0, headName, "", ParentID, 0, "MAT" & headID)
        Case Loss
            Amount = FormatField(rstAccTrans.Fields("Debit"))
            headID = rstAccTrans.Fields("HeadID")
            headName = rstAccTrans.Fields("HeadName")
            ParentID = rstAccTrans.Fields("ParentID")
            Call m_colMaterials.Add(0, 0, 0, 0, headID, "", Amount, 0, headName, "", ParentID, 0, "MAT" & headID)
    End Select
    
    rstAccTrans.MoveNext
Loop

With m_frmPurchase
    .optCredit.Value = True
    If UBound(arrTransID) > 0 Then .optCash.Value = True
    
End With
Set rstAccTrans = Nothing


GetDetailsFromAccTrans = True
End Function


Private Function GetCreditORCashFromAccTrans() As Boolean
'Declare the variables
Dim arrTransID As Variant
Dim loopCount As Integer

Dim rstAccTrans As ADODB.Recordset

Dim StrSql As String
Dim headID As Long

'Initialise the function
GetCreditORCashFromAccTrans = False

arrTransID = Split(m_TransIDString, ",")


For loopCount = LBound(arrTransID) To UBound(arrTransID) - 1
    StrSql = StrSql & "TransID =" & arrTransID(loopCount) & " OR "
Next loopCount

StrSql = StrSql & "TransID =" & arrTransID(loopCount)

gDbTrans.SqlStmt = " SELECT A.HeadID,A.Credit,A.Debit,B.HeadName,B.ParentID,AccountType" & _
                  " FROM AccTrans A,Heads B, ParentHeads C" & _
                  " WHERE ( " & StrSql & " )" & _
                  " AND B.ParentID=C.ParentID" & _
                  " AND A.HeadId=B.HeadID"

Call gDbTrans.Fetch(rstAccTrans, adOpenStatic)

m_frmPurchase.optCredit.Value = True
Do While Not rstAccTrans.EOF
    headID = rstAccTrans.Fields("HeadID")
    'Determine whether Transcation is credit or Cash
    If headID = wis_CashHeadID Then
        m_frmPurchase.optCash = True
        rstAccTrans.MoveLast
    End If
    rstAccTrans.MoveNext
Loop

Set rstAccTrans = Nothing


GetCreditORCashFromAccTrans = True
End Function

'This function will get Stock Details , by taking the ID of PurchaseTransID
' which is module level variable set after get the details from the Purchase table
Private Function GetStockDetailsAddedToCollection() As Boolean
'Declare the variables
Dim PurVoucher As Wis_VoucherTypes
Dim rstPurchase As ADODB.Recordset
Dim rstRelation As ADODB.Recordset

Dim Quantity As Double
Dim PurPrice As Currency
Dim RelationID As Long
Dim Amount As Currency
Dim TaxAmount As Currency
Dim TradePrice As Currency
Dim MRP As Currency
Dim headName As String
Dim ProductName As String
Dim ProductGroupName As String

Dim fldRelationID As ADODB.Field
Dim fldQuantity As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldTaxAmount As ADODB.Field
Dim fldGroup As ADODB.Field

'Trap an error
On Error GoTo ErrLine

'Iniitalse the function
GetStockDetailsAddedToCollection = False

'Validate the module level variables
If m_PurTransID = 0 Then Exit Function

PurVoucher = Purchase

gDbTrans.SqlStmt = " SELECT RelationID,Quantity AS QtyPurchase,UnitPrice,TaxAmount" & _
                   " FROM Stock " & _
                   " WHERE PurOrSaleID = " & m_PurTransID & _
                   " AND VoucherType=" & PurVoucher

Call gDbTrans.CreateView("QryPurchase")

gDbTrans.SqlStmt = " SELECT * FROM QryPurchase"
                   
If gDbTrans.Fetch(rstPurchase, adOpenForwardOnly) < 1 Then Exit Function


If Not GetRstFromRelationMaster(rstRelation) Then Exit Function

Set fldRelationID = rstPurchase.Fields("RelationID")
Set fldQuantity = rstPurchase.Fields("QtyPurchase")
Set fldUnitPrice = rstPurchase.Fields("UnitPrice")
Set fldTaxAmount = rstPurchase.Fields("TaxAmount")


If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials

Do While Not rstPurchase.EOF
    RelationID = fldRelationID.Value
    Quantity = fldQuantity.Value
    PurPrice = fldUnitPrice.Value
    TaxAmount = FormatField(fldTaxAmount)
    
    Amount = Quantity * PurPrice
    If Val(m_frmPurchase.txtInvoiceAmount.Text) = 0 Then Amount = 0
    
    TradePrice = 0: MRP = 0: headName = "": ProductName = "": ProductGroupName = ""
    
    If gDbTrans.FindRecord(rstRelation, "RelationID= " & RelationID) Then
        TradePrice = FormatField(rstRelation.Fields("TradingPrice"))
        MRP = FormatField(rstRelation.Fields("MRP"))
'        HeadName = FormatField(rstRelation.Fields("HeadName"))
        ProductName = FormatField(rstRelation.Fields("ProductName"))
        ProductGroupName = FormatField(rstRelation.Fields("GroupName"))
    End If
    
    'Add To the collection
    Call m_colMaterials.Add(Quantity, PurPrice, TradePrice, MRP, RelationID, headName, Amount, TaxAmount, ProductName, ProductGroupName, 0, 0, "MAT" & RelationID)
    
    'Move the recordset
    rstPurchase.MoveNext
Loop

Set rstRelation = Nothing
Set rstPurchase = Nothing

GetStockDetailsAddedToCollection = True

Exit Function

ErrLine:
    If Err Then MsgBox "GetStockeDetailsAddedToCollection" & vbCrLf & Err.Description
    

End Function


Private Function GetRstFromRelationMaster(rstRelation As ADODB.Recordset) As Boolean

GetRstFromRelationMaster = False

'gDbTrans.SqlStmt = " SELECT A.RelationID,TradingPrice,MRP,ProductName " & _
                   " FROM RelationMaster AS A,Products C,QryPurchase D " & _
                   " WHERE A.RelationID =D.RelationID " & _
                   " AND A.ProductID=C.ProductID "
gDbTrans.SqlStmt = " SELECT A.RelationID,TradingPrice,MRP,ProductName,GroupName " & _
                   " FROM RelationMaster AS A,Products C,QryPurchase D,ProductGroup E " & _
                   " WHERE A.RelationID =D.RelationID " & _
                   " AND A.ProductID=C.ProductID " & _
                   " AND A.GroupID=E.GroupID "

If gDbTrans.Fetch(rstRelation, adOpenStatic) < 1 Then Exit Function

GetRstFromRelationMaster = True
End Function


'This function will get the details from the Purchase table
' where module level variables are set
' Here : m_PurTransID will be set which used to get the details from Stock Table
'      " m_TransIDString will be set which used to get the details from AccTrans Table
Private Function GetPurchaseTableDetails() As Boolean
'Declare the variables
Dim rstInvoices As ADODB.Recordset
Dim USFromDate As String
Dim USToDate As String
Dim PurVoucher As Wis_VoucherTypes
Dim FillViewClass As clsFillView
Dim loopCount As Integer

GetPurchaseTableDetails = False

With m_frmPurchase
    USFromDate = GetSysFormatDate(FinIndianFromDate)
    USToDate = GetSysFormatDate(.txtInvoiceDate.Text)
End With

PurVoucher = Purchase

gDbTrans.SqlStmt = " SELECT TransID,InvoiceNo,TransDate,InvoiceDate," & _
                   " InvoiceAmount,RedirectHeadID,InvoiceType,Trans_TransID" & _
                   " FROM Purchase " & _
                   " WHERE HeadID = " & m_VendorID & _
                   " AND InvoiceDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USToDate & "#" & _
                   " AND  VoucherType = " & PurVoucher & _
                   " ORDER BY InvoiceDate,TransDate "

If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Set FillViewClass = New clsFillView

m_InvoiceNo = ""

If Not FillViewClass.FillViewForPurchase(m_frmLookUp.lvwReport, rstInvoices, "TransID", True) Then Exit Function
m_frmLookUp.Show vbModal

Set FillViewClass = Nothing
If m_InvoiceNo = "" Then Exit Function

With m_frmPurchase
    .txtInvoiceNumber = m_InvoiceNo
    .txtInvoiceAmount = m_InvoiceAmount
    .txtInvoiceDate = m_InvoiceDate
    .txtTransDate = m_TransDate
    Select Case m_InvoiceType
        Case InvoiceNumber
            .optInvoiceNo.Value = True
        Case RONumber
            .optROno.Value = True
        Case STANumber
            .optSTANo.Value = True
    End Select
    If m_ReDirectHeadID > 0 Then
        .chkRedirect.Value = vbChecked
        For loopCount = 0 To .cmbRedirect.ListCount - 1
            If m_ReDirectHeadID = .cmbRedirect.ItemData(loopCount) Then
                .cmbRedirect.ListIndex = loopCount
                Exit For
            End If
        Next loopCount
    End If
End With

GetPurchaseTableDetails = True


End Function

Private Sub m_frmPurchase_OkClicked()

'Validate the data
If Not Validated Then Exit Sub

If m_frmPurchase Is Nothing Then Set m_frmPurchase = New frmPurchase

'Save the data to the database
If SaveDetails Then ClearPurchaseForm
End Sub



Private Function SaveDetails() As wis_FunctionReturned
'Declare the Variables
Dim RowNum As Integer
Dim lpCount As Integer
Dim lpCount_2 As Integer
Dim MaxCount As Integer

Dim ProductHeadID() As Long
Dim ProductAmount() As Currency
Dim TaxAmount() As Currency
Dim taxHeadID As Long
Dim transHeadID As Long
Dim TotalTaxAmount As Currency
Dim TotalProductAmount As Currency

Dim headID As Long
Dim arrHeadID() 'Two Dimensional Array
Dim arrParentID()
Dim PLIDCount As Integer
Dim RelationID As Long
Dim ParentID As Long
Dim TransDate As Date

Dim InvoiceAmount As Currency
Dim TotalAmount As Currency

Dim PurchaseAmount As Currency
Dim Amount As Currency

Dim GroupName As String
Dim GroupNameEnglish As String
Dim IsCashPurchase As Boolean

Dim AccTransClass As clsAccTrans
Dim LedgerClass As clsLedger
Dim MaterialClass As clsMaterial
Dim colLedgers As clsLedgers

Dim lKeys As Variant
Dim InTrans As Boolean

'Trap an error
On Error GoTo ErrLine

If m_UpdateStatus Then
    gDbTrans.BeginTrans
    InTrans = True
    If Not DeleteTransaction Then GoTo Exit_Line
End If

With m_frmPurchase
    TransDate = GetSysFormatDate(.txtTransDate)
    GroupName = .cmbGroup.Text
    'GroupNameEnglish = .cmbGroupEnglish.ItemData(.cmbGroup.ListIndex)
    IsCashPurchase = False
    If .optCash.Value Then IsCashPurchase = True
    If .optCredit.Value = False Then IsCashPurchase = True
End With

If IsCashPurchase Then
    Dim RegCashTrans As Boolean
    Dim SetupClass As clsSetup
    Set SetupClass = New clsSetup
    RegCashTrans = IIf(UCase(SetupClass.ReadSetupValue("Trading", "RegisterCashTransaction", "False")) = "TRUE", True, False)
    Set SetupClass = Nothing
End If

'Initialise the objects
Set AccTransClass = New clsAccTrans
Set LedgerClass = New clsLedger
Set colLedgers = New clsLedgers
Set MaterialClass = New clsMaterial

If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials

'Initialise the function
SaveDetails = Failure

MaxCount = m_colMaterials.count

lKeys = m_colMaterials.GetKeys

PLIDCount = 0

'Begin Transaction
If Not InTrans Then gDbTrans.BeginTrans: InTrans = True
ReDim ProductHeadID(0)
ReDim ProductAmount(0)
ReDim TaxAmount(0)

For lpCount = 0 To MaxCount - 1
    Set MaterialClass = m_colMaterials.Item(lKeys(lpCount))
    Amount = MaterialClass.Amount
    ParentID = MaterialClass.ParentID
    'Get the total amount
    Select Case ParentID
        Case 0 ' will be given zero for stock transactions
            
            TotalAmount = TotalAmount + Amount
            TotalTaxAmount = TotalTaxAmount + MaterialClass.TaxAmount
            PurchaseAmount = PurchaseAmount + Amount
            
            ProductHeadID(UBound(ProductHeadID)) = LedgerClass.GetHeadIDCreated(wis_PurchaseParentID, MaterialClass.ProductGroupName & " " & GetResourceString(176), 0, "")
            ProductAmount(UBound(ProductAmount)) = MaterialClass.Amount
            TaxAmount(UBound(TaxAmount)) = MaterialClass.TaxAmount
            
            ReDim Preserve ProductHeadID(UBound(ProductHeadID) + 1)
            ReDim Preserve ProductAmount(UBound(ProductAmount) + 1)
            ReDim Preserve TaxAmount(UBound(TaxAmount) + 1)
       
         Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
         
            TotalAmount = TotalAmount - Amount
            ReDim Preserve arrHeadID(1, PLIDCount)
            ReDim Preserve arrParentID(PLIDCount)
            arrHeadID(0, PLIDCount) = MaterialClass.RelationID 'Headid
            arrHeadID(1, PLIDCount) = MaterialClass.Amount
            arrParentID(PLIDCount) = ParentID
            PLIDCount = PLIDCount + 1
        
         Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
         
            TotalAmount = TotalAmount + Amount
            ReDim Preserve arrHeadID(1, PLIDCount)
            ReDim Preserve arrParentID(PLIDCount)
            arrHeadID(0, PLIDCount) = MaterialClass.RelationID 'Headid
            arrHeadID(1, PLIDCount) = MaterialClass.Amount
            arrParentID(PLIDCount) = ParentID
            PLIDCount = PLIDCount + 1
    End Select
Next lpCount

'August 2015
''Now Check for the duplicate Procut Items
If UBound(ProductHeadID) > 1 Then
    MaxCount = UBound(ProductHeadID) - 1
    For lpCount = 0 To MaxCount
        headID = ProductHeadID(0)
        For lpCount_2 = lpCount + 1 To MaxCount
            If headID = ProductHeadID(lpCount_2) Then
                ProductAmount(lpCount) = ProductAmount(lpCount) + ProductAmount(lpCount_2)
                ProductHeadID(lpCount_2) = 0
                ProductAmount(lpCount_2) = 0
            End If
        Next
    Next
End If
'END August 2015

'If he is entering Zero Amount Stock
'i.e Stock has to be increased but not the amount
If TotalAmount = 0 Then
    If Not InTrans Then gDbTrans.BeginTrans: InTrans = True
    If InsertToPurchase(False) <> Success Then GoTo Exit_Line
    If InsertToStock <> Success Then GoTo Exit_Line
    
    gDbTrans.CommitTrans
    InTrans = False
    SaveDetails = Success
    'MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


'Get the Purchase Account HeadID
'headID = LedgerClass.GetHeadIDCreated(wis_PurchaseParentID, GroupName & " " & GetResourceString(176), 0, GroupNameEnglish & "  " & LoadResString(176))  'In Kannada it is Kharidi

lpCount = 0
transHeadID = IIf(RegCashTrans Or (Not IsCashPurchase), m_VendorID, wis_CashHeadID)
'Call AccTransClass.AllTransHeadsAdd(transHeadId, 0, TotalAmount)
'Call AccTransClass.AllTransHeadsAdd(transHeadId, PurchaseAmount, 0)

'NEW LOGIC ADDED ON AUGUST 2015
MaxCount = UBound(ProductAmount) - 1
For lpCount = 0 To MaxCount

    ''We need to add only one Trans Per HeadID, so Adding the total amoount as the first iteration
    If lpCount = 0 Then Call AccTransClass.AllTransHeadsAdd(transHeadID, 0, TotalAmount + TotalTaxAmount)
    
    ''IF THERE IS HEADID THEN
    If (ProductHeadID(lpCount)) > 0 Then _
        Call AccTransClass.AllTransHeadsAdd(ProductHeadID(lpCount), ProductAmount(lpCount), 0)
    
Next lpCount

'' Add the Tax Amount
If TotalTaxAmount > 0 Then
    taxHeadID = LedgerClass.GetHeadIDCreated(parReceivable, GetResourceString(162, 173), 0, LoadResourceStringS(176, 173))
    Call AccTransClass.AllTransHeadsAdd(taxHeadID, TotalTaxAmount, 0)
End If
'END NEW LOGIC ADDED ON AUGUST 2015

lpCount = 0
If PLIDCount > 0 Then
    For lpCount = 0 To PLIDCount - 1
        'Get the total amount
        Select Case arrParentID(lpCount)
            Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
                Call AccTransClass.AllTransHeadsAdd(arrHeadID(0, lpCount), 0, arrHeadID(1, lpCount))
            Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
                Call AccTransClass.AllTransHeadsAdd(arrHeadID(0, lpCount), arrHeadID(1, lpCount), 0)
        End Select
    Next lpCount
End If

If AccTransClass.SaveVouchers(Purchase, TransDate, , True) <> Success Then GoTo Exit_Line

If RegCashTrans Then
    'Call AccTransClass.AllTransHeadsAdd(wis_CashHeadID, 0, TotalAmount)
    'Call AccTransClass.AllTransHeadsAdd(m_VendorID, TotalAmount, 0)
    Call AccTransClass.AllTransHeadsAdd(wis_CashHeadID, 0, TotalAmount + TotalTaxAmount)
    Call AccTransClass.AllTransHeadsAdd(m_VendorID, TotalAmount + TotalTaxAmount, 0)
    If AccTransClass.SaveVouchers(payment, TransDate, , True) <> Success Then GoTo Exit_Line
End If

' Insert the purchase details of VendorID(HeadID) by Invoice Amount
' In the Purchase table
If InsertToPurchase(RegCashTrans) <> Success Then Exit Function

' Insert the Qunatities of each product into Stock table and update the
' stock balance accordingly
If InsertToStock <> Success Then Exit Function

gDbTrans.CommitTrans: InTrans = False

Call ClearControls
SaveDetails = Success

'MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE


Exit_Line:
If InTrans Then gDbTrans.RollBack
Exit Function

ErrLine:
    MsgBox "SaveDetails:" & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    SaveDetails = FatalError

End Function

Private Function DeleteTransaction() As Boolean


DeleteTransaction = False

If Not DeletePurchaseTableDetails Then Exit Function
If Not DeleteStockTableDetails Then Exit Function
If Not DeleteAccTransDetails Then Exit Function
DeleteTransaction = True

End Function

Private Function DeleteAccTransDetails() As Boolean
'Declare the variables
Dim arrTransID As Variant
Dim StrSql As String
Dim loopCount As Integer

DeleteAccTransDetails = False

If m_TransIDString = "" Then Exit Function
If m_PurTransID <= 0 Then Exit Function

arrTransID = Split(m_TransIDString, ",")

For loopCount = LBound(arrTransID) To UBound(arrTransID) - 1
    StrSql = StrSql & "TransID =" & arrTransID(loopCount) & " OR "
Next loopCount

StrSql = StrSql & "TransID =" & arrTransID(loopCount)

gDbTrans.SqlStmt = " DELETE *" & _
                  " FROM AccTrans" & _
                  " WHERE ( " & StrSql & " )"

If Not gDbTrans.SQLExecute Then Exit Function

gDbTrans.SqlStmt = " DELETE *" & _
                  " FROM IDFromInventory" & _
                  " WHERE ( " & StrSql & " )"

If Not gDbTrans.SQLExecute Then Exit Function

DeleteAccTransDetails = True

End Function


Private Function DeleteStockTableDetails() As Boolean
'Declare the variables
Dim USTransDate As String
Dim rst As Recordset

DeleteStockTableDetails = False
If m_InvoiceNo = "" Then Exit Function
If m_TransDate = "" Then Exit Function
If m_PurTransID <= 0 Then Exit Function

USTransDate = GetSysFormatDate(m_TransDate)

gDbTrans.SqlStmt = " DELETE *" & _
                   " FROM Stock " & _
                   " WHERE PurORSaleID = " & m_PurTransID & _
                   " AND TransDate =#" & USTransDate & "#"

If Not gDbTrans.SQLExecute Then Exit Function

DeleteStockTableDetails = True
End Function



Private Function DeletePurchaseTableDetails() As Boolean
'Declarew the variables
DeletePurchaseTableDetails = False

If m_InvoiceNo = "" Then Exit Function
If m_TransDate = "" Then Exit Function
If m_VendorID = 0 Then Exit Function
If m_PurTransID = 0 Then Exit Function

gDbTrans.SqlStmt = " DELETE * " & _
                   " FROM Purchase " & _
                   " WHERE InvoiceNo = " & AddQuotes(m_InvoiceNo)
                   
If Not gDbTrans.SQLExecute Then Exit Function

gDbTrans.SqlStmt = " DELETE *" & _
                   " FROM InvoiceDetails" & _
                   " WHERE HeadID = " & m_VendorID & _
                   " AND PurchaseID =" & m_PurTransID
                   
If Not gDbTrans.SQLExecute Then Exit Function

DeletePurchaseTableDetails = True

End Function


Public Property Get VendorID() As Long
VendorID = m_VendorID
End Property

Public Property Let VendorID(ByVal vNewValue As Long)
m_VendorID = vNewValue
End Property

Private Sub m_frmPurchase_PLHeadsClicked(ByVal AccType As wis_AccountType)
'Validation are done at the Command Add CLickevent  no need to validate
'Add the data to the collection then add to grid
If AddPLHeadsToCollection(AccType) Then Call AddToGridFromCollection


End Sub

Private Sub m_frmPurchase_UnDoClicked()
If RemoveDataFromCollection Then
    m_frmPurchase.InitGrid
    AddToGridFromCollection
End If
End Sub


Private Sub m_frmPurchase_WindowClosed()

Set m_frmPurchase = Nothing
End Sub


Public Property Get ObjectHandle() As Long

    If Not m_frmPurchase Is Nothing Then _
                ObjectHandle = m_frmPurchase.hwnd

End Property

