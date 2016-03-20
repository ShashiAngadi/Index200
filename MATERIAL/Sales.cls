VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_frmSales As frmSales
Attribute m_frmSales.VB_VarHelpID = -1
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_dbOperation As wis_DBOperation
Private m_colMaterials As clsMaterials
Private m_GrdFunctions As clsGrdFunctions

Private m_VendorID As Long
Private m_RelationID As Long
Private m_SalesTransID As Long

Private m_isUpdatingGrid As Integer

Private m_UpdateStatus As Boolean
Private m_TransIDString As String 'From Purchase Table
Private m_InvoiceNo As String
Private m_TransDate As String
Private m_InvoiceAmount As Currency

Public Property Let isUpdatingGrid(NewValue As Integer)
    m_isUpdatingGrid = NewValue
If m_frmSales Is Nothing Then Exit Property

With m_frmSales
    .cmbGroup.Locked = NewValue
    .cmbProductName.Locked = NewValue
    .cmbUnit.Locked = NewValue
End With

End Property

Public Property Get isUpdatingGrid() As Integer
    isUpdatingGrid = m_isUpdatingGrid
End Property


Private Sub ClearControls()
'cmbGroup.ListIndex = -1
With m_frmSales
    isUpdatingGrid = 0
    .cmbProductName.ListIndex = -1
    .txtPrice.Text = ""
    .txtQuantity.Text = ""
    .txtAmount.Text = ""
    .txtSoot.Text = ""
    .cmdUndo.Enabled = True
    .cmdPrint.Enabled = True
    .cmdAccept.Enabled = True
    .cmdDelete.Enabled = False
    On Error Resume Next
    .cmbGroup.SetFocus
    .cmbBranch.Locked = True
    
    .lblTaxAmount = GetResourceString(173)
    .txtAmount.Tag = ""
    .txtAmount.Text = ""
End With

End Sub


Private Function AddDataToCollection() As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim lngRelationID As Long

Dim strManufacter As String
Dim strProductName As String
Dim strGroupName As String

Dim curTotal As Currency
Dim curTradingprice As Currency
Dim curMRP As Currency
Dim curAmount As Currency
Dim curTaxAmount As Currency

Dim curSalesPrice As Currency

Dim dblQuantity As Double
Dim eSoot As Wis_VoucherTypes
Dim Soot As Double
Dim GodownID As Byte
Dim rst As ADODB.Recordset

AddDataToCollection = False

'Get the required data from the sales form

With m_frmSales
    If .cmbProductName.ListIndex = -1 Then Exit Function
    If .cmbBranch.Enabled Then If .cmbBranch.ListIndex = -1 Then Exit Function
    
    lngRelationID = .cmbProductName.ItemData(.cmbProductName.ListIndex)
    GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
    
    strProductName = .cmbProductName.Text
    strGroupName = .cmbGroup.Text
    'strManufacter = .cmbManufacturer.Text
    curSalesPrice = Val(.txtPrice.Text)
    dblQuantity = Val(.txtQuantity.Text)
    Soot = Val(.txtSoot.Text)
    curAmount = Val(.txtAmount.Text)
    curTaxAmount = Val(.txtTaxAmount)
End With

'Get Trading Price & MRP from Data base
gDbTrans.SqlStmt = "SELECT TradingPrice,MRP FROM RelationMaster WHERE RelationID = " & lngRelationID
  
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

curTradingprice = FormatField(rst("TradingPrice"))
curMRP = FormatField(rst("MRP"))

'curAmount = dblQuantity * curSalesPrice

If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials

If m_colMaterials.Exists("MAT" & lngRelationID) And isUpdatingGrid = 0 Then
    MsgBox GetResourceString(752), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Add to the collection
If isUpdatingGrid Then _
    m_colMaterials.Remove ("MAT" & lngRelationID): isUpdatingGrid = 0
Call m_colMaterials.Add(dblQuantity, curSalesPrice, curTradingprice, _
                         curMRP, lngRelationID, strManufacter, curAmount, curTaxAmount, _
                        strProductName, strGroupName, 0, Soot, "MAT" & lngRelationID)
 
AddDataToCollection = True
Exit Function

ErrLine:
    If Err Then _
        MsgBox "AddDataToCollection: " & vbCrLf & Err.Description
    
End Function
Private Function AddToGridFromCollection() As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim MaxCount As Integer
Dim ItemCount As Integer

Dim lngRelationID As Long
Dim ParentID As Long

Dim lKeys As Variant

'Decalre the variables for the grid
Dim strManufacturer As String
Dim ProductName As String
Dim Quantity As Double
Dim UnitPrice As Currency
Dim curMRP As Currency
Dim TotalAmount As Currency
Dim TotalTaxAmount As Currency
Dim Amount As Currency
Dim TaxAmount As Currency
Dim QtySoot As Double


Dim MaterialClass As clsMaterial

'Initlaise the function
AddToGridFromCollection = Failure

If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions

Set m_GrdFunctions.fGrd = m_frmSales.grd

'Get the data from collection
MaxCount = m_colMaterials.count

lKeys = m_colMaterials.GetKeys

TotalAmount = 0
'Add data to the grid
For ItemCount = 0 To MaxCount - 1
    Set MaterialClass = m_colMaterials.Item(lKeys(ItemCount))
    With m_frmSales.grd
    
        m_GrdFunctions.fMoreRows (2)
        .Row = ItemCount + 1
        '.RowData (mid(m_colMaterials(0).
        .Col = 0: .Text = ItemCount + 1
        
        lngRelationID = MaterialClass.RelationID
        .Col = 1: .Text = MaterialClass.ProductName: .RowData(ItemCount + 1) = lngRelationID: .CellFontBold = False
        
        Quantity = MaterialClass.Quantity
        .Col = 2: .Text = Quantity '""
        'If Quantity > 0 Then .Text = Quantity
        
        UnitPrice = MaterialClass.UnitPrice
        .Col = 3: .Text = UnitPrice '""
        'If UnitPrice > 0 Then .Text = UnitPrice
        
        curMRP = MaterialClass.MRP
        .Col = 4: .Text = FormatCurrency(MaterialClass.Amount)
        'If curMRP > 0 Then .Col = 6: .Text = curMRP
        
        QtySoot = MaterialClass.Soot
        .Col = 5: .Text = ""
        If QtySoot > 0 Then .Text = QtySoot
        
        Amount = MaterialClass.Amount
        '.Col = 6: .Text = Amount: .CellFontBold = False
        
        TaxAmount = MaterialClass.TaxAmount
        .TextMatrix(.Row, 6) = TaxAmount
        
        .TextMatrix(.Row, 7) = FormatCurrency(Amount + TaxAmount)
         
        ParentID = MaterialClass.ParentID
        
        'Get the total amount
        Select Case ParentID
            Case 0 ' will be given zero for stock transactions
                TotalAmount = TotalAmount + Amount + TaxAmount
            Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
                TotalAmount = TotalAmount + Amount
            Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
                TotalAmount = TotalAmount - Amount
        End Select
        
    End With
Next ItemCount
  
'Print the total
With m_frmSales.grd

    m_GrdFunctions.fMoreRows (2)
    .Row = .Row + 1
            
    .Col = 1: .CellFontBold = True: .Text = GetResourceString(286) '52  Text Amount
    .Col = 7: .Text = TotalAmount: .CellFontBold = True
    
End With

Debug.Assert 1 = 0
''Allow user to add only soot value
'so he can deduct that much quantity from stock


'If TotalAmount = 0 Then m_frmSales.InitGrid
'If TotalAmount = 0 And Val(m_frmSales.txtSoot) = 0 Then m_frmSales.InitGrid

With m_frmSales
    .lblTallyFigure.Caption = (TotalAmount)
End With


AddToGridFromCollection = Success

Exit Function

ErrLine:
    MsgBox "AddToGRidFromCollection " & vbCrLf & Err.Description
    AddToGridFromCollection = FatalError
    

End Function

Private Sub ClearSalesForm()
With m_frmSales
    '.txtInvoiceDate.Text = ""
    .InitGrid
    
'    .tabSales.Tabs(1).Selected = True
    .cmbBranch.Locked = False
    .cmbBranch.SetFocus
    .cmdAccept.Caption = GetResourceString(4) '"&Accept"
    .cmdDelete.Enabled = False
    .lblTallyFigure.Caption = ""
    .cmbExpenseHead.ListIndex = -1
    .cmbIncomeHead.ListIndex = -1
    .txtExpense.Text = ""
    .txtIncome.Text = ""
    .txtTaxAmount.Tag = ""
    .txtTaxAmount = ""
    .lblTaxAmount = GetResourceString(173)
    .optCash.Value = False
    .optCredit.Value = False
    isUpdatingGrid = 0
End With
    
Set m_colMaterials = Nothing

m_UpdateStatus = False
    
End Sub

Private Sub frmGridClicked()

If isUpdatingGrid Then Exit Sub

Dim RelationID As Long
Dim MatClass As New clsMaterial

With m_frmSales.grd
    RelationID = .RowData(.Row)
    If RelationID <= 0 Then Exit Sub
    Set MatClass = m_colMaterials.Item("MAT" & RelationID)
End With

Dim I As Integer
Dim MaxI As Integer

'94482 45854 Babu Bagalkot
Dim GroupID As Long
Dim UnitID As Integer

Dim rst As Recordset
gDbTrans.SqlStmt = "select * From RelationMaster " & _
            " Where RelationID = " & RelationID
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

UnitID = rst("UnitID")
GroupID = rst("GroupID")

With m_frmSales
    'Set The Group name
    With .cmbGroup
        MaxI = .ListCount
        I = 0
        Do
            If .ItemData(I) = GroupID Then Exit Do
            I = I + 1
            If I >= MaxI Then Exit Do
        Loop
        If I < MaxI Then .ListIndex = I
    End With
    'Select the Unit Name
    With .cmbUnit
        MaxI = .ListCount
        I = 0
        Do
            If .ItemData(I) = UnitID Then Exit Do
            I = I + 1
            If I >= MaxI Then Exit Do
        Loop
        If I < MaxI Then .ListIndex = I
    End With
    With .cmbProductName
        MaxI = .ListCount
        I = 0
        Do
            If .ItemData(I) = RelationID Then Exit Do
            I = I + 1
        Loop
        If I = MaxI Then
            .AddItem MatClass.ProductName
            .ItemData(.newIndex) = RelationID
        End If
        .ListIndex = I
    End With
    
    'set the Slaes Price
    .txtPrice = MatClass.UnitPrice
    .txtQuantity = MatClass.Quantity
    .txtAmount = MatClass.Amount
    .txtTaxAmount = MatClass.TaxAmount
    .txtSoot = MatClass.Soot
    isUpdatingGrid = .grd.Row
End With
End Sub

Private Function InsertToSales(ByVal RegCashTrans As Boolean) As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim VoucherType As Wis_VoucherTypes

Dim Trans_TransID As Long
Dim strTransID As String
Dim InvoiceNo As String
Dim InvoiceAmount As Currency
Dim USInvoiceDate As String
Dim GodownID As Byte

Dim rst As ADODB.Recordset

Dim AccTransClass As clsAccTrans

'New code on 21/10/2004
'THis code is done for testing which might seems to be correct
'For further details go through the calling function Save Details

Dim SetupClass  As New clsSetup

RegCashTrans = RegCashTrans Or (IIf(UCase(SetupClass.ReadSetupValue("Trading", "RegisterCashTransaction", "False")) = "TRUE", True, False))
Set SetupClass = Nothing
''nnew code ends here

'Initialise the function
InsertToSales = Failure

Set AccTransClass = New clsAccTrans

'Trans_TransID = AccTransClass.GetMaxTransID(IIf(RegCashTrans, wis_CashHeadID, m_VendorID))
Trans_TransID = AccTransClass.GetMaxTransID(IIf(RegCashTrans, m_VendorID, wis_CashHeadID))

strTransID = Trans_TransID

If RegCashTrans Then strTransID = strTransID & "," & (Trans_TransID - 1)

Set AccTransClass = Nothing

'Get the Max TransID
gDbTrans.SqlStmt = "SELECT MAX(TransID) FROM Sales"

Call gDbTrans.Fetch(rst, adOpenForwardOnly)

m_SalesTransID = FormatField(rst(0)) + 1

With m_frmSales
    InvoiceNo = Trim$(.txtInvoiceNumber)
    USInvoiceDate = GetSysFormatDate(.txtInvoiceDate)
    If .cmbBranch.ListIndex = -1 Then Exit Function
    GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
End With

VoucherType = Sales

'Insert these values into Purchase, and Stock tables
gDbTrans.SqlStmt = " INSERT INTO Sales (HeadID,TransID,GodownID," & _
                  " InvoiceNo,InvoiceDate,InvoiceAmount,VoucherType,Trans_TransID) " & _
                  " VALUES ( " & _
                  m_VendorID & "," & _
                  m_SalesTransID & "," & _
                  GodownID & "," & _
                  "'" & InvoiceNo & "'," & _
                  "#" & USInvoiceDate & "#," & _
                  m_InvoiceAmount & "," & _
                  VoucherType & "," & _
                  AddQuotes(strTransID) & " ) "

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

InsertToSales = Success

Exit Function

ErrLine:
    MsgBox "InsertToSales" & vbCrLf & Err.Description, vbCritical
    
    
End Function

Public Property Get ObjectName()
    ObjectName = "SalesClass"
End Property


Private Function InsertToStock() As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rst As ADODB.Recordset
Dim lpCount As Integer
Dim MaxCount As Integer

Dim ParentID As Long
Dim RelationID As Long
Dim TransID As Long
Dim GodownID As Byte

Dim Quantity As Double 'local copy
Dim Soot As Double

Dim Amount As Currency 'local copy
Dim TaxAmount As Currency 'local copy

Dim PurchasePrice As Currency
Dim SalesPrice As Currency

Dim USTransDate As String

Dim VoucherType As Wis_VoucherTypes
Dim FreeVoucherType As Wis_VoucherTypes

Dim lKeys As Variant

Dim MaterialClass As clsMaterial
Dim colMaterials As clsMaterials

Set MaterialClass = New clsMaterial

'Initalise the function
InsertToStock = Failure
With m_frmSales
    USTransDate = GetSysFormatDate(.txtInvoiceDate)
    If .cmbBranch.ListIndex = -1 Then Exit Function
    GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
End With
MaxCount = m_colMaterials.count

lKeys = m_colMaterials.GetKeys

'FreeVoucherType = FreeSales

For lpCount = 0 To MaxCount - 1 'TempLoopNo
    Set MaterialClass = m_colMaterials.Item(lKeys(lpCount))
    
    RelationID = MaterialClass.RelationID
    Soot = MaterialClass.Soot
    ParentID = MaterialClass.ParentID
    TransID = 0
    
    If ParentID = 0 Then
        'Get maxTransID
        TransID = MaterialClass.GetMaxTransID()
        VoucherType = Sales
        
        SalesPrice = MaterialClass.UnitPrice
        
        Quantity = MaterialClass.Quantity
        Amount = MaterialClass.Amount
        TaxAmount = MaterialClass.TaxAmount
        
        TransID = TransID + 1
        
        gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID, " & _
                      " Quantity,UnitPrice,Amount,TaxAmount,TotalAmount, VoucherType," & _
                      " PurOrSaleID,TransDate ) " & _
                      " VALUES ( " & _
                      RelationID & "," & _
                      TransID & "," & _
                      GodownID & "," & _
                      Quantity & "," & SalesPrice & "," & _
                      Amount & "," & TaxAmount & "," & (Amount + TaxAmount) & "," & _
                      VoucherType & "," & _
                      m_SalesTransID & "," & _
                      "#" & USTransDate & "# ) "
                      
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
          
        
        If Soot > 0 Then
            SalesPrice = MaterialClass.UnitPrice
            TransID = TransID + 1
            VoucherType = StockSoot
    
            gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID, " & _
                            " Quantity,UnitPrice,VoucherType ," & _
                            " PurOrSaleID,TransDate ) " & _
                            " VALUES ( " & _
                            RelationID & "," & _
                            TransID & "," & _
                            GodownID & "," & _
                            Soot & "," & _
                            SalesPrice & "," & _
                            VoucherType & "," & _
                            m_SalesTransID & "," & _
                            "#" & USTransDate & "# ) "
                      
            If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
                
        End If
    End If
Next lpCount

InsertToStock = Success

Exit Function

ErrLine:
    MsgBox "InsertToStock" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function RemoveDataFromCollection() As Boolean
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim SelectedRow As Integer
Dim lngRelationID As Long
'Initialise the function
RemoveDataFromCollection = False

'Get the Selected row
With m_frmSales.grd
    SelectedRow = .Row
    .Col = 2
    'Get the RelationID from the selected row
    lngRelationID = .RowData(SelectedRow)
End With
If lngRelationID = 0 Then Exit Function

'remove the item from the collection
Call m_colMaterials.Remove("MAT" & lngRelationID)

RemoveDataFromCollection = True

Exit Function

ErrLine:
    
End Function

Public Sub Show()
Dim MaterialClass As clsMaterial

If MaterialClass Is Nothing Then Set MaterialClass = New clsMaterial

Dim headID As Long
headID = MaterialClass.GetHeadIDFromHeadsList(Enum_Customers)
If headID = 0 Then Exit Sub

m_VendorID = headID

If m_frmSales Is Nothing Then Set m_frmSales = New frmSales
m_frmSales.VendorID = m_VendorID
m_frmSales.lblCompanyName = MaterialClass.GetCompanyName(m_VendorID)
gWindowHandle = m_frmSales.hwnd
m_frmSales.Show 'vbModal

End Sub

Private Function Validated() As Boolean

Dim InvoiceAmount As Currency

Validated = False

If m_frmSales Is Nothing Then Set m_frmSales = New frmSales

With m_frmSales
    If .optCash.Value = False And .optCredit.Value = False Then
        MsgBox "Please select transaction type credit or cash", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
    If Not DateValidate(.txtInvoiceDate.Text, "/", True) Then
        'MsgBox "Invalid date specified!", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    If .txtInvoiceNumber.Text = "" Then Exit Function
    If Not m_UpdateStatus Then
        If InvoiceNoExists(.txtInvoiceNumber.Text) Then
            'MsgBox "Invoice Number already exits", vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(679), vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox .txtInvoiceNumber
            Exit Function
        End If
    End If
    
End With

Validated = True

End Function
Private Function InvoiceNoExists(ByVal InvoiceNo As String) As Boolean
'Declare the variables
Dim rst As Recordset

InvoiceNoExists = False
'Check the invoice no in the query
gDbTrans.SqlStmt = " SELECT InvoiceNo FROM Sales " & _
                   " WHERE InvoiceNO = '" & InvoiceNo & "'"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then InvoiceNoExists = True

End Function

Private Sub Class_Initialize()
'Initialise the objects
Set m_colMaterials = New clsMaterials


End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
If strSelection = "" Then Exit Sub
m_SalesTransID = CLng(strSelection)

End Sub

Private Sub m_frmLookUp_SubItems(strSubItem() As String)
m_InvoiceNo = strSubItem(0)
m_TransDate = strSubItem(1)
m_InvoiceAmount = strSubItem(2)
m_TransIDString = strSubItem(UBound(strSubItem) - 1)

End Sub


Private Sub m_frmSales_AddClicked()
'Validation are done at the Command Add CLickevent  no need to validate
'Add the data to the collection then add to grid
If AddDataToCollection Then
    If AddToGridFromCollection <> Success Then Exit Sub
    Call ClearControls
End If

End Sub

Private Sub m_frmSales_ChangeCustomer()

If Not m_colMaterials Is Nothing Then
    If m_colMaterials.count Then Exit Sub
End If

Dim MaterialClass As New clsMaterial

Dim headID As Long
headID = MaterialClass.GetHeadIDFromHeadsList(Enum_Customers)
If headID = 0 Then Set MaterialClass = Nothing: Exit Sub

m_VendorID = headID
m_frmSales.VendorID = m_VendorID
m_frmSales.lblCompanyName = MaterialClass.GetCompanyName(m_VendorID)

Set MaterialClass = Nothing

End Sub


Private Sub m_frmSales_ClearClicked()
Call ClearSalesForm
End Sub

Private Sub m_frmSales_DeleteClicked()
Call DeleteDetails
End Sub

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

Call ClearSalesForm

End Function

Private Sub m_frmSales_GridClicked()
Call frmGridClicked
End Sub

Private Sub m_frmSales_InvoiceClicked()
    Call ClearSalesForm
    Call LoadDetailsToUpdate
End Sub

Private Function LoadDetailsToUpdate() As Boolean
m_UpdateStatus = False
If Not GetSalesTableDetails Then Exit Function
If Not GetStockDetailsAddedToCollection Then Exit Function
If Not GetDetailsFromAccTrans Then Exit Function
If AddToGridFromCollection <> Success Then Exit Function
Call ClearControls

With m_frmSales
'    .tabSales.Tabs(1).Selected = True
    .cmdAccept.Caption = GetResourceString(171) ' "&Update"
    .cmdDelete.Enabled = True
End With

m_UpdateStatus = True
End Function
Private Function GetDetailsFromAccTrans() As Boolean
'Declare the variables
Dim arrTransID As Variant
Dim TotalAmount As Currency
Dim Amount As Currency
Dim TaxAmount As Currency

Dim headID As Long
Dim ParentID As Long
Dim headName As String

Dim rstAccTrans As ADODB.Recordset
Dim MaterialClass As clsMaterial

Dim Quantity As Double
Dim StrSql As String
Dim loopCount As Integer

'Initialise the function
GetDetailsFromAccTrans = False

If m_TransIDString = "" Then Exit Function

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
                  " AND ( C.AccountType = 4 OR C.AccountType= 8 OR C.AccountType= 1 )"

If gDbTrans.Fetch(rstAccTrans, adOpenStatic) < 1 Then Exit Function


Do While Not rstAccTrans.EOF

    Select Case rstAccTrans.Fields("AccountType")
        Case Profit
            Amount = FormatField(rstAccTrans.Fields("Credit"))
            headID = rstAccTrans.Fields("HeadID")
            headName = rstAccTrans.Fields("HeadName")
            ParentID = rstAccTrans.Fields("ParentID")
            Call m_colMaterials.Add(0, 0, 0, 0, headID, "", Amount, TaxAmount, headName, "", ParentID, 0, "MAT" & headID)
        Case Loss
            Amount = FormatField(rstAccTrans.Fields("Debit"))
            headID = rstAccTrans.Fields("HeadID")
            headName = rstAccTrans.Fields("HeadName")
            ParentID = rstAccTrans.Fields("ParentID")
            Call m_colMaterials.Add(0, 0, 0, 0, headID, "", Amount, TaxAmount, headName, "", ParentID, 0, "MAT" & headID)
    End Select
    
    rstAccTrans.MoveNext
Loop

With m_frmSales
    rstAccTrans.MoveFirst
    headID = rstAccTrans.Fields("HeadID")
    If VendorID <> headID Then
        
    End If
    .optCredit.Value = True
    If UBound(arrTransID) > 0 Then .optCash.Value = True
End With


Set rstAccTrans = Nothing

GetDetailsFromAccTrans = True
End Function



'This function will get Stock Details , by taking the ID of SalesTransID
' which is module level variable set after get the details from the Purchase table
Private Function GetStockDetailsAddedToCollection() As Boolean
'Declare the variables
Dim SalesVoucher As Wis_VoucherTypes
Dim rstSales As ADODB.Recordset
Dim rstRelation As ADODB.Recordset

Dim Quantity As Double
Dim SalesPrice As Currency
Dim RelationID As Long
Dim Amount As Currency
Dim TaxAmount As Currency
Dim TradePrice As Currency
Dim MRP As Currency
Dim headName As String
Dim ProductName As String
Dim ProductGroupName As String
Dim Soot As Double

Dim fldRelationID As ADODB.Field
Dim fldQuantity As ADODB.Field
Dim fldUnitPrice As ADODB.Field
Dim fldSootQuantity As ADODB.Field
Dim fldTax As ADODB.Field

'Trap an error
On Error GoTo ErrLine

'Iniitalse the function
GetStockDetailsAddedToCollection = False

'Validate the module level variables
If m_SalesTransID = 0 Then Exit Function

SalesVoucher = Sales

gDbTrans.SqlStmt = " SELECT RelationID,Quantity AS QtySales,UnitPrice,TaxAmount" & _
                   " FROM Stock " & _
                   " WHERE PurOrSaleID = " & m_SalesTransID & _
                   " AND VoucherType=" & SalesVoucher

Call gDbTrans.CreateView("QrySales")

SalesVoucher = StockSoot

gDbTrans.SqlStmt = " SELECT RelationID,Quantity AS QtySoot,UnitPrice,TaxAmount" & _
                   " FROM Stock " & _
                   " WHERE PurOrSaleID = " & m_SalesTransID & _
                   " AND VoucherType=" & SalesVoucher

Call gDbTrans.CreateView("QrySoot")


gDbTrans.SqlStmt = " SELECT A.*,B.QtySoot " & _
                   " FROM QrySales AS A " & _
                   " LEFT JOIN QrySoot AS B " & _
                   " ON A.RelationID=B.RelationID "

Call gDbTrans.CreateView("QrySalesMerge")

gDbTrans.SqlStmt = " SELECT * FROM QrySalesMerge "
                   
If gDbTrans.Fetch(rstSales, adOpenForwardOnly) < 1 Then Exit Function


If Not GetRstFromRelationMaster(rstRelation) Then Exit Function

Set fldRelationID = rstSales.Fields("RelationID")
Set fldQuantity = rstSales.Fields("QtySales")
Set fldUnitPrice = rstSales.Fields("UnitPrice")
Set fldSootQuantity = rstSales.Fields("QtySoot")
Set fldTax = rstSales.Fields("TaxAmount")
If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials

Do While Not rstSales.EOF
    RelationID = fldRelationID.Value
    Quantity = fldQuantity.Value
    Soot = FormatField(fldSootQuantity)
    SalesPrice = fldUnitPrice.Value
    Amount = Quantity * SalesPrice
    TaxAmount = FormatField(fldTax)
    
    TradePrice = 0: MRP = 0: headName = "": ProductName = ""
    
    If gDbTrans.FindRecord(rstRelation, "RelationID= " & RelationID) Then
        TradePrice = FormatField(rstRelation.Fields("TradingPrice"))
        MRP = FormatField(rstRelation.Fields("MRP"))
        ProductName = FormatField(rstRelation.Fields("ProductName"))
        ProductGroupName = FormatField(rstRelation.Fields("GroupName"))
    End If
    'Add To the collection
    Call m_colMaterials.Add(Quantity, SalesPrice, TradePrice, MRP, RelationID, headName, Amount, TaxAmount, ProductName, ProductGroupName, 0, Soot, "MAT" & RelationID)
    
    'Move the recordset
    rstSales.MoveNext
Loop


Set rstRelation = Nothing
Set rstSales = Nothing

GetStockDetailsAddedToCollection = True

Exit Function

ErrLine:
    If Err Then MsgBox "GetStockeDetailsAddedToCollection" & vbCrLf & Err.Description
    

End Function

Private Function GetRstFromRelationMaster(rstRelation As ADODB.Recordset) As Boolean

GetRstFromRelationMaster = False

gDbTrans.SqlStmt = " SELECT A.RelationID,TradingPrice,MRP,ProductName, GroupName" & _
                   " FROM RelationMaster AS A,Products C,QrySalesMerge D,ProductGroup E  " & _
                   " WHERE A.RelationID =D.RelationID " & _
                   " AND A.ProductID=C.ProductID " & _
                   " AND A.GroupID=E.GroupID "
If gDbTrans.Fetch(rstRelation, adOpenStatic) < 1 Then Exit Function

GetRstFromRelationMaster = True
End Function



'This function will get the details from the Sales table
' where module level variables are set
' Here : m_SalesTransID will be set which used to get the details from Stock Table
'      " m_TransIDString will be set which used to get the details from AccTrans Table
Private Function GetSalesTableDetails() As Boolean
'Declare the variables
Dim rstInvoices As ADODB.Recordset
Dim USFromDate As String
Dim USToDate As String
Dim GodownID As Byte
Dim SalesVoucher As Wis_VoucherTypes
Dim FreeSalesVoucher As Wis_VoucherTypes
Dim FillViewClass As clsFillView

GetSalesTableDetails = False
Dim SameVendor As Boolean
SameVendor = True
With m_frmSales
    USFromDate = GetSysFormatDate(FinIndianFromDate)
    USToDate = GetSysFormatDate(.txtInvoiceDate.Text)
    If .cmbBranch.ListIndex = -1 Then Exit Function
    If Val(.cmdSalesInvoice.Tag) Then SameVendor = False
    GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
End With

SalesVoucher = Sales

gDbTrans.SqlStmt = " SELECT TransID,InvoiceNo,InvoiceDate," & _
                   " InvoiceAmount,Trans_TransID" & _
                   " FROM Sales" & _
                   " WHERE GodownID = " & GodownID & _
                   " AND InvoiceDate BETWEEN #" & USFromDate & "#" & _
                   " AND #" & USToDate & "#" & _
                   " AND VoucherType = " & SalesVoucher
If SameVendor Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And HeadID = " & VendorID
End If

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY InvoiceDate "

If gDbTrans.Fetch(rstInvoices, adOpenForwardOnly) < 1 Then Exit Function

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Set FillViewClass = New clsFillView

m_InvoiceNo = ""

If Not FillViewClass.FillViewForPurchase(m_frmLookUp.lvwReport, rstInvoices, "TransID", True) Then Exit Function

m_frmLookUp.Show vbModal

Set m_frmLookUp = Nothing

Set FillViewClass = Nothing

If m_InvoiceNo = "" Then Exit Function

gDbTrans.SqlStmt = "SELECT * FROM Sales" & _
                   " WHERE InvoiceNo = " & AddQuotes(m_InvoiceNo)
Call gDbTrans.Fetch(rstInvoices, adOpenDynamic)

Dim headID As Long
With m_frmSales
    .txtInvoiceNumber = m_InvoiceNo
    headID = FormatField(rstInvoices("HeadID"))
    If VendorID <> headID Then
        Dim MaterialClass As New clsMaterial
        VendorID = headID
        .VendorID = m_VendorID
        .lblCompanyName = MaterialClass.GetCompanyName(m_VendorID)
        Set MaterialClass = Nothing
    End If
End With
Set rstInvoices = Nothing

GetSalesTableDetails = True

End Function

Private Sub m_frmSales_OkClicked()

'Validate the data
If Not Validated Then Exit Sub

If m_frmSales Is Nothing Then Set m_frmSales = New frmSales

'Save the data to the database
If SaveDetails Then ClearSalesForm
End Sub

Private Function SaveDetails() As wis_FunctionReturned
'Trap an error
On Error GoTo ErrLine

''August 2015
Dim ProductAmount() As Currency
Dim ProductHeadID() As Long
Dim TaxAmount As Currency
Dim TotalTaxAmount As Currency
''end August 2015

'Declare the Variables
Dim lpCount As Integer
Dim lpCount_2 As Integer
Dim MaxCount As Integer

Dim headID As Long
Dim ParentID As Long
Dim arrHeadID() 'Two Dimensional Array
Dim arrParentID()
Dim PLIDCount As Integer

Dim RelationID As Long
Dim GodownID As Byte

Dim TransDate As Date
Dim GroupName As String
Dim GroupNameEnglish As String
Dim IsCashSales As Boolean

Dim Amount As Currency
Dim TotalAmount As Currency
Dim SalesAmount As Currency

Dim AccTransClass As clsAccTrans
Dim LedgerClass As clsLedger
Dim MaterialClass As clsMaterial

Dim lKeys As Variant

Dim InTrans  As Boolean

If Not InTrans Then gDbTrans.BeginTrans: InTrans = True
If m_UpdateStatus Then
    'gDbTrans.BeginTrans
    'InTrans = True
    If Not DeleteTransaction Then GoTo Exit_Line
End If


With m_frmSales
    TransDate = GetSysFormatDate(.txtInvoiceDate)
    If .optCash.Value = True Then IsCashSales = True
    If .optCredit.Value = True Then IsCashSales = False
    GroupName = Trim$(.cmbGroup.Text)
    GroupNameEnglish = Trim$(.cmbGroupEnglish.ItemData(.cmbGroup.ListIndex))
End With

If IsCashSales Then
    Dim RegCashTrans As Boolean
    Dim SetupClass As clsSetup
    Set SetupClass = New clsSetup
    RegCashTrans = IIf(UCase(SetupClass.ReadSetupValue("Trading", "RegisterCashTransaction", "False")) = "TRUE", True, False)
    Set SetupClass = Nothing
End If

'Initialise the objects
Set AccTransClass = New clsAccTrans
Set LedgerClass = New clsLedger
Set MaterialClass = New clsMaterial

If m_colMaterials Is Nothing Then Set m_colMaterials = New clsMaterials

'Initialise the function
SaveDetails = Failure


ReDim ProductHeadID(0)
ReDim ProductAmount(0)

lKeys = m_colMaterials.GetKeys
MaxCount = m_colMaterials.count
For lpCount = 0 To MaxCount - 1
    Set MaterialClass = m_colMaterials.Item(lKeys(lpCount))
    
    Amount = MaterialClass.Amount
    ParentID = MaterialClass.ParentID
    TaxAmount = MaterialClass.TaxAmount
    'Get the total amount
    Select Case ParentID
        Case 0 ' will be given zero for stock transactions
            TotalAmount = TotalAmount + Amount
            SalesAmount = SalesAmount + Amount
            'August 2015
            TotalTaxAmount = TotalTaxAmount + TaxAmount
            ProductHeadID(UBound(ProductHeadID)) = LedgerClass.GetHeadIDCreated(wis_SalesParentID, MaterialClass.ProductGroupName & " " & GetResourceString(180), 0, "")
            ProductAmount(UBound(ProductAmount)) = Amount
            
            ReDim Preserve ProductHeadID(UBound(ProductHeadID) + 1)
            ReDim Preserve ProductAmount(UBound(ProductAmount) + 1)
            
            'END August 2015
        Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
            
            TotalAmount = TotalAmount + Amount
            ReDim Preserve arrHeadID(1, PLIDCount)
            ReDim Preserve arrParentID(PLIDCount)
            arrHeadID(0, PLIDCount) = MaterialClass.RelationID 'Headid
            arrHeadID(1, PLIDCount) = MaterialClass.Amount
            arrParentID(PLIDCount) = ParentID
            PLIDCount = PLIDCount + 1
        Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
            
            TotalAmount = TotalAmount - Amount
            ReDim Preserve arrHeadID(1, PLIDCount)
            ReDim Preserve arrParentID(PLIDCount)
            arrHeadID(0, PLIDCount) = MaterialClass.RelationID 'Headid
            arrHeadID(1, PLIDCount) = MaterialClass.Amount
            arrParentID(PLIDCount) = ParentID
            PLIDCount = PLIDCount + 1
    End Select
Next lpCount

m_InvoiceAmount = TotalAmount

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

'Begin Transaction
If Not InTrans Then gDbTrans.BeginTrans: InTrans = True

' debit the VendorID & and Credit the Sales Account in
' AccTrans table
'Get the Purchase Account HeadID
headID = LedgerClass.GetHeadIDCreated(wis_SalesParentID, GroupName & " " & GetResourceString(180), 0, GroupNameEnglish & " " & LoadResString(180))
Dim transHeadID As Long
transHeadID = IIf(RegCashTrans Or (Not IsCashSales), m_VendorID, wis_CashHeadID)
'Call AccTransClass.AllTransHeadsAdd(transHeadID, TotalAmount, 0)
'Call AccTransClass.AllTransHeadsAdd(headID, 0, SalesAmount)

'August 2015
MaxCount = UBound(ProductHeadID) - 1
For lpCount = 0 To MaxCount
    If lpCount = 0 Then Call AccTransClass.AllTransHeadsAdd(transHeadID, TotalAmount + TotalTaxAmount, 0)
    If ProductAmount(lpCount) > 0 Then Call AccTransClass.AllTransHeadsAdd(ProductHeadID(lpCount), 0, ProductAmount(lpCount))
Next

'Get the TaxgHead ID
'' Add the Tax Amount
If TotalTaxAmount > 0 Then
    transHeadID = LedgerClass.GetHeadIDCreated(parReceivable, GetResourceString(162, 173), 0, LoadResourceStringS(176, 173))
    Call AccTransClass.AllTransHeadsAdd(transHeadID, 0, TotalTaxAmount)
End If
'August 2015

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

'NEED NOT SAVE THE VOUCHER DETAILS IN CASE OF FREE TRASCATION
If (TotalAmount + TotalTaxAmount) > 0 Then
    If AccTransClass.SaveVouchers(Sales, TransDate, , True) <> Success Then
        gDbTrans.RollBack
        Exit Function
    End If
End If

'if the person is paid cash payment then Debit the cash in hand HeadID
'and Credit the SundryDebitor (or ) Customer
If RegCashTrans Then
    Call AccTransClass.AllTransHeadsAdd(wis_CashHeadID, m_InvoiceAmount, 0)
    Call AccTransClass.AllTransHeadsAdd(m_VendorID, 0, m_InvoiceAmount)
    
    If AccTransClass.SaveVouchers(Receipt, TransDate, , True) <> Success Then
        gDbTrans.RollBack
        Exit Function
    End If
End If

' Insert the purchase details of VendorID(HeadID) by Invoice Amount
' In the Purchase table
If InsertToSales((Not IsCashSales)) <> Success Then Exit Function

' Insert the Qunatities of each product into Stock table and update the
' stock balance accordingly
If InsertToStock() <> Success Then Exit Function

gDbTrans.CommitTrans
InTrans = False
SaveDetails = Success

'MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE

Exit Function

Exit_Line:
    If InTrans Then gDbTrans.RollBack
    Exit Function

ErrLine:
    If InTrans Then gDbTrans.RollBack
    MsgBox "SaveDetails:" & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    SaveDetails = FatalError
    
End Function

Private Function DeleteTransaction() As Boolean

DeleteTransaction = False

If Not DeleteSalesTableDetails Then Exit Function
If Not DeleteStockTableDetails Then Exit Function
If Not DeleteAccTransDetails Then Exit Function
DeleteTransaction = True

End Function

Private Function DeleteSalesTableDetails() As Boolean
'Declarew the variables
DeleteSalesTableDetails = False

If m_InvoiceNo = "" Then Exit Function
If m_TransDate = "" Then Exit Function

gDbTrans.SqlStmt = " DELETE * " & _
                   " FROM Sales " & _
                   " WHERE InvoiceNo = " & AddQuotes(m_InvoiceNo)
                   
If Not gDbTrans.SQLExecute Then Exit Function

DeleteSalesTableDetails = True
End Function

Private Function DeleteStockTableDetails() As Boolean
'Declare the variables
Dim USTransDate As String

DeleteStockTableDetails = False
If m_InvoiceNo = "" Then Exit Function
If m_TransDate = "" Then Exit Function
If m_SalesTransID <= 0 Then Exit Function

USTransDate = GetSysFormatDate(m_TransDate)

gDbTrans.SqlStmt = " DELETE *" & _
                   " FROM Stock " & _
                   " WHERE PurORSaleID = " & m_SalesTransID & _
                   " AND TransDate =#" & USTransDate & "#"

If Not gDbTrans.SQLExecute Then Exit Function

DeleteStockTableDetails = True
End Function


Private Function DeleteAccTransDetails() As Boolean
'Declare the variables
Dim arrTransID As Variant
Dim StrSql As String
Dim loopCount As Integer

DeleteAccTransDetails = False

If m_TransIDString = "" Then Exit Function
If m_SalesTransID <= 0 Then Exit Function

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




Public Property Get VendorID() As Long
VendorID = m_VendorID
End Property

Public Property Let VendorID(ByVal vNewValue As Long)
m_VendorID = vNewValue
End Property


Private Sub m_frmSales_PLHeadsClicked(ByVal AccType As wis_AccountType)
'Validation are done at the Command Add CLickevent  no need to validate
'Add the data to the collection then add to grid
If AddPLHeadsToCollection(AccType) Then Call AddToGridFromCollection
End Sub

Private Function AddPLHeadsToCollection(ByVal AccType As wis_AccountType) As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim HeadAmount As Currency
Dim TaxAmount As Currency
Dim headName As String
Dim headID As Long
Dim ParentID As Long

AddPLHeadsToCollection = False
'Assign the values
With m_frmSales
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
    
If m_colMaterials.Exists("MAT" & headID) Then
    MsgBox GetResourceString(752), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If
ParentID = GetParentID(headID)
Call m_colMaterials.Add(0, 0, 0, 0, headID, "", HeadAmount, TaxAmount, headName, "", ParentID, 0, "MAT" & headID)

AddPLHeadsToCollection = True

Exit Function

ErrLine:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"

End Function



Private Sub m_frmSales_UnDoClicked()

If RemoveDataFromCollection Then
    m_frmSales.InitGrid
    AddToGridFromCollection
End If
End Sub


Private Sub m_frmSales_WindowClosed()
    Set m_frmSales = Nothing
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmSales Is Nothing Then _
            ObjectHandle = m_frmSales.hwnd
End Property

