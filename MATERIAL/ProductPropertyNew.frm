VERSION 5.00
Begin VB.Form frmProductPropertyNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties of an Item or Product"
   ClientHeight    =   7230
   ClientLeft      =   1650
   ClientTop       =   2220
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProductPropertyNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8580
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   5895
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   8325
      Begin VB.TextBox txtTaxAmount 
         Height          =   375
         Left            =   6000
         TabIndex        =   23
         Top             =   5280
         Width           =   2175
      End
      Begin VB.TextBox txtStockValue 
         Height          =   360
         Left            =   6000
         TabIndex        =   19
         Top             =   4755
         Width           =   2175
      End
      Begin VB.TextBox txtTax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   21
         Top             =   5280
         Width           =   1365
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   ".."
         Height          =   315
         Left            =   7650
         TabIndex        =   28
         Top             =   2220
         Width           =   435
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   ".."
         Height          =   315
         Left            =   7650
         TabIndex        =   29
         Top             =   1650
         Width           =   435
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   ".."
         Height          =   315
         Left            =   7650
         TabIndex        =   27
         Top             =   930
         Width           =   435
      End
      Begin VB.TextBox txtOpBalance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   17
         Top             =   4755
         Width           =   1395
      End
      Begin VB.TextBox txtSalesPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   14
         Top             =   4004
         Width           =   1365
      End
      Begin VB.TextBox txtTradingPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   11
         Top             =   2970
         Width           =   1365
      End
      Begin VB.ComboBox cmbUnit 
         BeginProperty Font 
            Name            =   "Nudi B-Akshar"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "ProductPropertyNew.frx":030A
         Left            =   2640
         List            =   "ProductPropertyNew.frx":030C
         TabIndex        =   9
         Top             =   2175
         Width           =   4725
      End
      Begin VB.ComboBox cmbGroup 
         BeginProperty Font 
            Name            =   "Nudi B-Akshar"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "ProductPropertyNew.frx":030E
         Left            =   2640
         List            =   "ProductPropertyNew.frx":0310
         TabIndex        =   5
         Top             =   885
         Width           =   4725
      End
      Begin VB.ComboBox cmbProductName 
         BeginProperty Font 
            Name            =   "Nudi B-Akshar"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "ProductPropertyNew.frx":0312
         Left            =   2640
         List            =   "ProductPropertyNew.frx":0314
         TabIndex        =   7
         Top             =   1530
         Width           =   4725
      End
      Begin VB.TextBox txtMRP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2640
         TabIndex        =   12
         Top             =   3487
         Width           =   1365
      End
      Begin VB.ComboBox cmbBranch 
         BeginProperty Font 
            Name            =   "Nudi B-Akshar"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "ProductPropertyNew.frx":0316
         Left            =   2640
         List            =   "ProductPropertyNew.frx":0318
         TabIndex        =   3
         Top             =   240
         Width           =   5565
      End
      Begin VB.Label lblTaxAmount 
         Caption         =   "Tax Amount"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Stock Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label lblTax 
         Caption         =   "Tax in Percent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   20
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label lblPurchaseUnit 
         Caption         =   "PurchaseUnit"
         Height          =   225
         Left            =   4890
         TabIndex        =   33
         Top             =   2970
         Width           =   1815
      End
      Begin VB.Label lblMRPUnit 
         Caption         =   "per unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4890
         TabIndex        =   30
         Top             =   3427
         Width           =   1755
      End
      Begin VB.Label lblStockValue 
         Caption         =   "Stock Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   4890
         Width           =   1695
      End
      Begin VB.Label lblMRP 
         Caption         =   "MRP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   32
         Top             =   3517
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   7750
         X2              =   120
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label lblProductName 
         Caption         =   "Select Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   6
         Top             =   1635
         Width           =   2295
      End
      Begin VB.Label lblGroup 
         Caption         =   "Select Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   270
         TabIndex        =   4
         Top             =   1050
         Width           =   2175
      End
      Begin VB.Label lblUnit 
         Caption         =   "Select Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   8
         Top             =   2325
         Width           =   2295
      End
      Begin VB.Label lblTradingPrice 
         Caption         =   "Purchase Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   10
         Top             =   2970
         Width           =   2175
      End
      Begin VB.Label lblSalesPrice 
         Caption         =   "Sales Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   4064
         Width           =   2145
      End
      Begin VB.Label lblSalesUnit 
         Caption         =   "Sales Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   4004
         Width           =   1815
      End
      Begin VB.Label lblOpBalance 
         Caption         =   "Opening Stock Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   4365
         Width           =   2355
      End
      Begin VB.Label lblBranch 
         Caption         =   "Select Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   390
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5700
      TabIndex        =   25
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7140
      TabIndex        =   26
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4260
      TabIndex        =   24
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   60
      Width           =   2205
   End
End
Attribute VB_Name = "frmProductPropertyNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_groupAdd As clsAddGroup
Attribute m_groupAdd.VB_VarHelpID = -1
Private m_dbOperation As wis_DBOperation
Private m_RelationID As Long

Private m_GroupID As Integer
Private m_UnitID As Long
Private m_GodownID As Byte

Private Sub ClearTextBoxes()
With Me
    .txtTradingPrice.Text = ""
    .lblPurchaseUnit.Caption = ""
    .txtMRP.Text = ""
    .lblMRPUnit.Caption = ""
    .txtSalesPrice.Text = ""
    .lblSalesUnit.Caption = ""
    .txtOpBalance.Text = ""
    
    m_dbOperation = Insert
    m_RelationID = 0
    cmdOk.Caption = GetResourceString(4) ' "&Accept"
    cmdDelete.Enabled = False
End With

End Sub

Private Sub LoadDetails()

Dim GroupID As Integer
Dim ProductID As Long
Dim UnitID As Long
Dim GodownID As Byte

Dim rstRelation As ADODB.Recordset

If cmbBranch.Enabled Then If cmbBranch.ListIndex = -1 Then Exit Sub
If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbProductName.ListIndex = -1 Then Exit Sub
If cmbUnit.ListIndex = -1 Then Exit Sub

GodownID = cmbBranch.ItemData(cmbBranch.ListIndex)
GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
ProductID = cmbProductName.ItemData(cmbProductName.ListIndex)
UnitID = cmbUnit.ItemData(cmbUnit.ListIndex)

Call ClearTextBoxes
Call SetUnitNameToLabels

m_RelationID = 0

gDbTrans.SqlStmt = " SELECT RelationID FROM RelationMaster " & _
                   " WHERE GroupID = " & GroupID & _
                   " AND UnitID = " & UnitID & _
                   " AND ProductID = " & ProductID & _
                   " AND GodownID = " & GodownID & _
                   " AND PriceChanged = " & 0

If gDbTrans.Fetch(rstRelation, adOpenForwardOnly) < 1 Then Exit Sub

m_RelationID = rstRelation.Fields("RelationID")
Set rstRelation = Nothing

gDbTrans.SqlStmt = " SELECT * FROM RelationMaster WHERE RelationID = " & m_RelationID
    
'knowing that already this record exists in the database
'if fetch returns negative value then there must error in the database
Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)

txtTradingPrice.Text = rstRelation.Fields("TradingPrice")
txtMRP.Text = rstRelation.Fields("MRP")
txtSalesPrice.Text = FormatField(rstRelation("SalesPrice"))
txtTax.Text = FormatField(rstRelation("Tax"))

If MsgBox("Item already exists" & vbCrLf & "Are the prices changed?", _
               vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
    UpdatePriceChanged (GodownID)
Else
    LoadValuesForUpdation (GodownID)
End If

Call SetUnitNameToLabels

End Sub

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls
lblBranch.Caption = GetResourceString(227, 27)

lblGroup.Caption = GetResourceString(157, 27)
lblProductName.Caption = GetResourceString(158, 35, 27)
lblUnit.Caption = GetResourceString(161, 27)
lblTradingPrice.Caption = GetResourceString(162, 212) '
lblMRP.Caption = GetResourceString(163)
lblSalesPrice.Caption = GetResourceString(164)
lblOpBalance.Caption = GetResourceString(284, 306)
lblQuantity.Caption = GetResourceString(306)
lblStockValue.Caption = GetResourceString(284, 40)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
cmdDelete.Caption = GetResourceString(14)
lblTax.Caption = GetResourceString(173) + " (in %))"
lblTaxAmount.Caption = GetResourceString(173, 40)

End Sub
Private Function GetSalesFromTradingPrice(ByVal TradingPrice As Currency, ByVal Kst As Single, ByVal Cst As Single) As Currency
' Trap an error
On Error GoTo ErrLine
'Declare the variables

'Initialse the function
GetSalesFromTradingPrice = 0

'Validate the inputs

If TradingPrice = 0 Then Exit Function

TradingPrice = TradingPrice + (TradingPrice * Cst / 100)

GetSalesFromTradingPrice = FormatCurrency(TradingPrice + (TradingPrice * Kst / 100))

ErrLine:
    
End Function

Private Sub LoadValuesForUpdation(ByVal GodownID As Byte)
Dim rst As ADODB.Recordset
Dim lngTransID As Long

'get the opening stock from the stock table
gDbTrans.SqlStmt = " SELECT TOP 1 Quantity,TransID,Amount,TaxAmount FROM Stock " & _
                 " WHERE RelationID = " & m_RelationID & _
                 " and GodownID =   " & GodownID & _
                 " ORDER BY TransID ASC"
                 
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

txtOpBalance.Text = FormatField(rst("Quantity"))
txtStockValue.Text = FormatField(rst("Amount"))
txtTaxAmount.Text = FormatField(rst("TaxAmount"))
lngTransID = FormatField(rst("TransID"))

cmdOk.Caption = GetResourceString(171) '"&Update"
m_dbOperation = Update
cmdDelete.Enabled = True
txtOpBalance.Enabled = True
'If lngTransID > 1 Then txtOpBalance.Enabled = False
End Sub

Private Sub SetFont()
On Error Resume Next

Dim Ctrl As Control

For Each Ctrl In Me
      Ctrl.FontName = "Arial"
      Ctrl.FONTSIZE = 10
Next Ctrl


End Sub


Private Sub ClearControls()

'cmbProductName.ListIndex = -1
'cmbUnit.ListIndex = -1
txtTradingPrice.Text = ""
txtMRP.Text = ""
txtSalesPrice.Text = ""
txtOpBalance.Text = ""
lblSalesUnit.Caption = ""
lblPurchaseUnit.Caption = ""
lblMRPUnit.Caption = ""
txtTax.Text = ""
txtTaxAmount.Text = ""
txtStockValue.Text = ""
On Error Resume Next
'cmbGroup.SetFocus

m_dbOperation = Insert
cmdOk.Caption = "&Ok"
cmdDelete.Enabled = False
End Sub

Private Sub LoadBranches()
Dim rstBranches As ADODB.Recordset
Dim GodownID As ADODB.Field
Dim GodownName As ADODB.Field




gDbTrans.SqlStmt = " SELECT GodownID,GodownName FROM GodownDet " & _
                   " ORDER BY GodownID "
        
Call gDbTrans.Fetch(rstBranches, adOpenStatic)

Set GodownID = rstBranches.Fields("GodownID")
Set GodownName = rstBranches.Fields("GodownName")

cmbBranch.Clear

Do While Not rstBranches.EOF
   
   cmbBranch.AddItem GodownName.Value
   cmbBranch.ItemData(cmbBranch.newIndex) = GodownID.Value
   
   'Move to next record
   rstBranches.MoveNext
Loop

If rstBranches.RecordCount = 1 Then
    cmbBranch.ListIndex = 0
    cmbBranch.Enabled = False
End If


End Sub
Private Sub LoadProductGroups()
Dim rstGroups As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT GroupID,GroupName FROM ProductGroup " & _
                   " ORDER BY GroupID "
        
Call gDbTrans.Fetch(rstGroups, adOpenForwardOnly)

cmbGroup.Clear

Do While Not rstGroups.EOF
   
   cmbGroup.AddItem FormatField(rstGroups.Fields("GroupName"))
   cmbGroup.ItemData(cmbGroup.newIndex) = FormatField(rstGroups.Fields("GroupID"))
   
   'Move to next record
   rstGroups.MoveNext
Loop

End Sub


Private Sub LoadUnits()

Dim rstUnits As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT UnitID,UnitName FROM Units " & _
                   " ORDER BY UnitID "
        
Call gDbTrans.Fetch(rstUnits, adOpenForwardOnly)

cmbUnit.Clear

Do While Not rstUnits.EOF
   cmbUnit.AddItem FormatField(rstUnits.Fields("UnitName"))
   cmbUnit.ItemData(cmbUnit.newIndex) = FormatField(rstUnits.Fields("UnitID"))
   
   'Move to next record
   rstUnits.MoveNext
Loop

End Sub


Private Sub LoadProducts()
Dim rstProducts As ADODB.Recordset
Dim intGroupID As Integer

If cmbGroup.ListIndex = -1 Then Exit Sub

intGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)

cmbProductName.Clear

gDbTrans.SqlStmt = " SELECT ProductID,ProductName FROM Products " & _
                  " WHERE GroupID = " & intGroupID & _
                  " ORDER BY ProductID "
        
If gDbTrans.Fetch(rstProducts, adOpenForwardOnly) < 1 Then Exit Sub

Do While Not rstProducts.EOF
   cmbProductName.AddItem FormatField(rstProducts("ProductName"))
   cmbProductName.ItemData(cmbProductName.newIndex) = FormatField(rstProducts("ProductID"))
   
   'Move to next record
   rstProducts.MoveNext
Loop

End Sub



Private Function SaveProductPropertyDetails() As Boolean
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
'Dim sglCST As Single
'Dim sglKST As Single

Dim isPriceChanged As Byte
Dim USTransDate As String

Dim opVoucherType As Wis_VoucherTypes

'Initialise the function
SaveProductPropertyDetails = False

'Exit the function if comboboxes are empty
If cmbBranch.Enabled Then If cmbBranch.ListIndex = -1 Then Exit Function
If cmbGroup.ListIndex = -1 Then Exit Function
If cmbProductName.ListIndex = -1 Then Exit Function
If cmbUnit.ListIndex = -1 Then Exit Function

isPriceChanged = 0
On Error GoTo ErrLine


USTransDate = GetSysFormatDate(FinIndianFromDate)
btGodownID = cmbBranch.ItemData(cmbBranch.ListIndex)
intGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
lngProductID = cmbProductName.ItemData(cmbProductName.ListIndex)
lngUnitID = cmbUnit.ItemData(cmbUnit.ListIndex)

curTradingprice = Val(Trim$(txtTradingPrice.Text))
curMRP = Val(Trim$(txtMRP.Text))
curSalesPrice = Val(Trim$(txtSalesPrice.Text))
dblOpStockBalance = Val(txtOpBalance.Text)
'sglKST = Val(txtKst.Text)
'sglCST = Val(txtCst.Text)

lngTransID = 1

'curPurchasePrice = GetAddedPercentage(curTradingprice, sglKST)
'If sglCST > 0 Then curPurchasePrice = GetAddedPercentage(curTradingprice, sglCST)


opVoucherType = Purchase
OpStockValue = curTradingprice * dblOpStockBalance

gDbTrans.SqlStmt = "SELECT MAX(RelationID) FROM RelationMaster "
         
Call gDbTrans.Fetch(rstRelation, adOpenForwardOnly)
lngRelationID = FormatField(rstRelation(0)) + 1

Set rstRelation = Nothing
   
gDbTrans.BeginTrans
   
If m_RelationID Then
    gDbTrans.SqlStmt = "UPDATE RelationMaster SET " & _
                  " PriceChanged = " & 1 & _
                  " WHERE RelationID = " & m_RelationID
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
End If
   
gDbTrans.SqlStmt = "INSERT INTO RelationMaster " & _
            "(RelationID,GodownID," & _
            " GroupID,ProductID,UnitID,TradingPrice," & _
            "MRP,SalesPrice, PriceChanged,Tax ) " & _
            " VALUES ( " & _
            lngRelationID & "," & _
            btGodownID & "," & _
            intGroupID & "," & _
            lngProductID & "," & _
            lngUnitID & "," & _
            curTradingprice & "," & _
            curMRP & "," & _
            curSalesPrice & "," & _
            isPriceChanged & "," & _
            Val(txtTax.Text) & " ) "

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
        
If txtOpBalance.Enabled = True Then
    'Insert this OpBalance into stock table as its first entry
    gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID,Quantity," & _
                       " UnitPrice,VoucherType,PurORSaleID,TransDate,Amount,TaxAmount ) " & _
                       " VALUES ( " & _
                       lngRelationID & "," & _
                       lngTransID & "," & _
                       btGodownID & "," & _
                       dblOpStockBalance & "," & _
                       curTradingprice & "," & _
                       opVoucherType & "," & _
                       0 & "," & _
                       "#" & USTransDate & "#," & _
                       Val(txtStockValue.Text) & "," & _
                       Val(txtTaxAmount.Text) & " ) "
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
       
End If

gDbTrans.CommitTrans

SaveProductPropertyDetails = True

'MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE

Exit Function

ErrLine:
   MsgBox "SaveProductPropertyDetails: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
   
End Function



Private Sub SetUnitNameToLabels()
lblPurchaseUnit.Caption = ""
lblSalesUnit.Caption = ""
lblMRPUnit.Caption = ""
lblPurchaseUnit.Caption = GetResourceString(439) & " " & cmbUnit.Text
lblSalesUnit.Caption = GetResourceString(439) & " " & cmbUnit.Text
lblMRPUnit.Caption = GetResourceString(439) & " " & cmbUnit.Text
End Sub

Private Function UpdateProductProperties() As Boolean
'Trap an error
On Error GoTo ErrLine

'Declare the Variables
Dim btGodownID As Byte
Dim intGroupID As Integer
Dim lngProductID As Long
Dim lngUnitID As Long
Dim lngRelationID As Long
Dim lngTransID As Long

Dim rstRelation As ADODB.Recordset

'Dim curPurchasePrice As Currency
Dim curTradingprice As Currency
Dim curMRP As Currency
Dim curSalesPrice As Currency
Dim OpStockValue As Currency

Dim dblOpStockBalance As Double

Dim isPriceChanged As Byte
Dim opVoucherType As Wis_VoucherTypes


UpdateProductProperties = False

If cmbBranch.Enabled Then If cmbBranch.ListIndex = -1 Then Exit Function
If cmbGroup.ListIndex = -1 Then Exit Function
If cmbProductName.ListIndex = -1 Then Exit Function
If cmbUnit.ListIndex = -1 Then Exit Function

If MsgBox("You are going to update the previous data " & vbCrLf & _
          GetResourceString(541), _
          vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function

isPriceChanged = 0
btGodownID = cmbBranch.ItemData(cmbBranch.ListIndex)
intGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
lngProductID = cmbProductName.ItemData(cmbProductName.ListIndex)
lngUnitID = cmbUnit.ItemData(cmbUnit.ListIndex)

curTradingprice = Val(Trim$(txtTradingPrice.Text))
curSalesPrice = Val(Trim$(txtSalesPrice.Text))
curMRP = Val(Trim$(txtMRP.Text))

dblOpStockBalance = Val(txtOpBalance.Text)
lngTransID = 1

'curPurchasePrice = GetAddedPercentage(curTradingprice, sglKST)
'If sglCST > 0 Then curPurchasePrice = GetAddedPercentage(curTradingprice, sglCST)

opVoucherType = Purchase

OpStockValue = curTradingprice * dblOpStockBalance

gDbTrans.SqlStmt = " UPDATE RelationMaster SET " & _
                " GodownID = " & btGodownID & "," & _
                " GroupID =" & intGroupID & "," & _
                " ProductID = " & lngProductID & "," & _
                " UnitID = " & lngUnitID & "," & _
                " TradingPrice = " & curTradingprice & "," & _
                " MRP = " & curMRP & "," & _
                " SalesPrice = " & curSalesPrice & "," & _
                " PriceChanged = " & isPriceChanged & "," & _
                " Tax = " & Val(txtTax.Text) & _
                " WHERE RelationID = " & m_RelationID

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
    
If txtOpBalance.Enabled Then
    gDbTrans.SqlStmt = " UPDATE Stock SET " & _
                    " Quantity = " & dblOpStockBalance & "," & _
                    " UnitPrice = " & curTradingprice & "," & _
                    " Amount = " & Val(txtStockValue.Text) & "," & _
                    " TaxAmount = " & Val(txtTaxAmount.Text) & _
                    " WHERE RelationID = " & m_RelationID & _
                    " AND TransID = " & lngTransID
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
       
End If
    
gDbTrans.CommitTrans

UpdateProductProperties = True

'MsgBox "Details Updated", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(707), vbInformation, wis_MESSAGE_TITLE

Exit Function

ErrLine:
      MsgBox "UpdateProductProperties: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE

End Function

Private Function Validated() As Boolean
'Trap an error
On Error GoTo ErrLine
'Declare the variables
'Dim rstRelation As ADODB.Recordset

'Initalise the function
Validated = False

If cmbBranch.ListIndex = -1 Then
   'MsgBox "Select the Branch", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetChangeString(GetResourceString(230), GetResourceString(227)), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbBranch.SetFocus
   Exit Function
End If
If cmbGroup.ListIndex = -1 Then
  ' MsgBox "Select Group name", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(786), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbGroup.SetFocus
   Exit Function
End If
If cmbProductName.ListIndex = -1 Then
  ' MsgBox "Select Product Name", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetChangeString(GetResourceString(230), GetResourceString(158)), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbProductName.SetFocus
   Exit Function
End If
If cmbUnit.ListIndex = -1 Then
   'MsgBox "Select Unit", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(161, 27), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbUnit.SetFocus
   Exit Function
End If
If Not CurrencyValidate(Trim$(txtTradingPrice), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtTradingPrice
   Exit Function
End If
If Not CurrencyValidate(Trim$(txtMRP), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtMRP
   Exit Function
End If
If Not CurrencyValidate(Trim$(txtSalesPrice), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtSalesPrice
   Exit Function
End If
If Not CurrencyValidate(Trim$(txtOpBalance), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtOpBalance
   Exit Function
End If

If Not CurrencyValidate(Trim$(txtStockValue), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtStockValue
   Exit Function
End If

If txtTax.Text <> "" Then
    If IsNumeric(txtTax.Text) Then
        If Val(txtTax.Text) > 25 Then
            MsgBox GetResourceString(833), vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtTax
            Exit Function
        End If
    Else
        MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtTax
        Exit Function
    End If
End If

If Not CurrencyValidate(Trim$(txtTaxAmount), True) Then
   'MsgBox "Invalid currency value specified!", vbExclamation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtTaxAmount
   Exit Function
End If


Validated = True

ErrLine:

End Function


Private Sub cmbBranch_Click()
Call ClearControls
End Sub


Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
Call LoadProducts
If m_dbOperation = Update Then Call ClearTextBoxes
End Sub


Private Sub cmbProductName_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbProductName.ListIndex = -1 Then Exit Sub
'If cmbUnit.ListIndex = -1 Then Exit Sub

Call LoadUnits

End Sub


Private Sub cmbUnit_Click()
Call LoadDetails
End Sub

Private Sub UpdatePriceChanged(ByVal GodownID As Byte)
Dim rst As ADODB.Recordset
Dim lngTransID As Long

On Error GoTo ErrLine

'get the opening stock from the stock table
gDbTrans.SqlStmt = "SELECT TOP 1 Quantity,TransID,Amount,TaxAmount FROM Stock " & _
                   " WHERE RelationID = " & m_RelationID & _
                   " AND GodownID = " & GodownID & _
                   " ORDER BY TransID ASC "
                 
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

txtOpBalance.Text = FormatField(rst("Quantity"))
txtStockValue.Text = FormatField(rst("Amount"))
txtTaxAmount.Text = FormatField(rst("TaxAmount"))

lngTransID = FormatField(rst("TransID"))

cmdOk.Caption = GetResourceString(4) '"&Ok" & Accept
m_dbOperation = Insert
cmdDelete.Enabled = False

txtOpBalance.Enabled = True
If lngTransID > 1 Then txtOpBalance.Enabled = False

Exit Sub

ErrLine:
    MsgBox "UpdatePriceChanged" & vbCrLf & Err.Description, vbCritical
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
If m_RelationID < 1 Then Exit Sub

If MsgBox("This will delete your Product Setting." & vbCrLf & "Do you Delete?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Product Property") = vbNo Then Exit Sub


If Not DeleteRelationID Then
    MsgBox "This Product Property has Stock Transactions." & vbCrLf & "So will not be Deleted", vbInformation, "Delete Transaction"
    Exit Sub
End If

MsgBox "Product Properties Deleted", vbInformation

Call ClearControls

End Sub
Private Function DeleteRelationID() As Boolean
'Declare the variables
DeleteRelationID = False

On Error GoTo ErrLine

'Check this RelationID Has any Transctions
If RelationIDHasTransactions Then Exit Function

gDbTrans.BeginTrans
    

gDbTrans.SqlStmt = " DELETE FROM" & _
                   " Stock" & _
                   " WHERE RelationID= " & m_RelationID

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
gDbTrans.SqlStmt = " DELETE FROM" & _
                   " RelationMaster" & _
                   " WHERE RelationID = " & m_RelationID

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
gDbTrans.CommitTrans

DeleteRelationID = True

Exit Function

ErrLine:
    MsgBox "DelateRelationID" & vbCrLf & Err.Description, vbCritical
    
End Function


Private Function RelationIDHasTransactions() As Boolean
'Declare the Variables
Dim rstStock As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT TransID" & _
                  " FROM Stock" & _
                  " WHERE PurOrSaleID > 0 " & _
                  " AND RelationID = " & m_RelationID
                  
If gDbTrans.Fetch(rstStock, adOpenForwardOnly) < 1 Then Exit Function

RelationIDHasTransactions = True

End Function


Private Sub cmdGroup_Click()

If m_groupAdd Is Nothing Then Set m_groupAdd = New clsAddGroup

m_groupAdd.ShowAddGroup grpProduct

End Sub

Private Sub cmdItem_Click()

frmCreateItem.Show 1

End Sub


Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

If m_dbOperation = Insert Then
    If Not SaveProductPropertyDetails Then Exit Sub
ElseIf m_dbOperation = Update Then
    If Not UpdateProductProperties Then Exit Sub
End If

cmbProductName.ListIndex = -1
cmbUnit.ListIndex = -1
Call ClearControls

End Sub




Private Sub cmdUnit_Click()
If m_groupAdd Is Nothing Then Set m_groupAdd = New clsAddGroup

m_groupAdd.ShowAddGroup grpUnit

End Sub

Private Sub Form_Load()
'Declare the varriables
Dim MaterialClass As clsMaterial

'Center the form
CenterMe Me

'set icon for the form caption
Me.Icon = LoadResPicture(147, vbResIcon)
  
Call SetKannadaCaption

'LoadBranches
Call LoadBranches

'Load the Groups Combo box
Call LoadProductGroups

'Load the Units Combo box
Call LoadUnits

If MaterialClass Is Nothing Then Set MaterialClass = New clsMaterial

'If MaterialClass.isCompanyWithinState(m_CompanyID) Then
'    'set Kst to 10.5%
'    txtKst.Text = 10.5
'Else
'    'set Kst to 10.5%
'    txtKst.Text = 10.5
'    'txtCst.Text = 4
'End If

m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"
cmdDelete.Enabled = False
txtOpBalance.Enabled = True


End Sub


Public Property Get GroupID() As Long
GroupID = m_GroupID
End Property

Public Property Let GroupID(ByVal vNewValue As Long)
m_GroupID = vNewValue
End Property

Public Property Get UnitID() As Long
UnitID = m_UnitID
End Property

Public Property Let UnitID(ByVal vNewValue As Long)
m_UnitID = vNewValue
End Property

Public Property Get GodownID() As Byte
GodownID = m_GodownID
End Property

Public Property Let GodownID(ByVal vNewValue As Byte)
m_GodownID = vNewValue
End Property

Private Sub CalculateStockValue()
txtStockValue.Text = FormatCurrency(Val(txtTradingPrice.Text) * Val(txtOpBalance.Text))
End Sub


Private Sub Form_Resize()
lblCompanyName.Left = (Me.Width - lblCompanyName.Width) / 2
End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set frmProductPropertyNew = Nothing
End Sub









Private Sub m_groupAdd_ItemAdded(strAddItem As String, NewID As Long)
Dim cmb As ComboBox
If m_groupAdd.GroupType = grpProduct Then Set cmb = cmbGroup
If m_groupAdd.GroupType = grpUnit Then Set cmb = cmbUnit

With cmb
    .AddItem strAddItem
    .ItemData(.newIndex) = NewID
End With

End Sub

Private Sub m_groupAdd_ItemDeleted(strDelItem As String)
Dim cmb As ComboBox
Dim count As Integer
Dim MaxCount As Integer
If m_groupAdd.GroupType = grpProduct Then Set cmb = cmbGroup
If m_groupAdd.GroupType = grpUnit Then Set cmb = cmbUnit

With cmb
    MaxCount = .ListCount - 1
    For count = 0 To MaxCount
        If StrComp(strDelItem, .List(count), vbTextCompare) = 0 Then
            If .ListIndex = count Then .ListIndex = .ListIndex - 1
            .RemoveItem count
            Exit For
        End If
    Next
End With

End Sub


Private Sub m_groupAdd_ItemDeleting(strDelItem As String, Cancel As Integer)
Dim cmb As ComboBox
Dim count As Integer
Dim MaxCount As Integer
Dim ItemId As Long

If m_groupAdd.GroupType = grpProduct Then Set cmb = cmbGroup
If m_groupAdd.GroupType = grpUnit Then Set cmb = cmbUnit

'Now Get the Item Id
With cmb
    MaxCount = .ListCount - 1
    For count = 0 To MaxCount
        If StrComp(strDelItem, .List(count), vbTextCompare) = 0 Then
            ItemId = .ItemData(count)
            Exit For
        End If
    Next
End With

If ItemId = 0 Then Exit Sub

'Now Check Whether This Id is Used in any transaction
If m_groupAdd.GroupType = grpProduct Then _
    gDbTrans.SqlStmt = "SELECT GroupID From Products " & _
        " Where GroupID = " & ItemId

If m_groupAdd.GroupType = grpUnit Then _
    gDbTrans.SqlStmt = "SELECT UnitID From RelationMaster " & _
        " Where UnitID = " & ItemId

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Cancel = True
Set rst = Nothing
End Sub


Private Sub txtOpBalance_Change()
Call CalculateStockValue
End Sub

Private Sub txtStockValue_Change()
    Call txtTax_Change
End Sub

Private Sub txtTax_Change()
    If Val(txtTax.Text) < 25 Then txtTaxAmount.Text = FormatCurrency(Val(txtTax.Text) / 100 * Val(txtStockValue.Text))
End Sub

Private Sub txtTradingPrice_Change()
Dim TRP As Currency
Dim Kst As Single
Dim Cst As Single

TRP = Val(txtTradingPrice.Text)
'Kst = Val(txtKst.Text)
'Cst = Val(txtCst.Text)
If Not CurrencyValidate(txtTradingPrice.Text, 0) Then Exit Sub
If Kst > 0 Then If Kst > 100 Then Exit Sub
If Cst > 0 Then If Cst > 100 Then Exit Sub
txtSalesPrice = GetSalesFromTradingPrice(TRP, Kst, Cst)

Call CalculateStockValue
End Sub


