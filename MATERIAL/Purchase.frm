VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Purchase"
   ClientHeight    =   8340
   ClientLeft      =   1230
   ClientTop       =   585
   ClientWidth     =   9705
   Icon            =   "Purchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbGroupEnglish 
      Height          =   315
      ItemData        =   "Purchase.frx":030A
      Left            =   0
      List            =   "Purchase.frx":030C
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   450
      Left            =   6870
      TabIndex        =   43
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   450
      Left            =   5580
      TabIndex        =   42
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ComboBox cmbExpenseHead 
      Height          =   315
      Left            =   210
      TabIndex        =   36
      Text            =   "Expense"
      Top             =   7260
      Width           =   2325
   End
   Begin VB.ComboBox cmbIncomeHead 
      Height          =   315
      Left            =   210
      TabIndex        =   32
      Text            =   "Income"
      Top             =   6810
      Width           =   2325
   End
   Begin VB.CommandButton cmdExpense 
      Caption         =   "Add"
      Height          =   405
      Left            =   3660
      TabIndex        =   38
      Top             =   7260
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddIncome 
      Caption         =   "Less"
      Height          =   400
      Left            =   3660
      TabIndex        =   34
      Top             =   6780
      Width           =   1065
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Enabled         =   0   'False
      Height          =   405
      Left            =   8430
      TabIndex        =   30
      Top             =   5115
      Width           =   1110
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   405
      Left            =   8430
      TabIndex        =   29
      Top             =   4635
      Width           =   1110
   End
   Begin VB.TextBox txtIncome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   33
      Top             =   6780
      Width           =   915
   End
   Begin VB.TextBox txtExpense 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   37
      Top             =   7260
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   8250
      TabIndex        =   44
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Acce&pt"
      Default         =   -1  'True
      Height          =   450
      Left            =   4260
      TabIndex        =   41
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2145
      Left            =   210
      TabIndex        =   31
      Top             =   4590
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3784
      _Version        =   393216
   End
   Begin VB.Frame fra 
      Caption         =   "Products Details"
      Height          =   4035
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9315
      Begin VB.TextBox txtTax 
         Height          =   315
         Left            =   2100
         TabIndex        =   50
         Top             =   3600
         Width           =   1545
      End
      Begin VB.TextBox txtTaxAmount 
         Height          =   315
         Left            =   7170
         TabIndex        =   49
         Top             =   3600
         Width           =   1305
      End
      Begin VB.CheckBox chkRedirect 
         Caption         =   "Redirect to"
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   790
         Width           =   1725
      End
      Begin VB.ComboBox cmbRedirect 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2130
         TabIndex        =   5
         Top             =   790
         Width           =   5115
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash"
         Height          =   300
         Left            =   7680
         TabIndex        =   47
         Top             =   790
         Width           =   1575
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "Credit"
         Height          =   300
         Left            =   7680
         TabIndex        =   46
         Top             =   330
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdInvoiceDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3870
         TabIndex        =   7
         Top             =   1250
         Width           =   315
      End
      Begin VB.TextBox txtInvoiceDate 
         Height          =   345
         Left            =   2130
         TabIndex        =   8
         Top             =   1250
         Width           =   1575
      End
      Begin VB.TextBox txtInvoiceAmount 
         Height          =   315
         Left            =   7170
         TabIndex        =   16
         Top             =   1710
         Width           =   1845
      End
      Begin VB.TextBox txtInvoiceNumber 
         Height          =   345
         Left            =   2130
         TabIndex        =   14
         Top             =   1710
         Width           =   1575
      End
      Begin VB.TextBox txtTransDate 
         Height          =   345
         Left            =   7170
         TabIndex        =   11
         Top             =   1250
         Width           =   1395
      End
      Begin VB.CommandButton cmdInvoice 
         Caption         =   "..."
         Height          =   315
         Left            =   3870
         TabIndex        =   13
         Top             =   1710
         Width           =   315
      End
      Begin VB.CommandButton cmdTransDate 
         Caption         =   "..."
         Height          =   315
         Left            =   8730
         TabIndex        =   10
         Top             =   1250
         Width           =   315
      End
      Begin VB.OptionButton optSTANo 
         Caption         =   "STA No"
         Height          =   300
         Left            =   2760
         TabIndex        =   2
         Top             =   330
         Width           =   2085
      End
      Begin VB.OptionButton optInvoiceNo 
         Caption         =   "Invoice No"
         Height          =   300
         Left            =   5340
         TabIndex        =   3
         Top             =   330
         Width           =   1695
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "Release Order No"
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   2295
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         ItemData        =   "Purchase.frx":030E
         Left            =   2100
         List            =   "Purchase.frx":0310
         TabIndex        =   18
         Top             =   2345
         Width           =   2265
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   315
         ItemData        =   "Purchase.frx":0312
         Left            =   7170
         List            =   "Purchase.frx":0314
         TabIndex        =   20
         Top             =   2345
         Width           =   1875
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   7170
         TabIndex        =   24
         Top             =   2763
         Width           =   1305
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   7170
         TabIndex        =   28
         Top             =   3181
         Width           =   1305
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   2100
         TabIndex        =   26
         Top             =   3181
         Width           =   1545
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   315
         ItemData        =   "Purchase.frx":0316
         Left            =   2100
         List            =   "Purchase.frx":0318
         TabIndex        =   22
         Top             =   2763
         Width           =   2265
      End
      Begin VB.Label lblTax 
         Caption         =   "Tax in %"
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblTaxAmount 
         Caption         =   "Tax Amount"
         Height          =   300
         Left            =   5430
         TabIndex        =   51
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   150
         X2              =   9070
         Y1              =   2170
         Y2              =   2170
      End
      Begin VB.Label lblInvoiceDate 
         Caption         =   "Invoice Date"
         Height          =   300
         Left            =   210
         TabIndex        =   6
         Top             =   1250
         Width           =   1905
      End
      Begin VB.Label lblInvoiceAmount 
         Caption         =   "Invoice Amount"
         Height          =   300
         Left            =   5160
         TabIndex        =   15
         Top             =   1710
         Width           =   2055
      End
      Begin VB.Label lblInvoiceNumber 
         Caption         =   "Invoice Number"
         Height          =   300
         Left            =   210
         TabIndex        =   12
         Top             =   1710
         Width           =   1875
      End
      Begin VB.Label lblTransDate 
         Caption         =   "Transaction Date"
         Height          =   300
         Left            =   5190
         TabIndex        =   9
         Top             =   1245
         Width           =   2025
      End
      Begin VB.Label lblGroup 
         Caption         =   "Select Group"
         Height          =   300
         Left            =   210
         TabIndex        =   17
         Top             =   2345
         Width           =   1305
      End
      Begin VB.Label lblUnit 
         Caption         =   "Unit"
         Height          =   300
         Left            =   5460
         TabIndex        =   19
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label lblPerUnit 
         Caption         =   "....."
         Height          =   255
         Left            =   8610
         TabIndex        =   35
         Top             =   2820
         Width           =   585
      End
      Begin VB.Label lblPrice 
         Caption         =   "Unit Price"
         Height          =   300
         Left            =   5460
         TabIndex        =   23
         Top             =   2760
         Width           =   1545
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   300
         Left            =   5460
         TabIndex        =   27
         Top             =   3180
         Width           =   1515
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   300
         Left            =   270
         TabIndex        =   25
         Top             =   3221
         Width           =   1335
      End
      Begin VB.Label lblProductName 
         Caption         =   "Product Name"
         Height          =   420
         Left            =   120
         TabIndex        =   21
         Top             =   2723
         Width           =   1605
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   9370
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Label lblTallyFigure 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7710
      TabIndex        =   40
      Top             =   6780
      Width           =   1545
   End
   Begin VB.Label lblTotalAmount 
      Caption         =   "Total Amount"
      Height          =   345
      Left            =   5790
      TabIndex        =   39
      Top             =   6930
      Width           =   1485
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
      Height          =   360
      Left            =   2790
      TabIndex        =   45
      Top             =   30
      Width           =   3480
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkClicked()
Public Event AddClicked()
Public Event UnDoClicked()
Public Event InvoiceClicked()
Public Event ClearClicked()
Public Event PLHeadsClicked(ByVal AccType As wis_AccountType)
Public Event DeleteClicked()
Public Event GridClicked()  'Shashi 12/9/03
Public Event WindowClosed() 'SDA 2/10/2003

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_curTalliedAmount As Currency

Private m_GrdFunctions As clsGrdFunctions
Private m_VendorID As Long

Private m_InvoiceNo As String
Private m_dbOperation As wis_DBOperation

Private Function InvoiceClickValidated() As Boolean
InvoiceClickValidated = False
If Not DateValidate(txtInvoiceDate, "/", True) Then
    ActivateTextBox txtInvoiceDate
    Exit Function
End If
InvoiceClickValidated = True
End Function

Private Function PLHeadsValidated(ByVal AccType As wis_AccountType) As Boolean
PLHeadsValidated = False

If AccType = Profit Then
    If cmbIncomeHead.ListIndex = -1 Then Exit Function
    If Not TextBoxCurrencyValidate(txtIncome, True, True) Then Exit Function
ElseIf AccType = Loss Then
    If cmbExpenseHead.ListIndex = -1 Then Exit Function
    If Not TextBoxCurrencyValidate(txtExpense, True, True) Then Exit Function
End If

PLHeadsValidated = True
End Function
Private Sub SetInvoiceLabel()
If optRONo.Value Then
    lblInvoiceDate.Caption = GetResourceString(225) & " " & _
                             GetResourceString(37) '"RO Date"
    lblInvoiceAmount.Caption = GetResourceString(225) & " " & _
                               GetResourceString(40) '"RO Amount"
    lblInvoiceNumber.Caption = GetResourceString(225) & " " & _
                               GetResourceString(60) '"RO Number"
ElseIf optSTANo.Value Then
    lblInvoiceDate.Caption = GetResourceString(226) & " " & _
                             GetResourceString(37) '"STA Date"
    lblInvoiceAmount.Caption = GetResourceString(226) & " " & _
                               GetResourceString(40) '"STA Amount"
    lblInvoiceNumber.Caption = GetResourceString(226) & " " & _
                               GetResourceString(60) '"STA Number"
Else
    lblInvoiceDate.Caption = GetResourceString(172) & " " & _
                             GetResourceString(37) '"Invoice Date"
    lblInvoiceAmount.Caption = GetResourceString(172) & " " & _
                               GetResourceString(40) '"Invoice Amount"
    lblInvoiceNumber.Caption = GetResourceString(172) & " " & _
                               GetResourceString(60) '"Invoice Number"
End If


End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls
'set the title for the Tabs
'tabPurchase.Font.Name = gFontName

'fra(0).Caption = GetResourceString(172,60) & _
                                                          GetResourceString(295)

fra.Caption = GetResourceString(158, 295)
lblCompanyName.Caption = GetResourceString(138, 35)


optRONo.Caption = GetResourceString(225, 60)
optSTANo.Caption = GetResourceString(226, 60)
optInvoiceNo.Caption = GetResourceString(172, 60)
chkRedirect.Caption = GetResourceString(207)

lblInvoiceDate.Caption = GetResourceString(172, 37)
lblTransDate.Caption = GetResourceString(38, 37)
lblInvoiceNumber.Caption = GetResourceString(172, 60)
lblInvoiceAmount.Caption = GetResourceString(172, 40)
lblTotalAmount.Caption = GetResourceString(52, 40)
cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)
cmdAdd.Caption = GetResourceString(10)
cmdCancel.Caption = GetResourceString(11)
cmdClear.Caption = GetResourceString(8)
cmdDelete.Caption = GetResourceString(14)
cmdAddIncome.Caption = GetResourceString(10)
cmdExpense.Caption = GetResourceString(255)

'lblManufacturer.Caption = GetResourceString(174)
lblGroup.Caption = GetResourceString(786)
lblUnit.Caption = GetResourceString(161)
lblProductName.Caption = GetResourceString(158)
lblPrice.Caption = GetResourceString(305)
lblQuantity.Caption = GetResourceString(306)
'lblFreeItem.Caption = GetResourceString(105)
lblAmount.Caption = GetResourceString(40)
lblTax.Caption = GetResourceString(173) & "(%)"
lblTaxAmount.Caption = GetResourceString(173, 40)
'Label3.Caption = GetResourceString(105)
'lblFreeItem.Caption = GetResourceString(105)
End Sub


Private Sub CalculateInvoiceAmount()
Dim curDiscount As Currency
Dim curTaxPaid As Currency

If Not CurrencyValidate(txtIncome.Text, True) Then
    'MsgBox "Invalid amount specified!", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If Not CurrencyValidate(txtExpense.Text, True) Then
    'MsgBox "Invalid amount specified!", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
curDiscount = Val(txtIncome.Text)

curTaxPaid = Val(txtExpense)

lblTallyFigure.Caption = m_curTalliedAmount - curDiscount + curTaxPaid

End Sub


Private Sub LoadProductDetails()
'Declare the variables
Dim lngRelationID As Long
Dim curTradingprice As Currency

'Dim sglCST As Single
'Dim sglKST As Single

Dim rstPurchasePrice As ADODB.Recordset
Dim MaterialClass As clsMaterial
Dim TransDate As String
 
With cmbGroup
    If .ListIndex = -1 Then Exit Sub
    'lngRelationID = .ItemData(.ListIndex)
End With
With cmbUnit
    If .ListIndex = -1 Then Exit Sub
    'lngRelationID = .ItemData(.ListIndex)
End With
With cmbProductName
    If .ListIndex = -1 Then Exit Sub
    lngRelationID = .ItemData(.ListIndex)
End With

gDbTrans.SqlStmt = " SELECT TradingPrice,MRP,SalesPrice,Tax FROM RelationMaster " & _
                   " WHERE RelationID = " & lngRelationID
  
If gDbTrans.Fetch(rstPurchasePrice, adOpenForwardOnly) < 1 Then Exit Sub

curTradingprice = FormatField(rstPurchasePrice("TradingPrice"))

'sglCST = FormatField(rstPurchasePrice("CST"))
'sglKST = FormatField(rstPurchasePrice("KST"))

'txtPrice.Text = GetAddedPercentage(curTradingprice, sglKST)

'If sglCST > 0 Then txtPrice.Text = GetAddedPercentage(curTradingprice, sglCST)
txtPrice.Text = curTradingprice
txtPrice.Tag = curTradingprice
txtTax.Text = FormatField(rstPurchasePrice("tax"))
lblPerUnit.Caption = GetResourceString(439) & " " & cmbUnit.Text

Set MaterialClass = New clsMaterial

TransDate = Trim$(txtTransDate.Text)
If Not DateValidate(TransDate, "/", True) Then
    'MsgBox "Invalid Date Specifed ", vbInformation
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Sub
End If
'Godown ID 1 Because HO is Always 1
txtQuantity.Text = MaterialClass.GetItemOnDateClosingStock(lngRelationID, TransDate, 1)

Set MaterialClass = Nothing

End Sub

Private Sub LoadProducts()

'Declare the variables
Dim lngGroupID As Long
Dim UnitID As Integer

Dim rstProduct As ADODB.Recordset
Dim GodownID As Byte
cmbProductName.Clear

With cmbGroup
    If .ListIndex = -1 Then Exit Sub
    'Get the respective Ids from Item data
    lngGroupID = .ItemData(.ListIndex)
End With
With cmbUnit
    If .ListIndex = -1 Then Exit Sub
    'Get the respective Ids from Item data
    UnitID = .ItemData(.ListIndex)
End With


gDbTrans.SqlStmt = "SELECT Distinct A.ProductID,B.ProductName,RelationID " & _
                    " FROM RelationMaster A, Products B" & _
                    " WHERE A.GroupID = " & lngGroupID & _
                    " AND A.GroupID = B.GroupID " & _
                    " AND A.ProductID = B.ProductID " & _
                    " AND PriceChanged = " & 0 & _
                    " AND A.UnitID = " & UnitID & _
                    " AND GodownID = " & 1

cmbProductName.Clear
If gDbTrans.Fetch(rstProduct, adOpenForwardOnly) < 1 Then Exit Sub

'Load the data to the combo box
With cmbProductName
    Do While Not rstProduct.EOF
       .AddItem FormatField(rstProduct("ProductName"))
       .ItemData(.newIndex) = FormatField(rstProduct("RelationID"))
       'Move the recordset
       rstProduct.MoveNext
    Loop
    If .ListCount = 1 Then .ListIndex = 0
End With

End Sub

Private Sub SetLastTransactionDate()
'Declare the variables
Dim rst As ADODB.Recordset
Dim headID As Long

If m_VendorID = 0 Then Exit Sub

gDbTrans.SqlStmt = " SELECT TOP 1 InvoiceNo,TransDate,InvoiceDate" & _
                " FROM Purchase" & _
                " WHERE HeadID = " & m_VendorID & _
                " ORDER BY TransDate DESC"
                  
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

txtInvoiceDate.Text = FormatField(rst.Fields("InvoiceDate"))
txtTransDate.Text = FormatField(rst.Fields("TransDate"))
txtInvoiceNumber.Text = FormatField(rst.Fields("InvoiceNo"))

End Sub

Public Property Get VendorID() As Long
VendorID = m_VendorID
End Property

Public Property Let VendorID(ByVal vNewValue As Long)
m_VendorID = vNewValue
End Property
Private Sub ClearControls()

'cmbGroup.ListIndex = -1
cmbUnit.ListIndex = -1
cmbProductName.ListIndex = -1
txtPrice.Text = ""
txtQuantity.Text = ""
'txtFreeQuantity.Text = ""
txtAmount.Text = ""

cmdUndo.Enabled = True
cmdAccept.Enabled = True
cmdDelete.Enabled = False
On Error Resume Next
cmbGroup.SetFocus
End Sub

Private Function DeleteSelectedRow() As Boolean

Dim SelectedRow As Integer
Dim TotalAmount As Currency
Dim Amount As Currency

Dim count As Integer
Dim InvoiceNo As String
Dim ProductName As String
Dim Quantity As String
Dim FreeQuantity As String
Dim Rate As String
Dim Total As String
Dim ParentID As Long

If grd.Row = 0 Then Exit Function

SelectedRow = grd.Row

' if The selected row is not containing any item then exit

grd.Col = 0

If Trim$(grd.Text) = "" Then Exit Function

grd.Col = 1

'first  store total in a variable
With grd

   .Row = SelectedRow
   Do
      .Row = .Row + 1
      .Col = 0
      If Trim$(.Text) = "" Then
         .Col = .Cols - 1
         TotalAmount = Val(.Text)
         Exit Do
      End If
   Loop
   
   'Get Selected AMount
   .Row = SelectedRow
   .Col = .Cols - 1
   Amount = Val(.Text)
   
   
   TotalAmount = TotalAmount - Amount
   m_curTalliedAmount = TotalAmount
   lblTallyFigure.Caption = m_curTalliedAmount
   'Now rearrange the Grid
   If TotalAmount = 0 Then
      InitGrid
      Exit Function
   End If
   

   m_GrdFunctions.fMoreRows (2)
   
   .Row = .Row + 1
    
   For count = SelectedRow To grd.Rows - 2
      'Get the Values of Next row
      .Row = count + 1: .Col = 0
      If .Text <> "" Then
         .Col = 1: InvoiceNo = .Text: .Text = ""
         .Col = 2: ProductName = .Text: ParentID = .RowData(2): .Text = ""
         .Col = 3: Quantity = .Text: .Text = ""
         .Col = 4: FreeQuantity = .Text: .Text = ""
         .Col = 5: Rate = .Text: .Text = ""
         .Col = 6: Total = .Text: .Text = ""
      Else
         'if it this is last row then Asign the total value and exit sub
         .Row = count
         .Col = 0: .Text = ""
         .Col = 1: .Text = "Total Amount": .CellFontBold = True
         .Col = 2: .Text = ""
         .Col = 3: .Text = ""
         .Col = 4: .Text = ""
         .Col = 5: .Text = ""
         
         .Col = .Cols - 1: .Text = FormatCurrency(TotalAmount): .CellFontBold = True
         'Clear the next row
         .Row = .Row + 1
         .Col = 1: .Text = ""
         .Col = 6: .Text = ""
         
         'Reduce the Rows by 1 No
         .Rows = .Rows - 1
         
         Exit Function
      End If
      
      'Assign the values of Next row to the current row
      .Row = count:
      .Col = 0: .Text = count
      .Col = 1: .Text = InvoiceNo
      .Col = 2: .Text = ProductName: .RowData(2) = ParentID
      .Col = 3: .Text = Quantity
      .Col = 4: .Text = FreeQuantity
      .Col = 5: .Text = Rate
      .Col = 6: .Text = Total
   Next count
End With
End Function

'private grdFunctions as clsgr
Private Sub GridResize()
    Dim Ratio  As Single
    
    With grd
       Ratio = .Width / .Cols
       
       .ColWidth(0) = Ratio * 0.5
       '.ColWidth(1) = Ratio * 1.45
       .ColWidth(1) = Ratio * 0.95
       .ColWidth(2) = Ratio * 0.95
       .ColWidth(3) = Ratio * 0.9
       .ColWidth(4) = Ratio * 0.9
       .ColWidth(5) = Ratio * 0.95
       .ColWidth(6) = Ratio * 0.95
       .ColWidth(7) = Ratio * 0.95
    End With

End Sub

Public Sub InitGrid()

With grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 8
    .FixedRows = 1
    .Row = 0
    
    .Col = 0: .Text = GetResourceString(33): .CellFontBold = True '"SlNo"
    '.Col = 1: .Text = GetResourceString(174): .CellFontBold = True '"Manufacturer"
    .Col = 1: .Text = GetResourceString(39): .CellFontBold = True '"Description"
    .Col = 2: .Text = GetResourceString(306): .CellFontBold = True '"Quantity"
    '.Col = 3: .Text = "Free Quantity": .CellFontBold = True
    .Col = 3: .Text = GetResourceString(176) & " " & _
                      GetResourceString(305): .CellFontBold = True '"Purchase Price"
    
    .Col = 4: .Text = GetResourceString(163): .CellFontBold = True '"MRP"
    .Text = GetResourceString(40) 'Amount
    
    .Col = 5: .Text = GetResourceString(162, 212): .CellFontBold = True '"TRP"
    
    .Col = 6: .Text = GetResourceString(173): .CellFontBold = True '"TAx"
    .Col = 7: .Text = GetResourceString(304): .CellFontBold = True '"Total Amount"
End With

End Sub

Private Sub LoadProductGroups()
Dim rstGroups As ADODB.Recordset

cmbGroup.Clear
cmbGroupEnglish.Clear
gDbTrans.SqlStmt = "SELECT GroupID,GroupName,GroupNameEnglish FROM ProductGroup " & _
                 " ORDER BY GroupID;"
        
Call gDbTrans.Fetch(rstGroups, adOpenForwardOnly)

Do
   If rstGroups.EOF Then Exit Sub
   cmbGroup.AddItem FormatField(rstGroups("GroupName"))
   cmbGroup.ItemData(cmbGroup.newIndex) = FormatField(rstGroups("GroupID"))
   cmbGroupEnglish.AddItem FormatField(rstGroups("GroupNameEnglish"))
   cmbGroupEnglish.ItemData(cmbGroupEnglish.newIndex) = FormatField(rstGroups("GroupID"))
   
   'Move to next record
   rstGroups.MoveNext
Loop

End Sub
Private Sub LoadUnits()

Dim rstUnits As ADODB.Recordset

'Declare the variables
Dim lngGroupID As Long
Dim GodownID As Byte

If cmbGroup.ListIndex = -1 Then Exit Sub

'Get the respective Ids from Item data
lngGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)

gDbTrans.SqlStmt = "SELECT Distinct UnitID,UnitName FROM Units"
If lngGroupID Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where UnitID IN " & _
            "(select Distinct UnitId From RelationMAster)"

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order By UnitId"

cmbUnit.Clear
If gDbTrans.Fetch(rstUnits, adOpenForwardOnly) < 1 Then Exit Sub

With cmbUnit
    Do
       If rstUnits.EOF Then Exit Sub
       .AddItem FormatField(rstUnits("UnitName"))
       .ItemData(.newIndex) = FormatField(rstUnits("UnitID"))
       'Move to next record
       rstUnits.MoveNext
    Loop
    If .ListCount = 1 Then .ListIndex = 0
End With

End Sub

Private Function LoadDataToGrid() As Boolean
'Declare the variables
Dim lngRelationID As Long
Dim RowNum As Integer

Dim curTotal As Currency
Dim InvoiceAmount As Currency
Dim curTradingprice As Currency
Dim curMRP As Currency
Dim curAmount As Currency

Dim rstPurchasePrice As ADODB.Recordset

Dim dblQuantity As Double

If cmbProductName.ListIndex = -1 Then Exit Function

lngRelationID = cmbProductName.ItemData(cmbProductName.ListIndex)

gDbTrans.SqlStmt = " SELECT TradingPrice,MRP FROM RelationMaster " & _
                   " WHERE RelationID = " & lngRelationID
  
If gDbTrans.Fetch(rstPurchasePrice, adOpenForwardOnly) < 1 Then Exit Function

curTradingprice = FormatField(rstPurchasePrice("TradingPrice"))
curMRP = FormatField(rstPurchasePrice("MRP"))
dblQuantity = Val(txtQuantity.Text)
curAmount = dblQuantity * curTradingprice


RowNum = 1
With grd
   For RowNum = 1 To .Rows
      .Row = RowNum
      'First Check For the mateial existance in the grid
      .Col = 0
      If Trim$(.Text) = "" Then Exit For
      
      .Col = .Cols - 1
      curTotal = curTotal + Val(.Text)
   Next RowNum
   
   .Col = 0: .Text = RowNum: .CellFontBold = False
   .Col = 1: .Text = cmbProductName.Text: .RowData(RowNum) = cmbProductName.ItemData(cmbProductName.ListIndex): .CellFontBold = False
   .Col = 2: .Text = dblQuantity
   '.Col = 3: .Text = txtFreeQuantity.Text
   .Col = 3: .Text = txtPrice.Text
   '.Col = 4: .Text = txtFreeAmount.Text
   
   .Col = 4: .Text = ""
   If curMRP > 0 Then .Col = 6: .Text = curMRP
   
   .Col = 5: .Text = curTradingprice
   .Col = 6: .Text = curAmount: .CellFontBold = False
   
   'Show the total
   RowNum = RowNum + 1
   m_GrdFunctions.fMoreRows (2)
   .Row = .Row + 1
   .Col = 1: .Text = "Total Amount": .CellFontBold = True
   curTotal = curTotal + curAmount
   .Col = 6: .Text = curTotal: .CellFontBold = True
   m_curTalliedAmount = curTotal
End With

InvoiceAmount = Val(txtInvoiceAmount.Text)

If m_curTalliedAmount >= InvoiceAmount Then _
                txtIncome.Text = m_curTalliedAmount - InvoiceAmount

lblTallyFigure.Caption = m_curTalliedAmount - Val(txtIncome.Text)

LoadDataToGrid = True
End Function
Private Function Validated() As Boolean

Validated = False

If Not TextBoxDateValidate(txtInvoiceDate, "/", True, True) Then Exit Function
   
If Not TextBoxDateValidate(txtTransDate, "/", True, True) Then Exit Function
   
If txtInvoiceNumber.Text = "" Then
   'MsgBox "Please specify the invoice number", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(813), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtInvoiceNumber
   Exit Function
End If

If Not TextBoxCurrencyValidate(txtInvoiceAmount, True, True) Then Exit Function
  
If cmbGroup.ListIndex = -1 Then
   'MsgBox "Please select the Group", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(786), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbGroup.SetFocus
   Exit Function
End If

If cmbUnit.ListIndex = -1 Then
   'MsgBox "Please Select the Unit", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(820), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbUnit.SetFocus
   Exit Function
End If

If cmbProductName.ListIndex = -1 Then
   'MsgBox "Please Select the Product name", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetChangeString(GetResourceString(230), GetResourceString(158)), vbInformation, wis_MESSAGE_TITLE
   On Error Resume Next
   cmbGroup.SetFocus
   Exit Function
End If

If Not TextBoxCurrencyValidate(txtPrice, True, True) Then Exit Function
   
If Not IsNumeric(txtQuantity.Text) Then
   'MsgBox "Invalid Quantity sepecified!", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(776), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtQuantity
   Exit Function
End If


If Not TextBoxCurrencyValidate(txtAmount, True, True) Then Exit Function

If Val(txtInvoiceAmount.Text) < Val(txtAmount.Text) Then
    MsgBox "Invoice Amount is Lesser than Item Amount"
    Exit Function
End If

Validated = True

End Function

Private Sub chkRedirect_Click()

Dim MaterialClass As clsMaterial

Set MaterialClass = New clsMaterial

cmbRedirect.Enabled = False

If chkRedirect.Value = vbChecked Then
    cmbRedirect.Enabled = True
    'cmbRedirect.Visible = True
    Call MaterialClass.LoadCompaniesToCombo(Enum_Customers, cmbRedirect)
End If

Set MaterialClass = Nothing

End Sub

Private Sub cmbExpenseHead_Click()
If cmbExpenseHead.ListIndex = -1 Then Exit Sub
txtExpense.Text = ""


End Sub


Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
'Load the Products to the products combo box
Call LoadUnits
'If cmbUnit.ListCount = 1 Then Call ActivateTextBox(txtQuantity)
Call LoadProducts

End Sub


Private Sub cmbIncomeHead_Click()
If cmbIncomeHead.ListIndex = -1 Then Exit Sub
txtIncome.Text = ""


End Sub


Private Sub cmbProductName_Click()

Call LoadProductDetails

End Sub




Private Sub cmbUnit_Click()
'Load the details of the product
Call LoadProducts
If cmbProductName.ListIndex = 1 Then Call ActivateTextBox(txtQuantity)
'Call LoadProductDetails
End Sub


Private Sub cmdAccept_Click()

RaiseEvent OkClicked
grd.Row = 1
grd.Col = 0
If Len(grd.Text) = 0 Then ReDim m_NewRelationId(0, 1)

Me.MousePointer = vbDefault

End Sub

'
Private Sub cmdAdd_Click()

If Not Validated Then Exit Sub

RaiseEvent AddClicked

End Sub

Private Sub cmdAddIncome_Click()
If Not PLHeadsValidated(Profit) Then Exit Sub

RaiseEvent PLHeadsClicked(Profit)

End Sub

'
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdClear_Click()
RaiseEvent ClearClicked

Dim MaxCount As Integer
Dim count As Integer


End Sub

Private Sub cmdDelete_Click()
RaiseEvent DeleteClicked
End Sub

Private Sub cmdExpense_Click()
If Not PLHeadsValidated(Loss) Then Exit Sub

RaiseEvent PLHeadsClicked(Loss)
End Sub

'
Private Sub cmdInvoice_Click()

If Not InvoiceClickValidated Then Exit Sub

RaiseEvent InvoiceClicked

End Sub

Private Sub cmdInvoicedate_Click()
With Calendar
    .selDate = IIf(DateValidate(txtInvoiceDate, "/", True), txtInvoiceDate, gStrDate)
    
    .Left = Left + fra.Left + cmdInvoicedate.Left
    .Top = Top + fra.Top + cmdInvoicedate.Top - .Height / 2
    .Show vbModal
    txtInvoiceDate = .selDate
End With

End Sub

Private Sub cmdTransDate_Click()
With Calendar
    If DateValidate(txtTransDate, "/", True) Then .selDate = txtTransDate
    .Left = Left + fra.Left + cmdTransDate.Left
    .Top = Top + fra.Top + cmdTransDate.Top - .Height / 2
    .Show vbModal
    txtTransDate = .selDate
End With
End Sub

'
Private Sub cmdUndo_Click()
'Call DeleteSelectedRow
RaiseEvent UnDoClicked
End Sub



Private Sub Form_Initialize()
If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
Set m_GrdFunctions.fGrd = grd
End Sub

'
Private Sub Form_Resize()
GridResize

'Set the focus on to first tabs Invoice date

'Set Last Transaction Date
Call SetLastTransactionDate

'ActivateTextBox txtInvoiceDate
optRONo.SetFocus

End Sub

'
Private Sub Form_Unload(Cancel As Integer)
RaiseEvent WindowClosed

End Sub


Private Sub grd_Click()
RaiseEvent GridClicked
End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_InvoiceNo = strSelection
End Sub

'
Private Sub optInvoiceNo_Click()
Call SetInvoiceLabel
End Sub

'
Private Sub optRONo_Click()
Call SetInvoiceLabel
End Sub


'
Private Sub optSTANo_Click()
Call SetInvoiceLabel
End Sub


Private Sub txtDiscountAmount_LostFocus()
If txtIncome = "" Then Exit Sub
If Val(txtIncome) = 0 Then Exit Sub

CalculateInvoiceAmount

End Sub







Private Sub txtAmount_Change()
Call txtTax_Change
'If Val(txtQuantity.Text) = 0 Then Exit Sub

'txtPrice.Text = FormatCurrency(Val(txtAmount.Text) / Val(txtQuantity.Text), True)

End Sub

'
Private Sub txtInvoiceAmount_LostFocus()
On Error Resume Next
cmbGroup.SetFocus


End Sub

'
Private Sub txtInvoiceDate_GotFocus()
ActivateTextBox txtInvoiceDate
End Sub


'
Private Sub txtInvoiceNumber_GotFocus()
ActivateTextBox txtInvoiceNumber
End Sub


Private Sub txtPrice_Change()
txtAmount.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtPrice_GotFocus()
With txtPrice
    txtPrice.SelStart = 0
    txtPrice.SelLength = Len(txtPrice)
End With
End Sub


Private Sub txtQuantity_Change()
    txtAmount.Text = FormatCurrency(Val(txtQuantity.Text) * Val(txtPrice.Text))
End Sub


Private Sub txtQuantity_GotFocus()
    With txtQuantity
        txtQuantity.SelStart = 0
        txtQuantity.SelLength = Len(txtQuantity)
    End With

End Sub











Private Sub Form_Load()
'Declare the variables
Dim MaterialClass As clsMaterial
Dim CompanyType As wis_CompanyType

'Center the form
CenterMe Me

'fraRedirect.BorderStyle = 0
'set icon for the form caption
Me.Icon = LoadResPicture(147, vbResIcon)

Call SetKannadaCaption

If MaterialClass Is Nothing Then Set MaterialClass = New clsMaterial

'Get the company type
CompanyType = MaterialClass.GetCompanyType(m_VendorID)
' if the company type is Stockist  the load manufacturers

'cmbManufacturer.Enabled = False
'lblManufacturer.Enabled = False

If CompanyType = Enum_Stockist Then
'    cmbManufacturer.Enabled = True
'    lblManufacturer.Enabled = True
    Debug.Print "MFR"
    'Call MaterialClass.LoadCompaniesToCombo(Enum_Manufacturer, cmbManufacturer)
End If


'Load Group Combo box
LoadProductGroups

'Load Unit combo box
LoadUnits

'Initalise the grid.
InitGrid
txtInvoiceDate = gStrDate
txtTransDate = txtInvoiceDate

'Load Income Heads
'Call LoadHeadsToCombo(cmbIncomeHead, Profit)
'Load Trading Income Heads
Dim showAllHeads As Boolean
showAllHeads = False
Dim SetUp As New clsSetup
If UCase(SetUp.ReadSetupValue("Trading", "ShowAllExpenseHeads", "True")) = "TRUE" Then showAllHeads = True
Set SetUp = Nothing

Call LoadLedgersToCombo(cmbIncomeHead, parTradingIncome, True)
'Then Load other Income Heads
If showAllHeads Then Call LoadLedgersToCombo(cmbIncomeHead, parIncome, False)

'Load Expense Heads
'Call LoadHeadsToCombo(cmbExpenseHead, Loss)
'Load Trading Expense Heads
Call LoadLedgersToCombo(cmbExpenseHead, parTradingExpense, True)
'Then Load other Income Heads
If showAllHeads Then Call LoadLedgersToCombo(cmbExpenseHead, parExpense, False)

cmdUndo.Enabled = False
cmdAccept.Enabled = False
cmdDelete.Enabled = False

Set MaterialClass = Nothing

ReDim m_NewRelationId(1, 0)

End Sub
Public Property Get TallyAmount() As Currency
TallyAmount = m_curTalliedAmount
End Property

Public Property Let TallyAmount(ByVal vNewValue As Currency)
m_curTalliedAmount = vNewValue
End Property

Private Sub txtTax_Change()
If Val(txtTax.Text) > 0 Then txtTaxAmount.Text = FormatCurrency(Val(txtTax.Text) / 100 * Val(txtAmount.Text))

End Sub

Private Sub txtTransDate_GotFocus()
ActivateTextBox txtTransDate
End Sub



Public Property Get DBOperation() As wis_DBOperation
DBOperation = m_dbOperation
End Property

Public Property Let DBOperation(ByVal vNewValue As wis_DBOperation)
m_dbOperation = vNewValue
End Property
