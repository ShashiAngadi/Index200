VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Sales"
   ClientHeight    =   8415
   ClientLeft      =   1830
   ClientTop       =   990
   ClientWidth     =   10050
   Icon            =   "Sales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbGroupEnglish 
      Height          =   315
      ItemData        =   "Sales.frx":030A
      Left            =   720
      List            =   "Sales.frx":030C
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Frame fra 
      Caption         =   "Product Details"
      Height          =   3375
      Left            =   210
      TabIndex        =   0
      Top             =   480
      Width           =   9705
      Begin VB.TextBox txtTaxAmount 
         Height          =   315
         Left            =   7020
         TabIndex        =   25
         Top             =   2400
         Width           =   1425
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "Credit"
         Height          =   195
         Left            =   7050
         TabIndex        =   44
         Top             =   3030
         Width           =   1575
      End
      Begin VB.OptionButton optCash 
         Caption         =   "Cash"
         Height          =   195
         Left            =   1770
         TabIndex        =   43
         Top             =   3030
         Width           =   1605
      End
      Begin VB.CommandButton cmdInvoicedate 
         Caption         =   "..."
         Height          =   315
         Left            =   4020
         TabIndex        =   4
         Top             =   690
         Width           =   315
      End
      Begin VB.TextBox txtInvoiceNumber 
         Height          =   315
         Left            =   7020
         TabIndex        =   8
         Top             =   690
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalesInvoice 
         Caption         =   "..."
         Height          =   315
         Left            =   8850
         TabIndex        =   7
         Top             =   690
         Width           =   330
      End
      Begin VB.TextBox txtInvoiceDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   690
         Width           =   2055
      End
      Begin VB.ComboBox cmbBranch 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   210
         Width           =   7455
      End
      Begin VB.TextBox txtSoot 
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   2400
         Width           =   1305
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         ItemData        =   "Sales.frx":030E
         Left            =   1800
         List            =   "Sales.frx":0310
         TabIndex        =   10
         Top             =   1140
         Width           =   2505
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   315
         ItemData        =   "Sales.frx":0312
         Left            =   7020
         List            =   "Sales.frx":0314
         TabIndex        =   12
         Top             =   1110
         Width           =   2235
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   7020
         TabIndex        =   16
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   7020
         TabIndex        =   21
         Top             =   1980
         Width           =   1425
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   1980
         Width           =   1335
      End
      Begin VB.ComboBox cmbProductName 
         Height          =   315
         ItemData        =   "Sales.frx":0316
         Left            =   1800
         List            =   "Sales.frx":0318
         TabIndex        =   14
         Top             =   1560
         Width           =   2505
      End
      Begin VB.Label lblTaxAmount 
         Caption         =   "Tax Amount"
         Height          =   225
         Left            =   5070
         TabIndex        =   24
         Top             =   2430
         Width           =   1275
      End
      Begin VB.Label lblInvoiceNumber 
         Caption         =   "Invoice Number"
         Height          =   285
         Left            =   5010
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblInvoiceDate 
         Caption         =   "Invoice Date"
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lblBranch 
         Caption         =   "Select Branch"
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   2235
      End
      Begin VB.Label lblSoot 
         Caption         =   "Soot"
         Height          =   315
         Left            =   210
         TabIndex        =   22
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label lblGroup 
         Caption         =   "Select Group"
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   1140
         Width           =   1485
      End
      Begin VB.Label lblUnit 
         Caption         =   "Unit"
         Height          =   225
         Left            =   5040
         TabIndex        =   11
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label lblPerUnit 
         Caption         =   "....."
         Height          =   225
         Left            =   8610
         TabIndex        =   17
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblPrice 
         Caption         =   "Unit Price"
         Height          =   225
         Left            =   5010
         TabIndex        =   15
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5040
         TabIndex        =   20
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   405
         Left            =   210
         TabIndex        =   18
         Top             =   1980
         Width           =   1485
      End
      Begin VB.Label lblProductName 
         Caption         =   "Product Name"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdPrint 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5070
      Width           =   570
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   400
      Left            =   7020
      TabIndex        =   39
      Top             =   7890
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   400
      Left            =   5550
      TabIndex        =   38
      Top             =   7890
      Width           =   1335
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
      TabIndex        =   33
      Top             =   7110
      Width           =   915
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
      TabIndex        =   30
      Top             =   6690
      Width           =   915
   End
   Begin VB.CommandButton cmdAddIncome 
      Caption         =   "Add"
      Height          =   400
      Left            =   4080
      TabIndex        =   31
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton cmdExpense 
      Caption         =   "Less"
      Height          =   400
      Left            =   4110
      TabIndex        =   34
      Top             =   7140
      Width           =   915
   End
   Begin VB.ComboBox cmbIncomeHead 
      Height          =   315
      Left            =   270
      TabIndex        =   29
      Text            =   "Income"
      Top             =   6690
      Width           =   2325
   End
   Begin VB.ComboBox cmbExpenseHead 
      Height          =   315
      Left            =   270
      TabIndex        =   32
      Text            =   "Expense"
      Top             =   7140
      Width           =   2325
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Enabled         =   0   'False
      Height          =   400
      Left            =   8880
      TabIndex        =   27
      Top             =   4575
      Width           =   1020
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   400
      Left            =   8880
      TabIndex        =   26
      Top             =   4095
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8490
      TabIndex        =   40
      Top             =   7890
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Acce&pt"
      Height          =   400
      Left            =   4080
      TabIndex        =   37
      Top             =   7890
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2625
      Left            =   270
      TabIndex        =   28
      Top             =   3900
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   4630
      _Version        =   393216
   End
   Begin VB.Label lblTotalAmount 
      Caption         =   "Total Amount"
      Height          =   255
      Left            =   5580
      TabIndex        =   35
      Top             =   6660
      Width           =   1425
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
      ForeColor       =   &H80000007&
      Height          =   345
      Left            =   7410
      TabIndex        =   36
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   270
      X2              =   9370
      Y1              =   7680
      Y2              =   7680
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
      Left            =   3390
      TabIndex        =   41
      Top             =   90
      Width           =   2190
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Save to the data bse
Public Event OkClicked()
Public Event InvoiceClicked()
'Add to the grid
Public Event AddClicked()
Public Event UnDoClicked()
Public Event ClearClicked()
Public Event PLHeadsClicked(ByVal AccType As wis_AccountType)
Public Event DeleteClicked()
Public Event PrintClicked()

Public Event GridClicked() 'Shashi 4/9/03
Public Event WindowClosed()
Public Event ChangeCustomer()

Private m_GrdFunctions As clsGrdFunctions

Private m_VendorID As Long
Private m_InvoiceNo As String
Private m_InvoiceAmount As Currency
Private m_StockAvailable As Double

Private Function InvoiceClickValidated() As Boolean

InvoiceClickValidated = False

If cmbBranch.ListIndex = -1 Then
    MsgBox "Please Select the branch", vbInformation, wis_MESSAGE_TITLE
    cmbBranch.SetFocus
    Exit Function
End If
If Not DateValidate(txtInvoiceDate, "/", True) Then
    MsgBox "Please Specify the Invoice Date", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtInvoiceDate
    Exit Function
End If
InvoiceClickValidated = True
End Function

Private Function InvoiceNoExists(ByVal InvoiceNum As String) As Boolean
Dim rst As ADODB.Recordset

InvoiceNoExists = False

gDbTrans.SqlStmt = " SELECT InvoiceNo FROM Sales " & _
                   " WHERE InvoiceNo =" & AddQuotes(InvoiceNum, True)

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

InvoiceNoExists = True

Set rst = Nothing

End Function

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls
fra.Caption = GetResourceString(158, 295)

lblInvoiceDate.Caption = GetResourceString(172, 37)
lblInvoiceNumber.Caption = GetResourceString(172, 60)

cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)
cmdAdd.Caption = GetResourceString(10)
cmdCancel.Caption = GetResourceString(11)
cmdClear.Caption = GetResourceString(8)
cmdDelete.Caption = GetResourceString(14)
cmdExpense.Caption = GetResourceString(10)
cmdAddIncome.Caption = GetResourceString(255)

lblBranch.Caption = GetResourceString(227) '& " " & GetResourceString(27)
lblGroup.Caption = GetResourceString(157)
lblUnit.Caption = GetResourceString(161)
lblProductName.Caption = GetResourceString(158)
lblPrice.Caption = GetResourceString(305)
lblQuantity.Caption = GetResourceString(306)
lblAmount.Caption = GetResourceString(40)
lblTotalAmount = GetResourceString(52, 40)

End Sub

Private Sub DeleteDiscountInGrid()
'declare the variables
Dim SelectedRow As Integer
Dim TotalAmount As Currency
Dim Amount As Currency

Dim ItemName As String
Dim SlNo As String
With grd
    SelectedRow = .Row
    'If the selected row is TotalAmount then exit
    .Col = 0
    If Trim$(.Text) = "" Then Exit Sub
    
    'Get the amount
    .Col = 0: SlNo = .Text
    .Col = .Cols - 1
    Amount = Val(.Text)
    
    'Get the Total amount
    .Row = .Row + 1
    .Col = .Cols - 1
    TotalAmount = Val(.Text)
        
    'Get the new Total Amount
    TotalAmount = TotalAmount + Amount
    
    'Now delete the selected row
    'Move to the Next row if there are no data then just delete  the selected row
    .Row = .Row + 1
    .Col = 2
    If Trim$(.Text) = "" Then
        .Row = SelectedRow
        .Col = 0: .Text = ""
        .Col = 2: .Text = ""
        .Col = .Cols - 1: .Text = ""
        
        'Move to the next row
        .Row = .Row + 1
        .Col = 0: .Text = ""
        .Col = 2: .Text = ""
        .Col = .Cols - 1: .Text = ""
        
        'Move to the Previous than the selectedrow
        .Row = SelectedRow - 1
        .Col = 2: .Text = "Total Amount ": .CellFontBold = True
        .Col = .Cols - 1: .Text = TotalAmount: .CellFontBold = True
    Else
        'Now store the values
        .Row = .Row + 1
        .Col = 0: SlNo = .Text: .Text = ""
        .Col = 2: ItemName = .Text = ""
        .Col = .Cols - 1: Amount = Val(.Text): .Text = ""
        
        'Goto to the previvois two rows
        .Row = .Row - 1
        .Col = 0: .Text = SlNo
        .Col = 2: .Text = ItemName
        .Col = .Cols - 1: .Text = Amount
        
        'Move to next row
        .Row = .Row + 1
        .Col = 2: .Text = "Total Amount": .CellFontBold = True
        .Col = .Cols - 1: .Text = TotalAmount: .CellFontBold = True
        
        'reduce the rows
        .Rows = .Rows - 1
    End If
    
End With


End Sub

Private Function DeleteSelectedRow() As Boolean
'Declare the variables
Dim SelectedRow As Integer
Dim lpCount As Integer

Dim Amount As Currency
Dim TotalAmount As Currency
Dim Total As Currency 'Inclusive of Discount deduction

Dim RelationID As Long

SelectedRow = grd.Row

With grd
    'Get the slected amount
    .Row = SelectedRow
    .Col = .Cols - 1: Amount = Val(.Text)
    
    .Col = 0
    Do
        .Row = .Row + 1
        If Trim$(.Text) = "" Then
            .Col = .Cols - 1: TotalAmount = Val(.Text)
            Exit Do
        End If
    Loop
    
   TotalAmount = TotalAmount - Amount
    'Now rearrange the Grid
   If TotalAmount = 0 Then
      InitGrid
      Exit Function
   End If

   .RemoveItem (SelectedRow)

End With


End Function

Private Function GetInvoiceNumber() As String
'Trap en error
On Error GoTo ErrLine
'Declare the variables
Dim rst As ADODB.Recordset
Dim InvoiceNumber As String
Dim tempStr As String
'Initialise the function
GetInvoiceNumber = ""

'Get the Max(Invoice Number the Query)
gDbTrans.SqlStmt = " SELECT MAX(InvoiceNo) FROM Sales"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

InvoiceNumber = FormatField(rst(0))

If Len(InvoiceNumber) > 2 Then
    tempStr = Mid$(InvoiceNumber, 3, Len(InvoiceNumber))
    GetInvoiceNumber = Left$(InvoiceNumber, 2) & (Val(tempStr) + 1)
Else
    GetInvoiceNumber = Val(InvoiceNumber) + 1
End If


Exit Function
ErrLine:
    
End Function

Private Sub LoadProductDetails()
Dim lngRelationID As Long
Dim btGodownID As Byte

Dim curTradingprice As Currency
Dim curSalesPrice As Currency

Dim MaterialClass As clsMaterial

Dim rst As ADODB.Recordset

If cmbProductName.ListIndex = -1 Then Exit Sub

lngRelationID = cmbProductName.ItemData(cmbProductName.ListIndex)
btGodownID = 1

If cmbBranch.Enabled Then If cmbBranch.ListIndex = -1 Then Exit Sub
btGodownID = cmbBranch.ItemData(cmbBranch.ListIndex)


Set MaterialClass = New clsMaterial

If Not TextBoxDateValidate(txtInvoiceDate, "/", True, True) Then Exit Sub

m_StockAvailable = MaterialClass.GetItemOnDateClosingStock(lngRelationID, txtInvoiceDate.Text, btGodownID)

If m_StockAvailable > 0 Then
    txtQuantity.Text = m_StockAvailable
Else
    'MsgBox "There is No Stock. Please purchse the product.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(614), vbInformation, wis_MESSAGE_TITLE
    cmbProductName.ListIndex = -1
    txtPrice.Text = ""
    Exit Sub
End If

'txtPrice.Text = FormatCurrency(MaterialClass.GetSalesPrice(lngRelationID, btGodownID))

'If Val(txtPrice.Text) <> 0 Then Exit Sub


gDbTrans.SqlStmt = " SELECT SalesPrice,Tax FROM RelationMaster " & _
                   " WHERE RelationID = " & lngRelationID
  
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

txtPrice.Text = FormatField(rst.Fields("SalesPrice"))
lblTaxAmount.Caption = LoadResString(173) & " " & FormatField(rst.Fields("Tax")) & "%"
txtTaxAmount.Tag = FormatField(rst.Fields("Tax"))
lblPerUnit.Caption = LoadResString(439) & " " & cmbUnit.Text

txtPrice.Text = FormatCurrency(MaterialClass.GetSalesPrice(lngRelationID, btGodownID))

Set rst = Nothing

End Sub

Private Sub LoadProducts()
'Declare the variables
Dim lngGroupID As Long

Dim rstRelation As ADODB.Recordset
Dim ProductName As ADODB.Field

Dim RelationID As ADODB.Field
Dim GodownID As Byte
Dim UnitID As Integer

cmbProductName.Clear

With cmbBranch
    If .ListIndex = -1 Then Exit Sub
    GodownID = .ItemData(.ListIndex)
End With
With cmbGroup
    If .ListIndex = -1 Then Exit Sub
    lngGroupID = .ItemData(.ListIndex)
End With
With cmbUnit
    If .ListIndex = -1 Then Exit Sub
    UnitID = .ItemData(.ListIndex)
End With

  
gDbTrans.SqlStmt = "SELECT A.ProductID,B.ProductName,RelationID " & _
                " FROM RelationMaster A, Products B" & _
                " WHERE A.GroupID = " & lngGroupID & _
                " And GodownId = " & GodownID & _
                " AND A.ProductID = B.ProductID " & _
                " AND A.GroupID = B.GroupID " & _
                " AND A.UnitID= " & UnitID

If gDbTrans.Fetch(rstRelation, adOpenForwardOnly) < 1 Then Exit Sub

Set ProductName = rstRelation.Fields("ProductName")
Set RelationID = rstRelation.Fields("RelationID")

'Load the data to the combo box
With cmbProductName
    Do While Not rstRelation.EOF
       .AddItem ProductName.Value
       .ItemData(.newIndex) = RelationID.Value
       
       'Move the recordset
       rstRelation.MoveNext
    Loop
    If .ListCount = 1 Then .ListIndex = 0
End With

End Sub


Public Property Get VendorID() As Long
VendorID = m_VendorID
End Property

Public Property Let VendorID(ByVal vNewValue As Long)
m_VendorID = vNewValue
End Property

'private grdFunctions as clsgr
Private Sub GridResize()
Dim Ratio  As Single

With grd
   Ratio = .Width / .Cols
   
   .ColWidth(0) = Ratio * 0.5
   .ColWidth(1) = Ratio * 0.95
   .ColWidth(2) = Ratio * 1.45
   .ColWidth(3) = Ratio * 0.95
   .ColWidth(4) = Ratio * 0.9
   .ColWidth(5) = Ratio * 0.9
   .ColWidth(6) = Ratio * 0.95
   .ColWidth(7) = Ratio * 0.95
End With

End Sub


Public Sub InitGrid()
With grd
    .Clear
    .Rows = 1
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 8
    .FixedRows = 1
    .Row = 0
    
    .Col = 0: .Text = GetResourceString(33): .CellFontBold = True '"SlNo"
    .Col = 1: .Text = GetResourceString(39): .CellFontBold = True '"Description"
    .Col = 2: .Text = GetResourceString(306): .CellFontBold = True '"Quantity"
    .Col = 3: .Text = GetResourceString(164): .CellFontBold = True '"sales Price"
    .Col = 4: .Text = GetResourceString(163): .CellFontBold = True '"MRP"
    .Text = GetResourceString(40) '"Amoount"
    .Col = 5: .Text = "Soot": .CellFontBold = True '"TRP" QtySoot
    '.Col = 6: .Text = GetResourceString(40): .CellFontBold = True '"Amount"
    .Col = 6: .Text = GetResourceString(173): .CellFontBold = True '"Tax"
    .Col = 7: .Text = GetResourceString(304): .CellFontBold = True '"Amount"
    
    
End With

End Sub

Private Sub LoadProductGroups()
Dim rstGroups As ADODB.Recordset

cmbGroup.Clear
cmbGroupEnglish.Clear
gDbTrans.SqlStmt = " SELECT GroupID,GroupName,GroupNameEnglish FROM ProductGroup " & _
                 " ORDER BY GroupID "
        
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

Dim lngGroupID As Long

cmbUnit.Clear
With cmbGroup
    If .ListIndex >= 0 Then lngGroupID = .ItemData(.ListIndex)
End With

gDbTrans.SqlStmt = " SELECT UnitID,UnitName FROM Units"
If lngGroupID Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where UnitID In " & _
            "(Select Distinct UnitId From RelationMaster " & _
                " Where GroupId = " & lngGroupID & ")"

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY UnitID "
        
Call gDbTrans.Fetch(rstUnits, adOpenForwardOnly)

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


Private Sub SetFont()
On Error Resume Next

Dim Ctrl As Control

For Each Ctrl In Me
      Ctrl.FontName = "Arial"
      Ctrl.FONTSIZE = 10
Next Ctrl


End Sub

Private Function Validated() As Boolean
'Declare the variables
Dim Quantity As Double

Validated = False


If Not DateValidate(txtInvoiceDate.Text, "/", True) Then
   'MsgBox "Invalid date specified!", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtInvoiceDate
   Exit Function
End If

If txtInvoiceNumber.Text = "" Then
   'MsgBox "Please specify the invoice number", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(813), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtInvoiceNumber
   Exit Function
End If

If cmbBranch.Enabled Then
    If cmbBranch.ListIndex = -1 Then
       ' MsgBox "Please select the Branch", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetChangeString(GetResourceString(230), GetResourceString(227)), vbInformation, wis_MESSAGE_TITLE
        cmbBranch.SetFocus
        Exit Function
    End If
End If
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

If Not CurrencyValidate(txtPrice.Text, False) Then
   'MsgBox "Invalid currency specified!", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
   'ActivateTextBox txtPrice
   'Exit Function
End If

If Not IsNumeric(txtQuantity.Text) Then
   'MsgBox "Invalid Quantity sepecified!", vbInformation, wis_MESSAGE_TITLE
   MsgBox GetResourceString(776), vbInformation, wis_MESSAGE_TITLE
   ActivateTextBox txtQuantity
   Exit Function
End If

If Not CurrencyValidate(txtAmount.Text, True) Then
    'MsgBox "Invalid currency specified!", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtAmount
    Exit Function
End If

'Check for the quantity
Quantity = txtQuantity
If Quantity > m_StockAvailable Then
    'MsgBox "You can not sell the Quantity more than Available Stock.", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(615), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Validated = True

End Function

Private Sub cmbBranch_Click()
'Set the last transaction date
Call SetLastTransactionDate

End Sub


Private Sub cmbExpenseHead_Click()
If cmbExpenseHead.ListIndex = -1 Then Exit Sub
txtExpense.Text = ""
End Sub


Private Sub cmbGroup_Click()
If cmbGroup.ListIndex = -1 Then Exit Sub
'Load the Products to the products combo box
Call LoadUnits
'LoadProducts

End Sub

Private Sub cmbIncomeHead_Click()
If cmbIncomeHead.ListIndex = -1 Then Exit Sub
txtIncome.Text = ""
End Sub


Private Sub cmbManufacturer_Click()
cmbUnit.ListIndex = -1

End Sub

Private Sub cmbProductName_Click()
    If cmbProductName.ListIndex = -1 Then Exit Sub
    Call LoadProductDetails

End Sub

Private Sub cmbUnit_Click()
Call LoadProducts 'Details
End Sub


Private Sub cmdAccept_Click()
RaiseEvent OkClicked
Me.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
If Not Validated Then Exit Sub
RaiseEvent AddClicked

End Sub

Private Sub cmdAddIncome_Click()
If Not PLHeadsValidated(Profit) Then Exit Sub

RaiseEvent PLHeadsClicked(Profit)
End Sub

Private Sub cmdCancel_Click()
Unload Me
'RaiseEvent WindowClosed
End Sub

Private Sub cmdClear_Click()

RaiseEvent ClearClicked
End Sub

Private Sub cmdDelete_Click()
RaiseEvent DeleteClicked
End Sub

Private Sub cmdExpense_Click()
If Not PLHeadsValidated(Loss) Then Exit Sub

RaiseEvent PLHeadsClicked(Loss)
End Sub

Private Sub cmdInvoicedate_Click()
With Calendar
    .selDate = gStrDate
    If DateValidate(txtInvoiceDate, "/", True) Then .selDate = txtInvoiceDate
    .Left = Left + fra.Left + cmdInvoicedate.Left
    .Top = Top + fra.Top + cmdInvoicedate.Top - .Height / 2
    .Show vbModal
    txtInvoiceDate = .selDate
End With

End Sub

Private Sub cmdPrint_Click()

RaiseEvent PrintClicked

End Sub

Private Sub cmdSalesInvoice_Click()

If Not InvoiceClickValidated Then Exit Sub

RaiseEvent InvoiceClicked
End Sub

Private Sub cmdUndo_Click()
RaiseEvent UnDoClicked
End Sub

Private Sub Form_Initialize()
If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
Set m_GrdFunctions.fGrd = grd

End Sub

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
cmdSalesInvoice.Tag = IIf(Shift And vbCtrlMask, 1, 0)

If KeyCode = vbKeyF6 Then RaiseEvent ChangeCustomer

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
cmdInvoicedate.Tag = 0
End Sub

Private Sub Form_Resize()

GridResize

'tabSales.Tabs(1).Selected = True
cmbBranch.SetFocus

End Sub

Private Sub SetLastTransactionDate()
'Declare the variables
Dim rst As ADODB.Recordset
Dim headID As Long
Dim GodownID  As Integer

If cmbBranch.ListIndex = -1 Then Exit Sub
If m_VendorID = 0 Then Exit Sub

GodownID = cmbBranch.ItemData(cmbBranch.ListIndex)

gDbTrans.SqlStmt = " SELECT TOP 1 InvoiceNo,InvoiceDate " & _
                  " FROM Sales " & _
                  " WHERE HeadID = " & m_VendorID & _
                  " AND GodownID = " & GodownID & _
                  " ORDER BY InvoiceDate DESC "
                
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

txtInvoiceDate.Text = FormatField(rst.Fields("InvoiceDate"))
txtInvoiceNumber.Text = FormatField(rst.Fields("InvoiceNo"))

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPurchase = Nothing
RaiseEvent WindowClosed
End Sub

Private Sub grd_DblClick()
RaiseEvent GridClicked
End Sub

Private Sub lblCompanyName_DblClick()
RaiseEvent ChangeCustomer
End Sub

Private Sub txtAmount_Change()
    txtTaxAmount.Text = FormatCurrency(Val(txtTaxAmount.Tag) / 100 * Val(txtAmount.Text))
End Sub

Private Sub txtInvoiceDate_GotFocus()
ActivateTextBox txtInvoiceDate
End Sub

Private Sub txtInvoiceNumber_GotFocus()
    ActivateTextBox txtInvoiceNumber
End Sub

Private Sub txtInvoiceNumber_LostFocus()

On Error Resume Next
cmbGroup.SetFocus
If cmbBranch.Enabled And cmbBranch.ListIndex < 0 Then cmbBranch.SetFocus

End Sub

Private Sub txtPrice_Change()
txtAmount.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtPrice_GotFocus()
On Error Resume Next
With txtPrice
    .SelStart = 0
    .SelLength = Len(txtPrice)
End With
End Sub

Private Sub txtQuantity_Change()
txtAmount.Text = FormatCurrency(Val(txtQuantity.Text) * Val(txtPrice.Text))
End Sub

Private Sub Form_Load()
'Declare the variables
Dim MaterialClass As clsMaterial

'Center the form
CenterMe Me

'set icon for the form caption
Me.Icon = LoadResPicture(147, vbResIcon)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
  
Call SetKannadaCaption

If MaterialClass Is Nothing Then Set MaterialClass = New clsMaterial

'Load the branches
Call MaterialClass.LoadAllBranches(cmbBranch)

'Load Group Combo box
LoadProductGroups

'Load Unit combo box
LoadUnits

'Initalise the grid.
InitGrid

Dim showAllHeads As Boolean
showAllHeads = False
Dim SetUp As New clsSetup
If UCase(SetUp.ReadSetupValue("Trading", "ShowAllExpenseHeads", "True")) = "TRUE" Then showAllHeads = True
Set SetUp = Nothing

Call LoadLedgersToCombo(cmbIncomeHead, parTradingIncome, True)
'Then Load other Income Heads
If showAllHeads Then Call LoadLedgersToCombo(cmbIncomeHead, parIncome, False)

Call LoadLedgersToCombo(cmbExpenseHead, parTradingExpense, True)
'Then Load other Income Heads
If showAllHeads Then Call LoadLedgersToCombo(cmbExpenseHead, parExpense, False)

'optCredit.Value = True
cmdUndo.Enabled = False
cmdPrint.Enabled = False
cmdAccept.Enabled = False
cmbBranch.Locked = False
cmdDelete.Enabled = False

End Sub

Private Sub txtQuantity_GotFocus()
On Error Resume Next
With txtQuantity
    .SelStart = 0
    .SelLength = Len(txtQuantity)
End With
End Sub

