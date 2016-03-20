VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TITLE_TOP As Integer = 50
Private Const LEFT_MARGIN As Integer = 400
Private Const LINE_MARGIN As Integer = 60
Private Const FIELD_MARGIN  As Integer = 200

Private Const HEADING_HEIGHT As Integer = 700
Private Const ROW_HEIGHT As Integer = 125
Private Const INVOICE_TEXT As String = "INVOICE"
Private Const INVOICE_NO_POS As Integer = 6900
Private Const INVOICE_VALUE_POS As Integer = 8700
Private Const INVOICE_DATE_POS As Integer = 9500
Private Const BOTTOM_LINE_POS As Integer = 11500
Private Const BOTTOM_TEXT_POS As Integer = 11800
Private Const EXPENSE_HEAD_POS As Integer = 4000
Private Const INCOME_HEAD_VAL_POS As Integer = 2400
Private Const EXPENSE_HEAD_VAL_POS As Integer = 7500
Private Const MAX_ITEMS_ON_PAGE As Integer = 30

'Private m_lhPrinter As Long
Private my_Printer As Printer
Private m_DataSource As MSFlexGrid
Private m_CollectionClass As clsMaterials

Private m_rstPrintProperty As ADODB.Recordset

Private m_PrintQuality As Integer
Private m_PrinterHeight As Long
Private m_PrinterWidth As Long
Private m_ScaleMode As Integer
Private m_PaperSize As Integer
Private m_PrintDate As String
Private m_InvoiceNo As String
Private m_KSTPercent As Single

Private m_CustomerHeadID As Long
Private m_CompanyName As String

'Declare the Inventory Items Printing
Private m_Total As Currency

Private m_TotalQuantity As Double
Private m_TotalFreeQuantity As Double

Private m_Qty_XPos As Single
Private m_Total_XPos As Single
Private m_FreeQty_XPos As Single
Private m_Pos_Calculate As Single
Private m_ExpenseY As Double
Private m_IncomeY As Double

Private m_NumOfPages As Integer
Private IsNextPageRequired As Boolean

Public Property Get ObjectName()
    ObjectName = "PrintObjectClass"
End Property


Private Function GetCustomerRecordSet(ByRef rstCustomer As ADODB.Recordset) As Boolean


gDbTrans.SqlStmt = " SELECT CompanyName,Address,DLNo,KST,CST" & _
                   " FROM CompanyCreation" & _
                   " WHERE HeadID=" & m_CustomerHeadID

Call gDbTrans.Fetch(rstCustomer, adOpenForwardOnly)


End Function

Private Function GetCompanyTitleRecordSet(ByRef rstCompanyTitle As ADODB.Recordset) As Boolean

Dim CompanyType As wis_CompanyType

CompanyType = Enum_Self


gDbTrans.SqlStmt = " SELECT CompanyName,Address,PhoneNo,DLNo,KST,CST" & _
                   " FROM CompanyCreation" & _
                   " WHERE CompanyType=" & CompanyType

Call gDbTrans.Fetch(rstCompanyTitle, adOpenForwardOnly)


End Function
Private Sub PrintAccountData(MaterialClass As clsMaterial, ByVal ParentID As Long, ByVal ColWid As Single)
'Declare the variables
Dim Amount As Currency

'Setup an error handler...
On Error GoTo ErrLine

With my_Printer
    Select Case ParentID
        Case wis_IncomeParentID, parTradingIncome, 80200, 80300, 80400, 80500, 80600, 80700, 80800, 80900
            Amount = MaterialClass.Amount
            m_Total = m_Total + Amount
            .CurrentX = LEFT_MARGIN
            .CurrentY = m_IncomeY
            
            .FontBold = False
            my_Printer.Print MaterialClass.ProductName
            
            .CurrentY = m_IncomeY
            .CurrentX = INCOME_HEAD_VAL_POS
            my_Printer.Print AlignData(FormatCurrency(Amount), ColWid)
            m_IncomeY = .CurrentY
        Case wis_ExpenseParentID, parTradingExpense, 90200, 90300, 90400, 90500, 90600, 90700, 90800, 90900
            Amount = MaterialClass.Amount
            m_Total = m_Total - Amount
            .CurrentX = EXPENSE_HEAD_POS
            .CurrentY = m_ExpenseY
            .FontBold = False
            my_Printer.Print MaterialClass.ProductName
            
            .CurrentY = m_ExpenseY
            .CurrentX = EXPENSE_HEAD_VAL_POS
            my_Printer.Print AlignData(FormatCurrency(Amount), ColWid)
            m_ExpenseY = .CurrentY
    End Select
End With

Exit Sub

ErrLine:
    MsgBox "PrintAccountData: " & Err.Description, vbCritical
    
End Sub

Private Sub PrintColumnHeader()
Dim curX As Single
Dim CurY As Single
Dim Ratio As Single
Dim ColWid As Single

With my_Printer
    .FONTSIZE = gFontSize + 2
    .FontBold = True
    .CurrentX = LEFT_MARGIN
    CurY = .CurrentY
    Ratio = .ScaleWidth / 11 '( 11  Total Columns)
End With
curX = LEFT_MARGIN

my_Printer.Print "Mfr"

With my_Printer
    .CurrentX = curX + Ratio * 0.45 '871 twips
    curX = .CurrentX
    .CurrentY = CurY
End With

my_Printer.Print "Particulars"
    
With my_Printer
    .CurrentX = curX + Ratio * 2.6 '3594 twips
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print "BatchNo "

With my_Printer
    .CurrentX = curX + Ratio * 0.8 '4431 twips
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print "ExpDate "

With my_Printer
    .CurrentX = curX + Ratio * 0.9
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print "Qty "

With my_Printer
    .CurrentX = curX + Ratio * 0.55
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print "Bonus "

With my_Printer
    .CurrentX = curX + Ratio * 0.45
    ColWid = Ratio - 50
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print AlignData("S.Price", ColWid)

With my_Printer
    .CurrentX = curX + Ratio * 1
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print AlignData("TRP", ColWid)

With my_Printer
    .CurrentX = curX + Ratio * 1
    curX = .CurrentX
    .CurrentY = CurY
End With

my_Printer.Print AlignData("MRP", ColWid)

With my_Printer
    .CurrentX = curX + Ratio * 1.2
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print "KST"


With my_Printer
    .CurrentX = curX + Ratio * 0.6
    curX = .CurrentX
    .CurrentY = CurY
End With
my_Printer.Print AlignData("Amount", ColWid)

With my_Printer
    .CurrentY = .CurrentY + LINE_MARGIN
    Call PrintLine(LEFT_MARGIN, .CurrentY, m_PrinterWidth, .CurrentY)
    .CurrentY = .CurrentY + LINE_MARGIN
End With

End Sub

Private Sub PrintCompanyTitle()
Dim rst As ADODB.Recordset
Dim CompanyAddress As String
Dim PhoneNum As String
Dim DLNo As String
Dim Kst As String
Dim Cst As String

'Get the Title Address from the database
Call GetCompanyTitleRecordSet(rst)
If rst.RecordCount < 0 Then
    MsgBox "Print Errors: Contact Vendor:" & vbCritical
    Exit Sub
End If

m_CompanyName = rst.Fields("CompanyName")
CompanyAddress = FormatField(rst.Fields("Address"))
PhoneNum = FormatField(rst.Fields("PhoneNo"))
DLNo = FormatField(rst.Fields("DLNo"))
Kst = FormatField(rst.Fields("KST"))
Cst = FormatField(rst.Fields("CST"))

With my_Printer
    .FONTSIZE = 14
    .FontBold = True
    .CurrentX = (m_PrinterWidth - .TextWidth(m_CompanyName)) / 2
    .CurrentY = TITLE_TOP
End With

my_Printer.Print m_CompanyName

Call SetFont

my_Printer.CurrentX = (m_PrinterWidth - my_Printer.TextWidth(CompanyAddress)) / 2
my_Printer.Print CompanyAddress

my_Printer.CurrentX = (m_PrinterWidth - my_Printer.TextWidth("Phone No : " & PhoneNum)) / 2
my_Printer.Print "Phone No : " & PhoneNum

my_Printer.CurrentX = (m_PrinterWidth - my_Printer.TextWidth("Drug License No : " & DLNo)) / 2
my_Printer.Print "Drug License No : " & DLNo

my_Printer.CurrentX = (m_PrinterWidth - my_Printer.TextWidth("KST No : " & Kst & ", CST No : " & Cst)) / 2
my_Printer.Print "KST No : " & Kst & ", CST No : " & Cst

With my_Printer
    .CurrentY = .CurrentY + LINE_MARGIN
    Call PrintLine(.ScaleLeft + LEFT_MARGIN, .CurrentY, m_PrinterWidth, .CurrentY)
End With


End Sub

Private Sub PrintDeclaration()
Dim rst As ADODB.Recordset
Dim ret As Long
'Get the Declaration From the Database
gDbTrans.SqlStmt = " SELECT ValueData" & _
                   " FROM Install" & _
                   " WHERE KeyData LIKE 'Declaration%'"

ret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If ret <= 0 Then Exit Sub

my_Printer.CurrentY = my_Printer.CurrentY + LINE_MARGIN

Do While Not rst.EOF
    With my_Printer
        .FONTSIZE = gFontSize
        .CurrentX = LEFT_MARGIN
    End With
    
    my_Printer.Print FormatField(rst.Fields(0))
    
    rst.MoveNext
Loop


gDbTrans.SqlStmt = " SELECT ValueData" & _
                   " FROM Install" & _
                   " WHERE KeyData= 'Jurisdiction'"

ret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If ret <= 0 Then Exit Sub

my_Printer.CurrentY = my_Printer.CurrentY + LINE_MARGIN


With my_Printer
    .FONTSIZE = gFontSize
    .CurrentX = LEFT_MARGIN
End With

my_Printer.Print " Subject to " & FormatField(rst.Fields(0)) & " Jurisdiction"
    

End Sub



Private Sub PrintFirstSecondSalesAmount(ByVal FirstSalesAmount As Currency, ByVal SecondSalesAmount As Currency, ByVal ColWid As Single)
Dim CurY As Single

With my_Printer
    CurY = m_IncomeY '+ LINE_MARGIN
    .CurrentY = CurY
    .CurrentX = LEFT_MARGIN
    my_Printer.Print "First Sales "
    
    If Not IsNextPageRequired Then
        .CurrentY = CurY
        .CurrentX = INCOME_HEAD_VAL_POS
        my_Printer.Print AlignData(FormatCurrency(FirstSalesAmount), ColWid)
    End If
    
    CurY = m_ExpenseY
    .CurrentY = CurY
    .CurrentX = EXPENSE_HEAD_POS
    my_Printer.Print "Second Sales"
    
    If Not IsNextPageRequired Then
        .CurrentY = CurY
        .CurrentX = EXPENSE_HEAD_VAL_POS
        my_Printer.Print AlignData(FormatCurrency(SecondSalesAmount), ColWid)
    End If
    m_ExpenseY = .CurrentY
    m_IncomeY = .CurrentY
End With

End Sub

Private Sub PrintFooter()

With my_Printer
    .FONTSIZE = gFontSize + 3
    .FontBold = True
    .CurrentX = (m_PrinterWidth - LEFT_MARGIN) - .TextWidth("For        " & m_CompanyName)
    .CurrentY = .CurrentY + 5 * LINE_MARGIN
End With

my_Printer.Print "For " & m_CompanyName

With my_Printer
    .FONTSIZE = gFontSize + 2
    .FontBold = False
    .CurrentY = .CurrentY + 10 * LINE_MARGIN
    .CurrentX = (m_PrinterWidth - LEFT_MARGIN) - .TextWidth("For   " & m_CompanyName)
    '.CurrentX = (m_PrinterWidth - .TextWidth("Authorised Signatory    ")) / 2 + 5 * LEFT_MARGIN
    
End With

my_Printer.Print "Authorised Signatory"


End Sub

Private Sub PrintInventoryItems(MaterialClass As clsMaterial, ByVal Ratio As Single)
Dim curX As Single
Dim CurY As Single
Dim ColWid As Single

Dim Amount As Currency

Dim Kst As Single
Dim Quantity As Single
Dim FreeQuantity As Single

With my_Printer
    CurY = .CurrentY
    ColWid = Ratio - 50
    
    If IsNextPageRequired Then
        .CurrentX = m_FreeQty_XPos
        .FONTSIZE = .FONTSIZE + 1
        .FontBold = True
        my_Printer.Print "B/F from Previous Page"
        .CurrentX = m_Total_XPos
        .CurrentY = CurY
        my_Printer.Print AlignData(FormatCurrency(m_Total), ColWid)
        CurY = .CurrentY
        .FONTSIZE = .FONTSIZE - 2
        .FontBold = False
    End If
    
    curX = LEFT_MARGIN
    .CurrentX = curX
    .CurrentY = CurY
     'my_Printer.Print MaterialClass.Manufacterer
            
    'Print Particulars
    .CurrentX = curX + Ratio * 0.45
    curX = .CurrentX
    .CurrentY = CurY
    my_Printer.Print MaterialClass.ProductName
            
    'Print Batch no
    .CurrentX = curX + Ratio * 2.6
    curX = .CurrentX
    .CurrentY = CurY
    'my_Printer.Print MaterialClass.BatchNo
            
    'Print Exp Date
    .CurrentX = curX + Ratio * 0.8
    curX = .CurrentX
    .CurrentY = CurY
    'my_Printer.Print FormatDateToMonthYear(MaterialClass.ExpDate, True)
            
    'Print the Quantity
    .CurrentX = curX + Ratio * 0.9
    curX = .CurrentX
    m_Qty_XPos = curX
    .CurrentY = CurY
    Quantity = MaterialClass.Quantity
    m_TotalQuantity = m_TotalQuantity + Quantity
    my_Printer.Print Quantity
            
            
    'Print the Free Quantity
    .CurrentX = curX + Ratio * 0.55
    curX = .CurrentX
    m_FreeQty_XPos = curX
    .CurrentY = CurY
    'FreeQuantity = MaterialClass.FreeQuantity
    m_TotalFreeQuantity = m_TotalFreeQuantity + FreeQuantity
    If FreeQuantity > 0 Then my_Printer.Print FreeQuantity
            
    'Print the SalePrice
    .CurrentX = curX + Ratio * 0.45
    
    curX = .CurrentX
    .CurrentY = CurY
    my_Printer.Print AlignData(FormatCurrency(MaterialClass.UnitPrice), ColWid)
            
            
    'Print the TRP
    .CurrentX = curX + Ratio * 1
    curX = .CurrentX
    .CurrentY = CurY
    my_Printer.Print AlignData(FormatCurrency(MaterialClass.TradingPrice), ColWid)
            
    'Print the MRP
    .CurrentX = curX + Ratio * 1
    curX = .CurrentX
    .CurrentY = CurY
    my_Printer.Print AlignData(FormatCurrency(MaterialClass.MRP), ColWid)
            
    'Print the KST
    .CurrentX = curX + Ratio * 1.2
    curX = .CurrentX
    .CurrentY = CurY
    'Kst = MaterialClass.Kst
    If Kst > 0 Then my_Printer.Print Kst
            
    'Print the Amount
    .CurrentX = curX + Ratio * 0.6
    curX = .CurrentX
    m_Total_XPos = curX
    .CurrentY = CurY
    Amount = MaterialClass.Amount
    my_Printer.Print AlignData(FormatCurrency(Amount), ColWid)
    m_Total = m_Total + Amount
End With

End Sub

Private Sub PrintLine(ByVal x1Pos As Single, ByVal y1Pos As Single, _
            ByVal x2Pos As Single, ByVal y2Pos As Single)

'Setup an error handler
On Error GoTo ErrLine

my_Printer.Line (x1Pos, y1Pos)-(x2Pos, y2Pos)

Exit Sub

ErrLine:
    MsgBox "PrintLine: " & Err.Description
    
End Sub


Private Sub PrintHeading()
Dim TitlePrint As wis_PrintTitle

TitlePrint = FormatField(m_rstPrintProperty.Fields("TitlePrint"))

If TitlePrint = Enum_PrintTitle Then PrintCompanyTitle

With my_Printer
'    .ScaleMode = m_ScaleMode
    .CurrentX = LEFT_MARGIN
    If HEADING_HEIGHT > .CurrentY Then .CurrentY = HEADING_HEIGHT
End With


End Sub


Private Sub PrintSubTotalLine(ByVal Ratio As Single)
Dim CurY As Single
Dim curX As Single

With my_Printer
    CurY = BOTTOM_LINE_POS
    .CurrentX = LEFT_MARGIN
    .CurrentY = CurY
    Call PrintLine(LEFT_MARGIN, CurY, m_PrinterWidth, CurY)
    
    .FontBold = True
    
    If Not IsNextPageRequired Then
        .CurrentX = m_Qty_XPos - .TextWidth("Total Qty     ")
        CurY = .CurrentY + LINE_MARGIN
        .CurrentY = CurY
        my_Printer.Print "Total Qty"
    
    
        .CurrentY = CurY
        .CurrentX = m_Qty_XPos '+ .TextWidth("Tot")
        curX = .CurrentX
        my_Printer.Print (m_TotalQuantity + m_TotalFreeQuantity)
    
        .CurrentX = curX + 2000
        .CurrentY = CurY
        my_Printer.Print "Sub Total"
    
        .CurrentY = CurY
        .CurrentX = m_Total_XPos
        my_Printer.Print AlignData(FormatCurrency(m_Total), (Ratio - 50))
        CurY = .CurrentY
        Call PrintLine(LEFT_MARGIN, CurY, m_PrinterWidth, CurY)
    
        CurY = .CurrentY + LINE_MARGIN
    End If
    m_IncomeY = CurY
    m_ExpenseY = CurY
End With

End Sub

Private Sub PrintTotalAmount(ByVal ColWid As Single)
'Dim CurFigureClass As clsCurFigure
Dim CurFigures As String
Dim CurY As Single
With my_Printer
    my_Printer.FontBold = True
    .FONTSIZE = .FONTSIZE + 4
    .CurrentX = m_Total_XPos - .TextWidth("Total Total To") - LEFT_MARGIN
    CurY = .CurrentY
    If IsNextPageRequired Then
        .FONTSIZE = .FONTSIZE - 2
        my_Printer.Print "Continued... " & m_NumOfPages + 1
    Else
        my_Printer.Print "Total Amount"
    End If
    
    
    .CurrentX = m_Total_XPos
    '.CurrentY = m_ExpenseY
    If Not IsNextPageRequired Then
        .CurrentY = CurY
        
'        my_Printer.Print AlignData(CurrencyRoundOff(m_Total), ColWid)
        
        my_Printer.FontBold = False
        'Print the text in Figures
'        Set CurFigureClass = New clsCurFigure
        
        .CurrentX = LEFT_MARGIN
        .FONTSIZE = .FONTSIZE - 3
'        CurFigures = CurFigureClass.NumberInFigure(m_Total)
        
        my_Printer.Print "Rupees " & CurFigures & " Only"
        
'        Set CurFigureClass = Nothing
    End If
End With

End Sub

Private Function SetPrintPropertyRecordset(ByRef rstProperty As ADODB.Recordset, ByVal SaleTransID As Long) As Long

gDbTrans.SqlStmt = " SELECT *,TransModeName" & _
                   " FROM InvoicePrint A,TransportMode B" & _
                   " WHERE SaleTransID=" & SaleTransID & _
                   " AND A.TransportMode=B.TransModeID"

SetPrintPropertyRecordset = gDbTrans.Fetch(rstProperty, adOpenStatic)


End Function

Private Sub PrintDetails(ByVal FirstSalesAmount As Currency, ByVal SecondSalesAmount As Currency, ByRef ItemCount As Integer)
'Declare the variables
Dim MaxCount As Integer

Dim CurY As Single
Dim curX As Single
Dim Ratio  As Single
Dim ColWid As Single

Dim ParentID As Long
Dim LinePrinted As Boolean

Dim Particulars As String
Dim CurFigures As String

Dim MaterialClass As clsMaterial
'Dim CurFigureClass As clsCurFigure

Dim lKeys

With my_Printer
    .FONTSIZE = gFontSize + 2
    .FontBold = False
    Ratio = .ScaleWidth / 11 '( 11  Total Columns)
    ColWid = Ratio - 50
End With

MaxCount = m_CollectionClass.count

lKeys = m_CollectionClass.GetKeys
If IsNextPageRequired Then
    ItemCount = ItemCount + 1
    CurY = my_Printer.CurrentY
End If


Do
    If ItemCount >= MaxCount Then Exit Do

    Set MaterialClass = m_CollectionClass.Item(lKeys(ItemCount))
    
    ParentID = MaterialClass.ParentID
    
    With my_Printer
        If ItemCount = 0 Then CurY = .CurrentY + LINE_MARGIN
        If ItemCount > 0 Then .CurrentY = CurY
        
        If ParentID = 0 Then Call PrintInventoryItems(MaterialClass, Ratio)
        
        If ParentID > 0 Then
            If Not LinePrinted Then
                Call PrintSubTotalLine(Ratio)
                LinePrinted = True
            End If
            Call PrintAccountData(MaterialClass, ParentID, ColWid)
        End If
        
        CurY = .CurrentY
        
    End With
    'my_Printer.EndDoc
    IsNextPageRequired = False
    
    If Not LinePrinted Then
        If CurY >= 11000 Then
            m_NumOfPages = m_NumOfPages + 1
            IsNextPageRequired = True
            Exit Do
        End If
    End If
    ItemCount = ItemCount + 1
    
Loop

'If ItemCount = MAX_ITEMS_ON_PAGE Then m_NumofPages = m_NumofPages + 1

'If Parentid does not contain any wis_ExpenseID, wis_TradingExpenseID
'wis_IncomeID, wis_TradingIncomeID
If Not LinePrinted Then
    Call PrintSubTotalLine(Ratio)
    LinePrinted = True
End If

'Print First Sales and Second Sales Amount
Call PrintFirstSecondSalesAmount(FirstSalesAmount, SecondSalesAmount, ColWid)

MaxCount = m_CollectionClass.count - 1

If ItemCount >= MaxCount Then
    IsNextPageRequired = False
    'If my_Printer.CurrentY >= m_IncomeY Then m_IncomeY = my_Printer.CurrentY
    'If my_Printer.CurrentY >= m_ExpenseY Then m_ExpenseY = my_Printer.CurrentY
End If


Call PrintTotalAmount(ColWid)

With my_Printer
    .CurrentX = LEFT_MARGIN
    .CurrentY = .CurrentY + LINE_MARGIN
    Call PrintLine(LEFT_MARGIN, .CurrentY, m_PrinterWidth, .CurrentY)
    .CurrentY = .CurrentY + LINE_MARGIN
End With

End Sub

Private Function AlignData(ByVal strSource As String, ByVal Wid As Single) As String

With my_Printer
    If .TextWidth(strSource) < Wid Then
        Do While .TextWidth(strSource) <= Wid
            'strSource = " " & strSource
            strSource = Space(1) & strSource
        Loop
    End If
End With

AlignData = strSource

End Function


Public Function PrintInvoice(ByVal SaleTransID As Long, ByVal FirstSalesAmount As Currency, _
           ByVal SecondSalesAmount As Currency, ByVal ItemCount As Integer) As Boolean


'Setup error handler...
On Error GoTo ErrLine

' If no recordset available, exit.
If m_DataSource Is Nothing Then Err.Raise vbObjectError + 513, , "No Data"

If m_CollectionClass Is Nothing Then Err.Raise vbObjectError + 513, , "No Data"

If SetPrintPropertyRecordset(m_rstPrintProperty, SaleTransID) < 1 Then
    MsgBox "Please Set the print Properties", vbInformation
    Exit Function
End If

'Set the Fonts of the grid
SetFont

'Prnt the Company Title Address etc
PrintHeading

'Print the Properties Such as Invoice Number, Consignee etc
PrintProperties

'PRINT the column headers
PrintColumnHeader

'Print the Details Such as Quantity, price etc.
Call PrintDetails(FirstSalesAmount, SecondSalesAmount, ItemCount)

'Print the Declaration
PrintDeclaration

'Print Footer
If IsNextPageRequired Then
    my_Printer.newPage
    my_Printer.EndDoc
    
    Call PrintInvoice(SaleTransID, FirstSalesAmount, SecondSalesAmount, ItemCount)
Else
    PrintFooter
    my_Printer.newPage
    my_Printer.EndDoc
End If

PrintInvoice = True

Exit Function

ErrLine:
    MsgBox "PrintInvoice: " & vbCrLf & Err.Description, vbCritical
    
End Function

Private Sub PrintProperties()
'Declare the variables
Dim curX As Single
Dim CurY As Single
Dim New_YPos As Single
Dim ColWid As Single
Dim InvValue_Pos As Single
Dim VertLine_YPos As Single
Dim PayTerm As wis_PaymentTerm
Dim PayTermText As String


Dim rstCustomer As ADODB.Recordset

With my_Printer
    .FONTSIZE = 14
    .FontBold = True
    .CurrentX = (m_PrinterWidth - .TextWidth(INVOICE_TEXT)) / 2
    CurY = .CurrentY
    CurY = CurY + .TextHeight(INVOICE_TEXT)
    .CurrentY = CurY + LINE_MARGIN
End With

my_Printer.Print INVOICE_TEXT

With my_Printer
    CurY = .CurrentY
    CurY = CurY + .TextHeight(INVOICE_TEXT)
    .CurrentY = CurY + LINE_MARGIN
    Call PrintLine(LEFT_MARGIN, .CurrentY, m_PrinterWidth, .CurrentY)
    VertLine_YPos = .CurrentY
    CurY = .CurrentY + ROW_HEIGHT
End With

'Get  ConsigneeName Address CST
'CopanyName,Address,DLNo,KST,CST
Call GetCustomerRecordSet(rstCustomer)

If rstCustomer.RecordCount < 0 Then
    MsgBox "Error in Getting Customer Details", vbCritical
    Exit Sub
End If

With my_Printer
    .CurrentX = LEFT_MARGIN
    .CurrentY = CurY
    .FONTSIZE = m_DataSource.CellFontSize + 2
End With

my_Printer.Print "Consignee:"

With my_Printer
    .CurrentX = LEFT_MARGIN
End With

my_Printer.Print rstCustomer.Fields("CompanyName")
With my_Printer
    .FONTSIZE = m_DataSource.CellFontSize + 3
    .CurrentX = LEFT_MARGIN
End With

'Set the fonts
SetFont

my_Printer.Print FormatField(rstCustomer.Fields("Address"))
my_Printer.CurrentX = LEFT_MARGIN
my_Printer.Print FormatField(rstCustomer.Fields("DLNo"))
my_Printer.CurrentX = LEFT_MARGIN
my_Printer.Print FormatField(rstCustomer.Fields("KST"))
my_Printer.CurrentX = LEFT_MARGIN
my_Printer.Print FormatField(rstCustomer.Fields("CST"))

'Print Invoice No date etc

With my_Printer
    ColWid = .TextWidth(rstCustomer.Fields("CompanyName"))
    InvValue_Pos = INVOICE_VALUE_POS
    .CurrentX = INVOICE_NO_POS
    .CurrentY = CurY
    New_YPos = CurY
    .FontBold = True
End With

my_Printer.Print "Invoice No "
With my_Printer
    .CurrentX = InvValue_Pos
    .CurrentY = CurY
    .FontBold = True
End With
my_Printer.Print m_InvoiceNo

With my_Printer
    CurY = .CurrentY
    .CurrentX = INVOICE_NO_POS
    .FontBold = False
End With
my_Printer.Print "Quotation No "

With my_Printer
    .CurrentX = InvValue_Pos
    .CurrentY = CurY
End With
my_Printer.Print FormatField(m_rstPrintProperty.Fields("QuotationNo"))

With my_Printer
    CurY = .CurrentY
    .CurrentX = INVOICE_NO_POS
End With
my_Printer.Print "Payment Term"

With my_Printer
    .CurrentX = InvValue_Pos
    .CurrentY = CurY
End With

PayTerm = FormatField(m_rstPrintProperty.Fields("PaymentTerm"))

Select Case PayTerm
    Case Enum_Cash
        PayTermText = "Cash"
    Case Enum_Cheque
        PayTermText = "Cheque"
    Case Enum_Credit
        PayTermText = "Credit"
    Case Enum_DD
        PayTermText = "DD"
End Select

my_Printer.Print PayTermText

With my_Printer
    CurY = .CurrentY
    .CurrentX = INVOICE_NO_POS
End With
my_Printer.Print "Transport Mode "

With my_Printer
    .CurrentX = InvValue_Pos
    .CurrentY = CurY
End With

my_Printer.Print FormatField(m_rstPrintProperty.Fields("TransModeName"))

With my_Printer
    CurY = .CurrentY
    .CurrentX = INVOICE_NO_POS
End With
my_Printer.Print "DC No"

With my_Printer
    .CurrentX = InvValue_Pos
    .CurrentY = CurY
End With
my_Printer.Print FormatField(m_rstPrintProperty.Fields("DCNo"))

'Now Print the Invoice Date infront of invoice No
With my_Printer
    CurY = .CurrentY
    .CurrentX = INVOICE_DATE_POS
    .CurrentY = New_YPos
End With
my_Printer.Print " Inv Date " & m_PrintDate

With my_Printer
    .CurrentX = INVOICE_DATE_POS
End With
my_Printer.Print " Qtn Date " & FormatField(m_rstPrintProperty.Fields("QuotationDate"))

With my_Printer
    .CurrentX = INVOICE_DATE_POS
End With
my_Printer.Print " L.R.No " & FormatField(m_rstPrintProperty.Fields("LRNo"))

With my_Printer
    .CurrentY = CurY + LINE_MARGIN
    CurY = .CurrentY
    Call PrintLine(LEFT_MARGIN, CurY, m_PrinterWidth, CurY)
    .CurrentY = .CurrentY + LINE_MARGIN
End With

End Sub

Public Property Set DataSource(ByVal grd As MSFlexGrid)
Set m_DataSource = grd

End Property
Public Property Set CollectionClass(ByVal clsMaterialCol As clsMaterials)
Set m_CollectionClass = clsMaterialCol

End Property
Private Sub SetFont()
On Error GoTo ErrLine

With my_Printer
    .FontName = m_DataSource.CellFontName
    .FONTSIZE = m_DataSource.CellFontSize + 2
    .FontBold = m_DataSource.CellFontBold
    .FontItalic = m_DataSource.CellFontItalic
    .FontUnderline = m_DataSource.CellFontUnderline
    .FontStrikethru = m_DataSource.CellFontStrikeThrough
End With

Exit Sub

ErrLine:
   MsgBox "SetFont: " & vbCrLf _
            & Err.Description, vbCritical
    

End Sub
Public Property Get PrintQuality() As Integer
PrintQuality = m_PrintQuality
End Property

Public Property Let PrintQuality(ByVal vNewValue As Integer)
m_PrintQuality = vNewValue
End Property

Public Property Get ScaleMode() As Integer
ScaleMode = m_ScaleMode
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
m_ScaleMode = vNewValue
End Property
Public Property Get PaperSize() As Integer
PaperSize = m_PaperSize
End Property

Public Property Let PaperSize(ByVal vNewValue As Integer)
m_PaperSize = vNewValue
End Property

Public Property Get PrinterHeight() As Long
PrinterHeight = m_PrinterHeight
End Property

Public Property Let PrinterHeight(ByVal vNewValue As Long)
m_PrinterHeight = vNewValue
End Property

Public Property Get PrinterWidth() As Long
PrinterWidth = m_PrinterWidth
End Property

Public Property Let PrinterWidth(ByVal vNewValue As Long)
m_PrinterWidth = vNewValue
End Property

Private Sub Class_Initialize()

Set my_Printer = Printer

End Sub

Private Sub Class_Terminate()
Set my_Printer = Nothing
Set m_CollectionClass = Nothing
Set m_DataSource = Nothing
Set m_rstPrintProperty = Nothing

End Sub

Public Property Get PrintDate() As String
PrintDate = m_PrintDate
End Property

Public Property Let PrintDate(ByVal vNewValue As String)
m_PrintDate = vNewValue
End Property

Public Property Get InvoiceNo() As String
InvoiceNo = m_InvoiceNo
End Property

Public Property Let InvoiceNo(ByVal vNewValue As String)
m_InvoiceNo = vNewValue
End Property

Public Property Get CustomerHeadID() As Long
CustomerHeadID = m_CustomerHeadID
End Property

Public Property Let CustomerHeadID(ByVal vNewValue As Long)
m_CustomerHeadID = vNewValue
End Property

Public Property Get KSTPercent() As Single
KSTPercent = m_KSTPercent
End Property

Public Property Let KSTPercent(ByVal vNewValue As Single)
m_KSTPercent = vNewValue
End Property
