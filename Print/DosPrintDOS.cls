VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDosPrintDos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare the consts
Private Const TITLE_TOP As Integer = 50
Private Const LEFT_MARGIN As Integer = 400
Private Const LINE_MARGIN As Integer = 60
Private Const FIELD_MARGIN  As Integer = 200

Private Const PAPER_WIDTH = 88 '88 CHARACTERS
Private Const MAX_ITEMS_ON_PAGE As Integer = 30



'Declare the objects to Wrtie to the file
Private m_FileSystemObject As FileSystemObject
Private m_TextStream As TextStream
Private my_Printer As Printer
Private m_CollectionClass As clsMaterials
Private m_DataSource As MSFlexGrid
Private m_rstPrintProperty As ADODB.Recordset

'Declare the Printer variables
Private m_PrinterHeight As Long
Private m_PrinterWidth As Long
Private m_PrintDate As String
Private m_InvoiceNo As String
Private m_CustomerHeadID As Long
Private m_KSTPercent As Single
Private m_CompanyName As String
Private m_InvoiceName As String

Private m_NumOfPages As Integer
Private IsNextPageRequired As Boolean

Private m_Total As Currency
Private m_TotalQuantity As Double
Private m_TotalFreeQuantity As Double
Private m_IncomeHeadPrinted As Boolean
Private m_ExpenseHeadPrinted As Boolean


'''
' These variables are required for Column alignment
Private MfrStr As String * 4
Private ParticularStr As String * 28
Private BatchNoStr As String * 6
Private ExpDateStr As String * 7
Private QtyStr As String * 5
Private BonusStr As String * 4
Private SPriceStr As String * 7
Private TRPStr As String * 7
Private MRPStr As String * 7
Private KSTStr As String * 3
Private AmountStr As String * 10
Private Sub ClearLocalVariables()

End Sub

Public Property Set DataSource(ByVal grd As MSFlexGrid)
Set m_DataSource = grd

End Property


Private Sub OpenTextFile()

Set m_FileSystemObject = New FileSystemObject

Set m_TextStream = m_FileSystemObject.CreateTextFile(App.Path & "\InvPrint.txt", True, False)

End Sub

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

'Open the textfile
OpenTextFile

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
    m_TextStream.Close
    
    Shell App.Path & "\PrintBatch.bat", vbNormalFocus
    
    Call PrintInvoice(SaleTransID, FirstSalesAmount, SecondSalesAmount, ItemCount)
Else
    PrintFooter
    Shell App.Path & "\PrintBatch.bat", vbNormalFocus
    'my_Printer.NewPage
    'my_Printer.EndDoc
End If

PrintInvoice = True

Exit Function

ErrLine:
    MsgBox "PrintInvoice: " & vbCrLf & Err.Description, vbCritical
    
End Function
Private Sub PrintFooter()

With m_TextStream
    
    .WriteLine GetRightAlignedText("For " & m_CompanyName, 70 + Len("For " & m_CompanyName))
    .WriteBlankLines (2)
    .WriteLine GetRightAlignedText("Authorised Signatory", 70 + Len("Authorised Signatory"))
    
End With

End Sub


Private Sub PrintDeclaration()
Dim Rst As ADODB.Recordset
Dim Ret As Long
'Get the Declaration From the Database
gDbTrans.SqlStmt = " SELECT ValueData" & _
                   " FROM Install" & _
                   " WHERE KeyData LIKE 'Declaration%'"

Ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)

If Ret <= 0 Then Exit Sub


With m_TextStream
    Do While Not Rst.EOF
        .WriteLine FormatField(Rst.Fields(0))
        
        Rst.MoveNext
    Loop
End With


gDbTrans.SqlStmt = " SELECT ValueData" & _
                   " FROM Install" & _
                   " WHERE KeyData= 'Jurisdiction'"

Ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)

If Ret <= 0 Then Exit Sub


m_TextStream.WriteLine "Subject to " & FormatField(Rst.Fields(0)) & " Jurisdiction"
    

End Sub




Private Sub PrintDetails(ByVal FirstSalesAmount As Currency, ByVal SecondSalesAmount As Currency, ByRef ItemCount As Integer)
'Declare the variables
Dim MaxCount As Integer

Dim ParentID As Long
Dim LinePrinted As Boolean

Dim Particulars As String
Dim CurFigures As String
Dim LineNum As Integer

Dim MaterialClass As clsMaterial
Dim CurFigureClass As clsCurFigure

Dim lKeys

MaxCount = m_CollectionClass.Count

lKeys = m_CollectionClass.GetKeys
If IsNextPageRequired Then ItemCount = ItemCount + 1
    
Do
    If ItemCount >= MaxCount Then Exit Do

    Set MaterialClass = m_CollectionClass.Item(lKeys(ItemCount))
    
    ParentID = MaterialClass.ParentID
    
    If ParentID = 0 Then Call PrintInventoryItems(MaterialClass)
    
    If ParentID > 0 Then
        If Not LinePrinted Then
            Call PrintSubTotalLine
            LinePrinted = True
        End If
        
        Call PrintAccountData(MaterialClass, ParentID)
        
    End If
    
    'my_Printer.EndDoc
    IsNextPageRequired = False
    
    If Not LinePrinted Then
        
        If MaxCount > MAX_ITEMS_ON_PAGE Then
            m_NumOfPages = m_NumOfPages + 1
            IsNextPageRequired = True
            Exit Do
        End If
    End If
    ItemCount = ItemCount + 1
    
Loop


'If Parentid does not contain any wis_ExpenseID, wis_TradingExpenseID
'wis_IncomeID, wis_TradingIncomeID
If Not LinePrinted Then
    Call PrintSubTotalLine
    LinePrinted = True
End If

'Print First Sales and Second Sales Amount
Call PrintFirstSecondSalesAmount(FirstSalesAmount, SecondSalesAmount)

MaxCount = m_CollectionClass.Count - 1

If ItemCount >= MaxCount Then IsNextPageRequired = False
    
PrintTotalAmount


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



Private Sub PrintTotalAmount()
'Declare the variables
Dim CurFigureClass As clsCurFigure
Dim CurFigures As String

With m_TextStream
    MfrStr = "":  ParticularStr = "": BatchNoStr = "": ExpDateStr = ""
    QtyStr = "": BonusStr = "": SPriceStr = ""
    AmountStr = FormatCurrency(m_Total)
    
    If Not IsNextPageRequired Then
        .WriteLine Chr(14) & MfrStr & ParticularStr & BatchNoStr & ExpDateStr & QtyStr & BonusStr & "         Total Amount " & GetRightAlignedText(AmountStr, Len(AmountStr))
        Set CurFigureClass = New clsCurFigure
        
        CurFigures = CurFigureClass.NumberInFigure(m_Total)
        .WriteLine "Rupees " & CurFigures & " Only"

        Set CurFigureClass = Nothing
    Else
        .WriteLine Chr(14) & MfrStr & ParticularStr & BatchNoStr & ExpDateStr & QtyStr & BonusStr & "Continued... " & m_NumOfPages + 1
    End If
        
    .WriteLine GetHorizontalLineText
End With

End Sub

Private Sub PrintCompanyTitle()
Dim Rst As ADODB.Recordset
Dim CompanyAddress As String
Dim PhoneNum As String
Dim DLNo As String
Dim Kst As String
Dim Cst As String

Dim strPrint As String

'Get the Title Address from the database
Call GetCompanyTitleRecordSet(Rst)

If Rst.RecordCount < 0 Then
    MsgBox "Print Errors: Contact Vendor:" & vbCritical
    Exit Sub
End If

m_CompanyName = Rst.Fields("CompanyName")
CompanyAddress = FormatField(Rst.Fields("Address"))
PhoneNum = FormatField(Rst.Fields("PhoneNo"))
DLNo = FormatField(Rst.Fields("DLNo"))
Kst = FormatField(Rst.Fields("KST"))
Cst = FormatField(Rst.Fields("CST"))

With m_TextStream
    
    .WriteLine GetCentreAlignedText(Chr(14) & m_CompanyName)
    .WriteLine GetCentreAlignedText(CompanyAddress) & Chr(19)
    .WriteLine GetCentreAlignedText("Phone No : " & PhoneNum) & Chr(19)
    .WriteLine GetCentreAlignedText("Drug License No : " & DLNo) & Chr(19)
    .WriteLine GetCentreAlignedText("KST No : " & Kst & ", CST No : " & Cst) & Chr(19)
    .WriteLine GetHorizontalLineText
End With
End Sub
Private Function GetCompanyTitleRecordSet(ByRef rstCompanyTitle As ADODB.Recordset) As Boolean

Dim CompanyType As wis_CompanyType

CompanyType = Enum_Self


gDbTrans.SqlStmt = " SELECT CompanyName,Address,PhoneNo,DLNo,KST,CST" & _
                   " FROM CompanyCreation" & _
                   " WHERE CompanyType=" & CompanyType

Call gDbTrans.Fetch(rstCompanyTitle, adOpenForwardOnly)


End Function

Private Sub PrintFirstSecondSalesAmount(ByVal FirstSalesAmount As Currency, ByVal SecondSalesAmount As Currency)
Dim strText As String * 22

With m_TextStream
    .WriteBlankLines (1)
    strText = "First Sales"
    AmountStr = FormatCurrency(FirstSalesAmount)
    .Write strText & GetRightAlignedText(AmountStr, Len(AmountStr))
    
    strText = "    Second Sales"
    AmountStr = FormatCurrency(SecondSalesAmount)
    .WriteLine strText & GetRightAlignedText(AmountStr, Len(AmountStr))
End With

End Sub


Private Sub PrintAccountData(MaterialClass As clsMaterial, ByVal ParentID As Long)
'Declare the variables
Dim Amount As Currency

Dim IncomeLine As Boolean
Dim ExpLine As Boolean
Dim strText As String * 22

'Setup an error handler...
On Error GoTo ErrLine


With m_TextStream
    Select Case ParentID
        Case wis_IncomeParentID, wis_TradingIncomeParentID
            Amount = MaterialClass.Amount
            m_Total = m_Total + Amount
            strText = "    " & MaterialClass.ProductName
            AmountStr = FormatCurrency(Amount)
            
            If m_ExpenseHeadPrinted Then
                .WriteLine strText & GetRightAlignedText(AmountStr, Len(AmountStr))
                m_IncomeHeadPrinted = True
                m_ExpenseHeadPrinted = True
            Else
                .Write strText & GetRightAlignedText(AmountStr, Len(AmountStr))
                m_IncomeHeadPrinted = True
            End If
            
            
        Case wis_ExpenseParentID, wis_TradingExpenseParentID
        
            Amount = MaterialClass.Amount
            m_Total = m_Total - Amount
            strText = MaterialClass.ProductName
            AmountStr = FormatCurrency(Amount)
            If m_ExpenseHeadPrinted Then
                .WriteLine strText & GetRightAlignedText(AmountStr, Len(AmountStr))
                m_ExpenseHeadPrinted = False
            Else
                .Write strText & GetRightAlignedText(AmountStr, Len(AmountStr))
                m_ExpenseHeadPrinted = True
            End If
            
    End Select
    
End With

Exit Sub

ErrLine:
    MsgBox "PrintAccountData: " & Err.Description, vbCritical

End Sub


Private Sub PrintSubTotalLine()
Dim LineNum As Integer
With m_TextStream
    LineNum = .Line
    If LineNum < 50 Then .WriteBlankLines (50 - LineNum)
    .WriteLine GetHorizontalLineText
    
    MfrStr = "": ParticularStr = "": BatchNoStr = ""
    If Not IsNextPageRequired Then
        .WriteLine MfrStr & ParticularStr & BatchNoStr & "Total Qty   " & (m_TotalQuantity + m_TotalFreeQuantity) & "       Sub Total  " & GetRightAlignedText(FormatCurrency(m_Total), 13)
        .WriteLine GetHorizontalLineText
    End If
    
End With

End Sub


Private Sub PrintInventoryItems(MaterialClass As clsMaterial)
'Declare the Vraibles
Dim Quantity As Double
Dim FreeQuantity As Double
Dim ExpDate As String
Dim Amount As Currency

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    If IsNextPageRequired Then
        .Write GetRightAlignedText("B/F from Previous Page", Len(MfrStr & ParticularStr & BatchNoStr & ExpDateStr & QtyStr & "B/F from Previous Page"))
        .WriteLine GetRightAlignedText(FormatCurrency(m_Total), Len(m_Total) + 2)
    End If
    
    MfrStr = MaterialClass.Manufacterer
    ParticularStr = MaterialClass.ProductName
    BatchNoStr = MaterialClass.BatchNo
    ExpDate = MaterialClass.ExpDate
    ExpDateStr = ""
    If ExpDate <> "" Then ExpDateStr = Right$(ExpDate, Len(ExpDate) - 3)
    
    Quantity = MaterialClass.Quantity
    QtyStr = Quantity
    FreeQuantity = MaterialClass.FreeQuantity
    BonusStr = FreeQuantity
    
    SPriceStr = FormatCurrency(MaterialClass.UnitPrice)
    TRPStr = FormatCurrency(MaterialClass.TradingPrice)
    MRPStr = FormatCurrency(MaterialClass.MRP)
    KSTStr = MaterialClass.Kst
    Amount = MaterialClass.Amount
    AmountStr = FormatCurrency(Amount)
    
    .Write MfrStr & ParticularStr & BatchNoStr & ExpDateStr
    .Write GetRightAlignedText(QtyStr, Len(QtyStr)) & GetRightAlignedText(BonusStr, Len(BonusStr))
    .Write GetRightAlignedText(SPriceStr, Len(SPriceStr)) & GetRightAlignedText(TRPStr, Len(TRPStr))
    .Write GetRightAlignedText(MRPStr, Len(MRPStr)) & GetRightAlignedText(KSTStr, Len(KSTStr) + 1)
    .WriteLine GetRightAlignedText(AmountStr, Len(AmountStr))
    
        
End With
    
m_TotalQuantity = m_TotalQuantity + Quantity
m_TotalFreeQuantity = m_TotalFreeQuantity + FreeQuantity
m_Total = m_Total + Amount


Exit Sub

ErrLine:
    MsgBox "PrintInventoryItems() " & Err.Description
    

End Sub


Private Sub PrintColumnHeader()
'Declare the variables
Dim curX As Single
Dim CurY As Single
Dim Ratio As Single
Dim ColWid As Single

MfrStr = "Mfr"
ParticularStr = "Particulars"
BatchNoStr = "BatchNo "
ExpDateStr = "ExpDate "
QtyStr = "Qty "
BonusStr = "Bon "
SPriceStr = "S.Price"
TRPStr = "TRP"
MRPStr = "MRP"
KSTStr = "KST"
AmountStr = "Amount"

With m_TextStream
    .Write MfrStr & ParticularStr & BatchNoStr & ExpDateStr & GetRightAlignedText(QtyStr, Len(QtyStr)) & GetRightAlignedText(BonusStr, Len(BonusStr))
    .Write GetRightAlignedText(SPriceStr, Len(SPriceStr))
    .Write GetRightAlignedText(TRPStr, Len(TRPStr))
    .Write GetRightAlignedText(MRPStr, Len(MRPStr))
    .Write GetRightAlignedText(KSTStr, Len(KSTStr))
    .WriteLine GetRightAlignedText(AmountStr, Len(AmountStr))
    .WriteLine GetHorizontalLineText
End With


End Sub


Private Sub PrintProperties()
'Declare the variables

Dim PayTerm As wis_PaymentTerm
Dim PayTermText As String

Dim rstCustomer As ADODB.Recordset

Dim CustString As String * 43
Dim InvString As String * 17
Dim InvNoString As String * 10
Dim InvDateString As String * 9

With m_TextStream
    .WriteLine GetCentreAlignedText(m_InvoiceName)
    .WriteBlankLines (1)
    .WriteLine GetHorizontalLineText
End With

'Get  ConsigneeName Address CST
'CopanyName,Address,DLNo,KST,CST
Call GetCustomerRecordSet(rstCustomer)

If rstCustomer.RecordCount < 0 Then
    MsgBox "Error in Getting Customer Details", vbCritical
    Exit Sub
End If

With m_TextStream
    
    CustString = "Consignee:"
    InvString = "Invoice No"
    InvNoString = m_InvoiceNo
    InvDateString = "Inv Date"
    .WriteLine CustString & InvString & InvNoString & InvDateString & m_PrintDate
    
    CustString = rstCustomer.Fields("CompanyName")
    InvString = "Quatation No"
    InvNoString = FormatField(m_rstPrintProperty.Fields("QuotationNo"))
    InvDateString = "Qtn Date"
    .WriteLine CustString & InvString & InvNoString & InvDateString & m_PrintDate & FormatField(m_rstPrintProperty.Fields("QuotationDate"))
    
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

    CustString = FormatField(rstCustomer.Fields("Address"))
    InvString = "Payment Term"
    InvNoString = PayTermText
    InvDateString = "L.R.No"
    .WriteLine CustString & InvString & InvNoString & InvDateString & FormatField(m_rstPrintProperty.Fields("LRNo"))
    
    CustString = FormatField(rstCustomer.Fields("DLNo"))
    InvString = "Transport Mode "
    InvNoString = FormatField(m_rstPrintProperty.Fields("TransModeName"))
    .WriteLine CustString & InvString & InvNoString & FormatField(m_rstPrintProperty.Fields("LRNo"))
    
    CustString = FormatField(rstCustomer.Fields("KST")) & " " & FormatField(rstCustomer.Fields("CST"))
    InvString = "DC No"
    InvNoString = FormatField(m_rstPrintProperty.Fields("DCNo"))
    .WriteLine CustString & InvString & InvNoString
    
    .WriteLine GetHorizontalLineText
End With

End Sub

Private Function GetCustomerRecordSet(ByRef rstCustomer As ADODB.Recordset) As Boolean


gDbTrans.SqlStmt = " SELECT CompanyName,Address,DLNo,KST,CST" & _
                   " FROM CompanyCreation" & _
                   " WHERE HeadID=" & m_CustomerHeadID

Call gDbTrans.Fetch(rstCustomer, adOpenForwardOnly)


End Function


Private Sub PrintHeading()
Dim TitlePrint As wis_PrintTitle

TitlePrint = FormatField(m_rstPrintProperty.Fields("TitlePrint"))

If TitlePrint = Enum_PrintTitle Then
    PrintCompanyTitle
Else
    m_TextStream.WriteBlankLines (8)
End If
End Sub


Private Function SetPrintPropertyRecordset(ByRef rstProperty As ADODB.Recordset, ByVal SaleTransID As Long) As Long

gDbTrans.SqlStmt = " SELECT *,TransModeName" & _
                   " FROM InvoicePrint A,TransportMode B" & _
                   " WHERE SaleTransID=" & SaleTransID & _
                   " AND A.TransportMode=B.TransModeID"

SetPrintPropertyRecordset = gDbTrans.Fetch(rstProperty, adOpenStatic)


End Function


Private Function GetHorizontalLineText() As String
'Declare the variables
Dim SourceWidth As Single
Dim strSource As String
Dim Wid As Single
'Setup an error handler
On Error GoTo ErrLine
 
strSource = "----------------------"


Do
    If SourceWidth >= PAPER_WIDTH Then Exit Do
    
    strSource = "-----" & strSource & "-----"
    
    SourceWidth = Len(strSource)
    
Loop

GetHorizontalLineText = strSource
Exit Function

ErrLine:
    MsgBox "GetHorizontalLineText: " & Err.Description
    
End Function




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
Private Function GetRightAlignedText(ByVal strSource As String, TextWid As Integer) As String

'Setup an error handler...
On Error GoTo ErrLine

strSource = Trim$(strSource)

Do
    If TextWid <= Len(strSource) Then Exit Do
    
    strSource = Space(1) & strSource
    
Loop

GetRightAlignedText = strSource

Exit Function

ErrLine:
    MsgBox "GetRightAlignedText() " & Err.Description, vbCritical


End Function

Private Function GetCentreAlignedText(ByVal strSource As String) As String


'Setup an error handler...
On Error GoTo ErrLine

GetCentreAlignedText = Space((PAPER_WIDTH - Len(strSource)) / 2) & strSource & Space((PAPER_WIDTH - Len(strSource)) / 2)

Exit Function

ErrLine:
    MsgBox "CentreAlignData() " & Err.Description, vbCritical


End Function
Private Function AlignText(ByVal strSource As String, ByVal Wid As Single) As String
'Declare the variables
Dim SourceWidth As Single

'Setup an error handler...
On Error GoTo ErrLine

SourceWidth = my_Printer.TextWidth(strSource)

Do
    If SourceWidth >= Wid Then Exit Do
    strSource = Space(1) & strSource
    SourceWidth = my_Printer.TextWidth(strSource)
Loop
    
        
AlignText = strSource

Exit Function

ErrLine:
    MsgBox "AlignText() " & Err.Description, vbCritical
    


End Function
Public Sub WriteToFile()

   Dim fld As ADODB.Field
   Dim strRow As String
   Dim strField As String
   Dim rsCustomers As ADODB.Recordset
   
   Open "Customers.txt" For Output As #1

gDbTrans.SqlStmt = " SELECT *" & _
                  " FROM CompanyCreation" & _
                  " WHERE CompanyType=" & 2

Call gDbTrans.Fetch(rsCustomers, adOpenDynamic)

   With rsCustomers

      .MoveFirst
      Do While Not .EOF
         For Each fld In .Fields
            ' If a field has a value, add quotation marks.
            If Len(fld.Value) > 0 Then
               strField = Chr(34) & fld.Value & Chr(34)
            Else
               strField = ""
            End If
            ' Add the field value and a tab delimeter
            ' to the output string.
            strRow = strRow & strField & Chr(9)
         Next
         ' Strip off the end tab character.
         strRow = Left(strRow, Len(strRow) - 1)
         ' Print the output string.
         Print #1, strRow
         strRow = ""
      .MoveNext
      Loop
   End With
   Close

strRow = GetCentreAlignedText("M/s R.Y.Shettar And Company")
End Sub


Private Sub Class_Initialize()
Set my_Printer = Printer
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
Public Property Get InvoiceName() As String

InvoiceName = m_InvoiceName
End Property

Public Property Let InvoiceName(ByVal vNewValue As String)
m_InvoiceName = vNewValue
End Property
Public Property Set CollectionClass(ByVal clsMaterialCol As clsMaterials)
Set m_CollectionClass = clsMaterialCol

End Property

Private Sub Class_Terminate()
Set my_Printer = Nothing
Set m_FileSystemObject = Nothing
Set m_TextStream = Nothing
Set m_CollectionClass = Nothing
Set m_DataSource = Nothing
Set m_rstPrintProperty = Nothing


End Sub


