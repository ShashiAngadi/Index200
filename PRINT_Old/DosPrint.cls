VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDosPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare the consts
Private Const MIN_ITEMS_ON_PAGE As Integer = 30
Private Const MAX_ITEMS_ON_PAGE As Integer = 40
Private m_ItemsInSecondLastPage As Integer

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

Private m_TotalPages As Integer
Private m_PageNo As Integer
Private IsNextPageRequired As Boolean
Private m_NoOfLines As Integer

Private m_Total As Currency
Private m_TotalQuantity As Double
Private m_TotalFreeQuantity As Double

'''
' These variables are required for Column alignment
Private ParticularStr As String * 24
Private PackingStr As String * 7
Private QtyStr As String * 5
Private SPriceStr As String * 7
Private AmountStr As String * 13

Private Enum wis_PrintFont_Types
    Enum_Roman = 0
    Enum_Sanseriff = 1
    Enum_Courier = 2
    Enum_Prestige = 3
    Enum_Script = 4
    Enum_Orator = 7
End Enum
Private Enum wis_PrintEnlarged
    Normal_WtHt = 0
    Double_Height = 1
    Quadruple_Height = 2
    Double_width = 3
    Quadruple_Width = 4
    Double_HtWt = 5
    Quadruple_HtWt = 6
End Enum


Private Sub AssignPaperSizeToLocalVariables()
'Declare the variables
Dim Rst As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

gDbTrans.SqlStmt = " SELECT ValueData" & _
                  " FROM Install" & _
                  " WHERE KeyData='PageHeight'"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

m_PrinterHeight = FormatField(Rst.Fields(0))

gDbTrans.SqlStmt = " SELECT ValueData" & _
                  " FROM Install" & _
                  " WHERE KeyData='PageWidth'"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

m_PrinterWidth = FormatField(Rst.Fields(0))

Set Rst = Nothing

Exit Sub

ErrLine:
        MsgBox "AssignPaperSizeToLocalVariables " & Err.Description
        

End Sub


Public Property Set DataSource(ByVal grd As MSFlexGrid)
Set m_DataSource = grd

End Property


Private Sub InitialisePrinter()
With m_TextStream
    .write Chr$(27) & Chr$(64) '@

'Set Line spacing n/72
'Here  n=8
'    .write Chr$(27) & Chr$(65) & Chr$(10)

End With
End Sub

Private Sub OpenTextFile()

If m_FileSystemObject Is Nothing Then Set m_FileSystemObject = New FileSystemObject

If IsNextPageRequired Then
    Set m_TextStream = m_FileSystemObject.OpenTextFile(App.Path & "\InvPrint.txt", ForAppending, True)
Else
    m_NoOfLines = 0
    Set m_TextStream = m_FileSystemObject.CreateTextFile(App.Path & "\InvPrint.txt", True, False)
End If


End Sub

Public Function PrintInvoice(ByVal SaleTransID As Long, ByVal ItemCount As Integer) As Boolean


'Setup error handler...
On Error GoTo ErrLine

' If no recordset available, exit.
If m_DataSource Is Nothing Then Err.Raise vbObjectError + 513, , "No Data"

If m_CollectionClass Is Nothing Then Err.Raise vbObjectError + 513, , "No Data"

If SetPrintPropertyRecordset(m_rstPrintProperty, SaleTransID) < 1 Then
    MsgBox "Please Set the print Properties", vbInformation
    Exit Function
End If

Dim MaxCount As Integer
Dim ActualDiff As Integer



'Open the textfile
OpenTextFile

If Not IsNextPageRequired Then
    InitialisePrinter
    
    'Set Paper height
    SetPaperSize
    
    'Calculate item on lAst Page
    MaxCount = m_CollectionClass.Count
    m_TotalPages = 1
    If MaxCount > MIN_ITEMS_ON_PAGE Then
        m_TotalPages = ((MaxCount / MAX_ITEMS_ON_PAGE) + 0.5) / 1
        If m_TotalPages = 1 Then m_TotalPages = 2
        ActualDiff = MaxCount - MAX_ITEMS_ON_PAGE * (m_TotalPages - 1)
        If ActualDiff > MIN_ITEMS_ON_PAGE Then
            m_TotalPages = m_TotalPages + 1
            ActualDiff = ActualDiff - MIN_ITEMS_ON_PAGE
            If ActualDiff <= MAX_ITEMS_ON_PAGE - MIN_ITEMS_ON_PAGE Then _
                        m_ItemsInSecondLastPage = MIN_ITEMS_ON_PAGE
        Else
            m_ItemsInSecondLastPage = MAX_ITEMS_ON_PAGE
            If ActualDiff <= 0 Then _
                m_ItemsInSecondLastPage = MIN_ITEMS_ON_PAGE
        End If
    End If

End If

SetPrintQuality False
'Prnt the Company Title Address etc
PrintHeading

'Print the Properties Such as Invoice Number, Consignee etc
PrintProperties

'PRINT the column headers
PrintColumnHeader

'Print the Details Such as Quantity, price etc.
Call PrintDetails(ItemCount)

'Print the Declaration
PrintDeclaration

'Print Footer
If IsNextPageRequired Then
    m_TextStream.write Chr$(12)
    m_TextStream.Close
    
'    Shell App.Path & "\PrintBatch.bat", vbHide
       
       
    Call PrintInvoice(SaleTransID, ItemCount)
Else
    PrintFooter
    
    m_TextStream.write Chr$(12)
    m_TextStream.Close
    Shell App.Path & "\PrintBatch.bat", vbHide
    m_NoOfLines = 0
End If

PrintInvoice = True

Exit Function

ErrLine:
    MsgBox "PrintInvoice: " & vbCrLf & Err.Description, vbCritical
    
End Function
Private Sub SetDoubleHeightOfChar(ByVal IsSet As Boolean)
With m_TextStream
    If IsSet Then
        .write Chr$(27) & Chr$(119) & Chr$(1)
    Else
        .write Chr$(27) & Chr$(119) & Chr$(0)
    End If
End With
End Sub

Private Sub PrintFooter()

With m_TextStream

    SetPrintQuality True
    SetMasterTypeStyles (60)
    AlignRight
    .WriteLine "For " & m_CompanyName & "     "
    .WriteBlankLines (2)
    .WriteLine "Authorised Signatory          "
    SetMasterTypeStyles (1)
    AlignLeft
End With

End Sub


Private Sub PrintDeclaration()
Dim Rst As ADODB.Recordset
Dim Ret As Long


If IsNextPageRequired Then Exit Sub
AlignLeft
SetPrintQuality True
With m_TextStream
    .WriteLine "E & O.E"
End With

'Get the Declaration From the Database
gDbTrans.SqlStmt = " SELECT ValueData" & _
                   " FROM Install" & _
                   " WHERE KeyData LIKE 'Declaration%'"

Ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)

If Ret <= 0 Then Exit Sub

SetMasterTypeStyles 4

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




Private Sub PrintDetails(ByRef ItemCount As Integer)
'Declare the variables
Dim MaxCount As Integer
Dim ItemsOnPage As Integer
Dim MaxPerPage  As Integer

Dim ParentID As Long
Dim LinePrinted As Boolean

Dim MaterialClass As clsMaterial

Dim lKeys

MaxCount = m_CollectionClass.Count
lKeys = m_CollectionClass.GetKeys
ItemsOnPage = 0

If IsNextPageRequired Then ItemCount = ItemCount + 1
    
SetPrintQuality True

MaxPerPage = MAX_ITEMS_ON_PAGE
If m_PageNo + 1 = m_TotalPages - 1 Then MaxPerPage = m_ItemsInSecondLastPage

Do
    If ItemCount >= MaxCount Then Exit Do

    Set MaterialClass = m_CollectionClass.Item(lKeys(ItemCount))
    
    ParentID = MaterialClass.ParentID
    
    If ParentID = 0 Then Call PrintInventoryItems(MaterialClass)
    
        
    If ParentID > 0 Then
        
        Dim ArrCount As Integer
        Dim arrParentID() As Long
        Dim arrHeadName() 'Two DImensional Array


        ReDim Preserve arrParentID(ArrCount)
        ReDim Preserve arrHeadName(1, ArrCount)

          
        arrParentID(ArrCount) = ParentID
        arrHeadName(0, ArrCount) = MaterialClass.ProductName
        arrHeadName(1, ArrCount) = MaterialClass.Amount
        
        ArrCount = ArrCount + 1
    
    End If
    
    IsNextPageRequired = False
    
    ItemsOnPage = ItemsOnPage + 1
    If Not LinePrinted Then
        If ItemsOnPage >= MaxPerPage Then
            m_PageNo = m_PageNo + 1
            IsNextPageRequired = True
            Exit Do
        End If
    End If
    
    ItemCount = ItemCount + 1
    
Loop


SetPrintQuality False

'If Parentid does not contain any wis_ExpenseID, wis_TradingExpenseID
'wis_IncomeID, wis_TradingIncomeID
If Not LinePrinted Then
    Call PrintSubTotalLine(ItemCount)
    LinePrinted = True
End If


'Print Account Data
If ArrCount > 0 Then Call PrintAccountData(arrParentID, arrHeadName)

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
    ParticularStr = "": PackingStr = ""
    QtyStr = "": SPriceStr = ""
    AmountStr = CurrencyRoundOff(m_Total)
    
    SetPrintQuality True
    If Not IsNextPageRequired Then
        AlignRight
        
        'SetDoubleHeightOfChar True
        SetMasterTypeStyles (44)
        .WriteLine "       GRAND TOTAL " & AmountStr
        'SetDoubleHeightOfChar False
        AlignLeft
        Set CurFigureClass = New clsCurFigure
        
        CurFigures = CurFigureClass.NumberInFigure(CCur(AmountStr))
        .WriteLine "Rupees " & CurFigures & " Only"
        SetMasterTypeStyles (1)
        Set CurFigureClass = Nothing
    Else
        SetMasterTypeStyles (44)
        AlignRight
        '.WriteLine "Continued to " & m_PageNo + 1 & "    "
        .WriteLine "C/F Rs. " & FormatCurrency(m_Total) & "    "
        SetMasterTypeStyles (1)
        AlignLeft
    End If
        
    .WriteLine GetHorizontalLineText
    

    
End With

End Sub

Private Sub PrintCompanyTitle()
'Declare the variables
Dim Rst As ADODB.Recordset
Dim CompanyAddress As String
Dim PhoneNum As String
Dim DLNo As String
Dim Kst As String
Dim Cst As String
Dim DOB As String

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
Cst = FormatField(Rst.Fields("CSTNo"))
DOB = FormatField(Rst.Fields("DOB"))

With m_TextStream
    SetDoubleHeightOfChar (True)
    SetMasterTypeStyles (44)
    SetPrintQuality (False)
    
    AlignCenter
    .WriteLine m_CompanyName
    SetMasterTypeStyles (1)
    SetDoubleHeightOfChar False
    
    SetPrintQuality (True)
    .WriteLine CompanyAddress
    .WriteLine "Phone No : " & PhoneNum & "     Drug License No : " & DLNo
    .WriteLine "KST No : " & Kst & ", CST No : " & Cst & ",Date " & DOB
    
    'Move Printer Head to the Left Margin
    AlignLeft
    .WriteLine GetHorizontalLineText
End With

End Sub
Private Function GetCompanyTitleRecordSet(ByRef rstCompanyTitle As ADODB.Recordset) As Boolean

Dim CompanyType As wis_CompanyType

CompanyType = Enum_Self

gDbTrans.SqlStmt = " SELECT CompanyName,Address,PhoneNo,DLNo,VATNo as KST,CSTNo,DOB" & _
                   " FROM CompanyDet"

Call gDbTrans.Fetch(rstCompanyTitle, adOpenForwardOnly)


End Function




Private Sub PrintAccountData(arrParentID() As Long, arrHeadNames())
'Declare the variables
Dim Amount As Currency
Dim ParentID As Long

Dim ItemCount As Integer
Dim MaxCount As Integer

Dim strText As String * 17

Dim IncomePrinted As Boolean
Dim ExpensePrinted As Boolean

'Setup an error handler...
On Error GoTo ErrLine

Dim MidCount As Byte
Dim RightCount As Byte
Dim LeftCount As Byte
Dim ArrCount As Byte
Dim StrToPrint() As String

ReDim StrToPrint(0)

MaxCount = UBound(arrParentID)

For ItemCount = 0 To MaxCount

    ParentID = arrParentID(ItemCount)
    strText = arrHeadNames(0, ItemCount)
    Amount = arrHeadNames(1, ItemCount)
    
    AmountStr = FormatCurrency(Amount)
    AmountStr = GetRightAlignedText(AmountStr, Len(AmountStr))
    'Get the total amount
    
    Select Case ParentID
        
        Case wis_IncomeParentID To (wis_IncomeParentID + HEAD_OFFSET - 1)
            m_Total = m_Total + Amount
            If MidCount > ArrCount Then _
                ArrCount = ArrCount + 1: ReDim Preserve StrToPrint(ArrCount)
            
            If Len(StrToPrint(MidCount)) > 66 Then
                'the Right string Already exists
                StrToPrint(ArrCount) = Left(StrToPrint(ArrCount), 33) & _
                    strText & AmountStr & Mid(StrToPrint(ArrCount), 68)
            ElseIf Len(StrToPrint(MidCount)) > 1 Then
                'the left string Already exists
                StrToPrint(MidCount) = StrToPrint(MidCount) & _
                     Space(3) & strText & AmountStr
            Else
                StrToPrint(MidCount) = String(34, " ") & strText & AmountStr
            End If
            MidCount = MidCount + 1
            
        Case wis_ExpenseParentID To (wis_ExpenseParentID + HEAD_OFFSET - 1)
            
            m_Total = m_Total - Amount
            
            If LeftCount > ArrCount Then _
                ArrCount = ArrCount + 1: ReDim Preserve StrToPrint(ArrCount)
            
            If Len(StrToPrint(LeftCount)) > 33 Then
                'the Right or middle string Already exists
                StrToPrint(ArrCount) = strText & AmountStr & _
                        Mid(StrToPrint(LeftCount), 34)
            Else
                StrToPrint(LeftCount) = strText & AmountStr
            End If
            
            LeftCount = LeftCount + 1
        
        Case wis_DrNoteParentID
            
            m_Total = m_Total + Amount
            
            If RightCount > ArrCount Then _
                ArrCount = ArrCount + 1: ReDim Preserve StrToPrint(ArrCount)
            
            
            StrToPrint(RightCount) = StrToPrint(RightCount) & _
                String(66 - Len(StrToPrint(ArrCount)), " ") & strText & AmountStr
            
            RightCount = RightCount + 1

        Case wis_CrNoteParentID
        
            m_Total = m_Total - Amount
            
            StrToPrint(RightCount) = StrToPrint(RightCount) & _
                String(66 - Len(StrToPrint(RightCount)), " ") & strText & AmountStr
            
            RightCount = RightCount + 1
        
            
    End Select

Next ItemCount

ItemCount = 0

With m_TextStream
    For ItemCount = 0 To ArrCount
        .WriteLine StrToPrint(ItemCount)
        
    Next ItemCount
End With

Exit Sub

ErrLine:
    MsgBox "PrintAccountData: " & Err.Description, vbCritical

End Sub


Private Sub PrintSubTotalLine(ByVal TotalItems As Integer)
Dim LineNum As Integer


With m_TextStream
    LineNum = .Line
    
    If LineNum < 44 + m_NoOfLines Then .WriteBlankLines (44 + m_NoOfLines - LineNum)
    m_NoOfLines = LineNum
    
    
    .WriteLine GetHorizontalLineText
    
    ParticularStr = "": PackingStr = "":
    If Not IsNextPageRequired Then
        .WriteLine ParticularStr & "  Items = " & TotalItems & "            Qty = " & (m_TotalQuantity + m_TotalFreeQuantity) & "                    Sub Total = " & FormatCurrency(m_Total)
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
Dim tmpStr As String
Dim lKST As Single

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    If IsNextPageRequired Then
        SetMasterTypeStyles (44)
        AlignRight
        .WriteLine "      B/F From Page No " & m_PageNo & "     " & FormatCurrency(m_Total) & "    "
        SetMasterTypeStyles (1)
        AlignLeft
    End If
    
    ParticularStr = MaterialClass.ProductName
    
    Quantity = MaterialClass.Quantity
    QtyStr = Quantity
    SPriceStr = FormatCurrency(MaterialClass.UnitPrice)
    Amount = MaterialClass.Amount
    AmountStr = FormatCurrency(Amount)
    
    .write ParticularStr & PackingStr
    .write GetRightAlignedText(QtyStr, Len(QtyStr))
    .write GetRightAlignedText(SPriceStr, Len(SPriceStr))
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

ParticularStr = "Particulars"
PackingStr = "Pack"
QtyStr = " Qty   "
SPriceStr = "Rate"
AmountStr = "Amount"

AlignLeft

With m_TextStream

    .write ParticularStr & PackingStr & GetRightAlignedText(QtyStr, Len(QtyStr))
    .write GetRightAlignedText(SPriceStr, Len(SPriceStr))
    .WriteLine GetRightAlignedText(AmountStr, Len(AmountStr))
    .WriteLine GetHorizontalLineText
End With

End Sub


Private Sub PrintProperties()
'Declare the variables
Dim PayTerm As wis_PaymentTerm
Dim PayTermText As String

Dim rstCustomer As ADODB.Recordset

Dim CustString As String * 47
Dim InvString As String * 17
Dim InvNoString As String * 10
Dim InvDateString As String * 10

'Get  ConsigneeName Address CST
'CopanyName,Address,DLNo,KST,CST
Call GetCustomerRecordSet(rstCustomer)

If rstCustomer.RecordCount < 0 Then
    MsgBox "Error in Getting Customer Details", vbCritical
    Exit Sub
End If

With m_TextStream
    
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

    InvNoString = PayTermText
    
    InvDateString = "INVOICE NO:"
    
    SetPrintQuality False
    SetPitchOfLetter
    
    SetMasterTypeStyles (44)
    AlignLeft
    
    .WriteLine "TERMS: " & InvNoString & "      " & InvDateString & ": " & m_InvoiceNo & "             DATE: " & m_PrintDate
    
    SetMasterTypeStyles 1
    
    
    AlignLeft
    
    SetPrintQuality True
    
    CustString = rstCustomer.Fields("CompanyName")
    SetMasterTypeStyles 44
    .WriteLine CustString & "     Page No " & m_PageNo + 1 & " of " & m_TotalPages
    SetMasterTypeStyles 1
    
    CustString = FormatField(rstCustomer.Fields("Address"))
    InvString = "Quatation No"
    InvNoString = FormatField(m_rstPrintProperty.Fields("QuotationNo"))
    InvDateString = "Qtn Date"
    
    .WriteLine CustString & InvString & InvNoString & InvDateString & FormatField(m_rstPrintProperty.Fields("QuotationDate"))
    
    CustString = ""
    InvString = "Transport Mode "
    InvNoString = FormatField(m_rstPrintProperty.Fields("TransModeName"))
    InvDateString = "L.R.No"
    .WriteLine CustString & InvString & InvNoString & InvDateString & FormatField(m_rstPrintProperty.Fields("LRNo"))
    
    
    CustString = FormatField(rstCustomer.Fields("KST")) & " " & FormatField(rstCustomer.Fields("CST"))
    InvString = "DC No"
    InvNoString = FormatField(m_rstPrintProperty.Fields("DCNo"))
    .WriteLine CustString & InvString & InvNoString
    
    .WriteLine GetHorizontalLineText
    
End With

End Sub

Private Function GetCustomerRecordSet(ByRef rstCustomer As ADODB.Recordset) As Boolean


gDbTrans.SqlStmt = " SELECT CompanyName,Address,KST,CST" & _
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


Private Sub SetEnlargedPrint(ByVal PrintType As wis_PrintEnlarged)

With m_TextStream
    .write Chr$(27) & Chr$(126) & Chr$(49) & Chr$(PrintType)
End With
'Decimal 27 126 49 n
End Sub

Private Sub SetFontName(ByVal FontNames As wis_PrintFont_Types)

Call SetPrintQuality(False)

With m_TextStream
    .write Chr$(27) & "k" & FontNames
End With

End Sub

Private Sub SetMasterTypeStyles(ByVal MasterSelectNo As Integer)
'Here Master Select No can be selected as follows
'Feature            DecimalCode (MasterSelectNo
'Pica               0 '10 Charecters Per Inch
'Elite              1 '12 Characters Per INch
'Proportional       2
'Condensed          4
'Emphasized         8
'double-Strike      16
'Dowble width       32
'Italic             64
'Underline          128
'If u want the Emphasizedand double strike and doublewidth format then (8+16+31)
' Is the masterselectNo
With m_TextStream
    
    .write Chr$(27) & Chr$(33) & Chr$(MasterSelectNo)
    
End With

End Sub

Private Sub SetRightMargin()

With m_TextStream
    'Sets the no of Columns
    .write Chr$(27) & Chr$(81) & Chr$(89) 'Ends with 88th Columns
End With

End Sub

Private Sub SetLeftMargin()

With m_TextStream
    'Sets the no of Columns
    .write Chr$(27) & Chr$(108) & Chr$(3) 'It starts with 3rd columns
End With

End Sub

Private Sub SetPaperSize()
'Declarte the variables

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    
    'Set the Page Length to Inches
    .write Chr$(27) & Chr$(67) & Chr$(0) & Chr$(12)  'It is 12 Inches
    
End With

'Set Horizontal widths
SetLeftMargin

SetRightMargin


Exit Sub

ErrLine:
        MsgBox "SetPaperSize() " & Err.Description
        
End Sub
Private Sub SetPitchOfLetter()


With m_TextStream
    '.Write Chr$(27) & Chr$("80") 'Sets the pitch to 10 characters Per Inch
    .write Chr$(27) & Chr$("77")  'Sets the pitch to 12 characters Per Inch
End With

End Sub


Private Function SetPrintPropertyRecordset(ByRef rstProperty As ADODB.Recordset, ByVal SaleTransID As Long) As Long

gDbTrans.SqlStmt = " SELECT A.*,TransModeName" & _
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
 
strSource = String$(98, "-")

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

Private Sub AlignLeft()
'DATA
'0 Left Justified
'1 Centered
'2 Right Aligned
'3 Justified both the sides

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    .write Chr$(27) & Chr$(97) & Chr$(0)
End With


Exit Sub

ErrLine:
    MsgBox "AlignLeft() " & Err.Description, vbCritical


End Sub
Private Sub AlignRight()
'DATA
'0 Left Justified
'1 Centered
'2 Right Aligned
'3 Justified both the sides

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    .write Chr$(27) & Chr$(97) & Chr$(2)
End With


Exit Sub

ErrLine:
    MsgBox "AlignRight() " & Err.Description, vbCritical

End Sub

Private Sub AlignCenter()

'DATA
'0 Left Justified
'1 Centered
'2 Right Aligned
'3 Justified both the sides

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    .write Chr$(27) & Chr$(97) & Chr$(1)
End With


Exit Sub

ErrLine:
    MsgBox "CentreAlignData() " & Err.Description, vbCritical


End Sub
Private Function GetCentreAlignedText(ByVal strSource As String) As String
'DATA
'0 Left Justified
'1 Centered
'2 Right Aligned
'3 Justified both the sides

'Setup an error handler...
On Error GoTo ErrLine

With m_TextStream
    .write Chr$(27) & Chr$(97) & Chr$(1) & strSource
End With


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
Private Sub SetPrintQuality(ByVal DraftQality As Boolean)

With m_TextStream
    If DraftQality Then
        .write Chr$(27) & "x0"
    Else
        .write Chr$(27) & "x1"
    End If
End With

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
