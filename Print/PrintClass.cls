VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                            phPrinter As Long, ByVal pDefault As Long) As Long

Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" _
            (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, _
                ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
   (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
   (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, _
   ByVal Level As Long, pJob As Byte, ByVal cdBuf As Long, pcbNeeded As Long, _
   pcReturned As Long) As Long



Private Const TITLE_TOP As Integer = 50
Private Const FIELD_MARGIN  As Integer = 200
Private Const PAGE_MARGIN As Integer = 500
Private Const RECT_MARGIN As Integer = 15
''Printer STATUS  ' Chaned Public to private

Private m_lhPrinter As Long
Private my_Printer As Printer
Private m_DataSource As Object

Private m_FileSysObject As FileSystemObject
Private m_TextStream As TextStream

Public Property Get ObjectName()
    ObjectName = "PrintClass"
End Property


Public Sub ClosePrinterConnection()
'Declare the variables
Dim lReturn As Long

On Error GoTo ErrLine

lReturn = EndPagePrinter(m_lhPrinter)
If lReturn = 0 Then Err.Raise vbObjectError + 513

lReturn = EndDocPrinter(m_lhPrinter)
If lReturn = 0 Then Err.Raise vbObjectError + 513

lReturn = ClosePrinter(m_lhPrinter)
If lReturn = 0 Then Err.Raise vbObjectError + 513

Set my_Printer = Nothing



Exit Sub

ErrLine:
    MsgBox "ClosePrinterConnection:" & vbCrLf & Err.Description
    

End Sub


Public Function OpenPrinterConnection() As Boolean
'Declare the variables
Dim lReturn As Long
Dim MyDocInfo As DOCINFO
    
On Error GoTo ErrLine

lReturn = OpenPrinter(my_Printer.DeviceName, m_lhPrinter, 0)
If lReturn = 0 Then Err.Raise vbObjectError + 513
    
MyDocInfo.pDocName = "Invoice Printing"
MyDocInfo.pOutputFile = vbNullString
MyDocInfo.pDatatype = vbNullString

lReturn = StartDocPrinter(m_lhPrinter, 1, MyDocInfo)
If lReturn = 0 Then Err.Raise vbObjectError + 513

lReturn = StartPagePrinter(m_lhPrinter)
If lReturn = 0 Then Err.Raise vbObjectError + 513

OpenPrinterConnection = True

Exit Function

ErrLine:
    MsgBox "PRINTER ERROR" & vbCrLf & Err.Description


End Function

Private Sub PrintColHeading(strData() As String)
Dim ItemCount As Integer
Dim PrintData As String


If UBound(strData) = 0 Then Exit Sub


For ItemCount = LBound(strData) To UBound(strData)
    
    PrintData = PrintData & " " & strData(ItemCount)
    
Next ItemCount


End Sub


Private Function PrintData(strData As Variant) As Long
Dim PrintBytes As Long

If strData Is Nothing Then Exit Function

PrintData = WritePrinter(m_lhPrinter, strData, Len(strData), PrintBytes)

End Function


Private Sub PrintLine(ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single)

If my_Printer Is Nothing Then Exit Sub

my_Printer.Line (X1, Y1)-(X2, Y2)

my_Printer.CurrentX = X2
my_Printer.CurrentY = Y2

End Sub

Private Sub PrintRectangle(ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single)

If my_Printer Is Nothing Then Exit Sub

my_Printer.Line (X1, Y1)-(X2, Y2), , B

my_Printer.CurrentX = X2
my_Printer.CurrentY = Y2

End Sub


Private Sub PrintRow()
Dim I As Integer
Dim curRow As Single
Dim curX As Single
Dim strData As String
Dim MaxCount As Integer
Dim lKeys

Dim MaterialClass As clsMaterial

On Error GoTo ErrLine

' Save the current row.
curRow = my_Printer.CurrentY
curX = FIELD_MARGIN
If m_DataSource Is Nothing Then Exit Sub

'Get the data from collection
MaxCount = m_DataSource.count

lKeys = m_DataSource.GetKeys

For I = 0 To MaxCount - 1
    Set MaterialClass = m_DataSource.Item(lKeys(I))
        
Next I



Exit Sub

ErrLine:

    MsgBox "Printrow: " & vbCrLf & Err.Description, vbCritical '
    

End Sub

Public Sub Test()

If m_FileSysObject Is Nothing Then Set m_FileSysObject = New FileSystemObject

Set m_TextStream = m_FileSysObject.CreateTextFile(App.Path & "\PharmPrint.txt", True, False)


With m_TextStream
    .WriteLine ("                        M/s R.Y.Shettar And Company")
    .WriteLine ("                        Veera Narayana Temple Road Gadag")
    .WriteBlankLines (2)
    .WriteLine ("-----------------------------------------------------------")
    .WriteBlankLines (1)
    .WriteLine ("                        INVOICE                             ")
    .WriteBlankLines (1)
    .WriteLine ("-----------------------------------------------------------")
    
    .write ("SlNo Mfr  Particulars                   Batch ExpDate Qty  Bonus TRP SalePrice MRP KST Amount")
    
    .Close
End With


'Call Shell(App.Path & "\PrintBatch.bat", vbMinimizedNoFocus)
End Sub

