VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ColStruct
    ColLength As Integer
    ColText As String
    COlHeader As String
End Type

Private Col() As ColStruct
Private m_Cols As Integer
Private m_Header As String
Private NewPageStarted As Boolean

Public Property Get ObjectName()
    ObjectName = "TransPrintClass"
End Property

Public Property Let COlHeader(Index As Integer, Hdr As String)
    Col(Index).COlHeader = Hdr
End Property

Public Property Get COlHeader(Index As Integer) As String
    COlHeader = Col(Index).COlHeader
End Property

Public Property Get Cols() As Integer
    Cols = m_Cols
End Property
Public Property Let Cols(NewValue As Integer)
    If NewValue < 0 Then Err.Raise 10001, "Invalid Cols value specified", "Print Class"
    m_Cols = NewValue
    ReDim Preserve Col(m_Cols)
End Property

Public Property Get isNewPage() As Boolean
'NewPageStarted = True
    isNewPage = NewPageStarted
End Property

Public Property Let isNewPage(newPage As Boolean)
    NewPageStarted = newPage
End Property

Public Property Get ColText(Index As Integer) As String
    ColText = Col(Index).ColText
End Property
Public Property Let ColText(Index As Integer, ColText As String)
    On Error GoTo ErrLine
    Col(Index).ColText = ColText
    Exit Property
ErrLine:
    Err.Raise 10001, "Invalid property assignment", "Print Class"
End Property

Public Property Get ColWidth(Index As Integer) As Integer
    ColWidth = Col(Index).ColLength
End Property

Public Property Let ColWidth(Index As Integer, Width As Integer)
    On Error GoTo ErrLine
    Col(Index).ColLength = Width
    Exit Property
ErrLine:
    Err.Raise 10001, "Invalid property assignment", "Print Class"
End Property
Public Property Let Header(NewValue As String)
m_Header = NewValue
End Property

'Public Property Get Header() As String
'Header = m_Header
'End Property

Public Sub newPage()
    On Error Resume Next
    Printer.newPage
    NewPageStarted = True
End Sub
Public Function NewLines(n As Integer)
    Dim count As Integer
    On Error Resume Next
    For count = 1 To n
        Printer.Print ""
    Next count
End Function

Public Function PrintText()
Dim TmpArr() As String
Dim strToPrint As String
Dim Flag As Boolean
Dim count As Integer
Dim TmpStr As String
'Printer.Font = "Courier"

If NewPageStarted Then
    On Error Resume Next
    NewPageStarted = False
End If


'You have a structure called col(count)
    ReDim TmpArr(m_Cols)
    Printer.Font = "Courier New"
    Printer.Font.Size = 9
    Printer.FontBold = False
     For count = 0 To m_Cols
        TmpArr(count) = Col(count).ColText
     Next count
     
     With Me
        .ColWidth(0) = 15
        .ColWidth(1) = 8
        .ColWidth(2) = 22
        .ColWidth(3) = 13
        .ColWidth(4) = 14
        .ColWidth(5) = 15
    End With
    
    Flag = True
    While Flag
        Flag = False
        strToPrint = ""
        For count = 0 To m_Cols
            If (count <= 1) Then
                TmpStr = TmpArr(count) & Space(Col(count).ColLength)
            Else
                TmpStr = Space(Col(count).ColLength - Len(TmpArr(count))) & TmpArr(count)
            End If
            strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
            TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
            If Len(TmpArr(count)) > 0 Then
                Flag = True
            End If
        Next count
        
        On Error Resume Next
        Printer.Print strToPrint
        
        'frmMain.txt.Text = frmMain.txt.Text & StrToPrint
    Wend
    Err.Clear
End Function
' Prints one txn row of data
Public Sub PrintRow()
    
    Dim strToPrint, TmpStr As String
    Dim count As Integer
    Dim TmpArr() As String
    ReDim TmpArr(m_Cols)
    
    Printer.Font = "Courier New"
    Printer.Font.Size = 9
    
    strToPrint = ""
    For count = 0 To m_Cols
        If (count <= 1) Then
            'tmpStr = TmpArr(count) & Space(Col(count).ColLength)
            TmpStr = Col(count).ColText & Space(Col(count).ColLength)
        Else
            'tmpStr = Space(Col(count).ColLength - Len(TmpArr(count))) & TmpArr(count)
            TmpStr = Space(Col(count).ColLength - IIf(Len(Col(count).ColText) < Col(count).ColLength, Len(Col(count).ColText), Col(count).ColLength)) & Col(count).ColText
        End If
    'strToPrint = strToPrint & Left(TmpStr, Col(Count).ColLength) & " | "
        'TmpArr(Count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(Count).ColLength))
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & "  "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count
    
    On Error Resume Next
    Printer.Print strToPrint

End Sub
' Prints one txn row of data
Public Sub PrintRow2()
    
    Dim strToPrint, TmpStr As String
    Dim count As Integer
    Dim TmpArr() As String
    Dim spaceCount As Integer
    ReDim TmpArr(m_Cols)
    
    Printer.Font = "Courier New"
    Printer.Font.Size = 9
    
    strToPrint = ""
    'Printer.CurrentX = 220
    For count = 0 To m_Cols
            If (count = 1) Then
            'tmpStr = TmpArr(count) & Space(Col(count).ColLength)
            TmpStr = Col(count).ColText & Space(Col(count).ColLength)
            TmpStr = Col(count).ColText
            Dim L As Integer
            L = Col(count).ColLength - Len(Col(count).ColText)
            If (L > 0) Then
                TmpStr = TmpStr & Space(L)
            End If
        Else
            'TmpStr = Space(Col(Count).ColLength - Len(TmpArr(Count))) & TmpArr(Count)
            spaceCount = Col(count).ColLength - Len(Col(count).ColText)
            If spaceCount < 0 Then spaceCount = 0
            TmpStr = Space(spaceCount) & Col(count).ColText
        End If
        'strToPrint = strToPrint & Left(tmpStr, Col(count).ColLength) & " | "
        'TmpArr(count) = Trim$(Right(tmpStr, Len(tmpStr) - Col(count).ColLength))
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & "  "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count
    
    On Error Resume Next
    Printer.Print strToPrint
       Printer.Font.Size = 11
    Printer.Print ""

End Sub
' Prints one txn row of data
Public Sub PrintRows()
    
    Dim strToPrint, TmpStr As String
    Dim count As Integer
    Dim TmpArr() As String
    ReDim TmpArr(m_Cols)
    
    Printer.Font = "Courier New"
    Printer.Font.Size = 9
    Printer.CurrentX = 160
    strToPrint = ""
    For count = 0 To m_Cols
        If (count <= 0) Then
            'tmpStr = TmpArr(count) & Space(Col(count).ColLength)
            TmpStr = Col(count).ColText & Space(Col(count).ColLength)
        Else
            'TmpStr = Space(Col(Count).ColLength - Len(TmpArr(Count))) & TmpArr(Count)
        TmpStr = Space(Col(count).ColLength - Len(Col(count).ColText)) & Col(count).ColText & ".00 "
       ' Debug.Print TmpStr
        End If
        'strToPrint = strToPrint & Left(tmpStr, Col(count).ColLength) & " | "
        'TmpArr(count) = Trim$(Right(tmpStr, Len(tmpStr) - Col(count).ColLength))
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & "  "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
        Debug.Print TmpArr(count)
        Debug.Print "this is expected" & strToPrint
    Next count
    
    On Error Resume Next
    Printer.Print ""
    Printer.Print strToPrint
    Printer.Font.Size = 6
     Printer.Print ""

End Sub
' Prints the header for passbook transactions.
Public Sub printHeader(Optional ModuleId As wisModules)

Dim YLocation As Integer
Dim count As Integer

    'Print column headers
    ' i.e.
    ' Print a line
    ' print column headers
    ' print another line (closing line)
    '
    With Me
        
        ' set printer column widths for header.
        .ColWidth(0) = IIf(gLangOffSet, 32, 20)
        .COlHeader(0) = GetResourceString(37) 'Date
        .ColWidth(1) = IIf(gLangOffSet, 50, 34)
        .COlHeader(1) = GetResourceString(39) 'Particulars/ShareValue
        .ColWidth(2) = IIf(gLangOffSet, 23, 16)
        .COlHeader(2) = GetResourceString(275) 'Cheque
        .ColWidth(3) = IIf(gLangOffSet, 32, 21)
        .COlHeader(3) = GetResourceString(276) 'Debit
        .ColWidth(4) = IIf(gLangOffSet, 32, 24)
        .COlHeader(4) = GetResourceString(277) 'Credit
        .ColWidth(5) = IIf(gLangOffSet, 32, 23)
        .COlHeader(5) = GetResourceString(42) 'Balance
        If ModuleId = wis_Members Then
            .COlHeader(0) = GetResourceString(37) '"Date"
            .COlHeader(1) = GetResourceString(53, 212) '"Share(Face) value"
            .COlHeader(2) = GetResourceString(334) '"Sold"
            .COlHeader(3) = GetResourceString(53, 40) '"Share Amount"
            .COlHeader(4) = GetResourceString(335) '"Quantity"
            .COlHeader(5) = GetResourceString(53, 42) ''"Balance"
        End If
        
    End With
    
    ReDim TmpArr(m_Cols)
    
    For count = 0 To m_Cols
        TmpArr(count) = Col(count).COlHeader
    Next count

    ' accumulate col headers text.
    Dim strToPrint, TmpStr As String
    strToPrint = ""
    For count = 0 To m_Cols
        TmpStr = TmpArr(count) & Space(Col(count).ColLength)
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count
    strToPrint = Left(strToPrint, Len(strToPrint) - 3)

    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    
    Printer.Print ""
    Printer.Print ""
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    Printer.CurrentX = 0
        
    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    ''COMMENTING BELOW LINE, AS PRINTING HAPPENING BEWLO AGAIN WITH COORDINATES
    'Printer.Print strToPrint
    
    '
    YLocation = Printer.CurrentY
    Printer.CurrentX = 10: Printer.Print Col(0).COlHeader 'date
    Printer.CurrentX = 1750: Printer.CurrentY = YLocation: Printer.Print "|   " + Col(1).COlHeader 'parti
    Printer.CurrentX = 4450: Printer.CurrentY = YLocation: Printer.Print "| " + Col(2).COlHeader 'Check
    Printer.CurrentX = 5650: Printer.CurrentY = YLocation: Printer.Print "|   " + Col(3).COlHeader 'Credi
    Printer.CurrentX = 7350: Printer.CurrentY = YLocation: Printer.Print "|   " + Col(4).COlHeader 'debit
    Printer.CurrentX = 9200: Printer.CurrentY = YLocation: Printer.Print "|   " + Col(5).COlHeader 'Balance
    Printer.CurrentX = 11200: Printer.CurrentY = YLocation: Printer.Print "|"
       
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    YLocation = Printer.CurrentY
    Printer.CurrentX = 0
    Printer.CurrentY = YLocation
    '// After printing header part, reset column widths
    '// to match the widths required for data rows printing.
    With Me
    .ColWidth(0) = 15
    .ColWidth(1) = 22
    .ColWidth(2) = 8
    .ColWidth(3) = 13
    .ColWidth(4) = 14
    .ColWidth(5) = 15
    End With
End Sub
' Prints the header for passbook transactions.
Public Sub printHeader1()

   '
    'Print column headers
    ' i.e.
    ' Print a line
    ' print column headers
    ' print another line (closing line)
    '
    With Me
        ' set printer column widths for header.
        .ColWidth(0) = 37: .COlHeader(0) = GetResourceString(37) 'Date
        '.ColWidth(1) = 20: .COlHeader(1) = GetResourceString(275) 'Cheque
        .ColWidth(1) = 55: .COlHeader(1) = GetResourceString(39) 'Particulars
        .ColWidth(2) = 35: .COlHeader(2) = GetResourceString(276)  'Debit
        .ColWidth(3) = 35: .COlHeader(3) = GetResourceString(277) 'Credit
        .ColWidth(4) = 40: .COlHeader(4) = GetResourceString(42) 'Balance
    End With
    
    ReDim TmpArr(m_Cols)
    Dim count As Integer
    For count = 0 To m_Cols
        TmpArr(count) = Col(count).COlHeader
    Next count

    ' accumulate col headers text.
    Dim strToPrint, TmpStr As String
    strToPrint = ""
    For count = 0 To m_Cols
        TmpStr = TmpArr(count) & Space(Col(count).ColLength)
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count

    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    

     Printer.Print ""
     Printer.Print ""
    'code for printing a line.
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    Printer.Print strToPrint
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    
    '// After printing header part, reset column widths
    '// to match the widths required for data rows printing.
    With Me
    .ColWidth(0) = 17
    .ColWidth(1) = 24
    .ColWidth(2) = 16
    .ColWidth(3) = 16
    .ColWidth(4) = 17
    '.ColWidth(5) = 15
    End With
End Sub
' Prints the header for passbook transactions.
Public Sub printHeader2()

   '
    'Print column headers
    ' i.e.
    ' Print a line
    ' print column headers
    ' print another line (closing line)
    '
    With Me
        ' set printer column widths for header.
        .ColWidth(0) = 32: .COlHeader(0) = GetResourceString(37) 'Date
        '.ColWidth(1) = 20: .COlHeader(1) = GetResourceString(275) 'Cheque
        .ColWidth(1) = 30: .COlHeader(1) = GetResourceString(39) 'particulars
        .ColWidth(2) = 40: .COlHeader(2) = GetResourceString(275) 'cheque
        .ColWidth(3) = 30: .COlHeader(3) = GetResourceString(276) 'Debit
        .ColWidth(4) = 30: .COlHeader(4) = GetResourceString(277) 'Credit
        .ColWidth(5) = 35: .COlHeader(5) = GetResourceString(42) 'Balance
    End With
    
    ReDim TmpArr(m_Cols)
    Dim count As Integer
    For count = 0 To m_Cols
        TmpArr(count) = Col(count).COlHeader
    Next count

    ' accumulate col headers text.
    Dim strToPrint, TmpStr As String
    strToPrint = ""
    For count = 0 To m_Cols
        TmpStr = TmpArr(count) & Space(Col(count).ColLength)
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count

    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    

     Printer.Print ""
     Printer.Print ""
    'code for printing a line.
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    Printer.Print strToPrint
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    
    '// After printing header part, reset column widths
    '// to match the widths required for data rows printing.
    With Me
 .ColWidth(0) = 15
    .ColWidth(1) = 13
    .ColWidth(2) = 17
    .ColWidth(3) = 13
    .ColWidth(4) = 14
    .ColWidth(5) = 15
    End With
End Sub
' Prints the header for passbook transactions.
Public Sub printHeader3()

   '
    'Print column headers
    ' i.e.
    ' Print a line
    ' print column headers
    ' print another line (closing line)
    '
    With Me
        ' set printer column widths for header.
        .ColWidth(0) = 32: .COlHeader(0) = GetResourceString(37) 'Date
        .ColWidth(1) = 20: .COlHeader(1) = GetResourceString(39) 'particulars
        .ColWidth(2) = 30: .COlHeader(2) = GetResourceString(276) 'Debit
        .ColWidth(3) = 40: .COlHeader(3) = GetResourceString(277) 'credit
       ' .ColWidth(4) = 40: .COlHeader(4) = GetResourceString(277) 'credit
        .ColWidth(4) = 35: .COlHeader(4) = GetResourceString(42) 'Balance
    End With
    
    ReDim TmpArr(m_Cols)
    Dim count As Integer
    For count = 0 To m_Cols
        TmpArr(count) = Col(count).COlHeader
    Next count

    ' accumulate col headers text.
    Dim strToPrint, TmpStr As String
    strToPrint = ""
    For count = 0 To m_Cols
        TmpStr = TmpArr(count) & Space(Col(count).ColLength)
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count

    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    

     Printer.Print ""
     Printer.Print ""
    'code for printing a line.
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    Printer.Print strToPrint
    'Printer.Print " _____________________________________________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    
    '// After printing header part, reset column widths
    '// to match the widths required for data rows printing.
    With Me
    .ColWidth(0) = 17
    .ColWidth(1) = 24
    .ColWidth(2) = 15
    .ColWidth(3) = 17
    .ColWidth(4) = 17
    End With
End Sub
' Prints the header for passbook transactions.
Public Sub printHead()

   '
    'Print column headers
    ' i.e.
    ' Print a line
    ' print column headers
    ' print another line (closing line)
    '
    With Me
        ' set printer column widths for header.
         '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    .ColWidth(0) = 13: .COlHeader(0) = GetResourceString(37) 'Date
    .ColWidth(1) = 11: .COlHeader(1) = GetResourceString(488) 'Particulars
    .ColWidth(2) = 15: .COlHeader(2) = GetResourceString(310) 'Debit
    .ColWidth(3) = 15: .COlHeader(3) = GetResourceString(344) 'Credit
    .ColWidth(4) = 13: .COlHeader(4) = GetResourceString(345) 'Interest
    .ColWidth(5) = 15: .COlHeader(5) = GetResourceString(285) 'Interest
    .ColWidth(6) = 14: .COlHeader(6) = GetResourceString(42) 'Balance
    End With
    
    ReDim TmpArr(m_Cols)
    Dim count As Integer
    For count = 0 To m_Cols
        TmpArr(count) = Col(count).COlHeader
    Next count

    ' accumulate col headers text.
    Dim strToPrint, TmpStr As String
    strToPrint = ""
    For count = 0 To m_Cols
        TmpStr = TmpArr(count) & Space(Col(count).ColLength)
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count

    Printer.Font = gFontName
    Printer.Font.Size = 12  'gFontSize
    

     Printer.Print ""
     Printer.Print ""
    'code for printing a line.
    'Printer.Print " _______________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    Printer.Print strToPrint
    'Printer.Print " _______________________________________________________________________________________________________________"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.ScaleWidth, Printer.CurrentY)
    'Printer.Print ""
    Printer.CurrentX = 0
    
    '// After printing header part, reset column widths
    '// to match the widths required for data rows printing.
    With Me
         .ColWidth(0) = 8
        .ColWidth(1) = 8
        .ColWidth(2) = 8
        .ColWidth(3) = 8
        .ColWidth(4) = 8
        .ColWidth(5) = 8
        .ColWidth(6) = 6
    End With
End Sub
' Prints one txn row of data
Public Sub printRow1()
    
    Dim strToPrint, TmpStr As String
    Dim count As Integer
    Dim TmpArr() As String
    ReDim TmpArr(m_Cols)
    
    Printer.Font = "Courier New"
    Printer.Font.Size = 9
    
    strToPrint = ""
    For count = 0 To m_Cols
        If (count <= 1) Then
            'tmpStr = TmpArr(count) & Space(Col(count).ColLength)
            TmpStr = Col(count).ColText & Space(Col(count).ColLength)
        Else
            'tmpStr = Space(Col(count).ColLength - Len(TmpArr(count))) & TmpArr(count)
            TmpStr = Space(Col(count).ColLength - Len(Col(count).ColText)) & Col(count).ColText
        End If
        'strToPrint = strToPrint & Left(tmpStr, Col(count).ColLength) & " | "
        'TmpArr(count) = Trim$(Right(tmpStr, Len(tmpStr) - Col(count).ColLength))
        strToPrint = strToPrint & Left(TmpStr, Col(count).ColLength) & " | "
        TmpArr(count) = Trim$(Right(TmpStr, Len(TmpStr) - Col(count).ColLength))
    Next count
    
    On Error Resume Next
    Printer.Print strToPrint

End Sub

Private Sub Class_Initialize()
    'NewPageStarted = True
End Sub


'
' to Print the first page information of the passbook.
Public Sub printPassbookPage(AccType As wisModules, AccId As Long)
Dim strToPrint
Dim metaRst As ADODB.Recordset
'--> Find the customer id  given the account id and account type.
'

' Find the linking table for query join, based on accType
Dim TableName As String
Dim accountString  As String
Printer.Font = gFontName
Printer.Font.Size = 16
Select Case AccType
    Case wis_SBAcc
        TableName = "SBMaster"
        accountString = GetResourceString(421)    ' Savings account string
    Case wis_CAAcc
        TableName = "CAMaster"
        accountString = GetResourceString(422)    ' current account string
    Case wis_PDAcc
        TableName = "PDMaster"
        accountString = GetResourceString(425)    ' pigmy account string
    Case wis_BKCC
        TableName = "BKCCMaster"
        accountString = GetResourceString(229)    ' BKCC account string
    Case wis_RDAcc
        TableName = "RDMaster"
        accountString = GetResourceString(424)    ' RD account string
    Case wis_Members
        TableName = "MemMaster"
        accountString = GetResourceString(49)    '  account string
    Case wis_Loans
        TableName = "LoanMaster" '"???"
        accountString = GetResourceString(441)    '  account string
End Select
'select the required fields from table nametab by accid from type of table
Dim SqlStr As String
SqlStr = "SELECT A.Title," & _
            "A.FirstName, " & _
            "A.MiddleName, " & _
            "A.LastName, " & _
            "A.HomeAddress, " & _
            "A.OfficeAddress, " & _
            "A.HomePhone, " & _
            "A.eMail, " & _
            "A.Place, " & _
            "B.AccNum " & _
            " FROM NameTab A, " & _
            TableName & " B " & _
            "WHERE A.CustomerId = B.CustomerId " & _
            "AND B.AccId = " & AccId

' Execute this statmtto get customer details
Dim rs As ADODB.Recordset


gDbTrans.SqlStmt = SqlStr
'rs = gDbTrans.SQLExecute
Call gDbTrans.Fetch(rs, adOpenDynamic)
    
' TODO: print the details of the retrieved person record
' on the passbook page
'Before Printing Move half page
Dim I As Integer
Dim MaxI As Integer
MaxI = IIf(gLangOffSet = wis_NoLangOffset, (wis_ROWS_PER_PAGE / 2 - 3), (wis_ROWS_PER_PAGE / 2 - 3))

' Print the customer's name
Printer.CurrentY = 6000
Printer.CurrentX = 1005


'Printer.Print "Name: " & (rs("Title") & rs("FirstName") & rs("MiddleName") & rs("LastName")) & Space(50) & accountString & ": " & AccID
Dim Str As String, strTmp As String
Str = GetResourceString(35) & ": " & FormatField(rs("Title"))
strTmp = Trim$(FormatField(rs("FirstName")))
If Len(strTmp) > 0 Then Str = Str + " " + strTmp
strTmp = Trim$(FormatField(rs("MiddleName")))
If Len(strTmp) > 0 Then Str = Str + " " + strTmp
strTmp = Trim$(FormatField(rs("LastName")))
If Len(strTmp) > 0 Then Str = Str + " " + strTmp

If Printer.TextWidth(Str & accountString & ": " & FormatField(rs("AccNum"))) < (0.9 * Printer.Width) Then
    Printer.Print Str & Space(40) & accountString & ": " & FormatField(rs("AccNum"))
Else
    Printer.Print Str
    Printer.CurrentX = 1005
    Printer.Print accountString & ": " & FormatField(rs("AccNum"))
End If
' Print account id details, depending on account type

'Printer.Print accountString & ": " & AccID

Printer.CurrentX = 1005
Dim Address As String
Dim strArr() As String

strArr = Split(FormatField(rs("HomeAddress")), ";")
For I = 0 To UBound(strArr)
    If Len(Trim$(strArr(I))) > 0 Then Address = Address + " " + Trim$(strArr(I))
Next
If Len(Trim$(Address)) < 0 Then
    strArr = Split(FormatField(rs("OfficeAddress")), ";")
    For I = 0 To UBound(strArr)
        If Len(Trim$(strArr(I))) > 0 Then Address = Address + " " + Trim$(strArr(I))
    Next
End If

Printer.Print GetResourceString(130); ":" & Address

Printer.CurrentX = 1005
If (Len(Trim(FormatField(rs("HomePhone")))) <> 0) Then Printer.Print (rs("HomePhone"))

Printer.CurrentX = 1005
If (Len(Trim(FormatField(rs("Email")))) > 0) Then Printer.Print (rs("Email"))

Printer.CurrentX = 1005
If (Len(Trim(FormatField(rs("Place")))) > 0) Then Printer.Print (rs("Place"))

Printer.EndDoc

End Sub
' to Print the FDpage Certificate information. 'Mrudu june 7th 2014.
Public Sub printFDPage(AccId As Integer)
Dim strToPrint
Dim metaRst As ADODB.Recordset
Dim AccType As Long
'Dim AccID As Long
'--> Find the customer id  given the account id and account type.
'

' Find the linking table for query join, based on accType
Dim TableName As String
Dim accountString  As String
Dim ClosedDate As String
Dim MaturedOn As String
Dim MaturityAmount As String
Printer.Font = gFontName
Printer.Font.Size = 16
AccType = wis_Deposits
TableName = "FDMaster"
accountString = GetResourceString(423)
'select the required fields from table nametab by accid from type of table
Dim SqlStr As String
SqlStr = "SELECT A.Title," & _
             "A.FirstName," & _
             "A.MiddleName," & _
             " A.LastName, " & _
             "A.HomeAddress," & _
            "A.OfficeAddress," & _
             "A.Place," & _
            "B.AccNum," & _
            "B.CertificateNo," & _
            "B.RateOfInterest," & _
            "B.CreateDate ," & _
            "B.EffectiveDate," & _
            "B.DepositAmount," & _
            "B.MaturityDate," & _
            "B.MaturedOn," & _
            "B.MaturityAmount, " & _
            "B.ClosedDate" & _
  " FROM NameTab as A, FDMaster as B" & _
 " WHERE A.CustomerId = B.CustomerId " & _
            "AND B.AccId = " & AccId

    ' Execute this statmtto get customer details
Dim rs As ADODB.Recordset


gDbTrans.SqlStmt = SqlStr
'rs = gDbTrans.SQLExecute
Call gDbTrans.Fetch(rs, adOpenDynamic)
    
' TODO: print the details of the retrieved person record
' on the passbook page
Const START_ROW = 4
Dim I As Integer
For I = 0 To START_ROW
    Printer.Print ""
Next I

' Print the customer's name
Printer.CurrentX = 9400
Printer.Print (rs("CertificateNo"))


Printer.CurrentX = 2300


'Printer.Print (rs("AccNum") & Space(50) & rs("DepositAmount"))
 Printer.Print (rs("AccNum") & Space(58) & rs("DepositAmount") & Space(36) & rs("CreateDate"))

Printer.CurrentX = 2300

Printer.Print (rs("Title") & rs("FirstName") & rs("MiddleName") & rs("LastName")) & Space(70) & rs("EffectiveDate")
 Printer.CurrentX = 2300
 
 Dim Address As String
Address = IIf(LenB(rs("HomeAddress")) <> 0, rs("HomeAddress"), rs("OfficeAddress"))
Printer.Print Address & Space(90) & (rs("DepositAmount"))
 
 'Printer.Print (rs("DepositAmount"))
Printer.CurrentX = 9400
 Printer.Print (rs("MaturityDate"))
 Printer.CurrentX = 9400
 'rs("Days") = IIf(IsNull(rs("Days")), 0, rs("Days"))
 MaturedOn = IIf(IsNull(rs("MaturedOn")), " ", rs("MaturedOn"))
 Printer.Print MaturedOn
 Printer.CurrentX = 9400
 MaturityAmount = IIf(IsNull(rs("MaturityAmount")), " ", rs("MaturityAmount"))
 Printer.Print MaturityAmount
 Printer.CurrentX = 9400
 Printer.Print (rs("RateOfInterest"))
 Printer.CurrentX = 9400
 Printer.Print ("")
 Printer.CurrentX = 9400
 ClosedDate = IIf(IsNull(rs("ClosedDate")), " ", rs("ClosedDate"))
 Printer.Print ClosedDate
End Sub
' to Print the first page information of the passbook.' Mrudu july 17th 2014.
Public Sub printPassbookPage1(AccType As wisModules, LoanID As Long)
Dim strToPrint
Dim metaRst As ADODB.Recordset
Dim TableName As String
Dim accountString  As String
Printer.Font = gFontName
Printer.Font.Size = 16
AccType = wis_Deposits
TableName = "LoanMaster"
accountString = GetResourceString(423)
'select the required fields from table nametab by accid from type of table
Dim SqlStr As String
SqlStr = "SELECT A.Title," & _
             "A.FirstName," & _
             "A.MiddleName," & _
             " A.LastName, " & _
             "A.HomeAddress," & _
            "A.OfficeAddress," & _
             "A.Place," & _
             "A.Profession," & _
            "B.LoanId," & _
            "B.MemID," & _
            "B.IntRate" & _
  " FROM NameTab as A, " & _
  "LoanMaster as B" & _
 " WHERE A.CustomerId = B.CustomerId " & _
            "AND B.LoanId = " & LoanID

Dim rs As ADODB.Recordset


gDbTrans.SqlStmt = SqlStr
'rs = gDbTrans.SQLExecute
Call gDbTrans.Fetch(rs, adOpenDynamic)
    
' TODO: print the details of the retrieved person record
' on the passbook page
Const START_ROW = 3
Dim I As Integer
For I = 0 To START_ROW
    Printer.Print ""
Next I
Dim Address As String
' Print the customer's name
Printer.CurrentX = 5800
If (Len(Trim(rs("Title") & rs("FirstName") & rs("MiddleName") & rs("LastName"))) > 40) Then
Printer.Print "name :" & (rs("Title") & rs("FirstName") & rs("MiddleName") & rs("LastName"))
Printer.Print ""
'Debug.Print Name

Else
Printer.Print (rs("Title") & rs("FirstName"))
Printer.CurrentX = 5400
Printer.Print (rs("MiddleName") & rs("LastName"))
End If

Printer.CurrentX = 6000
Printer.Print (rs("Profession"))

Printer.CurrentX = 6000
Address = IIf(LenB(rs("HomeAddress")) <> 0, rs("HomeAddress"), rs("OfficeAddress"))
Printer.Print Address
Printer.Print ""

Printer.CurrentX = 6300
Printer.Print (rs("MemID"))

End Sub




