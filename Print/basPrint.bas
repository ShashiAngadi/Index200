Attribute VB_Name = "basPrint"
Public Sub PrintBetweenDates(ModuleId As wisModules, AccId As Long, StartIndiandate As String, EndIndianDate As String)

Dim masterTable As String
Dim transTable As String
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintRow As Integer
Const HEADER_ROWS = 3
Dim curPrintRow As Integer
Dim YLocation As Integer
Dim BeginY As Integer
Dim AccTypeName As String
Dim checkField As String

checkField = "ChequeNo"
If ModuleId = wis_SBAcc Then
    masterTable = "SBMASTER"
    transTable = "SBTRANS"
    AccTypeName = GetResourceString(421, 60)
ElseIf ModuleId = wis_CAAcc Then
    masterTable = "CAMASTER"
    transTable = "CATRANS"
    AccTypeName = GetResourceString(422, 60)
ElseIf ModuleId = wis_PDAcc Then
    masterTable = "PDMASTER"
    transTable = "PDTRANS"
    checkField = " VoucherNo as ChequeNo"
    AccTypeName = GetResourceString(425, 36, 60)
ElseIf ModuleId = wis_RDAcc Then
    masterTable = "RDMASTER"
    transTable = "RDTRANS"
    checkField = " VoucherNo as ChequeNo"
    AccTypeName = GetResourceString(424, 36, 60)
ElseIf ModuleId = wis_Members Then
    masterTable = "MemMASTER"
    transTable = "MemTRANS"
    'checkField = " VoucherNo as ChequeNo"
    AccTypeName = GetResourceString(53, 36, 60)
End If

'1.select the details of theaccount holder
SqlStr = "SELECT A.Name,B.AccNum FROM QryName A, " & _
             masterTable & " B " & _
            " WHERE A.CustomerId = B.CustomerId " & _
            " AND B.AccId = " & AccId
            
            
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint

'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From " & transTable & " WHERE AccId = " & AccId
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
    If (rst(0) = 0) Then
        MsgBox "There are no transactions available for printing."
        Exit Sub
    End If

''As A4 Printinint Start with New Page
clsPrint.isNewPage = True

'3. Getting matching records for passbook printing
    SqlStr = "SELECT AccId,TransDate,TransType,Amount,Balance,Particulars," & checkField & _
        " From " & transTable & " WHERE AccId = " & AccId & _
        " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
        " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#"
If ModuleId = wis_Members Then
     SqlStr = " Select Count(*) as ShareCount, A.AccId,A.TransID,TransDate,FaceValue as Particulars,FaceValue* Count(*) as Amount, Balance, " & wDeposit & " as TransType " & _
        " from MemTrans A Left join ShareTrans B " & _
        " On  A.AccID=B.AccID and A.TransID=B.SaleTransID " & _
        " where A.AccID= " & AccId & _
        " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
        " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
        " And (A.TransType= " & wDeposit & " or A.TransType= " & wContraDeposit & ")" & _
        " Group By A.AccID,TransID,TransDate,FaceValue,Balance"
    SqlStr = SqlStr & " UNION " & _
        " Select Count(*) as ShareCunt, A.AccId,A.TransID,TransDate,FaceValue as Particulars,FaceValue*Count(8) as Amount, Balance, " & wWithdraw & " as TransType " & _
        " from MemTrans A Left join ShareTrans B " & _
        " On  A.AccID=B.AccID and A.TransID=B.ReturnTransID " & _
        " where A.AccID = " & AccId & _
        " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
        " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
        " And (A.TransType= " & wWithdraw & " or A.TransType= " & wContraWithdraw & ")" & _
        " Group By A.AccID,TransID,TransDate,FaceValue,Balance"
End If
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If

'Printer.PaperSize = 9
'Printer.Font.Name = gFontName
'Printer.Font.Size = 12 'gFontSize
Printer.Font = gFontName '"Courier New"
Printer.FONTSIZE = 12

With clsPrint
    '.Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    
    .Cols = 5
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .COlHeader(1) = GetResourceString(275) 'Cheque
    '.ColWidth(2) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(3) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(4) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(5) = 15: .COlHeader(5) = GetResourceString(42) 'Balance
    
    If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE_A4) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
    End If
    
    ' column widths for printing txn rows.
        .ColWidth(0) = 15
        .ColWidth(1) = 22
        .ColWidth(2) = 8
        .ColWidth(3) = 13
        .ColWidth(4) = 14
        .ColWidth(5) = 15
    BeginY = Printer.CurrentY
    While Not rst.EOF
        If .isNewPage Then
            ''Print the BANK Name
            Printer.Font.name = gFontName
            Printer.Font.Size = Printer.Font.Size + 2
            Printer.CurrentY = 1000
            Printer.CurrentX = 5000 - (Printer.TextWidth(gCompanyName) / 2)
            Printer.Font.Bold = True
            Printer.Print gCompanyName
            Printer.Font.Bold = False
            Printer.Font.Size = Printer.Font.Size - 2
            
            BeginY = Printer.CurrentY + 150
            Printer.CurrentY = BeginY
            Printer.Print (FormatField(metaRst("Name")))
            Printer.CurrentX = 9500 - Printer.TextWidth(AccTypeName)
            Printer.CurrentY = BeginY
            Printer.Print AccTypeName + ":" & FormatField(metaRst("AccNum"))
            
            .printHeader (ModuleId)
            
            BeginY = Printer.CurrentY
            .isNewPage = False
        End If
        .ColText(0) = FormatField(rst("TransDate"))
        .ColText(1) = FormatField(rst("Particulars"))
        .ColText(2) = FormatField(rst("ChequeNo"))
        If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
            .ColText(3) = FormatField(rst("Amount"))
            .ColText(4) = " "
        Else
            .ColText(4) = FormatField(rst("Amount"))
            .ColText(3) = " "
        End If
        .ColText(5) = FormatField(rst("Balance"))
        .PrintRow
        YLocation = Printer.CurrentY + 100
        ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE_A4) Then
            ' since we have to print now in a new page,
            ' we need to print the header.' So, set columns widths for header.
            
            Printer.Line (1800, BeginY)-(1800, YLocation)
            Printer.Line (4500, BeginY)-(4500, YLocation)
            Printer.Line (5700, BeginY)-(5650, YLocation)
            Printer.Line (7400, BeginY)-(7400, YLocation)
            Printer.Line (9250, BeginY)-(9250, YLocation)
            'Printer.Line (11200, 1500)-(11200, YLocation)
            
            Printer.CurrentX = 0
            .newPage
           
            curPrintRow = 1
        End If
        rst.MoveNext
    Wend
    '.newPage
End With

    Printer.Line (1800, BeginY)-(1800, YLocation)
    Printer.Line (4500, BeginY)-(4500, YLocation)
    Printer.Line (5700, BeginY)-(5700, YLocation)
    Printer.Line (7400, BeginY)-(7400, YLocation)
    Printer.Line (9250, BeginY)-(9250, YLocation)
    'Printer.Line (11200, 1500)-(11200, YLocation)
    
Printer.EndDoc

Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing

End Sub




