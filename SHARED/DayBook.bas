Attribute VB_Name = "basDayBook"
Option Explicit

Dim SQLStmt As String
Dim Rst As Recordset
Dim TransType As wisTransactionTypes
Dim Debit As Currency
Dim Credit As Currency
Dim HeadPrint As Boolean
Dim M_SlNo As Integer

Public Function MMAccTrans(AsOnDate As Date, webTable As HTMLTable, Optional ShowTotal As Boolean) As Boolean

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency
SQLStmt = "SELECT MMTrans.AccId,Amount,TransType,Title +' '+FirstName +' '+ " & _
        " MiddleName +' '+LastName as Name FROM MemTrans A,MemMaster B,NameTab C " & _
        " WHERE TransDate = #" & AsOnDate & "#" & _
        " And A.AccID = MMTrans.AccID And " & _
        " NameTab.CustomerId = MMMaster.CustomerID ORDER BY MMTrans.AccID"
gDbTrans.SQLStmt = SQLStmt

If gDbTrans.SQLFetch <= 0 Then GoTo CalculateMembershipFee

While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If Not HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = LoadResString(gLangOffSet + 409): .CellFontBold = True
                HeadPrint = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = FormatField(Rst("Title")) & " " & _
                                    FormatField(Rst("FirstName")) & " " & _
                                    FormatField(Rst("MiddleName")) & " " & _
                                    FormatField(Rst("LastName"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            ElseIf TransType = wWithDraw Then
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatField(Rst("Amount")): .CellAlignment = 7
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'***********************************
CalculateMembershipFee:
SQLStmt = "SELECT MMTrans.AccId,Amount,TransID,Title,FirstName, " & _
        " MiddleName,LastName FROM MMTrans,MMMaster,NameTab WHERE TransDate = " & _
        "#" & FormatDate(AsonIndianDate) & "#" & _
        " And TransID = 1 And TransType = " & wCharges & _
        " And MMMaster.AccID = MMTrans.AccID And " & _
        " NameTab.CustomerId = MMMaster.CustomerID ORDER BY MMTrans.AccID"
gDbTrans.SQLStmt = SQLStmt
If gDbTrans.SQLFetch <= 0 Then
    GoTo CalculateShareFee
End If

Set Rst = gDbTrans.Rst.Clone
HeadPrint = False

While Not Rst.EOF
TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If Not HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = LoadResString(gLangOffSet + 195): .CellFontBold = True
                HeadPrint = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Title")) & " " & _
                                    FormatField(Rst("FirstName")) & " " & _
                                    FormatField(Rst("MiddleName")) & " " & _
                                    FormatField(Rst("LastName"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'***********************************
CalculateShareFee:
SQLStmt = "SELECT MMTrans.AccId,Amount,TransID,Title,FirstName, " & _
        " MiddleName,LastName FROM MMTrans,MMMaster,NameTab WHERE TransDate = " & _
        "#" & FormatDate(AsonIndianDate) & "#" & _
        " And TransID <> 1 And TransType = " & wCharges & _
        " And MMMaster.AccID = MMTrans.AccID And " & _
        " NameTab.CustomerId = MMMaster.CustomerID ORDER BY MMTrans.AccID"
gDbTrans.SQLStmt = SQLStmt
If gDbTrans.SQLFetch <= 0 Then Exit Function

Set Rst = gDbTrans.Rst.Clone
HeadPrint = False

While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If Not HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = LoadResString(gLangOffSet + 198): .CellFontBold = True
                HeadPrint = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Title")) & " " & _
                                    FormatField(Rst("FirstName")) & " " & _
                                    FormatField(Rst("MiddleName")) & " " & _
                                    FormatField(Rst("LastName"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

ExitLine:
MMAccTrans = True

End Function
Public Function LoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

Dim AsOnDate As Date
Dim LoanHeadPrint As Boolean

AsOnDate = FormatDate(AsonIndianDate)
gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.transtype, a.amount, b.accid,C.LoanAccNo FROM loantrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & AsOnDate & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND (a.TransType = " & wDeposit & _
        " Or a.TransType = " & wWithDraw & ")"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY c.SchemeID,a.loanid, a.transid"
' Execute the query...
If gDbTrans.SQLFetch < 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
LoanHeadPrint = True
Dim SchemeName As String

While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 49) & " " & _
                    LoadResString(gLangOffSet + 215)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            ElseIf TransType = wWithDraw Then
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'****************************
CalculateInterest:
gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.amount, b.accid,C.LoanAccNo FROM loantrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & FormatDate(AsonIndianDate) & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.TransType = " & wCharges
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY a.loanid, a.transid"
' Execute the query...
If gDbTrans.SQLFetch < 0 Then GoTo LastLine

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 49) & " " & _
                    LoadResString(gLangOffSet + 236)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
LoanTrans = True

LastLine:

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
End Function

Public Function BKCCDepositTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency
Dim AsOnDate As Date

Dim SchemeName As String

AsOnDate = FormatDate(AsonIndianDate)

gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.transtype, a.amount, b.accid,LoanAccNo FROM BKCCtrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & AsOnDate & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND (a.TransType = " & _
            wBKCCDeposit & " Or a.TransType = " & wBKCCWithDraw & ")"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY a.loanid, a.transid"
' Execute the query...
If gDbTrans.SQLFetch < 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
Dim LoanHeadPrint As Boolean
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 229) & " " & _
                    LoadResString(gLangOffSet + 43)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst("TransType"))
            If TransType = wBKCCDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            ElseIf TransType = wBKCCWithDraw Then
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'****************************
CalculateInterest:
gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.amount, b.accid,LoanAccNo FROM bkcctrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & AsOnDate & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.TransType = " & wInterest & _
        " AND PARTICULARS = " & AddQuotes(LoadResString(gLangOffSet + 233), True)  'Interest on Deposit
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY a.loanid, a.transid"
' Execute the query...
If gDbTrans.SQLFetch < 0 Then GoTo LastLine

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
LoanHeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 229) & " " & _
                    LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 236)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
BKCCDepositTrans = True

LastLine:

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function

Public Function BKCCLoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency
Dim AsOnDate As Date

AsOnDate = FormatDate(AsonIndianDate)

gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.transtype, a.amount, b.accid,C.LoanAccNo FROM Bkcctrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & AsOnDate & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND (a.TransType = " & _
            wDeposit & " Or a.TransType = " & wWithDraw & ")"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY a.loanid, a.transid"
' Execute the query...

If gDbTrans.SQLFetch < 0 Then GoTo CalculateInterest

Dim SchemeName As String
Dim LoanHeadPrint As Boolean

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 229) & " " & _
                    LoadResString(gLangOffSet + 215)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            ElseIf TransType = wWithDraw Then
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'****************************
CalculateInterest:
gDbTrans.SQLStmt = "SELECT SchemeName, c.MemberID, a.loanID, d.title + space(1) " _
        & "+ d.firstname + space(1) + d.middlename +space(1) + d.lastname " _
        & "as custname, a.amount, b.accid,C.LoanAccNo FROM bkcctrans a, mmmaster b, " _
        & "loanmaster c, nametab d, loantypes e WHERE a.loanid = c.loanid AND c.memberid = " _
        & "b.accid AND b.customerid = d.customerid AND c.schemeid = e.schemeid "

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate = #" & AsOnDate & "#"
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.TransType = " & wCharges
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " ORDER BY a.loanid, a.transid"

' Execute the query...
If gDbTrans.SQLFetch < 0 Then
    ' Error in database.
    GoTo LastLine
End If
Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
LoanHeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 229) & " " & _
                    LoadResString(gLangOffSet + 236)
                .CellFontBold = True
                HeadPrint = False
            End If
            If SchemeName <> Rst("SchemeName") Then _
                LoanHeadPrint = True: SchemeName = Rst("SchemeName")
            If LoanHeadPrint Then
                LoanHeadPrint = False
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1: .Text = FormatField(Rst("SchemeName"))
                .CellAlignment = 1: .CellFontBold = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Custname"))
            .Col = 2: .Text = " " & FormatField(Rst("LoanAccNo")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
BKCCLoanTrans = True

LastLine:

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
End Function
Public Function MatFDTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency


gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from MFDTrans A, FDMaster B, NameTab C where " & _
    " TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID " & _
    " And B.CustomerId = C.CustomerID " & _
    " And (TransType = " & wDeposit & " Or TransType = " & wWithDraw & ")" & _
    " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line 'CalculateInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = False
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If Not HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 46) & " " & LoadResString(gLangOffSet + 412)
                .CellFontBold = True
                HeadPrint = True
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

'******************************************************************
CalculateInterest:

Exit_Line:
MatFDTrans = True
If Not ShowTotal Then Exit Function

If DebitAmount <> 0 Or CreditAmount <> 0 Then
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(DebitAmount): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(CreditAmount): .CellFontBold = True
    End With
End If

End Function

Public Function FDTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from FDTrans A, FDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And (TransType = " & wDeposit & _
    " Or TransType = " & wWithDraw & _
    ") order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 412)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If grd.Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: grd.Text = FormatCurrency(CreditAmount): grd.CellAlignment = 7: grd.CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'******************************************************************
CalculateInterest:
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, " & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from FDTrans A, FDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And TransType = " & wInterest & _
    " order by A.AccID "
    
If gDbTrans.SQLFetch <= 0 Then
    GoTo CalculateInterestPayable
End If

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 412) & _
                    " " & LoadResString(gLangOffSet + 487)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 5: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        CreditAmount = CreditAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If


'**************************************
CalculateInterestPayable:

gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType, " & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from FDTrans A, FDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And (TransType = " & wContraInterest & ")" & _
    " order by A.AccID, A.DepositId "


If gDbTrans.SQLFetch <= 0 Then Exit Function

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 412) & _
                    " " & LoadResString(gLangOffSet + 449)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 270): .CellAlignment = 4
            TransType = FormatField(Rst("TransType"))
            If TransType = wContraInterest Then
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            Else
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
FDTrans = True
If Not ShowTotal Then Exit Function

If DebitAmount <> 0 Or CreditAmount <> 0 Then
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(DebitAmount): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(CreditAmount): .CellFontBold = True
    End With
End If


End Function

Public Function MatDLTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

Dim DLClass As clsDLAcc
Dim DepositName As String
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from MDLTrans A, DLMaster B, NameTab C " & _
    " where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And A.AccId = B.AccId And A.DepositID = B.DepositID " & _
    " And B.CustomerId = C.CustomerID " & _
    " And (TransType = " & wDeposit & " Or TransType = " & wWithDraw & _
    ") order by A.AccID,A.DepositID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line

Set Rst = gDbTrans.Rst.Clone
Set DLClass = New clsDLAcc
DepositName = LoadResString(gLangOffSet + 46) & " " & DLClass.DepositName

Set DLClass = Nothing
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = DepositName
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatField(Rst("Amount")): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If


Exit_Line:
MatDLTrans = True
If Not ShowTotal Then Exit Function

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function

Public Function RDTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency


gDbTrans.SQLStmt = "Select A.AccId, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from RDTrans A, RDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And (TransType = " & wDeposit & _
    " Or TransType = " & wWithDraw & _
    ") order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 413)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatField(Rst("Amount")): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'****************************************************************
CalculateInterest:
gDbTrans.SQLStmt = "Select A.AccId, Amount, " & _
        " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
        " from RDTrans A, RDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
        " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
        " And Loan = " & False & " And TransType = " & wInterest & _
        " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo CalculateInterestPayable

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 413) & _
                    " " & LoadResString(gLangOffSet + 487)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 5: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        CreditAmount = CreditAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'**************************************
CalculateInterestPayable:
gDbTrans.SQLStmt = "Select A.AccId, Amount, " & _
        " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
        " from RDTrans A, RDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
        " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
        " And Loan = " & False & " And TransType = " & wContraInterest & _
        " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo LastLine:

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 413) & _
                    " " & LoadResString(gLangOffSet + 449)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 270): .CellAlignment = 4
            .Col = 5: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        CreditAmount = CreditAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

LastLine:
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If


Exit_Line:

End Function

Public Function PDTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency
Dim UtilClass As clsUtils

'PDAccount
    Set UtilClass = New clsUtils
    Debit = UtilClass.Deposits(AsonIndianDate, AsonIndianDate, wis_PD)
    Credit = UtilClass.WithDrawals(AsonIndianDate, AsonIndianDate, wis_PD)
    If Debit <> 0 Or Credit <> 0 Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .CellAlignment = 2
            .Text = LoadResString(gLangOffSet + 414) '" Pigmy Deposits"
            .Col = 4: .Text = FormatCurrency(Debit): .CellAlignment = 7
            .Col = 5: .Text = FormatCurrency(Credit): .CellAlignment = 7
        End With
    End If
    Debit = 0: Credit = 0
    Set UtilClass = Nothing

    Set UtilClass = New clsUtils
    Debit = UtilClass.Profit(AsonIndianDate, AsonIndianDate, wis_PD)
    Credit = UtilClass.Loss(AsonIndianDate, AsonIndianDate, wis_PD)
    If Debit > 0 Or Credit > 0 Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .CellAlignment = 2: .CellFontName = gFontName
            .Text = "PD " & LoadResString(gLangOffSet + 414) & " " & LoadResString(gLangOffSet + 487)  '"Interest Issued Deposits"
            .Col = 4: .Text = FormatCurrency(Debit): .CellAlignment = 7
            .Col = 5: .Text = FormatCurrency(Credit): .CellAlignment = 7
        End With
        Debit = 0: Credit = 0
    End If
    Debit = UtilClass.Deposits(AsonIndianDate, AsonIndianDate, wis_PDLoan)
    Credit = UtilClass.WithDrawals(AsonIndianDate, AsonIndianDate, wis_PDLoan)
    If Debit <> 0 Or Credit <> 0 Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .CellAlignment = 2
            .Text = LoadResString(gLangOffSet + 414) & " " & LoadResString(gLangOffSet + 58)  '"Pigmy Deposits" Loans
            .Col = 4: .Text = FormatCurrency(Debit): .CellAlignment = 7
            .Col = 5: .Text = FormatCurrency(Credit): .CellAlignment = 7
        End With
    End If
    Debit = 0: Credit = 0
    Set UtilClass = Nothing

    Set UtilClass = New clsUtils
    Debit = UtilClass.Profit(AsonIndianDate, AsonIndianDate, wis_PDLoan)
    Credit = UtilClass.Loss(AsonIndianDate, AsonIndianDate, wis_PDLoan)
    If Debit > 0 Or Credit > 0 Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .CellAlignment = 2: .CellFontName = gFontName
            .Text = LoadResString(gLangOffSet + 414) & " " & LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 483) '"Interest Issued Deposits Loans"
            .Col = 2: .Text = FormatCurrency(Debit): .CellAlignment = 7
            .Col = 3: .Text = FormatCurrency(Credit): .CellAlignment = 7
        End With
        Debit = 0: Credit = 0
    End If
    Set UtilClass = Nothing

End Function

Public Function FDLoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
                   " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
                   " from FDTrans A, FDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
                   " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
                   " And Loan = " & True & " And (TransType = " & wDeposit & _
                   " Or TransType = " & wWithDraw & _
                   ") order by A.AccID "


If gDbTrans.SQLFetch <= 0 Then GoTo ShowInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 412) & " " & LoadResString(gLangOffSet + 58)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'***********************************************
ShowInterest:
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, " & _
        " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
        " from FDTrans A, FDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
        " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
        " And Loan = " & True & " And TransType = " & wCharges & _
        " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then Exit Function

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 412) & _
                    " " & LoadResString(gLangOffSet + 483)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
FDLoanTrans = True
If Not ShowTotal Then Exit Function


If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function


Public Function DLLoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency


Dim DLClass As clsDLAcc
Dim DepositName As String
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from DLTrans A, DLMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & True & " And (TransType = " & wDeposit & _
    " Or TransType = " & wWithDraw & _
    ") order by A.AccID "

If gDbTrans.SQLFetch <= 0 Then GoTo ShowInterest

Set Rst = gDbTrans.Rst.Clone
Set DLClass = New clsDLAcc
DepositName = DLClass.DepositName
Set DLClass = Nothing
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = DepositName
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'***********************************************
ShowInterest:
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, " & _
        " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
        " from DLTrans A, DLMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
        " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
        " And Loan = " & True & " And TransType = " & wCharges & _
        " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then
    Exit Function
End If

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = DepositName & " " & LoadResString(gLangOffSet + 483)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:

DLLoanTrans = True
If Not ShowTotal Then Exit Function

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function

Public Function RDLoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

gDbTrans.SQLStmt = "Select A.AccId, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from RDTrans A, RDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " And Loan = " & True & " And (TransType = " & wDeposit & _
    " Or TransType = " & wWithDraw & _
    ") order by A.AccID "

If gDbTrans.SQLFetch <= 0 Then GoTo ShowInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 413) & " " & LoadResString(gLangOffSet + 58)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = TransAmount + DebitAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        If grd.Rows = grd.Row + 2 Then grd.Rows = grd.Rows + 2
        grd.Row = grd.Row + 1
        grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 304): grd.CellAlignment = 7: grd.CellFontBold = True
        grd.Col = 4: grd.Text = FormatCurrency(DebitAmount): grd.CellAlignment = 7: grd.CellFontBold = True
        grd.Col = 5: grd.Text = FormatCurrency(CreditAmount): grd.CellAlignment = 7: grd.CellFontBold = True
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'***********************************************
ShowInterest:
gDbTrans.SQLStmt = "Select A.AccId, Amount, " & _
            " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
            " from RDTrans A, RDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
            " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
            " And Loan = " & True & " And TransType = " & wCharges & _
            " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo LastLine

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 413) & " " & LoadResString(gLangOffSet + 58) & _
                    " " & LoadResString(gLangOffSet + 483)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close


LastLine:
Set Rst = Nothing
RDLoanTrans = True

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

Exit_Line:

End Function


Public Function PDLoanTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency


gDbTrans.SQLStmt = "Select A.AccId, Amount, TransType ," & _
        " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
        " from PDTrans A, PDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
        " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
        " And Loan = " & True & " And (TransType = " & wDeposit & _
        " Or TransType = " & wWithDraw & _
        ") order by A.AccID "

If gDbTrans.SQLFetch <= 0 Then GoTo ShowInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    If FormatField(Rst("Amount")) <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 414) & " " & LoadResString(gLangOffSet + 58)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + FormatField(Rst("Amount"))
            Else
                .Col = 5
                CreditAmount = CreditAmount + FormatField(Rst("Amont"))
            End If
            .Text = FormatField(Rst("Amount")): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
'***********************************************
ShowInterest:
gDbTrans.SQLStmt = "Select A.AccId, Amount, " & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from PDTrans A, PDMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " And Loan = " & True & " And TransType = " & wCharges & _
    " And A.UserID = B.UserID" & _
    " order by A.AccID "

If gDbTrans.SQLFetch <= 0 Then Exit Function

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    If FormatField(Rst("Amount")) <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 414) & " " & _
                    LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 483)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 4: .Text = FormatField(Rst("Amount")): .CellAlignment = 7
        End With
        DebitAmount = DebitAmount + FormatField(Rst("Amount"))
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

Exit_Line:
PDLoanTrans = True
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
    End If
End If


End Function


Public Function SBTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency


gDbTrans.SQLStmt = "Select SBTrans.Accid, Amount, TransType, " & _
        "  title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
        " From SBTrans,SBMaster,NameTab where TransDate = #" & FormatDate(AsonIndianDate) & "# " & _
        " And SbMaster.AccId = SBTrans.AccId  And NameTab.CustomerId = SBMaster.CustomerID" & _
        " And (TransType = " & wDeposit & " Or TransType = " & wWithDraw & _
        " or TransType = " & wContraDeposit & ")" & _
        " order by SBTrans.AccID"

If gDbTrans.SQLFetch <= 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 410)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Or TransType = wContraDeposit Then
                If TransType = wDeposit Then
                    .Text = LoadResString(gLangOffSet + 381)
                Else
                    .Text = LoadResString(gLangOffSet + 270)
                End If
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                If TransType = wWithDraw Then
                    .Text = LoadResString(gLangOffSet + 381)
                Else
                    .Text = LoadResString(gLangOffSet + 270)
                End If
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

CalculateInterest:
'***************************************************************
gDbTrans.SQLStmt = "Select SBTrans.Accid, Amount," & _
        "  title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
        " From SBTrans,SBMaster,NameTab where TransDate = #" & FormatDate(AsonIndianDate) & "# " & _
        " And SbMaster.AccId = SBTrans.AccId  And NameTab.CustomerId = SBMaster.CustomerID" & _
        " And TransType = " & wContraInterest & _
        " order by SBTrans.AccID"
                   
If gDbTrans.SQLFetch <= 0 Then GoTo LastLine

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 436) & " " & LoadResString(gLangOffSet + 487)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 270): .CellAlignment = 4
            .Col = 5: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        CreditAmount = CreditAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close

LastLine:
Set Rst = Nothing

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function


Public Function CATrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

gDbTrans.SQLStmt = "Select CATrans.Accid, Amount, TransType, " & _
        "  title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
        " From CATrans,CAMaster,NameTab where TransDate = #" & FormatDate(AsonIndianDate) & "# " & _
        " And CAMaster.AccId = CATrans.AccId  And NameTab.CustomerId = CAMaster.CustomerID" & _
        " And (TransType = " & wDeposit & " Or TransType = " & wWithDraw & ")" & _
        " order by CATrans.AccID"
                   
If gDbTrans.SQLFetch <= 0 Then GoTo LastLine

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = LoadResString(gLangOffSet + 411)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing

LastLine:
CATrans = True

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

ExitLine:

End Function

Public Function DLTrans(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean
Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency

Dim DLClass As clsDLAcc
Dim DepositName As String
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, TransType ," & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from DLTrans A, DLMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And (TransType = " & wDeposit & _
    " Or TransType = " & wWithDraw & _
    ") order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo CalculateInterest

Set Rst = gDbTrans.Rst.Clone
Set DLClass = New clsDLAcc
DepositName = DLClass.DepositName
Set DLClass = Nothing
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = DepositName
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            TransType = FormatField(Rst.Fields("TransType"))
            If TransType = wDeposit Then
                .Col = 4
                DebitAmount = DebitAmount + TransAmount
            Else
                .Col = 5
                CreditAmount = CreditAmount + TransAmount
            End If
            .Text = FormatField(Rst("Amount")): .CellAlignment = 7
        End With
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If
'***********************************************************
CalculateInterest:
gDbTrans.SQLStmt = "Select A.AccId,A.DepositID, Amount, " & _
    " Title + ' ' + FirstName + ' ' +  MiddleName + ' ' + LastName as Name " & _
    " from DLTrans A, DLMaster B, NameTab C where TransDate = #" & FormatDate(AsonIndianDate) & "#" & _
    " And  A.AccId = B.AccId And A.DepositID = B.DepositID  And B.CustomerId = C.CustomerID " & _
    " And Loan = " & False & " And TransType = " & wInterest & _
    " order by A.AccID "
                   
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line

Set Rst = gDbTrans.Rst.Clone
HeadPrint = True
While Not Rst.EOF
    TransAmount = FormatField(Rst("Amount"))
    If TransAmount <> 0 Then
        With grd
            If HeadPrint Then
                If .Rows = .Row + 2 Then .Rows = .Rows + 2
                .Row = .Row + 1
                .Col = 1
                .Text = DepositName & " " & LoadResString(gLangOffSet + 487)
                .CellFontBold = True
                HeadPrint = False
            End If
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            M_SlNo = M_SlNo + 1
            .Col = 0: .Text = Format(M_SlNo)
            .Col = 1: .Text = FormatField(Rst("Name"))
            .Col = 2: .Text = " " & FormatField(Rst("AccID")): .CellAlignment = 4
            .Col = 3: .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            .Col = 5: .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        End With
        CreditAmount = CreditAmount + TransAmount
    End If
    Rst.MoveNext
Wend
Rst.Close
Set Rst = Nothing
'**************************************

Exit_Line:
DLTrans = True
If Not ShowTotal Then Exit Function

If ShowTotal Then
'Now Print The Total Debit & Credit  SHashi 9/9/01
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        With grd
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 1: .Text = LoadResString(gLangOffSet + 304): .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(DebitAmount): .CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(CreditAmount): .CellAlignment = 7: .CellFontBold = True
        End With
        DebitAmount = 0
        CreditAmount = 0
    End If
End If

End Function


Public Function BankAccounts(AsonIndianDate As String, grd As MSFlexGrid, Optional ShowTotal As Boolean) As Boolean

Dim NameStr() As String
Dim HeadId() As Long
Dim BankClass As clsBankAcc
Dim LoopCount As Integer
Dim AccountID As Long
Dim TransType As wisTransactionTypes

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim TransAmount As Currency


''''''''''''''''''''''''The Below Set OF Code Has Witten By shashi
'''''''''''''''''''''''' On 11/9/2001
#If Completed Then

gDbTrans.SQLStmt = "SELECT AccID, Amount, Particulars, ChequeNo, TransType FROM AccTrans" & _
  " WHERE AccID > " & wis_BankHead & " And AccID < " & wis_BankHead + wis_BankHeadOffSet & _
  " And TransDate = #" & FormatDate(AsonIndianDate) & "# ORDER BY AccID"

If gDbTrans.SQLFetch < 1 Then Exit Sub

Dim rstHeadName As Recordset
Dim MainHeadID As Long
Dim MainHeadPrint As Boolean
Dim SubHeadPrint As Boolean
Dim SubHeadId As Long

'Get The Head name Form Account Master
gDbTrans.SQLStmt = "SELECT AccID, AccName FROM " & _
  " AccMaster ORDER BY AccID"
Call gDbTrans.SQLFetch
Set rstHeadName = gDbTrans.Rst.Clone
    
    Set Rst = gDbTrans.Rst.Clone
    If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1: grd.Col = 1
    grd.CellAlignment = 2: grd.CellFontBold = True
    While Not Rst.EOF
        AccountID = FormatField(Rst("ACCID"))
    Wend

Exit Function
#End If
'''''''''''''''''''Shashi's code ends heare

Dim IDOfHead As Long
Dim PreviousHeadID As Long
Dim PreviousAccountID As Long
On Error GoTo ErrorLine
gDbTrans.SQLStmt = "SELECT AccID, Amount, Particulars, ChequeNo, TransType FROM AccTrans" & _
  " WHERE TransDate = #" & FormatDate(AsonIndianDate) & "# ORDER BY AccID"
    
If gDbTrans.SQLFetch < 1 Then GoTo ErrorLine

With grd
    Set Rst = gDbTrans.Rst.Clone
    DebitAmount = 0: CreditAmount = 0
    While Not Rst.EOF
        IDOfHead = CLng(Left(CStr(Rst("AccID")), Len(CStr(Rst("AccID"))) - 3) & "000")
        If IDOfHead <> PreviousHeadID Then
            If DebitAmount <> 0 Or CreditAmount <> 0 Then
                If .Rows <= .Row + 2 Then .Rows = .Rows + 1
                .Row = .Row + 1
                .Col = 1: .Text = LoadResString(gLangOffSet + 304)
                .CellAlignment = 7: .CellFontBold = True
                .Col = 4: .Text = FormatCurrency(DebitAmount)
                .CellAlignment = 7: .CellFontBold = True: DebitAmount = 0
                .Col = 5: .Text = FormatCurrency(CreditAmount)
                .CellAlignment = 7: .CellFontBold = True: CreditAmount = 0
            End If
            gDbTrans.SQLStmt = "Select AccName FROM AccMaster Where AccID = " & IDOfHead
            
            If gDbTrans.SQLFetch < 1 Then Exit Function
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 1: .Text = FormatField(gDbTrans.Rst(0)): .CellFontBold = True
            PreviousHeadID = IDOfHead
            gDbTrans.Rst.Close
            Set gDbTrans.Rst = Nothing
        End If
        AccountID = FormatField(Rst.Fields("AccID"))
        If AccountID <> PreviousAccountID Then
            If DebitAmount <> 0 Or CreditAmount <> 0 Then
                If .Rows <= .Row + 2 Then .Rows = .Rows + 1
                .Row = .Row + 1
                .Col = 1: .Text = LoadResString(gLangOffSet + 304)
                .CellAlignment = 7: .CellFontBold = True
                .Col = 4: .Text = FormatCurrency(DebitAmount)
                .CellAlignment = 7: .CellFontBold = True: DebitAmount = 0
                .Col = 5: .Text = FormatCurrency(CreditAmount)
                .CellAlignment = 7: .CellFontBold = True: CreditAmount = 0
            End If
            gDbTrans.SQLStmt = "Select AccName FROM AccMaster Where AccID = " & AccountID
            
            If gDbTrans.SQLFetch < 1 Then Exit Function
            
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 1: .Text = FormatField(gDbTrans.Rst(0))
            PreviousAccountID = AccountID
            gDbTrans.Rst.Close
            Set gDbTrans.Rst = Nothing
        End If
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        M_SlNo = M_SlNo + 1
        .Col = 0: .Text = Format(M_SlNo)
        TransAmount = FormatField(Rst.Fields("Amount"))
        TransType = FormatField(Rst.Fields("TransType"))
        .Col = 1: .Text = FormatField(Rst.Fields("Particulars"))
        .Col = 3
        If TransType = wDeposit Or TransType = wCharges Or TransType = wContraCharges _
                    Or TransType = wContraDeposit Or TransType = wRPCharges Then
            If TransType = wDeposit Or TransType = wRPCharges Or TransType = wCharges Then
                .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            Else
                .Text = LoadResString(gLangOffSet + 270): .CellAlignment = 4
            End If
            .Col = 4
            DebitAmount = DebitAmount + TransAmount
        Else
            If TransType = wWithDraw Or TransType = wRPInterest Or TransType = wInterest Then
                .Text = LoadResString(gLangOffSet + 381): .CellAlignment = 4
            Else
                .Text = LoadResString(gLangOffSet + 270): .CellAlignment = 4
            End If
            .Col = 5
            CreditAmount = CreditAmount + TransAmount
        End If
        .Text = FormatCurrency(TransAmount): .CellAlignment = 7
        Rst.MoveNext
    Wend
    
    If DebitAmount <> 0 Or CreditAmount <> 0 Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 1: .Text = LoadResString(gLangOffSet + 304)
        .CellAlignment = 7: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(DebitAmount)
        .CellAlignment = 7: .CellFontBold = True: DebitAmount = 0
        .Col = 5: .Text = FormatCurrency(CreditAmount)
        .CellAlignment = 7: .CellFontBold = True: CreditAmount = 0
    End If
End With


ErrorLine:
    If Err.Number = 9 Then
        HeadPrint = False
        LoopCount = -1
        Resume Next
    End If
End Function



