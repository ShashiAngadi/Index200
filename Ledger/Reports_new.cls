VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private m_frmWeb As frmWeb

Private WithEvents m_frmRptDt As frmRptDt
Attribute m_frmRptDt.VB_VarHelpID = -1
Private WithEvents m_frmAccReports As frmAccReports
Attribute m_frmAccReports.VB_VarHelpID = -1
Private WithEvents m_WebDoc As HTMLDocument
Attribute m_WebDoc.VB_VarHelpID = -1

Private m_WebTable As HTMLTable
Private m_WebTBody As HTMLTableSection
Private m_WebRow As HTMLTableRow
Private m_WebCell As HTMLTableCell
Private m_DOmElement
'To make the total for the report
Dim m_withdrawTotal As Currency
Dim m_depositTotal As Currency
Dim m_contraWithdrawTotal As Currency
Dim m_contraDepositTotal As Currency

'''USED ONLY FOR UPDATELEDGER METHOD
Dim m_rstLedgerTrans As Recordset

' To Handle Report Dates
Private m_FromRpDate As String
Private m_ToRpDate As String
Private m_FromUSDate As Date
Private m_ToUSDate As Date

Private m_RpHeadId As Long
Private m_RpParentID As Long

Private m_FromRpAmount As Currency
Private m_ToRpAmount As Currency

Private m_OtherDetail

Private m_AccountReportList As Wis_AccountReportList



' This will fill the imablanced row of the Balance Sheet
Private Sub BALFillEmptyRows(Optional NewBook As Boolean = False)

Dim CurrentRow As HTMLTableRow
Dim CurrentTable As HTMLTable

Dim LiabilityRows As Integer
Dim AssetRows As Integer
Dim RowsToAdd As Integer
Dim Row As Integer

LiabilityRows = m_WebRow.cells(0).All.Liability.Rows.Length - 1
AssetRows = m_WebRow.cells(1).All.Asset.Rows.Length - 1

If LiabilityRows > AssetRows Then

    RowsToAdd = LiabilityRows - AssetRows
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    CurrentTable.Style.fontFamily = gFontName
    CurrentTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
Else
    
    RowsToAdd = AssetRows - LiabilityRows
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    CurrentTable.Style.fontFamily = gFontName
    CurrentTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)

End If

For Row = 1 To RowsToAdd

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
    
Next Row
    

End Sub

Private Sub BALONstartReport()

Dim DOMElement

'With m_WebDoc
'    Set m_WebTable = m_WebDoc.All("HeaderTbl")
'    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
'    With m_WebRow.cells(0)
'        .innerText = GetResourceString(67) & " " & ToRpDate '"Balance as on " & ToRpDate
'        .Style.fontWeight = "Bold"
'        .Style.FontSize = "11"
'        .Align = "Center"
'    End With
'
'End With
Dim strHead As String
strHead = GetResourceString(67) & " " & GetFromDateString(m_ToRpDate)
Call SetReportTitle(strHead, False)

Call StartHtmlTable("Balance", m_WebDoc.body, 4)
    
With m_WebTable
    .Width = "500"
End With
m_WebTable.Style.fontFamily = gFontName
m_WebTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = GetResourceString(33) '"Sl No"
        .Width = "10%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = GetResourceString(36, 35) '"Head Name"
        .Width = "65%"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = GetResourceString(91) '40 '"Amount"
        .Width = "20%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = " "
        .Width = "5%"
        .Align = "Left"
    End With
    
End With

End Sub

Private Function InsertCashBookTotals(CurrentTable As HTMLTable, NewBook As Boolean, SlNo As Integer, name As String, _
    Deposit As Currency, contraDeposit As Currency, WithDraw As Currency, contraWithdraw As Currency, Optional ShowTotal As Boolean = True)


        Dim CurrentRow As HTMLTableRow
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 9, 7))
        Set InsertCashBookTotals = CurrentRow
        Set m_WebRow = CurrentRow
        CurrentRow.Style.fontWeight = "Bold"
        If SlNo > 0 And NewBook = False Then
            With CurrentRow.cells(1)
                .innerText = name
                .Align = "Left"
            End With
        End If
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = name
            .Align = "Left"
        End With
        If Deposit Then
        With CurrentRow.cells(IIf(NewBook, 2, 3))
            .innerText = FormatCurrency(Deposit)
            .Align = "Right"
        End With
        End If
        If contraDeposit Then
        With CurrentRow.cells(IIf(NewBook, 1, 4))
            .innerText = FormatCurrency(contraDeposit)
            .Align = "Right"
        End With
        End If
        If NewBook And ShowTotal And (contraDeposit Or Deposit) Then
            With CurrentRow.cells(0)
                .innerText = FormatCurrency(Deposit + contraDeposit)
                .Align = "Right"
            End With
        End If

        If WithDraw Then
        With CurrentRow.cells(IIf(NewBook, 6, 5))
            .innerText = FormatCurrency(WithDraw)
            .Align = "Right"
        End With
        End If
        If contraWithdraw Then
        With CurrentRow.cells(IIf(NewBook, 7, 6))
            .innerText = FormatCurrency(contraWithdraw)
            .Align = "Right"
        End With
        End If
        If NewBook And ShowTotal And (WithDraw Or contraWithdraw) Then
            With CurrentRow.cells(8)
                .innerText = FormatCurrency(WithDraw + contraWithdraw)
                .Align = "Right"
            End With
        End If

        Deposit = 0: contraDeposit = 0: WithDraw = 0: contraWithdraw = 0
        
End Function

Private Sub SetReportTitle(strTitle As String, Optional ShowDate As Boolean = True)

Dim DOMElement
Dim strDate As String

With m_WebDoc
   
   '.body.appendChild DOMElement
    'Set DOMElement = .createElement("H3")
    '.body.appendChild DOMElement
    
    Call StartHtmlTable("TitleTbl", m_WebDoc.body, 1)
    
    'Set m_WebTable = m_WebDoc.All("HeaderTbl", 0)
    'm_WebTable.border = 0
    'Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    Set DOMElement = .createElement("H1")
    .body.appendChild DOMElement
    With DOMElement
        .innerText = strTitle
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FONTSIZE = IIf(gLangOffSet, "19", "17")
    End With
    '.body.appendChild DOMElement
    If ShowDate Then
        Set DOMElement = .createElement("H3")
        strDate = GetFromDateString(FromRpDate, ToRpDate)
        .body.appendChild DOMElement
        With DOMElement
            .innerText = strDate
            .Align = "Center"
            .Style.fontFamily = gFontName
            .Style.FONTSIZE = IIf(gLangOffSet, "17", "15")
        End With
    End If
'    Set m_WebTable = .createElement("Table")
'    Set DOMElement = .createElement("Tbody")
'    m_WebTable.appendChild DOMElement
End With

End Sub

Private Function InsertTransDetails(rstTrans As Recordset, SlNo As Long, NewBook As Boolean) As Boolean

Dim CurrentRow As HTMLTableRow
Dim Amount As Currency
Dim transType As wisTransactionTypes
Dim colno As Byte
Dim TotalDeposit As Currency
Dim TotalContraDeposit As Currency
Dim TotalWithDraw As Currency
Dim TotalContraWithDraw As Currency

Dim firstRound As Boolean

firstRound = True

Repeat:

Do While Not rstTrans.EOF
    
    SlNo = SlNo + 1
    transType = rstTrans("TransType")
    Amount = FormatField(rstTrans("Amount"))
    If NewBook Then
        
        If firstRound Then
            If (transType = wWithdraw Or transType = wContraWithdraw) Then GoTo nextRecord
        Else
            If (transType = wDeposit Or transType = wContraDeposit) Then GoTo nextRecord
        End If
        
        If transType = wDeposit Then colno = 2: TotalDeposit = TotalDeposit + Amount
        If transType = wContraDeposit Then colno = 1: TotalContraDeposit = TotalContraDeposit + Amount
        If transType = wWithdraw Then colno = 6: TotalWithDraw = TotalWithDraw + Amount
        If transType = wContraWithdraw Then colno = 7: TotalContraWithDraw = TotalContraWithDraw + Amount
        
        Set CurrentRow = InsertRowAndCellSToTable(m_WebTable, IIf(NewBook, 9, 7))
        With CurrentRow.cells(0)
            '.innerText = SlNo
            .Align = "Right"
        End With
    
        With CurrentRow.cells(IIf(firstRound, 3, 5))
            .innerText = FormatField(rstTrans("AccNum"))
            .Align = "Left"
        End With
        With CurrentRow.cells(4)
            .innerText = FormatField(rstTrans("Name"))
            .Align = "Left"
        End With
        
        With CurrentRow.cells(colno)
            .innerText = FormatCurrency(Amount)
            .Align = "Right"
        End With
    Else
        If transType = wDeposit Then colno = 3: TotalDeposit = TotalDeposit + Amount
        If transType = wContraDeposit Then colno = 4: TotalContraDeposit = TotalContraDeposit + Amount
        If transType = wWithdraw Then colno = 5: TotalWithDraw = TotalWithDraw + Amount
        If transType = wContraWithdraw Then colno = 6: TotalContraWithDraw = TotalContraWithDraw + Amount
    
        Set CurrentRow = InsertRowAndCellSToTable(m_WebTable, IIf(NewBook, 9, 7))
    
        With CurrentRow.cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With
    
        With CurrentRow.cells(1)
            .innerText = FormatField(rstTrans("Name"))
            .Align = "Left"
        End With
        With CurrentRow.cells(2)
            .innerText = FormatField(rstTrans("AccNum"))
            .Align = "Left"
        End With
    
        With CurrentRow.cells(colno)
            .innerText = FormatCurrency(Amount)
            .Align = "Right"
        End With
    
    End If
    
    
nextRecord:

    rstTrans.MoveNext
Loop

If firstRound And NewBook Then
    firstRound = False
    rstTrans.MoveFirst
    GoTo Repeat
End If
Set CurrentRow = InsertRowAndCellSToTable(m_WebTable, IIf(NewBook, 9, 7))
CurrentRow.Style.fontWeight = "Bold"

If NewBook Then
    With CurrentRow.cells(1)
        '.innerText = GetResourceString(42)
        .Align = "Left"
    End With
    If TotalDeposit Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(TotalDeposit)
        .Align = "Right"
    End With
    End If
    If TotalContraDeposit Then
    With CurrentRow.cells(1)
        .innerText = FormatCurrency(TotalContraDeposit)
        .Align = "Right"
    End With
    End If
    If TotalDeposit Or TotalContraDeposit Then
    With CurrentRow.cells(0)
        .innerText = FormatCurrency(TotalDeposit + TotalContraDeposit)
        .Align = "Right"
    End With
    End If
    If TotalWithDraw Then
    With CurrentRow.cells(6)
        .innerText = FormatCurrency(TotalWithDraw)
        .Align = "Right"
    End With
    End If
    If TotalContraWithDraw Then
    With CurrentRow.cells(7)
        .innerText = FormatCurrency(TotalContraWithDraw)
        .Align = "Right"
    End With
    End If
    If TotalWithDraw Or TotalContraWithDraw Then
    With CurrentRow.cells(8)
        .innerText = FormatCurrency(TotalWithDraw + TotalContraWithDraw)
        .Align = "Right"
    End With
    End If
    
Else

    With CurrentRow.cells(1)
        .innerText = GetResourceString(42)
        .Align = "Left"
    End With
    
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(TotalDeposit)
        .Align = "Right"
    End With
    With CurrentRow.cells(4)
        .innerText = FormatCurrency(TotalContraDeposit)
        .Align = "Right"
    End With
    With CurrentRow.cells(5)
        .innerText = FormatCurrency(TotalWithDraw)
        .Align = "Right"
    End With
    With CurrentRow.cells(6)
        .innerText = FormatCurrency(TotalContraWithDraw)
        .Align = "Right"
    End With
End If

End Function

Public Function ShowBalancing() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstBalances As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstOpBalance As ADODB.Recordset

Dim AccTransClass As clsAccTrans
'Dim MaterialClass As clsMaterial

Dim Balance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim StockValue As Currency

Dim DiffBalance As Currency

Dim Balance_Actual As Currency
Dim StockValue_Actual As Currency

Dim GrandTotal(2 To 7) As Currency
Dim colno As Integer

Dim opBalance As Currency

Dim headID As Long
Dim ParentID As Long
Dim PrevParentID As Long

Dim SlNo As Long

Dim headName As String
Dim ParentName As String
Dim fromDate As String
Dim toDate As String

Dim DRCRType As wis_DrCrType

Dim fldHeadID As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldParentName As ADODB.Field
Dim fldParentID As ADODB.Field
Dim fldAccountType As ADODB.Field

Dim isHeadID As Boolean

Dim AccountType As wis_AccountType

' Intialise the Function
ShowBalancing = True

Call BalancingstartReport

Set AccTransClass = New clsAccTrans

ParentID = RpParentID
headID = RpHeadID

fromDate = FromRpDate
toDate = ToRpDate

gDbTrans.SqlStmt = " SELECT HeadID,OpAmount" & _
        " FROM OpBalance WHERE OpDate = " & "#" & m_FromUSDate & "#" '& _

Call gDbTrans.Fetch(rstOpBalance, adOpenDynamic)

gDbTrans.SqlStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal,HeadId" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                   " AND #" & GetSysFormatDate(toDate) & "#" & _
                   " GROUP BY HeadID"
                   
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)

Dim AccType As wis_AccountType
AccType = Asset Or Liability
gDbTrans.SqlStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " And (AccountType = " & Asset & _
                        " OR AccountType = " & Liability & ")" & _
                   " AND B.UserCreated > 2 " & _
                   " ORDER BY b.ParentID"


' Get the Recordset

Call gDbTrans.Fetch(rstBalances, adOpenForwardOnly)

Set fldHeadID = rstBalances.Fields("HeadID")
Set fldHeadName = rstBalances.Fields("HeadName")
Set fldParentName = rstBalances.Fields("ParentName")
Set fldParentID = rstBalances.Fields("ParentID")
Set fldAccountType = rstBalances.Fields("AccountType")

' Start the Loop

SlNo = 0
Do While Not rstBalances.EOF
    
    headID = fldHeadID.Value
    headName = fldHeadName.Value
    ParentName = fldParentName.Value
    ParentID = fldParentID.Value
    AccountType = fldAccountType.Value
    
    ' fill the parentname
    If PrevParentID <> ParentID Then
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        m_WebRow.Style.fontWeight = "Bold"
        m_WebRow.cells(1).innerText = ParentName
        PrevParentID = ParentID
    End If
    
    opBalance = 0
    If gDbTrans.FindRecord(rstOpBalance, "HeadID=" & headID) Then _
                         opBalance = rstOpBalance.Fields("OpAmount")
                    
    DebitTotal = 0
    CreditTotal = 0
    If gDbTrans.FindRecord(rstTrans, "HeadID=" & headID) Then
        DebitTotal = FormatField(rstTrans.Fields("DebitTotal"))
        CreditTotal = FormatField(rstTrans.Fields("CreditTotal"))
    End If
                
    Balance = AccTransClass.GetClosingBalanceonAccountType _
                    (opBalance, DebitTotal, CreditTotal, AccountType)
    
    Balance_Actual = GetAccountBalance(headID, m_ToUSDate)
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

    SlNo = SlNo + 1

    With m_WebRow
        With .cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With

        With .cells(1)
            .innerText = headName
            .Align = "Left"
        End With
        
        DRCRType = GetDRCRType(Balance, AccountType)
        colno = IIf(DRCRType = enumDebit, 2, 3)
        With .cells(colno)
            .innerText = FormatCurrency(Abs(Balance))
            .Align = "Right"
        End With
        GrandTotal(colno) = GrandTotal(colno) + Abs(Balance)
        
        DRCRType = GetDRCRType(Balance_Actual, AccountType)
        colno = IIf(DRCRType = enumDebit, 4, 5)
        With .cells(colno)
            .innerText = FormatCurrency(Abs(Balance_Actual))
            .Align = "Right"
        End With
        GrandTotal(colno) = GrandTotal(colno) + Abs(Balance_Actual)
        
        DiffBalance = Balance - Balance_Actual
        DRCRType = GetDRCRType(DiffBalance, AccountType)
                colno = IIf(DRCRType = enumDebit, 6, 7)
        With .cells(colno)
            .innerText = FormatCurrency(Abs(DiffBalance))
            .Align = "Right"
        End With
        GrandTotal(colno) = GrandTotal(colno) + Abs(DiffBalance)

    End With
        
    rstBalances.MoveNext
Loop


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

SlNo = SlNo + 1

With m_WebRow

    With .cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = GetResourceString(175) '"Stock In Hand"
        .Align = "Left"
    End With
    
    DRCRType = GetDRCRType(StockValue, AccountType)
    
    colno = IIf(DRCRType = enumDebit, 2, 3)
    With .cells(colno)
        .innerText = FormatCurrency(Abs(StockValue))
        .Align = "Right"
    End With
    GrandTotal(colno) = GrandTotal(colno) + Abs(StockValue)
    
End With


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

' Exit function
Set AccTransClass = Nothing

Set rstBalances = Nothing
Set rstTrans = Nothing
Set rstOpBalance = Nothing

ShowBalancing = True

Exit Function

Hell:

    If Err Then MsgBox "ShowBalancing : " & vbCrLf & Err.Description
    Resume
    Set AccTransClass = Nothing
    
    Set rstBalances = Nothing
    Set rstTrans = Nothing
    Set rstOpBalance = Nothing

    Exit Function
    
End Function

Public Function GetAccountBalance(ByVal AccHeadID As Long, ByVal AsOnDate As Date) As Currency

On Error Resume Next

Dim rst As Recordset
Dim SqlStr As String
Dim pos As Long

Set rst = Nothing
On Error GoTo Exit_Line

Dim ModuleId As wisModules

ModuleId = GetModuleIDFromHeadID(AccHeadID)

Set rst = Nothing

SqlStr = ""
'Members
If ModuleId >= wis_Members And ModuleId < wis_Members + 100 Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From MemMaster A,MemTrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From MemTrans C Where C.AccID= A.AccID)"
            

'BKcc Account
ElseIf ModuleId = wis_BKCCLoan Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From BKCCMaster A,BKCCTrans B Where " & _
        " A.LoanID = B.LoanID And Deposit = " & False & _
        " And TransID = (Select Max(TransID) From BKCCTrans C " & _
            "Where C.LoanID= A.LoanID And Deposit = " & False & ")"

ElseIf ModuleId = wis_BKCC Then
    SqlStr = "Select Sum(Balance) * -1 as AcountBalance " & _
        " From BKCCMaster A,BKCCTrans B Where " & _
        " A.LoanID = B.LoanID And Deposit = True " & _
        " ANd TransID = (Select Max(TransID) From BKCCTrans C " & _
            "Where C.LoanID= A.LoanID And Deposit = true )"

'Current Account
ElseIf ModuleId = wis_CAAcc Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From CAMaster A,CATrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From CATrans C Where C.AccID= A.AccID)"

'DepositLoans
ElseIf ModuleId >= wis_DepositLoans And ModuleId < wis_DepositLoans + 100 Then
    SqlStr = ""
    If ModuleId > wis_DepositLoans Then _
        SqlStr = " ANd A.DepositType = " & ModuleId - wis_DepositLoans
    SqlStr = "Select Sum(Balance) * -1 as AcountBalance " & _
        " From DepositLoanMaster A,DepositLoanTrans B Where " & _
        " A.LoanID = B.LoanID " & SqlStr & _
        " ANd TransID = (Select Max(TransID) From DepositLoanTrans C " & _
            "Where C.LoanID= A.LoanID )"

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    SqlStr = ""
    If ModuleId > wis_Deposits Then _
        SqlStr = " ANd A.DepositType = " & ModuleId - wis_Deposits
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From FDMaster A,FDTrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From FDTrans C Where C.AccID= A.AccID)" & SqlStr

'Loan Accounts
ElseIf ModuleId = wis_Loans And ModuleId < wis_Loans + 100 Then
    If ModuleId > wis_Loans Then _
        SqlStr = " AND A.SchemeID = " & ModuleId - wis_Loans
    SqlStr = "Select Sum(Balance) * -1 as AcountBalance " & _
        " From LoanMaster A, LoanTrans B Where " & _
        " A.LoanID = B.LoanID " & SqlStr & _
        " And TransID = (Select Max(TransID) From LoanTrans C " & _
            "Where C.LoanID = A.LoanID )"

'Pigmy Accounts
ElseIf ModuleId = wis_PDAcc Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From PDMaster A,PDTrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From PDTrans C Where C.AccID= A.AccID)"

'Recurring Accounts
ElseIf ModuleId = wis_RDAcc Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From RDMaster A,RDTrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From RDTrans C Where C.AccID= A.AccID)"

ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    SqlStr = "Select Sum(Balance) as AcountBalance " & _
        " From SBMaster A,SBTrans B Where " & _
        " A.AccID = B.AccID And TransID = (Select " & _
            " Max(TransID) From SBTrans C Where C.AccID= A.AccID)"

Else
    'MsgBox "Plese select the account type", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


If SqlStr = "" Then Exit Function
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
    GetAccountBalance = FormatField(rst(0)) '"AccountBalance"))

Exit_Line:

End Function


Public Property Get ObjectName()
    ObjectName = "ReportClass"
End Property

Private Sub ShowDailyCashBook_DepositLoan(CurrentTable As HTMLTable, SlNo As Long, NewBook As Boolean, DepositType As Integer, DepositName As String, AmountType As wis_AmountType)
    
    Dim CurrentRow As HTMLTableRow
    Dim rstRPDetails As Recordset
    Dim FieldName As String
    Dim transTable As String
    If AmountType = wisPrincipal Then
        FieldName = "Amount"
        transTable = "DepositLoanTrans"
    ElseIf AmountType = wisRegularInt Then
        FieldName = "Amount"
        transTable = "DepositLoanIntTrans"
    ElseIf AmountType = wisPenalInt Then
        FieldName = "PenalAmount"
        transTable = "DepositLoanIntTrans"
    ElseIf AmountType = wisMisc Then
        FieldName = "MiscAmount"
        transTable = "DepositLoanIntTrans"
    ElseIf AmountType = wisPayable Then
        FieldName = ""
        transTable = ""
    End If
    
    
    gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId, " & FieldName & " as Amount,TransType,Name" & _
            " FROM " & transTable & " A,DepositLoanMaster B,QryName C " & _
            " WHERE B.DepositType = " & DepositType & " And TransDate = #" & m_ToUSDate & "#" & _
            " And A.LoanID = B.LoanID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 9)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = DepositName
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If

End Sub

Private Sub TRIALBALstartReport()

Dim DOMElement

With m_WebDoc
   
    Set m_WebTable = .All("HeaderTbl", 0)
   
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        '.innerText = "Trial Balance as on " & ToRpDate
        .innerText = GetResourceString(377) & " " & GetFromDateString(m_ToRpDate)
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = "16"
        .Align = "Center"
    End With
    
    Call StartHtmlTable("Balance", .body, 4)
    
End With

With m_WebTable
    .Width = "500"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = GetResourceString(33) '"Sl No"
        .Width = "10%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = GetResourceString(36, 35) '"Head Name"
        .Width = "50%"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = "DR " & GetResourceString(67) '"DR Amount"
        .Width = "20%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = "CR " & GetResourceString(67)  '"CR Amount "
        .Width = "20%"
        .Align = "Right"
    End With
    
End With

End Sub

Private Sub BalancingstartReport()

'Dim DOMElement

Call SetReportTitle(GetResourceString(377) & " " & GetFromDateString(ToRpDate), False)

Call StartHtmlTable("Balance", m_WebDoc.body, 8)
    
With m_WebTable
    .Width = "650"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    
    With .cells(0)
        .rowSpan = 2
        .innerText = GetResourceString(33) '"Sl No"
        .Width = "8%"
        .Align = "Center"
    End With
    
    With .cells(1)
        .rowSpan = 2
        .innerText = GetResourceString(36, 35) '"Head Name"
        .Width = "23%"
        .Align = "Left"
    End With
    
    
    With .cells(2)
        .colSpan = 2
        .innerText = "General ledger Balance"
        .Width = "23%"
        .Align = "Center"
    End With
    
    
    With .cells(3)
        .colSpan = 2
        .innerText = "Balance as per register"
        .Width = "23%"
        .Align = "Center"
    End With
    
    With .cells(4)
        .colSpan = 2
        .innerText = "DIfferece"
        .Width = "23%"
        .Align = "Center"
    End With
    
End With


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 6)
m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = "DR " & GetResourceString(67)    '"CR Amount"
        .Align = "Center"
    End With
    
    With .cells(1)
        .innerText = "CR " & GetResourceString(67)    '"CR Amount"
        .Align = "Center"
    End With
    
    With .cells(2)
        .innerText = "DR " & GetResourceString(67) '"DR Amount "
        .Align = "Center"
    End With
    With .cells(3)
        .innerText = "CR " & GetResourceString(67)  '"CR Amount "
        .Align = "Center"
    End With
    
    With .cells(4)
        .innerText = "DR " & GetResourceString(67)
        .Align = "Center"
    End With
    With .cells(5)
        .innerText = "CR " & GetResourceString(67)
        .Align = "Center"
    End With
    
End With

End Sub

Private Sub CashBooKStartReport(NewBook As Boolean)

'Dim DOMElement
Dim rst As Recordset

Call SetReportTitle(GetResourceString(419) & " " & GetFromDateString(m_ToRpDate), False)

Call StartHtmlTable("CashBook", m_WebDoc.body, IIf(NewBook, 19, 7))
'Set m_WebTable = m_WebDoc.body.All("HeaderTbl", 7)
If m_WebTable Is Nothing Then Set m_WebTable = m_WebDoc.body.All("HeaderTbl", IIf(NewBook, 9, 7))

With m_WebTable
    .Width = "1100"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, IIf(NewBook, 7, 5))

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    If NewBook Then
        With .cells(0)
            .rowSpan = 2
            .innerText = GetResourceString(52) '"total"
            .Width = "7%"
            .Align = "Center"
        End With
        With .cells(1)
            .colSpan = 2
            .innerText = GetResourceString(196) '"Receipts"
            .Align = "Center"
        End With
        With .cells(2)
            .rowSpan = 2
            .innerText = GetResourceString(36) & " " & _
                                GetResourceString(60) '"AccNo"
            .Width = "5%"
            .Align = "Center"
        End With
        With .cells(3)
            .rowSpan = 2
            .innerText = GetResourceString(35) '"Name"
            .Width = "35%"
            .Align = "Center"
        End With
        
        With .cells(4)
            .rowSpan = 2
            .innerText = GetResourceString(36) & " " & _
                                GetResourceString(60) '"AccNo"
            .Width = "5%"
            If m_RpParentID Then .Width = "10%"
            .Align = "Center"
        End With
        With .cells(5)
            .colSpan = 2
            .innerText = GetResourceString(197) '"Payments"
            .Width = "7%"
            If m_RpParentID Then .Width = "10%"
            .Align = "Center"
        End With
        With .cells(6)
            .rowSpan = 2
            .innerText = GetResourceString(52) '"Total"
            '.Width = "13%"
            If m_RpParentID Then .Width = "10%"
            .Align = "Center"
        End With
    Else
        With .cells(0)
            .rowSpan = 2
            .innerText = GetResourceString(33) '"Sl No"
            .Width = "5%"
            .Align = "Center"
        End With
        
        With .cells(1)
            .rowSpan = 2
            .innerText = GetResourceString(35) '"Name"
            .Width = "35%"
            .Align = "Center"
        End With
        
        With .cells(2)
            .rowSpan = 2
            .innerText = GetResourceString(36) & " " & _
                                GetResourceString(60) '"AccNo"
            .Width = "9%"
            .Align = "Center"
        End With
        
        With .cells(3)
            .colSpan = 2
            '.innerText = GetResourceString(40) '"Cash"
            .innerText = GetResourceString(196) '"Receipts"
            '.Width = "13%"
            .Align = "Center"
        End With
        With .cells(4)
            .colSpan = 2
            .innerText = GetResourceString(197) '"Payments"
            .Width = "15%"
            If m_RpParentID Then .Width = "15%"
            .Align = "Center"
        End With
    End If
    
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    
    With .cells(0)
        '.innerText = GetResourceString(197) '"Payments"
        .innerText = GetResourceString(270) '"contra"
        .Width = "13%"
        If m_RpParentID Then .Width = "15%"
        .Align = "Center"
    End With
    With .cells(1)
        '.innerText = GetResourceString(196) '"Receipts"
        .innerText = GetResourceString(40) '"Cash"
        .Width = "13%"
        .Align = "Center"
    End With
    
    With .cells(2)
        .innerText = GetResourceString(196) '"Receipts"
        .innerText = GetResourceString(40) '"Cash"
        .Width = "13%"
        .Align = "Center"
    End With
    
    With .cells(3)
        '.innerText = GetResourceString(197) '"Payments"
        .innerText = GetResourceString(270) '"contra"
        .Width = "13%"
        .Align = "Center"
    End With
    
End With

End Sub

Private Sub GLStartReport()

'Dim DOMElement
Dim strText As String
Dim rst As Recordset
strText = ""
If RpHeadID Then
    gDbTrans.SqlStmt = "Select HeadName From Heads " & _
        " where HeadID = " & RpHeadID
    If gDbTrans.Fetch(rst, adOpenDynamic) Then strText = FormatField(rst(0))
ElseIf RpParentID Then
    gDbTrans.SqlStmt = "Select ParentName From ParentHeads " & _
        " where ParentID = " & RpParentID
    If gDbTrans.Fetch(rst, adOpenDynamic) Then strText = FormatField(rst(0))
End If

strText = strText & " " & GetResourceString(93)
Call SetReportTitle(strText)

Call StartHtmlTable("GLedger", m_WebDoc.body, 6)
    
With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = GetResourceString(33) '"Sl No"
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = GetResourceString(39) '"Particulars"
        .Width = "38%"
        If m_RpParentID Then
            .innerText = GetResourceString(37) '"Date
            .Width = "15%"
        End If
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = GetResourceString(284) '"OP Balance"
        .Width = "13%"
        If m_RpParentID Then .Width = "17%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = GetResourceString(196) '"Receipts"
        .Width = "13%"
        If m_RpParentID Then .Width = "15%"
        .Align = "Right"
    End With
    
    With .cells(4)
        .innerText = GetResourceString(197) '"Payments"
        .Width = "13%"
        If m_RpParentID Then .Width = "15%"
        .Align = "Right"
    End With
    
    With .cells(5)
        .innerText = GetResourceString(285) '"CL Balance"
        If m_RpParentID Then .Width = "17%"
        .Width = "13%"
        .Align = "Right"
    End With
End With

End Sub



Private Sub LDCheckPageLength(ByRef SlNo As Long, _
                              ByVal ClosingBalance As Currency, _
                              ByVal AccountType As wis_AccountType)
If SlNo = 56 Then

    ' This will print the Closing Balance
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 5)
    
    Call LDSetHeaderFooter(ClosingBalance, AccountType, "Amount Carried Forward")
    
    ' This will start the Header
   
    LDPrintHeader (True)
   
    Call LDSetHeaderFooter(ClosingBalance, AccountType, "Amount Brought Forward")
   
    SlNo = 0
    
End If

End Sub

Private Sub LDCreateViews(ByVal headID As Long, _
                         rstTrans As ADODB.Recordset, _
                         rstLedger As ADODB.Recordset)
                          

' Set the Sql Statement

gDbTrans.SqlStmt = " SELECT TransID,Debit,Credit,TransDate" & _
                    " FROM AccTrans " & _
                    " WHERE HeadID = " & headID & _
                    " AND TransDate BETWEEN #" & m_FromUSDate & "#" & _
                    " AND #" & m_ToUSDate & "#" & _
                    " ORDER BY TransDate,TransID"

If Not gDbTrans.CreateView("QryLedger") Then Exit Sub
    
gDbTrans.SqlStmt = " SELECT * " & _
                  " FROM QryLedger"
                
'Fetch the Data to Recordset based on the Condition
Call gDbTrans.Fetch(rstTrans, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT a.*,HeadName" & _
                    " FROM AccTrans a, QryLedger b,Heads c " & _
                    " WHERE a.HeadID <> " & headID & _
                    " AND a.TransID=b.TransID" & _
                    " AND a.HeadID=c.HeadID"
                    
Call gDbTrans.Fetch(rstLedger, adOpenStatic)

End Sub

Private Sub LDSetRowDetails(ByVal TransDate As Date, _
                           ByVal headName As String, _
                           ByVal VoucherName As String, _
                           ByVal Debit As Currency, _
                           ByVal Credit As Currency, _
                           ByVal ClosingBalance As Currency)

'Chane by string to date (TransDate)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow.cells(0)
    .innerText = GetIndianDate(TransDate)
    .Align = "Right"
End With
With m_WebRow.cells(1)
    .innerText = headName
    .Align = "Left"
End With
'With m_WebRow.cells(2)
'    .innerText = VoucherName
'    .Align = "Left"
'End With

If Debit > 0 Then
    With m_WebRow.cells(2)
        .innerText = FormatCurrency(Debit)
        .Align = "Right"
    End With
End If

If Credit > 0 Then
    With m_WebRow.cells(3)
        .innerText = FormatCurrency(Credit)
        .Align = "Right"
    End With
End If

With m_WebRow.cells(4)
    .innerText = FormatCurrency(ClosingBalance)
    .Align = "Right"
End With

End Sub

Private Sub LoadFrmWeb()

If Not m_frmWeb Is Nothing Then Exit Sub

Set m_frmWeb = New frmWeb
Load m_frmWeb

End Sub

Private Sub SubDaystartReport()

Dim DOMElement

With m_WebDoc

    Set m_WebTable = .body.All("HeaderTbl")
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        '.innerText = "Sub Day Book from " & FromRpDate & " to " & ToRpDate
        .innerText = GetResourceString(390, 63) & " " & _
                        GetFromDateString(FromRpDate, m_ToRpDate)
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    
    Call StartHtmlTable("SubDayBook", .body, 4)
    
End With

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = GetResourceString(33) '"Sl No"
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = GetResourceString(39) '"Particulars"
        .Width = "45%"
        .Align = "Left"
    End With
    
    
    With .cells(2)
        .innerText = GetResourceString(196) ' "Receipts"
        .Width = "12%"
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = GetResourceString(197) '"Payments"
        .Width = "12%"
        .Align = "Right"
    End With
    
    
End With

End Sub

Private Sub DBStartReport()

Call SetReportTitle(GetResourceString(390, 63) & " " & GetFromDateString(m_ToRpDate), False) '"Day Book For : " & ToRpDate
Call StartHtmlTable("DayBook", m_WebDoc.body, 5)

With m_WebTable
    .Width = "700"
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 6)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow
    
    With .cells(0)
        .innerText = GetResourceString(33)
        .Width = "7%"
        .Align = "Right"
    End With
    
    With .cells(1)
        .innerText = GetResourceString(37)
        .Width = "12%"
        .Align = "Right"
    End With
    
    
    With .cells(2)
        .innerText = GetResourceString(39)
        .Width = "45%"
        .Align = "Left"
    End With
    
    
    With .cells(3)
        .innerText = GetResourceString(272) '"Debit Amt"
        .Width = "12%"
        .Align = "Right"
    End With
    
    With .cells(4)
        .innerText = GetResourceString(271) '"Credit Amt"
        .Width = "12%"
        .Align = "Right"
    End With
    
    
End With

End Sub

' This sub will process the selected report in the from
Public Function ProcessSelectedReport() As Boolean

Dim newReport As Boolean

newReport = CBool(GetConfigValue("NewReport"))

' Handle Error
On Error GoTo Hell:

' Declarations
ProcessSelectedReport = False
If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb
Load m_frmWeb
' Selection
Select Case m_AccountReportList
    Case AccountLedger
        If Not ShowLedgerToHtml Then Exit Function
     
    Case AccountLedgerOnDate
        If Not ShowLedgerToHtml Then Exit Function
    
    Case AccountsClosed
    
    Case BalancesAsON
        If Not ShowBalancesAsON Then Exit Function
        
    Case DayBook
        If Not ShowDayBook Then Exit Function
        
    Case SubDayBook
        If Not ShowSubDayBook Then Exit Function
        
    Case GeneralLedger
        If Not ShowGeneralLedger(newReport) Then Exit Function
    
    Case AccountGeneralLedger
        If Not ShowAccountGeneralLedger Then Exit Function
        
    Case ProfitandLossTrans
    
    Case ReportNothing
    
    Case TotalTransActionsMade
    
End Select

ProcessSelectedReport = True
Exit Function

Hell:
    If Err Then MsgBox "ProcessSelectedReport : " & vbCrLf & Err.Description
   
End Function

' This function will show the Transactions of the day selected
' This will read torpDate property
Private Function ShowDayBook() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstDayBook As ADODB.Recordset

Dim headID As Long
Dim SlNo As Long
Dim TransID As Long

Dim headName As String
Dim TransDate As String
Dim VoucherName As String

Dim Debit As Currency
Dim Credit As Currency

Dim VoucherType As Wis_VoucherTypes

Dim LedgerClass As clsLedger

' Intialise the Function
ShowDayBook = False

' Check the Inputs
If ToRpDate = "" Then Exit Function

' Set the Sql Stmt

gDbTrans.SqlStmt = "SELECT a.TransID, a.HeadID, a.Debit, a.Credit, a.TransDate, " & _
                  " a.VoucherType, b.HeadName " & _
                  " FROM AccTrans a, Heads b " & _
                  " WHERE Transdate = # " & m_ToUSDate & "#" & _
                  " and a.HeadID=b.HeadID and a.headid <> " & wis_CashHeadID & _
                  " ORDER BY a.TransID"

' Fetch Recordset
If gDbTrans.Fetch(rstDayBook, adOpenForwardOnly) < 1 Then Exit Function
' Start Report and its Headers
DBStartReport

' Start Loop

Set LedgerClass = New clsLedger

SlNo = 0

Do While Not rstDayBook.EOF
    
    headName = FormatField(rstDayBook.Fields("HeadName"))
    TransDate = FormatField(rstDayBook.Fields("TransDate"))
    VoucherType = FormatField(rstDayBook.Fields("VoucherType"))
    Debit = FormatField(rstDayBook.Fields("Debit"))
    Credit = FormatField(rstDayBook.Fields("Credit"))
    TransID = FormatField(rstDayBook.Fields("TransID"))
    SlNo = SlNo + 1
    VoucherName = LedgerClass.GetVoucherName(VoucherType)
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With
    With m_WebRow.cells(1)
        .innerText = TransDate
        .Align = "Right"
    End With
    With m_WebRow.cells(2)
        .innerText = headName
        .Align = "Left"
    End With
    If Debit Then
    With m_WebRow.cells(3)
        .innerText = FormatCurrency(Debit)
        .Align = "Right"
    End With
    End If
    If Credit > 0 Then
        With m_WebRow.cells(4)
            .innerText = FormatCurrency(Credit)
            .Align = "Right"
        End With
    End If
    
    If Credit = 0.001 Then
        With m_WebRow.cells(5)
            .innerText = FormatCurrency(Credit)
            .Align = "Right"
        End With
    End If
    
    Do
        If rstDayBook.EOF Then Exit Do
        If FormatField(rstDayBook.Fields("TransID")) <> TransID Then Exit Do
        rstDayBook.MoveNext
    Loop
Loop

Set LedgerClass = Nothing
Set rstDayBook = Nothing

ShowDayBook = True

Exit Function

Hell:

    If Err Then MsgBox "DayBook : " & vbCrLf & Err.Description
    
    Set LedgerClass = Nothing
    Set rstDayBook = Nothing
    
End Function

Private Sub IntialiseWebDocument()

If m_frmWeb Is Nothing Then Exit Sub

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

m_WebDoc.body.Style.fontFamily = gFontName

SetCompanyTitle

End Sub


Private Sub LDPrintHeader(isInsertPageBreak As Boolean)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

m_WebRow.Style.fontWeight = "Bold"

If isInsertPageBreak Then m_WebRow.Style.pageBreakBefore = "Always"
    
With m_WebRow
    With .cells(0)
        .innerText = GetResourceString(37) '"Date"
        .Align = "Right"
        .Width = "11%"
    End With
    With .cells(1)
        .innerText = GetResourceString(36, 35) '"Ledger Name"
        .Align = "Left"
        .Width = "35%"
    End With
'    With .cells(2)
'        .innerText = "Vouch Type"
'        .Align = "Left"
'        .Width = "14%"
'    End With
    With .cells(2)
        .innerText = GetResourceString(271) '"Debit"
        .Align = "Right"
        .Width = "14%"
    End With
    With .cells(3)
        .innerText = GetResourceString(272) '"Credit"
        .Align = "Right"
        .Width = "14%"
    End With
    With .cells(4)
        .innerText = GetResourceString(67) '"Balance"
        .Align = "Right"
        .Width = "15%"
    End With

End With

End Sub

Private Sub LDStartLedger()

Dim DOMElement
Dim headName As String

Dim AccTransClass As clsAccTrans


Set m_WebDoc = m_frmWeb.web.document

Set AccTransClass = New clsAccTrans

headName = AccTransClass.GetHeadName(RpHeadID)

Set AccTransClass = Nothing

headName = headName & " " & GetResourceString(36, 295)
Call SetReportTitle(headName)

Call StartHtmlTable("LedgerTable", m_WebDoc.body, 5)

LDPrintHeader (False)

End Sub

'This will fill the imablanced row of the Balance Sheet
Private Sub PLFillEmptyRows()

Dim CurrentRow As HTMLTableRow
Dim CurrentTable As HTMLTable

Dim ExpenseRows As Integer
Dim IncomeRows As Integer
Dim Row As Integer
Dim RowsToAdd As Integer

ExpenseRows = m_WebRow.cells(0).All.Expense.Rows.Length - 1
IncomeRows = m_WebRow.cells(1).All.Income.Rows.Length - 1

If ExpenseRows > IncomeRows Then

    RowsToAdd = ExpenseRows - IncomeRows
    Set CurrentTable = m_WebRow.cells(1).All.Income
    
Else
    RowsToAdd = IncomeRows - ExpenseRows
    Set CurrentTable = m_WebRow.cells(0).All.Expense

End If

For Row = 1 To RowsToAdd

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(m_OtherDetail, 3, 2))

Next Row
    
End Sub

'This function will show the balances of the specified heads to the web control
Private Function ShowBalancesAsON() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstBalances As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstOpBalance As ADODB.Recordset

Dim AccTransClass As clsAccTrans
Dim DRCRType As wis_DrCrType

Dim Balance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim opBalance As Currency

Dim headID As Long
Dim ParentID As Long
Dim PrevParentID As Long

Dim SlNo As Long

Dim headName As String
Dim ParentName As String
Dim fromDate As Date
Dim toDate As Date

Dim fldHeadID As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldParentName As ADODB.Field
Dim fldParentID As ADODB.Field
Dim fldAccountType As ADODB.Field

Dim isHeadID As Boolean

Dim AccountType As wis_AccountType

' Intialise the Function
ShowBalancesAsON = False

Set AccTransClass = New clsAccTrans

' Start the report
BALONstartReport

ParentID = RpParentID
headID = RpHeadID

fromDate = m_FromUSDate
toDate = m_ToUSDate

' Set the Sql
If RpHeadID = 0 Then

    gDbTrans.SqlStmt = " SELECT HeadID,OpAmount" & _
                       " FROM OpBalance" & _
                       " WHERE OpDate=" & "#" & FinUSFromDate & "#"
                     
    Call gDbTrans.Fetch(rstOpBalance, adOpenDynamic)
    
    gDbTrans.SqlStmt = " SELECT SUM(a.Debit) as DebitTotal, " & _
                       " SUM(a.Credit) as CreditTotal,HeadId" & _
                       " FROM AccTrans a" & _
                       " WHERE TransDate BETWEEN #" & FinUSFromDate & "#" & _
                       " AND #" & toDate & "#" & _
                       " GROUP BY HeadID"
                       
    Call gDbTrans.Fetch(rstTrans, adOpenDynamic)
    
    isHeadID = False
    
End If

If RpHeadID > 0 Then isHeadID = True

If RpParentID = 0 And RpHeadID = 0 Then _
    gDbTrans.SqlStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.ParentID=b.ParentID" & _
                       " ORDER BY b.ParentID"

If RpParentID > 0 And RpHeadID = 0 Then _
    gDbTrans.SqlStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.ParentID=b.ParentID" & _
                       " AND b.ParentID=" & ParentID

If RpParentID > 0 And RpHeadID > 0 Then _
    gDbTrans.SqlStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                       " FROM Heads a,ParentHeads b" & _
                       " WHERE a.HeadID=" & headID & _
                       " AND a.ParentID=b.ParentID"

' Get the Recordset

Call gDbTrans.Fetch(rstBalances, adOpenForwardOnly)

Set fldHeadID = rstBalances.Fields("HeadID")
Set fldHeadName = rstBalances.Fields("HeadName")
Set fldParentName = rstBalances.Fields("ParentName")
Set fldParentID = rstBalances.Fields("ParentID")
Set fldAccountType = rstBalances.Fields("AccountType")

' Start the Loop

SlNo = 0
Do While Not rstBalances.EOF
    
    headID = fldHeadID.Value
    headName = fldHeadName.Value
    ParentName = fldParentName.Value
    ParentID = fldParentID.Value
    AccountType = fldAccountType.Value
    
    'fill the parentname
    If PrevParentID <> ParentID Then
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        m_WebRow.Style.fontWeight = "Bold"
        m_WebRow.cells(1).innerText = ParentName
        PrevParentID = ParentID
    End If
    
    If Not isHeadID Then
        opBalance = 0
        If gDbTrans.FindRecord(rstOpBalance, "HeadID = " & headID) Then _
                             opBalance = rstOpBalance.Fields("OpAmount")
        'rstOpBalance.Find "HeadID = " & HeadId, , adSearchForward
        DebitTotal = 0
        CreditTotal = 0
    
        If gDbTrans.FindRecord(rstTrans, "HeadID = " & headID) Then
            DebitTotal = FormatField(rstTrans.Fields("DebitTotal"))
            CreditTotal = FormatField(rstTrans.Fields("CreditTotal"))
        End If
                    
        Balance = AccTransClass.GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)

    Else
        Balance = AccTransClass.GetOpBalance(headID, m_ToUSDate)
    End If
    
    If Balance = 0 Then GoTo NextID
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

    SlNo = SlNo + 1

    With m_WebRow
        With .cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With

        With .cells(1)
            .innerText = headName
            .Align = "Left"
        End With

        With .cells(2)
            .innerText = FormatCurrency(Abs(Balance))
            .Align = "Right"
        End With
        
        DRCRType = GetDRCRType(Balance, AccountType)
        
        If DRCRType = enumDebit Then
            With .cells(3)
                .innerText = "DR"
                .Align = "Right"
            End With
        End If
        
        If DRCRType = enumCredit Then
            With .cells(3)
                .innerText = "CR"
                .Align = "Right"
            End With
        End If
        
    End With
        
NextID:
    rstBalances.MoveNext
Loop


' Exit function
Set AccTransClass = Nothing

Set rstBalances = Nothing
Set rstTrans = Nothing
Set rstOpBalance = Nothing

ShowBalancesAsON = True

Exit Function

Hell:

    If Err Then MsgBox "ShowBalancesAsOn : " & vbCrLf & Err.Description
    
    Set AccTransClass = Nothing
    
    Set rstBalances = Nothing
    Set rstTrans = Nothing
    Set rstOpBalance = Nothing

    Exit Function
    
End Function
'This function will show the balances of the specified heads to the web control
Public Function ShowTrialBalance() As Boolean

' Handle the Error
On Error GoTo Hell:

' Declarations
Dim rstBalances As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstOpBalance As ADODB.Recordset

Dim AccTransClass As clsAccTrans
'Dim MaterialClass As clsMaterial

Dim Balance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim StockValue As Currency
Dim GrDebitTotal As Currency
Dim GrCreditTotal As Currency

Dim opBalance As Currency

Dim headID As Long
Dim ParentID As Long
Dim PrevParentID As Long

Dim SlNo As Long

Dim headName As String
Dim ParentName As String

Dim DRCRType As wis_DrCrType

Dim fldHeadID As ADODB.Field
Dim fldHeadName As ADODB.Field
Dim fldParentName As ADODB.Field
Dim fldParentID As ADODB.Field
Dim fldAccountType As ADODB.Field

Dim isHeadID As Boolean

Dim AccountType As wis_AccountType

' Intialise the Function
ShowTrialBalance = False

Set AccTransClass = New clsAccTrans

' Start the report
TRIALBALstartReport

ParentID = RpParentID
headID = RpHeadID



gDbTrans.SqlStmt = " SELECT HeadID,OpAmount" & _
                   " FROM OpBalance" & _
                   " WHERE OpDate=" & "#" & m_FromUSDate & "#"
Call gDbTrans.Fetch(rstOpBalance, adOpenDynamic)

gDbTrans.SqlStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal," & _
                   " HeadId  FROM AccTrans a" & _
                   " WHERE TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " GROUP BY HeadID"
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)

gDbTrans.SqlStmt = " SELECT a.HeadID,a.HeadName,b.ParentID,b.ParentName,AccountType" & _
                   " FROM Heads a,ParentHeads b" & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " And b.ParentID <> " & parProfitORLoss & _
                   " ORDER BY b.ParentID"

' Get the Recordset
Call gDbTrans.Fetch(rstBalances, adOpenForwardOnly)

Set fldHeadID = rstBalances.Fields("HeadID")
Set fldHeadName = rstBalances.Fields("HeadName")
Set fldParentName = rstBalances.Fields("ParentName")
Set fldParentID = rstBalances.Fields("ParentID")
Set fldAccountType = rstBalances.Fields("AccountType")

' Start the Loop

SlNo = 0
Do While Not rstBalances.EOF
    
    headID = fldHeadID.Value
    headName = fldHeadName.Value
    ParentName = fldParentName.Value
    ParentID = fldParentID.Value
    AccountType = fldAccountType.Value
    
    ' fill the parentname
    If PrevParentID <> ParentID Then
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
        m_WebRow.Style.fontWeight = "Bold"
        m_WebRow.cells(1).innerText = ParentName
        PrevParentID = ParentID
    End If
    
    opBalance = 0
    If gDbTrans.FindRecord(rstOpBalance, "HeadID=" & headID) Then _
                         opBalance = rstOpBalance.Fields("OpAmount")
                    
    DebitTotal = 0
    CreditTotal = 0
    
    If gDbTrans.FindRecord(rstTrans, "HeadID=" & headID) Then
        DebitTotal = FormatField(rstTrans.Fields("DebitTotal"))
        CreditTotal = FormatField(rstTrans.Fields("CreditTotal"))
    End If
                
    Balance = AccTransClass.GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
    
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

    SlNo = SlNo + 1

    With m_WebRow
        With .cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With

        With .cells(1)
            .innerText = headName
            .Align = "Left"
        End With
        
        
        DRCRType = GetDRCRType(Balance, AccountType)
        
        If DRCRType = enumDebit Then
        
            With .cells(2)
                .innerText = FormatCurrency(Abs(Balance))
                .Align = "Right"
            End With
            
            GrDebitTotal = GrDebitTotal + Abs(Balance)
            
        End If
        
        If DRCRType = enumCredit Then
        
            With .cells(3)
                .innerText = FormatCurrency(Abs(Balance))
                .Align = "Right"
            End With
            
            GrCreditTotal = GrCreditTotal + Abs(Balance)
            
        End If
        
    End With
        
    rstBalances.MoveNext
Loop

Dim MaterialClass As clsMaterial
Set MaterialClass = New clsMaterial

StockValue = MaterialClass.GetOnDateClosingStockValue(m_ToUSDate)
If StockValue = 0 Then MaterialClass.GetOpeningStockValue (1)
Set MaterialClass = Nothing

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

SlNo = SlNo + 1

With m_WebRow

    With .cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = GetResourceString(175) '"Stock In Hand"
        .Align = "Left"
    End With
    
    DRCRType = GetDRCRType(StockValue, AccountType)
    
    If DRCRType = enumDebit Then
    
        With .cells(2)
            .innerText = FormatCurrency(Abs(StockValue))
            .Align = "Right"
        End With
        
        GrDebitTotal = GrDebitTotal + Abs(StockValue)
        
    End If
    
    If DRCRType = enumCredit Then
    
        With .cells(3)
            .innerText = FormatCurrency(Abs(StockValue))
            .Align = "Right"
        End With
        
        GrCreditTotal = GrCreditTotal + Abs(StockValue)
        
    End If
    
End With

'not to show profit & Loss here
If 2 < 1 Then
''Put the Profit/Loss
'Now Get the Last Year Profit Or Loss
Dim PrevProfit As Currency
Dim Profit As Currency
'Now Get the Current Yer Profit
Call BALCreateViewsForConditionalBalanceSheet
PrevProfit = GetPreviousYearProfit
Profit = GetNetProfit + PrevProfit
SlNo = SlNo + 1

With m_WebRow

    With .cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = GetResourceString(332, IIf(Profit > 0, 403, 404)) '"Profit/Loss"
        .Align = "Left"
    End With
    
    DRCRType = IIf(Profit < 0, wis_DrCrType.enumDebit, wis_DrCrType.enumCredit)
    
    If DRCRType = enumDebit Then
    
        With .cells(2)
            .innerText = FormatCurrency(Abs(Profit))
            .Align = "Right"
        End With
        
        GrDebitTotal = GrDebitTotal + Abs(Profit)
        
    End If
    
    If DRCRType = enumCredit Then
    
        With .cells(3)
            .innerText = FormatCurrency(Abs(Profit))
            .Align = "Right"
        End With
        
        GrCreditTotal = GrCreditTotal + Abs(Profit)
        
    End If
    
End With
End If


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

Dim DiffAmount As Currency

DiffAmount = GrCreditTotal - GrDebitTotal

If DiffAmount Then
    
    If DiffAmount > 0 Then
        GrDebitTotal = GrDebitTotal + DiffAmount
    Else
        GrCreditTotal = GrCreditTotal + Abs(DiffAmount)
    End If
    
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    With m_WebRow
        .Style.fontWeight = "Bold"
        
        With .cells(0)
            .innerText = ""
            .Align = "Right"
        End With
    
        With .cells(1)
            .innerText = "Difference Amount"
            .Align = "Left"
        End With
        
        With .cells(2)
            .innerText = IIf(DiffAmount > 0, FormatCurrency(DiffAmount), "")
            .Align = "Right"
        End With
        
        With .cells(3)
            .innerText = IIf(DiffAmount > 0, "", FormatCurrency(DiffAmount))
            .Align = "Right"
        End With
    End With
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
End If

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow

    .Style.fontWeight = "Bold"
    
    With .cells(0)
        .innerText = ""
        .Align = "Right"
    End With

    With .cells(1)
        .innerText = GetResourceString(286) '"Grand Total"
        .Align = "Left"
    End With
    
    With .cells(2)
        .innerText = FormatCurrency(GrDebitTotal)
        .Align = "Right"
    End With
    
    With .cells(3)
        .innerText = FormatCurrency(GrCreditTotal)
        .Align = "Right"
    End With
End With


' Exit function
Set AccTransClass = Nothing

Set rstBalances = Nothing
Set rstTrans = Nothing
Set rstOpBalance = Nothing

ShowTrialBalance = True

Exit Function

Hell:

    If Err Then MsgBox "ShowTrialBalance : " & vbCrLf & Err.Description
    
    Set AccTransClass = Nothing
    
    Set rstBalances = Nothing
    Set rstTrans = Nothing
    Set rstOpBalance = Nothing

    Exit Function
    
End Function

' This will show the General Ledger for the Given Period
' This will read the torpDate,fromrpdate,rpParentid,rpHeadid properties
Private Function ShowAccountGeneralLedger() As Boolean

' Handle Error
On Error GoTo Hell:

Dim rstTrans As ADODB.Recordset

Dim Receipts As Currency
Dim Payments As Currency
Dim opBalance As Currency
Dim CLBalance As Currency
Dim SumReceipts As Currency
Dim OpeningCash As Currency
Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim headID As Long
Dim ParentID As Long

Dim AsOnDate As Date
Dim headName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans

Const TableName = "AGLedger"
Const QRYRP = "QryAccGLDetails"
Const totalCols = 6

Dim sqlSupport As String

ShowAccountGeneralLedger = False

' Validataions
If FromRpDate = "" Then Exit Function
If ToRpDate = "" Then Exit Function

' This is the date to query
AsOnDate = m_ToUSDate

gDbTrans.SqlStmt = "Select AccountType From ParentHeads " & _
                        "Where ParentID = " & RpParentID
Call gDbTrans.Fetch(rstTrans, adOpenDynamic)
AccountType = FormatField(rstTrans("AccountType"))

Set AccTransClass = New clsAccTrans
If RpHeadID Then
    opBalance = AccTransClass.GetOpBalance(RpHeadID, m_FromUSDate)
Else
    gDbTrans.SqlStmt = "Select HeadID From Heads " & _
                    " Where ParentID = " & RpParentID
    If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then Exit Function
    While Not rstTrans.EOF
        opBalance = opBalance + AccTransClass.GetOpBalance(rstTrans("HeadID"), AsOnDate)
        rstTrans.MoveNext
    Wend
End If

Dim SqlStr As String
' create views for the receipts and payments
SqlStr = "SELECT Sum(Credit) AS Receipts, SUM(Debit) AS Payments," & _
        " TransDate FROM AccTrans WHERE " & _
        " TransDate BETWEEN #" & m_FromUSDate & "#" & _
        " AND #" & m_ToUSDate & "#"

If RpHeadID Then
    SqlStr = SqlStr & " AND HeadID = " & RpHeadID
Else
    SqlStr = SqlStr & " AND HeadID IN (Select HeadID From Heads " & _
                " Where ParentID = " & RpParentID & ")"
End If

gDbTrans.SqlStmt = SqlStr & " Group By TransDate Order By TransDate"
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Exit Function

GLStartReport

Set CurrentTable = m_WebTable

' Start the Receipts and Payments
SlNo = 0

Do While Not rstTrans.EOF
    Receipts = FormatField(rstTrans("Receipts"))
    Payments = FormatField(rstTrans("Payments"))

    CLBalance = AccTransClass.GetClosingBalanceonAccountType _
                        (opBalance, Payments, Receipts, AccountType)

    SlNo = SlNo + 1
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
    With CurrentRow.cells(0)
        .innerText = SlNo
        .Align = "Right"
    End With

    With CurrentRow.cells(1)
        .innerText = FormatField(rstTrans("Transdate"))
        .Align = "Left"
    End With

    With CurrentRow.cells(2)
        .innerText = FormatCurrency(opBalance)
        .Align = "Right"
    End With

    With CurrentRow.cells(3)
        .innerText = FormatCurrency(Receipts)
        .Align = "Right"
    End With

    With CurrentRow.cells(4)
        .innerText = FormatCurrency(Payments)
        .Align = "Right"
    End With

    With CurrentRow.cells(5)
        .innerText = FormatCurrency(Abs(CLBalance))
        .Align = "Right"
    End With

    SumReceipts = SumReceipts + Receipts
    SumPayments = SumPayments + Payments
    opBalance = CLBalance
    rstTrans.MoveNext
Loop
 
Set rstTrans = Nothing
Set AccTransClass = Nothing

If SumReceipts > Receipts Or SumPayments > Payments Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
    
    With CurrentRow.cells(1)
        .innerText = GetResourceString(286)
        .Style.fontWeight = "Bold"
        .Align = "Right"
    End With

    With CurrentRow.cells(3)
        .innerText = FormatCurrency(SumReceipts)
        .Style.fontWeight = "Bold"
        .Align = "Right"
    End With

    With CurrentRow.cells(4)
        .innerText = FormatCurrency(SumPayments)
        .Style.fontWeight = "Bold"
        .Align = "Right"
    End With

End If

ShowAccountGeneralLedger = True

Exit Function

Hell:

    MsgBox "ShowAccountGeneralLedger: " & vbCrLf & Err.Description
    'Resume
    Set rstTrans = Nothing
    Set AccTransClass = Nothing
    
End Function

' This will show the General Ledger for the Given Period
' This will read the torpDate,fromrpdate,rpParentid,rpHeadid properties

Public Function ShowGeneralLedger(NewBook As Boolean) As Boolean

' Handle Error
On Error GoTo Hell:

Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset
Dim rstRPNoTrans As ADODB.Recordset

Dim Receipts As Currency
Dim Payments As Currency
Dim opBalance As Currency
Dim CLBalance As Currency
Dim SumReceipts As Currency
Dim OpeningCash As Currency
Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim headID As Long
Dim ParentID As Long

Dim headName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans

Const TableName = "GLedger"
Const QRYRP = "QryRPDetails"
Const totalCols = 6

ShowGeneralLedger = False

' Validataions
If FromRpDate = "" Then Exit Function
If ToRpDate = "" Then Exit Function

' This is the date to query
'AsOnDate = ToRpDate

' Start the Receipts and Payments
If RpParentID Then
    ShowGeneralLedger = ShowAccountGeneralLedger
    Exit Function
End If

GLStartReport

' Put the Opening Cash to the Html
Set AccTransClass = New clsAccTrans

OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, IIf(NewBook, m_FromUSDate, m_ToUSDate))

Call PutAmountInBold(TableName, ToRpDate, 2, totalCols, GetResourceString(284), OpeningCash)

' create views for the receipts and payments
Call GLCreateViews(QRYRP, NewBook)

gDbTrans.SqlStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName" & _
                   " FROM " & QRYRP & " AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"

'' Create a View of this transactions
gDbTrans.CreateView ("qryGenLedger")

gDbTrans.SqlStmt = "select * from qryGenLedger"
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)


gDbTrans.SqlStmt = " SELECT b.ParentID, a.HeadID as RpHeadid,b.HeadName" & _
                   " FROM OpBalance AS a, Heads AS b" & _
                   " WHERE a.HeadId = b.HeadID and a.HeadID not in (Select distinct RPHeadID from qryGenLedger)" & _
                   " and a.HeadID <> " & wis_CashHeadID & " and a.OpDate = #" & FinUSFromDate & "# and opAmount > 0" & _
                   " GROUP BY b.ParentID,a.HeadID,b.HeadName"
                   
Call gDbTrans.Fetch(rstRPNoTrans, adOpenDynamic)

Set CurrentTable = m_WebTable
Set AccTransClass = New clsAccTrans

SlNo = 0

' Main Loop
Do While Not rstParents.EOF

    ParentID = rstParents.Fields("ParentID")
    ParentName = rstParents.Fields("ParentName")
    AccountType = rstParents.Fields("AccountType")
    
    With rstRPNoTrans
        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID
    End With
    
    With rstRPDetails

        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID

        If rstRPDetails.RecordCount > 0 Or rstRPNoTrans.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)

            CurrentRow.Style.fontWeight = "Bold"

            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
            End With
        End If

        Do While Not .EOF

            headID = FormatField(.Fields("RPHeadID"))
            headName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
 
            opBalance = AccTransClass.GetOpBalance(headID, IIf(NewBook, m_FromUSDate, m_ToUSDate))
            
            CLBalance = AccTransClass.GetClosingBalanceonAccountType(opBalance, Payments, Receipts, AccountType)

            SlNo = SlNo + 1
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
            With CurrentRow.cells(0)
                .innerText = SlNo
                .Align = "Right"
            End With

            With CurrentRow.cells(1)
                .innerText = headName
                .Align = "Left"
            End With

            With CurrentRow.cells(2)
                .innerText = FormatCurrency(opBalance)
                .Align = "Right"
            End With

            With CurrentRow.cells(3)
                .innerText = FormatCurrency(Receipts)
                .Align = "Right"
            End With

            With CurrentRow.cells(4)
                .innerText = FormatCurrency(Payments)
                .Align = "Right"
            End With

            With CurrentRow.cells(5)
                .innerText = FormatCurrency(Abs(CLBalance))
                .Align = "Right"
            End With

            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments

            .MoveNext
        Loop
    End With
    
    ''Put the Heads which has not made no tranasctions
    With rstRPNoTrans
        Do While Not .EOF

            headID = FormatField(.Fields("RPHeadID"))
            headName = FormatField(.Fields("HeadName"))
            
            opBalance = AccTransClass.GetOpBalance(headID, IIf(NewBook, m_FromUSDate, m_ToUSDate))
            
            SlNo = SlNo + 1
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
            With CurrentRow.cells(0)
                .innerText = SlNo
                .Align = "Right"
            End With

            With CurrentRow.cells(1)
                .innerText = headName
                .Align = "Left"
            End With

            With CurrentRow.cells(2)
                .innerText = FormatCurrency(opBalance)
                .Align = "Right"
            End With

            With CurrentRow.cells(3)
                .innerText = "0.00"
                .Align = "Right"
            End With

            With CurrentRow.cells(4)
                .innerText = "0.00"
                .Align = "Right"
            End With

            With CurrentRow.cells(5)
                .innerText = FormatCurrency(Abs(opBalance))
                .Align = "Right"
            End With
            .MoveNext
        Loop
    End With
    rstParents.MoveNext

Loop

' Now Put Closing Balance
Call PutAmountInBold(TableName, ToRpDate, totalCols - 1, totalCols, GetResourceString(285), (OpeningCash + SumReceipts) - SumPayments)
 
Set rstRPDetails = Nothing
Set rstParents = Nothing
Set AccTransClass = Nothing

ShowGeneralLedger = True

Exit Function

Hell:

    MsgBox "ShowGeneralLedger: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    Set AccTransClass = Nothing
    
End Function



' This will show the General Ledger for the Given Period
' This will read the torpDate,fromrpdate,rpParentid,rpHeadid properties

Private Function ShowSubDayBook() As Boolean

Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset

Dim Receipts As Currency
Dim Payments As Currency
Dim opBalance As Currency
Dim CLBalance As Currency
Dim SumReceipts As Currency
Dim OpeningCash As Currency
Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim headID As Long
Dim ParentID As Long

Dim AsonIndianDate As String
Dim headName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans

Const TableName = "SubDayBook"
Const QRYRP = "QryRPDetails"
Const totalCols = 6

' Handle Error
On Error GoTo Hell:

ShowSubDayBook = False

' Validataions
If FromRpDate = "" Then Exit Function
If ToRpDate = "" Then Exit Function
If RpParentID = 0 Then Exit Function

' This is the date to query
AsonIndianDate = ToRpDate

' Start the Receipts and Payments
SubDaystartReport

' Put the Opening Cash to the Html

OpeningCash = RPPutOpeningCash(TableName, AsonIndianDate)

' create views for the receipts and payments
SubDayCreateViews (QRYRP)

gDbTrans.SqlStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName" & _
                   " FROM " & QRYRP & " a, Heads b ,ParentHeads c" & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " AND b.ParentID=c.ParentID" & _
                   " AND c.ParentID=" & RpParentID & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"
                   
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable
Set AccTransClass = New clsAccTrans

SlNo = 0

' Main Loop
Do While Not rstParents.EOF

    ParentID = rstParents.Fields("ParentID")
    ParentName = rstParents.Fields("ParentName")
    AccountType = rstParents.Fields("AccountType")
    
    With rstRPDetails

        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID

        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

            CurrentRow.Style.fontWeight = "Bold"

            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
            End With
        End If

        Do While Not .EOF

            headID = FormatField(.Fields("RPHeadID"))
            headName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
 
            SlNo = SlNo + 1
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)
            With CurrentRow.cells(0)
                .innerText = SlNo
                .Align = "Right"
            End With

            With CurrentRow.cells(1)
                .innerText = headName
                .Align = "Left"
            End With

            With CurrentRow.cells(2)
                .innerText = FormatCurrency(Receipts)
                .Align = "Right"
            End With

            With CurrentRow.cells(3)
                .innerText = FormatCurrency(Payments)
                .Align = "Right"
            End With

            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments

            .MoveNext
        Loop
    End With

    rstParents.MoveNext

Loop

' Now Put Closing Balance
Call RPputClosingCash((OpeningCash + SumReceipts) - SumPayments, TableName)

Set rstRPDetails = Nothing
Set rstParents = Nothing

ShowSubDayBook = True

Exit Function

Hell:
    
    MsgBox "ShowSubDayBook: " & vbCrLf & Err.Description
    'Resume
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    
End Function

Public Sub SetCompanyTitle()

Dim rstCompany As ADODB.Recordset
Dim eCompanyType As wis_CompanyType

Dim DOMElement

eCompanyType = Enum_Self

gDbTrans.SqlStmt = " SELECT CompanyName,Address " & _
                   " FROM CompanyCreation " & _
                   " WHERE CompanyType = " & eCompanyType

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 1 Then Exit Sub

If m_frmWeb Is Nothing Then Set m_frmWeb = New frmWeb
Load m_frmWeb
If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    
    Call StartHtmlTable("HeaderTbl", m_WebDoc.body, 1)
    
    m_WebTable.border = 0
        
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    
    With m_WebRow.cells(0)
        .innerText = rstCompany.Fields("CompanyName").Value
        .Style.FONTSIZE = IIf(gLangOffSet, "25", "23")
        .Align = "Center"
        .Style.fontWeight = "Bold"
    End With
    If Len(rstCompany.Fields("Address")) Then
        If Len(rstCompany("CompanyName")) Then Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    With m_WebRow.cells(0)
        .innerText = FormatField(rstCompany.Fields("Address"))
        .Style.FONTSIZE = IIf(gLangOffSet, "19", "17")
        .Align = "Center"
        If Len(.innerText) Then Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)
    End With
    End If
End With

End Sub

' This fetch the data from Base table and load it to web control
'
Private Function ShowLedgerToHtml() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstLedger As ADODB.Recordset
Dim rstTrans As ADODB.Recordset

Dim headName As String
Dim TransDate As String
Dim VoucherName As String

Dim AccountType As wis_AccountType
Dim VoucherType As Wis_VoucherTypes

Dim Debit As Currency
Dim Credit As Currency
Dim opBalance As Currency
Dim DebitTotal As Currency
Dim HeadIDDebit As Currency
Dim HeadIDCredit As Currency
Dim CreditTotal As Currency
Dim ClosingBalance As Currency

Dim SlNo As Long
Dim headID As Long
Dim TransID As Long
Dim ToHeadID As Long

Dim VoucherNames

Dim LedgerClass As clsLedger
Dim AccTransClass As clsAccTrans

ShowLedgerToHtml = False

' Validations
If RpHeadID = 0 Then Exit Function

headID = RpHeadID

Set AccTransClass = New clsAccTrans
Set LedgerClass = New clsLedger

VoucherNames = AccTransClass.GetVoucherNametoArray

LDStartLedger
 
AccountType = AccTransClass.GetAccountType(headID)

opBalance = AccTransClass.GetOpBalance(headID, FinUSFromDate)

Call LDSetHeaderFooter(opBalance, AccountType, GetResourceString(284)) ' "Opening Balance")

Call LDCreateViews(headID, rstTrans, rstLedger)

SlNo = 0

DebitTotal = 0
CreditTotal = 0

Do While Not rstTrans.EOF
    
    TransID = rstTrans.Fields("TransID")
    
    HeadIDDebit = rstTrans.Fields("Debit")
    HeadIDCredit = rstTrans.Fields("Credit")
    
    rstLedger.Filter = adFilterNone
    If HeadIDDebit > 0 Then
        rstLedger.Filter = "TransID=" & TransID & " AND Credit > " & 0
    Else
        rstLedger.Filter = "TransID=" & TransID & " AND Debit > " & 0
    End If
    
    Do While Not rstLedger.EOF
            
        ' This will check the page and reallocate the slno
        Call LDCheckPageLength(SlNo, ClosingBalance, AccountType)
    
        Debit = 0
        Credit = 0
    
        If HeadIDDebit > 0 Then
            If HeadIDDebit < FormatField(rstLedger.Fields("Credit")) Then
                Debit = HeadIDDebit
            Else
                Debit = FormatField(rstLedger.Fields("Credit"))
            End If
            Credit = FormatField(rstLedger.Fields("debit"))
        End If
        
        If HeadIDCredit > 0 Then
            If HeadIDCredit < FormatField(rstLedger.Fields("Debit")) Then
                Credit = HeadIDCredit
            Else
                Credit = FormatField(rstLedger.Fields("Debit"))
            End If
            Debit = FormatField(rstLedger.Fields("Credit"))
        End If
        
        Credit = rstLedger.Fields("Debit")
        ToHeadID = rstLedger.Fields("HeadID")
        headName = rstLedger.Fields("HeadName")
        TransDate = rstLedger.Fields("TransDate")
        VoucherType = rstLedger.Fields("VoucherType")
                
        VoucherName = LedgerClass.GetVoucherName(VoucherType)
        
        DebitTotal = DebitTotal + Debit
        CreditTotal = CreditTotal + Credit
        
        ClosingBalance = AccTransClass.GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
        
        SlNo = SlNo + 1
        
        Call LDSetRowDetails(TransDate, headName, VoucherName, Debit, Credit, ClosingBalance)
            
        rstLedger.MoveNext
        
    Loop
   
   
   rstTrans.MoveNext
    
Loop

ClosingBalance = AccTransClass.GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)

Call LDSetHeaderFooter(ClosingBalance, AccountType, GetResourceString(285))


Set AccTransClass = Nothing
Set LedgerClass = Nothing
Set rstLedger = Nothing
Set rstTrans = Nothing


' Validations
ShowLedgerToHtml = True
Exit Function

Hell:
    
    MsgBox "ShowLedgerToHtml:" & vbCrLf & Err.Description
    
    Set rstLedger = Nothing
    Set AccTransClass = Nothing
    Set LedgerClass = Nothing
  
End Function
Private Sub LDSetHeaderFooter(ByVal Balance As Currency, AccountType As wis_AccountType, MsgString As String)

Dim ColToPrint As Integer
Dim ColorToPrint As String

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable)

With m_WebRow
        
    .cells(1).innerText = MsgString
    .cells(1).Style.fontWeight = "Bold"
        
      Select Case AccountType
        Case Asset
        
            If Balance >= 0 Then
                ColToPrint = 2
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If
            
        Case Liability
            
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 2
                ColorToPrint = "Red"
            End If

        Case Loss
        
            If Balance >= 0 Then
                ColToPrint = 2
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If
            
        Case Profit
        
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 2
                ColorToPrint = "Red"
            End If

            
        Case ItemSales
        
            If Balance >= 0 Then
                ColToPrint = 3
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 2
                ColorToPrint = "Red"
            End If
            
        Case ItemPurchase
        
            If Balance >= 0 Then
                ColToPrint = 2
                ColorToPrint = "Black"
            End If
            
            If Balance < 0 Then
                ColToPrint = 3
                ColorToPrint = "Red"
            End If

    End Select
               
    With .cells(ColToPrint)
        .innerText = FormatCurrency(Abs(Balance))
        .Style.Color = ColorToPrint
        .Style.fontWeight = "Bold"
        .Align = "Right"
    End With
    
End With

End Sub
'
'This will return the current dr or cr balance for the accounttype
Private Function GetDRCRType(ByVal Balance As Currency, _
                              ByVal AccountType As wis_AccountType) As wis_DrCrType
                              
  Select Case AccountType
    Case Asset
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
        
    Case Liability
        
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If

    Case Loss
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
        
    Case Profit
    
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If

        
    Case ItemSales
    
        If Balance >= 0 Then
            GetDRCRType = enumCredit
        End If
          
        If Balance < 0 Then
            GetDRCRType = enumDebit
        End If
        
    Case ItemPurchase
    
        If Balance >= 0 Then
            GetDRCRType = enumDebit
        End If
        
        If Balance < 0 Then
            GetDRCRType = enumCredit
        End If
    
End Select
                   

End Function
Private Sub BALPrintGrandTotals(GrandTotal() As Currency, Optional NewBook As Boolean = False)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

PrintString = GetResourceString(286) '"Grand Total"

Set CurrentTable = m_WebRow.cells(0).All.Liability

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal(IIf(NewBook, 0, 1)))
    .Align = "Right"
End With
If NewBook Then
With CurrentRow.cells(4)
    .innerText = FormatCurrency(GrandTotal(1))
    .Align = "Right"
End With

End If

Set CurrentTable = m_WebRow.cells(1).All.Asset

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal(IIf(NewBook, 0, 1)))
    .Align = "Right"
End With
If NewBook Then
With CurrentRow.cells(4)
    .innerText = FormatCurrency(GrandTotal(1))
    .Align = "Right"
End With
End If

End Sub
Private Sub PLPrintGrandTotals(GrandTotal As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

PrintString = GetResourceString(286) '"Net  Total"

Set CurrentTable = m_WebRow.cells(0).All.Expense

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With


Set CurrentTable = m_WebRow.cells(1).All.Income

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(m_OtherDetail, 3, 2))

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(m_OtherDetail, 3, 2))
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(IIf(m_OtherDetail, 2, 1))
    .innerText = FormatCurrency(GrandTotal)
    .Align = "Right"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

Private Sub BALPrintProfitorLoss(Profit As Currency)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim PrintString As String

If Profit >= 0 Then
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    PrintString = GetResourceString(403) '"Profit"
End If

If Profit < 0 Then
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
    PrintString = GetResourceString(404) '"Loss"
End If
    
CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(Abs(Profit))
    .Align = "Right"
End With

End Sub
Private Sub PrintBalProfitorLoss_1(PreviousProfit As Currency, CurrentProfit As Currency, Optional NewBook As Boolean = False)



Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow
Dim prevCurrentTable As HTMLTable
Dim prevParentRow As HTMLTableRow

Dim PrintString As String
Dim NetProfit As Currency

NetProfit = PreviousProfit + CurrentProfit

    If NetProfit >= 0 Then
        Set CurrentTable = m_WebRow.cells(0).All.Liability
        PrintString = GetResourceString(332, 403) '"Profit"
    End If
    
    If NetProfit < 0 Then
        Set CurrentTable = m_WebRow.cells(1).All.Asset
        PrintString = GetResourceString(332, 404) '"Loss"
    End If
    ' Start Putting Parentid's heads
    Set ParentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

ParentRow.Style.fontWeight = "bold"
With ParentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With
With ParentRow.cells(IIf(NewBook, 4, 2))
    .innerText = FormatCurrency(Abs(NetProfit))
    .Align = "Right"
    .Style.fontWeight = "bold"
End With
If NewBook Then
    With ParentRow.cells(2)
        .innerText = FormatCurrency(Abs(PreviousProfit))
        .Align = "Right"
        .Style.fontWeight = "bold"
    End With
End If

If PreviousProfit Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
    
    PrintString = GetResourceString(250, 251)
    PrintString = PrintString & " " & GetResourceString(IIf(PreviousProfit > 0, 403, 404))
    'printstring=previous year's profit
    'CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(0)
        .innerText = PrintString
        .Align = "Left"
    End With
    
    With CurrentRow.cells(IIf(NewBook, 1, 1))
        .innerText = FormatCurrency(PreviousProfit)
        .Align = "Right"
    End With
    If NewBook Then
        With CurrentRow.cells(2)
            '.innerText = FormatCurrency(PreviousProfit)
            .Align = "Right"
            .Style.fontWeight = "bold"
        End With
        With CurrentRow.cells(3)
            .innerText = FormatCurrency(PreviousProfit)
            .Align = "Right"
            '.Style.fontWeight = "bold"
        End With
    End If
    
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
    
    PrintString = GetResourceString(374, 251)
    PrintString = PrintString & " " & GetResourceString(IIf(PreviousProfit > 0, 403, 404))
    'printstring=Currnet year's profit
    'CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(0)
        .innerText = PrintString
        .Align = "Left"
    End With
    
    With CurrentRow.cells(IIf(NewBook, 3, 1))
        .innerText = FormatCurrency(CurrentProfit)
        .Align = "Right"
    End With

End If

End Sub



Private Sub PrintBalProfitorLoss(PreviousProfit As Currency, CurrentProfit As Currency, ProfitDist As Currency, NewBook As Boolean)

Dim assetTable As HTMLTable
Dim liabilityTable As HTMLTable

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow
Dim prevCurrentTable As HTMLTable
Dim prevParentRow As HTMLTableRow
Dim prevRow As HTMLTableRow

Dim PrintString As String
Dim prevPrintString As String
Dim NetProfit As Currency

NetProfit = PreviousProfit + CurrentProfit

Set liabilityTable = m_WebRow.cells(0).All.Liability
Set assetTable = m_WebRow.cells(1).All.Asset

If NetProfit >= 0 Then
    Set ParentRow = InsertRowAndCellSToTable(assetTable, IIf(NewBook, 5, 3))
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    PrintString = GetResourceString(332, 403) '"Profit"
Else
    Set ParentRow = InsertRowAndCellSToTable(liabilityTable, IIf(NewBook, 5, 3))
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    PrintString = GetResourceString(332, 404) '"Loss"
End If
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
ParentRow.Style.fontWeight = "bold"

With ParentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With
With ParentRow.cells(IIf(NewBook, 4, 2))
    .innerText = FormatCurrency(Abs(NetProfit))
    .Align = "Right"
    .Style.fontWeight = "bold"
End With


If PreviousProfit < 0 Then
    Set prevCurrentTable = m_WebRow.cells(1).All.Asset
    prevPrintString = GetResourceString(250, 251, 404)
Else
    Set prevCurrentTable = m_WebRow.cells(0).All.Liability
    prevPrintString = GetResourceString(250, 251, 403)
End If
Set prevRow = InsertRowAndCellSToTable(prevCurrentTable, IIf(NewBook, 5, 3))
With prevRow.cells(0)
    .innerText = prevPrintString
    .Align = "Left"
End With

With prevRow.cells(1)
    .innerText = FormatCurrency(Abs(PreviousProfit))
    .Align = "Right"
End With
If NewBook Then
    With prevRow.cells(1)
        .innerText = FormatCurrency(Abs(PreviousProfit + ProfitDist))
        .Align = "Right"
    End With
    With prevRow.cells(3)
        .innerText = FormatCurrency(Abs(PreviousProfit))
        .Align = "Right"
    End With
End If
If CurrentProfit < 0 Then
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    PrintString = GetResourceString(374, 251, 404)
Else
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    PrintString = GetResourceString(374, 251, 403)
End If
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With CurrentRow.cells(IIf(NewBook, 3, 1))
    .innerText = FormatCurrency(Abs(CurrentProfit))
    .Align = "Right"
End With


End Sub


Private Sub BALPrintClosingStockValue(ByVal StockValue As Currency, ByVal opBalance As Currency, Optional NewBook As Boolean = False)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow
Dim PrintString As String

Set CurrentTable = m_WebRow.cells(1).All.Asset
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

PrintString = GetResourceString(175) '"Stock in Hand"

ParentRow.Style.fontWeight = "Bold"

With ParentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
End With

With ParentRow.cells(2)
    .innerText = FormatCurrency(StockValue)
    .Align = "Right"
    .Style.fontWeight = "Bold"
End With
If NewBook Then
    With ParentRow.cells(2)
        .innerText = FormatCurrency(opBalance)
        .Align = "Right"
        .Style.fontWeight = "Bold"
    End With
    With ParentRow.cells(4)
        .innerText = FormatCurrency(StockValue)
        .Align = "Right"
        .Style.fontWeight = "Bold"
    End With

End If

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

With CurrentRow.cells(0)
    .innerText = "1. " & PrintString
    .Align = "Left"
End With

With CurrentRow.cells(1)
    .innerText = FormatCurrency(StockValue)
    .Align = "Right"
End With

If NewBook Then
    With CurrentRow.cells(1)
        .innerText = FormatCurrency(opBalance)
        .Align = "Right"
    End With
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(StockValue)
        .Align = "Right"
    End With

End If

End Sub

Private Sub PLPrintProfitorLoss(Profit As Currency, PrintString As String, isRight As Boolean)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

If isRight Then
    Set CurrentTable = m_WebRow.cells(1).All.Income
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
Else
    Set CurrentTable = m_WebRow.cells(0).All.Expense
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 3)
End If
    
With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(Abs(Profit))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

Private Sub PrintPLorOP(Amount As Currency, PrintString As String, isRight As Boolean)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

If isRight Then
    Set CurrentTable = m_WebRow.cells(1).All.Income
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(m_OtherDetail, 3, 2))
Else
    Set CurrentTable = m_WebRow.cells(0).All.Expense
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(m_OtherDetail, 3, 2))
End If
    
With CurrentRow.cells(0)
    .innerText = PrintString
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(IIf(m_OtherDetail, 2, 1))
    .innerText = FormatCurrency(Abs(Amount))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub


Private Sub BALSetBalanceSheetDifference(AssetTotal() As Currency, LiabilityTotal() As Currency, Optional NewBook As Boolean = False)

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim ParentAmount As Currency

' If same then not necessary to publish difference
If AssetTotal(1) = LiabilityTotal(1) Then Exit Sub

If AssetTotal(1) > LiabilityTotal(1) Then

    ParentAmount = AssetTotal(1) \ 1 - LiabilityTotal(1) \ 1
    
    Set CurrentTable = m_WebRow.cells(0).All.Liability
    
ElseIf LiabilityTotal(1) > AssetTotal(1) Then

    ParentAmount = LiabilityTotal(1) - AssetTotal(1)
    
    Set CurrentTable = m_WebRow.cells(1).All.Asset
    
End If
    
If (Not CurrentTable Is Nothing) And ParentAmount > 0 Then

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))

    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 5, 3))
    
    CurrentRow.Style.fontWeight = "bold"
    
    With CurrentRow.cells(0)
        .innerText = "Difference in BalanceSheet"
        .Align = "Left"
    End With
    
    With CurrentRow.cells(IIf(NewBook, 4, 2))
        .innerText = FormatCurrency(ParentAmount)
        .Align = "Right"
    End With
    
End If
    
' This Equalise the rows of both the tables
BALFillEmptyRows (NewBook)

End Sub

Private Function BALCreateViewsForBalanceSheet() As Boolean

On Error GoTo Hell:

BALCreateViewsForBalanceSheet = False

gDbTrans.SqlStmt = " SELECT a.HeadID as OpHeadID,OpAmount,ParentID,HeadName" & _
                   " FROM OpBalance a,Heads b " & _
                   " WHERE OpDate = " & "#" & GetSysFormatDate(FinIndianFromDate) & "#" & _
                   " AND a.HeadID=b.HeadID" & _
                   " ORDER BY a.HeadID"
                   
gDbTrans.CreateView ("qryOPHeads")

gDbTrans.SqlStmt = " SELECT (sum(Credit)-sum(Debit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=2 " & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryLiabilities")

gDbTrans.SqlStmt = " SELECT (sum(Debit)-sum(Credit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID, c.ParentName " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=1 " & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryAssets")

BALCreateViewsForBalanceSheet = True

Exit Function

Hell:
    
End Function
Private Function BALCreateViewsForConditionalBalanceSheet() As Boolean

On Error GoTo Hell:

BALCreateViewsForConditionalBalanceSheet = False

gDbTrans.SqlStmt = " SELECT a.HeadID as OpHeadID,OpAmount,ParentID,HeadName" & _
                   " FROM OpBalance a,Heads b " & _
                   " WHERE OpDate = " & "#" & FinUSFromDate & "#" & _
                   " AND a.HeadID=b.HeadID" & _
                   " ORDER BY a.HeadID"
                   
gDbTrans.CreateView ("qryOPHeads")

gDbTrans.SqlStmt = " SELECT (sum(Credit)-sum(Debit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID" & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType= " & Liability & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " GROUP BY c.ParentID,a.HeadID, Headname"

                   
gDbTrans.CreateView ("qryLiabilities")

gDbTrans.SqlStmt = " SELECT (sum(Debit)-sum(Credit)) AS Amount, a.Headid, " & _
                   " HeadName, c.ParentID" & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType= " & Asset & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#" & _
                   " GROUP BY c.ParentID, a.HeadID, Headname"
                   
gDbTrans.CreateView ("qryAssets")

BALCreateViewsForConditionalBalanceSheet = True

Exit Function

Hell:
    
End Function

Private Function CreateViewsForReport2(RepType As wisReports) As Boolean

On Error GoTo Hell:

CreateViewsForReport2 = False
Dim SqlStr As String


SqlStr = " SELECT sum(Credit) AS Amount, a.Headid, " & _
        " HeadName, c.ParentID, c.ParentName " & _
        " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
        " WHERE a.HeadID = b.HeadID " & _
        " AND b.ParentID=c.ParentID " & _
        " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
        " AND #" & m_ToUSDate & "#"

If RepType = wisProfitLossStatement Then
    SqlStr = SqlStr & _
            " And (C.AccountType = " & Profit & _
                " OR C.AccountType = " & Loss & ") " & _
            " AND C.ParentID <> " & wis_TradingIncParentID & _
            " AND C.ParentID <> " & wis_TradingExpParentID
Else
    SqlStr = SqlStr & " And a.HeadID <> " & wis_CashHeadID
End If

gDbTrans.SqlStmt = SqlStr & " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("QryReceipts")

SqlStr = " SELECT sum(Debit) AS Amount, a.Headid , " & _
            " HeadName, c.ParentID, c.ParentName " & _
            " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
            " WHERE a.HeadID = b.HeadID " & _
            " AND b.ParentID=c.ParentID " & _
            " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
            " AND #" & m_ToUSDate & "#"

If RepType = wisProfitLossStatement Then
    SqlStr = SqlStr & _
            " And (C.AccountType = " & Profit & _
                " OR C.AccountType = " & Loss & ") " & _
            " AND C.ParentID <> " & wis_TradingIncParentID & _
            " AND C.ParentID <> " & wis_TradingExpParentID
Else
    SqlStr = SqlStr & " And a.HeadID <> " & wis_CashHeadID
End If
                   
gDbTrans.SqlStmt = SqlStr & " GROUP BY c.ParentID, c.ParentName, a.HeadID, Headname"
                   
gDbTrans.CreateView ("QryPayments")

CreateViewsForReport2 = True

Exit Function

Hell:
    
End Function

Private Function NetProfitViews() As Boolean

On Error GoTo Hell:

NetProfitViews = False

gDbTrans.SqlStmt = " SELECT (sum(Credit)-sum(Debit)) As NetIncome " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=8" & _
                   " AND C.ParentID <> " & wis_TradingIncParentID & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#"
                   
gDbTrans.CreateView ("QryIncomes")

gDbTrans.SqlStmt = " SELECT (sum(Debit)-sum(Credit)) As NetExpenses " & _
                   " FROM AccTrans AS a, Heads AS b, ParentHeads AS c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND c.AccountType=4 " & _
                   " AND C.ParentID <> " & wis_TradingExpParentID & _
                   " AND a.TransDate BETWEEN #" & m_FromUSDate & "#" & _
                   " AND #" & m_ToUSDate & "#"
                   
gDbTrans.CreateView ("QryExpenses")

                   
NetProfitViews = True

Exit Function

Hell:
    
End Function

Private Function InsertRowAndCellSToTable(Table As HTMLTable, _
                            Optional CellsCount As Integer) As HTMLTableRow

Dim Cell As Integer
Dim Row As HTMLTableRow

If m_WebDoc Is Nothing Then Set m_WebDoc = New HTMLDocument

With m_WebDoc

    'Set Row = Table.tBodies(0).insertRow(-1)
    Set Row = .createElement("TR")
    Row.Height = 19
    Table.tBodies(0).appendChild Row
    
    'if user has not specified the cols then get the columns from the table
    If CellsCount = 0 Then CellsCount = Table.Cols
    
    For Cell = 0 To CellsCount - 1
        Set m_WebCell = .createElement("TD")
        With Row
            .appendChild m_WebCell
            '.Style.fontFamily = gFontName
            '.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
            
        End With
    Next Cell
    'Row.Style.fontFamily = gFontName
    'Row.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    m_WebCell.innerText = " "

End With

Set InsertRowAndCellSToTable = Row

End Function
Private Sub BALPutAssetorLiability(rstSumParents As ADODB.Recordset, _
                                        rstOBHeads As ADODB.Recordset, _
                                        rstOBParents As ADODB.Recordset, _
                                        rstHeads As ADODB.Recordset, _
                                        LiabilityTotal() As Currency, _
                                        AssetTotal() As Currency, _
                                        ShowNegative As Boolean, Optional NewBook As Boolean = False)
On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim headID As Long
Dim HeadsCount As Long

Dim opBalance As Currency
Dim showOPBalance As Currency
Dim currentBalance As Currency
Dim ParentAmount As Currency
Dim ParentOPAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim headName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow

Dim FirstLoop As Boolean
Dim Repeat As Boolean
Dim isAsset As Boolean
Dim PrintStatus As wis_PrintStatus

Const cnstLIABILITY = 0
Const cnstASSET = 1
Const cnstHEAD = 0

Dim totalCols As Integer
Dim cnstHEAD_AMOUNT As Integer
Dim cnstAMOUNT As Integer

cnstHEAD_AMOUNT = IIf(NewBook, 3, 1)
cnstAMOUNT = IIf(NewBook, 4, 2)
totalCols = IIf(NewBook, 5, 3)

ParentID = FormatField(rstOBParents.Fields("ParentID"))
opBalance = FormatField(rstOBParents.Fields("ParentOB"))
ParentName = FormatField(rstOBParents.Fields("ParentName"))
isAsset = IIf(FormatField(rstOBParents.Fields("AccountType")) = Asset, True, False)
PrintStatus = FormatField(rstOBParents.Fields("PrintDetailed"))

FirstLoop = True
If PrintStatus <> PrintDetailed Then ShowNegative = True

' Filter the table
rstOBHeads.Filter = adFilterNone
rstOBHeads.Filter = "ParentID=" & ParentID

' This is total count
HeadsCount = rstHeads.RecordCount

RepeatLine:

' get the parent details
ParentAmount = 0
ParentOPAmount = 0
Repeat = False

If rstSumParents.RecordCount > 0 Then _
    If gDbTrans.FindRecord(rstSumParents, "ParentID=" & ParentID) Then _
                ParentAmount = FormatField(rstSumParents.Fields("SumAmount"))

ParentAmount = ParentAmount + opBalance

If ParentAmount <= 0 And FirstLoop Then
    Dim NegativeNos As Integer
    
    If rstOBHeads.RecordCount = 0 Then ShowNegative = True
    NegativeNos = 1
    Do
        If rstOBHeads.EOF Then Exit Do
        Amount = FormatField(rstOBHeads.Fields("opAmount"))
        If Not rstHeads Is Nothing Then
            If rstHeads.RecordCount Then
                rstHeads.MoveFirst
                rstHeads.Find "HeadID = " & rstOBHeads("opHeadID")
                If Not rstHeads.EOF Then Amount = FormatField(rstHeads("Amount")) + Amount
            End If
        End If
        If Amount > 0 Then NegativeNos = 0: Exit Do
        rstOBHeads.MoveNext
    Loop
    rstOBHeads.MoveFirst
    If NegativeNos Then isAsset = Not isAsset: FirstLoop = False
End If
                        
If isAsset Then Set CurrentTable = m_WebRow.cells(cnstASSET).All.Asset
If Not isAsset Then Set CurrentTable = m_WebRow.cells(cnstLIABILITY).All.Liability

If ParentAmount <= 0 And ShowNegative Then isAsset = Not isAsset

If Not FirstLoop And Not ShowNegative Then ParentAmount = 0

' Start Putting Parentid's heads
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)

ParentRow.Style.fontWeight = "bold"

With ParentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
End With

SlNo = 0
ParentAmount = 0
Do While Not rstOBHeads.EOF
    
    opBalance = FormatField(rstOBHeads.Fields("OpAmount"))
    headID = FormatField(rstOBHeads.Fields("OpHeadID"))
    headName = FormatField(rstOBHeads.Fields("HeadName"))
        
    Amount = 0: currentBalance = 0
        
    If HeadsCount > 0 Then If gDbTrans.FindRecord(rstHeads, "HeadID=" & headID) Then _
            Amount = FormatField(rstHeads.Fields("Amount"))
    
    If Amount = 0 And opBalance = 0 Then GoTo NextHead
    
    Amount = Amount + opBalance:    currentBalance = Amount + opBalance
    'Debug.Assert Amount > -1
    If Amount - opBalance < 0 And Amount > 0 Then NegativeNos = NegativeNos - 1
    If (Amount < 0 Or opBalance < 0) And FirstLoop Then Repeat = True
    
    showOPBalance = opBalance
    If Not ShowNegative Then
        If NewBook Then
            If Amount = 0 And opBalance = 0 Then GoTo NextHead
            If FirstLoop Then
                If Amount <= 0 And opBalance <= 0 Then GoTo NextHead
                If Amount < 0 And opBalance > 0 Then Amount = 0
                If opBalance < 0 And Amount > 0 Then opBalance = 0
            Else
                If Amount >= 0 And opBalance >= 0 Then GoTo NextHead
                If Amount > 0 And opBalance < 0 Then Amount = 0
                If opBalance > 0 And Amount < 0 Then opBalance = 0
                'If Amount < 0 Then Amount = 0
                'If opBalance < 0 Then opBalance = 0
            End If
            If Amount >= 0 And opBalance = 0 And Not FirstLoop Then GoTo NextHead
            If Amount <= 0 And opBalance = 0 And FirstLoop Then GoTo NextHead Else Amount = Abs(Amount): showOPBalance = Abs(opBalance)
        Else
            If Amount = 0 Then GoTo NextHead
            If Amount >= 0 And Not FirstLoop Then GoTo NextHead
            If Amount <= 0 And FirstLoop Then GoTo NextHead Else Amount = Abs(Amount): showOPBalance = Abs(opBalance)
        End If
    End If
    'It Will Print Only if Detailed Print is Enabled
    If PrintStatus = PrintDetailed Then
        
        If Amount <> 0 Or (NewBook = True And opBalance <> 0) Then
            
            SlNo = SlNo + 1
            
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
            
            With CurrentRow.cells(cnstHEAD)
                .innerText = SlNo & "." & headName
                .Align = "Left"
            End With
            
            With CurrentRow.cells(cnstHEAD_AMOUNT)
                .innerText = FormatCurrency(Amount)
                .Align = "Right"
            End With
            
            If NewBook Then
            Debug.Assert opBalance > -1
                With CurrentRow.cells(1)
                    .innerText = FormatCurrency(showOPBalance)
                    .Align = "Right"
                End With
            End If
        End If
        
    End If
    ParentAmount = ParentAmount + Amount
    'ParentAmount = ParentAmount + Abs(Amount)
    'ParentOPAmount = ParentOPAmount + opBalance
    ParentOPAmount = ParentOPAmount + Abs(opBalance)
NextHead:

    rstOBHeads.MoveNext
Loop

' Now put the parenthead Total
With ParentRow.cells(cnstAMOUNT)
    .innerText = FormatCurrency(Abs(ParentAmount))
    .Align = "Right"
    .Style.fontWeight = "bold"
End With
If NewBook Then
    With ParentRow.cells(2)
        .innerText = FormatCurrency(Abs(ParentOPAmount))
        .Align = "Right"
        .Style.fontWeight = "bold"
    End With
End If
            
If isAsset And ParentAmount > 0 Then
    AssetTotal(1) = AssetTotal(1) + Abs(ParentAmount)
    AssetTotal(0) = AssetTotal(0) + Abs(ParentOPAmount)
Else
    LiabilityTotal(1) = LiabilityTotal(1) + Abs(ParentAmount)
    LiabilityTotal(0) = LiabilityTotal(0) + Abs(ParentOPAmount)
End If
    
If PrintStatus = NoPrintDetailed Then Exit Sub


If Repeat And Not ShowNegative And FirstLoop Then
    Repeat = False
    FirstLoop = False
    isAsset = Not isAsset
    rstOBHeads.MoveFirst
    GoTo RepeatLine
End If

Hell:
    'Resume

End Sub
Private Sub BALPutAssetorLiability_OLD(rstSumParents As ADODB.Recordset, _
                                        rstOBHeads As ADODB.Recordset, _
                                        rstOBParents As ADODB.Recordset, _
                                        rstHeads As ADODB.Recordset, _
                                        LiabilityTotal As Currency, _
                                        AssetTotal As Currency, _
                                        ShowNegative As Boolean)
On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim headID As Long
Dim HeadsCount As Long

Dim opBalance As Currency
Dim ParentAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim headName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow
Dim ParentRow As HTMLTableRow

Dim FirstLoop As Boolean
Dim Repeat As Boolean
Dim isAsset As Boolean
Dim PrintStatus As wis_PrintStatus

Const cnstHEAD = 0
Const cnstHEAD_AMOUNT = 1
Const cnstAMOUNT = 2

Const cnstLIABILITY = 0
Const cnstASSET = 1

Const cnstTOTALCOLS = 3


ParentID = FormatField(rstOBParents.Fields("ParentID"))
opBalance = FormatField(rstOBParents.Fields("ParentOB"))
ParentName = FormatField(rstOBParents.Fields("ParentName"))
isAsset = IIf(FormatField(rstOBParents.Fields("AccountType")) = Asset, True, False)
PrintStatus = FormatField(rstOBParents.Fields("PrintDetailed"))

FirstLoop = True
If PrintStatus <> PrintDetailed Then ShowNegative = True

' Filter the table
rstOBHeads.Filter = adFilterNone
rstOBHeads.Filter = "ParentID=" & ParentID

' This is total count
HeadsCount = rstHeads.RecordCount

RepeatLine:

' get the parent details
ParentAmount = 0
Repeat = False

If rstSumParents.RecordCount > 0 Then _
    If gDbTrans.FindRecord(rstSumParents, "ParentID=" & ParentID) Then _
                ParentAmount = FormatField(rstSumParents.Fields("SumAmount"))

ParentAmount = ParentAmount + opBalance

If ParentAmount <= 0 And FirstLoop Then
    Dim NegativeNos As Integer
    
    If rstOBHeads.RecordCount = 0 Then ShowNegative = True
    NegativeNos = 1
    Do
        If rstOBHeads.EOF Then Exit Do
        Amount = FormatField(rstOBHeads.Fields("opAmount"))
        If Not rstHeads Is Nothing Then
        If rstHeads.RecordCount Then
            rstHeads.MoveFirst
            rstHeads.Find "HeadID = " & rstOBHeads("opHeadID")
            If Not rstHeads.EOF Then Amount = FormatField(rstHeads("Amount")) + Amount
        End If
        End If
        If Amount > 0 Then NegativeNos = 0: Exit Do
        rstOBHeads.MoveNext
    Loop
    rstOBHeads.MoveFirst
    If NegativeNos Then isAsset = Not isAsset: FirstLoop = False
End If
                        
If isAsset Then Set CurrentTable = m_WebRow.cells(cnstASSET).All.Asset
If Not isAsset Then Set CurrentTable = m_WebRow.cells(cnstLIABILITY).All.Liability

If ParentAmount <= 0 And ShowNegative Then isAsset = Not isAsset

If Not FirstLoop And Not ShowNegative Then ParentAmount = 0

' Start Putting Parentid's heads
Set ParentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)

ParentRow.Style.fontWeight = "bold"

With ParentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
End With

SlNo = 0
ParentAmount = 0
Do While Not rstOBHeads.EOF
    
    opBalance = FormatField(rstOBHeads.Fields("OpAmount"))
    headID = FormatField(rstOBHeads.Fields("OpHeadID"))
    headName = FormatField(rstOBHeads.Fields("HeadName"))
        
    Amount = 0
        
    If HeadsCount > 0 Then If gDbTrans.FindRecord(rstHeads, "HeadID=" & headID) Then _
            Amount = FormatField(rstHeads.Fields("Amount"))
    
    Amount = Amount + opBalance
    If Amount - opBalance < 0 And Amount > 0 Then NegativeNos = NegativeNos - 1
    If Amount < 0 And FirstLoop Then Repeat = True
    
    If Not ShowNegative Then
        If Amount = 0 Then GoTo NextHead
        'If Amount < 0 Then ParentAmount = ParentAmount + Abs(Amount)
        If Amount >= 0 And Not FirstLoop Then GoTo NextHead
        If Amount <= 0 And FirstLoop Then GoTo NextHead Else Amount = Abs(Amount)
    End If
    'It Will Print Only if Detailed Print is Enabled
    If PrintStatus = PrintDetailed Then
        
        If Amount <> 0 Then
            
            SlNo = SlNo + 1
            
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)
            
            With CurrentRow.cells(cnstHEAD)
                .innerText = SlNo & "." & headName
                .Align = "Left"
            End With
            
            With CurrentRow.cells(cnstHEAD_AMOUNT)
                .innerText = FormatCurrency(Amount)
                .Align = "Right"
            End With
        End If
        
    End If
    ParentAmount = ParentAmount + Abs(Amount)
NextHead:

    rstOBHeads.MoveNext
Loop

' Now put the parenthead Total
With ParentRow.cells(cnstAMOUNT)
    .innerText = FormatCurrency(Abs(ParentAmount))
    .Align = "Right"
    .Style.fontWeight = "bold"
End With

If isAsset And ParentAmount > 0 Then
    AssetTotal = AssetTotal + Abs(ParentAmount)
Else
    LiabilityTotal = LiabilityTotal + Abs(ParentAmount)
End If
    
If PrintStatus = NoPrintDetailed Then Exit Sub


If Repeat And Not ShowNegative And FirstLoop Then
    Repeat = False
    FirstLoop = False
    isAsset = Not isAsset
    rstOBHeads.MoveFirst
    GoTo RepeatLine
End If

Hell:
    'Resume

End Sub
Private Sub PrintTagsFromWebDocument()

Dim TagsCount As Integer
Dim Tag As Integer

With m_frmWeb.web.document
    TagsCount = .All.Length
'    For Tag = 0 To TagsCount - 1
'        Debug.Print .All(Tag).tagName
'    Next Tag
End With

End Sub

Private Sub PutHtmlStyle()

With m_frmWeb.web.document

    .write "<style>"
    .write ".clsContentText         {font-family:verdana;font-size:x-small;padding-left:11px;text-align:center;}"
    
    .write ".clsRowBold         {font-weight: bold}"
    
    .write " P  {font: 10pt/12pt; verdana;color: black;"
    .write "    font-weight: bold;"
    .write "    text-align:center;"
    .write "}"
    
    .write " table {"
    .write " font-weight: bold;"
    '.Write "TABLE-LAYOUT:fixed;"
    .write "border-collapse:collapse;"
    .write "cellPadding=0;width=80% "
    .write "border=1;"
    .write "}"
    
    .write "tr {"
    .write "font-weight: normal;"
    .write "height=20;"
    .write "align=center;"
    .write "}"
    
    .write "td {"
    .write "font-weight: normal;"
    .write "height=20;"
    .write "width=50;"
    .write "}"
    
    .write "</style>"

End With

End Sub

Private Function PLPutIncomeorExpenses(rstSum As ADODB.Recordset, _
                               rstHeads As ADODB.Recordset, isExpense As Boolean) As Currency

On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim headID As Long

Dim ParentAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim headName As String

Dim isIncome As Boolean

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const cnstHEAD = 0
Const cnstHEAD_AMOUNT = 1
Const cnstAMOUNT = 2

Const cnstEXPENSES = 0
Const cnstINCOME = 1

Const cnstTOTALCOLS = 3

PLPutIncomeorExpenses = 0

ParentID = FormatField(rstSum.Fields("ParentID"))
ParentAmount = FormatField(rstSum.Fields("SumAmount"))
ParentName = FormatField(rstSum.Fields("ParentName"))

If ParentAmount = 0 Then GoTo Hell

If ParentAmount > 0 Then

    If isExpense Then
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    Else
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    End If
    
Else

    If isExpense Then
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    Else
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    End If
    
End If
                        
' Start Putting parentid's heads
        
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)

        
With CurrentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With


' Filter the table
rstHeads.Filter = adFilterNone
rstHeads.Filter = "ParentID=" & ParentID

SlNo = 0

Do While Not rstHeads.EOF

    
    Amount = FormatField(rstHeads.Fields("Amount"))
    headID = FormatField(rstHeads.Fields("HeadID"))
    headName = FormatField(rstHeads.Fields("HeadName"))
    
    If Amount <> 0 Then
    
        SlNo = SlNo + 1
        
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)
        
        With CurrentRow.cells(cnstHEAD)
            .innerText = SlNo & "." & headName
            .Align = "Left"
            .Style.FONTSIZE = 13
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
        End With
        
        With CurrentRow.cells(cnstHEAD_AMOUNT)
            .innerText = FormatCurrency(Amount)
            .Align = "Right"
            .Style.FONTSIZE = 13
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
        End With
        
    End If
    
    rstHeads.MoveNext
    
Loop

' Now put the parenthead Total

With CurrentRow.cells(cnstAMOUNT)
    .innerText = FormatCurrency(Abs(ParentAmount))
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

PLPutIncomeorExpenses = ParentAmount

Exit Function

Hell:
    
End Function

Private Function PutReceiptORPayment(rstSum As ADODB.Recordset, _
                               rstHeads As ADODB.Recordset, isRightSide As Boolean) As Currency

On Error GoTo Hell:

Dim SlNo As Integer

Dim ParentID As Long
Dim headID As Long

Dim ParentAmount As Currency
Dim Amount As Currency

Dim ParentName As String
Dim headName As String

Dim isIncome As Boolean

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const cnstHEAD = 0
Const cnstHEAD_AMOUNT = 1
Const cnstAMOUNT = 2

Const cnstEXPENSES = 0
Const cnstINCOME = 1

'Const cnstTOTALCOLS = 3
Dim cnstTOTALCOLS  As Integer
cnstTOTALCOLS = IIf(m_OtherDetail, 3, 2)

PutReceiptORPayment = 0

ParentID = FormatField(rstSum.Fields("ParentID"))
Debug.Assert ParentID <> 210000
Debug.Assert ParentID <> 220000
ParentAmount = FormatField(rstSum.Fields("SumAmount"))
ParentName = FormatField(rstSum.Fields("ParentName"))

If ParentAmount = 0 Then GoTo Hell

If ParentAmount > 0 Then
    If isRightSide Then
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    Else
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    End If
    
Else
    If isRightSide Then
        Set CurrentTable = m_WebRow.cells(cnstINCOME).All.Income
    Else
        Set CurrentTable = m_WebRow.cells(cnstEXPENSES).All.Expense
    End If
    
End If
                        
' Start Putting parentid's heads
        
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)

        
With CurrentRow.cells(cnstHEAD)
    .innerText = ParentName
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With


' Filter the table
rstHeads.Filter = adFilterNone
rstHeads.Filter = "ParentID=" & ParentID

SlNo = 0

Do While Not rstHeads.EOF

    
    Amount = FormatField(rstHeads.Fields("Amount"))
    headID = FormatField(rstHeads.Fields("HeadID"))
    headName = FormatField(rstHeads.Fields("HeadName"))
    
    If Amount <> 0 Then
    
        SlNo = SlNo + 1
        
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, cnstTOTALCOLS)
        
        With CurrentRow.cells(cnstHEAD)
            .innerText = SlNo & "." & headName
            .Align = "Left"
            .Style.FONTSIZE = 13
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
        End With
        
        With CurrentRow.cells(cnstHEAD_AMOUNT)
            .innerText = FormatCurrency(Amount)
            .Align = "Right"
            .Style.FONTSIZE = 13
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
        End With
        
    End If
    
    rstHeads.MoveNext
    
Loop

' Now put the parenthead Total
If m_OtherDetail Then
    With CurrentRow.cells(cnstAMOUNT)
        .innerText = FormatCurrency(Abs(ParentAmount))
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If

PutReceiptORPayment = ParentAmount

Exit Function

Hell:
    
End Function


Private Sub PLPutPandLTableStructure()

Dim CurrentRow As HTMLTableRow

Set CurrentRow = InsertRowAndCellSToTable(m_WebTable, 4)

With CurrentRow.cells(0)
    .innerText = GetResourceString(368) '"Expenses"
    .Align = "Left"
    .Width = "40%"
End With

With CurrentRow.cells(1)
    .innerText = GetResourceString(40) '"Amount"
    .Align = "Left"
    .Width = "10%"
End With

With CurrentRow.cells(2)
    .innerText = GetResourceString(366) '"Income"
    .Align = "Left"
    .Width = "40%"
End With

With CurrentRow.cells(3)
    .innerText = GetResourceString(40) '"Amount"
    .Align = "Left"
    .Width = "10%"
End With

End Sub



Private Sub PutTradingProfitORLoss(ByVal Purchase As Currency, _
                                   ByVal Sales As Currency)

Dim Profit As Currency

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

With m_frmWeb.web.document

    If Purchase >= Sales Then
    
        Profit = Purchase - Sales
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
                
        If Profit <> 0 Then
            m_WebRow.Style.fontWeight = "Bold"
            
            With m_WebRow.cells(2)
                .innerText = "Gross " & GetResourceString(404)  '"Gross Loss"
                .Align = "Left"
            End With
            
            With m_WebRow.cells(3)
                .innerText = FormatCurrency(Profit)
                .Align = "Right"
            End With
        End If
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
                
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = GetResourceString(286) '"Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Purchase)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(2)
            .innerText = GetResourceString(286) '"Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Purchase)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17

        End With
                
    Else
    
        Profit = Sales - Purchase
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = GetResourceString(403) '"Gross Profit"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Profit)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
        
        m_WebRow.Style.fontWeight = "Bold"
        
        With m_WebRow.cells(0)
            .innerText = GetResourceString(286) '"Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Sales)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(2)
            .innerText = GetResourceString(286) '"Gross Total"
            .Align = "Left"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Sales)
            .Align = "Right"
            .Style.fontWeight = "Bold"
            .Style.FONTSIZE = 15
            If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
        End With
    End If
    
End With


End Sub

Private Sub PLPutProfitORLoss(ByVal Income As Currency, _
                              ByVal Expenses As Currency)
                            
Dim Profit As Currency
Dim Total As Currency
Dim Cell As Integer

Const TotalCells = 4
If m_WebTable Is Nothing Then Exit Sub

With m_WebTable

    'just an extra line
    Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
    
    If Expenses >= Income Then
    
        Profit = Expenses - Income
                    
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
        m_WebRow.Style.fontWeight = "Bold"
        With m_WebRow.cells(2)
            .innerText = GetResourceString(332, 404) '"Net Loss"
            .Align = "Left"
        End With
        
        With m_WebRow.cells(3)
            .innerText = FormatCurrency(Profit)
            .Align = "right"
        End With
                    
        Total = Expenses
    Else
    
        Profit = Income - Expenses
        Total = Income
        
        Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
        m_WebRow.Style.fontWeight = "Bold"
        With m_WebRow.cells(0)
            .innerText = GetResourceString(332, 403) '"Net Profit"
            .Align = "Left"
        End With
        
        With m_WebRow.cells(1)
            .innerText = FormatCurrency(Profit)
            .Align = "right"
        End With
        
    End If
    
End With

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)

' print the total
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, TotalCells)
m_WebRow.Style.fontWeight = "Bold"
With m_WebRow.cells(0)
    .innerText = "Net Total"
    .Align = "Left"
End With

With m_WebRow.cells(1)
    .innerText = FormatCurrency(Total)
    .Align = "right"
End With

With m_WebRow.cells(2)
    .innerText = "Net Total"
    .Align = "Left"
End With

With m_WebRow.cells(3)
    .innerText = FormatCurrency(Total)
    .Align = "right"
End With

End Sub


Private Sub BALSetAssetHead(Optional NewBook As Boolean = False)

Dim CurrentRow As HTMLTableRow
Dim assetTable As HTMLTable

Set assetTable = m_WebRow.cells(1).All.Asset

assetTable.Style.fontFamily = gFontName
assetTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
assetTable.Style.BorderColor = "Black"

If NewBook Then
    Set CurrentRow = InsertRowAndCellSToTable(assetTable, 3)
    CurrentRow.Style.fontWeight = "bold"
    With CurrentRow.cells(0)
        .innerText = GetResourceString(347) '"Asset"
        .Align = "Center"
        .noWrap = True
        .rowSpan = 2
        .Width = "22%"
    End With
    With CurrentRow.cells(1)
        .colSpan = 2
        .innerText = FinIndianFromDate
        .Align = "Center"
        .Width = "20%"
    End With
    With CurrentRow.cells(2)
        .colSpan = 2
        .innerText = GetFromDateString(ToRpDate)
        .Align = "center"
        .Width = "20%"
    End With
    
    Set CurrentRow = InsertRowAndCellSToTable(assetTable, 4)
    CurrentRow.Style.fontWeight = "bold"
    
    With CurrentRow.cells(0)
        .innerText = GetResourceString(40) '"amout"
        '.Align = "Left"
        .Align = "center"
        .Width = "19%"
    End With
    
    With CurrentRow.cells(1)
        .innerText = GetResourceString(52) '"Head Total"
        .Align = "center"
        .Width = "19%"
    End With
    
    With CurrentRow.cells(2)
        .innerText = GetResourceString(40) '"Amount"
        .Align = "center"
        .Width = "19%"
    End With
    With CurrentRow.cells(3)
        .innerText = GetResourceString(52) '"Amount"
        .Align = "center"
        .Width = "22%"
    End With
    'Set CurrentRow = InsertRowAndCellSToTable(AssetTable, 3)
    
Else
    Set CurrentRow = InsertRowAndCellSToTable(assetTable, IIf(NewBook, 5, 3))
CurrentRow.Style.fontWeight = "bold"

    With CurrentRow.cells(0)
        .innerText = GetResourceString(347) '"Asset"
        .Align = "Left"
        .noWrap = True
        .Width = IIf(NewBook, "28%", "55%") '"55%"
    End With
    
    With CurrentRow.cells(1)
        .innerText = GetResourceString(52) '"Head Total"
        .Align = "Left"
        .Width = IIf(NewBook, "18%", "22%") '"22%"
    End With
    
    With CurrentRow.cells(2)
        .innerText = GetResourceString(40) '"Amount"
        .Align = "Left"
        .Width = IIf(NewBook, "18%", "22%") '"23%"
    End With
End If

End Sub
Private Sub PLSetIncomeHead()

Dim CurrentRow As HTMLTableRow
Dim IncomeTable As HTMLTable

Set IncomeTable = m_WebRow.cells(1).All.Income
Set CurrentRow = InsertRowAndCellSToTable(IncomeTable, 3)

'IncomeTable.Style.BorderColor = "Black"
With CurrentRow.cells(0)
    .innerText = GetResourceString(366) '"Income"
    .Align = "Left"
    .Width = "57%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(1)
    .innerText = GetResourceString(52)    '"Head Total"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(2)
    .innerText = GetResourceString(40) '"Amount"
    .Align = "Left"
    .Width = "23%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

Private Sub SetReceiptHeader(RepType As wisReports)

Dim CurrentRow As HTMLTableRow
Dim IncomeTable As HTMLTable

Set IncomeTable = m_WebRow.cells(1).All.Income
Set CurrentRow = InsertRowAndCellSToTable(IncomeTable, IIf(m_OtherDetail, 3, 2))

'IncomeTable.Style.BorderColor = "Black"
With CurrentRow.cells(0)
    .innerText = GetResourceString(IIf(RepType = wisProfitLossStatement, 366, 272)) '"Income"
    .Align = "Left"
    .Width = IIf(m_OtherDetail, "57%", "68%")
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(1)
    .innerText = GetResourceString(52)    '"Head Total"
    .Align = "Left"
    .Width = IIf(m_OtherDetail, "20%", "32%")
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
If m_OtherDetail Then
    With CurrentRow.cells(2)
        .innerText = GetResourceString(40) '"Amount"
        .Align = "Left"
        .Width = "23%"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
         If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If

End Sub

Private Sub BALSetLiabilityHeader(Optional NewBook As Boolean = False)

Dim CurrentRow As HTMLTableRow
Dim liabilityTable As HTMLTable

Set liabilityTable = m_WebRow.cells(0).All.Liability
liabilityTable.Style.fontFamily = gFontName
liabilityTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
If NewBook Then
    
    Set CurrentRow = InsertRowAndCellSToTable(liabilityTable, 3)
    CurrentRow.Style.fontWeight = "bold"
    With CurrentRow.cells(0)
        .innerText = GetResourceString(346) '"Liability"
        .Align = "Center"
        .noWrap = False
        .rowSpan = 2
        .Width = "22%"
    End With
    With CurrentRow.cells(1)
        .colSpan = 2
        .innerText = FinIndianFromDate
        .Align = "Center"
        .Width = "20%"
    End With
    With CurrentRow.cells(2)
        .colSpan = 2
        .innerText = GetFromDateString(ToRpDate)
        .Align = "center"
        .Width = "20%"
    End With
    
    Set CurrentRow = InsertRowAndCellSToTable(liabilityTable, 4)
    
    liabilityTable.Style.BorderColor = "Black"
    CurrentRow.Style.fontWeight = "bold"
    
    With CurrentRow.cells(0)
        .innerText = GetResourceString(40) '"amount"
        .Align = "Center"
        .Width = "19%"
    End With
    
    With CurrentRow.cells(1)
        .innerText = GetResourceString(52) '"Head Total"
        .Align = "Center"
        .Width = "19%"
    End With
    
    With CurrentRow.cells(2)
        .innerText = GetResourceString(40) '"Amount"
        .Align = "Center"
        .Width = "19%"
    End With
    
    With CurrentRow.cells(3)
        .innerText = GetResourceString(52) '"Amount"
        .Align = "Center"
        .Width = "22%"
    End With
    
Else
    Set CurrentRow = InsertRowAndCellSToTable(liabilityTable, 3)
    
    liabilityTable.Style.BorderColor = "Black"
    
    CurrentRow.Style.fontWeight = "bold"
    
    With CurrentRow.cells(0)
        .innerText = GetResourceString(346) '"Liability"
        .Align = "Left"
        .noWrap = False
        .Width = IIf(NewBook, "28%", "55%") '"55%"
    End With
    
    With CurrentRow.cells(1)
        .innerText = GetResourceString(52) '"Head Total"
        .Align = "Left"
        .Width = IIf(NewBook, "18%", "22%") '"22%"
    End With
    With CurrentRow.cells(2)
        .innerText = GetResourceString(40) '"Amount"
        .Align = "Left"
        .Width = IIf(NewBook, "18%", "23%") '"23%"
    End With

End If

End Sub
Private Sub PLSetExpenseHeader()

Dim CurrentRow As HTMLTableRow
Dim ExpenseTable As HTMLTable

Set ExpenseTable = m_WebRow.cells(0).All.Expense
Set CurrentRow = InsertRowAndCellSToTable(ExpenseTable, 3)

CurrentRow.Style.fontWeight = "bold"

With CurrentRow.cells(0)
    .innerText = GetResourceString(368) '"Expenses"
    .Align = "Left"
    .Width = "57%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(1)
    .innerText = GetResourceString(52) '"Head Total"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(2)
    .innerText = GetResourceString(40) ' "Amount"
    .Align = "Left"
    .Width = "23%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub
Private Sub SetPaymentHeader(RepType As wisReports)

Dim CurrentRow As HTMLTableRow
Dim ExpenseTable As HTMLTable

Set ExpenseTable = m_WebRow.cells(0).All.Expense
Set CurrentRow = InsertRowAndCellSToTable(ExpenseTable, IIf(m_OtherDetail, 3, 2))


CurrentRow.Style.fontWeight = "bold"

With CurrentRow.cells(0)
    .innerText = GetResourceString(IIf(RepType = wisProfitLossStatement, 368, 271))
    .Align = "Left"
    .Width = IIf(m_OtherDetail, "57%", "68%")
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(1)
    .innerText = GetResourceString(40) ' "Amount"
    .Align = "Center"
    .Width = IIf(m_OtherDetail, "20%", "32%")
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
If m_OtherDetail Then
    With CurrentRow.cells(2)
        .innerText = GetResourceString(52) '"Head Total"
        .Align = "Center"
        .Width = "23%"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
         If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If

End Sub


'' This will create the views required for the Receipts and Payments
Private Sub RPCreateViews(ByVal qryName As String)

' This Qurey will only load the Receipts and Payments details of the given period

gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID, b.debit AS Receipt, " & _
                   " b.Credit AS Payment" & _
                   " FROM QryAccTransMerge  " & _
                   " WHERE TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND " & "#" & m_ToUSDate & "#" & _
                   " AND  a.HeadID<>b.Headid And a.Headid<>" & wis_CashHeadID
                 
gDbTrans.CreateView (qryName)


End Sub

'' This will create the views required for the Receipts and Payments
Private Sub SubDayCreateViews(ByVal qryName As String)

' This Qurey will only load the Receipts and Payments details of the given period

gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID, b.debit AS Receipt, " & _
                   " b.Credit AS Payment" & _
                   " FROM QryAccTransMerge,ParentHeads c  " & _
                   " WHERE TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND " & "#" & m_ToUSDate & "#" & _
                   " AND  a.HeadID<>b.Headid And a.Headid<>" & wis_CashHeadID

gDbTrans.CreateView (qryName)


End Sub
'' This will create the views required for the Receipts and Payments
Private Sub GLCreateViews(ByVal qryName As String, NewBook As Boolean)

' This Qurey will only load the Receipts and Payments details of the given period
                 
gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID, a.Credit AS Receipt, " & _
                   " a.Debit AS Payment" & _
                   " FROM AccTrans a  " & _
                   " WHERE TransDate BETWEEN " & "#" & IIf(NewBook, m_FromUSDate, m_ToUSDate) & "#" & _
                   " AND " & "#" & m_ToUSDate & "#" & _
                   " AND a.Headid<>" & wis_CashHeadID

gDbTrans.CreateView (qryName)

End Sub

Private Sub RPprintCash(ByVal PrintString, ByVal CashAmount As Currency, ByVal TableName As String)

If TableName = "" Then Exit Sub
If PrintString = "" Then Exit Sub

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Set CurrentTable = m_WebDoc.All(TableName)
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, 4)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = PrintString
    .Align = "Left"
End With

If CashAmount >= 0 Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
    End With
Else
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
    End With
End If

End Sub

'This will print the cash Closing Balance
Private Sub RPputClosingCash(ByVal CashAmount As Currency, ByVal TableName As String)

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

' Validations

If TableName = "" Then Exit Sub

Set CurrentTable = m_WebTable

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

With CurrentRow.cells(1)
    .innerText = GetResourceString(285)  '"Closing Balance"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

If CashAmount >= 0 Then
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
Else
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(CashAmount)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If

End Sub
Private Function RPPutOpeningCash(ByVal TableName As String, AsOnIndainDate As String) As Currency

' Handle Error
On Error GoTo Hell:

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim OpeningCash As Currency

Dim AccTransClass As clsAccTrans

RPPutOpeningCash = 0

Set AccTransClass = New clsAccTrans

OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, GetSysFormatDate(AsOnIndainDate))

Set AccTransClass = Nothing

Set CurrentTable = m_WebTable
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = GetResourceString(284) '"Opening Cash Balance"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

If OpeningCash >= 0 Then
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
Else
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If

RPPutOpeningCash = OpeningCash

Exit Function

Hell:

End Function
Private Sub RPPutCashinField(ByVal TableName As String, _
                                 ByVal AsOnIndainDate As String, _
                                 ByVal FieldNo As Long, _
                                 ByVal MsgStr As String, _
                                 ByVal Amount As Currency)

Debug.Print "Check Code"
' Handle Error
On Error GoTo Hell:

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Set CurrentTable = m_WebTable
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(FieldNo >= 8, FieldNo + 1, 8))

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = MsgStr
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(FieldNo)
    .innerText = FormatCurrency(Amount)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With


Exit Sub

Hell:

End Sub

Private Sub PutAmountInBold(ByVal TableName As String, _
                                 ByVal AsOnIndainDate As String, _
                                 ByVal FieldNo As Long, ByVal NoOfCols As Integer, _
                                 ByVal MsgStr As String, _
                                 ByVal Amount As Currency)

Debug.Print "Check Code"
' Handle Error
On Error GoTo Hell:

' Declarations
Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Set CurrentTable = m_WebTable
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, NoOfCols)

CurrentRow.Style.fontWeight = "Bold"

With CurrentRow.cells(1)
    .innerText = MsgStr
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With CurrentRow.cells(FieldNo)
    .innerText = FormatCurrency(Amount)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With


Exit Sub

Hell:

End Sub
Private Function FormatCurrency(ByVal Curr As Currency) As String
    FormatCurrency = Format(Abs(Curr), "##########0.00")
    If Curr < 0 Then FormatCurrency = "(-)" & FormatCurrency
    
End Function


Public Function ShowDailyCashBook(NewBook As Boolean) As Boolean

' Handle Error
On Error GoTo Hell:

Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset
Dim rstDeposits As Recordset
Dim rstMemType As Recordset
Dim DepName As String

'Shashi June 2013 Added for Total amount display
Dim rstTotal As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
''

Dim DepositType As Integer
Dim Receipts As Currency
Dim Payments As Currency
Dim opBalance As Currency
Dim CLBalance As Currency
'Dim SumReceipts As Currency
Dim OpeningCash As Currency
'Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim headID As Long
Dim ParentID As Long

Dim headName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans
Dim ColsInTable As Integer
Dim isModuleHasTrans As Boolean

ShowDailyCashBook = False

ColsInTable = IIf(NewBook, 9, 7)
' Validataions
If ToRpDate = "" Then Exit Function
'Set the Total as 0
SetTotals

CashBooKStartReport (NewBook)

' Put the Opening Cash to the Html
Set AccTransClass = New clsAccTrans

OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, m_ToUSDate)

Set CurrentTable = m_WebTable
Set AccTransClass = New clsAccTrans

If NewBook Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 9, 7))
    
    CurrentRow.Style.fontWeight = "Bold"
    
    With CurrentRow.cells(4)
        .innerText = GetResourceString(284)
        .Align = "Left"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
    
    With CurrentRow.cells(0)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
    
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(OpeningCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
Else
    Call PutAmountInBold("CashBook", ToRpDate, 3, ColsInTable, GetResourceString(284), OpeningCash)
End If
''''

SlNo = 0

'Get the Member TYpes
gDbTrans.SqlStmt = "SELECT * From MemberTypeTab"
Call gDbTrans.Fetch(rstMemType, adOpenDynamic)

'First select The the Details Of Member
'Get the total amount
If CheckAndAddTransTotal("MemTrans") Then
    While rstMemType.EOF = False
        DepositType = FormatField(rstMemType("memberType"))
        'Get the each transcations
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM MemTrans A,MemMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID " & _
                " And C.CustomerId = B.CustomerID "
        
        If DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.MemberType = " & DepositType
            
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = IIf(DepositType > 0, FormatField(rstMemType("MemberTypeName")) & " ", GetResourceString(49)) + _
                            GetResourceString(53)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        rstMemType.MoveNext
    Wend
End If

''Member fees
If CheckAndAddTransTotal("MemIntTrans") Then
    If Not rstMemType.BOF Then rstMemType.MoveFirst
    While rstMemType.EOF = False
        DepositType = FormatField(rstMemType("memberType"))
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM MemIntTrans A,MemMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And TransID = 1 And A.AccID = B.AccID " & _
                " And Amount > 0 And C.CustomerId = B.CustomerID"
        If DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.MemberType = " & DepositType
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = IIf(DepositType > 0, FormatField(rstMemType("MemberTypeName")) & " ", "") + _
                            GetResourceString(79, 191)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        rstMemType.MoveNext
    Wend
'End If
'Share Fees
'If CheckAndAddTransTotal("MemIntTrans") Then
    If Not rstMemType.BOF Then rstMemType.MoveFirst
    While rstMemType.EOF = False
        DepositType = FormatField(rstMemType("memberType"))
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM MemIntTrans A,MemMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And TransID > 1 And A.AccID = B.AccID " & _
                " And Amount > 0 And C.CustomerId = B.CustomerID"
        If DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.MemberType = " & DepositType
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = IIf(DepositType > 0, FormatField(rstMemType("MemberTypeName")) & " ", "") + _
                            GetResourceString(53, 191) 'Share Fee 53
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        rstMemType.MoveNext
    Wend
End If

''Member payble
If CheckAndAddTransTotal("MemIntPayable") Then
    If Not rstMemType.BOF Then rstMemType.MoveFirst
    While rstMemType.EOF = False
        DepositType = FormatField(rstMemType("memberType"))
    
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM MemIntPayable A,MemMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID " & _
                " And C.CustomerId = B.CustomerID"
        If DepositType > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And B.MemberType = " & DepositType
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = IIf(DepositType > 0, FormatField(rstMemType("MemberTypeName")) & " ", "") + _
                            GetResourceString(53, 191)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        rstMemType.MoveNext
    Wend
End If


'Deposit Loans
gDbTrans.SqlStmt = "SELECT * From DepositName"
If gDbTrans.Fetch(rstDeposits, adOpenDynamic) < 1 Then Set rstDeposits = Nothing

If CheckAndAddTransTotal("DepositLoanTrans") Then
    'Get the details of RD deposit Loans
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_RD, GetResourceString(424, 58), wisPrincipal)
    
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_PD, GetResourceString(425, 58), wisPrincipal)
    If Not rstDeposits Is Nothing Then
        rstDeposits.MoveFirst
        While Not rstDeposits.EOF
            DepName = FormatField(rstDeposits("DepositName")) & " " & GetResourceString(58)
            Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, rstDeposits(0), DepName, wisPrincipal)
            rstDeposits.MoveNext
        Wend
    End If
End If

'Interest on deposit loan
If CheckAndAddTransTotal("DepositLoanIntTrans") Then
    'Get the details of RD deposit Loans
    DepName = GetResourceString(424, 58, 47)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_RD, DepName, wisRegularInt)
    
    DepName = GetResourceString(425, 58, 47)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_PD, DepName, wisRegularInt)
    If Not rstDeposits Is Nothing Then
        rstDeposits.MoveFirst
        While Not rstDeposits.EOF
            DepName = FormatField(rstDeposits("DepositName")) & " " & GetResourceString(58, 47)
            Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, rstDeposits(0), DepName, wisRegularInt)
            rstDeposits.MoveNext
        Wend
    End If
End If
'Penal Interest on deposit loan
If CheckAndAddTransTotal_FieldName("DepositLoanIntTrans", "PenalAmount") Then

    'Get the details of RD deposit Loans
    DepName = GetResourceString(424, 58, 345)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_RD, DepName, wisPenalInt)
    
    DepName = GetResourceString(425, 58, 345)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_PD, DepName, wisPenalInt)
    If Not rstDeposits Is Nothing Then
        rstDeposits.MoveFirst
        While Not rstDeposits.EOF
            DepName = FormatField(rstDeposits("DepositName")) & " " & GetResourceString(58, 345)
            Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, rstDeposits(0), DepName, wisPenalInt)
            rstDeposits.MoveNext
        Wend
    End If
End If
'Penal Interest on deposit loan
If CheckAndAddTransTotal_FieldName("DepositLoanIntTrans", "MisCAmount") Then
    'Get the details of RD deposit Loans
    DepName = GetResourceString(424, 58, 327)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_RD, DepName, wisMisc)
    
    DepName = GetResourceString(425, 58, 327)
    Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, wisDeposit_PD, DepName, wisMisc)
    If Not rstDeposits Is Nothing Then
        rstDeposits.MoveFirst
        While Not rstDeposits.EOF
            DepName = FormatField(rstDeposits("DepositName")) & " " & GetResourceString(58, 327)
            Call ShowDailyCashBook_DepositLoan(CurrentTable, SlNo, NewBook, rstDeposits(0), DepName, wisMisc)
            rstDeposits.MoveNext
        Wend
    End If

End If

'BKcc Loan Details
gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
        " FROM BKCCTrans WHERE Deposit = " & False & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name" & _
        " FROM BKCCTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " And A.LoanID = B.LoanID And Deposit = " & False & _
        " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"

If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = GetResourceString(229, 58)
        .Align = "Left"
    End With
    Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
End If

gDbTrans.SqlStmt = "SELECT Sum(IntAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT Sum(PenalIntAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,(IntAmount+PenalIntAmount) as Amount,TransType,Name" & _
        " FROM BKCCIntTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " AND (IntAmount+PenalIntAmount) > 0 And A.LoanID = B.LoanID And Deposit = " & False & _
        " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"

If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = GetResourceString(229, 47)
        .Align = "Left"
    End With
    Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
End If
''Misca Amount from BKCC Loan
gDbTrans.SqlStmt = "SELECT Sum(MiscAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,MiscAmount as Amount,TransType,Name" & _
        " FROM BKCCIntTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " AND MiscAmount > 0 And A.LoanID = B.LoanID And Deposit = " & False & _
        " And C.CustomerId = B.CustomerID and MiscAmount > 0 ORDER BY Transdate,val(AccNum)"

If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = GetResourceString(229, 327)
        .Align = "Left"
    End With
    Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
End If


'Loan Details
gDbTrans.SqlStmt = "SELECT Sum(IntAmount) as Total,TransType" & _
        " FROM LoanIntTrans WHERE TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)
gDbTrans.SqlStmt = "SELECT Sum(PenalIntAmount) as Total,TransType" & _
        " FROM LoanIntTrans WHERE TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)
gDbTrans.SqlStmt = "SELECT Sum(MiscAmount) as Total,TransType" & _
        " FROM LoanIntTrans WHERE TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

isModuleHasTrans = CheckAndAddTransTotal("LoanTrans")
isModuleHasTrans = CheckAndAddTransTotal("LoanIntReceivable")
''Check for the Interest

gDbTrans.SqlStmt = "Select * From LoanScheme"
If gDbTrans.Fetch(rstParents, adOpenDynamic) > 0 Then
    While Not rstParents.EOF
        headName = FormatField(rstParents("SchemeName"))
        gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name" & _
                " FROM LoanTrans A,LoanMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.LoanID = B.LoanID ANd SchemeID = " & rstParents("SchemeId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId," & _
                " (IntAmount+PenalIntAmount) as amount,TransType,Name" & _
                " FROM LoanIntTrans A,LoanMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And (IntAmount+PenalIntAmount) > 0 And A.LoanID = B.LoanID ANd SchemeID = " & rstParents("SchemeId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName & " " & GetResourceString(47)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name" & _
                " FROM LoanIntReceivable A,LoanMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.LoanID = B.LoanID ANd SchemeID = " & rstParents("SchemeId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName & " " & _
                    GetResourceString(376, 47)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId," & _
                " MiscAmount as amount,TransType,Name" & _
                " FROM LoanIntTrans A,LoanMaster B,QryName C " & _
                " WHERE MiscAmount > 0 And TransDate = #" & m_ToUSDate & "#" & _
                " And A.LoanID = B.LoanID and SchemeID = " & rstParents("SchemeId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName & " " & GetResourceString(327)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        rstParents.MoveNext
    Wend
End If

'Savings Bank
If CheckAndAddTransTotal("SbTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM SbTrans A,SBMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(421)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If
If CheckAndAddTransTotal("SbPLTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM SbPLTrans A,SBMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(421, 47)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Current account Bank
If CheckAndAddTransTotal("CATrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType, Name" & _
            " FROM CATrans A,CAMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(422)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If
If CheckAndAddTransTotal("CAPLTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType, Name" & _
            " FROM CAPLTrans A,CAMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(422) & " Interest"
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Fd Account
'Check is there any transaction happened in FD

isModuleHasTrans = CheckAndAddTransTotal("FDTrans")
isModuleHasTrans = isModuleHasTrans Or CheckAndAddTransTotal("FDIntTrans")
isModuleHasTrans = isModuleHasTrans Or CheckAndAddTransTotal("FDIntPayable")
isModuleHasTrans = isModuleHasTrans Or CheckAndAddTransTotal("MatFDTrans")

If isModuleHasTrans Then
  gDbTrans.SqlStmt = "Select * From DepositName"
  If gDbTrans.Fetch(rstParents, adOpenDynamic) > 0 Then
    While Not rstParents.EOF
        headName = FormatField(rstParents("DepositName"))
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM FDTrans A,FDMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID ANd DepositType = " & rstParents("DepositId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM FDIntTrans A,FDMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID ANd DepositType = " & rstParents("DepositId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName & " " & GetResourceString(47)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM FDIntPayable A,FDMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID ANd DepositType = " & rstParents("DepositId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = headName & " " & _
                    GetResourceString(47, 450)
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
                " FROM MatFDTrans A,FDMaster B,QryName C " & _
                " WHERE TransDate = #" & m_ToUSDate & "#" & _
                " And A.AccID = B.AccID ANd DepositType = " & rstParents("DepositId") & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
        If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
            CurrentRow.Style.fontWeight = "Bold"
            With CurrentRow.cells(IIf(NewBook, 4, 1))
                .innerText = GetResourceString(220) & " " & headName '43
                .Align = "Left"
            End With
            Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
        End If
        
        rstParents.MoveNext
    Wend
  End If
End If

'recurring Deposit
If CheckAndAddTransTotal("RDTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM RDTrans A,RDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(424)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

If CheckAndAddTransTotal("RDIntTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM RDIntTrans A,RDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
     If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(424, 27)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
     End If
End If


'RD interest payble
If CheckAndAddTransTotal("RDIntPayable") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM RDIntPayable A,RDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(424) & " " & _
                    GetResourceString(450)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Pigmy Agent commision
If CheckAndAddTransTotal("AgentTrans") Then
    gDbTrans.SqlStmt = "SELECT AgentID as AccNum,Amount,TransType, Name" & _
            " FROM AgentTrans A,UserTab B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AgentID = B.UserID " & _
            " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
    'Below We are going to Update for the Individual Pigmy User Transaction
    'So keep the Totals of Deposits(and contadeposits) here
   
End If
Receipts = m_depositTotal
Payments = m_contraDepositTotal
'Pigmy Transaction

If CheckAndAddTransTotal("PDTrans") Then
    ''Put back Deplosit totals other wise Pigmy Deposit Transaction becomes double
        m_depositTotal = Receipts
        m_contraDepositTotal = Payments
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType, Name" & _
            " FROM PDTrans A,PDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And (TransType = " & wWithdraw & " OR TransType = " & wContraWithdraw & ")" & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Pigmy interest & fees
If CheckAndAddTransTotal("PDIntTrans") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM PDIntTrans A,PDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425, 47)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Pigmy interest payble
If CheckAndAddTransTotal("PDIntPayable") Then
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name" & _
            " FROM PDIntPayable A,PDMaster B,QryName C " & _
            " WHERE TransDate = #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And (TransType = " & wWithdraw & " OR TransType = " & wContraWithdraw & ")" & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425) & " " & _
                GetResourceString(450)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

'Bkcc Details
gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
        " FROM BKCCTrans WHERE Deposit = " & True & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name" & _
        " FROM BKCCTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " And Amount > 0 And A.LoanID = B.LoanID And Deposit = " & True & _
        " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"

If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    CurrentRow.Style.fontWeight = "Bold"
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = GetResourceString(229, 43)
        .Align = "Left"
    End With
    Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
End If

'BKCC int trans
gDbTrans.SqlStmt = "SELECT Sum(IntAmount+PenalIntAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & True & " And TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,(IntAmount+PenalIntAmount) as Amount,TransType,Name" & _
        " FROM BKCCIntTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " and (IntAmount+PenalIntAmount) > 0 And A.LoanID = B.LoanID And Deposit = " & True & _
        " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"

If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then _
                Call InsertTransDetails(rstRPDetails, SlNo, NewBook)

' create views for the receipts and payments
gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID,VoucherType, " & _
                   " a.Credit AS Receipt, a.Debit AS Payment,A.TransID,Particulars " & _
                   " FROM AccTrans A Left join TransParticulars B  " & _
                   " On A.TransID = B.TransID " & _
                   " WHERE TransDate = #" & m_ToUSDate & "#" & _
                   " AND a.Headid <> " & wis_CashHeadID & _
                   " And a.HeadID Not In (Select HeadId From BankHeadIDs )"

gDbTrans.CreateView ("QryCashTrans")

gDbTrans.SqlStmt = " SELECT Receipt AS Receipts, Payment AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName,a.VoucherType,a.Particulars,a.TransID" & _
                   " FROM QryCashTrans AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " ORDER BY b.ParentID,a.RPHeadID,b.HeadName,a.TransID"
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)
                   
gDbTrans.SqlStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"
Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

' Main Loop
Dim VoucherType As Wis_VoucherTypes
Dim colno As Byte
Dim Amount As Currency
Dim strPart As String
Dim subTotalDeposit As Currency
Dim subTotalContraDeposit As Currency
Dim subTotalWithDraw As Currency
Dim subTotalContraWithDraw As Currency

Dim printHeadTotal As Boolean
Dim printParentHeadTotal As Boolean

Dim headTotalDeposit As Currency
Dim headTotalContraDeposit As Currency
Dim headTotalWithDraw As Currency
Dim headTotalContraWithDraw As Currency

Dim parentHeadTotalDeposit As Currency
Dim parentHeadTotalContraDeposit As Currency
Dim parentHeadTotalWithDraw As Currency
Dim parentHeadTotalContraWithDraw As Currency
'Dim parentName As String

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
Do While Not rstRPDetails.EOF
    strPart = ""
    If rstRPDetails("ParentID") <> ParentID Then
        If printParentHeadTotal = True Then
            Set CurrentRow = InsertCashBookTotals(CurrentTable, NewBook, 0, ParentName + " " + GetResourceString(42), parentHeadTotalDeposit, parentHeadTotalContraDeposit, parentHeadTotalWithDraw, parentHeadTotalContraWithDraw)
            CurrentRow.cells(IIf(NewBook, 4, 1)).Align = "Center"
        Else
            parentHeadTotalDeposit = 0: parentHeadTotalContraDeposit = 0: parentHeadTotalWithDraw = 0: parentHeadTotalContraWithDraw = 0
        End If
        printParentHeadTotal = False
        printHeadTotal = False
        ParentID = rstRPDetails("ParentID")
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        rstParents.Find "ParentID = " & ParentID
        ParentName = FormatField(rstParents("ParentName"))
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = ParentName
            .Align = "Left"
        End With
    Else
        printParentHeadTotal = True
    End If
    
    If headID = FormatField(rstRPDetails("RPHeadID")) Then
        printHeadTotal = True
    ElseIf printHeadTotal Then
        printHeadTotal = False
        Call InsertCashBookTotals(CurrentTable, NewBook, 0, GetResourceString(42), headTotalDeposit, headTotalContraDeposit, headTotalWithDraw, headTotalContraWithDraw, False)
    Else
        headTotalDeposit = 0: headTotalContraDeposit = 0: headTotalWithDraw = 0: headTotalContraWithDraw = 0
    End If

    strPart = FormatField(rstRPDetails("Particulars"))
    headID = FormatField(rstRPDetails("RPHeadID"))
    headName = FormatField(rstRPDetails("HeadName"))
    If Len(Trim$(strPart)) > 0 Then headName = headName + "     (" + strPart + ")"
    Receipts = FormatField(rstRPDetails("Receipts"))
    Payments = FormatField(rstRPDetails("Payments"))
    VoucherType = FormatField(rstRPDetails("VoucherType"))
    If Receipts Then colno = IIf(NewBook, 2, 3): Amount = Receipts
    If Payments Then colno = IIf(NewBook, 6, 5): Amount = Payments
    If VoucherType = Journal Then
        colno = colno + IIf(Receipts, IIf(NewBook, -1, 1), 1)
        headTotalContraDeposit = headTotalContraDeposit + Receipts
        headTotalContraWithDraw = headTotalContraWithDraw + Payments
        parentHeadTotalContraDeposit = parentHeadTotalContraDeposit + Receipts
        parentHeadTotalContraWithDraw = parentHeadTotalContraWithDraw + Payments
        subTotalContraDeposit = subTotalContraDeposit + Receipts
        subTotalContraWithDraw = subTotalContraWithDraw + Payments
        m_contraDepositTotal = m_contraDepositTotal + Receipts
        m_contraWithdrawTotal = m_contraWithdrawTotal + Payments
    Else
        headTotalDeposit = headTotalDeposit + Receipts
        headTotalWithDraw = headTotalWithDraw + Payments
        parentHeadTotalDeposit = parentHeadTotalDeposit + Receipts
        parentHeadTotalWithDraw = parentHeadTotalWithDraw + Payments
        
        subTotalDeposit = subTotalDeposit + Receipts
        subTotalWithDraw = subTotalWithDraw + Payments
        m_depositTotal = m_depositTotal + Receipts
        m_withdrawTotal = m_withdrawTotal + Payments
    End If
    SlNo = SlNo + 1
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    
    If Not NewBook Then
        With CurrentRow.cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With
    End If
    
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = headName
        .Align = "Left"
    End With

    With CurrentRow.cells(colno)
        .innerText = FormatCurrency(Amount)
        .Align = "Right"
    End With
    
    rstRPDetails.MoveNext
Loop


If (headTotalDeposit = parentHeadTotalDeposit And headTotalContraDeposit = parentHeadTotalContraDeposit And headTotalWithDraw = parentHeadTotalWithDraw And headTotalContraWithDraw = parentHeadTotalContraWithDraw) Then printHeadTotal = False

If printHeadTotal Then Call InsertCashBookTotals(CurrentTable, NewBook, 0, GetResourceString(42), headTotalDeposit, headTotalContraDeposit, headTotalWithDraw, headTotalContraWithDraw)
If printParentHeadTotal = True Then Call InsertCashBookTotals(CurrentTable, NewBook, 0, ParentName + " " + GetResourceString(42), parentHeadTotalDeposit, parentHeadTotalContraDeposit, parentHeadTotalWithDraw, parentHeadTotalContraWithDraw): m_WebRow.cells(1).Align = "Center"


'''Conclusion

'Now Put Closing Balance
Dim closingCash As Currency
closingCash = AccTransClass.GetOpBalance(wis_CashHeadID, DateAdd("d", 1, m_ToUSDate))
If NewBook Then
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, IIf(NewBook, 9, 7))
    
    CurrentRow.Style.fontWeight = "Bold"
    
    With CurrentRow.cells(4)
        .innerText = GetResourceString(285)
        .Align = "Left"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
    
    With CurrentRow.cells(6)
        .innerText = FormatCurrency(closingCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
    
    With CurrentRow.cells(8)
        .innerText = FormatCurrency(closingCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
    ''''
Else
    Call PutAmountInBold("CashBook", ToRpDate, ColsInTable - 1, ColsInTable, GetResourceString(285), OpeningCash)
End If




'If subTotalDeposit Or subTotalContraDeposit Or subTotalWithDraw Or subTotalWithDraw Then _
 '   Call InsertCashBookTotals(m_WebTable, NewBook, 0, GetResourceString(42), subTotalDeposit, subTotalContraDeposit, subTotalWithDraw, subTotalContraWithDraw)


    ''Now Put the Total of Deposit & Withdrwal
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    headTotalContraDeposit = m_depositTotal + m_contraDepositTotal
    headTotalContraWithDraw = m_withdrawTotal + m_contraWithdrawTotal
    OpeningCash = OpeningCash + headTotalContraDeposit
    closingCash = closingCash + headTotalContraWithDraw
    Set CurrentRow = InsertCashBookTotals(m_WebTable, NewBook, 0, GetResourceString(52), m_depositTotal, m_contraDepositTotal, m_withdrawTotal, m_contraWithdrawTotal)
    If NewBook Then
        With CurrentRow.cells(0)
            .innerText = FormatCurrency(headTotalContraDeposit)
            .Align = "Right"
            .Style.fontWeight = "Bold"
        End With
        With CurrentRow.cells(8)
            .innerText = FormatCurrency(headTotalContraWithDraw)
            .Align = "Right"
            .Style.fontWeight = "Bold"
        End With
        
        'Show the Total including opening/closing cash
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
         With CurrentRow.cells(4)
            .innerText = GetResourceString(286)
            .Align = "center"
            .Style.fontWeight = "Bold"
        End With
        With CurrentRow.cells(0)
            .innerText = FormatCurrency(OpeningCash)
            .Align = "Right"
            .Style.fontWeight = "Bold"
        End With
        With CurrentRow.cells(8)
            .innerText = FormatCurrency(closingCash)
            .Align = "Right"
            .Style.fontWeight = "Bold"
        End With
    End If
    

Set rstRPDetails = Nothing
Set rstParents = Nothing
Set AccTransClass = Nothing


' Exit Function
ShowDailyCashBook = True

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "Show Daily cash book:" & vbCrLf & Err.Description
'Resume
    
    
End Function

Private Sub UpdateContraLedger(headID As Long, LedgerDate As Date, ByRef contraDepositTotal As Currency, ByRef contraDepositId As Long, ByRef contraWithdrawTotal As Currency, contraWithrawId As Long, ByRef bankClass As clsBankAcc)
    If contraDepositTotal Then
        If contraWithrawId Then
            Call bankClass.UpdateContraTrans(headID, contraWithrawId, contraDepositTotal, LedgerDate)
            contraWithrawId = 0
        Else
            contraDepositId = headID
        End If
    End If
    If contraWithdrawTotal Then
        If contraDepositId Then
            Call bankClass.UpdateContraTrans(contraDepositId, headID, contraWithdrawTotal, LedgerDate)
            contraDepositId = 0
        Else
            contraWithrawId = headID
        End If
    End If
End Sub

Public Function UpdateLedgerBook(LedgerDate As Date) As Boolean

' Handle Error
On Error GoTo Hell:

Dim contraWithdraw As Currency
Dim contraDeposit As Currency
Dim contraWithrawId As Long
Dim contraDepositId As Long
Dim DepositId As Integer


Dim ContraDepositAmount As Currency
Dim ContraWithdrawAmount As Currency

Dim miscDeposit As Currency, miscContraDeposit As Currency
Dim miscWithDraw As Currency, miscContraWithdraw As Currency


Dim rstParents As ADODB.Recordset
Dim rstRPDetails As ADODB.Recordset
Dim bankClass As New clsBankAcc

'Shashi June 2013 Added for Total amount display
Dim rstTotal As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
''

Dim DepositType As Integer
Dim Receipts As Currency
Dim Payments As Currency
Dim opBalance As Currency
Dim CLBalance As Currency
'Dim SumReceipts As Currency
Dim OpeningCash As Currency
'Dim SumPayments As Currency
Dim ClosingStock As Currency

Dim SlNo As Long
Dim headID As Long
Dim ParentID As Long

Dim headName As String
Dim ParentName As String

Dim AccountType As wis_AccountType

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Dim AccTransClass As clsAccTrans
Dim ColsInTable As Integer
Dim isModuleHasTrans As Boolean

UpdateLedgerBook = False
Dim NewBook As Boolean

ColsInTable = IIf(NewBook, 9, 7)
' Validataions
If ToRpDate = "" Then Exit Function

' Put the Opening Cash to the Html
Set AccTransClass = New clsAccTrans
SlNo = 0

'First select The the Details Of Member
'Get the total amount
'Set the Total as 0
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("MemTrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(53, 191))  'Share Fees
    headID = bankClass.GetHeadIDCreated(GetResourceString(79, 191))  'mem Feee
    
    headID = bankClass.GetHeadIDCreated(GetResourceString(53, 36))   'Share account
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashWithDrawls(headID, m_withdrawTotal, LedgerDate)
End If

''Member fees
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("MemIntTrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(79, 191))  'mem Feee
    headID = bankClass.GetHeadIDCreated(GetResourceString(53, 191))  'Share Fees
    
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

''Member payble
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("MemIntPayable") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(49, 200))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

'Deposit Loans
Dim rstDeposits As Recordset
Dim DepName As String


gDbTrans.SqlStmt = "SELECT " & wisDeposit_RD & " as DepositId,' " & LoadResString(424) & " ' as DepositName  From DepositName " & _
    " Union SELECT " & wisDeposit_PD & " as DepositId,' " & LoadResString(425) & " ' as DepositName  From DepositName " & _
    " Union SELECT depositid,DepositName From DepositName"
    
If gDbTrans.Fetch(rstDeposits, adOpenDynamic) < 1 Then Set rstDeposits = Nothing
If Not rstDeposits Is Nothing Then
    rstDeposits.MoveFirst
    Dim strDepType As String
    While Not rstDeposits.EOF
        ' DepName & " " & GetResourceString(450) ' Int Payable
        'DepName & " " & GetResourceString(487) ' Int Paid
        'DepName & " " & GetResourceString(375, 47) ' Int Provision
        'DepName & " " & GetResourceString(58) ' Loan headID
        'DepName & " " & GetResourceString(58, 483) ' Int Provision
        'GetResourceString(46) & " " & Depname  ' Matrured De[poist
        
        DepName = FormatField(rstDeposits("DepositName"))
        DepositId = FormatField(rstDeposits("DepositId"))
        
        strDepType = IIf(DepositId = wisDeposit_PD, "PD", IIf(DepositId = wisDeposit_RD, "RD", "FD"))
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM " & strDepType & "Trans WHERE TransDate = #" & LedgerDate & "#" & _
                " And AccIdId In (Select AccId from " & strDepType & "Master where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(DepName)
            If (DepositId <> wisDeposit_PD) Then
                'For Pigmy Do not take deposits from Pigmy Trans,
                'Take it from AgentTrans to match the Cash in Hand
                Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            End If
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM " & strDepType & "IntTrans WHERE TransDate = #" & LedgerDate & "#" & _
                " And AccIdId In (Select AccId from " & strDepType & "Master where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(DepName & " " & GetResourceString(487))
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM " & strDepType & "IntPayable WHERE TransDate = #" & LedgerDate & "#" & _
                " And AccIdId In (Select AccId from " & strDepType & "Master where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(DepName & " " & GetResourceString(450))
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM MatFDTrans WHERE TransDate = #" & LedgerDate & "#" & _
                " And AccIdId In (Select AccId from FDMaster where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(GetResourceString(46) & " " & DepName)
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        
        ''Deposit Loans
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM DepositLoanTrans WHERE TransDate = #" & LedgerDate & "#" & _
                " And LoanId In (Select LoanId from DepositLoanMaster where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(DepName & " " & GetResourceString(58))
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
                " FROM DepositLoanIntTrans WHERE TransDate = #" & LedgerDate & "#" & _
                " And LoanId In (Select LoanId from DepositLoanMaster where DepositType = " & DepositId & ")" & _
                " Group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            Set m_rstLedgerTrans = rstTrans
            headID = bankClass.GetHeadIDCreated(DepName & " " & GetResourceString(58))
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        
        rstDeposits.MoveNext
    Wend
End If

'BKcc Loan Details
gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
        " FROM BKCCTrans WHERE Deposit = " & False & " And TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    Call SetTotals  ' TO Clear the Deposit/WitDraw amount
    Call AddToTotal(rstTrans)
    Set m_rstLedgerTrans = rstTrans
    headID = bankClass.GetHeadIDCreated(GetResourceString(229, 58))  '43
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If
'BKcc DEPOSIT
gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
        " FROM BKCCTrans WHERE Deposit = " & True & " And TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    Call SetTotals  ' TO Clear the Deposit/WitDraw amount
    Call AddToTotal(rstTrans)
    Set m_rstLedgerTrans = rstTrans
    headID = bankClass.GetHeadIDCreated(GetResourceString(229, 43))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

gDbTrans.SqlStmt = "SELECT Sum(IntAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    Call SetTotals  ' TO Clear the Deposit/WitDraw amount
    Call AddToTotal(rstTrans)
    Set m_rstLedgerTrans = rstTrans
    headID = bankClass.GetHeadIDCreated(GetResourceString(229, 43, 487))  'BKCC Deposit Interest
    headID = bankClass.GetHeadIDCreated(GetResourceString(229, 58, 344))  'BKCC Reg Interest
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

gDbTrans.SqlStmt = "SELECT Sum(PenalIntAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    Call SetTotals  ' TO Clear the Deposit/WitDraw amount
    Call AddToTotal(rstTrans)
    Set m_rstLedgerTrans = rstTrans
    headID = bankClass.GetHeadIDCreated(GetResourceString(229, 58, 345))  'BKCC Loan Penal Interest
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

''Misc Amount from BKCC Loan
gDbTrans.SqlStmt = "SELECT Sum(MiscAmount) as Total,TransType" & _
        " FROM BKCCIntTrans WHERE Deposit = " & False & " And TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    Call SetTotals  ' TO Clear the Deposit/WitDraw amount
    Call AddToTotal(rstTrans)
    headID = bankClass.GetHeadIDCreated(GetResourceString(327))  'Misc Amount
    
End If

'Loan Details
gDbTrans.SqlStmt = "SELECT Sum(MiscAmount) as Total,TransType" & _
        " FROM LoanIntTrans WHERE TransDate = #" & LedgerDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then Call AddToTotal(rstTrans)

'bankClass.GetHeadIDCreated(FormatField(rst("SchemeName")), parMemberLoan)
gDbTrans.SqlStmt = "Select * From LoanScheme"
If gDbTrans.Fetch(rstParents, adOpenDynamic) > 0 Then
    While Not rstParents.EOF
        headName = FormatField(rstParents("SchemeName"))
        
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
            " FROM LoanTrans WHERE TransDate = #" & LedgerDate & "#" & _
            " And LoanId in (Select Distinct LoanId from LoanMaster where SchemeId = " & rstParents("schemeid") & " )" & _
            " group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            headID = bankClass.GetHeadIDCreated(headName)    'Regular Int
            Set m_rstLedgerTrans = rstTrans
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        gDbTrans.SqlStmt = "SELECT Sum(IntAmount) as Total,TransType" & _
            " FROM LoanIntTrans WHERE TransDate = #" & LedgerDate & "#" & _
            " And LoanId in (Select Distinct LoanId from LoanMaster where SchemeId = " & rstParents("schemeid") & " )" & _
            " group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            headID = bankClass.GetHeadIDCreated(headName + " " + GetResourceString(344))   'Regular Int
            Set m_rstLedgerTrans = rstTrans
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        gDbTrans.SqlStmt = "SELECT Sum(PenalIntAmount) as Total,TransType" & _
            " FROM LoanIntTrans WHERE TransDate = #" & LedgerDate & "#" & _
            " And LoanId in (Select Distinct LoanId from LoanMaster where SchemeId = " & rstParents("schemeid") & " )" & _
            " group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            headID = bankClass.GetHeadIDCreated(headName + " " + GetResourceString(345))   'Penal Int
            Set m_rstLedgerTrans = rstTrans
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        'Interest Recievable
        gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
            " FROM LoanIntReceivable WHERE TransDate = #" & LedgerDate & "#" & _
            " And LoanId in (Select Distinct LoanId from LoanMaster where SchemeId = " & rstParents("schemeid") & " )" & _
            " group by TransType"
        If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
            Call SetTotals  ' TO Clear the Deposit/WitDraw amount
            Call AddToTotal(rstTrans)
            headID = bankClass.GetHeadIDCreated(headName + " " + GetResourceString(376, 47))   'Int Payable
            Set m_rstLedgerTrans = rstTrans
            Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
            Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
        End If
        
        rstParents.MoveNext
    Wend
End If

'Savings Bank
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("SbTrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(421))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("SbPLTrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(421, 487))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

'Current account Bank
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("CATrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(422))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("CAPLTrans") Then
    headID = bankClass.GetHeadIDCreated(GetResourceString(422, 487))
    Call bankClass.AddCashDeposits(headID, m_depositTotal, LedgerDate)
    Call bankClass.AddCashDeposits(headID, m_withdrawTotal, LedgerDate)
End If

'Fd Account

'Pigmy Agent commision
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("AgentTrans") Then
    gDbTrans.SqlStmt = "SELECT AgentID as AccNum,Amount,TransType, Name" & _
            " FROM AgentTrans A,UserTab B,QryName C " & _
            " WHERE TransDate = #" & LedgerDate & "#" & _
            " And A.AgentID = B.UserID " & _
            " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
    'Below We are going to Update for the Individual Pigmy User Transaction
    'So keep the Totals of Deposits(and contadeposits) here
   
End If

Receipts = m_depositTotal
Payments = m_contraDepositTotal
'Pigmy Transaction
Call SetTotals  ' TO Clear the Deposit/WitDraw amount
If CheckAndAddTransTotal("PDTrans") Then
    ''Put back Deplosit totals other wise Pigmy Deposit Transaction becomes double
        m_depositTotal = Receipts
        m_contraDepositTotal = Payments
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType, Name" & _
            " FROM PDTrans A,PDMaster B,QryName C " & _
            " WHERE TransDate = #" & LedgerDate & "#" & _
            " And A.AccID = B.AccID " & _
            " And (TransType = " & wWithdraw & " OR TransType = " & wContraWithdraw & ")" & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(rstRPDetails, adOpenDynamic) > 0 Then
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = GetResourceString(425)
            .Align = "Left"
        End With
        Call InsertTransDetails(rstRPDetails, SlNo, NewBook)
    End If
End If

' create views for the receipts and payments
gDbTrans.SqlStmt = " SELECT a.HeadID AS RPHeadID,VoucherType, " & _
                   " a.Credit AS Receipt, a.Debit AS Payment,A.TransID,Particulars " & _
                   " FROM AccTrans A Left join TransParticulars B  " & _
                   " On A.TransID = B.TransID " & _
                   " WHERE TransDate = #" & LedgerDate & "#" & _
                   " AND a.Headid <> " & wis_CashHeadID & _
                   " And a.HeadID Not In (Select HeadId From " & _
                        " BankHeadIDs )"

gDbTrans.CreateView ("QryCashTrans")

gDbTrans.SqlStmt = " SELECT Receipt AS Receipts, Payment AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName,a.VoucherType,a.Particulars" & _
                   " FROM QryCashTrans AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " ORDER BY b.ParentID,a.RPHeadID,b.HeadName"
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)
''Get Particulars for TransID
gDbTrans.SqlStmt = " SELECT Particulars " & _
                   " FROM TransParticulars A Inner Join QryCashTrans B" & _
                   " On A.TransID ="
Dim rstParticulars As Recordset
'Call gDbTrans.Fetch(rstParticulars, adOpenDynamic)
                   
gDbTrans.SqlStmt = " SELECT ParentName,ParentID,AccountType" & _
                   " FROM Parentheads " & _
                   " ORDER BY ParentID"
Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

' Main Loop
Dim VoucherType As Wis_VoucherTypes
Dim colno As Byte
Dim Amount As Currency
Dim strPart As String
Dim subTotalDeposit As Currency
Dim subTotalContraDeposit As Currency
Dim subTotalWithDraw As Currency
Dim subTotalContraWithDraw As Currency

Dim printHeadTotal As Boolean
Dim printParentHeadTotal As Boolean

Dim headTotalDeposit As Currency
Dim headTotalContraDeposit As Currency
Dim headTotalWithDraw As Currency
Dim headTotalContraWithDraw As Currency

Dim parentHeadTotalDeposit As Currency
Dim parentHeadTotalContraDeposit As Currency
Dim parentHeadTotalWithDraw As Currency
Dim parentHeadTotalContraWithDraw As Currency
'Dim parentName As String

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
Do While Not rstRPDetails.EOF
    strPart = ""
    If rstRPDetails("ParentID") <> ParentID Then
        If printParentHeadTotal = True Then
            Set CurrentRow = InsertCashBookTotals(CurrentTable, NewBook, 0, ParentName + " " + GetResourceString(42), parentHeadTotalDeposit, parentHeadTotalContraDeposit, parentHeadTotalWithDraw, parentHeadTotalContraWithDraw)
            CurrentRow.cells(IIf(NewBook, 4, 1)).Align = "Center"
        Else
            parentHeadTotalDeposit = 0: parentHeadTotalContraDeposit = 0: parentHeadTotalWithDraw = 0: parentHeadTotalContraWithDraw = 0
        End If
        printParentHeadTotal = False
        printHeadTotal = False
        ParentID = rstRPDetails("ParentID")
        Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
        CurrentRow.Style.fontWeight = "Bold"
        rstParents.Find "ParentID = " & ParentID
        ParentName = FormatField(rstParents("ParentName"))
        With CurrentRow.cells(IIf(NewBook, 4, 1))
            .innerText = ParentName
            .Align = "Left"
        End With
    Else
        printParentHeadTotal = True
    End If
    
    If headID = FormatField(rstRPDetails("RPHeadID")) Then
        printHeadTotal = True
    ElseIf printHeadTotal Then
        printHeadTotal = False
        Call InsertCashBookTotals(CurrentTable, NewBook, 0, GetResourceString(42), headTotalDeposit, headTotalContraDeposit, headTotalWithDraw, headTotalContraWithDraw, False)
    Else
        headTotalDeposit = 0: headTotalContraDeposit = 0: headTotalWithDraw = 0: headTotalContraWithDraw = 0
    End If

    strPart = FormatField(rstRPDetails("Particulars"))
    headID = FormatField(rstRPDetails("RPHeadID"))
    headName = FormatField(rstRPDetails("HeadName"))
    If Len(Trim$(strPart)) > 0 Then headName = headName + "     (" + strPart + ")"
    Receipts = FormatField(rstRPDetails("Receipts"))
    Payments = FormatField(rstRPDetails("Payments"))
    VoucherType = FormatField(rstRPDetails("VoucherType"))
    If Receipts Then colno = IIf(NewBook, 2, 3): Amount = Receipts
    If Payments Then colno = IIf(NewBook, 6, 5): Amount = Payments
    If VoucherType = Journal Then
        colno = colno + IIf(Receipts, IIf(NewBook, -1, 1), 1)
        headTotalContraDeposit = headTotalContraDeposit + Receipts
        headTotalContraWithDraw = headTotalContraWithDraw + Payments
        parentHeadTotalContraDeposit = parentHeadTotalContraDeposit + Receipts
        parentHeadTotalContraWithDraw = parentHeadTotalContraWithDraw + Payments
        subTotalContraDeposit = subTotalContraDeposit + Receipts
        subTotalContraWithDraw = subTotalContraWithDraw + Payments
        m_contraDepositTotal = m_contraDepositTotal + Receipts
        m_contraWithdrawTotal = m_contraWithdrawTotal + Payments
    Else
        headTotalDeposit = headTotalDeposit + Receipts
        headTotalWithDraw = headTotalWithDraw + Payments
        parentHeadTotalDeposit = parentHeadTotalDeposit + Receipts
        parentHeadTotalWithDraw = parentHeadTotalWithDraw + Payments
        
        subTotalDeposit = subTotalDeposit + Receipts
        subTotalWithDraw = subTotalWithDraw + Payments
        m_depositTotal = m_depositTotal + Receipts
        m_withdrawTotal = m_withdrawTotal + Payments
    End If
    SlNo = SlNo + 1
    Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, ColsInTable)
    
    If Not NewBook Then
        With CurrentRow.cells(0)
            .innerText = SlNo
            .Align = "Right"
        End With
    End If
    
    With CurrentRow.cells(IIf(NewBook, 4, 1))
        .innerText = headName
        .Align = "Left"
    End With

'            With CurrentRow.cells(2)
'                .innerText = FormatCurrency(OpBalance)
'                .Align = "Right"
'            End With

    With CurrentRow.cells(colno)
        .innerText = FormatCurrency(Amount)
        .Align = "Right"
    End With

    'With CurrentRow.cells(4)
    '    .innerText = FormatCurrency(Payments)
    '    .Align = "Right"
    'End With

    
    'SumReceipts = SumReceipts + Receipts
    'SumPayments = SumPayments + Payments

    rstRPDetails.MoveNext
Loop


If (headTotalDeposit = parentHeadTotalDeposit And headTotalContraDeposit = parentHeadTotalContraDeposit And headTotalWithDraw = parentHeadTotalWithDraw And headTotalContraWithDraw = parentHeadTotalContraWithDraw) Then printHeadTotal = False

If printHeadTotal Then Call InsertCashBookTotals(CurrentTable, NewBook, 0, GetResourceString(42), headTotalDeposit, headTotalContraDeposit, headTotalWithDraw, headTotalContraWithDraw)
If printParentHeadTotal = True Then Call InsertCashBookTotals(CurrentTable, NewBook, 0, ParentName + " " + GetResourceString(42), parentHeadTotalDeposit, parentHeadTotalContraDeposit, parentHeadTotalWithDraw, parentHeadTotalContraWithDraw): m_WebRow.cells(1).Align = "Center"


'''Conclusion

Set rstRPDetails = Nothing
Set rstParents = Nothing
Set AccTransClass = Nothing


' Exit Function
UpdateLedgerBook = True

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "Show Daily cash book:" & vbCrLf & Err.Description
'Resume
    
    
End Function

Private Function CheckAndAddTransTotal(strTableName As String, Optional addTotal = True) As Boolean
CheckAndAddTransTotal = False
Dim rstTrans As Recordset
gDbTrans.SqlStmt = "SELECT Sum(Amount) as Total,TransType" & _
        " FROM " & strTableName & " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " group by TransType"
If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    CheckAndAddTransTotal = True
    If addTotal Then Call AddToTotal(rstTrans)
End If
Set m_rstLedgerTrans = rstTrans
End Function
Private Function CheckAndAddTransTotal_FieldName(strTableName As String, FieldName As String, Optional addTotal = True) As Boolean

CheckAndAddTransTotal_FieldName = False
Dim rstTrans As Recordset
gDbTrans.SqlStmt = "SELECT Sum(" & FieldName & ") as Total,TransType" & _
        " FROM " & strTableName & " WHERE TransDate = #" & m_ToUSDate & "#" & _
        " Group by TransType having Sum(" & FieldName & ") > 0 "

If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) > 0 Then
    CheckAndAddTransTotal_FieldName = True
    If addTotal Then Call AddToTotal(rstTrans)
End If

Set m_rstLedgerTrans = rstTrans

End Function

Private Sub AddToTotal(rstTrans As Recordset)

On Error GoTo Err_line:
Dim transType As wisTransactionTypes

While rstTrans.EOF = False
    transType = FormatField(rstTrans("TransType"))
    If transType = wWithdraw Then m_withdrawTotal = m_withdrawTotal + FormatField(rstTrans(0))
    If transType = wDeposit Then m_depositTotal = m_depositTotal + FormatField(rstTrans(0))
    If transType = wContraWithdraw Then m_contraWithdrawTotal = m_contraWithdrawTotal + FormatField(rstTrans(0))
    If transType = wContraDeposit Then m_contraDepositTotal = m_contraDepositTotal + FormatField(rstTrans(0))
    rstTrans.MoveNext
Wend

Err_line:

End Sub
Private Sub SetTotals()
m_withdrawTotal = 0
m_depositTotal = 0
m_contraWithdrawTotal = 0
m_contraDepositTotal = 0

End Sub

Public Function ShowConditionBalanceSheet() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstOBHeads As ADODB.Recordset
Dim rstHeads As ADODB.Recordset

Dim rstSum As ADODB.Recordset

Dim rstOBParents As ADODB.Recordset

Dim AssetTotal(2) As Currency
Dim LiabilityTotal(2) As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim StockValue As Currency

Dim MaterialClass As clsMaterial
Dim PrintStatus As wis_PrintStatus

Dim SqlStr As String

Const cnstLIABILITY = 0
Const cnstASSET = 1

' Intialiase the Function
ShowConditionBalanceSheet = False


'Now Check whether the Negaitive amount has to shown in in
'same side or other side
Dim SetUp As New clsSetup
m_OtherDetail = CBool(SetUp.ReadSetupValue("General", "NegativeInBalanceSheet", "False"))
Set SetUp = Nothing


Call BALCreateViewsForConditionalBalanceSheet

' Opening Balances
gDbTrans.SqlStmt = " SELECT * FROM qryOPHeads"
If gDbTrans.Fetch(rstOBHeads, adOpenDynamic) < 1 Then Exit Function

SqlStr = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName," & _
        " b.PrintOrder,b.PrintDetailed,b.AccountType" & _
        " FROM qryOPHeads a,ParentHeads b" & _
        " WHERE a.ParentID=b.ParentID" & _
        " And (AccountType = " & Asset & " OR AccountType = " & Liability & ")" & _
        " And a.ParentID <> " & parProfitORLoss & _
        " GROUP BY a.ParentID,b.ParentName," & _
         " b.PrintDetailed,AccountType,b.PrintOrder " & _
        " ORDER By PrintOrder"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstOBParents, adOpenDynamic) < 1 Then Exit Function


SqlStr = " SELECT * FROM qryLiabilities"
SqlStr = SqlStr & " UNION " & "SELECT * FROM qryAssets"

gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstHeads, adOpenDynamic)
'Get the Liability Heads Except
SqlStr = " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryLiabilities" & _
        " Where parentID <> " & parProfitORLoss & _
        " GROUP BY ParentID"
SqlStr = SqlStr & " UNION " & _
        " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryAssets" & _
        " GROUP BY ParentID"

gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstSum, adOpenDynamic)

' Start the Html
BALStartBalanceSheet

' Start the Both Headers
BALSetAssetHead
BALSetLiabilityHeader

' Main Loop
Do

    'If rstOBAssetParents.EOF And rstOBLiabilityParents.EOF Then Exit Do
    If rstOBParents.EOF Then Exit Do
    
    If Not rstOBParents.EOF Then
        
        PrintStatus = rstOBParents.Fields("PrintDetailed")
    
        'Now Filter the Heads Record set for the bala sheet
        rstHeads.Filter = adFilterNone
        rstHeads.Filter = "ParentID=" & rstOBParents.Fields("ParentID")
        
        'If rstHeads.RecordCount Then
            Call BALPutAssetorLiability(rstSum, rstOBHeads, _
                    rstOBParents, rstHeads, LiabilityTotal, AssetTotal, CBool(m_OtherDetail))
        'End If
        
        If Not rstOBParents.EOF Then rstOBParents.MoveNext
    End If
    
Loop


'Get the Closing Stock value from the database
Set MaterialClass = New clsMaterial
Dim OpStock As Currency
OpStock = MaterialClass.GetOpeningStockValue(1)
StockValue = MaterialClass.GetOnDateClosingStockValue(m_ToUSDate)
If StockValue = 0 Then StockValue = OpStock

'print the closing stock value
If StockValue Or OpStock Then Call BALPrintClosingStockValue(StockValue, OpStock, False)

'Add the Stockvalue to the asset Total
AssetTotal(1) = AssetTotal(1) + StockValue



' This will equalise the rows of both the tables
BALFillEmptyRows

' Print the the Profit
'BALPrintProfitorLoss (Profit)

'Now Get the Last Year Profit Or Loss
gDbTrans.SqlStmt = "SELECT SUM(OPAmount) as ParentOB" & _
        " FROM qryOPHeads a,ParentHeads b" & _
        " WHERE a.ParentID=b.ParentID" & _
        " And a.ParentID = " & parProfitORLoss
Dim PrevProfit As Currency
Dim ProfitDist As Currency
If gDbTrans.Fetch(rstOBParents, adOpenDynamic) > 0 Then
    PrevProfit = FormatField(rstOBParents("ParentOB"))
    If rstSum.RecordCount > 0 Then _
        If gDbTrans.FindRecord(rstSum, "ParentID=" & parProfitORLoss) Then _
                        Profit = FormatField(rstSum.Fields("SumAmount"))

    'Now Get the Disrtibuted amount of the Previous profit
    If PrevProfit > 0 Then
        gDbTrans.SqlStmt = "SELECT SUM(Debit) as TotalAmount FROM ACCTRANS" & _
                " WHERE TransDate >= #" & FinUSFromDate & "#" & _
                " AND TransDate <= #" & m_ToUSDate & "#" & _
                " AND HeadID = " & wis_PrevProfitHeadID
        If gDbTrans.Fetch(rstSum, adOpenForwardOnly) > 0 Then
            ProfitDist = FormatField(rstSum("TotalAmount"))
            PrevProfit = PrevProfit - ProfitDist
        End If
    
    
    End If
    PrevProfit = PrevProfit + Profit
    
End If
'Now Get the Current Yer Profit
Profit = GetNetProfit
Call PrintBalProfitorLoss(PrevProfit, Profit, ProfitDist, False)

Profit = Profit + PrevProfit
If Profit >= 0 Then
    LiabilityTotal(1) = LiabilityTotal(1) + Abs(Profit)
Else
    AssetTotal(1) = AssetTotal(1) + Abs(Profit)
End If

' This will equalise the rows of both the tables
BALFillEmptyRows

Call BALSetBalanceSheetDifference(AssetTotal, LiabilityTotal)
    
If AssetTotal(1) > LiabilityTotal(1) Then

    Call BALPrintGrandTotals(AssetTotal)
    
ElseIf LiabilityTotal(1) >= AssetTotal(1) Then

    Call BALPrintGrandTotals(LiabilityTotal)
    
End If

Set rstOBHeads = Nothing
Set rstSum = Nothing
Set rstHeads = Nothing
Set rstOBParents = Nothing

Set MaterialClass = Nothing


' Exit Function
ShowConditionBalanceSheet = True

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowConditionBalanceSheet :" & vbCrLf & Err.Description
        
    Set rstOBHeads = Nothing
    Set rstSum = Nothing
    Set rstHeads = Nothing
    Set rstOBParents = Nothing
    
End Function


Public Function ShowConditionBalanceSheet_New() As Boolean

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstOBHeads As ADODB.Recordset
Dim rstHeads As ADODB.Recordset

Dim rstSum As ADODB.Recordset

Dim rstOBParents As ADODB.Recordset


Dim AssetTotal(2) As Currency
Dim LiabilityTotal(2) As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim StockValue As Currency

Dim MaterialClass As clsMaterial
Dim PrintStatus As wis_PrintStatus

Dim SqlStr As String

Const cnstLIABILITY = 0
Const cnstASSET = 1

' Intialiase the Function
ShowConditionBalanceSheet_New = False


'Now Check whether the Negaitive amount has to shown in in
'same side or other side
Dim SetUp As New clsSetup
m_OtherDetail = CBool(SetUp.ReadSetupValue("General", "NegativeInBalanceSheet", "False"))
Set SetUp = Nothing


Call BALCreateViewsForConditionalBalanceSheet

' Opening Balances
gDbTrans.SqlStmt = " SELECT * FROM qryOPHeads"
If gDbTrans.Fetch(rstOBHeads, adOpenDynamic) < 1 Then Exit Function

SqlStr = " SELECT SUM(OPAmount) as ParentOB,a.ParentID,b.ParentName," & _
        " b.PrintOrder,b.PrintDetailed,b.AccountType" & _
        " FROM qryOPHeads a,ParentHeads b" & _
        " WHERE a.ParentID=b.ParentID" & _
        " And (AccountType = " & Asset & " OR AccountType = " & Liability & ")" & _
        " And a.ParentID <> " & parProfitORLoss & _
        " GROUP BY a.ParentID,b.ParentName," & _
         " b.PrintDetailed,AccountType,b.PrintOrder " & _
        " ORDER By PrintOrder"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstOBParents, adOpenDynamic) < 1 Then Exit Function


SqlStr = " SELECT * FROM qryLiabilities"
SqlStr = SqlStr & " UNION " & "SELECT * FROM qryAssets"

gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstHeads, adOpenDynamic)
'Get the Liability Heads Except
SqlStr = " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryLiabilities" & _
        " Where parentID <> " & parProfitORLoss & _
        " GROUP BY ParentID"
SqlStr = SqlStr & " UNION " & _
        " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryAssets" & _
        " GROUP BY ParentID"


gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstSum, adOpenDynamic)

' Start the Html
BALStartBalanceSheet

' Start the Both Headers
BALSetAssetHead (True)
BALSetLiabilityHeader (True)

' Main Loop
Do

    'If rstOBAssetParents.EOF And rstOBLiabilityParents.EOF Then Exit Do
    If rstOBParents.EOF Then Exit Do
    
    If Not rstOBParents.EOF Then
        
        PrintStatus = rstOBParents.Fields("PrintDetailed")
    
        'Now Filter the Heads Record set for the bala sheet
        rstHeads.Filter = adFilterNone
        rstHeads.Filter = "ParentID=" & rstOBParents.Fields("ParentID")
        
        'If rstHeads.RecordCount Then
            Call BALPutAssetorLiability(rstSum, rstOBHeads, _
                    rstOBParents, rstHeads, LiabilityTotal, AssetTotal, CBool(m_OtherDetail), True)
        'End If
        
        If Not rstOBParents.EOF Then rstOBParents.MoveNext
    End If
    
Loop


'Get the Closing Stock value from the database
Set MaterialClass = New clsMaterial
Dim OpStock As Currency
OpStock = MaterialClass.GetOpeningStockValue(1)
StockValue = MaterialClass.GetOnDateClosingStockValue(m_ToUSDate)
If StockValue = 0 Then StockValue = OpStock

'print the closing stock value
If StockValue Then Call BALPrintClosingStockValue(StockValue, OpStock, True)

'Add the Stockvalue to the asset Total
AssetTotal(1) = AssetTotal(1) + StockValue


' This will equalise the rows of both the tables
BALFillEmptyRows (True)

' Print the the Profit
'BALPrintProfitorLoss (Profit)

'Now Get the Last Year Profit Or Loss
gDbTrans.SqlStmt = "SELECT SUM(OPAmount) as ParentOB" & _
        " FROM qryOPHeads a,ParentHeads b" & _
        " WHERE a.ParentID=b.ParentID" & _
        " And a.ParentID = " & parProfitORLoss
Dim PrevProfit As Currency
Dim ProfitDist As Currency
If gDbTrans.Fetch(rstOBParents, adOpenDynamic) > 0 Then
    PrevProfit = FormatField(rstOBParents("ParentOB"))
    If rstSum.RecordCount > 0 Then _
        If gDbTrans.FindRecord(rstSum, "ParentID=" & parProfitORLoss) Then _
                        Profit = FormatField(rstSum.Fields("SumAmount"))

    'Now Get the Disrtibuted amount of the Previous profit
    If PrevProfit > 0 Then
        gDbTrans.SqlStmt = "SELECT SUM(Debit) as TotalAmount FROM ACCTRANS" & _
                " WHERE TransDate >= #" & FinUSFromDate & "#" & _
                " AND TransDate <= #" & m_ToUSDate & "#" & _
                " AND HeadID = " & wis_PrevProfitHeadID
        If gDbTrans.Fetch(rstSum, adOpenForwardOnly) > 0 Then
            ProfitDist = FormatField(rstSum("TotalAmount"))
            PrevProfit = PrevProfit - ProfitDist
        End If
    
    End If
    PrevProfit = PrevProfit + Profit
    
End If
'Now Get the Current Yer Profit
Profit = GetNetProfit
Call PrintBalProfitorLoss(PrevProfit, Profit, ProfitDist, True)

Profit = Profit + PrevProfit
If Profit >= 0 Then
    LiabilityTotal(1) = LiabilityTotal(1) + Abs(Profit)
Else
    AssetTotal(1) = AssetTotal(1) + Abs(Profit)
End If
If PrevProfit >= 0 Then
    LiabilityTotal(0) = LiabilityTotal(0) + Abs(PrevProfit)
Else
    AssetTotal(0) = AssetTotal(0) + Abs(PrevProfit)
End If

' This will equalise the rows of both the tables
BALFillEmptyRows (True)

Call BALSetBalanceSheetDifference(AssetTotal, LiabilityTotal, True)
    
If AssetTotal(1) > LiabilityTotal(1) Then

    Call BALPrintGrandTotals(AssetTotal, True)
    
ElseIf LiabilityTotal(1) >= AssetTotal(1) Then

    Call BALPrintGrandTotals(LiabilityTotal, True)
    
End If

Set rstOBHeads = Nothing
Set rstSum = Nothing
Set rstHeads = Nothing
Set rstOBParents = Nothing

Set MaterialClass = Nothing


' Exit Function
ShowConditionBalanceSheet_New = True


'm_frmWeb.web.Document.body.innerhtml
Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowConditionBalanceSheet_new :" & vbCrLf & Err.Description
        
    Set rstOBHeads = Nothing
    Set rstSum = Nothing
    Set rstHeads = Nothing
    Set rstOBParents = Nothing
    
End Function

'This will equaliase the row heights
Private Sub BALEqualRowHeights()

Dim CurrentRow As HTMLTableRow
Dim assetTable As HTMLTable
Dim liabilityTable As HTMLTable

Dim SpaceToAdd As Integer

Dim Rows As Integer
Dim Row As Integer

With m_WebRow
    
    If .cells(0).All.Liability.Rows.Length <> .cells(1).All.Asset.Rows.Length Then Exit Sub
    
    Set liabilityTable = .cells(0).All.Liability
    Set assetTable = .cells(1).All.Asset
    
    Rows = .cells(0).All.Liability.Rows.Length
    
End With


For Row = 0 To Rows - 1
    
    'If Len(LiabilityTable.Rows(Row).cells(1).innerText) <> Len(AssetTable.Rows(Row).cells(1).innerText) Then
                    
    SpaceToAdd = Len(liabilityTable.Rows(Row).cells(0).innerText) - Len(assetTable.Rows(Row).cells(0).innerText)
    
    If SpaceToAdd <> 0 Then
        If SpaceToAdd < 0 Then
            liabilityTable.Rows(Row).cells(0).innerText = liabilityTable.Rows(Row).cells(0).innerText & Space(Abs(SpaceToAdd))
        Else
            assetTable.Rows(Row).cells(0).innerText = assetTable.Rows(Row).cells(0).innerText & Space(Abs(SpaceToAdd))
        End If
    End If
Next Row
    
Set m_WebRow = InsertRowAndCellSToTable(liabilityTable, 3)

End Sub

Public Function ShowReport2(RepType As wisReports) As Currency

' Handle Error
On Error GoTo Hell:

' Declarations

Dim rstExpense As ADODB.Recordset
Dim rstIncome As ADODB.Recordset

Dim rstSumExpense As ADODB.Recordset
Dim rstSumIncome As ADODB.Recordset

Dim IncomeTotal As Currency
Dim ExpenseTotal As Currency
Dim ParentAmount As Currency
Dim Profit As Currency
Dim GrossProfit As Currency
Dim strPass As String

Const cnstExpense = 0
Const cnstINCOME = 1

Set m_WebDoc = Nothing

' Create Required Views for the Balance Sheet Report
If RepType = wisProfitLossStatement Then
    GrossProfit = GetTradingProfit
Else
    Dim AccTransClass As New clsAccTrans
    GrossProfit = AccTransClass.GetOpBalance(wis_CashHeadID, m_FromUSDate)
    Set AccTransClass = Nothing
End If
Call CreateViewsForReport2(RepType)

' Opening Balances
gDbTrans.SqlStmt = "SELECT * FROM QryPayments"

Call gDbTrans.Fetch(rstExpense, adOpenDynamic)


gDbTrans.SqlStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryPayments" & _
                   " GROUP BY ParentID,ParentName"

Call gDbTrans.Fetch(rstSumExpense, adOpenDynamic)

                   
' Assets
gDbTrans.SqlStmt = " SELECT * FROM QryReceipts"
                   
Call gDbTrans.Fetch(rstIncome, adOpenDynamic)

gDbTrans.SqlStmt = " SELECT SUM(Amount) as SumAmount,ParentID,ParentName" & _
                   " FROM QryReceipts" & _
                   " GROUP BY ParentID,ParentName"
                   
Call gDbTrans.Fetch(rstSumIncome, adOpenDynamic)

'Get whether to Show the Total Or No
Dim SetUp As New clsSetup
m_OtherDetail = IIf(UCase(SetUp.ReadSetupValue("General", "ShowRPorPLTotal", "True")) = "TRUE", True, False)
Set SetUp = Nothing

' Start the Html
'PLStartProfitAndLoss
Call StartReport2(RepType)

' Start the Both Headers
'PLSetExpenseHeader
SetPaymentHeader (RepType)

'PLSetIncomeHead
SetReceiptHeader (RepType)

' If profit is loss
strPass = GetResourceString(284)

If RepType = wisProfitLossStatement Then
    strPass = GetResourceString(162, IIf(GrossProfit >= 0, 403, 404))
Else
    GrossProfit = GrossProfit * -1
End If
Dim isRightSide  As Boolean

Call PrintPLorOP(Abs(GrossProfit), strPass, IIf(GrossProfit > 0, True, False))
    
If GrossProfit <= 0 Then
    ExpenseTotal = ExpenseTotal + Abs(GrossProfit)
Else
    IncomeTotal = IncomeTotal + Abs(GrossProfit)
End If

' This will equalise the rows of both the tables
PLFillEmptyRows

isRightSide = IIf(RepType = wisProfitLossStatement, True, False)
' Main Loop
Do

    If rstSumIncome.EOF And rstSumExpense.EOF Then Exit Do
    
    If Not rstSumExpense.EOF Then
    
        ParentAmount = PutReceiptORPayment(rstSumExpense, rstExpense, isRightSide)
        
        If RepType <> wisProfitLossStatement Then ParentAmount = ParentAmount * -1
        If ParentAmount > 0 Then
            ExpenseTotal = ExpenseTotal + Abs(ParentAmount)
        Else
            IncomeTotal = IncomeTotal + Abs(ParentAmount)
        End If
        
    End If
    
    If Not rstSumIncome.EOF Then
    
        ParentAmount = PutReceiptORPayment(rstSumIncome, rstIncome, Not isRightSide)
        If RepType <> wisProfitLossStatement Then ParentAmount = ParentAmount * -1
        If ParentAmount > 0 Then
            IncomeTotal = IncomeTotal + Abs(ParentAmount)
        Else
            ExpenseTotal = ExpenseTotal + Abs(ParentAmount)
        End If
    End If
    
    
    If Not rstSumExpense.EOF Then rstSumExpense.MoveNext
    If Not rstSumIncome.EOF Then rstSumIncome.MoveNext
    
Loop


Profit = IncomeTotal - ExpenseTotal

'Print the the Profit
strPass = GetResourceString(285)
If RepType = wisProfitLossStatement Then
    strPass = GetResourceString(332) & " " & _
                GetResourceString(IIf(Profit >= 0, 403, 404))
End If

Call PrintPLorOP(Abs(Profit), strPass, IIf(Profit >= 0, False, True))


' This will equalise the rows of both the tables
PLFillEmptyRows

If IncomeTotal > ExpenseTotal Then

    Call PLPrintGrandTotals(IncomeTotal)
    
ElseIf ExpenseTotal >= IncomeTotal Then

    Call PLPrintGrandTotals(ExpenseTotal)
    
End If

Set rstExpense = Nothing
Set rstIncome = Nothing

Set rstSumExpense = Nothing
Set rstSumIncome = Nothing


' Exit Function
ShowReport2 = Profit


Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "ShowReport2 :" & vbCrLf & Err.Description
        
    Set rstExpense = Nothing
    Set rstIncome = Nothing
    
    Set rstSumExpense = Nothing
    Set rstSumIncome = Nothing
    
End Function


Private Function GetPreviousYearProfit() As Currency

Dim rstOBParents As Recordset
Dim rstSum As Recordset
Dim Profit As Currency
Dim PrevProfit As Currency

'Get the Liability Heads Except
gDbTrans.SqlStmt = " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryLiabilities" & _
        " Where parentID <> " & parProfitORLoss & _
        " GROUP BY ParentID"
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " UNION " & _
        " SELECT SUM(Amount) as SumAmount,ParentID" & _
        " FROM QryAssets" & _
        " GROUP BY ParentID"

Call gDbTrans.Fetch(rstSum, adOpenDynamic)

'Now Get the Last Year Profit Or Loss
gDbTrans.SqlStmt = "SELECT SUM(OPAmount) as ParentOB" & _
        " FROM qryOPHeads a,ParentHeads b" & _
        " WHERE a.ParentID=b.ParentID" & _
        " And a.ParentID = " & parProfitORLoss

If gDbTrans.Fetch(rstOBParents, adOpenDynamic) > 0 Then
    PrevProfit = FormatField(rstOBParents("ParentOB"))
    If rstSum.RecordCount > 0 Then _
        If gDbTrans.FindRecord(rstSum, "ParentID=" & parProfitORLoss) Then _
                        Profit = FormatField(rstSum.Fields("SumAmount"))

    'Now Get the Disrtibuted amount of the Previous profit
    If PrevProfit > 0 Then
        gDbTrans.SqlStmt = "SELECT SUM(Debit) as TotalAmount FROM ACCTRANS" & _
                " WHERE TransDate >= #" & FinUSFromDate & "#" & _
                " AND TransDate <= #" & m_ToUSDate & "#" & _
                " AND HeadID = " & wis_PrevProfitHeadID
        If gDbTrans.Fetch(rstSum, adOpenForwardOnly) > 0 Then _
            PrevProfit = PrevProfit - FormatField(rstSum("TotalAmount"))
    
    End If
    PrevProfit = PrevProfit + Profit
End If

GetPreviousYearProfit = PrevProfit

End Function

'This will return the net proft or net loss
Private Function GetNetProfit() As Currency

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstExpense As ADODB.Recordset
Dim rstIncome As ADODB.Recordset

Dim IncomeTotal As Currency
Dim ExpenseTotal As Currency
Dim Profit As Currency
Dim GrossProfit As Currency

' Intialiase the Function
GetNetProfit = 0

' Create Required Views for the Balance Sheet Report

'GrossProfit = ShowTradingAccount
GrossProfit = GetTradingProfit
 
Call NetProfitViews

' Opening Balances

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryExpenses"
                   
Call gDbTrans.Fetch(rstExpense, adOpenDynamic)

                   
' Assets
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryIncomes"
                   
Call gDbTrans.Fetch(rstIncome, adOpenDynamic)

ExpenseTotal = FormatField(rstExpense.Fields("NetExpenses"))
IncomeTotal = FormatField(rstIncome.Fields("NetIncome"))

If GrossProfit <= 0 Then
    ' If profit is loss
    ExpenseTotal = ExpenseTotal + Abs(GrossProfit)
Else
    'if profit is Profit
    IncomeTotal = IncomeTotal + Abs(GrossProfit)
End If
  
'incometotal=grossprofit+
Profit = IncomeTotal - ExpenseTotal

Set rstExpense = Nothing
Set rstIncome = Nothing

' Exit Function
GetNetProfit = Profit

Exit Function


Hell:
    If Err.Number > 0 Then MsgBox "Getnetprofit :" & vbCrLf & Err.Description
       
    Set rstExpense = Nothing
    Set rstIncome = Nothing
    
    
End Function

Public Function ShowReportDate(Optional BothDates As Boolean = True) As Boolean

On Error GoTo Hell:

ShowReportDate = False

Set m_frmRptDt = New frmRptDt

With m_frmRptDt
    .txtStDate = FinIndianFromDate
    .txtEndDate = gStrDate
    .txtStDate.Enabled = BothDates
    .txtStDate.BackColor = IIf(BothDates, wisWhite, wisGray)
End With
' User should not change the opening date
'm_frmRptDt.txtStDate.Locked = True

m_frmRptDt.Show vbModal

Set m_frmRptDt = Nothing

' If no date is entered then
If FromRpDate = "" And ToRpDate = "" Then Exit Function

LoadFrmWeb
IntialiseWebDocument

ShowReportDate = True

Exit Function

Hell:

    Set m_frmRptDt = Nothing
    
End Function
Public Function ShowReportForm() As Boolean

On Error GoTo Hell:

ShowReportForm = False

If m_frmAccReports Is Nothing Then Set m_frmAccReports = New frmAccReports

FromRpDate = ""
ToRpDate = ""

m_frmAccReports.txtDate1 = FinIndianFromDate
m_frmAccReports.txtDate2 = gStrDate

Restart:
Set m_frmWeb = New frmWeb
Unload m_frmWeb

LoadFrmWeb
Load m_frmWeb
m_frmAccReports.Show vbModal
    
' If No Date is Entered Then
If FromRpDate = "" And ToRpDate = "" Then
    Set m_frmAccReports = Nothing
    Exit Function
End If

' This Will Initialise the Web Document
IntialiseWebDocument

' this will process the selected report and will feed it to webcontrol

If Not ProcessSelectedReport Then GoTo Restart

ShowReportForm = True

ShowCurrentReport

Set m_WebDoc = Nothing
Set m_frmWeb = Nothing
GoTo Restart

Exit Function

Hell:
    MsgBox " There is error"
    
    'Set m_frmAccReports = Nothing
        
End Function

Private Sub PauseApplication()

Dim Start

Start = Timer
Do
    'it will be in the loop
    If Timer >= Start + 1 Then Exit Do
Loop



End Sub

Public Sub ShowCurrentReport()

With m_frmWeb
    .Show vbModal
End With

End Sub

Public Function ShowRPReport() As Boolean

Dim rstRPDetails As ADODB.Recordset
Dim rstParents As ADODB.Recordset

Dim OpeningCash As Currency
Dim ClosingStock As Currency
Dim Receipts As Currency
Dim Payments As Currency
Dim SumReceipts As Currency
Dim SumPayments As Currency

Dim ParentID As Long
Dim headID As Long
Dim SlNo As Long

Dim ParentName As String
Dim headName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const TableName = "RPTable"
Const QRYRP = "QryRPDetails"
Const totalCols = 4

' Handle Error
On Error GoTo Hell:

ShowRPReport = False

' Start the Receipts and Payments
RPStartRP

' Put the Opening Cash to the Html
OpeningCash = RPPutOpeningCash(TableName, FromRpDate)

' create views for the receipts and payments
RPCreateViews (QRYRP)

gDbTrans.SqlStmt = " SELECT ParentName,ParentID " & _
                   " FROM Parentheads WHERE AccountType > 0  " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT Sum(a.Receipt) AS Receipts, SUM(a.Payment) AS Payments," & _
                   " b.ParentID, a.RpHeadid,b.HeadName " & _
                   " FROM " & QRYRP & " AS a, Heads AS b " & _
                   " WHERE a.RPHeadId = b.HeadID " & _
                   " GROUP BY b.ParentID,a.RPHeadID,b.HeadName"

gDbTrans.SqlStmt = "  SELECT SUM(Debit) As Payments,SUM(Credit) as Receipts,a.Headid as RpHeadID,b.HeadName,b.ParentID" & _
                   " FROM AccTrans a,Heads b,ParentHeads c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND a.HeadID <> " & wis_CashHeadID & _
                   " AND TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND " & "#" & m_ToUSDate & "#" & _
                   " GROUP BY a.HeadId,b.HeadName,b.ParentID " & _
                   " ORDER BY b.ParentID,a.HeadID"
                   
Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable

SlNo = 0

' Main Loop
Do While Not rstParents.EOF
    
    ParentID = FormatField(rstParents.Fields("ParentID"))
    ParentName = FormatField(rstParents.Fields("ParentName"))
    
    With rstRPDetails
        
        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID
        
        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
            
            CurrentRow.Style.fontWeight = "Bold"
            
            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
                .Style.fontWeight = "Bold"
                .Style.FONTSIZE = 14
                If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
            End With
        End If
        Do While Not .EOF
            
            headID = FormatField(.Fields("RPHeadID"))
            headName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
            
            If Receipts Or Payments Then
                SlNo = SlNo + 1
                Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
                With CurrentRow.cells(0)
                    .innerText = SlNo
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(1)
                    .innerText = headName
                    .Align = "Left"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(2)
                    .innerText = FormatCurrency(Receipts)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(3)
                    .innerText = FormatCurrency(Payments)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
            End If
            
            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments
            
            .MoveNext
        Loop
    End With
    
    rstParents.MoveNext
    
Loop


' Now Put Closing Balance
Call RPputClosingCash((OpeningCash + SumReceipts) - SumPayments, TableName)

Set rstRPDetails = Nothing
Set rstParents = Nothing

''Mow Show the Total Amount

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)

With CurrentRow.cells(1)
    .innerText = GetResourceString(286)
    .Align = "Left"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With CurrentRow.cells(2)
    .innerText = FormatCurrency(OpeningCash + SumReceipts)
    .Align = "Right"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With CurrentRow.cells(3)
    .innerText = FormatCurrency(OpeningCash + SumReceipts)
    .Align = "Right"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

ShowRPReport = True

Exit Function

Hell:
    MsgBox "ShowRPReport: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    
End Function


Public Function ShowReport1(RepType As wisReports) As Boolean

Dim rstRPDetails As ADODB.Recordset
Dim rstParents As ADODB.Recordset

Dim OpeningCash As Currency
Dim ClosingStock As Currency
Dim Receipts As Currency
Dim Payments As Currency
Dim SumReceipts As Currency
Dim SumPayments As Currency

Dim ParentID As Long
Dim headID As Long
Dim SlNo As Long

Dim ParentName As String
Dim headName As String

Dim CurrentTable As HTMLTable
Dim CurrentRow As HTMLTableRow

Const TableName = "RPTable"
Const QRYRP = "QryRPDetails"
Const totalCols = 4

' Handle Error
On Error GoTo Hell:

ShowReport1 = False

' Start the Receipts and Payments
'RPStartRP
StartReport1 (RepType)

' Put the Opening Cash to the Html
If RepType = wisDebitCreditStatement Then
    Dim AccTransClass  As clsAccTrans
    Set AccTransClass = New clsAccTrans
    OpeningCash = AccTransClass.GetOpBalance(wis_CashHeadID, m_FromUSDate)
    Set AccTransClass = Nothing
'    OpeningCash = RPPutOpeningCash(TableName, FromRpDate)
Else
    OpeningCash = GetTradingProfit
End If
' create views for the receipts and payments
'RPCreateViews (QRYRP)
'Call CreateViewsForReport1(RepType)

gDbTrans.SqlStmt = " SELECT ParentName,ParentID " & _
                   " FROM Parentheads WHERE AccountType > 0  " & _
                   " ORDER BY ParentID"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

gDbTrans.SqlStmt = "SELECT SUM(Debit) As Payments,SUM(Credit) as Receipts, " & _
                        " a.Headid as RpHeadID,b.HeadName,b.ParentID" & _
                   " FROM AccTrans a,Heads b,ParentHeads c " & _
                   " WHERE a.HeadID = b.HeadID " & _
                   " AND b.ParentID=c.ParentID " & _
                   " AND a.HeadID <> " & wis_CashHeadID & _
                   " AND TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND " & "#" & m_ToUSDate & "#"
If RepType = wisProfitLossStatement Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
                    " And (c.AccountType = " & Profit & _
                        " OR c.AccountType = " & Loss & ")" & _
                    " And c.parentID <> " & parTradingExpense & _
                    " And c.ParentID <> " & parTradingIncome
End If

gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
                   " GROUP BY a.HeadId,b.HeadName,b.ParentID " & _
                   " ORDER BY b.ParentID,a.HeadID"

Call gDbTrans.Fetch(rstRPDetails, adOpenDynamic)

Set CurrentTable = m_WebTable

SlNo = 1

'Put Opening cash or Trading Profit
headName = GetResourceString(284) '"Opening Cash Balance"
If RepType = wisProfitLossStatement Then
    headName = GetResourceString(162, IIf(OpeningCash >= 0, 403, 404))  '"Trading profit
End If

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
With CurrentRow.cells(0)
    .innerText = SlNo
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = IIf(gLangOffSet, 16, 14)
    
End With

With CurrentRow.cells(1)
    .innerText = headName
    .Align = "Left"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = IIf(gLangOffSet, 15, 13)
    
End With

With CurrentRow.cells(IIf(OpeningCash >= 0, 2, 3))
    .innerText = FormatCurrency(Abs(OpeningCash))
    .Align = "Right"
    .Style.fontWeight = "Normal"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With



' Main Loop
Do While Not rstParents.EOF
    
    ParentID = FormatField(rstParents.Fields("ParentID"))
    ParentName = FormatField(rstParents.Fields("ParentName"))
    
    With rstRPDetails
        
        .Filter = adFilterNone
        .Filter = "ParentID=" & ParentID
        
        If rstRPDetails.RecordCount > 0 Then
            Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
            
            CurrentRow.Style.fontWeight = "Bold"
            
            With CurrentRow.cells(1)
                .innerText = ParentName
                .Align = "Left"
                .Style.fontWeight = "Bold"
                .Style.FONTSIZE = 14
                If gLangOffSet <> 0 Then .Style.FONTSIZE = 16
            End With
        End If
        Do While Not .EOF
            
            headID = FormatField(.Fields("RPHeadID"))
            headName = FormatField(.Fields("HeadName"))
            Receipts = FormatField(.Fields("Receipts"))
            Payments = FormatField(.Fields("Payments"))
            
            If Receipts Or Payments Then
                SlNo = SlNo + 1
                Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
                With CurrentRow.cells(0)
                    .innerText = SlNo
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(1)
                    .innerText = headName
                    .Align = "Left"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(2)
                    .innerText = FormatCurrency(Receipts)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
                
                With CurrentRow.cells(3)
                    .innerText = FormatCurrency(Payments)
                    .Align = "Right"
                    .Style.fontWeight = "Normal"
                    .Style.FONTSIZE = 13
                    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
                End With
            End If
            
            SumReceipts = SumReceipts + Receipts
            SumPayments = SumPayments + Payments
            
            .MoveNext
        Loop
    End With
    
    rstParents.MoveNext
    
Loop


' Now Put Closing Balance
'Call RPputClosingCash((OpeningCash + SumReceipts) - SumPayments, TableName)
Dim closingCash As Currency

closingCash = (OpeningCash + SumReceipts) - SumPayments
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable)

headName = GetResourceString(285)
If RepType = wisProfitLossStatement Then
headName = GetResourceString(332) & " " & _
        GetResourceString(IIf(OpeningCash >= 0, 403, 404))
End If
With CurrentRow.cells(1)
    .innerText = headName  '"Closing Balance"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

If closingCash >= 0 Then
    With CurrentRow.cells(3)
        .innerText = FormatCurrency(closingCash)
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
Else
    With CurrentRow.cells(2)
        .innerText = FormatCurrency(Abs(closingCash))
        .Align = "Right"
        .Style.fontWeight = "Bold"
        .Style.FONTSIZE = 15
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
    End With
End If


Set rstRPDetails = Nothing
Set rstParents = Nothing

''Mow Show the Total Amount

Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)
Set CurrentRow = InsertRowAndCellSToTable(CurrentTable, totalCols)

With CurrentRow.cells(1)
    .innerText = GetResourceString(286)
    .Align = "Left"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

closingCash = IIf(OpeningCash + SumReceipts > SumPayments, OpeningCash + SumReceipts, SumPayments)
With CurrentRow.cells(2)
    .innerText = FormatCurrency(closingCash)
    .Align = "Right"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

With CurrentRow.cells(3)
    .innerText = FormatCurrency(closingCash)
    .Align = "Right"
    .Style.fontWeight = "bold"
    .Style.FONTSIZE = 13
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
End With

ShowReport1 = True

Exit Function

Hell:
    MsgBox "ShowRPReport: " & vbCrLf & Err.Description
    
    Set rstRPDetails = Nothing
    Set rstParents = Nothing
    
End Function



' This function will show the TradingAccount
'Lingappa Sindhanur
Public Function ShowTradingAccount() As Currency

' Handle Error
On Error GoTo Hell:

Dim rstTradingIncome As ADODB.Recordset
Dim rstTradingExpense As ADODB.Recordset

Dim Profit As Currency
Dim Purchase As Currency
Dim Sales As Currency
Dim OpeningStock As Currency
Dim ClosingStock As Currency

Dim TotalPurchase As Currency
Dim TotalSales As Currency

Dim IncomeHeadName As String
Dim ExpenseHeadName As String

Dim MaterialClass As clsMaterial

' Intialiase the Function
ShowTradingAccount = 0

gDbTrans.SqlStmt = " SELECT (SUM(Debit)-SUM(Credit)) AS Purchase,A.HeadID,HeadName" & _
                    " FROM AccTrans AS a, Heads AS b " & _
                    " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_PurchaseParentID & _
                    " OR b.ParentID  = " & wis_TradingExpParentID & " ) " & _
                    " AND  a.TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                    " AND # " & m_ToUSDate & "#" & _
                    " GROUP BY A.HeadID, HeadName "

Call gDbTrans.Fetch(rstTradingExpense, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT (SUM(Credit)-SUM(Debit)) AS Sales ,A.HeadID,HeadName" & _
                   " FROM AccTrans AS a, Heads AS b " & _
                   " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_SalesParentID & _
                   " OR b.ParentID  = " & wis_TradingIncParentID & " ) " & _
                   " AND  a.TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND # " & m_ToUSDate & "#" & _
                   " GROUP BY A.HeadID, HeadName "

Call gDbTrans.Fetch(rstTradingIncome, adOpenForwardOnly)


Set MaterialClass = New clsMaterial
If m_FromRpDate = FinIndianFromDate Then
    OpeningStock = MaterialClass.GetOpeningStockValue
Else
    OpeningStock = MaterialClass.GetOnDateClosingStockValue(DateAdd("d", -1, m_FromUSDate))
End If

ClosingStock = MaterialClass.GetOnDateClosingStockValue(m_ToUSDate) '(ToRpDate)

Set MaterialClass = Nothing

' Start the Page and set the Headers
Call TRStartHtmlPage

Call TRPutOpeningStock(OpeningStock)

Do

    If rstTradingExpense.EOF And rstTradingIncome.EOF Then Exit Do
    
    
    'Get expense Details
    If Not rstTradingExpense.EOF Then
        Purchase = FormatField(rstTradingExpense.Fields("Purchase"))
        ExpenseHeadName = FormatField(rstTradingExpense.Fields("HeadName"))
    End If
    
    'Get Income  Details
    If Not rstTradingIncome.EOF Then
        Sales = FormatField(rstTradingIncome.Fields("Sales"))
        IncomeHeadName = FormatField(rstTradingIncome.Fields("HeadName"))
    End If
    
    Call TRPutSalesPurchases(Purchase, Sales, ExpenseHeadName, IncomeHeadName)
    ExpenseHeadName = ""
    IncomeHeadName = ""
    TotalPurchase = TotalPurchase + Purchase: Purchase = 0
    TotalSales = TotalSales + Sales: Sales = 0 '+ ClosingStock

    If Not rstTradingExpense.EOF Then rstTradingExpense.MoveNext
    If Not rstTradingIncome.EOF Then rstTradingIncome.MoveNext
    
Loop


Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

Call TRPutClosingStock(ClosingStock)


TotalPurchase = TotalPurchase + OpeningStock
TotalSales = TotalSales + ClosingStock
Call PutTradingProfitORLoss(TotalPurchase, TotalSales)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

ShowTradingAccount = TotalSales - TotalPurchase

Exit Function

Hell:

    If Err.Number > 0 Then MsgBox "Trading Account:" & vbCrLf & Err.Description
    
'Resume
End Function


' This function will return the Trading Profit or Loss
'Lingappa Sindhanur
Private Function GetTradingProfit() As Currency

' Handle Error
On Error GoTo Hell:

Dim rstTradingIncome As ADODB.Recordset
Dim rstTradingExpense As ADODB.Recordset

Dim OpeningStock As Currency
Dim ClosingStock As Currency
Dim TotalPurchase As Currency
Dim TotalSales As Currency

Dim MaterialClass As clsMaterial

' Intialiase the Function
GetTradingProfit = 0

gDbTrans.SqlStmt = "SELECT (SUM(Debit)-SUM(Credit)) As TradingExpense" & _
                    " FROM AccTrans AS a, Heads AS b " & _
                    " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_PurchaseParentID & _
                    " OR b.ParentID  = " & wis_TradingExpParentID & " ) " & _
                    " AND  a.TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                    " AND # " & m_ToUSDate & "#"
                    
Call gDbTrans.Fetch(rstTradingExpense, adOpenForwardOnly)

gDbTrans.SqlStmt = " SELECT (SUM(Credit)-SUM(Debit)) AS TradingIncome" & _
                   " FROM AccTrans AS a, Heads AS b " & _
                   " WHERE a.HeadID = b.HeadID AND ( b.ParentID =" & wis_SalesParentID & _
                   " OR b.ParentID  = " & wis_TradingIncParentID & " ) " & _
                   " AND  a.TransDate BETWEEN " & "#" & m_FromUSDate & "#" & _
                   " AND # " & m_ToUSDate & "#"

Call gDbTrans.Fetch(rstTradingIncome, adOpenForwardOnly)

TotalPurchase = FormatField(rstTradingExpense.Fields("TradingExpense"))
TotalSales = FormatField(rstTradingIncome.Fields("TradingIncome"))

Set MaterialClass = New clsMaterial
If DateDiff("D", m_FromUSDate, FinUSFromDate) = 0 Then
    OpeningStock = MaterialClass.GetOpeningStockValue
Else
    OpeningStock = MaterialClass.GetOnDateClosingStockValue(DateAdd("D", -1, m_FromUSDate))
End If
ClosingStock = MaterialClass.GetOnDateClosingStockValue(m_ToUSDate)

Set MaterialClass = Nothing
If TotalPurchase = 0 And TotalSales = 0 And ClosingStock = 0 Then ClosingStock = OpeningStock
TotalPurchase = TotalPurchase + OpeningStock
TotalSales = TotalSales + ClosingStock

GetTradingProfit = TotalSales - TotalPurchase

Exit Function

Hell:

    If Err.Number > 0 Then MsgBox "Trading Account:" & vbCrLf & Err.Description
    
    
End Function

Private Sub BALStartBalanceSheet()

Dim strHeader As String
Dim DOMElement
Dim DOMTbody

' set the web document to htmldoc

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document

strHeader = GetResourceString(408) '"Balance Sheet for the Duration"


Dim tmpDate As Date
tmpDate = m_FromRpDate
m_FromRpDate = m_ToRpDate
Call SetReportTitle(strHeader)
m_FromRpDate = tmpDate
With m_WebDoc
    Set m_WebTable = .createElement("Table")
    
    Set DOMElement = .createElement("Tbody")
    m_WebTable.appendChild DOMElement
End With

With m_WebTable
    
    .border = 0
    .Width = "1400"
    .cellPadding = 0
    .cellSpacing = 0
    .Id = "BalanceSheet"
    .Align = "Center"
    
End With

m_WebDoc.body.appendChild m_WebTable

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 2)

With m_WebRow.cells(0)
    .innerText = ""
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
    
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Liability"
        
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .borderRightStyle = "none"
            .FONTSIZE = 13
            .paddingLeft = 5
            .paddingRight = 5
        End With
        
    End With
    
     .vAlign = "Top"
     .Width = "50%"
End With

With m_WebRow.cells(1)

    .innerText = ""
    
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Asset"
        
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .borderLeftStyle = "none"
            .FONTSIZE = 13
            .paddingLeft = 5
            .paddingRight = 5
        End With
        
    End With
     
     .vAlign = "Top"
     .Width = "50%"
     
End With

End Sub
Private Sub RPStartRP()

Dim strHeader As String
Dim HtmlItem
Set m_WebDoc = m_frmWeb.web.document


With m_WebDoc

    Set HtmlItem = .createElement("H3")
   .body.appendChild HtmlItem
    strHeader = GetResourceString(196) & " - " & GetResourceString(197) _
            & " " & GetResourceString(430) '"Receipts and Payments for the Duration"
    With HtmlItem
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FONTSIZE = "17"
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    
    Set HtmlItem = .createElement("H3")
   
    'strHeader = "From : " & FromRpDate & "  To : " & ToRpDate
    strHeader = GetFromDateString(FromRpDate, ToRpDate)
    .body.appendChild HtmlItem
    With HtmlItem
        .innerText = strHeader
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FONTSIZE = "17"
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    
    
    Set m_WebTable = .createElement("Table")
   
    With m_WebTable
        .Id = "RPTable"
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FONTSIZE = 10
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = 4
    End With
    
    Set m_WebTBody = .createElement("TBody")
    
    m_WebTable.appendChild m_WebTBody
    
    .body.appendChild m_WebTable
    
    m_WebTable.Style.fontFamily = gFontName
    m_WebTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

With m_WebRow.cells(0)
    .innerText = GetResourceString(33) '"Sl No"
    .Align = "Left"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(1)
    .innerText = GetResourceString(39) '"Particualrs"
    .Align = "Left"
    .Width = "50%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(2)
    .innerText = GetResourceString(196) '"Receipts"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With
With m_WebRow.cells(3)
    .innerText = GetResourceString(197) '"Payments"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

Private Sub StartReport1(RepType As wisReports)

Dim strHeader As String
Dim HtmlItem
Dim fntSize As Single

Set m_WebDoc = m_frmWeb.web.document

With m_WebDoc
    fntSize = IIf(gLangOffSet, 19, 17)
'    Set HtmlItem = .createElement("H3")
'   .body.appendChild HtmlItem
    If RepType = wisDebitCreditStatement Then
        '"Receipts and Payments for the Duration"
        strHeader = GetResourceString(196) & " - " & GetResourceString(197, 430)
    Else
        'Profit and loss statment
        strHeader = GetResourceString(403, 244, 404)
    End If
'    With HtmlItem
'        .innerText = strHeader
'        .Align = "Center"
'        .Style.fontFamily = gFontName
'        .Style.FontSize = fntSize
'    End With
'
'    Set HtmlItem = .createElement("H3")
'
'    'strHeader = "From : " & FromRpDate & "  To : " & ToRpDate
'    strHeader = GetFromDateString(FromRpDate, ToRpDate)
'    .body.appendChild HtmlItem
'    With HtmlItem
'        .innerText = strHeader
'        .Align = "Center"
'        .Style.fontFamily = gFontName
'        .Style.FontSize = fntSize
'    End With
Call SetReportTitle(strHeader)
    Set m_WebTable = .createElement("Table")
   
    With m_WebTable
        .Id = "RPTable"
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FONTSIZE = 10
        .Width = "650"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = 4
    End With
    
    Set m_WebTBody = .createElement("TBody")
    
    m_WebTable.appendChild m_WebTBody
    
    .body.appendChild m_WebTable
    
End With
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

fntSize = IIf(gLangOffSet, 17, 15)
With m_WebRow.cells(0)
    .innerText = GetResourceString(33) '"Sl No"
    .Align = "Left"
    .Width = "10%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = fntSize
End With
With m_WebRow.cells(1)
    .innerText = GetResourceString(39) '"Particualrs"
    .Align = "Left"
    .Width = "50%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = fntSize
End With
With m_WebRow.cells(2)
    strHeader = GetResourceString(IIf(RepType = wisDebitCreditStatement, 196, 403))
    .innerText = strHeader '"Receipts"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = fntSize
End With
With m_WebRow.cells(3)
    strHeader = GetResourceString(IIf(RepType = wisDebitCreditStatement, 197, 404))
    .innerText = strHeader '"Payments"
    .Align = "Left"
    .Width = "20%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = fntSize
End With

End Sub


Private Sub PLStartProfitAndLoss()

Dim strHeader As String
Dim DOMElement
Dim DOMTbody

' set the web document to htmldoc

Set m_WebDoc = m_frmWeb.web.document

'strHeader = "Profit and Loss for the Duration"
strHeader = GetResourceString(403, 244, 404)
Call SetReportTitle(strHeader)

Set m_WebTable = m_WebDoc.createElement("Table")
With m_WebTable
    .border = 0
    .Width = "900"
    .cellPadding = 0
    .cellSpacing = 0
    .Id = "PLTable"
    .Align = "Center"
    .Style.fontFamily = gFontName
    .Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
End With

m_WebDoc.body.appendChild m_WebTable
    
Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 2)

With m_WebRow.cells(0)
    
    .innerText = ""
    
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Expense"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FONTSIZE = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
        
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

With m_WebRow.cells(1)
    .innerText = ""
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Income"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FONTSIZE = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

End Sub

Private Sub StartReport2(RepType As wisReports)

Dim strHeader As String
Dim DOMElement
Dim DOMTbody

' set the web document to htmldoc
Set m_WebDoc = m_frmWeb.web.document

If RepType = wisProfitLossStatement Then
    'strHeader = "Profit and Loss for the Duration"
    strHeader = GetResourceString(403, 244, 404)
Else
    '"Receipts and Payments for the Duration"
    strHeader = GetResourceString(196) & " - " & GetResourceString(197) & _
                " " & GetResourceString(430)
End If
Call SetReportTitle(strHeader)

With m_WebTable
    .border = 0
    .Width = IIf(m_OtherDetail, "900", "600")
    .cellPadding = 0
    .cellSpacing = 0
    .Id = "PLTable"
    .Align = "Center"
End With

m_WebDoc.body.appendChild m_WebTable

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 2)

With m_WebRow.cells(0)
    
    .innerText = ""
    
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Expense"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FONTSIZE = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
        
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

With m_WebRow.cells(1)
    .innerText = ""
    Set DOMElement = m_frmWeb.web.document.createElement("Table")
    Set DOMTbody = m_frmWeb.web.document.createElement("TBODY")
    
    .appendChild DOMElement
    
    With DOMElement
        .appendChild DOMTbody
        .border = 1
        .Width = "100%"
        .cellSpacing = 0
        .cellPadding = 0
        .Id = "Income"
        .Align = "Center"
        With .Style
            .tableLayout = "fixed"
            .borderCollapse = "Collapse"
            .BorderColor = "Black"
            .FONTSIZE = 13
            .paddingLeft = 3
            .paddingRight = 3
        End With
    End With
        
     .vAlign = "Top"
     .Width = "50%"
End With

End Sub

Private Sub StartHtmlPage()

With m_frmWeb.web
        .document.write "<Html>" & "<Body>"
End With

End Sub

Private Sub StartHtmlTable(ByVal TableID As String, ObjectToAppend, _
            ByVal ColsCount As Integer)

With m_WebDoc

    Set m_WebTable = .createElement("Table")
    
    With m_WebTable
        .Id = TableID
        .border = 1
        .cellPadding = 0
        .cellSpacing = 0
        .Style.BorderColor = "Black"
        .Style.borderCollapse = "Collapse"
        .Align = "Center"
        .Style.FONTSIZE = 10
        .Width = "1050"
        .Style.paddingLeft = 5
        .Style.paddingRight = 5
        .Cols = ColsCount
    End With
    
    ObjectToAppend.appendChild m_WebTable
    
    Set m_WebTBody = .createElement("TBody")
    m_WebTable.appendChild m_WebTBody
    
    m_WebTable.Style.fontFamily = gFontName
    m_WebTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
    
End With

End Sub
Private Sub TRPutOpeningStock(OpeningStock As Currency)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = GetResourceString(284, 175) '"Opening Stock"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = FormatCurrency(OpeningStock)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub
Private Sub TRPutClosingStock(ClosingStock As Currency)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)
m_WebRow.Style.fontWeight = "Bold"
With m_WebRow.cells(2)
    .innerText = GetResourceString(175) & " " & _
            GetResourceString(285) '"Closing Stock"
    .Align = "Left"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = FormatCurrency(ClosingStock)
    .Align = "Right"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
    If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

Private Sub TRPutSalesPurchases(Purchase As Currency, Sales As Currency, _
                                                    ByVal PurchaseHead As String, ByVal SalesHead As String)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

If PurchaseHead <> "" Then
    With m_WebRow.cells(0)
        .innerText = PurchaseHead
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
End If
If Purchase > 0 Then
    With m_WebRow.cells(1)
        .innerText = FormatCurrency(Purchase)
        .Align = "Right"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
End If

If SalesHead <> "" Then
    With m_WebRow.cells(2)
        .innerText = SalesHead
        .Align = "Left"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
End If

If Sales > 0 Then
    With m_WebRow.cells(3)
        .innerText = FormatCurrency(Sales)
        .Align = "Right"
        .Style.fontWeight = "Normal"
        .Style.FONTSIZE = 13
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 15
    End With
End If
End Sub
'
Private Sub TRStartHtmlPage()

' This will put the trading header

Dim eTbody
Dim DOMElement

Set m_WebDoc = m_frmWeb.web.document

If m_WebDoc Is Nothing Then Set m_WebDoc = m_frmWeb.web.document
m_WebDoc.body.Style.fontFamily = gFontName

With m_WebDoc
    
    Set DOMElement = .createElement("H3")
    .body.appendChild DOMElement
    With DOMElement
        '.innerText = "Trading Account From " & FromRpDate & "  To " & ToRpDate
        .innerText = GetResourceString(444) & " " & GetFromDateString(FromRpDate, ToRpDate)
        .Align = "Center"
        .Style.fontFamily = gFontName
        .Style.FONTSIZE = "17"
        If gLangOffSet <> 0 Then .Style.FONTSIZE = 19
    End With
    
    Set m_WebTable = .createElement("Table")
    Set eTbody = .createElement("TBody")
    m_WebTable.Style.fontFamily = gFontName
    m_WebTable.Style.FONTSIZE = IIf(gLangOffSet <> 0, 15, 13)
End With

Call StartHtmlTable("Trading", m_WebDoc.body, 4)

Set m_WebRow = InsertRowAndCellSToTable(m_WebTable, 4)

m_WebRow.Style.fontWeight = "Bold"

With m_WebRow.cells(0)
    .innerText = GetResourceString(39) '"Particulars"
    .Align = "Left"
    .Width = "35%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(1)
    .innerText = GetResourceString(40) '"Amount"
    .Align = "Left"
    .Width = "15%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(2)
    .innerText = GetResourceString(39) '"Particulars"
    .Align = "Left"
    .Width = "35%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

With m_WebRow.cells(3)
    .innerText = GetResourceString(40) '"Amount"
    .Align = "Left"
    .Width = "15%"
    .Style.fontWeight = "Bold"
    .Style.FONTSIZE = 15
     If gLangOffSet <> 0 Then .Style.FONTSIZE = 17
End With

End Sub

' This function is used to test the html functions
Public Sub WebHtml()

Dim rstLedger As ADODB.Recordset
Dim headID As Long
Dim FldCount As Integer
Dim FldLength As Integer

gDbTrans.SqlStmt = " SELECT * FROM AccTrans"

'Fetch the Data to Recordset based on the Condition
Call gDbTrans.Fetch(rstLedger, adOpenForwardOnly)

StartTimer

FldLength = rstLedger.Fields.count - 1

Do While Not rstLedger.EOF
    For FldCount = 0 To FldLength
        headID = FormatField(rstLedger.Fields(0))
    Next FldCount
    rstLedger.MoveNext
Loop

StopTimer

Call rstLedger.Save("c:\windows\desktop\tr", adPersistXML)

MsgBox "Hoho"
End Sub

' This function is used to test the html functions

Public Sub WebHtmlthatworked()

Dim Ele
Dim EleCell
Dim EleTable
Dim EleTbody

With m_frmWeb.web.document

    .Clear
           
    Set Ele = .createElement("Html")
    .body.appendChild Ele
    
    Set Ele = .createElement("Body")
    .body.appendChild Ele
    
    Set Ele = .createElement("H1")
    .body.appendChild Ele
    
    Ele.innerText = "This is test"
    
    Set Ele = .createElement("H1")
    .body.appendChild Ele
    
    Ele.innerText = "This is test 1"
    
    Set EleTable = .createElement("Table")

    EleTable.Id = "PLTable"
    EleTable.border = 1
    EleTable.cellSpacing = 0
    EleTable.cellPadding = 0

    Set EleTbody = .createElement("TBody")
    EleTable.appendChild EleTbody

    Dim I As Integer
    Dim j As Integer

    For I = 0 To 4
        Set Ele = .createElement("TR")
        For j = 0 To 4

            Set EleCell = .createElement("TD")
            EleCell.innerText = "Test" & j
            Ele.appendChild (EleCell)
        Next j
        EleTbody.appendChild (Ele)
    Next I

    .body.appendChild EleTable
    
End With

    
End Sub

Private Sub Class_Initialize()

Set m_frmWeb = New frmWeb
Load m_frmWeb

End Sub

Private Sub Class_Terminate()

On Error Resume Next

Set m_frmWeb = Nothing

End Sub

Private Sub m_frmAccReports_OKClick(StDate As String, EndDate As String, ParentID As Long, headID As Long, ReportSelected As Wis_AccountReportList)

FromRpDate = StDate
ToRpDate = EndDate
RpParentID = ParentID
RpHeadID = headID
m_AccountReportList = ReportSelected

End Sub

Private Sub m_frmRptDt_OKClick(StDate As String, EndDate As String)
    FromRpDate = StDate
    ToRpDate = EndDate
End Sub
Public Property Get FromRpDate() As String
    FromRpDate = m_FromRpDate
End Property
Public Property Let FromRpDate(ByVal NewValue As String)
    m_FromRpDate = ""
    m_FromUSDate = vbNull
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_FromRpDate = NewValue
    m_FromUSDate = GetSysFormatDate(m_FromRpDate)
End Property
Public Property Get ToRpDate() As String
    ToRpDate = m_ToRpDate
End Property
Public Property Get RpHeadID() As Long
    RpHeadID = m_RpHeadId
End Property
Public Property Let RpHeadID(ByVal NewValue As Long)
    m_RpHeadId = NewValue
End Property
Public Property Get RpParentID() As Long
    RpParentID = m_RpParentID
End Property
Public Property Let RpParentID(ByVal NewValue As Long)
    m_RpParentID = NewValue
End Property
Public Property Get FromRpAmount() As Currency
    FromRpAmount = m_FromRpAmount
End Property
Public Property Let FromRpAmount(ByVal NewValue As Currency)
    m_FromRpAmount = 0
    If Not CurrencyValidate(Str(NewValue), True) Then Exit Property
    m_FromRpAmount = NewValue
End Property
Public Property Get ToRpAmount() As Currency
    ToRpAmount = m_ToRpAmount
End Property
Public Property Let ToRpAmount(ByVal NewValue As Currency)
    m_ToRpAmount = 0
    If Not CurrencyValidate(Str(NewValue), True) Then Exit Property
    m_ToRpAmount = NewValue
End Property

Public Property Let ToRpDate(ByVal NewValue As String)
    m_ToRpDate = ""
    m_ToUSDate = vbNull
    If Not DateValidate(NewValue, "/", True) Then Exit Property
    m_ToRpDate = NewValue
    m_ToUSDate = GetSysFormatDate(m_ToRpDate)
End Property

