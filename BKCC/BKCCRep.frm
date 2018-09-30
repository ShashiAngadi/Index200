VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBkccReport 
   Caption         =   "BKCC Loan Reports .."
   ClientHeight    =   6015
   ClientLeft      =   2550
   ClientTop       =   1860
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   6675
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   5130
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   2520
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   1080
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   400
         Left            =   3780
         TabIndex        =   2
         Top             =   210
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4605
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8123
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   "Report Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2340
      TabIndex        =   4
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmBkccReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ParentForm As frmBKCCAcc

Private m_FromIndianDate As String
Private m_FromDate As Date
Private m_ToIndianDate As String
Private m_ToDate As Date
Private m_FromAmt As Currency
Private m_ToAmt As Currency
Private m_Caste As String
Private m_Place As String
Private m_ReportType As wis_BKCCReports
Private m_ReportOrder As wis_ReportOrder
Private m_repType As wis_LoanReports
Private m_FarmerType As wisFarmerClassification
Private m_Gender As wis_Gender

'Declare the object
Private xlWorkBook As Object
Private xlWorkSheet As Object

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_printCount As Long
Private m_TotalCount As Long
Private WithEvents m_frmCancel As frmCancel
Attribute m_frmCancel.VB_VarHelpID = -1
Private m_Cancel As Boolean
Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property

Private Function HasOverDueByYearEnd(LoanID As Long, ByVal TransDate As Date) As Boolean

Dim LastIssueDate  As Date
Dim transType As wisTransactionTypes
Dim rstPay As Recordset
Dim PrevLoanBalance As Currency

HasOverDueByYearEnd = False

LastIssueDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))

 'If he has not paid any amount, then consider whole amount till the End of Year
       'He has the Previous Loan so Check whether He will over due by end of year.
       gDbTrans.SqlStmt = "Select Amount,Balance,TransType,TransDate from BKCCTrans where LoanId = " & LoanID & _
           " And Deposit = False And TransDate > #" & DateAdd("yyyy", -1, LastIssueDate) & "#  And TransDate <= #" & LastIssueDate & "#" & _
           " order by TransDate, TransID"
       If gDbTrans.Fetch(rstPay, adOpenDynamic) < 1 Or rstPay.recordCount < 1 Then Exit Function
       transType = FormatField(rstPay("TransType"))
       
       ''Check All the Transactions till begin of year
       'Take the Previos balance from First Transcation
       PrevLoanBalance = FormatField(rstPay("Balance")) + _
           (FormatField(rstPay("Amount")) * IIf(transType = wDeposit Or transType = wContraDeposit, 1, -1))

    If PrevLoanBalance > 0 Then
       While Not rstPay.EOF
           transType = FormatField(rstPay("TransType"))
           PrevLoanBalance = PrevLoanBalance + _
               (FormatField(rstPay("Amount")) * IIf(transType = wDeposit Or transType = wContraDeposit, 1, -1))
           
           If PrevLoanBalance <= 0 Then rstPay.MoveLast
           rstPay.MoveNext
       Wend
    End If
       HasOverDueByYearEnd = PrevLoanBalance > 0
            
End Function

Public Property Let FarmerType(NewValue As wisFarmerClassification)
    m_FarmerType = NewValue
End Property

Public Property Let Gender(NewValue As wis_Gender)
    m_Gender = NewValue
End Property


Private Sub GetLoanDetailsForCliamBill(ByVal LoanID As Long, ByVal TransDate As Date, ByRef RepaidAmount As Currency, ByRef payDate() As Date, ByRef PayAmount() As Currency)
    Dim LastIssueDate As Date
    ReDim Preserve payDate(0)
    ReDim Preserve PayAmount(0)
    payDate(0) = TransDate
    PayAmount(0) = 0
    Dim rstPay As Recordset
    Dim PrevLoanBalance As Currency
    Dim lastTransType As wisTransactionTypes
    Dim transType As wisTransactionTypes, ContraTransType As wisTransactionTypes
    
    
    LastIssueDate = CDate("31/3/" & Year(TransDate) - IIf(Month(TransDate) < 4, 1, 0))
    transType = wWithdraw
    ContraTransType = wContraWithdraw
    
    gDbTrans.SqlStmt = "Select Amount,Balance,TransType,TransDate from BkCCTrans where LoanId = " & LoanID & _
        " And Deposit = False And TransDate < #" & TransDate & "#  And TransDate > #" & DateAdd("yyyy", -1, TransDate) & "#" & _
        " Union " & _
        " Select 0 as Amount,Balance, 100 as TransType,TransDate from BKCCTrans where loanid = " & LoanID & _
        " And TransID = (Select max(transId) from bkcctrans where loanid = " & LoanID & _
                " And transDate <= #" & DateAdd("yyyy", -1, TransDate) & "# )" & _
        " order by TransDate"
    
    If gDbTrans.Fetch(rstPay, adOpenDynamic) < 1 Then Exit Sub
    If rstPay.recordCount < 1 Then Exit Sub
    'Do while
    lastTransType = FormatField(rstPay("TransType"))
    Do
        If rstPay("TransType") = 100 Then
            PrevLoanBalance = FormatField(rstPay("Balance"))
        ElseIf rstPay("transType") = wDeposit Or rstPay("transType") = wContraDeposit Then
            'If PrevLoanBalance > 0 Then RepaidAmount = RepaidAmount - PrevLoanBalance
            'If PrevLoanBalance > 0 And PrevLoanBalance - FormatField(rstPay("Balance")) <= 0 Then
            If FormatField(rstPay("Balance")) <= 0 Then
                ReDim payDate(0)
                ReDim PayAmount(0)
                payDate(0) = TransDate
                PayAmount(0) = 0
            End If
            PrevLoanBalance = PrevLoanBalance - FormatField(rstPay("Amount"))
        Else
           If rstPay("transDate") <= LastIssueDate Then
            'PrevLoanBalance = PrevLoanBalance + FormatField(rstPay("Balance"))
            PayAmount(UBound(payDate)) = FormatField(rstPay("Amount"))
            payDate(UBound(payDate)) = rstPay("TransDate")
            ReDim Preserve payDate(UBound(payDate) + 1)
            ReDim Preserve PayAmount(UBound(PayAmount) + 1)
            If FormatField(rstPay("Balance")) <= 0 Then rstPay.MoveLast
            End If
        End If
        
        rstPay.MoveNext
        If rstPay.EOF Then Exit Do
    Loop
    
    If PrevLoanBalance >= RepaidAmount Then
        'That mean he has a balance which is due more than a Year
        ReDim payDate(0)
        ReDim PayAmount(0)
        payDate(0) = TransDate
        PayAmount(0) = 0
    ElseIf PrevLoanBalance > 0 Then
        RepaidAmount = RepaidAmount - PrevLoanBalance
    End If
    
    If UBound(payDate) > 0 Then
        ReDim Preserve payDate(UBound(payDate) - 1)
        ReDim Preserve PayAmount(UBound(PayAmount) - 1)
    End If

End Sub


Private Sub GetRepayDetailsForCliamBill(ByVal LoanID As Long, ByVal TransDate As Date, ByVal Balance As Currency, ByRef loanIssueAmount As Currency, ByRef payDate() As Date, ByRef repayAmount() As Currency, VoucherNo() As String)
    Dim LastIssueDate As Date
    Dim loanIssueBalance As Currency
    ReDim payDate(0)
    ReDim repayAmount(0)
    ReDim VoucherNo(0)
    Dim transAmount As Currency
    
    payDate(0) = "1/1/2000"
    repayAmount(0) = 0
    VoucherNo(0) = "NORECORDS"
    Dim rstPay As Recordset
    Dim PrevLoanBalance As Currency
    Dim transType As wisTransactionTypes, ContraTransType As wisTransactionTypes
    
    LastIssueDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))
    loanIssueBalance = loanIssueAmount
    PrevLoanBalance = Balance - loanIssueAmount
    
    If PrevLoanBalance > 0 Then
        'If HasOverDueByYearEnd(LoanID, TransDate) Then repayAmount(0) = 0: Exit Sub
    End If
    
    transType = wDeposit
    ContraTransType = wContraDeposit
    
    gDbTrans.SqlStmt = "Select Amount,Balance,TransType,TransDate,voucherNo from BKCCTrans where LoanId = " & LoanID & _
        " And Deposit = False And TransDate >= #" & TransDate & "#  And TransDate <= #" & LastIssueDate & "#" & _
        " order by TransDate,TransID"
    
    
    If gDbTrans.Fetch(rstPay, adOpenDynamic) < 1 Or rstPay.recordCount < 1 Then
        'If he has not paid any amount, then consider whole amount till the End of Year
        payDate(0) = "1/1/2000"
        repayAmount(0) = loanIssueAmount
        Exit Sub
    End If
    
    
    'Do while
    Do
        transType = FormatField(rstPay("TransType"))
        If rstPay("TransType") = wDeposit Or rstPay("TransType") = wContraDeposit Then
            transAmount = FormatField(rstPay("Amount"))
            If PrevLoanBalance > 0 Then
                transAmount = 0
                PrevLoanBalance = PrevLoanBalance - FormatField(rstPay("Amount"))
                If PrevLoanBalance < 0 Then transAmount = Abs(PrevLoanBalance)
            End If
            If transAmount > 0 Then
                loanIssueBalance = loanIssueBalance - transAmount
                repayAmount(UBound(payDate)) = transAmount
                payDate(UBound(payDate)) = rstPay("TransDate")
                VoucherNo(UBound(VoucherNo)) = FormatField(rstPay("VoucherNo"))
                ReDim Preserve payDate(UBound(payDate) + 1)
                ReDim Preserve repayAmount(UBound(repayAmount) + 1)
                ReDim Preserve VoucherNo(UBound(VoucherNo) + 1)
            End If
        End If
        
        rstPay.MoveNext
        If rstPay.EOF Then Exit Do
    Loop
    
   
    If loanIssueBalance > 0 Then
        payDate(UBound(payDate)) = "1/1/2000" 'LastIssueDate
        repayAmount(UBound(payDate)) = loanIssueBalance
        VoucherNo(UBound(VoucherNo)) = "NORECORDS"
    ElseIf UBound(payDate) > 0 Then
        ReDim Preserve payDate(UBound(payDate) - 1)
        ReDim Preserve repayAmount(UBound(repayAmount) - 1)
        ReDim Preserve VoucherNo(UBound(VoucherNo) - 1)
    End If

    
End Sub
Private Sub GetRepayDetailsForPrevCliamBill(ByVal LoanID As Long, ByVal TransDate As Date, ByVal Balance As Currency, ByRef loanIssueAmount As Currency, ByRef payDate() As Date, ByRef repayAmount() As Currency, VoucherNo() As String)
    Dim LastIssueDate As Date
    Dim YearEndDate As Date
    Dim loanIssueBalance As Currency
    ReDim payDate(0)
    ReDim repayAmount(0)
    ReDim VoucherNo(0)
    Dim transAmount As Currency
    
    payDate(0) = "1/1/2000"
    repayAmount(0) = 0
    VoucherNo(0) = "NORECORDS"
    Dim rstPay As Recordset
    Dim PrevLoanBalance As Currency
    Dim transType As wisTransactionTypes, ContraTransType As wisTransactionTypes
    
    YearEndDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))
    LastIssueDate = DateAdd("yyyy", 1, TransDate)
    loanIssueBalance = loanIssueAmount
    PrevLoanBalance = Balance - loanIssueAmount
    
    If PrevLoanBalance > 0 Then
        'If HasOverDueByYearEnd(LoanID, TransDate) Then repayAmount(0) = 0: Exit Sub
    End If
    
    transType = wDeposit
    ContraTransType = wContraDeposit
    
    gDbTrans.SqlStmt = "Select Amount,Balance,TransType,TransDate,voucherNo from BKCCTrans where LoanId = " & LoanID & _
        " And Deposit = False And TransDate >= #" & TransDate & "#  And TransDate <= #" & LastIssueDate & "#" & _
        " order by TransDate,TransID"
    
    
    If gDbTrans.Fetch(rstPay, adOpenDynamic) < 1 Or rstPay.recordCount < 1 Then
        'If he has not paid any amount, then consider whole amount till the End of Year
        payDate(0) = "1/1/2000"
        repayAmount(0) = loanIssueAmount
        Exit Sub
    End If
    
    
    'Do while
    Do
        transType = FormatField(rstPay("TransType"))
        If rstPay("TransType") = wDeposit Or rstPay("TransType") = wContraDeposit Then
            transAmount = FormatField(rstPay("Amount"))
            If PrevLoanBalance > 0 Then
                transAmount = 0
                PrevLoanBalance = PrevLoanBalance - FormatField(rstPay("Amount"))
                If PrevLoanBalance < 0 Then transAmount = Abs(PrevLoanBalance)
            End If
            If transAmount > 0 Then
                loanIssueBalance = loanIssueBalance - transAmount
                
                If rstPay("TransDate") > YearEndDate Then
                repayAmount(UBound(payDate)) = transAmount
                payDate(UBound(payDate)) = rstPay("TransDate")
                VoucherNo(UBound(VoucherNo)) = FormatField(rstPay("VoucherNo"))
                ReDim Preserve payDate(UBound(payDate) + 1)
                ReDim Preserve repayAmount(UBound(repayAmount) + 1)
                ReDim Preserve VoucherNo(UBound(VoucherNo) + 1)
                End If
                If loanIssueBalance <= 0 Then Exit Do
            End If
        End If
        
        rstPay.MoveNext
        If rstPay.EOF Then Exit Do
    Loop
    
   
    If loanIssueBalance > 0 Then
        payDate(UBound(payDate)) = "1/1/2000" 'LastIssueDate
        repayAmount(UBound(payDate)) = 0 'loanIssueBalance
        VoucherNo(UBound(VoucherNo)) = "NORECORDS"
    ElseIf UBound(payDate) > 0 Then
        ReDim Preserve payDate(UBound(payDate) - 1)
        ReDim Preserve repayAmount(UBound(repayAmount) - 1)
        ReDim Preserve VoucherNo(UBound(VoucherNo) - 1)
    End If

    
End Sub

Private Sub InitMonthRegGrid()
    
    Dim ColCount As Long
    Dim Wid As Single
    For ColCount = 0 To grd.Cols - 1
        Wid = GetSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth" & ColCount, grd.Width / grd.Cols) * grd.Width
        If Wid >= grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
        If Wid < 20 And Wid <> 0 Then Wid = grd.Width / grd.Cols * 2
        grd.ColWidth(ColCount) = Wid
    Next ColCount

End Sub

Private Sub ReportLoanMonthlyBalance()
Dim SqlStr As String
Dim rstMain As ADODB.Recordset

Dim fromDate As Date
Dim toDate As Date

'Get the Firs Day Of the Financial Year
fromDate = DateAdd("d", -1, FinUSFromDate)
'Last Day Of the ToDate
toDate = GetSysLastDate(m_ToDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463, 67, 42) & _
    " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

SqlStr = "select AccNum,LoanId, Name as CustName,MemberNum " & _
    " FROM BkccMaster A INNER JOIN QryMemName B ON A.MemID = B.MemID " & _
    " WHERE (ClosedDate IS NULL Or ClosedDate < #" & toDate & "#)"

If m_ReportType = repBkccDepMonBal Then
    SqlStr = SqlStr & " And LoanID In (Select Distinct LoanId From BkccTrans " & _
        " Where TransDate Between #" & fromDate & "# AND #" & toDate & "# And Balance < 0) "
End If
'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType

SqlStr = SqlStr & " ORDER BY " & IIf(m_ReportOrder = wisByName, "IsciName", "Val(AccNum)")

'First Fetch the Master Records
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Sub

Dim SlNo As Long
Dim count As Long
Dim MaxCount As Long
Dim rowno As Integer, colno As Integer

count = DateDiff("m", fromDate, toDate)
count = count + 2
MaxCount = rstMain.recordCount * count

RaiseEvent Initialise(0, MaxCount)

Call InitGrid
count = 0
rowno = grd.Row

While Not rstMain.EOF
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        colno = 0: .TextMatrix(rowno, colno) = .Row
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rstMain("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rstMain("MemberNum"))
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rstMain("CustName"))
        colno = 4
    End With
    count = count + 1
    RaiseEvent Processing("Collecting Information ", count / MaxCount)
    rstMain.MoveNext
Wend
With grd
   ' .Cols = .Cols + 1
    .Row = rowno
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
    
End With

Dim rstBalance As Recordset
Dim AccNum As String
Dim TotalBalance As Currency

fromDate = "4/30/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))


While DateDiff("d", fromDate, toDate) > 0
  With grd
    .Cols = .Cols + 1
    .Col = .Cols - 1
    .Row = 0: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID,LoanID From BKCcTrans" & _
            " WHERE TransDate <= #" & fromDate & "# GROUP BY LoanID"
    gDbTrans.CreateView ("qryMaxID")
    
    SqlStr = "Select A.LoanID,AccNum, Balance From (BKCCMaster A " & _
        " INNER JOIN BkccTrans B ON A.LoanID=B.LoanID) INNER JOIN " & _
        " qryMaxID C ON B.TransID = C.MaxTransID AND B.LoanID = C.LoanID" & _
        " WHERE Balance " & IIf(m_ReportType = repBkccDepMonBal, "< 0", "> 0")
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rstBalance, adOpenDynamic) < 1 Then GoTo NextMonth
    SqlStr = ""
    
    rstMain.MoveFirst: rowno = 0
    Do
        If rowno = .Rows - 1 Or rstMain.EOF Then Exit Do
        rowno = rowno + 1
        rstBalance.MoveFirst
        rstBalance.Find "LoanId = " & rstMain("LoanID")
        If Not rstBalance.EOF Then
            .TextMatrix(rowno, colno) = FormatCurrency(Abs(rstBalance("Balance")))
            TotalBalance = TotalBalance + Val(.TextMatrix(rowno, colno))
        End If
        count = count + 1
        RaiseEvent Processing("Calulating monthly Balance", count / MaxCount)
        rstMain.MoveNext
    Loop
    
    rowno = rowno + 2
    .Row = rowno: .Col = colno
    .TextMatrix(rowno, colno) = FormatCurrency(TotalBalance)
    .CellFontBold = True: TotalBalance = 0
    colno = colno + 1
NextMonth:
    fromDate = DateAdd("d", 1, fromDate)
    fromDate = DateAdd("d", -1, DateAdd("m", 1, fromDate))
  End With
Wend

End Sub

Private Sub ReportLoanMonthlyTrans()
Dim SqlStr As String
Dim rstMain As Recordset
Dim rstTrans As Recordset
Dim fromDate As Date
Dim toDate As Date
Dim NoCount As Integer
Dim Deposit As Boolean

'Last Day Prevoius Finanacil Year
fromDate = DateAdd("d", -1, FinUSFromDate)

SqlStr = "Select AccNum,A.LoanId,SanctionAmount,CurrentSanction,ExtraSanction, " & _
    " Name as CustName FROM BkccMaster A, QryName C " & _
    " WHERE C.CustomerId = A.CustomerID " & _
    " AND (ClosedDate IS NULL Or ClosedDate < #" & m_ToDate & "#) "

'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType

SqlStr = SqlStr & " ORDER BY " & IIf(m_ReportOrder = wisByName, "IsciName", "Val(AccNum)")

RaiseEvent Initialise(0, 10)

'First Fetch the MAster Records
gDbTrans.SqlStmt = SqlStr
'DISPALY THE PROCESS BAR BEACUSE FETCHING MAY TAKE LOT OF TIME
RaiseEvent Processing("Fetching the master account", 0.25)
If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Sub

'Then Fetch the Balance as on 31/3
fromDate = DateAdd("d", -1, FinUSFromDate)

gDbTrans.SqlStmt = "SELECT Max(TransID)as MaxTransID,LoanID From BKCcTrans C " & _
            " Where TransDate <= #" & fromDate & "# GROUP BY LOanID"
gDbTrans.CreateView ("qryMaxID")

SqlStr = "Select A.LoanID,AccNum, Balance " & _
        " From (BKCCMAster A INNER JOIN BkccTrans B ON A.LoanID = B.LoanID)" & _
        " INNER JOIN qryMaxID C On B.LoanID=C.LoanID AND B.TransID = C.MAxTransID" & _
        " WHERE Balance > 0 "


gDbTrans.SqlStmt = SqlStr
RaiseEvent Processing("Fetching the master account", 0.55)
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then GoTo NextMonth
RaiseEvent Processing("Fetching the master account", 0.95)

Dim SlNo As Long
Dim count As Long
Dim MaxCount As Long
Dim rowno As Integer, colno As Integer

count = DateDiff("m", fromDate, m_ToDate)
count = count + 2
MaxCount = rstMain.recordCount * count

Call InitGrid
Dim TotAmount() As Currency
ReDim TotAmount(3 To grd.Cols - 1)

grd.Row = grd.FixedRows - 1
rowno = grd.Row
count = 0


While Not rstMain.EOF
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        colno = 0: .TextMatrix(rowno, colno) = rowno
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rstMain("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rstMain("CustName"))
        
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rstMain("SanctionAmount"))
        TotAmount(colno) = TotAmount(colno) + Val(.TextMatrix(rowno, colno))
        colno = 4: .TextMatrix(rowno, colno) = FormatField(rstMain("CurrentSanction"))
        TotAmount(colno) = TotAmount(colno) + Val(.TextMatrix(rowno, colno))
        colno = 5: .TextMatrix(rowno, colno) = FormatField(rstMain("ExtraSanction"))
        TotAmount(colno) = TotAmount(colno) + Val(.TextMatrix(rowno, colno))
        colno = 6: .TextMatrix(rowno, colno) = Val(FormatField(rstMain("CurrentSanction"))) + Val(FormatField(rstMain("ExtraSanction")))
        TotAmount(colno) = TotAmount(colno) + Val(.TextMatrix(rowno, colno))
        
        If Not rstTrans Is Nothing Then
            rstTrans.MoveFirst
            rstTrans.Find "LoanId = " & rstMain("loanID")
            If Not rstTrans.EOF Then .Col = 7: .Text = FormatField(rstTrans("Balance"))
        End If
    End With
    count = count + 1
    RaiseEvent Processing("Collecting Information ", count / MaxCount)
    rstMain.MoveNext
Wend


Dim AccNum As String
Dim TotalBalance As Currency
Dim ColCount As Integer

Dim Balance As Currency
Dim DepAmount As Currency
Dim WithDraw As Currency
Dim Interest As Currency
Dim transType As wisTransactionTypes

Dim rstTemp As Recordset

'Start Of the Financial Year
fromDate = FinUSFromDate

'Last Day Of the Todate
toDate = GetSysLastDate(m_ToDate)

grd.Rows = rstMain.recordCount + grd.FixedRows + 2

While DateDiff("d", fromDate, toDate) > 0
  With grd
    
    ColCount = .Cols - 1
    .Cols = .Cols + 4
    ReDim Preserve TotAmount(3 To .Cols - 1)
    .Col = .Cols - 1
    .Row = 0
    .Col = ColCount + 1: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 4: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Row = 1
    .Col = ColCount + 1: .Text = GetResourceString(272)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetResourceString(271)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetResourceString(271)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 4: .Text = GetResourceString(67, 58)
    .CellAlignment = 4: .CellFontBold = True
    .Row = 2
    .Col = ColCount + 1: .Text = GetResourceString(272)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetResourceString(271)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetResourceString(47)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 4: .Text = GetResourceString(67, 58)
    .CellAlignment = 4: .CellFontBold = True
    .Row = 3
    For NoCount = ColCount + 1 To .Cols - 1
        .Col = NoCount: .Text = NoCount + 1
        .CellAlignment = 4: .CellFontBold = True
    Next
    
    SqlStr = "Select 'PRINCIPAL',LoanID,TransType,Sum(Amount) As TotalAmount" & _
        " From BkccTrans WHERE TransDate >= #" & fromDate & "# " & _
        " AND TransDate < #" & DateAdd("m", 1, fromDate) & "# " & _
        " And Deposit = " & Deposit & " Group By LoanID,TransType"
    SqlStr = SqlStr & " UNION " & _
        "Select 'INTEREST',LoanID,TransType,Sum(IntAmount+PenalIntAmount) As TotalAmount" & _
        " From BkccIntTrans WHERE TransDate >= #" & fromDate & "# " & _
        " AND TransDate < #" & DateAdd("m", 1, fromDate) & "# " & _
        " ANd Deposit = " & Deposit & " Group By LoanID,TransType"
    
    gDbTrans.SqlStmt = SqlStr
    Set rstTrans = Nothing
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then GoTo NextMonth
    
    rstMain.MoveFirst: .Row = .FixedRows - 1
    rowno = .Row
    Do
        If rowno = .Rows - 1 Or rstMain.EOF Then Exit Do
        rowno = rowno + 1
        DepAmount = 0: WithDraw = 0: Interest = 0
        colno = ColCount
        Balance = Val(.TextMatrix(rowno, colno))
        
    'rstTrans.FindFirst "LoanId = " & rstMain("loanID")
    'CODE Begin
        Set rstTemp = Nothing
        Set rstTemp = rstTrans
        rstTemp.Filter = "LoanId = " & rstMain("loanID")
    
    'Code END
        'rstTrans.Requery
        Do
            If rstTemp.EOF Then Exit Do
            transType = rstTemp("TransType")
            If rstTemp(0) = "INTEREST" Then
                If transType = wDeposit Or transType = wContraDeposit Then Interest = Interest + FormatField(rstTemp("TotalAmount"))
            Else
                If transType = wDeposit Or transType = wContraDeposit Then DepAmount = DepAmount + FormatField(rstTemp("TotalAmount"))
                If transType = wWithdraw Or transType = wContraWithdraw Then WithDraw = WithDraw + FormatField(rstTemp("TotalAmount"))
            End If
            rstTemp.MoveNext
        Loop
        'End If
        Balance = Balance - DepAmount + WithDraw
        colno = ColCount + 1: .TextMatrix(rowno, colno) = WithDraw
        TotAmount(colno) = TotAmount(colno) + WithDraw
        colno = ColCount + 2: .TextMatrix(rowno, colno) = DepAmount
        TotAmount(colno) = TotAmount(colno) + DepAmount
        colno = ColCount + 3: .TextMatrix(rowno, colno) = Interest
        TotAmount(colno) = TotAmount(colno) + Interest
        colno = ColCount + 4: .TextMatrix(rowno, colno) = Balance
        TotAmount(colno) = TotAmount(colno) + Balance
        
        count = count + 1
        RaiseEvent Processing("Calulating monthly Transactions", count / MaxCount)
        rstMain.MoveNext
    Loop
    rowno = rowno + 2
    
    .Row = rowno: .Col = colno
    
    .Text = FormatCurrency(TotalBalance)
    .CellFontBold = True: TotalBalance = 0
NextMonth:
    fromDate = DateAdd("d", 1, fromDate)
    fromDate = DateAdd("d", -1, DateAdd("m", 1, fromDate))
  End With
Wend


With grd
    .Rows = rstMain.recordCount + .FixedRows + 2
    .Row = .Rows - 1
    .Col = 2: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
    For ColCount = 3 To .Cols - 1
        .Col = ColCount: .Text = FormatCurrency(TotAmount(ColCount))
        .CellFontBold = True: .CellAlignment = 4
    Next
End With

End Sub

Private Sub ReportDepositMonthlyTrans()
Dim SqlStr As String
Dim rstMain As Recordset
Dim rstTrans As Recordset
Dim fromDate As Date
Dim toDate As Date
Dim NoCount As Integer
Dim Deposit As Boolean

Deposit = True
'Last Day Prevoius Finanacil Year
fromDate = DateAdd("d", -1, FinUSFromDate)
      
SqlStr = "Select AccNum,A.LoanId, Name as CustName" & _
    " FROM BkccMaster A Inner Join QryName B " & _
    " On A.CustomerId = B.CustomerID " & _
    " WHERE (ClosedDate IS NULL Or ClosedDate < #" & m_ToDate & "#) " & _
    " And LoanId IN (Select Distinct LoanID From BkccTrans " & _
        " Where Deposit = " & Deposit & " And TransDate > #" & fromDate & "#)"

'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType

SqlStr = SqlStr & " ORDER BY " & IIf(m_ReportOrder = wisByName, "IsciName", "Val(AccNum)")

RaiseEvent Initialise(0, 10)

'First Fetch the MAster Records
gDbTrans.SqlStmt = SqlStr
'DISPALY THE PROCESS BAR BEACUSE FETCHING MAY TAKE LOT OF TIME
RaiseEvent Processing("Fetching the master account", 0.25)
If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Sub

'Then Fetch the Balance as on 31/3
fromDate = DateAdd("d", -1, FinUSFromDate)
gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID,LoanID" & _
            " From BKCcTrans C Where TransDate <= #" & fromDate & "#" & _
            " GROUP BY LoanID"
gDbTrans.CreateView ("qryMaxID")

SqlStr = "Select A.LoanID,AccNum,Balance " & _
        " FROM (BKCCMaster A INNER JOIN BKCCTrans B" & _
        " ON A.LoanID = B.LoanID) INNER JOIN qryMaxID C" & _
        " On B.LoanID = C.LoanID AND B.TransID = C.MaxTransID" & _
        " WHERE Balance <0"

gDbTrans.SqlStmt = SqlStr
RaiseEvent Processing("Fetching the master account", 0.55)
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then GoTo NextMonth
RaiseEvent Processing("Fetching the master account", 0.95)

Dim SlNo As Long
Dim count As Long
Dim MaxCount As Long

count = DateDiff("m", fromDate, m_ToDate)
count = count + 2
MaxCount = rstMain.recordCount * count

Call InitGrid
Dim TotAmount() As Currency
Dim rowno As Integer, colno As Integer

grd.Row = grd.FixedRows - 1
rowno = grd.Row: colno = 0
count = 0

While Not rstMain.EOF
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        colno = 0: .TextMatrix(rowno, colno) = rowno
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rstMain("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rstMain("CustName"))
        If Not rstTrans Is Nothing Then
            rstTrans.MoveFirst
            rstTrans.Find "LoanId = " & rstMain("LoanID")
            If Not rstTrans.EOF Then colno = 3: .TextMatrix(rowno, colno) = Abs(FormatField(rstTrans("Balance")))
        End If
    End With
    count = count + 1
    RaiseEvent Processing("Collecting Information ", count / MaxCount)
    rstMain.MoveNext
Wend

Dim AccNum As String
Dim TotalBalance As Currency
Dim ColCount As Integer

Dim Balance As Currency
Dim DepAmount As Currency
Dim WithDraw As Currency
Dim Interest As Currency
Dim transType As wisTransactionTypes

Dim rstTemp As Recordset

'Start Of the Financial Year
fromDate = FinUSFromDate

'Last Day Of the Todate
toDate = GetSysLastDate(m_ToDate)

grd.Rows = rstMain.recordCount + grd.FixedRows + 2

While DateDiff("d", fromDate, toDate) > 0
  With grd
    
    ColCount = .Cols - 1
    .Cols = .Cols + 3
    ReDim Preserve TotAmount(3 To .Cols - 1)
    .Col = .Cols - 1
    .Row = 0
    .Col = ColCount + 1: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    
    .Row = 1
    .Col = ColCount + 1: .Text = GetResourceString(271)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetResourceString(272)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetResourceString(67, 43)
    .CellAlignment = 4: .CellFontBold = True
    
    .Row = 2
    .Col = ColCount + 1: .Text = GetResourceString(271)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 2: .Text = GetResourceString(272)
    .CellAlignment = 4: .CellFontBold = True
    .Col = ColCount + 3: .Text = GetResourceString(67, 43)
    .CellAlignment = 4: .CellFontBold = True
    .Row = 3
    For NoCount = ColCount + 1 To .Cols - 1
        .Col = NoCount: .Text = NoCount + 1
        .CellAlignment = 4: .CellFontBold = True
    Next
    
    SqlStr = "Select 'PRINCIPAL',LoanID,TransType,Sum(Amount) As TotalAmount" & _
        " From BkccTrans WHERE TransDate >= #" & fromDate & "# " & _
        " AND TransDate < #" & DateAdd("m", 1, fromDate) & "# " & _
        " And Deposit = " & Deposit & " Group By LoanID,TransType"
    
    gDbTrans.SqlStmt = SqlStr
    Set rstTrans = Nothing
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then GoTo NextMonth
    
    rstMain.MoveFirst: .Row = .FixedRows - 1
    rowno = .Row
    Do
        If rowno = .Rows - 1 Or rstMain.EOF Then Exit Do
        rowno = rowno + 1
        DepAmount = 0: WithDraw = 0: Interest = 0
        colno = ColCount
        Balance = Val(.TextMatrix(rowno, colno))
        
    'rstTrans.FindFirst "LoanId = " & rstMain("loanID")
    'CODE Begin
        Set rstTemp = Nothing
        Set rstTemp = rstTrans
        rstTemp.Filter = "LoanId = " & rstMain("LoanID")
    
    'Code END
        'rstTrans.Requery
        Do
            If rstTemp.EOF Then Exit Do
            transType = rstTemp("TransType")
            If rstTemp(0) = "INTEREST" Then
                If transType = wDeposit Or transType = wContraDeposit Then Interest = Interest + FormatField(rstTemp("TotalAmount"))
            Else
                If transType = wDeposit Or transType = wContraDeposit Then DepAmount = DepAmount + FormatField(rstTemp("TotalAmount"))
                If transType = wWithdraw Or transType = wContraWithdraw Then WithDraw = WithDraw + FormatField(rstTemp("TotalAmount"))
            End If
            rstTemp.MoveNext
        Loop
        Balance = Balance - DepAmount + WithDraw
        colno = ColCount + 1: .Text = WithDraw
        TotAmount(colno) = TotAmount(colno) + WithDraw
        colno = ColCount + 2: .TextMatrix(rowno, colno) = DepAmount
        TotAmount(colno) = TotAmount(colno) + DepAmount
        colno = ColCount + 3: .TextMatrix(rowno, colno) = Abs(Balance)
        TotAmount(colno) = TotAmount(colno) + Balance
        
        count = count + 1
        RaiseEvent Processing("Calulating monthly Transactions", count / MaxCount)
        rstMain.MoveNext
    Loop
    rowno = rowno + 2
    
    .Row = rowno
    .Text = FormatCurrency(TotalBalance)
    .CellFontBold = True: TotalBalance = 0
NextMonth:
    fromDate = DateAdd("d", 1, fromDate)
    fromDate = DateAdd("d", -1, DateAdd("m", 1, fromDate))
  End With
Wend


With grd
    .Rows = rstMain.recordCount + .FixedRows + 2
    .Row = .Rows - 1
    .Col = 2: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
    For ColCount = 3 To .Cols - 1
        .Col = ColCount: .Text = FormatCurrency(TotAmount(ColCount))
        .CellFontBold = True: .CellAlignment = 4
    Next
End With

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463, 67, 42) & _
    " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

End Sub

Private Function ReportReceivables()
'Now Get the Head ID Of the Bkcc
Dim AccHeadID As Long
Dim rstMain As Recordset

AccHeadID = GetHeadID(GetResourceString(229) & " " & _
                        GetResourceString(58), parMemberLoan)
If AccHeadID = 0 Then gCancel = 2: Exit Function

'NOw fetch the details from From
gDbTrans.SqlStmt = "Select A.*,B.HeadName From AmountReceivable A,Heads B " & _
            " Where AccHeadID = " & AccHeadID & _
            " And B.HeadID = A.DueHeadID" & _
            " And (TransType = " & wWithdraw & _
                " OR TransType = " & wContraWithdraw & ")" & _
            " AND AccTransID >= (Select Max(TransID) " & _
                " From BKCCTrans C Where C.LoanID=A.AccID " & _
                " And C.TransDate <= A.TransDate )"

If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Function
Dim rstName As Recordset
        
gDbTrans.SqlStmt = "Select AccNum,LoanId,A.CustomerId,Name as CustName" & _
        " From BKCCMaster A Inner Join QryName B " & _
        " ON A.CustomerId = B.customerId " & _
        " WHERE LoanID in (Select Distinct AccID as LoanID " & _
            " From AmountReceivAble Where AccHeadID = " & AccHeadID & ")"
            
Call gDbTrans.Fetch(rstName, adOpenDynamic)

With grd
    .Clear
    .Rows = 10
    .Cols = 5
    .FixedCols = 1
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36, 60) '"AccNum"
    .Col = 2: .Text = GetResourceString(35)  '"Name"
    .Col = 3: .Text = GetResourceString(36, 35) '"HeadNAme"
    .Col = 4: .Text = GetResourceString(40)  '"Amount"
End With

Dim rowno As Integer, colno As Integer
rowno = grd.Row

While Not rstMain.EOF
    With grd
        rowno = rowno + 1
        colno = 0: .TextMatrix(rowno, colno) = rowno
        rstName.MoveFirst
        rstName.Find "LoanID = " & rstMain("AccID")
        If Not rstName.EOF Then
            colno = 1: .TextMatrix(rowno, colno) = FormatField(rstName("AccNum"))
            colno = 2: .TextMatrix(rowno, colno) = FormatField(rstName("CustNAme"))
        End If
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rstMain("HeadNAme"))
        colno = 4: .TextMatrix(rowno, colno) = FormatField(rstMain("Amount"))
    End With
    rstMain.MoveNext
Wend
End Function

Public Property Let ToAmount(curTo As Currency)
    m_ToAmt = curTo
End Property

Public Property Let FromAmount(curFrom As Currency)
    m_FromAmt = curFrom
End Property

Public Property Let ToIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_ToIndianDate = NewDate
        m_ToDate = GetSysFormatDate(NewDate)
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_FromIndianDate = NewDate
        m_FromDate = GetSysFormatDate(NewDate)
    Else
        m_FromIndianDate = ""
        m_FromDate = vbNull
    End If
End Property

Public Property Let Place(NewPlace As String)
    m_Place = NewPlace
End Property

Public Property Let ReportOrder(RepOrder As wis_ReportOrder)
    m_ReportOrder = RepOrder
End Property

Public Property Let ReportType(RepType As wis_BKCCReports)
    m_ReportType = RepType
End Property

Public Sub InitGrid()
    Dim ColCount As Integer
    Dim ColWid As Single
    
    Dim IntClass As New clsInterest
    Dim rebateRate As String
    
        grd.Clear
        grd.Rows = 10


Select Case m_ReportType
    Case repBkccDepBalance, repBKCCLoanBalance
        With grd
            .Cols = 5: .Rows = 5
            .FixedCols = 2: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '"Sl No
            .Col = 1: .Text = GetResourceString(36, 60) 'Account No
            .Col = 2: .Text = GetResourceString(49, 60) 'Memebr No
            .Col = 3: .Text = GetResourceString(35) 'Name
            .Col = 4: .Text = GetResourceString(42) '"Balance"
        End With

    Case repBKCCLoanIssued, repBKCCLoanReturned
        With grd
            .Cols = 6: .Rows = 5
            .FixedCols = 2: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '"Sl No
            .Col = 1: .Text = GetResourceString(36, 60) 'Account No
            .Col = 2: .Text = GetResourceString(49, 60) 'Member No
            .Col = 3: .Text = GetResourceString(35) 'Name
            .Col = 4: .Text = GetResourceString(340)  ' "Date
            .Col = 5: .Text = GetResourceString(42) '"Balance"
        End With

    Case repBkccDepHolder, repBkccLoanHolder
        With grd
            .Cols = 11
            .Rows = 10
            .FixedCols = 1: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .Col = 1: .Text = GetResourceString(36, 60) 'Account No
            .Col = 2: .Text = GetResourceString(49, 60) 'Member No
            .Col = 3: .Text = GetResourceString(35) 'Name
            .Col = 4: .Text = GetResourceString(340) ' "Date
            .Col = 5: .Text = GetResourceString(111) ' "Caste
            .Col = 6: .Text = GetResourceString(112) ' "Place
            .Col = 7: .Text = GetResourceString(67, 58) '"Loan Balance"
            .Col = 8: .Text = GetResourceString(344) '"Regular Interest"
            .Col = 9: .Text = GetResourceString(345) '"Penal Interest"
            .Col = 10: .Text = GetResourceString(52) '"total"
            If m_ReportType = repBkccDepHolder Then
                .Cols = 10
                .Col = 8: .Text = GetResourceString(47) '"Interest"
                .Col = 9: .Text = GetResourceString(52) '"total"
            End If
        End With

    Case repBKCCLoanDailyCash, repBkccDepDailyCash
        With grd
            .Cols = 11: .Rows = 11
            .FixedCols = 1: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Slno
            .Col = 1: .Text = GetResourceString(37) 'Date
            .Col = 2: .Text = GetResourceString(58, 60) 'Loan No
            .Col = 3: .Text = GetResourceString(49, 60) 'Memeber No
            .Col = 4: .Text = GetResourceString(35) 'Name
            .Col = 5: .Text = GetResourceString(271)  '"Deposit
            .Col = 6: .Text = GetResourceString(272) '"WithDraw
            .Col = 7: .Text = GetResourceString(344) 'Regular Interest"
            .Col = 8: .Text = GetResourceString(345) 'PenalInterest"
            .Col = 9: .Text = GetResourceString(487) 'Interest"
            .Col = 10: .Text = GetResourceString(67, 58) 'Loan Balance
            If m_ReportType = repBkccDepDailyCash Then
                .Col = 7: .Text = GetResourceString(487) 'Interest"
                .Col = 8: .Text = GetResourceString(483) 'Balance"
                .Col = 9: .Text = GetResourceString(42) 'Balance"
                .Cols = 10
            End If
        End With

    Case repBKCCDepDayBook, repBKCCLoanDayBook
        With grd
            .Cols = 15: .Rows = 10
            .FixedCols = 1: .FixedRows = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Slno
            .Col = 1: .Text = GetResourceString(37) 'Date
            .Col = 2: .Text = GetResourceString(36, 60) 'AccNum
            .Col = 3: .Text = GetResourceString(49, 60) 'AccNum
            .Col = 4: .Text = GetResourceString(35) 'Name
            .Col = 5: .Text = GetResourceString(271) '"Deposit"
            .Col = 6: .Text = GetResourceString(271) '"Deposit"
            .Col = 7: .Text = GetResourceString(272)  '"With drwa
            .Col = 8: .Text = GetResourceString(272)  '"With drwa
            .Col = 9: .Text = GetResourceString(344) 'REGULAR Interest"
            .Col = 10: .Text = GetResourceString(344) 'Regular Interest"
            .Col = 11: .Text = GetResourceString(345) 'PenalInterest
            .Col = 12: .Text = GetResourceString(345) 'Penal Intere
            .Col = 13: .Text = GetResourceString(487) 'Interest Paid
            .Col = 14: .Text = GetResourceString(67, 58) 'Loan Balance
            If m_ReportType = repBKCCDepDayBook Then
                .Col = 11: .Text = GetResourceString(43) & " " & _
                    GetResourceString(42) 'Deposit Balance
            End If

            .Row = 1
            .Col = 0: .Text = GetResourceString(33) 'Slno
            .Col = 1: .Text = GetResourceString(37) 'Date
            .Col = 2: .Text = GetResourceString(36, 60) '"AccNum
            .Col = 3: .Text = GetResourceString(49, 60) '"Member Num
            .Col = 4: .Text = GetResourceString(35) 'Name
            .Col = 5: .Text = GetResourceString(269)  '"Cash
            .Col = 6: .Text = GetResourceString(270)  '"Contra
            .Col = 7: .Text = GetResourceString(269) '"Cash
            .Col = 8: .Text = GetResourceString(270) '"Contra
            .Col = 9: .Text = GetResourceString(269) '"Cash
            .Col = 10: .Text = GetResourceString(270) '"Contra
            .Col = 11: .Text = GetResourceString(269) '"Cash
            .Col = 12: .Text = GetResourceString(270) '"Contra
            .Col = 13: .Text = GetResourceString(487) 'Interest Paid
            .Col = 14: .Text = GetResourceString(67, 58) 'Loan Balance
            If m_ReportType = repBKCCDepDayBook Then
                .Col = 10: .Text = GetResourceString(43) & " " & _
                    GetResourceString(42) 'Deposit Balance
                .Cols = 11
            End If
        End With

    Case repBkccOD
        With grd
            .Clear
            .Cols = 13
            .FixedCols = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Sl No "
            .Col = 1: .Text = GetResourceString(36, 60) '" Loan No "
            .Col = 2: .Text = GetResourceString(49, 60) '" Member No "
            .Col = 3: .Text = GetResourceString(35) ' "Loan Holder" '
            .Col = 4: .Text = GetResourceString(209)  '" Loan due date" '
            .Col = 5: .Text = GetResourceString(290, 37)  ''" Loan issue date" '
            .Col = 6: .Text = GetResourceString(67, 58) '"Loan Amount"
            .Col = 7: .Text = GetResourceString(84, 18) '"Od Amount"
            .Col = 8: .Text = GetResourceString(344) '"REGULAR INTEREST"
            .Col = 9: .Text = GetResourceString(345) '"PENALINTEREST"
            .Col = 10: .Text = GetResourceString(52) '"Total"
            .Col = 11: .Text = GetResourceString(389) & 1 '"GUarantor"
            .Col = 12: .Text = GetResourceString(389) & 2 '"GUarantor"
        End With

     Case repBkccDepIntPaid, repBKCCLoanIntCol
        With grd
            .Cols = 8
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(37)  '" Date"
            .Col = 2: .Text = GetResourceString(36, 60)  '" Loan No "
            .Col = 3: .Text = GetResourceString(49, 60)  '" Member No "
            .Col = 4: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 5: .Text = GetResourceString(344)    '''"REGULAR INTEREST"
            .Col = 6: .Text = GetResourceString(345)    '''"PENALINTEREST"
            .Col = 7: .Text = GetResourceString(487)    '''"INTEREST PAID"
            If m_ReportType = repBkccDepIntPaid Then
                .Cols = 7
                .Col = 5: .Text = GetResourceString(487)    '''"INTEREST PAID"
                .Col = 6: .Text = GetResourceString(483)    '''"INTEREST RECEIVED"
            End If
        End With

    Case repBkccDepGLedger, repBKCCLoanGLedger
        With grd
            .Cols = 6: .Rows = 5
            .FixedRows = 1: .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) ' "Sl No"
            .Col = 1: .Text = GetResourceString(37) ' "Date"
            .Col = 2: .Text = GetResourceString(284)  '"Opening Blance
            .Col = 3: .Text = GetResourceString(271)  '"Deposit
            .Col = 4: .Text = GetResourceString(272)  '"Withdraw"
            .Col = 5: .Text = GetResourceString(285)  '"Closing Balanc"
        End With
    Case repBkccGuarantor
        With grd
            .Cols = 6
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan Id "
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 3: .Text = GetResourceString(67, 58) '"Loan Balance" '
            .Col = 4: .Text = GetResourceString(389) & "1"  '' Guarantors1
            .Col = 5: .Text = GetResourceString(389) & "2"  '' Guarantors2
        End With
    Case repBkccLoanMonBal, repBkccDepMonBal
        With grd
            .Cols = 4
            .FixedCols = 2: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan Id "
            .Col = 2: .Text = GetResourceString(49, 60)  '" Member Id "
            .Col = 3: .Text = GetResourceString(35) '"Loan Holder" '
        End With
    Case repBkccMonTrans
        With grd
            .Cols = 8
            .FixedCols = 2: .FixedRows = 4
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 3: .Text = GetResourceString(247) '"Sanction Amount
            .Col = 4: .Text = GetResourceString(247) '"Sanction Amount
            .Col = 5: .Text = GetResourceString(247) '"Sanction Amount
            .Col = 6: .Text = GetResourceString(52) 'Total
            .Col = 7: .Text = GetResourceString(284) 'Opening Balance
            If m_ReportType = repBkccDepMonTrans Then _
                .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            .Row = 1
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 3: .Text = "3 " & GetResourceString(251, 247) '"Sanction Amount
            .Col = 4: .Text = GetResourceString(374, 247) '"Sanction Amount
            .Col = 5: .Text = "Extra " & GetResourceString(247) '"Sanction Amount
            .Col = 6: .Text = GetResourceString(52) 'Total
            .Col = 7: .Text = GetResourceString(284) 'Balance
            If m_ReportType = repBkccDepMonTrans Then _
                .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            .Row = 2
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 3: .Text = "3 " & GetResourceString(251) & " " & _
                                GetResourceString(247) '"Sanction Amount
            .Col = 4: .Text = GetResourceString(374) & " " & _
                                GetResourceString(247) '"Sanction Amount
            .Col = 5: .Text = "Extra " & GetResourceString(247) '"Sanction Amount
            .Col = 6: .Text = GetResourceString(52) 'Total
            .Col = 7: .Text = GetResourceString(284) 'Balance
            If m_ReportType = repBkccDepMonTrans Then _
                .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            If m_ReportType = repBkccDepMonTrans Then .Cols = .Cols - 4
            .Row = 3
            For ColCount = 0 To .Cols - 1
                .Col = ColCount
                .Text = ColCount + 1
            Next
        End With
    
    Case repBkccDepMonTrans
        With grd
            .Cols = 4
            .FixedCols = 2: .FixedRows = 4
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            .Row = 1
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            
            .Row = 2
            .Col = 0: .Text = GetResourceString(33) '" Slno"
            .Col = 1: .Text = GetResourceString(36, 60)  '" Loan No
            .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
            .Col = 3: .Text = GetResourceString(42) & " " & GetFromDateString(FinIndianFromDate)   'Balance AS on 31/3
            
            .Row = 3
            For ColCount = 0 To .Cols - 1
                .Col = ColCount
                .Text = ColCount + 1
            Next
        End With
        
    Case repBKCCMemberTrans
        With grd
            .Clear
            .Rows = 5
            .Cols = 8
            .FixedCols = 2
            .FixedRows = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .Col = 1: .Text = GetResourceString(58) & " " & _
                                GetResourceString(60)   '"Loan No
            .Col = 2: .Text = GetResourceString(49, 60) 'Memeber ID
            .Col = 3: .Text = GetResourceString(35) 'Name
            .Col = 4: .Text = GetResourceString(272) 'Withdraw
            .Col = 5: .Text = GetResourceString(271) 'Deposit
            .Col = 6: .Text = GetResourceString(271) 'Deposit
            .Col = 7: .Text = GetResourceString(271) 'Deposit
            .Row = 1
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .Col = 1: .Text = GetResourceString(58) & " " & _
                                GetResourceString(60)   '"Loan No
            .Col = 2: .Text = GetResourceString(49, 60) 'Memeber ID
            .Col = 3: .Text = GetResourceString(35) 'Name
            .Col = 4: .Text = GetResourceString(272) 'Withdraw
            .Col = 5: .Text = GetResourceString(310) 'Deposit
            .Col = 6: .Text = GetResourceString(344) 'Reg Interesr
            .Col = 7: .Text = GetResourceString(345) 'Penal INterest
        End With
    Case repBKCCLoanClaimBill
        With grd
            .Clear
            .Rows = 5
            .Cols = 23
            .FixedCols = 2
            .FixedRows = 3
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .ColAlignment(0) = flexAlignLeftCenter
            .Col = 1: .Text = GetResourceString(35) 'Name
            .ColAlignment(1) = flexAlignLeftCenter
            .Col = 2: .Text = GetResourceString(112) 'Place
            .ColAlignment(2) = flexAlignLeftCenter
            
            .Col = 3: .Text = GetResourceString(58, 60) '"Loan No
            .ColAlignment(3) = flexAlignLeftCenter
            .Col = 4: .Text = GetResourceString(58, 37) '"Loan Date
            .ColAlignment(4) = flexAlignLeftCenter
            .Col = 5: .Text = "Crop" '"Loan No
            .Col = 6: .Text = GetResourceString(58, 40) '"Loan Amount
            .ColAlignment(6) = flexAlignRightCenter
            .Col = 7: .Text = GetResourceString(43) '"Deposit
            .Col = 8: .Text = GetResourceString(436, 91) '"Subsidy Amount
            .ColAlignment(8) = flexAlignRightCenter
            .Col = 9: .Text = GetResourceString(209) ' & " " & GetResourceString(37) '"Due Date
            .ColAlignment(9) = flexAlignRightCenter
            
            '----
            .Col = 10: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(10) = flexAlignLeftCenter
            .Col = 11: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(11) = flexAlignLeftCenter
            .Col = 12: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(12) = flexAlignRightCenter
            .Col = 13: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(13) = flexAlignRightCenter
            '----
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Subsidy", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "10"
            .Col = 14: .Text = rebateRate & "% " & GetResourceString(47, 436, 91) 'Interest Subsidy
            .ColAlignment(14) = flexAlignRightCenter
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "3"
            .Col = 15: .Text = rebateRate & "% " & GetResourceString(47, 437, 91) 'Interest Rebate
            .ColAlignment(15) = flexAlignRightCenter
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "NabardRate", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "2.5"
            .Col = 16: .Text = "Nabard  " & rebateRate & "% " & GetResourceString(47, 437, 91) 'Interest Rebate
            .ColAlignment(16) = flexAlignRightCenter
            
            '.Col = 16: .Text = GetResourceString(261) 'Remark
            
            .Row = 1
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .Col = 1: .Text = GetResourceString(35) 'Name
            .Col = 2: .Text = GetResourceString(112) 'Place
            .Col = 3: .Text = GetResourceString(58, 60) '"Loan No
            .Col = 4: .Text = GetResourceString(58, 37) '"Loan Date
            .Col = 5: .Text = "Crop" '"Loan No
            .Col = 6: .Text = GetResourceString(58, 40) '"Loan Amount
            .Col = 7: .Text = GetResourceString(43) '"Deposit
            .Col = 8: .Text = GetResourceString(436, 91) '"Subsidy Amount
            .Col = 9: .Text = GetResourceString(209) ' & " " & GetResourceString(37) '"Due Date
            '----
            .Col = 10: .Text = GetResourceString(41) '& " " & GetResourceString(60) '"Voucher No
            .Col = 11: .Text = GetResourceString(37) '"Repayment Date
            .Col = 12: .Text = GetResourceString(91) '"Repayment Amount
            .Col = 13: .Text = GetResourceString(44) & GetResourceString(92) '"Days
            '----
            .Col = 14: .Text = GetResourceString(47, 436, 91) 'Interest Subsidy
            .Col = 15: .Text = GetResourceString(47, 437, 91) 'Interest Rebate
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Subsidy", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "10"
            .Col = 14: .Text = rebateRate & "% " & GetResourceString(47, 436, 91) 'Interest Subsidy
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "3"
            .Col = 15: .Text = rebateRate & "% " & GetResourceString(47, 437, 91) 'Interest Rebate
            rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "NabardRate", FinUSFromDate))
            If Val(rebateRate) = 0 Then rebateRate = "2.5"
            .Col = 16: .Text = "Nabard  " & rebateRate & "% " & GetResourceString(47, 437, 91) 'Interest Rebate
            
            '.Col = 16: .Text = GetResourceString(261) 'Remark
            .Col = 17: .Text = GetResourceString(464, 60) '"AAdar Number
            .Col = 18: .Text = GetResourceString(239, 60) 'mObile Number
            .Col = 19: .Text = GetResourceString(446, 36, 60) '"DCC Bank AccOunt
            .Col = 20: .Text = GetResourceString(446, 482) 'DCC Bank IFSC
            .Col = 21: .Text = GetResourceString(237, 418, 36)  '"Other Bank AccOunt
            .Col = 22: .Text = GetResourceString(237, 418, 482) 'OTHR Bank IFSC
                        
            .Row = 2
            For ColCount = 0 To .Cols - 1
                .Col = ColCount
                .Text = ColCount + 1
            Next
            .TextMatrix(2, 0) = "  1"
        End With

    Case repBKCCLoanClaimBill_Yearly, repBKCCLoanClaimBill_PrevYearly
        With grd
            .Clear
            .Rows = 5
            .Cols = 20
            .FixedCols = 2
            .FixedRows = 3
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .ColAlignment(0) = flexAlignLeftCenter
            .Col = 1: .Text = GetResourceString(35) 'Name
            .ColAlignment(1) = flexAlignLeftCenter
            '.Col = 2: .Text = GetResourceString(112) 'Place .ColAlignment(2) = flexAlignLeftCenter
            
            .Col = 2: .Text = GetResourceString(58, 60) '"Loan No
            .ColAlignment(2) = flexAlignLeftCenter
            .Col = 3: .Text = GetResourceString(58, 37) '"Loan Date
            .ColAlignment(3) = flexAlignLeftCenter
            .Col = 4: .Text = GetResourceString(58, 40) '"Loan Amount
            .ColAlignment(4) = flexAlignRightCenter
            .Col = 5: .Text = GetResourceString(209) ' & " " & GetResourceString(37) '"Due Date
            .ColAlignment(5) = flexAlignRightCenter
            '----
            
            .Col = 6: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(6) = flexAlignLeftCenter
            .Col = 7: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(7) = flexAlignLeftCenter
            .Col = 8: .Text = GetResourceString(216, 295) '"Repayment Details
            .ColAlignment(8) = flexAlignRightCenter
            '----
            rebateRate = IntClass.InterestRate(wis_BKCCLoan, "Rebate1", FinUSFromDate)
            If Val(rebateRate) = 0 Then rebateRate = "2"
            .Col = 9: .Text = rebateRate & "% " & GetResourceString(47, 436) 'Interest Subsidy
            .ColAlignment(9) = flexAlignRightCenter
            .Col = 10: .Text = rebateRate & "% " & GetResourceString(47, 436) 'Interest Subsidy
            
            rebateRate = IntClass.InterestRate(wis_BKCCLoan, "Rebate2", FinUSFromDate)
            If Val(rebateRate) = 0 Then rebateRate = "3"
            
            .Col = 11: .Text = rebateRate & "% " & GetResourceString(47, 436) 'Interest Subsidy
            .ColAlignment(11) = flexAlignRightCenter
            .Col = 12: .Text = rebateRate & "% " & GetResourceString(47, 436) 'Interest Subsidy
            .Col = 13: .Text = GetResourceString(261) 'Remark
            
            .Row = 1
            .Col = 0: .Text = GetResourceString(33) 'Sl No
            .Col = 1: .Text = GetResourceString(35) 'Name
            .Col = 2: .Text = GetResourceString(58, 60) '"Loan No
            .Col = 3: .Text = GetResourceString(58, 37) '"Loan Date
            .Col = 4: .Text = GetResourceString(58, 40) '"Loan Amount
            .Col = 5: .Text = GetResourceString(209)
            '----
            .Col = 6: .Text = GetResourceString(41) '& " " & GetResourceString(60) '"Voucher No
            .Col = 7: .Text = GetResourceString(37) '"Repayment Date
            .Col = 8: .Text = GetResourceString(91) '"Repayment Amount
            '----
            
            .Col = 9: .Text = GetResourceString(44) & GetResourceString(92) '"Days
            .Col = 10: .Text = GetResourceString(91) 'Interest Subsidy
            .Col = 11: .Text = GetResourceString(44) & GetResourceString(92) '"Days
            .Col = 12: .Text = GetResourceString(91) 'Interest Subsidy
            
            .Col = 13: .Text = GetResourceString(261) 'Remark
            
            '''KYC Columns
            .Col = 13: .Text = GetResourceString(464, 60) '"Adhaar Number
            .Col = 14: .Text = GetResourceString(239, 60) 'mObile Number
            .Col = 15: .Text = GetResourceString(446, 36, 60) '"DCC Bank AccOunt
            .Col = 16: .Text = GetResourceString(446, 482) 'DCC Bank IFSC
            .Col = 17: .Text = GetResourceString(237, 418, 36) & GetResourceString(92) '"Other Bank AccOunt
            .Col = 18: .Text = GetResourceString(237, 418, 482) 'OTHR Bank IFSC
            
            
            .Row = 2
            For ColCount = 0 To .Cols - 1
                .Col = ColCount
                .Text = ColCount + 1
            Next
            .TextMatrix(2, 0) = "  1"
        End With

End Select

Dim count As Integer
Dim RowCount As Integer

With grd
    If .FixedRows > 1 Then .MergeCells = flexMergeFree
    For RowCount = 0 To .FixedRows - 1
        .WordWrap = True
        .Row = RowCount
        If .FixedRows > 1 Then .MergeRow(RowCount) = True
        For count = 0 To .Cols - 1
            .Col = count
            .MergeCol(count) = IIf(.FixedRows > 1, True, False)
            .CellFontBold = True
            .CellAlignment = 4
        Next
    Next
End With

End Sub

Private Sub ReportInterestRecieved()
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim transType As wisTransactionTypes
Dim TotalPenal As Currency
Dim TotalReg As Currency

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)
Dim Deposit As Boolean
Deposit = IIf(m_ReportType = repBkccDepIntPaid, True, False)
Dim SqlStmt

' Display status.
' Build the report query.
SqlStmt = "SELECT A.LoanID,AccNum, A.TransDate,name as CustName," _
        & " A.TransType, Intamount, PenalIntAmount,MiscAmount,MemberNum " _
        & " FROM (BkccIntTrans A INNER JOIN BkccMaster B On  " _
        & " A.loanid = B.loanid) INNER JOIN qryMemName C " _
        & " ON B.Memid = C.Memid " _
        & " WHERE Deposit = " & Deposit _
        & " AND TransDate >= #" & m_FromDate & "#" _
        & " AND TransDate <= #" & m_ToDate & "#"
                 
If m_ReportType = repBkccDepIntPaid Then
    SqlStmt = SqlStmt & " And Deposit = True "
Else
    SqlStmt = SqlStmt & " And Deposit = False "
End If
If m_FarmerType Then _
    SqlStmt = SqlStmt & " And FarmerType = " & m_FarmerType
If m_Gender Then _
    SqlStmt = SqlStmt & " And Gender = " & m_Gender

' Finally, add the sorting clause.
'andassign to db class
If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & " ORDER BY a.TransDate,IsciName"
Else
    SqlStmt = SqlStmt & " ORDER BY a.TransDate,val(AccNum)"
End If

'Execute the query...
gDbTrans.SqlStmt = SqlStmt
SqlStmt = ""
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

'Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data to be written into the grid.", 0)

Call InitGrid

Dim SlNo As Long
Dim rowno As Integer, colno As Integer

' Fill the rows
SlNo = 1
grd.Rows = 20
rowno = grd.Row
    
Do While Not rptRS.EOF
    If FormatField(rptRS("IntAmount")) = 0 And FormatField(rptRS("PenalIntAmount")) = 0 Then GoTo nextRecord
    With grd
        ' Set the row.
        If .Rows <= rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        transType = FormatField(rptRS("TransType"))
        'Fill the loan id.
        colno = 0
        .TextMatrix(rowno, colno) = SlNo: SlNo = SlNo + 1
        
        ' Fill the loan holder name.
        colno = 1
        .TextMatrix(rowno, colno) = FormatField(rptRS("transdate"))
        colno = 2
        .TextMatrix(rowno, colno) = FormatField(rptRS("AccNum"))
        colno = 3
        .TextMatrix(rowno, colno) = FormatField(rptRS("MemberNum"))
        colno = 4
        .TextMatrix(rowno, colno) = FormatField(rptRS("custname"))
    
        ' Fill the transaction date.
        
        transType = FormatField(rptRS("TransType"))
        If m_ReportType = repBKCCLoanIntCol Then
            If transType = wDeposit Or transType = wContraDeposit Then
                colno = 5: .TextMatrix(rowno, colno) = FormatField(rptRS("IntAmount"))
                TotalReg = TotalReg + Val(.TextMatrix(rowno, colno))
                colno = 6: .TextMatrix(rowno, colno) = FormatField(rptRS("PenalIntAmount"))
                TotalPenal = TotalPenal + Val(.TextMatrix(rowno, colno))
            Else
                colno = 7: .TextMatrix(rowno, colno) = FormatField(rptRS("IntAmount"))
            End If
        Else
            If transType = wDeposit Or transType = wContraDeposit Then
                colno = 6: .TextMatrix(rowno, colno) = FormatField(rptRS("IntAmount"))
                TotalPenal = TotalPenal + Val(.TextMatrix(rowno, colno))
            Else
                colno = 5: .TextMatrix(rowno, colno) = FormatField(rptRS("IntAmount"))
                TotalReg = TotalReg + Val(.TextMatrix(rowno, colno))
            End If
        End If
    End With
    
    SlNo = SlNo + 1
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / rptRS.recordCount)
    rptRS.MoveNext
Loop

With grd
    .Row = rowno
    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .Col = 1: .Text = GetResourceString(286)
    .CellFontBold = True
    If m_ReportType = repBKCCLoanIntCol Then
        .Col = 5: .Text = FormatCurrency(TotalReg): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalPenal): .CellFontBold = True
    Else
        .Col = 5: .Text = FormatCurrency(TotalReg): .CellFontBold = True
    End If
    ' Display the grid.
    .Visible = True
End With

Me.Caption = "INDEX-2000  [List of Interests made...]"

Exit_Line:
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line



End Sub

Private Sub SetKannadaCaption()
    Call SetFontToControls(Me)
    
    'Set kannada caption for the all the controls
    cmdOk.Caption = GetResourceString(11)
    cmdPrint.Caption = GetResourceString(23)
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Set m_frmCancel = New frmCancel
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .ReportTitle = Me.lblReportTitle
    .GridObject = grd
    m_frmCancel.Show
    m_frmCancel.PicStatus.Visible = True
    UpdateStatus m_frmCancel.PicStatus, 0
    .PrintGrid
    Unload m_frmCancel
    
End With

End Sub

Private Sub cmdWeb_Click()

Dim clswebGrid As New clsgrdWeb
With clswebGrid
    Set .GridObject = grd
    .CompanyAddress = ""
    .CompanyName = gCompanyName
    .ReportTitle = lblReportTitle
    Call clswebGrid.ShowWebView '(grd)

End With

End Sub

Private Sub Form_Click()
    Call grd_LostFocus
End Sub

Private Sub Form_Load()

Dim ReportNo As Integer
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'set kannada caption
Call SetKannadaCaption
lblReportTitle.FONTSIZE = 16
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

With grd
    'Init the grid
    .Clear
    .Rows = 20
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    .Text = "No Records Available"
End With

Select Case m_ReportType
    Case repBKCCLoanIssued
        Call ReportLoanTransaction(wWithdraw)
    Case repBkccGuarantor
        lblReportTitle.Caption = GetResourceString(389, 49, 295) '"Guarantor list
        Call ReportGuarantors
    
    Case repBKCCLoanDayBook
        lblReportTitle.Caption = GetResourceString(229) & " " & _
                GetResourceString(63) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)  'Loan Transaction
        Call ReportLoanDayBook
    Case repBKCCDepDayBook
        lblReportTitle.Caption = GetResourceString(43) & " " & _
                GetResourceString(63) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)   'Deposit Transction
        Call ReportDepositDayBook
    
    Case repBKCCLoanBalance
        lblReportTitle.Caption = GetResourceString(58) & " " & _
                GetResourceString(42) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate) 'Loan Balance
        Call ReportBalance
    Case repBkccDepBalance
        lblReportTitle.Caption = GetResourceString(43) & " " & _
            GetResourceString(42) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate) '"Deposit Balance
        Call ReportBalance
    
    Case repBkccOD
        lblReportTitle.Caption = GetResourceString(84) & " " & _
            GetResourceString(18) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate) 'Over Due Loans
        Call ReportOverdueLoans
    
    Case repBKCCLoanDailyCash
        lblReportTitle.Caption = GetResourceString(58) & " " & _
            GetResourceString(85) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
            Call ReportDailyCashbook
    Case repBkccDepDailyCash
        lblReportTitle.Caption = GetResourceString(43) & " " & _
            GetResourceString(85) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
            Call ReportDailyCashbook
    
    Case repBKCCLoanGLedger
        lblReportTitle.Caption = GetResourceString(229) & " " & _
            GetResourceString(93) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate) '"DEPositBlance
            Call ReportGeneralLedger
    Case repBkccDepGLedger
        lblReportTitle.Caption = GetResourceString(43) & " " & _
            GetResourceString(93) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate) '"DEPositBlance
            Call ReportGeneralLedger
            
    Case repBKCCLoanIssued
            ReportSanctionedLoans
    
    Case repBKCCLoanIntCol
        lblReportTitle.Caption = GetResourceString(58) & " " & _
            GetResourceString(487) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ReportInterestRecieved
    Case repBkccDepIntPaid
        lblReportTitle.Caption = GetResourceString(43) & " " & _
            GetResourceString(483) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ReportInterestRecieved
    
    Case repBkccLoanHolder
        lblReportTitle.Caption = GetResourceString(58) & " " & _
                GetResourceString(295) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate) 'Loan Detail
        Call ReportLoanDetails
    Case repBkccDepHolder
        lblReportTitle.Caption = GetResourceString(43) & " " & _
                GetResourceString(295) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate) 'Deposit Details
        Call ReportLoanDetails
    
    Case repBkccLoanMonBal, repBkccDepMonBal
        Call ReportLoanMonthlyBalance
     
    Case repBkccMonTrans
        lblReportTitle.Caption = GetResourceString(58, 231) '& " " & GetResourceString(58)
        Call ReportLoanMonthlyTrans
    
    Case repBkccDepMonTrans
        lblReportTitle.Caption = GetResourceString(58, 43, 231)
        Call ReportDepositMonthlyTrans
    
    Case repBKCCMonthlyRegister
        Call ShowMeetingRegistar
        
    Case repBKCCShedule_1
        Call ShowShed1
    Case repBKCCMemberTrans
        Call ReportCustomerTransaction
    Case repBKCCReceivable
        Call ReportReceivables
    Case repBKCCLoanReturned
        Call ReportLoanTransaction(wDeposit)
    Case repBKCCLoanClaimBill
        lblReportTitle.Caption = GetResourceString(414, 438, 295) '"Claim Stament
        Call ReportClaimBill
    Case repBKCCLoanClaimBill_Yearly
        lblReportTitle.Caption = GetResourceString(436, 438, 295) '"Claim Stament
        Call ReportClaimBill_Yearly
    Case repBKCCLoanClaimBill_PrevYearly
        lblReportTitle.Caption = GetResourceString(436, 438, 295) '"Claim Stament
        Call ReportClaimBill_PrevYearly
    
End Select

Screen.MousePointer = vbDefault

End Sub

Private Function ShowShed1() As Boolean

RaiseEvent Processing("Fetching the records", 0)

On Error GoTo ErrLine
'Declarations
Dim rstLoan As Recordset
Dim RstRepay  As Recordset
'Dim rstAdvance As Recordset

Dim SqlStr As String
Dim count As Integer

Dim SlNo As Integer
Dim strSocName As String
Dim strLoanName As String
Dim strBranchName As String
Dim BankID As Long
Dim OBalance As Currency
Dim RefID As Long
Dim TransID As Long
Dim LoanSchemeID

Dim RowNum As Integer
Dim ColNum As Integer

'Decalration By SHashi
Dim ColAmount() As Currency
Dim SubTotal() As Currency
Dim GrandTotal() As Currency
Dim fromDate As Date
Dim LastDate As Date
Dim ObDate As Date

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed1 = False

Dim LoanCategary As wisLoanCategories
LoanCategary = wisAgriculural

'This Report Includes Only agricultural loans so
ObDate = GetSysFormatDate("1/7/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1))


'Get The Details Of Agri Loans, and Balance as ondate
gDbTrans.SqlStmt = "SELECT Max(TransID) as MAxTransID, LoanID" & _
        " From BKCCTrans Where TransDate <= #" & m_ToDate & "#" & _
        " GROUP BY LoanID"
gDbTrans.CreateView ("qryMaxID")

SqlStr = "SELECT A.LoanId, AccNum, Balance, Name As Name From" & _
    " ((BKCCMaster A Inner Join BKCCTrans B ON A.LoanId = B.LoanId)" & _
    " Inner Join QryName C ON A.CustomerID = C.CustomerId)" & _
    " INNER JOIN qryMaxID D ON B.LoanID = D.LoanID " & _
    " AND B.TransID = D.MaxTransID"

'Select the Farmer Type
If m_FarmerType <> NoFarmer Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If Trim$(m_Place) <> "" Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender

    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoan, adOpenDynamic) < 1 Then gCancel = True: Exit Function
RaiseEvent Initialise(0, rstLoan.recordCount * 2)

'Now fix the headers for the shedule
Call Shed1RowCOl

'Now Get the Loan RepayMent Of the during this period
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where " & _
    " (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " AND TransDate >= #" & ObDate & "# " & _
    " AND TransDate <= #" & m_ToDate & "# " & _
    " GROUP BY LoanID"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(RstRepay, adOpenDynamic) < 1 Then Set RstRepay = Nothing

'Now Get the Loan Advance Of the during this period
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans " & _
    " Where (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " AND TransDate >= #" & ObDate & "# AND TransDate <= #" & m_ToDate & "#" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
'If gDbTrans.Fetch(rstAdvance, adOpenDynamic) < 1 Then GoTo ErrLine
SqlStr = ""

ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim LoanID As Long
Dim L_clsLoan As New clsBkcc
Dim curRepay As Currency
Dim ODAmount As Currency
Dim PrevOD As Currency
Dim AddRow As Boolean

Dim rowno As Integer, colno As Integer

grd.Row = grd.FixedRows
rowno = grd.Row

SlNo = 0

While Not rstLoan.EOF
    'Initialise the Varible
    LoanID = FormatField(rstLoan("LoanID"))
    ColAmount(3) = L_clsLoan.OverDueAmount(LoanID, ObDate)    'Over due as on Opeing date
    ColAmount(4) = L_clsLoan.LoanDemand(LoanID, ObDate, m_ToDate) 'Amount which falls as over due between OB Date & today
    ColAmount(5) = L_clsLoan.OverDueAmount(LoanID, m_ToDate)      'Over due as on today.
    
    ColAmount(5) = ColAmount(3) + ColAmount(4)
    
    'The difference between two date is the Loan demand of that period
    'Therefore
    'ColAmount(4) = ColAmount(5) - ColAmount(3)
    
    curRepay = 0: ' curAdvance = 0
    If Not RstRepay Is Nothing Then
        RstRepay.MoveFirst
        RstRepay.Find "LoanID = " & LoanID
        If Not RstRepay.EOF Then
            If RstRepay("LoanID") = LoanID Then curRepay = FormatField(RstRepay(0))
        End If
    End If
'    Debug.Assert ColAmount(5) - curRepay = L_clsLoan.OverDueAmount(LoanID, , m_toIndianDate)
'    If Not rstAdvance Is Nothing Then
'        rstRepay.FindFirst "LoanID = " & LoanID
'        If Not rstAdvance.NoMatch then
'            if rstAdvance("LoanID") = LoanID Then curAdvance = FormatField(rstAdvance(0))
'        End If
'    End If
    ColAmount(9) = FormatCurrency(curRepay)
    ColAmount(10) = 0: ColAmount(6) = 0: ColAmount(7) = 0
    ColAmount(8) = 0
    'Now Calculate the Recovery demand
    If curRepay > 0 Then
        'Now calculate the recovery against arrears demand
        If ColAmount(3) >= curRepay Then
            ColAmount(6) = curRepay: curRepay = 0
        Else
            ColAmount(6) = ColAmount(3): curRepay = curRepay - ColAmount(6)
        End If
        'Now calculate the recovery against current demand
        If ColAmount(4) >= curRepay Then
            ColAmount(7) = curRepay: curRepay = 0
        Else
            ColAmount(7) = ColAmount(4): curRepay = curRepay - ColAmount(7)
        End If
        'Remianinog amount is the advance recovery
        ColAmount(8) = curRepay
    End If
   
    ColAmount(10) = ColAmount(5) - ColAmount(6) - ColAmount(7)  'OVer due amount as on date
    ODAmount = ColAmount(10)
    
    'Over due amount as on date
    ODAmount = L_clsLoan.OverDueAmount(LoanID, m_ToDate)
    PrevOD = 0
    ColAmount(16) = L_clsLoan.OverDueSince(5, LoanID, m_ToDate)
    If ColAmount(16) > ODAmount Then ColAmount(16) = ODAmount 'Over due since 5 & above 5 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(4, LoanID, m_ToDate) - ColAmount(16)
    If ColAmount(15) > ODAmount Then ColAmount(15) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(3, LoanID, m_ToDate) - ColAmount(15)   'Over due since 3 Years
    If ColAmount(14) > ODAmount Then ColAmount(14) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(2, LoanID, m_ToDate) - ColAmount(14)  'Over due since 2 Years
    If ColAmount(13) > ODAmount Then ColAmount(13) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(1, LoanID, m_ToDate) - ColAmount(13)
    If ColAmount(12) > ODAmount Then ColAmount(12) = ODAmount 'Over due since 1 Year
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = ODAmount 'Over due under one year
    AddRow = False
    For count = 3 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            Exit For
        End If
    Next
    
    If AddRow Then
        With grd
            If .Rows <= rowno + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            SlNo = SlNo + 1
            colno = 0: .TextMatrix(rowno, colno) = SlNo
            colno = 1: .TextMatrix(rowno, colno) = FormatField(rstLoan("AccNum"))
            colno = 2: .TextMatrix(rowno, colno) = FormatField(rstLoan("Name"))
            For count = 3 To .Cols - 1
                If ColAmount(count) < 0 Then ColAmount(count) = 0
                colno = count: .TextMatrix(rowno, colno) = FormatCurrency(ColAmount(count))
                GrandTotal(count) = GrandTotal(count) + ColAmount(count)
            Next
        End With
    End If
    DoEvents
    If gCancel Then rstLoan.MoveLast
    RaiseEvent Processing("Writing the record", (rstLoan.AbsolutePosition / rstLoan.recordCount))
    rstLoan.MoveNext
Wend

AddRow = False
For count = 3 To grd.Cols - 1
    If GrandTotal(count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        colno = 2: .TextMatrix(rowno, colno) = "Grand Total": .CellFontBold = True
        .Row = rowno
        For count = 3 To .Cols - 1
            If GrandTotal(count) < 0 Then ColAmount(count) = 0
            .Col = count: .CellFontBold = True
            .Text = FormatCurrency(GrandTotal(count))
        Next
    End With
End If

Set L_clsLoan = Nothing
If grd.Row = 3 Then gCancel = True: Exit Function
'lblReportTitle = "Demand, collecion and Balance register for the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)
lblReportTitle = GetResourceString(397, 398, 244, 244, 67, 417) & " " & _
        GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)

ShowShed1 = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function

Private Sub Shed1RowCOl()

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Cols = 17: .Rows = 10
    .FixedCols = 2: .FixedRows = 3
    .WordWrap = True: .AllowUserResizing = flexResizeBoth
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of the customer"
    .Col = 3: .Text = GetResourceString(397) '"Demand"
    .Col = 4: .Text = GetResourceString(397) '"Demand"
    .Col = 5: .Text = GetResourceString(397) '"Demand"
    .Col = 6: .Text = GetResourceString(398) '"Recovery against"
    .Col = 7: .Text = GetResourceString(398) '"Recovery against"
    .Col = 8: .Text = GetResourceString(398) '"Recovery against"
    .Col = 9: .Text = GetResourceString(398) '"Recovery against"
    .Col = 10: .Text = GetResourceString(84) '"Overdue"
    .Col = 11: .Text = GetResourceString(84) '"Overdue"
    .Col = 12: .Text = GetResourceString(84) '"Overdue"
    .Col = 13: .Text = GetResourceString(84) '"Overdue"
    .Col = 14: .Text = GetResourceString(84) '"Overdue"
    .Col = 15: .Text = GetResourceString(84) '"Overdue"
    .Col = 16: .Text = GetResourceString(84) '"Overdue"
    
    .Row = 1
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of the customer"
    .Col = 3: .Text = GetResourceString(399) '"Arrears"
    .Col = 4: .Text = GetResourceString(374) '"Current"
    .Col = 5: .Text = GetResourceString(52) & vbCrLf & "(3+4)"  'tOTAL
    .Col = 6: .Text = GetResourceString(399) ' "Arrears Demand"
    .Col = 7: .Text = GetResourceString(374) '"Current Demand"
    .Col = 8: .Text = GetResourceString(355) '"Advance Recovery If any"
    .Col = 9: .Text = GetResourceString(52) & vbCrLf & "(6 + 7 + 8)" 'tOTAL
    .Col = 10: .Text = GetResourceString(84, 67) & vbCrLf & "(5-(6+7))" 'od bALANCE
    .Col = 11: .Text = "0 " & GetResourceString(107) & " 1" & _
            GetResourceString(108, 208) '"Less than One year"
    .Col = 12: .Text = "1 " & GetResourceString(107) & " 2" & _
            GetResourceString(108, 208) '"1 to 2 years"
    .Col = 13: .Text = "2 " & GetResourceString(107) & " 3" & _
            GetResourceString(108, 208) '"2 to 3 years"
    .Col = 14: .Text = "3 " & GetResourceString(107) & " 4" & _
            GetResourceString(108, 208) '"3 to 4 years"
    .Col = 15: .Text = "4 " & GetResourceString(107) & " 5" & _
            GetResourceString(108, 208)  '"4 to 5 years"
    .Col = 16: .Text = GetChangeString(GetResourceString(193), "5 " & GetResourceString(208))  '"Above 5 Years"
    .RowHeight(1) = 800
    
    Dim I As Integer
    Dim j As Integer
    .Row = 2
    For j = 3 To .Cols - 1
        .Col = j: .Text = Format(j, "00")
    Next
    .Col = 0: .Text = "01"
    .Col = 1: .Text = "02"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeRestrictRows
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j: .MergeCol(j) = True
            .CellFontBold = True
            .CellAlignment = 4
        Next
        .MergeRow(I) = True
    Next
End With

End Sub

Private Function ShowMeetingRegistar() As Boolean

Dim SqlPrin As String
Dim SqlInt As String
Dim SqlStr As String
Dim PrinRepay As Currency
Dim IntRepay As Currency

Dim rst As Recordset
Dim rstLoanScheme As Recordset
Dim SchemeName  As String
Dim Date31_3 As Date
Dim DateLastMonth As Date
'INDIAN dATE FORMAT OF ABOVE VARIABLES
Dim IndDate31_3  As String
'Dim IndDateLastMonth As String

Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes

Err.Clear
On Error GoTo ErrLine

'Get all date in the format of system format
IndDate31_3 = "31/3/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))
Date31_3 = GetSysFormatDate(IndDate31_3)

DateLastMonth = GetSysFormatDate("1/" & Month(m_ToDate) & "/" & Year(m_ToDate))
DateLastMonth = DateAdd("d", -1, DateLastMonth)

RaiseEvent Processing("Fetching the records", 0)

m_repType = repMonthlyRegisterAll

Dim rstMaster As Recordset

Dim rstBalance31_3 As Recordset
Dim rstIntBalance31_3 As Recordset

Dim rstBalanceLastMonth As Recordset
Dim rstIntBalLastMonth As Recordset

Dim rstBalanceAsOn As Recordset
Dim rstIntBalanceAsOn As Recordset

Dim rstPrinTransLast As Recordset
Dim rstIntTransLast As Recordset

Dim rstPrinTransAsOn As Recordset
Dim rstIntTransAsOn As Recordset

Screen.MousePointer = vbHourglass

'Get The details of BKCC
SqlStr = "SELECT * From BKCCMaster WHERE " & _
    " LoanId IN (SELECT Distinct LoanID From BKCCTrans) "
'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType

SqlStr = SqlStr & " Order By Val(AccNum)"

DoEvents
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching the record", 0.1)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstMaster, adOpenDynamic) <= 0 Then GoTo ErrLine

'Create the view to get the Max Transaction of the Trans Table
gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID,LoanID FROM BKCCTrans" & _
            " WHERE Deposit = False AND TransDate <= #" & Date31_3 & "#" & _
            " GROUP BY LoanID"
gDbTrans.CreateView ("qry31_3ID")

gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID, LoanID FROM BKCCIntTrans" & _
        " WHERE DEPOSIT = FALSE AND TransDate <= #" & Date31_3 & "# " & _
        " GROUP BY LoanID"
gDbTrans.CreateView ("qry31_3IntID")

'Get the Loan balance on 31/3/yyyy
SqlPrin = "SELECT A.LoanID,Balance,Deposit FROM BKCCTrans A " & _
        " Inner Join qry31_3ID B On A.TransID = B.MaxTransID " & _
        " ANd A.LoanId = B.LoanID WHERE Deposit = False " & _
        " ORDER BY A.LoanId, TransID Desc"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstBalance31_3, adOpenDynamic) < 1 Then Set rstBalance31_3 = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.15)
If gCancel Then Exit Function

'Get the Interest Balance 31/3/yyyy
SqlInt = "SELECT A.LoanID,IntBalance,TransDate FROM BKCCIntTrans A" & _
        " Inner Join qry31_3IntID B ON A.TransId = B.MaxTransID " & _
        " AND A.LoanID = B.LoanID WHERE DEPOSIT = FALSE"

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntBalance31_3, adOpenDynamic) < 1 Then Set rstIntBalance31_3 = Nothing
SqlInt = ""

gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID, LoanID " & _
        " FROM BKCCTrans WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND DEPOSIT = FALSE GROUP BY LOanID"
gDbTrans.CreateView ("qryLastID")
gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID, LoanID " & _
        " FROM BKCCIntTrans WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND DEPOSIT = FALSE GROUP BY LOanID"
gDbTrans.CreateView ("qryLastIntID")
        
'Get the Loan balance as on lastMonth
SqlPrin = "SELECT A.LoanID,TransDate,Balance,Deposit " & _
    " FROM BKCCTrans A Inner Join qryLastID B " & _
    " ON A.TransId = B.MaxTransID" & _
    " AND A.LoanID = B.LoanID WHERE DEPOSIT = FALSE"
'Get the Interest Balance ON  LAST MONTH
SqlInt = "SELECT B.LoanID,IntBalance,TransDate,Deposit FROM BKCCIntTrans A" & _
    " Inner Join qrylastIntID B ON A.TransId = B.MaxTransID" & _
        " AND A.LoanID = B.LoanID WHERE DEPOSIT = FALSE"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstBalanceLastMonth, adOpenDynamic) < 1 Then _
                Set rstBalanceLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.25)
If gCancel Then Exit Function
    
gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntBalLastMonth, adOpenDynamic) < 1 Then _
                Set rstIntBalLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.35)
If gCancel Then Exit Function

'Create the Views to get the max transaction as on to date
gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID,LoanID FROM " & _
        " BKCCTrans WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND DEPOSIT = FALSE GROUP BY LoanID"
gDbTrans.CreateView ("qryToDateID")

gDbTrans.SqlStmt = "SELECT MAX(TransID) as MaxTransID,LoanID FROM " & _
        " BKCCIntTrans WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND DEPOSIT = FALSE GROUP BY LoanID"
gDbTrans.CreateView ("qryToDateIntID")

'Get the Loan balance as on date
SqlPrin = "SELECT A.LoanID,TransDate,Balance,Deposit FROM BKCCTrans A " & _
     " Inner Join qryToDateId B On A.TransId = B.MaxTransID " & _
        " AND A.LoanID = B.LoanID WHERE Deposit = FALSE "
'Get the Interest Balance ON  Date
SqlInt = "SELECT A.LoanID,IntBalance,TransDate,Deposit FROM BKCCIntTrans A" & _
    " Inner Join qryToDateIntID B On A.TransId = B.MaxTransID " & _
        " AND A.LoanID = B.LoanID Where DEPOSIT = FALSE"
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstBalanceAsOn, adOpenDynamic) < 0 Then Set rstBalanceAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.45)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntBalanceAsOn, adOpenDynamic) < 1 Then Set rstIntBalanceAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.5)
If gCancel Then Exit Function

'GEt the Transacted amount After 31/3/yyyy till last month
SqlPrin = "SELECT SUM(AMOUNT) as SumAmount,LoanID,TransType FROM BKCCTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " And Deposit = False GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM BKCCIntTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " And Deposit = False GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransLast, adOpenDynamic) < 1 Then Set rstPrinTransLast = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.55)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransLast, adOpenDynamic) < 1 Then Set rstIntTransLast = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.65)
If gCancel Then Exit Function

'GEt the Transacted amount From last month to till Today
SqlPrin = "SELECT SUM(AMOUNT)as SumAmount,LoanID,TransType FROM BKCCTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " And Deposit = False GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM BKCCIntTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " And Deposit = False GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransAsOn, adOpenDynamic) < 1 Then Set rstPrinTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.75)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransAsOn, adOpenDynamic) < 1 Then Set rstIntTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.85)
If gCancel Then Exit Function

'Now Initialise the grid
With grd
    .Clear
    .Cols = 23
    .Rows = 20
End With

Dim SlNo As Integer
Dim LoanID As Long
Dim L_clsCust As New clsCustReg
Dim L_clsLoan As New clsBkcc
Dim retstr As String
Dim strArr() As String
Dim TransDate As Date
Dim IntRate As Single
Dim Amount As Currency
Dim IntAmount As Currency
Dim Balance As Currency
Dim PrevDate As Date
Dim Balance31_3 As Currency
Dim BalanceLastMonth As Currency
Dim BalanceNow As Currency
Dim IntBal31_3 As Currency
Dim IntBalLastMonth As Currency
Dim IntBalNow As Currency


Dim ODAmount As Currency
Dim ODInt As Currency

Call SetGrid

Call InitMonthRegGrid

RaiseEvent Initialise(0, rstMaster.recordCount)

grd.Row = grd.FixedRows
lblReportTitle = "Meeting register As on " & m_ToIndianDate

Dim rowno As Integer, colno As Integer
RaiseEvent Initialise(0, rstMaster.recordCount)

rowno = grd.Row
Do
    If rstMaster.EOF Then Exit Do
    
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    LoanID = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
    
  With grd
    If .Rows < rowno + 3 Then .Rows = .Rows + 3
    rowno = rowno + 1
    .MergeRow(rowno) = False
    colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00")
    colno = 1: .TextMatrix(rowno, colno) = FormatField(rstMaster("AccNum"))
    colno = 2: .TextMatrix(rowno, colno) = L_clsCust.CustomerName(FormatField(rstMaster("CustomerID")))
    retstr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    If Val(retstr) Then
        colno = 3: .TextMatrix(rowno, colno) = L_clsCust.CustomerName(Val(retstr))
        rowno = rowno + 1: .TextMatrix(rowno, colno) = "": rowno = rowno - 1
    End If
    retstr = FormatField(rstMaster("Guarantor2"))
    If Val(retstr) Then
        colno = 4: .TextMatrix(rowno, colno) = L_clsCust.CustomerName(Val(retstr))
        rowno = rowno + 1: .TextMatrix(rowno, colno) = "": rowno = rowno - 1
    End If
    On Error GoTo ErrLine
    'Loan Advance Date
    colno = 5: .TextMatrix(rowno, colno) = FormatField(rstMaster("IssueDate"))
    'Colno = 6: .TextMatrix(rowNo, colNo)= FormatField(rstMaster("LoanAmount"))
    
    'Loan Balance as on 31/3
    PrevDate = Date31_3
    Balance31_3 = 0: IntBal31_3 = 0
    If Not rstBalance31_3 Is Nothing Then
        rstBalance31_3.MoveFirst
        rstBalance31_3.Find " LoanID = " & LoanID
        If Not rstBalance31_3.EOF Then
            If rstBalance31_3("LoanID") = LoanID Then
                Balance31_3 = FormatField(rstBalance31_3("Balance"))
            End If
        End If
    End If
    If Not rstIntBalance31_3 Is Nothing Then
        rstIntBalance31_3.MoveFirst
        rstIntBalance31_3.Find " LoanID = " & LoanID
        If Not rstIntBalance31_3.EOF Then
            If rstIntBalance31_3("LoanID") = LoanID Then
                TransDate = FormatField(rstIntBalance31_3("TransDate"))
                IntBal31_3 = FormatField(rstIntBalance31_3("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If

    IntBal31_3 = IntBal31_3 + L_clsLoan.RegularInterest(LoanID, Date31_3)
    IntBal31_3 = IntBal31_3 + L_clsLoan.PenalInterest(LoanID, Date31_3)
    If Balance31_3 Then colno = 7: .TextMatrix(rowno, colno) = Balance31_3
    If IntBal31_3 Then colno = 8: .TextMatrix(rowno, colno) = IntBal31_3
    If Balance31_3 Then colno = 9: .TextMatrix(rowno, colno) = Val(Balance31_3 + IntBal31_3)
            
    'Recovery from 1/4/yyyy to LastMonth
    PrinRepay = 0: IntRepay = 0
    transType = wDeposit: ContraTransType = wContraDeposit
    If Not rstPrinTransLast Is Nothing Then
        If gDbTrans.FindRecord(rstPrinTransLast, "LoanID = " & LoanID & ",Transtype = " & transType) Then _
            PrinRepay = FormatField(rstPrinTransLast("SumAmount"))
        If gDbTrans.FindRecord(rstPrinTransLast, "LoanID = " & LoanID & ",Transtype = " & ContraTransType) Then _
            PrinRepay = PrinRepay + FormatField(rstPrinTransLast("SumAmount"))
    End If
    If Not rstIntTransLast Is Nothing Then
        If gDbTrans.FindRecord(rstIntTransLast, "LoanID = " & LoanID & ",Transtype = " & transType) Then _
            IntRepay = FormatField(rstIntTransLast("SumIntAmount"))
        If gDbTrans.FindRecord(rstIntTransLast, "LoanID = " & LoanID & ",Transtype = " & ContraTransType) Then _
            IntRepay = IntRepay + FormatField(rstIntTransLast("SumIntAmount"))
    End If

    If PrinRepay Then colno = 10: .TextMatrix(rowno, colno) = FormatCurrency(PrinRepay)
    If IntRepay Then colno = 11: .TextMatrix(rowno, colno) = FormatCurrency(IntRepay)
    If PrinRepay + IntRepay Then colno = 12: .TextMatrix(rowno, colno) = FormatCurrency(PrinRepay + IntRepay)
            
    'Loan Balance as on end of last month
    BalanceLastMonth = Balance31_3: IntBalLastMonth = 0
    
    If Not rstBalanceLastMonth Is Nothing Then
        rstBalanceLastMonth.MoveFirst
        rstBalanceLastMonth.Find "LoanID = " & LoanID
        If Not rstBalanceLastMonth.EOF Then
            BalanceLastMonth = rstBalanceLastMonth("Balance")
            TransDate = rstBalanceLastMonth("TransDate")
            PrevDate = TransDate
        End If
    End If
    If Not rstIntBalLastMonth Is Nothing Then
        rstIntBalLastMonth.MoveFirst
        rstIntBalLastMonth.Find " LoanID = " & LoanID
        If Not rstIntBalLastMonth.EOF Then _
            If rstIntBalLastMonth("LoanID") = LoanID Then _
                IntBalLastMonth = FormatField(rstIntBalLastMonth("IntBalance"))
    End If

    IntBalLastMonth = IntBalLastMonth + L_clsLoan.RegularInterest(LoanID, DateLastMonth)
    If BalanceLastMonth Then colno = 13: .TextMatrix(rowno, colno) = FormatCurrency(BalanceLastMonth)
    If IntBalLastMonth Then colno = 14: .TextMatrix(rowno, colno) = FormatCurrency(IntBalLastMonth)
    If BalanceLastMonth Then colno = 15: .TextMatrix(rowno, colno) = FormatCurrency(BalanceLastMonth + IntBalLastMonth)
    
    colno = 10: .TextMatrix(rowno, colno) = Balance31_3 - BalanceLastMonth
    If Val(.TextMatrix(rowno, colno)) < 0 Then .TextMatrix(rowno, colno) = "0.00"
    
    'Recovery during this month
    PrinRepay = 0: IntRepay = 0
    transType = wDeposit: ContraTransType = wContraDeposit
    If Not rstPrinTransAsOn Is Nothing Then
        If gDbTrans.FindRecord(rstPrinTransAsOn, "LoanID=" & LoanID & ",Transtype = " & transType) Then _
                PrinRepay = FormatField(rstPrinTransAsOn("SumAmount"))
        If gDbTrans.FindRecord(rstPrinTransAsOn, "LoanID=" & LoanID & ",Transtype = " & ContraTransType) Then _
                PrinRepay = PrinRepay + FormatField(rstPrinTransAsOn("SumAmount"))
    End If
    If Not rstIntTransAsOn Is Nothing Then
        If gDbTrans.FindRecord(rstIntTransAsOn, _
            "LoanID=" & LoanID & ",Transtype = " & transType) Then _
                IntRepay = FormatField(rstIntTransAsOn("SumIntAmount"))
        If gDbTrans.FindRecord(rstIntTransAsOn, _
            "LoanID=" & LoanID & ",Transtype = " & ContraTransType) Then _
                IntRepay = IntRepay + FormatField(rstIntTransAsOn("SumIntAmount"))
    End If
        
    If PrinRepay Then colno = 16: .TextMatrix(rowno, colno) = FormatCurrency(PrinRepay)
    If IntRepay Then colno = 17: .TextMatrix(rowno, colno) = FormatCurrency(IntRepay)
    If PrinRepay + IntRepay <> 0 Then colno = 18: .TextMatrix(rowno, colno) = FormatCurrency(PrinRepay + IntRepay)
    
    'Balance as of now
    BalanceNow = BalanceLastMonth: IntBalNow = 0
    If Not rstBalanceAsOn Is Nothing Then
        rstBalanceAsOn.MoveFirst
        rstBalanceAsOn.Find " LoanID = " & LoanID '& " AND TransType = " & TransType
        If Not rstBalanceAsOn.EOF Then
            If rstBalanceAsOn("LoanID") = LoanID Then
                rstIntBalanceAsOn.MoveFirst
                rstIntBalanceAsOn.Find " LoanID = " & LoanID
                BalanceNow = rstBalanceAsOn("Balance")
                TransDate = rstBalanceAsOn("TransDate")
                IntBalNow = FormatField(rstIntBalanceAsOn("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBalNow = IntBalNow + L_clsLoan.RegularInterest(LoanID, m_ToDate)
    If BalanceNow Then colno = 19: .TextMatrix(rowno, colno) = BalanceNow
    If IntBalNow Then colno = 20: .TextMatrix(rowno, colno) = IntBalNow
    If BalanceNow Then colno = 21: .TextMatrix(rowno, colno) = Val(BalanceNow + IntBalNow)
    
    'Recovery during this Month
    colno = 16
    'DONT REMOVE THE FOLLOWING DEBUG STATEMENT
    'BECAUSE THIS IS VERY IMP AND THIS
    'THIS FUNCTION HAS TO BE IMPROVED IN A SUCH WAY THAT
    'AT ANY COST CODE SHOULD NOT STOP AT THIS JUNCTION
    'Debug.Assert Val(.TextMatrix(rowno, colno)) = BalanceLastMonth - BalanceNow
    '.Col = 17: .Text = IntBalLastMonth - IntBalNow
    '.Col = 18: .Text = Val((BalanceLastMonth - BalanceNow) + (IntBalLastMonth - IntBalNow))

  End With
    DoEvents
    RaiseEvent Processing("Writing the record", rstMaster.AbsolutePosition / rstMaster.recordCount)
    If gCancel Then rstMaster.MoveLast
    rstMaster.MoveNext
Loop

Set L_clsLoan = Nothing
ShowMeetingRegistar = True
Screen.MousePointer = vbDefault
ErrLine:
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, , wis_MESSAGE_TITLE
       'Resume
        Exit Function
    End If

End Function

Private Sub SetGrid()
Dim count As Integer
Dim rst As Recordset

Dim strFirstDay As String
Dim strYear As String
Dim strLastMonth As String
Dim strCurrentMonth As String

strYear = Year(FinUSFromDate)

strFirstDay = FinIndianFromDate

'Now Get the Last Date of the Last Month
strLastMonth = GetIndianDate(GetSysLastDate(DateAdd("m", -1, m_ToDate)))
strCurrentMonth = GetAppLastDate(m_ToIndianDate)

With grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .FixedCols = 0
    .FixedRows = 0
    .MergeCells = flexMergeFree
    .Cols = 1: .Row = 1
End With

'Get the Details of the Loan scheme
CommonSetting:
With grd
    .Cols = 22
    .Rows = 10
    .FixedCols = 2
    .FixedRows = 3
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    .Row = 0:
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3: .Text = GetResourceString(389) '"Guarantor Name & Address"
    .Col = 4: .Text = GetResourceString(389) ' "Guarantor Name & Address"
    .Col = 5: .Text = GetResourceString(340) '"Issue Date"
    .Col = 6: .Text = GetResourceString(80, 40) '"Loan Amount"
LoadResString (gLangOffSet + 80)
    .Col = 7: .Text = GetResourceString(67) & " " & GetFromDateString("31/3/" & strYear)
                        '"Out standing as 31/3/" & strYear & ""
    .Col = 8: .Text = GetResourceString(67) & " " & GetFromDateString("31/3/" & strYear)
                        '"Out standing as 31/3/" & strYear & ""
    .Col = 9: .Text = GetResourceString(67) & " " & GetFromDateString("31/3/" & strYear)
                        '"Out standing as 31/3/" & strYear & ""
    
    '.Col = 10: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 10: .Text = GetResourceString(20) & " " & _
                            GetFromDateString(strFirstDay, strLastMonth)
    .Col = 11: .Text = GetResourceString(20) & " " & _
                            GetFromDateString(strFirstDay, strLastMonth)
    .Col = 12: .Text = GetResourceString(20) & " " & _
                            GetFromDateString(strFirstDay, strLastMonth)
    
    .Col = 13: .Text = "Balance OutStanding as on the end of last month"
    .Col = 13: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strLastMonth)
    .Col = 14: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strLastMonth)
    .Col = 15: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strLastMonth)
                        
    .Col = 16: .Text = "Repayments during this month"
    .Col = 16: .Text = GetResourceString(374) & " " & _
                        GetResourceString(192) & " " & _
                        GetResourceString(20)
    .Col = 17: .Text = GetResourceString(374) & " " & _
                        GetResourceString(192) & " " & _
                        GetResourceString(20)
    .Col = 18: .Text = GetResourceString(374) & " " & _
                        GetResourceString(192) & " " & _
                        GetResourceString(20)
                       
    .Col = 19: .Text = "Balance OutStanding end of this month"
    .Col = 19: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strCurrentMonth)
    .Col = 20: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strCurrentMonth)
    .Col = 21: .Text = GetResourceString(67) & " " & _
                            GetFromDateString(strCurrentMonth)
    
    ''2nd row
    .Row = 1: .MergeRow(3) = True
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3:  .Text = "1" '"Guarantor 1"
    .Col = 4:  .Text = "2" '"Guarantor 2"
    .Col = 5:  .Text = GetResourceString(340) '"Issue Date"
    .Col = 6: .Text = GetResourceString(80) & " " & _
                            GetResourceString(40) '"Loan Amount"

    '.Col = 6: .Text = "Disbursuments from 1/4/" & strYear & " to upto end of last month"
    .Col = 7: .Text = GetResourceString(310) '"Principal"
    .Col = 8: .Text = GetResourceString(47) '"Interest"
    .Col = 9: .Text = GetResourceString(52) '"Total"
    
    .Col = 10: .Text = GetResourceString(310) '"Principal"
    .Col = 11: .Text = GetResourceString(47) '"Interest"
    .Col = 12: .Text = GetResourceString(52) '"Total"
    .Col = 13: .Text = GetResourceString(310) '"Principal"
    .Col = 14: .Text = GetResourceString(47) '"Interest"
    .Col = 15: .Text = GetResourceString(52) '"Total"
    
    .Col = 16: .Text = GetResourceString(310) '"Principal"
    .Col = 17: .Text = GetResourceString(47) '"Interest"
    .Col = 18: .Text = GetResourceString(52) '"Total"
    .Col = 19: .Text = GetResourceString(310) '"Principal"
    .Col = 20: .Text = GetResourceString(47) '"Interest"
    .Col = 21: .Text = GetResourceString(52) '"Total"
    
    .Row = 2: .MergeRow(4) = True
    .MergeCells = flexMergeFree
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    For count = 3 To .Cols - 1
        .Col = count: .Text = (count)
        .CellAlignment = 4
        .MergeCol(count) = True
    Next
    '.Col = 3: .Text = "2b"
    .Col = 2: .Text = "2a"
    .Col = 1: .Text = "2"
    .Col = 0: .Text = "1"
    
    Dim I As Integer, j As Integer
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j
            .CellAlignment = 4: .CellFontBold = True
        Next
    Next
End With

End Sub

Private Sub ReportDailyCashbook()
 
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim SqlStr As String
Dim PrevMemberID As Long
Dim LoanID As Long

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)
Dim Deposit As Boolean
Deposit = IIf(m_ReportType = repBkccDepDailyCash, True, False)

' Display status.
' Build the report query.

SqlStr = "SELECT 'PRINCIPAL',Name as CustName, AccNum, A.loanID," & _
        " A.TransDate,TransID,TransType, Amount, Balance,MemberNum   FROM" & _
        " (BKCCMaster B Inner Join BkCCTrans A On B.LOanID = A.LoanID) " & _
        " Inner Join QryMemName C On B.MemID = C.MemID " & _
        " WHERE Deposit = " & Deposit & _
        " AND TransDate >= #" & m_FromDate & "# " & _
        " AND TransDate <= #" & m_ToDate & "# "

If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)

SqlStr = SqlStr & " UNION " & _
        "SELECT 'INTEREST', Name as CustName, AccNum, A.loanID, C.TransDate,TransID, " & _
        " TransType, IntAmount as Amount, PenalIntAmount as Balance, MemberNum" & _
        " FROM (BkccMaster A Inner Join BKCCIntTrans C On A.LoanID = C.LoanID)" & _
        " Inner Join QryMemName B On A.MemID =B.MemID " & _
        " WHERE Deposit = " & Deposit & _
        " AND TransDate >= #" & m_FromDate & "# " & _
        " AND TransDate <= #" & m_ToDate & "# "

If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)

' Finally, add the sorting clause.
gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate,A.LoanID,TransID "
SqlStr = ""
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
'Set rptRS = gDBTrans.Rst.Clone

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim transType As wisTransactionTypes
Dim TransID As Long

Dim SlNo As Long
Dim TransDate As Date
Dim SubDeposit As Currency
Dim SubWithdraw As Currency
Dim SubInterest As Currency
Dim SubPenalInt As Currency
Dim TotalWithDraw As Currency
Dim TotalDeposit As Currency
Dim TotalInterest As Currency
Dim TotalPenalInt As Currency
Dim SubBalance As Currency
Dim TotalBalance As Currency
Dim PRINTTotal As Boolean
Dim rowno As Integer, colno As Integer

' Fill the rows
SlNo = 0
grd.Rows = 40
rowno = grd.Row

TransDate = rptRS("TransDate")
Do While Not rptRS.EOF
    ' Set the row.
  With grd
    If TransDate <> rptRS("Transdate") Then
        PRINTTotal = True
        SlNo = 0: TransID = 0
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        
        .Row = rowno
        .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
        TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        .Col = 6: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
        TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
        .Col = 7: .Text = FormatCurrency(SubInterest): .CellFontBold = True
        TotalInterest = TotalInterest + SubInterest: SubInterest = 0
        .Col = 8: .Text = FormatCurrency(SubPenalInt): .CellFontBold = True
        TotalPenalInt = TotalPenalInt + SubPenalInt: SubPenalInt = 0
        .Col = .Cols - 1: .Text = FormatCurrency(SubBalance): .CellFontBold = True
        TotalBalance = TotalBalance + SubBalance: SubBalance = 0
        TransDate = rptRS("Transdate")
    End If
    
    If LoanID <> rptRS("LoanID") Then TransID = 0
    If TransID <> rptRS("TransID") Then
        TransID = rptRS("TransID")
        If .Rows <= rowno + 2 Then .Rows = .Rows + 2
        rowno = rowno + 1: SlNo = SlNo + 1
        colno = 0: .TextMatrix(rowno, colno) = SlNo
        colno = 1: .TextMatrix(rowno, colno) = GetIndianDate(TransDate)
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rptRS("AccNum"))
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rptRS("MemberNum"))
        colno = 4: .TextMatrix(rowno, colno) = FormatField(rptRS("custname"))
    End If
    
    'If FormatField(rptRS("Amount")) = 0 Then GoTo NextRecord
    transType = FormatField(rptRS("TransType"))
    
    ' Fill the loan holder name & ' Fill the transaction date.
    transType = FormatField(rptRS("TransType"))
    If rptRS(0) = "PRINCIPAL" Then
        If transType = wDeposit Or transType = wContraDeposit Then
            colno = 5: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubWithdraw = SubWithdraw + Val(.TextMatrix(rowno, colno))
        Else
            colno = 6: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubDeposit = SubDeposit + Val(.TextMatrix(rowno, colno))
        End If
        colno = .Cols - 1
        .TextMatrix(rowno, colno) = FormatCurrency(Abs(rptRS("Balance")))
    Else
        If m_ReportType = repBkccDepDailyCash Then
            If transType = wWithdraw Or transType = wContraWithdraw Then
                colno = 7: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
                SubInterest = SubInterest + Val(.TextMatrix(rowno, colno))
            Else
                colno = 8: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
                SubPenalInt = SubPenalInt + Val(.TextMatrix(rowno, colno))
            End If
        Else
            If transType = wContraDeposit Or transType = wDeposit Then
                colno = 7: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
                SubInterest = SubInterest + Val(.Text)
                colno = 8: .TextMatrix(rowno, colno) = FormatField(rptRS("Balance"))
                SubPenalInt = SubPenalInt + Val(.TextMatrix(rowno, colno))
            Else
                colno = 9: .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            End If
        End If
    End If
  End With

nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / rptRS.recordCount)
    
    LoanID = rptRS("LoanID")
    rptRS.MoveNext
Loop

With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    
    .Row = rowno
    .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
    TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
    .Col = 6: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
    TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
    .Col = 7: .Text = FormatCurrency(SubInterest): .CellFontBold = True
    TotalInterest = TotalInterest + SubInterest: SubInterest = 0
    .Col = 8: .Text = FormatCurrency(SubPenalInt): .CellFontBold = True
    TotalPenalInt = TotalPenalInt + SubPenalInt
    .Col = .Cols - 1: .Text = ""
    TotalBalance = TotalBalance + SubBalance: SubBalance = 0
    If PRINTTotal Then
        If .Rows <= .Row + 3 Then .Rows = .Row + 4
        .Row = .Row + 2
        .Col = 4: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalDeposit):  .CellFontBold = True
        .Col = 7: .Text = FormatCurrency(TotalInterest): .CellFontBold = True
        .Col = 8: .Text = FormatCurrency(TotalPenalInt): .CellFontBold = True
        .Col = .Cols - 1: .Text = ""
    End If
End With
' Display the grid.
grd.Visible = True
Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    GoTo Exit_Line

End Sub

Private Sub ReportLoanDayBook()

' Declare variables...
Dim count As Long
Dim rptRS As Recordset
Dim SqlStr As String
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.

SqlStr = "SELECT 'PRINCIPAL', AccNum, A.loanID, TransDate, TransID," _
    & " name as custname, TransType, Amount, Balance,C.MemberNum FROM (BKCCTrans a" _
    & " INNER JOIN BKCCMaster B ON A.loanid = B.Loanid) " _
    & " INNER JOIN QryMemName C ON B.MemID = C.MemID" _
    & " WHERE Deposit = False AND Amount > 0 " _
    & " AND TransDate >= #" & m_FromDate & "# " _
    & " AND TransDate <= #" & m_ToDate & "# "

'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)

SqlStr = SqlStr & " UNION " _
    & "SELECT 'INTEREST', AccNum, A.loanID, TransDate,TransID," _
    & " Name as Custname, TransType, IntAmount as Amount, PenalIntAmount as Balance,C.MemberNum" _
    & " FROM (BKCCIntTrans A INNER JOIN BKCCMaster B ON a.loanid = B.loanid)" _
    & " INNER JOIN QryMemName C ON B.MemID = C.MemID " _
    & " WHERE (IntAmount > 0 OR PenalIntAmount > 0) And Deposit = False" _
    & " AND TransDate >= #" & m_FromDate & "# " _
    & " AND TransDate <= #" & m_ToDate & "# "

'Select the Farmer Type
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)
    

' Finally, add the sorting clause.
gDbTrans.SqlStmt = SqlStr & " ORDER BY a.TransDate, a.loanid, a.transid"
' Execute the query...
count = gDbTrans.Fetch(rptRS, adOpenDynamic)
If count < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf count = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
'Set rptRS = gDBTrans.Rst.Clone

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim transType As wisTransactionTypes
Dim SlNo As Long
Dim TransDate As Date
Dim SubTotal(5 To 13) As Currency
Dim GrandTotal(5 To 13) As Currency
Dim LoanID As Long
Dim TransID As Long
Dim PRINTTotal As Boolean
Dim rowno As Integer, colno As Integer

TransDate = rptRS("Transdate")
' Fill the rows
SlNo = 0
grd.Row = 1
rowno = 1
grd.Rows = 10
Do While Not rptRS.EOF
    With grd
        'Put the Sub Total
        If TransDate <> rptRS("Transdate") Then
            PRINTTotal = True
            SlNo = 0
            If .Rows <= rowno + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            .Row = rowno
            .Col = 4: .Text = GetResourceString(304)
            .CellFontBold = True: .CellAlignment = 4
            For count = 5 To 13
                .Col = count
                If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
                .CellFontBold = True: .CellAlignment = 7
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End If
        TransDate = rptRS("Transdate")
        If LoanID <> FormatField(rptRS("LoanID")) Then TransID = 0
        LoanID = FormatField(rptRS("LoanID"))
        If TransID <> rptRS("TransId") Then
            SlNo = SlNo + 1
            If .Rows <= rowno + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = SlNo
            colno = 1: .TextMatrix(rowno, colno) = GetIndianDate(TransDate)
            colno = 2: .TextMatrix(rowno, colno) = rptRS("AccNum")
            colno = 3: .TextMatrix(rowno, colno) = rptRS("MemberNum")
            colno = 4: .TextMatrix(rowno, colno) = FormatField(rptRS("CustNAme"))
        End If
        
        TransID = rptRS("TransId")
        TransDate = rptRS("transdate")
        transType = FormatField(rptRS("TransType"))
        If rptRS(0) = "PRINCIPAL" Then
            If transType = wDeposit Then colno = 5
            If transType = wContraDeposit Then colno = 6
            If transType = wWithdraw Then colno = 7
            If transType = wContraWithdraw Then colno = 8
            .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            colno = 14: .TextMatrix(rowno, colno) = FormatField(rptRS("Balance"))
        Else
            If transType = wDeposit Or transType = wContraDeposit Then
                If transType = wDeposit Then colno = 9
                If transType = wContraDeposit Then colno = 10
                .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
                SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
                If transType = wDeposit Then colno = 11
                If transType = wContraDeposit Then colno = 12
                .TextMatrix(rowno, colno) = FormatField(rptRS("Balance"))
                SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            Else
                colno = 13
                .TextMatrix(rowno, colno) = Val(FormatField(rptRS("Amount")) + FormatField(rptRS("Balance")))
                SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            End If
        End If
   
   End With

nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / rptRS.recordCount)
    rptRS.MoveNext
Loop
rptRS.MoveLast

With grd
    .Row = rowno
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(304)
    .CellFontBold = True: .CellAlignment = 4
    For count = 5 To 13
        .Col = count
        If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
        .CellFontBold = True: .CellAlignment = 7
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        SubTotal(count) = 0
    Next

    'Put GrandTotal
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 4: .Text = GetResourceString(286)
        .CellFontBold = True: .CellAlignment = 4
        For count = 5 To 13
            .Col = count
            If GrandTotal(count) Then .Text = FormatCurrency(GrandTotal(count))
            .CellFontBold = True: .CellAlignment = 7
        Next
    End If
End With

' Display the grid.
grd.Visible = True
Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Sub

Private Sub ReportDepositDayBook()

' Declare variables...
Dim count As Long
Dim rptRS As Recordset
Dim SqlStr As String
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.

SqlStr = "SELECT 'PRINCIPAL', B.AccNum, A.loanID, TransDate, TransID," _
    & " name as custname, TransType, Amount, Balance " _
    & " FROM (BKCCTrans A Inner JOin BKCCMaster B On A.LoanID=B.LoanId)" _
    & " Inner Join QryName C On B.CustomerID=C.CustomerID" _
    & " WHERE Amount > 0 And Deposit = True " _
    & " AND TransDate >= #" & m_FromDate & "# " _
    & " AND TransDate <= #" & m_ToDate & "# "
     
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)

SqlStr = SqlStr & " UNION " _
    & "SELECT 'INTEREST', B.AccNum, A.loanID, TransDate,TransID," _
    & " name as custname, TransType, IntAmount as Amount, PenalIntAmount as Balance" _
    & " FROM (BKCCIntTrans A Inner Join BKCCMaster B On A.LoanId=B.LoanID)" _
    & " Inner Join QryName C On B.CustomerID =C.CustomerID" _
    & " WHERE (IntAmount > 0 OR PenalIntAmount >0 )" _
    & " And Deposit = True " _
    & " AND TransDate >= #" & m_FromDate & "# " _
    & " AND TransDate <= #" & m_ToDate & "# "
     
If m_FarmerType Then SqlStr = SqlStr & " And FarmerType = " & m_FarmerType
If m_Gender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If Len(m_Place) Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste)

' Finally, add the sorting clause.
gDbTrans.SqlStmt = SqlStr & " ORDER BY A.TransDate, a.loanid, a.transid"
SqlStr = ""
' Execute the query...

count = gDbTrans.Fetch(rptRS, adOpenDynamic)
If count < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf count = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
'Set rptRS = gDBTrans.Rst.Clone

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim transType As wisTransactionTypes
Dim SlNo As Long
Dim TransDate As Date
Dim SubTotal(4 To 9) As Currency
Dim GrandTotal(4 To 9) As Currency
Dim LoanID As Long
Dim TransID As Long
Dim PRINTTotal As Boolean
Dim rowno As Integer, colno As Integer

TransDate = rptRS("Transdate")
' Fill the rows
SlNo = 0
grd.Row = 1: rowno = 1
grd.Rows = 10
Do While Not rptRS.EOF
    With grd
        'Put the Sub Total
        If TransDate <> rptRS("Transdate") Then
            PRINTTotal = True
            SlNo = 0
            If .Rows <= rowno + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304)
            .CellFontBold = True: .CellAlignment = 4
            For count = 4 To 9
                .Col = count
                If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
                .CellFontBold = True: .CellAlignment = 7
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End If
        TransDate = rptRS("Transdate")
        If LoanID <> FormatField(rptRS("LoanID")) Then TransID = 0
        LoanID = FormatField(rptRS("LoanID"))
        If TransID <> rptRS("TransId") Then
            SlNo = SlNo + 1
            If .Rows <= rowno + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = SlNo
            colno = 1: .TextMatrix(rowno, colno) = GetIndianDate(TransDate)
            colno = 2: .TextMatrix(rowno, colno) = rptRS("AccNum")
            colno = 3: .TextMatrix(rowno, colno) = FormatField(rptRS("CustNAme"))
        End If
        
        TransID = rptRS("TransId")
        TransDate = rptRS("transdate")
        transType = FormatField(rptRS("TransType"))
        If rptRS(0) = "PRINCIPAL" Then
            If transType = wDeposit Then colno = 4
            If transType = wContraDeposit Then colno = 5
            If transType = wWithdraw Then colno = 6
            If transType = wContraWithdraw Then colno = 7
            .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            colno = 10: .TextMatrix(rowno, colno) = FormatCurrency(Abs(rptRS("Balance")))
        Else
            If transType = wWithdraw Then colno = 8
            If transType = wContraWithdraw Then colno = 9
            If transType = wDeposit Or transType = wContraDeposit Then colno = 10
            If transType = wDeposit Or transType = wContraDeposit Then GoTo nextRecord
            .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            If transType = wDeposit Or transType = wContraDeposit Then GoTo nextRecord
        End If
   End With

nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / rptRS.recordCount)
    rptRS.MoveNext
Loop
rptRS.MoveLast

With grd
    .Row = rowno
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(304)
    .CellFontBold = True: .CellAlignment = 4
    For count = 4 To 9
        .Col = count
        If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
        .CellFontBold = True: .CellAlignment = 7
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        SubTotal(count) = 0
    Next

    'Put GrandTotal
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(286)
        .CellFontBold = True: .CellAlignment = 4
        For count = 4 To 9
            .Col = count
            If GrandTotal(count) Then .Text = FormatCurrency(GrandTotal(count))
            .CellFontBold = True: .CellAlignment = 7
        Next
    End If
End With

' Display the grid.
grd.Visible = True
Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
    'Resume
    GoTo Exit_Line

End Sub
Private Sub ReportLoanTransaction(transType As wisTransactionTypes)
 
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim PrevMemberID As String
Dim PrevLoanID As Long
Dim ContraTransType As wisTransactionTypes
Dim actualTransType As wisTransactionTypes
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
ContraTransType = IIf(transType = wWithdraw, wContraWithdraw, wContraDeposit)

gDbTrans.SqlStmt = "SELECT c.customerID,C.MemberNum, a.loanID, a.transDate,B.AccNum, " & _
        "Name as custname, a.transtype, a.amount, a.balance, Remarks " & _
        "FROM (BKCCTrans A INNER JOIN BKCCMaster B ON A.loanid = B.loanid) " & _
        "INNER JOIN QryMemName C ON B.MemID = C.MemID " & _
        "WHERE (TransType = " & transType & " Or TransType = " & ContraTransType & ") AND Transdate >= #" & m_FromDate & "#" & _
        "AND Transdate <= #" & m_ToDate & "#"

If Trim$(m_FromAmt) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount >= " & Val(m_FromAmt)

If Trim$(m_ToAmt) <> "" And Val(m_ToAmt) > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount <= " & Val(m_ToAmt)

If Trim$(m_Place) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And FarmerType = " & m_FarmerType

' Finally, add the sorting clause.
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY a.TransDate, " & _
    IIf(m_ReportOrder = wisByAccountNo, "  val(b.AccNum)", " IsciName") & ", A.Transid"
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
'Set rptRS = gDBTrans.Rst.Clone

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst


' Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.recordCount + 1
    If .Rows < 50 Then .Rows = 50
    .FixedRows = 1
    .FixedCols = 0
    .FormatString = "Sl NO|LoanID  | Name  |Date     |Loan amount"
End With
Call InitGrid

Dim SlNo As Long
Dim TransDate As String
Dim SubWithdraw As Currency
Dim SubDeposit As Currency
Dim SubInterest As Currency
Dim TotalWithDraw As Currency
Dim TotalDeposit As Currency
Dim TotalInterest As Currency
'Dim SubBalance As Currency
'Dim TotalBalance As Currency
Dim rowno As Integer, colno As Integer
rowno = 1
' Fill the rows
SlNo = 1
grd.Rows = 4

Do While Not rptRS.EOF
    With grd
        ' Set the row.
        '.Row = .AbsolutePosition + 1
        If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
        .Rows = rowno + 2
        .Row = rowno
        If TransDate <> "" And TransDate <> FormatField(rptRS("transdate")) Then
            .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            '.Col = .Cols - 1: .Text = FormatCurrency(SubBalance): .CellFontBold = True
            'TotalBalance = TotalBalance + SubBalance: SubBalance = 0
            rowno = rowno + 1
            .Row = rowno
        End If
        TransDate = FormatField(rptRS("transdate"))
        actualTransType = FormatField(rptRS("TransType"))
        
        .Col = 0: .Text = SlNo
        SlNo = SlNo + 1
            
        ' Fill the loan id.
        colno = 1
        If PrevLoanID <> FormatField(rptRS("LoanID")) Then
            .TextMatrix(rowno, colno) = FormatField(rptRS("AccNum"))
            PrevLoanID = FormatField(rptRS("Loanid"))
        End If
        'colNo = 1: .TextMatrix(RowNo, colNo) = FormatField(rptRS("Remarks")): .CellAlignment = 4
    
        ' Fill the loan holder name.
        colno = 3
        If PrevMemberID <> rptRS("AccNum") Then
            .TextMatrix(rowno, colno) = FormatField(rptRS("custname"))
            PrevMemberID = FormatField(rptRS("AccNum"))
             'SubBalance = SubBalance + FormatField(rptRS("Balance"))
        End If
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rptRS("MemberNum"))
        
        ' Fill the transaction date.
        colno = 4
        .TextMatrix(rowno, colno) = TransDate 'FormatField(rptRS("transdate"))
        'TransDate = .Text
        
        If transType = transType Or transType = ContraTransType Then
            colno = 5
            .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
            SubWithdraw = SubWithdraw + Val(.TextMatrix(rowno, colno))
        End If
        
    End With
    ' Move to next row.
    rowno = rowno + 1

nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", SlNo / rptRS.recordCount + 2)
    rptRS.MoveNext
    
Loop

With grd
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    .Row = rowno
    '.Row = .Row + 1
    .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
    '.Col = 6: .Text = FormatCurrency(SubBalance): .CellFontBold = True
    TotalWithDraw = TotalWithDraw + SubWithdraw
    'TotalBalance = TotalBalance + SubBalance
    
    If SubWithdraw <> TotalWithDraw Then
        If .Rows <= .Row + 3 Then .Rows = .Row + 3
        .Row = .Row + 2
        .Col = 3: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True

    End If
    ' Display the grid.
    .Visible = True
    If grd.Rows < 40 Then grd.Rows = 40
End With

Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line
End Sub



Private Sub ReportClaimBill()
 
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim rstKYC As Recordset
Dim LoanID As Long
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes
Dim actualTransType As wisTransactionTypes
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading && Verifying the records ", 0)

' Display status.
' Build the report query.

transType = wDeposit
ContraTransType = IIf(transType = wWithdraw, wContraWithdraw, wContraDeposit)

gDbTrans.SqlStmt = "SELECT c.customerID,C.MemberNum, a.loanID, a.transDate,B.AccNum, " & _
        "Name as custname,Place, a.transtype, a.amount, a.balance, Remarks,A.VoucherNo " & _
        "FROM (BKCCTrans A INNER JOIN BKCCMaster B ON A.loanid = B.loanid) " & _
        "INNER JOIN QryMemName C ON B.MemID = C.MemID " & _
        "WHERE Deposit = False and (TransType = " & transType & " Or TransType = " & ContraTransType & ") AND Transdate >= #" & m_FromDate & "#" & _
        "AND Transdate <= #" & m_ToDate & "#" ' And Balance = 0"

If Trim$(m_FromAmt) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount >= " & Val(m_FromAmt)

If Trim$(m_ToAmt) <> "" And Val(m_ToAmt) > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount <= " & Val(m_ToAmt)

If Trim$(m_Place) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And FarmerType = " & m_FarmerType

' Finally, add the sorting clause.
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY A.TransDate,val(B.AccNum) "
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'GET THE KYC Details
gDbTrans.SqlStmt = "Select * from  KYCTab order by customerid"
Call gDbTrans.Fetch(rstKYC, adOpenDynamic)

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst
rstKYC.MoveLast
rstKYC.MoveFirst
' Initialize the grid.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.recordCount + 1
    If .Rows < 50 Then .Rows = 50
    .FixedRows = 1
    .FixedCols = 0
    .FormatString = "Sl NO|LoanID  | Name  |Date     |Loan amount"
End With
Call InitGrid

Dim SlNo As Long
Dim TransDate As Date
Dim SubLoanAmount As Currency
Dim SubLoanRepay As Currency
Dim SubSubsidy As Currency
Dim SubRebate As Currency
Dim SubNabardAmount As Currency
Dim TotalLoanAmount As Currency
Dim TotalLoanRepay As Currency
Dim TotalSubsidy As Currency
Dim TotalRebate As Currency
Dim TotalNabardAmount As Currency

Dim rowno As Integer, Days As Integer
Dim LoanDate() As Date, LoanAmount() As Currency
Dim repayAmount As Currency, currRepay As Currency
Dim subsidyRate As Single, rebateRate As Single, nabardRate As Single
Dim showSubTotal  As Boolean
Dim suffix As String, payCount As Integer
suffix = " "
Dim IntClass As New clsInterest
subsidyRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Subsidy", FinUSFromDate))
rebateRate = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate", FinUSFromDate))
nabardRate = Val(IntClass.InterestRate(wis_BKCCLoan, "NabardRate", FinUSFromDate))
If subsidyRate = 0 Then subsidyRate = 10
If rebateRate = 0 Then rebateRate = 3
If nabardRate = 0 Then nabardRate = 2.5
Set IntClass = Nothing

rowno = grd.FixedRows
'Fill the rows
SlNo = 1
grd.Rows = 10

Do While Not rptRS.EOF
    With grd
        ' Set the row.
        '.Row = .AbsolutePosition + 1
        If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
        'TransDate = FormatField(rptRS("transdate"))
        TransDate = rptRS("transdate")
        repayAmount = FormatField(rptRS("Amount"))
        
        'Get the Loan Amount for this Loan Id
        'LoanAmount(0) = 0
        LoanID = FormatField(rptRS("LoanId"))
        Call GetLoanDetailsForCliamBill(LoanID, TransDate, repayAmount, LoanDate(), LoanAmount())
        If LoanAmount(0) <= 0 Then GoTo nextRecord
        
        payCount = 0
        suffix = IIf(suffix = " ", "", " ")
        Do
            If LoanAmount(payCount) <= 0 Then Exit Do
            .Rows = rowno + 2
            .Row = rowno
            .TextMatrix(rowno, 0) = SlNo
            SlNo = SlNo + 1
            ' Fill the loan Num, Membername,Place.
            .TextMatrix(rowno, 1) = FormatField(rptRS("custname"))
            .TextMatrix(rowno, 2) = FormatField(rptRS("Place")) & suffix
            .TextMatrix(rowno, 3) = FormatField(rptRS("AccNum"))
            Debug.Assert FormatField(rptRS("AccNum")) <> 2
            ''KYC
            rstKYC.MoveFirst
            rstKYC.Find "CustomerID = " & FormatField(rptRS("CustomerID"))
            If Not rstKYC.EOF Then
                .TextMatrix(rowno, 17) = FormatField(rstKYC("IDText1"))
                .TextMatrix(rowno, 18) = FormatField(rstKYC("CustPhone"))
                .TextMatrix(rowno, 19) = FormatField(rstKYC("ExtAccNum1"))
                .TextMatrix(rowno, 20) = FormatField(rstKYC("ExtIFSC1"))
                .TextMatrix(rowno, 21) = FormatField(rstKYC("ExtAccNum2"))
                .TextMatrix(rowno, 22) = FormatField(rstKYC("ExtIFSC2"))
            rstKYC.MoveFirst
            End If
            'KYC
            .TextMatrix(rowno, 4) = suffix & GetIndianDate(LoanDate(payCount))
            '.TextMatrix(rowno, 5) = "Crop"
            '.TextMatrix(rowno, 7) = "Deposit Amount"
            currRepay = repayAmount
            If repayAmount > LoanAmount(payCount) Then currRepay = LoanAmount(payCount)
            'If repayAmount < LoanAmount(payCount) Then LoanAmount(payCount) = currRepay
            
            .TextMatrix(rowno, 6) = LoanAmount(payCount): SubLoanAmount = SubLoanAmount + LoanAmount(payCount)
            
            .TextMatrix(rowno, 8) = currRepay: SubLoanRepay = SubLoanRepay + currRepay
            
            .TextMatrix(rowno, 9) = GetIndianDate(DateAdd("yyyy", 1, DateAdd("d", -1, LoanDate(payCount)))) & suffix
            .TextMatrix(rowno, 10) = FormatField(rptRS("VoucherNo")) & suffix
            .TextMatrix(rowno, 11) = FormatField(rptRS("transdate")) & suffix
            '.TextMatrix(rowno, 12) = currRepay & suffix
            .TextMatrix(rowno, 12) = suffix & FormatField(rptRS("Amount"))
             
            Days = DateDiff("d", LoanDate(payCount), TransDate)
            .TextMatrix(rowno, 13) = suffix & Days
            'Interest subsidy amount
            .TextMatrix(rowno, 14) = FormatCurrency(currRepay * Days / 365 * subsidyRate / 100)
            .TextMatrix(rowno, 15) = FormatCurrency(currRepay * Days / 365 * rebateRate / 100)
            .TextMatrix(rowno, 16) = FormatCurrency(currRepay * Days / 365 * nabardRate / 100)
            
            SubSubsidy = SubSubsidy + Val(.TextMatrix(rowno, 14))
            SubRebate = SubRebate + Val(.TextMatrix(rowno, 15))
            SubNabardAmount = SubNabardAmount + Val(.TextMatrix(rowno, 16))
            If (SlNo - 1) Mod 20 = 0 Then showSubTotal = True
                            
            ' Move to next row.
            rowno = rowno + 1

            payCount = payCount + 1
            repayAmount = repayAmount - currRepay
            If payCount > UBound(LoanAmount) Or repayAmount <= 0 Then Exit Do
        Loop
        
        If showSubTotal Then
            showSubTotal = False
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            .TextMatrix(rowno, 1) = GetResourceString(304)
            .TextMatrix(rowno, 6) = SubLoanAmount
            .TextMatrix(rowno, 8) = SubLoanRepay
            .TextMatrix(rowno, 12) = SubLoanRepay
            .TextMatrix(rowno, 14) = SubSubsidy
            .TextMatrix(rowno, 15) = SubRebate
            .TextMatrix(rowno, 16) = SubNabardAmount
            
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate
            TotalNabardAmount = TotalNabardAmount + SubNabardAmount: TotalLoanAmount = TotalLoanAmount + SubLoanAmount
            TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            SubSubsidy = 0: SubRebate = 0: SubLoanAmount = 0: SubLoanRepay = 0: SubNabardAmount = 0
            rowno = rowno + 1
            .Rows = rowno + 2
            .Row = rowno
        End If
    End With
    
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / (rptRS.recordCount + 2))
    rptRS.MoveNext
Loop

    If SubLoanRepay > 0 Then
        With grd
        .Row = rowno
        For rowno = 0 To .Cols - 1
            .Col = rowno
            .CellFontBold = True
        Next
        rowno = .Row
        .TextMatrix(rowno, 1) = GetResourceString(304)
        .TextMatrix(rowno, 6) = SubLoanAmount
        .TextMatrix(rowno, 8) = SubLoanRepay
        .TextMatrix(rowno, 12) = SubLoanRepay
        .TextMatrix(rowno, 14) = SubSubsidy
        .TextMatrix(rowno, 15) = SubRebate
        .TextMatrix(rowno, 16) = SubNabardAmount
        
        rowno = rowno + 2
        .Rows = rowno + 2
        .Row = rowno
        
        End With
    End If
    If TotalLoanRepay > 0 Then
        With grd
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate
            TotalNabardAmount = TotalNabardAmount + SubNabardAmount: TotalLoanAmount = TotalLoanAmount + SubLoanAmount
            TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            .TextMatrix(rowno, 1) = GetResourceString(286)
            .TextMatrix(rowno, 6) = TotalLoanAmount:
            .TextMatrix(rowno, 8) = TotalLoanRepay
            .TextMatrix(rowno, 12) = TotalLoanRepay
            .TextMatrix(rowno, 14) = TotalSubsidy
            .TextMatrix(rowno, 15) = TotalRebate
            .TextMatrix(rowno, 16) = TotalNabardAmount
        End With
    End If

        
    
    If grd.Rows < 40 Then grd.Rows = 40

Me.Caption = "INDEX-2000  [Subsidy Claim bill]"

Exit_Line:
    grd.Visible = True
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line
End Sub

Private Sub ReportClaimBill_PrevYearly()
 
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim rstKYC As Recordset
Dim LoanID As Long
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes
Dim actualTransType As wisTransactionTypes
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading && Verifying the records ", 0)

' Display status.
' Build the report query.

transType = wWithdraw
ContraTransType = IIf(transType = wWithdraw, wContraWithdraw, wContraDeposit)

gDbTrans.SqlStmt = "SELECT c.customerID,C.MemberNum, a.loanID, a.transDate,B.AccNum, " & _
        "Name as custname,Place, a.transtype, a.amount, a.balance, Remarks,A.VoucherNo " & _
        "FROM (BKCCTrans A INNER JOIN BKCCMaster B ON A.loanid = B.loanid) " & _
        "INNER JOIN QryMemName C ON B.MemID = C.MemID " & _
        "WHERE Deposit = False and (TransType = " & transType & " Or TransType = " & ContraTransType & ") " & _
        "AND Transdate >= #" & DateAdd("yyyy", -1, m_FromDate) & "# AND Transdate <= #" & DateAdd("yyyy", -1, m_ToDate) & "#"

If Trim$(m_FromAmt) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount >= " & Val(m_FromAmt)

If Trim$(m_ToAmt) <> "" And Val(m_ToAmt) > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount <= " & Val(m_ToAmt)

If Trim$(m_Place) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And FarmerType = " & m_FarmerType

' Finally, add the sorting clause.
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY A.TransDate,val(B.AccNum) "
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'GET THE KYC Details
gDbTrans.SqlStmt = "Select * from  KYCTab order by customerid"
Call gDbTrans.Fetch(rstKYC, adOpenDynamic)

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.recordCount + 1
    If .Rows < 50 Then .Rows = 50
    .FixedRows = 1
    .FixedCols = 0
    .FormatString = "Sl NO|LoanID  | Name  |Date     |Loan amount"
End With
Call InitGrid

Dim SlNo As Long
Dim TransDate As Date
Dim SubLoanAmount As Currency
Dim SubLoanRepay As Currency
Dim SubSubsidy As Currency
Dim SubRebate As Currency
Dim TotalLoanAmount As Currency
Dim TotalLoanRepay As Currency
Dim TotalSubsidy As Currency
Dim TotalRebate As Currency

Dim rowno As Integer, Days As Integer
Dim LastIssueDate As Date, CalcualteDate As Date
Dim loanRepayDate() As Date, LoanRepayAmount() As Currency, repayVoucher() As String
Dim loanIssueAmount As Currency, currRepay As Currency
Dim rebateRate1  As Single, rebateRate2 As Single
Dim showSubTotal  As Boolean
Dim suffix As String, payCount As Integer
suffix = " "
Dim IntClass As New clsInterest

rebateRate1 = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate1", FinUSFromDate))
rebateRate2 = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate2", FinUSFromDate))

If rebateRate1 = 0 Then rebateRate1 = 2
If rebateRate2 = 0 Then rebateRate2 = 3
Set IntClass = Nothing

rowno = grd.FixedRows
'Fill the rows
SlNo = 1
grd.Rows = 10

Do While Not rptRS.EOF
    With grd
        ' Set the row.
        '.Row = .AbsolutePosition + 1
        If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
        
        'TransDate = FormatField(rptRS("transdate"))
        TransDate = rptRS("transdate")
        loanIssueAmount = FormatField(rptRS("Amount"))
        
        'Get the Loan Amount for this Loan Id
        'LoanRepayAmount(0) = 0
        Debug.Assert FormatField(rptRS("AccNum")) <> 224
        LoanID = FormatField(rptRS("LoanId"))
        Call GetRepayDetailsForPrevCliamBill(LoanID, TransDate, FormatField(rptRS("Balance")), loanIssueAmount, loanRepayDate(), LoanRepayAmount(), repayVoucher())
        If LoanRepayAmount(0) <= 0 Then
            'CHECK WITH OTHER PKPS & VEERANNA
            .Rows = rowno + 2
            .Row = rowno
            .TextMatrix(rowno, 0) = SlNo
            SlNo = SlNo + 1
            ' Fill the loan Num, Membername,Place.
            .TextMatrix(rowno, 1) = FormatField(rptRS("custname"))
            .TextMatrix(rowno, 2) = FormatField(rptRS("AccNum"))
            .TextMatrix(rowno, 3) = GetIndianDate(TransDate) & suffix
            .TextMatrix(rowno, 4) = loanIssueAmount
            SubLoanAmount = SubLoanAmount + Val(.TextMatrix(rowno, 4))
            LastIssueDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))
            Days = DateDiff("d", LastIssueDate, DateAdd("YYYY", 1, TransDate)) - 1
            '2% Interest subsidy amount
            .TextMatrix(rowno, 9) = suffix & Days
            .TextMatrix(rowno, 10) = FormatCurrency(loanIssueAmount * Days / 365 * 2 / 100)
            SubSubsidy = SubSubsidy + Val(.TextMatrix(rowno, 10))
                    rowno = rowno + 1
            GoTo nextRecord
        End If
        LastIssueDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))
        payCount = 0
        suffix = IIf(suffix = " ", "", " ")
        Do
            If LoanRepayAmount(payCount) <= 0 Then Exit Do
            .Rows = rowno + 2
            .Row = rowno
            .TextMatrix(rowno, 0) = SlNo
            SlNo = SlNo + 1
            ' Fill the loan Num, Membername,Place.
            .TextMatrix(rowno, 1) = FormatField(rptRS("custname"))
            .TextMatrix(rowno, 2) = FormatField(rptRS("AccNum"))
            .TextMatrix(rowno, 3) = GetIndianDate(TransDate) & suffix
            ''KYC
            rstKYC.MoveFirst
            rstKYC.Find "CustomerID = " & FormatField(rptRS("CustomerID"))
            If Not rstKYC.EOF Then
                .TextMatrix(rowno, 14) = FormatField(rstKYC("IDText1"))
                .TextMatrix(rowno, 15) = FormatField(rstKYC("CustPhone"))
                .TextMatrix(rowno, 16) = FormatField(rstKYC("ExtAccNum1"))
                .TextMatrix(rowno, 17) = FormatField(rstKYC("ExtIFSC1"))
                .TextMatrix(rowno, 18) = FormatField(rstKYC("ExtAccNum2"))
                .TextMatrix(rowno, 19) = FormatField(rstKYC("ExtIFSC2"))
            
            End If
            'KYC
            
            
            CalcualteDate = IIf(loanRepayDate(payCount) = "1/1/2000", LastIssueDate, loanRepayDate(payCount))
            currRepay = loanIssueAmount
            If loanIssueAmount > LoanRepayAmount(payCount) Then currRepay = LoanRepayAmount(payCount)
            If payCount = 0 Then
                .TextMatrix(rowno, 4) = loanIssueAmount
                SubLoanAmount = SubLoanAmount + Val(.TextMatrix(rowno, 4))
            End If
            .TextMatrix(rowno, 5) = suffix & GetIndianDate(DateAdd("yyyy", 1, DateAdd("d", -1, TransDate)))
            
            If UCase(repayVoucher(payCount)) <> "NORECORDS" Then
                .TextMatrix(rowno, 6) = repayVoucher(payCount) & suffix
                .TextMatrix(rowno, 7) = GetIndianDate(CalcualteDate) & suffix
                .TextMatrix(rowno, 8) = suffix & LoanRepayAmount(payCount)
                SubLoanRepay = SubLoanRepay + LoanRepayAmount(payCount)
            End If
           
            Days = DateDiff("d", LastIssueDate, CalcualteDate)
            '2% Interest subsidy amount
            .TextMatrix(rowno, 9) = suffix & Days
            .TextMatrix(rowno, 10) = FormatCurrency(currRepay * Days / 365 * rebateRate1 / 100)
            SubSubsidy = SubSubsidy + Val(.TextMatrix(rowno, 10))
            
            '3% Interest subsidy amount
            If loanRepayDate(payCount) <> "1/1/2000" Then
                Days = DateDiff("d", TransDate, CalcualteDate)
                .TextMatrix(rowno, 11) = suffix & Days
                .TextMatrix(rowno, 12) = FormatCurrency(currRepay * Days / 365 * rebateRate2 / 100)
                SubRebate = SubRebate + Val(.TextMatrix(rowno, 12))
            End If
                            
            If (SlNo - 1) Mod 20 = 0 Then showSubTotal = True
            ' Move to next row.
            rowno = rowno + 1

            payCount = payCount + 1
            loanIssueAmount = loanIssueAmount - currRepay
            If payCount > UBound(LoanRepayAmount) Or loanIssueAmount <= 0 Then Exit Do
        Loop
        
        If showSubTotal Then
            showSubTotal = False
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            .TextMatrix(rowno, 1) = GetResourceString(304)
            .TextMatrix(rowno, 4) = SubLoanAmount
            .TextMatrix(rowno, 8) = SubLoanRepay
            .TextMatrix(rowno, 10) = SubSubsidy
            .TextMatrix(rowno, 12) = SubRebate
            
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate: TotalLoanAmount = TotalLoanAmount + SubLoanAmount: TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            SubSubsidy = 0: SubRebate = 0: SubLoanAmount = 0: SubLoanRepay = 0
            rowno = rowno + 1
            .Rows = rowno + 2
            .Row = rowno
        End If
    End With
    
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / (rptRS.recordCount + 2))
    rptRS.MoveNext
Loop

    If SubLoanAmount > 0 Then
        With grd
        .Row = rowno
        For rowno = 0 To .Cols - 1
            .Col = rowno
            .CellFontBold = True
        Next
        rowno = .Row
            
        .TextMatrix(rowno, 1) = GetResourceString(304)
        .TextMatrix(rowno, 4) = SubLoanAmount
        .TextMatrix(rowno, 8) = SubLoanRepay
        .TextMatrix(rowno, 10) = SubSubsidy
        .TextMatrix(rowno, 12) = SubRebate
        
         rowno = rowno + 2
        .Rows = rowno + 2
        .Row = rowno
        
        End With
    End If
    If TotalLoanAmount > 0 Then
        With grd
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate: TotalLoanAmount = TotalLoanAmount + SubLoanAmount: TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            .TextMatrix(rowno, 1) = GetResourceString(286)
            .TextMatrix(rowno, 4) = TotalLoanAmount
            .TextMatrix(rowno, 8) = TotalLoanRepay
            .TextMatrix(rowno, 10) = TotalSubsidy
            .TextMatrix(rowno, 12) = TotalRebate
        End With
    End If

        
    
    If grd.Rows < 40 Then grd.Rows = 40

Me.Caption = "INDEX-2000  [Yearly Subsidy Claim bill]"

Exit_Line:
    grd.Visible = True
    Exit Sub

Err_line:
    If Err Then
        MsgBox "Report Yearly clain bill: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line
End Sub

Private Sub ReportClaimBill_Yearly()
 
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim rstKYC As Recordset
Dim LoanID As Long
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes
Dim actualTransType As wisTransactionTypes
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading && Verifying the records ", 0)

' Display status.
' Build the report query.

transType = wWithdraw
ContraTransType = IIf(transType = wWithdraw, wContraWithdraw, wContraDeposit)

gDbTrans.SqlStmt = "SELECT c.customerID,C.MemberNum, a.loanID, a.transDate,B.AccNum, " & _
        "Name as custname,Place, a.transtype, a.amount, a.balance, Remarks,A.VoucherNo " & _
        "FROM (BKCCTrans A INNER JOIN BKCCMaster B ON A.loanid = B.loanid) " & _
        "INNER JOIN QryMemName C ON B.MemID = C.MemID " & _
        "WHERE Deposit = False and (TransType = " & transType & " Or TransType = " & ContraTransType & ") AND Transdate >= #" & m_FromDate & "#" & _
        "AND Transdate <= #" & m_ToDate & "#"

If Trim$(m_FromAmt) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount >= " & Val(m_FromAmt)

If Trim$(m_ToAmt) <> "" And Val(m_ToAmt) > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND a.amount <= " & Val(m_ToAmt)

If Trim$(m_Place) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And FarmerType = " & m_FarmerType

' Finally, add the sorting clause.
gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER BY A.TransDate,val(B.AccNum) "
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'KYC
gDbTrans.SqlStmt = "Select * from KYCTab order by CustomerID"
Call gDbTrans.Fetch(rstKYC, adOpenDynamic)

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.recordCount + 1
    If .Rows < 50 Then .Rows = 50
    .FixedRows = 1
    .FixedCols = 0
    .FormatString = "Sl NO|LoanID  | Name  |Date     |Loan amount"
End With
Call InitGrid

Dim SlNo As Long
Dim TransDate As Date
Dim SubLoanAmount As Currency
Dim SubLoanRepay As Currency
Dim SubSubsidy As Currency
Dim SubRebate As Currency
Dim TotalLoanAmount As Currency
Dim TotalLoanRepay As Currency
Dim TotalSubsidy As Currency
Dim TotalRebate As Currency

Dim rowno As Integer, Days As Integer
Dim LastIssueDate As Date, CalcualteDate As Date
Dim loanRepayDate() As Date, LoanRepayAmount() As Currency, repayVoucher() As String
Dim loanIssueAmount As Currency, currRepay As Currency
Dim rebateRate1  As Single, rebateRate2 As Single
Dim showSubTotal  As Boolean
Dim suffix As String, payCount As Integer
suffix = " "
Dim IntClass As New clsInterest

rebateRate1 = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate1", FinUSFromDate))
rebateRate2 = Val(IntClass.InterestRate(wis_BKCCLoan, "Rebate2", FinUSFromDate))
If rebateRate1 = 0 Then rebateRate1 = 2
If rebateRate2 = 0 Then rebateRate2 = 3
Set IntClass = Nothing

rowno = grd.FixedRows
'Fill the rows
SlNo = 1
grd.Rows = 10

Do While Not rptRS.EOF
    With grd
        ' Set the row.
        '.Row = .AbsolutePosition + 1
        If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
        'TransDate = FormatField(rptRS("transdate"))
        TransDate = rptRS("transdate")
        loanIssueAmount = FormatField(rptRS("Amount"))
        
        'Get the Loan Amount for this Loan Id
        'LoanRepayAmount(0) = 0
        Debug.Assert FormatField(rptRS("AccNum")) <> 510
        LoanID = FormatField(rptRS("LoanId"))
        Call GetRepayDetailsForCliamBill(LoanID, TransDate, FormatField(rptRS("Balance")), loanIssueAmount, loanRepayDate(), LoanRepayAmount(), repayVoucher())
        If LoanRepayAmount(0) <= 0 Then GoTo nextRecord
        LastIssueDate = CDate("31/3/" & Year(TransDate) + IIf(Month(TransDate) < 4, 0, 1))
        payCount = 0
        suffix = IIf(suffix = " ", "", " ")
        Do
            If LoanRepayAmount(payCount) <= 0 Then Exit Do
            .Rows = rowno + 2
            .Row = rowno
            .TextMatrix(rowno, 0) = SlNo
            SlNo = SlNo + 1
            ' Fill the loan Num, Membername,Place.
            .TextMatrix(rowno, 1) = FormatField(rptRS("custname"))
            .TextMatrix(rowno, 2) = FormatField(rptRS("AccNum"))
            .TextMatrix(rowno, 3) = GetIndianDate(TransDate) & suffix
            ''KYC
            rstKYC.MoveFirst
            rstKYC.Find "CustomerID = " & FormatField(rptRS("CustomerID"))
            If Not rstKYC.EOF Then
                .TextMatrix(rowno, 14) = FormatField(rstKYC("IDText1"))
                .TextMatrix(rowno, 15) = FormatField(rstKYC("CustPhone"))
                .TextMatrix(rowno, 16) = FormatField(rstKYC("ExtAccNum1"))
                .TextMatrix(rowno, 17) = FormatField(rstKYC("ExtIFSC1"))
                .TextMatrix(rowno, 18) = FormatField(rstKYC("ExtAccNum2"))
                .TextMatrix(rowno, 19) = FormatField(rstKYC("ExtIFSC2"))
            
            End If
            'KYC
            
            
            CalcualteDate = IIf(loanRepayDate(payCount) = "1/1/2000", LastIssueDate, loanRepayDate(payCount))
            currRepay = loanIssueAmount
            If loanIssueAmount > LoanRepayAmount(payCount) Then currRepay = LoanRepayAmount(payCount)
            If payCount = 0 Then
                .TextMatrix(rowno, 4) = loanIssueAmount
                SubLoanAmount = SubLoanAmount + Val(.TextMatrix(rowno, 4))
            End If
            .TextMatrix(rowno, 5) = suffix & GetIndianDate(DateAdd("yyyy", 1, DateAdd("d", -1, TransDate)))
            
            If UCase(repayVoucher(payCount)) <> "NORECORDS" Then
                .TextMatrix(rowno, 6) = repayVoucher(payCount) & suffix
                .TextMatrix(rowno, 7) = GetIndianDate(CalcualteDate) & suffix
                .TextMatrix(rowno, 8) = suffix & LoanRepayAmount(payCount)
                SubLoanRepay = SubLoanRepay + LoanRepayAmount(payCount)
            End If
           
            Days = DateDiff("d", TransDate, CalcualteDate)
            '2% Interest subsidy amount
            .TextMatrix(rowno, 9) = suffix & Days
            .TextMatrix(rowno, 10) = FormatCurrency(currRepay * Days / 365 * rebateRate1 / 100)
            SubSubsidy = SubSubsidy + Val(.TextMatrix(rowno, 10))
            
            '3% Interest subsidy amount
            If loanRepayDate(payCount) <> "1/1/2000" Then
                .TextMatrix(rowno, 11) = suffix & Days
                .TextMatrix(rowno, 12) = FormatCurrency(currRepay * Days / 365 * rebateRate2 / 100)
                SubRebate = SubRebate + Val(.TextMatrix(rowno, 12))
            End If
                            
            If (SlNo - 1) Mod 20 = 0 Then showSubTotal = True
            ' Move to next row.
            rowno = rowno + 1

            payCount = payCount + 1
            loanIssueAmount = loanIssueAmount - currRepay
            If payCount > UBound(LoanRepayAmount) Or loanIssueAmount <= 0 Then Exit Do
        Loop
        
        If showSubTotal Then
            showSubTotal = False
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            .TextMatrix(rowno, 1) = GetResourceString(304)
            .TextMatrix(rowno, 4) = SubLoanAmount
            .TextMatrix(rowno, 8) = SubLoanRepay
            .TextMatrix(rowno, 10) = SubSubsidy
            .TextMatrix(rowno, 12) = SubRebate
            
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate: TotalLoanAmount = TotalLoanAmount + SubLoanAmount: TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            SubSubsidy = 0: SubRebate = 0: SubLoanAmount = 0: SubLoanRepay = 0
            rowno = rowno + 1
            .Rows = rowno + 2
            .Row = rowno
        End If
    End With
    
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", rptRS.AbsolutePosition / (rptRS.recordCount + 2))
    rptRS.MoveNext
Loop

    If SubLoanAmount > 0 Then
        With grd
        .Row = rowno
        For rowno = 0 To .Cols - 1
            .Col = rowno
            .CellFontBold = True
        Next
        rowno = .Row
            
        .TextMatrix(rowno, 1) = GetResourceString(304)
        .TextMatrix(rowno, 4) = SubLoanAmount
        .TextMatrix(rowno, 8) = SubLoanRepay
        .TextMatrix(rowno, 10) = SubSubsidy
        .TextMatrix(rowno, 12) = SubRebate
        
         rowno = rowno + 2
        .Rows = rowno + 2
        .Row = rowno
        
        End With
    End If
    If TotalLoanAmount > 0 Then
        With grd
            .Row = rowno
            For rowno = 0 To .Cols - 1
                .Col = rowno
                .CellFontBold = True
            Next
            rowno = .Row
            
            TotalSubsidy = TotalSubsidy + SubSubsidy: TotalRebate = TotalRebate + SubRebate: TotalLoanAmount = TotalLoanAmount + SubLoanAmount: TotalLoanRepay = TotalLoanRepay + SubLoanRepay
            .TextMatrix(rowno, 1) = GetResourceString(286)
            .TextMatrix(rowno, 4) = TotalLoanAmount
            .TextMatrix(rowno, 8) = TotalLoanRepay
            .TextMatrix(rowno, 10) = TotalSubsidy
            .TextMatrix(rowno, 12) = TotalRebate
        End With
    End If

        
    
    If grd.Rows < 40 Then grd.Rows = 40

Me.Caption = "INDEX-2000  [Yearly Subsidy Claim bill]"

Exit_Line:
    grd.Visible = True
    Exit Sub

Err_line:
    If Err Then
        MsgBox "Report Yearly clain bill: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line
End Sub



Private Sub ReportGeneralLedger()

' Declare variables...
Dim Lret As Long
Dim SqlStr As String
Dim rptRS As Recordset
Dim PrevDate As String
Dim TotAmt As Currency

' Setup error handler.
On Error GoTo Err_line

RaiseEvent Processing("Reading & Verifying the data ", 0)
Dim Deposit As Boolean

Deposit = IIf(m_ReportType = repBkccDepGLedger, True, False)

SqlStr = "Select 'PRINCIPAL',sum(Amount) as TotalAmount," & _
    " TransDate, TransType From BKCCTrans Where" & _
    " TransDate >= #" & m_FromDate & "# " & _
    " AND TransDate <= #" & m_ToDate & "# And Deposit = " & Deposit
If m_FarmerType Then
    SqlStr = SqlStr & " And LoanID In (Select LoanId From BkccMaster " & _
            "Where farmerType = " & m_FarmerType & ") "
End If

SqlStr = SqlStr & " Group By TransDate, TransType"
gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"
SqlStr = ""
If gDbTrans.Fetch(rptRS, adOpenDynamic) <= 0 Then GoTo Exit_Line

' Populate the record set.
If rptRS.EOF Then Exit Sub
rptRS.MoveLast
rptRS.MoveFirst

Dim TotalWithDraw As Currency
Dim TotalDeposit As Currency
Dim SubWithdraw As Currency
Dim SubDeposit As Currency
Dim OpeningBalance As Currency

' Initialize the grid.
Call InitGrid
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Aligning the data ", 0)
' Fill the rows
Dim transType  As wisTransactionTypes
Dim TransDate As Date
Dim SlNo As Long
Dim PRINTTotal As Boolean
Dim AccClass As New clsAccTrans
Dim rowno As Integer, colno As Integer


If m_ReportType = repBkccDepGLedger Then
    OpeningBalance = AccClass.GetOpBalance( _
        GetIndexHeadID(GetResourceString(229, 43)), m_FromDate)
Else
    OpeningBalance = AccClass.GetOpBalance( _
        GetIndexHeadID(GetResourceString(229, 58)), m_FromDate)
End If
Set AccClass = Nothing

TransDate = rptRS("TransDate")
PRINTTotal = False
With grd
    .Row = .FixedRows
    .Col = 1: .Text = GetResourceString(284)
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .CellFontBold = True
    
    rowno = .Row: colno = .Col
End With

Do While Not rptRS.EOF
    With grd
        ' Set the row.
        If rptRS("Transdate") <> TransDate Then
            PRINTTotal = True
            SlNo = SlNo + 1
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = SlNo
            colno = 1: .TextMatrix(rowno, colno) = GetIndianDate(TransDate)
            colno = 2: .TextMatrix(rowno, colno) = FormatCurrency(OpeningBalance)
            colno = 3: .TextMatrix(rowno, colno) = FormatCurrency(SubDeposit)
            colno = 4: .TextMatrix(rowno, colno) = FormatCurrency(SubWithdraw)
            
            OpeningBalance = OpeningBalance - SubDeposit + SubWithdraw
            colno = 5: .TextMatrix(rowno, colno) = FormatCurrency(OpeningBalance)
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            TransDate = rptRS("TransDate")
            
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0

        End If
        transType = rptRS("TransType")
        If transType = wWithdraw Or transType = wContraWithdraw Then
            SubWithdraw = SubWithdraw + FormatField(rptRS("TotalAmount"))
        Else
            SubDeposit = SubDeposit + FormatField(rptRS("TotalAmount"))
        End If
    End With

    
    DoEvents
    Me.Refresh
    If gCancel Then rptRS.MoveLast
    RaiseEvent Processing("Writing the data ", rptRS.AbsolutePosition / rptRS.recordCount)
    ' Move to next row.
    rptRS.MoveNext
    
Loop

'Now Print the last day's receipt & Payment
With grd
    .Row = rowno
    SlNo = SlNo + 1
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 0: .Text = SlNo
    .Col = 1: .Text = GetIndianDate(TransDate)
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .Col = 3: .Text = FormatCurrency(SubDeposit)
    .Col = 4: .Text = FormatCurrency(SubWithdraw)
    
    OpeningBalance = OpeningBalance - SubDeposit + SubWithdraw
    .Col = 5: .Text = FormatCurrency(OpeningBalance)
    
    TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
    TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    '.Col = 0: .Text = SlNo
    .Col = 4: .Text = GetResourceString(285)
    .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(OpeningBalance)
    .CellFontBold = True
    
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        .Col = 1: .Text = GetResourceString(286) 'Grand total
        .CellFontBold = True
        .Col = 3: .Text = FormatCurrency(TotalDeposit)
        .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(TotalWithDraw)
        .CellFontBold = True
    End If
End With

grd.Visible = True
Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportGenLedger: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
 Resume
    GoTo Exit_Line
End Sub

Private Sub ReportLoanDetails()

Dim Lret As Long
Dim rptRS As Recordset
Dim Guarantor As Long
Dim TotalBalance As Currency
Dim TotalIssue As Currency
Dim SqlStr As String
Me.MousePointer = vbHourglass

RaiseEvent Processing("Reading the records ", 0)
SqlStr = "SELECT Max(TransID) AS MaxTransID, LoanID" & _
    " FROM BkccTrans WHERE TransDate <= #" & m_ToDate & "#" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr

If Not gDbTrans.CreateView("QryLoanTrans") Then Exit Sub

'Build query.
SqlStr = "SELECT A.LoanID,AccNum,IssueDate,SanctionAmount,Balance, " _
    & " Name As CustName, Caste,Place, Guarantor1, Guarantor2, MemberNum " _
    & " FROM ((BKCCMaster A INNER JOIN qryMemName C ON A.MemID=C.MemID)" _
    & " INNER JOIN BKCCTrans B ON A.LoanID = B.LoanID)" _
    & " INNER JOIN QryLoanTrans D ON B.TransID = D.MaxTransID AND B.LoanID = D.LoanID "
    
If m_ReportType = repBkccDepHolder Then
    SqlStr = SqlStr & " WHERE Balance < 0 "
Else
    SqlStr = SqlStr & " WHERE Balance > 0 "
End If

If Trim$(m_Place) <> "" Then _
        SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
        SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then _
    SqlStr = SqlStr & " And Gender = " & m_Gender
If m_FarmerType Then _
    SqlStr = SqlStr & " And FarmerType = " & m_FarmerType

gDbTrans.SqlStmt = SqlStr
SqlStr = ""
Lret = gDbTrans.Fetch(rptRS, adOpenDynamic)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

rptRS.MoveLast
rptRS.MoveFirst

'Raise event to access frmcancel.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

' Initialize the grid.
With grd
End With

Call InitGrid

Dim SlNo As Integer
Dim TotalAmount As Currency
Dim RegularInterest As Currency
Dim PenalInterest As Currency
Dim TillDate As String
Dim TotalRegularInterest As Currency
Dim TotalPenalInterest As Currency
Dim rowno As Integer, colno As Integer

rowno = grd.Row
TotalAmount = 0
' Fill the rows
Do While Not rptRS.EOF
    
    RegularInterest = BKCCRegularInterest(m_ToDate, rptRS("LoanID"))
    TotalRegularInterest = TotalRegularInterest + (Abs(RegularInterest) \ 1)
    PenalInterest = BKCCPenalInterest(m_ToDate, rptRS("LoanID"))
    TotalPenalInterest = TotalPenalInterest + (PenalInterest \ 1)
    
    With grd
        ' Set the row.
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        SlNo = SlNo + 1
        colno = 0
        .TextMatrix(rowno, colno) = SlNo
    
        ' Fill the loanid.
        colno = 1
        .TextMatrix(rowno, colno) = rptRS("AccNum")
        colno = 2
        .TextMatrix(rowno, colno) = rptRS("MemberNum")
    
       ' Fill the loan holder name.
        colno = 3
        .TextMatrix(rowno, colno) = rptRS("CustName")
        colno = 4
        .TextMatrix(rowno, colno) = FormatField(rptRS("IssueDate"))
        colno = 5
        .TextMatrix(rowno, colno) = rptRS("Caste")
        colno = 6
        .TextMatrix(rowno, colno) = rptRS("Place")
    
    ' Fill the loan issue date.
        ' Fill the loan amount.
        colno = 7
        .TextMatrix(rowno, colno) = FormatCurrency(Abs(rptRS("Balance"))): .CellAlignment = 7
        TotalAmount = TotalAmount + Val(Abs(rptRS("Balance")))
        colno = 8
        .TextMatrix(rowno, colno) = FormatCurrency(Abs(RegularInterest) \ 1): .CellAlignment = 7
        colno = 9
        .TextMatrix(rowno, colno) = FormatCurrency(PenalInterest \ 1): .CellAlignment = 7
        If m_ReportType = repBkccDepHolder Then
            colno = 9
            .TextMatrix(rowno, colno) = FormatCurrency(Abs(rptRS("Balance")) + RegularInterest)
        Else
            colno = 10
            .TextMatrix(rowno, colno) = FormatCurrency(FormatField(rptRS("Balance")) + RegularInterest + _
                PenalInterest): .CellAlignment = 7
        End If
        TotalBalance = TotalBalance + Val(.Text)
    End With
    
nextRecord:
    ' Move to next row.
    If gCancel Then rptRS.MoveLast

    rptRS.MoveNext
    DoEvents
    RaiseEvent Processing("Writing the data ", rptRS.AbsolutePosition / rptRS.recordCount)
Loop

With grd
    .Row = rowno
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    '.Col = 7: .Text = FormatCurrency(TotalIssue): .CellFontBold = True: .CellAlignment = 7
    .Col = 7: .Text = FormatCurrency(TotalAmount): .CellFontBold = True: .CellAlignment = 7
    .Col = 8: .Text = FormatCurrency(TotalRegularInterest): .CellFontBold = True: .CellAlignment = 7
    .Col = 9: .Text = FormatCurrency(TotalPenalInterest): .CellFontBold = True: .CellAlignment = 7
    If m_ReportType = repBkccLoanHolder Then
    .Col = 10: .Text = FormatCurrency(TotalBalance): .CellFontBold = True: .CellAlignment = 7 ' Display the grid.
    End If
    .Visible = True
End With

Me.Caption = "INDEX-2000  [List of loans issued...]"

Exit_Line:
    Set rptRS = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
    GoTo Exit_Line

End Sub
Private Sub ReportOverdueLoans()

Dim Lret As Long
Dim rstPrev As Recordset
Dim rstPresent As Recordset
Dim rstTrans As Recordset
Dim OverDueLoan As Boolean

' Setup error handler.
On Error GoTo Err_line
Me.MousePointer = vbHourglass

Dim transType As wisTransactionTypes
'
RaiseEvent Processing("Reading & Verifying the data ", 0)
'Build the report query. to Get Instalment Loan
'Get the Balance of each account before one year
'If tha amount repaid is more or equal to tha balance
'then this deposit is not OD

'Create the view for max TransID
gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID,LoanID" & _
            " FROM BKCCTrans Where TransDate < #" & DateAdd("d", 2, DateAdd("yyyy", -1, m_ToDate)) & "#" & _
            " GROUP BY LoanID"
gDbTrans.CreateView ("qryMaxID")

transType = wWithdraw
'Get the Balance Before On year
gDbTrans.SqlStmt = "SELECT A.Loanid, AccNum, Balance, Amount, TransDate, " _
        & " TransID, Name as CustName,Guarantor1, Guarantor2,MemberNum " _
        & " FROM ((BkccMaster A INNER JOIN BKCCTrans B ON A.LoanID= B.LoanID)" _
        & " INNER JOIN QryMemName C ON A.MemID = C.MemID) " _
        & " INNER JOIN qryMaxID D ON B.TransID = D.MaxTransID AND B.LoanID = D.LoanID " _
        & " WHERE Balance > 0 "

If Trim$(m_Place) <> "" Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)

If m_Gender Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
                " And FarmerType = " & m_FarmerType
If m_ReportOrder = wisByName Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER By IsciName,TransID"
Else
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ORDER By val(AccNum),TransID"
End If

Lret = gDbTrans.Fetch(rstPrev, adOpenDynamic)
If Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
   GoTo Exit_Line
End If

'Create Temporary view to Get the Max TransID
gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID, LoanID FROM BKCCTrans" _
            & " Where TransDate <= #" & m_ToDate & "# GROUP BY LoanID"
gDbTrans.CreateView ("qryMaxTransID")

'Get the Present Balance
gDbTrans.SqlStmt = "SELECT A.Loanid, Balance FROM " _
        & " (BKCCTrans A INNER JOIN BKCCMaster B ON A.LoanID= B.LoanID) " _
        & " INNER JOIN qryMaxTransID C ON A.TransID = C.MaxTransID" _
        & " AND A.LoanID = C.LoanID WHERE Balance > 0 "

If gDbTrans.Fetch(rstPresent, adOpenDynamic) < 1 Then Set rstPresent = Nothing

'Get the TransCtion Details since a year
gDbTrans.SqlStmt = "SELECT Loanid, Sum(Amount) as RepaidAmount FROM " & _
        " BKCCTrans WHERE TransDate > #" & DateAdd("yyyy", -1, m_ToDate) & "# " & _
        " And TransDate <= #" & m_ToDate & "#" & _
        " And (TransType = " & wDeposit & " Or TransType = " & wContraDeposit & ")" & _
        " Group BY LoanID "

If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Set rstTrans = Nothing
    

' Initialize the grid.
RaiseEvent Initialise(0, rstPrev.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

Dim SlNo As Integer
Dim Balance As Currency
Dim TotalAmount As Currency
Dim RegularInterest As Currency
Dim PenalInterest As Currency
Dim TotalRegularInterest As Currency
Dim TotalPenalInterest As Currency
Dim PaidAmount As Currency
Dim ODAmount As Currency
Dim GuarantorId As Long
Dim lCust As New clsCustReg
Dim rowno As Integer, colno As Integer

' Fill the rows
Call InitGrid
        
rowno = grd.Row

Do While Not rstPrev.EOF
    Balance = FormatField(rstPrev("Balance"))
    Debug.Assert rstPrev("LoanId") <> 1875
    PaidAmount = 0
    If Not rstTrans Is Nothing Then
        rstTrans.MoveFirst
        rstTrans.Find "LoanID = " & rstPrev("LoaniD")
        If Not rstTrans.EOF Then PaidAmount = FormatField(rstTrans("RepaidAmount"))
    End If
    ODAmount = Balance - PaidAmount
    If ODAmount <= 0 Then GoTo nextRecord
    RegularInterest = (BKCCRegularInterest(m_ToDate, rstPrev("LoanID"))) \ 1
    PenalInterest = (BKCCPenalInterest(m_ToDate, rstPrev("LoanID"))) \ 1
    
    If Not rstPresent Is Nothing Then
        rstPresent.MoveFirst
        rstPresent.Find "LoanID = " & rstPrev("LoaniD")
        If Not rstPresent.EOF Then Balance = FormatField(rstPresent("Balance"))
    End If
    
    'Set the row.
    With grd
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: SlNo = SlNo + 1
        colno = 0: .TextMatrix(rowno, colno) = SlNo
        ' Fill the loanid.
        colno = 1: .TextMatrix(rowno, colno) = rstPrev("AccNum")
        colno = 2: .TextMatrix(rowno, colno) = rstPrev("MemberNum")
        ' Fill the loan holder name.
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rstPrev("CustName"))
        colno = 4: .TextMatrix(rowno, colno) = GetIndianDate(DateAdd("yyyy", 1, rstPrev("TransDate")))
        colno = 5: .TextMatrix(rowno, colno) = GetIndianDate(rstPrev("TransDate"))
        colno = 6: .TextMatrix(rowno, colno) = FormatCurrency(Balance)
        colno = 7: .TextMatrix(rowno, colno) = FormatCurrency(ODAmount)
        TotalAmount = TotalAmount + Val(.TextMatrix(rowno, colno))
   
        colno = 8: .TextMatrix(rowno, colno) = FormatCurrency(RegularInterest): .CellAlignment = 7
        TotalRegularInterest = TotalRegularInterest + RegularInterest
        colno = 9: .TextMatrix(rowno, colno) = FormatCurrency(PenalInterest): .CellAlignment = 7
        TotalPenalInterest = TotalPenalInterest + PenalInterest
        TotalPenalInterest = TotalPenalInterest + PenalInterest
        colno = 10: .TextMatrix(rowno, colno) = FormatCurrency(ODAmount + RegularInterest + PenalInterest)
        .CellAlignment = 7
        GuarantorId = FormatField(rstPrev("Guarantor1"))
        If GuarantorId Then colno = 11: .TextMatrix(rowno, colno) = lCust.CustomerName(GuarantorId)
        GuarantorId = FormatField(rstPrev("Guarantor2"))
        If GuarantorId Then colno = 12: .TextMatrix(rowno, colno) = lCust.CustomerName(GuarantorId)
        
    End With

nextRecord:
    
    DoEvents
    If gCancel Then rstPrev.MoveLast
    
    RaiseEvent Processing("Writing the data ", rstPrev.AbsolutePosition / rstPrev.recordCount)
    
    ' Move to next row.
    rstPrev.MoveNext

Loop
        
With grd
    .Row = rowno
    If .Rows <= .Row + 1 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
    .Col = 7: .Text = FormatCurrency(TotalAmount)
        .CellFontBold = True: .CellAlignment = 7
    .Col = 8: .Text = FormatCurrency(TotalRegularInterest)
        .CellFontBold = True: .CellAlignment = 7
    .Col = 9: .Text = FormatCurrency(TotalPenalInterest)
        .CellFontBold = True: .CellAlignment = 7
    ' Display the grid.
    .Visible = True
End With

Me.Caption = "INDEX-2000  [over due loans]"

Exit_Line:
    Set rstPrev = Nothing
    Set lCust = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
       MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    GoTo Exit_Line

End Sub

Private Sub ReportBalance()
Dim SqlStmt As String
Dim rst As Recordset
Dim TotalBalance As Currency
Dim SlNo As Long

'raise event to update frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create the query for Max TransID
gDbTrans.SqlStmt = "Select Max(TransId) As MaxTransID,LoanID" & _
        " From BKCCTrans Where TransDate <= #" & m_ToDate & "#" & _
        " Group By LoanId"
Call gDbTrans.CreateView("qryBkCCMaxTransID")

SlNo = IIf(m_ReportType = repBkccDepBalance, -1, 1)
SqlStmt = "Select A.LoanId,A.AccNum,B.MemberNum,Balance * " & SlNo & " As Balance," & _
    " Name From ((BKCCMaster A Inner Join QryMemName B On A.MemID = B.MemID ) " & _
    " Inner Join qryBkCcMaxTransID C On A.LoanId = C.LoanId) Inner Join  BKCCTrans D" & _
        " On C.MaxTransId = D.TransID ANd C.LoanID = D.LoanID"
SlNo = 0

Dim sqlClause As String
sqlClause = ""
If m_ReportType = repBkccDepBalance Then
    sqlClause = sqlClause & " And  Balance < " & -1 * m_FromAmt
    If m_ToAmt <> 0 Then sqlClause = sqlClause & " And Balance > " & -1 * m_ToAmt
Else
    sqlClause = sqlClause & " And  Balance > " & (m_FromAmt - 0.01)
    If m_ToAmt <> 0 Then sqlClause = sqlClause & " And Balance <= " & m_ToAmt
End If

If Trim$(m_Place) <> "" Then sqlClause = sqlClause & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then sqlClause = sqlClause & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then sqlClause = sqlClause & " And Gender = " & m_Gender

If m_FarmerType <> NoFarmer Then sqlClause = sqlClause & " And FarmerType = " & m_FarmerType

sqlClause = Trim$(sqlClause)
If Len(sqlClause) Then sqlClause = " WHERE " & Mid(sqlClause, 4)

If m_ReportOrder = wisByName Then
    gDbTrans.SqlStmt = SqlStmt & sqlClause & " ORDER BY IsciName"
Else
    gDbTrans.SqlStmt = SqlStmt & sqlClause & " ORDER BY Val(AccNum)"
End If

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
Call InitGrid

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

SlNo = 1
Dim rowno As Integer, colno As Integer


While Not rst.EOF
    grd.ColData(0) = 1
    If FormatField(rst("Balance")) <> 0 Then
    With grd
        If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
        rowno = SlNo
        colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00"): .CellAlignment = 1
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("AccNum")): .CellAlignment = 4
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("MemberNum")): .CellAlignment = 4
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rst("Name")): .CellAlignment = 1
        colno = 4: .TextMatrix(rowno, colno) = FormatField(rst("Balance")): .CellAlignment = 7
        TotalBalance = TotalBalance + Val(.TextMatrix(rowno, colno))
    End With
    SlNo = SlNo + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / (rst.recordCount + 1))
    End If
    rst.MoveNext
Wend

With grd
    .Row = rowno
    If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(52, 42) ' "Total Balance"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub

Private Function ReportCustomerTransaction() As Boolean

ReportCustomerTransaction = False

' Declare variables...
Dim Lret As Long
Dim rstTrans As Recordset
Dim RstCust As Recordset

Dim PrinSql  As String
Dim IntSql As String

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Build the report query.
PrinSql = "Select 'PRINCIPAL',sum(Amount) as RegInt,'0' as PenalInt, " & _
        "LoanID,TransType From BKCCTrans Where " & _
        "TransDate >= #" & m_FromDate & "# And TransDate <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then PrinSql = PrinSql & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then PrinSql = PrinSql & " AND Amount <= " & m_ToAmt
PrinSql = PrinSql & " Group by LoanID,TransType "

IntSql = "Select 'INTEREST',sum(IntAmount) as RegInt,sum(PenalIntAmount) as PenalInt, " & _
        " LoanID,TransType From BKCCIntTrans Where " & _
        " TransDate >= #" & m_FromDate & "# And TransDate <= #" & m_ToDate & "# " & _
        " Group by LoanID,TransType "

' Finally, add the sorting clause.
gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & _
    " Order by LoanID"
    
' Execute the query...
Lret = gDbTrans.Fetch(rstTrans, adOpenStatic)
If Lret <= 0 Then GoTo Exit_Line

Dim sqlSupport  As String
sqlSupport = ""
If Trim$(m_Place) <> "" Then _
    sqlSupport = sqlSupport & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    sqlSupport = sqlSupport & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then _
    sqlSupport = sqlSupport & " And Gender = " & m_Gender
If Len(sqlSupport) > 0 Then
    'Make View on the basis of the above condition
    'and use this view in while fetching details
    gDbTrans.SqlStmt = "Select Title +' '+FirstNAme+' '" & _
            " +MiddleName+' '+LAstNAme as Name From NameTab" & _
            " WHERE " & Mid(Trim$(sqlSupport), 4)
    'Create query
    gDbTrans.CreateView ("qryName1")
    sqlSupport = ""
End If

grd.Clear
grd.Cols = 5
grd.FixedCols = 1
grd.Rows = 20

Dim TotalBalance As Currency

Lret = rstTrans.recordCount + 2
RaiseEvent Initialise(0, Lret)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

' Initialize the grid.
Call InitGrid

Dim transType As wisTransactionTypes
Dim SlNo As Long
Dim TransDate As Date
Dim TransID As Long
Dim SubWithdraw As Currency
Dim SubDeposit As Currency
Dim SubInterest As Currency
Dim TotalWithDraw As Currency
Dim TotalDeposit As Currency
Dim TotalInterest As Currency
Dim SubPenal As Currency
Dim TotalPenal As Currency

Dim Amount As Currency
Dim Balance As Currency
Dim SubBalance As Currency
' Fill the rows
grd.Rows = 10
grd.Row = grd.FixedRows - 1

Dim LoanID As Long
Dim loopCount As Integer
Dim rowno As Integer, colno As Integer

rowno = grd.Row

'SlNo = 1
LoanID = rstTrans("LoanID")
Do While Not rstTrans.EOF
    With grd
        If LoanID <> rstTrans("LoanID") Then
            SlNo = SlNo + 1
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1
            gDbTrans.SqlStmt = "Select Name as CustName,AccNum,MemberNum" & _
                " From BKCCMaster A Inner Join QryMemName B" & _
                " On A.MemID = B.MemID " & _
                " Where LoanId = " & LoanID
            If gDbTrans.Fetch(RstCust, adOpenDynamic) > 0 Then
            
            colno = 0: .TextMatrix(rowno, colno) = SlNo
            colno = 1: .TextMatrix(rowno, colno) = RstCust("AccNum")
            colno = 2: .TextMatrix(rowno, colno) = RstCust("MemberNum")
            ' Fill the loan holder name.
            colno = 3: .TextMatrix(rowno, colno) = Trim$(FormatField(RstCust("CustName")))
            End If
            If SubWithdraw Then colno = 4: .TextMatrix(rowno, colno) = FormatCurrency(SubWithdraw)
            If SubDeposit Then colno = 5: .TextMatrix(rowno, colno) = FormatCurrency(SubDeposit)
            If SubInterest Then colno = 6: .TextMatrix(rowno, colno) = FormatCurrency(SubInterest)
            If SubPenal Then colno = 7: .TextMatrix(rowno, colno) = FormatCurrency(SubPenal)
            
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
            TotalInterest = TotalInterest + SubInterest: SubInterest = 0
            TotalPenal = TotalPenal + SubPenal: SubPenal = 0
            
            LoanID = rstTrans("LOanID")
        End If
        
        transType = rstTrans("TransType")
        Amount = FormatField(rstTrans("RegInt"))
        Balance = FormatField(rstTrans("PenalInt"))
        
        If rstTrans(0) = "PRINCIPAL" Then  'If it is principal details
            If transType = wWithdraw Or transType = wContraWithdraw Then
                SubWithdraw = SubWithdraw + Amount
            Else
                SubDeposit = SubDeposit + Amount
            End If
        Else
            If transType = wDeposit Or transType = wContraDeposit Then
                SubInterest = SubInterest + Amount
                SubPenal = SubPenal + Balance
            End If
        End If
   End With

nextRecord:
    loopCount = loopCount + 1
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", loopCount / Lret)
    DoEvents
    If gCancel Then rstTrans.MoveLast
    
    rstTrans.MoveNext
Loop
    
    With grd
        SlNo = SlNo + 1
        If .Rows <= rowno + 2 Then .Rows = .Rows + 2
        rowno = rowno + 1
        gDbTrans.SqlStmt = "Select Name as CustName,AccNum,MemberNum" & _
            " From BKCCMaster A Inner Join QryMemName B" & _
            " on A.MemID = B.MemID" & _
            " Where LoanId = " & LoanID
        If gDbTrans.Fetch(RstCust, adOpenDynamic) > 0 Then
        
        
        colno = 0: .TextMatrix(rowno, colno) = SlNo
        colno = 1: .TextMatrix(rowno, colno) = RstCust("AccNum")
        colno = 2: .TextMatrix(rowno, colno) = RstCust("MemberNum")
        ' Fill the loan holder name.
        colno = 3: .TextMatrix(rowno, colno) = Trim$(FormatField(RstCust("CustName")))
        End If
        If SubWithdraw Then colno = 4: .TextMatrix(rowno, colno) = FormatCurrency(SubWithdraw)
        If SubDeposit Then colno = 5: .TextMatrix(rowno, colno) = FormatCurrency(SubDeposit)
        If SubInterest Then colno = 6: .TextMatrix(rowno, colno) = FormatCurrency(SubInterest)
        If SubPenal Then colno = 7: .TextMatrix(rowno, colno) = FormatCurrency(SubPenal)
        
        TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
        TotalInterest = TotalInterest + SubInterest: SubInterest = 0
        TotalPenal = TotalPenal + SubPenal: SubPenal = 0
            
        'Put GrandTotal
        .Rows = rowno + 3
        rowno = rowno + 2
        .Row = rowno
        .Col = 3: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalDeposit): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalInterest): .CellFontBold = True
        .Col = 7: .Text = FormatCurrency(TotalPenal): .CellFontBold = True
                   
        If TotalInterest = 0 Then .ColWidth(5) = 5
        If TotalPenal = 0 Then .ColWidth(6) = 5
        
        'Added on 29/11/00(For Allignment)
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 8
        .ColAlignment(5) = 7
        .ColAlignment(6) = 8
        .ColAlignment(7) = 9
    End With

' Display the grid.
grd.Visible = True
Me.Caption = "LOANS  [customer Transaction ...]"

lblReportTitle.Caption = GetResourceString(205) & " " & _
    GetResourceString(28) & " " & _
    GetFromDateString(m_FromIndianDate, m_ToIndianDate)

ReportCustomerTransaction = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportCustomerTransaction: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        
    End If
'Resume
    GoTo Exit_Line

End Function

Private Sub ReportSanctionedLoans()
  
Dim rstMaster As Recordset
Dim MemID As Integer
Dim I As Integer
    
RaiseEvent Processing("Reading & Verifying the data ", 0)
 
gDbTrans.SqlStmt = "SELECT Loanid, AccNum, SanctionAmount," & _
    " Name as CustName From BKCCMaster A Inner Join" & _
    " QryMemName B ON A.MemID = B.MemID " & _
    " WHERE IssueDate >= #" & m_FromDate & "# AND IssueDate <= #" & m_ToDate & "# "

If Trim$(m_Place) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)

If m_Gender Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And Gender = " & m_Gender
'Select the Farmer Type
If m_FarmerType <> NoFarmer Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And FarmerType = " & m_FarmerType
  
If gDbTrans.Fetch(rstMaster, adOpenDynamic) < 1 Then Exit Sub
 
'Set RstMaster = gDBTrans.Rst.Clone
Call InitGrid
Dim colno As Integer

For I = 1 To rstMaster.recordCount
    If rstMaster.EOF Then Exit For
    With grd
        '.Row = I
        
        .TextMatrix(I, 0) = CStr(I)
        .TextMatrix(I, 1) = rstMaster("AccNum")
        .TextMatrix(I, 2) = rstMaster("MemberNum")
        .TextMatrix(I, 3) = FormatField(rstMaster("CustName"))
        .TextMatrix(I, 4) = FormatField(rstMaster("SanctionAmount"))
        If .Rows = I + 1 Then .Rows = .Rows + 1
    End With
    DoEvents
    If gCancel Then rstMaster.MoveLast
    RaiseEvent Processing("Writing the data ...", I / rstMaster.recordCount)
    
    rstMaster.MoveNext

Next I

grd.Visible = True
ErrLine:
  If Err Then
      MsgBox Err.Description
      'Resume
  End If
End Sub

Private Sub ReportGuarantors()
Dim SqlStmt As String

'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create view from Max TransID
gDbTrans.SqlStmt = "SELECT Max(TransID) as MaxTransID, LoanID " & _
        " From BkccTrans D WHERE TransDate <= #" & m_ToDate & "#" & _
        " GROUP BY LoanID"
'Create View
gDbTrans.CreateView ("qryBKCCMaxTransID")
     
gDbTrans.SqlStmt = "SELECT A.LoanID, A.Balance FROM BKCCTrans A " & _
        " Inner Join qryBKCCMaxTransID B " & _
        " ON A.LOanID=B.LoanID AND A.TransID = B.MaxTransID " & _
        " Where Balance > 0"
gDbTrans.CreateView ("qryBKCCLoanBalance")

gDbTrans.SqlStmt = "SELECT AccNum, Guarantor1, Guarantor2, A.LoanId, Name,Place,Caste,Gender " & _
    " ,FarmerType  From BKCCMaster A INNER JOIN QryName B" & _
    " ON B.CustomerId = A.CustomerId " & _
    " WHERE (Guarantor1 > 0 OR  Guarantor2 >0 ) "
gDbTrans.CreateView ("qryBKCCCustName")

SqlStmt = "SELECT AccNum, Guarantor1, Guarantor2, A.LoanId, Name, " & _
    " Balance From qryBKCCCustName A INNER JOIN qryBKCCLoanBalance B" & _
    " ON A.LoanID = B.LoanID "

Dim SqlStmt1 As String
SqlStmt1 = "SELECT AccNum, Guarantor1, Guarantor2, A.LoanId, Name, " & _
    " Balance From ((BKCCMaster A INNER JOIN QryName B" & _
    " ON B.CustomerId = A.CustomerId) INNER JOIN BKCCTrans C" & _
    " ON C.LoanID = A.LoanID) INNER JOIN qryBkccMaxTransID D " & _
    " ON A.LOanID=d.LoanID AND C.TransID = D.MaxTransID" & _
    " WHERE (Guarantor1 > 0 OR  Guarantor2 >0 ) AND Balance > 0 "
     
If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_FarmerType Then SqlStmt = SqlStmt & " And FarmerType = " & m_FarmerType

If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & " ORDER BY IsciName"
Else
    SqlStmt = SqlStmt & " ORDER BY Val(AccNum)"
End If

gDbTrans.SqlStmt = SqlStmt: SqlStmt = ""
 
Dim rst As Recordset

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

Dim SlNo As Long
Dim CustClass As New clsCustReg
    Call InitGrid
    
    RaiseEvent Initialise(0, rst.recordCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    If Trim$(m_Place) <> "" Then
         grd.Cols = grd.Cols + 1
         grd.Col = grd.Cols - 1
         grd.Text = GetResourceString(270)
    End If
    If Trim$(m_Caste) <> "" Then
         grd.Cols = grd.Cols + 1
         grd.Col = grd.Cols - 1
         grd.Text = GetResourceString(100)
    End If
   
   Dim FirstGuarantorID As Long
   Dim SecondGuarantorID As Long
   Dim rowno As Integer, colno As Integer
   
   rowno = grd.Row
    SlNo = 1
    While Not rst.EOF
        With grd
            If .Rows <= .Rows + 2 Then .Rows = .Rows + 1
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00"): .CellAlignment = 1
            colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("AccNum")): .CellAlignment = 4
            colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("Name")): .CellAlignment = 1
            colno = 3: .TextMatrix(rowno, colno) = FormatField(rst("Balance")): .CellAlignment = 1
            FirstGuarantorID = FormatField(rst("Guarantor1"))
            SecondGuarantorID = FormatField(rst("Guarantor2"))
            If FirstGuarantorID Then _
                colno = 4: .TextMatrix(rowno, colno) = CustClass.CustomerName(FirstGuarantorID)
            If SecondGuarantorID Then _
                colno = 5: .TextMatrix(rowno, colno) = CustClass.CustomerName(SecondGuarantorID)
        End With
        SlNo = SlNo + 1
        
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.recordCount)
        rst.MoveNext
        
    Wend

ExitLine:
    
End Sub

Private Sub Form_Resize()

Screen.MousePointer = vbDefault
On Error Resume Next
lblReportTitle.Top = 0
lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
grd.Left = 0
grd.Top = lblReportTitle.Top + lblReportTitle.Height
grd.Width = Me.Width - 150
fra.Top = Me.ScaleHeight - fra.Height
fra.Left = Me.Width - fra.Width
grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height
cmdOk.Left = fra.Width - cmdOk.Width - (cmdOk.Width / 4)
cmdPrint.Left = cmdOk.Left - cmdPrint.Width - (cmdPrint.Width / 4)
cmdWeb.Top = cmdPrint.Top
cmdWeb.Left = cmdPrint.Left - cmdWeb.Width - (cmdPrint.Width / 4)

Dim I As Integer
Dim ColWid As Single
For I = 0 To grd.Cols - 1
    ColWid = GetSetting(App.EXEName, "BKCCReport" & m_ReportType, _
        "ColWidth" & I, 1 / grd.Cols) * grd.Width
    If ColWid < 10 Or ColWid > grd.Width * 0.9 Then ColWid = grd.Width / grd.Cols
    grd.ColWidth(I) = ColWid
Next I

End Sub

Private Sub Form_Unload(cancel As Integer)
    Set frmBkccReport = Nothing
End Sub

Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "BKCCReport" & m_ReportType, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount
End Sub

Private Sub m_frmCancel_CancelClicked()
    m_grdPrint.CancelProcess
End Sub

Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
    
    m_TotalCount = MaxCount
    If m_frmCancel Is Nothing Then Set m_frmCancel = New frmCancel
    m_frmCancel.PicStatus.Visible = True
    m_frmCancel.PicStatus.ZOrder 0

End Sub

Private Sub m_grdPrint_Message(strMessage As String)
    m_frmCancel.lblMessage = strMessage
End Sub

Private Sub m_grdPrint_ProcessCount(count As Long)

On Error Resume Next

    If (count / m_TotalCount) > 0.95 Then
        Unload m_frmCancel
        Exit Sub
    End If
    UpdateStatus m_frmCancel.PicStatus, count / m_TotalCount
    Err.Clear

End Sub

