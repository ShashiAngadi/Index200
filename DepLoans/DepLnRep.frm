VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDepLoanReport 
   Caption         =   "Loan Reports .."
   ClientHeight    =   5850
   ClientLeft      =   1485
   ClientTop       =   1770
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1320
      TabIndex        =   1
      Top             =   5070
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   540
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   1680
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
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
      Top             =   390
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
      Top             =   0
      Width           =   1635
   End
End
Attribute VB_Name = "frmDepLoanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public ParentForm As frmDepLoan

Private m_ReportType As wis_DepLoanReports
Private m_FromIndianDate As String
Private m_ToIndianDate As String
Private m_FromDate As Date
Private m_ToDate As Date

Private m_FromAmount As Currency
Private m_ToAmount As Currency

Private m_Caste As String
Private m_Place As String
Private m_Gender As wis_Gender

Private m_DepositType As wis_DepositType
Private m_Order As wis_ReportOrder

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)



Private Sub ReportLoanMonthlyBalance()

Dim SqlStr As String
Dim rstMain As ADODB.Recordset
Dim fromDate As Date
Dim toDate As Date
Dim sqlDepType As String

'NowGet the LAst days of the given dates
toDate = GetSysLastDate(m_ToDate)
fromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(43, 58, 67, 42) & " " & _
                GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))

sqlDepType = ""
If m_DepositType Then
    sqlDepType = " AND A.LoanID In (Select LoanId From DepositLoanMaster " & _
                " WHERE DepositType = " & m_DepositType & ")"
End If

SqlStr = "Select AccNum, LoanId, Name as CustName " & _
    " FROM DepositLoanMaster A, QryName B WHERE B.CustomerId = A.CustomerID "
SqlStr = SqlStr & " And LoanID In (Select Distinct LoanId From DepositLoanTrans " & _
    " Where TransDate Between #" & fromDate & "# AND #" & toDate & "# And Balance > 0) "

'Select the Farmer Type
If m_DepositType Then SqlStr = SqlStr & " And DepositType = " & m_DepositType
SqlStr = SqlStr & " ORDER BY " & IIf(m_Order = wisByName, "IsciName", "Val(AccNum)")

'First Fetch the MAster Records
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Sub

Dim SlNo As Long
Dim count As Long
Dim maxCount As Long

count = DateDiff("m", fromDate, toDate)
count = count + 2
maxCount = rstMain.recordCount * count

RaiseEvent Initialise(0, maxCount)

With grd
    .Clear
    .Cols = 3
    .FixedCols = 2: .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '" Slno"
    .Col = 1: .Text = GetResourceString(36, 60)  '" Loan Id "
    .Col = 2: .Text = GetResourceString(35) '"Loan Holder" '
End With
Dim rowno As Long, colno As Byte
rowno = grd.Row

count = 0
While Not rstMain.EOF
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        .TextMatrix(rowno, 0) = rowno
        .TextMatrix(rowno, 1) = FormatField(rstMain("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rstMain("CustName"))
    End With
    count = count + 1
    RaiseEvent Processing("Collecting Information ", count / maxCount)
    rstMain.MoveNext
Wend
With grd
    If .Rows <= rowno + 2 Then .Rows = .Rows + 1
    rowno = rowno + 1
    If .Rows <= rowno + 2 Then .Rows = .Rows + 1
    rowno = rowno + 1
    .Row = rowno
    
    .Col = 2: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
End With

Dim rstBalance As Recordset
Dim AccNum As String
Dim TotalBalance As Currency

While DateDiff("d", fromDate, toDate) >= 0
  With grd
    .Cols = .Cols + 1
    .Col = .Cols - 1
    colno = .Col
    .Row = 0: rowno = 0
    .TextMatrix(rowno, 0) = GetMonthString(Month(fromDate))
    .CellAlignment = 4: .CellFontBold = True
    SqlStr = "Select A.LoanID,AccNum, Balance From DepositLoanMaster A, " & _
        " DepositLoanTrans B WHERE A.LoanID = B.LoanID " & _
        " And TransID = (SELECT Max(TransID) From DepositLoanTrans C Where " & _
            " C.LoanID = B.LoanID And TransDate <= #" & fromDate & "# )" & _
        " And Balance > 0"
    gDbTrans.SqlStmt = SqlStr & sqlDepType
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
        RaiseEvent Processing("Calulating monthly Balance", count / maxCount)
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
End Sub

Public Property Let Caste(StrCaste As String)
    m_Caste = StrCaste
End Property


Public Property Let DepositType(NewdepType As wis_DepositType)
    m_DepositType = IIf(NewdepType > 100, NewdepType Mod 100, NewdepType)
End Property

Public Property Let FromAmount(newCurr As Currency)
    m_FromAmount = newCurr
End Property


Public Property Let FromIndainDate(strDate As String)
    If DateValidate(strDate, "/", True) Then
        m_FromDate = GetSysFormatDate(strDate)
        m_FromIndianDate = m_FromDate
    Else
        m_FromIndianDate = ""
        m_FromDate = vbNull
    End If

End Property

Public Property Let Gender(NewValue As wis_Gender)
    m_Gender = NewValue
End Property


Public Property Let Place(strPLace As String)
m_Place = strPLace
End Property

Public Property Let ReportOrder(newOrder As wis_ReportOrder)
    m_Order = newOrder
End Property


Public Property Let ReportType(NewType As wis_DepLoanReports)
    m_ReportType = NewType
End Property


Public Property Let ToAmount(newCurr As Currency)
    m_ToAmount = newCurr
End Property

Public Property Let ToIndainDate(strDate As String)
    If DateValidate(strDate, "/", True) Then
        m_ToDate = GetSysFormatDate(strDate)
        m_ToIndianDate = strDate
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property


Public Sub InitGrid()
Dim ReportNo As Integer
Dim ColCount As Integer
Dim ColWid As Single
For ColCount = 0 To grd.Cols - 1
    grd.ColWidth(ColCount) = GetSetting(App.EXEName, lblReportTitle.Caption, _
        "ColWidth" & ColCount, 1 / grd.Cols) * grd.Width
Next ColCount
grd.ColWidth(0) = 600

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'Set kannada caption for the all the controls
Me.cmdOk.Caption = GetResourceString(11)
Me.cmdPrint.Caption = GetResourceString(23)
End Sub


Private Sub cmdOk_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
 Set m_grdPrint = wisMain.grdPrint
 With m_grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
    .GridObject = grd
    .ReportTitle = lblReportTitle
    .PrintGrid
End With
 
Set m_grdPrint = Nothing
 
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

'Init the grid
Call PrintNoRecords(Me.grd)
'm_ReportType = 4
'Set the Referance

If m_ReportType = repDepLnBalance Then Call ReportLoanBalance
If m_ReportType = repDepLnDetail Then
    'set the report title
    lblReportTitle.Caption = GetResourceString(80) & " " & _
                           GetResourceString(295) & " " & GetFromDateString(m_ToIndianDate)  'Loan Details
    
        Call ReportLoanDetails
End If
If m_ReportType = repDepLnOverDue Then Call ReportOverdueLoans
If m_ReportType = repDepSubDayBook Then Call ReportSubDayBook
If m_ReportType = repDepLnGenLedger Then Call ReportGeneralLedger
If m_ReportType = repDepLnCashBook Then Call ReportCashBook
If m_ReportType = repDepLnMonthlyBalance Then Call ReportLoanMonthlyBalance

Screen.MousePointer = vbDefault

End Sub


Private Sub ReportGeneralLedger()
Dim SqlStmt As String
Dim OpeningBalance As Currency
Dim rst As Recordset
Dim TransDate As Date

'Build the SQL
SqlStmt = " SELECT SUM(Amount) As TotalAmount,TransDate,TransType " & _
        " FROM DepositLoanTrans " & _
        " WHERE TransDate >= #" & m_FromDate & "# " & _
        " AND TransDate <= #" & m_ToDate & "#"

If m_FromAmount > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmount
If m_ToAmount > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmount
    
If m_DepositType Then SqlStmt = SqlStmt & " AND LoanID IN " & _
    " (Select LoanID From DepositLoanMaster where DepositType  = " & m_DepositType & " )"

SqlStmt = SqlStmt & " GROUP by TransDate, TransType"

gDbTrans.SqlStmt = SqlStmt

    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Verifying records", 0)
     
'Fire the query
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    
    Dim count As Integer
    
Dim SubTotal As Currency, GrandTotal As Currency
Dim WithDraw As Currency, Deposit As Currency
Dim TotalWithDraw As Currency, TotalDeposit As Currency
Dim SlNo As Long
Dim transType As wisTransactionTypes
Dim rowno As Long, colno As Byte

   'COmpute liability (Opening Balance) as on this date
OpeningBalance = GetDepositLoanOpBalance(m_FromDate, m_DepositType)
    
'Initialize the grid
With grd
    .Cols = 6
    .Rows = 2
    .FixedCols = 1: .FixedRows = 1
    .Row = 0
    
    .Col = 0: .Text = GetResourceString(33) 'Sl NO
    .Col = 1: .Text = GetResourceString(37) '"Date"
    .Col = 2: .Text = GetResourceString(284) '"Opening Balnace
    .Col = 3: .Text = GetResourceString(271) '"Deposited"
    .Col = 4: .Text = GetResourceString(279) '"Withdrawn"
    .Col = 5: .Text = GetResourceString(42) '"Balance"
    For SlNo = 0 To .Cols - 1
        .Col = SlNo
        .CellAlignment = 4
        .CellFontBold = True
    Next
    RaiseEvent Initialise(0, rst.recordCount)
    
    .Row = 0
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0
    'ContraWithDraw = 0: ContraDeposit = 0
    TransDate = vbNull
    
    .Row = 0
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1: .Text = GetResourceString(284) '"Opening Balance"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

'Initialize Some Sub Total Varialbles
TotalDeposit = 0: TotalWithDraw = 0

TransDate = rst("TransDate")
SlNo = 0
'Fill the grid
rowno = grd.Row
While Not rst.EOF
    If TransDate <> rst("TransDate") Then
        With grd
            SlNo = SlNo + 1
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            .TextMatrix(rowno, 0) = Format(SlNo, "00")
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = FormatCurrency(OpeningBalance)

            .TextMatrix(rowno, 3) = FormatCurrency(Deposit)
            .TextMatrix(rowno, 4) = FormatCurrency(WithDraw)
            
            OpeningBalance = OpeningBalance - Deposit + WithDraw
            .TextMatrix(rowno, 5) = FormatCurrency(OpeningBalance)
            TotalWithDraw = TotalWithDraw + WithDraw
            TotalDeposit = TotalDeposit + Deposit
            WithDraw = 0: Deposit = 0
        End With
    End If
    transType = rst("TransType")
    
    If transType = wWithdraw Or transType = wContraWithdraw Then
        WithDraw = WithDraw + FormatField(rst("TotalAmount"))
    Else
        Deposit = Deposit + FormatField(rst("TotalAmount"))
    End If
    TransDate = rst("TransDate")
nextRecord:
    
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Formatting the Data ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext
    DoEvents
    Me.Refresh
Wend
    

With grd
    SlNo = SlNo + 1
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    'With loans
    .Row = rowno
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = TransDate
    .Col = 2: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    .Col = 3: .CellAlignment = 7: .Text = FormatCurrency(Deposit)
    .Col = 4: .CellAlignment = 7: .Text = FormatCurrency(WithDraw)
    
    OpeningBalance = OpeningBalance + WithDraw - Deposit
    .Col = 5: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    TotalWithDraw = TotalWithDraw + WithDraw
    TotalDeposit = TotalDeposit + Deposit
    WithDraw = 0: Deposit = 0
    
    'Display The Closing Balance
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(285) '"Closing Balance"
    .CellAlignment = 1: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(OpeningBalance)
    .CellFontBold = True: .CellAlignment = 7
    
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    'Display The Totals
    .Col = 1: .Text = GetResourceString(286) 'Grand Total
    .CellFontBold = True: .CellAlignment = 1
    .Col = 3: .Text = FormatCurrency(TotalDeposit)
    .CellFontBold = True: .CellAlignment = 7
    .Col = 4: .Text = FormatCurrency(TotalWithDraw)
    .CellFontBold = True: .CellAlignment = 7
    
End With

If DateDiff("d", m_FromDate, m_ToDate) = 0 Then
    lblReportTitle.Caption = GetResourceString(43, 58, 93) & " " & m_FromIndianDate
Else
    lblReportTitle.Caption = GetResourceString(43, 58, 93) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
End If

End Sub
Private Function GetDepositLoanOpBalance(AsOnDate As Date, _
                Optional Deptype As wis_DepositType) As Currency

Dim SqlStr As String

Dim Balance As Currency
Dim headName As String
Dim AccClass As New clsAccTrans


GetDepositLoanOpBalance = 0

'SqlStr = "SELECT LoanID, Max(TransID) As MaxTransID " & _
        " FROM DepositLoanTrans WHERE TransDate <= #" & AsOnDate & "#"

If Deptype Then
'    SqlStr = SqlStr & " ANd LoanID IN (Select LoanID From " & _
        " DepositLoanMaster Where DepositType = " & Deptype & " )"
    
    headName = GetDepositTypeText(CInt(Deptype)) & " " & GetResourceString(58)
    Balance = AccClass.GetOpBalance(GetIndexHeadID(headName), AsOnDate)
    GetDepositLoanOpBalance = Balance
    Set AccClass = Nothing
    Exit Function
End If

Dim rst As Recordset
gDbTrans.SqlStmt = "Select HeadID From Heads Where PArentID = " & parMemDepLoan
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    While Not rst.EOF
        Balance = Balance + AccClass.GetOpBalance(rst("HeadId"), AsOnDate)
        rst.MoveNext
    Wend
    GetDepositLoanOpBalance = Balance
    Set AccClass = Nothing
    Set rst = Nothing
    Exit Function
End If

SqlStr = SqlStr & " GROUP BY LoanID"

gDbTrans.SqlStmt = SqlStr
gDbTrans.CreateView ("qryTemp")

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM DepositLoanTrans A, qryTemp B " & _
    " WHERE A.LoanID=B.LoanID And A.TransID = B.MaxTransID"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then GetDepositLoanOpBalance = FormatField(rst(0))

Set AccClass = Nothing

End Function
Private Sub ReportCashBook()
Dim SqlStmt As String
Dim TmpStr As String
Dim rst As Recordset
Dim TransDate As Date

'.Clear
RaiseEvent Processing("Verifyinng records", 0)
SqlStmt = "Select A.Loanid,AccNum,TransID,Particulars," & _
        " Balance,TransDate,Amount,TransType, Name " & _
        " From DepositLoanTrans A,DepositLoanMaster B,QryName C " & _
        " WHERE B.LoanID=A.LoanID AND C.CustomerId = B.CustomerId " & _
        " AND TransDate >= #" & m_FromDate & "# AND Transdate  <= #" & m_ToDate & "# "
'Check for the deposit Type
If m_DepositType Then SqlStmt = SqlStmt & " ANd B.DepositType = " & m_DepositType

If m_FromAmount > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmount
If m_ToAmount > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmount

If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender >= " & m_Gender


'Build the Final Query
If m_Order = wisByName Then
    SqlStmt = SqlStmt & " order by TransDate, IsciName"
Else
    SqlStmt = SqlStmt & " order by TransDate, val(B.AccNum)"
End If

gDbTrans.SqlStmt = SqlStmt

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    
'Initialize the grid
    Dim SubTotal As Currency, GrandTotal As Currency
    Dim WithDraw As Currency, Deposit As Currency
    Dim ContraWithDraw As Currency, ContraDeposit As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency
    Dim TotalContraWithDraw As Currency, TotalContraDeposit As Currency
    Dim TotalBankBalance As Currency
    Dim count As Integer
    Dim SlNo As Long
    Dim rowno As Long, colno As Byte
    
    
    With grd
        .Clear: .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35): .ColAlignment(2) = 2     '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        '.Col = 5: .Text = LoadResString(gLangOffset + 41) '"Voucher No
        .Col = 5: .Text = GetResourceString(271) '"Deposited"
        '.Col = 6: .Text = GetResourceString(271) '"Deposited"
        '.Col = 7: .Text = GetResourceString(279) '"Withdrawn"
        '.Col = 8: .Text = GetResourceString(279) '"Withdrawn"
        .Col = 6: .Text = GetResourceString(279) '"Withdrawn"
         RaiseEvent Initialise(0, rst.recordCount)
         RaiseEvent Processing("Aliging the data ", 0)
        .ColAlignment(0) = 0
        .Row = 0
        For SlNo = 0 To .Cols - 1
            .Col = SlNo
             .CellAlignment = 4: .CellFontBold = True
            '.Row = 0: .CellAlignment = 4: .CellFontBold = True
            '.Row = 1: .CellAlignment = 4: .CellFontBold = True
        Next
    End With
   
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: ContraWithDraw = 0: ContraDeposit = 0

Dim PrintSubTotal As Boolean
TransDate = m_FromDate
rst.MoveFirst
TransDate = rst("TransDate")
grd.Row = 1: SlNo = 0
rowno = 1: colno = 0
While Not rst.EOF
    With grd
        'Set next row
        If TransDate <> rst("TransDate") Then
            PrintSubTotal = True
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            If Deposit Then .Col = 5: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
            'If ContraDeposit Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(ContraDeposit): .CellAlignment = 7
            If WithDraw Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
            'If ContraWithDraw Then .Col = 8: .CellFontBold = True: .Text = FormatCurrency(ContraWithDraw): .CellAlignment = 7
            TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
            'TotalContraDeposit = TotalContraDeposit + ContraDeposit: TotalContraWithDraw = TotalContraWithDraw + ContraWithDraw
            WithDraw = 0: Deposit = 0: ContraWithDraw = 0: ContraDeposit = 0
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            'rowNo = rowNo + 1
        End If
        
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate"))
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum")): .CellAlignment = 7
        .TextMatrix(rowno, 3) = FormatField(rst("Name")): .CellAlignment = 1
        .TextMatrix(rowno, 4) = FormatField(rst("Particulars")): .CellAlignment = 1
        '.Col = 5: .Text = FormatField(Rst("VoucherNo")): .CellAlignment = 4
        
        Dim transType As wisTransactionTypes
        transType = FormatField(rst("TransType"))
        
        If transType = wDeposit Or transType = wContraDeposit Then
            colno = 5: Deposit = Deposit + FormatField(rst("Amount"))
        'ElseIf TransType = wContraDeposit Then
        '    ColNo = 6: ContraDeposit = ContraDeposit + FormatField(Rst("Amount"))
        ElseIf transType = wWithdraw Or transType = wContraWithdraw Then
            colno = 6 'Colno = 7
            WithDraw = WithDraw + FormatField(rst("Amount"))
        'ElseIf TransType = wContraWithdraw Then
        '    Colno = 8: ContraWithDraw = ContraWithDraw + FormatField(Rst("Amount"))
        End If
        '.CellAlignment = 6
        .TextMatrix(rowno, colno) = FormatField(rst("Amount"))
    End With
    TransDate = rst("TransDate")
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.recordCount)
    
    rst.MoveNext
Wend

lblReportTitle.Caption = "DailyCashBookBetween" & " " & m_FromIndianDate & " " & " AND " & m_ToIndianDate

'Show sub Total
With grd
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    .Row = rowno
    .Col = 3: .CellAlignment = 4: .CellFontBold = True: .Text = GetResourceString(304) '"Sub Total "
    If Deposit Then .Col = 5: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
    'If ContraDeposit Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(ContraDeposit): .CellAlignment = 7
    If WithDraw Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
    'If ContraWithDraw Then .Col = 8: .CellFontBold = True: .Text = FormatCurrency(ContraWithDraw): .CellAlignment = 7
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    TotalContraDeposit = TotalContraDeposit + ContraDeposit: TotalContraWithDraw = TotalContraWithDraw + ContraWithDraw

'Show Grand Total
    If PrintSubTotal Then
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        .Row = rowno
        .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) '"Sub Total "
        If TotalDeposit Then .Col = 5: .CellFontBold = True: .Text = FormatCurrency(TotalDeposit): .CellAlignment = 7
        'If TotalContraDeposit Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(TotalContraDeposit): .CellAlignment = 7
        If TotalWithDraw Then .Col = 6: .CellFontBold = True: .Text = FormatCurrency(TotalWithDraw): .CellAlignment = 7
        'If TotalContraWithDraw Then .Col = 8: .CellFontBold = True: .Text = FormatCurrency(TotalContraWithDraw): .CellAlignment = 7
    End If
End With


End Sub

Private Function ReportSubDayBook() As Boolean

' Declare variables...
Dim SqlStr As String
Dim Lret As Long
Dim rptRS As Recordset
Dim PrevLoanID As Long

' Setup error handler.
On Error GoTo Err_line

lblReportTitle.Caption = GetResourceString(58) & " " & _
    GetResourceString(28) & " " & _
    GetFromDateString(m_FromIndianDate, m_ToIndianDate)  ' Loan Transaction

'raise event to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)


' Display status.
' Build the report query.
SqlStr = "SELECT 'PRINCIPAL',AccNum,DepositType, Amount,TransDate," & _
    " TransType,Balance,A.LoanID, Name as CustName," & _
    " TransID From DepositLoanMaster A,DepositLoanTrans B, QryName C WHERE " & _
    " TransDate >= #" & m_FromDate & "# AND Transdate <= #" & m_ToDate & "#" & _
    " AND A.LoanId = B.LoanID AND c.CustomerId = a.CustomerID "

'Check For Deposit Type
If m_DepositType Then SqlStr = SqlStr & " ANd A.DepositType = " & m_DepositType

If m_FromAmount > 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmount
If m_ToAmount > 0 Then SqlStr = SqlStr & " AND Amount <= " & m_ToAmount

If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender

SqlStr = SqlStr & " UNION " & _
    " SELECT 'INTEREST',AccNum,DepositType, Amount,TransDate," & _
    " TransType,Balance, A.Loanid, Name as CustName," & _
    " TransID From DepositLoanMaster A,DepositLoanIntTrans B, QryName C WHERE " & _
    " TransDate >= #" & m_FromDate & "# AND Transdate <= #" & m_ToDate & "#" & _
    " AND A.LoanId = B.LoanID AND c.CustomerId = a.CustomerID "

'Check For Deposit Type
If m_DepositType Then SqlStr = SqlStr & " ANd A.DepositType = " & m_DepositType

If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)

If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender


SqlStr = SqlStr & " ORDER BY TransDate,TransID"

If m_Order = wisByAccountNo Then SqlStr = SqlStr & ", AccNum"
If m_Order = wisByName Then SqlStr = SqlStr '& ", FirstNAme"

' Execute the query...
gDbTrans.SqlStmt = SqlStr
Lret = gDbTrans.Fetch(rptRS, adOpenForwardOnly)
'Free the memory
SqlStr = ""

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Initialize the grid.
'Call InitGrid
Dim rowno As Long, colno As Byte

RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.recordCount + 1
    .Cols = 9
    If .Rows < 20 Then .Rows = 20
    .FixedRows = 2
    .FixedCols = 2
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'Sl No
    .Col = 1: .Text = GetResourceString(37) 'Date
    .Col = 2: .Text = GetResourceString(58, 36, 60) 'Loan Account No
    .Col = 3: .Text = GetResourceString(35) 'Name
    .Col = 4: .Text = GetResourceString(272) 'Withdraw
    .Col = 5: .Text = GetResourceString(272) 'Withdraw
    .Col = 6: .Text = GetResourceString(271) 'Deposit
    .Col = 7: .Text = GetResourceString(271) 'Deposit
    .Col = 7: .Text = GetResourceString(271) 'deposit
    .Col = 8: .Text = GetResourceString(47) 'Interest
    .Row = 1
    .Col = 0: .Text = GetResourceString(33) 'Sl No
    .Col = 1: .Text = GetResourceString(37) 'Date
    .Col = 2: .Text = GetResourceString(58, 36, 60) 'Loan Account No
    .Col = 3: .Text = GetResourceString(35) 'Name
    .Col = 4: .Text = GetResourceString(269) 'Cash
    .Col = 5: .Text = GetResourceString(270) 'Contra
    .Col = 6: .Text = GetResourceString(269) 'Cash
    .Col = 7: .Text = GetResourceString(270) 'Contra
    .Col = 8: .Text = GetResourceString(47) 'Interest
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = False
    .MergeCol(5) = False
    .MergeCol(5) = False
    .MergeCol(6) = False
    .MergeCol(7) = False
    
    .MergeRow(0) = True
    .MergeRow(1) = True
    For Lret = 0 To .Cols - 1
        .Col = Lret
        .Row = 0
        .CellAlignment = 4: .CellFontBold = True
        .Row = 1
        .CellFontBold = 4: .CellAlignment = 4
    Next
    
End With

Call InitGrid

Dim transType As wisTransactionTypes
Dim SlNo As Long
Dim TransDate As Date
Dim TransID As Long
Dim SubTotal() As Currency
Dim Total() As Currency
ReDim SubTotal(4 To grd.Cols - 1)
ReDim Total(4 To grd.Cols - 1)

' Fill the rows
SlNo = 0
grd.Rows = 20
grd.Row = grd.FixedRows
rowno = grd.FixedRows
PrevLoanID = FormatField(rptRS.Fields("LoanID"))
TransDate = rptRS.Fields("TransDate")

Do While Not rptRS.EOF
    If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
    With grd
        ' Fill the loan id.
        'Put the Sub Total
        If TransDate <> rptRS("TransDate") Then
            With grd
                If .Rows <= rowno + 2 Then .Rows = rowno + 2
                rowno = rowno + 1
                colno = 3
                'To make the cell bold go to the particular cell
                .Row = rowno: .Col = colno
                .Text = GetResourceString(42)
                .CellFontBold = True
                For SlNo = 4 To .Cols - 1
                    .Col = SlNo: .Text = FormatCurrency(SubTotal(SlNo))
                    .CellFontBold = True
                    Total(SlNo) = Total(SlNo) + SubTotal(SlNo)
                    SubTotal(SlNo) = 0
                Next
            End With
            SlNo = 0
        End If
        SlNo = SlNo + 1
            
        TransDate = rptRS("TransDate")
        transType = FormatField(rptRS("TransType"))
        If PrevLoanID <> rptRS("LoanID") Then
            PrevLoanID = rptRS("LoanID")
            TransID = 0
        End If
        If TransID <> rptRS("TransID") Then
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
        End If
            
        ' Fill the transaction date.
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
         
        'Fill the loan holder name.
        .TextMatrix(rowno, 2) = FormatField(rptRS("AccNum"))
        .TextMatrix(rowno, 3) = FormatField(rptRS("Custname"))
        
        transType = FormatField(rptRS("TransType"))
        If transType = 255 Then transType = wWithdraw
        If rptRS.Fields(0) = "INTEREST" Then
            If transType = wContraDeposit Or transType = wDeposit Then colno = 8
        Else
            If transType = wWithdraw Then colno = 4
            If transType = wContraWithdraw Then colno = 5
            If transType = wDeposit Then colno = 6
            If transType = wContraDeposit Then colno = 7
        End If
        
        ' Fill the balance amount.
        .TextMatrix(rowno, colno) = FormatField(rptRS("Amount"))
        SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
    End With
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", _
                rptRS.AbsolutePosition / rptRS.recordCount)
    rptRS.MoveNext
Loop
rptRS.MoveLast

With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    
    .Row = rowno: .Col = 3
    .Text = GetResourceString(304): .CellFontBold = True
    For SlNo = 4 To .Cols - 1
        '.MergeCol(SlNo) = True
        .Col = SlNo: .Text = FormatCurrency(SubTotal(SlNo))
        .CellFontBold = True
        Total(SlNo) = Total(SlNo) + SubTotal(SlNo)
        SubTotal(SlNo) = 0
    Next
End With

    'Put GrandTotal
    With grd
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        .Row = rowno
        For SlNo = 4 To .Cols - 1
            .Col = SlNo: .Text = FormatCurrency(Total(SlNo))
            .CellFontBold = True
        Next
    End With
    
' Display the grid.
grd.Visible = True
Me.Caption = "INDEX-2000  [List of payments made...]"

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function

Private Sub ReportOverdueLoans()
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim Guarantor As Long
Dim GRName As String
Dim SqlStr As String
' Setup error handler.
On Error GoTo Err_line
Me.MousePointer = vbHourglass

lblReportTitle.Caption = GetResourceString(84) & " " & _
    GetResourceString(18) & " " & GetFromDateString(m_ToIndianDate)      'Over due loans

RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStr = "SELECT a.AccNum, A.Loanid, A.loanduedate, Balance, Name as CustName, " & _
    " a.LoanAmount FROM DepositLoanMaster A, DepositLoanTrans B, QryName C " & _
    " WHERE TransId = (Select max(TransID) From DepositLoanTrans D  Where " & _
        " D.LoanId= a.LoanId And TransDate <= #" & m_ToDate & "#) " & _
    " ANd A.LoanID= B.LoanID AND C.Customerid = A.Customerid " & _
    " AND A.loanduedate <= #" & m_FromDate & "# And Balance > 0"

'Check for the deposit Type
If m_DepositType Then SqlStr = SqlStr & " And A.DepositType = " & m_DepositType

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender


If m_FromAmount > 0 Then SqlStr = SqlStr & " AND Balance >= " & m_FromAmount
If m_ToAmount > 0 Then SqlStr = SqlStr & " AND Balance <= " & m_ToAmount

If m_Order = wisByAccountNo Then SqlStr = SqlStr & " ORDER BY val(AccNum)"
If m_Order = wisByName Then SqlStr = SqlStr & " ORDER BY IsciName"

gDbTrans.SqlStmt = SqlStr

Lret = gDbTrans.Fetch(rptRS, adOpenForwardOnly)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
Dim rowno As Long, colno As Byte

'Raise event to access frmcancel.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

' Initialize the grid.
With grd
    .Visible = False
    .Clear
    .Rows = 10
    .Cols = 5
    If .Rows < 50 Then .Rows = 10
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'SlnO
    .Col = 1: .Text = GetResourceString(58, 36, 60) 'Loan Acc NO
    .Col = 2: .Text = GetResourceString(35) 'Name
    .Col = 3: .Text = GetResourceString(209) 'Due Date
    .Col = 4: .Text = GetResourceString(80, 91) 'Amoun
    For Lret = 0 To .Cols - 1
        .Col = Lret
        .CellAlignment = 4
        .CellFontBold = True
    Next
End With

'Call InitGrid

Dim SlNo As Integer
Dim TotalAmount As Currency
Dim RegularInterest As Currency
Dim PenalInterest As Currency
Dim TillDate As String
Dim TotalRegularInterest As Currency
Dim TotalPenalInterest As Currency
TillDate = m_FromIndianDate
TotalAmount = 0
' Fill the rows
rowno = grd.Row

With rptRS
    Do While Not .EOF
        ' Set the row.
        If grd.Rows <= rowno + 2 Then grd.Rows = rowno + 2
        rowno = rowno + 1
        SlNo = SlNo + 1
        grd.Col = 0
        grd.TextMatrix(rowno, 0) = SlNo

        ' Fill the loanid.
        grd.TextMatrix(rowno, 0) = Format(SlNo, "00")

        grd.TextMatrix(rowno, 1) = .Fields("AccNum")

        ' Fill the loan holder name.
        grd.TextMatrix(rowno, 2) = .Fields("CustName")
    
    ' Fill the loan issue date.
        grd.TextMatrix(rowno, 3) = FormatField(.Fields("LoanDueDate"))

        ' Fill the loan amount.
        grd.TextMatrix(rowno, 4) = FormatCurrency(.Fields("Balance")): grd.CellAlignment = 7
        TotalAmount = TotalAmount + Val(.Fields("Balance"))
        
nextRecord:
        ' Move to next row.
        
        DoEvents
        If gCancel Then .MoveLast
        RaiseEvent Processing("Writing the data ", .AbsolutePosition / .recordCount)
        
        .MoveNext
    Loop
End With
        grd.Row = rowno
        If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Row + 2
        grd.Row = grd.Row + 1: grd.Col = 4
        grd.Text = FormatCurrency(TotalAmount): grd.CellFontBold = True
        grd.CellAlignment = 7
        
' Display the grid.
grd.Visible = True
Me.Caption = "INDEX-2000  [List of loans issued...]"

Exit_Line:
    'rptRS.Close
    Set rptRS = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Sub


'
Private Sub ReportLoanDetails()
Dim Lret As Long
Dim rptRS As Recordset
Dim Guarantor As Long
Dim GRName As String
Dim TotalBalance As Currency
Dim TotalIssue As Currency

Dim SqlStr As String
Me.MousePointer = vbHourglass


RaiseEvent Processing("Reading the records ", 0)
SqlStr = "SELECT Max(TransID) AS MaxTransID, LoanID FROM DepositLoanTrans A" & _
    " WHERE TransDate <= #" & m_ToDate & "#"

If m_DepositType Then
    SqlStr = SqlStr & " AND LoanID in (Select LoanID From DepositLoanMaster " & _
            " Where DepositType = " & m_DepositType & " )"
End If

SqlStr = SqlStr & " GROUP BY A.LoanID"

gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.CreateView("QryTemp") Then Exit Sub

SqlStr = "SELECT A.LoanID,LOanIssueDate,LoanDueDate,LoanAmount," & _
    " Balance,Caste,Place,TransDate,AccNum,InterestRate,Name as CustName" & _
    " FROM DepositLoanMaster A, DepositLoanTrans B,QryName C, QryTEMP D WHERE" & _
    " B.TransID = D.MaxTransID AND B.LoanID = D.LoanID ANd A.LoanID= B.LoanID " & _
    " AND C.CustomerID = A.CustomerID AND Balance > 0"
'Add the clause if place or caste s specified.
If Trim$(m_Place) <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)

'Add the amount specified
If Trim$(m_FromAmount) <> 0 Then SqlStr = SqlStr & " and amount>=" & m_FromAmount
If Trim$(m_ToAmount) <> 0 Then SqlStr = SqlStr & " and amount<=" & m_ToAmount

If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender

If m_Order = wisByAccountNo Then SqlStr = SqlStr & " ORDER BY val(A.AccNum)"
If m_Order = wisByName Then SqlStr = SqlStr ' & " ORDER BY A.IsciName"

gDbTrans.SqlStmt = SqlStr
Lret = gDbTrans.Fetch(rptRS, adOpenForwardOnly)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

'Raise event to access frmcancel.
RaiseEvent Initialise(0, rptRS.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

' Initialize the grid.
With grd
    .Visible = False
    .Clear
    .Rows = 20: .Cols = 11
    .FixedRows = 1
    .FixedCols = 2
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'SlNO
    .Col = 1: .Text = GetResourceString(58, 36, 60)   'Loan No
    .Col = 2: .Text = GetResourceString(35) 'Name
    .Col = 3: .Text = GetResourceString(111) 'Caste
    .Col = 4: .Text = GetResourceString(112) 'Place
    .Col = 5: .Text = GetResourceString(80, 91) 'LOan Amount
    .Col = 6: .Text = GetResourceString(340) 'Issued On
    .Col = 7: .Text = "Pledge Deposits" 'GetResourceString(pl) 'Pladge details
    .Col = 8: .Text = GetResourceString(209) 'Due Date
    .Col = 9: .Text = GetResourceString(42) 'Balance
    .Col = 10: .Text = GetResourceString(80, 47) 'Interest on loan
    For Lret = 0 To .Cols - 1
        .Col = Lret
        .CellAlignment = 4
        .CellFontBold = True
    Next
End With

'Call InitGrid
Dim SlNo As Integer
Dim TotalAmount As Currency
Dim RegularInterest As Currency
Dim LastDate As Date
Dim TotalRegularInterest As Currency
Dim Days As Integer
Dim Balance As Currency
Dim IntRate As Double
Dim rowno As Long, colno As Byte

'TillDate = m_FromIndianDate
TotalAmount = 0
' Fill the rows
rowno = grd.Row
With rptRS
    Do While Not .EOF
        ' Set the row.
        If grd.Rows <= rowno + 2 Then grd.Rows = rowno + 2
        rowno = rowno + 1
        
        LastDate = .Fields("Transdate")
        SlNo = SlNo + 1
        grd.TextMatrix(rowno, 0) = SlNo

        grd.TextMatrix(rowno, 1) = .Fields("AccNum")

       ' Fill the loan holder name.
        grd.TextMatrix(rowno, 2) = .Fields("CustName")
    
        grd.TextMatrix(rowno, 3) = .Fields("Caste")
        grd.TextMatrix(rowno, 4) = .Fields("PLace")

        grd.TextMatrix(rowno, 5) = FormatField(.Fields("LoanAmount"))
        TotalIssue = TotalIssue + Val(grd.Text)
    
       ' Fill the loan issue date.
        grd.TextMatrix(rowno, 6) = FormatField(.Fields("LoanIssueDate"))
        
        grd.TextMatrix(rowno, 8) = FormatField(.Fields("LoanDueDate"))

        ' Fill the loan amount.
        Balance = .Fields("Balance")
        grd.TextMatrix(rowno, 9) = FormatCurrency(Balance)
        TotalAmount = TotalAmount + Balance
        
        IntRate = FormatField(.Fields("InterestRate"))
        Days = DateDiff("d", LastDate, m_ToIndianDate)
        RegularInterest = Balance * (Days / 365) * (IntRate / 100)
        If RegularInterest > 0 Then
            grd.TextMatrix(rowno, 10) = FormatCurrency(RegularInterest \ 1): grd.CellAlignment = 7
            TotalRegularInterest = TotalRegularInterest + (RegularInterest \ 1)
        End If
        TotalBalance = TotalBalance + Val(grd.Text)
        
nextRecord:
       'Move to next row.
        DoEvents
        If gCancel Then .MoveLast
        RaiseEvent Processing("Writing the data ", .AbsolutePosition / .recordCount)
        
        .MoveNext

    Loop
End With


grd.Row = rowno
If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Row + 2
grd.Row = grd.Row + 1

grd.Col = 5: grd.Text = FormatCurrency(TotalIssue): grd.CellFontBold = True: grd.CellAlignment = 7
grd.Col = 10: grd.Text = FormatCurrency(TotalRegularInterest): grd.CellFontBold = True: grd.CellAlignment = 7
grd.Col = 9: grd.Text = FormatCurrency(TotalBalance): grd.CellFontBold = True: grd.CellAlignment = 7 ' Display the grid.
grd.Visible = True

Me.Caption = "INDEX-2000  [List of loans issued...]"

Exit_Line:
    Set rptRS = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line

End Sub
Private Sub ReportLoanBalance()

Dim SqlStmt As String
Dim TotDepBal As Currency
Dim Balance As Currency
Dim StrAmt As String
Dim rstDeposit As Recordset
Dim rst As Recordset
Dim TotalBalance As Currency
Dim SlNo As Long

'raiseevent to access frmcancel
lblReportTitle.Caption = GetResourceString(58) & " " & _
                GetResourceString(42) & GetFromDateString(m_ToIndianDate) 'Dep Balance

RaiseEvent Processing("Reading & Verifying the data ", 0)
SqlStmt = "Select A.LoanId,Balance, AccNum, Place,Caste,Name as CustName" & _
    " From DepositLoanMaster A,DepositLoanTrans B, QryName C WHERE TransID = " & _
        "(SELECT Max(transID) From DepositLoanTrans D Where D.LoanID = A.LoanID " & _
        " And TransDate <= #" & m_ToDate & "#)" & _
    " AND B.LoanID = A.LoanID AND C.CustomerID = A.CustomerID "

'check for the individual report--FD,RD,PD,Dep
If m_DepositType Then SqlStmt = SqlStmt & " AND A.DepositType = " & m_DepositType
    
If m_FromAmount > 0 Then SqlStmt = SqlStmt & " AND Balance >= " & m_FromAmount
If m_ToAmount > 0 Then SqlStmt = SqlStmt & " AND Balance <= " & m_ToAmount

If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender

SqlStmt = SqlStmt & " ORder BY " & IIf(m_Order = wisByAccountNo, " Val(AccNum)", " IsciName")
gDbTrans.SqlStmt = SqlStmt
SqlStmt = ""
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim rowno As Integer, colno As Byte

' InitGird
With grd
    .Clear
    .Cols = 4
    .FixedCols = 1
    .Rows = 20
    
    RaiseEvent Initialise(0, rst.recordCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36, 60) '" Loan Id "
    .ColAlignment(1) = 7
    .Col = 2: .Text = GetResourceString(35)
    .ColAlignment(2) = 1
    .Col = 3: .Text = GetResourceString(42)
    If Trim$(m_Place) <> "" Then
        .Cols = .Cols + 1
        .Col = .Cols - 1
        .Text = GetResourceString(270)
    End If
    If Trim$(m_Caste) <> "" Then
        .Cols = .Cols + 1
        .Col = .Cols - 1
        .Text = GetResourceString(100)
    End If
    For SlNo = 0 To .Cols - 1
        .Col = SlNo
        .CellAlignment = 4
        .CellFontBold = True
    Next
    
End With

SlNo = 1
rowno = grd.Row
While Not rst.EOF
    If FormatField(rst("Balance")) <> 0 Then
    With grd
        If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
        .Row = SlNo
        .TextMatrix(SlNo, 0) = Format(SlNo, "00"): .CellAlignment = 1
        .TextMatrix(SlNo, 1) = FormatField(rst("AccNum")): .CellAlignment = 4
        .TextMatrix(SlNo, 2) = FormatField(rst("CustName")): .CellAlignment = 1
        .TextMatrix(SlNo, 3) = FormatField(rst("Balance")): .CellAlignment = 7
        colno = 3
        TotalBalance = TotalBalance + Val(.Text)
        If Trim$(m_Place) <> "" Then
           colno = colno + 1
           .TextMatrix(SlNo, colno) = FormatField(rst("Place"))
        End If
        If Trim$(m_Caste) <> "" Then
           colno = colno + 1
           .TextMatrix(SlNo, colno) = FormatField(rst("Caste"))
        End If
    End With
    
    SlNo = SlNo + 1
    DoEvents
    If gCancel Then rst.MoveNext
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.recordCount)
    End If
    rst.MoveNext
Wend

With grd
    If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
    SlNo = SlNo + 1
    If .Rows <= SlNo + 2 Then .Rows = .Rows + 1
    
    SlNo = SlNo + 1
    .Row = SlNo
    .Col = 2: .Text = GetResourceString(52, 42) ' "Total Balance"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

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
cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)

Dim Wid As Single
Dim I As Integer
Wid = (grd.Width - 185) / grd.Cols
Call InitGrid
End Sub


Private Sub grd_LostFocus()

Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, lblReportTitle.Caption, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount
End Sub


Private Sub m_grdPrint_MaxProcessCount(maxCount As Long)
m_TotalCount = maxCount
Set m_frmCancel = New frmCancel
m_frmCancel.Show
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


