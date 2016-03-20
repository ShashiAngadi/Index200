VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCAReport 
   Caption         =   "Current Account Reports.."
   ClientHeight    =   5895
   ClientLeft      =   1125
   ClientTop       =   1905
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   6465
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   5160
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   3240
         TabIndex        =   6
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   1845
         TabIndex        =   4
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   400
         Left            =   3780
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.CheckBox chkDetails 
         Caption         =   "Show details"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   3285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8334
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2670
      TabIndex        =   5
      Top             =   60
      Width           =   1365
   End
End
Attribute VB_Name = "frmCAReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_FromIndianDate As String
Dim m_ToIndianDate As String
Dim m_FromDate As Date
Dim m_ToDate As Date
Dim m_FromAmt As Currency
Dim m_ToAmt As Currency
Dim m_Caste As String
Dim m_Place As String
Dim m_Gender As wis_Gender
Dim m_ReportOrder As wis_ReportOrder
Dim m_ReportType As wis_CAReports
Private m_AccGroup As Integer

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Public Event Initialize(Min As Integer, Max As Integer)
Public Event Processing(strMessage As String, Ratio As Single)



Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroup = NewValue
End Property


Public Property Let Caste(StrCaste As String)
    m_Caste = StrCaste
End Property

Public Property Let FromAmount(newAmount As Currency)
m_FromAmt = newAmount
End Property

Public Property Let Gender(NewGender As wis_Gender)
    m_Gender = NewGender
End Property

Public Property Let ReportOrder(newOrder As wis_ReportOrder)
    m_ReportOrder = newOrder
End Property

Public Property Let ReportType(NewReportType As wis_CAReports)
 m_ReportType = NewReportType
End Property

Public Property Let ToAmount(newAmount As Currency)
m_ToAmt = newAmount
End Property

Public Property Let ToIndianDate(strDate As String)
    If DateValidate(strDate, "/", True) Then
        m_ToIndianDate = strDate
        m_ToDate = GetSysFormatDate(strDate)
        'm_ToIndianDate = GetAppFormatDate(m_ToDate)

    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(strDate As String)
    If DateValidate(strDate, "/", True) Then
        m_FromIndianDate = strDate
        m_FromDate = GetSysFormatDate(strDate)
        'm_FromIndianDate = GetAppFormatDate(m_FromDate)
    Else
        m_FromIndianDate = ""
        m_FromDate = vbNull
    End If
End Property

Public Property Let Place(strPLace As String)
    m_Place = strPLace
End Property


Private Sub InitGrid(Optional Resize As Boolean)

Dim ColCount As Integer
For ColCount = 0 To grd.Cols - 1
    With grd
        .ColWidth(ColCount) = GetSetting(App.EXEName, "CAReport" & m_ReportType, "ColWidth" & ColCount, 1 / .Cols) * .Width
        If .ColWidth(ColCount) >= .Width Then .ColWidth(ColCount) = .Width / 3
        If .ColWidth(ColCount) <= 0 Then .ColWidth(ColCount) = .Width / .Cols
    End With
Next ColCount

ErrLine:
    If Err Then
        MsgBox "InitGrid Error :" & vbCrLf & Err.Description
    End If

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

chkDetails.Caption = GetResourceString(295) 'Details
cmdOk.Caption = GetResourceString(11)
cmdPrint.Caption = GetResourceString(23)

End Sub


Private Sub ShowAccountsClosed()

Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim NoAnd As Boolean   'Stupid variable

RaiseEvent Processing("Reading & Verifying the records. ", 0)

'Fire SQL
SqlStmt = "Select AccId,AccNum,ClosedDate," & _
    " Title+' '+FirstName+' '+MiddleName+' '+LastName as Name" & _
    " FROM CAmaster A, NameTab B" & _
    " WHERE ClosedDate >= #" & m_FromDate & "#" & _
    " And ClosedDate <= #" & m_ToDate & "#" & _
    " And B.CustomerID = A.CustomerID "
If m_Place <> "" Then SqlStmt = SqlStmt & " AND PLACE = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND CASTE = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupID = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & " order by ClosedDate, val(AccNum)"
Else
    SqlStmt = SqlStmt & " order by ClosedDate, IsciName"
End If

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

'Initialize the grid
grd.Cols = 4
grd.FixedCols = 1
grd.Row = 0
grd.Col = 0: grd.Text = GetResourceString(33)  '"Sl No"
grd.CellAlignment = 4: grd.CellFontBold = True
grd.Col = 1: grd.Text = GetResourceString(36, 60) '"Acc No"
grd.CellAlignment = 4: grd.CellFontBold = True
grd.Col = 2: grd.Text = GetResourceString(35)  '"Name"
grd.CellAlignment = 4: grd.CellFontBold = True
grd.Col = 3: grd.Text = GetResourceString(281)   '"Create Date"
grd.CellAlignment = 4: grd.CellFontBold = True

grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 1
grd.ColAlignment(2) = 0
grd.ColAlignment(3) = 2

RaiseEvent Initialize(0, rst.RecordCount)
RaiseEvent Processing("Reading the data to write into the grid.", 0)

Dim rowno As Long

rowno = grd.Row

'Fill the grid
While Not rst.EOF
    DoEvents
    Me.Refresh
    
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        
        .TextMatrix(rowno, 0) = " " & Format(rowno + 1, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = " " & FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("ClosedDate"))
    
nextRecord:
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Writing into grid.", rst.AbsolutePosition / rst.RecordCount)
        rst.MoveNext
    End With
Wend
lblReportTitle.Caption = GetResourceString(422) & " " & _
        GetResourceString(65) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)


End Sub

Private Sub ShowAccountsCreated()

'Declarig the variables
Dim SqlStmt As String
Dim rst As ADODB.Recordset

RaiseEvent Processing("Reading and verifying the records ", 0)
'Fire SQL
SqlStmt = "Select AccID,AccNum,CreateDate, " & _
    " Title+' '+firstName+' '+MiddleName+' '+LastName as Name " & _
    " FROM CAMaster A, NameTab B WHERE " & _
    " CreateDate >= #" & m_FromDate & "#" & _
    " AND CreateDate <= #" & m_ToDate & "#" & _
    " AND B.CustomerID = A.CustomerID "

If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender

If m_Place <> "" Then SqlStmt = SqlStmt & " AND PLACE = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND CASTE = " & AddQuotes(m_Caste, True)
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupID = " & m_AccGroup

If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & " order by CreateDate, IsciName"
Else
    SqlStmt = SqlStmt & " order by CreateDate, val(AccNum)"
End If

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
        
'Initialize the grid
With grd
    .Cols = 4
    .FixedCols = 1
    .Row = 0
    .Col = 0: .CellFontBold = True
        .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .CellFontBold = True
        .Text = GetResourceString(36, 60) '"Acc No"
    .Col = 2: .CellFontBold = True
        .Text = GetResourceString(35) '"Name"
    .Col = 3: .CellFontBold = True
        .Text = GetResourceString(281)  '"Create Date"
End With

    RaiseEvent Initialize(0, rst.RecordCount)
    RaiseEvent Processing("Arranging the data to write into the grid. ", 0)
    
Dim rowno As Long
    
    
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    DoEvents
    Me.Refresh
    'Set next row
    With grd
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        
        .TextMatrix(rowno, 0) = " " & Format(rowno, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("CreateDate"))
    End With
nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend

lblReportTitle.Caption = GetResourceString(64) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)


End Sub

Private Sub ShowBalances()
'Declare the variables
Dim rst As ADODB.Recordset
Dim I As Long
Dim Total As Currency
Dim SQL As String
    
RaiseEvent Processing("Reading & Verifying the data.", 0)
 
SQL = "Select A.AccId, AccNum, Balance, " & _
    " Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
    " From CAtrans A, CaMaster B,NameTab C where TransID = " & _
        "(Select MAX(TransID) from CAtrans D WHERE D.AccID = A.AccID " & _
            " AND TransDate <= #" & m_ToDate & "#)" & _
    " And A.AccId = B.AccId And B.CustomerId = C.CustomerId "

If m_FromAmt > 0 Then SQL = SQL & " And Balance >= " & m_FromAmt
If m_ToAmt > 0 Then SQL = SQL & " And Balance <=  " & m_ToAmt

'Query by caste
If m_Caste <> "" Then SQL = SQL & " And Caste =  " & AddQuotes(m_Caste, True)
'Query by PLACE
If m_Place <> "" Then SQL = SQL & " And Place=  " & AddQuotes(m_Place, True)
'Query by Gender
If m_Gender <> wisNoGender Then SQL = SQL & " And Gender =  " & m_Gender
If m_AccGroup Then SQL = SQL & " AND AccGroupID = " & m_AccGroup

If m_ReportOrder = wisByName Then
    SQL = SQL & " Order By IsciName"
Else
    SQL = SQL & " Order By val(AccNum)"
End If

gDbTrans.SqlStmt = SQL
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

'Initialize the grid
With grd
    .Cols = 4
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 1: .Text = GetResourceString(36) & " " & _
                GetResourceString(60)  ' "Acc No"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 2: .Text = GetResourceString(35) ' "Name"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 3: .Text = GetResourceString(42)  ' "Balance"
    .CellFontBold = True: .CellAlignment = 4
    .ColAlignment(0) = 1
    .ColAlignment(2) = 0
    .ColAlignment(3) = 1
End With
RaiseEvent Initialize(0, rst.RecordCount)
RaiseEvent Processing("Arranging the data to write into the grid.", 0)

Dim SlNo As Long
Dim rowno As Integer, colno As Byte
rowno = grd.Row

SlNo = 0
While Not rst.EOF
    DoEvents
    Me.Refresh
    'See if you have to show this record
    If FormatField(rst("Balance")) = 0 Then GoTo nextRecord
    
    With grd
        'Set next row
        If .Rows = rowno + 2 Then .Rows = .Rows + 1
        rowno = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("Balance"))
        Total = Total + FormatField(rst("Balance"))
        'Total = Total + Val(.Text) 'FormatField(Rst("Balance"))
        If Val(.Text) < 0 Then
            .Row = rowno: .Col = 3
            .Text = FormatCurrency(Abs(.Text))
            .CellForeColor = vbRed
        End If
        '.CellAlignment = 7
    End With
    
nextRecord:
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing into the grid. ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Move next row
With grd

    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    .Row = rowno
    .Col = 2: .CellFontBold = True
    .Text = GetResourceString(52, 42) '"Total Balances"
    .Col = 3: .CellFontBold = True
    .Text = FormatCurrency(Total): .CellAlignment = 7
End With
    lblReportTitle.Caption = GetResourceString(422) & " " & _
                GetResourceString(67) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
            


End Sub


Private Sub ShowMonthlyBalances()

Dim count As Long
Dim totalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SqlStmt As String

Dim fromDate As Date
Dim toDate As Date

'Get the Last day of the given month
toDate = GetSysLastDate(m_ToDate)


fromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(m_FromDate)), GetMonthString(Month(m_ToDate)))

SqlStmt = "SELECT A.AccNum,A.AccID, A.CustomerID," & _
    " Title + ' '+ FirstNAme +' '+ MiddleName +' '+ LastName as CustNAme " & _
    " From CAMaster A,NameTab B,CAMaster C WHERE A.CreateDate <= #" & toDate & "#" & _
    " AND (A.ClosedDate Is NULL OR A.Closeddate >= #" & fromDate & "#)" & _
    " AND  B.CustomerID = A.CustomerID" & _
    " AND C.AccID = A.AccID Order By val(a.ACCNUM)"

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub
'Set rstMain = gDBTrans.Rst.Clone
count = DateDiff("M", fromDate, toDate) + 2
totalCount = (count + 1) * rstMain.RecordCount
'RaiseEvent Initialize(0, TotalCount)

Dim prmAccId As Parameter
Dim prmdepositId As Parameter

With grd
    .Clear
    .Cols = 3
    .Rows = 5
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'Sl No
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(36) 'AccountNo
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(35) 'Name
    .CellAlignment = 4: .CellFontBold = True
End With

Dim rowno As Integer, colno As Byte

count = 0
While Not rstMain.EOF
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: count = count + 1
        .TextMatrix(rowno, 0) = count
        .TextMatrix(rowno, 1) = FormatField(rstMain("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rstMain("CustName"))
    End With
    rstMain.MoveNext
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
Wend
    
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        .Row = rowno + 1
        .Col = 2: .Text = GetResourceString(286) 'Grand Total
        .CellFontBold = True
    End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

'FromDate = "4/30/" & Year(FromDate)
'FromDate = m_FromDate

Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    SqlStmt = "SELECT [AccId], Max([TransID]) AS MaxTransID" & _
            " FROM CATrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY [AccID];"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("CAMonBal")
    SqlStmt = "SELECT A.AccId,Balance From CATrans A,CAMonBal B " & _
        " Where B.AccId = A.AccID ANd  TransID =MaxTransID"
    gDbTrans.SqlStmt = SqlStmt
    
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then ProcCount = ProcCount + rstMain.RecordCount: GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0
        .Col = .Cols - 1: .Text = GetMonthString(Month(fromDate)) & _
                " " & GetResourceString(42)
        .CellAlignment = 4: .CellFontBold = True
        rowno = 0
        colno = .Cols - 1
    End With
    
    rstMain.MoveFirst
    TotalBalance = 0
    
    While Not rstMain.EOF
        'grd.Row = grd.Row + 1
        rowno = rowno + 1
        Balance = 0
        rstBalance.MoveFirst
        rstBalance.Find "ACCID = " & rstMain("AccID")
        If Not rstBalance.EOF Then Balance = FormatField(rstBalance("Balance"))
        
        grd.TextMatrix(rowno, colno) = FormatCurrency(Balance)
        rstMain.MoveNext
        TotalBalance = TotalBalance + Balance
        
        ProcCount = ProcCount + 1
        DoEvents
        If gCancel Then Exit Do
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
    Wend
    If grd.Rows < rowno + 3 Then grd.Rows = rowno + 3
    grd.Row = rowno + 2
    grd.Text = FormatCurrency(TotalBalance)
    grd.CellFontBold = True

NextMonth:
    RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
'    rstBalance.MoveFirst
    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

Exit Sub
ErrLine:
    If Err Then MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear

End Sub


Private Sub ShowCALedger()
Dim SqlStmt As String
Dim OpeningDate As Date
Dim OpeningBalance As Currency
Dim rst As ADODB.Recordset
Dim TransDate As Date

'Get liability on a day before m_fromdate
OpeningDate = DateAdd("d", -1, m_FromDate)

lblReportTitle.Caption = GetResourceString(422) & " " & _
                    GetResourceString(93)   '(CA) Ledger

'Build the SQL
SqlStmt = " SELECT SUM(Amount) As TotalAmount,TransDate,TransType FROM CATrans " & _
        " WHERE TransDate >= #" & m_FromDate & "#" & _
        " and TransDate <= #" & m_ToDate & "#"

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

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
Dim PrintSubTotal As Boolean
    
'COmpute liability (Opening Balance) as on this date
'OpeningBalance = ComputeTotalCALiability(OpeningDate)
Dim AccClass As New clsAccTrans
OpeningBalance = AccClass.GetOpBalance( _
        GetIndexHeadID(GetResourceString(422)), m_FromDate)
Set AccClass = Nothing

'Initialize the grid
With grd
    .Clear
    .Cols = 6
    .Rows = 2
    .FixedCols = 1: grd.FixedRows = 1
    .Row = 0
    
    .Col = 0: .Text = GetResourceString(33) 'Sl NO
    .CellFontBold = True: .CellAlignment = 4
    .Col = 1: .Text = GetResourceString(37) '"Date"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 2: .Text = GetResourceString(284) '"OPening Blance
    .CellFontBold = True: .CellAlignment = 4
    .Col = 3: .Text = GetResourceString(271) '"Deposited"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 4: .Text = GetResourceString(272)  '"Withdrawn"
    .CellFontBold = True: .CellAlignment = 4
    .Col = 5: .Text = GetResourceString(285) '"Closing Balance"
    .CellFontBold = True: .CellAlignment = 4

'    RaiseEvent Initialise(0, Rst.RecordCount)
    
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

'Written On 25th Aug 2000 ----Vinay
'Initialize Some Sub Total Varialbles
TotalDeposit = 0: TotalWithDraw = 0

Dim rowno As Integer, colno As Byte

TransDate = rst("TransDate")
SlNo = 0
rowno = grd.Row

'Fill the grid
While Not rst.EOF
    If TransDate <> rst("TransDate") Then
        With grd
            SlNo = SlNo + 1
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            'OpeningBalance = ComputeTotalSBLiability(Transdate)
            .TextMatrix(rowno, 0) = Format(SlNo, "00")
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = FormatCurrency(OpeningBalance)
            .TextMatrix(rowno, 3) = FormatCurrency(Deposit)
            .TextMatrix(rowno, 4) = FormatCurrency(WithDraw)
            OpeningBalance = OpeningBalance + Deposit - WithDraw
            .TextMatrix(rowno, 0) = FormatCurrency(OpeningBalance)
            TotalWithDraw = TotalWithDraw + WithDraw
            TotalDeposit = TotalDeposit + Deposit
            WithDraw = 0: Deposit = 0
        End With
    End If
    transType = rst("TransType")
    
    If transType = wWithdraw Or transType = wContraWithdraw Then
        WithDraw = WithDraw + FormatField(rst("TotalAmount"))
    Else 'If TransType = wDeposit or TransType = wcontraDeposit Then
        Deposit = Deposit + FormatField(rst("TotalAmount"))
    End If
    TransDate = rst("TransDate")

nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Formatting the Data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend


lblReportTitle.Caption = GetResourceString(422) & " " & _
        GetResourceString(93) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)


With grd
    .Row = rowno
    SlNo = SlNo + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = GetIndianDate(TransDate)
    .Col = 2: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    .Col = 3: .CellAlignment = 7: .Text = FormatCurrency(Deposit)
    .Col = 4: .CellAlignment = 7: .Text = FormatCurrency(WithDraw)
    OpeningBalance = OpeningBalance + Deposit - WithDraw
    .Col = 5: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    TotalWithDraw = TotalWithDraw + WithDraw
    TotalDeposit = TotalDeposit + Deposit
    WithDraw = 0: Deposit = 0
    
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(285)
    .CellAlignment = 7: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 7: .CellFontBold = True
    
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

End Sub

Private Function ShowJointAccounts() As Boolean
Dim SqlStr As String
Dim rst As ADODB.Recordset

'SET THE CAPTION
lblReportTitle.Caption = GetResourceString(265) & " " & _
    GetResourceString(36)  'Joint account

SqlStr = "SELECT A.AccID, B.CustomerID as MainCustID, A.CustomerID as JointCustID, " & _
    " AccNum FROM CAJoint A,CAMaster B Where B.AccID = A.AccID " & _
    " AND (B.ClosedDate is NULL OR B.ClosedDate > #" & m_ToDate & "#)"

gDbTrans.SqlStmt = SqlStr & " ORDER BY AccNum"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

Dim SlNo As Long
Dim CustClass As clsCustReg
Set CustClass = New clsCustReg

'Now List
With grd
    .Clear
    .Cols = 3
    .FixedCols = 1
    .Rows = 10
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '" SlNO
    .Col = 1: .Text = GetResourceString(36, 60) '" AccNum
    .Col = 2: .Text = GetResourceString(35) '" Name
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
End With

SlNo = 0: grd.Row = 0
Dim AccNum As String
Dim rowno As Integer, colno As Byte

rowno = grd.Row

While Not rst.EOF
    With grd
        If AccNum <> FormatField(rst("AccNum")) Then
            AccNum = FormatField(rst("AccNum"))
            If .Rows <= rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1: SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = Format(SlNo, "00")
            .TextMatrix(rowno, 1) = AccNum
            .TextMatrix(rowno, 2) = CustClass.CustomerName(FormatField(rst("MainCustID")))
        End If
        If .Rows <= rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        .TextMatrix(rowno, 1) = AccNum
        .TextMatrix(rowno, 2) = CustClass.CustomerName(FormatField(rst("JointCustID")))
    End With
    rst.MoveNext
Wend

Set CustClass = Nothing

lblReportTitle.Caption = GetResourceString(422) & " " & _
        GetResourceString(177)

ShowJointAccounts = True

End Function


Private Sub ShowProductsAndInterests()

Dim AccIDs() As Long
Dim Products() As Currency
Dim Mn As Integer
Dim Yr As Integer
Dim I As Integer
Dim This_date As Date
Dim Rate As Double
Dim Total() As Currency

ReDim Total(0)
'Prelim checks
If m_FromIndianDate = "" Then Exit Sub
If m_ToIndianDate = "" Then Exit Sub
    
'Initialize the grid
'Initialize the grid
With grd
    .Clear
    .Cols = 2
    .Rows = 12
    .FixedRows = 2
    .FixedCols = 1
    .Row = 0: .Col = 0
    .Text = GetResourceString(36, 60) ' "Acc No"
    .CellAlignment = 4: .CellFontBold = True
    
    .Row = 1: .Col = 0
    .Text = GetResourceString(36, 60) ' "Acc No"
    .CellAlignment = 4: .CellFontBold = True
End With

'Get the month start and number of months
If DateDiff("D", m_FromDate, m_ToDate) < 0 Then Exit Sub

'Get the interest rate from setup
    Dim l_Setup As New clsSetup
    Rate = Val(l_Setup.ReadSetupValue("SBAcc", "RateOfInterest", 0))
RaiseEvent Processing("Reading the setup value for interest ", 0)
'Loop through all the months till To_Date starting with From_Date
    This_date = m_FromDate
    Dim MonthCount As Integer
    Dim TotalMonth As Integer
    Dim rowno As Integer, colno As Byte
    
    TotalMonth = DateDiff("m", m_FromDate, m_ToDate)
    RaiseEvent Initialize(0, TotalMonth)
    
    While DateDiff("m", This_date, m_ToDate) >= 0
        ReDim AccIDs(0)
        ReDim Products(0)
        DoEvents
        Me.Refresh
        Mn = Month(This_date)
        Yr = Year(This_date)
        If TotalMonth <= 0 Then TotalMonth = 1
        MonthCount = MonthCount + 1
        RaiseEvent Processing("Computing the interest for the month  " & GetMonthString(Mn), MonthCount / TotalMonth)
        
        Call ComputeCAProducts(AccIDs(), Mn, Yr, Products())
        RaiseEvent Processing("Computing the interest for the month  " & GetMonthString(Mn), MonthCount / TotalMonth)
        'NOTE:
        'I have deliberately put the foll 3 separate loops to minimize switching
        'between columns and to make code more readable. GIRISH
        
        'LOOP1:     Print the AccIDs the first time
        If m_FromDate = This_date Then
            'Reset the row
          With grd
            .Row = 0: .Col = 0
            .Text = GetResourceString(36, 60) '"Account No."
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .Col = 0
            .Text = GetResourceString(36, 60) '"Account No."
            .CellAlignment = 4: .CellFontBold = True
            
            .Row = 0: .Col = 1
            .Text = GetResourceString(35) 'Name
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .Col = 1
            .Text = GetResourceString(35) 'Name
            .CellAlignment = 4: .CellFontBold = True
            
            .Col = 0
            rowno = .Row: colno = 0
            For I = 0 To UBound(AccIDs) - 1
                If .Rows = rowno + 1 Then .Rows = .Rows + 1
                rowno = rowno + 1
                .TextMatrix(rowno, 0) = AccIDs(I)
            Next I
            .Col = 1: colno = 1
          End With
        End If
        
        'LOOP 2:    Print the products
        With grd
            .Row = 0: rowno = 0
            If .Cols = .Col + 1 Then .Cols = .Cols + 1
            .Col = .Col + 1
            .Text = GetMonthString(Mn) & " " & Yr
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .Text = GetResourceString(66) ''Product
            .CellAlignment = 4: .CellFontBold = True
            For I = 0 To UBound(AccIDs) - 1
                .Row = .Row + 1
                .Text = FormatCurrency(Products(I))
            Next I
            
            'LOOP 3:    Print the interest values
            .Row = 0
            If .Cols = .Col + 1 Then .Cols = .Cols + 1
            .Col = .Col + 1
            .Text = GetMonthString(Mn) & " " & Yr
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .Text = GetResourceString(47) ''INterest
            .CellAlignment = 4: .CellFontBold = True
        End With
        colno = grd.Col: rowno = grd.Row
        For I = 0 To UBound(AccIDs) - 1
            DoEvents
            If gCancel Then Exit Sub
            Me.Refresh
            rowno = rowno + 1
            grd.TextMatrix(rowno, colno) = FormatCurrency(ComputeCAInterest(Products(I), Rate))
            If UBound(Total) < I Then ReDim Preserve Total(I)
            Total(I) = Total(I) + Val(grd.Text)
        Next I
        'Move to next month
        This_date = DateAdd("m", 1, This_date)
    Wend
    
With grd
    If .Cols = .Col + 1 Then .Cols = .Cols + 1
    .Col = .Col + 1
    .Row = 0
    .Text = GetResourceString(52)
    .CellFontBold = True
    .Row = 1
    .Text = GetResourceString(52)
    .CellFontBold = True
    
    Dim Grand As Currency
    
    For I = 0 To UBound(AccIDs) - 1
        .Row = .Row + 1: .Text = FormatCurrency(Total(I) \ 1)
        Grand = Grand + Val(.Text)
    Next I
    
End With
    
'Now Get the Account No & name of the accountHolder
gDbTrans.SqlStmt = "SeLECT AccId,AccNum, " & _
    " FirstNAme + ' ' + MiddleNAme +' ' + LastNAme as Name " & _
    " FROM CAMaster A,NAmeTab B Where B.CustomerId = A.CustomerId " & _
    " ORDER BY AccID"

Dim rst As ADODB.Recordset
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

grd.Row = 1: grd.Col = 0
I = 0
rowno = 1
While Not rst.EOF
    With grd
        rowno = rowno + 1
        If .TextMatrix(rowno, 0) <> rst("AccID") Then GoTo nextRecord
        .TextMatrix(rowno, colno) = rst("AccNUm")
        .TextMatrix(rowno, 1) = FormatField(rst("Name"))
        .Col = .Cols - 1 = Total(I)
        Grand = Grand + Total(I)
    End With
nextRecord:
    rst.MoveNext
    I = I + 1
Wend

lblReportTitle.Caption = GetResourceString(422, 66, 47, 295)


'Now write the grand total
If grd.Rows > grd.Row Then grd.Rows = grd.Rows + 1
grd.Row = grd.Row + 1: grd.Text = FormatCurrency(Grand): grd.CellFontBold = True

grd.MergeCells = flexMergeFree
grd.MergeCol(0) = True
grd.MergeCol(1) = True
grd.MergeRow(0) = True
grd.MergeRow(1) = True
 
End Sub

Private Sub ShowSubDayBook()

Dim SqlStmt As String
Dim TmpStr As String
Dim rst As ADODB.Recordset
Dim TransDate As Date
Dim CAClass As ClsCAAcc


'.Clear
RaiseEvent Processing("Verifyinng records", 0)

'SET THE CAPTION
lblReportTitle.Caption = GetResourceString(390, 63) 'Sub day book

SqlStmt = "Select A.Accid,AccNum,TransID,Particulars,Balance," & _
    " TransDate,Amount, VoucherNo,TransType, " & _
    " Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
    " From CATrans A,CAMaster B,NameTab C WHERE B.AccID=A.AccID " & _
    " AND C.CustomerId = B.CustomerId " & _
    " AND TransDate >= #" & m_FromDate & "# " & _
    " AND Transdate  <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender

If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupID = " & m_AccGroup

'Build the Final Query
If m_ReportOrder = wisByName Then
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
    Dim rowno  As Integer, colno As Byte
    Dim PrintSubTotal As Boolean
    
    With grd
        .Clear: .Cols = 10
        .FixedCols = 1
        .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        .Col = 5: .Text = GetResourceString(41) '"Voucher No
        .Col = 6: .Text = GetResourceString(271) '"Deposited"
        .Col = 7: .Text = GetResourceString(271) '"Deposited"
        .Col = 8: .Text = GetResourceString(279) '"Withdrawn"
        .Col = 9: .Text = GetResourceString(279) '"Withdrawn"
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35): .ColAlignment(2) = 2     '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        .Col = 5: .Text = GetResourceString(41) '"Voucher No
        .Col = 6: .Text = GetResourceString(269) 'CAsh
        .Col = 7: .Text = GetResourceString(270) 'Contra
        .Col = 8: .Text = GetResourceString(269) 'CAsh
        .Col = 9: .Text = GetResourceString(270) 'Contra
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
         
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
'         RaiseEvent Initialise(0, Rst.RecordCount)
'         RaiseEvent Processing("Aliging the data ", 0)
        .ColAlignment(0) = 0
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        For SlNo = 0 To .Cols - 1
            .Col = SlNo
            .Row = 0: .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .CellAlignment = 4: .CellFontBold = True
        Next
    End With
   
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: ContraWithDraw = 0: ContraDeposit = 0

rst.MoveFirst
TransDate = rst("TransDate")
grd.Row = 1: SlNo = 0
rowno = 1
While Not rst.EOF
    With grd
        'Set next row
        If TransDate <> rst("TransDate") Then
            PrintSubTotal = True
            If rowno < .Rows + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
            .Col = 7: .CellFontBold = True: .Text = FormatCurrency(ContraDeposit): .CellAlignment = 7
            .Col = 8: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
            .Col = 9: .CellFontBold = True: .Text = FormatCurrency(ContraWithDraw): .CellAlignment = 7
            
            TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
            TotalContraDeposit = TotalContraDeposit + ContraDeposit: TotalContraWithDraw = TotalContraWithDraw + ContraWithDraw
            WithDraw = 0: Deposit = 0: ContraWithDraw = 0: ContraDeposit = 0
            If .Rows < rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
        End If
        
        If rowno < .Rows + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate"))
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 3) = FormatField(rst("Name"))
        .TextMatrix(rowno, 4) = FormatField(rst("Particulars"))
        .TextMatrix(rowno, 5) = FormatField(rst("VoucherNo"))
        
        Dim transType As wisTransactionTypes
        Dim Amount As Currency
        transType = rst("TransType")
        Amount = FormatField(rst("Amount"))
        If transType = wDeposit Then
            colno = 6: Deposit = Deposit + Amount
        ElseIf transType = wContraDeposit Then
            colno = 7: ContraDeposit = ContraDeposit + Amount
        ElseIf transType = wWithdraw Then
            colno = 8: WithDraw = WithDraw + Amount
        ElseIf transType = wContraWithdraw Then
            colno = 9: ContraWithDraw = ContraWithDraw + Amount
        End If
        If Amount > 0 Then .TextMatrix(rowno, colno) = FormatCurrency(Amount)
    End With
    TransDate = rst("TransDate")
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend


lblReportTitle.Caption = GetResourceString(422, 390, 63) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)

'Show sub Total
With grd
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    .Row = rowno
    .Col = 3: .CellAlignment = 4: .CellFontBold = True: .Text = GetResourceString(304) '"Sub Total "
    .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
    .Col = 7: .CellFontBold = True: .Text = FormatCurrency(ContraDeposit): .CellAlignment = 7
    .Col = 8: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
    .Col = 9: .CellFontBold = True: .Text = FormatCurrency(ContraWithDraw): .CellAlignment = 7
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    TotalContraDeposit = TotalContraDeposit + ContraDeposit: TotalContraWithDraw = TotalContraWithDraw + ContraWithDraw

'Show Grand Total
'If it has alredy printed the sub total then print the sub total
'of the next day 'if we are showing the deatil of one day then
'need not show the sub total ' so overwrite gransd total on sub total
        
    If PrintSubTotal Then
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
    End If
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) '"Sub Total "
    .Col = 6: .CellFontBold = True: .Text = FormatCurrency(TotalDeposit): .CellAlignment = 7
    .Col = 7: .CellFontBold = True: .Text = FormatCurrency(TotalContraDeposit): .CellAlignment = 7
    .Col = 8: .CellFontBold = True: .Text = FormatCurrency(TotalWithDraw): .CellAlignment = 7
    .Col = 9: .CellFontBold = True: .Text = FormatCurrency(TotalContraWithDraw): .CellAlignment = 7
End With


End Sub


Private Sub ShowSubCashBook()

Dim SqlStmt As String
Dim TmpStr As String
Dim rst As ADODB.Recordset
Dim TransDate As Date
Dim CAClass As ClsCAAcc

'.Clear
RaiseEvent Processing("Verifyinng records", 0)

'SET THE CAPTION
lblReportTitle.Caption = GetResourceString(390, 63) 'Sub day book

SqlStmt = "Select A.Accid,AccNum,TransID,Particulars,Balance," & _
    " TransDate,Amount, VoucherNo,TransType, " & _
    " Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName As Name " & _
    " From CATrans A,CAMaster B,NameTab C WHERE B.AccID=A.AccID " & _
    " AND C.CustomerId = B.CustomerId " & _
    " AND TransDate >= #" & m_FromDate & "# " & _
    " AND Transdate  <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender

If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupID = " & m_AccGroup

'Build the Final Query
If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & " order by TransDate, IsciName"
Else
    SqlStmt = SqlStmt & " order by TransDate, val(B.AccNum)"
End If

gDbTrans.SqlStmt = SqlStmt

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    
'Initialize the grid
    Dim SubTotal As Currency, GrandTotal As Currency
    Dim WithDraw As Currency, Deposit As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency
    Dim TotalBankBalance As Currency
    Dim count As Integer
    Dim SlNo As Long
    Dim rowno As Integer, colno As Byte
    Dim PrintSubTotal As Boolean
    
    With grd
        .MergeCells = flexMergeNever
        .Clear: .Cols = 8
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        .Col = 5: .Text = GetResourceString(41) '"Voucher No
        .Col = 6: .Text = GetResourceString(271) '"Deposited"
        .Col = 7: .Text = GetResourceString(279) '"Withdrawn"
        
        For SlNo = 0 To .Cols - 1
            .Col = SlNo
            .CellAlignment = 4: .CellFontBold = True
        Next
    End With
   
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0

rst.MoveFirst
TransDate = rst("TransDate")
grd.Row = 1: SlNo = 0
rowno = 1: colno = 0
While Not rst.EOF
    With grd
        'Set next row
        If TransDate <> rst("TransDate") Then
            SlNo = 0
            PrintSubTotal = True
            If rowno > .Rows - 2 Then .Rows = .Rows + 2
            rowno = rowno + 1
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
            .Col = 7: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
            
            TotalWithDraw = TotalWithDraw + WithDraw
            TotalDeposit = TotalDeposit + Deposit
            WithDraw = 0: Deposit = 0
            If rowno > .Rows + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
        End If
        
        If rowno > .Rows - 2 Then .Rows = .Rows + 2
        rowno = rowno + 1: SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate"))
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 3) = FormatField(rst("Name"))
        .TextMatrix(rowno, 4) = FormatField(rst("Particulars"))
        .TextMatrix(rowno, 5) = FormatField(rst("VoucherNo"))
        
        Dim transType As wisTransactionTypes
        Dim Amount As Currency
        transType = rst("TransType")
        Amount = FormatField(rst("Amount"))
        If transType = wDeposit Or transType = wContraDeposit Then
            colno = 6: Deposit = Deposit + Amount
        Else
            colno = 7: WithDraw = WithDraw + Amount
        End If
        
        If Amount > 0 Then .TextMatrix(rowno, colno) = FormatCurrency(Amount)
    End With
    TransDate = rst("TransDate")
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend


lblReportTitle.Caption = GetResourceString(422, 390, 85) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)

'Show sub Total
With grd
    If rowno > .Rows - 2 Then .Rows = .Rows + 2
    rowno = rowno + 1
    .Row = rowno
    .Col = 3: .CellAlignment = 4: .CellFontBold = True: .Text = GetResourceString(304) '"Sub Total "
    .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
    .Col = 7: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    
'Show Grand Total
'If it has alredy printed the sub total then print the sub total
'of the next day 'if we are showing the deatil of one day then
'need not show the sub total ' so overwrite gransd total on sub total
        
    If PrintSubTotal Then
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
    End If
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) '"Sub Total "
    .Col = 6: .CellFontBold = True: .Text = FormatCurrency(TotalDeposit): .CellAlignment = 7
    .Col = 7: .CellFontBold = True: .Text = FormatCurrency(TotalWithDraw): .CellAlignment = 7
End With


End Sub


Private Sub cmdOk_Click()
Unload Me
End Sub




Private Sub cmdPrint_Click()

Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .Font.name = gFontName
    .Font.Size = gFontSize
    .CompanyName = gCompanyName
    .GridObject = grd
    .ReportTitle = lblReportTitle.Caption
    .PrintGrid
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
 'Center the form
 Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption

'Init the grid
    grd.Rows = 50
    grd.Cols = 1
    grd.FixedCols = 0
    grd.Row = 1
    
    grd.Text = GetResourceString(278)   '"No Records Available"
    grd.CellAlignment = 4: grd.CellFontBold = True
Screen.MousePointer = vbHourglass
        
    If m_ReportType = repCABalance Then
        Call ShowBalances
    ElseIf m_ReportType = repCADayBook Then
        Call ShowSubDayBook
    ElseIf m_ReportType = repCACashBook Then
        Call ShowSubCashBook
    ElseIf m_ReportType = repCALedger Then
        Call ShowCALedger
    ElseIf m_ReportType = repCAAccOpen Then
        Call ShowAccountsCreated
    ElseIf m_ReportType = repCAAccClose Then
        Call ShowAccountsClosed
    ElseIf m_ReportType = repCAProduct Then
        Call ShowProductsAndInterests
    ElseIf m_ReportType = repCAJoint Then
        Call ShowJointAccounts
    ElseIf m_ReportType = repCAMonthlyBalance Then
        Call ShowMonthlyBalances
    End If
    
    
Screen.MousePointer = vbNormal
Me.lblReportTitle.FONTSIZE = 16

End Sub

Private Sub Form_Resize()
Screen.MousePointer = vbHourglass
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
InitGrid (True)

Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False, Cancel)

Set frmCAReport = Nothing
   
End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "CAReport" & m_ReportType, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount
End Sub


Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_TotalCount = MaxCount
Set m_frmCancel = New frmCancel
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


