VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSBReport 
   Caption         =   "Savings Bank Reports.."
   ClientHeight    =   6090
   ClientLeft      =   2115
   ClientTop       =   1935
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   6540
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   5580
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web View"
         Height          =   450
         Left            =   720
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   450
         Left            =   2430
         TabIndex        =   3
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   450
         Left            =   3780
         TabIndex        =   2
         Top             =   30
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8334
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
      Height          =   195
      Left            =   1650
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frmSBReport"
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

Dim m_Place As String
Dim m_Caste As String

Dim m_Gender As wis_Gender
Dim m_ReportType As wis_SBReports
Dim m_ReportOrder As wis_ReportOrder
Dim m_AccGroup As Integer

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)
Public Event WindowClosed()
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

Public Property Let ReportType(NewReportType As wis_SBReports)
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



Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'Me.chkDetails.Caption = GetResourceString(295)  ' Details
Me.cmdOK.Caption = GetResourceString(11)  '
cmdPrint.Caption = GetResourceString(23)
End Sub

'
Private Sub ShowAccountsClosed()

Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim NoAnd As Boolean   'Stupid variable
RaiseEvent Processing("Verifying records", 0)
'Fire SQL
SqlStmt = "Select AccId,AccNum, CreateDate, Name, " & _
    " ClosedDate  From SBMaster A Inner Join QryName B" & _
    " ON A.CustomerID = B.CustomerID " & _
    " Where ClosedDate  <= #" & m_ToDate & "#" & _
    " AND ClosedDate >= #" & m_FromDate & "#"

If m_Caste <> "" Then _
    SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then _
    SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then _
    SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_AccGroup Then _
    SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByName Then
    gDbTrans.SqlStmt = SqlStmt & " Order by CreateDate,IsciName"
Else
    gDbTrans.SqlStmt = SqlStmt & " order by CreateDate,VAL(AccNum)"
End If

    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
     DoEvents
     If gCancel Then Exit Sub
     RaiseEvent Processing("Verifying records", 0)
        
'Initialize the grid
    With grd
        .Cols = 4
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)   '"Sl No"
        .Col = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(282)  '"Closed Date"
    End With
    
    RaiseEvent Initialise(0, rst.RecordCount)
    RaiseEvent Processing("Reading  Data", 0)
    'Dim SlNo As Long
    Dim rowno As Integer, colno As Integer
'Fill the grid
    rowno = grd.Row: colno = 0
    While Not rst.EOF
        With grd
            'Set next row
            'If .Rows = .Row + 1 Then .Rows = .Rows + 1
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            '.Row = .Row + 1: SlNo = SlNo + 1
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = Format(rowno, "00")
            colno = 1: .TextMatrix(rowno, colno) = " " & FormatField(rst("AccNum"))
            colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("Name"))
            colno = 3: .TextMatrix(rowno, colno) = " " & FormatField(rst("ClosedDate"))
        End With
nextRecord:
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Reading record", rst.AbsolutePosition / rst.RecordCount)
        rst.MoveNext
    Wend
    
    If WisDateDiff(m_FromIndianDate, m_ToIndianDate) = 0 Then
        lblReportTitle.Caption = GetResourceString(421) & " " & _
            GetResourceString(65) & " " & m_FromIndianDate
    Else
        lblReportTitle.Caption = GetResourceString(421) & " " & _
            GetResourceString(65) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    End If

End Sub

Private Sub ShowAccountsCreated()

Dim SqlStmt As String
Dim rst As Recordset
Dim NoAnd As Boolean
'Fire SQL
    
SqlStmt = "Select AccNum,AccId,Name," & _
    " Createdate from SBMaster A Inner Join qryName B" & _
    " On A.CustomerID = B.CustomerID " & _
    " Where CreateDate  <= #" & m_ToDate & "#" & _
    " AND CreateDate >= #" & m_FromDate & "#"

If m_Caste <> "" Then _
    SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then _
    SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then _
    SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_AccGroup Then _
    SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & " order by  val(AccNum),CreateDate"
Else
    gDbTrans.SqlStmt = SqlStmt & " order by  IsciName, CreateDate"
End If
    
DoEvents
 If gCancel Then Exit Sub
 RaiseEvent Processing("Verifying records", 0)
 
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

DoEvents
 If gCancel Then Exit Sub
 RaiseEvent Processing("Verifying records", 0)
 
'Initialize the grid
    RaiseEvent Initialise(0, rst.RecordCount)
    With grd
        .Cols = 4: .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)  '"Sl No"
        .Col = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(281)  '"Create Date"
    End With
    'Dim SlNo As Long
    Dim rowno As Integer, colno As Integer
'Fill the grid
    While Not rst.EOF
        With grd
            'Set next row
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            
            colno = 0: .TextMatrix(rowno, colno) = Format(rowno, "00")
            colno = 1: .TextMatrix(rowno, colno) = " " & FormatField(rst("AccNum"))
            colno = 2: .TextMatrix(rowno, colno) = " " & FormatField(rst("Name"))
            colno = 3: .TextMatrix(rowno, colno) = " " & FormatField(rst("CreateDate"))
        End With
nextRecord:
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Verifying records", rst.AbsolutePosition / rst.RecordCount)
        rst.MoveNext
    Wend

If WisDateDiff(m_FromIndianDate, m_ToIndianDate) = 0 Then
    lblReportTitle.Caption = GetResourceString(421) & " " & _
        GetResourceString(64) & " " & m_FromIndianDate
Else
    lblReportTitle.Caption = GetResourceString(421) & " " & _
        GetResourceString(64) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
End If

End Sub
Private Function ShowAccountsWithCheque() As Boolean

Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim AccHeadID As Long


SqlStr = "SELECT Distinct AccID From ChequeMaster WHERE " & _
    " AccHeadId = " & GetIndexHeadID(GetResourceString(421)) & _
        " AND Trans = " & wischqIssue

SqlStr = "SELECT AccId,AccNum, Name" & _
    " FROM SBMaster A Inner Join QryName B" & _
    " On A.CustomerID = B.CustomerID " & _
    " Where ACCID in (" & SqlStr & ")"

If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup


gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

'Now Display the Details of the account
With grd
    .Clear
    .Rows = 15
    .Cols = 3
    .FixedCols = 1
    .FixedRows = 1
    .Row = 0
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36) & " " & _
            GetResourceString(60)
    .Col = 2: .Text = GetResourceString(35)
    .Col = 0: .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .CellAlignment = 4: .CellFontBold = True
End With

'Dim SlNo As Long
Dim rowno As Integer, colno As Integer

grd.Row = 0: rowno = 0

While Not rst.EOF
    With grd
        If .Rows <= rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        colno = 0: .TextMatrix(rowno, colno) = Format(rowno, "00")
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("Name"))
    End With
    rst.MoveNext
Wend

lblReportTitle.Caption = GetResourceString(421) & " " & _
            GetResourceString(177) & " " & GetFromDateString(m_ToIndianDate)

End Function
Private Function ShowJointAccounts() As Boolean
Dim AccNum As String
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim SlNo As Long
Dim CustClass As New clsCustReg

'(select Count(AccID) From SBJoint) > 0

SqlStr = "SELECT A.AccID,B.CustomerID as MainCustID, A.CustomerID as JointCustId, AccNum" & _
    " FROM SBMAster A Inner Join SBJoint B On A.AccID = B.AccID " & _
    " Where (A.ClosedDate Is NULL OR A.ClosedDate > #" & m_ToDate & "# )" & _
    " ORDER BY val(A.AccNum)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

'Now List
With grd
    .Clear
    .Cols = 3
    .Rows = 10
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '" SlNO
    .Col = 1: .Text = GetResourceString(36, 60) '" AccNum
    .Col = 2: .Text = GetResourceString(35) '" Name
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
End With

SlNo = 0: grd.Row = 0
While Not rst.EOF
    With grd
        If AccNum <> FormatField(rst("AccNum")) Then
            AccNum = FormatField(rst("AccNum"))
            SlNo = SlNo + 1
            If .Rows <= .Row + 1 Then .Rows = .Row + 2
            .Row = .Row + 1
            .Col = 0: .Text = Format(SlNo, "00")
            .Col = 1: .Text = AccNum
            .Col = 2: .Text = CustClass.CustomerName(FormatField(rst("MainCustID")))
        End If
        .Col = 1: .Text = AccNum
        If .Rows <= .Row + 1 Then .Rows = .Row + 2
        .Row = .Row + 1
        .Col = 2: .Text = CustClass.CustomerName(FormatField(rst("JointCustID")))
    End With
    rst.MoveNext
Wend

Set CustClass = Nothing

'lblReportTitle.Caption = "SB Joint accounts"
lblReportTitle.Caption = GetResourceString(265) & " " & _
    GetResourceString(421)

ShowJointAccounts = True

End Function
'
Private Sub ShowSBLedger()
Dim SqlStmt As String
Dim OpeningDate As Date
Dim OpeningBalance As Currency
Dim rst As ADODB.Recordset
Dim transdate As String

'Get liability on a day before fromdate ---siddu
OpeningDate = DateAdd("d", -1, m_FromDate)

'Check for the Sql
SqlStmt = " SELECT SUM(Amount) As TotalAmount,TransDate,TransType FROM SBTrans " & _
    " WHERE TransDate >= #" & m_FromDate & "#" & _
    " and TransDate <= #" & m_ToDate & "#"

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

SqlStmt = SqlStmt & " GROUP by TransDate, TransType"
gDbTrans.SqlStmt = SqlStmt

DoEvents
If gCancel Then Exit Sub
RaiseEvent Processing("Verifying records", 0)

'Fetch the query
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim count As Integer
Dim SubTotal As Currency, GrandTotal As Currency
Dim WithDraw As Currency, Deposit As Currency
Dim TotalWithDraw As Currency, TotalDeposit As Currency
Dim SlNo As Long
Dim transType As wisTransactionTypes

'COmpute liability (Opening Balance) as on this date
OpeningBalance = ComputeTotalSBLiability(OpeningDate)

'Initialize the grid
With grd
    .Clear
    .Cols = 6
    .Rows = 2
    .FixedCols = 1: .FixedRows = 1
    .Row = 0
    
    .Col = 0: .Text = GetResourceString(33) 'Sl NO
    .Col = 1: .CellAlignment = 4: .Text = GetResourceString(37) '"Date"
    .Col = 2: .Text = GetResourceString(284) ' OpeningBalance
    .Col = 3: .CellAlignment = 4: .Text = GetResourceString(271) '"Deposited"
    .Col = 4: .CellAlignment = 4: .Text = GetResourceString(272)  '"Withdrawn" Debit
    .Col = 5: .Text = GetResourceString(285) '"Closing Balance" 'GetResourceString(42) '"Closing Balance" Credit
End With

RaiseEvent Initialise(0, rst.RecordCount)
grd.Row = 0
SubTotal = 0: GrandTotal = 0
WithDraw = 0: Deposit = 0
'ContraWithDraw = 0: ContraDeposit = 0
transdate = ""
grd.Row = 0

TotalDeposit = 0: TotalWithDraw = 0
'TotalContraDeposit = 0: TotalContraWithDraw = 0

transdate = FormatField(rst("TransDate"))
SlNo = 0

'FIll Up THe Opening balance
grd.Row = grd.FixedRows
grd.Col = 1
grd.Text = GetResourceString(284)
grd.CellFontBold = True
grd.Col = 2
grd.Text = FormatCurrency(OpeningBalance)
grd.CellFontBold = True

Dim rowno As Integer, colno As Integer
rowno = grd.Row: colno = grd.Col

'Fill the grid
While Not rst.EOF
    If transdate <> FormatField(rst("TransDate")) Then
        With grd
            SlNo = SlNo + 1
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00")
            colno = 1: .TextMatrix(rowno, colno) = transdate
            colno = 2: .TextMatrix(rowno, colno) = FormatCurrency(OpeningBalance)
            colno = 3: .CellAlignment = 7: .TextMatrix(rowno, colno) = FormatCurrency(Deposit)
            .Row = rowno
            .Col = 4: .CellAlignment = 7: .TextMatrix(rowno, .Col) = FormatCurrency(WithDraw)
            OpeningBalance = OpeningBalance + Deposit - WithDraw
            .Col = 5: .CellAlignment = 7: .TextMatrix(rowno, .Col) = FormatCurrency(OpeningBalance)
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
    transdate = FormatField(rst("TransDate"))
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Formatting the Data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend


With grd
    SlNo = SlNo + 1
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    
    'Last  OB
    .Col = 2: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    
    'Get the new Ob
    OpeningBalance = OpeningBalance + Deposit - WithDraw
   
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = transdate
    .Col = 3: .CellAlignment = 7: .Text = FormatCurrency(Deposit)
    .Col = 4: .CellAlignment = 7: .Text = FormatCurrency(WithDraw)
    .Col = 5: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    TotalWithDraw = TotalWithDraw + WithDraw
    TotalDeposit = TotalDeposit + Deposit
    WithDraw = 0: Deposit = 0
    
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

    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4
    .Text = GetResourceString(285) 'CLosing Balance
    .CellFontBold = True
    .Col = .Cols - 1
    .Text = FormatCurrency(OpeningBalance)
    .CellFontBold = True

End With
    
    lblReportTitle.Caption = GetResourceString(93) & " " & _
          GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    lblReportTitle.FontBold = True

End Sub

Private Sub ShowBalance()
Dim rst As ADODB.Recordset
Dim I As Long
Dim SqlStmt As String
Dim Total As Currency
Dim SubTotal As Currency

If m_ToIndianDate = "" Then ToIndianDate = gStrDate

'Fire SQL Query

gDbTrans.SqlStmt = "SELECT AccID, Max(TransID) AS MaxTransID FROM SBTrans" & _
        " WHERE TransDate <= #" & m_ToDate & "# GROUP BY AccID"
gDbTrans.CreateView ("qrySBMaxTransID")

SqlStmt = "SELECT A.CustomerID,Name,A.AccId,AccNum,Balance " & _
    " from (SBMaster A  Inner Join qryName B On A.CustomerID=B.CustomerID)" & _
    " Inner Join (qrySBMaxTransID C Inner Join SBTrans D " & _
        " On  C.MaxTransID = D.TransID And C.AccID = D.AccID)" & _
    " On A.AccID =C.AccID"
    
    Dim sqlClause As String
    'sqlClause = GetReportFilter(m_Place, m_Caste, "B")
    'sqlClause = IIf(Len(sqlClause) > 0, " AND ", "") & sqlClause
    
    If m_Place <> "" Then sqlClause = sqlClause & " And Place = " & AddQuotes(m_Place, True)
    If m_Caste <> "" Then sqlClause = sqlClause & " And Caste = " & AddQuotes(m_Caste, True)
    
    If m_Gender <> wisNoGender Then sqlClause = sqlClause & " And Gender = " & m_Gender
    If m_AccGroup Then sqlClause = sqlClause & " AND AccGroupId = " & m_AccGroup
    
    If m_FromAmt > 0 Then sqlClause = sqlClause & " And BALANCE >= " & m_FromAmt
    If m_ToAmt > 0 Then sqlClause = sqlClause & " And BALANCE <= " & m_ToAmt
    
    sqlClause = Trim$(sqlClause)
    If Len(sqlClause) Then sqlClause = " WHERE " & Mid(sqlClause, 4)
    
    If m_ReportOrder = wisByName Then
        gDbTrans.SqlStmt = SqlStmt & sqlClause & " ORDER BY IsciName "
    Else
        gDbTrans.SqlStmt = SqlStmt & sqlClause & " ORDER BY val(A.AccNum)"
    End If
    
    DoEvents
    If gCancel Then Exit Sub
    RaiseEvent Processing("Verifying records", 0)
     
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    DoEvents
     If gCancel Then Exit Sub
     RaiseEvent Processing("Verifying records", 0)
     
'Initialize the Grid
    Dim count As Integer
    With grd
        .Clear
        .Cols = 4
        .FixedRows = 1: .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) '"SlNo"
        .Col = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(42) '"Balance"
    End With
    
Dim AccId As Long


RaiseEvent Initialise(0, 100)
    
grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 0
grd.ColAlignment(2) = 1
Dim name As String
Dim SlNo As Long
Dim rowno As Integer, colno As Integer

SlNo = 0
While Not rst.EOF
    'See if account is closed
    If FormatField(rst("Balance")) = 0 Then  'Closed account
        AccId = FormatField(rst("AccID"))
        GoTo nextRecord
    End If
    
    'See if you have to show this record
    If AccId = FormatField(rst("AccID")) Then GoTo nextRecord
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        name = FormatField(rst("Name"))
        colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00")
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = name
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rst("Balance"))
        .CellAlignment = 7
    End With
    AccId = FormatField(rst("AccID"))
    Total = Total + FormatField(rst("Balance"))
    SubTotal = SubTotal + FormatField(rst("Balance"))

nextRecord:
    If rowno Mod 50 = 0 Then
        With grd
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            .Col = 2: .TextMatrix(rowno, .Col) = GetResourceString(304) '"Grand Total"
            .CellAlignment = 4: .CellFontBold = True
            .Col = 3: .CellFontBold = True: .TextMatrix(rowno, .Col) = FormatCurrency(SubTotal): .CellFontBold = True
            .CellAlignment = 7: SubTotal = 0
        End With
    End If
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Reading records", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Set next row and print grand total
With grd
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1
    .Col = 2: .TextMatrix(rowno, .Col) = GetResourceString(286) '"Grand Total"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .CellFontBold = True: .TextMatrix(rowno, .Col) = FormatCurrency(Total): .CellFontBold = True
    .CellAlignment = 7
End With


lblReportTitle.Caption = GetResourceString(421) & " " & _
            GetResourceString(67) & " " & GetFromDateString(m_ToIndianDate)


End Sub


'
Private Sub ShowMonthlyBalances()

Dim count As Long
Dim totalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SqlStmt As String

Dim fromDate As Date
Dim toDate As Date

'Get the Date
toDate = m_ToDate
'Get the Last day of the given month
toDate = GetSysLastDate(m_ToDate)
fromDate = GetSysLastDate(m_FromDate)


'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463) & " " & _
        GetResourceString(67) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))

SqlStmt = "SELECT A.AccNum,A.AccID, A.CustomerID, Name as CustName " & _
    " From SBMaster A Inner Join QryName B On A.CustomerID = B.CustomerID " & _
    " WHERE A.CreateDate <= #" & toDate & "#" & _
    " AND (A.ClosedDate Is NULL OR A.Closeddate >= #" & fromDate & "#)"


If m_AccGroup Then SqlStmt = SqlStmt & " And A.AccGroupId = " & m_AccGroup
SqlStmt = SqlStmt & " Order By " & _
        IIf(m_ReportOrder = wisByAccountNo, "val(A.ACCNUM)", "IsciName")
        
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub
'Set rstMain = gDBTrans.Rst.Clone
count = DateDiff("M", fromDate, toDate) + 2
totalCount = (count + 1) * rstMain.RecordCount
RaiseEvent Initialise(0, totalCount)

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
    .Col = 1: .Text = GetResourceString(36) 'AccountNo
    .Col = 2: .Text = GetResourceString(35) 'Name
    .Row = .FixedRows - 1: count = 0
End With

Dim rowno As Integer, colno As Integer

While Not rstMain.EOF
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: count = count + 1
        colno = 0: .TextMatrix(rowno, colno) = count
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rstMain("AccNum"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rstMain("CustNAme"))
        .RowData(.Row) = 0
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
    rstMain.MoveNext
Wend
    
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        colno = 2: .TextMatrix(rowno, colno) = GetResourceString(286) 'Grand Total
    End With


Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    
    SqlStmt = "SELECT AccId, Max(TransID) AS MaxTransID" & _
            " FROM SBTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY AccID"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("SBMonBal")
    
    SqlStmt = "SELECT A.AccId,Balance From SBTrans A,SbMonBal B " & _
        " Where B.AccId = A.AccID ANd TransID =MaxTransID"
    
    gDbTrans.SqlStmt = SqlStmt
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        rowno = 0
        colno = .Cols - 1: .TextMatrix(rowno, colno) = GetMonthString(Month(fromDate)) & _
                " " & GetResourceString(42)
    End With
    
    rstMain.MoveFirst
    TotalBalance = 0
    
    While Not rstMain.EOF
        rowno = rowno + 1
        rstBalance.MoveFirst
        rstBalance.Find "ACCID = " & rstMain("AccID")
        If rstBalance.EOF Then GoTo NextAccount
        If rstBalance("Balance") = 0 Then GoTo NextAccount
        With grd
            .TextMatrix(rowno, colno) = FormatField(rstBalance("Balance"))
            .RowData(rowno) = 1
        End With
        Balance = rstBalance("Balance")
        TotalBalance = TotalBalance + Balance
        
        
        DoEvents
        If gCancel Then rstMain.MoveLast
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
                
NextAccount:
        rstMain.MoveNext
        ProcCount = ProcCount + 1
    Wend
    
    With grd
        rowno = rowno + 2
        .TextMatrix(rowno, colno) = FormatCurrency(TotalBalance)
        .Col = colno: .Row = rowno
        .CellFontBold = True
        .RowData(rowno) = 1
    End With
    
NextMonth:

'    rstBalance.MoveFirst
    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

If gCancel Then Exit Sub

''Now Checkall the accounts
'Delete the account from grid which are not having any balance
With grd
    count = 0
    Do
        count = count + 1
        If count >= .Rows Then Exit Do
        If .RowData(count) = 0 Then .RemoveItem (count): count = count - 1
    Loop
    
End With


Exit Sub
ErrLine:
    MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear
End Sub

Private Sub ShowProductsAndInterests()

Dim AccIDs() As Long
Dim Products() As Currency
Dim Total() As Currency
Dim Mn As Integer
Dim Yr As Integer
Dim I As Integer
Dim This_date As Date
Dim Rate As Double
Dim l_Setup As New clsSetup
Dim rst As ADODB.Recordset
Dim rowno As Integer, colno As Integer
Dim strAccNum As String

lblReportTitle.Caption = GetResourceString(66)

'Prelim checks
If Not DateValidate(m_FromIndianDate, "/", True) Then Exit Sub
If Not DateValidate(m_ToIndianDate, "/", True) Then Exit Sub
    
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
DoEvents
If gCancel Then Exit Sub
RaiseEvent Processing("Verifying records", 0)

'Get the interest rate from setup
Rate = Val(l_Setup.ReadSetupValue("SBAcc", "RateOfInterest", 0))

'Loop through all the months till To_Date starting with From_Date
This_date = m_FromDate
RaiseEvent Initialise(0, 12)
ReDim Total(0)

'Get the Record setfor the a/c holdeer name
gDbTrans.SqlStmt = "Select AccId,AccNum, Name " & _
    " FROM SBMaster A Inner Join QryName B On A.CustomerId = B.CustomerId " & _
    " WHERE A.AccID In ( Select Distinct AccID from SbTrans)" & _
    " ORDER BY A.AccID"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim SetUp As New clsSetup
Dim noIntOnMin As Boolean
noIntOnMin = CBool(SetUp.ReadSetupValue("SBAcc", "NoInterestOnMinBalance", "False"))
Set SetUp = Nothing

While DateDiff("m", This_date, m_ToDate) >= 0
    ReDim AccIDs(0)
    ReDim Products(0)
    Mn = Month(This_date)
    RaiseEvent Processing("Claculating Interest for month " & _
        GetMonthString(Mn), Mn / 12)
    Yr = Year(This_date)
    
    'Call ComputeSBProducts_New(AccIDs(), Mn, Yr, Products(), 16,noIntOnMin )
    Call ComputeSBProducts_Daily(AccIDs(), Products(), This_date, m_ToDate, Rate, noIntOnMin)
    'NOTE:
    'I have deliberately put the full 3 separate loops to minimize switching
    'between columns and to make code more readable. GIRISH
    'LOOP1:     Print the AccIDs the first time
    If m_FromDate = This_date Then
        'Reset the row
        With grd
            .Row = 0: .Col = 0
            .Text = GetResourceString(36, 60) '"Account No."
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1
            .Text = GetResourceString(36, 60) '"Account No."
            .CellAlignment = 4: .CellFontBold = True
            
            .Row = 0: .Col = 1
            .Text = GetResourceString(35) 'Name
            .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .Col = 1
            .Text = GetResourceString(35) 'Name
            .CellAlignment = 4: .CellFontBold = True
            
            'Set the Number of rows for the grid
            .Rows = UBound(AccIDs) + 1
            
            .Col = 0
            
            colno = 0: rowno = .Row
            For I = 0 To UBound(AccIDs) - 1
                rowno = rowno + 1 'Identify the next row
                If .Rows = rowno + 1 Then .Rows = .Rows + 1
                '.Row = .Row + 1
                '.Text = AccIDs(I)
                'Find the Account number
                rst.Find ("AccID = " & AccIDs(I))
                If Not rst.EOF Then
                    .TextMatrix(rowno, 0) = FormatField(rst("AccNum")) 'AccIDs(I)
                    .TextMatrix(rowno, 1) = FormatField(rst("Name"))
                Else
                    .TextMatrix(rowno, 0) = AccIDs(I)
                End If
                rst.MoveFirst
            Next I
            .Col = 1: colno = 1
        End With
    End If
    
    'LOOP 2:    Print the products
    With grd
        .Row = 0
        If .Cols = .Col + 1 Then .Cols = .Cols + 1
        .Col = .Col + 1
        
        rowno = 0
        colno = .Col
        
        .Text = GetMonthString(Mn) & " " & Yr
        .CellAlignment = 4: .CellFontBold = True
        .Row = 1
        .Text = GetResourceString(66) ''Product
        .CellAlignment = 4: .CellFontBold = True
        
        rowno = .Row
        For I = 0 To UBound(AccIDs) - 1
            '.Row = .Row + 1
            '.Text = FormatCurrency(Products(I))
            rowno = rowno + 1
            .TextMatrix(rowno, colno) = FormatCurrency(Products(I))
        Next I
        
        'LOOP 3:    Print the interest values
        .Row = 0
        If .Cols = .Col + 1 Then .Cols = .Cols + 1
        .Col = .Col + 1
        .Text = GetMonthString(Mn) & " " & Yr
        .CellAlignment = 4: .CellFontBold = True
        .Row = 1: .Text = GetResourceString(47) ''INterest
        .CellAlignment = 4: .CellFontBold = True
        'RaiseEvent Initialise(0, UBound(AccIDs) - 1)
    
        rowno = .Row: colno = .Col
    End With
    
    For I = 0 To UBound(AccIDs) - 1
        DoEvents
        If gCancel Then Exit For
        Me.Refresh
        'grd.Row = grd.Row + 1
        rowno = rowno + 1
        'grd.Text = FormatCurrency(ComputeSBInterest(Products(I), Rate))
        grd.TextMatrix(rowno, colno) = FormatCurrency(ComputeSBInterest(Products(I), Rate))
        If UBound(Total) < I Then ReDim Preserve Total(I)
        Total(I) = Total(I) + Val(grd.Text)
    Next I
    
    'Move to next month
    This_date = DateAdd("m", 1, This_date)
Wend
    
With grd
    .Row = 1
    If .Cols = .Col + 1 Then .Cols = .Cols + 1
    .Col = .Col + 1
    .Text = GetResourceString(52)
    .CellFontBold = True
    Dim Grand As Currency
    
    rowno = .Row: colno = .Col
    
    For I = 0 To UBound(AccIDs) - 1
        '.Row = .Row + 1: .Text = FormatCurrency(Total(I) \ 1)
        rowno = rowno + 1
        .TextMatrix(rowno, colno) = FormatCurrency(Total(I) \ 1)
        'Grand = Grand + Val(.Text)
        Grand = Grand + Total(I)
    Next I
    'If .Rows > .Row Then .Rows = .Rows + 1
    '.Row = .Row + 1: .Text = FormatCurrency(Grand): .CellFontBold = True
    If .Rows > rowno Then .Rows = .Rows + 1
    rowno = rowno + 1
    .TextMatrix(rowno, colno) = FormatCurrency(Grand): .CellFontBold = True
End With
    
'Now Get the Account No & name of the accountHolder
''gDbTrans.SQLStmt = "Select AccId,AccNum, Name " & _
''    " FROM SBMaster A Inner Join QryNAme B On A.CustomerId = B.CustomerId " & _
''    " ORDER BY val(AccNum)"
''
''If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
''grd.Row = 1
''While Not Rst.EOF
''    With grd
''        .Row = .Row + 1
''        .Col = 0
''        If .Text <> Rst("AccID") Then GoTo NextRecord
''        .Text = Rst("AccNUm")
''        .Col = 1: .Text = FormatField(Rst("Name"))
''        If .Row = .Rows - 1 Then Rst.MoveLast
''    End With
''NextRecord:
''    Rst.MoveNext
''Wend
'9448822168
grd.MergeCells = flexMergeFree
grd.MergeCol(0) = True
grd.MergeCol(1) = True
grd.MergeRow(0) = True
grd.MergeRow(1) = True


lblReportTitle.Caption = GetResourceString(66, 47, 295) & " " & _
    GetFromDateString(m_FromIndianDate, m_ToIndianDate)

End Sub


'
Private Sub ShowSubDayBook()

Dim SqlStmt As String
Dim TmpStr As String
Dim rst As ADODB.Recordset
Dim transdate As Date
Dim SBClass As clsSBAcc

'.Clear
RaiseEvent Processing("Verifying records", 0)

SqlStmt = "Select A.Accid,AccNum,TransID,Particulars,Balance,TransDate, " & _
        " Amount, VoucherNo,TransType, Name " & _
        " From (SBTrans A Inner Join SBMaster B On A.AccID = B.AccID)" & _
        " Inner Join QryName C On B.CustomerId = C.CustomerId " & _
        " Where TransDate >= #" & m_FromDate & "# " & _
        " AND Transdate  <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

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
    Dim contraWithdraw As Currency, contraDeposit As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency
    Dim TotalContraWithDraw As Currency, TotalContraDeposit As Currency
    Dim TotalBankBalance As Currency
    Dim count As Integer
    Dim SlNo As Long
    Dim Amount As Currency
    
    With grd
        .Clear: .Cols = 10
        .FixedCols = 1
        .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35): .ColAlignment(2) = 2     '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        .Col = 5: .Text = GetResourceString(41) '"Voucher No
        .Col = 6: .Text = GetResourceString(271) '"Deposited"
        .Col = 7: .Text = GetResourceString(271) '"Deposited"
        .Col = 8: .Text = GetResourceString(279) '"Withdrawn"
        .Col = 9: .Text = GetResourceString(279) '"Withdrawn"
        '.Col = 10: .Text = GetResourceString(42) '"Balance"
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35): .ColAlignment(2) = 2     '"Name"
        .Col = 4: .Text = GetResourceString(39) 'Particulars
        .Col = 5: .Text = GetResourceString(41) '"Voucher No
        .Col = 6: .Text = GetResourceString(269) 'Cash
        .Col = 7: .Text = GetResourceString(270) 'Contra
        .Col = 8: .Text = GetResourceString(269) 'Cash
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
         RaiseEvent Initialise(0, rst.RecordCount)
         RaiseEvent Processing("Aliging the data ", 0)
        .ColAlignment(0) = 0
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        For SlNo = 0 To .Cols - 1
            .Col = SlNo
            .Row = 0: .CellAlignment = 4: .CellFontBold = True
            .Row = 1: .CellAlignment = 4: .CellFontBold = True
        Next
    End With
   
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: contraWithdraw = 0: contraDeposit = 0

Dim PrintSubTotal As Boolean
Dim rowno As Integer, colno As Integer

transdate = m_FromDate
rst.MoveFirst
transdate = rst("TransDate")
grd.Row = 1: SlNo = 0
rowno = 1
While Not rst.EOF
    With grd
        'Set next row
        If transdate <> rst("TransDate") Then
            PrintSubTotal = True
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
            .Col = 7: .CellFontBold = True: .Text = FormatCurrency(contraDeposit): .CellAlignment = 7
            .Col = 8: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
            .Col = 9: .CellFontBold = True: .Text = FormatCurrency(contraWithdraw): .CellAlignment = 7
            
            TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
            TotalContraDeposit = TotalContraDeposit + contraDeposit: TotalContraWithDraw = TotalContraWithDraw + contraWithdraw
            WithDraw = 0: Deposit = 0: contraWithdraw = 0: contraDeposit = 0
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
        End If
        
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00")
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("TransDate"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("AccNum")): .CellAlignment = 7
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rst("Name")): .CellAlignment = 1
        colno = 4: .TextMatrix(rowno, colno) = FormatField(rst("Particulars")): .CellAlignment = 1
        colno = 5: .TextMatrix(rowno, colno) = FormatField(rst("VoucherNo")): .CellAlignment = 4
        
        Dim transType As wisTransactionTypes
        transType = rst("TransType")
        Amount = FormatField(rst("Amount"))
        If transType = wDeposit Then
            colno = 6: Deposit = Deposit + Amount
        ElseIf transType = wContraDeposit Then
            colno = 7: contraDeposit = contraDeposit + Amount
        ElseIf transType = wWithdraw Then
            colno = 8: WithDraw = WithDraw + Amount
        ElseIf transType = wContraWithdraw Then
            colno = 9: contraWithdraw = contraWithdraw + Amount
        End If
        If Amount > 0 Then .TextMatrix(rowno, colno) = FormatCurrency(Amount)

    End With
    transdate = rst("TransDate")
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

lblReportTitle.Caption = GetResourceString(421, 390, 63) & " " & m_FromIndianDate

'Show sub Total
With grd
    .Row = rowno
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .CellAlignment = 4: .CellFontBold = True: .Text = GetResourceString(304) '"Sub Total "
    .Col = 6: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
    .Col = 7: .CellFontBold = True: .Text = FormatCurrency(contraDeposit): .CellAlignment = 7
    .Col = 8: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
    .Col = 9: .CellFontBold = True: .Text = FormatCurrency(contraWithdraw): .CellAlignment = 7
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    TotalContraDeposit = TotalContraDeposit + contraDeposit: TotalContraWithDraw = TotalContraWithDraw + contraWithdraw
        
'Show Grand Total
    If PrintSubTotal Then
        lblReportTitle.Caption = GetResourceString(421, 390, 63) & " " & m_FromIndianDate & " " _
            & GetResourceString(108) & " " & m_ToIndianDate

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


'
Private Sub ShowSubCashBook()
Dim SqlStmt As String
Dim TmpStr As String
Dim rst As ADODB.Recordset
Dim transdate As Date
Dim SBClass As clsSBAcc
''Clear
RaiseEvent Processing("Verifyinng records", 0)

SqlStmt = "Select A.Accid,AccNum,TransID,Particulars,Balance," & _
        " TransDate,Amount, VoucherNo,TransType, Name " & _
        " From (SBTrans A Inner Join SBMaster B On A.AccID = B.AccID)" & _
        " Inner Join QryName C On B.CustomerId = C.CustomerId " & _
        " Where TransDate >= #" & m_FromDate & "# " & _
        " AND Transdate  <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " AND Amount <= " & m_ToAmt

If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

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
    Dim contraWithdraw As Currency, contraDeposit As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency
    Dim TotalContraWithDraw As Currency, TotalContraDeposit As Currency
    Dim TotalBankBalance As Currency
    Dim count As Integer
    Dim SlNo As Long
    Dim Amount As Currency
    
    With grd
        .Clear: .Cols = 6
        .FixedCols = 1
        .FixedRows = 1 ' 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' Sl No"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35): .ColAlignment(2) = 2     '"Name"
        .Col = 4: .Text = GetResourceString(269) 'Cash
        .Col = 5: .Text = GetResourceString(269) 'Cash'

        .Col = 4: .Text = GetResourceString(271) '"Deposited"
        .Col = 5: .Text = GetResourceString(279) '"Withdrawn"

        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
         
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
         RaiseEvent Initialise(0, rst.RecordCount)
         RaiseEvent Processing("Aliging the data ", 0)
        .ColAlignment(0) = 0
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7

'        .ColAlignment(10) = 1
        For SlNo = 0 To .Cols - 1
            .Col = SlNo
            .Row = 0: .CellAlignment = 4: .CellFontBold = True
            '.Row = 1: .CellAlignment = 4: .CellFontBold = True
        Next
    End With
    
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: contraWithdraw = 0: contraDeposit = 0

Dim PrintSubTotal As Boolean
transdate = m_FromDate
rst.MoveFirst
transdate = rst("TransDate")
grd.Row = 0: SlNo = 0

On Error Resume Next
grd.Row = grd.FixedRows - 1
On Error GoTo 0

Dim rowno As Integer, colno As Integer
rowno = grd.Row
While Not rst.EOF
    With grd
        'Set next row
        If transdate <> rst("TransDate") Then
            PrintSubTotal = True
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1
            
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            .Col = 4: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
            .Col = 5: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
            
            TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
            Deposit = 0: WithDraw = 0
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            .Row = .Row + 1
        End If
        
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        colno = 0: .TextMatrix(rowno, colno) = Format(SlNo, "00")
        colno = 1: .TextMatrix(rowno, colno) = FormatField(rst("TransDate"))
        colno = 2: .TextMatrix(rowno, colno) = FormatField(rst("AccNum")): .CellAlignment = 7
        colno = 3: .TextMatrix(rowno, colno) = FormatField(rst("Name")): .CellAlignment = 1
        
        Dim transType As wisTransactionTypes
        transType = rst("TransType")
        Amount = FormatField(rst("Amount"))
        If transType = wDeposit Or transType = wContraDeposit Then
            colno = 4: Deposit = Deposit + Amount
        ElseIf transType = wWithdraw Or transType = wContraWithdraw Then
            colno = 5: WithDraw = WithDraw + Amount
        End If
        If Amount > 0 Then .TextMatrix(rowno, colno) = FormatCurrency(Amount)
        
        
        
    End With
    transdate = rst("TransDate")
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
  
Wend

lblReportTitle.Caption = GetResourceString(421, 390, 63) & " " & m_FromIndianDate

'Show sub Total
With grd
    .Row = rowno
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .CellAlignment = 4: .CellFontBold = True: .Text = GetResourceString(304) '"Sub Total "
    .Col = 4: .CellFontBold = True: .Text = FormatCurrency(Deposit): .CellAlignment = 7
    .Col = 5: .CellFontBold = True: .Text = FormatCurrency(WithDraw): .CellAlignment = 7
    
    TotalWithDraw = TotalWithDraw + WithDraw: TotalDeposit = TotalDeposit + Deposit
    
    'Show Grand Total
    If PrintSubTotal Then
             lblReportTitle.Caption = GetResourceString(62) & " " & _
             m_FromIndianDate & " " _
             & GetResourceString(108) & " " & m_ToIndianDate

        If .Rows <= .Row + 2 Then .Rows = .Row + 3
        .Row = .Row + 2
    End If
    
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) '"Sub Total "
    .Col = 4: .CellFontBold = True: .Text = FormatCurrency(TotalDeposit): .CellAlignment = 7
    .Col = 5: .CellFontBold = True: .Text = FormatCurrency(TotalWithDraw): .CellAlignment = 7
    
End With

End Sub

Private Sub cmdOk_Click()
'Call ShowGeneralLedger
Unload Me
End Sub



Private Sub cmdPrint_Click()

Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
    .ReportTitle = lblReportTitle
    .GridObject = grd
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

Private Sub Form_Activate()
    Call grd_LostFocus
End Sub

Private Sub Form_Click()
    Call grd_LostFocus
End Sub


Private Sub Form_Load()
'Set icon for the form caption
DoEvents
RaiseEvent Processing("Initailising ", 0)
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption
Me.lblReportTitle.FONTSIZE = 16
'Me.lblReportTitle.Caption = GetResourceString(421)
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

'Init the grid
With grd
    .Clear
    .MergeCells = flexMergeNever
    .Rows = 25
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    .Col = 0
    .ColWidth(0) = .Width - 100
    .Text = GetResourceString(278) '"No Records Available"
    .CellAlignment = 4: .CellFontBold = True
End With

    Screen.MousePointer = vbHourglass
    If m_ReportType = repSBJoint Then
        Call ShowJointAccounts
    ElseIf m_ReportType = repSBBalance Then
        Call ShowBalance
    ElseIf m_ReportType = repSBDayBook Then
        Call ShowSubDayBook
    ElseIf m_ReportType = repSBLedger Then
        Call ShowSBLedger
    ElseIf m_ReportType = repSBAccOpen Then
        Call ShowAccountsCreated
    ElseIf m_ReportType = repSBAccClose Then
        Call ShowAccountsClosed
    ElseIf m_ReportType = repSBProduct Then
        Call ShowProductsAndInterests
    ElseIf m_ReportType = repSBCheque Then
        Call ShowAccountsWithCheque
    ElseIf m_ReportType = repSbMonthlyBalance Then
        Call ShowMonthlyBalances
    ElseIf m_ReportType = repSBSubCashBook Then
        Call ShowSubCashBook
    End If

Screen.MousePointer = vbNormal

End Sub


Private Sub Form_Resize()
On Error Resume Next
    lblReportTitle.Top = 0
    lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
    fra.Top = Me.ScaleHeight - fra.Height
    fra.Left = Me.Width - fra.Width
    grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height - 100
    grd.Width = Me.ScaleWidth - 100
    cmdOK.Left = fra.Width - cmdOK.Width - (cmdOK.Width / 4)
    cmdPrint.Left = cmdOK.Left - cmdPrint.Width - (cmdPrint.Width / 4)
    cmdWeb.Top = cmdPrint.Top
    cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)
    
    'Dim wid As Single
    Dim ColCount As Integer
For ColCount = 0 To grd.Cols - 1
    With grd
        .ColWidth(ColCount) = GetSetting(App.EXEName, "SBReport" & m_ReportType, "ColWidth" & ColCount, 1 / .Cols) * .Width
        If .ColWidth(ColCount) >= .Width * 0.9 Then .ColWidth(ColCount) = .Width / 3
        If .ColWidth(ColCount) <= 0 Then .ColWidth(ColCount) = .Width / .Cols
    End With
Next


End Sub


Private Sub Form_Unload(Cancel As Integer)

RaiseEvent WindowClosed
End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
For ColCount = 0 To grd.Cols - 1
    Call SaveSetting(App.EXEName, "SBReport" & m_ReportType, _
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


