VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMMReport 
   Caption         =   "Members Report ..."
   ClientHeight    =   5790
   ClientLeft      =   1860
   ClientTop       =   1410
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   6480
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   5100
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   390
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   1950
         TabIndex        =   3
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   400
         Left            =   3780
         TabIndex        =   2
         Top             =   210
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4575
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8070
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
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
      Left            =   2310
      TabIndex        =   4
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmMMReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_FromIndianDate As String
Dim m_ToIndianDate As String
Dim m_FromDate As Date
Dim m_ToDate As Date

Dim m_FromAmt As String
Dim m_ToAmt As String

'To Get the Type Of Member
'Dim m_MemberType As Integer
Dim m_MemberType As wis_MemberType
Dim m_ReportType As wis_MemReports
Dim m_ReportOrder As wis_ReportOrder
Dim m_AccGroup As Integer

'TO Get Place or Caste
Dim m_Place As String
Dim m_Caste As String
Dim m_Gender As Byte
'To raise event
Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)


Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroup = NewValue
End Property


Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property

Public Property Let Gender(NewValue As wis_Gender)
    m_Gender = NewValue
End Property


Public Property Let memberTYpe(newMem As wis_MemberType)
    m_MemberType = newMem
End Property

Private Sub ShowShareCertificate()
Dim SqlStmt As String
Dim rst As Recordset

Dim Total As Currency

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Query Without customer name
SqlStmt = "Select C.CustomerId, C.AccNum, A.AccId, " & _
    " B.TransDate as IssueDate, FaceValue, CertNo,ReturnTransID " & _
    " from ShareTrans A Inner Join (MemTrans B " & _
    " Inner Join MemMaster C On C.AccId = B.AccID) " & _
    " On A.SaleTransID = B.TransID AND A.AccId = B.accID " & _
    " WHERE B.TransDate >= #" & m_FromDate & "#" & _
    " AND B.TransDate <= #" & m_ToDate & "#"

If m_FromAmt <> "" And m_FromAmt <> "0" Then SqlStmt = SqlStmt & " AND CertNo >= " & AddQuotes(m_FromAmt, True)
If m_ToAmt <> "" And m_ToAmt <> "0" Then SqlStmt = SqlStmt & " AND CertNo <= " & AddQuotes(m_ToAmt, True)

Dim sqlSupport As String
sqlSupport = ""
If m_Place <> "" Then sqlSupport = sqlSupport & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlSupport = sqlSupport & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then sqlSupport = sqlSupport & " AND Gender = " & m_Gender
If m_MemberType Then sqlSupport = sqlSupport & " and MemberType = " & m_MemberType
If m_AccGroup Then sqlSupport = sqlSupport & " AND AccGroupId = " & m_AccGroup

If Len(sqlSupport) > 0 Then
    sqlSupport = Mid(Trim$(sqlSupport), 4)
    sqlSupport = " AND C.CustomerID In " & _
        "(Select CustomerID From NameTab Where " & sqlSupport & ")"
End If

gDbTrans.SqlStmt = SqlStmt & sqlSupport & " order by B.TransDate"
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
SqlStmt = ""

RaiseEvent Initialise(0, rst.recordCount + 2)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
Dim SlNo As Long
Dim rowno As Long
grd.Row = grd.FixedRows
grd.MergeCells = flexMergeNever
rowno = grd.Row
SlNo = 1

While Not rst.EOF
    With grd
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("CertNo"))
        .TextMatrix(rowno, 2) = FormatField(rst("FaceValue"))
        .TextMatrix(rowno, 3) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 4) = FormatField(rst("IssueDate"))
        If FormatField(rst("ReturnTransID")) > 0 Then
            Dim rstTemp As Recordset
            gDbTrans.SqlStmt = "Select TransDate From MemTrans Where " & _
                " AccID = " & rst("AccID") & " And TransID = " & rst("ReturnTransID")
            If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                .TextMatrix(rowno, 5) = FormatField(rstTemp("TransDate"))
        End If
    End With
    
    SlNo = SlNo + 1
    If gCancel Then rst.MoveLast
    rst.MoveNext
    RaiseEvent Processing("Writing the records data ", SlNo / rst.recordCount)
    DoEvents
    
Wend

End Sub

Public Property Let ToAmount(curTo As String)
    m_ToAmt = curTo
End Property


Public Property Let FromAmount(curFrom As String)
    m_FromAmt = curFrom
End Property

Public Property Let ToIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_ToIndianDate = NewDate
        m_ToDate = GetSysFormatDate(NewDate)
        'm_ToIndianDate = GetAppFormatDate(m_ToDate)
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_FromIndianDate = NewDate
        m_FromDate = GetSysFormatDate(NewDate)
        'm_FromIndianDate = GetAppFormatDate(m_FromDate)
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

Public Property Let ReportType(RepType As wis_MemReports)
    m_ReportType = RepType
End Property

Private Sub InitGrid()
    
    Dim ColWid As Single
    Dim count As Long
    
With grd
    .Clear
    .MergeCells = flexMergeNever
    .Rows = 1
    .Rows = 15
    .Cols = 2
    .FixedCols = 0
    If m_ReportType = repMemBalance Or m_ReportType = repMemNonLoanMembers Or m_ReportType = repMemLoanMembers Then ''Show Balances
        .Cols = 4
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(49) & " " & _
               GetResourceString(60) '"Member No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(42) '"Balance"
        GoTo LastLine
    ElseIf m_ReportType = repMembers Then    ''Show Laibility
        .Cols = 5
        .Row = 0
        .FixedCols = 1: .FixedRows = 1
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(49) & " " & _
                GetResourceString(60) '"Member No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(37) '"Date"
        If m_Place <> "" Then
            .Cols = .Cols + 1
            .Col = .Col + 1: .Text = GetResourceString(270)
        End If
        If m_Caste <> "" Then
            .Cols = .Cols + 1
            .Col = .Col + 1
            .Text = GetResourceString(100)
        End If
        .Col = .Col + 1: .Text = GetResourceString(42) '"Balance"
        GoTo LastLine
    ElseIf m_ReportType = repMemDayBook Then    'ShowTransactions
        .Cols = 8: .Rows = 10
        .FixedCols = 1: .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' "Sl No"
        .Col = 1: .Text = GetResourceString(37) ' "Date"
        .Col = 2: .ColAlignment(1) = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(279) '"Withdrawn"
        .Col = 5: .Text = GetResourceString(279) '"Withdrawn"
        .Col = 6: .Text = GetResourceString(271) '"Deposited"
        .Col = 7: .Text = GetResourceString(271) '"Deposited"
'        .Col = 8: .Text = GetResourceString(53,191) 'Share/Meme Fe Fee
'        .Col = 9: .Text = GetResourceString(280) '"Balance"
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) ' "Sl No"
        .Col = 1: .Text = GetResourceString(37) ' "Date"
        .Col = 2: .ColAlignment(1) = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(269) '"Cash "
        .Col = 5: .Text = GetResourceString(270) '"Contra"
        .Col = 6: .Text = GetResourceString(269) '"Cash"
        .Col = 7: .Text = GetResourceString(270) '"Contra"
'        .Col = 8: .Text = GetResourceString(53,191) 'Share/Meme Fe Fee
'        .Col = 9: .Text = GetResourceString(280) '"Balance"
        
        GoTo LastLine
    
    ElseIf m_ReportType = repMemSubCashBook Then    'ShowTransactions
        .Cols = 6: .Rows = 10
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) ' "Sl No"
        .Col = 1: .Text = GetResourceString(37) ' "Date"
        .Col = 2: .ColAlignment(1) = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(271) '"Deposited"
        .Col = 5: .Text = GetResourceString(279) '"Withdrawn"
        
        GoTo LastLine
    
    ElseIf m_ReportType = repMemLedger Then   ''Show General Ledger
        
        .Clear: .Cols = 6
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) '"SL NO"
        .Col = 1: .Text = GetResourceString(37) '"Date"
        .Col = 2: .Text = GetResourceString(284) '"Opening Balance"
        .Col = 3: .Text = GetResourceString(302) '"Shares Issued"
        .Col = 4: .Text = GetResourceString(303) '"Share Returned"
        .Col = 5: .Text = GetResourceString(285) '"Closing Balance"
        GoTo LastLine
    ElseIf m_ReportType = repMemOpen Then      ''ShowAccounts Create
        .Cols = 4
        .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)  '"Sl No"
        .Col = 1: .ColAlignment(1) = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(281)  '"Create Date"
        GoTo LastLine
    ElseIf m_ReportType = repMemClose Then       ''ShowAccounts Closed
        .Cols = 4
        .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) '"Sl No"
        .Col = 1: .Text = GetResourceString(36, 60) '"Acc No"
        .Col = 2: .Text = GetResourceString(35) '"Name"
        .Col = 3: .Text = GetResourceString(282) '"Closed Date"
        GoTo LastLine
    ElseIf m_ReportType = repFeeCol Then
        .Cols = 5
        .Rows = 5
        .Row = 0
         .Col = 0: .Text = GetResourceString(38) + GetResourceString(37) '"Transaction Date"
         .Col = 1: .Text = GetResourceString(49) + GetResourceString(60) '"Member No"
         .Col = 2: .Text = GetResourceString(35) '"Name"
         .Col = 3: .ColAlignment(3) = 1: .Text = GetResourceString(53, 191) '"Share Fee"
         .Col = 4: .ColAlignment(4) = 1: .Text = GetResourceString(79, 191) '"Member Fee"
         GoTo LastLine
    
    ElseIf m_ReportType = repMemShareCert Then   ''Show share certificate
        .Cols = 6
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) '"SL No
        .Col = 1: .Text = GetResourceString(337, 60) '"Certificate No
        .Col = 2: .Text = GetResourceString(53, 140) '"Share value"
        .Col = 3: .Text = GetResourceString(49, 60) '"memeber No
        .Col = 4: .Text = GetResourceString(302)  '"Share issued "
        .Col = 5: .Text = GetResourceString(303)  '"Share Returned "
        '.Col = 6: .Text = GetResourceString(53,191) '"Share Fee"
        
        GoTo LastLine
    End If
End With

LastLine:

With grd
    Dim RowCount As Integer
    .MergeCells = flexMergeFree
    For RowCount = 0 To .FixedRows - 1
        .Row = RowCount
        .MergeRow(RowCount) = True
        For count = 0 To .Cols - 1
          .Col = count
          .CellAlignment = 4:  .CellFontBold = True
          .MergeCol(count) = True
        Next
    Next
End With

Exit Sub
ExitLine:
    ColWid = 0
    For count = 0 To grd.Cols - 2
        ColWid = ColWid + grd.ColWidth(count)
    Next count
    grd.ColWidth(grd.Cols - 1) = grd.Width - ColWid - grd.Width * 0.04 'Me.ScaleWidth * 0.03
    
End Sub

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

'chkDetails.Caption = GetResourceString(295)
cmdOk.Caption = GetResourceString(11)
cmdPrint.Caption = GetResourceString(23)
ErrLine:

End Sub

Private Sub ShowMembersCancelled()

Dim SqlStmt As String
Dim rst As Recordset
Dim NoAnd As Boolean   'Stupid variable

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Fire SQL
SqlStmt = "Select AccId,AccNum,ClosedDate,Name as CustName " & _
    " FROM MemMaster A Inner Join QryName B ON B.CustomerID = A.CustomerID " & _
    " WHERE ClosedDate >= #" & m_FromDate & "#" & _
    " AND ClosedDate  <= #" & m_ToDate & "#"

'Format the string to Get Name Details From NameTab
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_MemberType Then SqlStmt = SqlStmt & " and MemberType = " & m_MemberType
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

SqlStmt = SqlStmt & " ORDER BY ClosedDate, AccID"
    
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
           
'Initialize the grid
RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)
 
Dim count As Integer
Dim rowno As Long
Call InitGrid
rowno = grd.Row
Dim SlNo As Long

'Fill the grid
While Not rst.EOF
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = " " & FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = " " & FormatField(rst("CustName"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("ClosedDate"))
    End With

nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext

Wend

End Sub

Private Sub ShowMembersAdmitted()

Dim SqlStmt As String
Dim rst As Recordset
Dim NoAnd As Boolean

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Fire SQL
SqlStmt = "Select AccId,AccNum,CReateDate,Name as CustName " & _
    " FROM QryName A Inner JOin MemMaster B ON B.CustomerID = A.CustomerID " & _
    " WHERE CreateDate >= #" & m_FromDate & "#" & _
    " AND CreateDate  <= #" & m_ToDate & "#"
    
'Format the string to Get Name Details From NameTab
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_MemberType Then SqlStmt = SqlStmt & " and MemberType = " & m_MemberType
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup
    
gDbTrans.SqlStmt = SqlStmt & " order by CreateDate, AccNum"

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub

'Initialize the grid
   RaiseEvent Initialise(0, rst.recordCount)
   RaiseEvent Processing(" Aligning the data ", 0)
   
    Dim count As Integer
    Dim rowno As Long
    Call InitGrid
    
grd.MergeCells = flexMergeNever
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        count = count + 1
        .TextMatrix(rowno, 0) = count
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("CustName"))
        .TextMatrix(rowno, 3) = FormatField(rst("CreateDate"))
    End With
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then Exit Sub
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext

Wend
    
End Sub
Private Sub ShowBalances()
Dim SqlStmt As String
Dim rst As Recordset
Dim I As Long

Dim Total As Currency

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create the Supporting query
gDbTrans.SqlStmt = "Select MAX(TransID) as MaxTransID,AccID from MemTrans" & _
        " where TransDate <= #" & m_ToDate & "# GROUP By AccID"
Call gDbTrans.CreateView("qryMaxMemTransID")

SqlStmt = "SELECT A.CustomerID, Name,A.AccId,AccNum,Balance,ClosedDate " & _
    " from (MemMaster A  Inner Join qryName D On A.CustomerID=D.CustomerID)" & _
    " Inner Join (qryMaxMemTransID B Inner Join MemTrans C " & _
        " On  B.MaxTransID = C.TransID And B.AccID = C.AccID)" & _
    " On A.AccID =B.AccID"
    
Dim sqlClause As String
sqlClause = " And (ClosedDate >= #" & m_ToDate & "# OR ClosedDate is NULL )"

If Val(m_FromAmt) > 0 Then sqlClause = sqlClause & " AND Balance >= " & m_FromAmt
If Val(m_ToAmt) > 0 Then sqlClause = sqlClause & " AND Balance <= " & m_ToAmt

If m_Place <> "" Then sqlClause = sqlClause & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " AND Caste = " & AddQuotes(m_Caste, True)
If m_MemberType Then sqlClause = sqlClause & " and MemberType = " & m_MemberType
If m_Gender Then sqlClause = sqlClause & " AND Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " AND AccGroupId = " & m_AccGroup

sqlClause = Trim$(sqlClause)
If Len(sqlClause) Then sqlClause = " WHERE " & Mid(sqlClause, 4)

If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & sqlClause & " order by IsciName"
Else
    SqlStmt = SqlStmt & sqlClause & " order by val(A.AccNum)"
End If
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
SqlStmt = ""

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)
    
Dim count As Integer
Call InitGrid
    
grd.Row = grd.FixedRows - 1
I = 0
While Not rst.EOF
    
    'See if you have to show this record
    'If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
    
    'Set next row
    With grd
        If .Rows <= I + 2 Then .Rows = I + 2
        I = I + 1
        .TextMatrix(I, 0) = I
        .TextMatrix(I, 1) = FormatField(rst("AccNUm"))
        .TextMatrix(I, 2) = FormatField(rst("Name"))
        .TextMatrix(I, 3) = FormatField(rst("Balance"))
        
    End With
    Total = Total + FormatField(rst("Balance"))
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    
    rst.MoveNext
    
Wend

'Move next row
With grd
    grd.Row = I
    grd.MergeCells = flexMergeRestrictRows
    grd.ColAlignment(3) = 7
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(52, 42) '"Total Balances"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub

Private Sub ShowLoanMemberBalances()
Dim SqlStmt As String

Dim rst As Recordset
Dim I As Long

m_FromAmt = "0"
m_ToAmt = "0"
    
Dim Total As Currency

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create the Supporting query
gDbTrans.SqlStmt = "Select MAX(TransID) as MaxTransID,AccID from MemTrans" & _
        " where TransDate <= #" & m_ToDate & "# GROUP By AccID"
If DateDiff("d", m_ToDate, GetSysFormatDate("31/3/2011")) > 1 Then _
    gDbTrans.SqlStmt = "Select MAX(TransID) as MaxTransID,AccID from MemTrans GROUP By AccID"

Call gDbTrans.CreateView("qryMaxMemTransID")

SqlStmt = "SELECT distinct B.CustomerID as CustID, Name FROM LoanMaster A, " & _
    " QryMemName B, LoanTrans D WHERE B.MemID = A.MemID " & _
    " And TransID = (SELECT Max(transID) " & _
        " From LoanTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
    " AND D.LoanID = A.LoanID And Balance > 0"
SqlStmt = SqlStmt & " UNION " & _
    "SELECT distinct B.CustomerID as CustID, Name FROM BKCCMaster A, " & _
    " QryMemName B, BKCCTrans D WHERE B.MemID = A.MemID " & _
    " And TransID = (SELECT Max(transID) " & _
        " From BKCCTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
    " AND D.LoanID = A.LoanID And Balance > 0"

gDbTrans.SqlStmt = SqlStmt & " UNION " & _
    "SELECT distinct B.CustomerID as CustID, Name FROM DepositLoanMaster A, " & _
    " QryMemName B, DepositLoanTrans D WHERE B.CustomerID = A.CustomerID " & _
    " And TransID = (SELECT Max(transID) " & _
        " From DepositLoanTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
    " AND D.LoanID = A.LoanID And Balance > 0"
    
If DateDiff("d", m_ToDate, GetSysFormatDate("31/3/2011")) > 1 Then

    SqlStmt = "SELECT distinct B.CustomerID as CustID, Name FROM LoanMaster A, " & _
        " QryMemName B WHERE B.MemID = A.MemID "
    SqlStmt = SqlStmt & " UNION " & _
        "SELECT distinct B.CustomerID as CustID, Name FROM BKCCMaster A, " & _
        " QryMemName B, BKCCTrans D WHERE B.MemID = A.MemID " & _
        " And TransID = (SELECT Max(transID) " & _
        " From BKCCTrans E WHERE E.LoanId = A.LoanID and Balance > 0)" & _
        " AND D.LoanID = A.LoanID"
    
    gDbTrans.SqlStmt = SqlStmt & " UNION " & _
        "SELECT distinct B.CustomerID as CustID, Name FROM DepositLoanMaster A, " & _
        " QryMemName B, DepositLoanTrans D WHERE B.CustomerID = A.CustomerID " & _
        " And TransID = (SELECT Max(transID) " & _
        " From DepositLoanTrans E WHERE E.LoanId = A.LoanID and Balance > 0)" & _
        " AND D.LoanID = A.LoanID"


End If

Call gDbTrans.CreateView("qryLoanCustomers")

Dim sqlClause As String
sqlClause = ""
    SqlStmt = "SELECT I.CustomerID, J.Name,I.AccId,AccNum,Balance " & _
        " from (MemMaster I Inner Join qryLoanCustomers J On I.CustomerID=J.CustID)" & _
        " Inner Join (qryMaxMemTransID K Inner Join MemTrans L " & _
            " On  K.MaxTransID = L.TransID And K.AccID = L.AccID)" & _
        " On I.AccID =K.AccID "
    'sqlClause = sqlClause & " AND Balance > " & Val(m_FromAmt)
    
sqlClause = " AND I.ClosedDate Is NULL AND Balance > 0 "
If Val(m_FromAmt) > 0 Then sqlClause = sqlClause & " AND Balance >= " & m_FromAmt
If Val(m_ToAmt) > 0 Then sqlClause = sqlClause & " AND Balance <= " & m_ToAmt

If m_Place <> "" Then sqlClause = sqlClause & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " AND Caste = " & AddQuotes(m_Caste, True)
If m_MemberType Then sqlClause = sqlClause & " and MemberType = " & m_MemberType
If m_Gender Then sqlClause = sqlClause & " AND Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " AND AccGroupId = " & m_AccGroup

sqlClause = Trim$(sqlClause)
If Len(sqlClause) Then sqlClause = " WHERE " & Mid(sqlClause, 4)

If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & sqlClause & " order by IsciName"
Else
    SqlStmt = SqlStmt & sqlClause & " order by val(I.AccNum)"
End If
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
SqlStmt = ""

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)
    
Dim count As Integer
Call InitGrid
    
grd.Row = grd.FixedRows - 1
I = 0
While Not rst.EOF
    
    'See if you have to show this record
    'If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
    
    'Set next row
    With grd
        If .Rows <= I + 2 Then .Rows = I + 2
        I = I + 1
        .TextMatrix(I, 0) = I
        .TextMatrix(I, 1) = FormatField(rst("AccNUm"))
        .TextMatrix(I, 2) = FormatField(rst("Name"))
        .TextMatrix(I, 3) = FormatField(rst("Balance"))
        
    End With
    Total = Total + FormatField(rst("Balance"))
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    
    rst.MoveNext
    
Wend

'Move next row
With grd
    grd.Row = I
    grd.MergeCells = flexMergeRestrictRows
    grd.ColAlignment(3) = 7
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(52, 42) '"Total Balances"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub

Private Sub ShowNonLoanMembers()
Dim SqlStmt As String
Dim rst As Recordset
Dim rstLoan As Recordset
Dim I As Long
Dim Total As Currency
m_FromAmt = "0"
m_ToAmt = "0"

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create the Supporting query
gDbTrans.SqlStmt = "Select MAX(TransID) as MaxTransID,AccID from MemTrans" & _
        " where TransDate <= #" & m_ToDate & "# GROUP By AccID"
If DateDiff("d", m_ToDate, GetSysFormatDate("31/3/2011")) > 1 Then _
    gDbTrans.SqlStmt = "Select MAX(TransID) as MaxTransID,AccID from MemTrans GROUP By AccID"

Call gDbTrans.CreateView("qryMaxMemTransID")

If DateDiff("d", m_ToDate, GetSysFormatDate("31/3/2011")) > 1 Then
    SqlStmt = "SELECT distinct B.CustomerID as CustID, Name FROM LoanMaster A, " & _
        " QryMemName B WHERE B.MemID = A.MemID "
    
    gDbTrans.SqlStmt = SqlStmt & " UNION " & _
    "SELECT distinct B.CustomerID as CustID, Name FROM BKCCMaster A, " & _
        " QryMemName B, BKCCTrans D WHERE B.MemID = A.MemID " & _
        " And TransID = (SELECT Max(transID) " & _
            " From BKCCTrans E WHERE E.LoanId = A.LoanID and balance > 0)" & _
        " AND D.LoanID = A.LoanID "
    
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " UNION " & _
        "SELECT distinct B.CustomerID as CustID, Name FROM DepositLoanMaster A, " & _
        " QryMemName B, DepositLoanTrans D WHERE B.CustomerID = A.CustomerID " & _
        " And TransID = (SELECT Max(transID) " & _
            " From DepositLoanTrans E WHERE E.LoanId = A.LoanID and balance > 0)" & _
        " AND D.LoanID = A.LoanID "
    

Else

    SqlStmt = "SELECT distinct B.CustomerID as CustID, Name FROM LoanMaster A, " & _
        " QryMemName B, LoanTrans D WHERE B.MemID = A.MemID " & _
        " And TransID = (SELECT Max(transID) " & _
            " From LoanTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
        " AND D.LoanID = A.LoanID And Balance > 0"
    gDbTrans.SqlStmt = SqlStmt & " UNION " & _
        "SELECT distinct B.CustomerID as CustID, Name FROM BKCCMaster A, " & _
        " QryMemName B, BKCCTrans D WHERE B.MemID = A.MemID " & _
        " And TransID = (SELECT Max(transID) " & _
            " From BKCCTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
        " AND D.LoanID = A.LoanID And Balance > 0"
        
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " UNION " & _
        "SELECT distinct B.CustomerID as CustID, Name FROM DepositLoanMaster A, " & _
        " QryMemName B, DepositLoanTrans D WHERE B.CustomerID = A.CustomerID " & _
        " And TransID = (SELECT Max(transID) " & _
            " From DepositLoanTrans E WHERE E.LoanId = A.LoanID and TransDate <= #" & m_ToDate & "#)" & _
        " AND D.LoanID = A.LoanID And Balance > 0"
End If
Call gDbTrans.Fetch(rstLoan, adOpenDynamic)
    
    
SqlStmt = "SELECT A.CustomerID, Name,A.AccId,AccNum,Balance " & _
    " from (MemMaster A  Inner Join qryName D On A.CustomerID=D.CustomerID)" & _
    " Inner Join (qryMaxMemTransID B Inner Join MemTrans C " & _
        " On  B.MaxTransID = C.TransID And B.AccID = C.AccID)" & _
    " On A.AccID =B.AccID "
    
Dim sqlClause As String
sqlClause = ""

sqlClause = " AND A.ClosedDate Is NULL AND Balance > 0 "
If m_Place <> "" Then sqlClause = sqlClause & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " AND Caste = " & AddQuotes(m_Caste, True)
If m_MemberType Then sqlClause = sqlClause & " and MemberType = " & m_MemberType
If m_Gender Then sqlClause = sqlClause & " AND Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " AND AccGroupId = " & m_AccGroup

sqlClause = Trim$(sqlClause)
If Len(sqlClause) Then sqlClause = " WHERE " & Mid(sqlClause, 4)

If m_ReportOrder = wisByName Then
    SqlStmt = SqlStmt & sqlClause & " order by IsciName"
Else
    SqlStmt = SqlStmt & sqlClause & " order by val(A.AccNum)"
End If
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
SqlStmt = ""

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)
    
Dim count As Integer
Call InitGrid
    
grd.Row = grd.FixedRows - 1
I = 0
While Not rst.EOF
    
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        rstLoan.Find "CustID = " & FormatField(rst("CustomerID"))
        If rstLoan.EOF = False Then GoTo nextRecord
    End If
    
    'See if you have to show this record
    'If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
    
    'Set next row
    With grd
        If .Rows <= I + 2 Then .Rows = I + 2
        I = I + 1
        .TextMatrix(I, 0) = I
        .TextMatrix(I, 1) = FormatField(rst("AccNUm"))
        .TextMatrix(I, 2) = FormatField(rst("Name"))
        .TextMatrix(I, 3) = FormatField(rst("Balance"))
        
    End With
    Total = Total + FormatField(rst("Balance"))
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    
    rst.MoveNext
    
Wend

'Move next row
With grd
    grd.Row = I
    grd.MergeCells = flexMergeRestrictRows
    grd.ColAlignment(3) = 7
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(52, 42) '"Total Balances"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub
Private Sub ShowMemberAndShareFee()
'Declare variables
Dim SqlStmt As String
Dim TmpStr As String
Dim rst As Recordset
Dim TransDate As Date
Dim AccId As Long

RaiseEvent Processing("Reading & Verifying the data ", 0)

SqlStmt = "Select B.AccId,AccNum,Amount,TransDate,TransId,TransType," & _
    " Name as CustName from QryName A Inner Join (MemMaster B" & _
    " Inner Join MemIntTrans C ON C.AccId = B.AccId) " & _
    " ON B.CustomerId = A.CustomerId WHERE Amount > 0 " & _
    " AND TransDate >= #" & m_FromDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#"
    
'Format the string to Get Name Details From NameTab
If m_Place <> "" Then SqlStmt = SqlStmt & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStmt = SqlStmt & " AND Gender = " & m_Gender
If m_MemberType Then SqlStmt = SqlStmt & " and MemberType = " & m_MemberType
If m_AccGroup Then SqlStmt = SqlStmt & " AND AccGroupId = " & m_AccGroup

'Get The Share Fee record Set
gDbTrans.SqlStmt = SqlStmt: SqlStmt = ""
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub
    
TransDate = rst("TransDate")
AccId = rst("AccId")

' initialise the grid
Dim MemFee As Currency, ShareFee As Currency
Dim GrandMemFee  As Currency, GrandShareFee As Currency
Dim count As Integer
Dim rowno As Long, colno As Byte

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid

Dim PRINTTotal As Boolean
'Fill in to the grid
grd.Row = grd.FixedRows
rowno = grd.Row

Do
    If rst.EOF = True Then Exit Do
    With grd
        
        If DateDiff("d", TransDate, rst("TransDate")) <> 0 Then
            PRINTTotal = True
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1: .Row = rowno
            .Col = 0: .Text = GetIndianDate(TransDate)
            .CellAlignment = 4: .CellFontBold = True
            .Col = 2: .Text = GetResourceString(52) '"Sub Total"
            .CellAlignment = 4: .CellFontBold = True
            .Col = 3: .Text = FormatCurrency(ShareFee): .CellAlignment = 7
            GrandShareFee = GrandShareFee + ShareFee: ShareFee = 0
            .CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(MemFee)
            GrandMemFee = GrandMemFee + MemFee: MemFee = 0
            .CellAlignment = 7: .CellFontBold = True
        End If
    
        If .Rows <= rowno + 2 Then .Rows = .Rows + 2
        rowno = rowno + 1
        AccId = rst("AccId"): TransDate = rst("TransDate")
        .TextMatrix(rowno, 0) = FormatField(rst("TransDate"))
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("CustName"))
        If Val(rst("TransId")) = 1 Then
            .TextMatrix(rowno, 4) = FormatField(rst("Amount"))
            MemFee = MemFee + Val(rst("Amount"))
        Else
            .TextMatrix(rowno, 3) = FormatField(rst("Amount"))
            ShareFee = ShareFee + Val(rst("Amount"))
        End If
        
    End With
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext
Loop

With grd
    .MergeCells = flexMergeNever
    .ColAlignment(3) = 7: .ColAlignment(4) = 7
    If .Rows <= rowno + 2 Then .Rows = .Rows + 2
    rowno = rowno + 1
    .Row = rowno
    .Col = 0: .Text = GetIndianDate(TransDate)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(52) '"Sub Total"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(ShareFee)
    GrandShareFee = GrandShareFee + ShareFee: ShareFee = 0
    .CellAlignment = 7: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(MemFee)
    GrandMemFee = GrandMemFee + MemFee: MemFee = 0
    .CellAlignment = 7: .CellFontBold = True
End With

If PRINTTotal Then
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 3
        rowno = rowno + 2
        .Row = rowno
        .Col = 2: .Text = GetResourceString(286) '"Grand Total"
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = FormatCurrency(GrandShareFee):
        .CellAlignment = 7: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(GrandMemFee): MemFee = 0
        .CellAlignment = 7: .CellFontBold = True
    End With
End If

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
'Get the LAst of the April Month
fromDate = GetSysLastDate(FinUSFromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463, 67, 42) & " " & _
    GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))

SqlStmt = "SELECT A.AccNum,A.AccID, A.CustomerID,Name as CustName " & _
    " From MemMaster A Inner Join QryName B On B.CustomerID = A.CustomerID" & _
    " WHERE A.CreateDate <= #" & toDate & "#" & _
    " AND (A.ClosedDate Is NULL OR A.Closeddate >= #" & fromDate & "#)" & _
    " "

SqlStmt = SqlStmt & " Order By " & _
    IIf(m_ReportOrder = wisByAccountNo, "val(A.ACCNUM)", "IsciName")
    
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub
SqlStmt = ""

count = DateDiff("M", fromDate, toDate) + 2
totalCount = (count + 1) * rstMain.recordCount
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
End With

grd.Row = 0: count = 0
Dim rowno As Long

While Not rstMain.EOF
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        .TextMatrix(rowno, 0) = rowno
        .TextMatrix(rowno, 1) = FormatField(rstMain("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rstMain("CustNAme"))
        .RowData(rowno) = 0
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
    
    rstMain.MoveNext
Wend
    
    With grd
        rowno = rowno + 2
        If .Rows <= rowno + 3 Then .Rows = rowno + 3
        .Row = rowno
        .RowData(rowno - 1) = 1
        .RowData(rowno) = 1
        .Col = 2: .Text = GetResourceString(286) 'Grand Total
        .CellFontBold = True
        .Rows = .Row + 1
    End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

fromDate = "4/30/" & Year(fromDate)
Do
    
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    
    SqlStmt = "SELECT AccId, Max(TransID) AS MaxTransID" & _
            " FROM MemTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY AccID"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("MemMonBal")
    
    SqlStmt = "Select C.AccNum, A.AccID, Balance From MemMonBal A " & _
        " Inner Join (MemTrans B Inner Join " & _
        " (MemMaster C "
    If m_ReportOrder = wisByName Then _
        SqlStmt = SqlStmt & " Inner Join QryName D On D.CustomerID = C.CustomerID"
    
    SqlStmt = SqlStmt & ") On C.AccID  = B.AccID ) " & _
        " On A.AccID = B.AccID and B.TransID = A.MaxTransID "
    
    ''Add the Order By Condition
    SqlStmt = SqlStmt & " Order By " & _
        IIf(m_ReportOrder = wisByAccountNo, "val(C.ACCNUM)", "IsciName")
    gDbTrans.SqlStmt = SqlStmt
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0
        .Col = .Cols - 1
        .Text = GetMonthString(Month(fromDate)) & " " & GetResourceString(42)
        .Col = .Cols - 1
        rowno = 0
    End With
    
    TotalBalance = 0
    
    ''Now Put the Montly Balance to the respective rows
    While rowno < grd.Rows - 1
        rowno = rowno + 1
        rstBalance.MoveFirst
        rstBalance.Find "AccNum = '" & grd.TextMatrix(rowno, 1) & "'", , adSearchForward
        If rstBalance.EOF Then GoTo NextAccount
        With grd
            .TextMatrix(rowno, .Col) = FormatField(rstBalance("Balance"))
            .RowData(rowno) = 1
        End With
        Balance = rstBalance("Balance")
        TotalBalance = TotalBalance + Balance
        
        DoEvents
        If gCancel Then fromDate = toDate: rowno = grd.Rows
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
        
NextAccount:
        ProcCount = ProcCount + 1
    Wend
    
    With grd
        .Row = .Rows - 1
        .Text = FormatCurrency(TotalBalance)
        .CellFontBold = True
    End With
    
NextMonth:

    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

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

Private Sub ShowMemLedger()
Dim SqlPrin As String
Dim SqlInt As String

'Prelim check
RaiseEvent Processing("Reading & Verifying the data", 0)

'Build the SQL
SqlPrin = "SELECT 'SHARE', Sum(Amount) as TotalAmount,TransDate, " & _
    " TransType From MemTrans WHERE TransID > 0" & _
    " AND TransDate >= #" & m_FromDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#"
SqlInt = "SELECT 'FEE', Sum(amount) as TotalAmount,TransDate, " & _
    " TransType From MemIntTrans WHERE TransId > 1" & _
    " AND TransDate >= #" & m_FromDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#"

If m_MemberType Then
    'SqlPrin = SqlPrin & " AND MemberType  <= " & m_MemberType
    'SqlInt = SqlInt & " AND MemberType  <= " & m_MemberType
End If

SqlPrin = SqlPrin & " GROUP BY TransDate,TransType"
SqlInt = SqlInt & " GROUP BY TransDate,TransType"

gDbTrans.SqlStmt = SqlPrin & " UNION " & SqlInt & " Order by TransDate"

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim OpeningBalance As Currency
Dim TransDate As Date
'Initialize the grid
    Dim SubTotal As Currency, GrandTotal As Currency
    Dim WithDraw As Currency, Deposit As Currency, Interest As Currency, Charges As Currency
    Dim TotalWithDraw As Currency, TotalDeposit As Currency, TotalInterest As Currency, TotalCharges As Currency
    Dim Balance As Currency
'Fire the query
   RaiseEvent Initialise(0, rst.recordCount)
   RaiseEvent Processing("Aligning the data ", 0)
   
   'Get liability on a day before m_fromdate
    OpeningBalance = ComputeTotalMMLiability(DateAdd("d", -1, m_ToDate))
    
    Dim count As Integer
    Dim SlNo As Integer
    Dim rowno As Long
    
    Call InitGrid
    grd.Row = 0
    SubTotal = 0: GrandTotal = 0
    WithDraw = 0: Deposit = 0: Interest = 0: Charges = 0

With grd
    .MergeCells = flexMergeNever
    .Row = .FixedRows
    
    .Col = 1: .Text = GetResourceString(284) '"Opening Balance"
    .CellAlignment = 4: .CellFontBold = True
    
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 4: .CellFontBold = True
    .ColAlignment(3) = 7
    .ColAlignment(4) = 7
    .ColAlignment(5) = 7
    
End With

TransDate = rst("TransDate")
Dim TotalPrint As Boolean
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    If TransDate <> rst("TransDate") Then
        TotalPrint = True
        With grd
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1: SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = FormatCurrency(OpeningBalance)
            .TextMatrix(rowno, 3) = FormatCurrency(Deposit)
            .TextMatrix(rowno, 4) = FormatCurrency(WithDraw)
            OpeningBalance = OpeningBalance + Deposit - WithDraw
            .TextMatrix(rowno, 5) = FormatCurrency(OpeningBalance)
            TotalWithDraw = TotalWithDraw + WithDraw
            TotalDeposit = TotalDeposit + Deposit
            TotalCharges = TotalCharges + Charges
            
        End With
        TransDate = rst("TransDate")
        WithDraw = 0: Deposit = 0: Interest = 0: Charges = 0
    End If
    
    Dim transType As wisTransactionTypes
    transType = FormatField(rst("TransType"))
    
    If rst(0) = "SHARE" Then
        If transType = wWithdraw Or transType = wContraWithdraw Then
            WithDraw = WithDraw + FormatField(rst("TotalAmount"))
        ElseIf transType = wDeposit Or transType = wContraDeposit Then
            Deposit = Deposit + FormatField(rst("TotalAmount"))
        End If
    Else
        If transType = wDeposit Or transType = wContraDeposit Then
            Charges = Charges + FormatField(rst("TotalAmount"))
        End If
    End If
    
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    
    rst.MoveNext
Wend
With grd
    If .Rows = rowno + 1 Then .Rows = .Rows + 1
    rowno = rowno + 1: SlNo = SlNo + 1
    .Row = rowno
    .Col = 0: .Text = SlNo
    .Col = 1: .Text = TransDate
    .Col = 2: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    .Col = 3: .CellAlignment = 7: .Text = FormatCurrency(Deposit)
    .Col = 4: .CellAlignment = 7: .Text = FormatCurrency(WithDraw)
    OpeningBalance = OpeningBalance + Deposit - WithDraw
    .Col = 5: .CellAlignment = 7: .Text = FormatCurrency(OpeningBalance)
    
    TotalWithDraw = TotalWithDraw + WithDraw
    TotalDeposit = TotalDeposit + Deposit
    TotalCharges = TotalCharges + Charges
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(285)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 4: .CellFontBold = True
    
    If TotalPrint Then
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = FormatCurrency(TotalDeposit)
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(TotalWithDraw)
        .CellAlignment = 4: .CellFontBold = True
    End If
End With
End Sub

Private Sub ShowLiability()
Dim SqlStmt As String
Dim rst As Recordset
Dim I As Long
Dim Total As Currency

RaiseEvent Processing("Reading & Verifying the data ", 0)

'Create a view to get max transid
SqlStmt = "Select max(TransId) as MaxTransID,AccID " & _
    " from MemTrans Group By AccId "
gDbTrans.SqlStmt = SqlStmt
gDbTrans.CreateView ("MemMaxTransID")

SqlStmt = "Select B.AccId,AccNum,CreateDate,Balance, Place, Caste, Name " & _
    " from qryName A Inner join (MemMaster B Inner Join (Memtrans C" & _
    " Inner Join MemMaxTransID D ON D.MaxTransID = C.TransID AND C.AccID = D.AccID)" & _
    " On  C.AccID = B.AccID) ON  B.CustomerID = A.CustomerID"

'Query on Caste,Place,Gender
Dim sqlClause As String
If m_Place <> "" Then sqlClause = sqlClause & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " And Caste = " & AddQuotes(m_Caste, True)
If m_MemberType Then sqlClause = sqlClause & " and MemberType = " & m_MemberType
If m_Gender Then sqlClause = sqlClause & " And Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " And AccGroupId = " & m_AccGroup
If Len(sqlClause) > 0 Then
    sqlClause = Trim$(sqlClause)
    'Replcace first 'and' with 'where'
    sqlClause = " WHERE " & Mid(sqlClause, 4)
End If
'Query Order by
If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & sqlClause & " order by  val(AccNum) "
Else
    SqlStmt = SqlStmt & sqlClause & " order by IsciName,val(B.AccNum)"
End If
    
gDbTrans.SqlStmt = SqlStmt: SqlStmt = ""
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Alignign the data ", 0)

Dim count As Integer
Dim rowno As Long, colno As Byte

Call InitGrid
grd.MergeCells = flexMergeNever
rowno = grd.Row

Dim AccId As Long
count = 0
While Not rst.EOF
    DoEvents
    Me.Refresh
    'See if you have to show this record
    If AccId = FormatField(rst("AccID")) Then GoTo nextRecord
   
    'Set next row
    With grd
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1
        count = count + 1
        .TextMatrix(rowno, 0) = count
        .TextMatrix(rowno, 1) = " " & FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("CreateDate"))
        colno = 3
        If m_Place <> "" Then
            colno = colno + 1
            .TextMatrix(rowno, colno) = FormatField(rst("Place"))
        End If
        If m_Caste <> "" Then
            colno = colno + 1
            .TextMatrix(rowno, colno) = FormatField(rst("Caste"))
        End If
        colno = colno + 1
        .TextMatrix(rowno, colno) = FormatField(rst("Balance"))
        
    End With
    AccId = FormatField(rst("AccID"))
    Total = Total + FormatField(rst("Balance"))
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext
Wend

grd.ColAlignment(grd.Cols - 1) = flexAlignRightCenter
'Set next row and print grand total
With grd
    rowno = rowno + 2
    If .Rows <= rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    .Col = 2: .Text = GetResourceString(286) '"Grand Total"
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub

Private Sub ShowDayBook()

Dim PrinSql As String
Dim IntSql As String
Dim rst As Recordset
Dim TransDate As Date

RaiseEvent Processing("Reading & Verifying the data ", 0)
PrinSql = "Select 'SHARE' , B.AccId,AccNum, TransID,Amount,Balance," & _
    " TransDate,TransType, Name as CustName from MemTrans A " & _
    " Inner Join (MemMaster B Inner Join QryName C ON C.CustomerId = B.CustomerId )" & _
    " On B.AccID = A.AccID Where Amount > 0 " & _
    " AND TransDate >= #" & m_FromDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#"


'Format the string to Get Name Details From NameTab
If m_Place <> "" Then PrinSql = PrinSql & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then PrinSql = PrinSql & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then PrinSql = PrinSql & " AND Gender = " & m_Gender
If m_MemberType Then PrinSql = PrinSql & " and MemberType = " & m_MemberType
If m_AccGroup Then PrinSql = PrinSql & " AND AccGroupId = " & m_AccGroup

If Val(m_FromAmt) <> 0 Then PrinSql = PrinSql & " AND Amount >= " & m_FromAmt
If Val(m_ToAmt) <> 0 Then PrinSql = PrinSql & " AND Amount <= " & m_ToAmt

gDbTrans.SqlStmt = PrinSql & " ORDER BY TransDate,A.AccID, TransID"
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid
Dim SubTotal(4 To 8) As Currency, GrandTotal(4 To 8) As Currency

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid

grd.MergeCol(4) = False
grd.MergeCol(5) = False

Dim transType As wisTransactionTypes
Dim TransID As Long
Dim AccId As Long
Dim SlNo As Long
Dim PRINTTotal As Boolean
Dim count  As Integer
Dim rowno As Long, colno As Byte

TransDate = rst("TransDate")
grd.Row = grd.FixedRows
rowno = grd.Row

While Not rst.EOF
    With grd
        If TransDate <> rst("TransDate") Then
            TransDate = rst("TransDate")
            PRINTTotal = True
            .Row = rowno
            If .Rows = .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1: SlNo = 0: TransID = 0
            .Col = 3: .Text = GetResourceString(52) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            rowno = .Row
            For count = 4 To 7
                .Col = count
                .TextMatrix(rowno, count) = FormatCurrency(SubTotal(count))
                .CellFontBold = True
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End If
        
        If AccId <> rst("AccId") Then TransID = 0
        AccId = rst("AccId")
        'Set next row
        If TransID <> rst("TransID") Then
            TransID = rst("TransID")
            rowno = rowno + 1
            If .Rows = rowno + 1 Then .Rows = rowno + 2
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = " " & FormatField(rst("AccNum"))
            .TextMatrix(rowno, 3) = FormatField(rst("CustName"))
        End If
        
        If FormatField(rst("Amount")) <> 0 Then
            transType = rst("TransType")
            If transType = wWithdraw Then colno = 4
            If transType = wContraWithdraw Then colno = 5
            If transType = wDeposit Then colno = 6
            If transType = wContraDeposit Then colno = 7
            .TextMatrix(rowno, colno) = FormatField(rst("Amount"))
            SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
            
        End If
    End With
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext
Wend

'Show grand Total
With grd
    If .Rows <= rowno + 1 Then .Rows = .Rows + 1
    .Row = rowno
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(52) '"Sub Total "
    .CellAlignment = 4: .CellFontBold = True
    For SlNo = 4 To 7
        .Col = SlNo
        .Text = FormatCurrency(SubTotal(SlNo))
        .CellFontBold = True
        GrandTotal(SlNo) = GrandTotal(SlNo) + SubTotal(SlNo)
        SubTotal(SlNo) = 0
    Next
    
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(286) '"Grand Total "
        .CellAlignment = 4: .CellFontBold = True
        For SlNo = 4 To 7
            .Col = SlNo
            .Text = FormatCurrency(GrandTotal(SlNo))
            .CellFontBold = True
        Next
    End If
End With

End Sub


Private Sub ShowSubCashBook()

Dim PrinSql As String
Dim IntSql As String
Dim rst As Recordset
Dim TransDate As Date

RaiseEvent Processing("Reading & Verifying the data ", 0)

PrinSql = "Select 'SHARE',B.AccId,AccNum,TransID,Amount,Balance, " & _
    " TransDate,TransType, Name as CustName from " & _
    " QryName A Inner Join (MemMaster B Inner Join " & _
    " MemTrans C On C.AccId = B.AccId) ON B.CustomerId = A.CustomerId " & _
    " Where Amount > 0 " & _
    " AND TransDate >= #" & m_FromDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#"

'Format the string to Get Name Details From NameTab
If m_Place <> "" Then PrinSql = PrinSql & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then PrinSql = PrinSql & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then PrinSql = PrinSql & " AND Gender = " & m_Gender
If m_MemberType Then PrinSql = PrinSql & " and MemberType = " & m_MemberType
If m_AccGroup Then PrinSql = PrinSql & " AND AccGroupId = " & m_AccGroup

If Val(m_FromAmt) <> 0 Then PrinSql = PrinSql & " AND Amount >= " & m_FromAmt
If Val(m_ToAmt) <> 0 Then PrinSql = PrinSql & " AND Amount <= " & m_ToAmt

gDbTrans.SqlStmt = PrinSql & " ORDER BY TransDate,B.AccID, TransID"
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid
Dim SubTotal(4 To 5) As Currency, GrandTotal(4 To 5) As Currency

RaiseEvent Initialise(0, rst.recordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid

Dim transType As wisTransactionTypes
Dim TransID As Long
Dim AccId As Long
Dim SlNo As Long
Dim PRINTTotal As Boolean
Dim count  As Integer
Dim rowno As Long

TransDate = rst("TransDate")
grd.Row = grd.FixedRows
grd.MergeCells = flexMergeNever
rowno = grd.Row

While Not rst.EOF
    With grd
        If TransDate <> rst("TransDate") Then
            TransDate = rst("TransDate")
            PRINTTotal = True
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1: SlNo = 0: TransID = 0
            .Row = rowno
            .Col = 3: .Text = GetResourceString(52) '"Sub Total "
            .CellAlignment = 4: .CellFontBold = True
            For count = 4 To 5
                .Col = count
                .Text = FormatCurrency(SubTotal(count))
                .CellFontBold = True
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End If
        
        If AccId <> rst("AccId") Then TransID = 0
        AccId = rst("AccId")
        'Set next row
        If TransID <> rst("TransID") Then
            TransID = rst("TransID")
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = " " & FormatField(rst("AccNum"))
            .TextMatrix(rowno, 3) = FormatField(rst("CustName"))
        End If
        
        If FormatField(rst("Amount")) <> 0 Then
            Dim colno As Byte
            transType = rst("TransType")
            colno = 4
            If transType = wWithdraw Or transType = wContraWithdraw Then colno = 5
            .TextMatrix(rowno, colno) = FormatField(rst("Amount"))
            SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
        End If
    End With
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.recordCount)
    rst.MoveNext
Wend

'Show grand Total
With grd
    .Row = rowno
    If .Rows <= .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(52) '"Sub Total "
    .CellAlignment = 4: .CellFontBold = True
    For SlNo = 4 To 5
        .Col = SlNo
        .Text = FormatCurrency(SubTotal(SlNo))
        .CellFontBold = True
        GrandTotal(SlNo) = GrandTotal(SlNo) + SubTotal(SlNo)
        SubTotal(SlNo) = 0
    Next
    
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(286) '"Grand Total "
        .CellAlignment = 4: .CellFontBold = True
        For SlNo = 4 To 5
            .Col = SlNo
            .Text = FormatCurrency(GrandTotal(SlNo))
            .CellFontBold = True
        Next
    End If
End With

End Sub

'
Private Sub cmdOk_Click()
Unload Me
End Sub



Private Sub cmdPrint_Click()
With wisMain.grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
    .GridObject = Me.grd
    .ReportTitle = lblReportTitle
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

lblReportTitle.FONTSIZE = 14
lblReportTitle.Caption = GetResourceString(49) & " " & _
    GetResourceString(283) & GetResourceString(92)

'Init the grid
With grd
    .Clear
    .Rows = 50
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    .Text = GetResourceString(278) ' "No Records Available"
    .CellFontBold = True: .CellAlignment = 4
End With

Screen.MousePointer = vbHourglass

    If m_ReportType = repMemBalance Then
         lblReportTitle.Caption = GetResourceString(61)
         Call ShowBalances
    ElseIf m_ReportType = repMembers Then
        lblReportTitle.Caption = GetResourceString(61)
        Call ShowLiability
    ElseIf m_ReportType = repMemSubCashBook Then
        lblReportTitle.Caption = GetResourceString(390) & " " & _
            GetResourceString(85) & " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowSubCashBook
    ElseIf m_ReportType = repMemDayBook Then
        lblReportTitle.Caption = GetResourceString(62) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDayBook
    ElseIf m_ReportType = repMemLedger Then
        lblReportTitle.Caption = GetResourceString(93) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMemLedger
    ElseIf m_ReportType = repMemOpen Then
        lblReportTitle.Caption = GetResourceString(94) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMembersAdmitted
    ElseIf m_ReportType = repMemClose Then
        lblReportTitle.Caption = GetResourceString(95) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMembersCancelled
    ElseIf m_ReportType = repFeeCol Then
        lblReportTitle.Caption = GetResourceString(96) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMemberAndShareFee
    ElseIf m_ReportType = repMonthlyBalance Then
        Call ShowMonthlyBalances
    ElseIf m_ReportType = repMemLoanMembers Then
         lblReportTitle.Caption = GetResourceString(426)
         Call ShowLoanMemberBalances
    ElseIf m_ReportType = repMemNonLoanMembers Then
         lblReportTitle.Caption = GetResourceString(427)
         Call ShowNonLoanMembers
    Else 'If .optReports(7).value Then
        lblReportTitle.Caption = GetResourceString(53, 337, 295)
        Call ShowShareCertificate
    End If

Screen.MousePointer = vbNormal
    

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
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Wid = GetSetting(App.EXEName, "MemReport" & m_ReportType, _
                "ColWidth" & ColCount, 1 / grd.Cols) * grd.Width
    If Wid > grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
    grd.ColWidth(ColCount) = Wid
        
    Next ColCount

End Sub


Private Sub Form_Unload(cancel As Integer)
Set frmMMReport = Nothing

End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "MemReport" & m_ReportType, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount

End Sub


