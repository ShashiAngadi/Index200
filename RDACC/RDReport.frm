VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRDReport 
   Caption         =   "Reccuring Deposits Report.."
   ClientHeight    =   6075
   ClientLeft      =   1470
   ClientTop       =   1710
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   6555
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1860
      TabIndex        =   1
      Top             =   5430
      Width           =   4605
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   450
         Left            =   330
         TabIndex        =   5
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   450
         Left            =   1740
         TabIndex        =   3
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   450
         Left            =   3120
         TabIndex        =   2
         Top             =   90
         Width           =   1215
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
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2550
      TabIndex        =   4
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmRDReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_FromIndianDate As String
Dim m_ToIndianDate As String
Dim m_FromDate As Date
Dim m_ToDate As Date
Dim m_ReportType As wis_RDReports
Dim m_ReportOrder As wis_ReportOrder

Dim m_Gender As wis_Gender
Dim m_Caste As String
Dim m_Place As String
Dim m_AccGroup As Integer

Dim m_FromAmt As Currency
Dim m_ToAmt As Currency

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private WithEvents m_frmCancel As frmCancel
Attribute m_frmCancel.VB_VarHelpID = -1
Private m_TotalCount As Long

Public Event Initialising(Min As Long, Max As Long)
Public Event Processing(strMessages As String, Ratio As Single)
Public Event WindowClosed()



Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroup = NewValue
End Property


Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property


Public Property Let FromAmount(newAmount As Currency)
    m_FromAmt = newAmount
End Property

Public Property Let Gender(NewGender As Integer)
    m_Gender = NewGender
End Property

Public Property Let Place(NewPlace As String)
    m_Place = NewPlace
End Property

Public Property Let ReportOrder(NewRP As wis_ReportOrder)
    m_ReportOrder = NewRP
End Property

Public Property Let ReportType(newRT As wis_RDReports)
    m_ReportType = newRT
End Property

Public Property Let ToAmount(newAmount As Currency)
    m_ToAmt = newAmount
End Property

Public Property Let ToIndianDate(NewStrdate As String)
    If Not DateValidate(NewStrdate, "/", True) Then
        Err.Raise 5002, , "Invalid Date"
        Exit Property
    End If
    m_ToIndianDate = NewStrdate
    m_ToDate = GetSysFormatDate(NewStrdate)
    'm_ToIndianDate = GetAppFormatDate(m_ToDate)
End Property

Public Property Let FromIndianDate(NewStrdate As String)
    If Not DateValidate(NewStrdate, "/", True) Then Exit Property
    
    m_FromDate = GetSysFormatDate(NewStrdate)
    m_FromIndianDate = NewStrdate
    
End Property

Private Sub InitGrid(Optional Resize As Boolean)

If Resize Then Exit Sub

Dim ColWid As Single
Dim count As Integer
Dim colno As Byte
Dim rowno As Byte

    'ColWid = (grd.Width - 200) / grd.Cols + 1
    With grd
        .Clear
        .Rows = 50
        .Cols = 2
        .FixedCols = 0
    End With
    
On Error Resume Next
grd.Row = 0
    If m_ReportType = repRDBalance Then
        With grd
            .Cols = 4
            .FixedCols = 2
            .Row = 0: count = 0
            .Col = 0: .Text = GetResourceString(33)   '"Sl No "
            .Col = 1: .Text = GetResourceString(36) & " " & _
                GetResourceString(60)
            .Col = 2: .Text = GetResourceString(35)   '"Name"
            .Col = 3: .Text = GetResourceString(42)  '"Balance"
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            
        End With
    
    ElseIf m_ReportType = repRDCashBook Then
        With grd
            .Cols = 7: .Rows = 5
            .FixedCols = 1: .FixedRows = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)    ' "SL No"
            .Col = 1: .Text = GetResourceString(37)    ' "Date"
            .Col = 2: .Text = GetResourceString(36, 60) '"Acc NO"
            .Col = 3: .Text = GetResourceString(35)    '"Name"
            .Col = 4: .Text = GetResourceString(41)   '"Voucher"
            .Col = 5: .Text = GetResourceString(271)   '"Deposit"
            .Col = 6: .Text = GetResourceString(272)   '"Payment"
            .ColAlignment(0) = 1
            .ColAlignment(1) = 7
            .ColAlignment(2) = 7
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            
            
        End With
    ElseIf m_ReportType = repRDDayBook Then
        With grd
            .Cols = 11: .Rows = 5
            .FixedCols = 1: .FixedRows = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)    ' "SL No"
            .Col = 1: .Text = GetResourceString(37)    ' "Date"
            .Col = 2: .Text = GetResourceString(36, 60) '"Acc NO"
            .Col = 3: .Text = GetResourceString(35)    '"Name"
            .Col = 4: .Text = GetResourceString(271)   '"Deposit"
            .Col = 5: .Text = GetResourceString(271)   '"Deposit"
            .Col = 6: .Text = GetResourceString(289)   '"Payment"
            .Col = 7: .Text = GetResourceString(289)   '"Payment"
            .Col = 8: .Text = GetResourceString(47)   '"Interest"
            .Col = 9: .Text = GetResourceString(483)  '"Interest Received"
            .Row = 1
            .Col = 0: .Text = GetResourceString(33)    ' "SL No"
            .Col = 1: .Text = GetResourceString(37)    ' "Date"
            .Col = 2: .Text = GetResourceString(36) & " " & _
                        GetResourceString(60)   '"Acc NO"
            .Col = 3: .Text = GetResourceString(35)    '"Name"
            .Col = 4: .Text = GetResourceString(269)   '"cash"
            .Col = 5: .Text = GetResourceString(270)   '"Contra"
            .Col = 6: .Text = GetResourceString(269)   '"cash"
            .Col = 7: .Text = GetResourceString(270)   '"Contra"
            .Col = 8: .Text = GetResourceString(47)   '"Interest"
            .Col = 9: .Text = GetResourceString(483)  '"Interest Received"
            .MergeCells = flexMergeFree
            .MergeRow(0) = True: .MergeRow(1) = True
            .MergeCol(0) = True: .MergeCol(1) = True
            .MergeCol(2) = True: .MergeCol(3) = True
            .MergeCol(8) = True: .MergeCol(9) = True
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 7
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            .ColAlignment(7) = 1
            .ColAlignment(8) = 7
            .ColAlignment(9) = 7
            
        End With
    End If
    
    If m_ReportType = repRDLaib Then
        With grd
            .Cols = 8
            .FixedCols = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)   ' "Sl No"
            .Col = 1: .Text = GetResourceString(36, 60)   ' "Acc No"
            .Col = 2: .Text = GetResourceString(35)  '"Name"
            .Col = 3: .Text = GetResourceString(43) & " " & _
                GetResourceString(37) '"Deposit Date"
            .Col = 4: .Text = GetResourceString(48) & " " & _
                GetResourceString(37)   'MAturity date
            .Col = 5: .Text = GetResourceString(186) '"Interest Rate"
            .Col = 6: .Text = GetResourceString(42)  '"Balance"
            .Col = 7: .Text = GetResourceString(77)    '"Liability"
            
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            .ColAlignment(7) = 1
            .ColAlignment(8) = 7
            .ColAlignment(9) = 7
            
        End With
    End If
    
    If m_ReportType = repRDMat Then
        With grd
            .Cols = 5
            .Row = 0
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(36, 60) ' "Account No"
            .Col = 1: .Text = GetResourceString(35)    '"Name"
            .Col = 2: .Text = GetResourceString(48) & " " & _
                    GetResourceString(37)   'MAturity date
            .Col = 3: .Text = GetResourceString(186)  '"RateOfInterest"
            .Col = 4: .Text = GetResourceString(43, 42) 'Deposited amount
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            
        End With
    End If
    
    If m_ReportType = repRDAccClose Then
        With grd
            .Cols = 6
            .FixedCols = 2
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'SlnO
            .Col = 1: .Text = GetResourceString(36, 60)   '"AccNo"
            .Col = 2: .Text = GetResourceString(35)   '"Name"
            .Col = 3: .Text = GetResourceString(282)   '"Closed Date"
            .Col = 4: .Text = GetResourceString(46) & " " & _
                        GetResourceString(40)   'MAtured amount
            .Col = 5: .Text = GetResourceString(47)
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7
            
        End With
    End If
    
    If m_ReportType = repRDAccOpen Then
        With grd
            .Cols = 4
            .FixedCols = 2
            .FixedRows = 1
            .Row = 0: count = 0
            .Col = 0: .Text = GetResourceString(33) 'SlnO
            .Col = 1: .Text = GetResourceString(36, 60)   '"AccNo"
            .Col = 2: .Text = GetResourceString(35)   '"Name"
            .Col = 3: .Text = GetResourceString(281)   '"CreateDate"
            
            .ColAlignment(0) = 1
            .ColAlignment(1) = 7
            .ColAlignment(2) = 1
            .ColAlignment(3) = 1
            
        End With
    End If
    
    If m_ReportType = repRDLedger Then
        With grd
            .MergeCells = flexMergeNever
            .Cols = 6
            .FixedRows = 1
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)    '"Sl No":
            .Col = 1: .Text = GetResourceString(37)    '"Date":
            .Col = 2: .Text = GetResourceString(284) '"OPening Balance
            .Col = 3: .Text = GetResourceString(271) '"Deposit"
            .Col = 4: .Text = GetResourceString(272)  '"WithDraw"
            .Col = 5: .Text = GetResourceString(285)  '"Closing Balance"
            
                        
            .ColAlignment(0) = 1
            .ColAlignment(1) = 7
            .ColAlignment(2) = 7
            .ColAlignment(3) = 7
            .ColAlignment(4) = 7
            .ColAlignment(5) = 7

        End With
    End If
    
Dim RowCount As Integer
With grd
    For RowCount = 0 To .FixedRows - 1
        .Row = RowCount
        For count = 0 To .Cols - 1
            .Col = count
            .CellAlignment = 4
            .CellFontBold = True
        Next count
    Next
End With

LastLine:
    ColWid = 0

    For count = 0 To grd.Cols - 2
        ColWid = ColWid + grd.ColWidth(count)
    Next count
    grd.ColWidth(grd.Cols - 1) = grd.Width - ColWid - grd.Width * 0.04 'Me.ScaleWidth * 0.03
    
End Sub

Private Sub MaturedDeposits()

Dim count As Integer
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim SecondRst As ADODB.Recordset
Dim Days As Integer
Dim DepAmt As Currency, MatAmt As Currency
Dim Interest As Double
Dim DepTotal As Currency, MatTotal As Currency
Dim DepDate As String, MatDate As String

RaiseEvent Processing("Reading and Verifying the records ", 0)

SqlStr = "Select AccId,AccNum, CreateDate, " & _
    " MaturityDate,ClosedDate,RateOfInterest, Name" & _
    " From RDMaster A Inner JOin QryName B On B.CustomerId = A.CustomerId" & _
    " WHERE MaturityDate >= #" & m_FromDate & "#" & _
    " AND (ClosedDate > #" & m_ToDate & "# OR ClosedDate is NULL) " & _
    " and MaturityDate <= #" & m_ToDate & "# "
    
If m_FromAmt > 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " AND Amount <= " & m_ToAmt

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    SqlStr = SqlStr & " Order By CreateDate, AccNum"
Else
    SqlStr = SqlStr & " Order By CreateDate, IsciName"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
SqlStr = ""

    RaiseEvent Initialising(0, rst.RecordCount)
    RaiseEvent Processing("Reading the data ", 0)

'Init the grid
Call InitGrid
Dim rowno As Long
rowno = grd.FixedRows

While Not rst.EOF
    gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount From RDTrans " & _
                " Where AccId = " & FormatField(rst("Accid"))
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo nextRecord
    'Set next row
    
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: count = 0
        DepDate = FormatField(rst("CreateDate"))
        MatDate = FormatField(rst("MaturityDate"))
        Days = WisDateDiff(DepDate, MatDate)
        DepAmt = Val(FormatField(SecondRst("TotalAmount")))
        Interest = Val(FormatField(rst("RateOfInterest")))
        MatAmt = FormatCurrency(DepAmt)
        MatTotal = MatTotal + MatAmt
        DepTotal = DepTotal + DepAmt
        .TextMatrix(rowno, 0) = FormatField(rst("AccID"))
        .TextMatrix(rowno, 1) = FormatField(rst("Name"))
        .TextMatrix(rowno, 2) = MatDate
        .TextMatrix(rowno, 3) = Interest
        .TextMatrix(rowno, 4) = FormatCurrency(DepAmt)
    End With
    
nextRecord:
        DoEvents
        If gCancel = True Then rst.MoveLast
        RaiseEvent Processing("Writing the record to the grid ", rowno / rst.RecordCount)
        rst.MoveNext
    Wend
'End With
'Set last
    Set rst = Nothing
    Set SecondRst = Nothing

With grd
    rowno = rowno + 2
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    .Col = 0: .Text = GetResourceString(52) '"Totals"
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(DepTotal)
    .CellAlignment = 4: .CellFontBold = True
End With

End Sub

Private Sub ShowDepositsOpened()

'Declaring the variables
Dim rst As Recordset
Dim Total As Currency
Dim SqlStr As String

'Get date part of SQL
RaiseEvent Processing("Reading and verifying the records ", 0)
'Build the FINAL SQL
SqlStr = "Select AccId,AccNUm,CreateDate,MaturityDate, Name" & _
    " From QryName A Inner join RDMaster B" & _
    " ON B.CustomerID = A.CustomerID " & _
    " WHERE CreateDate <= #" & m_ToDate & "#" & _
    " And CreateDate >= #" & m_FromDate & "#"

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By CreateDate, val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By CreateDate, IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Initialize the Grid
 Call InitGrid
    
    RaiseEvent Initialising(0, rst.RecordCount)
    RaiseEvent Processing("Verifying the data to write into the grid  ", 0)
  
Dim AccId As Long
Dim SlNo As Integer
Dim rowno As Long
Dim MaxRows As Long
MaxRows = rst.RecordCount + 2
    
grd.Row = grd.FixedRows
rowno = grd.FixedRows
While Not rst.EOF
    'Set next row
    With grd
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("CreateDate"))
    End With
    
nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Wring data to the grid. ", rowno / MaxRows)
    
    rst.MoveNext
Wend
Set rst = Nothing
End Sub

Private Sub ShowDepositsClosed()

Dim SqlStr As String
Dim rst As Recordset

'Get date part of SQL
RaiseEvent Processing("Reading and Verifying the data ", 0)

gDbTrans.SqlStmt = "Select Max(TransId) as MaxTransID,AccID from RDTrans Group By AccId "
gDbTrans.CreateView ("RDMaxTrans")
gDbTrans.SqlStmt = "Select Amount,A.AccID from RDTrans A Inner Join " & _
    " RDMaxTrans B On B.AccID = A.AccID And B.MaxTransID=A.TransID "
gDbTrans.CreateView ("RDMatAmount")

gDbTrans.SqlStmt = "Select Max(TransId) as MaxTransID,AccID from RDIntTrans Group By AccId "
gDbTrans.CreateView ("RDIntMaxTrans")
gDbTrans.SqlStmt = "Select Amount ,A.AccID from RDIntTrans A Inner Join " & _
    " RDIntMaxTrans B On B.AccID = A.AccID And B.MaxTransID=A.TransID "
gDbTrans.CreateView ("RDIntAmount")

gDbTrans.SqlStmt = "Select Max(TransId) as MaxTransID,AccID from RDIntPayable Group By AccId "
gDbTrans.CreateView ("RDPayMaxTrans")
gDbTrans.SqlStmt = "Select A.AccID,Amount from RDIntPayable A Inner Join " & _
    " RDPayMaxTrans B On B.AccID = A.AccID And B.MaxTransID=A.TransID "
gDbTrans.CreateView ("RDPayAmount")


SqlStr = "Select B.AccId,AccNum, MaturityDate, ClosedDate,RateOfInterest, Name,IsciName," & _
    " iif(isnull(C.Amount),0,C.Amount) as MatAmount," & _
    " iif(isnull(D.Amount),0,D.Amount) as IntAmount, " & _
    " iif(isnull(E.Amount),0,E.Amount) as PayAmount " & _
    " From QryName A Inner Join " & _
    " (RDMaster B Left join (RDMatAmount C  " & _
        " Left join (RDIntAmount D Left join RDPayAmount E On E.AccID= D.AccID)" & _
    " On C.AccID= D.AccID) On C.AccID=B.AccID) ON A.CustomerId = B.CustomerID" & _
    " WHERE ClosedDate >= #" & m_FromDate & "#" & _
    " AND ClosedDate <= #" & m_ToDate & "#"
 


If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup


If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By ClosedDate, val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By ClosedDate, IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Initialize the Grid
Call InitGrid
    
    RaiseEvent Initialising(0, rst.RecordCount)
    RaiseEvent Processing("Reading the data to write into the grid. ", 0)

Dim AccId As Long
Dim TransID As Long
Dim SlNo As Integer
Dim Amount As Currency
Dim TotalAmount As Currency
Dim IntAmount As Currency
Dim TotalIntAmount As Currency
Dim rstTemp As ADODB.Recordset
Dim rowno As Long

rowno = grd.Row
While Not rst.EOF
    Amount = 0: IntAmount = 0
    'Get Returned Amount
    Amount = FormatField(rst("MatAmount"))
    IntAmount = FormatField(rst("IntAmount"))
    IntAmount = IntAmount + FormatField(rst("PayAmount"))
         
    'Set next row
    With grd
        rowno = rowno + 1
        If .Rows <= rowno + 1 Then .Rows = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("ClosedDate"))
        .TextMatrix(rowno, 4) = FormatCurrency(Amount)
        .TextMatrix(rowno, 5) = FormatCurrency(IntAmount)
    End With
    TotalAmount = TotalAmount + Amount
    TotalIntAmount = TotalIntAmount + IntAmount
    Amount = 0
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid . ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

Set rst = Nothing
Set rstTemp = Nothing
With grd
    'Set next row
    rowno = rowno + 2
    If .Rows <= rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    
    .Col = 2: .Text = GetResourceString(52, 43)   ' "Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(TotalAmount)
    .CellAlignment = 7: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(TotalIntAmount)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub

Private Sub ShowLiabilities()

'Declaring the variables
Dim SqlStr As String
Dim SecondRst As Recordset
Dim rst As Recordset
Dim count As Integer
                    
RaiseEvent Processing("Reading & Verifying the records. ", 0)

SqlStr = "Select A.AccID,AccNum, MaturityDate, CreateDate," & _
    " RateOfInterest, CLosedDate, Name " & _
    " From QryName B Inner join RDMaster A ON B.CustomerId = A.CustomerId" & _
    " Where A.AccId NOT In (Select AccId From RDMaster" & _
        " Where ClosedDate < #" & m_ToDate & "#" & ")"

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By Val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

'Init the grid
Call InitGrid

'''RaiseEvent Initialising(0, Rst.RecordCount)
RaiseEvent Processing("Reading the records to write into the grid ", 0)

Dim Amount As Currency
Dim SlNo As Integer
Dim Liability As Currency
Dim GrandTotal As Currency

Dim CustName As String
Dim rowno As Long

grd.Row = grd.FixedRows - 1
rowno = grd.FixedRows - 1
While Not rst.EOF
    gDbTrans.SqlStmt = "Select sum(Amount) as TotalAmount from RDTrans " & _
        " where TransDate <= #" & m_ToDate & "# AND AccId = " & rst("AccId")
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo nextRecord
    
    If Val(FormatField(SecondRst("TotalAmount"))) = 0 Then GoTo nextRecord
    
    Amount = FormatField(SecondRst("TotalAmount"))
    Liability = Amount + ComputeRDInterest(Amount, FormatField(rst("RateOfInterest")))
    If m_FromAmt > 0 And Liability < m_FromAmt Then GoTo nextRecord
    If m_ToAmt > 0 And Liability > m_ToAmt Then GoTo nextRecord

    SlNo = SlNo + 1
    With grd
       'Set next row
        rowno = rowno + 1
        If .Rows <= rowno + 1 Then .Rows = rowno + 1
        count = 0
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("CreateDate"))
        .TextMatrix(rowno, 4) = FormatField(rst("MaturityDate"))
        .TextMatrix(rowno, 5) = FormatField(rst("RateOfInterest"))
        .TextMatrix(rowno, 6) = FormatCurrency(Amount)
        .TextMatrix(rowno, 7) = FormatCurrency(Liability)
    End With
    GrandTotal = GrandTotal + Liability
nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid . ", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend
'Fill In total Liability

Set rst = Nothing
Set SecondRst = Nothing

With grd
    rowno = rowno + 2
    If .Rows <= rowno + 1 Then .Rows = rowno + 1
    .Row = rowno

    .Col = 2: .Text = GetResourceString(77) '"TOTAL LIABILITIES"
    .CellFontBold = True
    .Col = 7: .Text = FormatCurrency(GrandTotal)
    .CellFontBold = True: .CellAlignment = 7
    
End With

End Sub

Private Sub ShowDayBook()

'Declring the variables
Dim rst As Recordset
Dim transType As wisTransactionTypes

RaiseEvent Processing("Reading & verifying the records. ", 0)

Dim SqlStr As String
SqlStr = "Select 'PRINCIPAL',AccNum , A.AccId,TransDate, Amount," & _
    " TransType, Name as CustName, IsciName," & _
    " A.TransID FROM QryName c Inner join " & _
    " (RDMaster B Inner join RDTrans A ON A.AccId = B.AccId) " & _
    " ON B.CustomerId = C.CustomerID where " & _
    " TransDate >= #" & m_FromDate & "# AND TransDate <= #" & m_ToDate & "#"

If m_FromAmt > 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " AND Amount <= " & m_ToAmt

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

SqlStr = SqlStr & " UNION " & "Select 'INTEREST',AccNum, A.AccId,TransDate, " & _
    " Amount,TransType,Name as CustName, IsciName," & _
    " A.TransID FROM QryName c inner join " & _
    " (RDMaster B inner join RDIntTrans A on A.AccId = B.AccId)" & _
    " ON B.CustomerId = C.CustomerID  where " & _
    " TransDate >= #" & m_FromDate & "# AND TransDate <= #" & m_ToDate & "#"

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By TransDate,AccNum"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By TransDate,IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

SqlStr = ""

Call InitGrid

RaiseEvent Initialising(0, rst.RecordCount)
RaiseEvent Processing("Verifying the data to write into the grid. ", 0)

Dim SubTotal(4 To 11) As Currency
Dim GrandTotal(4 To 11) As Currency
Dim I As Integer
Dim SlNo As Integer
Dim TransID As Long
Dim AccNum As String
Dim TransDate As Date
Dim Amount As Currency
Dim PRINTTotal As Boolean
Dim rowno As Long

I = 1
AccNum = FormatField(rst("AccNum"))
TransDate = rst("TransDate")
grd.Row = grd.FixedRows
rowno = grd.FixedRows
While Not rst.EOF
    With grd
        'See if you have to calculate sub totals
        If DateDiff("D", TransDate, rst("TransDate")) Then
            PRINTTotal = True
            SlNo = 0
            'Set next row
            rowno = rowno + 1
            If .Rows < rowno + 2 Then .Rows = rowno + 2
            .Row = rowno
            .Col = 1: .CellFontBold = True: .Text = ". " & GetIndianDate(TransDate)
            .Col = 3: .CellFontBold = True: .Text = GetResourceString(304)
            For I = 4 To 9
                .Col = I: .CellFontBold = True
                If SubTotal(I) Then .Text = FormatCurrency(SubTotal(I))
                GrandTotal(I) = GrandTotal(I) + SubTotal(I)
                SubTotal(I) = 0
            Next
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
        End If
        
        If AccNum <> FormatField(rst("AccNum")) Then TransID = 0
        AccNum = FormatField(rst("AccNum"))
        'Set next row
        If TransID <> rst("TransID") Then
            If .Rows < rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = FormatField(rst("TransDate")): .CellAlignment = 7
            .TextMatrix(rowno, 2) = AccNum
            .TextMatrix(rowno, 3) = FormatField(rst("CustName"))
        End If
        TransID = rst("TransID")
        transType = rst("TransType")
        TransDate = rst("TransDate")
        If rst(0) = "INTEREST" Then
            If transType = wWithdraw Or transType = wContraWithdraw Then I = 8
            If transType = wDeposit Or transType = wContraDeposit Then I = 9
        Else
            If transType = wDeposit Then I = 4
            If transType = wContraDeposit Then I = 5
            If transType = wWithdraw Then I = 6
            If transType = wContraWithdraw Then I = 7
        End If
        Amount = FormatField(rst("Amount"))
        .TextMatrix(rowno, I) = FormatCurrency(Amount)
        SubTotal(I) = SubTotal(I) + Amount
        
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Writing the data to the grid . ", rowno / rst.RecordCount)
        rst.MoveNext
    End With
Wend

Set rst = Nothing

  With grd
        'Set next row
        
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 1: .CellFontBold = True: .Text = ". " & GetIndianDate(TransDate)
        .Col = 3: .CellFontBold = True: .Text = GetResourceString(304)
        For I = 4 To 9
            .Col = I: .CellFontBold = True
            If SubTotal(I) Then .Text = FormatCurrency(SubTotal(I))
            GrandTotal(I) = GrandTotal(I) + SubTotal(I)
            SubTotal(I) = 0
        Next
        If PRINTTotal Then
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1:
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 1: .CellFontBold = True: .Text = ". " & GetIndianDate(TransDate)
            .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) ' "Grand Totals"
            For I = 4 To 9
                .Col = I: .CellFontBold = True
                If GrandTotal(I) Then .Text = FormatCurrency(GrandTotal(I))
            Next
        End If
  End With
    
End Sub

Private Sub ShowSubCashBook()

'Declring the variables
Dim rst As Recordset
Dim transType As wisTransactionTypes

RaiseEvent Processing("Reading & verifying the records. ", 0)

Dim SqlStr As String
SqlStr = "Select AccNum , A.AccId,TransDate, Amount, " & _
    " TransType,VoucherNo, Name as CustName, " & _
    " A.TransID FROM QryName c inner join " & _
    " (RDMaster B inner join RDTrans A ON A.AccId = B.AccId )" & _
    " ON B.CustomerId = C.CustomerID  " & _
    " Where TransDate >= #" & m_FromDate & "# AND TransDate <= #" & m_ToDate & "#"

If m_FromAmt > 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " AND Amount <= " & m_ToAmt

If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By TransDate,AccNum"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By TransDate"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

SqlStr = ""

Call InitGrid

RaiseEvent Initialising(0, rst.RecordCount)
RaiseEvent Processing("Verifying the data to write into the grid. ", 0)

Dim SubTotal(5 To 6) As Currency
Dim GrandTotal(5 To 6) As Currency
Dim I As Integer
Dim SlNo As Integer
Dim TransID As Long
Dim AccNum As String
Dim TransDate As Date
Dim Amount As Currency
Dim PRINTTotal As Boolean
Dim rowno As Long
Dim MaxRows As Long
MaxRows = rst.RecordCount + 2

I = 1
AccNum = FormatField(rst("AccNum"))
TransDate = rst("TransDate")
grd.Row = grd.FixedRows
rowno = grd.Row
While Not rst.EOF
    With grd
        'See if you have to calculate sub totals
        If DateDiff("D", TransDate, rst("TransDate")) Then
            PRINTTotal = True
            SlNo = 0
            'Set next row
            rowno = rowno + 1
            If .Rows < rowno + 1 Then .Rows = rowno + 1
            .Row = rowno
            .Col = 1: .CellFontBold = True: .Text = ". " & GetIndianDate(TransDate)
            .Col = 3: .CellFontBold = True: .Text = GetResourceString(304)
            For I = 5 To 6
                .Col = I: .CellFontBold = True
                If SubTotal(I) Then .Text = FormatCurrency(SubTotal(I))
                GrandTotal(I) = GrandTotal(I) + SubTotal(I)
                SubTotal(I) = 0
            Next
            rowno = rowno + 1
        End If
        
        If AccNum <> FormatField(rst("AccNum")) Then TransID = 0
        AccNum = FormatField(rst("AccNum"))
        'Set next row
        If TransID <> rst("TransID") Then
            rowno = rowno + 1
            If .Rows <= rowno + 1 Then .Rows = rowno + 1
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = FormatField(rst("TransDate")): .CellAlignment = 7
            .TextMatrix(rowno, 2) = AccNum: .CellAlignment = 7
            .TextMatrix(rowno, 3) = FormatField(rst("CustName"))
            .TextMatrix(rowno, 4) = FormatField(rst("VoucherNo"))
        End If
        TransID = rst("TransID")
        transType = rst("TransType")
        TransDate = rst("TransDate")
        I = 6
        If transType = wDeposit Or transType = wContraDeposit Then I = 5
        Amount = FormatField(rst("Amount"))
        .TextMatrix(rowno, I) = FormatCurrency(Amount)
        SubTotal(I) = SubTotal(I) + Amount
        
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Writing the data to the grid . ", rowno / MaxRows)
        rst.MoveNext
    End With
Wend

Set rst = Nothing

  With grd
        'Set next row
        rowno = rowno + 1
        If .Rows <= rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 1: .CellFontBold = True: .Text = ". " & GetIndianDate(TransDate)
        .Col = 3: .CellFontBold = True: .Text = GetResourceString(304)
        For I = 5 To 6
            .Col = I: .CellFontBold = True
            If SubTotal(I) Then .Text = FormatCurrency(SubTotal(I))
            GrandTotal(I) = GrandTotal(I) + SubTotal(I)
            SubTotal(I) = 0
        Next
        If PRINTTotal Then
            rowno = rowno + 2
            If .Rows <= rowno + 1 Then .Rows = rowno + 1
            .Row = rowno
            .Col = 1: .CellFontBold = True: .Text = "." & GetIndianDate(TransDate)
            .Col = 3: .CellFontBold = True: .Text = GetResourceString(286) ' "Grand Totals"
            For I = 5 To 6
                .Col = I: .CellFontBold = True
                If GrandTotal(I) Then .Text = FormatCurrency(GrandTotal(I))
            Next
        End If
  End With
  
  
End Sub



Private Sub ShowDepositBalances()

'Declaring the variables
Dim rst As Recordset
Dim SqlStmt As String

RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStmt = "Select Max(TransId) as MaxTransID,AccID From RDTrans Where " & _
            " TransDate <= #" & m_ToDate & "# Group BY AccId"
gDbTrans.SqlStmt = SqlStmt
'Create view to get Max Transaction ID of RD Trans
gDbTrans.CreateView ("RDMaxTrans")
'Build Next Querry Using Above Querry

SqlStmt = "Select Balance,A.AccID, B.AccNum, Name " & _
    " From QryName C inner join (RDMaster B " & _
    " inner join (RDtrans A inner join RDMaxTrans D" & _
    " ON D.AccID = A.AccID AND D.MaxTransID = A.TransID)" & _
    " ON B.AccId = A.Accid) ON B.CustomerId = C.CustomerId "

Dim sqlClause As String
sqlClause = ""
sqlClause = sqlClause & " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then sqlClause = sqlClause & " And Balance < " & m_ToAmt
If m_Place <> "" Then sqlClause = sqlClause & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then sqlClause = sqlClause & " AND Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " AND AccGroupId = " & m_AccGroup
If Len(sqlClause) Then
    sqlClause = Trim$(sqlClause)
    sqlClause = " WHERE " & Mid(sqlClause, 4)
    SqlStmt = SqlStmt & sqlClause
End If
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & " Order By val(B.AccNum)"
Else
    gDbTrans.SqlStmt = SqlStmt & " Order By IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Call InitGrid

RaiseEvent Initialising(0, rst.RecordCount)
RaiseEvent Processing("Reading and Verifying the data ", 0)

Dim TotalAmount As Currency
Dim SlNo As Integer
Dim rowno As Long

grd.Row = grd.FixedRows - 1
rowno = grd.FixedRows - 1

While Not rst.EOF
    With grd
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        SlNo = SlNo + 1
         .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        .TextMatrix(rowno, 3) = FormatField(rst("Balance"))
        TotalAmount = TotalAmount + Val(.TextMatrix(rowno, 3))
    End With

nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Wring data to the grid. ", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend
    
    Set rst = Nothing

    With grd
        rowno = rowno + 2
        If .Rows < rowno + 3 Then .Rows = rowno + 3
        .Row = rowno
        .Col = 2: .CellFontBold = True
        .Text = GetResourceString(52, 42)   '"Totals Balance "
        .Col = .Cols - 1: .CellFontBold = True
        .Text = FormatCurrency(TotalAmount)
    End With
    
End Sub

Private Sub cmdOk_Click()

Unload Me
RaiseEvent WindowClosed
End Sub

Private Sub cmdPrint_Click()
 ' Call the print class services...
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    Set m_frmCancel = New frmCancel
    Load m_frmCancel
        
    With m_frmCancel
        .PicStatus.Visible = True
        .Show
    End With
    
    .GridObject = grd
    .ReportTitle = Me.lblReportTitle.Caption
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
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

'Center the form
 Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'set kannada fonts
Call SetKannadaCaption
'Init the grid
With grd
    .Clear
    .Rows = 35
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    
    .Text = GetResourceString(278) '"No Records Available"
    .CellAlignment = 4: .CellFontBold = True
End With

'Show report
    If m_ReportType = repRDBalance Then
        lblReportTitle.Caption = GetResourceString(70)
        Call ShowDepositBalances
    End If
    If m_ReportType = repRDDayBook Then
        lblReportTitle.Caption = GetResourceString(424, 85) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDayBook
    End If
    If m_ReportType = repRDCashBook Then
        lblReportTitle.Caption = GetResourceString(424, 63) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowSubCashBook
    End If
    If m_ReportType = repRDLaib Then
        Me.lblReportTitle.Caption = GetResourceString(77)
        Call ShowLiabilities
    End If
    If m_ReportType = repRDMat Then
        lblReportTitle.Caption = GetResourceString(72) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call MaturedDeposits
    End If
    If m_ReportType = repRDAccClose Then
        lblReportTitle.Caption = GetResourceString(78) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        
        Call ShowDepositsClosed
    End If
    If m_ReportType = repRDAccOpen Then
        lblReportTitle.Caption = GetResourceString(64) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositsOpened
    End If
    'Get the monthly balancess
    If m_ReportType = repRDMonbal Then Call ShowMonthlyBalances
    
    If m_ReportType = repRDLedger Then
        If m_FromIndianDate = "" Or m_ToIndianDate = "" Then
            lblReportTitle.Caption = GetResourceString(43) & " " & _
                GetResourceString(93) '"Deposit GeneralLegder
        Else
            lblReportTitle.Caption = GetResourceString(43) & " " & _
                GetResourceString(93) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        End If
        Call ShowDepositGeneralLedger
    End If
        
   lblReportTitle.FONTSIZE = 14
   
End Sub

Private Sub Form_Resize()

Screen.MousePointer = vbDefault
On Error Resume Next
lblReportTitle.Top = 0
lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
With grd
.Left = 0
.Top = lblReportTitle.Top + lblReportTitle.Height
.Width = Me.Width - 150
End With
fra.Top = Me.ScaleHeight - fra.Height
fra.Left = Me.Width - fra.Width
grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height
cmdOk.Left = fra.Width - cmdOk.Width - (cmdOk.Width / 4)
cmdPrint.Left = cmdOk.Left - cmdPrint.Width - (cmdPrint.Width / 4)
cmdWeb.Top = cmdPrint.Top
cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)

Dim ColCount As Integer
Dim Wid As Single
With grd
For ColCount = 0 To grd.Cols - 1
    Wid = GetSetting(App.EXEName, "RDReport" & m_ReportType, "ColWidth" & ColCount, 1 / .Cols) * .Width
    If Wid >= .Width * 0.9 Then Wid = .Width / .Cols
    If Wid <= 0 Then Wid = .Width / .Cols
    
    grd.ColWidth(ColCount) = Wid
Next ColCount
End With
End Sub
Private Sub ShowDepositGeneralLedger()

'Declaring the variables
Dim count As Integer
Dim SqlStr As String
Dim rst As Recordset
Dim TransDate As Date

RaiseEvent Processing("Verifying the data ", 0)

SqlStr = "Select Sum(Amount) as TotalAmount, " & _
    " TransDate,TransType From RDTrans where " & _
    " TransDate >= #" & m_FromDate & "# AND TransDate <= #" & m_ToDate & "#" & _
    " Group By TransDate,TransType"

gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

grd.Clear

Call InitGrid

Dim DepAmount As Double
Dim WithdrawAmount As Double
Dim TotalDepAmount As Double
Dim TotalWithdrawAmount As Double

Dim I As Integer
Dim Balance As Currency
Dim transType As wisTransactionTypes
    
RaiseEvent Initialising(0, rst.RecordCount)
RaiseEvent Processing("Reading the data ", 0)
    
TransDate = DateAdd("d", -1, m_FromDate)
Balance = GetRDBalance(TransDate)
TransDate = rst("TransDate")

Dim PRINTTotal As Boolean
Dim rowno As Long

With grd
    .Row = .FixedRows
    .Col = 1: .Text = GetResourceString(284)
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(Balance)
    .CellFontBold = True
    .Col = 3: .CellFontBold = False
    rowno = .Row
End With

While Not rst.EOF
    If DateDiff("d", TransDate, rst("TransDate")) <> 0 Then
        PRINTTotal = True
        With grd
            rowno = rowno + 1
            If .Rows <= rowno + 1 Then .Rows = rowno + 1
            .TextMatrix(rowno, 0) = rowno - 1
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = FormatCurrency(Balance)
            
            If DepAmount Then
                .TextMatrix(rowno, 3) = FormatCurrency(DepAmount)
                TotalDepAmount = TotalDepAmount + DepAmount
            End If
            If WithdrawAmount Then
                .TextMatrix(rowno, 4) = FormatCurrency(DepAmount)
                TotalWithdrawAmount = TotalWithdrawAmount + WithdrawAmount
            End If
            Balance = Balance + DepAmount - WithdrawAmount
            DepAmount = 0: WithdrawAmount = 0
            .TextMatrix(rowno, 5) = FormatCurrency(Balance)
            TransDate = FormatField(rst("TransDate"))

        End With
    End If
    transType = rst("TransType")
    If transType = wDeposit Or transType = wContraDeposit Then
        DepAmount = DepAmount + FormatField(rst("TotalAmount"))
    Else
        WithdrawAmount = WithdrawAmount + FormatField(rst("TotalAmount"))
    End If
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing data to the grid. ", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend
    
Set rst = Nothing

With grd
    rowno = rowno + 1
    If .Rows <= rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    .Col = 0: .Text = (.Row - 1)
    .Col = 1: .Text = GetIndianDate(TransDate)
    .Col = 2: .Text = FormatCurrency(Balance)
    If DepAmount Then
        .Col = 3
        .Text = FormatCurrency(DepAmount)
        TotalDepAmount = TotalDepAmount + DepAmount
    End If
    If WithdrawAmount Then
        .Col = 4
        .Text = FormatCurrency(WithdrawAmount)
        TotalWithdrawAmount = TotalWithdrawAmount + WithdrawAmount
    End If
    Balance = Balance + DepAmount - WithdrawAmount
    DepAmount = 0: WithdrawAmount = 0
    .Col = 5: .Text = FormatCurrency(Balance)
    rowno = .Row
End With

If PRINTTotal Then
    With grd
        rowno = rowno + 2
        If .Rows <= rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        
        .Col = 3: .CellFontBold = True
        .Text = FormatCurrency(TotalDepAmount)
        .Col = 4: .CellFontBold = True
        .Text = FormatCurrency(TotalWithdrawAmount)
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 4: grd.Text = GetResourceString(285)
        .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(Balance)
        .CellFontBold = True
    End With
Else
    'Remove the Firs row
    grd.RemoveItem 1
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
fromDate = FinUSFromDate
If Len(m_FromIndianDate) Then fromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463) & " " & _
            GetResourceString(67) & " " & _
            GetResourceString(42) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)

SqlStmt = "SELECT A.AccNum,A.AccID, A.CustomerID,Name as CustName " & _
        " From QryName B Inner Join RDMaster A" & _
            " ON A.CustomerID = B.CustomerID" & _
        " WHERE A.CreateDate <= #" & toDate & "#" & _
        " AND (A.ClosedDate Is NULL OR A.Closeddate >= #" & fromDate & "#)" & _
        " Order By val(A.ACCNUM)"

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub

count = DateDiff("M", fromDate, toDate) + 2
totalCount = (count + 1) * rstMain.RecordCount
RaiseEvent Initialising(0, totalCount)

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
    .Row = 0
End With

count = 0
Dim rowno As Long

While Not rstMain.EOF
    With grd
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .Col = 0: .TextMatrix(rowno, 0) = rowno
        .TextMatrix(rowno, 1) = FormatField(rstMain("AccId"))
        .TextMatrix(rowno, 2) = FormatField(rstMain("CustName"))
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
    rstMain.MoveNext
Wend
    
    With grd
        rowno = rowno + 2
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 2: .Text = GetResourceString(286) 'Grand Total
        .CellAlignment = 4: .CellFontBold = True
    End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    SqlStmt = "SELECT [AccId], Max([TransID]) AS MaxTransID" & _
            " FROM RDTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY [AccID];"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("RDMonBal")
    SqlStmt = "SELECT A.AccId,Balance From RDTrans A,RDMonBal B " & _
        " Where B.AccId = A.AccID ANd  TransID =MaxTransID"
    gDbTrans.SqlStmt = SqlStmt
    
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0: rowno = 0
        .Col = .Cols - 1: .Text = GetMonthString(Month(fromDate)) & _
                " " & GetResourceString(42)
        .CellAlignment = 4: .CellFontBold = True
    End With
    
    rstMain.MoveFirst
    grd.Row = 0: rowno = 0
    TotalBalance = 0
    grd.Col = grd.Cols - 1
    
    While Not rstMain.EOF
        rowno = rowno + 1
        Balance = 0
        rstBalance.MoveFirst
        rstBalance.Find " ACCID = " & rstMain("AccID")
        If Not rstBalance.EOF Then Balance = FormatField(rstBalance("Balance"))
        
        grd.TextMatrix(rowno, grd.Col) = FormatCurrency(Balance)
        TotalBalance = TotalBalance + Balance
        
        ProcCount = ProcCount + 1
        DoEvents
        If gCancel Then rstMain.MoveLast
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
        rstMain.MoveNext
    Wend
    rowno = rowno + 2
    grd.Row = rowno
    grd.Text = FormatCurrency(TotalBalance)
    grd.CellFontBold = True

NextMonth:

    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

Exit Sub
ErrLine:
    MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRDReport = Nothing
End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdOk.Caption = GetResourceString(11)
cmdPrint.Caption = GetResourceString(23)

End Sub

Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "RDReport" & m_ReportType, _
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


