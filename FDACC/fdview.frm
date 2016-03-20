VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFDReport 
   Caption         =   "Fixed Deposit Reports.."
   ClientHeight    =   5805
   ClientLeft      =   2655
   ClientTop       =   2175
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
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
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   900
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
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8334
      _Version        =   393216
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frmFDReport"
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
Dim m_ReportOrder As wis_ReportOrder
Dim m_ReportType As wis_FDReports
Dim m_DepositType As Long
Dim m_Gender As wis_Gender
Dim m_AccGroup As Integer

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)


Public Property Let AccountGroup(NewValue As Integer)
m_AccGroup = NewValue
End Property


Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property

Public Property Let FromAmount(curFrom As Currency)
    m_FromAmt = curFrom
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


Public Property Let Gender(NewValue As wis_Gender)
    m_Gender = NewValue
End Property


Public Property Let Place(NewPlace As String)
    m_Place = NewPlace
End Property
Public Property Let ReportOrder(RepOrder As wis_ReportOrder)
    m_ReportOrder = RepOrder
End Property

Public Property Let ReportType(RepType As wis_FDReports)
    m_ReportType = RepType
End Property

Private Sub SelectAndShow()

Select Case m_ReportType
    Case repFDBalance
        lblReportTitle.Caption = GetResourceString(70)
        Call ShowDepositBalances
    Case repMFDBalance
        lblReportTitle.Caption = GetResourceString(220)
        Call ShowMFDBalances
    
    Case repFDDayBook, repMFDDayBook
        Call ShowDayBook
    
    Case repFDCashBook, repMFDCashBook
        Call ShowCashBook
    
    Case repFDLaib
        lblReportTitle.Caption = GetResourceString(77)
        Call ShowLiabilities
    
    Case repFDMat
        If m_FromIndianDate = "" Or m_ToIndianDate = "" Then
            Me.lblReportTitle.Caption = GetResourceString(72)
        Else
            Me.lblReportTitle.Caption = GetResourceString(72) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        End If
        Call MaturedDeposits
    
    Case m_ReportType = repFDAccOpen
        lblReportTitle.Caption = GetResourceString(78) & " " & _
                        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositsOpened
    
    Case repFDAccClose
        lblReportTitle.Caption = GetResourceString(65) & " " & _
                            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositsClosed
    
    Case repFDLedger, repMFDLedger
        Call ShowGeneralLedger
    
    Case repFDTrans
        lblReportTitle.Caption = GetResourceString(43) & " " & _
            GetResourceString(62) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositTransMade
    Case repMFDTrans
        lblReportTitle.Caption = GetResourceString(220) & " " & _
            GetResourceString(62) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMatDepositTransMade
    
    Case repFDMonbal
        lblReportTitle.Caption = GetResourceString(463) & " " & _
            GetResourceString(42) & " " & _
            GetFromDateString(GetMonthString(Month(m_FromDate)), GetMonthString(Month(m_ToDate)))
        Call ShowMonthlyBalance
    Case repFDJoint
        Call ShowJointAccounts
    Case repFDAccOpen
        Call ShowDepositsOpened
End Select

End Sub

Private Sub ShowJointAccounts()
Dim SqlStr As String
Dim rst As ADODB.Recordset

lblReportTitle.Caption = GetResourceString(265, 43)  '"Joint Deposit
SqlStr = "Select A.AccNum,A.AccID,CreateDate,Balance,A.CustomerID as MainCustID," & _
    " B.CustomerID as CustID FROM  FDJoint B, FDMaster A,FDTrans C WHERE " & _
    " A.AccID = B.AccId and C.AccID = A.AccID" & _
    " AND TransID = (SELECT MAX(TransID) From FDTrans D" & _
        " WHERE D.AccID = A.AccID And TransDate <= #" & m_FromDate & "#)" & _
    " AND A.DepositType = " & m_DepositType & _
    " AND Balance > 0 ORDER BY val(A.AccNum),A.AccID"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
Dim CustClass As New clsCustReg
Dim SlNo As Integer
Dim CustNo As Byte
Dim AccNum As String
Call InitGrid
Dim rowno As Long, colno As Byte

rowno = grd.FixedRows - 1
While Not rst.EOF
    With grd
        If AccNum <> rst("AccNum") Then
            AccNum = rst("accNum")
            If .Rows < rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            .Row = rowno: .Col = 0
            SlNo = SlNo + 1: CustNo = 1
            .TextMatrix(rowno, 0) = SlNo & "  " & CustNo
            .TextMatrix(rowno, 1) = rst("AccNum")
            .TextMatrix(rowno, 2) = CustClass.CustomerName(rst("MainCustID"))
        End If
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        CustNo = CustNo + 1
        .TextMatrix(rowno, 0) = CustNo
        .TextMatrix(rowno, 1) = rst("AccNum")
        .TextMatrix(rowno, 2) = CustClass.CustomerName(rst("CustID"))
        .TextMatrix(rowno, 3) = FormatCurrency(rst("Balance"))
    End With
    rst.MoveNext
Wend

End Sub

Private Sub InitGrid()
Dim count As Integer
    
    'ColWid = (grd.Width - 200) / grd.Cols + 1
On Error Resume Next
grd.Row = 0
grd.MergeCells = flexMergeRestrictRows

If m_ReportType = repFDBalance Or m_ReportType = repMFDBalance Then
    'Or m_ReportType = repFDMat
    With grd
        .Cols = 5
        .FixedCols = 1
        .Row = 0: count = 0
        .Col = 0: .Text = GetResourceString(33) ' "Sl NO"
        .Col = 1: .Text = GetResourceString(36, 60) ' "Account No"
        .Col = 2: .Text = GetResourceString(43, 37) 'DepostDate
        .Col = 3: .Text = GetResourceString(35) '"Name"
        .Col = 4: .Text = GetResourceString(42) '"Balance"
        .ColAlignment(3) = 1
    End With
End If

If m_ReportType = repFDCashBook Or m_ReportType = repMFDCashBook Then
    With grd
        .Cols = 7
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)    ' "Sl NOp"
        .Col = 1: .Text = GetResourceString(37)    ' "Date"
        .Col = 2: .Text = GetResourceString(36, 60)    '"Acc NO"
        .Col = 3: .Text = GetResourceString(35)    '"Name"
        .Col = 4: .Text = GetResourceString(41)   '"Vouche No"
        .Col = 5: .Text = GetResourceString(271)   '"Deposit"
        .Col = 6: .Text = GetResourceString(272)   '"Payment"
        
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        
    End With

End If
If m_ReportType = repFDDayBook Or m_ReportType = repMFDDayBook Then
    With grd
        .Cols = 10
        .FixedCols = 1
        .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)    ' "sl No"
        .Col = 1: .Text = GetResourceString(37)    ' "Date"
        .Col = 2: .Text = GetResourceString(36, 60)    '"Acc NO"
        .Col = 3: .Text = GetResourceString(35)    '"Name"
        .Col = 4: .Text = GetResourceString(271)   '"Deposit"
        .Col = 5: .Text = GetResourceString(271)   '"Deposit"
        .Col = 6: .Text = GetResourceString(272)   '"Payment"
        .Col = 7: .Text = GetResourceString(272)   '"Payment"
        .Col = 8: .Text = GetResourceString(272)   '"Interest"
        .Col = 9: .Text = GetResourceString(272)   '"Payable"
        .Row = 1
        .Col = 0: .Text = GetResourceString(33)    ' "sl No"
        .Col = 1: .Text = GetResourceString(37)    ' "Date"
        .Col = 2: .Text = GetResourceString(36, 60)    '"Acc NO"
        .Col = 3: .Text = GetResourceString(35)    '"Name"
        .Col = 4: .Text = GetResourceString(269)   '"Cash"
        .Col = 5: .Text = GetResourceString(270)   '"Contra"
        .Col = 6: .Text = GetResourceString(269)   '"Cash"
        .Col = 7: .Text = GetResourceString(270)   '"Contra"
        .Col = 8: .Text = GetResourceString(47)   '"Interest"
        .Col = 9: .Text = GetResourceString(450)   '"Payable"
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = False
        .MergeCol(4) = False
        .MergeCol(5) = False
        .MergeCol(6) = False
        .MergeCol(7) = True
        .MergeCol(8) = True
        
    End With
End If
If m_ReportType = repFDLaib Then
    With grd
        .Cols = 7
        .FixedCols = 1
        .Row = 0
        count = 0
        .Col = 0: .Text = GetResourceString(36, 60)   ' "Acc No"
        .Col = 1: .Text = GetResourceString(35)  '"Name"
        .Col = 2: .Text = GetResourceString(43) & _
                GetResourceString(37) '"Deposit Date"
        .Col = 3: .Text = GetResourceString(48, 37)  'MAturity date
        .Col = 4: .Text = GetResourceString(186) '"Interest Rate"
        .Col = 5: .Text = GetResourceString(43, 42) '"Deposited Amount"
        .Col = 6: .Text = GetResourceString(77)  '"Liability"
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(2) = 1
    End With
End If
If m_ReportType = repFDMat Then
    With grd
        .Cols = 6
        .Row = 0
        .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)   '"sL No"
        .Col = 1: .Text = GetResourceString(36, 60)   '"Account No"
        .Col = 2: .Text = GetResourceString(35)    '"Name"
        .Col = 3: .Text = GetResourceString(48, 37)  'MAturity date
        .Col = 4: .Text = GetResourceString(186)  '"RateOfInterest"
        .Col = 5: .Text = GetResourceString(43, 42) 'Deposited amount
        .ColAlignment(4) = 1
    End With
End If

If m_ReportType = repFDAccClose Then
    With grd
        .Cols = 5
        .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(36) & " " & _
                    GetResourceString(60)  '"AccNo"
        .Col = 1: .Text = GetResourceString(35)   '"Name"
        .Col = 2: .Text = GetResourceString(282)   '"Closed Date"
        .Col = 3: .Text = GetResourceString(43) & " " & _
                GetResourceString(40)   'Deposit amount
        .Col = 4: .Text = GetResourceString(46) & " " & _
                GetResourceString(40)   'MAtured amount
    End With
End If

If m_ReportType = repFDAccOpen Then
    grd.Cols = 6
    grd.FixedCols = 1
    grd.WordWrap = True
    grd.Row = 0: count = 0
    grd.Col = 0: grd.Text = GetResourceString(33)   '"Sl NO"
    grd.Col = 1: grd.Text = GetResourceString(36, 60)   '"AccNo"
    grd.Col = 2: grd.Text = GetResourceString(35)   '"Name"
    grd.Col = 3: grd.Text = GetResourceString(43, 60)  '"Deposit nO"
    grd.Col = 4: grd.Text = GetResourceString(43, 37) '"Deposit Date"
    grd.Col = 5: grd.Text = GetResourceString(43, 42) '"Deposited Amount"
    grd.ColAlignment(4) = 1
End If

If m_ReportType = repFDLedger Then 'm_ReportType = repMFDLedger Or
    With grd
        .Clear
        .Cols = 6
        .Rows = 5
        .FixedRows = 1
        .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)    '"Sl NO":
        .Col = 1: .Text = GetResourceString(37)    '"Date":
        .Col = 2: .Text = GetResourceString(284) '"OB"
        .Col = 3: .Text = GetResourceString(271) '"Deposit"
        .Col = 4: .Text = GetResourceString(279)  '"WithDraw"
        .Col = 5: .Text = GetResourceString(285)  '"Closing Balance"
        '.Col = 6: .Text = GetResourceString(273)  '"Cb"
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
'        .ColAlignment(5) = 1
    End With
End If
If m_ReportType = repFDTrans Then 'm_ReportType = repMFDLedger Or
    With grd
        .Clear
        .Cols = 6
        .Rows = 5
        .FixedRows = 1: .FixedCols = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)    '"Sl NO":
        .Col = 1: .Text = GetResourceString(37)    '"Date":
        .Col = 2: .Text = GetResourceString(271) '"Deposit"
        .Col = 3: .Text = GetResourceString(279)  '"WithDraw"
        .Col = 4: .Text = GetResourceString(47)  '"Interst "
        .Col = 5: .Text = GetResourceString(273)  '"Charges"
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
'        .ColAlignment(5) = 1
    End With
End If

If m_ReportType = repFDMonbal Then
    With grd
        .Clear
        .Cols = 4: .Rows = 5
        .FixedCols = 2
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)   ' "Sl No"
        .Col = 1: .Text = GetResourceString(36, 60)   ' "Acc No"
        .Col = 2: .Text = GetResourceString(35)  '"Name"
        .Col = 3: .Text = GetResourceString(43) & _
                        GetResourceString(37) '"Deposit Date"
        
    End With
End If
If m_ReportType = repFDJoint Then
    With grd
        .Clear
        .Cols = 4: .Rows = 5
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)   ' "Sl No"
        .Col = 1: .Text = GetResourceString(36, 60)   ' "Acc No"
        .Col = 2: .Text = GetResourceString(35)  '"Name"
        .Col = 3: .Text = GetResourceString(42) 'Balance '& _
                        GetResourceString(37) '"Deposit Date"
        
    End With
End If

Dim I As Integer
With grd
    For I = 0 To .FixedRows - 1
        .Row = I
        For count = 0 To .Cols - 1
            .MergeCol(count) = True
            .Col = count
            .CellAlignment = 4
            .CellFontBold = True
        Next count
        .MergeRow(I) = True
    Next
End With
        
LastLine:

End Sub


'
Private Sub ShowMonthlyBalance()
Dim count As Integer
Dim StartDate As Date
Dim EndDate As Date
Dim SqlStr As String
Dim rstMain As ADODB.Recordset

'Get Last Day of the FromDate
StartDate = GetSysLastDate(m_FromDate)

'Get Last date of the Todate
EndDate = GetSysLastDate(m_ToDate)

Call InitGrid

SqlStr = "SELECT A.AccNum,A.AccID,Createdate,DepositAmount,Name as CustName" & _
        " From FDMaster A, FDTrans B, QryName C WHERE CreateDate <= #" & EndDate & "#" & _
        " AND (ClosedDate Is NULL Or ClosedDate > #" & EndDate & "#)" & _
        " AND A.AccId = B.AccID AND TransID = 1 AND A.CustomerID = C.CustomerID" & _
        " AND A.DepositType = " & m_DepositType

If m_AccGroup Then SqlStr = SqlStr & " AND AccGroupId = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " ORDER BY val(AccNum),A.Accid "
Else
    gDbTrans.SqlStmt = SqlStr & " ORDER BY IsciName,A.Accid "
End If

If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub

count = DateDiff("M", StartDate, EndDate)
Dim ProcCount As Long
Dim TotalProcCount As Long
Dim rowno As Long, colno As Byte

TotalProcCount = rstMain.RecordCount * (count + 1)
RaiseEvent Initialise(0, TotalProcCount)
ProcCount = 0
rowno = grd.FixedRows - 1

While Not rstMain.EOF
    ProcCount = ProcCount + 1
    With grd
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        .TextMatrix(rowno, 0) = ProcCount
        .TextMatrix(rowno, 1) = FormatField(rstMain("AccNum"))
        .TextMatrix(rowno, 2) = Trim$(FormatField(rstMain("CustName")))
        .TextMatrix(rowno, 3) = FormatField(rstMain("CreateDate"))
    End With
    rstMain.MoveNext
    RaiseEvent Processing("Collecting account information", ProcCount / TotalProcCount)
Wend

rowno = rowno + 2
If grd.Rows < rowno + 1 Then grd.Rows = rowno + 2

grd.Row = rowno
grd.Col = 2: grd.Text = GetResourceString(52, 42)
grd.CellFontBold = True
Dim rst As ADODB.Recordset
Dim TotalBalance As Currency

Do
    If DateDiff("d", StartDate, EndDate) < 0 Then Exit Do
    grd.Cols = grd.Cols + 1
    SqlStr = "SELECT A.AccNum,A.AccID,Balance,TransID From FDMAster A, FdTrans B " & _
        " WHERE B.AccID = A.AccID And TransID = (SELECT MAX(TransID) " & _
        " From FDTrans D Where D.AccID = A.AccID And TransDate <= " & _
        "#" & StartDate & "#)" & _
        " AND A.DepositType = " & m_DepositType '& _
        " AND A.CustomerID = C.CustomerID"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then GoTo NextMonth
    TotalBalance = 0
    grd.Row = grd.FixedRows - 1
    grd.Col = grd.Cols - 1
    
    rstMain.MoveFirst
    grd.Text = GetMonthString(Month(StartDate))
    grd.CellFontBold = True
    grd.CellAlignment = 4
    
    rowno = grd.Row
    colno = grd.Col
    While Not rstMain.EOF
        count = Val(rstMain("AccNum"))
        With grd
            rowno = rowno + 1
            rst.MoveFirst
            rst.Find "AccID = " & rstMain("AccId")
            If Not rst.EOF Then
                .TextMatrix(rowno, colno) = FormatField(rst("Balance"))
                TotalBalance = TotalBalance + Val(.Text)
            End If
        End With

NextAccount:
        rstMain.MoveNext
        ProcCount = ProcCount + 1
'        RaiseEvent Processing("Collecting account information", ProcCount / TotalProcCount)
    Wend
    grd.Row = grd.Rows - 1
    grd.Text = FormatCurrency(TotalBalance)
    grd.CellFontBold = True
NextMonth:
    StartDate = DateAdd("D", 1, StartDate)
    StartDate = DateAdd("M", 1, StartDate)
    StartDate = DateAdd("D", -1, StartDate)
Loop

Set rst = Nothing
Set rstMain = Nothing

End Sub

Private Sub ShowMonthlyBalance_Temp()
Dim DateMin As Date
Dim DateMax As Date
Dim rstMain As ADODB.Recordset
Dim ProcCount As Integer
Dim rst As ADODB.Recordset
Dim TotalBalance As Currency

'Get Last Day of month of FromDate
DateMin = GetSysLastDate(m_FromDate)
'Get Last date of month of Todate
DateMax = GetSysLastDate(m_ToDate)


'create the necessary queries
gDbTrans.SqlStmt = "SELECT A.[CreateDate], A.[AccNum], " & _
                "A.[AccID], Name " & _
                "From FDMaster A, QryName B, FDTrans C " & _
                "WHERE (((A.[ClosedDate]) Is Null Or " & _
                "(A.[ClosedDate])>=#" & DateMin & "#) And " & _
                "((A.[MaturedOn]) Is Null Or " & _
                "(A.[MaturedOn])>=#" & DateMin & "#)) And " & _
                "A.[CustomerID]=B.[CustomerID] And " & _
                "A.[DepositType]=" & m_DepositType & _
                " And A.AccID = C.AccID And C.TransDate <= #" & DateMax & "#;"

If Not gDbTrans.CreateView("QryFDMaster") Then Exit Sub

gDbTrans.SqlStmt = "SELECT Max([TransID]) AS MaxTransID, [B].[AccID] " & _
                "From FDTrans A, QryFDMaster B " & _
                "WHERE A.AccID = B.AccID " & _
                "GROUP BY [B].[AccID], Month([TransDate]);"
If Not gDbTrans.CreateView("QryMonthMaxTransID") Then Exit Sub

gDbTrans.SqlStmt = "SELECT B.AccID, [Balance], A.TransDate, B.MaxTransID " & _
                "FROM FDTrans AS A, QryMonthMaxTransID AS B " & _
                "WHERE A.AccID=B.AccID And A.TransID=B.MaxTransID;"
If Not gDbTrans.CreateView("QryMonthBalance") Then Exit Sub

'initialise the query
Call InitGrid

'find the records and put into grid
gDbTrans.SqlStmt = "SELECT DISTINCT * FROM QryFDMaster ORDER BY AccID"
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub

While Not rstMain.EOF
    ProcCount = ProcCount + 1
    With grd
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = ProcCount
        .Col = 1: .Text = FormatField(rstMain("AccNum"))
        .Col = 2: .Text = Trim$(FormatField(rstMain("Name")))
        .Col = 3: .Text = FormatField(rstMain("CreateDate"))
    End With
    rstMain.MoveNext
'    RaiseEvent Processing("Collecting account information", ProcCount / RstMain.RecordCount)
Wend

With grd
    If .Rows < .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows < .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(52, 42)
    .CellFontBold = True
End With

Do
    If DateDiff("d", DateMin, DateMax) < 0 Then Exit Do
    With grd
        .Cols = .Cols + 1
        .Row = .FixedRows - 1
        .Col = .Cols - 1
        .Text = GetMonthString(Month(DateMin))
        .CellFontBold = True
        .CellAlignment = 4
    End With
    
    gDbTrans.SqlStmt = "SELECT Max(MaxTransID) As TempTransID,Balance,AccID " & _
                "FROM QryMonthBalance WHERE TransDate <= #" & DateMin & _
                "# Group By AccID, Balance"
    If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub
    rstMain.MoveFirst
    TotalBalance = 0
    Do While Not rstMain.EOF
        With grd
            .Row = .Row + 1
            rst.MoveFirst
            rst.Find "AccID = " & rstMain("accID")
            If Not rst.EOF Then
                .Text = FormatField(rst("Balance"))
                TotalBalance = TotalBalance + Val(.Text)
            End If
        End With
        rstMain.MoveNext
    Loop
    grd.Row = grd.Rows - 1
    grd.Text = FormatCurrency(TotalBalance)
    grd.CellFontBold = True
    DateMin = DateAdd("D", 1, DateMin)
    DateMin = DateAdd("M", 1, DateMin)
    DateMin = DateAdd("D", -1, DateMin)
Loop

Set rst = Nothing
Set rstMain = Nothing

End Sub

Public Property Let ToAmount(curTo As Currency)
    m_ToAmt = curTo
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

Private Sub MaturedDeposits()
Dim rst As ADODB.Recordset

lblReportTitle.Caption = GetResourceString(72)   '"Deposits That Mature"

RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim SqlStmt As String

SqlStmt = "Select AccId,AccNum,CreateDate,DepositAmount, " & _
        " MaturityDate,ClosedDate,RateOfInterest,MaturityAmount,Name as CustName" & _
        " FROM FDMaster A, QryName B" & _
        " where MaturityDate BETWEEN #" & m_FromDate & "#" & _
        " AND #" & m_ToDate & "# " & _
        " AND A.CustomerId= B.CustomerId" & _
        " AND A.DepositType = " & m_DepositType

If m_Place <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " And MaturityAmount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " And Maturity <= " & m_ToAmt


If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & " Order By MaturityDate,AccNum,AccID"
Else
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order By MaturityDate,IsciName,AccID"
End If
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

'Init the grid With grd
Dim SlNo As Integer
Dim Days As Long
Dim DepAmt As Currency, MatAmt As Currency
Dim Interest As Double
Dim DepTotal As Currency, MatTotal As Currency
Dim DepDate As Date, MatDate As Date
Dim rowno As Long, colno As Byte
Call InitGrid

RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data to write into the grid .", 0)

grd.Row = 0
rowno = 0: colno = 0
While Not rst.EOF
    'Set next row
    If grd.Rows = grd.Row + 1 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    rowno = rowno + 1
    SlNo = SlNo + 1
    DepDate = rst("CreateDate")
    MatDate = rst("MaturityDate")
    
    Days = DateDiff("D", DepDate, MatDate)
    DepAmt = CCur(FormatField(rst("DepositAmount")))
    Interest = CCur(FormatField(rst("RateOfInterest")))
    MatAmt = CCur(FormatField(rst("MaturityAmount")))
    If MatAmt = 0 Then MatAmt = _
        DepAmt + ComputeFDInterest(DepAmt, DepDate, MatDate, CSng(Interest))
    MatTotal = MatTotal + MatAmt
    DepTotal = DepTotal + DepAmt
    With grd
        .Row = rowno
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum")): .CellAlignment = flexAlignRightCenter
        .Col = 2: .Text = FormatField(rst("CustName")): .CellAlignment = 1
        .Col = 3: .Text = GetIndianDate(MatDate): .CellAlignment = 4
        .Col = 4: .Text = Interest: .CellAlignment = flexAlignRightCenter
        .Col = 5: .Text = FormatCurrency(DepAmt): .CellAlignment = flexAlignRightCenter
    End With
    
    Me.Refresh
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data into the grid.", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend
    
    Set rst = Nothing

'Set last
With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    
    .Col = 1: .CellFontBold = True
    .Text = GetResourceString(52)
    .CellAlignment = flexAlignRightCenter '"Totals"
    .Col = .Cols - 1: .CellFontBold = True
    .Text = FormatCurrency(DepTotal): .CellAlignment = flexAlignRightCenter
End With

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdOK.Caption = GetResourceString(11)       '`"«Œâ³ú"
cmdPrint.Caption = GetResourceString(23)

End Sub

Private Sub ShowDepositBalances()

Dim count As Integer
Dim rst As ADODB.Recordset
Dim SqlStmt As String
Dim StrAmount As String
Dim transType As wisTransactionTypes
Dim TotalAmount As Currency

gDbTrans.SqlStmt = "SELECT Max(TransID) As MaxTransID, A.AccID" & _
    " FROM FDTrans A, FDMaster B" & _
    " WHERE A.TransDate  <= #" & m_ToDate & "#" & _
    " And A.AccID = B.AccID AND B.DepositType = " & _
    m_DepositType & " GROUP BY A.AccID"
If Not gDbTrans.CreateView("QryTemp") Then Exit Sub

SqlStmt = "Select Balance,AccNum,CreateDate, Name" & _
    " From FDTrans A,FDMaster B,QryName C, QryTemp D" & _
    " Where A.TransID = D.MaxTransID And D.AccID = A.AccID AND A.Balance <> 0" & _
    " And A.AccID = B.AccId And C.CustomerId = B.CustomerId"

If m_Place <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroup

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " And A.Amount  > " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " And A.Amount < " & m_ToAmt


'Build the FINAL qUERY &  aSSIGN TO dBcLASS
If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & " Order by val(AccNum),B.CustomerID "
Else
    SqlStmt = SqlStmt & " Order by IsciName,B.CustomerID "
End If

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Call InitGrid
With grd
    .Row = 0: count = 0
    While Not rst.EOF
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1: count = count + 1
        .Col = 0: .Text = Format(count, "00"): .CellAlignment = 7
        .Col = 1: .Text = FormatField(rst("AccNum")): .CellAlignment = flexAlignLeftCenter
        .Col = 2: .Text = FormatField(rst("CreateDate"))
        .Col = 3: .Text = FormatField(rst("Name"))
        .Col = 4: .Text = FormatField(rst("Balance")): .CellAlignment = 7
        TotalAmount = TotalAmount + CCur(FormatField(rst("Balance")))
        rst.MoveNext
    Wend
End With

Set rst = Nothing

DoEvents
If gCancel Then Exit Sub
Me.Refresh

With grd
    .Rows = .Rows + 2
    .Row = .Row + 2: count = 0
    .Col = 2: .Text = GetResourceString(52) & " " & _
                GetResourceString(42): .CellAlignment = 7
    .CellFontBold = True ' "Totals Balance "
    .Col = .Cols - 1: .Text = FormatCurrency(TotalAmount)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
End With

End Sub
Private Sub ShowMFDBalances()

Dim count As Integer
Dim rst As ADODB.Recordset
Dim SqlStmt As String
Dim transType As wisTransactionTypes
Dim TotalAmount As Currency

gDbTrans.SqlStmt = "SELECT Max(TransID) As MaxTransID, A.AccID " & _
    " FROM MatFDTrans A, FDMaster B WHERE " & _
    " A.TransDate  <= #" & m_ToDate & "#" & _
    " AND B.DepositType = " & m_DepositType & _
    " And A.AccID = B.AccID GROUP BY A.AccId"

If Not gDbTrans.CreateView("QryTemp") Then Exit Sub

SqlStmt = "Select A.Balance, B.AccNum, A.AccID,Name" & _
        " From MatFDTrans A,FDMaster B, QryName C, QryTemp D" & _
        " WHERE A.TransID = D.MaxTransID And A.AccID = D.AccID" & _
        " AND B.DepositType = " & m_DepositType & " AND Balance <> 0" & _
        " And A.AccID = B.AccId And B.CustomerId = C.CustomerId"

If m_FromAmt > 0 Then SqlStmt = SqlStmt & " And A.Amount  > " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " And A.Amount < " & m_ToAmt

If m_Place <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroup

'Build the fINAL qUERY &  aSSIGN TO dBcLASS
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & " Order by VAL(B.AccNum),A.AccID,B.CustomerID "
Else
    gDbTrans.SqlStmt = SqlStmt & " Order by IsciName,A.AccID,B.CustomerID "
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
Call InitGrid

grd.Row = 0
While Not rst.EOF
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1: count = 0
        .Col = 0: .Text = FormatField(rst("AccNum")): .CellAlignment = 7
        .Col = 1: .Text = FormatField(rst("Name")): .CellAlignment = 1
        .Col = 2: .Text = FormatField(rst("Balance")): .CellAlignment = 7
    End With
    TotalAmount = TotalAmount + CCur(FormatField(rst("Balance")))


    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    rst.MoveNext
Wend

Set rst = Nothing
With grd
    .Rows = .Rows + 2
    .Row = .Row + 2: count = 0
    .Col = 1: .Text = GetResourceString(52, 42)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(TotalAmount)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
End With

End Sub
Private Sub ShowDepositsOpened()
'Declare variables
Dim Dt As Boolean
Dim Amt As Boolean
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim Total As Currency

RaiseEvent Processing("Reading & Verifying the records ", 0)

lblReportTitle.Caption = GetResourceString(64)
'Build the FINAL SQL
SqlStr = "Select AccId ,AccNum,CreateDate ,MaturityDate," & _
    " DepositAmount,Closeddate ,RateOfInterest ,CertificateNo, Name" & _
    " From FDMaster A, QryName B" & _
    " Where B.CustomerID = A.CustomerID " & _
    " AND A.DepositType = " & m_DepositType

If m_FromIndianDate <> "" Then _
    SqlStr = SqlStr & " AND CreateDate >= #" & m_FromDate & "#"
If m_ToIndianDate <> "" Then _
    SqlStr = SqlStr & " AND CreateDate <= #" & m_ToDate & "#"

If m_FromAmt <> 0 Then _
    SqlStr = SqlStr & " AND DepositAmount >= " & m_FromAmt
If m_ToAmt <> 0 Then _
    SqlStr = SqlStr & " AND DepositAmount <= " & m_ToAmt

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order By CreateDate, val(AccNum),AccId"
Else
    gDbTrans.SqlStmt = SqlStr & " Order By CreateDate, IsciName,AccId"
End If

If m_Place <> "" Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroup

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

Dim count As Integer
Dim AccId As Long
Dim Amount As Currency

Call InitGrid
RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data to write into the grid", 0)

While Not rst.EOF
    'Set next row
    With grd
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1: count = count + 1
        AccId = FormatField(rst("AccID"))
        .Col = 0: .Text = Format(count, "00"): .CellAlignment = 7
        .Col = 1: .Text = FormatField(rst("AccNum")): .CellAlignment = 7
        .Col = 2: .Text = FormatField(rst("Name")): .CellAlignment = 1
        .Col = 3: .Text = FormatField(rst("CertificateNo")): .CellAlignment = 7
        .Col = 4: .Text = FormatField(rst("CreateDate")): .CellAlignment = 4
        .Col = 5: .Text = FormatField(rst("DepositAmount")): .CellAlignment = 7
    End With
    Total = Total + FormatField(rst("DepositAmount"))
nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    If grd.Row Mod 100 = 0 Then _
        RaiseEvent Processing("Writing the data to the grid.", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend

Set rst = Nothing
'Set next row
With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .CellFontBold = True
    .Text = GetResourceString(52, 43)
    .CellAlignment = 1  ' "Total Deposits "
    .Col = .Cols - 1: .CellFontBold = True
    grd.Text = FormatCurrency(Total): .CellAlignment = flexAlignRightCenter
End With

End Sub

Private Sub ShowDepositsClosed()

Dim SqlStmt As String
Dim Dt As Boolean
Dim Amt As Boolean
Dim DtStr As String
Dim AmtStr As String
Dim rst As ADODB.Recordset
Dim Total As Currency
Dim TotalDeposit As Currency

RaiseEvent Processing("Reading & Verifying the records ", 0)

'Build the FINAL SQL
SqlStmt = "Select A.AccID,AccNum,CertificateNo,DepositAmount," & _
    " Amount as PrinAmount,ClosedDate,C.TransID, Name" & _
    " From FDMaster A,QryName B,FDTrans C " & _
    " WHERE A.CustomerId = B.CustomerID AND C.AccID = A.AccId" & _
    " AND A.DepositType = " & m_DepositType & _
    " AND ClosedDate >= #" & m_FromDate & "#" & _
    " AND ClosedDate <= #" & m_ToDate & "#" & _
    " And TransID = (Select Max(TransID) From FDTrans D" & _
            " Where D.AccID = A.AccID)"

If m_Place <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & " Order By val(AccNum)"
Else
    SqlStmt = SqlStmt & " Order By IsciName,val(AccNum)"
End If

gDbTrans.SqlStmt = SqlStmt
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
gDbTrans.SqlStmt = SqlStmt
If Not gDbTrans.CreateView("QryFDClose1") Then Exit Sub


SqlStmt = "Select A.AccID,AccNum,CertificateNo,Name,DepositAmount," & _
    " A.TransID,A.PrinAmount,B.Amount As IntAmount,ClosedDate " & _
    " From QryFDClose1 A Left Join FdIntTrans B " & _
    " ON A.AccId = B.AccID And A.TransID = B.TransID "

gDbTrans.SqlStmt = SqlStmt
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
gDbTrans.SqlStmt = SqlStmt
If Not gDbTrans.CreateView("qryFDClose2") Then Exit Sub

SqlStmt = "Select A.AccID,AccNum,Name,CertificateNo,DepositAmount," & _
    " A.PrinAmount,A.IntAmount,B.Amount As PayableAmount,ClosedDate " & _
    " From QryFDClose2 A Left Join FdIntPayable B " & _
    " ON A.AccId = B.AccID And A.TransID = B.TransID "

gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Initialize the Grid
Dim AccId As Long
Dim DepositType As Long
Dim Amount As Currency

Call InitGrid
RaiseEvent Initialise(0, rst.RecordCount)

With grd
    While Not rst.EOF
        'Get Returned Amount
        Amount = Val(FormatField(rst("PrinAmount"))) + _
            Val(FormatField(rst("IntAmount"))) + Val(FormatField(rst("PayableAmount")))
        'Set next row
        If .Rows = .Row + 1 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 0: .Text = FormatField(rst("AccNum"))
        .CellAlignment = flexAlignRightCenter: .CellAlignment = 1
        .Col = 1: .Text = FormatField(rst("Name"))
        .CellAlignment = 1
        
        .Col = 2: .Text = FormatField(rst("ClosedDate"))
        .CellAlignment = flexAlignRightCenter
        .Col = 3: .Text = FormatField(rst("DepositAmount"))
        .CellAlignment = flexAlignRightCenter
        TotalDeposit = TotalDeposit + Val(.Text)
        .Col = 4: .Text = FormatCurrency(Amount): .CellAlignment = flexAlignRightCenter
        .CellAlignment = flexAlignRightCenter
        Total = Total + Amount
        
        rst.MoveNext
    Wend
    Set rst = Nothing

    
    'Set next row and put the total
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    
    .Col = 1: .Text = GetResourceString(52) & " " & _
            GetResourceString(43)  '"Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalDeposit): .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(Total): .CellFontBold = True
End With

End Sub


Private Sub ShowLiabilities()
'declare the variables
Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim toDate As Date

RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStmt = "Select A.AccID,A.AccNum,DepositAmount,DepositType," & _
        " CreateDate,EffectiveDate,ClosedDate," & _
        " MaturityDate,MaturityAmount,RateOfInterest, Name" & _
        " From FDMaster A ,QryName B Where" & _
        " (ClosedDate > #" & m_ToDate & "# Or ClosedDate Is NULL )" & _
        " And CreateDate <= #" & m_ToDate & "#" & _
        " And B.CustomerId = A.CustomerId " & _
        " AND A.DepositType = " & m_DepositType

If m_Place <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroup

If m_ReportOrder = wisByAccountNo Then
     gDbTrans.SqlStmt = SqlStmt & " Order By val(AccNum)"
Else
     gDbTrans.SqlStmt = SqlStmt & " Order By IsciName"
End If
    
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

'Init the grid

Dim Liability As Currency
Dim GrandTotal As Currency
Dim TotalLiability As Currency

Call InitGrid

RaiseEvent Initialise(0, rst.RecordCount)

Dim CustName As String
grd.Row = 0
While Not rst.EOF
  With grd
    'Set next row
    If .Rows = .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1:
    .Col = 0: .Text = FormatField(rst("AccNum")): .CellAlignment = 7
    .Col = 1: .Text = FormatField(rst("Name"))
    .Col = 2: .Text = FormatField(rst("EffectiveDate"))
        .CellAlignment = flexAlignCenterCenter
    .Col = 3: .Text = FormatField(rst("MaturityDate"))
        .CellAlignment = flexAlignCenterCenter
    If .Text <> "" Then
        .Col = 4: .Text = FormatField(rst("RateOfInterest"))
        .Col = 5: .Text = FormatField(rst("DepositAmount")): .CellAlignment = flexAlignRightCenter
        GrandTotal = GrandTotal + Val(.Text)
        toDate = IIf(DateDiff("d", m_ToDate, rst("MaturityDate")) > 0, m_ToDate, rst("MaturityDate"))
        Liability = FormatField(rst("DepositAmount")) + _
            ComputeFDInterest(Val(FormatField(rst("DepositAmount"))), FormatField(rst("EffectiveDate")), _
                  toDate, FormatField(rst("DepositType")), CSng(FormatField(rst("RateOfInterest"))))
        .Col = 6: .Text = FormatCurrency(Liability): .CellAlignment = flexAlignRightCenter
        
        TotalLiability = TotalLiability + Liability
    End If
  End With
  
  DoEvents
  If gCancel Then rst.MoveLast
  Me.Refresh
  If grd.Row Mod 50 = 0 Then _
      RaiseEvent Processing("Writing the data into the grid .", _
          rst.AbsolutePosition / rst.RecordCount)
  rst.MoveNext

Wend
  
Set rst = Nothing

'Fill In total Liability
With grd
    If .Rows <= .Row + 2 Then .Rows = .Row + 3
    .Row = .Row + 2
    .Col = 1: .Text = GetResourceString(52) & " " & _
            GetResourceString(42): .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    
    .Col = 5
    .Text = FormatCurrency(GrandTotal)
    .CellFontBold = True
    .Col = 6
    .Text = FormatCurrency(TotalLiability)
    .CellFontBold = True
    .CellAlignment = flexAlignRightCenter
End With

End Sub

Private Sub ShowDayBook()

Dim rst As ADODB.Recordset
Dim transType As wisTransactionTypes
Dim count As Integer
Dim CustName As String
Dim SqlStr As String
Dim sqlClause As String

Dim strTableName As String
strTableName = IIf(m_ReportType = repFDDayBook, "FDTrans", "MatFdTrans")

'here check the Condition
sqlClause = ""
If m_Place <> "" Then sqlClause = sqlClause & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then sqlClause = sqlClause & " And Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " And AccGroupID = " & m_AccGroup

SqlStr = "Select 'PRINCIPLE', B.AccNum,val(B.AccNum) as ac, B.AccID," & _
    " TransDate,Amount, TransType, Name " & _
    " from " & strTableName & " A, FDMaster B, QryName C where " & _
    " TransDate BETWEEN #" & m_FromDate & "#" & _
    " And #" & m_ToDate & "# " & _
    " And A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " AND B.DepositType = " & m_DepositType & _
    sqlClause

If m_FromAmt > 0 Then SqlStr = SqlStr & " And Place = " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " And Caste = " & m_ToAmt

If m_ReportType = repFDDayBook Then
    SqlStr = SqlStr & " UNION " & "Select 'INTEREST',B.AccNum, val(B.AccNum) as ac, B.AccID," & _
        " TransDate, Amount, TransType, Name " & _
        " FROM FDIntTrans A, FDMaster B, QryName C where " & _
        " TransDate BETWEEN #" & m_FromDate & "#" & _
        " And #" & m_ToDate & "# " & _
        " And A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
        " AND B.DepositType = " & m_DepositType & _
        sqlClause
    
    SqlStr = SqlStr & " UNION " & "Select 'PAYABLE',B.AccNum, val(B.AccNum) as ac, B.AccID," & _
        " TransDate,Amount,TransType, Name " & _
        " from FDIntPayable A, FDMaster B, QryName C where " & _
        " TransDate BETWEEN #" & m_FromDate & "#" & _
        " And #" & m_ToDate & "# " & _
        " And A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
        " AND B.DepositType = " & m_DepositType & _
        sqlClause
End If

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,ac"
Else
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,IsciName,B.AccNum "
End If
SqlStr = "": sqlClause = ""

If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
Dim MaxColNo As Integer
Dim SubTotal(4 To 9) As Currency
Dim GrandTotal(4 To 9) As Currency
Dim Amount As Currency
Dim TransDate As String
Dim AccId As String
Dim PRINTTotal As Boolean
Dim totalCount As Long
Dim loopCount As Integer
Dim SlNo As Integer
Dim rowno As Long, colno As Byte

MaxColNo = IIf(m_ReportType = repFDDayBook, 9, 7)

totalCount = rst.RecordCount + 2
Call InitGrid
RaiseEvent Initialise(0, totalCount)
loopCount = 1
With grd

  rowno = .FixedRows - 1
  While Not rst.EOF
    'See if you have to calculate sub totals
    If TransDate <> "" And TransDate <> FormatField(rst("TransDate")) Then
        AccId = 0: PRINTTotal = True
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        .Row = rowno: .Col = 3
        .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
        For colno = 4 To MaxColNo
            .Col = colno
            If SubTotal(colno) Then .Text = FormatCurrency(SubTotal(colno))
            GrandTotal(colno) = GrandTotal(colno) + SubTotal(colno)
            SubTotal(colno) = 0
            .CellFontBold = True
        Next colno
    End If
    
    TransDate = FormatField(rst("TransDate"))
    If AccId <> rst("Accid") Then
        AccId = rst("AccID")
        TransDate = FormatField(rst("TransDate"))
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate")): .CellAlignment = 4
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum")): .CellAlignment = flexAlignRightCenter
        .TextMatrix(rowno, 3) = FormatField(rst("Name"))
    End If
    
    
    
    transType = FormatField(rst("TransType"))
    Amount = FormatField(rst("Amount"))
    If rst(0) = "PAYABALE" Then
        colno = 9
    ElseIf rst(0) = "INTEREST" Then
        colno = 8
    Else
        If transType = wDeposit Then colno = 4
        If transType = wContraDeposit Then colno = 5
        If transType = wWithdraw Then colno = 6
        If transType = wContraWithdraw Then colno = 7
    End If
    '.Col = colNo
    .TextMatrix(rowno, colno) = FormatCurrency(Amount): .CellAlignment = flexAlignRightCenter
    SubTotal(colno) = SubTotal(colno) + Amount
    
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
  
    RaiseEvent Processing("Writing the data into the grid .", loopCount / totalCount)
    rst.MoveNext
    loopCount = loopCount + 1
  Wend
  
  Set rst = Nothing
End With
  
With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    .Row = rowno: .Col = 3
    .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
    For count = 4 To MaxColNo
        .Col = count
        If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        .CellAlignment = flexAlignRightCenter
        .CellFontBold = True
    Next count
        
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(286):
        .CellFontBold = True
        For count = 4 To MaxColNo
            .Col = count
            If GrandTotal(count) Then .Text = FormatCurrency(GrandTotal(count))
            .CellAlignment = flexAlignRightCenter
            .CellFontBold = True
        Next count
    End If
End With


If m_ReportType = repFDDayBook Then
    lblReportTitle.Caption = GetResourceString(43)
Else
    grd.Cols = 8
    lblReportTitle.Caption = GetResourceString(220)
End If


lblReportTitle.Caption = lblReportTitle.Caption & " " & GetResourceString(63) & _
                        " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

End Sub

Private Sub ShowCashBook()

Dim rst As ADODB.Recordset
Dim transType As wisTransactionTypes
Dim count As Integer
Dim CustName As String
Dim SqlStr As String
Dim sqlClause As String

'here check the Condition
sqlClause = ""
If m_Place <> "" Then sqlClause = sqlClause & " And Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then sqlClause = sqlClause & " And Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " And AccGroupID = " & m_AccGroup

Dim strTableName As String
strTableName = IIf(m_ReportType = repFDCashBook, " FDTrans ", " MatFDTrans ")

SqlStr = "Select B.AccNum, TransDate,Amount, TransType,VoucherNo, Name " & _
    " from " & strTableName & " A, FDMaster B, QryName C where " & _
    " TransDate BETWEEN #" & m_FromDate & "#" & _
    " And #" & m_ToDate & "# " & _
    " And A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " AND B.DepositType = " & m_DepositType & _
    sqlClause

If m_FromAmt > 0 Then SqlStr = SqlStr & " And Place = " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " And Caste = " & m_ToAmt
    
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,val(B.AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,IsciName,B.AccNum "
End If
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim SubTotal(5 To 6) As Currency
Dim GrandTotal(5 To 6) As Currency
Dim Amount As Currency
Dim TransDate As String
Dim AccNum As String
Dim PRINTTotal As Boolean
Dim totalCount As Long
Dim loopCount As Integer
Dim SlNo As Integer
Dim rowno As Long, colno As Byte


totalCount = rst.RecordCount + 2
Call InitGrid
RaiseEvent Initialise(0, totalCount)

loopCount = 1
With grd
  rowno = .FixedRows - 1
  While Not rst.EOF
    'See if you have to calculate sub totals
    If TransDate <> "" And TransDate <> FormatField(rst("TransDate")) Then
        AccNum = "": PRINTTotal = True
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        .Row = rowno: .Col = 3
        .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
        For count = 5 To 6
            .Col = count
            If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
            GrandTotal(count) = GrandTotal(count) + SubTotal(count)
            SubTotal(count) = 0
            .CellFontBold = True
        Next count
    End If
    TransDate = FormatField(rst("TransDate"))
    If AccNum <> FormatField(rst("AccNUm")) Then
        AccNum = FormatField(rst("AccNum"))
        TransDate = FormatField(rst("TransDate"))
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate")): .CellAlignment = 4
        .TextMatrix(rowno, 2) = AccNum: .CellAlignment = flexAlignRightCenter
        .TextMatrix(rowno, 3) = FormatField(rst("Name"))
        .TextMatrix(rowno, 4) = FormatField(rst("VoucherNo"))
    End If
    
    transType = FormatField(rst("TransType"))
    Amount = FormatField(rst("Amount"))
    .Row = rowno
    count = 6
    If transType = wDeposit Or transType = wContraDeposit Then count = 5
    
    .Col = count
    .Text = FormatCurrency(Amount): .CellAlignment = flexAlignRightCenter
    SubTotal(count) = SubTotal(count) + Amount
    
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
  
    RaiseEvent Processing("Writing the data into the grid .", loopCount / totalCount)
    rst.MoveNext
    loopCount = loopCount + 1
  Wend
  
  Set rst = Nothing
End With
  
With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    .Row = rowno
    .Col = 3: .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
    For count = 5 To 6
        .Col = count
        If SubTotal(count) Then .Text = FormatCurrency(SubTotal(count))
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        .CellAlignment = flexAlignRightCenter
        .CellFontBold = True
    Next count
        
    If PRINTTotal Then
        rowno = rowno + 2
        If .Rows <= rowno + 2 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 3: .Text = GetResourceString(286):
        .CellFontBold = True
        For count = 5 To 6
            .Col = count
            If GrandTotal(count) Then .Text = FormatCurrency(GrandTotal(count))
            .CellAlignment = flexAlignRightCenter
            .CellFontBold = True
        Next count
    End If
End With

If m_ReportType = repFDCashBook Then
    lblReportTitle.Caption = GetResourceString(43)
Else
    lblReportTitle.Caption = GetResourceString(220)
End If

lblReportTitle.Caption = lblReportTitle.Caption & " " & GetResourceString(390, 85) & _
        " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)


End Sub


Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

' Call the print class services...
If m_grdPrint Is Nothing Then Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .ReportTitle = lblReportTitle.Caption
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .Font.Size = gFontSize
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

Private Sub Form_Click()
Call grd_LostFocus
End Sub

Private Sub Form_Load()

'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'now check the Whether thr  Report type has set or not
If m_ReportType = 0 Then
    Err.Raise 50201, "FD Report", "Report type not set"
    Exit Sub
End If

Call SetKannadaCaption

'Init the grid
With grd
    .Clear
    .Cols = 1
    .FixedCols = 0
    .Rows = 20
    .Row = 1
    .Text = GetResourceString(278)
    .CellAlignment = 4: .CellFontBold = True
End With

Me.lblReportTitle.FONTSIZE = 14

'Show report
SelectAndShow

End Sub


Private Function GetCustomerName(ByVal CustomerID As Long) As String
Dim rst As ADODB.Recordset
    GetCustomerName = ""
    gDbTrans.SqlStmt = "Select Title,FirstName,MiddleName, LastName from NameTab " & _
            " Where CustomerId = " & CustomerID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        GetCustomerName = FormatField(rst(0)) + FormatField(rst(1)) & " " & _
        FormatField(rst(2)) & " " & FormatField(rst(3))
    End If
    Set rst = Nothing

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gWindowHandle = 0
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
   cmdOK.Left = fra.Width - cmdOK.Width - (cmdOK.Width / 4)
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - (cmdPrint.Width / 4)
    cmdWeb.Top = cmdPrint.Top
    cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)

   Dim Wid As Single
   Dim I As Integer
   
    With grd
        For I = 0 To .Cols - 1
            Wid = 1 / .Cols
            Wid = GetSetting(App.EXEName, "FDReport" & m_ReportType, "ColWidth" & I, Wid) * .Width
            If Wid >= .Width * 0.95 Then Wid = .Width / .Cols
            .ColWidth(I) = Wid
        Next
         
    End With
    
   Exit Sub
   Wid = (grd.Width - 185) / grd.Cols
   For I = 0 To grd.Cols - 1
       grd.ColWidth(I) = Wid - 30
   Next I
End Sub

Private Sub ShowGeneralLedger()
'Declare variables
Dim count As Integer
Dim rst As ADODB.Recordset
Dim SqlStr As String
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim FdBalance As Currency
Dim SubTotal(2 To 3) As Currency
Dim GrandTotal(2 To 3) As Currency
Dim PRINTTotal As Boolean

Dim SlNo As Integer

RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim strTableName As String
strTableName = IIf(m_ReportType = repMFDLedger, "MatFDTrans", "FDTrans")

SqlStr = "Select 'PRINCIPLE',Sum(Amount)as TotalAmount ,TransDate," & _
    " TransType From FDMaster A, " & strTableName & " B where" & _
    " TransDate BETWEEN #" & m_FromDate & "# And" & _
    " #" & m_ToDate & "# AND A.AccID = B.AccID" & _
    " AND A.DepositType = " & m_DepositType

SqlStr = SqlStr & " Group By TransDate,TransType"
gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

SqlStr = ""

grd.Clear
Call InitGrid

Dim loopCount As Integer
Dim totalCount As Integer
Dim rowno As Long, colno As Byte

totalCount = rst.RecordCount + 2
RaiseEvent Initialise(0, rst.RecordCount)

ReDim TotalAmount(grd.Cols - 1)

TransDate = DateAdd("d", -1, m_FromDate)
Dim AccClass As clsAccTrans
Set AccClass = New clsAccTrans
If m_ReportType = repMFDLedger Then
    FdBalance = AccClass.GetOpBalance(GetIndexHeadID(GetResourceString(220) & " " & GetDepositTypeText(CInt(m_DepositType))), _
            m_FromDate)
Else
'    FdBalance = FdClass.Balance(CDate(TransDate), CInt(m_DepositType))
    FdBalance = AccClass.GetOpBalance(GetIndexHeadID(GetDepositTypeText(CInt(m_DepositType))), _
                    m_FromDate)
End If
Set AccClass = Nothing

With grd
    .Row = .FixedRows
    .Col = 1: .CellFontBold = True: .Text = GetResourceString(284) '"Opening Balnce"
    .Col = 2: .CellFontBold = True: .Text = FormatCurrency(FdBalance)
    TransDate = rst("TransDate")
    rowno = .FixedRows
End With

While Not rst.EOF
    If DateDiff("d", TransDate, rst("TransDate")) Then 'If  The Dates are not Same
        PRINTTotal = True
        With grd
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            TransDate = rst("TransDate")
            
            .TextMatrix(rowno, 2) = FormatCurrency(FdBalance)
            .TextMatrix(rowno, 3) = FormatCurrency(SubTotal(2))
            .TextMatrix(rowno, 4) = FormatCurrency(SubTotal(3))
            
            FdBalance = FdBalance + SubTotal(2) - SubTotal(3)
            .TextMatrix(rowno, 5) = FormatCurrency(FdBalance)

            GrandTotal(2) = GrandTotal(2) + SubTotal(2): SubTotal(2) = 0
            GrandTotal(3) = GrandTotal(3) + SubTotal(3): SubTotal(3) = 0
        End With
    End If

    transType = FormatField(rst("TransType"))
    Amount = FormatField(rst("TotalAmount"))
    
    colno = 2
    If transType = wWithdraw Or transType = wContraWithdraw Then colno = 3
    
    'grd.Col = colNo
    SubTotal(colno) = SubTotal(colno) + Amount
    
    DoEvents
    If gCancel Then rst.MoveLast
    loopCount = loopCount + 1
    RaiseEvent Processing("Writing the data to the grid.", loopCount / totalCount)
    
    rst.MoveNext
Wend
Set rst = Nothing
    
With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    SlNo = SlNo + 1
    
    .Row = rowno
    .TextMatrix(rowno, 0) = SlNo
    .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
    .TextMatrix(rowno, 2) = FormatCurrency(FdBalance)
    .TextMatrix(rowno, 3) = FormatCurrency(SubTotal(2))
    .TextMatrix(rowno, 4) = FormatCurrency(SubTotal(3))
    
    FdBalance = FdBalance + SubTotal(2) - SubTotal(3)
    .TextMatrix(rowno, 5) = FormatCurrency(FdBalance)

    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    .Row = rowno
    .Col = 4: .Text = GetResourceString(285): .CellFontBold = True  '"closing Balnce"
    .Col = 5: .Text = FormatCurrency(FdBalance): .CellFontBold = True

    GrandTotal(2) = GrandTotal(2) + SubTotal(2): SubTotal(2) = 0
    GrandTotal(3) = GrandTotal(3) + SubTotal(3): SubTotal(3) = 0
    
    If PRINTTotal Then
        rowno = rowno + 1
        If .Rows <= rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 3: .Text = FormatCurrency(GrandTotal(2)): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(GrandTotal(3)): .CellFontBold = True
    End If
End With

lblReportTitle.Caption = GetResourceString(IIf(m_ReportType = repFDLedger, 43, 220)) & " " & _
            GetResourceString(93) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)


End Sub

Private Sub ShowDepositTransMade()
'Declare variables
Dim count As Integer
Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim SqlStr As String
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim FdBalance As Currency
Dim SubTotal(2 To 5) As Currency
Dim GrandTotal(2 To 5) As Currency
Dim PRINTTotal As Boolean

RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStr = "Select 'PRINCIPLE',Sum(Amount)as TotalAmount,TransDate," & _
    " TransType From FDMaster A, FDTrans B where " & _
    " TransDate BETWEEN #" & m_FromDate & "# And " & _
    "#" & m_ToDate & "# AND A.AccID = B.AccID AND A.DepositType = " & m_DepositType

SqlStr = SqlStr & " Group By TransDate,TransType"

SqlStr = SqlStr & " UNION " & "Select 'INTEREST',Sum(Amount)as TotalAmount,TransDate," & _
    " TransType From FDMaster A, FDIntTrans B where " & _
    " TransDate BETWEEN #" & m_FromDate & "# And #" & m_ToDate & "#" & _
    " AND A.AccID = B.AccID AND A.DepositType = " & m_DepositType
SqlStr = SqlStr & " Group By TransDate,TransType"

gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

grd.Clear
Call InitGrid

Dim loopCount As Integer
Dim totalCount As Integer
totalCount = rst.RecordCount + 2
RaiseEvent Initialise(0, rst.RecordCount)

ReDim TotalAmount(grd.Cols - 1)

TransDate = DateAdd("d", -1, m_FromDate)
Dim FdClass As clsFDAcc
Set FdClass = New clsFDAcc

FdBalance = FdClass.Balance(TransDate, CInt(m_DepositType))
Set FdClass = Nothing

While Not rst.EOF
    If DateDiff("d", TransDate, rst("TransDate")) Then 'If  The Dates are not Same
        With grd
            PRINTTotal = True
            If .Rows <= .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 0: .Text = .Row
            .Col = 1
            .Text = GetIndianDate(TransDate)
            TransDate = rst("TransDate")
            FdBalance = FdBalance + SubTotal(2) - SubTotal(3)
            For count = 2 To 5
                .Col = count
                .Text = FormatCurrency(SubTotal(count)): .CellAlignment = flexAlignRightCenter
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End With
    End If

    transType = FormatField(rst("TransType"))
    Amount = FormatField(rst("TotalAmount"))
    If transType = wDeposit Or transType = wContraDeposit Then
        count = 2
        If rst(0) = "INTEREST" Then count = 5
    ElseIf transType = wWithdraw Or transType = wContraWithdraw Then
        count = 3
        If rst(0) = "INTEREST" Then count = 4
    End If
    
    grd.Col = count
    SubTotal(count) = SubTotal(count) + Amount
    
    DoEvents
    If gCancel Then rst.MoveLast
    loopCount = loopCount + 1
    RaiseEvent Processing("Writing the data to the grid.", loopCount / totalCount)
    
    rst.MoveNext
Wend


Set rst = Nothing
    
With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .Text = GetIndianDate(TransDate)
    FdBalance = FdBalance + SubTotal(2) - SubTotal(3)
    For count = 2 To 5
        .Col = count
        .Text = FormatCurrency(SubTotal(count)): .CellAlignment = flexAlignRightCenter
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
    Next
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        For count = 2 To 5
            .Col = count
            .Text = FormatCurrency(GrandTotal(count)): .CellAlignment = flexAlignRightCenter
            .CellFontBold = True
        Next
    End If

End With


End Sub

Private Sub ShowMatDepositTransMade()
'Declare variables
Dim count As Integer
Dim SqlStmt As String
Dim rst As ADODB.Recordset
Dim SqlStr As String

RaiseEvent Processing("Reading & Verifying the records ", 0)
SqlStr = "Select 'PRINCIPLE',Sum(Amount)as TotalAmount ,TransDate," & _
    " TransType From MatFDTrans where " & _
    " TransDate >= #" & m_FromDate & "# And " & _
    "TransDate <= #" & m_ToDate & "#"
SqlStr = SqlStr & " Group By TransDate,TransType"

gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

grd.Clear
Dim Amount As Currency
Dim TransDate As Date
Dim transType As wisTransactionTypes
Dim FdBalance As Currency
Dim SubTotal(1 To 4) As Currency
Dim GrandTotal(1 To 4) As Currency
Dim PRINTTotal As Boolean
Call InitGrid
RaiseEvent Initialise(0, rst.RecordCount)

ReDim TotalAmount(grd.Cols - 1)
TransDate = DateAdd("d", -1, m_FromDate)

Dim FdClass As clsFDAcc
Dim rowno As Long

Set FdClass = New clsFDAcc

FdBalance = FdClass.BalanceMaturedFD(TransDate, CInt(m_DepositType))
Set FdClass = Nothing

With grd
    .Row = 1
    .Col = 0: .CellFontBold = True: .Text = GetResourceString(284)
        .CellAlignment = 1     '"Opening Balnce"
    .Col = 1: .CellFontBold = True: .Text = FormatCurrency(FdBalance)
        .CellAlignment = flexAlignRightCenter
    TransDate = rst("TransDate")
    rowno = .FixedRows
End With

While Not rst.EOF
    If DateDiff("d", TransDate, rst("TransDate")) Then 'If  The Dates are not Same
        With grd
            PRINTTotal = True
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1
            .TextMatrix(rowno, 0) = GetIndianDate(TransDate)
            TransDate = rst("TransDate")
            FdBalance = FdBalance + SubTotal(1) - SubTotal(2)
            For count = 1 To 4
                .TextMatrix(rowno, count) = FormatCurrency(SubTotal(count))
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
        End With
    End If

    transType = rst("TransType")
    Amount = FormatField(rst("TotalAmount"))
    If transType = wDeposit Or transType = wContraDeposit Then
        count = 1
        If rst(0) = "INTEREST" Then count = 4
    ElseIf transType = wWithdraw Or transType = wContraWithdraw Then
        count = 2
        If rst(0) = "INTEREST" Then count = 3
    End If
    
    SubTotal(count) = SubTotal(count) + Amount
    
    DoEvents
    If gCancel Then rst.MoveLast
    If grd.Row Mod 50 = 0 Then _
        RaiseEvent Processing("Writing the data to the grid.", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

Set rst = Nothing
    
With grd
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    .TextMatrix(rowno, 0) = GetIndianDate(TransDate)
    FdBalance = FdBalance + SubTotal(1) - SubTotal(2)
    For count = 1 To 4
        .TextMatrix(rowno, count) = FormatCurrency(SubTotal(count))
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
    Next
    If PRINTTotal Then
        rowno = rowno + 2
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        
        .Row = rowno
        For count = 1 To 4
            .Col = count
            .Text = FormatCurrency(GrandTotal(count))
            .CellFontBold = True
        Next
    End If
    If .Rows <= rowno + 2 Then .Rows = rowno + 2
    rowno = rowno + 1
    .Row = rowno
    .Col = 0: .CellFontBold = True: .Text = GetResourceString(285): .CellAlignment = 1     '"closing Balnce"
    .Col = 2: .CellFontBold = True: .Text = FormatCurrency(FdBalance): .CellAlignment = flexAlignRightCenter
End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmFDReport = Nothing
End Sub

Private Sub grd_LostFocus()
Dim I As Integer
Dim Wid As Single
With grd
For I = 0 To .Cols - 1
    Wid = Format(.ColWidth(I) / .Width, "##.####")
    Call SaveSetting(App.EXEName, "FDReport" & m_ReportType, "ColWidth" & I, Wid)
Next I
End With
End Sub

Private Sub ShowMatDepDayBook()

Dim rst As ADODB.Recordset
Dim transType As wisTransactionTypes
Dim count As Integer
Dim CustName As String
Dim SqlStr As String
Dim sqlClause As String

'To Get Deposits & Payments of of PD Account
sqlClause = ""
If m_Place <> "" Then sqlClause = sqlClause & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then sqlClause = sqlClause & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then sqlClause = sqlClause & " AND Gender = " & m_Gender
If m_AccGroup Then sqlClause = sqlClause & " And AccGroupID = " & m_AccGroup

SqlStr = "Select 'PRINCIPLE', A.AccId,B.AccNum, TransDate,Amount, TransType,Name " & _
    " FROM MatFDTrans A, FDMaster B, QryName C where " & _
    " TransDate >= #" & m_FromDate & "#" & _
    " And TransDate <= #" & m_ToDate & "# " & _
    " And A.AccId = B.AccId And B.CustomerId = C.CustomerID " & _
    " AND B.DepositType = " & m_DepositType

If m_FromAmt > 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " AND Amount <= " & m_ToAmt

SqlStr = SqlStr & sqlClause
    
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,B.AccNum,A.AccID "
Else
    gDbTrans.SqlStmt = SqlStr & " Order by TransDate,IsciName, A.AccID "
End If
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim SubTotal(3 To 8) As Currency
Dim GrandTotal(3 To 8) As Currency
Dim Amount As Currency
Dim TransDate As String
Dim AccNum As String
Dim PRINTTotal As Boolean

Call InitGrid

Dim loopCount As Integer
Dim totalCount As Integer
totalCount = rst.RecordCount + 2
RaiseEvent Initialise(0, rst.RecordCount)

With grd
  .Row = .FixedRows
  While Not rst.EOF
    'See if you have to calculate sub totals
    If TransDate <> "" And TransDate <> FormatField(rst("TransDate")) Then
        AccNum = "": PRINTTotal = True
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 2: .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
        For count = 3 To 8
            .Col = count
            .Text = FormatCurrency(SubTotal(count))
            GrandTotal(count) = GrandTotal(count) + SubTotal(count)
            SubTotal(count) = 0
            .CellFontBold = True
        Next count
    End If
    TransDate = FormatField(rst("TransDate"))
    If AccNum <> FormatField(rst("AccNUm")) Then
        AccNum = FormatField(rst("AccNum"))
        TransDate = FormatField(rst("TransDate"))
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = FormatField(rst("TransDate")): .CellAlignment = 4
        .Col = 1: .Text = AccNum: .CellAlignment = flexAlignRightCenter
        .Col = 2: .Text = FormatField(rst("Name"))
    End If
    If Val(FormatField(rst("Amount"))) = 0 Then GoTo nextRecord
    
    transType = rst("TransType")
    Amount = FormatField(rst("Amount"))
    If rst(0) = "PAYABALE" Then
        count = 8
    ElseIf rst(0) = "INTEREST" Then
        count = 7
    Else
        If transType = wDeposit Then count = 3
        If transType = wContraDeposit Then count = 4
        If transType = wWithdraw Then count = 5
        If transType = wContraWithdraw Then count = 6
    End If
    .Col = count
    .Text = FormatCurrency(Amount): .CellAlignment = flexAlignRightCenter
    SubTotal(count) = SubTotal(count) + Amount
nextRecord:
    
    loopCount = loopCount + 1
    RaiseEvent Processing("Writing the data into the grid .", loopCount / totalCount)
    DoEvents
    If gCancel Then rst.MoveLast
    
    rst.MoveNext
    Me.Refresh
  Wend
  
  Set rst = Nothing
End With
  
With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(304): .CellFontBold = True: .CellAlignment = flexAlignRightCenter
    For count = 3 To 8
        .Col = count
        .Text = FormatCurrency(SubTotal(count))
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        .CellAlignment = flexAlignRightCenter
        .CellFontBold = True
    Next count
        
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 2: .Text = GetResourceString(286):
        For count = 3 To 8
            .Col = count
            .Text = FormatCurrency(SubTotal(count))
            .Text = GrandTotal(count)
            .CellAlignment = flexAlignRightCenter
            .CellFontBold = True
        Next count
    End If
End With

End Sub


Public Property Let DepositType(ByVal NewValue As Long)
    m_DepositType = NewValue
End Property

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


