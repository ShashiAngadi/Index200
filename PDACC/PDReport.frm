VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPDReport 
   Caption         =   "Pigmy Deposit Reports ..."
   ClientHeight    =   6000
   ClientLeft      =   1725
   ClientTop       =   1845
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1320
      TabIndex        =   1
      Top             =   5280
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   3720
         TabIndex        =   6
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   2160
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
      Begin VB.CheckBox chkAgent 
         Caption         =   "Show Agent Name"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4785
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8440
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
      Left            =   2640
      TabIndex        =   5
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmPDReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_FromIndianDate As String
Dim m_ToIndianDate As String
Dim m_FromDate As Date
Private m_ToDate As Date
Private m_AccID As Long
Private m_AgentID As Integer

Dim m_FromAmt As Currency
Dim m_ToAmt As Currency
Dim m_Gender As Integer
Dim m_Caste As String
Dim m_Place As String
Dim m_AgentNameShow As Boolean
Dim m_ReportOrder As wis_ReportOrder
Dim m_ReportType As wis_PDReports
Dim m_AccGroup As Integer


Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private WithEvents m_frmCancel As frmCancel
Attribute m_frmCancel.VB_VarHelpID = -1



Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroup = NewValue
End Property


Public Property Let AgentID(NewValue As Integer)
    m_AgentID = NewValue
End Property

Public Property Let Caste(NewCaste As String)
m_Caste = NewCaste
End Property

Public Property Let DisplayAgentName(NewValue As Boolean)
    m_AgentNameShow = NewValue
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

Private Sub ShowDailyTransaction_PrintMonthHeader(fromDate As Date)
    With grd
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = GetResourceString(36, 60)
        .Col = 2: .Text = GetResourceString(35)
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = GetResourceString(36, 60)
        .Col = 2: .Text = GetResourceString(35)
        
        'Move to Previous ROw of print header
        .Row = .Row - 1
        .MergeRow(.Row) = True
        Dim colno As Integer
        For colno = 1 To 31
            .Col = colno + 3
            .CellAlignment = vbCenter
            .TextMatrix(.Row, colno + 3) = GetResourceString(410, 38) & " " & GetFromDateString(GetIndianDate(fromDate), GetIndianDate(DateAdd("D", -1, DateAdd("m", 1, fromDate))))
            .TextMatrix(.Row + 1, .Col) = colno
        Next
        
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        'If .Rows < .Row + 2 Then .Rows = .Rows + 1
        '.Row = .Row + 1
    End With
End Sub

Public Property Let ReportOrder(NewRP As wis_ReportOrder)
    m_ReportOrder = NewRP
End Property

Public Property Let ReportType(newRT As wis_PDReports)
    m_ReportType = newRT
End Property

Private Sub ShowDailyTransaction_AccountTotal(custRst As Recordset, currRowNo As Long, accountWithdraw As Currency, accountDeposit As Currency, accountBalance As Currency)
        'custRst.Filter = "AccId = " & AccId & " and Month1 = " & Month(Rst("TransDate")) & " and Year1 = " & Year(Rst("TransDate"))
        If accountDeposit > 0 Then grd.TextMatrix(currRowNo, 35) = FormatCurrency(accountDeposit)
        If accountWithdraw > 0 Then grd.TextMatrix(currRowNo, 37) = FormatCurrency(accountWithdraw)
        If Not custRst.EOF Then
            accountBalance = FormatField(custRst("Balance"))
            'grd.TextMatrix(currRowNo, 1) = FormatField(custRst("AgentId"))
            grd.TextMatrix(currRowNo, 1) = FormatField(custRst("AccNum"))
            grd.TextMatrix(currRowNo, 2) = FormatField(custRst("CustName"))
            
            grd.TextMatrix(currRowNo, 36) = FormatCurrency(accountBalance - accountDeposit + accountWithdraw)
            grd.TextMatrix(currRowNo, 38) = FormatCurrency(accountBalance)
        End If
        custRst.Filter = adFilterNone
        accountBalance = 0: accountDeposit = 0: accountWithdraw = 0
End Sub

Private Function ShowDailyTransaction_MonthTotal(firstCol As Integer, firstRowNo As Integer, currRowNo As Long) As Currency
    Dim accountBalance As Currency
    Dim rowLoop As Long
    Dim colLoop As Integer
    
    If grd.Rows < currRowNo + 3 Then grd.Rows = currRowNo + 4
    grd.TextMatrix(currRowNo, 2) = GetResourceString(52)
    For colLoop = firstCol To 31 + firstCol
        accountBalance = 0
        For rowLoop = firstRowNo To currRowNo - 1
            accountBalance = accountBalance + Val(grd.TextMatrix(rowLoop, colLoop))
        Next
       
        grd.TextMatrix(currRowNo, colLoop) = FormatCurrency(accountBalance)
    Next

    ShowDailyTransaction_MonthTotal = accountBalance
End Function


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
    m_FromIndianDate = NewStrdate
    m_FromDate = GetSysFormatDate(NewStrdate)
    'm_FromIndianDate = GetAppFormatDate(m_FromDate)
    
End Property

Private Sub ShowAgentTransaction()

Dim SqlStmt As String
Dim rst As Recordset
Dim count As Integer
Dim transType As wisTransactionTypes
Dim PigmyCommission As Single
    
    RaiseEvent Processing("Reading & Verifying the data ", 0)
    transType = wDeposit
    gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount,AgentId,TransDate  " & _
            " From AgentTrans Where TransDate >= #" & m_FromDate & "# " & _
            " and TransDate <= #" & m_ToDate & "# " & _
            " Group By AgentId, TransDate "
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
    chkAgent.Enabled = False
Call InitGrid


RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Dim l_AgentID As Integer
Dim AgentAmount As Currency
Dim TotalAmount As Currency
Dim PigmyAmount As Currency

Dim SetupClass As New clsSetup
PigmyCommission = SetupClass.ReadSetupValue("PDAcc", "PigmyCommission", "03")
If PigmyCommission > 1 Then PigmyCommission = PigmyCommission / 100
Set SetupClass = Nothing

Dim rowno As Long
rowno = grd.Row
While Not rst.EOF
    With grd
        If .Rows <= rowno + 2 Then .Rows = .Rows + 2
        rowno = rowno + 1
        If l_AgentID <> FormatField(rst("AgentID")) Then
            .Row = rowno
            If .Rows <= rowno + 1 Then .Rows = .Rows + 1
            If l_AgentID <> 0 Then
                .Col = 0: .Text = GetResourceString(304) '"Sub Total"
                .CellAlignment = 7: .CellFontBold = True
                .Col = 2: .Text = FormatCurrency(PigmyAmount)
                .CellAlignment = 7: .CellFontBold = True
                .Col = 3: .Text = FormatCurrency(AgentAmount)
                .CellAlignment = 7: .CellFontBold = True
                TotalAmount = TotalAmount + PigmyAmount: PigmyAmount = 0
                AgentAmount = 0
            Else
                .Row = 0: rowno = 0
            End If
            If .Rows = rowno + 1 Then .Rows = .Rows + 2
            rowno = rowno + 1
            .Row = rowno
            l_AgentID = Val(FormatField(rst("AgentId")))
            .Col = 0: .Text = GetAgentName(CLng(l_AgentID))
            .CellFontBold = True
        End If
        .TextMatrix(rowno, 1) = FormatField(rst("TransDate"))
        .TextMatrix(rowno, 2) = FormatField(rst("TotalAmount"))
        .TextMatrix(rowno, 3) = FormatCurrency(FormatField(rst("TotalAmount")) * PigmyCommission)
    End With
    
    AgentAmount = AgentAmount + Val(grd.TextMatrix(rowno, 3))
    PigmyAmount = PigmyAmount + FormatField(rst("totalAmount"))
    
nextRecord:
    rst.MoveNext
    
    DoEvents
    Me.Refresh

    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
Wend

With grd
    If .Rows <= rowno + 1 Then .Rows = rowno + 2
    .Row = rowno + 1
    .Col = 0: .Text = GetResourceString(52)
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(PigmyAmount): .CellFontBold = True: .CellAlignment = 7
    .Col = 3: .Text = FormatCurrency(AgentAmount): .CellFontBold = True: .CellAlignment = 7
    TotalAmount = TotalAmount + PigmyAmount
        
  If PigmyAmount <> TotalAmount Then
    If .Rows <= .Row + 2 Then .Rows = .Row + 3
    .Row = .Row + 2
    .Col = 0: .Text = GetResourceString(286) '"Grand Total"
    .CellFontBold = True
    .Col = 2: .Text = FormatCurrency(TotalAmount): .CellAlignment = 7
    .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalAmount * PigmyCommission): .CellAlignment = 7
    .CellFontBold = True
  End If
End With

End Sub
Private Sub InitGrid()
gCancel = 0

Dim count As Integer
Dim ColWid As Single
Dim I As Integer
    'ColWid = (grd.Width - 200) / grd.Cols + 1
With grd
    .Clear
    .Rows = 20
    .Cols = 2
    .FixedCols = 0
End With
        
On Error Resume Next
If m_ReportType = repPDMonTrans Then
    
    With grd
        .Cols = 5
        'If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        count = 0
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = 0: .Text = GetResourceString(33) '"sL No"
        .Col = 1: .Text = GetResourceString(330, 35)
        .Col = 2: .Text = GetResourceString(36, 60) '"Account No"
        .Col = 3: .Text = GetResourceString(35) '"Name"
        Dim TransDate As Date
        TransDate = m_FromDate
        Do
            .Cols = .Cols + 1
            .Col = .Col + 1
            .Text = GetMonthString(Month(TransDate))
            TransDate = DateAdd("m", 1, TransDate)
            If TransDate > m_ToDate Then Exit Do
        Loop
    End With
    
    GoTo BoldLine
End If

If m_ReportType = repPDBalance Then
    With grd
        .Cols = 4
        If chkAgent.Value = vbChecked Then .Cols = .Cols + 1
        count = 0
        .Row = 0
        .FixedRows = 1
        .FixedCols = 1
        .Col = count: .Text = GetResourceString(33): count = count + 1 '"sL No"
        .Col = count: .Text = GetResourceString(36, 60): count = count + 1 '"Account No"
        .Col = count: .Text = GetResourceString(35): count = count + 1 '"Name"
        If m_AgentNameShow Then
            .Col = count
            .Text = GetResourceString(330, 35): count = count + 1 '"Agent Name"
        End If
        .Col = count: .Text = GetResourceString(42): count = count + 1    '"Balance"
    End With
    GoTo BoldLine
End If
    
If m_ReportType = repPDLedger Then
    With grd
        .Clear
        .Cols = 6: .Rows = 10
        .FixedCols = 1: .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)  '"Slno":
        .Col = 1: .Text = GetResourceString(37)  '"Date":
        .Col = 2: .Text = GetResourceString(284) 'Opening balance
        .Col = 3: .Text = GetResourceString(271) 'Withdraw
        .Col = 4: .Text = GetResourceString(272) '"Repayment"
        .Col = 5: .Text = GetResourceString(285) '"Closing balance"
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDDayBook Then
    With grd
        .Clear
        .Cols = 10
        .MergeCells = flexMergeFree
        .FixedCols = 1
        .FixedRows = 2
        Dim TmpStr As String
        .Row = 0: count = 0
        .MergeRow(0) = True
        .Col = count: .Text = GetResourceString(33): count = count + 1 ' "Sl NO"
        .Col = count: .Text = GetResourceString(37): count = count + 1 ' "Date"
        .Col = count: .Text = GetResourceString(36, 60): count = count + 1 '"Acc NO"
        .Col = count: .Text = GetResourceString(35): count = count + 1 '"Name"
        If chkAgent Then
            .Cols = .Cols + 1
            .Col = count: .Text = GetResourceString(330) & " " & _
                GetResourceString(35): count = count + 1 '"Agent Name"
        End If
        .Col = count: .Text = GetResourceString(271): count = count + 1 '"Deposit"
        .Col = count: .Text = GetResourceString(271): count = count + 1 '"Deposit"
        .Col = count: .Text = GetResourceString(272): count = count + 1  '"Payment"
        .Col = count: .Text = GetResourceString(272): count = count + 1  '"Payment"
        .Col = count: .Text = GetResourceString(274): count = count + 1 '"Interest"
        .Col = count: .Text = GetResourceString(274): count = count + 1 '"Interest"
        .Row = 1
        .MergeRow(2) = True
        For count = 0 To .Cols - 1
            .MergeCol(count) = True
            .Row = 0
            .Col = count
            .CellAlignment = 4: .CellFontBold = True
            TmpStr = .Text
            .Row = 1
            .Text = TmpStr
            .Col = count
            .CellAlignment = 4: .CellFontBold = True
        Next
        
        I = 0: .Row = 1
        For count = .Cols - 1 To .Cols - 6 Step -1
            I = I + 1
            .Col = count
            .MergeCol(count) = False
            .Text = GetResourceString(269 + I Mod 2)
        Next
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDCashBook Then
    With grd
        .Clear
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0: count = 0
        .MergeRow(0) = True
        .Col = 0: .Text = GetResourceString(33)   ' "Sl NO"
        .Col = 1: .Text = GetResourceString(37)   ' "Date"
        .Col = 2: .Text = GetResourceString(36, 60)  '"Acc NO"
        .Col = 3: .Text = GetResourceString(35)   '"Name"
        .Col = 4: .Text = GetResourceString(41):  'Voucher No
        .Col = 5: .Text = GetResourceString(271)  '"Deposit"
        .Col = 6: .Text = GetResourceString(272)  '"Payment"
        .Row = 1
        .MergeRow(2) = True
        For count = 0 To .Cols - 1
            .MergeCol(count) = True
            .Row = 0
        Next
    End With
    
    GoTo BoldLine
End If
    
    
If m_ReportType = repPDAccClose Then
    With grd
        .Cols = 5
        .FixedCols = 1
        .Row = 0: count = 0
        .Col = count: .Text = GetResourceString(33): count = count + 1 '"Sl No"
        .Col = count: .Text = GetResourceString(36, 60): count = count + 1 '"AccNo"
        .Col = count: .Text = GetResourceString(35): count = count + 1 '"Name"
        If chkAgent.Value = vbChecked Then
            .Cols = .Cols + 1
            .Col = count: .Text = GetResourceString(330, 35): count = count + 1 '"AgentName"
        End If
        .Col = count: .Text = GetResourceString(282): count = count + 1 '"Closed Date"
        .Col = count: .Text = GetResourceString(292): count = count + 1 '"MaturedAmount"
        .Row = 0
    End With
    
    GoTo BoldLine

End If
    
If m_ReportType = repPDAccOpen Then
    With grd
        .Rows = 25
        .Cols = 4
        If chkAgent Then .Cols = .Cols + 1
        .FixedCols = 0
        .WordWrap = True
        .Row = 0: count = 0
        .Col = count: .Text = GetResourceString(36, 60): count = count + 1: .ColWidth(count) = .Width / .Cols '"AccNo"
        .Col = count: .Text = GetResourceString(35): count = count + 1: .ColWidth(count) = .Width / .Cols '"Name"
        If chkAgent.Value = vbChecked Then
            .Col = count: .Text = GetResourceString(330, 35): count = count + 1: .ColWidth(count) = .Width / .Cols '"Agent Name"
        End If
        .Col = count: .Text = GetResourceString(281): count = count + 1: .ColWidth(count) = .Width / .Cols '"CreateDate"
        .Col = count: .Text = GetResourceString(226): count = count + 1: .ColWidth(count) = .Width / .Cols '"Deposited Amount"
    End With
    GoTo BoldLine
End If
        
If m_ReportType = repPDAgentTrans Then
    With grd
        .Cols = 4
        .Row = 0
        .Col = 0: .Text = GetResourceString(330, 35) '"Agent Name"
        .Col = 1: .Text = GetResourceString(38) + GetResourceString(37) '"Transaction Date"
        .Col = 2: .Text = GetResourceString(40)   '"Amount Collected"
        .Col = 3: .Text = GetResourceString(328)   '"Pigmy Commission"
    End With
    GoTo BoldLine
End If

If m_ReportType = repPDMonBal Then
    With grd
        .Clear
        .Rows = 5: .Cols = 4
        .FixedRows = 2: .FixedCols = 1
        .Cols = 4 + DateDiff("M", m_FromDate, m_ToDate) * 2
        If .Cols = 4 Then .Cols = 6
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = "Agent ID"
        .Col = 2: .Text = GetResourceString(36, 60)
        .Col = 3: .Text = GetResourceString(35)
        .Row = 1
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = "Agent ID"
        .Col = 2: .Text = GetResourceString(36, 60)
        .Col = 3: .Text = GetResourceString(35)
        count = Month(m_FromDate)
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        Do
            .Col = .Col + 1
            .Row = 0: .Text = GetMonthString(count)
            .Row = 1: .Text = GetResourceString(271) 'Deposit
            .Col = .Col + 1
            
            .Row = 0: .Text = GetMonthString(count)
            .Row = 1: .Text = GetResourceString(289) 'With draw
            If .Col = .Cols - 1 Then Exit Do
            count = count + 1
        Loop
    End With
End If
    
    
If m_ReportType = repPDDailyTrans Then
    With grd
        .Clear
        .Rows = 15: .Cols = 39
        .FixedRows = 2: .FixedCols = 2
        
        
        .Row = 0
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = GetResourceString(36, 60)
        .Col = 2: .Text = GetResourceString(35)
        .Col = 35: .Text = GetResourceString(271) 'Deposit
        .Col = 36: .Text = GetResourceString(250, 42) 'Prev Balance 374 Current
        .Col = 37: .Text = GetResourceString(272) 'WithDraw
        .Col = 38: .Text = GetResourceString(52, 42) 'Prev Balance 374 Current
        .MergeRow(0) = True
        .Row = 1
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = GetResourceString(36, 60)
        .Col = 2: .Text = GetResourceString(35)
        .Col = 35: .Text = GetResourceString(271) 'Deposit
        .Col = 36: .Text = GetResourceString(250, 42) 'Prev Balance 374 Current
        .Col = 37: .Text = GetResourceString(272) 'WithDraw
        .Col = 38: .Text = GetResourceString(52, 42) 'Prev Balance 374 Current
        Dim colno As Integer
        .Row = 0
        For colno = 1 To 31
            .Col = colno + 3
            .CellAlignment = vbCenter
            .TextMatrix(.Row, colno + 3) = GetResourceString(410, 38) & " " & GetFromDateString(GetIndianDate(GetSysFirstDate(m_FromDate)), GetIndianDate(DateAdd("D", -1, DateAdd("m", 1, GetSysFirstDate(m_FromDate)))))
            .TextMatrix(.Row + 1, colno + 3) = colno
        Next
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(35) = True
        .MergeCol(36) = True
        .MergeCol(37) = True
        .MergeCol(38) = True
        .MergeCells = flexMergeFree
    End With
End If

BoldLine:

With grd
    .Row = 0
    Do
        If .Row = .FixedRows Then Exit Do
        For count = 0 To .Cols - 1
            .Col = count
            .CellAlignment = 4
            .CellFontBold = True
        Next count
        .Row = .Row + 1
    Loop
End With

    Exit Sub

ExitLine:

With grd
    ColWid = 0
    For count = 0 To .Cols - 2
        ColWid = ColWid + .ColWidth(count)
        '.CellFontBold = True
    Next count
    .ColWidth(grd.Cols - 1) = .Width - ColWid - TextWidth(.ScrollBars) - 250
End With

End Sub

Private Sub MaturedDeposits()
Dim count As Integer
Dim rst As Recordset
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Balance < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And B.Gender= " & m_Gender
If m_AgentID Then strClause = strClause & " And A.AgentID = " & m_AgentID

gDbTrans.SqlStmt = "Select AccId,AccNum, AgentId,CreateDate,MaturityDate," & _
    " ClosedDate,RateOfInterest, B.Name from PDMaster A Inner join " & _
    " QryName B On A.CustomerId= B.CustomerId " & _
    " where MaturityDate Between #" & m_FromDate & "# " & _
    " and #" & m_ToDate & "# " & strClause

If chkAgent.Value Then
        
    gDbTrans.SqlStmt = "Select AccId,A.AccNum,A.AgentId,A.CreateDate,MaturityDate," & _
        " A.ClosedDate,RateOfInterest, B.Name, C.Name  as AgentName " & _
        " From QryName B Inner join (PDMaster A " & _
            " Inner join (UserTab D inner join QryName C " & _
        " On C.CustomerId = D.CustomerId )On A.AgentId = D.UserId) " & _
        " On A.CustomerId= B.CustomerId " & _
        " Where MaturityDate Between #" & m_FromDate & "#  " & _
        " And #" & m_ToDate & "# " & strClause
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub


'initialise the grid.
With grd
    count = 0
    .Clear
    .Rows = 2: .Rows = 25
    .Cols = 5
    If chkAgent Then .Cols = .Cols + 1
    .Row = 0
    .FixedCols = 0
    .Row = 0
    .Col = count: .Text = GetResourceString(36, 60): count = count + 1 ' "Account No"
    .Col = count: .Text = GetResourceString(35): count = count + 1 '"Name"
    If chkAgent.Value = vbChecked Then
        .Col = count: .Text = GetResourceString(330, 35)
        count = count + 1 '"Agent Name "
    End If
    .Col = count: .Text = GetResourceString(291): count = count + 1 ' "Maturity Date"
    .Col = count: .Text = GetResourceString(186): count = count + 1 '"RateOfInterest"
    .Col = count: .Text = GetResourceString(43, 40): count = count + 1 '"Deposited Amount"
    .Row = 0
    For count = 0 To .Cols - 1
        .Col = count
        .CellAlignment = 4: .CellFontBold = True
    Next
End With


Dim SecondRst As Recordset
Dim Days As Integer
Dim DepAmt As Currency, MatAmt As Currency
Dim Interest As Double
Dim DepTotal As Currency, MatTotal As Currency
Dim DepDate As String, MatDate As String

    
    RaiseEvent Initialise(0, rst.RecordCount)
    RaiseEvent Processing("Aligning  the data ", 0)

Dim rowno As Long
grd.Row = 0

While Not rst.EOF
    gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount From PDTrans " & _
                " Where AccId = " & FormatField(rst("Accid"))
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo nextRecord
    With grd
        'Set next row
        If .Rows <= rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: count = 0
        DepDate = FormatField(rst("CreateDate"))
        MatDate = FormatField(rst("MaturityDate"))
        Days = WisDateDiff(DepDate, MatDate)
        DepAmt = Val(FormatField(SecondRst("TotalAmount")))
        Interest = Val(FormatField(rst("RateOfInterest")))
        MatAmt = FormatCurrency(DepAmt + ComputePDInterest(DepAmt, Interest))
        MatTotal = MatTotal + MatAmt
        DepTotal = DepTotal + DepAmt
        .Col = count: .TextMatrix(rowno, count) = FormatField(rst("AccNUM")): count = count + 1
        .Col = count: .TextMatrix(rowno, count) = FormatField(rst("Name")): count = count + 1
        If chkAgent.Value = vbChecked Then
            .Col = count: .TextMatrix(rowno, count) = FormatField(rst("AgentName")) ''GetAgentName(FormatField(Rst("UserId")))
            count = count + 1
        End If
        .Col = count: .TextMatrix(rowno, count) = MatDate: count = count + 1
        .Col = count: .TextMatrix(rowno, count) = Interest: count = count + 1
        .Col = count: .TextMatrix(rowno, count) = FormatCurrency(DepAmt): count = count + 1
    End With
    
nextRecord:
       
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext

Wend

'Set last
With grd
    .Row = rowno
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    
    .Col = 1: .Text = GetResourceString(52) '"Totals"
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(DepTotal)
    .CellAlignment = 7: .CellFontBold = True
End With
    
lblReportTitle.Caption = GetResourceString(72) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    
   
End Sub

'
Private Sub ShowDepositBalances()
Dim I As Integer
Dim rst As Recordset
Dim SqlStmt As String
Dim StrAmount As String

'
RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStmt = "Select Max(TransId) AS MaxTransID, A.AccID" & _
    " From PDTrans B Inner Join PDMaster A On B.AccId = A.AccId " & _
    " Where TransDate <= #" & m_ToDate & "#" & _
    " GROUP BY A.AccID"
     
gDbTrans.SqlStmt = SqlStmt
If Not gDbTrans.CreateView("QryTemp") Then Exit Sub

SqlStmt = "Select  B.Balance, A.AgentId, A.AccID,A.AccNum,A.CustomerId, Name " & _
    " From QryName C Inner join (PDMaster A inner join " & _
    " (PDtrans B Inner join QryTemp D ON B.TransId = D.MaxTransID AND D.AccID = B.AccID )" & _
        " On A.AccID = B.AccId )" & _
    " ON C.CustomerId = A.CustomerId "

StrAmount = ""
If m_FromAmt > 0 Then StrAmount = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then StrAmount = StrAmount & " And Balance < " & m_ToAmt
If Len(m_Place) Then StrAmount = StrAmount & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then StrAmount = StrAmount & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then StrAmount = StrAmount & " And C.Gender= " & m_Gender
If m_AccGroup Then StrAmount = StrAmount & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then StrAmount = StrAmount & " And A.AgentID = " & m_AgentID

If Len(StrAmount) Then
    StrAmount = " WHERE " & Mid(Trim$(StrAmount), 4)
End If

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & StrAmount & " Order by A.AgentID,val(A.AccNum)"
Else
    gDbTrans.SqlStmt = SqlStmt & StrAmount & " Order by A.AgentID,C.IsciName"
End If

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
RaiseEvent Initialise(0, rst.RecordCount + 1)
RaiseEvent Processing("Aligning the data ", 0)

Dim TotalAmount As Currency
Dim AgentID As Long
Dim AgentName As String
Dim Total As Currency

Call InitGrid
Dim SlNo As Integer

grd.Row = 0
rst.MoveFirst
AgentID = Val(FormatField(rst("AgentID")))
AgentName = GetAgentName(AgentID)
I = 0
If m_AgentNameShow Then I = 1
Dim rowno As Long

While Not rst.EOF
    With grd
        If AgentID <> Val(FormatField(rst("AgentID"))) And TotalAmount > 0 Then
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1
            .Row = rowno: .Col = .Cols - 1
            .Text = FormatCurrency(TotalAmount): .CellAlignment = 7: .CellFontBold = True
            Total = Total + TotalAmount
            .Col = .Cols - 2
            .Text = GetResourceString(52, 42)
            .CellFontBold = True
            AgentID = Val(FormatField(rst("AgentID"))): TotalAmount = 0
            AgentName = GetAgentName(AgentID)
        End If
        
        If FormatField(rst("Balance")) = 0 Then GoTo nextRecord
        If .Rows = rowno + 2 Then .Rows = .Rows + 2
        rowno = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        If chkAgent.Value = vbChecked Then
          .TextMatrix(rowno, 3) = AgentName
        End If
        .TextMatrix(rowno, 3 + I) = FormatField(rst("Balance"))
    End With
    TotalAmount = TotalAmount + Val(FormatField(rst("Balance")))

nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext

Wend
    
With grd
    rowno = rowno + 2
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    .Col = 2: .Text = GetResourceString(52, 42)   ' "Totals Balance "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(TotalAmount)
    .CellAlignment = 7: .CellFontBold = True
    .Row = .Row: .Text = FormatCurrency(TotalAmount + Total)
    .CellAlignment = 7: .CellFontBold = True
End With

lblReportTitle.Caption = GetResourceString(70)

End Sub

Private Sub ShowDepositsOpened()

Dim Dt As Boolean
Dim Amt As Boolean
Dim AmtStr As String
Dim rst As Recordset
Dim Total As Currency
Dim SqlStr As String
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

Dim strClause As String
If Len(m_Place) Then strClause = strClause & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And B.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup
If m_AgentID Then strClause = strClause & " And A.AgentID= " & m_AgentID

'Build the FINAL SQL
SqlStr = " Select AccNum,AccId,AgentID,CreateDate,MaturityDate, " & _
    " Closeddate,RateOfInterest, Name" & _
    " From PDMaster A Inner join QryName B " & _
    " ON B.CustomerID = A.CustomerID " & _
    " where CreateDate <= #" & m_ToDate & "#" & _
    " And CreateDate >= #" & m_FromDate & "# " & strClause & _
    " order by val(AccNum)"
 
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
        
Dim count As Integer
RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Alignign the data ", 0)

Call InitGrid
Dim AccId As Long
Dim Amount As Currency
Dim SecondRst As Recordset
Dim rowno As Long
rowno = grd.Row
While Not rst.EOF
    With grd
        gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount " & _
            " From PDTrans Where AccId = " & FormatField(rst("AccId"))
        If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo nextRecord
        'Set next row
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        count = 0
        AccId = FormatField(rst("AccID"))
        .TextMatrix(rowno, count) = FormatField(rst("AccNum")): count = count + 1
        .TextMatrix(rowno, count) = FormatField(rst("Name")): count = count + 1
        If chkAgent.Value = vbChecked Then
            .TextMatrix(rowno, count) = FormatField(rst("AgentName")): count = count + 1
        End If
        .TextMatrix(rowno, count) = FormatField(rst("CreateDate")): count = count + 1
        .TextMatrix(rowno, count) = FormatField(SecondRst("TotalAmount")): count = count + 1
        Total = Total + Val(FormatField(SecondRst("TotalAmount")))
    End With
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Set next row
With grd
    rowno = rowno + 2
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    
    .Col = 0: .Text = GetResourceString(52, 43)   '"Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: grd.Text = FormatCurrency(Total): .CellAlignment = 4: .CellFontBold = True
End With

End Sub

Private Sub ShowDepositsClosed()

Dim SqlStr As String
Dim rst As Recordset
Dim Total As Currency


RaiseEvent Processing("Reading & Verifying the records", 0)

SqlStr = "Select A.AccId,A.Amount as PrinAmount,B.Amount as IntAmount," & _
        " A.TransID,B.TransType " & _
        " From PDTrans A Left Join PDIntTrans B" & _
        " ON A.AccID = B.ACCID and A.TransID = B.TransID" & _
        " Where A.TransDate >= #" & m_FromDate & "#" & _
        " AND A.TransDate <= #" & m_ToDate & "#"
gDbTrans.SqlStmt = SqlStr
Call gDbTrans.CreateView("QryPDClose1")

SqlStr = "Select A.*,B.Amount as PayableAmount,  " & _
        " PrinAmount + IntAmount + Amount as MatAmount" & _
        " From qryPDClose1 A Left Join PDIntPayable B" & _
        " ON A.AccID = B.ACCID and A.TransID = B.TransID " & _
        " Where A.TransID = (Select Max(TransID) From PDTrans C " & _
            " Where C.AccId = A.AccID )"
        
gDbTrans.SqlStmt = SqlStr
Call gDbTrans.CreateView("QryPDClose")

'Build the SQL
SqlStr = "Select AccNum,AgentID,MaturityDate, " & _
        " PigmyAmount,ClosedDate, c.*, B.Name " & _
        " From QryName B Inner join (PDMaster A " & _
        " Inner join qryPdClose C ON C.AccId = A.AccID) " & _
            " On B.CustomerId = A.CustomerID " & _
        " WHERE ClosedDate >= #" & m_FromDate & "#" & _
        " AND ClosedDate <= #" & m_ToDate & "#"

If Len(m_Place) Then SqlStr = SqlStr & " And B.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStr = SqlStr & " And B.Caste = " & AddQuotes(m_Caste)
If m_Gender Then SqlStr = SqlStr & " And B.Gender= " & m_Gender
If m_AccGroup Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroup
If m_AgentID Then SqlStr = SqlStr & " And A.AgentID = " & m_AgentID

If m_ReportOrder = wisByName Then
    SqlStr = SqlStr & " ORDER BY ClosedDate,A.AgentID,B.IsciName"
Else
    SqlStr = SqlStr & " ORDER BY ClosedDate,A.AgentID,val(A.AccNum)"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

'Initialize the Grid
Call InitGrid

Dim I As Integer
Dim SlNo As Integer

Dim AccId As Long
Dim AgentID As Integer
Dim AgentName As String
Dim TransID As Integer
Dim Amount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim rstTemp As Recordset
Dim rowno As Long

I = IIf(chkAgent.Value = vbChecked, 1, 0)
rowno = grd.Row
While Not rst.EOF
    'Get Returned Amount
    If AgentID <> rst("AgentID") Then AgentName = GetAgentName(rst("AgentID"))
    AgentID = rst("AgentID")
    
    SlNo = SlNo + 1
    gDbTrans.SqlStmt = "select * From qryPDClose " & _
                        " Where AccID = " & rst("AccId")
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then GoTo nextRecord
    
    Amount = 0: IntAmount = 0
    PayableAmount = 0: TransID = 0
    
    Amount = FormatField(rst("PrinAmount"))
    'Get the the INterest Paid Amount
    IntAmount = FormatField(rst("IntAmount"))
    If rst("TransType") = wWithdraw Or rst("TransType") = wContraWithdraw _
                                Then IntAmount = IntAmount * -1
    PayableAmount = FormatField(rst("PayableAmount"))
    If IntAmount < 0 Then IntAmount = IntAmount + PayableAmount: PayableAmount = 0
    
    'Checkthe condition of the minimum amount
    If (Amount + IntAmount + PayableAmount) < m_FromAmt Then GoTo nextRecord
    'Check the condition of the maximum amount is given
    If m_ToAmt > 0 And (Amount + IntAmount + PayableAmount) > m_ToAmt Then GoTo nextRecord
    
    With grd
        'Set next row
        SlNo = SlNo + 1
        If .Rows < rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1:
        AccId = FormatField(rst("AccId"))
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("Name"))
        If chkAgent.Value = vbChecked Then .TextMatrix(rowno, 3) = AgentName
        
        .TextMatrix(rowno, 3 + I) = FormatField(rst("ClosedDate"))
        .TextMatrix(rowno, 4 + I) = FormatCurrency(Amount + IntAmount + PayableAmount)
        Total = Total + Val(.TextMatrix(rowno, 4 + I))
    End With
    
nextRecord:
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Set next row
With grd
    rowno = rowno + 2
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno + 1
    .Col = 1: .Text = GetResourceString(52) & " " & _
                    GetResourceString(43) ' "Total Deposits "
    .CellAlignment = 4: .CellFontBold = True
    .Col = .Cols - 1: .Text = FormatCurrency(Total)
    .CellAlignment = 7: .CellFontBold = True
End With

    lblReportTitle.Caption = GetResourceString(78) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        
End Sub

Private Sub ShowLiabilities()
Dim SqlStmt As String
Dim SecondRst As Recordset
Dim rst As Recordset
Dim count As Integer
'
RaiseEvent Processing("Reading & Verifying the records ", 0)

                    
SqlStmt = "Select A.AgentID,A.AccID,MaturityDate,CreateDate,RateOfInterest," & _
        " CLosedDate, Name " & _
        " From PDMaster A Inner join QryName B ON B.CustomerId = A.CustomerId" & _
        " Where A.AccId not In (Select AccId From PDMaster" & _
        " Where ClosedDate < #" & m_ToDate & "# )"

If chkAgent.Value Then
SqlStmt = "Select A.AgentID,A.AccID,MaturityDate,CreateDate,RateOfInterest," & _
        " ClosedDate, B.Name, D.Name as AgentName " & _
        " From QryName B Inner join (PDMaster A Inner join " & _
            " (UserTab C Inner join QryName D ON C.CustomerID = D.CustomerID)" & _
        " ON A.AgentID = C.UserID) ON B.CustomerId = A.CustomerId" & _
        " Where A.AccId not In (Select AccId From PDMaster" & _
            " Where ClosedDate < #" & m_ToDate & "#" & ")  "

End If

If m_FromAmt > 0 Then SqlStmt = " And Balance > " & m_FromAmt
If m_ToAmt > 0 Then SqlStmt = SqlStmt & " And Balance < " & m_ToAmt
If Len(m_Place) Then SqlStmt = SqlStmt & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then SqlStmt = SqlStmt & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then SqlStmt = SqlStmt & " And C.Gender= " & m_Gender
If m_AccGroup Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then SqlStmt = SqlStmt & " And A.AgentID= " & m_AgentID


If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStmt & " Order by A.UserId,val(A.AccNum)"
Else
    gDbTrans.SqlStmt = SqlStmt & " Order by IsciName "
End If
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    count = 100
    count = 0
    Exit Sub
End If
    
'Init the grid
Call InitGrid

Dim Days As Integer
Dim Liability As Currency
Dim GrandTotal As Currency

Dim CustName As String
grd.Row = 0
Dim rowno As Long

RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)
With grd
  While Not rst.EOF
    gDbTrans.SqlStmt = "Select sum(Amount) as TotalAmount " & _
        " from PDTrans where TransDate <=" & _
        " #" & m_FromDate & "# and AccId = " & rst("AccId")
    If gDbTrans.Fetch(SecondRst, adOpenForwardOnly) < 1 Then GoTo nextRecord
    'Set next row
    rowno = rowno + 1
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    count = 0
    .TextMatrix(rowno, count) = FormatField(rst("AccID")): count = count + 1
    .TextMatrix(rowno, count) = FormatField(rst("Name")): count = count + 1
    If chkAgent.Value = vbChecked Then
        .TextMatrix(rowno, count) = FormatField(rst("AgentName")) ''GetAgentName(Val(Rst("UserId")))
        count = count + 1
    End If
    .TextMatrix(rowno, count) = FormatField(rst("CreateDate")): count = count + 1
    .TextMatrix(rowno, count) = FormatField(rst("MaturityDate")): count = count + 1
    .TextMatrix(rowno, count) = FormatField(rst("RateOfInterest")): count = count + 1
    .TextMatrix(rowno, count) = FormatField(SecondRst("TotalAmount")): count = count + 1
    Liability = Val(FormatField(SecondRst("TotalAmount"))) + _
            ComputePDInterest(Val(FormatField(SecondRst("TotalAmount"))), Val(FormatField(rst("RateOfInterest"))))
    .TextMatrix(rowno, count) = FormatCurrency(Liability): count = count + 1
    GrandTotal = GrandTotal + Liability
nextRecord:
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
  Wend
'Fill In total Liability
    rowno = rowno + 2
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    
    .Row = rowno
    .Col = 1: .Text = GetResourceString(52) + " " + GetResourceString(405) '"TOTAL LIABILITIES":
    .CellAlignment = 4: .CellFontBold = True
    
    .Col = IIf(chkAgent.Value = vbChecked, 7, 6)
    .Text = FormatCurrency(GrandTotal)
    .CellAlignment = 7: .CellFontBold = True
End With

End Sub


Private Sub ShowDayBook()

Dim SqlStr As String
Dim rst As Recordset
Dim I As Integer
Dim TransDep As wisTransactionTypes
Dim transType As wisTransactionTypes
Dim count As Integer

'Report title
lblReportTitle.Caption = GetResourceString(71)

'To Get Deposits & Payments of of PD Account
TransDep = wDeposit
transType = wWithdraw

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Amount > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Amount < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And C.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup
If m_AgentID > 0 Then strClause = strClause & " And A.AgentID = " & m_AgentID

RaiseEvent Processing("Reading & Verifyig the records ", 0)
On Error GoTo ErrLine

SqlStr = "Select C.Name, A.AccId,A.AgentId,A.AccNum, Val(A.AccNum) as AcNum, TransID,TransDate, " & _
    " Amount, TransType,IsciName from  QryName C Inner join " & _
    " (PDMaster A Inner Join PDTrans B On B.AccId = A.AccId) " & _
    " ON C.CustomerId = A.CustomerID " & _
    " where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#" & _
 strClause

SqlStr = SqlStr & " UNION " & "Select  'INTEREST', A.AccId,A.AgentId,A.AccNum, Val(A.AccNum) as AcNum," & _
    " TransID,TransDate, Amount, TransType,IsciName from QryName C" & _
    " Inner join (PDMaster A Inner join PDIntTrans B On B.AccId = A.AccId)" & _
    " ON C.CustomerId = A.CustomerID " & _
    " Where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#"

SqlStr = SqlStr & strClause

If m_ReportOrder = wisByAccountNo Then
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,AcNum,TransID"
Else
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,C.IsciName,TransID" 'Isciname was not in the above query(Included)
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
    
RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
grd.Row = grd.FixedRows
Dim CustName As String
Dim TransDate As Date
Dim AgentName As String
Dim AgentDepositTotal As Currency
Dim AgentID As Long
Dim AccId As Long

Dim SubTotal() As Currency
Dim GrandTotal() As Currency

ReDim SubTotal(4 To grd.Cols - 1)
ReDim GrandTotal(4 To grd.Cols - 1)

AgentDepositTotal = 0
AgentID = Val(FormatField(rst("AgentId")))

Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim blInt As Boolean

TransDate = rst("TransDate")

'Put All The Agents name in a StringArray
grd.Row = grd.FixedRows
I = 0
If chkAgent.Value = vbChecked Then I = 1
AgentName = GetAgentName(AgentID)
Dim rowno As Long, colno As Byte
rowno = 1
While Not rst.EOF
    'See if you have to calculate sub totals
    With grd
        If chkAgent.Value = vbChecked Then
        If AgentID <> Val(FormatField(rst("AgentId"))) Then
            If .Rows <= rowno + 2 Then .Rows = .Rows + 4
            rowno = rowno + 1
            .Row = rowno
            .Col = 5: .Text = FormatCurrency(AgentDepositTotal)
            .CellAlignment = 7: .CellFontBold = True
            AgentID = Val(FormatField(rst("AgentId")))
            AgentName = GetAgentName(AgentID)
            If .Rows <= rowno + 2 Then .Rows = .Rows + 4
            rowno = rowno + 1
            AgentDepositTotal = 0
        End If
        End If
        If TransDate <> rst("TransDate") Then
            PRINTTotal = True
            'Set next row
            AccId = 0: SlNo = 0
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1: count = 0
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) & _
                " " & GetIndianDate(TransDate)
            .CellAlignment = 4: .CellFontBold = True
            For count = IIf(I, 5, 4) To .Cols - 1
                .Col = count
                .Text = FormatCurrency(SubTotal(count))
                .CellAlignment = 7: .CellFontBold = True
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
            If .Rows <= rowno + 2 Then .Rows = .Rows + 2
            rowno = rowno + 1: count = 0
            .Row = rowno
            TransDate = rst("transDate")
        End If
        'Set next row
        If .Rows <= rowno + 2 Then .Rows = rowno + 2
        rowno = rowno + 1: count = 0
        'if He has paid the interest amount then
        'need not to write into the grid 'so moveback one row
        If AccId = rst("AccId") Then rowno = rowno - 1 Else blInt = False
        
        If chkAgent.Value = vbChecked Then .TextMatrix(rowno, 4) = AgentName
        
        transType = FormatField(rst("TransType"))
        If FormatField(rst(0)) = "INTEREST" Then
            blInt = True
            If transType = wWithdraw Then colno = 8 + I
            If transType = wContraWithdraw Then colno = 9 + I
        Else
            If AccId = rst("AccId") And Not blInt Then rowno = rowno + 1
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = Format(SlNo, "00")
            .TextMatrix(rowno, 3) = FormatField(rst(0))
            If transType = wDeposit Then colno = 4 + I
            If transType = wContraDeposit Then colno = 5 + I
            If transType = wWithdraw Then colno = 6 + I
            If transType = wContraWithdraw Then colno = 7 + I
        End If
        
        .TextMatrix(rowno, colno) = FormatField(rst("Amount"))
        '.Row = rowno: .Col = colno
        SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
        .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum"))
        
    End With
        
        AccId = rst("AccId")
        If gCancel Then rst.MoveLast
        rst.MoveNext
        DoEvents
       
        Me.Refresh
        RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
Wend

'Now Print the Subtotal Of the Last day
'Set next row
With grd
    AccId = 0
    If .Rows <= rowno + 2 Then .Rows = .Rows + 2
    rowno = rowno + 1: count = 0
    .Row = rowno
    .Col = 2: .Text = GetResourceString(304) & _
        " " & GetIndianDate(TransDate): count = count + 1
    .CellAlignment = 4: .CellFontBold = True
    If chkAgent.Value = vbChecked Then count = count + 1
    For count = IIf(chkAgent.Value = vbChecked, 5, 4) To .Cols - 1
        .Col = count
        .Text = FormatCurrency(SubTotal(count))
        .CellAlignment = 7: .CellFontBold = True
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
    Next
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1: count = 0
End With
      
'Now Print the grand total Of the Last day
If PRINTTotal = True Then
    With grd
        'Set next row
        If .Rows <= rowno + 3 Then .Rows = rowno + 3
        rowno = rowno + 2: count = 0
        .Row = rowno
        .Col = 2: .Text = GetResourceString(286): count = count + 1
        .CellAlignment = 4: .CellFontBold = True
        If chkAgent.Value = vbChecked Then count = count + 1
        For count = IIf(chkAgent.Value = vbChecked, 5, 6) To .Cols - 1
            .Col = count
            .Text = FormatCurrency(GrandTotal(count))
            .CellAlignment = 7: .CellFontBold = True
        Next
    End With

End If
  
lblReportTitle.Caption = GetResourceString(390) & " " & _
        GetResourceString(63) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
  
Exit Sub
ErrLine:
'Resume
Exit Sub

End Sub

Private Sub ShowSubCashBook()

Dim SqlStr As String
Dim rst As Recordset
Dim I As Integer
Dim TransDep As wisTransactionTypes
Dim transType As wisTransactionTypes
Dim count As Integer

'Report title
lblReportTitle.Caption = GetResourceString(71)

'To Get Deposits & Payments of of PD Account
TransDep = wDeposit
transType = wWithdraw

Dim strClause As String
If m_FromAmt > 0 Then strClause = " And Amount > " & m_FromAmt
If m_ToAmt > 0 Then strClause = strClause & " And Amount < " & m_ToAmt
If Len(m_Place) Then strClause = strClause & " And C.Place = " & AddQuotes(m_Place)
If Len(m_Caste) Then strClause = strClause & " And C.Caste = " & AddQuotes(m_Caste)
If m_Gender Then strClause = strClause & " And C.Gender= " & m_Gender
If m_AccGroup Then strClause = strClause & " And AccGroupID = " & m_AccGroup

RaiseEvent Processing("Reading & Verifyig the records ", 0)
On Error GoTo ErrLine

SqlStr = "Select Name, A.AccId,A.AgentId,A.AccNum, TransID,TransDate, VoucherNo," & _
    " Amount, TransType,IsciName from QryName C" & _
    " Inner join (PDMaster A Inner join PDTrans B ON B.AccId = A.AccId)" & _
    " ON C.CustomerId = A.CustomerID  " & _
    " where TransDate >= #" & m_FromDate & "# " & _
    " And TransDate <= #" & m_ToDate & "#"
SqlStr = SqlStr & strClause

If m_ReportOrder = wisByAccountNo Then
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,Val(A.AccNum),TransID"
Else
    SqlStr = SqlStr & " Order By TransDate,A.AgentId,C.IsciName,TransID" 'Isciname was not in the above query(Included)
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
SqlStr = ""

RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
grd.Row = grd.FixedRows
Dim CustName As String
Dim TransDate As Date
Dim AgentName As String
Dim AgentDepositTotal As Currency
Dim AgentID As Long
Dim AccId As Long

Dim SubTotal() As Currency
Dim GrandTotal() As Currency

ReDim SubTotal(5 To grd.Cols - 1)
ReDim GrandTotal(5 To grd.Cols - 1)

AgentDepositTotal = 0

Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim blInt As Boolean
Dim rowno As Long, colno As Byte

TransDate = rst("TransDate")

'Put All The Agents name in a StringArray
grd.Row = grd.FixedRows - 1
rowno = grd.Row
I = 0
AgentName = GetAgentName(AgentID)

While Not rst.EOF
    'See if you have to calculate sub totals
    With grd
        If AgentID <> Val(FormatField(rst("AgentId"))) Then
            If .Rows <= rowno + 2 Then .Rows = rowno + 4
            rowno = rowno + 1
            .Row = rowno
            .Col = 5: .Text = FormatCurrency(AgentDepositTotal)
            .CellAlignment = 7: .CellFontBold = True
            AgentDepositTotal = 0
        End If
        If TransDate <> rst("TransDate") Then
            PRINTTotal = True: AgentID = 0
            'Set next row
            AccId = 0: SlNo = 0
            If .Rows < rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1: count = 0
            .Row = rowno
            .Col = 3: .Text = GetResourceString(304) & _
                            " " & GetIndianDate(TransDate)
            .CellAlignment = 4: .CellFontBold = True
            For count = IIf(I, 6, 5) To .Cols - 1
                .Col = count
                .Text = FormatCurrency(SubTotal(count))
                .CellAlignment = 7: .CellFontBold = True
                GrandTotal(count) = GrandTotal(count) + SubTotal(count)
                SubTotal(count) = 0
            Next
            If .Rows <= rowno + 2 Then .Rows = rowno + 2
            rowno = rowno + 1: count = 0
            TransDate = rst("transDate")
        End If
        If AgentID <> Val(FormatField(rst("AgentId"))) Then
            AgentID = Val(FormatField(rst("AgentId")))
            AgentName = GetAgentName(AgentID)
            If .Rows <= rowno + 2 Then .Rows = rowno + 4
            rowno = rowno + 1
            .Row = rowno
            .Col = 3: .Text = AgentName: .CellFontBold = True
        End If
        
        'Set next row
        rowno = rowno + 1: count = 0
        transType = FormatField(rst("TransType"))
        If AccId = rst("AccId") And Not blInt Then rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = Format(SlNo, "00")
        .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
        .TextMatrix(rowno, 2) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 3) = FormatField(rst("Name"))
        .TextMatrix(rowno, 4) = FormatField(rst("VoucherNo"))
        
        colno = 6
        If transType = wDeposit Or transType = wContraDeposit Then colno = 5
        .TextMatrix(rowno, colno) = FormatField(rst("Amount"))
        
        SubTotal(colno) = SubTotal(colno) + Val(.TextMatrix(rowno, colno))
        
    End With
        
    AccId = rst("AccId")
    rst.MoveNext
    DoEvents
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
Wend

'Now Print the Subtotal Of the Last day
'Set next row
With grd
    AccId = 0
    rowno = rowno + 1
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno: count = 0
    .Col = 3: .Text = GetResourceString(304) & _
        " " & GetIndianDate(TransDate): count = count + 1
    .CellAlignment = 4: .CellFontBold = True
    If chkAgent.Value = vbChecked Then count = count + 1
    For count = IIf(I, 6, 5) To .Cols - 1
        .Col = count
        .Text = FormatCurrency(SubTotal(count))
        .CellAlignment = 7: .CellFontBold = True
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
    Next
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1: count = 0
    rowno = .Row
End With
      
'Now Print the grand total Of the Last day
If PRINTTotal = True Then
    With grd
        'Set next row
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1: count = 0
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1: count = 0
        .Col = 3: .Text = GetResourceString(286): count = count + 1
        .CellAlignment = 4: .CellFontBold = True
        If chkAgent.Value = vbChecked Then count = count + 1
        For count = IIf(I, 6, 5) To .Cols - 1
            .Col = count
            .Text = FormatCurrency(GrandTotal(count))
            .CellAlignment = 7: .CellFontBold = True
        Next
    End With

End If
  
lblReportTitle.Caption = GetResourceString(390) & " " & _
        GetResourceString(85) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
  
Exit Sub
ErrLine:
'Resume
Exit Sub

End Sub

'
Private Sub ShowMonthlyTransaction(Optional Loan As Boolean)
Dim rst As Recordset
Dim transType As wisTransactionTypes
Dim count As Integer

'To Get Deposits & Payments of of PD Account
transType = wWithdraw

'Now Set the Date as on date
'
RaiseEvent Processing("Reading & Verifyig the records ", 0)
Err.Clear
grd.Clear
grd.Cols = 3
grd.Row = 0
grd.Col = 0: grd.Text = "Agent ID": grd.ColWidth(0) = 0.1
grd.Col = 1: grd.Text = GetResourceString(36, 60)
grd.Col = 2: grd.Text = GetResourceString(35)

'First Insert the AgentId, AccId  to the grid
gDbTrans.SqlStmt = "Select AgentID,AccID, AccNum, Name  as CustName " & _
    " From PDMaster A Inner join QryName B ON A.CustomerID = B.CustomerID " & _
    " Where (ClosedDate >= #" & m_ToDate & "# OR ClosedDate is NULL )" & _
    " Order By AgentId, val(AccNum)"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
grd.Row = 0

' Intially set the No col
Dim FirstDate As Date
Dim LastDate As Date
Dim colno As Integer
Dim AgentID As Long
Dim AccNum As String

FirstDate = GetSysFirstDate(m_FromDate)
grd.Cols = grd.Cols + DateDiff("m", m_FromDate, m_ToDate) * 2

'Now onwardss alll dates are in "mm/dd/yyyyyy" format
LastDate = DateAdd("m", 1, FirstDate)
Dim rowno As Long
rowno = grd.Row
While Not rst.EOF
    With grd
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .TextMatrix(rowno, 0) = FormatField(rst("AgentId"))
        .TextMatrix(rowno, 1) = FormatField(rst("AccNum"))
        .TextMatrix(rowno, 2) = FormatField(rst("CustName"))
    End With
    rst.MoveNext
Wend


' Now start to fill the grid
colno = 3
Do
        'Condition
        'Get One month transaction details of all accounts
        If DateDiff("M", FirstDate, m_ToDate) < 0 Then Exit Do
        Set rst = Nothing
        gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount,AgentId," & _
            " AccNum,TransType From PdTrans A Inner join PDMaster B ON " & _
            " A.AccID = B.accId WHERE Transdate >= #" & FirstDate & "# " & _
            " And TransDate < #" & LastDate & "# " & _
            " GROUP BY AgentId,AccNum,TransType" & _
            " ORDER BY AgentID,val(AccNum) "
    
        If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo NextMonth
        ' now Insert the Transaction
    With grd
        .Row = 0: rowno = 0
        If .Cols <= colno Then .Cols = colno + 2
        .Col = colno
        .Text = GetMonthString(Month(FirstDate)) & " " & GetResourceString(271)
        .Col = colno + 1
        .Text = GetMonthString(Month(FirstDate)) & " " & GetResourceString(272)
    End With
    
    While Not rst.EOF
        With grd
            transType = FormatField(rst("TransType"))
            'Now Get the Propre grid row to fit the Values
            Do
                AgentID = Val(.TextMatrix(rowno, 0))
                AccNum = .TextMatrix(rowno, 1)
                If AgentID = rst("AgentId") And AccNum = rst("AccNum") Then Exit Do
                If rowno = .Rows - 1 Then
                    rowno = 1
                    GoTo nextRecord
                End If
                If rowno = .Rows - 1 Then Exit Do
                rowno = rowno + 1
            Loop
            If transType = wDeposit Or transType = wContraDeposit Then
                .TextMatrix(rowno, colno) = FormatField(rst("TotalAmount"))
            Else 'If TransType = wWithDraw Then
                .TextMatrix(rowno, colno + 1) = FormatField(rst("TotalAmount"))
            End If
        End With
nextRecord:
        rst.MoveNext
    Wend
    
NextMonth:
    colno = colno + 2
    FirstDate = LastDate
    LastDate = DateAdd("m", 1, CDate(FirstDate))
Loop

End Sub

Private Sub ShowDailyTransaction(Optional Loan As Boolean)
Dim SqlStmt As String
Dim rst As Recordset
Dim transType As wisTransactionTypes
Dim fromDate As Date
Dim toDate As Date
Dim AccId As Long
Dim AgentID As Integer
Dim AgentName As String
Dim firstCol As Integer
Dim colno As Integer
Dim accountWithdraw As Currency
Dim accountDeposit As Currency
Dim accountBalance As Currency
Dim custRst As Recordset
Dim tmpCustRst As Recordset

RaiseEvent Processing("Reading && Verifyig the records ", 0)
Err.Clear

Call InitGrid
''Call Customer Details

gDbTrans.SqlStmt = "Select AgentID,AccID,AccNum, Name as CustName" & _
    " From PdMaster B inner join  QryName C ON B.CustomerID = C.CustomerID" & _
    " Order By Val(AccNum) "
If gDbTrans.Fetch(custRst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(278), vbOKOnly, wis_MESSAGE_TITLE
    Exit Sub
End If

'Monthly Max Trans ID
gDbTrans.SqlStmt = "Select AccId, max(TransID) as MonTransId, Month(TransDate) as Month1,year(TransDate) as Year1" & _
    " from PDTrans group by AccId, Month(TransDate),Year(transDate)"
gDbTrans.CreateView ("MonthTransIDs")

gDbTrans.SqlStmt = "Select C.AgentID,A.AccID,C.AccNum,Name as CustName, B.Balance,B.TransDate,Month1,Year1" & _
    " From MonthTransIDs A Inner join (PDTrans B  Inner Join " & _
    " (PdMaster C inner join  QryName D ON C.CustomerID = D.CustomerID) " & _
    " ON B.AccID = C.AccID )On A.AccID = B.AccID and A.MonTransID = B.TransID"
If m_AgentID <> 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " where C.AgentId = " & m_AgentID

If gDbTrans.Fetch(custRst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(278), vbOKOnly, wis_MESSAGE_TITLE
    Exit Sub
End If

fromDate = GetSysFirstDate(m_FromDate)
toDate = DateAdd("M", 1, fromDate)

grd.Row = 0
' Intially set the No col
Dim AccNum As String
Dim MonthNo As Integer
Dim DayNo As Integer
Dim SlNo As Integer
Dim firstRow As Integer
Dim rowno As Long
Dim colLoop As Integer
Dim rowLoop As Integer
Dim curMonth As Integer
Dim curYear As Integer
Dim accountRowNo As Long
Dim noTrans As Boolean

grd.Row = grd.FixedRows - 1
grd.Row = 1
rowno = 1
noTrans = False
While DateDiff("m", fromDate, m_ToDate) >= 0
    
    If gCancel Then GoTo NextMonth
    
    SqlStmt = "Select Distinct(F.AccID) From PDTrans F" & _
        " Where F.TransDate >= #" & fromDate & "# and F.TransDate < #" & toDate & "# "
    
    SqlStmt = "Select D.AgentId, D.AccID,D.AccNum, val(D.AccNum) as AcNum, #" & _
        DateAdd("D", -1, fromDate) & "# as TransDate, 0 as TransID,0 as Amount,0 as TransType,0 as Balance " & _
        " From PDMaster D Where  D.AccID not in (" & SqlStmt & ")"
    If m_AgentID <> 0 Then SqlStmt = SqlStmt & " And D.AgentId = " & m_AgentID
    
    gDbTrans.SqlStmt = SqlStmt & " UNION " & _
        "Select AgentId, A.AccID,B.AccNum, val(B.AccNum) as AcNum,TransDate,TransID,Amount,TransType,Balance " & _
        " From PDTrans A Inner Join PDMaster B On A.AccID = B.AccID " & _
        " Where TransDate >= #" & fromDate & "# and TransDate < #" & toDate & "# "
    If m_AgentID <> 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AgentId = " & m_AgentID
    
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order by AgentId, AcNum"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo NextMonth
    gDbTrans.SqlStmt = SqlStmt
    
    
    grd.Row = rowno
    If DateDiff("m", fromDate, m_FromDate) <> 0 Then ShowDailyTransaction_PrintMonthHeader (fromDate)
    rowno = grd.Row
    AccId = 0
    firstCol = 3
    SlNo = 0
    firstRow = rowno + 1
    
    While Not rst.EOF
        If AgentID <> rst("AgentID") And m_AgentNameShow Then
            If AgentID <> 0 Then
                If grd.Rows < rowno + 2 Then grd.Rows = rowno + 2
                rowno = rowno + 1
                'Print Total of the Each Day for the agent
                accountBalance = ShowDailyTransaction_MonthTotal(firstCol, firstRow, rowno)
                rowno = rowno + 1
            End If
            AgentID = FormatField(rst("AgentID"))
            accountBalance = 0: accountDeposit = 0: accountWithdraw = 0
            
            'Print the Agent Details
            If grd.Rows < rowno + 2 Then grd.Rows = rowno + 2
            rowno = rowno + 1
            grd.Row = rowno
            
            'Agent name label
            For SlNo = firstCol - 1 To firstCol + 1
                grd.Col = SlNo
                grd.TextMatrix(rowno, SlNo) = GetResourceString(330, 35)
                grd.CellAlignment = vbCenter
            Next
            'Agent name
            AgentName = GetAgentName(CLng(rst("AgentID")))
            For SlNo = firstCol + 1 To 31 + firstCol
                grd.TextMatrix(rowno, SlNo) = AgentName
                grd.Col = SlNo
                grd.CellAlignment = vbCenter
            Next
            grd.MergeRow(rowno) = True
           
            firstRow = rowno + 1
            SlNo = 0
        End If
        If AccId <> FormatField(rst("Accid")) Then
            'find the CustDetails
            If AccId <> 0 Then
                custRst.Filter = "AccId = " & AccId & " and Month1 = " & Month(rst("TransDate")) & " and Year1 = " & Year(rst("TransDate"))
                If noTrans Then
                    gDbTrans.SqlStmt = "Select A.AccID, AccNum , Balance,TransDate,Name as CustName " & _
                            " from PDTrans A inner join (PDMaster B inner join  QryName C ON C.CustomerID = B.CustomerID)" & _
                            " on A.AccID = B.AccID where A.AccID = " & AccId & _
                            " and TransID = (Select max(TransID) from PDTrans where TransDate < #" & fromDate & "# and AccID = " & AccId & ")"
                    If gDbTrans.Fetch(tmpCustRst, adOpenDynamic) < 1 Then
                        gDbTrans.SqlStmt = "Select B.AccID, AccNum ,0 as Balance,#1/1/2014# as TransDate,Name as CustName " & _
                            " from PDMaster B inner join  QryName C ON C.CustomerID = B.CustomerID" & _
                            " where B.AccID = " & AccId
                        Call gDbTrans.Fetch(tmpCustRst, adOpenDynamic)
                    End If
                    
                    Call ShowDailyTransaction_AccountTotal(tmpCustRst, accountRowNo, accountWithdraw, accountDeposit, accountBalance)
                Else
                    custRst.Filter = "AccId = " & AccId & " and Month1 = " & Month(fromDate) & " and Year1 = " & Year(fromDate)
                    Call ShowDailyTransaction_AccountTotal(custRst, accountRowNo, accountWithdraw, accountDeposit, accountBalance)
                End If
                custRst.Filter = adFilterNone
                accountBalance = 0: accountDeposit = 0: accountWithdraw = 0
            End If
            
            rowno = rowno + 1
            If grd.Rows < rowno + 1 Then grd.Rows = rowno + 1
            grd.Row = rowno
            
            SlNo = SlNo + 1
            grd.TextMatrix(rowno, 0) = Format(SlNo, "00")
            
            AccId = FormatField(rst("Accid"))
            curMonth = Month(rst("TransDate"))
            curYear = Year(rst("TransDate"))
            
            accountWithdraw = 0: accountDeposit = 0
            
        End If
        
        DayNo = Day(rst("TransDate"))
        transType = FormatField(rst("TransType"))
        noTrans = IIf(transType = 0, True, False)
        If transType = wDeposit Or transType = wContraDeposit Then
            grd.TextMatrix(rowno, firstCol + DayNo) = FormatField(rst("Amount"))
            accountDeposit = accountDeposit + FormatField(rst("Amount"))
        Else
            accountWithdraw = accountWithdraw + FormatField(rst("Amount"))
        End If
        accountRowNo = rowno
        
        If gCancel Then rst.MoveLast
        rst.MoveNext
        
    Wend
    
    'Print Total of the Each Day
    accountBalance = ShowDailyTransaction_MonthTotal(firstCol, firstRow, rowno + 1)
    
    custRst.Filter = "AccId = " & AccId & " and Month1 = " & curMonth & " and Year1 = " & curYear
    Call ShowDailyTransaction_AccountTotal(custRst, accountRowNo, accountWithdraw, accountDeposit, accountBalance)
    custRst.Filter = adFilterNone
    accountBalance = 0: accountDeposit = 0: accountWithdraw = 0
    AccId = 0
    AgentID = 0
NextMonth:
    rowno = rowno + 1
    If grd.Rows < rowno + 2 Then grd.Rows = rowno + 2
    fromDate = DateAdd("M", 1, fromDate)
    toDate = DateAdd("m", 1, toDate)
        
Wend

End Sub

Private Sub ShowMonthlyTransactionNew(Optional Loan As Boolean)
Dim rst As Recordset
Dim transType As wisTransactionTypes

RaiseEvent Processing("Reading && Verifyig the records ", 0)
Err.Clear

Call InitGrid

'First Insert the AgentId, AccId  to the grid
gDbTrans.SqlStmt = "Select AgentID,AccID, AccNum, Name  as CustName " & _
    " From PDMaster A Inner join QryName B ON A.CustomerID = B.CustomerID " & _
    " WHERE (ClosedDate >= #" & m_ToDate & "# OR ClosedDate is NULL )" & _
    " Order By AgentId, val(AccNum)"

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
grd.Row = 0

' Intially set the No col
Dim AccNum As String
Dim MonthNo As Integer
Dim SlNo As Integer
Dim rowno As Long

grd.Row = grd.FixedRows - 1
rowno = grd.Row
While Not rst.EOF
    rowno = rowno + 1
    If grd.Rows < rowno + 1 Then grd.Rows = rowno + 1
    SlNo = SlNo + 1
    grd.TextMatrix(rowno, 0) = Format(SlNo, "00")
    grd.TextMatrix(rowno, 1) = FormatField(rst("AgentId"))
    grd.TextMatrix(rowno, 2) = FormatField(rst("AccNum"))
    grd.TextMatrix(rowno, 3) = FormatField(rst("CustName"))
    rst.MoveNext
Wend

rowno = rowno + 2
If grd.Rows < rowno + 1 Then grd.Rows = rowno + 1

grd.TextMatrix(rowno, 3) = GetResourceString(52, 42)
    

' Now start to fill the grid
Dim DepAmount As Currency
Dim WithdrawAmount As Currency
Dim totalDepAmount As Currency
Dim totalWithdrawAmount As Currency

rowno = grd.FixedRows
While rowno < grd.Rows
    With grd
    AccNum = grd.TextMatrix(rowno, 2)
    
    gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount,month(Transdate) as MonthNo," & _
            " TransType From PdTrans A Inner Join PDMaster B " & _
            " ON A.AccID = B.accId WHERE Transdate >= #" & m_FromDate & "# " & _
            " And TransDate < #" & m_ToDate & "# " & _
            " AND AccNum = " & AddQuotes(AccNum, True) & _
            " GROUP BY month(Transdate), TransType" & _
            " Order By month(Transdate)"
            
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then GoTo NextAccount
    MonthNo = rst("MonthNo")
    grd.Col = 4
    Do
        On Error Resume Next
        If MonthNo <> rst("MonthNo") Or rst.EOF Then
            Do
                .Row = 0
                If .Text = GetMonthString(MonthNo) Then
                    .Row = rowno
                    .Text = FormatCurrency(DepAmount)
                    .Col = .Col + 1
                    .Text = FormatCurrency(WithdrawAmount)
                    Debug.Assert WithdrawAmount = 0
                    Exit Do
                End If
                If .Col = .Cols - 1 Then Exit Do
                .Col = .Col + 1
            Loop
            DepAmount = 0: WithdrawAmount = 0
            MonthNo = rst("MonthNo")
            If rst.EOF Then Exit Do
        End If
        transType = FormatField(rst("TransType"))
        If transType > 0 Then DepAmount = DepAmount + FormatField(rst("totalAmount"))
        If transType < 0 Then WithdrawAmount = WithdrawAmount + FormatField(rst("totalAmount"))
        
        RaiseEvent Processing("Writing Data", rowno / .Rows)
        rst.MoveNext
        
    Loop

NextAccount:
    
    End With
    
    rowno = rowno + 1
Wend

End Sub


Private Sub cmdOk_Click()
chkAgent.Enabled = True

Unload Me
End Sub



Private Sub cmdPrint_Click()

Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
        
    Set m_frmCancel = New frmCancel
    Load m_frmCancel
    
    With m_frmCancel
       .Visible = True
       '.Show
    End With
    
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .ReportTitle = lblReportTitle
    .GridObject = grd
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
On Error Resume Next
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
gCancel = 0
'set kannada fonts
Call SetKannadaCaption
'Init the grid
With grd
    .Clear
    .Rows = 20
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    .Text = "No Records Available"
    .CellAlignment = 4: .CellFontBold = True
End With

'Show report
    chkAgent.Value = IIf(m_AgentNameShow, vbChecked, vbUnchecked)
    
    If m_ReportType = repPDBalance Then Call ShowDepositBalances
    If m_ReportType = repPDDayBook Then Call ShowDayBook
    If m_ReportType = repPDCashBook Then Call ShowSubCashBook
    If m_ReportType = repPDMat Then Call MaturedDeposits
    If m_ReportType = repPDAccClose Then Call ShowDepositsClosed
    If m_ReportType = repPDAccOpen Then
        lblReportTitle.Caption = GetResourceString(64) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDepositsOpened
    End If
    
    If m_ReportType = repPDLedger Then Call ShowDepositGeneralLedger
    
    If m_ReportType = repPDAgentTrans Then
        lblReportTitle.Caption = GetResourceString(330) & " " & _
                GetResourceString(38) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowAgentTransaction
    End If
    
    If m_ReportType = repPDMonTrans Then   'This Report to take individua recipts & Pay ments of account Holders
        lblReportTitle.Caption = GetResourceString(330, 38) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowMonthlyTransactionNew
    End If
    If m_ReportType = repPDMonBal Then Call ShowMonthlyBalances
    
    If m_ReportType = repPDDailyTrans Then   'This Report to take individual recipts of account Holders
        lblReportTitle.Caption = GetResourceString(410, 38) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        Call ShowDailyTransaction
    End If
    lblReportTitle.FONTSIZE = 14

'Set the Caption here
'Me.lblReportTitle.Caption = GetResourceString(85)
End Sub

Private Sub ShowMonthlyBalances()
gCancel = 0
Dim count As Long
Dim totalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SqlStmt As String

Dim fromDate As Date
Dim toDate As Date

'Get the Last day of the given month
toDate = GetSysLastDate(m_ToDate)
'Get the Last day of first month to get the balance of that month
fromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463) & " " & _
        GetResourceString(67) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))

SqlStmt = "SELECT A.AccNum,A.AccID, A.CustomerID, Name as CustNAme " & _
        " From QryName B Inner join (PDMaster A inner join" & _
        " PDMaster C ON C.AccID = A.AccID) On B.CustomerID = A.CustomerID" & _
        " WHERE A.CreateDate <= #" & toDate & "#" & _
        " AND (A.ClosedDate Is NULL OR A.ClosedDate >= #" & fromDate & "#)"

SqlStmt = SqlStmt & " Order By A.AgentID," & _
        IIf(m_ReportOrder = wisByAccountNo, "val(a.ACCNUM)", "IsciName")
        
gDbTrans.SqlStmt = SqlStmt
If gDbTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Sub

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
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(36, 60) 'AccountNo
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(35) 'Name
    .CellAlignment = 4: .CellFontBold = True
End With

grd.Row = 0: count = 0
Dim rowno As Long

While Not rstMain.EOF
    With grd
        rowno = rowno + 1
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        count = count + 1
        .TextMatrix(rowno, 0) = count
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
    Dim totalRow As Integer
    With grd
        rowno = rowno + 2
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .RowData(rowno - 1) = 1
        .Col = 2: .Text = GetResourceString(286) 'Grand Total
        .CellFontBold = True
        totalRow = rowno
        .RowData(rowno) = 1
    End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    
    SqlStmt = "SELECT AccId, Max(TransID) AS MaxTransID" & _
            " FROM PDTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY AccID"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("PDMonBal")
    SqlStmt = "SELECT A.AccId,Balance From PDTrans A,PDMonBal B " & _
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
    TotalBalance = 0
    
    While Not rstMain.EOF
        rowno = rowno + 1
        
        rstBalance.MoveFirst
        rstBalance.Find "ACCID = " & rstMain("AccID")
        If rstBalance.EOF Then GoTo NextAccount
        If rstBalance("Balance") = 0 Then GoTo NextAccount
        With grd
            .TextMatrix(rowno, .Col) = FormatField(rstBalance("Balance"))
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
        .Row = totalRow
        .Text = FormatCurrency(TotalBalance)
        .CellFontBold = True
        .RowData(rowno) = 1
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
    For I = 0 To grd.Cols - 1
        Wid = GetSetting(App.EXEName, "PDReport" & m_ReportType, "ColWidth" & I, 1 / grd.Cols) * grd.Width
        If Wid > grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
        If Wid <= 15 Then Wid = 20
        grd.ColWidth(I) = Wid
    Next I

End Sub

Private Sub ShowDepositGeneralLedger()
Dim count As Integer
Dim SqlStr As String
Dim rst As Recordset
Dim TransDate As Date
Dim OpeningBalance As Currency
'
RaiseEvent Processing("Reading & Verifying the records", 0)

SqlStr = "Select 'PRINCIPAL',Sum(Amount) as TotalAmount,TransDate,TransType From PDTrans " & _
        " WHERE TransDate >= #" & GetSysFormatDate(m_FromIndianDate) & "# " & _
        " And TransDate <= #" & GetSysFormatDate(m_ToIndianDate) & "#" & _
        " Group By TransDate,TransType"

gDbTrans.SqlStmt = SqlStr & " ORDER BY TransDate"
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

chkAgent.Enabled = False
RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid

With grd
    .Row = .FixedRows
    .Col = 0: .Text = GetResourceString(284) '"Opening Balnce"
    .CellAlignment = 4: .CellFontBold = True
    
    OpeningBalance = GetPDBalance(GetIndianDate(DateAdd("D", -1, m_FromDate)))
    .Col = 2: .Text = FormatCurrency(OpeningBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

Dim transType As wisTransactionTypes
Dim DepositAmount As Currency
Dim WithdrawAmount As Currency
Dim TotalDepositAmount As Currency
Dim totalWithdrawAmount As Currency
Dim PRINTTotal As Boolean
Dim SlNo As Integer
Dim rowno As Long
TransDate = rst("TransDate")
While Not rst.EOF
    If TransDate <> rst("TransDate") <> 0 Then
        With grd
            PRINTTotal = True
            rowno = rowno + 1
            If .Rows = rowno + 2 Then .Rows = .Rows + 2
            .Row = rowno
            SlNo = SlNo + 1
            .TextMatrix(rowno, 0) = SlNo
            .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
            .TextMatrix(rowno, 2) = FormatCurrency(OpeningBalance)
            .TextMatrix(rowno, 3) = FormatCurrency(DepositAmount)
            .TextMatrix(rowno, 4) = FormatCurrency(WithdrawAmount)
            
            OpeningBalance = OpeningBalance + DepositAmount - WithdrawAmount
            .TextMatrix(rowno, 5) = FormatCurrency(OpeningBalance)
            TotalDepositAmount = TotalDepositAmount + DepositAmount
            totalWithdrawAmount = totalWithdrawAmount + WithdrawAmount
            WithdrawAmount = 0: DepositAmount = 0
            TransDate = rst("TransDate")
        End With
    End If
    
    transType = FormatField(rst("TransType"))
    If transType = wDeposit Or transType = wContraDeposit Then
        DepositAmount = DepositAmount + rst("TotalAmount")
    Else
        WithdrawAmount = WithdrawAmount + rst("TotalAmount")
    End If
    
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend
    
With grd
    rowno = rowno + 1
    If .Rows < rowno + 1 Then .Rows = rowno + 1
    .Row = rowno
    SlNo = SlNo + 1
    .TextMatrix(rowno, 0) = SlNo
    .TextMatrix(rowno, 1) = GetIndianDate(TransDate)
    .TextMatrix(rowno, 2) = FormatCurrency(OpeningBalance)
    .TextMatrix(rowno, 3) = FormatCurrency(DepositAmount)
    .TextMatrix(rowno, 4) = FormatCurrency(WithdrawAmount)
    
    OpeningBalance = OpeningBalance + DepositAmount - WithdrawAmount
    .TextMatrix(rowno, 5) = FormatCurrency(OpeningBalance): .CellAlignment = 7
    TotalDepositAmount = TotalDepositAmount + DepositAmount
    totalWithdrawAmount = totalWithdrawAmount + WithdrawAmount
    WithdrawAmount = 0: DepositAmount = 0

    If PRINTTotal Then
        rowno = rowno + 2
        If .Rows < rowno + 1 Then .Rows = rowno + 1
        .Row = rowno
        .Col = 3: .Text = FormatCurrency(TotalDepositAmount)
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(totalWithdrawAmount)
        .CellAlignment = 4: .CellFontBold = True
            
        If .Rows = .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 4
        .CellAlignment = 4: .CellFontBold = True
        .Text = GetResourceString(285)  '"Totals Amount"
        .Col = 5
        .CellAlignment = 7: .CellFontBold = True
        .Text = OpeningBalance
    Else
        .RemoveItem .FixedRows
    End If

End With

If DateDiff("D", m_FromDate, m_ToDate) = 0 Or Not PRINTTotal Then
    lblReportTitle.Caption = GetResourceString(425) & " " & _
        GetResourceString(93) '"Deposit GeneralLegder
Else
    Me.lblReportTitle.Caption = GetResourceString(425) & " " & _
        GetResourceString(93) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)
End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
'Set mfrmPDReport = Nothing

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdPrint.Caption = GetResourceString(23)
Me.cmdOk.Caption = GetResourceString(11)
End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
    
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "PDReport" & m_ReportType, _
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


