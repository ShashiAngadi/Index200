VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmShgReport 
   Caption         =   "SHG Reports.."
   ClientHeight    =   5895
   ClientLeft      =   1125
   ClientTop       =   1905
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
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
Attribute VB_Name = "frmShgReport"
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
Dim m_ReportType As wis_ShgReports
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

'This Function Get the record set of Sb Balance
'having sb accounts of selp help groups only
Private Function GetSBBalanceRstAsOn(AsOnDate As Date) As Recordset

Dim rst As Recordset
Set GetSBBalanceRstAsOn = Nothing

'Get the Sb BAlance as on Date
gDbTrans.SqlStmt = "Select Balance,SbAccID,A.AccID " & _
        " From SbMaster A,ShgMAster B, SbTrans C " & _
        " Where A.AccID = B.SbAccID And C.AccID = A.AccID " & _
        " AND TransID = (Select Max(TransID) From SbTrans D" & _
            " WHEre D.AccID = A.AccID And TransDate <= #" & AsOnDate & "#)"
        
            
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                            Set GetSBBalanceRstAsOn = rst

End Function


'This Function Get the Transaction Loan Account
'of selp help groups only
Private Function GetLoanTransRst(fromDate As Date, toDate As Date, _
                                Optional IsPayment As Boolean = False) As Recordset

Set GetLoanTransRst = Nothing

Dim rst As Recordset
Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes

'Now Get the Transaction types
transType = IIf(IsPayment, wWithdraw, wDeposit)
ContraTransType = IIf(IsPayment, wContraWithdraw, wContraDeposit)


'Get the Sb BAlance as on Date
gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount,LoanID" & _
        " From LoanTrans Where LoanID In " & _
            "(Select Distinct LoanID From ShgMaster)" & _
        " And TransDate >= #" & fromDate & "# And TransDate <= #" & toDate & "#" & _
        " And (TransType = " & transType & " OR TransType = " & ContraTransType & ")" & _
        " Group By LoanId"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                                Set GetLoanTransRst = rst

End Function


'This Function Get the Transaction Od Savings account
'of selp help groups only
Private Function GetSBTransRst(fromDate As Date, toDate As Date) As Recordset

Set GetSBTransRst = Nothing

Dim rst As Recordset

'Get the Sb BAlance as on Date
gDbTrans.SqlStmt = "Select Sum(Amount) as TotalAmount,AccID" & _
        " From SBTrans Where AccID In " & _
            "(Select Distinct SBAccID  as AccID From ShgMaster)" & _
        " And TransDate >= #" & fromDate & "# And TransDate <= #" & toDate & "#" & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
        " Group By AccId"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                                Set GetSBTransRst = rst

End Function



'This Function Get the Transaction Loan Interest Amount
'of selp help group's accounts only
Private Function GetLoanIntTransRst(fromDate As Date, toDate As Date) As Recordset

Set GetLoanIntTransRst = Nothing

Dim rst As Recordset

'Get the Loan interest payment
gDbTrans.SqlStmt = "Select Sum(IntAmount+PenalIntAmount) as TotalAmount,LoanID" & _
        " From LoanIntTrans Where LoanID In " & _
            "(Select Distinct LoanID From ShgMaster)" & _
        " And TransDate >= #" & fromDate & "# And TransDate <= #" & toDate & "#" & _
        " And (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
        " Group by LoanID"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                                Set GetLoanIntTransRst = rst

End Function



'This Function Get the record set of Loan Balance
'having loan accounts of selp help groups only
Private Function GetLoanBalanceRstAsOn(AsOnDate As Date) As Recordset

Dim rst As Recordset
Set GetLoanBalanceRstAsOn = Nothing

'Get the Sb BAlance as on Date
gDbTrans.SqlStmt = "Select Balance,A.LoanID " & _
        " From ShgMAster A, LoanTrans B " & _
        " Where B.LoanID = A.LoanID " & _
        " AND TransID = (Select Max(TransID) From LoanTrans C" & _
            " WHEre C.LoanID = A.LoanID And TransDate <= #" & AsOnDate & "#)"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                        Set GetLoanBalanceRstAsOn = rst


End Function


Public Property Let ReportOrder(newOrder As wis_ReportOrder)
    m_ReportOrder = newOrder
End Property

Public Property Let ReportType(NewReportType As wis_ShgReports)
 m_ReportType = NewReportType
End Property

Private Sub SetMonthlyReportGrid()
    
Dim colno  As Integer
Dim rowno As Integer
Dim MaxCol As Integer
Dim MaxRow As Integer
Dim strText As String
Dim strDate As String

With grd
    .Clear
    .Cols = 23
    .Rows = 6
    .FixedRows = 3
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    MaxCol = .Cols - 1
    MaxRow = .FixedRows - 1
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(371, 35) ' "Name"
    .Col = 2: .Text = GetResourceString(112)  ' "Place"
    .Col = 3: .Text = GetResourceString(281)  ' "Create Date"
    
    .Col = 4: .Text = GetResourceString(49, 60) '
    .Col = 5: .Text = GetResourceString(49, 60) '
    .Col = 6: .Text = GetResourceString(49, 60) '
    
    strText = GetResourceString(411, 387) 'GetResourceString(411) & IIf(gLangOffSet, "За ", " ") & _
               '     GetResourceString(387) 'Weekly Savings
    .Col = 7: .Text = strText   ' "weekly Savings "
    .Col = 8: .Text = strText   ' "weekly Savings "
    .Col = 9: .Text = strText   ' "weekly Savings "
    
    strText = GetResourceString(58) & "" & GetResourceString(289) ' "Loans Payments"
    .Col = 10: .Text = strText   ' "Loans Payments"
    .Col = 11: .Text = strText
    .Col = 12: .Text = strText
    
    strText = GetResourceString(58) & "" & GetResourceString(216) ' "Loans rePayments"
    .Col = 13: .Text = strText
    .Col = 14: .Text = strText
    .Col = 15: .Text = strText
    .Col = 16: .Text = strText
    .Col = 17: .Text = strText
    .Col = 18: .Text = strText

    .Col = 19: .Text = GetResourceString(80, 67)   ' "Balance"
    
    .Col = 20: .Text = GetResourceString(84)   ' "Over Due"
    .Col = 21: .Text = GetResourceString(84)   ' "Over Due"

    .Col = 22: .Text = "SB " & GetResourceString(67)   ' "SB Balance"
'SECOND ROW
    .Row = 1
    .RowHeight(1) = .RowHeight(3) * 1.5
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(371, 35) ' "Name"
    .Col = 2: .Text = GetResourceString(112)  ' "Place"
    .Col = 3: .Text = GetResourceString(281)  ' "Create Date"
    
    .Col = 4: .Text = GetResourceString(52)  ' Total
    .Col = 5: .Text = GetResourceString(384) 'Sc/St
    .Col = 6: .Text = GetResourceString(237, 49) 'Other Member
    
    strText = GetUptoString(GetResourceString(250, 192))
    .Col = 7: .Text = strText    ' "upto Previous month"
    .Col = 8: .Text = GetResourceString(374, 192) ' "Current Month"
    .Col = 9: .Text = GetResourceString(52)   ' "Savings Account"

    strText = GetUptoString(GetResourceString(250, 192))
    .Col = 10: .Text = strText   ' "Loans "
    strText = GetResourceString(374) & " " & _
                GetResourceString(192) ' "Current Month Advancs"
    .Col = 11: .Text = strText   ' "Current month "
    .Col = 12: .Text = GetResourceString(52)   ' "Total"
    
    strText = GetUptoString(GetResourceString(250, 192))
    .Col = 13: .Text = strText  ' "Recovery"
    .Col = 14: .Text = strText  ' "Recovery"
    strText = GetResourceString(374, 192) ' "Current Month "
    .Col = 15: .Text = strText   ' "Current Month
    .Col = 16: .Text = strText   ' "Current month "
    .Col = 17: .Text = GetResourceString(52)   ' "Total"
    .Col = 18: .Text = GetResourceString(52)   ' "Total"

    .Col = 19: .Text = GetResourceString(80, 67)  ' "Loan Balance"
    
    .Col = 20: .Text = GetResourceString(60)   ' "No"
    .Col = 21: .Text = GetResourceString(40)   ' "Amount"

    .Col = 22: .Text = "SB " & GetResourceString(67)   ' "SB Balance"
    
    
    .MergeCells = flexMergeFree
    For rowno = 0 To .FixedRows - 1
        .Row = rowno
        .MergeRow(rowno) = True
        For colno = 0 To MaxCol
            .Col = colno
            .CellAlignment = 4
            .CellFontBold = True
            If colno < 5 Then .MergeCol(colno) = True
            If .Row = .FixedRows - 1 Then .Text = (colno + 1)
        Next
    Next
    .MergeCol(19) = True: .MergeCol(22) = True

'THIRD ROW
    .Row = 2
    .Col = 13: .Text = GetResourceString(310)  'Princpal
    .Col = 14: .Text = GetResourceString(47)   'Interest
    .Col = 15: .Text = GetResourceString(310)  'Princpal
    .Col = 16: .Text = GetResourceString(47)   'Interest
    .Col = 17: .Text = GetResourceString(310)  'Princpal
    .Col = 18: .Text = GetResourceString(47)   'Interest
    
End With

End Sub

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
        .ColWidth(ColCount) = GetSetting(App.EXEName, "SHGReport" & m_ReportType, "ColWidth" & ColCount, 1 / .Cols) * .Width
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


Private Sub ShowAccountsCreated()
'This Function Show The Selef Help group
'Formed during the specified period

'Declarig the variables
Dim SqlStmt As String
Dim rst As ADODB.Recordset

RaiseEvent Processing("Reading and verifying the records ", 0)
'Fire SQL
SqlStmt = "Select AccID,AccNum,CreateDate, Name " & _
    " FROM ShgMaster A, qryName B WHERE " & _
    " CreateDate >= #" & m_FromDate & "#" & _
    " AND CreateDate <= #" & m_ToDate & "#" & _
    " AND B.CustomerID = A.CustomerID "

If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " AND A.Gender = " & m_Gender

If m_Place <> "" Then SqlStmt = SqlStmt & " AND A.PLACE = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStmt = SqlStmt & " AND A.CASTE = " & AddQuotes(m_Caste, True)

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
        .Text = GetResourceString(370, 60) '"Acc No"
    .Col = 2: .CellFontBold = True
        .Text = GetResourceString(35) '"Name"
    .Col = 3: .CellFontBold = True
        .Text = GetResourceString(281)  '"Create Date"
End With
    RaiseEvent Initialize(0, rst.RecordCount)
    RaiseEvent Processing("Arranging the data to write into the grid. ", 0)
    
    Dim SlNo As Long
'Fill the grid
While Not rst.EOF

    With grd
        'Set next row
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        
        .Col = 0: .Text = " " & Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("Name"))
        .Col = 3: .Text = " " & FormatField(rst("CreateDate"))
    End With
    
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid ", rst.AbsolutePosition / rst.RecordCount)
    
    rst.MoveNext
Wend


lblReportTitle.Caption = GetResourceString(64) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)


End Sub
Private Sub ShowSBBalances()
'Declare the variables
Dim rst As ADODB.Recordset
Dim I As Long
Dim Total As Currency
Dim SQL As String
    
RaiseEvent Processing("Reading & Verifying the data.", 0)
 
SQL = "Select A.AccId, AccNum, Balance, Name " & _
    " From SHGMaster A,QryName B,SBtrans C where C.TransID = " & _
        "(Select MAX(TransID) from SBtrans D WHERE D.AccID = C.AccID " & _
            " AND TransDate <= #" & m_ToDate & "#)" & _
    " And C.AccId = A.SBAccId And B.CustomerId = A.CustomerId "

If m_FromAmt > 0 Then SQL = SQL & " And Balance >= " & m_FromAmt
If m_ToAmt > 0 Then SQL = SQL & " And Balance <=  " & m_ToAmt

'Query by caste
If m_Caste <> "" Then SQL = SQL & " And A.Caste =  " & AddQuotes(m_Caste, True)
'Query by PLACE
If m_Place <> "" Then SQL = SQL & " And A.Place=  " & AddQuotes(m_Place, True)
'Query by Gender
If m_Gender <> wisNoGender Then SQL = SQL & " And A.Gender =  " & m_Gender
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
    .Clear
    .Cols = 4
    .Rows = 5
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
End With

RaiseEvent Initialize(0, rst.RecordCount)
RaiseEvent Processing("Arranging the data to write into the grid.", 0)

Dim SlNo As Long
SlNo = 0
While Not rst.EOF
    DoEvents
    Me.Refresh
    'See if you have to show this record
    If FormatField(rst("Balance")) = 0 Then GoTo nextRecord
    
    'Set next row
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("Name"))
        .Col = 3: .Text = FormatField(rst("Balance"))
        Total = Total + Val(.Text) 'FormatField(Rst("Balance"))
        If Val(.Text) < 0 Then
            .Text = FormatCurrency(Abs(.Text))
            .CellForeColor = vbRed
        End If
        .CellAlignment = 7
    End With
    
nextRecord:
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing into the grid. ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Move next row
With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .CellFontBold = True
    .Text = GetResourceString(52, 42) '"Total Balances"
    .Col = 3: .CellFontBold = True
    .Text = FormatCurrency(Total): .CellAlignment = 7
End With

lblReportTitle.Caption = GetResourceString(421) & " " & _
                GetResourceString(67) & " " & _
                GetFromDateString(m_ToIndianDate)
            


End Sub

Private Sub ShowSHGGroups()
'Declare the variables
Dim rst As ADODB.Recordset
'Dim I As Long
Dim Total As Currency
Dim SQL As String
    
RaiseEvent Processing("Reading & Verifying the data.", 0)
 
SQL = "Select A.AccId, AccNum, A.Caste,A.Place, " & _
    " TotalMembers,ScStMembers,FemaleMembers,FemaleScStMembers,Name " & _
    " From SHGMaster A,QryName B where B.CustomerId = A.CustomerId "

'Query by caste
If m_Caste <> "" Then SQL = SQL & " And A.Caste =  " & AddQuotes(m_Caste, True)
'Query by PLACE
If m_Place <> "" Then SQL = SQL & " And A.Place=  " & AddQuotes(m_Place, True)
'Query by Gender
If m_Gender <> wisNoGender Then SQL = SQL & " And A.Gender =  " & m_Gender
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
    .Clear
    .Cols = 8
    .Rows = 4
    .FixedRows = 2
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(35) ' "Name"
    .Col = 2: .Text = GetResourceString(111)  ' "Caste"
    .Col = 3: .Text = GetResourceString(112)  ' "Place"
    .Col = 4: .Text = GetResourceString(49)  ' "Memebres"
    .Col = 5: .Text = GetResourceString(49)  ' "Sc St Memebers"
    .Col = 6: .Text = GetResourceString(384, 49) ' "Female Memebres"
    .Col = 7: .Text = GetResourceString(384, 49) ' "Sc St Female Memebres"
    .Row = 1
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(35) ' "Name"
    .Col = 2: .Text = GetResourceString(111)  ' "Caste"
    .Col = 3: .Text = GetResourceString(112)  ' "Place"
    .Col = 4: .Text = GetResourceString(52)  ' "Total Memebres"
    .Col = 5: .Text = GetResourceString(386) 'Female
    .Col = 6: .Text = GetResourceString(52)  ' "Total Memebres"
    .Col = 7: .Text = GetResourceString(386) ' Female
    
    Dim I As Integer
    Dim j As Integer
    
    .MergeCells = flexMergeFree
    For I = 0 To 1
        .Row = I
        .MergeRow(I) = True
        For j = 0 To .Cols - 1
            .Col = j
            .CellAlignment = 4
            .CellFontBold = True
            If I Then .MergeCol(j) = True
        Next
    Next
    
End With

RaiseEvent Initialize(0, rst.RecordCount)
RaiseEvent Processing("Arranging the data to write into the grid.", 0)

Dim SlNo As Long
SlNo = 0
While Not rst.EOF
    DoEvents
    Me.Refresh
    'See if you have to show this record
    'If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
    
    'Set next row
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("Name"))
        .Col = 2: .Text = FormatField(rst("Caste"))
        .Col = 3: .Text = FormatField(rst("Place"))
        .Col = 4: .Text = FormatField(rst("TotalMembers"))
        .Col = 5: .Text = FormatField(rst("FemaleMembers"))
        .Col = 6: .Text = FormatField(rst("ScStMembers"))
        .Col = 7: .Text = FormatField(rst("FemaleScStMembers"))
    End With
    
nextRecord:
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing to the grid. ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Move next row
With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1: .CellFontBold = True
    .Text = GetResourceString(52)
'    .Col = 3: .CellFontBold = True
    .Text = FormatCurrency(Total): .CellAlignment = 7
End With

lblReportTitle.Caption = GetResourceString(370) & " " & _
                GetResourceString(67) & " " & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
            

End Sub

Private Sub ShowScStMembers()
'Declare the variables
Dim rst As ADODB.Recordset
'Dim I As Long
Dim Total As Currency
Dim SQL As String
    
RaiseEvent Processing("Reading & Verifying the data.", 0)
 
SQL = "Select A.AccId, AccNum, A.Caste,A.Place, " & _
    " TotalMembers,ScStMembers,FemaleMembers,FemaleScStMembers,B.Name " & _
    " From SHGMaster A,QryName B where B.CustomerId = A.CustomerId "

'Query by caste
If m_Caste <> "" Then SQL = SQL & " And A.Caste =  " & AddQuotes(m_Caste, True)
'Query by PLACE
If m_Place <> "" Then SQL = SQL & " And A.Place=  " & AddQuotes(m_Place, True)
'Query by Gender
If m_Gender <> wisNoGender Then SQL = SQL & " And A.Gender =  " & m_Gender
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
    .Clear
    .Cols = 5
    .Rows = 4
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(35) ' "Name"
    .Col = 2: .Text = GetResourceString(52, 49) ' "Memberes"
    .Col = 3: .Text = GetResourceString(384, 49) ' "Sc St Memebres"
    .Col = 4: .Text = GetResourceString(384, 386) ' "Sc St Female Memebres"
    
    Dim j As Integer
    
    .MergeCells = flexMergeFree
    For j = 0 To .Cols - 1
        .Col = j
        .CellAlignment = 4
        .CellFontBold = True
    Next
    
End With

RaiseEvent Initialize(0, rst.RecordCount)
RaiseEvent Processing("Arranging the data to write into the grid.", 0)

Dim SlNo As Long
Dim TotalMem As Integer
Dim TotalScSt As Integer
Dim TotalFemScSt As Integer

SlNo = 0
While Not rst.EOF
    DoEvents
    Me.Refresh
    'See if you have to show this record
    'If FormatField(Rst("Balance")) = 0 Then GoTo NextRecord
    
    'Set next row
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("Name"))
        .Col = 2: .Text = FormatField(rst("TotalMembers"))
        TotalMem = TotalMem + Val(.Text)
        
        .Col = 3: .Text = FormatField(rst("ScStMembers"))
        TotalScSt = TotalScSt + Val(.Text)
        .Col = 4: .Text = FormatField(rst("FemaleScStMembers"))
        TotalFemScSt = TotalFemScSt + Val(.Text)
    End With
    
nextRecord:
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing into the grid. ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

'Move next row
With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1: .CellFontBold = True
    .Text = GetResourceString(52)
    
    .Col = 2: .Text = TotalMem
    
    .Col = 3: .Text = TotalScSt
    .Col = 4: .Text = TotalFemScSt
End With

lblReportTitle.Caption = GetResourceString(371) & " - " & _
                GetResourceString(384, 49) & " " & _
                GetFromDateString(m_ToIndianDate)
            

End Sub


Private Sub ShowMonthlyReport()

'Declare the variables
Dim rstMain As ADODB.Recordset
Dim Total As Currency
Dim SQL As String

RaiseEvent Processing("Reading & Verifying the data.", 0)
 
SQL = "Select A.AccId, AccNum, A.Caste,A.Place, " & _
    " SbAccId,LoanID,a.CreateDate,TotalMembers,ScStMembers,B.Name " & _
    " From SHGMaster A,qryName B where B.CustomerId = A.CustomerId "

'Query by caste
If m_Caste <> "" Then SQL = SQL & " And A.Caste =  " & AddQuotes(m_Caste, True)
'Query by PLACE
If m_Place <> "" Then SQL = SQL & " And A.Place=  " & AddQuotes(m_Place, True)
'Query by Gender
If m_Gender <> wisNoGender Then SQL = SQL & " And A.Gender =  " & m_Gender

If m_ReportOrder = wisByName Then
    SQL = SQL & " Order By IsciName"
Else
    SQL = SQL & " Order By val(AccNum)"
End If

gDbTrans.SqlStmt = SQL
'Cret ShG Query
'If gDbTrans.CreateView("QRYShgList") <= 0 Then Exit Sub
If gDbTrans.Fetch(rstMain, adOpenForwardOnly) <= 0 Then Exit Sub


'Initialize the grid
Call SetMonthlyReportGrid

RaiseEvent Initialize(0, rstMain.RecordCount + 1)
RaiseEvent Processing("Arranging the data to write into the grid.", 0)

Dim SlNo As Long
Dim rstSbBalance As Recordset
Dim rstLoanBalance As Recordset
'Dim rstPrevSbBalance As Recordset
Dim rstPrevLoanBalance  As Recordset
Dim rstSavings As Recordset
Dim rstPrevSavings As Recordset

Dim rstLoanAdv As Recordset
Dim rstLoanRepay As Recordset
Dim rstPrevLoanAdv As Recordset
Dim rstPrevLoanRepay As Recordset
Dim rstInt As Recordset
Dim rstPrevInt As Recordset

Dim PrevDate As Date
Dim fromDate As Date

fromDate = m_ToDate
'Now Get the Last date of the Previous month
PrevDate = GetSysLastDate(DateAdd("m", -1, fromDate))

'Now Get All Record Sets
'Sb Balnce as oN date
Set rstSbBalance = GetSBBalanceRstAsOn(fromDate)
'sb Balance as on the end of last month
'Set rstPrevSbBalance = GetSBBalanceRstAsOn(PrevDate)

'Loan Balance
Set rstLoanBalance = GetLoanBalanceRstAsOn(fromDate)
'Loan balance as on end of last month
Set rstPrevLoanBalance = GetLoanBalanceRstAsOn(PrevDate)

'Now Get the Loan Transaction up to previous month
Set rstPrevSavings = GetSBTransRst(FinUSFromDate, PrevDate)
Set rstPrevLoanAdv = GetLoanTransRst(FinUSFromDate, PrevDate, True)
Set rstPrevLoanRepay = GetLoanTransRst(FinUSFromDate, PrevDate, False)
Set rstPrevInt = GetLoanIntTransRst(FinUSFromDate, PrevDate)

'Get the Loan Transaction of the current month
Set rstSavings = GetSBTransRst(DateAdd("d", 1, PrevDate), fromDate)
Set rstLoanAdv = GetLoanTransRst(DateAdd("d", 1, PrevDate), fromDate, True)
Set rstLoanRepay = GetLoanTransRst(DateAdd("d", 1, PrevDate), fromDate, False)
Set rstInt = GetLoanIntTransRst(DateAdd("d", 1, PrevDate), fromDate)

Dim TotalMem As Integer
Dim ScStMem As Integer

Dim SbACCID As Long
Dim SBBalance As Currency
'Dim PrevSbBalance As Currency
Dim Savings As Currency
Dim PrevSavings As Currency

Dim DueInst As Integer
Dim DueAmount As Currency

Dim LoanID As Long
Dim loanClass As New clsLoan
Dim LoanBalance As Currency
'Dim PrevLoanBalance As Currency
Dim LoanAdvance As Currency
Dim PrevLoanAdvance As Currency
Dim LoanRec As Currency
Dim PrevLoanRec As Currency
Dim IntAmount As Currency
Dim PrevIntAmount As Currency


SlNo = 0
grd.Row = grd.FixedRows
While Not rstMain.EOF
    'Get the Loan Account Id no
    SbACCID = FormatField(rstMain("SbAccID"))
    LoanID = FormatField(rstMain("LoanId"))
    
    'Savings account
    If SbACCID Then
        If Not rstSbBalance Is Nothing Then
            rstSbBalance.MoveFirst
            rstSbBalance.Find "AccID = " & SbACCID
            If Not rstSbBalance.EOF Then SBBalance = FormatField(rstSbBalance("Balance"))
        End If
        If Not rstPrevSavings Is Nothing Then
            rstPrevSavings.MoveFirst
            rstPrevSavings.Find "AccID = " & SbACCID
            If Not rstPrevSavings.EOF Then PrevSavings = FormatField(rstPrevSavings("TOtalAmount"))
        End If
        If Not rstSavings Is Nothing Then
            rstSavings.MoveFirst
            rstSavings.Find "AccID = " & SbACCID
            If Not rstSavings.EOF Then Savings = FormatField(rstSavings("TotalAmount"))
        End If
    End If
    
    'Loan Account details
    If LoanID Then
        If Not rstLoanBalance Is Nothing Then
            rstLoanBalance.MoveFirst
            rstLoanBalance.Find "LoanID = " & LoanID
            If Not rstLoanBalance.EOF Then _
                LoanBalance = FormatField(rstLoanBalance("Balance"))
        End If
        If Not rstPrevLoanAdv Is Nothing Then
            rstPrevLoanAdv.MoveFirst
            rstPrevLoanAdv.Find "LoanID = " & LoanID
            If Not rstPrevLoanAdv.EOF Then PrevLoanAdvance = FormatField(rstPrevLoanAdv("TotalAmount"))
        End If
        'Loan Recovery
        If Not rstPrevLoanRepay Is Nothing Then
            rstPrevLoanRepay.MoveFirst
            rstPrevLoanRepay.Find "LoanID = " & LoanID
            If Not rstPrevLoanRepay.EOF Then _
                PrevLoanRec = FormatField(rstPrevLoanRepay("TotalAmount"))
        End If
        If Not rstLoanRepay Is Nothing Then
            rstLoanRepay.MoveFirst
            rstLoanRepay.Find "LoanID = " & LoanID
            If Not rstLoanRepay.EOF Then LoanRec = FormatField(rstLoanRepay("TotalAmount"))
        End If
            
        'Loan INterest Recovery
        If Not rstPrevInt Is Nothing Then
            rstPrevInt.MoveFirst
            rstPrevInt.Find "LoanID = " & LoanID
            If Not rstPrevInt.EOF Then PrevIntAmount = FormatField(rstPrevInt("TotalAmount"))
        End If
        If Not rstInt Is Nothing Then
            rstInt.MoveFirst
            rstInt.Find "LoanID = " & LoanID
            If Not rstInt.EOF Then IntAmount = FormatField(rstInt("TotalAmount"))
        End If
        
        'Over Due Inst
        DueAmount = loanClass.OverDueAmount(LoanID, , m_ToDate)
        DueInst = loanClass.DueInstallments(LoanID, m_ToDate)
    End If
    
    DoEvents
    Me.Refresh
    'Set next row
    With grd
        If .Rows = .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rstMain("Name"))
        .Col = 2: .Text = FormatField(rstMain("Place"))
        .Col = 3: .Text = FormatField(rstMain("CreateDate"))
        TotalMem = FormatField(rstMain("TotalMembers"))
        ScStMem = FormatField(rstMain("ScStMembers"))
        .Col = 4: .Text = TotalMem
        .Col = 5: If ScStMem Then .Text = ScStMem
        .Col = 6: .Text = TotalMem - ScStMem
        
        If SbACCID Then
            .Col = 7: .Text = PrevSavings
            .Col = 8: .Text = Savings
            .Col = 9: .Text = (PrevSavings + Savings)
            
            .Col = 22: .Text = SBBalance
            
            Savings = 0: PrevSavings = 0: SBBalance = 0
        End If
        'Fill the Loan Deatils
        If LoanID Then
            'Loan Advance
            .Col = 10: .Text = PrevLoanAdvance
            .Col = 11: .Text = LoanAdvance
            .Col = 12: .Text = LoanAdvance + PrevLoanAdvance
            
            'Loan Recovery
            .Col = 13: .Text = PrevLoanRec
            .Col = 14: .Text = PrevIntAmount
            .Col = 15: .Text = LoanRec
            .Col = 16: .Text = IntAmount
            'Total Recovery
            .Col = 17: .Text = LoanRec + PrevLoanRec
            .Col = 18: .Text = IntAmount + PrevIntAmount
            
            .Col = 19: .Text = LoanBalance
            If DueInst Then .Col = 20: .Text = DueInst
            If DueAmount Then .Col = 21: .Text = DueAmount
            
            DueInst = 0: DueAmount = 0
            LoanRec = 0: PrevLoanRec = 0
            LoanAdvance = 0: PrevLoanAdvance = 0
            LoanBalance = 0
        End If
        
    End With
    
nextRecord:
    
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Writing into the grid. ", rstMain.AbsolutePosition / rstMain.RecordCount)
    rstMain.MoveNext
Wend


Set loanClass = Nothing

lblReportTitle.Caption = GetResourceString(371, 463, 430) & " " & GetMonthString(Month(m_ToDate))
            

End Sub


                
Private Sub ShowSBMonthlyBalances()

Dim count As Long
Dim totalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SqlStmt As String

Dim fromDate As Date
Dim toDate As Date

'Get the Last day of the given month
toDate = GetSysLastDate(m_ToDate)

'FromDate = "3/31/" & IIf(Month(ToDate) > 3, Year(ToDate), Year(ToDate) - 1)
fromDate = GetSysLastDate(m_FromDate)

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(463) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(m_FromDate)), GetMonthString(Month(m_ToDate)))

SqlStmt = "SELECT C.AccNum,A.SbAccID,A.CreateDate,A.Place, Name as CustNAme " & _
    " From SHGMaster A,qryName B,SBMaster C WHERE A.CreateDate <= #" & toDate & "#" & _
    " AND (C.ClosedDate Is NULL OR C.Closeddate >= #" & fromDate & "#)" & _
    " AND  B.CustomerID = A.CustomerID" & _
    " AND C.AccID = A.SbAccID "
    
If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & " Order By val(c.AccNum)"
Else
    SqlStmt = SqlStmt & " Order By IsciName"
End If

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
    .Cols = 5
    .Rows = 5
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'Sl No
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(35) 'Name
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(112) 'Place
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = GetResourceString(281) 'Create Date
    .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = GetResourceString(36, 60) 'Account No
    .CellAlignment = 4: .CellFontBold = True
End With

grd.Row = 0: count = 0
While Not rstMain.EOF
    With grd
        If .Rows < .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1: count = count + 1
        .Col = 0: .Text = count
        .Col = 1: .Text = FormatField(rstMain("CustName"))
        .Col = 2: .Text = FormatField(rstMain("Place"))
        .Col = 3: .Text = FormatField(rstMain("CreateDate"))
        .Col = 4: .Text = FormatField(rstMain("AccNum"))
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
    
    rstMain.MoveNext
Wend

With grd
    If .Rows < .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    If .Rows < .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(286) 'Grand Total
    .CellFontBold = True
End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

fromDate = "4/30/" & Year(fromDate)

Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    SqlStmt = "SELECT [AccId], Max([TransID]) AS MaxTransID" & _
            " FROM SBTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY [AccID];"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("SBMonBal")
    SqlStmt = "SELECT A.AccId,Balance From SBTrans A,SBMonBal B " & _
        " Where B.AccId = A.AccID ANd  TransID =MaxTransID"
    gDbTrans.SqlStmt = SqlStmt
    
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0
        .Col = .Cols - 1: .Text = GetMonthString(Month(fromDate)) & _
                " " & GetResourceString(42)
        .CellAlignment = 4: .CellFontBold = True
    End With
    
    rstMain.MoveFirst
    TotalBalance = 0
    With grd
        .Row = 0
        .Col = .Cols - 1
    End With
    
    While Not rstMain.EOF
        Balance = 0
        rstBalance.MoveFirst
        rstBalance.Find "ACCID = " & rstMain("SBAccID")
        If Not rstBalance.EOF Then Balance = FormatField(rstBalance("Balance"))
        
        With grd
            .Row = .Row + 1
            If Balance Then .Text = FormatCurrency(Balance)
        End With
        
        TotalBalance = TotalBalance + Balance
        
        ProcCount = ProcCount + 1
        DoEvents
        If gCancel Then rstMain.MoveLast
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
    
        rstMain.MoveNext
    Wend
    With grd
        .Row = .Row + 2
        .Text = FormatCurrency(TotalBalance)
        .CellFontBold = True
    End With
    
NextMonth:
'    rstBalance.MoveFirst
    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(421) & " " & _
        GetResourceString(463) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))


Exit Sub
ErrLine:
    MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear

End Sub

Private Sub ShowLoanMonthlyBalances()

Dim count As Long
Dim totalCount As Long
Dim ProcCount As Long

Dim rstMain As Recordset
Dim SqlStmt As String

Dim fromDate As Date
Dim toDate As Date

'Fet the Last day of the given month
toDate = GetSysLastDate(m_ToDate)
fromDate = GetSysLastDate(m_FromDate)

SqlStmt = "SELECT C.AccNum, A.LoanID,A.Place,C.IssueDate,LoanAmount,B.Name as CustNAme " & _
        " From SHGMaster A,QryName B,LoanMaster C WHERE " & _
        " B.CustomerID = A.CustomerID" & _
        " AND C.LoanID = A.LoanID "

If m_ReportOrder = wisByAccountNo Then
    SqlStmt = SqlStmt & " Order By val(C.ACCNUM)"
Else
    SqlStmt = SqlStmt & " Order By IsciName"
End If

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
    .Cols = 6
    .Rows = 5
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) 'Sl No
        .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(35) 'Name
        .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(80, 36, 60) 'LOan AccountNo
        .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = GetResourceString(112) 'Name
        .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = GetResourceString(340) 'Issue date
        .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .Text = GetResourceString(80, 40) 'Loan Amount
        .CellAlignment = 4: .CellFontBold = True
End With

grd.Row = 0: count = 0

While Not rstMain.EOF
    With grd
        If .Rows < .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1: count = count + 1
        .Col = 0: .Text = count
        .Col = 1: .Text = FormatField(rstMain("CustName"))
        .Col = 2: .Text = FormatField(rstMain("AccNum"))
        .Col = 3: .Text = FormatField(rstMain("Place"))
        .Col = 4: .Text = FormatField(rstMain("IssueDate"))
        .Col = 5: .Text = FormatField(rstMain("LoanAmount"))
    End With
    
    ProcCount = ProcCount + 1
    DoEvents
    If gCancel Then rstMain.MoveLast
    RaiseEvent Processing("Inserting customer Name", ProcCount / totalCount)
    rstMain.MoveNext
Wend

With grd
    If .Rows < .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    If .Rows < .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(286) 'Grand Total
    .CellFontBold = True
End With

Dim Balance As Currency
Dim TotalBalance As Currency
Dim rstBalance As Recordset

'Get the Last date of First Month
    'First get  the Next Month
    fromDate = GetSysLastDate(m_FromDate)
    
Do
    If DateDiff("d", fromDate, toDate) < 0 Then Exit Do
    SqlStmt = "SELECT [LoanId], Max([TransID]) AS MaxTransID" & _
            " FROM LoanTrans Where TransDate <= #" & fromDate & "# " & _
            " GROUP BY [LoanID];"
    gDbTrans.SqlStmt = SqlStmt
    gDbTrans.CreateView ("LoanMonBal")
    SqlStmt = "SELECT A.LoanId,Balance From LoanTrans A,LoanMonBal B " & _
        " Where B.LOanId = A.LoanID ANd  TransID =MaxTransID"
    gDbTrans.SqlStmt = SqlStmt
    
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then GoTo NextMonth
    
    With grd
        .Cols = .Cols + 1
        .Row = 0
        .Col = .Cols - 1: .Text = GetMonthString(Month(fromDate)) '& " " & GetResourceString(42)
        .CellAlignment = 4: .CellFontBold = True
    End With
    
    rstMain.MoveFirst
    TotalBalance = 0
    With grd
        .Row = 0
        .Col = .Cols - 1
    End With
    
    While Not rstMain.EOF
        grd.Row = grd.Row + 1
        Balance = 0
        rstBalance.MoveFirst
        rstBalance.Find "LoanID = " & rstMain("LOanID")
        If Not rstBalance.EOF Then Balance = FormatField(rstBalance("Balance"))
        
        grd.Text = FormatCurrency(Balance)
        TotalBalance = TotalBalance + Balance
        
        ProcCount = ProcCount + 1
        DoEvents
        If gCancel Then rstMain.MoveLast
        RaiseEvent Processing("Calculating deposit balance", ProcCount / totalCount)
        rstMain.MoveNext
    Wend
    With grd
        .Row = .Row + 2
        .Text = FormatCurrency(TotalBalance)
        .CellFontBold = True
    End With
    
NextMonth:
    
    fromDate = DateAdd("D", 1, fromDate)
    fromDate = DateAdd("m", 1, fromDate)
    fromDate = DateAdd("D", -1, fromDate)
Loop

'Set the Title for the Report.
lblReportTitle.Caption = GetResourceString(371) & " " & _
        GetResourceString(80) & " " & _
        GetResourceString(463) & " " & _
        GetResourceString(42) & " " & _
        GetFromDateString(GetMonthString(Month(fromDate)), GetMonthString(Month(toDate)))

Exit Sub


ErrLine:
    MsgBox "Error MonBalance", vbExclamation, wis_MESSAGE_TITLE
    Err.Clear

End Sub

Private Function ShowLoanBalance() As Boolean

ShowLoanBalance = False
Err.Clear
On Error GoTo ExitLine

Dim SqlStmt As String
'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)
    
SqlStmt = "Select B.LoanId, B.AccNum, Balance, A.Place, A.Caste,D.Name as CustName " & _
    " From ShgMaster A, LoanMaster B, LoanTrans C, QryName D " & _
    " Where B.LoanId = A.LoanId And C.LoanId = B.LoanId " & _
    " And D.CustomerId = A.CustomerId And TransId = " & _
        "(Select Max(TransId) From LoanTrans E Where E.LoanId = A.LoanId " & _
            " And TransDate <= #" & m_ToDate & "# ) "


If m_FromAmt <> 0 Then SqlStmt = SqlStmt & " And  Balance >= " & m_FromAmt
If m_ToAmt <> 0 Then SqlStmt = SqlStmt & " And Balance <= " & m_ToAmt

If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " And A.Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " And A.Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And A.Gender = " & m_Gender

'Quaring the loanmaster
gDbTrans.SqlStmt = SqlStmt & " Order by " & IIf(m_ReportOrder = wisByAccountNo, "val(B.AccNum)", "IsciName")

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function
    Dim TotalBalance As Currency
    Dim SubBalance As Currency
    Dim SlNo As Long
    
    
    RaiseEvent Initialize(0, rst.RecordCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    ' InitGird
With grd
    .Clear
    .Rows = 10
    .Cols = 4
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(35) '"Name"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .Text = GetResourceString(80, 60) '"Acc No"
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = GetResourceString(67) '"Balance"
    .CellAlignment = 4: .CellFontBold = True
    
End With
    
    RaiseEvent Initialize(0, rst.RecordCount)
    RaiseEvent Processing("Arranging the data to write into the grid. ", 0)
    
Dim TotalNo As Integer
TotalNo = rst.RecordCount + 2
SlNo = 1
While Not rst.EOF
    With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("CustName"))
        .Col = 2: .Text = FormatField(rst("AccNum"))
        .Col = 3: .Text = FormatField(rst("Balance"))
        TotalBalance = TotalBalance + Val(.Text)
        
    End With
    
    SlNo = SlNo + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", SlNo / TotalNo)
    rst.MoveNext
    
Wend

lblReportTitle.Caption = GetResourceString(67) & " " & _
        GetResourceString(58) & " " & _
            GetFromDateString(m_ToIndianDate)

With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1: .Text = GetResourceString(286)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

    ShowLoanBalance = True
    'Call grd_LostFocus

ExitLine:
    If Err Then
        MsgBox "ERROR ReportLoanBalance" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
    End If
End Function


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
With grd
    .Rows = 50
    .Cols = 1
    .FixedCols = 0
    .Row = 1
    
    .Text = GetResourceString(278)   '"No Records Available"
    .CellAlignment = 4: .CellFontBold = True
End With

Screen.MousePointer = vbHourglass
        
gCancel = 0
If m_ReportType = wisShgCreated Then
    Call ShowAccountsCreated
ElseIf m_ReportType = wisShgList Then
    Call ShowSHGGroups
ElseIf m_ReportType = wisShgSbMonBalance Then
    Call ShowSBMonthlyBalances
ElseIf m_ReportType = wisShgLoanMonBalnace Then
    Call ShowLoanMonthlyBalances
ElseIf m_ReportType = wisSHGSbBalance Then
    Call ShowSBBalances
ElseIf m_ReportType = wisShgLoanBalance Then
    Call ShowLoanBalance
ElseIf m_ReportType = wisShgMonthlyStmt Then
    Call ShowMonthlyReport
ElseIf m_ReportType = wisShgScStMembers Then
    Call ShowScStMembers
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
        Call SaveSetting(App.EXEName, "SHGReport" & m_ReportType, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount
End Sub


Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_TotalCount = MaxCount
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


