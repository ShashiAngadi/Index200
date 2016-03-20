VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLoanView 
   Caption         =   "Loan Reports .."
   ClientHeight    =   6360
   ClientLeft      =   1650
   ClientTop       =   1800
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   6735
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1170
      TabIndex        =   1
      Top             =   5580
      Width           =   5205
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   1680
         TabIndex        =   3
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Close"
         Height          =   400
         Left            =   3300
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4905
      Left            =   90
      TabIndex        =   0
      Top             =   570
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8652
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
      Left            =   2400
      TabIndex        =   4
      Top             =   90
      Width           =   1635
   End
End
Attribute VB_Name = "frmLoanView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)
Public Event WindowClosed()


Private WithEvents m_grdPrint  As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private WithEvents m_frmCancel As frmCancel
Attribute m_frmCancel.VB_VarHelpID = -1
Private m_Count As Long
Private m_MaxCount As Long

Private m_FromIndianDate As String
Private m_FromDate As Date
Private m_ToIndianDate As String
Private m_ToDate As Date
Private m_FromAmt As Currency
Private m_ToAmt As Currency

Private m_Caste As String
Private m_Place As String
Private m_Gender As Byte
Private m_Purpose As String
Private m_SqlCondition As String

Private m_SchemeId As Integer
Private m_ReportType As wis_LoanReports
Private m_ReportOrder As wis_ReportOrder
Private m_AccGroupID As Integer


Public Property Let AccountGroup(NewValue As Integer)
    m_AccGroupID = NewValue
End Property


Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property

Public Property Let LoanPurpose(strPurpose As String)
    m_Purpose = strPurpose
End Property

Private Function ReportInterestReceivable() As Boolean
Dim rst As Recordset
'Dim rstReceivable As Recordset
Dim LoanID As Long


ReportInterestReceivable = False
Err.Clear
On Error GoTo ExitLine

Dim SqlStmt As String
'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)
    
SqlStmt = "Select B.LoanId, AccNum,C.MemberNum, Balance as IntBalance,  " & _
    " Place, Caste,A.SchemeID, SchemeName, Name as CustName " & _
    " From LoanMaster A, LoanIntReceivable B, QryMemName C, LoanScheme D " & _
    " Where TransId = (Select Max(TransId) From LoanIntReceivable F " & _
        " Where F.LoanId = A.LoanId " & _
        " And TransDate <= #" & m_ToDate & "# ) " & _
    " AND C.MemID = A.MemID ANd A.LoanID = B.LoanID AND D.SchemeID = A.SchemeID "

If m_SchemeId Then SqlStmt = SqlStmt & " And A.SchemeId = " & m_SchemeId
If m_FromAmt <> 0 Then SqlStmt = SqlStmt & " And  Balance >= " & m_FromAmt
If m_ToAmt <> 0 Then SqlStmt = SqlStmt & " And Balance <= " & m_ToAmt

If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroupID

'Quaring the loanmaster

gDbTrans.SqlStmt = SqlStmt & " Order by A.SchemeId, val(A.AccNum)"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    Call PrintNoRecords(grd)
    Exit Function
End If
    

    Dim TotalReceivable As Currency
    Dim SubReceivable As Currency
    Dim subInt As Currency
    Dim TotalInt As Currency
    
    Dim SlNo As Long
    Dim l_SchemeID As Integer
    Dim SchemeName As String
    
    Dim totalCount As Long
    totalCount = rst.RecordCount + 2
    RaiseEvent Initialise(0, totalCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    ' InitGird
    grd.Clear
    grd.FixedRows = 1
    Dim ColWid As Single
    
    ColWid = grd.Width / grd.Cols
    grd.Row = 0
    grd.FormatString = ">Sl No|<Loan Id|<CustomerName|>Balance Interest|>Interest"
    Call InitGrid

grd.Row = grd.FixedRows
    
totalCount = rst.RecordCount + 2
Dim loanClass As clsLoan
Set loanClass = New clsLoan
SchemeName = FormatField(rst("SchemeName"))
l_SchemeID = FormatField(rst("SchemeID"))
SlNo = 1
While Not rst.EOF
    With grd
        If l_SchemeID <> FormatField(rst("SchemeID")) Then
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 3: .Text = SchemeName & " " & GetResourceString(304)
            grd.CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(SubReceivable)
            grd.CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(subInt)
            grd.CellAlignment = 7: .CellFontBold = True
            
            TotalInt = TotalInt + subInt
            TotalReceivable = TotalReceivable + SubReceivable
            SubReceivable = 0: subInt = 0
            
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1: SlNo = 1
            SchemeName = FormatField(rst("SchemeName"))
            l_SchemeID = FormatField(rst("SchemeID"))
        End If
        LoanID = FormatField(rst("LoanID"))
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("MemberNum"))
        .Col = 3: .Text = FormatField(rst("CustName"))
        .Col = 4: .Text = FormatField(rst("intBalance")): .CellAlignment = 7
        SubReceivable = SubReceivable + Val(.Text)
        .Col = 5: .Text = loanClass.RegularInterest(LoanID, , m_ToDate)
        .CellAlignment = 7
        subInt = subInt + Val(.Text)
        
    End With
    
    SlNo = SlNo + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", SlNo / totalCount)
    rst.MoveNext
Wend

Set loanClass = Nothing

With grd
    If m_SchemeId = 0 And TotalReceivable > 0 Then
        lblReportTitle.Caption = GetResourceString(376) & _
            " " & GetResourceString(47) & " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(SubReceivable): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(subInt): .CellFontBold = True
    Else
        lblReportTitle.Caption = SchemeName & " " & GetResourceString(376) & _
             " " & GetResourceString(47) & " " & GetFromDateString(m_ToIndianDate)
    End If
    TotalInt = TotalInt + subInt
    TotalReceivable = TotalReceivable + SubReceivable
            
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(286)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(TotalReceivable)
    .CellAlignment = 7: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(TotalInt)
    .CellAlignment = 7: .CellFontBold = True
End With

ReportInterestReceivable = True

ExitLine:
    Screen.MousePointer = vbDefault
End Function


Private Function ReportInterestCalculate() As Boolean
Dim rst As Recordset
'Dim rstReceivable As Recordset
Dim LoanID As Long


ReportInterestCalculate = False
Err.Clear
On Error GoTo ExitLine

Dim SqlStmt As String
'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)
    
SqlStmt = "Select B.LoanId, AccNum,C.MemberNum, Balance ,  " & _
    " Place, Caste,A.SchemeID, SchemeName, Name as CustName " & _
    " From LoanMaster A, LoanTrans B, QryMemName C, LoanScheme D " & _
    " Where TransId = (Select Max(TransId) From LoanTrans F " & _
        " Where F.LoanId = A.LoanId " & _
        " And TransDate <= #" & m_ToDate & "# ) " & _
    " AND Balance > 0 and C.MemID = A.MemID ANd A.LoanID = B.LoanID AND D.SchemeID = A.SchemeID "

If m_SchemeId Then SqlStmt = SqlStmt & " And A.SchemeId = " & m_SchemeId
If m_FromAmt <> 0 Then SqlStmt = SqlStmt & " And  Balance >= " & m_FromAmt
If m_ToAmt <> 0 Then SqlStmt = SqlStmt & " And Balance <= " & m_ToAmt

If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroupID

'Quaring the loanmaster

gDbTrans.SqlStmt = SqlStmt & " Order by A.SchemeId, val(A.AccNum)"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    Call PrintNoRecords(grd)
    Exit Function
End If
    

    Dim TotalReceivable As Currency
    Dim SubReceivable As Currency
    Dim subInt As Currency
    Dim TotalInt As Currency
    
    Dim SlNo As Long
    Dim l_SchemeID As Integer
    Dim SchemeName As String
    
    Dim totalCount As Long
    totalCount = rst.RecordCount + 2
    RaiseEvent Initialise(0, totalCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    ' InitGird
    grd.Clear
    grd.FixedRows = 1
    Dim ColWid As Single
    
    ColWid = grd.Width / grd.Cols
    grd.Row = 0
    grd.FormatString = ">Sl No|<Loan Id|<CustomerName|>Balance Interest|>Interest"
    Call InitGrid

grd.Row = grd.FixedRows
    
totalCount = rst.RecordCount + 2
Dim loanClass As clsLoan
Set loanClass = New clsLoan
SchemeName = FormatField(rst("SchemeName"))
l_SchemeID = FormatField(rst("SchemeID"))
SlNo = 1
While Not rst.EOF
    With grd
        If l_SchemeID <> FormatField(rst("SchemeID")) Then
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 3: .Text = SchemeName & " " & GetResourceString(304)
            grd.CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(SubReceivable)
            grd.CellAlignment = 7: .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(subInt)
            grd.CellAlignment = 7: .CellFontBold = True
            
            TotalInt = TotalInt + subInt
            TotalReceivable = TotalReceivable + SubReceivable
            SubReceivable = 0: subInt = 0
            
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1: SlNo = 1
            SchemeName = FormatField(rst("SchemeName"))
            l_SchemeID = FormatField(rst("SchemeID"))
        End If
        LoanID = FormatField(rst("LoanID"))
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("MemberNum"))
        .Col = 3: .Text = FormatField(rst("CustName"))
        .Col = 4: .Text = FormatField(rst("Balance")): .CellAlignment = 7
        SubReceivable = SubReceivable + Val(.Text)
        .Col = 5: .Text = loanClass.RegularInterest(LoanID, , m_ToDate)
        .CellAlignment = 7
        subInt = subInt + Val(.Text)
        
    End With
    
    SlNo = SlNo + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", SlNo / totalCount)
    rst.MoveNext
Wend

Set loanClass = Nothing

With grd
    If m_SchemeId = 0 And TotalReceivable > 0 Then
        lblReportTitle.Caption = GetResourceString(376) & _
            " " & GetResourceString(47) & " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(SubReceivable): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(subInt): .CellFontBold = True
    Else
        lblReportTitle.Caption = SchemeName & " " & GetResourceString(376) & _
             " " & GetResourceString(47) & " " & GetFromDateString(m_ToIndianDate)
    End If
    TotalInt = TotalInt + subInt
    TotalReceivable = TotalReceivable + SubReceivable
            
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 3: .Text = GetResourceString(286)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(TotalReceivable)
    .CellAlignment = 7: .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(TotalInt)
    .CellAlignment = 7: .CellFontBold = True
End With

ReportInterestCalculate = True

ExitLine:
    Screen.MousePointer = vbDefault
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
        'm_ToIndianDate = GetIndianDate(m_ToDate)
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_FromIndianDate = NewDate
        m_FromDate = GetSysFormatDate(NewDate)
        'm_FromIndianDate = GetIndianDate(m_FromDate)
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

Public Property Let ReportType(RepType As wis_LoanReports)
    m_ReportType = RepType
End Property

Public Property Let LoanSchemeType(LoanType As Integer)
    m_SchemeId = LoanType
End Property


Private Sub InitGrid(Optional Resize As Boolean)

If m_ReportType = repLoanDailyCash Then
    With grd
        .Clear
        .Rows = 15
        .Cols = 12
        .FixedCols = 1
        .FixedRows = 3
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'Date
        .Col = 2: .Text = GetResourceString(58, 60) 'Loan No
        .Col = 3: .Text = GetResourceString(49, 60) 'Memebr No
        .Col = 4: .Text = GetResourceString(35) 'Name
        .Col = 5: .Text = GetResourceString(81) 'Loan Issued
        .Col = 6: .Text = GetResourceString(81) 'LOan Issued
        .Col = 7: .Text = GetResourceString(82) 'repayment Made
        .Col = 8: .Text = GetResourceString(82) 'repayment Made
        .Col = 9: .Text = GetResourceString(344) 'Regular Interest
        .Col = 10: .Text = GetResourceString(345) 'Penal Interest
        .Col = 11: .Text = GetResourceString(58, 42)
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'Date
        .Col = 2: .Text = GetResourceString(58, 60) 'Loan No
        .Col = 3: .Text = GetResourceString(49, 60) 'Loan No
        .Col = 4: .Text = GetResourceString(35) 'Name
        .Col = 5: .Text = GetResourceString(269) 'Cash
        .Col = 6: .Text = GetResourceString(270) 'Contra
        .Col = 7: .Text = GetResourceString(269) 'Cash
        .Col = 8: .Text = GetResourceString(270) 'Contra
        .Col = 9: .Text = GetResourceString(344) 'Regular Interest
        .Col = 10: .Text = GetResourceString(345) 'Penal Interest
        .Col = 11: .Text = GetResourceString(58, 42) 'Loan Balance
        
        .ColAlignment(4) = 1
    End With
    GoTo ExitLine
End If
If m_ReportType = repLoanGLedger Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'Date
        .Col = 2: .Text = GetResourceString(272) 'LOan Issued
        .Col = 3: .Text = GetResourceString(271) 'repayment Made
        '.Col = 4: .Text = GetResourceString(344) 'Regular Interest
        '.Col = 5: .Text = GetResourceString(345) 'Penal Interest
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanGuarantor Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 6
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(389) & " 1"  'Guaranteers
        .Col = 5: .Text = GetResourceString(389) & " 2"  'Guaranteers
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanCustRP Then
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
        .Col = 2: .Text = GetResourceString(49) & " " & _
                            GetResourceString(60)   '"Member No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(272) 'Withdraw
        .Col = 5: .Text = GetResourceString(271) 'Deposit
        .Col = 6: .Text = GetResourceString(271) 'Deposit
        .Col = 7: .Text = GetResourceString(271) 'Deposit
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58) & " " & _
                            GetResourceString(60)   '"Loan No
        .Col = 2: .Text = GetResourceString(49) & " " & _
                            GetResourceString(60)   '"Mem No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(272) 'Withdraw
        .Col = 5: .Text = GetResourceString(310) 'Deposit
        .Col = 6: .Text = GetResourceString(344) 'Reg Interesr
        .Col = 7: .Text = GetResourceString(345) 'Penal INterest
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanIntCol Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 7
        .FixedCols = 2
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(37) 'Date
        .Col = 5: .Text = GetResourceString(344)   'Regular Interest
        .Col = 6: .Text = GetResourceString(345)   'Penal Interest
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanHolder Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 10
        .FixedCols = 3
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(111) 'Caste
        .Col = 5: .Text = GetResourceString(112) 'Place
        .Col = 6: .Text = GetResourceString(214) 'Loan Scheme
        .Col = 7: .Text = GetResourceString(340) 'Issue date
        .Col = 8: .Text = GetResourceString(80, 91) 'Loan Amount
        .Col = 9: .Text = GetResourceString(42) 'Loan Balance
        
    End With
    GoTo ExitLine
End If
    
'.FormatString = ">Loan Acc NO|<Loan holder Name|<Society Name|<Caste |<Place |" _
             & "<Loan Name|^Date |>Loan Amount |>Balance"
 
If m_ReportType = repLoanIssued Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 6
        .FixedCols = 3
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(340) 'Issue date
        .Col = 5: .Text = GetResourceString(80, 91) 'Loan Amount
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanInstOD Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 10
        .FixedCols = 1
        .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        '.Col = 3: .Text = GetResourceString(340) 'Issue date
        .Col = 4: .Text = GetResourceString(209) 'Due date
        .Col = 5: .Text = GetResourceString(42) 'Loan Balance
        .Col = 6: .Text = GetResourceString(84)   'Over Due AMount
        .Col = 7: .Text = GetResourceString(84)  'Reg Interesest
        .Col = 8: .Text = GetResourceString(84)  'Od Interesest
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        '.Col = 3: .Text = GetResourceString(340) 'Issue date
        .Col = 4: .Text = GetResourceString(209) 'Due date
        .Col = 5: .Text = GetResourceString(42) 'Loan Balance
        .Col = 6: .Text = GetResourceString(58)  'Over Due AMount
        .Col = 7: .Text = GetResourceString(344) ' & " " & GetResourceString(47) ' Interesest
        .Col = 8: .Text = GetResourceString(345) ' Interesest
        .MergeCells = flexMergeRestrictRows
        .MergeRow(1) = True
    End With
    GoTo ExitLine
End If
If m_ReportType = repLoanOD Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 10
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        '.Col = 3: .Text = GetResourceString(340) 'Issue date
        .Col = 4: .Text = GetResourceString(209) 'Due date
        .Col = 5: .Text = GetResourceString(42) 'Loan Balance
        .Col = 6: .Text = GetResourceString(84, 40) 'Over Due AMount
        '.Col = 6: .Text = GetResourceString(47) 'Od Interesest 'GetResourceString(84) & " " &
        .Col = 7: .Text = GetResourceString(344) 'Regular Interesest
        .Col = 8: .Text = GetResourceString(345) 'Od Interesest
        .Col = 9: .Text = GetResourceString(52) 'Total Amount
    End With
    GoTo ExitLine
End If
If m_ReportType = repLoanRepMade Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 9
        .FixedCols = 1
        .FixedRows = 2
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'date
        .Col = 2: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 4: .Text = GetResourceString(35) 'Name
        .Col = 5: .Text = GetResourceString(341) 'Reapid Amount
        .Col = 6: .Text = GetResourceString(341) 'Interest
        .Col = 7: .Text = GetResourceString(341)  'penal Interesest
        .Col = 8: .Text = GetResourceString(42) 'Loan Balance
        .Row = 1
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'date
        .Col = 2: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 4: .Text = GetResourceString(35) 'Name
        .Col = 5: .Text = GetResourceString(310) 'Princapal Amount
        .Col = 6: .Text = GetResourceString(344) 'Interest
        .Col = 7: .Text = GetResourceString(345)  'penal Interesest
        .Col = 8: .Text = GetResourceString(42) 'Loan Balance
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanSanction Or m_ReportType = repLoanBalance Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 5
        .FixedCols = 2
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Member No
        .Col = 3: .Text = GetResourceString(35) 'Name
        .Col = 4: .Text = GetResourceString(262) 'Sanction Loan
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 4
        .ColAlignment(3) = 1
        If m_ReportType = repLoanBalance Then .Text = GetResourceString(42) 'Balance
            
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanIntReceivable Or m_ReportType = repLoanIntReceivableTill Then
    With grd
        .Clear
        .Rows = 4
        .Cols = 6
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 2: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(35) 'Name
        If m_ReportType = repLoanIntReceivableTill Then
            .Col = 4: .Text = GetResourceString(42)
            .Col = 5: .Text = GetResourceString(47)
        Else
            .Col = 4: .Text = GetResourceString(376) & " " & _
                            GetResourceString(47)
            .Col = 5: .Text = GetResourceString(344)
        End If
    End With
    GoTo ExitLine
End If

If m_ReportType = repLoanCashBook Then
    With grd
        .Clear
        .Rows = 5
        .Cols = 11
        .FixedCols = 1
        .FixedRows = 1
        .Row = 0
        .Col = 0: .Text = GetResourceString(33) 'Sl No
        .Col = 1: .Text = GetResourceString(37) 'date
        .Col = 2: .Text = GetResourceString(58, 60)  '"Loan No
        .Col = 3: .Text = GetResourceString(49, 60)  '"Loan No
        .Col = 4: .Text = GetResourceString(35) 'Name
        .Col = 5: .Text = GetResourceString(41) 'Voucher
        .Col = 6: .Text = GetResourceString(58) 'Repaid Amount
        .Col = 7: .Text = GetResourceString(20) 'Repaid Amount
        .Col = 8: .Text = GetResourceString(344) 'Interest
        .Col = 9: .Text = GetResourceString(345)  'penal Interesest
        .Col = 10: .Text = GetResourceString(42) 'Loan Balance
        .ColAlignment(4) = 1
    End With
    GoTo ExitLine
End If

ExitLine:

Dim RowCount As Integer
Dim ColCount As Integer
Dim blMerge As Boolean

With grd

If .FixedRows > 1 Then blMerge = True
If blMerge Then .MergeCells = flexMergeFree
    For RowCount = 0 To .FixedRows - 1
        .Row = RowCount
        .MergeRow(RowCount) = blMerge
        For ColCount = 0 To .Cols - 1
            .Col = ColCount
            .MergeCol(ColCount) = blMerge
            .CellAlignment = 4
            .CellFontBold = True
        Next
    Next
End With

End Sub

Private Sub LoansAndPayments()
'm_ReportType = repLoanDailyCash

'Declare the varaibles
Dim rst As Recordset
Dim SqlStr As String

Dim fromDate As String
Dim toDate As String

RaiseEvent Processing("Reading & Verifying the records ", 0)

SqlStr = "Select 'Principal', A.LoanId,C.CustomerID,AccNum,val(AccNum) as AcNum,TransID," & _
    " Name as CustName, Transtype,TransDate,Amount,Particulars FROM LoanTrans A, " & _
    " LoanMaster B, QryName C WHERE B.LoanId = A.LoanId " & _
    " AND TransDate >= #" & m_FromDate & "# TransDate <= #" & m_ToDate & "#" & _
    " AND C.CustomerID = B.CustomerID"
    
    If m_Place <> "" Then SqlStr = SqlStr & "  AND Place = " & AddQuotes(m_Place, True)
    If m_Caste <> "" Then SqlStr = SqlStr & "  AND caste = " & AddQuotes(m_Caste, True)
    If m_FromAmt <> 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
    If m_ToAmt <> 0 Then SqlStr = SqlStr & "   AND Amount <= " & m_ToAmt
    If m_SchemeId Then SqlStr = SqlStr & " AND B.SchemeID = " & m_SchemeId

SqlStr = SqlStr & " UNION " & _
    " SELECT 'Interest', A.LoanId, CustomerID, AccNum,val(AccNum) as AcNum, TransID, " & _
    " Name as CustName, Transtype,TransDate,IntAmount as Amount,PenalIntAmount as Particulars " & _
    " FROM LoanTrans A,LoanMaster B,QryName C WHERE B.LoanId = A.LoanId " & _
    " AND TransDate >= #" & m_FromDate & "# TransDate <= #" & m_ToDate & "#" & _
    " AND C.CustomerID = B.CustomerID"

    If m_Place <> "" Then SqlStr = SqlStr & "  AND Place = " & AddQuotes(m_Place, True)
    If m_Caste <> "" Then SqlStr = SqlStr & "  AND caste = " & AddQuotes(m_Caste, True)
    If m_FromAmt <> 0 Then SqlStr = SqlStr & " AND Amount >= " & m_FromAmt
    If m_ToAmt <> 0 Then SqlStr = SqlStr & "   AND Amount <= " & m_ToAmt
    If m_SchemeId Then SqlStr = SqlStr & "     AND B.SchemeID = " & m_SchemeId
    
SqlStr = SqlStr & " ORDER By TransDate,AcNum,TransID,"

If gDbTrans.Fetch(rst, adOpenStatic) <= 0 Then
    Call PrintNoRecords(grd)
    Exit Sub
End If


Call InitGrid
Dim totalCount As Long
Dim loopCount As Integer
totalCount = rst.RecordCount + 2

RaiseEvent Initialise(0, totalCount)
RaiseEvent Processing("Aligning the data which is is being written to the grid.", 0)

Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim TransID As Long
Dim LoanID As Long
Dim Amount As Currency
Dim IntAmount As Currency
Dim SlNo As Long

Dim SubTotal(4 To 9) As Currency
Dim GrandTotal(4 To 9) As Currency
Dim PRINTTotal As Boolean

TransDate = rst("TransDate")
grd.Row = 0: TransID = 0
Dim I As Integer

While Not rst.EOF
    If TransDate <> rst("TransDate") Then
        PRINTTotal = True
        With grd
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 2: .Text = GetResourceString(304)
            .CellAlignment = 4: .CellFontBold = True
            For I = 4 To 9
                .Col = I: .Text = FormatCurrency(SubTotal(I))
                .CellAlignment = 7: .CellFontBold = True
                GrandTotal(I) = GrandTotal(I) = SubTotal(I)
                SubTotal(I) = 0
            Next
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            .Row = .Row + 1
        End With
    End If
    If LoanID = rst.Fields("LoanID") Then TransID = 0
    LoanID = rst.Fields("LOanID")
    If TransID <> FormatField(rst("TransID")) Then 'Set new row
        TransID = FormatField(rst("TransID"))
        With grd
            'Write all information
            If .Rows = .Row + 1 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 0: .Text = Format(SlNo, "00")
            .Col = 1: .Text = FormatField(rst("AccNum"))
            .Col = 2: .Text = FormatField(rst("CustName"))
            .Col = 3: .Text = GetIndianDate(TransDate)
        End With
    End If
    transType = rst("TransType")
    Amount = FormatField(rst("Amount"))
    
    With grd
        If FormatField(rst(0)) = "Principal" Then
            If transType = wDeposit Then .Col = 6
            If transType = wContraDeposit Then .Col = 7
            If transType = wWithdraw Then .Col = 4
            If transType = wContraWithdraw Then .Col = 5
            .Text = FormatCurrency(Amount): .CellAlignment = 7
        Else
            If transType = wWithdraw Or transType = wContraWithdraw Then
                .Col = 10
            Else
                .Col = 8
                .Text = FormatCurrency(Amount): .CellAlignment = 7
                .CellAlignment = 7
                SubTotal(.Col) = SubTotal(.Col) + Val(.Text)
                .Col = 9
                .Text = FormatField(rst("Particulars")): .CellAlignment = 7
            End If
        End If
        SubTotal(.Col) = SubTotal(.Col) + Val(.Text)
    End With
    
nextRecord:
    
    loopCount = loopCount + 1
    RaiseEvent Processing("Writing the data to the grid.", loopCount / totalCount)
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    rst.MoveNext

Wend

'Enter last totals
    With grd
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
            .Col = 2: .Text = GetResourceString(304)
        .CellAlignment = 4: .CellFontBold = True
        For I = 4 To 9
            .Col = I: .Text = FormatCurrency(SubTotal(I))
            .CellAlignment = 7
            .CellFontBold = True
            GrandTotal(I) = GrandTotal(I) = SubTotal(I)
            SubTotal(I) = 0
        Next
    End With

If PRINTTotal Then
    With grd
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 2: .Text = GetResourceString(286)
        .CellAlignment = 4: .CellFontBold = True
        For I = 4 To 9
            .Col = I: .Text = FormatCurrency(GrandTotal(I))
            .CellAlignment = 7
            .CellFontBold = True
        Next
    End With
End If

End Sub

Private Function ReportInterestRecieved() As Boolean
 
 ReportInterestRecieved = False

' Declare variables...
Dim Lret As Long
Dim SqlStr As String
Dim rptRS As Recordset
Dim PrevMemberID As Long
Dim PrevLoanID As Long
Dim TotalPenalInt As Currency
Dim TotalRegInt As Currency


' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
SqlStr = "SELECT AccNum, A.LoanID, A.TransDate,B.Customerid, A.TransType," & _
        " Name as CustName, A.IntAmount, A.PenalIntAmount,MemberNum " & _
        " FROM LoanIntTrans A, LoanMaster B, QryMemName C " & _
        " WHERE A.LoanId = B.Loanid AND C.MemID = B.MemID " & _
        " AND (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")" & _
        " AND a.Transdate >= #" & m_FromDate & "#" & _
        " AND A.TransDate <= #" & m_ToDate & "#"

' Add the WHERE clause restrictions, if specified.
If m_SchemeId Then SqlStr = SqlStr & " AND B.SchemeID = " & m_SchemeId

If Trim$(m_Place) <> "" Then SqlStr = SqlStr & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStr = SqlStr & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID

' Finally, add the sorting clause.
SqlStr = SqlStr & " ORDER BY A.TransDate, SchemeID"
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " , val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " , IsciName"
End If

' Execute the query...

Lret = gDbTrans.Fetch(rptRS, adOpenStatic)
If Lret < 0 Then
    Call PrintNoRecords(grd)
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    Call PrintNoRecords(grd)
    MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If


' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
Dim totalCount As Long
totalCount = rptRS.RecordCount + 2
RaiseEvent Initialise(0, totalCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim SlNo As Long
' Fill the rows

With grd
    .Rows = 20
    .Row = .FixedRows - 1
    .Visible = False
    Do While Not rptRS.EOF
        ' Set the row.
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        ' Fill the loan id.
        .Col = 0
        .Text = Format(SlNo, "00")
        .Col = 1
        .Text = FormatField(rptRS("AccNum"))
    ' Fill the loan holder name.
        .Col = 2: .Text = FormatField(rptRS("MemberNum"))
        .Col = 3: .Text = FormatField(rptRS("CustName"))
        ' Fill the transaction date.
        .Col = 4
        .Text = FormatField(rptRS("TransDate"))

        .Col = 5: .Text = FormatField(rptRS("IntAmount"))
        TotalRegInt = TotalRegInt + Val(.Text): .CellAlignment = 7
        .Col = 6: .Text = FormatField(rptRS("PenalIntAmount"))
        TotalPenalInt = TotalPenalInt + Val(.Text): .CellAlignment = 7
    
nextRecord:
        
        RaiseEvent Processing("Writing the data to the grid.", SlNo / totalCount)
        DoEvents
        If gCancel Then rptRS.MoveLast
        Me.Refresh
        rptRS.MoveNext
    Loop

    lblReportTitle.Caption = GetResourceString(485) & " " & _
                    GetFromDateString(m_FromIndianDate, m_ToIndianDate)

    If .Rows <= .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .Col = 1: .Text = "Grand Total"
    .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(TotalRegInt)
    .CellFontBold = True: .CellAlignment = 7
    .Col = 6: .Text = FormatCurrency(TotalPenalInt)
    .CellFontBold = True: .CellAlignment = 7
    
End With

ReportInterestRecieved = True


Exit_Line:
    grd.Visible = True
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    Err.Clear
    GoTo Exit_Line
End Function

Public Property Let SchemeID(NewID As Integer)
    m_SchemeId = NewID
End Property

Private Sub cmdOk_Click()

Unload Me
End Sub


Private Sub cmdPrint_Click()
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    Set m_frmCancel = New frmCancel
    Load m_frmCancel
    
    With m_frmCancel
        .PicStatus.Visible = True
        .Show
    End With
        
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .ReportTitle = Me.lblReportTitle.Caption
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
Dim SqlStmt As String
Dim rst As Recordset

Call CenterMe(Me)
Call SetKannadaCaption

'Init the grid
    grd.Clear
    grd.Rows = 20
    grd.Cols = 1
    grd.FixedCols = 0
    grd.Row = 1
    Dim SchemeID As Integer
    grd.Text = "No Records Available"
''Set the SQL condition
m_SqlCondition = ""
If Trim$(m_Place) <> "" Then m_SqlCondition = m_SqlCondition & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then m_SqlCondition = m_SqlCondition & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then m_SqlCondition = m_SqlCondition & " And Gender = " & m_Gender
'If m_AccGroupID Then m_SqlCondition = m_SqlCondition & " And AccGroupID = " & m_AccGroupID

    
    
    Dim RetBool As Boolean
    If m_ReportType = repLoanBalance Then RetBool = ReportLoanBalance
    If m_ReportType = repLoanDailyCash Then RetBool = ReportSubDayBook
    If m_ReportType = repLoanGLedger Then RetBool = ReportGeneralLedger
    If m_ReportType = repLoanGuarantor Then RetBool = ReportGuarantors
    If m_ReportType = repLoanHolder Then RetBool = ReportLoanHolders
    If m_ReportType = repLoanInstOD Then RetBool = ReportOverdueInstalments
    If m_ReportType = repLoanIntCol Then RetBool = ReportInterestRecieved
    If m_ReportType = repLoanIssued Then RetBool = ReportLoanIssued
    If m_ReportType = repLoanOD Then RetBool = ReportOverdueLoans
    If m_ReportType = repLoanRepMade Then RetBool = ReportRepaymentsMade
    If m_ReportType = repLoanSanction Then RetBool = ReportSanctionedLoans
    If m_ReportType = repLoanCashBook Then RetBool = ReportSubCashBook
    If m_ReportType = repLoanIntReceivable Then RetBool = ReportInterestReceivable
    If m_ReportType = repLoanIntReceivableTill Then RetBool = ReportInterestCalculate
    If m_ReportType = repLoanCustRP Then RetBool = ReportCustomerTransaction
    If m_ReportType = repLoanReceivable Then RetBool = ReportReceivables
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
Exit Sub
     Screen.MousePointer = vbDefault
End Sub

Private Function ReportReceivables() As Boolean
On Error GoTo Err_line
'Now Get the Head ID Of the Bkcc
Dim AccHeadID As Long
Dim rstMain As Recordset
If m_SchemeId = 0 Then Exit Function
gDbTrans.SqlStmt = "Select SchemeName From " & _
            " LoanScheme Where SchemeID = " & m_SchemeId
If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then gCancel = 2: Exit Function

AccHeadID = GetHeadID(FormatField(rstMain("SchemeName")), parMemberLoan)
If AccHeadID = 0 Then gCancel = 2: Exit Function

'NOw fetch the details from From
gDbTrans.SqlStmt = "Select A.*,B.HeadName From AmountReceivable A,Heads B " & _
            " Where AccHeadID = " & AccHeadID & _
            " And B.HeadID = A.DueHeadID" & _
            " And (TransType = " & wWithdraw & _
                " OR TransType = " & wContraWithdraw & ")" & _
            " AND AccTransID >= (Select Max(TransID) " & _
                " From LoanTrans C Where C.LoanID=A.AccID " & _
                " And C.TransDate <= A.TransDate )"

If gDbTrans.Fetch(rstMain, adOpenDynamic) < 1 Then Exit Function
'Dim rstHead As Recordset
'Dim DueHeadId As Long
Dim rstName As Recordset

        
gDbTrans.SqlStmt = "Select AccNum,MemberNum,LoanId,A.CustomerId,Name as CustName" & _
        " From LoanMaster A,QryMemName B " & _
        " Where A.MemID= b.MemID " & _
        " And LoanID in (Select Distinct AccID as LoanID " & _
            " From AmountReceivAble Where AccHeadID = " & AccHeadID & ")"
        
Call gDbTrans.Fetch(rstName, adOpenDynamic)


With grd
    .Clear
    .Rows = 10
    .Cols = 6
    .FixedCols = 1
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36, 60) '"AccNum"
    .Col = 2: .Text = GetResourceString(36, 60) '"AccNum"
    .Col = 3: .Text = GetResourceString(35)  '"Name"
    .Col = 4: .Text = GetResourceString(36, 35) '"HeadNAme"
    .Col = 5: .Text = GetResourceString(40)  '"Amount"
End With

While Not rstMain.EOF
    With grd
        .Row = .Row + 1
        .Col = 0: .Text = .Row
        rstName.MoveFirst
        rstName.Find "LoanID = " & rstMain("AccID")
        If Not rstName.EOF Then
            .Col = 1: .Text = FormatField(rstName("AccNum"))
            .Col = 2: .Text = FormatField(rstName("MemberNum"))
            .Col = 3: .Text = FormatField(rstName("CustNAme"))
        End If
        .Col = 4: .Text = FormatField(rstMain("HeadNAme"))
        .Col = 5: .Text = FormatField(rstMain("Amount"))
    End With
    rstMain.MoveNext
Wend

ReportReceivables = True
Err_line:
Err.Clear
End Function
Private Function ReportRepaymentsMade() As Boolean
 
m_ReportType = repLoanRepMade

ReportRepaymentsMade = False

' Declare variables...
Dim LoanID As Long
Dim TransID As Long

Dim PrinSql As String
Dim IntSql As String
' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
PrinSql = "SELECT 'PRINCIPAL',SchemeName,MemberNum,A.loanID,A.TransDate,TransID," _
        & " name as CustName, AccNum,val(AccNum) as AcNum,TransType,Amount as PrinAmount,Balance as Bal" _
        & " FROM LoanTrans A, LoanMaster B, QryMemName C,LoanScheme D " _
        & " WHERE B.Loanid = A.loanid AND C.MemID = B.MemID " _
        & " And D.SchemeId = B.SchemeId " _
        & " And (TransType = " & wDeposit & " or TransType = " & wContraDeposit & ")" _
        & " AND A.Transdate >= #" & m_FromDate & "#" _
        & " AND A.Transdate <= #" & m_ToDate & "#"
        
IntSql = "SELECT 'INTEREST',SchemeName,MemberNum,A.loanID,A.TransDate,TransID, " _
        & " Name as CustName," _
        & " AccNum,val(AccNum) as AcNum,TransType,IntAmount as PrinAmount,PenalIntAmount as Bal " _
        & " FROM LoanIntTrans A, LoanMaster B, QryMemName C,LoanScheme D " _
        & " WHERE B.Loanid = A.loanid AND C.MemID = B.MemID " _
        & " And D.SchemeId = B.SchemeId " _
        & " And (TransType = " & wDeposit & " or TransType = " & wContraDeposit & ")" _
        & " AND A.Transdate >= #" & m_FromDate & "#" _
        & " AND A.Transdate <= #" & m_ToDate & "#"

' Add the WHERE clause restrictions, if specified.
If m_SchemeId Then
    PrinSql = PrinSql & " AND B.SchemeID = " & m_SchemeId
    IntSql = IntSql & " AND B.SchemeID = " & m_SchemeId
End If
If m_FromAmt > 0 Then PrinSql = PrinSql & " AND A.Amount >= " & m_FromAmt
If m_ToAmt > 0 Then PrinSql = PrinSql & " AND A.Amount <= " & m_ToAmt

If Trim$(m_Place) <> "" Then PrinSql = PrinSql & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then PrinSql = PrinSql & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then PrinSql = PrinSql & " And Gender = " & m_Gender
If m_AccGroupID Then PrinSql = PrinSql & " And AccGroupID = " & m_AccGroupID

If Trim$(m_Place) <> "" Then IntSql = IntSql & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then IntSql = IntSql & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then IntSql = IntSql & " And Gender = " & m_Gender
If m_AccGroupID Then IntSql = IntSql & " And AccGroupID = " & m_AccGroupID

' Finally, add the sorting clause.
Dim rst As Recordset
gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & _
                    " ORDER BY TransDate,SchemeName,AcNum"
    
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function
PrinSql = "": IntSql = ""
    
Dim TotalBalance As Currency

' Populate the record set.
rst.MoveLast
rst.MoveFirst

' Initialize the grid.
Dim totalCount As Long
totalCount = rst.RecordCount + 2
RaiseEvent Initialise(0, totalCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim SlNo As Long
Dim TransDate As Date
Dim SubDeposit As Double
Dim subInt As Double
Dim SubPenalInt As Double
Dim TotalDeposit As Double
Dim TotalInt  As Double
Dim TotalPenalInt As Double
Dim transType As wisTransactionTypes

Dim Amount As Double
Dim Balance As Double
Dim PRINTTotal As Boolean

' Fill the rows
SlNo = 1
TransDate = rst("TransDate")
grd.Row = grd.FixedRows - 1

Do While Not rst.EOF
    ' Set the row.
    With grd
        transType = FormatField(rst("TransType"))
        If TransDate <> rst("TransDate") Then
            PRINTTotal = True
            ' Set the row.
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            transType = FormatField(rst("TransType"))
            .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
            TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
            .Col = 6: .Text = FormatCurrency(subInt): .CellFontBold = True
            TotalInt = TotalInt + subInt: subInt = 0
            .Col = 7: .Text = FormatCurrency(SubPenalInt): .CellFontBold = True
            TotalPenalInt = TotalPenalInt + SubPenalInt: SubPenalInt = 0
            TransDate = rst("TransDate")
        End If
        
        If LoanID <> FormatField(rst("LoanId")) Then
            LoanID = FormatField(rst("LoanId"))
            TransID = 0
        End If
        
        'Increase the Row By One set
        If TransID <> FormatField(rst("TransID")) Then
            TransID = FormatField(rst("TransID"))
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            ' Fill the loan id.
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = GetIndianDate(TransDate)
            .Col = 2: .Text = FormatField(rst("AccNum"))
            .Col = 3: .Text = FormatField(rst("MemberNum"))
            ' Fill the loan holder name.
            .Col = 4: .Text = FormatField(rst("Custname"))
            SlNo = SlNo + 1
        End If
        
        transType = rst("TransType")
        Amount = FormatField(rst("PrinAmount"))
        Balance = FormatField(rst("bal"))
        If rst(0) = "PRINCIPAL" Then
            .Col = 5: .Text = FormatCurrency(Amount)
            SubDeposit = SubDeposit + Amount
            ' Fill the balance amount.
            .Col = 6: .Text = FormatCurrency(Balance)
        Else
            .Col = 6: .Text = FormatCurrency(Amount)
            subInt = subInt + Amount
            ' Fill the balance amount.
            If Balance Then
                .Col = 7: .Text = FormatCurrency(Balance)
                SubPenalInt = SubPenalInt + Balance
            End If
        End If
        ' Move to next row.
    End With
    

nextRecord:
    DoEvents
    If gCancel Then rst.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", SlNo / totalCount)
    rst.MoveNext
Loop

With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
    TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
    .Col = 6: .Text = FormatCurrency(subInt): .CellFontBold = True
    TotalInt = TotalInt + subInt: subInt = 0
    .Col = 7: .Text = FormatCurrency(SubPenalInt): .CellFontBold = True
    TotalPenalInt = TotalPenalInt + SubPenalInt: SubPenalInt = 0
    
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 4: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalDeposit): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalInt): .CellFontBold = True
        .Col = 7: .Text = FormatCurrency(TotalPenalInt): .CellFontBold = True
        If TotalPenalInt = 0 Then .ColWidth(6) = 5
    End If
    ' Display the grid.
    .Visible = True
End With

'set the tile
lblReportTitle.Caption = GetResourceString(76) & " " & _
        GetFromDateString(m_FromIndianDate, m_ToIndianDate)

Me.Caption = "LOANS [Repayments made...]"

ReportRepaymentsMade = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        
    End If
'Resume
    GoTo Exit_Line

End Function


Private Function ReportSubDayBook() As Boolean
 
 ReportSubDayBook = False
' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim PrevMemberID As Long
Dim PrevLoanID As Long
Dim PrinSql As String
Dim IntSql As String
Dim I As Integer

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
'SQl to fetch the Principal
PrinSql = "SELECT 'PRINCIPAL',SchemeName, MemberNum, B.CustomerID, A.LoanID, TransID, " _
        & " TransType, TransDate, AccNum,val(AccNum) as AcNum, A.Amount, A.Balance,Name as CustName " _
        & " FROM LoanTrans a, LoanMaster B, LoanScheme C,QryMemName D  WHERE " _
        & " B.loanid = A.loanid" _
        & " AND D.MemID = B.MemID AND C.SchemeID = B.SchemeId" _
        & " AND A.TransDate >= #" & m_FromDate & "#" _
        & " AND A.TransDate <= #" & m_ToDate & "#"
If m_Caste <> "" Then PrinSql = PrinSql & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then PrinSql = PrinSql & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then PrinSql = PrinSql & " And Gender = " & m_Gender
If m_AccGroupID Then PrinSql = PrinSql & " And AccGroupID = " & m_AccGroupID

'Sql to fetch the Interest
IntSql = "SELECT 'INTEREST',SchemeName, MemberNum, B.CustomerID, A.LoanID, TransID, " _
        & " TransType, TransDate, AccNum,val(AccNum) as AcNum,A.IntAmount, A.PenalIntAmount," _
        & " Name as CustName " _
        & " FROM LoanIntTrans A, LoanMaster B, LoanScheme C, QryMemName D WHERE " _
        & " B.loanid = A.loanid" _
        & " AND D.MemID = B.MemID AND C.SchemeID = B.SchemeId" _
        & " AND A.TransDate >= #" & m_FromDate & "#" _
        & " AND A.TransDate <= #" & m_ToDate & "#"
If m_Caste <> "" Then IntSql = IntSql & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then IntSql = IntSql & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then IntSql = IntSql & " And Gender = " & m_Gender
If m_AccGroupID Then IntSql = IntSql & " And AccGroupID = " & m_AccGroupID


' Add the WHERE clause restrictions, if specified.
If m_SchemeId Then
    PrinSql = PrinSql & " AND B.SchemeID = " & m_SchemeId
    IntSql = IntSql & " AND B.SchemeID = " & m_SchemeId
End If
If m_FromAmt > 0 Then PrinSql = PrinSql & " AND A.Amount >= " & m_FromAmt
If m_ToAmt > 0 Then PrinSql = PrinSql & " AND A.Amount <= " & m_ToAmt

' Finally, add the Pricipal & interest sql and add sorting clause.
gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & _
    " ORDER BY TransDate, SchemeName,AcNum"
    
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenStatic)
If Lret <= 0 Then GoTo Exit_Line


' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.RecordCount + 1
    If .Rows < 50 Then .Rows = 50
    .FixedRows = 1
    .FormatString = ">SlNo |<Loan Acc NO |<Loan holder name|<Loan Name |>Date   |>Loan advence|" _
                & "Loan recovery| Interest |Penal Interest| Balance"
    .FixedCols = 1
End With

Lret = rptRS.RecordCount + 2
RaiseEvent Initialise(0, Lret)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

Dim transType As wisTransactionTypes
Dim SlNo As Long
Dim TransDate As String
Dim SubAmount(5 To 10)
Dim TotalAmount(5 To 10)
Dim TransID As Long
Dim LoanID As Long
Dim TotalPrint As Boolean
Dim loopCount As Integer

' Fill the rows
SlNo = 0: grd.Row = 0
grd.Rows = 40
grd.Row = grd.FixedRows - 1
TransDate = FormatField(rptRS.Fields("TransDate"))
Do While Not rptRS.EOF
  With grd
    If LoanID <> rptRS("LoanID") Then TransID = 0
    LoanID = rptRS("LOanID")
    ' Set the row.
    If TransDate <> FormatField(rptRS("TransDate")) Then
        TotalPrint = True
        TransDate = FormatField(rptRS("TransDate"))
        SlNo = 0: TransID = 0
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
        For I = 5 To 10
            .Col = I
            If SubAmount(I) Then .Text = FormatCurrency(SubAmount(I))
            TotalAmount(I) = TotalAmount(I) + SubAmount(I)
            SubAmount(I) = 0
            .CellFontBold = True
        Next
    End If
    transType = rptRS("TransType")
    'Fill the loanid.
    If TransID <> rptRS("TransID") Then
        TransID = rptRS("TransID"): SlNo = SlNo + 1
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = TransDate
        TransDate = .Text
        .Col = 2: .Text = rptRS("AccNum")
        ' Fill the loan  name.
    End If
    If rptRS(0) = "INTEREST" Then
        If transType = wDeposit Then
            .Col = 9: .Text = FormatField(rptRS("Amount"))
            SubAmount(.Col) = SubAmount(.Col) + Val(.Text): .CellAlignment = 7
            .Col = 10: .Text = FormatField(rptRS("Balance")): .CellAlignment = 7
            SubAmount(.Col) = SubAmount(.Col) + Val(.Text)
        End If
    Else
        If transType = wWithdraw Or transType = wContraWithdraw Then
            .Col = IIf(transType = wWithdraw, 5, 6)
            .Text = FormatField(rptRS("Amount")): .CellAlignment = 7
            SubAmount(.Col) = SubAmount(.Col) + Val(.Text)
        Else
            .Col = IIf(transType = wDeposit, 7, 8)
            .Text = FormatField(rptRS("Amount")): .CellAlignment = 7
            SubAmount(.Col) = SubAmount(.Col) + Val(.Text)
        End If
        ' Fill the loan  name.
        .Col = 11: .Text = FormatField(rptRS("Balance")): .CellAlignment = 7
    End If
    .Col = 3: .Text = FormatField(rptRS("MemberNum"))
    .Col = 4: .Text = FormatField(rptRS("CustName"))
    
nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    loopCount = loopCount + 1
    RaiseEvent Processing("Writing the data to the grid.", loopCount / Lret)
    rptRS.MoveNext
  
  End With
Loop
    
With grd
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 1: .Text = GetResourceString(304): .CellFontBold = True
    If .Rows <= .Row + 2 Then .Rows = .Row + 2
    .Row = .Row + 1
    .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
    For I = 5 To 10
        .Col = I
        If SubAmount(I) Then .Text = FormatCurrency(SubAmount(I))
        TotalAmount(I) = TotalAmount(I) + SubAmount(I)
        SubAmount(I) = 0
        .CellFontBold = True
    Next
            
    If TotalPrint Then
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        .Col = 4: .Text = GetResourceString(286): .CellFontBold = True
        For I = 5 To 10
            .Col = I
            If TotalAmount(I) Then .Text = FormatCurrency(TotalAmount(I))
            .CellFontBold = True
        Next
    End If
End With

'set the title
lblReportTitle.Caption = GetResourceString(390) & " " & _
    GetResourceString(63) & " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)

' Display the grid.
grd.Visible = True

Me.Caption = "Loans  [Sub day book...]"

ReportSubDayBook = True

'Call grd_LostFocus

Exit_Line:
    Exit Function
    
Err_line:
    If Err Then
        MsgBox "ReportTransactionMade: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        
    End If
'Resume
    GoTo Exit_Line

End Function

Private Function ReportCustomerTransaction() As Boolean

ReportCustomerTransaction = False

' Declare variables...
Dim Lret As Long
Dim rstTrans As Recordset
Dim RstCust As Recordset

Dim PrinSql  As String
Dim IntSql As String
Dim sqlSupport As String

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
sqlSupport = ""
If m_SchemeId Then _
    sqlSupport = sqlSupport & " and SchemeID = " & m_SchemeId
If m_AccGroupID Then _
    sqlSupport = sqlSupport & " And AccGroupID = " & m_AccGroupID

If Len(sqlSupport) Then
    sqlSupport = " And LOanID In (Select LoanID From LoanMAster " & _
        "WHERE " & Mid(Trim(sqlSupport), Len("and") + 1) & ") "
End If


' Build the report query.
PrinSql = "Select 'PRINCIPAL',sum(Amount) as RegInt,'0' as PenalInt, " & _
        "LoanID,TransType From LoanTrans Where " & _
        "TransDate >= #" & m_FromDate & "# And TransDate <= #" & m_ToDate & "# "

If m_FromAmt > 0 Then PrinSql = PrinSql & " AND Amount >= " & m_FromAmt
If m_ToAmt > 0 Then PrinSql = PrinSql & " AND Amount <= " & m_ToAmt
PrinSql = PrinSql & sqlSupport & " Group by LoanID,TransType "

IntSql = "Select 'INTEREST',sum(IntAmount) as RegInt,sum(PenalIntAmount) as PenalInt, " & _
        "LoanID,TransType From LoanIntTrans Where " & _
        "TransDate >= #" & m_FromDate & "# And TransDate <= #" & m_ToDate & "# " & _
        sqlSupport & " Group by LoanID,TransType "

' Finally, add the sorting clause.
gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & _
    " Order by LoanID"
    
' Execute the query...
Lret = gDbTrans.Fetch(rstTrans, adOpenStatic)
If Lret <= 0 Then
    Call PrintNoRecords(grd)
    GoTo Exit_Line
End If


sqlSupport = ""
If Trim$(m_Place) <> "" Then _
    sqlSupport = sqlSupport & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then _
    sqlSupport = sqlSupport & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then _
    sqlSupport = sqlSupport & " And Gender = " & m_Gender


grd.Clear
grd.Cols = 5
grd.FixedCols = 1
grd.Rows = 20


Dim TotalBalance As Currency

Lret = rstTrans.RecordCount + 2
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


'SlNo = 1
LoanID = rstTrans("LoanID")
'.Row = 0
Do While Not rstTrans.EOF
    With grd
        If LoanID <> rstTrans("LoanID") Then
            SlNo = SlNo + 1
            If .Rows <= .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            gDbTrans.SqlStmt = "Select Name as CustName,MemberNum,AccNum " & _
                "From LoanMaster A, QryMemName B where A.MemID = B.MemID " & _
                " and LoanId = " & LoanID
            Call gDbTrans.Fetch(RstCust, adOpenDynamic)
            
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = RstCust("AccNum")
            ' Fill the loan holder name.
            .Col = 2: .Text = Trim$(FormatField(RstCust("MemberNum")))
            .Col = 3: .Text = Trim$(FormatField(RstCust("CustName")))
            If SubWithdraw Then .Col = 4: .Text = FormatCurrency(SubWithdraw)
            If SubDeposit Then .Col = 5: .Text = FormatCurrency(SubDeposit)
            If SubInterest Then .Col = 6: .Text = FormatCurrency(SubInterest)
            If SubPenal Then .Col = 7: .Text = FormatCurrency(SubPenal)
            
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
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        gDbTrans.SqlStmt = "Select Name as CustName,MemberNum,AccNum " & _
            "From LoanMaster A Inner Join QryMemName B on B.MemID = A.MemID " & _
            "Where LoanId = " & LoanID
        If gDbTrans.Fetch(RstCust, adOpenDynamic) > 0 Then
        
        
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = RstCust("AccNum")
        ' Fill the loan holder name.
        .Col = 2: .Text = Trim$(FormatField(RstCust("MemberNum")))
        .Col = 3: .Text = Trim$(FormatField(RstCust("CustName")))
        End If
        If SubWithdraw Then .Col = 4: .Text = FormatCurrency(SubWithdraw)
        If SubDeposit Then .Col = 5: .Text = FormatCurrency(SubDeposit)
        If SubInterest Then .Col = 6: .Text = FormatCurrency(SubInterest)
        If SubPenal Then .Col = 7: .Text = FormatCurrency(SubPenal)
        
        TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
        TotalInterest = TotalInterest + SubInterest: SubInterest = 0
        TotalPenal = TotalPenal + SubPenal: SubPenal = 0
            
        'Put GrandTotal
        .Rows = .Row + 3
        .Row = .Row + 2
        .Col = 3: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalDeposit): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalInterest): .CellFontBold = True
        .Col = 7: .Text = FormatCurrency(TotalPenal): .CellFontBold = True
                   
        If TotalInterest = 0 Then .ColWidth(6) = 5
        If TotalPenal = 0 Then .ColWidth(7) = 5
        
        'Added on 29/11/00(For Allignment)
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
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

'Call grd_LostFocus

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

Private Function ReportSubCashBook() As Boolean

ReportSubCashBook = False

' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim PrinSql  As String
Dim IntSql As String


' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
PrinSql = "SELECT 'PRINCIPAL',B.CustomerID,AccNum,val(AccNum) as AcNum,MemberNum, A.LoanID,TransType,TransID," _
    & "TransDate,SchemeName, Amount As IntAmount, Balance as PenalBalance,VoucherNo, " _
    & " Name as CustName" _
    & " FROM LoanTrans A, LoanMaster B, LoanScheme C,QryMemName D WHERE " _
    & " B.LoanID = A.Loanid " _
    & " AND C.Schemeid = B.schemeid AND D.MemID = B.MemID" _
    & " AND A.TransDate >= #" & m_FromDate & "#" _
    & " AND A.TransDate <= #" & m_ToDate & "#"
    
IntSql = " SELECT 'INTEREST',B.CustomerID, AccNum,val(AccNum) as AcNum,MemberNum,A.LoanID,TransType,TransID," _
    & " TransDate,SchemeName,IntAmount, PenalIntAmount as PenalBalance,VoucherNo," _
    & " Name as CustName FROM LoanIntTrans A,LoanMaster B,LoanScheme C,QryMemName D " _
    & " WHERE B.LoanID = A.Loanid " _
    & " AND C.Schemeid = B.Schemeid AND D.MemID = B.MemID" _
    & " AND A.TransDate >= #" & m_FromDate & "#" _
    & " AND A.TransDate <= #" & m_ToDate & "#"
    
If m_SchemeId Then
    PrinSql = PrinSql & " and C.SchemeID = " & m_SchemeId
    IntSql = IntSql & " and C.SchemeID = " & m_SchemeId
End If


If m_FromAmt > 0 Then PrinSql = PrinSql & " AND A.amount >= " & m_FromAmt
If m_ToAmt > 0 Then PrinSql = PrinSql & " AND a.amount <= " & m_ToAmt

If Trim$(m_Place) <> "" Then PrinSql = PrinSql & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then PrinSql = PrinSql & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then PrinSql = PrinSql & " And Gender = " & m_Gender
If m_AccGroupID Then PrinSql = PrinSql & " And AccGroupID = " & m_AccGroupID

'Interest query
If Trim$(m_Place) <> "" Then IntSql = IntSql & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then IntSql = IntSql & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then IntSql = IntSql & " And Gender = " & m_Gender
If m_AccGroupID Then IntSql = IntSql & " And AccGroupID = " & m_AccGroupID

gDbTrans.SqlStmt = PrinSql & " Order by A.TransDate,B.SchemeID,Val(AccNum)" 'from loanTrans
    
Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

Dim TotalBalance As Currency
    ' InitGird
    grd.Clear
    grd.Cols = 6
    grd.FixedCols = 1
    grd.Rows = 20

' Finally, add the sorting clause.
gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & _
        " ORDER BY TransDate, SchemeName,loanid"
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenStatic)
If Lret <= 0 Then
    MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.RecordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

Call InitGrid

If m_SchemeId Then grd.ColWidth(3) = 5
'If RepSocDet = DirectLoans Then grd.ColWidth(2) = 5

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
grd.Rows = 40
grd.Row = 0
Dim L_clsCust As New clsCustReg
Dim LoanID As Long
TransDate = rptRS("TransDate")

'SlNo = 1
Do While Not rptRS.EOF
    With grd
        If LoanID <> rptRS("LoanID") Then TransID = 0
        LoanID = rptRS("LOanID")
        'Put the Sub Total
        If TransDate <> rptRS("TransDate") Then
            SlNo = 0
            If .Rows <= .Row + 2 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
            
            .Col = 6: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            
            .Col = 7: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
            TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
            
            .Col = 8: .Text = FormatCurrency(SubInterest): .CellFontBold = True
            TotalInterest = TotalInterest + SubInterest: SubInterest = 0
            .Col = 9: .Text = FormatCurrency(SubPenal): .CellFontBold = True
            TotalPenal = TotalPenal + SubPenal: SubPenal = 0
            TransDate = rptRS("TransDate"): TransID = 0
        End If
        
        If TransID <> rptRS("TransID") Then
            TransID = rptRS("TransID")
            If .Rows < .Row + 2 Then .Rows = .Row + 2
            .Row = .Row + 1
            SlNo = SlNo + 1
        End If
        transType = rptRS("TransType")
        Amount = FormatField(rptRS("IntAmount"))
        Balance = FormatField(rptRS("PenalBalance"))
        
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetIndianDate(TransDate)
        .Col = 2: .Text = rptRS("AccNum")
        .Col = 3: .Text = rptRS("MemberNum")
        ' Fill the loan holder name.
        .Col = 4: .Text = FormatField(rptRS("CustName"))
        .Col = 5: .Text = FormatField(rptRS("VoucherNo"))
        
        If rptRS(0) = "PRINCIPAL" Then  'If it is principal details
            If transType = wWithdraw Or transType = wContraWithdraw Then
                .Col = 6: .Text = FormatCurrency(Amount)  'LOan Advanced
                .Col = 10: .Text = FormatCurrency(Balance)  'Loan BAlance
                SubWithdraw = SubWithdraw + Amount
            Else
                .Col = 7: .Text = FormatCurrency(Amount) 'Loan Repaid
                .Col = 10: .Text = FormatCurrency(Balance) 'LOan Balance
                SubDeposit = SubDeposit + Amount
                SubBalance = SubBalance + Balance
            End If
        Else
            If transType = wDeposit Or transType = wContraDeposit Then
                .Col = 8: .Text = FormatCurrency(Amount)  'Regualr Interest
                .Col = 9: .Text = FormatCurrency(Balance)  'Penal Interest
                SubInterest = SubInterest + Amount
                SubPenal = SubPenal + Balance
            End If
        End If
   End With

nextRecord:
    DoEvents
    If gCancel Then rptRS.MoveLast
    Me.Refresh
    RaiseEvent Processing("Writing the data to the grid.", _
            rptRS.AbsolutePosition / rptRS.RecordCount)
    rptRS.MoveNext

Loop

rptRS.MoveLast
    
    With grd
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        .Col = 4: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
        TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        .Col = 7: .Text = FormatCurrency(SubDeposit): .CellFontBold = True
        TotalDeposit = TotalDeposit + SubDeposit: SubDeposit = 0
        .Col = 8: .Text = FormatCurrency(SubInterest): .CellFontBold = True
        TotalInterest = TotalInterest + SubInterest: SubInterest = 0
        .Col = 9: .Text = FormatCurrency(SubPenal): .CellFontBold = True
        TotalPenal = TotalPenal + SubPenal: SubPenal = 0
        
        'Put GrandTotal
        .Rows = .Row + 3
        .Row = .Row + 2
        .Col = 4: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 6: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True
        .Col = 7: .Text = FormatCurrency(TotalDeposit): .CellFontBold = True
        .Col = 8: .Text = FormatCurrency(TotalInterest): .CellFontBold = True
        .Col = 9: .Text = FormatCurrency(TotalPenal): .CellFontBold = True
        .Col = 10: .Text = FormatCurrency(SubBalance): .CellFontBold = True
                   
        If TotalInterest = 0 Then .ColWidth(7) = 5
        If TotalPenal = 0 Then .ColWidth(8) = 5
        If .Rows < .Row + 3 Then .Rows = .Row + 2
        
        'Added on 29/11/00(For Allignment)
        .ColAlignment(0) = 9
        .ColAlignment(1) = 9
        .ColAlignment(3) = 8
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 8
        .ColAlignment(7) = 9
        .ColAlignment(8) = 8
        .ColAlignment(9) = 7
    End With

' Display the grid.
grd.Visible = True

Me.Caption = "LOANS  [Sub Cash book...]"
lblReportTitle.Caption = GetResourceString(390, 85) & " " & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
ReportSubCashBook = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportSubDayBook: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        
    End If
'Resume
    GoTo Exit_Line

End Function

Private Function ReportLoanIssued() As Boolean
            
 ReportLoanIssued = False
 
' Declare variables...
Dim SqlStr As String
Dim Lret As Long
Dim rptRS As Recordset
Dim PrevMemberID As Long
Dim PrevLoanID As Long
Dim transType As wisTransactionTypes

' Setup error handler.
On Error GoTo Err_line

'raiseevent to run frmcancel
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Display status.
' Build the report query.
transType = wWithdraw
SqlStr = "SELECT SchemeName,C.CustomerID, MemberNum,A.loanID, A.TransDate," _
        & "Name  as custname, A.TransType, A.Amount, A.Balance,AccNum " _
        & " FROM LoanTrans A, LoanMaster B, QryMemName C, LoanScheme D WHERE " _
        & " A.LoanID = B.LoanID AND C.MemID = B.MemID " _
        & " And D.SchemeId = B.SchemeId " _
        & " And (TransType = " & wWithdraw & " OR TransType = " & wContraWithdraw & ")" _
        & " AND A.TransDate >= #" & m_FromDate & "#" _
        & " AND A.TransDate <= #" & m_ToDate & "#"


' Add the WHERE clause restrictions, if specified.
If m_SchemeId Then SqlStr = SqlStr & " AND B.SChemeID = " & m_SchemeId

If m_FromAmt > 0 Then SqlStr = SqlStr & " AND A.Amount >= " & m_FromAmt
If m_ToAmt > 0 Then SqlStr = SqlStr & " AND A.Amount <= " & m_ToAmt
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID
If Trim$(m_Purpose) <> "" Then SqlStr = SqlStr & " And B.LoanPurpose = " & AddQuotes(m_Purpose, True)
Dim TotalBalance As Currency

' Finally, add the sorting clause.
gDbTrans.SqlStmt = SqlStr & " ORDER BY A.TransDate, val(B.AccNum), A.TransId"
' Execute the query...
Lret = gDbTrans.Fetch(rptRS, adOpenStatic)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    GoTo Exit_Line
End If


' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst


' Initialize the grid.
'Call InitGrid
RaiseEvent Initialise(0, rptRS.RecordCount)
RaiseEvent Processing("Alignig the data tobe written into the grid.", 0)

With grd
    .Visible = False
End With
Call InitGrid

Dim SlNo As Long
Dim TransDate As Date
Dim SubWithdraw As Currency
Dim TotalWithDraw As Currency
Dim PRINTTotal As Boolean
' Fill the rows
SlNo = 1
grd.Rows = 40
grd.Row = grd.FixedRows - 1
TransDate = rptRS("TransDate")
Do While Not rptRS.EOF
    ' Set the row.
    If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
    With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If TransDate <> rptRS("TransDate") Then
            PRINTTotal = True
            'SlNo = SlNo + 1
            .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
            .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
            TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            TransDate = rptRS("TransDate")
        End If
        If FormatField(rptRS("Amount")) = 0 Then GoTo nextRecord
        transType = rptRS("TransType")
        ' Fill the loan id.
        If PrevLoanID <> FormatField(rptRS("LoanID")) Then
            .Col = 0: .Text = Format(SlNo, "00")
            .Col = 1: .Text = FormatField(rptRS("AccNum"))
            .Col = 2: .Text = FormatField(rptRS("MemberNum"))
            PrevLoanID = FormatField(rptRS("Loanid"))
            ' Fill the loan holder name.
            .Col = 3: .Text = FormatField(rptRS("CustName"))
            SlNo = SlNo + 1
        End If
        
        ' Fill the transaction date.
        .Col = 4: .Text = GetIndianDate(TransDate)
    End With
    'Show it in the repaid amount column.
    transType = rptRS("TransType")
    If transType = wWithdraw Or transType = wContraWithdraw Then
        grd.Col = 5: grd.Text = FormatField(rptRS("Amount"))
        SubWithdraw = SubWithdraw + Val(grd.Text)
    End If
    

nextRecord:
    DoEvents
    With rptRS
        If gCancel Then .MoveLast
        Me.Refresh
        RaiseEvent Processing("Writing the data to the grid.", .AbsolutePosition / .RecordCount)
        .MoveNext
    End With
Loop

With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(SubWithdraw): .CellFontBold = True
        TotalWithDraw = TotalWithDraw + SubWithdraw: SubWithdraw = 0
        
    If PRINTTotal Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(286): .CellFontBold = True
        .Col = 5: .Text = FormatCurrency(TotalWithDraw): .CellFontBold = True
    End If
    
    ' Display the grid.
    .Visible = True
End With

'Set the title
lblReportTitle = GetResourceString(290) & " " & _
    GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    
ReportLoanIssued = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    GoTo Exit_Line

End Function

Private Function ReportGeneralLedger() As Boolean

ReportGeneralLedger = False
' Declare variables...
Dim SqlStr As String
Dim PrinSql As String
Dim IntSql As String
Dim Lret As Long
Dim rst As Recordset
Dim opBalance As Currency

' Setup error handler.
On Error GoTo Err_line

RaiseEvent Processing("Reading & Verifying the data ", 0)

If m_SchemeId = 0 Then
    PrinSql = "Select 'Principal', SUM(Amount) as TotalAmount, " & _
            " TransDate, TransType From LoanTrans"
    IntSql = "Select sum(PenalIntAmount) as TotalIntAmount, " & _
            " sum(IntAmount) as TotalAmount ," & _
            " TransDate, TransType From LoanIntTrans"
Else
    PrinSql = " Select 'Principal', Sum(Amount) as TotalAmount, " & _
            " TransDate, TransType From LoanTrans Where LoanID in " & _
            " (Select LoanID From LoanMaster Where Schemeid = " & m_SchemeId & ")"
            
    IntSql = " Select Sum(PenalIntAmount) as PenalAmount,Sum(IntAmount) as TotalAmount, " & _
            " TransDate, TransType From LoanIntTrans Where LoanID in " & _
            " (Select LoanID From LoanMaster Where Schemeid = " & m_SchemeId & ")"
End If

    ' Add the Date Range
    If m_SchemeId = 0 Then
        PrinSql = PrinSql & " Where "
        IntSql = IntSql & " Where "
    Else
        PrinSql = PrinSql & " And "
        IntSql = IntSql & " And "
    End If
    
    PrinSql = PrinSql & " TransDate >= #" & m_FromDate & "#" & _
        " AND TransDate <= #" & m_ToDate & "# "
    
    IntSql = IntSql & " LoanIntTrans.TransDate >= #" & m_FromDate & "#" & _
        " AND LoanIntTrans.TransDate <= #" & m_ToDate & "#"
    
    PrinSql = PrinSql & " Group By TransDate, TransType"
    IntSql = IntSql & " Group By TransDate, TransType"

gDbTrans.SqlStmt = PrinSql & " UNION " & IntSql & " ORDER BY TransDate"

Lret = gDbTrans.Fetch(rst, adOpenStatic)
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error retrieving loan details.", vbCritical, wis_MESSAGE_TITLE
    Call PrintNoRecords(grd)
    GoTo Exit_Line
ElseIf Lret = 0 Then
    Call PrintNoRecords(grd)
    MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

' Initialize the grid.
Call InitGrid

Dim SubTotal(2 To 5) As Currency
Dim GrandTotal(2 To 5) As Currency

RaiseEvent Initialise(0, rst.RecordCount)
RaiseEvent Processing("Aligning the data ", 0)
' Fill the rows
Dim transType  As wisTransactionTypes
Dim TransDate As Date
Dim TransID As Long
Dim SlNo  As Integer
Dim I As Integer
Dim PRINTTotal As Boolean

TransDate = rst.Fields("TransDate")

With grd
    .Row = 1
    .Col = 1
    .Text = GetResourceString(284): .CellFontBold = True
    Dim loanClass As New clsLoan
    opBalance = loanClass.Balance(, m_SchemeId, m_FromDate)
    Set loanClass = Nothing
    .Col = 2: .Text = FormatCurrency(opBalance): .CellFontBold = True
End With

Do While Not rst.EOF
    ' Set the row.
    If rst.Fields("TransDate") <> TransDate Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Row + 2
            .Row = .Row + 1
            SlNo = SlNo + 1
            .Col = 0: .Text = Format(SlNo, "00")
            .Col = 1: .Text = GetIndianDate(TransDate)
            PRINTTotal = True
            For I = 2 To .Cols - 1
                .Col = I: .Text = FormatCurrency(SubTotal(I))
                .CellAlignment = 7
                GrandTotal(I) = GrandTotal(I) + SubTotal(I)
                SubTotal(I) = 0
            Next
        End With
        TransDate = rst.Fields("TransDate")
    End If
    transType = rst.Fields("TransType")
    With grd
        If rst.Fields(0) = "Principal" Then
            If transType = wWithdraw Or transType = wContraWithdraw Then
                SubTotal(2) = SubTotal(2) + FormatField(rst.Fields("TotalAmount"))
            Else
                SubTotal(3) = SubTotal(3) + FormatField(rst.Fields("TotalAmount"))
            End If
        End If
    End With
    ' Move to next row.
    DoEvents
    Me.Refresh
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Loop

    With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = GetIndianDate(TransDate)
        For I = 2 To .Cols - 1
            .Col = I: .Text = FormatCurrency(SubTotal(I))
            .CellAlignment = 7 ':  .CellFontBold = True
            GrandTotal(I) = GrandTotal(I) + SubTotal(I)
            SubTotal(I) = 0
        Next
    End With

If PRINTTotal Then
    With grd
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        '.Col = 0: .Text = Format(SlNo, "00")
        '.Col = 1: .Text = GetIndianDate(TransDate)
        For I = 2 To .Cols - 1
            .Col = I: .Text = FormatCurrency(GrandTotal(I))
            .CellAlignment = 7: .CellFontBold = True
        Next
        
        If .Rows <= .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 1: .Text = GetResourceString(285): .CellFontBold = True
        .Col = 3: .Text = FormatCurrency(opBalance + GrandTotal(2) - GrandTotal(3)): .CellFontBold = True
        
    End With
End If

'set the title
lblReportTitle.Caption = GetResourceString(93) & " " & _
                        GetFromDateString(m_FromIndianDate, m_ToIndianDate)

ReportGeneralLedger = True

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportDailyCashBook: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
''' Resume
    GoTo Exit_Line
End Function


Private Function ReportLoanHolders() As Boolean


ReportLoanHolders = False
' Set up error handler.
On Error GoTo Err_line

' Declare required variables.
Dim SqlStr As String
Dim SlNo As Long
RaiseEvent Processing("Reading & Verifying the records ", 0)

' Prepare the SQL query...
SqlStr = "SELECT A.IssueDate, A.LoanId, AccNum,MemberNum, SchemeName,A.LoanAmount, " & _
        " Caste,Place, Balance, Name as CustName FROM LoanMaster A, " & _
        " QryMemName B, LoanScheme C,LoanTrans D WHERE B.MemID = A.MemID " & _
        " AND A.SchemeId = C.SchemeId And TransID = (SELECT MAx(transID) " & _
            " From LoanTrans E WHERE E.LoanId = A.LoanID )" & _
        " AND d.LoanID = A.LoanID And Balance > 0"
        
'If scheme type is specified, add the scheme type restriction clause.
If m_SchemeId Then SqlStr = SqlStr & " AND C.schemeID = " & m_SchemeId

' If asOnIndianDate specified, add date clause.
If Trim$(m_FromIndianDate) <> "" Then _
    SqlStr = SqlStr & " AND A.IssueDate <= #" & m_ToDate & "#"

' If stamt specified, add amount clause.
If m_FromAmt <> 0 Then SqlStr = SqlStr & " AND a.LoanAmount >= " & m_FromAmt

' If endamt specified, add amount clause.
If m_ToAmt <> 0 Then SqlStr = SqlStr & " AND a.LoanAmount >= " & m_ToAmt


'Search for the Caste & Place
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID

If Trim$(m_Purpose) <> "" Then SqlStr = SqlStr & " And A.LoanPurpose = " & AddQuotes(m_Purpose, True)

'Qing the loanmaster
If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " order by A.SchemeID,val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " order by A.SchemeID,IsciName"
End If
    
Dim rst As Recordset
Dim rstTemp As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

Dim TotalBalance As Currency

' Finally, execute the query.
gDbTrans.SqlStmt = SqlStr
SlNo = gDbTrans.Fetch(rstTemp, adOpenStatic)
If SlNo < 0 Then
    MsgBox "Error getting loan transaction details.", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf SlNo = 0 Then
    MsgBox "No records.", vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Load the details to grid.

' Initialize the grid.
RaiseEvent Initialise(0, SlNo)
RaiseEvent Processing("Aligning the data", 0)

Call InitGrid
grd.Row = grd.FixedRows - 1
' Fill the rows
SlNo = 0
Do While Not rst.EOF
    With grd
        ' Set the row.
        If .Row < rst.AbsolutePosition + 2 Then .Rows = rst.AbsolutePosition + 2
        .Row = rst.AbsolutePosition
        
        SlNo = SlNo + 1
        .Col = 0
        .Text = Format(SlNo, "00")
        
        ' Fill the loan Account no.
        .Col = 1: .Text = rst("AccNum")
        .Col = 2: .Text = rst("MemberNum")
        'Fill the loan holder name.
        .Col = 3: .Text = rst("CustName")
        
        ' Fill the loan holder name Caste .
        .Col = 4: .Text = rst("Caste")
        
        ' Fill the loan holder name PLace .
        .Col = 5: .Text = rst("Place")

        ' Fill the loan scheme name.
        .Col = 6: .Text = rst("SchemeName")

        ' Fill the loan issue date.
        .Col = 7: .Text = FormatField(rst("IssueDate"))
                      
        ' Fill the loan amount.
        .Col = 8: .Text = rst("LoanAmount"): .CellAlignment = 7
        .Col = 9: .Text = rst("Balance"): .CellAlignment = 7
    End With

    ' Move to next row.
    DoEvents
    
    With rst
        If gCancel Then .MoveLast
        .MoveNext
        RaiseEvent Processing("Writing the data to the grid.", SlNo / .RecordCount)
    End With
    
Loop
    
    'Added for the allignment(19/11/01)
    With grd
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        
    End With
    
' Display the grid.
grd.Visible = True

'set the tilte
lblReportTitle.Caption = GetResourceString(83) & " " & _
    GetFromDateString(m_ToIndianDate)

Me.Caption = "Loans  [List of loans holders...]"

ReportLoanHolders = True
'Call grd_LostFocus

Exit_Line:
    Exit Function

Err_line:
    If Err Then
        MsgBox "ReportLoanHolders: " & vbCrLf _
                & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Function

Private Function ReportOverdueLoans() As Boolean

m_ReportType = repLoanOD

ReportOverdueLoans = False
' Declare variables...
Dim Lret As Long
Dim SqlStr As String

' Setup error handler.
On Error GoTo Err_line
Me.MousePointer = vbHourglass

RaiseEvent Processing("Reading & Verifying the records ", 0)

' Build the report query.
SqlStr = "SELECT A.LoanID, AccNum,MemberNum,LoanDueDate,Balance,SchemeName,Caste,Place, " _
    & " LoanAmount, Name as CustName " _
    & " FROM LoanMaster A, LoanScheme B, LoanTrans C, QryMemName D WHERE " _
    & " TransId = (Select MAX(TransID) From LoanTrans E  Where E.LoanId = A.LoanId " _
        & " And TransDate <= #" & m_ToDate & "# )" _
    & " And Balance > 0 AND LoanDueDate <= #" & m_ToDate & "# " _
    & " AND C.LoanID= A.LoanID " _
    & " AND B.SchemeID = A.SchemeiD AND D.MemID = A.MemID "
        
If m_SchemeId Then SqlStr = SqlStr & " AND A.SchemeID = " & m_SchemeId

If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID
If Trim$(m_Purpose) <> "" Then SqlStr = SqlStr & " And A.LoanPurpose = " & AddQuotes(m_Purpose, True)

gDbTrans.SqlStmt = SqlStr & " order by B.SchemeID,Val(A.AccNum)" 'from loanscheme

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

Dim TotalBalance As Currency
    
grd.Clear
grd.Cols = 5
grd.FixedCols = 1
grd.Rows = 20

'Raise event to access frmcancel.
RaiseEvent Initialise(0, rst.RecordCount + 1)
RaiseEvent Processing("Aligning the data ", 0)

' Initialize the grid.
Call InitGrid

Dim SlNo As Integer
Dim TotalAmount As Currency
Dim ODAmount As Currency
Dim TotalODAmount As Currency
Dim RegInt As Currency
Dim TotalRegInt As Currency
Dim PenalInt As Currency
Dim TotalPenalInt As Currency
Dim L_clsLoan As New clsLoan
Dim totalCount As Integer
' Fill the rows
SlNo = 1
        
totalCount = rst.RecordCount + 2
Do While Not rst.EOF
    With grd
        ' Set the row.
        If .Rows <= SlNo + 1 Then .Rows = .Rows + 2
        .Row = SlNo
        .Col = 0
        .Text = SlNo

        ' Fill the loanid.
        .Col = 1: .Text = rst("AccNum")
        .Col = 2: .Text = rst("MemberNum")

       'Fill the loan holder name.
        .Col = 3: .Text = FormatField(rst("CustName"))

        ' Fill the loan issue date.
        .Col = 4
        .Text = FormatField(rst("LoanDueDate"))
        
        ' Fill the loan amount.
        .Col = 5
        .Text = FormatCurrency(rst("Balance")): grd.CellAlignment = 7
        ODAmount = L_clsLoan.OverDueAmount(rst("LOanID"), , m_ToDate)
        
        'Check the OD AMount Condition
        If ODAmount < m_FromAmt Then GoTo nextRecord
        If m_ToAmt <> 0 And ODAmount > m_ToAmt Then GoTo nextRecord
        
        RegInt = L_clsLoan.RegularInterest(rst("LoanID"), , m_ToDate)
        PenalInt = L_clsLoan.PenalInterest(rst("LoanID"), , m_ToDate)
        If ODAmount = 0 And PenalInt = 0 Then
            .Row = .Row - 1
            GoTo nextRecord
        End If
        .Col = 6: .Text = FormatCurrency(ODAmount)
        TotalODAmount = TotalODAmount + ODAmount
        
        .Col = 7: .Text = FormatCurrency(RegInt)
        TotalRegInt = TotalRegInt + RegInt
        
        .Col = 8: .Text = FormatCurrency(PenalInt)
        TotalPenalInt = TotalPenalInt + PenalInt
        
        .Col = 9: .Text = FormatCurrency(ODAmount + RegInt + PenalInt)
        
        SlNo = SlNo + 1
        
nextRecord:
        
        DoEvents
        If gCancel Then rst.MoveLast
        RaiseEvent Processing("Writing the data ", rst.AbsolutePosition / rst.RecordCount)
        ' Move to next row.
        rst.MoveNext
   
    End With
Loop

    With grd
        .Row = .Row + 1: .Col = 2
        .Text = GetResourceString(286): .CellFontBold = True: .CellAlignment = 7
        .Col = 6: .Text = FormatCurrency(TotalODAmount)
            .CellFontBold = True: .CellAlignment = 7
        .Col = 7: .Text = FormatCurrency(TotalRegInt)
            .CellFontBold = True: .CellAlignment = 7
        .Col = 8: .Text = FormatCurrency(TotalPenalInt)
            .CellFontBold = True: .CellAlignment = 7
        .Col = 9: .Text = FormatCurrency(TotalODAmount + TotalRegInt + TotalPenalInt)
            .CellFontBold = True: .CellAlignment = 7
        'Display the grid.
        .Visible = True
    End With
Set L_clsLoan = Nothing
'Me.Caption = "LOANS [Over Due Loans]"

'set the title
lblReportTitle.Caption = GetResourceString(84) & " " & _
    GetResourceString(18) & " " & GetFromDateString(m_ToIndianDate)

ReportOverdueLoans = True
'Call grd_LostFocus
Exit_Line:
    Set rst = Nothing
    Me.MousePointer = vbDefault
    Exit Function
Err_line:
    If Err Then
        MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
    End If
'Resume
    GoTo Exit_Line

End Function

Private Function ReportOverdueInstalments() As Boolean

m_ReportType = repLoanInstOD

ReportOverdueInstalments = False

' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim OverDueLoan As Boolean
Dim SqlStr As String

' Setup error handler.
On Error GoTo Err_line
'initialise the grid
With grd
    .Clear
    .Cols = 1
    .FixedCols = 0
    .Rows = 30
    .Row = 1
    .Text = "No Records"
    .CellAlignment = 4: .CellFontBold = True
End With

Me.MousePointer = vbHourglass
'
RaiseEvent Processing("Reading & Verifying the data ", 0)
' Build the report query. to Get Instalment Loan
'New Code  Shashi 29/7/00
Dim InstType As wisInstallmentTypes
InstType = Inst_No
SqlStr = "SELECT A.Loanid, MemberNum,A.LoanDueDate,IssueDate,InstAmount, A.LoanAmount, " _
        & " A.InstMode, Name As CustName, A.CustomerID, " _
        & " SchemeName,AccNum FROM LoanMaster A, LoanScheme B,  QryMemName C " _
        & " WHERE InstMode <> " & InstType & " And B.SchemeId = A.SchemeId " _
        & " AND C.MemID = A.MemID "

' Add the "WHERE" clause if date/Scheme range is specified.
If m_SchemeId Then SqlStr = SqlStr & " AND B.SchemeID = " & m_SchemeId
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Gender Then SqlStr = SqlStr & " AND Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID

If m_ReportOrder = wisByAccountNo Then
    gDbTrans.SqlStmt = SqlStr & " order by B.SchemeID,val(AccNum)"
Else
    gDbTrans.SqlStmt = SqlStr & " order by B.SchemeID,IsciName"
End If

Dim rst As Recordset
Lret = gDbTrans.Fetch(rptRS, adOpenStatic)
If Lret < 1 Then Exit Function


Dim TotalBalance As Currency
    
    ' InitGird
    grd.Clear
    grd.Cols = 5
    grd.FixedCols = 1
    grd.Rows = 20

If Lret < 0 Then
    ' Error in database.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    'GoTo Exit_Line
End If

If Lret = 0 Then
    MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If


' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Initialize the grid.
RaiseEvent Initialise(0, rptRS.RecordCount + 1)
RaiseEvent Processing("Aligning the data ", 0)

Call InitGrid
'If RepSocDet = DirectLoans Then grd.ColWidth(3) = 5

Dim TotalNo As Integer
Dim SlNo As Integer
Dim LoanID As Long
Dim Balance As Currency
Dim ODAmount As Currency
Dim RegInt As Currency
Dim PenalInt As Currency

Dim TotalRegInt As Currency
Dim TotalODAmount As Currency
Dim TotalPenalInt As Currency

' Fill the rows
Dim L_clsLoan As New clsLoan

TotalNo = rptRS.RecordCount

Do While Not rptRS.EOF
    With grd
        'Now find the Whether he has paid the all the Instalments as on date _
        'or he has skipped any of them First Find the Mode Of Instalment
        LoanID = Val(FormatField(rptRS("LoanID")))
        InstType = Val(FormatField(rptRS("InstMode")))
        Balance = L_clsLoan.Balance(LoanID, , m_ToDate)
        ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
        RegInt = L_clsLoan.RegularInterest(LoanID, , m_ToDate)
        PenalInt = L_clsLoan.PenalInterest(LoanID, , m_ToDate)
        If ODAmount = 0 Then GoTo nextRecord
        ' Set the row.
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1
        SlNo = SlNo + 1
        .Col = 0
        .Text = SlNo
        ' Fill the loanid.
        .Col = 1: .Text = rptRS("AccNum")
        .Col = 2: .Text = rptRS("MemberNum")


       ' Fill the loan holder name.
        .Col = 3: .Text = FormatField(rptRS("CustName"))
        
        
    ' Fill the loan details.
        .Col = 4: .Text = FormatField(rptRS("LoanDueDate"))
        .Col = 5: .Text = FormatCurrency(Balance)
        .Col = 6: .Text = FormatCurrency(ODAmount)
        .Col = 7: .Text = FormatCurrency(RegInt)
        .Col = 8: .Text = FormatCurrency(PenalInt)
    End With
    
        TotalBalance = TotalBalance + Balance
        TotalODAmount = TotalODAmount + ODAmount
        TotalRegInt = TotalRegInt + RegInt
        TotalPenalInt = TotalPenalInt + PenalInt

nextRecord:
    
    DoEvents
    If gCancel Then rptRS.MoveLast
    RaiseEvent Processing("Writing the data ", rptRS.AbsolutePosition / TotalNo)
'    Debug.Assert SlNo <> 37
    ' Move to next row.
    rptRS.MoveNext
Loop

    If TotalODAmount > 0 Then
        With grd
            If .Rows <= .Row + 1 Then .Rows = .Rows + 2
            .Row = .Row + 1
            .Col = 3: .Text = GetResourceString(286): .CellFontBold = True: .CellAlignment = 7
            .Col = 5: .Text = FormatCurrency(TotalBalance): .CellFontBold = True: .CellAlignment = 7
            .Col = 6: .Text = FormatCurrency(TotalODAmount): .CellFontBold = True: .CellAlignment = 7
            .Col = 7: .Text = FormatCurrency(TotalRegInt): .CellFontBold = True: .CellAlignment = 7
            .Col = 8: .Text = FormatCurrency(TotalPenalInt): .CellFontBold = True: .CellAlignment = 7
        End With
    Else
        Call PrintNoRecords(grd)
    End If
' Display the grid.
grd.Visible = True
'set the title
lblReportTitle.Caption = GetResourceString(113) & " " & _
        GetFromDateString(m_ToIndianDate)

'Me.Caption = "Loans [OverDue installments ...]"
Set L_clsLoan = Nothing
ReportOverdueInstalments = True
'Call grd_LostFocus
Exit_Line:
    'rptRS.Close
    Set rptRS = Nothing
    Me.MousePointer = vbDefault
    Exit Function

Err_line:
    If Err Then
       MsgBox "ReportInstallmentOverdue: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function

Private Function ReportLoanBalance() As Boolean

ReportLoanBalance = False
Err.Clear
On Error GoTo ExitLine

Dim SqlStmt As String
'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)
    
SqlStmt = "Select B.LoanId, AccNum,C.MemberNum, Balance, SchemeName, " & _
    " Place, Caste, Name as CustName " & _
    " From LoanMaster A, LoanTrans B, QryMemName C, LoanScheme D Where TransId = " & _
        "(Select Max(TransId) From LoanTrans E Where E.LoanId = A.LoanId " & _
        " And TransDate <= #" & m_ToDate & "# ) " & _
    " ANd B.Balance <> 0 And B.LoanId = A.LoanId And C.MemID = A.MemID " _
    & " And D.SchemeId = A.SchemeID "

If m_SchemeId Then SqlStmt = SqlStmt & " And A.SchemeId = " & m_SchemeId
 
If m_FromAmt <> 0 Then SqlStmt = SqlStmt & " And  Balance >= " & m_FromAmt
If m_ToAmt <> 0 Then SqlStmt = SqlStmt & " And Balance <= " & m_ToAmt

If Trim$(m_Place) <> "" Then SqlStmt = SqlStmt & " And Place = " & AddQuotes(m_Place, True)
If Trim$(m_Caste) <> "" Then SqlStmt = SqlStmt & " And Caste = " & AddQuotes(m_Caste, True)
If m_Gender <> wisNoGender Then SqlStmt = SqlStmt & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStmt = SqlStmt & " And AccGroupID = " & m_AccGroupID

If Trim$(m_Purpose) <> "" Then SqlStmt = SqlStmt & " And A.LoanPurpose = " & AddQuotes(m_Purpose, True)

'Quaring the loanmaster

gDbTrans.SqlStmt = SqlStmt & " Order by A.SchemeId, " & IIf(m_ReportOrder = wisByAccountNo, "val(A.AccNum)", "IsciName")

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    Call PrintNoRecords(Me.grd)
    Exit Function
End If

    Dim TotalBalance As Currency
    Dim SubBalance As Currency
    Dim SlNo As Long
    Dim l_SchemeID As Integer
    Dim SchemeName As String
    
    
    RaiseEvent Initialise(0, rst.RecordCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    ' InitGird
    grd.Clear
    grd.FixedRows = 1
    Dim ColWid As Single
    
    ColWid = grd.Width / grd.Cols
    grd.Row = 0
    grd.FormatString = ">Sl No|<Loan Id|<CustomerName|>Balance"
    Call InitGrid
    
grd.Row = grd.FixedRows - 1
    
Dim l_Cust As New clsCustReg
Dim TotalNo As Integer
TotalNo = rst.RecordCount + 2
SchemeName = FormatField(rst("SchemeName"))
SlNo = 1
Dim ProcessNo As Integer

While Not rst.EOF
    With grd
        If SchemeName <> FormatField(rst("SchemeName")) Then
            
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 3: .Text = SchemeName & " " & GetResourceString(304)
            grd.CellAlignment = 7: .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(SubBalance)
            grd.CellAlignment = 7: .CellFontBold = True
            SubBalance = 0
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1: SlNo = 1
            SchemeName = FormatField(rst("SchemeName"))
        End If
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("MemberNum"))
        .Col = 3: .Text = FormatField(rst("CustName"))
        .Col = 4: .Text = FormatField(rst("Balance")): .CellAlignment = 7
        TotalBalance = TotalBalance + Val(.Text)
        SubBalance = SubBalance + Val(.Text)
        
    End With
    SlNo = SlNo + 1
    ProcessNo = ProcessNo + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", ProcessNo / TotalNo)
    rst.MoveNext
    
Wend

lblReportTitle.Caption = GetResourceString(67) & " " & _
        GetResourceString(58) & " " & _
            GetFromDateString(m_ToIndianDate)
With grd
    If m_SchemeId = 0 Then
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 3: .Text = GetResourceString(304): .CellFontBold = True
        .Col = 4: .Text = FormatCurrency(SubBalance): .CellFontBold = True
    End If
            
    If .Rows <= .Row + 3 Then .Rows = .Row + 4
    .Row = .Row + 2
    .Col = 3: .Text = GetResourceString(286)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

    ReportLoanBalance = True

ExitLine:
    If Err Then
        MsgBox "ERROR ReportLoanBalance" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
    End If
End Function

Private Function ReportSanctionedLoans() As Boolean

m_ReportType = repLoanSanction

ReportSanctionedLoans = False
  
Dim SqlStr As String
Dim rstMaster As Recordset
Dim MemID As Integer
Dim I As Long
Dim loopCount As Long
    
RaiseEvent Processing("Reading & Verifying the data ", 0)

SqlStr = "SELECT A.CustomerID, AccNum,Loanid,LoanAmount,Caste,Place,SchemeName," & _
            " Name as CustName From LoanMaster A,QryName B,LoanScheme C WHERE " & _
            " B.CustomerID =A.CustomerID AND C.SchemeID = A.SchemeID " & _
            " AND IssueDate >= #" & m_FromDate & "# " & _
            " AND IssueDate <= #" & m_ToDate & "#"

If m_SchemeId Then SqlStr = SqlStr & " AND A.Schemeid = " & m_SchemeId


If m_FromAmt <> 0 Then SqlStr = SqlStr & " AND LoanAmount >= " & m_FromAmt
If m_ToAmt <> 0 Then SqlStr = SqlStr & " AND LoanAmount <= " & m_ToAmt
If m_Caste <> "" Then SqlStr = SqlStr & " AND Caste = " & AddQuotes(m_Caste, True)
If m_Place <> "" Then SqlStr = SqlStr & " AND Place = " & AddQuotes(m_Place, True)
If m_Gender <> wisNoGender Then SqlStr = SqlStr & " And Gender = " & m_Gender
If m_AccGroupID Then SqlStr = SqlStr & " And AccGroupID = " & m_AccGroupID
If Trim$(m_Purpose) <> "" Then SqlStr = SqlStr & _
        " And A.LoanPurpose = " & AddQuotes(m_Purpose, True)

gDbTrans.SqlStmt = SqlStr & " order by " & _
    IIf(m_ReportOrder = wisByAccountNo, "val(A.AccNum)", "ISCIName")
'Dim Rst As Recordset
If gDbTrans.Fetch(rstMaster, adOpenStatic) < 1 Then Exit Function

Dim TotalBalance As Currency
    
' InitGird
grd.Clear
Call InitGrid

loopCount = rstMaster.RecordCount
  
For I = 1 To loopCount
    If rstMaster.EOF Then Exit For
    MemID = Val(FormatField(rstMaster("CustomerID")))
    'split the Instalmentname field
    With grd
        .Row = I
        .Col = 0: .Text = CStr(I)
        .Col = 1: .Text = FormatField(rstMaster("AccNum"))
        'Loan holder name
        .Col = 2: .CellAlignment = 1
        .Text = FormatField(rstMaster("CustName")) 'MemName
        .Col = 3: .CellAlignment = 7
        .Text = FormatField(rstMaster("LoanAmount"))
        TotalBalance = TotalBalance + Val(.Text)
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
    End With
    
    DoEvents
    If gCancel Then Exit For
    RaiseEvent Processing("Writing the data ", I / loopCount)
    rstMaster.MoveNext
Next I

With grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(286)
    .CellAlignment = 7: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 7: .CellFontBold = True
End With

'set the title
lblReportTitle.Caption = GetResourceString(262) & " " & _
                    GetFromDateString(m_FromIndianDate, m_ToIndianDate)

grd.Visible = True
ReportSanctionedLoans = True

'Call grd_LostFocus

ErrLine:
  If Err Then
      MsgBox Err.Description
      'Resume
  End If
End Function

Private Function ReportGuarantors() As Boolean

m_ReportType = repLoanGuarantor

ReportGuarantors = False

Dim SqlStmt As String
Dim PrinSql As String
Dim IntSql As String

'raiseevent to access frmcancel
RaiseEvent Processing("Reading & Verifying the data ", 0)

SqlStmt = "Select SchemeName,Guarantor1,Guarantor2,A.LoanId,AccNum," & _
    " Name as CustName From LoanMaster A, LoanScheme B,QryName C Where " & _
    " C.CustomerID = A.CustomerID AND B.SchemeID = A.SchemeID " & _
    " AND (Guarantor1 > 0 OR Guarantor2 > 0) "

If m_FromIndianDate <> "" Then
    SqlStmt = SqlStmt & " AND A.LoanId IN (SELECT DISTINCT LoanID FROM LoanTrans E " & _
        " Where Balance > 0 AND TransDate <=#" & m_ToDate & "# )"
End If

If m_SchemeId Then SqlStmt = SqlStmt & " And A.SchemeId = " & m_SchemeId
    
gDbTrans.SqlStmt = SqlStmt & " ORDER BY IssueDate "
    
    Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then
    Call PrintNoRecords(grd)
    Exit Function
End If
    Dim rstGuar As Recordset
    Dim SlNo As Long
    Dim Address As Boolean
    
    ' InitGird
    RaiseEvent Initialise(0, rst.RecordCount)
    RaiseEvent Processing("Aligning the data ", 0)
    
    Call InitGrid
    grd.Row = grd.FixedRows - 1
    
   Dim FirstGuarantor As String
   Dim SecondGuarantor As String
   
SlNo = 1
Dim count As Integer

While Not rst.EOF
    Address = False
    With grd
        FirstGuarantor = FormatField(rst("Guarantor1"))
        SecondGuarantor = FormatField(rst("Guarantor2"))
        
        If FirstGuarantor <= 0 And SecondGuarantor <= 0 Then GoTo nextRecord
        Debug.Assert SlNo <> 28
        If .Rows <= .Row + 3 Then .Rows = .Rows + 2
        .Row = .Row + 1
        .Col = 0: .Text = Format(SlNo, "00")
        .Col = 1: .Text = FormatField(rst("AccNum"))
        .Col = 2: .Text = FormatField(rst("CustName"))
        
        On Error Resume Next
        If FirstGuarantor > 0 Then
            gDbTrans.SqlStmt = "SELECT Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
                    " LastName as NAME, HomeAddress, OfficeAddress,Place " & _
                    " FROM NameTab Where CustomerID = " & FirstGuarantor
            If gDbTrans.Fetch(rstGuar, adOpenStatic) Then
                .Col = 3
                .Text = FormatField(rstGuar("Name"))
                .Row = .Row + 1
                .Col = 3
                .Text = FormatField(rstGuar("HomeAddress")) & " " & FormatField(rstGuar("OfficeAddress")) & _
                    FormatField(rstGuar("Place"))
                .RowHeight(.Row) = TextHeight(.Text)
                .Row = .Row - 1
                Address = True
            End If
        End If
        If SecondGuarantor > 0 Then
            gDbTrans.SqlStmt = "SELECT Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
                    " LastName as NAME, HomeAddress, OfficeAddress,Place " & _
                    " FROM NameTab Where CustomerID = " & SecondGuarantor
            If gDbTrans.Fetch(rstGuar, adOpenStatic) Then
                .Col = 4
                .Text = FormatField(rstGuar("Name"))
                .Row = .Row + 1: .Col = 4
                .Text = FormatField(rstGuar("HomeAddress")) & " " & FormatField(rstGuar("OfficeAddress")) & _
                    FormatField(rstGuar("Place"))
                .RowHeight(.Row) = TextHeight(.Text)
                .Row = .Row - 1
                Address = True
            End If
        End If
        On Error GoTo ExitLine
        If Address Then .Row = .Row + 1
    End With
    
    SlNo = SlNo + 1
nextRecord:
    
    count = count + 1
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", count / (rst.RecordCount + 1))
    rst.MoveNext
Wend

    lblReportTitle.Caption = GetResourceString(389) & " " & "list" & " " & _
                             GetFromDateString(m_FromIndianDate, m_ToIndianDate)
    ReportGuarantors = True
    Err.Clear

ExitLine:
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    End If

End Function

Private Sub Form_Resize()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    lblReportTitle.Top = 100
    lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
    grd.Left = 0
    grd.Top = lblReportTitle.Top + lblReportTitle.Height
    grd.Width = Me.Width - 150
    fra.Top = Me.ScaleHeight - fra.Height
    fra.Left = Me.Width - fra.Width
    grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height - lblReportTitle.Top
    cmdOk.Left = fra.Width - cmdOk.Width - (cmdOk.Width / 4)
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - (cmdPrint.Width / 4)
    cmdWeb.Top = cmdPrint.Top
    cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)
   
    Dim Wid As Single
    Dim ColCount As Long

For ColCount = 0 To grd.Cols - 1
    Wid = (grd.Width - 185) / grd.Cols
    grd.ColWidth(ColCount) = GetSetting(App.EXEName, "LoanReport" & m_ReportType, _
        "ColWidth" & ColCount, grd.Width / grd.Cols) * grd.Width
    
    If grd.ColWidth(ColCount) > grd.Width * 0.9 Then grd.ColWidth(ColCount) = grd.Width / grd.Cols
    If grd.ColWidth(ColCount) < 20 Then grd.ColWidth(ColCount) = 30
    Me.Refresh
Next ColCount

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set frmLoanView = Nothing
RaiseEvent WindowClosed
End Sub


Private Sub grd_LostFocus()
If grd.Cols = 1 Then Exit Sub
Dim ColCount As Integer
With grd
    If .Rows <= .FixedRows Then Exit Sub
    If .Row < .FixedRows Then .Row = .FixedRows
    For ColCount = 0 To .Cols - 1
        Call SaveSetting(App.EXEName, "LoanReport" & m_ReportType, _
                "ColWidth" & ColCount, .ColWidth(ColCount) / .Width)
        
    Next ColCount
End With
End Sub


Private Sub m_frmCancel_CancelClicked()
'If PrintClass Is Nothing Then Exit Sub
m_grdPrint.CancelProcess
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdPrint.Caption = GetResourceString(23)  'Print
cmdOk.Caption = GetResourceString(1) 'OK

End Sub

Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_MaxCount = MaxCount
End Sub

Private Sub m_grdPrint_Message(strMessage As String)
m_frmCancel.lblMessage = strMessage
End Sub


Private Sub m_grdPrint_ProcessCount(count As Long)
On Error Resume Next
'm_frmCancel.prg.Value = Count
UpdateStatus m_frmCancel.PicStatus, count / m_MaxCount
Err.Clear
End Sub

