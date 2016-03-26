VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCustSearch 
   ClientHeight    =   6255
   ClientLeft      =   1740
   ClientTop       =   1605
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   7365
   Begin VB.Frame fraName 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   180
      TabIndex        =   8
      Top             =   120
      Width           =   5325
      Begin VB.CheckBox chkLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name"
         Height          =   315
         Left            =   3450
         TabIndex        =   11
         Top             =   180
         Width           =   1725
      End
      Begin VB.CheckBox chkBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Balance"
         Height          =   315
         Left            =   3420
         TabIndex        =   4
         Top             =   600
         Width           =   1725
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   1440
         TabIndex        =   1
         Top             =   120
         Width           =   1875
      End
      Begin VB.ComboBox cmbPlace 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label lblAccType 
         Caption         =   "Label3"
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label lblCustName 
         Caption         =   "Name:"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblPlace 
         Caption         =   "Location"
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   630
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   5910
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   400
      Left            =   5910
      TabIndex        =   7
      Top             =   300
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4095
      Left            =   90
      TabIndex        =   10
      Top             =   1920
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   7223
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCustSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_MinWidth As Single
'Private m_Minheight As Single

Private Const m_MinWidth = 4305
Private Const m_Minheight = 3430
Private m_ClsObject As Object

Private Sub InitGrid()

Dim rst As Recordset
Dim AccHeadID As Long
Dim ParentID As Long
'Dim TotalCols As Integer
Dim ShowBalance As Boolean

ShowBalance = IIf(chkBalance.Value = vbChecked, -1, 0)

If cmbAccType.ListIndex < 1 Then
    ParentID = 0: AccHeadID = 0
Else
    AccHeadID = cmbAccType.ItemData(cmbAccType.ListIndex)
    If AccHeadID Mod SUB_HEAD_OFFSET = 0 Then ParentID = AccHeadID: AccHeadID = 0
End If

'TotalCols = 2
'If AccHeadId Then TotalCols = 3
If ParentID Then
    gDbTrans.SqlStmt = "Select * From Heads Where ParentID = " & ParentID
    Call gDbTrans.Fetch(rst, adOpenDynamic)
    'TotalCols = TotalCols + Rst.RecordCount
End If

Dim MaxI As Integer

With grd
    .Clear
    .Rows = 2: .FixedRows = 1
    .Cols = 2: .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(35)
    .MergeCells = flexMergeNever
    .AllowUserResizing = flexResizeBoth
    
    If chkBalance.Value = vbChecked Then
'        .Cols = .Cols + 1
        .Rows = 3: .FixedRows = 2
        .Row = 1
        .Col = 0: .Text = GetResourceString(33)
        .Col = 1: .Text = GetResourceString(35)
        .MergeCells = flexMergeRestrictAll
        .MergeRow(0) = True: .MergeRow(1) = True
        .MergeCol(0) = True: .MergeCol(1) = True
        .Row = 0
    End If
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    
    If AccHeadID Then
        .Cols = 3
        .Col = 2
        .Text = cmbAccType.Text
        .ColData(2) = AccHeadID
        If ShowBalance Then
            .Cols = 4
            .Col = 3
            .Text = cmbAccType.Text
            .Row = 1
            .Col = .Cols - 2: .Text = GetResourceString(36) & " " & _
                                GetResourceString(60)
            .Col = .Cols - 1: .Text = GetResourceString(67)
            .Row = 0
        End If
        GoTo LastLine
    End If
    If Not rst Is Nothing Then
        While Not rst.EOF
            .Cols = .Cols + 1
            .Col = .Cols - 1
            .Text = FormatField(rst("HeadName"))
            .ColData(.Col) = FormatField(rst("HeadID"))
            If ShowBalance Then
                .Cols = .Cols + 1
                .Col = .Cols - 1
                .Text = FormatField(rst("HeadName"))
                .Row = 1
                .Col = .Cols - 2: .Text = GetResourceString(36) & " " & _
                                    GetResourceString(60)
                .Col = .Cols - 1: .Text = GetResourceString(67)
                .Row = 0
            End If
            rst.MoveNext
        Wend
        GoTo LastLine
    End If
    
    Dim I As Integer
    'Load All Heads
    For I = 1 To cmbAccType.ListCount - 1
        AccHeadID = cmbAccType.ItemData(I)
        If AccHeadID Mod SUB_HEAD_OFFSET Then
            .Row = 0
            .Cols = .Cols + 1
            .Col = .Cols - 1
            .Text = cmbAccType.List(I)
            .ColData(.Col) = AccHeadID
            If ShowBalance Then
                .Cols = .Cols + 1
                .Col = .Cols - 1
                .Text = cmbAccType.List(I)
                .Row = 1
                .Col = .Cols - 2: .Text = GetResourceString(36) & " " & _
                                    GetResourceString(60)
                .Col = .Cols - 1: .Text = GetResourceString(67)
            End If
        End If
    Next
    AccHeadID = 0
LastLine:
    
.Row = 0
For I = 0 To .Cols - 1
    .Col = I
    .CellAlignment = 4
    .CellFontBold = True
Next

If ShowBalance Then
    .Row = 1
    For I = 0 To .Cols - 1
        .Col = I
        .CellAlignment = 4
        .CellFontBold = True
    Next
End If

End With

End Sub

Private Sub SearchCustomer()
Dim strSearch As String
Dim rst As Recordset

strSearch = Trim(txtName)

With gDbTrans
    .SqlStmt = "SELECT CustomerId, " & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name, " & _
        " Profession FROM NameTab "
    If chkLastName.Value = vbChecked Then
        .SqlStmt = .SqlStmt & " Where LastName like '" & strSearch & "%' "
    Else
        .SqlStmt = .SqlStmt & " Where FirstName like " & AddQuotes(strSearch & "%")
    End If
    .SqlStmt = .SqlStmt & " Order by IsciName"
    If .Fetch(rst, adOpenStatic) < 1 Then Set rst = Nothing
    
End With

If rst Is Nothing Then
    MsgBox "No customer found with such name", vbInformation, wis_MESSAGE_TITLE
    cmdSearch.Enabled = True
    Exit Sub
End If

'Search Module
Dim headID As Long
Dim custId As Long
Dim NextRow As Boolean
Dim showRow As Boolean
Dim I As Integer, SlNo As Long
Dim MaxI As Integer
Dim rstBalance As Recordset

'Now Set the Grid
grd.Visible = False
Call InitGrid
grd.Visible = True
grd.Visible = False
MaxI = grd.Cols - 1
grd.Row = grd.FixedRows
Dim colno As Integer, rowno As Long
rowno = grd.Row
While Not rst.EOF
    custId = FormatField(rst("CustomerID"))
    If custId <= 0 Then GoTo NextCustomer
    NextRow = False
    For I = 2 To MaxI
        If grd.Rows <= rowno + 1 Then grd.Rows = rowno + 1
        headID = grd.ColData(I)
        Set rstBalance = GetCustRecordSet(headID, custId)
        If rstBalance Is Nothing Then GoTo NextAccType
        With grd
            '.Col = I
            showRow = True
            showRow = IIf(rstBalance("Balance"), True, False)
            'If rstBalance("Balance") Then
                NextRow = True
                .TextMatrix(rowno, I) = FormatField(rstBalance("AccNum"))
                If chkBalance Then _
                    .TextMatrix(rowno, I + 1) = FormatField(rstBalance("Balance"))
            'End If
            If chkBalance Then I = I + 1
        End With
NextAccType:
    Next
    
    If Not NextRow Then GoTo NextCustomer
    
    With grd
        .RowData(rowno) = custId
        If .Rows = rowno + 1 Then .Rows = rowno + 1
        SlNo = SlNo + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = FormatField(rst("Name"))
        rowno = rowno + 1
    End With
    
NextCustomer:
    rst.MoveNext
Wend

'Now Check For the Coloumn where No Value has entered

For I = 2 To MaxI
    If chkBalance Then I = I + 1
    With grd
        NextRow = False
        .Col = I
        For rowno = .FixedRows To .Rows - 1
            If .TextMatrix(rowno, .Col) <> "" Then NextRow = True: Exit For
        Next
        .ColWidth(I) = IIf(NextRow, 1200, 0)
        If chkBalance Then .ColWidth(I - 1) = IIf(NextRow, 800, 0)
    End With
Next
grd.Visible = True

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdSearch.Caption = GetResourceString(183)
cmdClose.Caption = GetResourceString(11)

lblCustName = GetResourceString(35)
lblAccType = GetResourceString(36, 253)
lblPlace = GetResourceString(99)

chkBalance.Caption = GetResourceString(67, 13)
chkLastName.Caption = GetResourceString(122) '& " " & GetResourceString(13)

Me.Caption = "Search Customers"

End Sub

Private Sub ShowCustomerDetail(ByVal CustomerID As Long, ByVal AccHeadID As Long)


Dim ModuleID As wisModules

ModuleID = GetModuleIDFromHeadID(AccHeadID)
If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
If ModuleID = wis_None Then Exit Sub


'Members
If ModuleID >= wis_Members And ModuleID < wis_Members + 100 Then
    Set m_ClsObject = New clsMMAcc
'Bkcc Account
ElseIf ModuleID = wis_BKCC Or ModuleID = wis_BKCCLoan Then
    Set m_ClsObject = New clsBkcc
    
ElseIf ModuleID = wis_CAAcc Then
    Set m_ClsObject = New ClsCAAcc

'DepositLoans
ElseIf ModuleID >= wis_DepositLoans And ModuleID < wis_DepositLoans + 100 Then
    
    Set m_ClsObject = New clsDepLoan
    
'Pigmy Accounts
ElseIf ModuleID = wis_PDAcc Then
    Set m_ClsObject = New clsPDAcc
    

'Recurring Accounts
ElseIf ModuleID = wis_RDAcc Then
    Set m_ClsObject = New clsRDAcc

'Deposit Accounts like Fd
ElseIf ModuleID >= wis_Deposits And ModuleID < wis_Deposits + 100 Then
    Set m_ClsObject = New clsFDAcc
    
'Loan Accounts
ElseIf ModuleID >= wis_Loans And ModuleID < wis_Loans + 100 Then
    
    Set m_ClsObject = New clsLoan
    
ElseIf ModuleID >= wis_SBAcc And ModuleID < wis_SBAcc Then
    
    Set m_ClsObject = New clsSBAcc

Else
'    MsgBox "Plese select the account type", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

m_ClsObject.CustomerID = CustomerID
If ModuleID >= wis_Loans And ModuleID < wis_Loans + 100 Then
    m_ClsObject.ShowCreateLoanAccount
Else
    m_ClsObject.Show
End If
If gWindowHandle = 0 Then gWindowHandle = Me.hwnd

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()

If Trim(txtName) = "" Then Exit Sub

cmdSearch.Enabled = False
Me.MousePointer = vbHourglass

Call SearchCustomer

grd.Visible = True
Me.MousePointer = vbDefault
cmdSearch.Enabled = True

End Sub

Private Function GetCustRecordSet(AccHeadID As Long, CustomerID As Long) As Recordset

On Error Resume Next

Dim rstReturn As Recordset
Dim SqlStr As String
Dim pos As Long
Dim sqlClause As String
Dim DepType As Integer
Dim subDepType As Integer
Set rstReturn = Nothing
On Error GoTo Exit_Line

Dim ModuleID As wisModules

ModuleID = GetModuleIDFromHeadID(AccHeadID)
DepType = ModuleID Mod 100
subDepType = ModuleID Mod 10

Set rstReturn = Nothing
If ModuleID = wis_None Then Exit Function

'Members
If ModuleID >= wis_Members And ModuleID < wis_Members + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From MemMaster A,MemTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From MemTrans D " & _
            " Where D.AccID = A.AccID ) "
    'If ModuleID Mod 100 > 0 Then _
        SqlStr = SqlStr & " And A.MemberType= " & ModuleID - wis_Members
    If DepType > 0 Then SqlStr = SqlStr & " AND A.MemberType = " & DepType
'Bkcc Account
ElseIf ModuleID = wis_BKCC Then
    SqlStr = "Select AccNum, A.LoanID as ID From BKCCMaster A"
    SqlStr = "Select AccNum, A.LoanID as ID,Balance * -1 as Balance " & _
        " From BKCCMaster A,BKCCTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.LoanID = A.LoanID AND TransID = " & _
            "(Select Max(TransID) From BKCCTrans D " & _
            " Where D.LoanID = A.LoanID And Deposit = True ) "
'Bkcc Loan Account
ElseIf ModuleID = wis_BKCCLoan Then
    SqlStr = "Select AccNum, A.LoanID as ID From BKCCMaster A"
    SqlStr = "Select AccNum, A.LoanID as ID,Balance " & _
        " From BKCCMaster A,BKCCTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.LoanID = A.LoanID AND TransID = " & _
            "(Select Max(TransID) From BKCCTrans D " & _
            " Where D.LoanID = A.LoanID And Deposit = False ) "
'Current Account
ElseIf ModuleID = wis_CAAcc Then
    SqlStr = "Select AccNum, AccId as Id From CAMaster A"
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From CAMaster A,CATrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From CATrans D " & _
            " Where D.AccID = A.AccID ) "
    If DepType > 0 Then SqlStr = SqlStr & " AND A.DepositType = " & DepType
    
'DepositLoans
ElseIf ModuleID >= wis_DepositLoans And ModuleID < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID From DepositLoanMaster A"
    If DepType > 0 Then sqlClause = " AND A.DepositType = " & DepType
    'If ModuleID > wis_DepositLoans Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_DepositLoans
    SqlStr = "Select AccNum, A.LoanID as ID,Balance " & _
        " From DepositLoanMaster A,DepositLoanTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.LoanID = A.LoanID AND TransID = " & _
            "(Select Max(TransID) From DepositLoanTrans D " & _
            " Where D.LoanID = A.LoanID ) "
    If DepType > 0 Then SqlStr = SqlStr & " AND A.DepositType = " & DepType
    'If ModuleID > wis_DepositLoans Then _
        SqlStr = SqlStr & " ANd A.DepositType = " & ModuleID - wis_DepositLoans

'Pigmy Accounts
ElseIf ModuleID >= wis_PDAcc And ModuleID < wis_PDAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID From PDMaster A"
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From PDMaster A,PDTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From PDTrans D " & _
            " Where D.AccID = A.AccID ) "
    If DepType > 0 Then SqlStr = SqlStr & " AND A.DepositType = " & DepType
'Recurring Accounts
ElseIf ModuleID >= wis_RDAcc And ModuleID < wis_RDAcc + 100 Then
    
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From RDMaster A,RDTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From RDTrans D " & _
            " Where D.AccID = A.AccID ) "
    If DepType > 0 Then SqlStr = SqlStr & " AND A.DepositType = " & DepType

'Deposit Accounts like Fd
ElseIf ModuleID >= wis_Deposits And ModuleID < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, AccId as ID From FDMaster A"
    If ModuleID > wis_Deposits Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_Deposits
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From FDMaster A,FDTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From FDTrans D " & _
            " Where D.AccID = A.AccID ) "
    If ModuleID > wis_Deposits Then _
        SqlStr = SqlStr & " ANd A.DepositType = " & ModuleID - wis_Deposits

'Loan Accounts
ElseIf ModuleID >= wis_Loans And ModuleID < wis_Loans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID From LoanMaster A"
    If ModuleID > wis_Loans Then _
        sqlClause = " AND A.SchemeID = " & ModuleID - wis_Loans
    SqlStr = "Select AccNum, A.LoanID as ID,Balance " & _
        " From LoanMaster A,LoanTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.LoanID = A.LOanID AND TransID = " & _
            "(Select Max(TransID) From LoanTrans D " & _
            " Where D.LoanID = A.LoanID ) "
    If ModuleID > wis_Loans Then _
        SqlStr = SqlStr & " AND A.SchemeID = " & ModuleID - wis_Loans

ElseIf ModuleID >= wis_SBAcc And ModuleID < wis_SBAcc + 100 Then
    
    SqlStr = "Select AccNum, A.AccID as ID,Balance " & _
        " From SBMaster A,SBTrans B" & _
        " Where A.CustomerID = " & CustomerID & _
        " AND B.AccID = A.AccID AND TransID = " & _
            "(Select Max(TransID) From SBTrans D " & _
            " Where D.AccID = A.AccID ) "
    If DepType > 0 Then SqlStr = SqlStr & " AND A.DepositType = " & DepType
    
Else
'    MsgBox "Plese select the account type", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If
    
    SqlStr = Trim(SqlStr)
    gDbTrans.SqlStmt = SqlStr

    If gDbTrans.Fetch(rstReturn, adOpenStatic) < 1 Then Exit Function
'        MsgBox "There are no customers in the " & _
            GetHeadName(AccHeadId), vbInformation, wis_MESSAGE_TITLE
'        Exit Function
'    End If

Exit_Line:

Set GetCustRecordSet = rstReturn

End Function


Private Sub Form_Load()
Dim rst As Recordset
Dim headID As Long

Call CenterMe(Me)

Call SetKannadaCaption

'Call LoadHeadsToCombo(cmbAccType,0)
With cmbAccType
    .Clear
    .AddItem ""
    '338 =all,
    'Memeber
    gDbTrans.SqlStmt = "Select * From MemberTypeTab order by membertype"
        
    If gDbTrans.Fetch(rst, adOpenDynamic) > 1 Then
        While rst.EOF = False
            .AddItem FormatField(rst("MemberTypeName"))
            headID = GetIndexHeadID(FormatField(rst("MemberTypeName")))
            If headID = 0 Then headID = GetIndexHeadID(FormatField(rst("MemberTypeName")) & " " & GetResourceString(53, 36))
            .ItemData(.newIndex) = headID
            'Move to next record
            rst.MoveNext
        Wend
    Else
        .AddItem GetResourceString(53, 36) '& GetResourceString(92)
        .ItemData(.newIndex) = GetIndexHeadID(GetResourceString(53, 36)) ', parMemberShare)
    End If
    'Member End
    
    .AddItem GetResourceString(338, 42) & GetResourceString(92)
    .ItemData(.newIndex) = parMemberDeposit
    Call LoadLedgersToCombo(cmbAccType, parMemberDeposit, False)
    
    .AddItem GetResourceString(338, 58) & _
            GetResourceString(36) & GetResourceString(92)
    .ItemData(.newIndex) = parMemberLoan
    Call LoadLedgersToCombo(cmbAccType, parMemberLoan, False)
    
    .AddItem GetResourceString(338) & GetResourceString(43, 58) & _
            GetResourceString(36) & GetResourceString(92)
    .ItemData(.newIndex) = parMemDepLoan
    Call LoadLedgersToCombo(cmbAccType, parMemDepLoan, False)
    
End With

Call LoadPlaces(cmbPlace)

End Sub

Private Sub Form_Resize()

Dim Margin As Single
Dim CmdPos As Single

Margin = 100
If frmCustSearch.WindowState = vbMinimized Then Exit Sub

If frmCustSearch.Height < m_Minheight Then frmCustSearch.Height = m_Minheight

If frmCustSearch.Width < m_MinWidth Then frmCustSearch.Width = m_MinWidth

cmdSearch.Left = frmCustSearch.Width - ((Margin * 2) + cmdSearch.Width)
cmdClose.Left = frmCustSearch.Width - ((Margin * 2) + cmdSearch.Width)

fraName.Width = cmdSearch.Left - fraName.Left - Margin * 2
txtName.Width = (fraName.Width - txtName.Left - Margin * 2) / 2
cmbAccType.Width = txtName.Width 'fraName.Width - Margin * 2
cmbPlace.Width = (fraName.Width - cmbPlace.Left - Margin * 2) / 2
chkBalance.Left = cmbPlace.Width + cmbPlace.Left + Margin * 2
chkBalance.Width = cmbPlace.Width - Margin * 2
chkLastName.Width = chkBalance.Width
chkLastName.Left = chkBalance.Left

With grd
    .Width = Width - Margin * 2
    .Height = Height - grd.Top - Margin * 4
End With

End Sub

Private Sub Form_Unload(cancel As Integer)
Set frmCustSearch = Nothing
End Sub

Private Sub grd_DblClick()
With grd
    If .RowData(.Row) = 0 Or .ColData(.Col) = 0 Then Exit Sub
    Call ShowCustomerDetail(.RowData(.Row), .ColData(.Col))
End With


End Sub


