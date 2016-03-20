VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLoanInst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installment details"
   ClientHeight    =   6195
   ClientLeft      =   1695
   ClientTop       =   1515
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   400
      Left            =   4140
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5580
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame fraInstall 
      Caption         =   "Loan Installment"
      Height          =   4065
      Left            =   150
      TabIndex        =   0
      Top             =   1530
      Width           =   6645
      Begin VB.TextBox txtFstInstDate 
         Height          =   315
         Left            =   1590
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   400
         Left            =   5130
         TabIndex        =   12
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CheckBox chkEMI 
         Caption         =   "EMI"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3450
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtgrd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4650
         TabIndex        =   20
         Top             =   1980
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtNoOfINst 
         Height          =   315
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbInstType 
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdInst 
         Height          =   2205
         Left            =   150
         TabIndex        =   13
         Top             =   1650
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   3889
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         BackColor       =   -2147483628
         ForeColorSel    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFstInstDate 
         Caption         =   "&First Inst Date :"
         Height          =   300
         Left            =   150
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label txtIssueDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label lblIssueDate 
         AutoSize        =   -1  'True
         Caption         =   "Issued On :"
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label txtLoanAmount 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5130
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblLoanAmount 
         AutoSize        =   -1  'True
         Caption         =   "Loan &amount :"
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label lblInstType 
         Caption         =   "Inst &Mode"
         Height          =   300
         Left            =   270
         TabIndex        =   5
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label lblNoOfInst 
         Caption         =   "&Installments"
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   810
         Width           =   1035
      End
   End
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer"
      Height          =   1305
      Left            =   150
      TabIndex        =   16
      Top             =   210
      Width           =   6645
      Begin VB.ComboBox cmbLoanScheme 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   810
         Width           =   4785
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1710
         TabIndex        =   21
         Top             =   330
         Width           =   4785
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCustName 
         Caption         =   "Customer &Name :"
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label lblLoanScheme 
         Caption         =   "&Loan Scheme :"
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   810
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmLoanInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dim m_SchemeId  as  Integer
Public Event OkClicked(InstIndianDates() As String, InstAmounts() As Currency)
Public Event CancelClicked()
Dim m_SchemeId As Integer
Private m_LoanID As Integer
Public m_IntRate As Double

Public Enum FormStatus
   InstRepay = 1
   InstInsert = 2
End Enum

Public Operation As FormStatus
Public Event WindowClosed()


Private Sub GridInit(Optional rowno As Integer)

txtgrd.Visible = False
grdInst.Clear
grdInst.Rows = 3
grdInst.Cols = 3
grdInst.FixedRows = 0: grdInst.FixedCols = 0

If Operation = InstRepay Then grdInst.Cols = 5
If rowno Then grdInst.Rows = rowno
grdInst.Row = 0
grdInst.Col = 0: txtgrd = "Inst No": grdInst.CellFontBold = True
grdInst.Col = 1: txtgrd = "Inst Date": grdInst.CellFontBold = True
grdInst.Col = 2: txtgrd = "Inst Amount": grdInst.CellFontBold = True
If Operation = InstRepay Then
   grdInst.Col = 3: txtgrd = "Paid Date": grdInst.CellFontBold = True
   grdInst.Col = 4: txtgrd = "Balance": grdInst.CellFontBold = True
End If
grdInst.Row = 1
grdInst.FixedRows = 1
grdInst.FixedCols = 1
txtgrd.Visible = True

End Sub


Public Property Let InterestRate(NewValue As Single)
m_IntRate = NewValue
End Property

Private Sub LoadInstallmentDetails()

Dim LoanAmount As Currency
Dim InstAmount As Currency
Dim NoOfInst As Integer
Dim InstType As wisInstallmentTypes
Dim SqlStr As String
Dim rst As Recordset

NoOfInst = Val(txtNoOfINst)
LoanAmount = Val(txtLoanAmount)

txtFstInstDate.Locked = False
'Get the Loan schemeid * Installment type
SqlStr = "SELECT InstMode,SchemeID,EMI,IntRate,NoOfINstall FROM LoanMaster WHERE " & _
        " LoanID = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    txtFstInstDate.Locked = True
    InstType = FormatField(rst("Instmode"))
    txtNoOfINst = FormatField(rst("NoOfInstall"))
    m_IntRate = FormatField(rst("IntRate"))
    If FormatField(rst("EMI")) Then chkEMI.Value = vbChecked
    m_SchemeId = FormatField(rst("SchemeID"))
    For NoOfInst = 0 To cmbLoanScheme.ListCount - 1
      If cmbLoanScheme.ItemData(NoOfInst) = m_SchemeId Then
        cmbLoanScheme.ListIndex = NoOfInst
        Exit For
      End If
    Next NoOfInst
    For NoOfInst = 0 To cmbInstType.ListCount - 1
      If cmbInstType.ItemData(NoOfInst) = InstType Then
        cmbInstType.ListIndex = NoOfInst
        Exit For
      End If
    Next NoOfInst
Else
    Set rst = Nothing
End If

SqlStr = "SELECT * FROM LoanInst WHERE LoanID = " & m_LoanID & _
        " ORDER BY InstNo"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    txtFstInstDate = FormatField(rst("InstDate"))
    NoOfInst = rst.RecordCount
    txtNoOfINst = NoOfInst
Else
    Set rst = Nothing
End If

If Not rst Is Nothing Then
    'Load The Installment Details to grid & exit sub
    grdInst.Visible = False
    grdInst.Row = 0
    txtgrd.Visible = True
    While Not rst.EOF
      With grdInst
        If .Rows < .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 0: txtgrd.Text = rst("InstNo"): .Text = txtgrd
        grdInst.Col = 1: txtgrd.Text = FormatField(rst("InstDate"))
        grdInst.Col = 2: txtgrd.Text = FormatField(rst("InstAmount"))
        If Me.Operation = InstRepay Then
            grdInst.Col = 3: txtgrd.Text = FormatField(rst("PaidDate"))
            grdInst.Col = 4: txtgrd.Text = FormatField(rst("InstBalance"))
        End If
      End With
      rst.MoveNext
    Wend
    If grdInst.Rows > 6 Then grdInst.ScrollBars = flexScrollBarBoth
    grdInst.Visible = True
    Exit Sub
End If

If Val(txtLoanAmount) = 0 Then Exit Sub
If Not DateValidate(txtIssueDate, "/", True) Then Exit Sub
If Me.Operation = InstInsert Then
  Dim FortNight As Boolean
  Dim InstNo As Integer
  Dim TotalInstAmount As Currency
  Dim NextDate As Date
    
    If NoOfInst = 0 Then Exit Sub
    InstType = cmbInstType.ItemData(cmbInstType.ListIndex)
    NextDate = GetSysFormatDate(txtIssueDate)
    InstAmount = (LoanAmount / NoOfInst) \ 1
    If chkEMI.Value = vbChecked Then
        Dim Div As Byte
        Div = 0
        If InstType = Inst_Weekly Then Div = 52
        If InstType = Inst_FortNightly Then Div = 26
        If InstType = Inst_Monthly Then Div = 12
        If InstType = Inst_BiMonthly Then Div = 6
        If InstType = Inst_Quartery Then Div = 4
        If InstType = Inst_HalfYearly Then Div = 2
        If InstType = Inst_Yearly Then Div = 1
        
        m_IntRate = m_IntRate / 100
        InstAmount = Pmt(m_IntRate / Div, NoOfInst, -LoanAmount)
    End If
    InstAmount = InstAmount \ 1
    InstNo = 1: FortNight = True
    grdInst.Rows = 1
    grdInst.Visible = False
    Do
        If InstNo > NoOfInst Then Exit Do
        If TotalInstAmount >= LoanAmount And chkEMI.Value = vbUnchecked Then Exit Do
        'Get The Next INstallment date
        If InstType = Inst_Daily Then NextDate = DateAdd("d", 1, NextDate)
        If InstType = Inst_Weekly Then NextDate = DateAdd("WW", 1, NextDate)
        If InstType = Inst_FortNightly Then
            If FortNight Then
                FortNight = False
                NextDate = DateAdd("d", 15, NextDate)
            Else
                FortNight = True
                NextDate = DateAdd("d", -15, NextDate)
                NextDate = DateAdd("m", 1, NextDate)
            End If
        End If
        If InstType = Inst_Monthly Then NextDate = DateAdd("M", 1, NextDate)
        If InstType = Inst_BiMonthly Then NextDate = DateAdd("m", 2, NextDate)
        If InstType = Inst_Quartery Then NextDate = DateAdd("q", 1, NextDate)
        If InstType = Inst_HalfYearly Then
            If FortNight Then
                FortNight = False
                NextDate = DateAdd("M", 6, NextDate)
            Else
                FortNight = True
                NextDate = DateAdd("M", -6, NextDate)
                NextDate = DateAdd("YYYY", 1, NextDate)
            End If
        End If
        If InstType = Inst_Yearly Then NextDate = DateAdd("YYYY", 1, NextDate)
        If InstNo = 1 Then
            If DateValidate(txtFstInstDate, "/", True) Then
                NextDate = GetSysFormatDate(txtFstInstDate)
            Else
                txtFstInstDate = GetIndianDate(NextDate)
            End If
        End If
        With grdInst
            'WRITE Into the Grid
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0: txtgrd = Format(InstNo, "00")
            .Col = 1: txtgrd = GetIndianDate(NextDate)
            .Col = 2: txtgrd = FormatCurrency(InstAmount)
        End With
        TotalInstAmount = TotalInstAmount + InstAmount
        InstNo = InstNo + 1
    Loop
    
    grdInst.Visible = True
End If

End Sub

Public Property Let LoanID(NewValue As Long)
m_LoanID = NewValue
End Property

Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

'Now Set the caption for all Controls
fraCustomer.Caption = GetResourceString(205) 'Customer
lblCustName.Caption = GetResourceString(205, 35) 'Customer Name
lblLoanScheme.Caption = GetResourceString(214) 'Loan Scheme

fraInstall.Caption = GetResourceString(57, 295)
lblIssueDate = GetResourceString(340) 'Issued On
lblInstType = GetResourceString(57) 'INstallment
lblFstInstDate = GetResourceString(31, 57) 'First Installment
lblLoanAmount = GetResourceString(80, 91) 'LOan Amount
lblNoOfInst = GetResourceString(55)

cmdRefresh.Caption = GetResourceString(32)  'Refresh
cmdOk.Caption = GetResourceString(1)         'OK
cmdCancel.Caption = GetResourceString(2)     'Cancel


End Sub

Private Sub cmdCancel_Click()
   RaiseEvent CancelClicked
   Me.Hide

End Sub


Private Sub cmdOk_Click()
Dim StrRets() As String
Dim InstAmount() As Currency
Dim InstIndDate() As String
Dim InstBalance() As Currency
Dim PrevDate As String
Dim count As Integer

ReDim InstAmount(grdInst.Rows - 2)
ReDim InstIndDate(grdInst.Rows - 2)
ReDim InstBalance(grdInst.Rows - 2)
On Error Resume Next
grdInst.Visible = True
For count = 1 To grdInst.Rows - 1
    grdInst.Row = count
    'Validation Of the Date  & Amount
    grdInst.Col = 1
    'Validate the Prev date with currnwt date
    'dates should be in a ascending
    InstIndDate(count - 1) = grdInst.Text
    
    grdInst.Col = 2
    'Validate the currency
    InstAmount(count - 1) = grdInst.Text
    InstBalance(count - 1) = InstAmount(count - 1)
    If count = 1 Then
        If Val(InstAmount(count - 1)) = 0 Then
            MsgBox "you have not entered the amount", vbOKOnly, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
Next
grdInst.Visible = True

RaiseEvent OkClicked(InstIndDate(), InstAmount())
Me.Hide
End Sub


Private Sub cmdRefresh_Click()
If Not DateValidate(txtFstInstDate, "/", True) Then
    'MsgBox "Invalid date specifried", vbInformation, wis_MESSAGE_TITLE
     MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtFstInstDate
    Exit Sub
End If

Call LoadInstallmentDetails

End Sub

Private Sub Form_Activate()
Call LoadInstallmentDetails
End Sub

Private Sub Form_Load()
'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption

txtCustName.FONTSIZE = txtCustName.FONTSIZE + 1
Call LoadLoanSchemes(cmbLoanScheme)
Call GridInit(3)

cmdRefresh.Enabled = False

'Load The instalment types
Dim InstType As wisInstallmentTypes
With cmbInstType
    InstType = Inst_No
    .AddItem ""
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Daily
    .AddItem GetResourceString(410) '"Daily"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Weekly
    .AddItem GetResourceString(411) '"Weekly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_FortNightly
    .AddItem GetResourceString(412) '"Fortnightly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Monthly
    .AddItem GetResourceString(463) '"Monthly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_BiMonthly
    .AddItem GetResourceString(413) '"Bi-Monthly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Quartery
    .AddItem GetResourceString(414) '"Quarterly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_HalfYearly
    .AddItem "6 " & GetResourceString(463) '"Half Yearly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Yearly
    .AddItem "1 " & GetResourceString(208) '"Yearly"
    .ItemData(cmbInstType.newIndex) = InstType
    
End With


End Sub


Private Sub grdInst_Click()
Call grdInst_EnterCell
End Sub

Private Sub grdInst_EnterCell()
   'If Me.ActiveControl.Name <> grdInst.Name Then Exit Sub
With grdInst
   txtgrd.Text = .Text
   txtgrd.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
   txtgrd.Visible = .RowIsVisible(.Row)
End With
   On Error Resume Next
   ActivateTextBox txtgrd
   'txtgrd.SetFocus
   Err.Clear
End Sub


Private Sub grdInst_GotFocus()
    Call grdInst_EnterCell
End Sub

Private Sub grdInst_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
'      SendKeys "{RIGHT}"
   End If
End Sub


Private Sub grdInst_LeaveCell()

With txtgrd
   If .Visible = False Then
     '  Exit Sub
   End If
   grdInst.Text = .Text
   .Visible = False
   .Text = ""
End With

End Sub

Private Sub grdInst_Scroll()
With grdInst
    txtgrd.Visible = .RowIsVisible(.Row)
    If txtgrd.Visible Then txtgrd.Visible = .RowIsVisible(.Row)
    If Not txtgrd.Visible Then Exit Sub
    txtgrd.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) ', grdInst.CellWidth, grdInst.CellHeight
End With
'txtgrd.Move grdInst.Left + grdInst.CellLeft, grdInst.Top + grdInst.CellTop, grdInst.CellWidth, grdInst.CellHeight
'If txtgrd.Left < grdInst.Left + grdInst.ColPos(0) = txtgrd.Width Then txtgrd.Visible = False
'If txtgrd.Top < grdInst.Top + grdInst.RowPos(0) + txtgrd.Height Then txtgrd.Visible = False
End Sub

Private Sub txtFstInstDate_LostFocus()

On Error Resume Next
If DateValidate(txtFstInstDate, "/", True) Then
    With grdInst
        .Row = 1
        .Col = 1
        cmdRefresh.Enabled = IIf(.Text <> txtFstInstDate.Text, True, False)
    End With
End If

End Sub

Private Sub txtgrd_GotFocus()
txtgrd.SelStart = 0
txtgrd.SelLength = Len(txtgrd)
End Sub

Private Sub txtgrd_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

With grdInst
    
    If KeyCode = 37 Then 'Press Left Arrow
        If txtgrd.SelStart = 0 And .Col > .FixedCols Then .Col = .Col - 1
    
    ElseIf KeyCode = 39 Then 'Press Right Arrow
        If txtgrd.SelStart = Len(txtgrd.Text) And .Col < .Cols - 1 Then .Col = .Col + 1
    
    ElseIf KeyCode = 38 Then  'Press UpArrow
        If .Row > .FixedRows Then .Row = .Row - 1
    
    ElseIf KeyCode = 40 Then ' Press Doun Arroow
        If .Row < .Rows - 1 Then .Row = .Row + 1
    
    ElseIf KeyCode = 33 Then  'Press PageUp
        If .Row > 0 Then .Row = .Row - 1
        
    ElseIf KeyCode = 34 Then ' Press PageDown
        If .Row < .Rows - 1 Then .Row = .Row + 1
    
    End If
    
End With

End Sub

Private Sub txtgrd_LostFocus()
    grdInst.Text = txtgrd.Text
    txtgrd.Visible = False
    'grdInst.SetFocus
End Sub

Private Sub txtNoOfINst_LostFocus()
Dim NoOfInst As Integer

If Not IsNumeric(txtNoOfINst) Then
   MsgBox "Please enter Numeric values", vbInformation, wis_MESSAGE_TITLE
   Exit Sub
End If
NoOfInst = Val(txtNoOfINst)
If NoOfInst > 0 Then
    grdInst.Rows = NoOfInst + 1
End If
End Sub

'This function will load the loan schemes
' into the cmbobox

Private Sub GridResize()

Dim sngColWidth As Single

sngColWidth = grdInst.Width / grdInst.Cols

With grdInst
   .ColWidth(0) = sngColWidth * 0.5
   .ColWidth(1) = sngColWidth * 1.5
   .ColWidth(2) = sngColWidth * 1
   If Operation = InstRepay Then
    .ColWidth(3) = sngColWidth * 1.5
    .ColWidth(4) = sngColWidth * 1
   End If
End With

End Sub


