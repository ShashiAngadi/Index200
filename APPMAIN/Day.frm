VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmDayBegin 
   Caption         =   "Begin/End Day"
   ClientHeight    =   3000
   ClientLeft      =   2970
   ClientTop       =   2340
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   4650
      TabIndex        =   19
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Frame fraEndDay 
      Caption         =   "Frame1"
      Height          =   2205
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   5775
      Begin VB.CommandButton cmdTrans 
         Caption         =   "Transaction"
         Height          =   435
         Left            =   3300
         TabIndex        =   16
         Top             =   1650
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeposit 
         Caption         =   "Deposit"
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   1650
         Width           =   2295
      End
      Begin VB.TextBox txtClTransDate 
         Height          =   345
         Left            =   3840
         TabIndex        =   9
         Top             =   180
         Width           =   1335
      End
      Begin WIS_Currency_Text_Box.CurrText txtClStock 
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   960
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtCLBalance 
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   570
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblClBalance 
         Caption         =   "Closing Balance"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   690
         Width           =   3255
      End
      Begin VB.Label lblClStock 
         Caption         =   "Closing Stock Value"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1050
         Width           =   3285
      End
      Begin VB.Label lblTransDate 
         Caption         =   " Transaction Date"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   3330
      TabIndex        =   17
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Frame fraBegin 
      Caption         =   "Begin The Day"
      Height          =   2205
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5775
      Begin VB.CommandButton cmdMatDeposit 
         Caption         =   "&Deposit"
         Height          =   375
         Left            =   3780
         TabIndex        =   18
         Top             =   1650
         Width           =   1785
      End
      Begin VB.CommandButton cmdTransDate 
         Caption         =   ".."
         Height          =   315
         Left            =   5190
         TabIndex        =   7
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txtBeginDate 
         Height          =   345
         Left            =   3810
         TabIndex        =   6
         Top             =   210
         Width           =   1335
      End
      Begin WIS_Currency_Text_Box.CurrText txtOpStock 
         Height          =   375
         Left            =   3810
         TabIndex        =   4
         Top             =   1110
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtOpBalance 
         Height          =   375
         Left            =   3810
         TabIndex        =   3
         Top             =   630
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lbOplTransDate 
         Caption         =   "Transaction Date"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label lblOpStock 
         Caption         =   "Opening Stock Value"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1170
         Width           =   3285
      End
      Begin VB.Label lblOpBalance 
         Caption         =   "Opening Balance"
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   750
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmDayBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_retVar As Variant
Private WithEvents m_frmGrid As frmGrid
Attribute m_frmGrid.VB_VarHelpID = -1
Private FDChecked As Boolean
Private LoanChecked As Boolean
Private TransChecked As Boolean
Private LedgerPosted As Boolean
Private DayBegun As Boolean

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

cmdOk.Caption = GetResourceString(11)
End Sub


Public Sub ShowDayEnd()
    fraEndDay.Visible = True
    fraBegin.Visible = False
    fraEndDay.ZOrder 0
    Show 1
End Sub

Public Sub ShowDayBegin()
    Dim ObDate As Date
    
    fraEndDay.Visible = False
    fraBegin.Visible = True
    fraBegin.ZOrder 0
    'get the last transaction date
    ObDate = GetSysFormatDate(GetLastTransactionDate)
    ObDate = DateAdd("D", 1, ObDate)
    txtBeginDate = GetIndianDate(ObDate)
    
    'Get the Opening Balance
    Dim TransClass As Object
    
    Set TransClass = New clsAccTrans
    txtOpBalance = TransClass.GetOpBalance(wis_CashHeadID, ObDate)
    
    'get the opening stock value
    Set TransClass = New clsMaterial
    txtOpStock = TransClass.GetOnDateClosingStockValue(DateAdd("d", -1, ObDate))
    
    Set TransClass = Nothing
    
    Show 1
    
End Sub

Private Function GetLastTransactionDate() As String
'Declare the Variables
Dim rstLastTransDate As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

GetLastTransactionDate = FinIndianFromDate 'FinFromDate
'Fetch the data from the database
gDbTrans.SqlStmt = " SELECT MAX(TransDate)" & _
                                " FROM AccTrans"
If gDbTrans.Fetch(rstLastTransDate, adOpenForwardOnly) < 1 Then Exit Function

'return the data
GetLastTransactionDate = FormatField(rstLastTransDate.Fields(0))

Exit Function

ErrLine:
    MsgBox "GetLastTransactionDate: " & Err.Description, vbCritical
    

End Function



Private Sub ShowTransaction()
'This Shows the Transaction Details as On Today
If m_frmGrid Is Nothing Then Set m_frmGrid = New frmGrid
Load m_frmGrid
m_frmGrid.Tag = "Trans"

With m_frmGrid.grd
    .Clear
    .Cols = 6: .Rows = 3
    .FixedCols = 1: .FixedRows = 2: .MergeCells = flexMergeRestrictAll
    .Row = 0: .MergeRow(0) = True
    .Col = 0: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(33)
    .Col = 1: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(36, 253)
    .Col = 2: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(271)
    .Col = 3: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(271)
    .Col = 4: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(272)
    .Col = 5: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(272)
    
    .Row = 1: .MergeRow(1) = True
    .Col = 0: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(33)
    .Col = 1: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(36, 253)
    .Col = 2:: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(269)
    .Col = 3:: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(270)
    .Col = 4: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(269)
    .Col = 5: .CellAlignment = 4: .CellFontBold = True
    .Text = GetResourceString(270)
    
    .ColWidth(0) = 600: .MergeCol(0) = True
    .ColWidth(1) = 2000: .MergeCol(1) = True
    .ColWidth(2) = 1000: .MergeCol(2) = True
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    .Rows = .FixedRows + 2
    .Row = .FixedRows - 1: .Col = 1
End With

'Now Consider each Head
'First Saving Bank Account
Dim CashReceipt As Currency, CashPayment As Currency
Dim ContraReceipt As Currency, ContraPayment As Currency
Dim showRow As Boolean, SlNo As Integer
Dim ClsObject As Object, rstTemp As Recordset
Dim TransDate As Date

TransDate = CDate(gStrDate)
'Member Share
Set ClsObject = New clsMMAcc
Dim MemClass As clsMMAcc
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetResourceString(420)
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 3: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 4: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If
Set ClsObject = Nothing

'Saving Account
Set ClsObject = New clsSBAcc
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetResourceString(421)
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If
Set ClsObject = Nothing

'Current Account
Set ClsObject = New ClsCAAcc
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetResourceString(422)
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If
Set ClsObject = Nothing

'Deposits
gDbTrans.SqlStmt = "SELECT * From DepositName"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
  Set ClsObject = New clsFDAcc
  Do
    CashReceipt = ClsObject.CashDeposits(TransDate, TransDate, rstTemp("DepositID"))
    CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate, rstTemp("DepositID"))
    ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate, rstTemp("DepositID"))
    ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate, rstTemp("DepositID"))
    showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
    If showRow Then
        With m_frmGrid.grd
            .Rows = .Rows + 1
            .Row = SlNo + .FixedRows: SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstTemp("DepositNAme"))
            If CashReceipt Then _
                .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
            If CashPayment Then _
                .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
            If ContraReceipt Then _
                .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
            If ContraPayment Then _
                .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
        End With
    End If
    
  rstTemp.MoveNext
  If rstTemp.EOF Then Exit Do
  Loop

  Set ClsObject = Nothing
End If

'Recurring Deposit Account
Set ClsObject = New clsRDAcc
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetResourceString(424)
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If
Set ClsObject = Nothing

'Pigmy Account
Set ClsObject = New clsPDAcc
Dim PDClass As clsPDAcc
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = GetResourceString(425)
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If
Set ClsObject = Nothing

'Deposits
gDbTrans.SqlStmt = "SELECT * From LoanScheme"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
  Set ClsObject = New clsLoan
  Do
    CashReceipt = ClsObject.CashDeposits(TransDate, TransDate, , rstTemp("Schemeid"))
    CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate, , rstTemp("SchemeID"))
    ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate, , rstTemp("SchemeID"))
    ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate, , rstTemp("SchemeID"))
    showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
    If showRow Then
        With m_frmGrid.grd
            .Rows = .Rows + 1
            .Row = SlNo + .FixedRows: SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstTemp("SchemeName"))
            If CashReceipt Then _
                .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
            If CashPayment Then _
                .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
            If ContraReceipt Then _
                .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
            If ContraPayment Then _
                .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
        End With
    End If
    
  rstTemp.MoveNext
  If rstTemp.EOF Then Exit Do
  Loop
End If
'Deposits Loan
Set ClsObject = New clsDepLoan
Dim DepLnClass As clsDepLoan
CashReceipt = ClsObject.CashDeposits(TransDate, TransDate, wisDeposit_PD)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate, wisDeposit_PD)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate, wisDeposit_PD)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate, wisDeposit_PD)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstTemp("SchemeName"))
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If

CashReceipt = ClsObject.CashDeposits(TransDate, TransDate, wisDeposit_RD)
CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate, wisDeposit_RD)
ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate, wisDeposit_RD)
ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate, wisDeposit_RD)
showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
If showRow Then
    With m_frmGrid.grd
        .Rows = .Rows + 1
        .Row = SlNo + .FixedRows: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstTemp("SchemeName"))
        If CashReceipt Then _
            .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
        If CashPayment Then _
            .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
        If ContraReceipt Then _
            .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
        If ContraPayment Then _
            .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
    End With
End If

gDbTrans.SqlStmt = "SELECT * From DepositName"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
  Do
    CashReceipt = ClsObject.CashDeposits(TransDate, TransDate, rstTemp("DepositID"))
    CashPayment = ClsObject.CashWithdrawls(TransDate, TransDate, rstTemp("DepositID"))
    ContraReceipt = ClsObject.ContraDeposits(TransDate, TransDate, rstTemp("DepositID"))
    ContraPayment = ClsObject.ContraWithdrawls(TransDate, TransDate, rstTemp("DepositID"))
    showRow = (CashReceipt Or CashPayment Or ContraReceipt Or ContraPayment)
    If showRow Then
        With m_frmGrid.grd
            .Rows = .Rows + 1
            .Row = SlNo + .FixedRows: SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstTemp("SchemeName"))
            If CashReceipt Then _
                .Col = 2: .Text = FormatCurrency(CashReceipt): CashReceipt = 0
            If CashPayment Then _
                .Col = 4: .Text = FormatCurrency(CashPayment): CashPayment = 0
            If ContraReceipt Then _
                .Col = 3: .Text = FormatCurrency(ContraReceipt): ContraReceipt = 0
            If ContraPayment Then _
                .Col = 5: .Text = FormatCurrency(ContraPayment): ContraPayment = 0
        End With
    End If
    
  rstTemp.MoveNext
  If rstTemp.EOF Then Exit Do
  Loop

  Set ClsObject = Nothing
End If

'Now Interest Ofthe Above Heads

m_frmGrid.Show 1


End Sub

Private Sub cmdBegin_Click()

End Sub

Private Sub cmdCancel_Click()

Unload Me
End
End Sub

Private Sub cmdDeposit_Click()

Dim rstDeposit As Recordset

gDbTrans.SqlStmt = "Select AccId,AccNum,DepositAmount, " & _
        " Title+' '+FirstName+' '+MiddleName+' '+LastName as Name " & _
        " From FDMaster A,NameTab B" & _
        " Where ClosedDate is NULL And MaturedOn is NULL" & _
        " And MaturityDate <= #" & gStrDate & "#" & _
        " And B.CustomerID = A.CustomerId" & _
        " Order By DepositType,MaturityDate,val(AccNum)"

If gDbTrans.Fetch(rstDeposit, adOpenDynamic) < 1 Then _
        FDChecked = True: Exit Sub

If m_frmGrid Is Nothing Then Set m_frmGrid = New frmGrid
Load m_frmGrid
With m_frmGrid.grd
    .Clear
    .Cols = 5
    .FixedCols = 2
    .FixedRows = 1
    .Rows = rstDeposit.RecordCount + 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36, 60)
    .Col = 2: .Text = GetResourceString(35)
    .Col = 3: .Text = GetResourceString(43, 40)
    .AllowUserResizing = flexResizeRows
    '.RowHeightMin = 350
    .Col = 4: .Text = GetResourceString(46)
    .ColWidth(0) = 600
    .ColWidth(1) = 750
    .ColWidth(2) = 2500
    .ColWidth(3) = 1150
    .ColWidth(4) = 550
    m_frmGrid.SelectionColoumn = .Cols - 1
    
    Dim SlNo As Integer
    While Not rstDeposit.EOF
        If .Rows = .Row + 1 Then .Rows = 1
        .Row = .Row + 1: SlNo = SlNo + 1
        .RowData(.Row) = rstDeposit("AccID")
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstDeposit("AccNum"))
        .Col = 2: .Text = FormatField(rstDeposit("Name"))
        .Col = 3: .Text = FormatField(rstDeposit("DepositAmount"))
        m_frmGrid.Selected(.Row) = True
        rstDeposit.MoveNext
    Wend
    
End With

m_retVar = ""
m_frmGrid.Show 1
FDChecked = True

Dim MaxAcc As Integer
MaxAcc = rstDeposit.RecordCount
Set rstDeposit = Nothing

If Not m_retVar Then
    Unload m_frmGrid
    Set m_frmGrid = Nothing
    Exit Sub
End If

'Now Show the Fd Closing Form
SlNo = 1
Dim AccId As Long

For SlNo = 1 To MaxAcc
    With m_frmGrid
        If Not .Selected(SlNo) Then GoTo NextDeposit
        AccId = .grd.RowData(SlNo)
    End With
    
    With frmFdClose
        .AccountId = AccId
        .optClose.Enabled = False
        .Show 1
    End With
    'Unload frmFdClose
    
NextDeposit:

Next

Unload frmFdClose
Set frmFdClose = Nothing

End Sub

Private Sub cmdLedger_Click()

End Sub

Private Sub cmdMatDeposit_Click()

If m_frmGrid Is Nothing Then Set m_frmGrid = New frmGrid
Load m_frmGrid

MousePointer = vbHourglass

With m_frmGrid.grd
    .Clear
    .Cols = 6
    .Rows = 5
    .FixedCols = 2
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(36, 60)
    .Col = 2: .Text = GetResourceString(35)
    .Col = 3: .Text = GetResourceString(43, 40)
    .AllowUserResizing = flexResizeRows
    '.RowHeightMin = 350
    .Col = 4: .Text = GetResourceString(47)
    .Col = 5: .Text = GetResourceString(46, 40)
    
    .ColWidth(0) = 600
    .ColWidth(1) = 750
    .ColWidth(2) = 2500
    .ColWidth(3) = 1150
    .ColWidth(4) = 550
    .ColWidth(5) = 1250
End With

    'm_frmGrid.SelectionColoumn = .Cols - 1
    Dim SlNo As Integer
    Dim IntAmount As Currency
    Dim DepAmount As Currency
    Dim TotalDepAmount As Currency
    Dim TotalIntAmount As Currency

Dim rstDeposit As Recordset
gDbTrans.SqlStmt = "Select AccId,AccNum,DepositAmount, " & _
        " Title+' '+FirstName+' '+MiddleName+' '+LastName as Name " & _
        " From FDMaster A,NameTab B" & _
        " Where ClosedDate is NULL And MaturedOn is NULL" & _
        " And MaturityDate <= #" & gStrDate & "#" & _
        " And B.CustomerID = A.CustomerId" & _
        " Order By DepositType,MaturityDate,val(AccNum)"

If gDbTrans.Fetch(rstDeposit, adOpenDynamic) < 1 Then GoTo RDDEPOSIT
FDChecked = True


Dim FdClass As New clsFDAcc
With m_frmGrid.grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(423)
    .CellFontBold = True
    While Not rstDeposit.EOF
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        .RowData(.Row) = rstDeposit("AccID")
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstDeposit("AccNUm"))
        .Col = 2: .Text = FormatField(rstDeposit("Name"))
        DepAmount = FormatField(rstDeposit("DepositAmount"))
        .Col = 3: .Text = DepAmount
        IntAmount = FdClass.InterestAmount(rstDeposit("AccID"))
        .Col = 4: .Text = FormatCurrency(IntAmount)
        .Col = 5: .Text = FormatCurrency(DepAmount + IntAmount)
        TotalDepAmount = TotalDepAmount + DepAmount
        TotalIntAmount = TotalIntAmount + IntAmount
        'm_frmGrid.Selected(.Row) = True
        rstDeposit.MoveNext
    Wend
    
End With
Set FdClass = Nothing

RDDEPOSIT:

gDbTrans.SqlStmt = "Select A.AccId,AccNum,Balance as DepositAmount," & _
        " Title+' '+FirstName+' '+MiddleName+' '+LastName as Name " & _
        " From RDMaster A,RDTrans B,NameTab C" & _
        " Where ClosedDate is NULL " & _
        " And MaturityDate <= #" & gStrDate & "#" & _
        " And B.AccID = A.AccId " & _
        " And TransID = (Select Max(TransID) From RDTrans D " & _
            " WHERE D.AccID = A.AccId)" & _
        " And C.CustomerID = A.CustomerId" & _
        " Order By MaturityDate,val(AccNum)"

If gDbTrans.Fetch(rstDeposit, adOpenDynamic) < 1 Then GoTo PigmyDEPOSIT
FDChecked = True

Dim RDClass As New clsRDAcc
    
With m_frmGrid.grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(424)
    .CellFontBold = True
    While Not rstDeposit.EOF
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        .RowData(.Row) = rstDeposit("AccID")
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstDeposit("AccNUm"))
        .Col = 2: .Text = FormatField(rstDeposit("Name"))
        DepAmount = FormatField(rstDeposit("DepositAmount"))
        .Col = 3: .Text = DepAmount
        IntAmount = RDClass.InterestAmount(rstDeposit("AccID"), gStrDate)
        .Col = 4: .Text = FormatCurrency(IntAmount)
        .Col = 5: .Text = FormatCurrency(DepAmount + IntAmount)
        
        TotalDepAmount = TotalDepAmount + DepAmount
        TotalIntAmount = TotalIntAmount + IntAmount
'        m_frmGrid.Selected(.Row) = True
        rstDeposit.MoveNext
    Wend
    
End With
Set RDClass = Nothing

PigmyDEPOSIT:

gDbTrans.SqlStmt = "Select A.AccId,AccNum,Balance as DepositAmount," & _
        " Title+' '+FirstName+' '+MiddleName+' '+LastName as Name " & _
        " From PDMaster A,PDTrans B,NameTab C" & _
        " Where ClosedDate is NULL " & _
        " And MaturityDate <= #" & gStrDate & "#" & _
        " And B.AccID = A.AccId " & _
        " And TransID = (Select Max(TransID) From PDTrans D " & _
            " WHERE D.AccID = A.AccId)" & _
        " And C.CustomerID = A.CustomerId" & _
        " Order By MaturityDate,val(AccNum)"

If gDbTrans.Fetch(rstDeposit, adOpenDynamic) < 1 Then GoTo PigmyDEPOSIT:
FDChecked = True

Dim PDClass As New clsPDAcc
    
With m_frmGrid.grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(425)
    .CellFontBold = True
    While Not rstDeposit.EOF
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1: SlNo = SlNo + 1
        .RowData(.Row) = rstDeposit("AccID")
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstDeposit("AccNUm"))
        .Col = 2: .Text = FormatField(rstDeposit("Name"))
        DepAmount = FormatField(rstDeposit("DepositAmount"))
        .Col = 3: .Text = DepAmount
        IntAmount = PDClass.InterestAmount(rstDeposit("AccID"), DayBeginUSDate)
        .Col = 4: .Text = FormatCurrency(IntAmount)
        .Col = 5: .Text = FormatCurrency(DepAmount + IntAmount)
        
        TotalDepAmount = TotalDepAmount + DepAmount
        TotalIntAmount = TotalIntAmount + IntAmount
        rstDeposit.MoveNext
    Wend
    
End With
Set PDClass = Nothing

With m_frmGrid.grd
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows = .Row + 1 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .Text = GetResourceString(286)
    .CellFontBold = True
    .Col = 3: .Text = TotalDepAmount: .CellFontBold = True
    .Col = 4: .Text = FormatCurrency(TotalIntAmount): .CellFontBold = True
    .Col = 5: .Text = FormatCurrency(TotalDepAmount + TotalIntAmount): .CellFontBold = True
End With

m_retVar = ""
MousePointer = vbDefault

If Not FDChecked Then Exit Sub

m_frmGrid.Show 1

Dim MaxAcc As Integer
MaxAcc = rstDeposit.RecordCount
Set rstDeposit = Nothing

If Not m_retVar Then
    Unload m_frmGrid
    Set m_frmGrid = Nothing
    Exit Sub
End If

End Sub

Private Sub cmdOk_Click()
'Now Write Into the LOgin Details
If fraBegin.Visible Then
    gDbTrans.SqlStmt = "UpDate INstall Set ValueData = '" & txtBeginDate & "'" & _
                " WHEre KeyData = 'BeginDate'"
    gDbTrans.BeginTrans
    Call gDbTrans.SQLExecute
    gDbTrans.CommitTrans
End If




Unload Me
End Sub

Private Sub cmdTrans_Click()

Call ShowTransaction

End Sub


Private Sub cmdTransDate_Click()
With Calendar
    .Left = Left + cmdTransDate.Left
    .Top = Top + cmdTransDate.Top
    .Show 1
    txtBeginDate = .selDate
End With
End Sub

Private Sub Form_Load()
Call CenterMe(Me)
Call SetKannadaCaption

End Sub

Private Sub m_frmGrid_CancelClicked()
m_retVar = False
End Sub

Private Sub m_frmGrid_OKClicked()
m_retVar = True
End Sub


Private Sub txtBeginDate_LostFocus()

If Not DateValidate(txtBeginDate, "/", True) Then Exit Sub

Dim ObDate As Date
Dim TransClass As Object

    ObDate = GetSysFormatDate(txtBeginDate)
    Set TransClass = New clsAccTrans
    txtOpBalance = TransClass.GetOpBalance(wis_CashHeadID, ObDate)
    Set TransClass = New clsMaterial
    txtOpStock = TransClass.GetOnDateClosingStockValue(DateAdd("d", -1, ObDate))
    Set TransClass = Nothing
    
End Sub


