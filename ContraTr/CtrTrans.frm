VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmContra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Index 2000 - Contra Transactions Wizard"
   ClientHeight    =   8040
   ClientLeft      =   930
   ClientTop       =   930
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTransDate 
      Height          =   345
      Left            =   2160
      TabIndex        =   30
      Top             =   210
      Width           =   1455
   End
   Begin VB.TextBox txtvoucher 
      Height          =   345
      Left            =   6150
      TabIndex        =   29
      Top             =   210
      Width           =   1425
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   ".."
      Height          =   315
      Left            =   3720
      TabIndex        =   28
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   30
      TabIndex        =   23
      Top             =   4305
      Width           =   7605
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   400
         Left            =   6210
         TabIndex        =   25
         Top             =   2610
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   400
         Left            =   4770
         TabIndex        =   26
         Top             =   2610
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2295
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame fraTo 
      Caption         =   "Account To Transaction(Credit)"
      Height          =   1695
      Left            =   30
      TabIndex        =   9
      Top             =   2040
      Width           =   7575
      Begin WIS_Currency_Text_Box.CurrText txtToAmount 
         Height          =   345
         Left            =   6240
         TabIndex        =   16
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.OptionButton optPenalInt 
         Caption         =   "Penal Interest"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5250
         TabIndex        =   20
         Top             =   1200
         Width           =   2085
      End
      Begin VB.OptionButton optRegInt 
         Caption         =   "Regular Interest"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   19
         Top             =   1170
         Width           =   2505
      End
      Begin VB.OptionButton optPrincipal 
         Caption         =   "Principal"
         Height          =   255
         Left            =   210
         TabIndex        =   18
         Top             =   1200
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.TextBox txtToAccNo 
         Height          =   345
         Left            =   4380
         TabIndex        =   14
         Top             =   630
         Width           =   885
      End
      Begin VB.ComboBox cmbTo 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   3855
      End
      Begin VB.CommandButton cmdTo 
         Caption         =   "..."
         Height          =   315
         Left            =   5310
         TabIndex        =   13
         Top             =   630
         Width           =   315
      End
      Begin VB.Label lblToAccType 
         Caption         =   "Account Type"
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   2925
      End
      Begin VB.Label lblToAccNo 
         Caption         =   "Account no"
         Height          =   315
         Left            =   3780
         TabIndex        =   12
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label lbltoAmount 
         Caption         =   "Amount"
         Height          =   315
         Left            =   6210
         TabIndex        =   15
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   400
      Left            =   6360
      TabIndex        =   27
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   400
      Left            =   5010
      TabIndex        =   22
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   400
      Left            =   6390
      TabIndex        =   21
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Frame fraFrom 
      Caption         =   "Account From Transactin(Debit)"
      Height          =   1305
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   7575
      Begin VB.TextBox txtFromAccNo 
         Height          =   345
         Left            =   4380
         TabIndex        =   5
         Top             =   810
         Width           =   885
      End
      Begin VB.CommandButton cmdFrom 
         Caption         =   "..."
         Height          =   315
         Left            =   5340
         TabIndex        =   4
         Top             =   810
         Width           =   315
      End
      Begin VB.ComboBox cmbFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   810
         Width           =   3765
      End
      Begin WIS_Currency_Text_Box.CurrText txtFromAmount 
         Height          =   345
         Left            =   6150
         TabIndex        =   7
         Top             =   810
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblFromAmount 
         Caption         =   "Amount"
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label lblFromAccNo 
         Caption         =   "Acc From No :"
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label lblFromAccType 
         Caption         =   "Account Type"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   2925
      End
   End
   Begin VB.Label lbltransDate 
      Caption         =   "TransDate"
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Top             =   270
      Width           =   1485
   End
   Begin VB.Label lblVoucher 
      Caption         =   "Voucher No"
      Height          =   315
      Left            =   4440
      TabIndex        =   17
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "frmContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_InitLoad As Boolean
Private m_GridRow As Integer
Private m_AccHeadId() As Long
Private m_Id() As Long
Private m_FromAmount As Currency
Private m_ToAmount As Currency

Private m_retVar As Variant
Private m_RetAccId As Long
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_AmountType() As wis_AmountType
Private m_Amount() As Currency

Public Event AddClicked()
Public Event UnDoClicked()
Public Event SaveClicked()
Public Event CancelClicked()
Public Event ClearClicked()
Public Event WindowClosed()

'Public Event TransferAmount(AccTypes() As wisModules, AccId() As Long, Amount() As Currency, AmountType() As Wis_AmountType, TransDate As Date, VoucherNo As String, Cancel As Integer)
Public Event TransferAmount(AccHeadID() As Long, AccId() As Long, Amount() As Currency, AmountType() As wis_AmountType, TransDate As Date, VoucherNo As String, frmUnload As Integer)

Private Function AddToGrid() As Boolean

On Error GoTo ErrLine

AddToGrid = False

Dim GrdRow As Integer
Dim ArrNo As Integer

Dim AccHeadID As wisModules
Dim AccNum As String
Dim StrAcctype As String

If m_GridRow = 0 Then m_GridRow = 1

If m_GridRow = 1 Then
    grd.Row = 1
    AccHeadID = cmbFrom.ItemData(cmbFrom.ListIndex)
    AccNum = Trim(txtFromAccNo)
    StrAcctype = cmbFrom.Text
    m_Amount(0) = txtFromAmount.Value
Else
    'Get the grid row
    With grd
        .Col = 0
        For ArrNo = 1 To .Rows - 1
            .Row = ArrNo
            If .Text = "" Then Exit For
        Next
        If .Row = 1 And m_InitLoad Then .Row = 2
    End With
    AccHeadID = cmbTo.ItemData(cmbTo.ListIndex)
    AccNum = Trim(txtToAccNo)
    StrAcctype = cmbTo.Text
End If

GrdRow = grd.Row
ArrNo = grd.Row - 1

'Now check the whether entered Details are correct are not
'Now Check the account no w.r.t Account tYPE

'Now get the Account if Not specified
If UBound(m_Id) < ArrNo Then
    ReDim Preserve m_Id(ArrNo)
    ReDim Preserve m_AccHeadId(ArrNo)
    ReDim Preserve m_AmountType(ArrNo)
    ReDim Preserve m_Amount(ArrNo)
End If

'Id we do not have the unique Id of the account type then
'Get the unique id of the account type
Dim RstAcc As Recordset
Set RstAcc = GetAccRecordSet(AccHeadID, AccNum)
If RstAcc Is Nothing Then
    'MsgBox "Invalid account no specified", vbInformation, wis_MESSAGE_TITLE
    'MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
    Call ActivateTextBox(IIf(grd.Row = 1, txtFromAccNo, txtToAccNo))
    Exit Function
End If

m_Id(ArrNo) = RstAcc("Id")
m_AccHeadId(ArrNo) = AccHeadID
m_AmountType(ArrNo) = wisPrincipal

Dim IntHeadID As Long
Dim bankClass As New clsBankAcc

If optRegInt Then
    m_AmountType(ArrNo) = wisRegularInt
    IntHeadID = GetIndexHeadID(GetHeadName(AccHeadID) & " " & GetResourceString(344))
    'If IntHeadId = 0 Then _
            IntHeadId = GetHeadID(GetHeadName(AccHeadID) & " " & GetResourceString(483), parMemLoanIntReceived)
    If IntHeadID = 0 Then GoTo ErrLine
End If
If optPenalInt Then
    m_AmountType(ArrNo) = wisPenalInt
    IntHeadID = GetIndexHeadID(GetHeadName(AccHeadID) & " " & GetResourceString(345))
    If IntHeadID = 0 Then _
        IntHeadID = GetIndexHeadID(GetHeadName(AccHeadID) & " " & GetResourceString(483))
    If IntHeadID = 0 Then GoTo ErrLine
    m_AccHeadId(ArrNo) = IntHeadID
End If

Set bankClass = Nothing

If optPrincipal = False Then
    'Change the Account HeadID
    
End If


'Now add the details to the grid
With grd
    .Col = 0: .Text = .Row
    .Col = 1: .Text = StrAcctype
    .Col = 2: .Text = GetResourceString(IIf(optPrincipal, 310, IIf(optRegInt, 344, 345)))
    .Col = 3: .Text = RstAcc("AccNum")
    .Col = 4: .Text = RstAcc("CustName")
    .Col = 5
    If Not fraTo.Enabled Then
        .Text = txtFromAmount.Value
        .CellAlignment = 1
        m_FromAmount = m_FromAmount + txtFromAmount.Value
    Else
       .Text = txtToAmount
       m_ToAmount = m_ToAmount + txtToAmount.Value
       m_Amount(ArrNo) = txtToAmount.Value
    End If
End With
'Dim ModuleId As wisModules

txtToAmount.Value = m_FromAmount - m_ToAmount
cmdSave.Enabled = IIf(txtToAmount.Value, False, True)

ExitLine:
AddToGrid = True

m_GridRow = m_GridRow + 1

Exit Function

ErrLine:
If Err Then
    MsgBox "Error in AddtoGrid"
    'Resume
    Err.Clear
End If
    
AddToGrid = False

End Function

Public Sub Clear()
'first Intitielaise the Varible
ReDim m_AccHeadId(0)
ReDim m_Id(0)
ReDim m_AmountType(0)
ReDim m_Amount(0)
'After That initialize the grid

Call InitGrid
cmbFrom.ListIndex = -1
cmbTo.ListIndex = -1

fraTo.Enabled = False
cmbFrom.Enabled = True
txtFromAccNo.Enabled = True
txtFromAmount.Enabled = True
cmdFrom.Enabled = True
optPrincipal = True

m_FromAmount = 0
m_ToAmount = 0
cmdUndo.Enabled = False
cmdSave.Enabled = False

End Sub

'
Private Sub InitGrid()
With grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .SelectionMode = flexSelectionByRow
    .Rows = 2
    .Cols = 6
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33): .ColWidth(0) = 450 'SlnO
    .Col = 1: .Text = GetResourceString(36, 35): .ColWidth(1) = 1650 'Accoun Name
    .Col = 2: .Text = GetResourceString(36): .ColWidth(2) = 950 'Name
    .Col = 3: .Text = GetResourceString(36, 60): .ColWidth(3) = 850 'Account No
    .Col = 4: .Text = GetResourceString(35): .ColWidth(4) = 1850 'Name
    .Col = 5: .Text = GetResourceString(40): .ColWidth(5) = 1050 'Amount
    
End With
    
End Sub


Public Sub InitialiseValue(ByVal TransDate As Date, ByVal VoucherNo As String, ByVal AccHeadID As Long, ByVal AccNum As String, ByVal Amount As Currency, cancel As Integer, Optional AmountType As wis_AmountType = wisPrincipal)
'Public Sub InitialiseValue(ByVal TransDate As Date, ByVal VoucherNo As String, ByVal Module As wisModules, ByVal AccNum As String, ByVal Amount As Currency, Cancel As Integer, Optional AmountType As Wis_AmountType = wisPrincipal)
Dim StrAcctype As String

'if the form has not loaded 'first load the form
'to load the form fetch the window handle
Dim X As Long
X = Me.hwnd

If Amount = 0 Then GoTo Exit_Line
If AccHeadID = 0 Then GoTo Exit_Line

'If transferring amount already loaded
'then exit the sub
cancel = True

If fraTo.Enabled Then GoTo Exit_Line
m_InitLoad = True
Call Clear

txtTransDate.Text = GetIndianDate(CStr(TransDate)): txtTransDate.Locked = True
txtvoucher = VoucherNo: txtvoucher.Locked = True

txtFromAccNo = AccNum
StrAcctype = GetResourceString(307)
cmbFrom.Text = GetResourceString(307)
txtFromAmount = Amount
m_Amount(0) = Amount

'Id we do not have the unique Id of the account type then
'Get the unique id of the account type
Dim RstAcc As Recordset
Set RstAcc = GetAccRecordSet(AccHeadID, AccNum)
If RstAcc Is Nothing Then
    'MsgBox "Invalid account no specified", vbInformation, wis_MESSAGE_TITLE
    'MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
    Call ActivateTextBox(IIf(grd.Row = 1, txtFromAccNo, txtToAccNo))
    Exit Sub
End If

m_Id(0) = RstAcc("Id")
m_AccHeadId(0) = AccHeadID
'If Module > 100 Then m_AccType(0) = Module - Module Mod 100
m_AmountType(0) = wisPrincipal

'If optRegInt Then m_AmountType(ArrNo) = wisRegularInt
'If optPenalInt Then m_AmountType(ArrNo) = wisPenalInt

'Now add the details to the grid
With grd
    .Rows = 3
    .Row = 1
    .Col = 0: .Text = .Row
    .Col = 1: .Text = StrAcctype
    .Col = 2: .Text = GetResourceString(IIf(optPrincipal, 310, IIf(optRegInt, 344, 345)))
    .Col = 3: .Text = RstAcc("AccNum")
    .Col = 4: .Text = RstAcc("CustName")
    .Col = 5: .Text = txtFromAmount
    .CellAlignment = 1
    m_FromAmount = m_FromAmount + txtFromAmount
End With
txtToAmount = m_FromAmount - m_ToAmount

fraTo.Enabled = True
cmbFrom.Enabled = False
txtFromAccNo.Enabled = False
txtFromAmount.Enabled = False
cmdFrom.Enabled = False


cancel = False

Exit_Line:

End Sub

Private Sub LoadAccountBalance()
If cmbFrom.ListIndex < 0 Then Exit Sub

Dim ret As Integer
Dim BalanceAmount As Currency
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim AccType As wisModules
Dim DepositType As Integer
'Check for the Selected AccountType if Account type is Sb then
'get the SB balance and if it is CA then Get the CA Balance-siddu
Dim LstInd As Long

'AccType = GetModuleIDFromHeadID(AccHeadID)
AccType = GetModuleIDFromHeadID(cmbFrom.ItemData(cmbFrom.ListIndex))
DepositType = AccType Mod 100
If AccType > 100 Then AccType = AccType - AccType Mod 100
With cmbFrom
   LstInd = .ListIndex
   'If .ListIndex = 0 Then AccType = wis_SBAcc
   'If .ListIndex = 1 Then AccType = wis_CAAcc
End With



'DepositType = GetDepositTypeIDFromHeadID(cmbFrom.ItemData(cmbFrom.ListIndex))
             
'Check for the Account tyeps
'AccType = Accnum
'If it is Current Account then get the Balance
If AccType = wis_CAAcc Then
    gDbTrans.SqlStmt = "SELECT AccID FROM CAMaster " & _
        " WHERE AccNum = " & AddQuotes(Trim$(txtFromAccNo.Text), True)
    If DepositType > 0 Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & DepositType
        
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        txtFromAmount.Text = ""
        Exit Sub
    End If
   
    AccId = FormatField(rst("AccID"))
    
    'get teh TotalBalance Form the Particular Account
    gDbTrans.SqlStmt = "Select TOP 1 Balance from CATrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox "No Records"
    Else
        BalanceAmount = FormatField(rst(0))
        txtFromAmount.Value = BalanceAmount
        Set rst = Nothing
    End If
End If

'if the Account is SB Get the Sb Account balance Given from the AccountId
If AccType = wis_SBAcc Then
    'AccId=
    gDbTrans.SqlStmt = "Select * from SBMaster" & _
                " where AccNum = " & AddQuotes(txtFromAccNo, True)
    If DepositType > 0 Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And DepositType = " & DepositType
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        txtFromAmount.Text = ""
        txtFromAccNo.SetFocus
        Exit Sub
    End If
    
    AccId = FormatField(rst("AccId"))
    gDbTrans.SqlStmt = "Select TOP 1 Balance from SBTrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox GetResourceString(179), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        BalanceAmount = FormatField(rst(0))
        txtFromAmount.Value = BalanceAmount
        Set rst = Nothing
    End If
End If

End Sub

Private Sub LoadAccountType()

Dim rstDeposit As Recordset
Dim rstLoan As Recordset
Dim rstTemp As Recordset
Dim recCount As Integer
Dim loopCount As Integer

gDbTrans.SqlStmt = "Select * FROM DepositName"
If gDbTrans.Fetch(rstDeposit, adOpenStatic) < 1 Then Set rstDeposit = Nothing
gDbTrans.SqlStmt = "Select * FROM LoanScheme"
If gDbTrans.Fetch(rstLoan, adOpenStatic) < 1 Then Set rstLoan = Nothing

'Dim ModuleId As wisModules
Dim AccHeadID As Long
Dim ClsBank As clsBankAcc
Dim AccName As String

Set ClsBank = New clsBankAcc

With cmbFrom
    
    'Withdrawing of the amount willbe only from the
    'account type where cheque facility is given
    'Such type of account are only two(three) types
    'They are
    '1) Saving Bank Account
    '2) Current Account
    '3) Suspence account
    '4) Od Account 'Presently no such accounts are in our s/w
    
    .Clear
    'ModuleId = wis_SBAcc
    Dim strDeposits() As String
    'AccName = GetResourceString(421) '"SB Account
    strDeposits = GetDepositTypesList(wis_SBAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount
    'ModuleId = wis_CAAcc
    'AccName = GetResourceString(422) '"CA AccOUnt
    strDeposits = GetDepositTypesList(wis_CAAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount
    
    'ModuleId = wis_SuspAcc
    AccName = GetResourceString(365) '"
    AccHeadID = GetIndexHeadID(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    'Here also show the Bank Account & bank loan account
    'beacause there is a chance of transaction between
    'Bank Sb and member sb & between Bank loan & member loan
    'or vice versa
    AccHeadID = parBankAccount
    'gDbTrans.SqlStmt = "Select * from Heads Where ParentID = " & AccHeadID
    gDbTrans.SqlStmt = "Select * from Heads Where ParentID IN (" & AccHeadID & "," & parBankLoanAccount & ")"
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        While Not rstTemp.EOF
            .AddItem FormatField(rstTemp("HeadName"))
            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
            rstTemp.MoveNext
        Wend
    End If
    gDbTrans.SqlStmt = "Select * from Heads Where IsContraHead = 1 " & _
        " And ParentID not IN (" & AccHeadID & "," & parBankLoanAccount & ")"
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        While Not rstTemp.EOF
            .AddItem FormatField(rstTemp("HeadName"))
            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
            rstTemp.MoveNext
        Wend
    End If
'    'Here load the Bank Loan Account
'    AccHeadID = parBankLoanAccount
'    gDbTrans.SqlStmt = "Select * from Heads Where ParentID = " & AccHeadID
'    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
'        While Not rstTemp.EOF
'            .AddItem FormatField(rstTemp("HeadName"))
'            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
'            rstTemp.MoveNext
'        Wend
'    End If
End With

With cmbTo
    .Clear
    
    'Share Account
    gDbTrans.SqlStmt = "Select * From Heads where ParentID = " & parMemberShare  'parMemberDeposit
    recCount = gDbTrans.Fetch(rstTemp, adOpenDynamic)
    If recCount = 1 Then
        AccName = GetResourceString(53, 36) 'Share Account
        AccHeadID = GetIndexHeadID(AccName)
        If AccHeadID Then
            .AddItem AccName
            .ItemData(.newIndex) = AccHeadID
        End If
    Else
        While rstTemp.EOF = False
            .AddItem FormatField(rstTemp("HeadName"))
            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
            
            rstTemp.MoveNext
        Wend
    End If
    
    'AccName = GetResourceString(421) '"SB Account
    strDeposits = GetDepositTypesList(wis_SBAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount
    
    'AccName = GetResourceString(422) '"CA AccOUnt
    strDeposits = GetDepositTypesList(wis_CAAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount
    
    AccName = GetResourceString(424)  '"RD Account
    strDeposits = GetDepositTypesList(wis_RDAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount
    
    AccName = GetResourceString(425) '"Pigmy Account
    strDeposits = GetDepositTypesList(wis_PDAcc)
    recCount = UBound(strDeposits)
    For loopCount = 0 To recCount - 1
        'AccHeadID = GetIndexHeadID(AccName)
        AccHeadID = GetIndexHeadID(strDeposits(loopCount))
        If AccHeadID Then
            .AddItem strDeposits(loopCount)
            .ItemData(.newIndex) = AccHeadID 'ModuleId
        End If
    Next loopCount

    'Now Load All othe Deposits
    'ModuleId = wis_Deposits
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName"))
            AccHeadID = GetIndexHeadID(AccName)
            If AccHeadID Then
                .AddItem AccName
                .ItemData(.newIndex) = AccHeadID
            End If
            rstDeposit.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 43) '"BKCC
    AccHeadID = GetIndexHeadID(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If

    'Now Load All Loans
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        While Not rstLoan.EOF
            AccName = FormatField(rstLoan("SchemeName"))
            AccHeadID = GetIndexHeadID(AccName)
            If AccHeadID Then
                .AddItem AccName
                .ItemData(.newIndex) = AccHeadID
            End If
            rstLoan.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = GetResourceString(229, 58) '"BKCC Loan
    AccHeadID = GetIndexHeadID(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If

    'ModuleId = wis_DepositLoans
    AccName = GetResourceString(43, 58)
    AccHeadID = GetIndexHeadID(AccName)
    '.AddItem AccName
    '.ItemData(.NewIndex) = AccHeadID 'ModuleId
    If Not rstDeposit Is Nothing Then
        rstDeposit.MoveFirst
        While Not rstDeposit.EOF
            AccName = FormatField(rstDeposit("DepositName")) & " " & GetResourceString(58)
            AccHeadID = GetIndexHeadID(AccName)
            If AccHeadID Then
                .AddItem AccName
                .ItemData(.newIndex) = AccHeadID
            End If
            rstDeposit.MoveNext
        Wend
    End If
    
    'ModuleId = wis_SuspAcc
    AccName = GetResourceString(365) '"Suspence AccOUnt
    AccHeadID = GetIndexHeadID(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID
    End If
    
    'Here also, show the Bank Account & bank loan account
    'beacause there is a chance of transaction between
    'Bank Sb and member sb & between Bank loan & member loan
    'or vice versa
    AccHeadID = parBankAccount
    gDbTrans.SqlStmt = "Select * from Heads Where ParentID = " & AccHeadID
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        While Not rstTemp.EOF
            .AddItem FormatField(rstTemp("HeadName"))
            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
            rstTemp.MoveNext
        Wend
    End If
    'Here load the Bank Loan Account
    AccHeadID = parBankLoanAccount
    gDbTrans.SqlStmt = "Select * from Heads Where ParentID = " & AccHeadID
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        While Not rstTemp.EOF
            .AddItem FormatField(rstTemp("HeadName"))
            .ItemData(.newIndex) = FormatField(rstTemp("HeadID"))
            rstTemp.MoveNext
        Wend
    End If

End With

End Sub

'
Private Function SaveDetails() As Boolean
On Error GoTo Exit_Line

Dim PrincAmount As Currency
Dim Amount As Currency
Dim IntAmount As Currency
Dim PenalAmount As Currency
Dim BalanceAmount As Currency
Dim ModuleID As wisModules
Dim AccHeadID As Long
Dim InTrans As Boolean
Dim TransDate As Date
Dim count As Integer
Dim MaxCount As Integer
Dim VoucherNo As String

TransDate = GetSysFormatDate(txtTransDate)
VoucherNo = Trim$(txtvoucher.Text)

'First withdraw the amount from the account
grd.Row = 1
grd.Col = grd.Cols - 1
Amount = m_Amount(0)
BalanceAmount = Amount

Dim ClsObject As Object

AccHeadID = m_AccHeadId(0)

'Get the ModuleID
Dim rstTemp As ADODB.Recordset
ModuleID = GetModuleIDFromHeadID(AccHeadID)
If ModuleID > 100 Then ModuleID = ModuleID - ModuleID Mod 100

If Not m_InitLoad Then
    gDbTrans.BeginTrans
    InTrans = True
    If ModuleID = wis_SBAcc Or ModuleID = wis_CAAcc Then
        If ModuleID = wis_CAAcc Then Set ClsObject = New ClsCAAcc
        If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
        If ClsObject.WithdrawAmount(m_Id(0), Amount, "Tfr to ", _
                TransDate, VoucherNo) = 0 Then GoTo Exit_Line
                
        Set ClsObject = Nothing
    End If
End If

MaxCount = UBound(m_Id)
Dim I As Integer
For count = 1 To MaxCount
    If m_Amount(count) = 0 Then GoTo NextAmount
    
    IntAmount = 0: PenalAmount = 0
    Amount = m_Amount(count)
    
    ModuleID = GetModuleIDFromHeadID(m_AccHeadId(count))
    
    If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
    
    If ModuleID = wis_CAAcc Or ModuleID = wis_SBAcc Or _
        ModuleID = wis_PDAcc Or _
        ModuleID = wis_Deposits Or ModuleID = wis_RDAcc Then
        
        If ModuleID = wis_BKCC Or ModuleID = wis_BKCCLoan Then Set ClsObject = New clsBkcc
        If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
        If ModuleID = wis_Deposits Then Set ClsObject = New clsFDAcc
        If ModuleID = wis_PDAcc Then Set ClsObject = New clsPDAcc
        If ModuleID = wis_CAAcc Then Set ClsObject = New ClsCAAcc
        If ModuleID = wis_RDAcc Then Set ClsObject = New clsRDAcc
        
        BalanceAmount = BalanceAmount - Amount
        If ClsObject.DepositAmount(m_Id(count), Amount, _
                         "Tfr From ", TransDate, VoucherNo) = 0 Then GoTo Exit_Line
        Set ClsObject = Nothing
    End If
    
    If ModuleID = wis_DepositLoans Or ModuleID = wis_BKCCLoan Or _
                    ModuleID = wis_Loans Or ModuleID = wis_Members Then
        'Now search whether he is Paying any Interest Amount
        'on this loan account
        Amount = 0
        I = count
        Do
            If I > MaxCount Then Exit Do
            
            If m_Id(count) = m_Id(I) And m_AccHeadId(count) = m_AccHeadId(I) Then
                If m_AmountType(I) = wisPrincipal And Amount = 0 Then
                    Amount = m_Amount(I)
                    m_Amount(I) = 0
                ElseIf m_AmountType(I) = wisRegularInt And IntAmount = 0 Then
                    IntAmount = m_Amount(I)
                    m_Amount(I) = 0
                ElseIf m_AmountType(I) = wisPenalInt And PenalAmount = 0 Then
                    PenalAmount = m_Amount(I)
                    m_Amount(I) = 0
                End If
            End If
            I = I + 1
        Loop

        If ModuleID = wis_DepositLoans Then
            Set ClsObject = New clsDepLoan
            If ClsObject.DepositAmount(m_Id(count), Amount, _
                IntAmount, "Tfr From ", TransDate, VoucherNo) = 0 Then GoTo Exit_Line
            
            BalanceAmount = BalanceAmount - Amount - IntAmount
        
        ElseIf ModuleID = wis_Loans Or ModuleID = wis_BKCCLoan Then
            
            If ModuleID = wis_Loans Then Set ClsObject = New clsLoan
            If ModuleID = wis_BKCCLoan Then Set ClsObject = New clsBkcc
            If ClsObject.DepositAmount(CLng(m_Id(count)), Amount, _
                IntAmount, PenalAmount, "Tfr From ", TransDate, VoucherNo) = 0 Then GoTo Exit_Line
            
            BalanceAmount = BalanceAmount - Amount - IntAmount - PenalAmount
        
        ElseIf ModuleID = wis_Members Then
            Set ClsObject = New clsMMAcc
            If ClsObject.DepositAmount(CLng(m_Id(count)), Amount, _
                    IntAmount, "Tfr From ", TransDate, VoucherNo) = 0 Then GoTo Exit_Line
            
            BalanceAmount = BalanceAmount - Amount - IntAmount - PenalAmount
        End If
        Set ClsObject = Nothing
    End If
NextAmount:
Next

Debug.Assert BalanceAmount = 0
If BalanceAmount = 0 Then
    If Not m_InitLoad Then gDbTrans.CommitTrans
    InTrans = False
    SaveDetails = True
End If
'Exit Function

Exit_Line:
If InTrans Then gDbTrans.RollBack

If Err Then
    MsgBox "Error In save Details", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If
    
End Function

Private Sub cmbFrom_Click()
If cmbFrom.ListIndex < 0 Then Exit Sub

Dim headID  As Long
Dim parentHeadID  As Long
'Get the Head ID of selected head
headID = cmbFrom.ItemData(cmbFrom.ListIndex)

'Now get the PArent ID of the selected head
parentHeadID = GetParentID(headID)
'if selected head is not havin any accoutn heads then
'Disablle the Account number textbox
If parentHeadID = parBankAccount Or parentHeadID = parBankLoanAccount Then
    txtFromAccNo.Enabled = False
    txtFromAccNo.Text = ""
    cmdFrom.Enabled = False
    'Get the Balance
    Dim bankClass As New clsBankAcc
    Me.txtFromAmount = bankClass.Balance(txtTransDate.Text, headID)
Else
    txtFromAccNo.Enabled = True
    cmdFrom.Enabled = True
End If

End Sub


Private Sub cmbTo_Click()

If cmbTo.ListIndex < 0 Then Exit Sub

'Dim Moduleid As wisModules
Dim headID As Long
Dim ParentID As Long
'Get the head id of transaction head
headID = cmbTo.ItemData(cmbTo.ListIndex)
'Now Gert the PArent id of the Transaction head
ParentID = GetParentID(headID)

'If HeadID > 100 Then HeadID = HeadID - (HeadID Mod 100)

optRegInt.Caption = GetResourceString(344)
        
'If ModuleId = wis_DepositLoans Or ModuleId = wis_Members Then
If ParentID = parMemDepLoan Or ParentID = parMemberShare Then
    If ParentID = parMemberShare Then optRegInt.Caption = _
        GetResourceString(53, 191)
    optRegInt.Enabled = True
    optPenalInt.Enabled = False
'ElseIf ModuleId = wis_Loans Or ModuleId = wis_BKCC Then
ElseIf ParentID = parMemberLoan Then
    optRegInt.Enabled = True
    optPenalInt.Enabled = True
Else
    optRegInt.Enabled = False
    optPenalInt.Enabled = False
End If
optPrincipal = True

'if selected head is not havin any accoutn heads then
'Disable the Account number textbox
If ParentID = parBankAccount Or ParentID = parBankLoanAccount Then
    txtToAccNo.Enabled = False
    txtToAccNo.Text = ""
    cmdTo.Enabled = False
Else
    txtToAccNo.Enabled = True
    cmdTo.Enabled = True
End If

End Sub

Private Sub cmdAccept_Click()

    RaiseEvent AddClicked
    Exit Sub

End Sub

'
Private Sub cmdClear_Click()
RaiseEvent ClearClicked

End Sub

'
Private Sub cmdClose_Click()
RaiseEvent CancelClicked
Unload Me
End Sub

'
Private Sub cmdDate_Click()
With Calendar
    .selDate = gStrDate
    .Left = fraFrom.Left + Left + cmdDate.Left + cmdDate.Width
    .Top = fraFrom.Top + Top + txtTransDate.Top - .Height / 2 + cmdDate.Height
    .Show 1
    txtTransDate = .selDate
End With
End Sub

Private Sub cmdFrom_Click()

If cmbFrom.ListIndex < 0 Then Exit Sub

On Error GoTo Exit_Line

    cmdFrom.Enabled = False
    cmdTo.Enabled = False

Dim AccHeadID As Long
Dim SqlStr As String

AccHeadID = cmbFrom.ItemData(cmbFrom.ListIndex)
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

Dim RstCust As Recordset

MousePointer = vbHourglass
Set RstCust = GetAccRecordSet(AccHeadID)

If RstCust Is Nothing Then
    MsgBox "There are no customers in the " & cmbFrom.Text, _
            vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
MousePointer = vbHourglass
Call FillView(m_frmLookUp.lvwReport, RstCust, True)
m_retVar = ""

With m_frmLookUp
    .lvwReport.ColumnHeaders(2).Width = 0
    .Show 1
End With

MousePointer = vbDefault

txtFromAccNo = m_retVar
Me.MousePointer = vbDefault
    
Exit_Line:
    cmdFrom.Enabled = True
    cmdTo.Enabled = True
    MousePointer = vbDefault
    
End Sub

'
Private Sub cmdSave_Click()
'Check For the date
If Not DateValidate(txtTransDate, "/", True) Then
    'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Sub
End If

If Trim(txtvoucher) = "" Then
    'MsgBox "Invalid Voucher specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(511), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtvoucher
    Exit Sub
End If
RaiseEvent SaveClicked


End Sub

Private Sub cmdTo_Click()

If cmbTo.ListIndex < 0 Then Exit Sub

On Error GoTo Exit_Line
    
cmdFrom.Enabled = False
cmdTo.Enabled = False

Dim ModuleID As wisModules
Dim RstCust As Recordset

Dim SqlStr As String
Dim SearchString As String
Dim Lret As Integer

If Trim(SearchString) = "" Then _
    SearchString = InputBox("Eneter Name to search", "SearchString")

With gDbTrans
    .SqlStmt = "SELECT AccNum," & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name " & _
        " FROM SBMaster A, NameTab B WHERE A.CustomerID = B.CustomerID "
    If Trim(SearchString) <> "" Then
        .SqlStmt = .SqlStmt & " AND (FirstName like '" & SearchString & "%' " & _
            " Or MiddleName like '" & SearchString & "%' " & _
            " Or LastName like '" & SearchString & "%')"
        .SqlStmt = .SqlStmt & " Order by IsciName"
    Else
        .SqlStmt = .SqlStmt & " Order by AccNum"
    End If
    Lret = .Fetch(RstCust, adOpenStatic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        cmdTo.Enabled = True
        Exit Sub
    End If
End With

ModuleID = cmbTo.ItemData(cmbTo.ListIndex)
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

Me.MousePointer = vbHourglass
Set RstCust = Nothing
If ModuleID = wis_SuspAcc Then
    'Load the Detils from the FromAccount
    ModuleID = cmbFrom.ItemData(cmbFrom.ListIndex)
    txtToAccNo = txtFromAccNo
   ' Set RstCust = GetAccRecordSet(Moduleid, txtFromAccNo)
    
Else
    Set RstCust = GetAccRecordSet(ModuleID)
    If RstCust Is Nothing Then GoTo Exit_Line
    
    Call FillView(m_frmLookUp.lvwReport, RstCust, True)
    m_retVar = ""
    
    With m_frmLookUp
        .lvwReport.ColumnHeaders(2).Width = 0
        .Show 1
    End With
    Me.MousePointer = vbDefault
    txtToAccNo = m_retVar
End If
Exit_Line:
    cmdFrom.Enabled = True
    cmdTo.Enabled = True
    Me.MousePointer = vbDefault
End Sub

'
Private Sub cmdUndo_Click()

With grd
    If .Row < 2 Then
        If .Rows = 3 Then Call Clear
        Exit Sub
    End If
    
    'if the user selected more than one row
    'then need not to delete that row
    If .RowSel <> .Row Then Exit Sub
    If .RowSel = .Rows - 1 Then Exit Sub
End With

Dim RemRow As Integer
Dim MaxCount As Integer
Dim count As Integer


'If The deleting row is not tHe last row then
With grd
    RemRow = .Row
    MaxCount = UBound(m_Id)
    If .Rows - 2 <> .Rows Then
        .Col = 0
        For count = RemRow To MaxCount
            .Row = count
            .Text = Val(.Text) - 1
            m_Id(count - 1) = m_Id(count)
            m_AccHeadId(count - 1) = m_AccHeadId(count)
            m_AmountType(count - 1) = m_AmountType(count)
            m_Amount(count - 1) = m_Amount(count)
        Next
    End If
    
    .Row = RemRow
    .Col = .Cols - 1
    m_ToAmount = m_ToAmount - Val(.Text)
    'remove the row
    .RemoveItem RemRow
    'Redimension the variables
    ReDim Preserve m_Id(MaxCount - 1)
    ReDim Preserve m_AccHeadId(MaxCount - 1)
    ReDim Preserve m_AmountType(MaxCount - 1)
    ReDim Preserve m_Amount(MaxCount - 1)
    
    cmdSave.Enabled = IIf(m_ToAmount - m_FromAmount = 0, True, False)
    
End With

optPrincipal = True

End Sub

'
Private Sub Form_Load()

Call CenterMe(Me)
txtTransDate = gStrDate
Call SetKannadaCaption
'gDbTrans.BeginTrans
Call LoadAccountType
'gDbTrans.CommitTrans
Call Clear

If gOnLine Then
    txtTransDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub


'
Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)
    
' TransCtion Frame
fraFrom.Caption = GetResourceString(107)
lbltransDate.Caption = GetResourceString(37)
lblVoucher.Caption = GetResourceString(41)
lblFromAccType.Caption = GetResourceString(36, 35)
lblFromAccNo.Caption = GetResourceString(36, 60)
lblFromAmount.Caption = GetResourceString(40)

fraTo.Caption = GetResourceString(108)
lblToAccType.Caption = GetResourceString(36, 35)
lblToAccNo.Caption = GetResourceString(36, 60)
lbltoAmount.Caption = GetResourceString(40)
optPrincipal.Caption = GetResourceString(310)
optRegInt.Caption = GetResourceString(344)
optPenalInt.Caption = GetResourceString(345)

cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(14)

cmdSave.Caption = GetResourceString(7)
cmdClose.Caption = GetResourceString(11)
CmdClear.Caption = GetResourceString(8)    '


End Sub

Private Sub Form_Unload(cancel As Integer)
    RaiseEvent WindowClosed
End Sub

'
Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub

'
Private Sub m_frmLookUp_SubItems(strSubItem() As String)
On Error Resume Next
m_RetAccId = Val(strSubItem(1))
Err.Clear
On Error GoTo 0
End Sub

Private Sub txtFromAccNo_Change()
'Chnge when accountNumber Changes
cmdFrom.Enabled = IIf(Trim$(txtFromAccNo.Text) <> "", True, False)
txtFromAmount.Text = ""
End Sub


Private Sub txtFromAccNo_LostFocus()

'Check for the Account Number
If txtFromAccNo.Text = "" Then Exit Sub
Call LoadAccountBalance

End Sub

