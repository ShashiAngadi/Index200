VERSION 5.00
Object = "{8491A895-6031-11D5-A300-0080AD7CA942}#12.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmCashTrans 
   Caption         =   "Cash Rceipt details"
   ClientHeight    =   6885
   ClientLeft      =   780
   ClientTop       =   1005
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8025
   Begin VB.Frame fraPayment 
      Caption         =   "Payment Details"
      Height          =   2655
      Left            =   60
      TabIndex        =   29
      Top             =   3660
      Width           =   7755
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   345
         Index           =   1
         Left            =   5130
         TabIndex        =   41
         Top             =   2220
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Index           =   1
         Left            =   6360
         TabIndex        =   42
         Top             =   2220
         Width           =   1035
      End
      Begin VB.TextBox txtScroll 
         Height          =   345
         Index           =   1
         Left            =   2010
         TabIndex        =   31
         Top             =   240
         Width           =   1545
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Index           =   1
         Left            =   2010
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   750
         Width           =   5385
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Index           =   1
         Left            =   2010
         TabIndex        =   38
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmdCustName 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   7440
         TabIndex        =   37
         Top             =   1200
         Width           =   285
      End
      Begin VB.TextBox txtCustName 
         Height          =   345
         Index           =   1
         Left            =   4800
         TabIndex        =   36
         Top             =   1200
         Width           =   2565
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Index           =   1
         Left            =   2010
         TabIndex        =   40
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.CommandButton cmdPay 
         Caption         =   ".."
         Height          =   315
         Left            =   3690
         TabIndex        =   46
         Top             =   240
         Width           =   405
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   5910
         TabIndex        =   45
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance"
         Height          =   285
         Index           =   1
         Left            =   4830
         TabIndex        =   44
         Top             =   1710
         Width           =   1065
      End
      Begin VB.Label lblName 
         Caption         =   "Cust Name :"
         Height          =   225
         Index           =   1
         Left            =   3300
         TabIndex        =   35
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblScroll 
         Caption         =   "Scroll No:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   30
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label lblAccType 
         Caption         =   "Account type"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   32
         Top             =   750
         Width           =   1755
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account NO"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   1230
         Width           =   1695
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   39
         Top             =   1710
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   ".."
      Height          =   285
      Left            =   3540
      TabIndex        =   2
      Top             =   90
      Width           =   315
   End
   Begin VB.TextBox txtTransDate 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Clos&e"
      Height          =   315
      Left            =   6600
      TabIndex        =   43
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame fraReceipt 
      Caption         =   "Cash receipt Details"
      Height          =   3225
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   7755
      Begin VB.CheckBox chkNew 
         Alignment       =   1  'Right Justify
         Caption         =   "New Account"
         Height          =   345
         Left            =   5610
         TabIndex        =   8
         Top             =   630
         Width           =   2055
      End
      Begin VB.Frame fraInterest 
         BorderStyle     =   0  'None
         Caption         =   "fraInt"
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   1980
         Width           =   7515
         Begin WIS_Currency_Text_Box.CurrText txtMisc 
            Height          =   345
            Left            =   6450
            TabIndex        =   24
            Top             =   90
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtPenal 
            Height          =   345
            Left            =   4260
            TabIndex        =   22
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtRegInt 
            Height          =   345
            Left            =   1920
            TabIndex        =   20
            Top             =   90
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin VB.Label lblMisc 
            Caption         =   "Misceleneous"
            Height          =   255
            Left            =   5280
            TabIndex        =   23
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblPenalInt 
            Caption         =   "Penal Interest"
            Height          =   255
            Left            =   3060
            TabIndex        =   21
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblRegInt 
            Caption         =   "Regular interest"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   345
         Index           =   0
         Left            =   6480
         TabIndex        =   28
         Top             =   2760
         Width           =   1035
      End
      Begin VB.OptionButton optPassing 
         Caption         =   "To Passing Officer"
         Height          =   315
         Left            =   3030
         TabIndex        =   26
         Top             =   2760
         Width           =   2235
      End
      Begin VB.OptionButton optAccount 
         Caption         =   "To Account"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox txtScroll 
         Height          =   345
         Index           =   0
         Left            =   2040
         TabIndex        =   5
         Top             =   210
         Width           =   1245
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Top             =   1560
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         Value           =   12999
         FontSize        =   8.25
      End
      Begin VB.TextBox txtCustName 
         Height          =   345
         Index           =   0
         Left            =   4380
         TabIndex        =   12
         Top             =   1110
         Width           =   2985
      End
      Begin VB.CommandButton cmdCustName 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   7440
         TabIndex        =   13
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Index           =   0
         Left            =   2040
         TabIndex        =   10
         Top             =   1110
         Width           =   1005
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   660
         Width           =   3405
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Height          =   345
         Index           =   0
         Left            =   5310
         TabIndex        =   27
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   0
         Left            =   5940
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblBalance 
         Caption         =   "Balance"
         Height          =   285
         Index           =   0
         Left            =   4860
         TabIndex        =   16
         Top             =   1590
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblName 
         Caption         =   "Cust Name :"
         Height          =   225
         Index           =   0
         Left            =   3180
         TabIndex        =   11
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label lblScroll 
         Caption         =   "Scroll No:"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   1590
         Width           =   1635
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account NO"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label lblAccType 
         Caption         =   "Account type"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1755
      End
   End
   Begin VB.Label lblransDate 
      Caption         =   "Label1"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   1635
   End
End
Attribute VB_Name = "frmCashTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_retVar As Variant

Private m_rstPayment As Recordset

Private Sub EnableAccept()

Dim I As Byte

I = IIf(fraReceipt.Enabled, 0, 1)


If cmbAccType(I).ListIndex < 0 Then Exit Sub
If Trim(txtScroll(I)) = "" Then Exit Sub
If Trim(txtAccNo(I)) = "" Then Exit Sub
If txtAmount(I) = 0 Then Exit Sub

cmdAccept(I).Enabled = True


End Sub

Private Function GetAccRecordSet(AccHeadID As Long, _
    Optional AccNum As String = "", Optional SearchName As String = "") As Recordset

On Error Resume Next

Dim I As Integer
I = IIf(fraReceipt.Enabled, 0, 1)

Dim rstReturn As Recordset
Dim SqlStr As String
Dim Pos As Long
Dim sqlClause As String

Set rstReturn = Nothing
On Error GoTo Exit_line

Dim ModuleID As wisModules

ModuleID = GetModuleIDFromHeadID(AccHeadID)

Set rstReturn = Nothing

'Members
If ModuleID = wis_Members Then
    SqlStr = "Select AccNum, A.AccID as ID From MemMaster A"

'BKcc Account
ElseIf ModuleID = wis_BKCC Then
    SqlStr = "Select AccNum, A.LoanID as ID From BKCCMaster A"
'Current Account
ElseIf ModuleID = wis_CAAcc Then
    SqlStr = "Select AccNum, AccId as Id From CAMaster A"

'Deposit Loans
ElseIf ModuleID >= wis_DepositLoans And ModuleID < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID From DepositLoanMaster A"
    If ModuleID > wis_DepositLoans Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_DepositLoans

'Deposit Accounts like Fd
ElseIf ModuleID >= wis_Deposits And ModuleID < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, AccId as ID From FDMaster A"
    If ModuleID > wis_Deposits Then _
        sqlClause = " ANd A.DepositType = " & ModuleID - wis_Deposits
'Loan Accounts
ElseIf (ModuleID >= wis_Loans And ModuleID < wis_Loans + 100) Then
    SqlStr = "Select AccNum, LoanId as ID From LoanMaster A"
    If ModuleID > wis_Loans Then _
        sqlClause = " AND A.SchemeID = " & ModuleID - wis_Loans

'Pigmy Accounts
ElseIf ModuleID = wis_PDAcc Then
    SqlStr = "Select AccNum, AccId as ID From PDMaster A"

'Recurring Accounts
ElseIf ModuleID = wis_RDAcc Then
    SqlStr = "Select AccNum, AccId as ID From RDMaster A"

ElseIf ModuleID = wis_SBAcc Then
    SqlStr = "Select AccNum, AccId as ID From SBMaster A"
Else
    MsgBox "Plese select the account type", vbInformation, wis_MESSAGE_TITLE
    cmbAccType(I).SetFocus
    Exit Function
End If
    
    SqlStr = Trim(SqlStr)
    Pos = InStr(1, SqlStr, "FROM", vbTextCompare)
    
    If Pos Then
        SqlStr = Left(SqlStr, Pos - 1) & _
          ", Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
          " " & Mid(SqlStr, Pos)
    End If
    
    SqlStr = SqlStr & ", NameTab B WHERE B.CustomerID = A.CustomerID"
        
    If AccNum <> "" Then _
        SqlStr = SqlStr & " AND A.AccNum = " & AddQuotes(AccNum, True)
    
    If Trim(SearchName) <> "" Then
        SqlStr = SqlStr & " AND (FirstName like '" & SearchName & "%' " & _
            " Or MiddleName like '" & SearchName & "%' " & _
            " Or LastName like '" & SearchName & "*')"
    End If
    
    gDbTrans.SQLStmt = SqlStr & " " & sqlClause & " ORDER By FirstName"

    If gDbTrans.Fetch(rstReturn, adOpenStatic) < 1 Then
        MsgBox "There are no customers in the " & _
            cmbAccType(I).Text, vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If

Exit_line:

Set GetAccRecordSet = rstReturn

End Function

Private Function GetNextScroll(TransType As wisTransactionTypes) As String

gDbTrans.SQLStmt = "Select ScrollNo From CashTrans " & _
    " Where CashTransID = (Select Max(CashTransID) From CashTrans " & _
    " Where TransDate >= #" & FinUSFromDate & "# And TransType = " & TransType & " )" & _
    " And TransDate >= #" & FinUSFromDate & "#"

Dim Rst As Recordset
Dim ScrollNo As String

ScrollNo = "1"

If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then GoTo ExitLine

ScrollNo = FormatField(Rst(0))

If ScrollNo = "" Then GoTo ExitLine

If IsNumeric(ScrollNo) Then ScrollNo = CStr(Val(ScrollNo)): GoTo ExitLine

Dim I As Integer
Dim strAlpha As String

I = 1
If IsNumeric(Left(ScrollNo, 1)) Then
    Do
        If Not IsNumeric(Left(ScrollNo, I)) Then Exit Do
        I = I + 1
    Loop
    strAlpha = Mid(ScrollNo, I)
    
    ScrollNo = Left(ScrollNo, I - 1)
    ScrollNo = CStr(Val(ScrollNo) + 1)
    ScrollNo = strAlpha & ScrollNo
    GoTo ExitLine
End If
If IsNumeric(Right(ScrollNo, 1)) Then
    Do
        If Not IsNumeric(Right(ScrollNo, I)) Then Exit Do
        I = I + 1
    Loop
    strAlpha = Mid(ScrollNo, 1, Len(ScrollNo) - I - 1)
    ScrollNo = Right(ScrollNo, I - 1)
    ScrollNo = CStr(Val(ScrollNo) + 1)
    ScrollNo = ScrollNo & strAlpha
End If


ExitLine:
    GetNextScroll = ScrollNo

End Function

Private Function DepositCash() As Boolean

Dim VoucherNo As String
Dim AccHeadID As Long
Dim AccNum As String
Dim I As Integer

I = IIf(fraReceipt.Enabled, 0, I)

AccHeadID = cmbAccType(I).ItemData(cmbAccType(I).ListIndex)
AccNum = txtAccNo(I)

Dim Rst As Recordset
Set Rst = GetAccRecordSet(AccHeadID, AccNum)
If Rst Is Nothing Then Exit Function

Dim Accid As Long
Dim TransType As wisTransactionTypes

Accid = FormatField(Rst("Id"))
TransType = IIf(I = 0, wDeposit, wWithdraw)

Dim CashTransID As Long
'Get the Max TransCtion ID
gDbTrans.SQLStmt = "Select Max(CashTransID) From CashTrans " & _
        " Where TransDate >= #" & FinUSFromDate & "#"

CashTransID = 1
If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then CashTransID = FormatField(Rst(0)) + 1

Dim TransID As Long
Dim TransDate As Date
Dim CounterId As Integer
Dim Balance As Currency
Dim ModuleID As wisModules

'Get the Balance of the tranacting account
Balance = Val(txtBalance(I))

TransDate = FormatDate(txtTransDate)

ModuleID = GetModuleIDFromHeadID(AccHeadID)

Dim BankClass As clsBankAcc
Set BankClass = New clsBankAcc
Dim ClsObject As Object

gDbTrans.BeginTrans

If Not optAccount Then GoTo CashLine

Dim IntAmount As Currency
Dim PenalAmount As Currency
Dim PrincAmount As Currency

PrincAmount = txtAmount(I)
    
If Not BankClass.UpdateCashDeposits(AccHeadID, PrincAmount, _
                    TransDate) Then GoTo Exit_line

If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)

If ModuleID = wis_CAAcc Or ModuleID = wis_SBAcc Or _
    ModuleID = wis_PDAcc Or _
    ModuleID = wis_Deposits Or ModuleID = wis_RDAcc Then

    If ModuleID = wis_BKCC Then Set ClsObject = New clsBkcc
    If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
    If ModuleID = wis_Deposits Then Set ClsObject = New clsFDAcc
    If ModuleID = wis_PDAcc Then Set ClsObject = New clsPDAcc
    If ModuleID = wis_CAAcc Then Set ClsObject = New clsCAAcc
    If ModuleID = wis_RDAcc Then Set ClsObject = New clsRDAcc

    TransID = ClsObject.DepositAmount(Accid, PrincAmount, _
                     "From Cash Counter " & CounterId, TransDate, VoucherNo, True)
    If TransID < 1 Then GoTo Exit_line
    Set ClsObject = Nothing

ElseIf ModuleID = wis_DepositLoans Or ModuleID = wis_BKCC Or _
    ModuleID = wis_Loans Or ModuleID = wis_Members Then
    'Now search whether he is Paying any Interest Amount
    'on this loan account
    
    If ModuleID = wis_DepositLoans Then
        Set ClsObject = New clsDepLoan
        TransID = ClsObject.DepositAmount(Accid, PrincAmount, _
            IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_line
        
    ElseIf ModuleID = wis_Loans Or ModuleID = wis_BKCC Then
        If ModuleID = wis_Loans Then Set ClsObject = New clsLoan
        If ModuleID = wis_BKCC Then Set ClsObject = New clsBkcc
        TransID = ClsObject.DepositAmount(CLng(Accid), PrincAmount, _
            IntAmount, PenalAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_line
        
    ElseIf ModuleID = wis_Members Then
        Set ClsObject = New clsMMAcc
        TransID = ClsObject.DepositAmount(CLng(Accid), PrincAmount, _
                    IntAmount, "Tfr From ", TransDate, VoucherNo, True)
        If TransID < 1 Then GoTo Exit_line
        
    End If
    Set ClsObject = Nothing
End If

CashLine:

AccHeadID = cmbAccType(I).ItemData(cmbAccType(I).ListIndex)

gDbTrans.SQLStmt = "Insert Into CashTrans " & _
    "(TransDate,CashTransID,AccHeadID," & _
    " AccID,TransType,TransId,Amount," & _
    " ScrollNo,TransUserID,CounterID)" & _
" VALUES (#" & TransDate & "#, " & CashTransID & "," & AccHeadID & _
    " ," & Accid & "," & TransType & "," & TransID & "," & _
    PrincAmount + IntAmount + PenalAmount & "," & _
    AddQuotes(txtScroll(I).Text, True) & "," & _
    gUserID & "," & CounterId & " )"

If Not gDbTrans.SQLExecute Then
    MsgBox LoadResString(gLangOffSet + 535), vbInformation, wis_MESSAGE_TITLE
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.CommitTrans

DepositCash = True

Exit_line:

End Function
Private Sub LoadPaymentVoucher(Optional NextRecord As Boolean = True)

If m_rstPayment Is Nothing Then Exit Sub
If m_rstPayment.EOF Then Exit Sub

If NextRecord Then m_rstPayment.MoveNext

If m_rstPayment.EOF Then Exit Sub

Dim AccHeadID As Long
Dim Accid As Long
Dim ModuleID As wisModules
Dim AccNum As String
Dim TableName As String

AccHeadID = FormatField(m_rstPayment("AccheadID"))
Accid = FormatField(m_rstPayment("AccID"))

ModuleID = GetModuleIDFromHeadID(AccHeadID)

If ModuleID = wis_CAAcc Then TableName = "CAMaster"
If ModuleID = wis_SBAcc Then TableName = "SBMaster"
If ModuleID = wis_Deposits Then TableName = "FDMaster"
If ModuleID = wis_PDAcc Then TableName = "PDMaster"
If ModuleID = wis_RDAcc Then TableName = "RDMaster"
If ModuleID = wis_BKCC Then TableName = "BKCCMaster"
If ModuleID = wis_Members Then TableName = "MemMaster"

If ModuleID = wis_DepositLoans Then TableName = "DepositLoanMaster"
If ModuleID = wis_Loans Then TableName = "LoanMaster"

Debug.Assert TableName <> ""
If TableName = "" Then Exit Sub

gDbTrans.SQLStmt = "Select AccNum, " & _
    " Title +' '+FirstNAme+' '+MiddleName+' '+LastName As CustName" & _
    " From " & TableName & " A, NameTab B " & _
    " Where B.CustomerID = A.CustomerID "

If ModuleID = wis_DepositLoans Or ModuleID = wis_Loans Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & _
        " AND A.LoanID = " & m_rstPayment("AccID")
Else
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & _
        " AND A.AccID = " & m_rstPayment("AccID")
End If

Dim rstTemp As Recordset

If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Sub
Dim I As Integer
I = 1

txtAccNo(I) = FormatField(rstTemp("AccNum"))
txtCustName(I) = FormatField(rstTemp("CustName"))
txtAmount(I) = FormatField(m_rstPayment("Amount"))
txtScroll(I) = GetNextScroll(wWithdraw)

With cmbAccType(1)
    I = 0
    Do
        If I = .ListCount Then Exit Do
        If AccHeadID = .ItemData(I) Then .ListIndex = I: Exit Do
        I = I + 1
    Loop
End With

End Sub

Private Sub RefreshPaymentRecordset()

Dim TransType As wisTransactionTypes
TransType = wWithdraw

gDbTrans.SQLStmt = "Select ScrollNo,AccID,HeadName, " & _
        " Amount From CashTrans A,Heads B Where " & _
        " B.HeadID = A.AccHeadid And TransType = " & TransType & _
        " And TransID > 0 " & _
        " and (TransUserID is Null or TransUserId =0 )" & _
        " order by transDate Desc,ScrollNo"
        
Call gDbTrans.Fetch(m_rstPayment, adOpenDynamic)
 

End Sub

Private Sub cmbAccType_Click(Index As Integer)

'Only For Receipt
If Index Then fraInterest.Enabled = False: Exit Sub

If cmbAccType(Index).ListIndex < 0 Then Exit Sub

If Not optAccount Then fraInterest.Enabled = False: Exit Sub

Dim AccHeadID As Long
Dim ModuleID As wisModules
AccHeadID = cmbAccType(Index).ItemData(cmbAccType(Index).ListIndex)

ModuleID = GetModuleIDFromHeadID(AccHeadID)
If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)

fraInterest.Enabled = True
lblRegInt = LoadResString(gLangOffSet + 344)

If ModuleID = wis_BKCC Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True
ElseIf ModuleID = wis_DepositLoans Then
    txtPenal.Enabled = False
    txtMisc.Enabled = False
ElseIf ModuleID = wis_Loans Then
    txtPenal.Enabled = True
    txtMisc.Enabled = True
ElseIf ModuleID = wis_Members Then
    lblRegInt = LoadResString(gLangOffSet + 53) & " " & _
         LoadResString(gLangOffSet + 191) 'Share Fee
    txtPenal.Enabled = False
    txtMisc.Enabled = False
Else
    fraInterest.Enabled = False
End If

lblPenalInt.Enabled = txtPenal.Enabled
lblMisc.Enabled = txtMisc.Enabled

fraInterest.Visible = fraInterest.Enabled

End Sub
Private Sub cmdAccept_Click(Index As Integer)

If Not DateValidate(txtTransDate, "/", True) Then
    MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Sub
End If

If Index = 0 Then
    If Not DepositCash Then Exit Sub
Else

End If
Call cmdCancel_Click(Index)

End Sub

Private Sub cmdCancel_Click(Index As Integer)

cmbAccType(Index).ListIndex = -1
txtAccNo(Index).Text = ""
txtCustName(Index).Text = ""
txtAmount(Index) = 0
txtBalance(Index) = ""

txtRegInt = 0
txtPenal = 0
txtMisc = 0
fraInterest.Visible = False

If Index Then Call LoadPaymentVoucher

End Sub


Private Sub cmdCustName_Click(Index As Integer)

Dim I As Byte

I = IIf(fraReceipt, 1, 0)
I = Index

If cmbAccType(I).ListIndex < 0 Then Exit Sub

On Error GoTo Exit_line

Dim HeadId As wisModules
Dim SqlStr As String

HeadId = cmbAccType(I).ItemData(cmbAccType(I).ListIndex)

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

Dim RstCust As Recordset

MousePointer = vbHourglass
Set RstCust = GetAccRecordSet(HeadId, , IIf(Trim(txtAccNo(I)) = "", txtCustName(I), ""))

If RstCust Is Nothing Then
    MsgBox "There are no customers in the " & cmbAccType(I).Text, _
            vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_line
End If

MousePointer = vbHourglass
Call FillView(m_frmLookUp.lvwReport, RstCust, True)
m_retVar = ""

With m_frmLookUp
    .lvwReport.ColumnHeaders(2).Width = 0
    .Show 1
End With

MousePointer = vbDefault

txtAccNo(I) = m_retVar
'Call txtAccNo_LostFocus(I)

Call EnableAccept
Me.MousePointer = vbDefault
    
Exit_line:

    'cmdFrom.Enabled = True
    'cmdTo.Enabled = True
    MousePointer = vbDefault
    
End Sub

Private Sub cmdDate_Click()
With Calendar
    .Left = cmdDate.Left + cmdDate.Width
    .Top = cmdDate.Top - .Height / 2
    .SelDate = IIf(DateValidate(txtTransDate, "/", True), txtTransDate, gStrDate)
    .Show 1
    txtTransDate = .SelDate
End With

End Sub

Private Sub cmdPay_Click()
Call RefreshPaymentRecordset
If m_rstPayment Is Nothing Then Exit Sub

'Call FillViewNew(m_frmLookUp, m_rstPayment,  "ScrollNO")
If Not FillView(m_frmLookUp, m_rstPayment) Then Exit Sub

m_frmLookUp.Show 1



End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyF6 Then Exit Sub

fraPayment.FontStrikethru = Not fraPayment.FontStrikethru
fraReceipt.FontStrikethru = Not fraReceipt.FontStrikethru

fraPayment.Enabled = Not fraPayment.Enabled
fraReceipt.Enabled = Not fraReceipt.Enabled

cmdAccept(0).Enabled = False
cmdAccept(1).Enabled = False

End Sub

Private Sub LoadAccountBalance()

Dim I As Byte
I = IIf(fraReceipt.Enabled, 0, 1)

If cmbAccType(I).ListIndex < 0 Then Exit Sub

Dim Ret As Integer
Dim BalanceAmount As Currency
Dim AccHeadID As Long
Dim Accid As Long
Dim Rst As ADODB.Recordset
Dim ModuleID As wisModules

'Check for the Selected AccountType if Account type is Sb then
'get the SB balance and if it is CA then Get the CA Balance-siddu
Dim LstInd As Long

With cmbAccType(I)
    AccHeadID = .ItemData(.ListIndex)
End With

ModuleID = GetModuleIDFromHeadID(AccHeadID)

'Check for the Account tyeps
'If it is Current or savings Account then get the Balance
If ModuleID = wis_CAAcc Then
    gDbTrans.SQLStmt = "SELECT AccID FROM CAMaster " & _
        "WHERE AccNum = " & AddQuotes(Trim$(txtAccNo(I).Text), True)
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        'MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        If ActiveControl.Name = cmdAccept(0).Name Or ActiveControl.Name <> cmdClose.Name Then Exit Sub
        'txtAmount.Text = ""
        ActivateTextBox txtAccNo(I)
        Exit Sub
    End If
   
    Accid = FormatField(Rst("AccID"))
    
    'get teh TotalBalance Form the Particular Account
    gDbTrans.SQLStmt = "Select TOP 1 Balance from CATrans where AccID = " & _
        Accid & " order by TransID DESC"
    Ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)
    If Ret <= 0 Then
        MsgBox "No Records"
    Else
        BalanceAmount = FormatField(Rst(0))
        txtBalance(I) = BalanceAmount
        Set Rst = Nothing
    End If
End If

'if the Account is SB Get the Sb Account balance Given from the AccountId
If ModuleID = wis_SBAcc Then
    'AccId=
    gDbTrans.SQLStmt = "Select * from SBMaster" & _
                " where AccNum = " & AddQuotes(txtAccNo(I), True)
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        txtAmount(I).Text = ""
        txtAccNo(I).SetFocus
        Exit Sub
    End If
    
    Accid = FormatField(Rst("AccId"))
    gDbTrans.SQLStmt = "Select TOP 1 Balance from SBTrans where AccID = " & _
        Accid & " order by TransID DESC"
    Ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)
    If Ret <= 0 Then
        MsgBox LoadResString(gLangOffSet + 525), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        BalanceAmount = FormatField(Rst(0))
        txtBalance(I) = BalanceAmount
        Set Rst = Nothing
    End If
    Exit Sub
End If

If I Then Exit Sub

If Not DateValidate(txtTransDate, "/", True) Then Exit Sub

Set Rst = GetAccRecordSet(AccHeadID, txtAccNo(I))
If Rst Is Nothing Then Exit Sub

Accid = FormatField(Rst("ID"))

Dim TransDate As Date
Dim ClsObject As Object

TransDate = FormatDate(txtTransDate)

If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
If ModuleID = wis_BKCC Then
    'Dim ClsObject As clsBkcc
    Set ClsObject = New clsBkcc
    txtRegInt = ClsObject.RegularInterest(Accid, TransDate)
    txtPenal = ClsObject.PenalInterest(Accid, TransDate)
End If
If ModuleID = wis_DepositLoans Then
'    Dim ClsObject As clsDepLoan
    Set ClsObject = New clsDepLoan
    txtRegInt = ClsObject.RegularInterest(Accid, TransDate)
End If
If ModuleID = wis_Loans Then
    'Dim ClsObject As clsLoan
    Set ClsObject = New clsLoan
    txtRegInt = ClsObject.RegularInterest(Accid, , TransDate)
    txtPenal = ClsObject.PenalInterest(Accid, , TransDate)
End If

End Sub

Private Sub Form_Load()

'First cetre the form
Call CenterMe(Me)

'Now Set the kannad action to all controls
Call SetKannadaCaption

'Now Load the Account Type
Call LoadAccountType

txtTransDate = FormatDate(gStrDate)

txtTransDate.Locked = gOnLine
cmdDate.Enabled = Not gOnLine

fraPayment.Enabled = True
fraReceipt.Enabled = False
fraReceipt.FontStrikethru = True
txtScroll(0).Text = GetNextScroll(wDeposit)

fraInterest.Enabled = False
fraInterest.Visible = False



End Sub

Private Sub LoadAccountType()

Dim RstDeposit As Recordset
Dim rstLoan As Recordset

gDbTrans.SQLStmt = "Select * FROM DepositName"
Call gDbTrans.Fetch(RstDeposit, adOpenStatic)
gDbTrans.SQLStmt = "Select * FROM LoanScheme"
Call gDbTrans.Fetch(rstLoan, adOpenStatic)

'Dim ModuleId As wisModules
Dim AccHeadID As Long
Dim ClsBank As clsBankAcc
Dim AccName As String

Set ClsBank = New clsBankAcc


With cmbAccType(0)
    'Withdrawing of the amount willbe only from the
    'account type where cheque facility is given
    'Such type of account are only two(three) types
    'They are
    '1) Saving Bank Account
    '2) Current Account
    '3) Od Account 'Presetly no such accounts are in our s/w

    .Clear
    
    AccName = LoadResString(gLangOffSet + 53) & " " & LoadResString(gLangOffSet + 36) 'Share Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 421) '"SB Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    AccName = LoadResString(gLangOffSet + 422) '"CA AccOUnt
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    AccName = LoadResString(gLangOffSet + 424)  '"RD Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 425) '"Pigmy Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    'Now Load All othe Deposits
    'ModuleId = wis_Deposits
    If Not RstDeposit Is Nothing Then
        RstDeposit.MoveFirst
        While Not RstDeposit.EOF
            AccName = FormatField(RstDeposit("DepositName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID 'ModuleId + FormatField(RstDeposit("DepositID"))
            RstDeposit.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = LoadResString(gLangOffSet + 229) & " " & LoadResString(gLangOffSet + 43) '"BKCC
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    'Now Load All Loans
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        While Not rstLoan.EOF
            AccName = FormatField(rstLoan("SchemeName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID
            rstLoan.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = LoadResString(gLangOffSet + 229) & " " & LoadResString(gLangOffSet + 58) '"BKCC Loan
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    'ModuleId = wis_DepositLoans
    AccName = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 58)
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    '.AddItem AccName
    '.ItemData(.NewIndex) = AccHeadID 'ModuleId
    If Not RstDeposit Is Nothing Then
        RstDeposit.MoveFirst
        While Not RstDeposit.EOF
            AccName = FormatField(RstDeposit("DepositName")) & " " & LoadResString(gLangOffSet + 58)
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID
            RstDeposit.MoveNext
        Wend
    End If
    
End With


With cmbAccType(1)
    .Clear
    
    AccName = LoadResString(gLangOffSet + 53) & " " & LoadResString(gLangOffSet + 36) 'Share Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 421) '"SB Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 422) '"CA AccOUnt
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 424)  '"RD Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    AccName = LoadResString(gLangOffSet + 425) '"Pigmy Account
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    'Now Load All othe Deposits
    'ModuleId = wis_Deposits
    If Not RstDeposit Is Nothing Then
        RstDeposit.MoveFirst
        While Not RstDeposit.EOF
            AccName = FormatField(RstDeposit("DepositName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID 'ModuleId + FormatField(RstDeposit("DepositID"))
            RstDeposit.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = LoadResString(gLangOffSet + 229) & " " & LoadResString(gLangOffSet + 43) '"BKCC
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    'Now Load All Loans
    If Not rstLoan Is Nothing Then
        rstLoan.MoveFirst
        While Not rstLoan.EOF
            AccName = FormatField(rstLoan("SchemeName"))
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID
            rstLoan.MoveNext
        Wend
    End If
    
    'Add Bkcc Deposit Accounts
    AccName = LoadResString(gLangOffSet + 229) & " " & LoadResString(gLangOffSet + 58) '"BKCC Loan
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.NewIndex) = AccHeadID
    End If
    
    'ModuleId = wis_DepositLoans
    AccName = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 58)
    AccHeadID = ClsBank.GetHeadIDCreated(AccName)
    '.AddItem AccName
    '.ItemData(.NewIndex) = AccHeadID 'ModuleId
    If Not RstDeposit Is Nothing Then
        RstDeposit.MoveFirst
        While Not RstDeposit.EOF
            AccName = FormatField(RstDeposit("DepositName")) & " " & LoadResString(gLangOffSet + 58)
            AccHeadID = ClsBank.GetHeadIDCreated(AccName)
            .AddItem AccName
            .ItemData(.NewIndex) = AccHeadID
            RstDeposit.MoveNext
        Wend
    End If
    
End With


End Sub

Private Sub SetKannadaCaption()

Dim Ctrl As Control
On Error Resume Next
'Now Assign the Kannada fonts to the All controls
For Each Ctrl In Me
    If Not TypeOf Ctrl Is ProgressBar And _
        Not TypeOf Ctrl Is VScrollBar And _
            Not TypeOf Ctrl Is Line And _
                Not TypeOf Ctrl Is Image Then
                    Ctrl.Font.Name = gFontName
                    
                    If Not TypeOf Ctrl Is ComboBox Then
                        Ctrl.Font.Size = gFontSize
                    Else
                        Ctrl.Font.Size = Ctrl.Font.Size + 1
                    End If
                    
    End If
Next Ctrl

    
' TransCtion Frame
'fraFrom.Caption = LoadResString(gLangOffSet + 107)
'lbltransDate.Caption = LoadResString(gLangOffSet + 37)
lblScroll(0).Caption = LoadResString(gLangOffSet + 41)
lblAccType(0).Caption = LoadResString(gLangOffSet + 36) & " " & _
                                LoadResString(gLangOffSet + 35)
lblAccNo(0).Caption = LoadResString(gLangOffSet + 36) & " " & _
                            LoadResString(gLangOffSet + 60)
lblAmount(0).Caption = LoadResString(gLangOffSet + 40)
lblRegInt = LoadResString(gLangOffSet + 344)
lblPenalInt = LoadResString(gLangOffSet + 345)
lblMisc = LoadResString(gLangOffSet + 327)

fraPayment.Caption = LoadResString(gLangOffSet + 108)
lblScroll(1).Caption = LoadResString(gLangOffSet + 41)
lblAccType(1).Caption = LoadResString(gLangOffSet + 36) & " " & _
                LoadResString(gLangOffSet + 35)
lblAccNo(1).Caption = LoadResString(gLangOffSet + 36) & " " & _
                LoadResString(gLangOffSet + 60)
lblAmount(1).Caption = LoadResString(gLangOffSet + 40)

'optPrincipal.Caption = LoadResString(gLangOffSet + 310)
'optRegInt.Caption = LoadResString(gLangOffSet + 344)
'optPenalInt.Caption = LoadResString(gLangOffSet + 345)
'
'cmdAccept.Caption = LoadResString(gLangOffSet + 4)
'cmdUndo.Caption = LoadResString(gLangOffSet + 14)
'
'cmdSave.Caption = LoadResString(gLangOffSet + 7)
'cmdClose.Caption = LoadResString(gLangOffSet + 11)
'cmdClear.Caption = LoadResString(gLangOffSet + 8)    '
'
End Sub


Private Sub fraPayment_Click()
If Not fraPayment.Enabled Then Call Form_KeyDown(vbKeyTab, 0)
End Sub

Private Sub fraReceipt_Click()
If Not fraReceipt.Enabled Then Call Form_KeyDown(vbKeyTab, 0)
End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)

    m_retVar = strSelection

End Sub


Private Sub optAccount_Click()
With fraInterest
   If .Enabled Then .Visible = True
End With
End Sub

Private Sub optPassing_Click()
    fraInterest.Visible = False
End Sub


Private Sub txtAccNo_LostFocus(Index As Integer)

'Check for the Account Number
If txtAccNo(Index).Text = "" Then Exit Sub
'Now Get the Account No
Dim Rst As Recordset
Dim AccHeadID As Long

AccHeadID = cmbAccType(Index).ItemData(cmbAccType(Index).ListIndex)

Set Rst = GetAccRecordSet(AccHeadID, Trim(txtAccNo(Index)))

If Rst Is Nothing Then Exit Sub
txtCustName(Index) = FormatField(Rst("CustName"))

Call LoadAccountBalance

Call EnableAccept

End Sub


Private Sub txtAmount_Change(Index As Integer)
Call EnableAccept
End Sub

