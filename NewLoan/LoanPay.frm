VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmLoanPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Payment"
   ClientHeight    =   6885
   ClientLeft      =   1080
   ClientTop       =   1500
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6330
      TabIndex        =   39
      Top             =   6270
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5040
      TabIndex        =   38
      Top             =   6270
      Width           =   1215
   End
   Begin VB.Frame fraLoan 
      Height          =   2265
      Left            =   90
      TabIndex        =   40
      Top             =   30
      Width           =   7455
      Begin VB.Label txtLoanBalance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   13
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label txtLoanDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5790
         TabIndex        =   43
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label txtCustID 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5790
         TabIndex        =   42
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label txtLoanAccNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   41
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   400
         Left            =   1890
         TabIndex        =   4
         Top             =   810
         Width           =   5325
      End
      Begin VB.Label txtLoanName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   6
         Top             =   1350
         Width           =   5325
      End
      Begin VB.Label lblLoanDate 
         Caption         =   "Loan Date :"
         Height          =   300
         Left            =   3960
         TabIndex        =   8
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label lblLoanBalance 
         Caption         =   "Loan Balance :"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label lblLoanName 
         Caption         =   "Loan  Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   1350
         Width           =   1545
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "Loan Account No :"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label lblName 
         Caption         =   "Loan Holder Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label lblCustID 
         Caption         =   "Customer ID :"
         Height          =   300
         Left            =   3960
         TabIndex        =   2
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame fraInstall 
      Height          =   3975
      Left            =   90
      TabIndex        =   0
      Top             =   2160
      Width           =   7455
      Begin VB.TextBox txtVoucherNo 
         Height          =   345
         Left            =   5790
         TabIndex        =   16
         Top             =   210
         Width           =   1425
      End
      Begin VB.CommandButton cmdMisc 
         Caption         =   "..."
         Height          =   315
         Left            =   3000
         TabIndex        =   19
         Top             =   1740
         Width           =   315
      End
      Begin VB.TextBox txtPayAmount 
         Height          =   345
         Left            =   1890
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1425
      End
      Begin VB.TextBox txtNewBalance 
         Height          =   345
         Left            =   5790
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1425
      End
      Begin VB.CheckBox chkSb 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit To SB Account"
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   3420
         Width           =   3135
      End
      Begin VB.TextBox txtSbAccNum 
         Height          =   315
         Left            =   5760
         TabIndex        =   36
         Top             =   3420
         Width           =   1185
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "..."
         Height          =   315
         Left            =   6990
         TabIndex        =   35
         Top             =   3420
         Width           =   315
      End
      Begin VB.CheckBox chkInterest 
         Alignment       =   1  'Right Justify
         Caption         =   "&Deduct Interest amount"
         Height          =   300
         Left            =   3960
         TabIndex        =   26
         Top             =   1740
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.TextBox txtRemark 
         Height          =   345
         Left            =   5790
         TabIndex        =   33
         Top             =   2790
         Width           =   1425
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtDate 
         Height          =   345
         Left            =   1890
         TabIndex        =   37
         Top             =   240
         Width           =   1425
      End
      Begin WIS_Currency_Text_Box.CurrText txtIntBalance 
         Height          =   345
         Left            =   1890
         TabIndex        =   12
         Top             =   660
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPenalBalance 
         Height          =   345
         Left            =   5790
         TabIndex        =   15
         Top             =   630
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtInterest 
         Height          =   345
         Left            =   1890
         TabIndex        =   18
         Top             =   1230
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPenal 
         Height          =   345
         Left            =   5790
         TabIndex        =   21
         Top             =   1230
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtNewLoan 
         Height          =   345
         Left            =   1890
         TabIndex        =   28
         Top             =   2310
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMisc 
         Height          =   345
         Left            =   1890
         TabIndex        =   24
         Top             =   1710
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblVoucherNo 
         Caption         =   "Voucher No:"
         Height          =   300
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   1545
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   7305
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   7305
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   7305
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lblPenalBalance 
         Caption         =   "&Penal balance"
         Height          =   300
         Left            =   3960
         TabIndex        =   14
         Top             =   690
         Width           =   1725
      End
      Begin VB.Label lblPenal 
         Caption         =   "&Penal Interest "
         Height          =   300
         Left            =   3960
         TabIndex        =   20
         Top             =   1320
         Width           =   1875
      End
      Begin VB.Label lblMisc 
         Caption         =   "&Miscelaneous"
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   1770
         Width           =   1425
      End
      Begin VB.Label lblPayAmount 
         Caption         =   "&Amount paid :"
         Height          =   300
         Left            =   120
         TabIndex        =   31
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label lblIntBalance 
         Caption         =   "Interest &balance"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label lblNewBalance 
         Caption         =   "&New Balance :"
         Height          =   300
         Left            =   3960
         TabIndex        =   30
         Top             =   2400
         Width           =   1845
      End
      Begin VB.Label lblInterest 
         Caption         =   "&Interest till date"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label lblRemarks 
         Caption         =   "&Remarks"
         Height          =   300
         Left            =   3960
         TabIndex        =   32
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label lblAmtSanctioned 
         Caption         =   "Amount &Credited"
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label lblDate 
         Caption         =   "&Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLoanPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_LoanID As Long
Private m_SchemeName As String
Private m_SchemeNameEnglish As String
Private m_SchemeId As Integer
Private m_FormLoaded As Boolean
Private m_AccHeadId As Long
Private m_clsReceivable As clsReceive

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private m_retVar As Variant


Public Event AccountTransaction()
Public Event WindowClosed()

Public Property Let LoanAccountID(NewValue As Long)
    
    m_LoanID = NewValue
    If m_LoanID = 0 Then Exit Property
    If m_FormLoaded Then Call UpdateDetails
    
End Property


Private Function LoanTransaction() As Boolean

'Declarariotn Of HeadId as
Dim LoanHeadID As Long
Dim RegIntHeadID As Long
Dim PenalIntHeadID As Long
Dim MiscHeadId As Long
Dim Trans As wisTransactionTypes
Dim VoucherNo As String

'First Validate TextBoxes
Dim Amount As Currency
Dim TransDate As Date

TransDate = GetSysFormatDate(txtDate)
VoucherNo = Trim(txtVoucherNo.Text)

'Now Get the Transctin detail sof this loan
Dim RstLoanTrans As Recordset
Dim bankClass As clsBankAcc
Dim UserID As Integer

UserID = gCurrUser.UserID

gDbTrans.SqlStmt = "SELECT top 1 * From LoanTrans " & _
                " Where LoanId = " & m_LoanID & _
                " Order By TransId DESC"
Dim Balance As Currency

If gDbTrans.Fetch(RstLoanTrans, adOpenForwardOnly) > 0 Then _
                        Balance = FormatField(RstLoanTrans("Balance"))
    

Dim SBClass As clsSBAcc
Dim SbACCID As Long
Dim ContraID As Long
Dim rst As Recordset

If chkSb.Value = vbChecked Then
    If Trim(txtSbAccNum) <> "" Then
        'check the existance of sbaccount
        Set SBClass = New clsSBAcc
        SbACCID = SBClass.GetAccountID(txtSbAccNum)
        Set SBClass = Nothing
        If SbACCID = 0 Then
            'Invalid Account NO
            MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
            GoTo Err_line
        End If
        'Get the Contra TransCtion Id
        ContraID = GetMaxContraTransID + 1
    End If
End If

Dim Remarks As String
Dim LoanAmount As Currency
Dim IntAmount As Currency
Dim PenalInt As Currency
Dim MiscAmount As Currency
Dim IntBalance As Currency
Dim PenalBalance As Currency

Dim SqlStr As String

Dim transType  As wisTransactionTypes
Dim count As Integer

UserID = gCurrUser.UserID

Remarks = Me.txtRemark
If Len(Remarks) > 50 Then Remarks = Left(txtRemark, 50)
transType = wWithdraw
LoanAmount = txtNewLoan

'Now UpDate The Interest Balance In LoanMaster
If chkInterest.Value = vbChecked Then
    IntAmount = txtInterest
    PenalInt = txtPenal
    IntBalance = txtIntBalance - IntAmount + Val(txtInterest.Tag)
    PenalBalance = txtPenalBalance - PenalInt + Val(txtPenal.Tag)
    MiscAmount = txtMisc
    If Not m_clsReceivable Is Nothing Then MiscAmount = m_clsReceivable.TotalAmount
Else
    IntAmount = 0
    PenalInt = 0 'vsal(txtPenal)
    IntBalance = txtInterest + txtIntBalance
    PenalBalance = txtPenal + txtPenalBalance
    MiscAmount = txtMisc
    'Check whether user is entering the Receivable Amount
    If Not m_clsReceivable Is Nothing And LoanAmount = 0 Then MiscAmount = m_clsReceivable.TotalAmount
End If
If MiscAmount = 0 Then Set m_clsReceivable = Nothing

If Not m_clsReceivable Is Nothing And (LoanAmount <= MiscAmount) Then
    If LoanAmount And LoanAmount <> MiscAmount Then
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    Else
        LoanTransaction = TransactMiscAmount
    End If
    Exit Function
End If


'Get the maximum TransId before incrementing the transid
Dim TransID As Long
TransID = GetLoanMaxTransID(m_LoanID) + 1

'Begin the Trasnction
Dim InTrans As Boolean
gDbTrans.BeginTrans
InTrans = True

Set bankClass = New clsBankAcc
LoanHeadID = bankClass.GetHeadIDCreated(m_SchemeName, m_SchemeNameEnglish, _
                    parMemberLoan, 0, wis_Loans + m_SchemeId)

If LoanAmount <> 0 Then
    transType = IIf(SbACCID, wContraWithdraw, wWithdraw)
    Balance = Balance + LoanAmount
    
    'Insert INto the Loan Principle table.
    SqlStr = "Insert Into LoanTrans(LoanId,TransID,TransType," & _
            "Amount,TransDate, Balance,VoucherNo,UserID,Particulars) Values (" & _
            m_LoanID & ", " & _
            TransID & "," & _
            transType & "," & _
            LoanAmount & "," & _
            "#" & TransDate & "#," & _
            Balance & "," & _
            AddQuotes(VoucherNo) & "," & _
            UserID & ", " & AddQuotes(Remarks, True) & ")"
        
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If transType = wWithdraw Then
        If LoanAmount <> (IntAmount + PenalInt + MiscAmount) Then
            If Not bankClass.UpdateCashWithDrawls(LoanHeadID, LoanAmount, _
                TransDate) Then GoTo Err_line
        End If
    Else
        'Now Insert into Contra Trans
        SqlStr = "INSERT INTO ContraTrans " & _
                "(ContraID, AccHeadID,AccID, " & _
                "TransType ,TransID, Amount,VoucherNo,UserID) " & _
                "VALUES ( " & ContraID & "," & LoanHeadID & "," & _
                m_LoanID & "," & transType & ", " & TransID & "," & _
                LoanAmount & ",' '," & gUserID & "  )"
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo Err_line
    End If
End If

'Now insert the interest amount
If IntAmount > 0 Or PenalInt > 0 Or MiscAmount > 0 Then
    Dim engHeadName As String
    If Len(m_SchemeNameEnglish) > 0 Then engHeadName = m_SchemeNameEnglish & " " & LoadResString(344)
    RegIntHeadID = bankClass.GetHeadIDCreated(m_SchemeName & " " & GetResourceString(344), _
        engHeadName, parMemLoanIntReceived, 0, wis_Loans + m_SchemeId)
    
    If Len(m_SchemeNameEnglish) > 0 Then engHeadName = m_SchemeNameEnglish & " " & LoadResString(345)
    PenalIntHeadID = bankClass.GetHeadIDCreated(m_SchemeName & " " & GetResourceString(345), _
        engHeadName, parMemLoanPenalInt, 0, wis_Loans + m_SchemeId)
    If MiscAmount > 0 Then
        If m_clsReceivable Is Nothing Then
            MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), parBankIncome, 0, wis_None)
        End If
    End If
    'Insert Into the Loan Interest Table
    transType = IIf(SbACCID, wContraDeposit, wDeposit)
    If LoanAmount = (IntAmount + PenalInt + MiscAmount) Then transType = wContraDeposit
    
    SqlStr = "Insert Into LoanIntTrans(LoanId," & _
            "TransID,TransType," & _
            "IntAmount,PenalIntAmount,MiscAmount," & _
            "IntBalance,PenalIntBalance,TransDate,VoucherNo,UserId) Values (" & _
            m_LoanID & ", " & _
            TransID & "," & _
            transType & "," & _
            IntAmount & "," & _
            PenalInt & "," & _
            MiscAmount & "," & _
            IntBalance & "," & _
            PenalBalance & "," & _
            " #" & TransDate & "#, " & _
            AddQuotes(VoucherNo) & "," & _
            UserID & ")"
        
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    
    If SbACCID Then
        'Now Insert interest amount into Contra Trans
        SqlStr = "INSERT INTO ContraTrans " & _
                "(ContraID, AccHeadID,AccID, " & _
                "TransType ,TransID, Amount,VoucherNo,UserID) " & _
                "VALUES ( " & ContraID & "," & PenalIntHeadID & "," & _
                m_LoanID & "," & transType & ", " & TransID & "," & _
                Amount & "," & AddQuotes(VoucherNo) & "," & gUserID & " )"
            
        If IntAmount > 0 Then
            gDbTrans.SqlStmt = SqlStr
            If Not gDbTrans.SQLExecute Then GoTo Err_line
            If Not bankClass.UpdateContraTrans(LoanHeadID, _
                    RegIntHeadID, IntAmount, TransDate) Then GoTo Err_line
        End If
        
        'Now Insert Penal interest amount into Contra Trans
        If PenalInt > 0 Then
            SqlStr = "INSERT INTO ContraTrans " & _
                    "(ContraID, AccHeadID,AccID, " & _
                    "TransType ,TransID, Amount,VoucherNo,UserID) " & _
                    "VALUES ( " & ContraID & "," & PenalIntHeadID & "," & _
                    m_LoanID & "," & transType & ", " & TransID & "," & _
                    Amount & "," & AddQuotes(VoucherNo) & "," & gUserID & " )"
            gDbTrans.SqlStmt = SqlStr
            If Not gDbTrans.SQLExecute Then GoTo Err_line
            If Not bankClass.UpdateContraTrans(LoanHeadID, _
                    PenalIntHeadID, PenalInt, TransDate) Then GoTo Err_line
        End If
        'Now Insert misc amount into Contra Trans
        If MiscAmount > 0 Then
            Do
                If Not m_clsReceivable Is Nothing Then Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
                If MiscHeadId = 0 Then Exit Do
                SqlStr = "INSERT INTO ContraTrans " & _
                        "(ContraID, AccHeadID,AccID, " & _
                        "TransType ,TransID, Amount,VoucherNo,userID) " & _
                        "VALUES ( " & ContraID & "," & MiscHeadId & "," & _
                        m_LoanID & "," & transType & ", " & TransID & "," & _
                        Amount & "," & AddQuotes(VoucherNo) & "," & gUserID & " )"
                gDbTrans.SqlStmt = SqlStr
                If Not gDbTrans.SQLExecute Then GoTo Err_line
                If Not bankClass.UpdateContraTrans(LoanHeadID, _
                        MiscHeadId, MiscAmount, TransDate) Then GoTo Err_line
            Loop
        End If
    Else
        If LoanAmount = (IntAmount + PenalInt + MiscAmount) Then
            If Not bankClass.UpdateContraTrans(LoanHeadID, _
                    RegIntHeadID, IntAmount, TransDate) Then GoTo Err_line
            If Not bankClass.UpdateContraTrans(LoanHeadID, _
                    PenalIntHeadID, PenalInt, TransDate) Then GoTo Err_line
            If MiscAmount > 0 Then
                Do
                    Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
                    If MiscHeadId = 0 Then Exit Do
                    If Not bankClass.UpdateContraTrans(LoanHeadID, MiscHeadId, _
                                         MiscAmount, TransDate) Then GoTo Err_line
                Loop
            End If
        
        Else
            If Not bankClass.UpdateCashDeposits(RegIntHeadID, _
                                IntAmount, TransDate) Then GoTo Err_line
            If Not bankClass.UpdateCashDeposits(PenalIntHeadID, _
                                PenalInt, TransDate) Then GoTo Err_line
            If MiscAmount > 0 Then
                If m_clsReceivable Is Nothing Then
                    If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                                    MiscAmount, TransDate) Then GoTo Err_line
                Else
                  Do
                    Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
                    If MiscHeadId = 0 Then Exit Do
                    If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                                    MiscAmount, TransDate) Then GoTo Err_line
                  Loop
                End If
            End If
        
        End If
    End If
End If

If SbACCID Then
    Amount = txtPayAmount
    If chkInterest.Value = vbChecked Then Amount = Amount - PenalInt - IntAmount
    'Now Pay the Amount SHHead
    
    Dim SbHeadID As Long
    SbHeadID = bankClass.GetHeadIDCreated(GetResourceString(421), LoadResString(421), parMemberDeposit, 0, wis_SBAcc)
    
    If Not bankClass.UpdateContraTrans(LoanHeadID, SbHeadID, Amount, TransDate) Then GoTo Err_line
                            
    Set SBClass = New clsSBAcc
    If SBClass.DepositAmount(SbACCID, Amount, "Loan Advance", TransDate, VoucherNo) = 0 Then GoTo Err_line
    
End If

'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(m_AccHeadId, _
            m_LoanID, TransDate, TransID, Amount) < 1 Then GoTo Err_line
    
End If

gDbTrans.CommitTrans
InTrans = False

LoanTransaction = True
    
Exit_Line:

    If InTrans Then gDbTrans.RollBack: InTrans = False
    Set bankClass = Nothing
    Set SBClass = Nothing
    Set Cashclass = Nothing
    
    Exit Function

Err_line:
    
    MsgBox "Error In Loan pay Installment"
    GoTo Exit_Line

End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblLoanAccNo = GetResourceString(58, 36, 60)
lblCustID = GetResourceString(49, 60) 'Customer ID
lblName = GetResourceString(35) 'Name
lblLoanName = GetResourceString(214) 'loanScheme
lblLoanBalance = GetResourceString(58, 42) 'Loan Balance
lblLoanDate = GetResourceString(58, 37) 'Loan Date

lblDate = GetResourceString(37) 'Date
lblVoucherNo = GetResourceString(41, 60)
lblInterest = GetResourceString(344)
lblMisc = GetResourceString(327) 'Misceleneous
lblAmtSanctioned = GetResourceString(247) 'Sanctioned Amount
lblPayAmount.Caption = GetResourceString(375, 40) 'Payable Amount
lblRemarks = GetResourceString(261) 'Remarks
lblIntBalance = GetResourceString(67, 344) 'Reg Interest balance
lblPenalBalance = GetResourceString(67, 345) 'penal Interest balance
lblPenal = GetResourceString(345) '"PenalInterest
lblNewBalance = GetResourceString(260, 42) 'New BAlance
chkInterest.Caption = GetResourceString(308)  'Deduct Interest
'chk.Caption = GetResourceString(309)  'Deduct Legal Fee
chkSb.Caption = GetResourceString(421, 271)

cmdCancel.Caption = GetResourceString(2)  'Cancel
cmdOk.Caption = GetResourceString(1)      'OK

End Sub


Private Function TransactMiscAmount() As Boolean

On Error GoTo Err_line
 
'Dim LastTransDate As Date
Dim TransID As Long
Dim inTransaction As Boolean

Dim NewBalance As Currency
Dim Balance As Currency

Dim MiscAmount As Currency
Dim PrincAmount  As Currency

Dim TotalAmount As Currency
Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim Deposit As Boolean
Dim rst As Recordset
Dim Amount As Currency
    
Dim bankClass As clsBankAcc
Dim MiscHeadId As Long
Dim LoanHeadID As Long

Dim VoucherNo As String
Dim UserID As Integer
Dim strParticulars As String
Dim rstTemp As Recordset


'Get the Voucher No and Cheque No
UserID = gCurrUser.UserID
VoucherNo = Trim(txtVoucherNo.Text)

MiscAmount = txtMisc 'txtPenalInt = txtPenalInterest.Value

transType = wContraWithdraw
TransDate = GetSysFormatDate(txtDate.Text)

TotalAmount = txtNewLoan.Value


'Now check whether the amount collecting is
'Deducting or addinng to the Loan head or not
Dim IsAddingToAccount As Boolean
Dim I As Integer
Dim OtherTransType As wisTransactionTypes

IsAddingToAccount = m_clsReceivable.IsAddToAccount
If txtNewLoan = MiscAmount Then IsAddingToAccount = True

'If all the amount is adding to the Account then
OtherTransType = wContraDeposit

'now get the Balance and Get a new transactionID.
gDbTrans.SqlStmt = "SELECT Top 1 Balance,TransID,TransDate " & _
            " FROM LoanTrans WHERE Loanid = " & m_LoanID & _
            " ORDER BY TransID Desc"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Balance = Val(FormatField(rst("Balance")))
    TransID = Val(FormatField(rst("TransID")))
End If

'Begin the transactionhere
gDbTrans.BeginTrans
inTransaction = True
'First insert then Amount ot the Accounthead
LoanHeadID = GetHeadID(m_SchemeName, parMemberLoan)
If LoanHeadID = 0 Then Exit Function

TransID = TransID + 1
If IsAddingToAccount Then
    If bankClass Is Nothing Then Set bankClass = New clsBankAcc
    LoanHeadID = bankClass.GetHeadIDCreated(m_SchemeName, m_SchemeNameEnglish, parMemberLoan, 0, wis_Loans + m_SchemeId)

    Deposit = IIf(Balance < 0, True, False)
    'If customer is paying the amount to the bank,reduece the balance
    Balance = Balance + TotalAmount
    
    strParticulars = GetResourceString(327)
    gDbTrans.SqlStmt = "INSERT INTO LoanTrans " _
                & "(LoanID, TransID, TransType, Amount, " _
                & " TransDate, Balance, VoucherNO,UserId, Particulars) " _
                & "VALUES (" & m_LoanID & ", " & TransID & ", " _
                & transType & ", " & TotalAmount & ", " _
                & "#" & TransDate & "#, " _
                & Balance & ", " _
                & AddQuotes(VoucherNo) & "," & UserID & "," _
                & AddQuotes(strParticulars) & " )"
    
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    'Now get the Contra ID
    Dim ContraID As Long
    ContraID = GetMaxContraTransID + 1
    gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                    "( ContraID,AccHeadId,AccID," & _
                    " TransType,TransID,Amount," & _
                    " UserID,VoucherNo) VALUES (" & _
                    ContraID & "," & _
                    LoanHeadID & "," & m_LoanID & "," & _
                    transType & "," & TransID & "," & _
                    TotalAmount & "," & UserID & "," & _
                    AddQuotes(VoucherNo) & ")"
    If Not gDbTrans.SQLExecute Then GoTo ExitLine

End If

'Now Get the pending Balance From the Previous account
Balance = 0
gDbTrans.SqlStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & LoanHeadID & _
        " ANd AccId = " & m_LoanID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    rstTemp.MoveLast
    Balance = FormatField(rstTemp("Balance"))
End If

I = 0
Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
Do While MiscHeadId > 0
    'Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, Amount, MiscAmount)
    If MiscHeadId = 0 Then Exit Do
    
    If IsAddingToAccount Then
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(MiscHeadId) & "," & MiscHeadId & "," & _
                        OtherTransType & "," & TransID & "," & _
                        MiscAmount & "," & UserID & "," & _
                        AddQuotes(VoucherNo) & ")"
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        'Insert the Same Into Head Transaction class
        If Not bankClass.UpdateContraTrans(LoanHeadID, _
                        MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
    End If
    
    Balance = Balance + MiscAmount
    
    'Now insert this details Into amount receivable table
    If Not AddToAmountReceivable(LoanHeadID, m_LoanID, TransID, TransDate, MiscAmount, MiscHeadId) Then GoTo ExitLine
    gDbTrans.SqlStmt = "Insert Into AmountReceivAble" & _
                    "( AccHeadId,AccID," & _
                    " TransType,TransDate,TransID,Amount," & _
                    " Balance,UserID,DueHeadID) VALUES (" & _
                    LoanHeadID & "," & m_LoanID & "," & _
                    OtherTransType & ",#" & TransDate & "#," & TransID & "," & _
                    MiscAmount & "," & Balance & "," & UserID & "," & _
                    MiscHeadId & ")"
    
    'If Not gDbTrans.SQLExecute Then GoTo ExitLine
    I = I + 1
    Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
Loop

gDbTrans.CommitTrans

inTransaction = False

'Call ResetTransactionForm

TransactMiscAmount = True

Exit Function

ExitLine:
    
    If inTransaction Then gDbTrans.RollBack
    Exit Function
    
Err_line:

    MsgBox "Error in Misc transaction", vbInformation, wis_MESSAGE_TITLE
    
End Function
Private Sub UpdateDetails()

If m_LoanID = 0 Then
    Err.Raise 2002, "Loan Payment", "You have not set the LoanId or Account Id"
    Exit Sub
End If


Dim Retval As Long
Dim rstMaster As Recordset
Dim rstTrans As Recordset
Dim rstTemp As Recordset
'Set LoanMaster Recpords
gDbTrans.SqlStmt = "Select * From LoanMaster Where LoanId = " & m_LoanID
Call gDbTrans.Fetch(rstMaster, adOpenStatic)
txtLoanAccNo = FormatField(rstMaster("AccNum"))
txtLoanDate = FormatField(rstMaster("IssueDate"))

'Set the Member Details
txtCustID = GetMemberNumber(FormatField(rstMaster("CustomerId")))
m_SchemeId = FormatField(rstMaster("SchemeId"))

'Set LoanTrans Records
gDbTrans.SqlStmt = "SELECT * From LoanTrans Where LoanId = " & m_LoanID & _
                " Order By TransId "
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then
    txtNewLoan = FormatField(rstMaster("LoanAmount"))
    chkInterest.Value = vbUnchecked
    chkInterest.Enabled = False
    txtDate = txtLoanDate
    txtIntBalance.Enabled = False
    txtPenalBalance.Enabled = False
    txtInterest.Enabled = False
    txtPenal.Enabled = False
Else
    rstTrans.MoveLast
    txtLoanBalance = FormatField(rstTrans("Balance"))
    txtNewBalance = txtLoanBalance
End If

gDbTrans.SqlStmt = "SELECT * From LoanIntTrans Where " & _
            " LoanId = " & m_LoanID & " Order By TransId"
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then
    chkInterest.Value = vbUnchecked
    chkInterest.Enabled = False
Else
    rstTrans.MoveLast
    txtIntBalance = FormatField(rstTrans("IntBalance"))
    txtPenalBalance = FormatField(rstTrans("PenalIntBalance"))
    txtIntBalance.Tag = txtIntBalance
    txtPenalBalance.Tag = txtPenalBalance
End If

'Set LoanScheme  Records
gDbTrans.SqlStmt = "SELECT * From LoanScheme Where Schemeid = " & m_SchemeId
Call gDbTrans.Fetch(rstTemp, adOpenStatic)

m_SchemeName = FormatField(rstTemp("SchemeName"))
m_SchemeNameEnglish = FormatField(rstTemp("SchemeNameEnglish"))
txtLoanName = m_SchemeName

Dim l_CustClass As New clsCustReg
txtName = l_CustClass.CustomerName(rstMaster("CustomerID"))
Set l_CustClass = Nothing

End Sub

Private Function ValidControls() As Boolean
'First Validate TextBoxes
If Not DateValidate(txtDate, "/", True) Then
    'MsgBox "Invalid date Specified", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    GoTo Exit_Line
End If
'Validate Vocucher NO
If Trim(txtVoucherNo) = "" Then
    'MsgBox "Invalid Voucher Specified", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(755), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtVoucherNo
    GoTo Exit_Line
End If

If DateDiff("d", GetSysFormatDate(txtDate), GetLoanLastTransDate(m_LoanID)) > 0 Then
    'MsgBox "Specified transaction date is earlier than previous transaction date", _
        vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(572), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    GoTo Exit_Line
End If

If txtNewLoan <= 0 Then
    'The user may be marking the receivable amount so chck for it
    If m_clsReceivable Is Nothing Then
        'MsgBox "Invalid amount specified", vbExclamation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), vbExclamation, wis_MESSAGE_TITLE
        txtNewLoan.SetFocus
        GoTo Exit_Line
    End If
End If
If txtMisc < 0 Then
    'MsgBox "Invalid amount specified", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtMisc
    GoTo Exit_Line
End If
If txtInterest < 0 Then
    'MsgBox "Invalid amount specified", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtInterest
    GoTo Exit_Line
End If
If txtPenal < 0 Then
    'MsgBox "Invalid amount specified", vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtPenal
    GoTo Exit_Line
End If

ValidControls = True

Exit_Line:

End Function

Private Sub chkInterest_Click()
Call txtNewLoan_Change
Dim chk As Boolean
chk = IIf(chkInterest, False, True)

txtInterest.Locked = chk
txtPenal.Locked = chk

End Sub

Private Sub chkSb_Click()

If chkSb.Value = vbUnchecked Then Exit Sub
If m_LoanID = 0 Then Exit Sub

Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT AccNum From SBMASTER " & _
    " WHERE CustomerId = (SELECT CustomerID From LoanMAster " & _
        " WHERE LoanID = " & m_LoanID & ")"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    txtSbAccNum = FormatField(rst(0))

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDate_Click()
With Calendar
    
    .Left = Me.Left + fraInstall.Left + cmdDate.Left - (.Width / 2)
    .Top = Me.Top + fraInstall.Top + cmdDate.Top
    If DateValidate(txtDate.Text, "/", True) Then
        .selDate = txtDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtDate.Text = .selDate
End With

End Sub

Private Sub cmdMisc_Click()

If m_LoanID = 0 Then Exit Sub
If m_clsReceivable Is Nothing Then _
            Set m_clsReceivable = New clsReceive
    
m_clsReceivable.Show
    
If m_clsReceivable.TotalAmount Then
    txtMisc.Locked = True
    txtMisc.Value = m_clsReceivable.TotalAmount
Else
    txtMisc.Locked = False
    Set m_clsReceivable = Nothing
End If

End Sub

Private Sub cmdOk_Click()

If Not ValidControls Then Exit Sub

If LoanTransaction Then Unload Me
    
End Sub


Private Sub cmdSB_Click()
If m_LoanID = 0 Then Exit Sub

Dim SearchName As String
Dim rst As Recordset

Me.MousePointer = vbHourglass

gDbTrans.SqlStmt = "SELECT FirstName,MiddleName,LastName FROM NAMETAB " & _
    " WHERE CustomerId = (SELECT CustomerID From LoanMaster " & _
        " WHERE LoanID = " & m_LoanID & ")"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    SearchName = FormatField(rst(0))
    If SearchName = "" Then SearchName = FormatField(rst(1))
    If SearchName = "" Then SearchName = FormatField(rst(2))
End If
SearchName = InputBox("Enter the Name to search", "Account Search", SearchName)

Me.MousePointer = vbHourglass
On Error GoTo Err_line
Dim Lret As Long

' Query the database to get all the customer names...
With gDbTrans
    .SqlStmt = "SELECT AccNum, " & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name " & _
        " FROM SBMaster A, NameTab B WHERE A.CustomerID = B.CustomerID "
    If Trim(SearchName) <> "" Then
        .SqlStmt = .SqlStmt & " AND (FirstName like '" & SearchName & "%' " & _
            " Or MiddleName like '" & SearchName & "%' " & _
            " Or LastName like '" & SearchName & "%')"
        .SqlStmt = .SqlStmt & " Order by IsciName"
    Else
        .SqlStmt = .SqlStmt & " Order by A.CustomerId"
    End If
    Lret = .Fetch(rst, adOpenStatic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        GoTo Exit_Line
    End If
End With

'Create a report dialog.
If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

With m_LookUp
    .m_SelItem = ""
    ' Fill the data to report dialog.
    If Not FillView(.lvwReport, rst, True) Then
        'MsgBox "Error filling the customer details.", vbCritical
        MsgBox "Error filling the customer details.", vbCritical
        GoTo Exit_Line
    End If
    
    ' Display the dialog.
    m_retVar = ""
    .Show vbModal
    txtSbAccNum = m_retVar
End With

Exit_Line:
    
    Me.MousePointer = vbDefault
    Exit Sub

Err_line:
    
    If Err Then
        'MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
        MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
    End If
    GoTo Exit_Line

End Sub

Private Sub Form_Load()
'set icon for the form caption
'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption
txtName.FONTSIZE = txtName.FONTSIZE + 1
'Call SetKannadaCaption
txtDate.Text = gStrDate

m_FormLoaded = True
If m_LoanID Then Call UpdateDetails
If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If


End Sub



Private Sub Form_Unload(Cancel As Integer)
m_FormLoaded = False
End Sub


Private Sub lblPenalBalance_Change()
Call txtInterest_Change
End Sub

Private Sub m_LookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub


Private Sub txtDate_Change()
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
If txtDate.Text = txtDate.Tag Then Exit Sub
txtDate.Tag = txtDate.Text

Dim RegInterest As Currency
Dim PenalInterest As Currency
' To Get Calulation of Interest we Hve to Set the repay Date Of Loan Acc form To the PresentCalculation Date
Dim L_class As New clsLoan
Dim AsOnDate As Date
AsOnDate = GetSysFormatDate(txtDate.Text)
RegInterest = L_class.RegularInterest(m_LoanID, , AsOnDate)
PenalInterest = L_class.PenalInterest(m_LoanID, , AsOnDate)
PenalInterest = IIf(PenalInterest < 0, 0, PenalInterest)
PenalInterest = PenalInterest + RegInterest
txtInterest = PenalInterest \ 1
End Sub


Private Sub txtIntBalance_LostFocus()
Call txtInterest_Change
End Sub

Private Sub txtInterest_Change()

If chkInterest.Value = vbChecked Then
    txtPayAmount = FormatCurrency(txtNewLoan - txtIntBalance - txtInterest - txtPenal - txtMisc)
Else
    txtPayAmount = FormatCurrency(txtNewLoan - txtMisc)
End If

End Sub

Private Sub txtInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtMisc_Click()
Call txtInterest_Change
End Sub

Private Sub txtNewLoan_Change()

txtNewBalance = FormatCurrency(txtNewLoan + Val(txtLoanBalance))
Call txtInterest_Change

End Sub

Private Sub txtNewLoan_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtPenal_Change()
Call txtInterest_Change
End Sub

Private Sub txtPenalBalance_LostFocus()
Call txtInterest_Change
End Sub


Private Sub txtRemark_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtSbAccNum_LostFocus()
Call txtInterest_Change
End Sub

