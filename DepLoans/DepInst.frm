VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmDepLoanInst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dposit Loan Payment"
   ClientHeight    =   4485
   ClientLeft      =   2265
   ClientTop       =   1890
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   400
      Left            =   6510
      TabIndex        =   13
      Top             =   3990
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   400
      Left            =   5190
      TabIndex        =   16
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   3825
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox cmbSB 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtInterest 
         Height          =   315
         Left            =   2190
         TabIndex        =   10
         Top             =   2090
         Width           =   1245
      End
      Begin VB.TextBox txtSbAccNum 
         Height          =   315
         Left            =   6270
         TabIndex        =   19
         Top             =   3330
         Width           =   765
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "..."
         Height          =   315
         Left            =   7200
         TabIndex        =   22
         Top             =   2610
         Width           =   315
      End
      Begin VB.CheckBox chkSb 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit To SB Account"
         Height          =   300
         Left            =   150
         TabIndex        =   20
         Top             =   3360
         Width           =   3225
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2190
         TabIndex        =   6
         Top             =   970
         Width           =   1215
      End
      Begin WIS_Currency_Text_Box.CurrText txtSanctioned 
         Height          =   345
         Left            =   6270
         TabIndex        =   12
         Top             =   970
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtInterestamount 
         Height          =   345
         Left            =   6270
         TabIndex        =   15
         Top             =   1530
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtIssuedAmount 
         Height          =   345
         Left            =   6270
         TabIndex        =   18
         Top             =   2610
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMisc 
         Height          =   345
         Left            =   6270
         TabIndex        =   23
         Top             =   2090
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblMisc 
         Caption         =   "Miscellaneous"
         Height          =   300
         Left            =   3600
         TabIndex        =   24
         Top             =   2090
         Width           =   1545
      End
      Begin VB.Label txtLoan 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2190
         TabIndex        =   8
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblSbAccNo 
         Caption         =   "Sb Account No"
         Height          =   300
         Left            =   3990
         TabIndex        =   21
         Top             =   3330
         Width           =   2205
      End
      Begin VB.Line Line2 
         X1              =   150
         X2              =   7440
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   7440
         Y1              =   735
         Y2              =   720
      End
      Begin VB.Label txtAccNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acc NO"
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   345
         Left            =   4230
         TabIndex        =   4
         Top             =   240
         Width           =   3195
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   300
         Left            =   2430
         TabIndex        =   3
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   970
         Width           =   1815
      End
      Begin VB.Label lblPrevLoanAmt 
         Caption         =   "Previous loan amount : "
         Height          =   300
         Left            =   180
         TabIndex        =   7
         Top             =   1530
         Width           =   1875
      End
      Begin VB.Label lblLoanSanctioned 
         Caption         =   "Sanctioned amount : "
         Height          =   300
         Left            =   3510
         TabIndex        =   11
         Top             =   970
         Width           =   2265
      End
      Begin VB.Label lblRateofIntForLoans 
         Caption         =   "Rate of interest for loans:"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   2090
         Width           =   1905
      End
      Begin VB.Label lblDepositNo 
         Caption         =   "Account No: "
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblLessIntOnPrevLoan 
         Caption         =   "Less interest on previous loans: "
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3510
         TabIndex        =   14
         Top             =   1530
         Width           =   2295
      End
      Begin VB.Label lblTotAmtIssued 
         Caption         =   "Total amount to be issued:"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3480
         TabIndex        =   17
         Top             =   2610
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmDepLoanInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_LoanID As Long
Private m_Loaded As Boolean
Private m_CustID As Long

Private M_setUp As New clsSetup
Private m_DepositType As Integer
Private m_LoanHeadID As Long
Private m_IntHeadID As Long

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private Sub LoadAccountDetails()
    'Now Load The User Name And Othe Details
    Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "Select * from DepositLoanMaster where " & _
            " LoanID =  " & m_LoanID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    
    m_DepositType = FormatField(rst("DepositType"))
    Dim AccHeadName As String
    Dim AccHeadNameEnglish As String
    If m_DepositType Then
        m_DepositType = m_DepositType Mod 100
        AccHeadName = GetDepositTypeText(m_DepositType) & " " & GetResourceString(58)
    Else
        AccHeadName = GetResourceString(43, 58)
        AccHeadNameEnglish = LoadResourceStringS(43, 58)
    End If
    
    Dim bankClass As clsBankAcc
    gDbTrans.BeginTrans
    Set bankClass = New clsBankAcc
    m_LoanHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepLoan, _
            0, wis_DepositLoans + m_DepositType)
    'Interest head
    AccHeadName = AccHeadName & " " & GetResourceString(483)
    AccHeadNameEnglish = AccHeadNameEnglish & " " & LoadResString(483)
    m_IntHeadID = bankClass.GetHeadIDCreated(AccHeadName, AccHeadNameEnglish, parMemDepLoanIntReceived, _
        0, wis_DepositLoans + m_DepositType)
    gDbTrans.CommitTrans
    Set bankClass = Nothing
    
    Dim CustClass As clsCustReg
    Set CustClass = New clsCustReg
    txtName = CustClass.CustomerName(rst("customerId"))
    txtAccNo = rst("AccNum")
    Set CustClass = Nothing
    Set rst = Nothing

'Get the due date & Closed Date
    Dim ClosedDate As Date
    gDbTrans.SqlStmt = "Select * from DepositLoanMaster where " & _
            " LoanID =  " & m_LoanID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    m_CustID = FormatField(rst("CustomerID"))
    
    'LoanClosedDate is not in the table.
    ClosedDate = IIf(FormatField(rst("LoanClosedDate")) = "", "1/1/100", rst("LoanClosedDate"))
    
    'Update command buttons only if deposit is not closed
    cmdAccept.Enabled = Not CBool(FormatField(rst("LoanClosed")))
    
    txtInterest = FormatField(rst("InterestRate"))

    'Get The Total Deposited Amount
    Dim LoanAmount As Currency
    Dim I As Integer
    
    gDbTrans.SqlStmt = "Select Top 1 * from DepositLoanTrans " & _
            " where LoanID = " & m_LoanID & " order by TransID Desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub
    LoanAmount = FormatField(rst("Balance"))
    txtLoan = FormatCurrency(LoanAmount)
    
End Sub

Public Property Let LoanID(NewVal As Long)
m_LoanID = NewVal

If m_LoanID And m_Loaded Then Call LoadAccountDetails

End Property


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblDepositNo.Caption = GetResourceString(43, 60)
lblDate.Caption = GetResourceString(37)
lblRateofIntForLoans.Caption = GetResourceString(186) 'rATE F iNTEREST
lblPrevLoanAmt.Caption = GetResourceString(246)
lblLoanSanctioned.Caption = GetResourceString(247)
lblLessIntOnPrevLoan.Caption = GetResourceString(248)
lblTotAmtIssued.Caption = GetResourceString(249)
cmdAccept.Caption = GetResourceString(4)
cmdCancel.Caption = GetResourceString(11)
Call SetDepositCheckBoxCaption(wis_SBAcc, chkSb, cmbSb) 'chkSb.Caption = GetResourceString(421, 271)
lblMisc.Caption = GetResourceString(327)

End Sub

Private Sub chkSb_Click()

If m_LoanID = 0 Then Exit Sub
If chkSb.Value = vbUnchecked Then
    cmdSb.Enabled = False
    txtSbAccNum.Enabled = False
    Exit Sub
End If

cmdSb.Enabled = True
txtSbAccNum.Enabled = True

Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT AccNum,DepositType From SBMASTER " & _
    " WHERE CustomerId = (SELECT CustomerID From DepositLoanMaster " & _
        " WHERE LoanID = " & m_LoanID & ")"

Dim recCount As Integer
    recCount = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If recCount = 1 Then
        txtSbAccNum = FormatField(rst(0))
        If cmbSb.Visible Then Call SetComboIndex(cmbSb, , FormatField(rst("DepositTYpe")))
    End If
    If recCount > 1 Then
        If Not cmbSb.Visible Then
            txtSbAccNum = FormatField(rst(0))
            Exit Sub
        End If
        
        If cmbSb.ListIndex < 0 Then
            MsgBox "Please select deposit type", , wis_MESSAGE_TITLE
            Exit Sub
        Else
            gDbTrans.SqlStmt = "SELECT AccNum From SBMASTER " & _
                " WHERE DepositType = " & cmbSb.ItemData(cmbSb.ListIndex) & _
                " And CustomerId = (SELECT CustomerID From DepositLoanMaster " & _
                " WHERE LoanID = " & m_LoanID & ")"
            recCount = gDbTrans.Fetch(rst, adOpenForwardOnly)
            If recCount > 0 Then txtSbAccNum = FormatField(rst(0))
        End If
    End If

End Sub

Private Sub cmdAccept_Click()

'Check out the sanctioned amount
If txtSanctioned = 0 Then
    'MsgBox "Invalid amount sanctioned !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    txtSanctioned.SetFocus
    Exit Sub
End If

Dim TransDate As Date
Dim rst As ADODB.Recordset

'Check out the date
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date of transaction not in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
TransDate = GetSysFormatDate(txtDate.Text)

'See if account has already been matured or closed
gDbTrans.SqlStmt = "Select * from DepositLoanMaster where LoanID = " & m_LoanID
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
    'MsgBox "Error accessing data base !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(601), vbCritical, gAppName & " - Error"
    Exit Sub
End If

If FormatField(rst("LoanClosedDate")) <> "" Then
    'MsgBox "This deposit has already been closed !", vbExclamation, gAppName & " - Error"
    'MsgBox GetResourceString(524), vbExclamation, gAppName & " - Error"
    'Exit Sub
End If

'Check date range w.r.t to loan
gDbTrans.SqlStmt = "Select TOP 1 TransDate from DepositLoanTrans where " & _
        " LoanID = " & m_LoanID & " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    If DateDiff("D", rst("TransDate"), TransDate) < 0 Then
        'MsgBox "Date of transaction is lesser than the previous transaction date", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If
End If

'Check out if the interest rate is valid
If txtInterest <= 0 Then
    'MsgBox "Interest rate has not been specified for this period." & vbCrLf & vbCrLf & "Please set the value of interest for this period in the properties of this account !", vbInformation, gAppName & " - Error"
    MsgBox GetResourceString(579) & vbCrLf & vbCrLf & GetResourceString(659), vbInformation, gAppName & " - Error"
    Exit Sub
End If

    
'Get New transaction ID
    Dim TransID As Long
    Dim transType As wisTransactionTypes
    Dim Balance As Currency
    Dim IntBalance As Currency
    Dim LoanAmount As Currency
    Dim IntAmount As Currency
    Dim MiscAmount As Currency
    Dim MiscHeadId As Long
    
    
    LoanAmount = txtSanctioned
    IntAmount = txtInterestamount
    MiscAmount = txtMisc
    
    MiscHeadId = parIncome + 1 'Misceleneous
   
    gDbTrans.SqlStmt = "Select TOP 1 TransID,Balance " & _
                    " From DepositLoanIntTrans " & _
                    " where LoanID = " & m_LoanID & _
                    " Order by TransID desc"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        IntBalance = Val(FormatField(rst("Balance")))
        TransID = FormatField(rst("TransID"))
    End If
    
    gDbTrans.SqlStmt = "Select TOP 1 TransID,Balance " & _
                    " from DepositLoanTrans " & _
                    " where LoanID = " & m_LoanID & _
                    " Order by TransID desc"
            
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        Balance = Val(FormatField(rst("Balance")))
        If TransID < rst("TransID") Then TransID = rst("TransID")
    End If
    Balance = Balance + LoanAmount
    TransID = TransID + 1

'Now Get the Calculated Interest till date

    Dim ContraID As Long
    Dim SbACCID As Long
    Dim SBClass As clsSBAcc
    Dim SbHeadID As Long
    Dim bankClass As clsBankAcc
    
If chkSb.Value = vbChecked Then
    ContraID = GetMaxContraTransID + 1

    If Trim(txtSbAccNum) = "" Then
        'MsgBox "Invalid Account no sepcified", vbInformation, "SB Account"
        MsgBox GetResourceString(500), vbInformation, "SB Account"
        GoTo ExitLine
    End If
    
    'Check For the SB Number Existace
    Set SBClass = New clsSBAcc
    SbACCID = 0
    If cmbSb.Visible Then SbACCID = cmbSb.ItemData(cmbSb.ListIndex)
    SbACCID = SBClass.GetAccountID(txtSbAccNum, SbACCID)
    Set SBClass = Nothing
    
    If SbACCID = 0 Then
        'MsgBox "Invalid Account no sepcified", vbInformation, "SB Account"
        MsgBox GetResourceString(500), vbInformation, "SB Account"
        ActivateTextBox cmbSb
        GoTo ExitLine
    End If
End If

Dim InTrans As Boolean
'Start data base transactions
gDbTrans.BeginTrans
InTrans = True

    'Now insert record of the new loan
    transType = wWithdraw
    If chkSb = vbChecked Then transType = wContraWithdraw
    If transType = wWithdraw And LoanAmount = (IntAmount + MiscAmount) Then transType = wContraWithdraw
    
    gDbTrans.SqlStmt = "Insert into DepositLoanTrans " & _
                    " (LoanID, TransID, TransType, " & _
                    " TransDate, Amount, Balance, Particulars,UserID) values ( " & _
                    m_LoanID & "," & _
                    TransID & "," & _
                    transType & "," & _
                    "#" & TransDate & "#," & _
                    txtSanctioned & "," & _
                    Balance & ", " & _
                    "'To Loans' ," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    'First insert any interest of previous loans
    transType = wDeposit
    If chkSb = vbChecked Then transType = wContraDeposit
    If transType = wDeposit And LoanAmount = IntAmount + MiscAmount Then transType = wContraDeposit
    
    If IntAmount + MiscAmount > 0 Then
        Dim C_IntAmount As Currency
        C_IntAmount = ComputeDepLoanRegularInterest(TransDate, m_LoanID)
        If IntAmount < C_IntAmount Then
            If MsgBox(GetResourceString(668) & vbCrLf & GetResourceString(669), _
                    vbYesNo, wis_MESSAGE_TITLE) = vbYes Then _
            IntBalance = IntBalance + C_IntAmount - txtInterestamount
        End If
        gDbTrans.SqlStmt = "Insert into DepositLoanIntTrans (LoanID, " & _
                        " TransID, TransType, " & _
                        " TransDate,Amount,MiscAmount, Balance, " & _
                        " Particulars,UserID) values ( " & _
                        m_LoanID & "," & _
                        TransID & "," & _
                        transType & "," & _
                        "#" & TransDate & "#," & _
                        IntAmount & "," & _
                        MiscAmount & "," & _
                        IntBalance & ", " & _
                        "'By interest' ," & gUserID & ")"
            
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
    End If
    
    Set bankClass = New clsBankAcc
     If MiscAmount > 0 Then
        MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), _
                    parBankIncome, 0, wis_None)
    End If
    If chkSb.Value = vbChecked Then
        transType = wContraWithdraw
        Dim SBType As Integer
        Dim headName As String
        Dim headNameEnglish As String
        If cmbSb.Visible Then
            SBType = cmbSb.ItemData(cmbSb.ListIndex)
            headName = GetDepositName(wis_SBAcc, SBType, headNameEnglish)
            SbHeadID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parMemberDeposit, 0, wis_SBAcc + SBType)
        Else
            SbHeadID = bankClass.GetHeadIDCreated(GetResourceString(421), LoadResString(421), parMemberDeposit, 0, wis_SBAcc + SBType)
        End If
        
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    " (ContraID, AccHeadID,AccID," & _
                    " TransType,TransID,Amount,UserID )" & _
                    " VALUES (" & ContraID & "," & _
                    m_LoanHeadID & "," & _
                    m_LoanID & "," & transType & ", " & _
                    TransID & "," & txtIssuedAmount & "," & gUserID & " )"
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        transType = wContraDeposit
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    " (ContraID, AccHeadID,AccID," & _
                    " TransType,TransID,Amount,UserID )" & _
                    " VALUES (" & ContraID & "," & _
                    m_IntHeadID & "," & _
                    m_LoanID & "," & transType & ", " & _
                    TransID & "," & IntAmount & "," & gUserID & " )"
        
        'If He is repaying any interest then Only mak this transaction
        If IntAmount > 0 Then If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    " (ContraID, AccHeadID,AccID," & _
                    " TransType,TransID,Amount,UserID )" & _
                    " VALUES (" & ContraID & "," & _
                    MiscHeadId & "," & _
                    m_LoanID & "," & transType & ", " & _
                    TransID & "," & MiscAmount & "," & gUserID & " )"
        
        'If He is repaying any interest then Only mak this transaction
        If MiscAmount > 0 Then If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        Set SBClass = New clsSBAcc
        If SBClass.DepositAmount(SbACCID, txtIssuedAmount, _
                    "from Dep Loan AccNo " & txtAccNo, TransDate, " ") = 0 Then GoTo ExitLine
        
        'now make the receipt or payment to the ledger heads
        If txtInterestamount > 0 Then
            If Not bankClass.UpdateContraTrans(m_LoanHeadID, m_IntHeadID, _
                        txtInterestamount, TransDate) Then GoTo ExitLine
        End If
        If MiscAmount > 0 Then
            If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
        End If
        If Not bankClass.UpdateContraTrans(m_LoanHeadID, SbHeadID, _
                                txtIssuedAmount, TransDate) Then GoTo ExitLine
        
        Set SBClass = Nothing
    Else
        If transType = wDeposit Then
            If IntAmount > 0 Then
                If Not bankClass.UpdateCashDeposits(m_IntHeadID, IntAmount, TransDate) _
                                Then GoTo ExitLine
            End If
            If Not bankClass.UpdateCashWithDrawls(m_LoanHeadID, LoanAmount, TransDate) _
                                Then GoTo ExitLine
            If MiscAmount > 0 Then
                If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                            MiscAmount, TransDate) Then GoTo ExitLine
            End If
        Else
            'if he has updated the interest
            'and not issued any loan
            If Not bankClass.UpdateContraTrans(m_LoanHeadID, m_IntHeadID, _
                                IntAmount, TransDate) Then GoTo ExitLine
            If MiscAmount > 0 Then
                If Not bankClass.UpdateContraTrans(m_LoanHeadID, _
                        MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
            End If
        End If
    End If
    
    Set bankClass = Nothing
    
    'If transaction is cash withdraw & there is casier window
    'then transfer the While Amount cashier window
    If transType = wWithdraw And gCashier Then
        Dim Cashclass As clsCash
        Set Cashclass = New clsCash
        If Cashclass.TransferToCashier(m_LoanHeadID, _
                        m_LoanID, TransDate, TransID, txtSanctioned) < 1 Then GoTo ExitLine
        Set Cashclass = Nothing
    End If

'COmmit transactions
gDbTrans.CommitTrans
InTrans = False

Unload Me

ExitLine:

If InTrans Then gDbTrans.RollBack

'Udate date with todays date (By default)
    txtDate.Text = gStrDate
    Me.txtSanctioned = 0
'Update the details on the UI
    Call UpdateUserInterface

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUndo_Click()

Dim TransID As Long
Dim rst As ADODB.Recordset

'Get the last transaction ID
    gDbTrans.SqlStmt = "Select Top 1 TransID, TransType from DepositLoanTrans " & _
                        " where AccID = " & m_LoanID & " order by TransID desc"
    'Call gDBTrans.SQLFetch
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "You do not have any loans on this deposit !", vbInformation, gAppName & " - Message"
        MsgBox GetResourceString(582), vbInformation, gAppName & " - Message"
        Exit Sub
    End If
    TransID = FormatField(rst("TransID"))

'COnfirm about transaction
'If MsgBox("Are you sure you want to undo a previous loan transaction ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
If MsgBox(GetResourceString(583), _
    vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then _
    Exit Sub

gDbTrans.BeginTrans
    
    'Remove the last transaction
    gDbTrans.SqlStmt = "Delete from DepositLoanTrans where AccID = " & m_LoanID & _
                        " and TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
        MsgBox GetResourceString(609), vbCritical, gAppName & " - Critical Error"
        gDbTrans.RollBack
        Exit Sub
    End If
    'Check out the transaction before the last transaction, because it may the interest
    'added. Since we are performing interest charges automatically, we've got to remove
    'this also automatically
    'May Not Necessary
    
    gDbTrans.SqlStmt = "Delete from DepositLoanIntTrans where AccID = " & m_LoanID & _
                        " and TransID = " & TransID
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
        MsgBox GetResourceString(609), vbCritical, gAppName & " - Critical Error"
        gDbTrans.RollBack
        Exit Sub
    End If
gDbTrans.CommitTrans

'Udate date with todays date (By default)
    txtDate.Text = gStrDate

Call UpdateUserInterface
End Sub

Private Sub cmdSB_Click()
If m_LoanID = 0 Then Exit Sub

Dim SearchName As String
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT FirstName,MiddleName,LastName " & _
    " FROM NAMETAB WHERE CustomerID = " & m_CustID
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    SearchName = FormatField(rst(0))
    If SearchName = "" Then SearchName = FormatField(rst(1))
    If SearchName = "" Then SearchName = FormatField(rst(2))
End If
SearchName = InputBox("Enter the Name to search", "Account Search", SearchName)

Screen.MousePointer = vbHourglass
On Error GoTo Err_line
Dim Lret As Long

' Query the database to get all the customer names...
With gDbTrans
    .SqlStmt = "SELECT AccNum," & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name " & _
        " FROM SBMaster A, NameTab B WHERE A.CustomerID = B.CustomerID "
    If Trim(SearchName) <> "" Then
        .SqlStmt = .SqlStmt & " AND (FirstName like '" & SearchName & "%' " & _
            " Or MiddleName like '" & SearchName & "%' " & _
            " Or LastName like '" & SearchName & "%')"
        .SqlStmt = .SqlStmt & " Order by IsciName"
    Else
        .SqlStmt = .SqlStmt & " Order by AccNum"
    End If
    Lret = .Fetch(rst, adOpenStatic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        GoTo Exit_Line
    End If
End With

' Create a report dialog.
If m_LookUp Is Nothing Then Set m_LookUp = New frmLookUp

With m_LookUp
    .m_SelItem = ""
    ' Fill the data to report dialog.
    If Not FillView(.lvwReport, rst, True) Then
        'MsgBox "Error filling the customer details.", vbCritical
        MsgBox "Error filling the customer details.", vbCritical
        GoTo Exit_Line
    End If
    Screen.MousePointer = vbDefault
    ' Display the dialog.
    .Show vbModal
End With

Exit_Line:
    Screen.MousePointer = vbDefault
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

'Centre the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    txtDate.Text = gStrDate

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'set kannada fonts
Call SetKannadaCaption
'Initialize the grid
    
    
'Obtain the rate of interest as applicable to this deposit
    Dim Days As Long
    Dim rst As ADODB.Recordset

    gDbTrans.SqlStmt = "Select * from DepositLoanMaster where LoanID = " & m_LoanID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "This account does not exists!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
    'Check if deposit is closed
    If FormatField(rst("LOanClosed")) Then
        cmdAccept.Enabled = False
    Else
        cmdAccept.Enabled = True
    End If

'Now Load The User Name And Othe Details
'Dim Rst As ADODB.Recordset
gDbTrans.SqlStmt = "Select * from DepositLoanMaster where " & _
        " LoanID =  " & m_LoanID
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then Exit Sub

Dim CustClass As clsCustReg
Set CustClass = New clsCustReg
txtName = CustClass.CustomerName(rst("customerId"))
txtAccNo = rst("AccNum")
Set CustClass = Nothing

Set rst = Nothing
m_Loaded = True

If m_LoanID <> 0 Then Call LoadAccountDetails
Call UpdateUserInterface

If gOnLine Then txtDate.Locked = True


End Sub

Private Sub UpdateUserInterface()

'Calculate the interest for loan if a loan has been drawn previously
Dim IntAmount As Currency

On Error Resume Next
'See if deposit has matured
If Val(txtLoan) <= 0 Then Exit Sub

IntAmount = ComputeDepLoanRegularInterest(GetSysFormatDate(txtDate.Text), m_LoanID)
txtInterestamount = (IntAmount \ 1)
    
Err.Clear

End Sub


Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
m_Loaded = False
End Sub

Private Sub Form_Unload(cancel As Integer)
'""(Me.hwnd, False)
'Set frmPDLoans = Nothing
End Sub



Private Sub m_LookUp_SelectClick(strSelection As String)
txtSbAccNum = strSelection
End Sub


Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub

Call UpdateUserInterface
If Me.ActiveControl.name = cmdAccept.name Then txtIssuedAmount.SetFocus

End Sub


Private Sub txtInterestAmount_Change()
'COmpute the total amount to be actully issued
txtIssuedAmount = txtSanctioned.Text - txtInterestamount - txtMisc

End Sub

Private Sub txtInterestAmount_GotFocus()
With txtInterestamount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtIssuedAmount_GotFocus()
With txtIssuedAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtMisc_Change()
txtIssuedAmount = txtSanctioned - txtInterestamount - txtMisc
End Sub

Private Sub txtSanctioned_Change()
'COmpute the total amount to be actully issued
txtIssuedAmount = txtSanctioned - txtInterestamount - txtMisc
End Sub


Private Sub txtSanctioned_GotFocus()
With txtSanctioned
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


