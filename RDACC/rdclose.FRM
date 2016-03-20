VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmRDClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RDClose"
   ClientHeight    =   4635
   ClientLeft      =   1215
   ClientTop       =   2505
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdTfr 
      Caption         =   "Transfer"
      Height          =   450
      Left            =   4050
      TabIndex        =   13
      Top             =   4020
      Width           =   1335
   End
   Begin VB.Frame fraCharges 
      Caption         =   "Charges"
      Height          =   1485
      Left            =   4200
      TabIndex        =   25
      Top             =   2400
      Width           =   4095
      Begin WIS_Currency_Text_Box.CurrText txtCharges 
         Height          =   345
         Left            =   2850
         TabIndex        =   27
         Top             =   390
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTax 
         Height          =   345
         Left            =   2850
         TabIndex        =   30
         Top             =   840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblPreClose 
         Caption         =   "Premature closure charges: "
         Height          =   300
         Left            =   210
         TabIndex        =   26
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label lblOthers 
         Caption         =   "Other charges (Tax, etc.)"
         Height          =   300
         Left            =   210
         TabIndex        =   29
         Top             =   900
         Width           =   2145
      End
   End
   Begin VB.Frame fraLoanDet 
      Caption         =   "Loan Details"
      Height          =   2415
      Left            =   4200
      TabIndex        =   14
      Top             =   30
      Width           =   4095
      Begin VB.CheckBox ChkDeductLoan 
         Alignment       =   1  'Right Justify
         Caption         =   "Deduct Loan Amount :"
         Height          =   300
         Left            =   210
         TabIndex        =   24
         Top             =   1830
         Width           =   3765
      End
      Begin VB.TextBox txtLoanRate 
         Height          =   315
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   850
         Width           =   1125
      End
      Begin WIS_Currency_Text_Box.CurrText txtLoanInterest 
         Height          =   345
         Left            =   2850
         TabIndex        =   21
         Top             =   1340
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtLoanAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2850
         TabIndex        =   17
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblIntOnLoan 
         Caption         =   "Interest on loans: "
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   1340
         Width           =   2145
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Total loan amount:"
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lblLoanInt 
         Caption         =   "Rate of interest:"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   850
         Width           =   2115
      End
   End
   Begin VB.Frame fraDepDetail 
      Caption         =   "Deposit Details"
      Height          =   3855
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4215
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2550
         TabIndex        =   2
         Text            =   "12/12/2002"
         Top             =   390
         Width           =   1245
      End
      Begin VB.TextBox txtIntRate 
         Height          =   315
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   1312
         Width           =   1125
      End
      Begin WIS_Currency_Text_Box.CurrText txtIntpayable 
         Height          =   345
         Left            =   2880
         TabIndex        =   8
         Top             =   1770
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepIntAmount 
         Height          =   345
         Left            =   2880
         TabIndex        =   10
         Top             =   2265
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotalInt 
         Height          =   345
         Left            =   2880
         TabIndex        =   12
         Top             =   2760
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPayableAmount 
         Height          =   345
         Left            =   2430
         TabIndex        =   23
         Top             =   3420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Line Line1 
         X1              =   4110
         X2              =   150
         Y1              =   3228
         Y2              =   3228
      End
      Begin VB.Label lblNetAmount 
         Caption         =   "Net payable amount:"
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   3420
         Width           =   2145
      End
      Begin VB.Label txtDepositAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2910
         TabIndex        =   4
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest on deposit:"
         Height          =   300
         Left            =   270
         TabIndex        =   9
         Top             =   2282
         Width           =   2145
      End
      Begin VB.Label lblIntPayable 
         Caption         =   "Interest From Payable"
         Height          =   300
         Left            =   270
         TabIndex        =   7
         Top             =   1809
         Width           =   2145
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposited amount:"
         Height          =   300
         Left            =   270
         TabIndex        =   3
         Top             =   863
         Width           =   2115
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   300
         Left            =   270
         TabIndex        =   1
         Top             =   390
         Width           =   2115
      End
      Begin VB.Label lblDepInt 
         Caption         =   "Rate of Interest:"
         Height          =   300
         Left            =   270
         TabIndex        =   5
         Top             =   1336
         Width           =   2115
      End
      Begin VB.Label lblTotalInt 
         Caption         =   "Total Interest on deposit:"
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   2755
         Width           =   2145
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   6990
      TabIndex        =   28
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   450
      Left            =   5550
      TabIndex        =   32
      Top             =   4020
      Width           =   1335
   End
   Begin VB.Label lblDepDetail 
      Caption         =   "Deposit details:"
      Height          =   225
      Left            =   60
      TabIndex        =   31
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmRDClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public p_AccID As Long
Private M_setUp As New clsSetup
Private m_AccHeadId As Long
Private m_RdPayableId As Long
Private m_RdIntHeadId As Long

Private m_Transfer As Boolean

Private m_ContraClass As clsContra
'
Private Function RDClose() As Boolean

On Error GoTo ErrLine

Dim TransDate As Date
Dim rst As ADODB.Recordset
Dim Days As Integer
Dim MatDate As Date
Dim transType  As wisTransactionTypes
Dim TransID As Long
Dim Balance As Currency
Dim PayableBalance As Currency
'Dim IntBalance As Currency
Dim PayableAmount As Currency
Dim IntAmount As Currency
Dim Amount As Currency
Dim TotalIntAmount As Currency
Dim ContraID As Long


TransDate = GetSysFormatDate(txtDate.Text)

TransID = GetRDMaxTransID(p_AccID)

'Do not allow if this deposit has loans
If Val(txtLoanAmount) > 0 Then
    'MsgBox "The deposit you are trying to close has loans against it" & vbCrLf & _
            "You must first get the loan repayment and then close this deposit", vbInformation, gAppName & " - Message"
    MsgBox GetResourceString(574) & vbCrLf & _
            GetResourceString(575), vbInformation, gAppName & " - Message"
            Exit Function
End If

'Warn for premature closure
gDbTrans.SqlStmt = "Select MaturityDate From RDMaster where " & _
                    " AccID = " & p_AccID
If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
    'MsgBox "Deposit not found !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(658), vbCritical, gAppName & " - Error"
    Exit Function
End If

MatDate = rst("MaturityDate")
Days = DateDiff("D", MatDate, TransDate)
If Days < 0 Then
    'If MsgBox("You are attempting to close this deposit prematurely !" & vbCrLf & vbCrLf & "Are you sure you want to continue this operation ?", vbQuestion + vbYesNo + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then
    If MsgBox(GetResourceString(576) & vbCrLf & vbCrLf & GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then
        Unload Me
        Exit Function
    End If
End If

'Now Compare the Date of Transaction w.r.t Last transaction date
If DateDiff("D", TransDate, GetRDLastTransDate(p_AccID)) > 0 Then
    'Date of transction is earlier than Last date of transction
    'If MsgBox(GetResourceString(572) & vbCrLf & vbCrLf & _
            GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, _
            gAppName & " - Confirmation") = vbNo Then
        
     '   Unload Me
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Now Compare the Transction date w.r.t To Last Transction Of RD
gDbTrans.SqlStmt = "Select TOP 1 Balance from RDTrans " & _
             " where AccID = " & p_AccID & " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then Balance = FormatField(rst("Balance"))


'Now Compare the Transction date w.r.t To Last Transction Of Interest payable
gDbTrans.SqlStmt = "Select TOP 1 Balance from RDIntPayable " & _
        " where AccID = " & p_AccID & " order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then PayableBalance = FormatField(rst("Balance"))
    
Set rst = Nothing
PayableAmount = Val(txtIntPayable)
IntAmount = txtDepIntAmount - Val(txtCharges.Text)
TotalIntAmount = Val(txtTotalInt)
Amount = Val(txtDepositAmount)
'Now Get the INterest Amount deposited in InterestPayble Account
PayableBalance = PayableBalance - PayableAmount
If PayableBalance < 0 Then
    'with drawig more amount from payableaccount
    If MsgBox(GetResourceString(577) & vbCrLf & GetResourceString(541) & "?", _
        vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    PayableBalance = 0
End If

'if transferring the amoun then get the COntra ID
If m_Transfer Then ContraID = GetMaxContraTransID

Balance = Balance - Amount
If Balance < 0 Then Exit Function

Dim UserID As Integer
'userid=gCurrUser.userid

Dim BankCls As clsBankAcc

'Begin the TransCtion
Dim InTrans As Boolean

'Get free TransID
TransID = GetRDMaxTransID(p_AccID) + 1

gDbTrans.BeginTrans
InTrans = True

Set BankCls = New clsBankAcc
If PayableAmount Then
    'First WithDraw the Amount from Payable Account
    transType = wWithdraw
    If m_Transfer Then transType = wContraWithdraw
    
    gDbTrans.SqlStmt = "INSERT INTO RDIntPayable ( ACCID, TransID, TransDate," & _
            " Amount,TransType,Balance,Particulars,UserID )" & _
            " VALUES (" & _
            p_AccID & "," & _
            TransID & ", " & _
            " #" & TransDate & "#," & _
            PayableAmount & "," & _
            transType & "," & _
            PayableBalance & "," & _
            "'RD Closed', " & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If transType = wWithdraw Then
        If Not BankCls.UpdateCashWithDrawls(m_RdPayableId, _
                    PayableAmount, TransDate) Then GoTo ErrLine
    Else
        'Insert into the contras table
        gDbTrans.SqlStmt = "INSERT INTO ContraTrans " & _
                " (ContraID,AccHeadID,AccID,TransType, " & _
                " TransID, Amount,UserID ) VALUES " & _
                "( " & ContraID & "," & m_RdPayableId & "," & _
                p_AccID & "," & transType & "," & TransID & "," & _
                PayableAmount & "," & gCurrUser.UserID & ")"
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
End If

    
'Then WithDraw the Amount from Interest Account
If IntAmount Then
    transType = IIf(m_Transfer, wContraWithdraw, wWithdraw)
    
    If IntAmount < 0 Then _
        transType = IIf(transType = wWithdraw, wDeposit, wContraDeposit)
        
    gDbTrans.SqlStmt = "INSERT INTO RDIntTrans " & _
                    " (ACCID, TransID, TransDate," & _
                    " Amount,TransType,Balance,Particulars,UserID )" & _
                    " VALUES (" & p_AccID & "," & _
                    TransID & ", " & _
                    " #" & TransDate & "#," & _
                    IntAmount & "," & _
                    transType & ", 0 ," & _
                    "'RD Closed' , " & _
                    UserID & " )"

    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    'Update to the bankAccount
    If transType = wWithdraw Then
        If Not BankCls.UpdateCashWithDrawls(m_RdIntHeadId, IntAmount, _
                    TransDate) Then GoTo ErrLine
    ElseIf transType = wDeposit Then
        If Not BankCls.UpdateCashDeposits(m_RdIntHeadId, Abs(IntAmount), _
                    TransDate) Then GoTo ErrLine
    Else
        'Insert into the contras table
        gDbTrans.SqlStmt = "INSERT INTO COntraTrans " & _
            " (ContraID,AccHeadID,AccID,TransType, " & _
            " TransID, Amount,UserID ) VALUES " & _
            "( " & ContraID & "," & m_RdIntHeadId & "," & _
            p_AccID & "," & transType & "," & TransID & "," & _
            IntAmount & "," & gCurrUser.UserID & ")"
        If Not gDbTrans.SQLExecute Then GoTo ErrLine

    End If
End If

'Then WithDraw the Amount from RD Account
transType = wWithdraw
If m_Transfer Then transType = wContraWithdraw
gDbTrans.SqlStmt = "INSERT INTO RDTrans ( ACCID, TransID, TransDate," & _
                " Amount,TransType,Balance,Particulars,UserID )" & _
                " VALUES (" & _
                p_AccID & "," & _
                TransID & ", " & _
                " #" & TransDate & "#," & _
                Amount & "," & _
                transType & "," & _
                Balance & "," & _
                "'RD Closed' ," & _
                UserID & " )"

If Amount Then
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    If transType = wWithdraw Then
        If Not BankCls.UpdateCashWithDrawls(m_AccHeadId, Amount, TransDate) Then GoTo ErrLine
            
    Else
        'Insert into the contras table
        gDbTrans.SqlStmt = "INSERT INTO COntraTrans " & _
                " (ContraID,AccHeadID,AccID,TransType, " & _
                " TransID, Amount,UserID ) VALUES " & _
                "( " & ContraID & "," & m_AccHeadId & "," & _
                p_AccID & "," & transType & "," & TransID + 1 & "," & _
                Amount & "," & gCurrUser.UserID & ")"
        If Not gDbTrans.SQLExecute Then GoTo ErrLine
    End If
End If

gDbTrans.SqlStmt = "UPDATE RDMaster set ClosedDate = #" & TransDate & "#" & _
                    " WHERE Accid = " & p_AccID
If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Enter the details of the tranfer transction
If m_Transfer Then
    If m_ContraClass Is Nothing Then Call cmdTfr_Click
    If m_ContraClass.Status <> Success Then GoTo ErrLine
    'Save the traferred details
    m_ContraClass.TransDate = TransDate
    If Not m_ContraClass.SaveDetails Then GoTo ErrLine
End If


'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(m_AccHeadId, _
            p_AccID, TransDate, TransID, txtPayableAmount) < 1 Then GoTo ErrLine
    Set Cashclass = Nothing
End If

'Commit the transction
gDbTrans.CommitTrans
InTrans = False

RDClose = True

Exit Function

ErrLine:
    If InTrans Then gDbTrans.RollBack
    MsgBox "Unable to close the Recurring Deposit", vbInformation, wis_MESSAGE_TITLE
    
    If Err Then
        MsgBox "error In RDClose", vbInformation, wis_MESSAGE_TITLE
        Err.Clear
    End If

End Function

Private Function TransferToLoanAccount() As Boolean
Dim VoucherNo As String
Dim TransDate As Date
Dim TransID As Long

Dim DepAmount As Currency
Dim MatAmount As Currency
Dim IntAmount As Currency
Dim PayableAmount As Currency
Dim Balance As Currency
'Dim IntBalance As Currency
Dim PayableBalance As Currency
Dim customerID As Long

Dim rst As Recordset

TransDate = GetSysFormatDate(txtDate)
'Now Check THe Date fo Transactio w.r.t LastTransction date
If DateDiff("d", TransDate, GetRDLastTransDate(p_AccID)) > 0 Then
    'Earlier Date
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    Exit Function
End If

gDbTrans.SqlStmt = "Select * from RDTMaster " & _
                    " where AccID = " & p_AccID
If gDbTrans.Fetch(rst, adOpenForwardOnly) Then _
            customerID = FormatField(rst.Fields("CustomerID"))

'get the Balance of the account
gDbTrans.SqlStmt = "Select Balance From RDTrans where" & _
                " AccID = " & p_AccID & " order by transid desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) Then _
                Balance = FormatField(rst.Fields("Balance"))

gDbTrans.SqlStmt = "Select Balance from RDIntPayable where" & _
                    " AccID = " & p_AccID & " order by transid desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) Then _
        PayableBalance = FormatField(rst.Fields("Balance"))

Dim Trans As wisTransactionTypes

PayableAmount = Val(txtIntPayable)
IntAmount = Val(txtDepIntAmount)
DepAmount = Val(txtDepositAmount)
MatAmount = DepAmount + IntAmount + PayableAmount

PayableBalance = PayableBalance - PayableAmount
Balance = Balance - DepAmount

If PayableBalance < 0 Then
    If MsgBox("You are withdrawing more amount than that of deposited from Payble account" & _
        vbCrLf & GetResourceString(541), vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    PayableBalance = 0
End If

'If He closes the account before three months so Taking some Charges
'for Pigmy agent's commission
IntAmount = IntAmount - Val(txtCharges.Text)
' Now Get the INterest Amount deposited in InterestPayble Account

'Get the Next TransId
TransID = GetRDMaxTransID(p_AccID) + 1

Dim bankClass As clsBankAcc
Dim InTrans As Boolean

gDbTrans.BeginTrans

Set bankClass = New clsBankAcc

Dim UserID As Long
UserID = gCurrUser.UserID

'Close the account by giving deposit
Trans = wContraWithdraw
gDbTrans.SqlStmt = "Insert into RDTrans (AccID,TransID, TransType, " & _
        " TransDate, Amount, Balance,VoucherNo ,Particulars,UserID ) values ( " & _
        p_AccID & "," & _
        TransID & "," & _
        Trans & "," & _
        "#" & TransDate & "#," & _
        DepAmount & "," & Balance & "," & _
        AddQuotes(VoucherNo, True) & "," & _
        " 'By Deposit Repayment' ," & _
        UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

If IntAmount Then
    Trans = IIf(IntAmount > 0, wContraWithdraw, wContraDeposit)
    gDbTrans.SqlStmt = "Insert into RDIntTrans (AccID, TransID, TransType, " & _
            " TransDate, Amount,Balance, VoucherNo,Particulars,UserID ) values ( " & _
            p_AccID & "," & _
            TransID & "," & _
            Trans & "," & _
            "#" & TransDate & "#," & _
            Abs(IntAmount) & ", 0," & _
            AddQuotes(VoucherNo, True) & "," & _
            "'By interest' ," & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    If Trans = wContraWithdraw Then _
        If Not bankClass.UpdateContraTrans(m_RdIntHeadId, m_AccHeadId, IntAmount, _
                TransDate) Then GoTo ErrLine
    If Trans = wContraDeposit Then
        If Not bankClass.UpdateContraTrans(IIf(PayableAmount, m_RdPayableId, m_AccHeadId), _
             m_AccHeadId, Abs(IntAmount), TransDate) Then GoTo ErrLine
        PayableAmount = IIf(PayableAmount, PayableAmount + IntAmount, 0)
     End If
End If

If PayableAmount Then
    Trans = wContraWithdraw
    
    gDbTrans.SqlStmt = "Insert into RDIntPayable (AccID, TransID, TransType, " & _
            " TransDate, Amount,Balance, VoucherNo,Particulars,UserID ) values ( " & _
            p_AccID & "," & _
            TransID & "," & _
            Trans & "," & _
            "#" & TransDate & "#," & _
            PayableAmount & "," & _
            PayableBalance & "," & _
            AddQuotes(VoucherNo, True) & "," & _
            "'By interest' ," & _
            UserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

    If Not bankClass.UpdateContraTrans(m_RdPayableId, m_AccHeadId, _
                PayableAmount, TransDate) Then GoTo ErrLine
    
End If

'Update the first transaction with close = 1
gDbTrans.SqlStmt = "UPdate RDMaster Set ClosedDate = #" & TransDate & "# " & _
        " where AccID = " & p_AccID

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'/////Contra And Suspense account goes here
'if he is repaying the loan amount or trnsferring to other account
'then the transction will be contra
Dim ContraID As Long
'Get the Contra ID
 ContraID = GetMaxContraTransID + 1

'put withdrawal transction details int to contra Table
 gDbTrans.SqlStmt = "Insert into ContraTrans " & _
        "(ContraId,AccId,AccHeadID," _
         & " TransType,TransId,Amount,VoucherNo,UserId)" & _
        " Values (" _
         & ContraID & "," & _
         p_AccID & "," & _
         m_AccHeadId & "," & _
         Trans & ", " & TransID & "," & _
         DepAmount & "," & _
         AddQuotes(VoucherNo, True) & _
         "," & gUserID & " )"
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

'if he is transfering the Amount to the loan account,
'transfer it to loan account
Dim LoanAmount As Currency
Dim LoanIntAmount As Currency

MatAmount = txtDepositAmount + Me.txtTotalInt
LoanAmount = Val(txtLoanAmount)
LoanIntAmount = txtLoanInterest

Dim DepLOanClass As clsDepLoan
Set DepLOanClass = New clsDepLoan

'if loan amount is more than the matured amount
'then first take the interest then remaining amount as principal
'loanclass
MatAmount = MatAmount - LoanIntAmount
LoanAmount = IIf(LoanAmount < MatAmount, LoanAmount, MatAmount)
MatAmount = MatAmount - LoanAmount
If DepLOanClass.DepositAmount(CInt(rst(0)), LoanAmount, _
        LoanIntAmount, "RD Close", TransDate, VoucherNo) = 0 Then GoTo ErrLine

    Dim LoanIntHeadID As Long
    Dim LoanHeadID As Long
    Dim headName As String
    headName = GetResourceString(424, 58)
    LoanHeadID = GetIndexHeadID(headName)
    headName = headName & " " & GetResourceString(483)
    LoanIntHeadID = GetIndexHeadID(headName)

If LoanIntAmount Then _
    If Not bankClass.UpdateContraTrans(m_AccHeadId, LoanIntHeadID, _
        LoanIntAmount, TransDate) Then GoTo ErrLine
If Not bankClass.UpdateContraTrans(m_AccHeadId, LoanHeadID, _
    LoanAmount, TransDate) Then GoTo ErrLine

'if matured amount is more than the loan amount
'then transfer tht remianing amount to the suspence account
If MatAmount > 0 Then
    Dim SuspHeadID As Long
    SuspHeadID = GetIndexHeadID(GetResourceString(365))
    Debug.Assert MatAmount = 0
    SuspHeadID = bankClass.GetHeadIDCreated(GetResourceString(365), LoadResString(365), _
                                    parSuspAcc, 0, wis_SuspAcc)
    
    Dim SuspClass As New clsSuspAcc
    If SuspClass.DepositAmount(m_AccHeadId, p_AccID, customerID, "", _
                        TransDate, MatAmount, TransID, VoucherNo) < 1 Then GoTo ErrLine
    
    
    If Not bankClass.UpdateContraTrans(m_AccHeadId, SuspHeadID, _
        MatAmount, TransDate) Then GoTo ErrLine
    
End If

gDbTrans.CommitTrans
Set bankClass = Nothing

TransferToLoanAccount = True

ErrLine:
    If InTrans Then gDbTrans.RollBack
    'MsgBox "Unable to perform transaction !", vbCritical, gAppName & " - Critical Error"
    MsgBox GetResourceString(535), vbCritical, gAppName & " - Critical Error"
        
    Set bankClass = Nothing
    Exit Function


End Function

Private Sub UpdateDetails()

'Check for valid date
Dim IntRate As Single
Dim Days As Integer
Dim DepDate As Date
Dim DepAmt As Currency
Dim MatDate As Date
Dim rstTrans As ADODB.Recordset
Dim rstMaster As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim TransDate As Date

Dim ClsBank  As clsBankAcc
If m_AccHeadId = 0 Then
    Dim headName As String
    Dim headNameEnglish As String
    Set ClsBank = New clsBankAcc
    gDbTrans.BeginTrans
    headName = GetResourceString(424)
    headNameEnglish = LoadResString(424)
    m_AccHeadId = ClsBank.GetHeadIDCreated(headName, headNameEnglish, parMemberDeposit, 0, wis_RDAcc)
    
    headName = GetResourceString(424, 487)
    headNameEnglish = LoadResourceStringS(424, 487)
    m_RdIntHeadId = ClsBank.GetHeadIDCreated(headName, headNameEnglish, parMemDepIntPaid, 0, wis_RDAcc)
    
    headName = GetResourceString(424, 375, 47)
    headName = LoadResourceStringS(424, 375, 47)
    m_RdPayableId = ClsBank.GetHeadIDCreated(headName, headNameEnglish, parDepositIntProv, 0, wis_RDAcc)
    gDbTrans.CommitTrans
    Set ClsBank = Nothing
End If

If DateValidate(txtDate.Text, "/", True) Then
    TransDate = GetSysFormatDate(txtDate.Text)
Else
    TransDate = gStrDate
End If

'Get the deposited amount
gDbTrans.SqlStmt = "Select Balance from RDTrans where " & _
        " AccID = " & p_AccID & " And TRansID = " & _
        "(Select MAx(TransID) from RDTrans where AccID = " & p_AccID & ")"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        DepAmt = Val(FormatField(rst("Balance")))
    End If

    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & p_AccID
    If gDbTrans.Fetch(rstMaster, adOpenForwardOnly) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
    MatDate = rstMaster("MaturityDate")
    DepDate = rstMaster("CreateDate")
    IntRate = rstMaster("RateOfinterest")

'Now Get the Interest Payble deposited
gDbTrans.SqlStmt = "SELECT Balance From RDIntPayable WHERE " & _
    " AccID = " & p_AccID & " AND TransID = (SELECT MAx(TransID) From " & _
        " RDIntPayable WHERE AccID = " & p_AccID & ")"
If gDbTrans.Fetch(rst, adOpenForwardOnly) Then
    txtIntPayable = FormatField(rst("BALANCE"))
    txtIntPayable.Tag = FormatField(rst("BALANCE"))
End If

'Now Get the Interest Till Date
txtDepIntAmount = ComputeRDDepositInterestAmount(p_AccID, txtDate.Text) \ 1
'Now Get the Toal Interest Till Date
txtTotalInt = txtIntPayable + txtDepIntAmount

'Get total loan Details
Dim LnAmt As Currency
Dim LoanID As Long
Dim Duration As Integer
LoanID = FormatField(rstMaster("LoanID"))
'Get sum of withdrawals as loans drawn
'Extract the rate of interest from Setup values
    
    gDbTrans.SqlStmt = "SELECT * From DepositLoanMaster WHERE " & _
            " LoanID = " & LoanID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) Then
        'txtLoanRate = FormatField(Rst("RateOFInterest"))
        txtLoanRate = FormatField(rst("InterestRate"))
        gDbTrans.SqlStmt = "SELECT * FROM DepositLoanTrans WHERE " & _
            " LOanID = " & LoanID & " AND TransID =(Select Max(transID) From " & _
            " DepositLoanTrans WHERE LOanId = " & LoanID & ")"
        If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then GoTo NextLine
    
        Duration = DateDiff("D", rst("TransDate"), TransDate)
        txtLoanAmount = FormatField(rst("Balance"))
        fraLoanDet.Enabled = Val(txtLoanAmount)
        txtLoanInterest = FormatCurrency(Val(txtLoanAmount) * Duration / 365 * Val(txtLoanRate) / 100)
        If Val(txtLoanAmount) Then ChkDeductLoan.Value = vbChecked
    End If
    
NextLine:
'CALUCLATE INTEREST ON DEPOSIT
    gDbTrans.SqlStmt = "SELECT * FROM RDTrans WHERE " & _
            " AccID = " & p_AccID & " AND TransID =(Select Max(transID) From " & _
                " RDTrans WHERE AccId = " & p_AccID & ")"
    If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
'Extract the rate of interest from Setup values
    txtIntRate.Text = "0.00"
    If IntRate <= 0 Then _
        IntRate = M_setUp.ReadSetupValue("RDAcc", "Interest on RDDeposit", "7")
        
    txtIntRate = Format(IntRate, "00.00")
    txtDepositAmount = DepAmt

' Now check for the valid date
    
    'Calculate the number of days
    Days = DateDiff("D", GetSysFormatDate(txtDate.Text), MatDate)
    If Days > 0 Then  'Account being closed prematurely
        Days = DateDiff("D", DepDate, GetSysFormatDate(txtDate.Text))
    Else
        Days = DateDiff("D", DepDate, MatDate)
    End If
Dim IntAmount As Currency
IntAmount = ComputeRDDepositInterestAmount(p_AccID, txtDate.Text)

txtDepIntAmount = IntAmount \ 1
                    
End Sub

Private Sub cmdAccept_Click()
    
'Perform the transaction with closed flag and send the guy home
'Check date
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date not in dd/mm/yyyy format ", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
    
    
If ChkDeductLoan.Value = vbChecked Then
    If Not TransferToLoanAccount Then Exit Sub
Else
    If Not RDClose Then Exit Sub
End If

Unload Me
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


'
Private Sub cmdTfr_Click()
'Perform the transaction with closed flag and send the guy home
'Check date
Dim TransDate As Date
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date not in dd/mm/yyyy format ", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
TransDate = GetSysFormatDate(txtDate)

Dim AccNum As String
Dim rst As Recordset
Dim PayableAmount As Currency
Dim IntAmount As Currency
Dim Amount As Currency

Dim RDHeadID As Long
Dim IntHeadID As Long
Dim PayableHeadID As Long

Dim BankCls As clsBankAcc

gDbTrans.SqlStmt = "SELECT AccNum From RDMaster Where AccId = " & p_AccID
If gDbTrans.Fetch(rst, adOpenDynamic) Then AccNum = FormatField(rst(0))

PayableAmount = Val(txtIntPayable)
IntAmount = Val(txtDepIntAmount)
Amount = Val(txtDepositAmount)

Set m_ContraClass = New clsContra

With m_ContraClass
Set BankCls = New clsBankAcc
    Call .Transfer(TransDate, "12", m_AccHeadId, AccNum, Amount)
    If PayableAmount Then _
        Call .Transfer(TransDate, "12", m_RdPayableId, AccNum, PayableAmount)
    
    If IntAmount Then _
        Call .Transfer(TransDate, "12", m_RdIntHeadId, AccNum, IntAmount)
    'Below Commented Line moved above 'If PayableAmount Then' line
    'Call .Transfer(TransDate, "12", m_AccHeadId, AccNum, Amount)
    
    .Show
    If .Status <> Success Then Exit Sub
End With

m_Transfer = True

If RDClose Then Unload Me

End Sub
Private Sub Form_Load()
'Center the form
 Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'Set kannada fonts
Call SetKannadaCaption
         
'Todays date
    txtDate.Text = gStrDate


Call UpdateDetails
If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmRDClose = Nothing
End Sub




Private Sub txtCharges_Change()
txtPayableAmount = FormatCurrency(Val(txtDepositAmount) + Val(txtTotalInt) _
    - Val(txtLoanAmount) - Val(txtLoanInterest) - Val(txtCharges) - Val(txtTax))

End Sub

Private Sub txtDate_LostFocus()
Call UpdateDetails
End Sub


Private Sub txtDepositAmount_Change()
'txtPayableAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + Val(txtDepositInterest.Text) - Val(txtLoanAmount.Text) - Val(txtLoanInterest.Text) - Val(txtCharges.Text) - Val(txtTax.Text))
End Sub


Private Sub txtIntPayable_Change()
    txtTotalInt = FormatCurrency(Val(txtIntPayable) + Val(txtDepIntAmount))
End Sub

Private Sub txtLoanAmount_Change()
    'txtPayableAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + Val(txtDepositInterest.Text) - Val(txtLoanAmount.Text) - Val(txtLoanInterest.Text) - Val(txtCharges.Text) - Val(txtTax.Text))
End Sub


Private Sub txtLoanInterest_Change()
txtPayableAmount = FormatCurrency(Val(txtDepositAmount) + Val(txtTotalInt) _
    - Val(txtLoanAmount) - Val(txtLoanInterest) - Val(txtCharges) - Val(txtTax))

End Sub


Private Sub txtLoanRate_LostFocus()
Call UpdateDetails

If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub

Dim LoanDate As Date
Dim Days As Integer
Dim LoanAmount As Currency
Dim IntRate As Single
Dim rstMaster As Recordset
Dim rst As ADODB.Recordset
Dim MatDate As Date
Dim AsOnDate As Date
Dim LoanID As Long

IntRate = Val(Me.txtLoanRate.Text)
LoanAmount = Val(txtLoanAmount)
AsOnDate = GetSysFormatDate(txtDate)

    gDbTrans.SqlStmt = "Select * from RDMaster where AccID = " & p_AccID
                        
    If gDbTrans.Fetch(rstMaster, adOpenForwardOnly) <= 0 Then
        'MsgBox "No deposits listed !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(570), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        MatDate = rstMaster("MaturityDate")
        LoanID = FormatField(rstMaster("LoanID"))
    End If
    
    If LoanID = 0 Then Exit Sub

'Get date of last transaction
    gDbTrans.SqlStmt = "Select TOP 1 TransDate,Balance from DepositLoanTrans where " & _
        " LoanID = " & LoanID & " order by TransID desc"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) Then
        On Error Resume Next
        LoanDate = rst("TransDate")
        'See if the account has matured
        If DateDiff("D", MatDate, AsOnDate) > 0 Then 'Account has matured
                    'Consider loan till 'Matured date only
            Days = DateDiff("D", LoanDate, MatDate)
        Else
            Days = DateDiff("D", LoanDate, AsOnDate)
        End If
    End If

    'computetefdinterest
    txtLoanInterest.Text = FormatCurrency(Val(txtLoanAmount) * Days / 365 * IntRate)
    
    Set rst = Nothing

    Err.Clear
    
End Sub


Private Sub txtTax_Change()
'txtPayableAmount.Text = FormatCurrency(Val(txtDepositAmount.Text) + Val(txtDepositInterest.Text) - Val(txtLoanAmount) - Val(txtLoanInterest.Text) - Val(txtCharges.Text) - Val(txtTax.Text))
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

fraDepDetail = GetResourceString(43, 295)
fraLoanDet = GetResourceString(58, 295)
fraCharges = GetResourceString(237, 273)

'Set the Kannada caption to the Command buttons
cmdAccept.Caption = GetResourceString(11)
cmdCancel.Caption = GetResourceString(2)
cmdTfr.Caption = GetResourceString(307)

lblDepDetail = GetResourceString(43, 295)
lblDate = GetResourceString(37)
Me.lblDepAmount = GetResourceString(43, 42)
lblDepInt = GetResourceString(186)
lblIntPayable = GetResourceString(450) 'Interest Payable
lblInterest = GetResourceString(233) 'INterest on Deposit
lblTotalInt = GetResourceString(52, 233) 'Total INterest on Deposit

'lblLoanDet = GetResourceString(58,295)
lblLoanAmount = GetResourceString(80, 91)
lblLoanInt = GetResourceString(186)
lblIntOnLoan = GetResourceString(80, 47)

'lblCharges = loadresstring(glangOffset +237) & " " & loadresstring(glangOffset + 273)
lblPreClose = GetResourceString(238)
lblOthers = GetResourceString(237, 273)
lblNetAmount = GetResourceString(240)

End Sub

Private Sub txtTotalInt_Change()
txtPayableAmount = FormatCurrency(Val(txtDepositAmount) + Val(txtTotalInt) _
    - Val(txtLoanAmount) - Val(txtLoanInterest) - Val(txtCharges) - Val(txtTax))
End Sub


