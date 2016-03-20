VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmFDInterest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interest payment"
   ClientHeight    =   5250
   ClientLeft      =   3930
   ClientTop       =   1875
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   3405
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   4515
      Begin VB.CheckBox chkAdd 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Interest to deposit"
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   2880
         Width           =   3555
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2790
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtInterestRate 
         Height          =   315
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   643
         Width           =   1215
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   4050
         TabIndex        =   0
         Top             =   180
         Width           =   315
      End
      Begin WIS_Currency_Text_Box.CurrText txtPayable 
         Height          =   345
         Left            =   2790
         TabIndex        =   7
         Top             =   1509
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtInterestAmount 
         Height          =   345
         Left            =   2790
         TabIndex        =   10
         Top             =   1972
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotalInterest 
         Height          =   345
         Left            =   2790
         TabIndex        =   13
         Top             =   2435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblDate 
         Caption         =   "Transaction Date"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   210
         Width           =   1785
      End
      Begin VB.Label txtDepositAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2790
         TabIndex        =   5
         Top             =   1076
         Width           =   1215
      End
      Begin VB.Label lblTotalInterest 
         Caption         =   "Total Interest : "
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   2435
         Width           =   2025
      End
      Begin VB.Label lblPayable 
         Caption         =   "Payable :"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1545
         Width           =   2025
      End
      Begin VB.Label lblDepositAmount 
         Caption         =   "Depositt amount:"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1100
         Width           =   1995
      End
      Begin VB.Label lblInterestAmount 
         Caption         =   "Interest accrued : "
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1990
         Width           =   1965
      End
      Begin VB.Label lblInterestRate 
         Caption         =   "Rate of interest:"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   655
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   400
      Left            =   1560
      TabIndex        =   11
      Top             =   4650
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   2850
      TabIndex        =   14
      Top             =   4650
      Width           =   1215
   End
   Begin WIS_Currency_Text_Box.CurrText txtPaidAmount 
      Height          =   345
      Left            =   2880
      TabIndex        =   17
      Top             =   3540
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.Label txtBalanceAmount 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2880
      TabIndex        =   18
      Top             =   4110
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   4080
      Y1              =   4545
      Y2              =   4530
   End
   Begin VB.Label lblPaidAmount 
      Caption         =   "Repaid amount : "
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   3570
      Width           =   2325
   End
   Begin VB.Label lblBalanceAmount 
      Caption         =   "deposit balance after this trans"
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   4140
      Width           =   2325
   End
End
Attribute VB_Name = "frmFDInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_AccID As Long
Private m_DepositType As Integer


Public Property Let AccountId(NewValue As Long)
    m_AccID = NewValue
End Property

Private Function InterestTransaction() As Boolean

Dim transType As wisTransactionTypes
Dim ContraTransType As wisTransactionTypes
Dim TransID As Long
Dim TransDate As Date

Dim Amount As Currency
Dim IntAmount As Currency
'Dim IntBalance  As Currency
Dim PayableAmount As Currency
Dim PayableBalance  As Currency
Dim rst As ADODB.Recordset
Dim InTrans As Boolean

TransDate = GetSysFormatDate(txtDate)

'Get new transID
TransID = GetFDMaxTransID(m_AccID)
If TransID = 0 Then Exit Function

'Now Check the LAst date of transaction
If DateDiff("d", TransDate, GetFDMaxTransDate(m_AccID)) > 0 Then
    'date of transction
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Details of the Interest Payble Account
gDbTrans.SqlStmt = "Select TOP 1 Balance " & _
            " from FDIntPayable where AccID = " & m_AccID & _
            " Order by TransID desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
                PayableBalance = FormatField(rst("Balance"))

Set rst = Nothing

'Now  GET THE AMOUNT which has to be paid from Loss Account (of this year)
'AND GET THE AMOUNT which has to be interest Payable account
PayableAmount = txtPayable
IntAmount = txtInterestAmount

Amount = Val(txtBalanceAmount) - Val(txtDepositAmount)

''Now Get the interest Amount deposited in InterestPayble Account
PayableBalance = IIf(PayableBalance > 0, PayableBalance, 0)

If Amount < 0 Or PayableAmount < 0 Then
    'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If Amount > 0 And Amount <> txtTotalInterest Then
    'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Start database operations
Dim bankClass As clsBankAcc
Dim IntHeadID As Long
Dim PayableHeadID As Long
Dim AccHeadID As Long


gDbTrans.BeginTrans
InTrans = True
Set bankClass = New clsBankAcc

'Now Deduct the amount from interest payble account
If IntAmount < 0 Then PayableAmount = txtPaidAmount

'Now Get all the HeadIDs
Dim headName As String
Dim headNameEnglish As String
Dim engHeadName As String

headName = GetDepositTypeTextEnglish(m_DepositType, headNameEnglish)
AccHeadID = GetIndexHeadID(GetDepositTypeText(m_DepositType))
If Len(headNameEnglish) > 0 Then engHeadName = headNameEnglish & " " & LoadResString(487)
IntHeadID = bankClass.GetHeadIDCreated(headName & " " & GetResourceString(487), engHeadName, _
    parMemDepIntPaid, wis_Deposits + m_DepositType)
If PayableAmount Then
    headName = headName & " " & GetResourceString(375, 47)
    If Len(headNameEnglish) > 0 Then engHeadName = headNameEnglish & " " & LoadResourceStringS(375, 47)
    PayableHeadID = bankClass.GetHeadIDCreated(headName, engHeadName, parDepositIntProv, wis_Deposits + m_DepositType)
End If

TransID = TransID + 1
'Insert Interest first

If PayableAmount > 0 Then
    'withdraw the amount from IntPayble
    PayableBalance = PayableBalance - PayableAmount
    If PayableBalance < 0 Then
        If MsgBox(GetResourceString(577) & vbCrLf & _
                    GetResourceString(541), vbYesNo + vbQuestion + _
                    vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then GoTo Exit_Line
        
        PayableBalance = 0
    End If
    transType = IIf(chkAdd = vbChecked, wContraWithdraw, wWithdraw)
    gDbTrans.SqlStmt = "Insert into FDIntPayable (AccID,TransID,  " & _
            " TransType, TransDate, Amount, Balance, " & _
            " Particulars,UserID) values (" & _
            m_AccID & "," & _
            TransID & "," & _
            transType & "," & _
            "#" & TransDate & "#," & _
            PayableAmount & "," & _
            PayableBalance & ", " & _
            "'Interest paid' " & _
            "," & gUserID & " ) "
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If chkAdd.Value = vbChecked Then
        If Not bankClass.UpdateContraTrans(PayableHeadID, AccHeadID, _
            PayableAmount, TransDate) Then GoTo Exit_Line
    Else
        If Not bankClass.UndoCashWithdrawls(PayableHeadID, _
            PayableAmount, TransDate) Then GoTo Exit_Line
    End If
End If

If IntAmount > 0 Then
    transType = IIf(chkAdd = vbChecked, wContraWithdraw, wWithdraw)
    gDbTrans.SqlStmt = "Insert into FDIntTrans (AccID, " & _
            " TransID, TransType, TransDate, Amount, Balance," & _
            " Particulars,UserID) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            transType & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            IntAmount & ", " & _
            "'Interest Paid'" & _
            "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
    If chkAdd.Value = vbChecked Then
        If Not bankClass.UpdateContraTrans(IntHeadID, AccHeadID, _
            IntAmount, TransDate) Then GoTo Exit_Line
    Else
        If Not bankClass.UpdateCashWithDrawls(IntHeadID, _
            IntAmount, TransDate) Then GoTo Exit_Line
    End If

End If

'If This transaction has affected balacne then
'Make one more transaction of deposit
'i.e Interest added to the FD Balance
If Amount > 0 Then
    transType = wContraDeposit
    gDbTrans.SqlStmt = "Insert into FDTrans (AccID, " & _
            " TransID, TransType, TransDate, Amount, Balance," & _
            " Particulars,UserID) values ( " & _
            m_AccID & "," & _
            TransID & "," & _
            transType & "," & _
            "#" & TransDate & "#," & _
            IntAmount & "," & _
            Val(txtBalanceAmount) & ", " & _
            "'Interest Added'" & _
            "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    
End If

'Now Update the last int paid date
gDbTrans.SqlStmt = "Update FdMaster Set LastIntDate = #" & TransDate & "#" & _
    " WHERE AccID = " & m_AccID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

'If transaction is cash withdraw & there is casier window
'then transfer the While Amount cashier window
If transType = wWithdraw And gCashier Then
    Dim Cashclass As clsCash
    Set Cashclass = New clsCash
    If Cashclass.TransferToCashier(AccHeadID, _
            m_AccID, TransDate, TransID, IntAmount) < 1 Then GoTo Exit_Line
    
    Set Cashclass = Nothing
End If

gDbTrans.CommitTrans
InTrans = False


InterestTransaction = True

Exit_Line:
    
    Set bankClass = Nothing
    'Resume
    If InTrans Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
    End If
End Function
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblDate.Caption = GetResourceString(37)   '
lblInterestRate.Caption = GetResourceString(186)
lblDepositAmount.Caption = GetResourceString(43, 42)
lblPaidAmount.Caption = GetResourceString(205, 483)
lblBalanceAmount = GetResourceString(43, 42)
cmdAccept.Caption = GetResourceString(1) 'OK
cmdCancel.Caption = GetResourceString(2)  'Cancel
lblTotalInterest.Caption = GetResourceString(52, 47) & " :"
lblInterestAmount.Caption = GetResourceString(47) & " :"
lblPayable.Caption = GetResourceString(450) & " :"

End Sub

Private Sub UpdateDetails()
On Error GoTo ErrLine

Dim transType As wisTransactionTypes
Dim LoanAmt As Currency
Dim RepayAmt As Currency
Dim Balance As Currency
Dim Days As Long
Dim InterestRate As Double
Dim Payable As Currency
Dim LastDate As Date
Dim TransDate As Date
Dim rst As ADODB.Recordset

If DateValidate(Trim(txtDate.Text), "/", True) Then
    TransDate = GetSysFormatDate(txtDate)
Else
    TransDate = gStrDate
End If

'Get Rate of Interest For This deposits
gDbTrans.SqlStmt = "Select RateOfInterest, CreateDate,LastIntDate, " & _
    "EffectiveDate,DepositType From FDMaster Where AccID = " & m_AccID
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

m_DepositType = FormatField(rst("DepositType"))
InterestRate = FormatField(rst("RateOfInterest"))
txtInterestRate.Text = Format(InterestRate, "#0.00")

If FormatField(rst("LastIntDate")) <> "" Then
    LastDate = rst("LastIntDate")
Else
    LastDate = rst("EffectiveDate")
End If

gDbTrans.SqlStmt = "Select Cumulative From DepositName " & _
        "Where DepositId =  " & m_DepositType
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
If FormatField(rst("Cumulative")) Then
    chkAdd = vbChecked
    txtPayable.Locked = True
End If

Payable = 0
gDbTrans.SqlStmt = "Select Top 1 Balance from FDIntPayable where " & _
    " AccID = " & m_AccID & " Order By TransID Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then Payable = FormatField(rst(0))
txtPayable = Payable

'Get the last date of transaction
gDbTrans.SqlStmt = "Select TOP 1 TransDate, Balance, TransType " & _
        " From FDTrans Where AccID = " & m_AccID & _
        " AND Amount <> 0 Order by TransID desc "
Call gDbTrans.Fetch(rst, adOpenForwardOnly)

LastDate = rst("TransDate")
Balance = CCur(FormatField(rst("Balance")))
transType = rst("TransType")
    
gDbTrans.SqlStmt = "Select TOP 1 TransDate, Balance, TransType " & _
            " From FDIntTrans Where AccID = " & m_AccID & _
            " Order by TransID desc "
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
    LastDate = IIf(LastDate > rst("TransDate"), LastDate, rst("TransDate"))
    
'Calculate the interest till date given
Dim IntAmount As Currency
Dim FdClass As clsFDAcc
Set FdClass = New clsFDAcc
Days = DateDiff("d", LastDate, TransDate)
'IntAmount = ComputeFDInterest(Balance, LastDate, TransDate, _
                    m_DepositType, CSng(InterestRate))

IntAmount = FdClass.InterestAmount(m_AccID, LastDate, TransDate)
IntAmount = IntAmount - (IntAmount * DateDiff("D", LastDate, TransDate) / 365 * InterestRate / 100)
Set FdClass = Nothing
txtDepositAmount = FormatCurrency(Balance)
txtBalanceAmount = FormatCurrency(Balance)
txtInterestAmount = (IntAmount \ 1)
txtPaidAmount = txtTotalInterest


If txtPayable.Locked Then txtPayable = 0

Exit Sub

ErrLine:

'MsgBox "No loans have been issued on this deposit !", vbExclamation, gAppName & " - Error"
'MsgBox GetResourceString(582), vbExclamation, gAppName & " - Error"
'Resume
End Sub

Private Sub chkAdd_Click()
Call txtTotalInterest_Change
End Sub

Private Sub cmdAccept_Click()

'Check the date
If Not DateValidate(txtDate.Text, "/", True) Then
    'MsgBox "Date not specified in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
Dim TransDate As String
TransDate = GetSysFormatDate(txtDate)
    
' Check whether The int paid amount and Deposited amount is same or not
If Not (Val(txtBalanceAmount) <> Val(txtDepositAmount) + txtTotalInterest Xor Val(txtBalanceAmount) <> Val(txtDepositAmount)) Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    'ActivateTextBox txtPaidAmount
    txtPaidAmount.SetFocus
    Exit Sub
End If

If txtPaidAmount < 0 Then
    'MsgBox "Amount repaid is greater that total amount to be paid !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(585), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

If Val(txtBalanceAmount) < 0 Then
    'MsgBox "Amount repaid is greater that total amount to be paid !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(585), vbExclamation, gAppName & " - Error"
    'ActivateTextBox txtBalanceAmount
    Exit Sub
End If

'Get the Master detaisl
Dim rstMaster As ADODB.Recordset
gDbTrans.SqlStmt = "SELECT * FROM FDMAster Where AccId = " & m_AccID
Call gDbTrans.Fetch(rstMaster, adOpenForwardOnly)

Dim rst As ADODB.Recordset
'Check last date of transaction
gDbTrans.SqlStmt = "Select TOP 1 TransDate from FDIntTrans where " & _
    " AccID = " & m_AccID & _
    " And (TransType <> " & wWithdraw & _
        " OR TransType <> " & wContraWithdraw & ")" & _
    " order by TransID DESC"
Dim Compdate As Date

Compdate = rstMaster("EffectiveDate")
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    Compdate = rst("TransDate")
Else
    gDbTrans.SqlStmt = "Select TOP 1 TransDate from FDTrans where " & _
        " AccID = " & m_AccID & _
        " And (TransType <> " & wDeposit & " OR TransType <> " & wContraDeposit & ")" & _
        " order by TransID DESC"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        Compdate = rst("TransDate")
End If

If DateDiff("d", Compdate, TransDate) < 0 Then
    'MsgBox "Date specified should be later than the date of last transaction on this account !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtDate
    Exit Sub
End If
    
If DateDiff("d", rstMaster("MaturityDate"), TransDate) > 0 Then
    If MsgBox(GetResourceString(578, 541), _
        vbYesNo + vbQuestion, gAppName & " - Error") = vbNo Then _
                            ActivateTextBox txtDate: Exit Sub
End If
    
Dim transType As wisTransactionTypes
Dim TransID As Long
Dim Balance As Currency
Dim Loan As Boolean
 
'Get new transID
Loan = FormatField(rstMaster("LoanID"))
gDbTrans.SqlStmt = "Select TOP 1 TransID, Balance " & _
    " from FDTrans where " & _
    " AccID = " & m_AccID & " order by TransID desc"
    
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
TransID = FormatField(rst("TransID")) + 1
Balance = CCur(FormatField(rst("Balance")))

Set rst = Nothing

'Check whether The Deposit balance is correct or not
If Val(txtBalanceAmount) < Balance Then
    'MsgBox "Invalid amount specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
    Exit Sub
End If

If Val(txtBalanceAmount) > Balance And chkAdd.Value <> vbChecked Then
    'if MsgBox ("You have entered loan deposit balance more than previous balance !"
   '     " do You want to continue ?", vbExclamation+vbyesno, gAppName & " - Error") = vbno then
    If MsgBox(GetResourceString(584, 541), _
          vbYesNo + vbInformation + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
         Exit Sub
   End If
End If

If Not InterestTransaction Then
    txtDate.Text = gStrDate
    txtPaidAmount.Text = "0.00"
    Call UpdateDetails
    Exit Sub
End If

'MsgBox "Repayment made successfully !", vbInformation, gAppName & " - Error"
MsgBox GetResourceString(586), vbInformation, gAppName & " - Error"

Unload Me
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + cmdDate.Top
    .selDate = txtDate.Text
    .Show vbModal
    txtDate.Text = .selDate
End With
End Sub


Private Sub Form_Load()
    
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'Set kannada fonts
Call SetKannadaCaption
'Load values to the text box
txtDate.Text = gStrDate
    
Call UpdateDetails

If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFDInterest = Nothing
   
End Sub


Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub

If Not gOnLine Then Call UpdateDetails

End Sub

Private Sub txtInterestAmount_Change()
    On Error Resume Next
    txtPaidAmount = IIf(chkAdd = vbChecked, 0, txtInterestAmount)
    txtTotalInterest = txtInterestAmount + txtPayable
End Sub

Private Sub txtPayable_Change()
    txtTotalInterest = txtInterestAmount + txtPayable
End Sub




Private Sub txtTotalInterest_Change()
    txtBalanceAmount = FormatCurrency(Val(txtDepositAmount) + IIf(chkAdd = vbChecked, txtTotalInterest, 0))
End Sub

