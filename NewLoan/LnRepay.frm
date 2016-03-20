VERSION 5.00
Begin VB.Form frmLoanRepay 
   Caption         =   "Loan repayment"
   ClientHeight    =   5160
   ClientLeft      =   2700
   ClientTop       =   2535
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   5205
   Begin VB.Frame fraRepay 
      Caption         =   "Repayment details"
      Height          =   3405
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   5115
      Begin VB.TextBox txtRemark 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   2970
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3600
         TabIndex        =   29
         Top             =   2550
         Width           =   1350
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   90
         TabIndex        =   28
         Top             =   2460
         Width           =   4935
      End
      Begin VB.TextBox txtNextBalance 
         Height          =   345
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1710
         Width           =   1500
      End
      Begin VB.TextBox txtNextPenalInt 
         Height          =   315
         Left            =   3420
         TabIndex        =   12
         Top             =   1350
         Width           =   1500
      End
      Begin VB.TextBox txtNextRegInt 
         Height          =   315
         Left            =   3420
         TabIndex        =   9
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox txtMIsc 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   2070
         Width           =   1470
      End
      Begin VB.TextBox txtTotInst 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.TextBox txtPrincipal 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   1710
         Width           =   1470
      End
      Begin VB.TextBox txtPenalInt 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1350
         Width           =   1470
      End
      Begin VB.TextBox txtRegInt 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   990
         Width           =   1470
      End
      Begin VB.TextBox txtPayIntBalance 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   630
         Width           =   1470
      End
      Begin VB.TextBox txtNextIntBalance 
         Height          =   315
         Left            =   3420
         TabIndex        =   6
         Top             =   630
         Width           =   1500
      End
      Begin VB.CommandButton cmdRepayDate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
         Height          =   285
         Left            =   2910
         TabIndex        =   3
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtTransDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label lblRemarks 
         Caption         =   "&Remarks"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   3030
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   3300
         X2              =   3300
         Y1              =   315
         Y2              =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "Values after Transaction"
         ForeColor       =   &H00FF00FF&
         Height          =   525
         Left            =   3360
         TabIndex        =   20
         Top             =   150
         Width           =   1635
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   "&Miscellaneous"
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   1440
      End
      Begin VB.Label lblTotInst 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2610
         Width           =   1440
      End
      Begin VB.Label lblPrincipal 
         AutoSize        =   -1  'True
         Caption         =   "&Principal Amount"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Width           =   1440
      End
      Begin VB.Label lblPenal 
         AutoSize        =   -1  'True
         Caption         =   "&Penal Interest :"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Regular Interest :"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1035
         Width           =   1440
      End
      Begin VB.Label lblPayIntBalance 
         Caption         =   "Interest &Balance"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1440
      End
      Begin VB.Label lblRepayDate 
         AutoSize        =   -1  'True
         Caption         =   "&Date of repayment : "
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame fraCust 
      Caption         =   "Customer"
      Height          =   1365
      Left            =   60
      TabIndex        =   23
      Top             =   30
      Width           =   5115
      Begin VB.TextBox txtLoanAccNo 
         Height          =   315
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtCustName 
         Height          =   315
         Left            =   1530
         TabIndex        =   25
         Top             =   540
         Width           =   3405
      End
      Begin VB.ComboBox cmbLoanScheme 
         Height          =   315
         Left            =   1530
         TabIndex        =   24
         Top             =   900
         Width           =   3405
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "Loan Account No :"
         Height          =   270
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label lblCustName 
         Caption         =   "Customer &Name :"
         Height          =   225
         Left            =   90
         TabIndex        =   27
         Top             =   570
         Width           =   1365
      End
      Begin VB.Label lblLoanScheme 
         Caption         =   "&Loan Scheme :"
         Height          =   225
         Left            =   90
         TabIndex        =   26
         Top             =   930
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   3360
      TabIndex        =   22
      Top             =   4800
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4290
      TabIndex        =   21
      Top             =   4800
      Width           =   885
   End
End
Attribute VB_Name = "frmLoanRepay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public p_LoanId As Long
Private m_rstLoanMast As Recordset


Private Sub SetKannadaCaption()
Dim ctrl As Control
On Error Resume Next
For Each ctrl In Me
If Not TypeOf ctrl Is VScrollBar And _
   Not TypeOf ctrl Is Line And _
     Not TypeOf ctrl Is ProgressBar And _
      Not TypeOf ctrl Is Image Then
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then ctrl.Font.Size = gFontSize
End If
Next
Err.Clear

lblLoanAccNo = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 60)
lblCustName = LoadResString(gLangOffSet + 35)  'Name
lblLoanScheme = LoadResString(gLangOffSet + 214)  'loanScheme
lblPrincipal = LoadResString(gLangOffSet + 20) & " " & LoadResString(gLangOffSet + 40) 'Repay Amount

lblRepayDate = LoadResString(gLangOffSet + 38) & " " & LoadResString(gLangOffSet + 37)   'Date
'lblInterest = LoadResString(gLangOffSet + 344)
lblMisc = LoadResString(gLangOffSet + 327) 'Misceleneous
lblTotInst = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 42)  'Total Amount
lblRemarks = LoadResString(gLangOffSet + 261) 'Remarks
lblPayIntBalance = LoadResString(gLangOffSet + 67) & " " & LoadResString(gLangOffSet + 47)  'Interest balance
lblPenal = LoadResString(gLangOffSet + 345) '"PenalInterest
'lblNewBalance = LoadResString(gLangOffSet + 260) & " " & LoadResString(gLangOffSet + 42) 'New BAlance

cmdCancel.Caption = LoadResString(gLangOffSet + 2)  'Cancel
cmdOk.Caption = LoadResString(gLangOffSet + 1)      'OK

End Sub

Private Function LoadLoanDetials()
If p_LoanId = 0 Then Exit Function
Dim SqlStr As String
Dim Rst As Recordset

Dim obj As Object
Dim CustId As Long
Dim SchemeId As Integer

'Get the LOanDetails
SqlStr = "SELECT * From LoanMaster WHERE LOanID = " & p_LoanId
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(m_rstLoanMast, adOpenStatic) < 1 Then Exit Function

'Get Customer Name
CustId = FormatField(m_rstLoanMast("CustomerID"))

Set obj = New clsCustReg
txtCustName = obj.CustomerName(CustId)
Set obj = Nothing

'Get Loan NAme
SchemeId = FormatField(m_rstLoanMast("SchemeID"))
'set The Como Box
Dim Count As Integer
For Count = 0 To cmbLoanScheme.ListCount - 1
    If cmbLoanScheme.ItemData(Count) = SchemeId Then
        cmbLoanScheme.ListIndex = Count
        Exit For
    End If
Next

Set obj = New clsLoan
txtLoanAccNo = FormatField(m_rstLoanMast("AccNum"))
txtNextIntBalance = obj.InterestBalance(p_LoanId, txtTransDate)
txtNextIntBalance.Tag = txtNextIntBalance
txtPayIntBalance.Enabled = IIf(Val(txtNextIntBalance) > 0, True, False)

txtNextBalance = obj.Balance(p_LoanId, , txtTransDate)
txtNextBalance.Tag = txtNextBalance
txtPrincipal.Enabled = Val(txtNextBalance)

txtNextRegInt = obj.RegularInterest(p_LoanId, , txtTransDate)

txtNextRegInt.Tag = txtNextRegInt
txtRegInt.Enabled = Val(txtNextRegInt)

txtNextPenalInt = obj.PenalInterest(p_LoanId, , txtTransDate)
txtNextPenalInt.Tag = txtNextPenalInt
txtPenalInt.Enabled = Val(txtNextPenalInt)

Set obj = Nothing

End Function

Private Function LoanRepayment() As Boolean

Dim SqlStr As String
Dim Rst As Recordset
Dim TransDate As String
Dim TransId As Long
Dim TransType As wisTransactionTypes

Dim Amount As Currency
Dim PrinAmount As Currency
Dim IntAmount As Currency
Dim PenalIntAmount As Currency
Dim IntBalAmount As Currency
Dim Balance As Currency
Dim IntBalance As Currency
Dim MiscAmount As Currency
Dim Particulars As String

Dim InstType As wisInstallmentTypes
Dim InstAmount As Currency
Dim InstBalance As Currency
Dim InstPayment As Currency
Dim InstNo As Integer

TransDate = FormatDate(txtTransDate)
InstType = FormatField(m_rstLoanMast("InstMode"))
InstAmount = FormatField(m_rstLoanMast("InstAmount"))

Dim RstInst As Recordset
'Compare the date w.r.t. LastTransDate
SqlStr = "SELECT TransDate,TransID FROM LoanTrans WHERE LoanID = " & p_LoanId & _
            " AND TransID = (SELECT MAX(TransID) " & _
                " FROM LoanTrans WHERE LoanID = " & p_LoanId & ")"
'added on 1/10/01
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(Rst, adOpenStatic) <> 1 Then
   MsgBox "Error In DataBase - Contact vendor", vbCritical, wis_MESSAGE_TITLE
   Exit Function
End If
If DateDiff("d", TransDate, Rst("TransDate")) > 0 Then
   MsgBox "Specified transaction date is earlier than the " & _
      " last transaction date", vbInformation, wis_MESSAGE_TITLE
   Exit Function
End If

TransId = FormatField(Rst("TransID")) + 1
Amount = Val(txtPrincipal)
Balance = Val(txtNextBalance)
IntBalance = (Val(txtNextRegInt) + Val(txtNextIntBalance))

IntBalAmount = Val(txtPayIntBalance)
IntAmount = Val(txtRegInt)
PenalIntAmount = Val(txtPenalInt)

txtPayIntBalance.Text = IntBalance

    
Set RstInst = Nothing
If InstType <> Inst_No Then
    SqlStr = "SELECT * FROM LoanInst Where LoanID = " & p_LoanId & _
        " AND InstBalance > 0 ORDER BY InstDate"
    gDBTrans.SQLStmt = SqlStr
    Call gDBTrans.Fetch(RstInst, adOpenDynamic)
End If


TransType = wDeposit
Dim LCls As New clsLoan

LoanRepayment = LCls.RepayLoan(p_LoanId, txtTransDate, Val(txtPrincipal), "", Val(txtRegInt), Val(txtPenalInt), Val(txtPayIntBalance), , Val(txtMisc))

Set LCls = Nothing

'Inserting the value is did by the lkoan class
'so we exiting from this function
Exit Function



If Balance < 0 Then
    MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

SqlStr = " INSERT INTO LoanTrans(LoanId,TransDate,TransId," _
     & " TransType,Amount,Balance,Particulars)" _
     & " VALUES ( " _
     & p_LoanId & "," _
     & " #" & TransDate & "#," _
     & TransId & "," _
     & TransType & "," _
     & Amount & "," _
     & Balance & "," _
     & AddQuotes(Particulars, True) & _
      ")"
  
gDBTrans.BeginTrans
gDBTrans.SQLStmt = SqlStr
If Not gDBTrans.SQLExecute Then
    gDBTrans.RollBack
    Exit Function
End If


If Balance = 0 Then
    SqlStr = "UPDATE LoanMaster SET LoanClosed = True " & _
        " WHERE LoanID = " & p_LoanId
    gDBTrans.SQLStmt = SqlStr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
    'Reduce the amount of installment
    SqlStr = "UPDATE LoanMaster SET InstBlance = 0 " & _
        " WHERE LoanID = " & p_LoanId
    gDBTrans.SQLStmt = SqlStr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
End If
'Then Insert the reords into LoanInt Trans
SqlStr = " INSERT INTO LoanIntTrans(LoanId,TransDate,TransId," _
     & " TransType,IntAmount,PenalIntAmount,MiscAmount,IntBalance) " _
     & " VALUES ( " _
     & p_LoanId & "," _
     & " #" & TransDate & "#," _
     & TransId & "," _
     & TransType & "," _
     & IntAmount & "," _
     & PenalIntAmount & "," _
     & MiscAmount & "," _
     & IntBalance & _
      ")"
  
gDBTrans.SQLStmt = SqlStr
If Not gDBTrans.SQLExecute Then
    gDBTrans.RollBack
    Exit Function
End If

'If Loan HAs Instalments then
'Update the Installment details To The LoanInst table
If FormatField(m_rstLoanMast("EMI")) = True Then
    MsgBox "Doubts has to be Cleared about bifuracation of amount"
    'Amount = Val(txtPrincipal) + Val(txtRegInt)
End If

If Not RstInst Is Nothing Then
    Do
        If RstInst.EOF Then Exit Do
        If Amount <= 0 Then Exit Do
        InstAmount = FormatField(RstInst("InstBalance"))
        InstNo = FormatField(RstInst("InstNo"))
        If InstAmount >= Amount Then
            InstBalance = InstAmount - Amount
            Amount = 0 'Amount - Instp
        Else
            InstBalance = 0
            Amount = Amount - InstAmount
        End If
        SqlStr = "UPDATE LoanInst  Set InstBalance = " & InstBalance & _
            ", PaidDate = #" & TransDate & "#" & _
            " WHERE LoanID = " & p_LoanId & _
            " AND InstNo = " & InstNo
        gDBTrans.SQLStmt = SqlStr
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            Exit Function
        End If
        RstInst.MoveNext
    Loop
End If

gDBTrans.CommitTrans

LoanRepayment = True
End Function

Private Function ValidatingControls() As Boolean
Err.Clear
On Error GoTo ExitLine
 
 'validating for Interest balance
If txtRegInt.Enabled = True Then
    If txtRegInt.Text = "" Then txtRegInt = 0
    If Not CurrencyValidate(txtRegInt.Text, True) Then
        MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtRegInt
        Exit Function
    End If
  
  'validate the Regular interest
  '  If Val(txtRegInt) = 0 And Val(txtNextRegInt) > 0 Then
   If txtNextRegInt.Text = "" Then txtNextRegInt = 0
    If Not CurrencyValidate(txtNextRegInt.Text, True) Then
      MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
      txtRegInt.SetFocus
      Exit Function
    End If
End If
  
  'validate the Interest balance
 If txtPayIntBalance.Enabled = True Then
    If txtPayIntBalance.Text = "" Then txtPayIntBalance = 0
    If Not CurrencyValidate(txtPayIntBalance.Text, True) Then
        MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
        txtPayIntBalance.SetFocus
        Exit Function
    End If
    If Val(txtPayIntBalance) = 0 And Val(txtNextIntBalance) > 0 Then
        If MsgBox("You have not speicified the InterestBalance" & vbCrLf & _
           "Do you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, _
           wis_MESSAGE_TITLE) = vbNo Then
           txtPayIntBalance.SetFocus
           Exit Function
          End If
    End If
 End If
   'Validate The Date of Transaction
   If Not DateValidate(txtTransDate, "/", True) Then
      MsgBox LoadResString(gLangOffSet + 501), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtTransDate
   End If

    'If Trim(txtPrincipal.Text) = "" Then
  If txtPrincipal.Text = "" Then txtPrincipal = 0
  If Not CurrencyValidate(txtPrincipal.Text, True) Then
      MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtPrincipal
      Exit Function
  End If
  If Val(txtPrincipal.Text) = 0 Then
      If MsgBox("You are not speicified the Principal amount" & vbCrLf & _
         "Do you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, _
         wis_MESSAGE_TITLE) = vbNo Then
         ActivateTextBox txtPrincipal
         Exit Function
      End If
  End If
  
  If txtPenalInt.Enabled = True Then
  If txtPenalInt.Text = "" Then txtPenalInt = 0
    If Not CurrencyValidate(txtPenalInt.Text, True) Then
        MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtPenalInt
        Exit Function
    End If
    If Val(txtPenalInt.Text) = 0 And Val(txtNextPenalInt) > 0 Then
        If MsgBox("You are not speicified the Penal interest" & vbCrLf & _
           "Do you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, _
           wis_MESSAGE_TITLE) = vbNo Then
           ActivateTextBox txtPenalInt
           Exit Function
        End If
    End If
  End If
  
  
  If Not CurrencyValidate(txtMisc.Text, True) And txtMisc <> "" Then
      MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
      txtMisc.SetFocus
      Exit Function
  End If
  
  ValidatingControls = True

ExitLine:
  
 End Function








Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim SqlStr As String
If Not ValidatingControls Then Exit Sub
'ALL THE DATA BASE STRUCTURE IS GIVEN IN LOANS.TAB FILE

'Then Insert The recoreds Into LoanTrans
If LoanRepayment Then
    MsgBox "Loan amount repaid", vbInformation, wis_MESSAGE_TITLE
    Unload Me
End If
  
End Sub


Private Sub cmdRepayDate_Click()
Dim s_date As String
s_date = Format(Now, "DD/MM/YYYY")

Dim my_da As String

With Calendar
  .Left = Screen.Width / 2
  .Top = Screen.Height / 2
  .SelDate = Format(Now, "dd/mm/yyyy")
  .Show vbModal
   txtTransDate = .SelDate
   Call LoadLoanDetials
End With
 
End Sub

Private Sub Form_Load()
'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption

''LOad the Loan schemed
Call LoadLoanSchemes(Me.cmbLoanScheme)
If txtTransDate = "" Then txtTransDate = FormatDate(gStrDate)
Call LoadLoanDetials
 
End Sub










Private Sub txtMisc_Change()
txtTotInst = FormatCurrency(Val(txtPrincipal) + Val(txtPayIntBalance) + _
    Val(txtRegInt) + Val(txtPenalInt) + Val(txtMisc))

End Sub

Private Sub txtPayIntBalance_Change()
txtNextIntBalance = FormatCurrency(Val(txtNextIntBalance.Tag) - Val(txtPayIntBalance))
txtTotInst = FormatCurrency(Val(txtPrincipal) + Val(txtPayIntBalance) + _
    Val(txtRegInt) + Val(txtPenalInt) + Val(txtMisc))

End Sub

Private Sub txtPenalInt_Change()
txtNextPenalInt = FormatCurrency(Val(txtNextPenalInt.Tag) - Val(txtPenalInt))
txtTotInst = FormatCurrency(Val(txtPrincipal) + Val(txtPayIntBalance) + _
    Val(txtRegInt) + Val(txtPenalInt) + Val(txtMisc))
End Sub


Private Sub txtPrincipal_Change()
txtNextBalance = FormatCurrency(Val(txtNextBalance.Tag) - Val(txtPrincipal))
txtTotInst = FormatCurrency(Val(txtPrincipal) + Val(txtPayIntBalance) + _
    Val(txtRegInt) + Val(txtPenalInt) + Val(txtMisc))
End Sub


Private Sub txtRegInt_Change()
txtNextRegInt = FormatCurrency(Val(txtNextRegInt.Tag) - Val(txtRegInt))
txtTotInst = FormatCurrency(Val(txtPrincipal) + Val(txtPayIntBalance) + _
    Val(txtRegInt) + Val(txtPenalInt) + Val(txtMisc))
End Sub

Private Sub txtTotInst_GotFocus()
txtTotInst = Val(txtPayIntBalance) + Val(txtRegInt) + Val(txtPenalInt) + Val(txtPrincipal) + Val(txtMisc)
End Sub


Private Sub txtTransDate_LostFocus()
If Not DateValidate(txtTransDate, "/", True) Then Exit Sub
Call LoadLoanDetials
End Sub


