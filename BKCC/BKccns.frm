VERSION 5.00
Object = "{8491A895-6031-11D5-A300-0080AD7CA942}#10.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmBkccInst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Instalment"
   ClientHeight    =   4545
   ClientLeft      =   2280
   ClientTop       =   2085
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6465
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5310
      TabIndex        =   17
      Top             =   4080
      Width           =   885
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   4320
      TabIndex        =   16
      Top             =   4080
      Width           =   885
   End
   Begin VB.Frame fraInstall 
      Height          =   2175
      Left            =   60
      TabIndex        =   28
      Top             =   1860
      Width           =   6135
      Begin WIS_Currency_Text_Box.CurrText txtInterest 
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "..."
         Height          =   315
         Left            =   5640
         TabIndex        =   14
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtSBAccNum 
         Height          =   285
         Left            =   4590
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1800
         Width           =   945
      End
      Begin WIS_Currency_Text_Box.CurrText txtNewBalance 
         Height          =   315
         Left            =   4770
         TabIndex        =   10
         Top             =   990
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepositBalance 
         Height          =   315
         Left            =   4770
         TabIndex        =   6
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtNewLoan 
         Height          =   315
         Left            =   1770
         TabIndex        =   8
         Top             =   990
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.CheckBox chkSb 
         Caption         =   "Credit To SB Account"
         Height          =   255
         Left            =   1770
         TabIndex        =   13
         Top             =   1800
         Width           =   2325
      End
      Begin VB.TextBox txtRemark 
         Height          =   345
         Left            =   1770
         TabIndex        =   12
         Top             =   1380
         Width           =   4245
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   285
         Left            =   3150
         TabIndex        =   2
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1770
         TabIndex        =   1
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label lblDepositBalance 
         Caption         =   "Deposit Balance :"
         Height          =   255
         Left            =   3330
         TabIndex        =   5
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblNewBalance 
         Caption         =   "New Balance"
         Height          =   225
         Left            =   3390
         TabIndex        =   9
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest till date"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label lblRemarks 
         Caption         =   "Remarks"
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Top             =   1410
         Width           =   1485
      End
      Begin VB.Label lblAmtSanctioned 
         Caption         =   "Amount Sanctioned"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Caption         =   "Date :"
         Height          =   225
         Left            =   150
         TabIndex        =   0
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Frame fraLoan 
      Height          =   1935
      Left            =   60
      TabIndex        =   18
      Top             =   30
      Width           =   6135
      Begin WIS_Currency_Text_Box.CurrText txtLoanBalance 
         Height          =   345
         Left            =   1770
         TabIndex        =   29
         Top             =   1440
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.TextBox txtLoanName 
         Height          =   345
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1020
         Width           =   4215
      End
      Begin VB.TextBox txtLoanID 
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtMemID 
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   885
      End
      Begin VB.Label txtLoanDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4710
         TabIndex        =   31
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1770
         TabIndex        =   30
         Top             =   630
         Width           =   4215
      End
      Begin VB.Label lblLoanDate 
         Caption         =   "Loan Date :"
         Height          =   195
         Left            =   3360
         TabIndex        =   27
         Top             =   1500
         Width           =   1125
      End
      Begin VB.Label lblLoanBalance 
         Caption         =   "Loan Balance :"
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   1470
         Width           =   1545
      End
      Begin VB.Label lblLoanName 
         Caption         =   "Loan  Name :"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   1050
         Width           =   1545
      End
      Begin VB.Label lblLoanID 
         Caption         =   "Loan Id :"
         Height          =   270
         Left            =   3420
         TabIndex        =   24
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblName 
         Caption         =   "Loan Holder Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label lblMemID 
         Caption         =   "Memebre Id :"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmBkccInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public p_LoanId As Long
Public p_AccID As Long

Private m_rstLoanMast As Recordset
Private m_rstLoanTrans As Recordset
Private m_InstalmentDetails As String
Private m_LoanTransTable As String
Private m_IsBKCCDeposit As Boolean
Private m_DepositInterest As Currency
Private Function LoanTransaction() As Boolean
Dim Amount As Currency
Dim NewBalance As Currency
Dim DepBalance As Currency
Dim Ret As Long
Dim Rst As Recordset

Dim TransDate As Date

'First Validate TextBoxes
If Not DateValidate(txtDate.Text, "/", True) Then
    MsgBox LoadResString(gLangOffSet + 501), vbExclamation, wis_MESSAGE_TITLE
    ActivateTextBox txtDate
    GoTo Exit_Line
End If
TransDate = FormatDate(txtDate)

If txtNewLoan = 0 Then
    MsgBox LoadResString(gLangOffSet + 499), vbExclamation, wis_MESSAGE_TITLE
    txtNewLoan.SetFocus
    GoTo Exit_Line
End If

If txtInterest = 0 Then
    'If MsgBox("You have not specified Interest" & vbCrLf & "Do you want to continue", _
            vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
    If MsgBox(LoadResString(gLangOffSet + 506) & vbCrLf & LoadResString(gLangOffSet + 541), _
            vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
        txtNewLoan.SetFocus
        GoTo Exit_Line
    End If
End If


Dim TransID As Long
Dim TransType  As wisTransactionTypes
Dim Remarks As String
Dim Balance As Currency
Dim LoanAmount As Currency
Dim Deposit As Boolean


' Now check for the date of transaction
gDBTrans.SQLStmt = "Select top 1 TransId, TransDate From BKCCTrans " & _
    " Where Loanid = " & p_LoanId & " ORDER BY TransID Desc"
TransID = 100
If gDBTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    If DateDiff("d", Rst("TransDate"), TransDate) < 0 Then
        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate
        GoTo Exit_Line
    End If
    TransID = Rst("TransId")
    Balance = Rst("Balance")
End If

' Now check for the date of transaction w.r.t. interest
gDBTrans.SQLStmt = "Select Top 1 TransId, TransDate From BKCCIntTrans " & _
    " Where Loanid = " & p_LoanId & " ORDER BY TransID Desc"
TransID = 100
If gDBTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    If DateDiff("d", Rst("TransDate"), TransDate) < 0 Then
        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate
        GoTo Exit_Line
    End If
    TransID = IIf(TransID > Rst("TransId"), TransID, Rst("TransId"))
End If

LoanAmount = txtNewLoan

Remarks = txtRemark
If Len(Remarks) > 50 Then Remarks = Left(txtRemark.Text, 50)

TransType = wWithdraw

gDBTrans.BeginTrans

Deposit = IIf(Balance < 0, True, False)
DepBalance = IIf(Balance < 0, Abs(Balance), 0)

Amount = 0
If LoanAmount <> 0 Then
    'Insert the Loan Amount
    TransType = wWithdraw
    If DepBalance Then
        Deposit = True
        NewBalance = Balance + LoanAmount - txtInterest
        If NewBalance >= 0 Then
            Amount = DepBalance + txtInterest
            NewBalance = 0
        Else
            Amount = txtNewLoan
            NewBalance = -txtDepositBalance
        End If
        'if the Amount he is withdrawing is still less than then his deposit amount
        'then confirm to add the interest on deposit
        If ((DepBalance - LoanAmount) > 0) And txtInterest > 0 Then
            Ret = MsgBox("Do you want to add the Interest to the Deposit ?", _
                vbInformation + vbYesNo + vbDefaultButton2, "Add Interest to Deposit")
        End If
        
        'if the Amount withdrawing is more than that of Deposit AMount
        'then Add the interest on deposit to the Deposit
        If DepBalance - LoanAmount <= 0 Or Ret = vbYes Then
            'First Insert the Interest Amount
            'Debit to loss account and credit to BKCC account
            TransID = TransID + 1
            TransType = wContraWithDraw
            gDBTrans.SQLStmt = "Insert Into BKCCIntTrans (LoanId,TransID," & _
                "TransType,IntAmount,PenalIntAmount,TransDate" & _
                ") Values " & "(" & _
                p_LoanId & ", " & TransID & "," & _
                TransType & "," & txtInterest & ", 0," & _
                "#" & TransDate & "#, " & ")"
                
            If Not gDBTrans.SQLExecute Then
                gDBTrans.RollBack
                GoTo err_line
            End If
            
            'Noe depoist the Inteest amoun to the Deposit
            TransType = wContraDeposit
            gDBTrans.SQLStmt = "Insert Into BKCCTrans (LoanId,TransID," & _
                " TransType,Amount,TransDate," & _
                " Balance,Particulars) Values " & _
                "(" & p_LoanId & ", " & TransID & "," & _
                TransType & "," & txtInterest & "," & _
                "#" & TransDate & "#, " & Balance - txtInterest & "," & _
                "'Interest Deposited'  )"
                
            If Not gDBTrans.SQLExecute Then
                gDBTrans.RollBack
                GoTo err_line
            End If
            If NewBalance <> 0 Then NewBalance = NewBalance - txtInterest
        End If
        
        'Now with draw the Loanamount
        TransID = TransID + 1
        TransType = wWithdraw
        gDBTrans.SQLStmt = "Insert Into BKCCTrans (LoanId,TransID," & _
            "TransType,Amount,TransDate," & _
            " Balance,Particulars) Values " & _
            "(" & p_LoanId & ", " & TransID & "," & _
            TransType & "," & Amount & "," & _
            "#" & TransDate & "#," & NewBalance & "," & _
            "'Amount WithDrawn'  )"
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            GoTo err_line
        End If
        Balance = Balance + LoanAmount - txtInterest
        If Balance > 0 Then Amount = Balance
    Else
        Balance = Balance + LoanAmount
    End If
    
    TransType = wWithdraw
    If Balance > 0 Then
        TransID = TransID + 1
        gDBTrans.SQLStmt = "Insert Into BKCCTrans " & _
            " (LoanId,TransID,TransType," & _
            " Amount,TransDate,Balance,Particulars) " & _
            " Values (" & p_LoanId & ", " & TransID & "," & _
            TransType & "," & _
            IIf(Amount > 0, Amount, LoanAmount) & "," & _
            "#" & TransDate & "#," & Balance & "," & _
            AddQuotes(Remarks, True) & ")"
        
        Amount = Amount + LoanAmount
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            GoTo err_line
        End If
    End If
End If
gDBTrans.CommitTrans


LoanTransaction = True
' Credit The Corresponding Loan Amount Into SBTransTab
'Get The SB Accid For This MemberID
Dim MemID As Long
Dim CustId As Long
Dim SbAccID As Long
'Dim SbObj As clsSBAcc

If chkSb.value Then
    
    MemID = Val(FormatField(m_rstLoanMast("MemberID")))
    gDBTrans.SQLStmt = "Select CustomerID From MMMaster Where Accid = " & MemID
    
    If gDBTrans.Fetch(Rst, adOpenDynamic) <> 1 Then
        MsgBox "Unable to update SBAccount), vbExclamation, wis_MESSAGE_TITLE"
        GoTo Exit_Line
    End If
    CustId = Val(Rst("CustomerID"))
'    Set SbObj = New clsSBAcc
'    If SbObj.CustomerBalance(CustId, SbAccID) > 0 Then
'        If Not SbObj.DepositAmount(SbAccID, Amount, "Loan Sanctioned", txtDate.Text) Then
'            MsgBox "Cannot Update The SB Account ", vbExclamation, wis_MESSAGE_TITLE
'        End If
'    End If
'    Set SbObj = Nothing
    
End If
    
Exit_Line:
Exit Function

err_line:
    MsgBox "Error In BKCC LoanTransaction"

End Function

Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next Ctrl
lblMemID.Caption = LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 60)
lblLoanID.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 60)
lblName.Caption = LoadResString(gLangOffSet + 58)
lblLoanName.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 35)
lblLoanBalance.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 42)
lblLoanDate.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 37)
lblDate.Caption = LoadResString(gLangOffSet + 37)
lblInterest.Caption = LoadResString(gLangOffSet + 47)
lblAmtSanctioned.Caption = LoadResString(gLangOffSet + 247)
lblNewBalance.Caption = LoadResString(gLangOffSet + 260) & " " & LoadResString(gLangOffSet + 42)
lblRemarks.Caption = LoadResString(gLangOffSet + 261)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
cmdOk.Caption = LoadResString(gLangOffSet + 1)
Me.chkSb.Caption = LoadResString(gLangOffSet + 421) & "  " & LoadResString(gLangOffSet + 271) '"Credit To SB Account"
If m_IsBKCCDeposit Then
    lblDepositBalance.Caption = LoadResString(gLangOffSet + 43) & " " & LoadResString(gLangOffSet + 42)
    lblDepositBalance.Visible = True
    txtDepositBalance.Visible = True
End If
End Sub

Private Sub SetRecordSets()
Dim Retval As Long
Dim Rst As Recordset

'Set LoanMaster Recpords
gDBTrans.SQLStmt = "Select * From LoanMaster Where LoanId = " & p_LoanId
'Call gDBTrans.Fetch(m_rstLoanMast, adOpenDynamic)
Call gDBTrans.Fetch(Rst, adOpenDynamic)
txtLoanDate = FormatField(Rst("IssueDate"))

'Set the Mebber Details
txtMemID.Text = FormatField(Rst("MemberId"))

Dim MMObj As clsMMAcc
Set MMObj = New clsMMAcc
txtName = MMObj.MemberName(Val(txtMemID.Text))
Set MMObj = Nothing

'Set Loan Deatils
txtDate.Text = FormatDate(gStrDate)
gDBTrans.SQLStmt = "Select * From LoanTrans Where " & _
    " LoanId = " & p_LoanId & " ORDER BY TransID Desc"

If gDBTrans.Fetch(Rst, adOpenForwardOnly) Then
    txtLoanBalance = IIf(FormatField(Rst("Balance")) < 0, 0, FormatField(Rst("Balance")))
    txtDepositBalance = IIf(FormatField(Rst("Balance")) >= 0, _
                    FormatCurrency(0), FormatCurrency(Abs(Rst("Balance"))))
    Call PutInterest
    txtNewBalance = txtLoanBalance
End If

End Sub


Private Sub PutInterest()

Dim RegInterest As Currency
Dim PenalInterest As Currency
Dim NewLoanBalance As Currency
Dim Balance As Currency
Dim LastTransDate As Date
Dim Days As Long
Dim Rst As Recordset

' To Get Calulation of Interest we Hve to Set the repay
''Date Of Loan Acc form To the PresentCalculation Date
    gDBTrans.SQLStmt = "SELECT TOP 1 TransDate, Balance FROM " & _
        " BKCCTrans WHERE LoanID = " & p_LoanId & " ORDER BY TransID DESC"
    If gDBTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        Balance = FormatField(Rst("Balance"))
        If Balance < 0 Then
            LastTransDate = Rst("TransDate")
            Days = DateDiff("d", LastTransDate, FormatDate(txtDate.Text))
                PenalInterest = DepositInterest_New(p_LoanId, FormatDate(txtDate.Text))
                txtInterest.Text = FormatCurrency(PenalInterest)
                txtDepositBalance.Text = FormatCurrency(Abs(m_DepositBalance))
                NewLoanBalance = Val(txtDepositBalance.Text) - _
                    Val(txtNewLoan.Text)
                txtNewBalance.Text = IIf(NewLoanBalance < 0, FormatCurrency(Abs(NewLoanBalance)), _
                    FormatCurrency(0))
                txtDepositBalance.Text = IIf(NewLoanBalance >= 0, _
                    FormatCurrency(Abs(NewLoanBalance)), _
                    FormatCurrency(0))
                Exit Sub
        Else
            GoTo LoanInterest
        End If
    End If

LoanInterest:
'RegInterest = frmLoanAcc.ComputeRegularInterest(txtDate.Text, p_LoanId)
'PenalInterest = frmLoanAcc.ComputePenalInterest(txtDate.Text, p_LoanId)
PenalInterest = IIf(PenalInterest < 0, 0, PenalInterest)
PenalInterest = PenalInterest + RegInterest
txtInterest.Text = FormatCurrency(PenalInterest \ 1)

End Sub


Public Function DepositInterest(ByVal LoanID As Long, ByVal TransDate As Date) As Currency
Dim rst_PM As Recordset
Dim rst_1TO10 As Recordset
Dim rst_11TO31 As Recordset
Dim DateLimit As String
Dim YearLimit As Integer
Dim Balance As Currency
Dim Balance1TO10 As Currency
Dim Balance11TO31 As Currency
Dim TransID As Long
Dim Dy As Integer
Dim Mon As Integer
Dim MonLimit As Integer
Dim Yr As Integer
Dim Interest As Currency
Dim Rst As Recordset

gDBTrans.SQLStmt = "SELECT TOP 1 TransDate, Balance FROM BKCCTrans WHERE" & _
    " LoanID = " & LoanID & " And Balance < 0 ORDER BY TransID ASC"

Call gDBTrans.Fetch(Rst, adOpenDynamic)
DateLimit = Rst("TransDate")
Balance = Abs(gDBTrans.Rst("Balance"))

gDBTrans.SQLStmt = "SELECT Top 1 IntAmount, TransDate" & _
    " FROM BKCCIntTrans WHERE LoanID = " & LoanID & " ORDER BY TransID DESC"
If gDBTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
    Interest = FormatField(Rst("Amount"))

If Interest > 0 Then DateLimit = Rst("TransDate")
'''    Balance = Abs(gDBTrans.Rst("Balance"))

If DateDiff("d", DateLimit, TransDate) < 30 Then Exit Function

Mon = Month(DateLimit)
MonLimit = Month(TransDate)
Yr = Year(DateLimit)
YearLimit = Year(TransDate)
Dy = Day(TransDate)


If Dy < 30 Then MonLimit = MonLimit - 1

gDBTrans.SQLStmt = "SELECT Max(TransID) AS MaxTransID, Month(transdate) AS Months" & _
    " From BKCCTrans" & _
    " WHERE TransDate BETWEEN #" & DateLimit & "# And #" & _
    TransDate & "# And LoanID = " & AccountNo & _
    " And Balance < 0 GROUP BY Month(TransDate);"
Call gDBTrans.Fetch(rst_PM, adOpenDynamic)

'Balance before 10th of this month
gDBTrans.SQLStmt = "SELECT MAX(TransID) as MaxTransID, Month(TransDate) as Months" & _
    " from BKCCTrans WHERE Day(TransDate) < 11 and transdate between" & _
    " #" & DateLimit & "# And #" & TransDate & "# and LoanID = " & AccountNo & _
    " And Balance < 0 GROUP BY Month(transdate);"

Call gDBTrans.Fetch(rst_1TO10, adOpenDynamic)

'Minimum Balance between 11th and last Day of this month
gDBTrans.SQLStmt = "SELECT Min(Balance) as MinBalance, Month(TransDate) as Months" & _
    " from BKCCTrans WHERE Day(TransDate) >= 11 and transdate between" & _
    " #" & DateLimit & "# And #" & TransDate & "# and LoanID = " & AccountNo & _
    " And Balance < 0 GROUP BY Month(transdate);"
Call gDBTrans.Fetch(rst_11TO31, adOpenDynamic)

Interest = 0
While (Mon >= MonLimit And Yr < YearLimit) Or (Mon <= MonLimit And Yr = YearLimit)

    Balance1TO10 = 0: Balance11TO31 = 0: TransID = 0
    rst_PM.MoveFirst
    rst_PM.Find "Months=" & Mon - 1
    If Not rst_PM.EOF Then TransID = rst_PM.Fields("MaxTransID")
    If TransID > 0 Then
        gDBTrans.SQLStmt = "SELECT Balance FROM BKCCTrans WHERE LoanID = " & _
            AccountNo & " And TransID = " & TransID
        Call gDBTrans.Fetch(Rst, adOpenDynamic)
        Balance = Abs(Rst(0))
        TransID = 0
    End If
    
    If Not rst_1TO10 Is Nothing Then
        rst_1TO10.MoveFirst
        rst_1TO10.Find "Months=" & Mon
        If Not rst_1TO10.EOF Then TransID = rst_1TO10.Fields("MaxTransID")
        If TransID > 0 Then
            gDBTrans.SQLStmt = "SELECT Balance FROM BKCCTrans WHERE LoanID = " & _
                AccountNo & " And TransID = " & TransID
            Call gDBTrans.Fetch(Rst, adOpenDynamic)
            Balance1TO10 = Abs(Rst(0))
            Balance = Balance1TO10
        End If
    End If
    
    If Not rst_11TO31 Is Nothing Then
        rst_11TO31.MoveFirst
        rst_11TO31.Find "Months =" & Mon
        If Not rst_11TO31.EOF Then Balance11TO31 = Abs(rst_11TO31.Fields("MinBalance"))
    End If
    
    If Balance1TO10 > 0 Then
        If Balance11TO31 > 0 Then Balance = Balance11TO31
    ElseIf Balance11TO31 > 0 Then
        Balance = IIf(Balance <= Balance11TO31, Balance, Balance11TO31)
    End If

    Interest = Interest + (Balance * 1 * m_DepositInterest) / (100 * 12)
    Mon = Mon + 1
    If Mon > 12 Then
        Yr = Yr + 1
        Mon = 1
    End If
Wend

Interest = Interest \ 1
DepositInterest_New = Interest

End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDate_Click()
With Calendar
    
    .Left = Me.Left + fraInstall.Left + cmdDate.Left - (.Width / 2)
    .Top = Me.Top + fraInstall.Top + cmdDate.Top
    If DateValidate(txtDate.Text, "/", True) Then
        .SelDate = FormatDate(txtDate.Text)
    Else
        .SelDate = FormatDate(gStrDate)
    End If
    .Show vbModal
    txtDate.Text = FormatDate(.SelDate)
End With

End Sub

Private Sub cmdOK_Click()
    If LoanTransaction Then Unload Me
End Sub



Private Sub cmdSb_Click()
Dim SqlStr As String
Dim SearchName As String
Dim Rst As Recordset

SqlStr = "Select FirstName,MiddleName,LastName " & _
    "From NameTab Where CustomerId = (Select customerID From BKCCMAster " & _
        "WHERE LoanID = " & p_LoanId & " );"
gDBTrans.SQLStmt = SqlStr
SqlStr = ""
If gDBTrans.Fetch(Rst, adOpenForwardOnly) Then
    SearchName = Trim(FormatField(Rst(0)))
    If SearchName = "" Then SearchName = Trim(FormatField(Rst(1)))
    If SearchName = "" Then SearchName = Trim(FormatField(Rst(2)))
Else
    Exit Sub
End If

SqlStr = "Select AccNum, A.AccID ," & _
    " Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
    " From SBMaster A, NameTab B WHERE B.CustomerID = A.CustomerID"

If Trim(SearchName) <> "" Then
    SqlStr = SqlStr & " AND (FirstName like '" & SearchName & "%' " & _
        " Or MiddleName like '" & SearchName & "%' " & _
        " Or LastName like '" & SearchName & "*')"
End If

gDBTrans.SQLStmt = SqlStr & " ORDER By IsciName"

If gDBTrans.Fetch(rstReturn, adOpenStatic) < 1 Then
    MsgBox "There are no customers in the " & LoadResString(gLangOffSet + 245), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If



End Sub

Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
If p_LoanId = 0 Then
    Err.Raise 2002, "Loan Instalment", "You have not set the LoanId or Account Id"
    Exit Sub
End If

Dim s_SetUp As New clsSetup
m_DepositInterest = CCur(s_SetUp.ReadSetupValue(wis_BKCC, "DEPOSIT_INTEREST", _
    CStr(0)))
Set s_SetUp = Nothing

txtLoanID.Text = p_LoanId
txtMemID.Text = p_AccID
'm_LoanTransTable = frmLoanAcc.m_TransTable
Call SetRecordSets


End Sub

Private Sub txtDate_Change()
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
PutInterest
End Sub

Private Sub txtInterest_Change()
Dim NewLoanBalance As Currency
NewLoanBalance = Abs(m_DepositBalance) + Val(txtInterest.Text) - _
    Val(txtNewLoan.Text)
txtNewBalance.Text = IIf(NewLoanBalance < 0, FormatCurrency(Abs(NewLoanBalance)), _
    FormatCurrency(0))
If NewLoanBalance < 0 Then
    txtDepositBalance.Text = FormatCurrency(0)
ElseIf NewLoanBalance <= Val(txtInterest.Text) Then
    txtDepositBalance.Text = FormatCurrency(NewLoanBalance)
Else
    txtDepositBalance.Text = FormatCurrency(Abs(m_DepositBalance) - Val(txtNewLoan.Text))
End If

'''txtDepositBalance.Text = IIf(NewLoanBalance >= 0, _
'''    FormatCurrency(Abs(NewLoanBalance)), _
'''    FormatCurrency(0))
End Sub
Private Sub txtInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtLoanBalance_GotFocus()
'Me.ActiveControl.SelStart = 0
'Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub



Private Sub txtNewLoan_Change()
Dim NewLoanBalance As Currency
If m_IsBKCCDeposit Then
    NewLoanBalance = Abs(m_DepositBalance) + Val(txtInterest.Text) - _
        Val(txtNewLoan.Text)
    txtNewBalance.Text = IIf(NewLoanBalance < 0, FormatCurrency(Abs(NewLoanBalance)), _
        FormatCurrency(0))
    If NewLoanBalance < 0 Then
        txtDepositBalance.Text = FormatCurrency(0)
    ElseIf NewLoanBalance <= Val(txtInterest.Text) Then
        txtDepositBalance.Text = FormatCurrency(NewLoanBalance)
    Else
        txtDepositBalance.Text = FormatCurrency(Abs(m_DepositBalance) - _
            Val(txtNewLoan.Text))
    End If
Else
    txtNewBalance.Text = FormatCurrency(Val(txtNewLoan.Text) + _
        Val(txtLoanBalance.Text))
End If
End Sub
Private Sub txtNewLoan_GotFocus()
'Me.ActiveControl.SelStart = 0
'Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub

Private Sub txtRemark_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


