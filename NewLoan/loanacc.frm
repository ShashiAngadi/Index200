VERSION 5.00
Begin VB.Form frmLoanAcc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX-2000   -   Loan Wizard"
   ClientHeight    =   6705
   ClientLeft      =   765
   ClientTop       =   1170
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   315
      Left            =   6090
      TabIndex        =   0
      Top             =   6240
      Width           =   1365
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   5895
      Left            =   300
      TabIndex        =   1
      Top             =   180
      Width           =   7305
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report :"
         Height          =   2640
         Left            =   285
         TabIndex        =   22
         Top             =   660
         Width           =   6585
         Begin VB.Frame Frame2 
            Height          =   705
            Left            =   0
            TabIndex        =   28
            Top             =   1950
            Width           =   6585
            Begin VB.ComboBox cmbCaste 
               Height          =   315
               Left            =   4470
               TabIndex        =   32
               Top             =   240
               Width           =   1515
            End
            Begin VB.ComboBox cmbPlace 
               Height          =   315
               Left            =   1470
               TabIndex        =   31
               Top             =   270
               Width           =   1485
            End
            Begin VB.CheckBox chkPlace 
               Caption         =   " Place"
               Height          =   285
               Left            =   390
               TabIndex        =   30
               Top             =   270
               Width           =   1095
            End
            Begin VB.CheckBox chkCaste 
               Caption         =   "Caste"
               Height          =   225
               Left            =   3390
               TabIndex        =   29
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.OptionButton optGeneralLedger 
            Caption         =   "General Ledger"
            Height          =   270
            Left            =   4395
            TabIndex        =   9
            Top             =   675
            Width           =   2130
         End
         Begin VB.OptionButton optOverDueLoans 
            Caption         =   "Over due loans"
            Height          =   270
            Left            =   4395
            TabIndex        =   11
            Top             =   1320
            Width           =   2130
         End
         Begin VB.OptionButton optGuarantors 
            Caption         =   "Guarantor's List"
            Height          =   315
            Left            =   2610
            TabIndex        =   34
            Top             =   630
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.OptionButton optInterest 
            Caption         =   "Interest Collected"
            Height          =   315
            Left            =   405
            TabIndex        =   33
            Top             =   1620
            Width           =   2385
         End
         Begin VB.OptionButton optLoanHolders 
            Caption         =   "List of loan holders as on "
            Height          =   210
            Left            =   405
            TabIndex        =   4
            Top             =   1980
            Width           =   2430
         End
         Begin VB.OptionButton OptLoanSanction 
            Caption         =   "Sanctioned Loans"
            Height          =   315
            Left            =   4395
            TabIndex        =   27
            Top             =   1620
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.OptionButton optInstOverDue 
            Caption         =   "Instalment over due loans"
            Height          =   285
            Left            =   405
            TabIndex        =   7
            Top             =   1320
            Value           =   -1  'True
            Width           =   3690
         End
         Begin VB.OptionButton optTransaction 
            Caption         =   "Total transactions made"
            Height          =   255
            Left            =   405
            TabIndex        =   5
            Top             =   675
            Width           =   2295
         End
         Begin VB.OptionButton optLoanBalance 
            Caption         =   "Balances where"
            Height          =   300
            Left            =   405
            TabIndex        =   3
            Top             =   330
            Width           =   2010
         End
         Begin VB.OptionButton optDailyCashBook 
            Caption         =   "Daily cash book"
            Height          =   270
            Left            =   4395
            TabIndex        =   8
            Top             =   330
            Width           =   2040
         End
         Begin VB.OptionButton optRepaymentsMade 
            Caption         =   "Repayments made"
            Height          =   270
            Left            =   4395
            TabIndex        =   10
            Top             =   1005
            Width           =   2130
         End
         Begin VB.OptionButton optLoansIssued 
            Caption         =   "Loans issued"
            Height          =   270
            Left            =   405
            TabIndex        =   6
            Top             =   1005
            Width           =   1830
         End
      End
      Begin VB.Frame fraAmountRange 
         Caption         =   "Amount range :"
         Height          =   900
         Left            =   270
         TabIndex        =   24
         Top             =   4350
         Width           =   6615
         Begin VB.TextBox txtStartAmt 
            Height          =   285
            Left            =   1590
            TabIndex        =   16
            Top             =   345
            Width           =   1200
         End
         Begin VB.TextBox txtEndAmt 
            Height          =   285
            Left            =   4485
            TabIndex        =   17
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label lblAmt1 
            AutoSize        =   -1  'True
            Caption         =   "Between :"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   375
            Width           =   1440
         End
         Begin VB.Label lblAmt2 
            AutoSize        =   -1  'True
            Caption         =   "And :"
            Height          =   195
            Left            =   3615
            TabIndex        =   25
            Top             =   390
            Width           =   525
         End
      End
      Begin VB.ComboBox cmbLoanType 
         Height          =   315
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   4095
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   345
         Left            =   5625
         TabIndex        =   18
         Top             =   5460
         Width           =   1215
      End
      Begin VB.Frame fraDateRange 
         Caption         =   "Date range :"
         Height          =   930
         Left            =   285
         TabIndex        =   19
         Top             =   3390
         Width           =   6585
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   250
            Left            =   5700
            TabIndex        =   14
            Top             =   390
            Width           =   250
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   250
            Left            =   2835
            TabIndex        =   12
            Top             =   360
            Width           =   250
         End
         Begin VB.TextBox txtEndDate 
            Height          =   285
            Left            =   4440
            TabIndex        =   15
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtStartDate 
            Height          =   285
            Left            =   1590
            TabIndex        =   13
            Top             =   345
            Width           =   1200
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "Ending date :"
            Height          =   195
            Left            =   3210
            TabIndex        =   21
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "Starting date :"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   375
            Width           =   1170
         End
      End
      Begin VB.Label lblLoanType 
         AutoSize        =   -1  'True
         Caption         =   "Select a loan type :"
         Height          =   195
         Left            =   285
         TabIndex        =   23
         Top             =   390
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmLoanAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private M_InterestBalance As Boolean
Private m_AccNo As Long
Private m_LoanId As Long
Private m_SchemeId As Long
Private m_SchemeName As String
Public m_rstSearchResults As Recordset
Private m_rstLoanTrans As Recordset
Private m_rstLoanMast As Recordset
Private m_rstScheme As Recordset
Private m_ModuleID  As wisModules
Private m_TotalInterest As Currency

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1
Private WithEvents m_frmLoanReport  As frmLoanView
Attribute m_frmLoanReport.VB_VarHelpID = -1

'Private m_Notes As New clsNotes
'Public m_AccHolder As New clsCustReg
'Private WithEvents SearchDialog As frmSearch


Const CTL_MARGIN = 15
Private Recursive As Boolean
Private m_accUpdatemode As Integer

Private m_PrevTab As Integer
Private m_InstalmentDetails As String
Private m_LoanAmount() As Currency
' If schemeloaded is True, the Save option
' will update the record, else it will insert.
Private m_SchemeLoaded As Boolean
Private m_MemberObj As clsMMAcc

' Declare events.
Public Event SetStatus(strMsg As String)



' Computes the total amount of instalments due from a loan
' account holder as on date.
Public Function ComputeInstalmentTotal() As Currency
On Error GoTo Err_Line

' Declare variables...
Dim LoanIssueDate As String
Dim InstMode As Byte
Dim AddMonth As Byte
Dim ToDay As String
Dim LastPaidDate As String
Dim NextInstalmentDate As String
'Dim PaidInstalments As Byte
Dim PaidInstTot As Currency
Dim LastInstalmentDate As String
Dim InstalmentNo As Byte
Dim unpaidinstalments  As Byte
Dim InstTot As Currency
Dim InstBalance As Currency
Dim Lret As Long

' If the recordsets 'm_rstloanmast', 'm_rstloantrans' and
' 'm_rstscheme' have not been set, then exit.
    If m_rstLoanMast Is Nothing Then
        Err.Raise "Loan Master recordset 'm_rstLoanMast' is empty."
        GoTo Exit_Line
    End If
    If m_rstScheme Is Nothing Then
        Err.Raise "Loan Master recordset 'm_rstScheme' is empty."
        GoTo Exit_Line
    End If
    If m_rstLoanTrans Is Nothing Then
        Err.Raise "Loan Master recordset 'm_rstLoanTrans' is empty."
        GoTo Exit_Line
    End If

'
' Compute how many instalments lie between the loan issue date and today.
'
'NOTE:  Since this function involves date arithmatic,  all the dates
'-------    are being handled in mm/dd/yyyy format only.
Dim LoanDueDate As String
    ' Get the loanissue date.
    LoanIssueDate = (m_rstLoanMast("IssueDate"))
    LoanDueDate = FormatField(m_rstLoanMast("LoanDueDate"))
    If Not DateValidate(txtRepayDate.Text, "/", True) Then GoTo Exit_Line
    ToDay = FormatDate(txtRepayDate.Text)
    
    ' Get the last instalment paid date.
    gDbTrans.SQLStmt = "SELECT TOP 1 transdate FROM loantrans " _
            & "WHERE loanid = " & FormatField(m_rstLoanMast("LoanID")) _
            & "And TransType = -1 ORDER BY transid DESC"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        MsgBox "Error getting loan transaction details.", vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    LastPaidDate = gDbTrans.Rst(0)

    ' instalment mode.
    InstMode = FormatField(m_rstLoanMast("InstalmentMode"))
    If InstMode < 6 And InstMode > 2 Then
        AddMonth = InstMode - 2
    ElseIf InstMode = 6 Then
        AddMonth = 6
    Else
        AddMonth = 12
    End If

    If InstMode = 0 Then
        ' No instalments defined for this loan account.
        ComputeInstalmentTotal = 0
        GoTo Exit_Line
    End If
    ' Begin a loop for calculating the no. of instalments...
    NextInstalmentDate = LoanIssueDate
    Do
        If InstMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, CDate(NextInstalmentDate))
        End If

        'If DateDiff("d", CDate(LastPaidDate), CDate(NextInstalmentDate)) < 0 Then
        '    ' Increment paid instalments count.
        '    PaidInstalments = PaidInstalments + 1
        'End If

        If DateDiff("d", CDate(ToDay), CDate(NextInstalmentDate)) > 0 Then Exit Do
        LastInstalmentDate = NextInstalmentDate
        InstalmentNo = InstalmentNo + 1
'''        If LastInstalmentDate > LoanDueDate Then Exit Do
    Loop
    
    ' Compute the amount for all instalments until this date.
    InstTot = InstalmentNo * Val(FormatField(m_rstLoanMast("InstalmentAmt")))
    
    ' Get the total amount paid by the loan holder
    ' towards instalment until this date.
    gDbTrans.SQLStmt = "SELECT SUM(amount) FROM loantrans WHERE loanid = " & Val(FormatField(m_rstLoanMast("LoanID"))) & " AND transtype= " & wDeposit & " And TransDate >= #" & LastPaidDate & "#"
    Lret = gDbTrans.SQLFetch
    If Lret < 0 Then
        MsgBox "Error getting loan transaction details.", vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    PaidInstTot = CCur(Val(FormatField(gDbTrans.Rst(0))))
    ' Instalment balance.
    'unpaidinstalments = InstalmentNo - PaidInstalments
    InstBalance = InstTot - PaidInstTot
    ' If the due instalments are already paid, exit.
    If Val(txtBalance.Caption) < InstBalance Then
        ComputeInstalmentTotal = Val(txtBalance.Caption)
    ElseIf InstBalance <= 0 Then
        ComputeInstalmentTotal = 0
    Else
        ComputeInstalmentTotal = InstBalance
    End If
    
Exit_Line:
    Exit Function

Err_Line:
    If Err Then
        MsgBox "ComputeInstalmentTotal: " & vbCrLf _
                & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 702) & vbCrLf _
                & Err.Description, vbCritical
    End If
''''Resume Next
    GoTo Exit_Line
End Function
' Calculates the penal interest for defaulted repayments.
Public Function ComputePenalInterest(TillIndianDate As String, Optional LoanIDNo As Long) As Currency
If Not DateValidate(TillIndianDate, "/", True) Then Exit Function
' Setup error handler
On Error GoTo Err_Line
' Variables...
Dim LoanID As Long
Dim LoanDate As String
Dim LoanDueDate As String
Dim LoanAmount As Currency
Dim TransType As wisTransactionTypes
Dim Lret As Long, i As Integer
Dim InstalmentMode As String
Dim PenaltyRate As Single
Dim PenalAmount As Currency
Dim AmountRepaid As Currency
Dim PaidInstalments As Integer
Dim DelayPeriod As Integer
Dim DefaultdDate As String
Dim InstalmentAmt As Currency
Dim BalanceAmt As Currency
Dim lastDate As String

'Load the Repayment date to a local variable
If LoanIDNo > 0 Then
    gDbTrans.SQLStmt = "SELECT * FROM loanMaster WHERE " _
        & "loanID = " & LoanIDNo '& " AND NOT loanclosed"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then GoTo Exit_Line
    ' Save the resultset for future references.
    Set m_rstLoanMast = gDbTrans.Rst.Clone
End If

If m_SchemeId <= 0 Then
    ' Get the scheme details.
    gDbTrans.SQLStmt = "SELECT * FROM loantypes WHERE " _
            & "schemeID = (SELECT schemeID FROM loanmaster " _
            & "WHERE loanid = " & LoanIDNo & ")"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    Set m_rstScheme = gDbTrans.Rst
    m_SchemeId = Val(FormatField(m_rstScheme("SchemeId")))
End If
'Get Loan Date & LoanDueDate
LoanDate = FormatField(m_rstLoanMast("Issuedate"))
LoanDueDate = FormatField(m_rstLoanMast("LoanDueDate"))

' Save the loanid to a local variable.
LoanID = Val(FormatField(m_rstLoanMast("loanid")))
If LoanID = 0 Then LoanID = LoanIDNo
' Get the rate of penalty.
PenaltyRate = Val(FormatField(m_rstLoanMast("PenalInterestRate")))

'Get total Loan Amount
TransType = wWithDraw
gDbTrans.SQLStmt = "Select Sum(Amount) as LoanAmount From LoanTrans Where LoanId = " & _
                                    LoanID & " And TransType = " & TransType
If gDbTrans.SQLFetch > 0 Then
    LoanAmount = CCur(FormatField(gDbTrans.Rst("LoanAmount")))
Else
    MsgBox "Error retrieving loan details from database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Get the balance amount for this loan.
gDbTrans.SQLStmt = "SELECT TOP 1 balance,TransDate FROM LoanTrans WHERE " _
        & " loanid = " & LoanID & " ORDER BY Transid DESC"
Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    MsgBox "Error retrieving loan details from database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
Else
    BalanceAmt = CCur(FormatField(gDbTrans.Rst("Balance")))
    If BalanceAmt = 0 Then GoTo Exit_Line
    lastDate = FormatField(gDbTrans.Rst("TransDate"))
End If

' The computation of penalty interest will differ depending upon
' whether there is any instalments defined or not.
InstalmentAmt = CCur(FormatField(m_rstLoanMast("instalmentamt")))

If InstalmentAmt = 0 Then
    ' Check if the due date for this loan expired.
    DelayPeriod = WisDateDiff(FormatField(m_rstLoanMast("Loanduedate")), TillIndianDate)
    If DelayPeriod <= 0 Then GoTo Exit_Line
    ' Compute the penal rate of interest.
    If WisDateDiff(FormatField(m_rstLoanMast("Loanduedate")), lastDate) > 0 Then
        DelayPeriod = WisDateDiff(lastDate, TillIndianDate)
    End If
    ComputePenalInterest = CCur(BalanceAmt * DelayPeriod * PenaltyRate / 36500)
    TillIndianDate = FormatDate(TillIndianDate)
    GoTo Exit_Line

End If


Dim NextInstalmentDate As String
Dim LastInstalmentDate As String
Dim LastInstalmentPaidDate As String
    
    'Get the instlment Mode ************
    InstalmentMode = FormatField(m_rstLoanMast("InstalmentMode"))
    
    'Get The AmountRepaid till today
    TransType = wDeposit
    gDbTrans.SQLStmt = "Select Sum(Amount) as RepaidAmount From LoanTrans Where LoanId = " & _
                            LoanID & " And TransType = " & TransType
    Lret = gDbTrans.SQLFetch
    If Lret < 0 Then
        MsgBox "Error retrieving loan details from database.", _
                vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    Else
        AmountRepaid = CCur(Val(FormatField(gDbTrans.Rst("RepaidAmount"))))
        BalanceAmt = LoanAmount - AmountRepaid
    End If
    '''************

    'Get The Last Instalment Paid Date & Loan Balance as on today ***********
    Dim AddMonth As Byte
    If InstalmentMode >= 3 And InstalmentMode < 6 Then
        AddMonth = InstalmentMode - 2
    ElseIf InstalmentMode = 6 Then
        AddMonth = 6
    Else
        AddMonth = 12
    End If
    '''****************
    'Calulate the no of instalments till last paid date
    '**** WARNING *****
    ' Im using dateadd function So i'm converting all date s into American format
    ' Later I'm changing it to Indian Format???
    
    TillIndianDate = FormatDate(TillIndianDate)
    LoanDueDate = FormatDate(LoanDueDate)
    
    TransType = wDeposit
    gDbTrans.SQLStmt = "Select Top 1 TransDate,Balance From LoanTrans Where LoanId = " & LoanID & _
                " And TransType = " & TransType & " ORDER BY transid DESC"
    If gDbTrans.SQLFetch > 0 Then
        LastInstalmentPaidDate = FormatDate(FormatField(gDbTrans.Rst("TransDate")))
    Else
        LastInstalmentPaidDate = FormatDate(LoanDate)
    End If
   
   'Shashi 5/7/2000
   'If He is Keeping Date instead of Balance Interest Then
   If Not M_InterestBalance Then
      'So We are taking Direct English Format
       LastInstalmentPaidDate = FormatField(m_rstLoanMast("Interestbalance"))
   End If
   'End 5/7/2000

    Dim InstalmentNo As Integer
    InstalmentNo = 0
    NextInstalmentDate = FormatDate(LoanDate)
    'Find the NoOf instalmets over till LastInstamentpaid date And
    'laso the LastInstament Date & LasInstalment Due Date w.r.t LstInstamentPaidDate
    Do
        If InstalmentMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstalmentMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, NextInstalmentDate)
        End If
        
        ' Check if the 'NextInstalmentDate' has crossed the last instalment date.  If so, exit.
''''        If DateDiff("d", CDate(NextInstalmentDate), CDate(LastInstalmentPaidDate)) <= 0 Then Exit Do
        If DateDiff("d", CDate(NextInstalmentDate), CDate(TillIndianDate)) <= 0 Then Exit Do
        LastInstalmentDate = NextInstalmentDate
        InstalmentNo = InstalmentNo + 1
    Loop
    If LastInstalmentDate = "" Then
        LastInstalmentDate = FormatDate(LoanDate)
    End If
    'Get the Instalment Last & Next Due Date
        'NextInstalmentDate = LastInstalmentPaidDate        ??? Why this statement  - wonders ravindra...!!!
    Dim TobePaidInstAmount As Currency
    Dim Count As Integer
    Dim InstalmentBalance As Currency
    Dim PenalInterestAmount
    'Dim ActualPaidAmount As Currency
    
    PenalAmount = 0: Count = 1
     'NextInstalmentDate = LastInstalmentDate
    'Find panel interest till the next Instlament Date if it is
    TobePaidInstAmount = InstalmentAmt * InstalmentNo
        InstalmentBalance = TobePaidInstAmount '''- AmountRepaid
        If CDate(LoanDate) > CDate(LastInstalmentPaidDate) Then
                LastInstalmentPaidDate = FormatDate(LoanDate)
        End If
        If InstalmentBalance > 0 Then
            'calculate then Penalty Amount
            PenalAmount = InstalmentBalance * PenaltyRate / 100 _
                    * DateDiff("d", CDate(LastInstalmentPaidDate), CDate(TillIndianDate)) / 365
            PenalInterestAmount = PenalAmount + PenalInterestAmount
        End If
        ' We considered the Penal interest till next instamnet date
        'There fore
        LastInstalmentDate = NextInstalmentDate
    'Find the no of skipped instalments
    
    Do
        ' IF the next instalment date is greater than current date, then exit
        If DateDiff("d", CDate(NextInstalmentDate), CDate(TillIndianDate)) <= 0 Then
            Exit Do
        End If
        If InstalmentMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstalmentMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, NextInstalmentDate)
        End If
        'After Getting A next InstalmentDate
        'Calulate the amount  to be repaid on the previous instalment date
        TobePaidInstAmount = InstalmentAmt * (InstalmentNo + Count)
        InstalmentBalance = TobePaidInstAmount - AmountRepaid
        
        'if Next Instalment Date is graeter than LoanDueDate then Calculate  penal interest on bAlnace Loan
        If DateDiff("D", LoanDueDate, NextInstalmentDate) > 0 Then
            If InstalmentBalance > 0 Then
                'calulate then Penalty Amount
                PenalAmount = InstalmentBalance * PenaltyRate / 100 _
                        * DateDiff("d", CDate(LastInstalmentDate), CDate(LoanDueDate)) / 365
                PenalInterestAmount = PenalAmount + PenalInterestAmount
            End If
            PenalAmount = BalanceAmt * PenaltyRate / 100 _
                    * DateDiff("d", CDate(LoanDueDate), CDate(TillIndianDate)) / 365
            PenalInterestAmount = PenalAmount + PenalInterestAmount
            Exit Do
        End If
        
        
        If InstalmentBalance > 0 Then
            'calulate then Penalty Amount
            PenalAmount = InstalmentBalance * PenaltyRate / 100 _
                    * DateDiff("d", CDate(LastInstalmentDate), CDate(TillIndianDate)) / 365
            PenalInterestAmount = PenalAmount + PenalInterestAmount
        End If

        LastInstalmentDate = NextInstalmentDate: Count = Count + 1
    Loop
    If PenalInterestAmount = 0 Then
        If InstalmentMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstalmentMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, lastDate)
        End If
        If CDate(lastDate) < CDate(TillIndianDate) Then
        PenalInterestAmount = TobePaidInstAmount * PenaltyRate / 100 _
                    * DateDiff("d", CDate(NextInstalmentDate), CDate(TillIndianDate)) / 365
        End If
    End If
        
    If PenalInterestAmount = 0 Then
        PenalInterestAmount = BalanceAmt * PenaltyRate / 100 _
                * WisDateDiff(FormatDate(LoanDueDate), FormatDate(TillIndianDate)) / 365
    End If
    
    '''Convert All The Date into Indian Format
    LastInstalmentDate = FormatDate(LastInstalmentDate)
    NextInstalmentDate = FormatDate(NextInstalmentDate)
    ComputePenalInterest = IIf(PenalInterestAmount < 0, 0, PenalInterestAmount)
    LoanDueDate = FormatDate(LoanDueDate)
    
''********  *****
Exit_Line:
   TillIndianDate = FormatDate(TillIndianDate)
   Exit Function
Err_Line:
    If Err Then
        MsgBox "ComputePenalInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
     
    End If
'    Resume
    GoTo Exit_Line
        
End Function

Public Function ComputeRegularInterest(TillIndianDate As String, Optional LoanID As Long) As Currency
' Setup error handler.
If Not DateValidate(TillIndianDate, "/", True) Then Exit Function
On Error GoTo Err_Line

' Define varaibles.
Dim Lret As Long
Dim InstalmentAmt As Currency
Dim InterestAmt As Currency
Dim InterestRate As Single
Dim IntDiff As Single
Dim ActualIntRate As Single
Dim lLoanID As Long
Dim Balance As Long
Dim lastDate As String
Dim Duration As Long

If LoanID > 0 Then
    gDbTrans.SQLStmt = "SELECT * FROM loanMaster WHERE " _
        & "loanID = " & LoanID '& " AND NOT loanclosed"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then GoTo Exit_Line
    ' Save the resultset for future references.
    Set m_rstLoanMast = gDbTrans.Rst.Clone
End If

If m_SchemeId <= 0 Then
    ' Get the scheme details.
    gDbTrans.SQLStmt = "SELECT * FROM loantypes WHERE " _
            & "schemeID = (SELECT schemeID FROM loanmaster " _
            & "WHERE loanid = " & LoanID & ")"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    Set m_rstScheme = gDbTrans.Rst
    m_SchemeId = Val(FormatField(m_rstScheme("SchemeId")))
End If

' Get the rate of interest.
InterestRate = Val(FormatField(m_rstLoanMast("InterestRate")))

' Store the loanid to local variable.
lLoanID = Val(FormatField(m_rstLoanMast("loanid")))
If lLoanID = 0 Then lLoanID = LoanID

' Get the balance amount or/and  date till interest paid for this loan account.
'There is Only one option whehter He has to Consider the Date till int paid
'Or the Interest Balance 'The Necessary changes has to be done on UI it shoul be Decide while installing the Package


gDbTrans.SQLStmt = "SELECT TOP 1 balance, transdate FROM loantrans " _
        & "WHERE loanid = " & lLoanID & " ORDER BY transid DESC"
Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    MsgBox "Error retrieving loan details from database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
    
    Balance = CCur(FormatField(gDbTrans.Rst("balance")))
    lastDate = FormatField(gDbTrans.Rst("transdate"))

If WisDateDiff(lastDate, FormatField(m_rstLoanMast("IssueDate"))) > 0 Then
    lastDate = FormatField(m_rstLoanMast("IssueDate"))
End If
'Shashi 5/7/2000
If Not M_InterestBalance Then
    lastDate = FormatDate(FormatField(m_rstLoanMast("Interestbalance")))
End If
'End 5/7/2000

' Get the duration in days from the last transaction date
' until this day.
Duration = WisDateDiff(lastDate, TillIndianDate)

If Duration > 0 Then
   ' Compute interest on this balance.
   Dim ClsInt As New clsInterest
   Dim IntRate As Single
   Dim IntAmount As Currency
   Dim NextDate As String
   Dim SchemeName As String

      'Get The SchemeName From LoanTypes
      gDbTrans.SQLStmt = "SELECT SchemeName From LoanTypes Where SchemeID = " & m_SchemeId
      If gDbTrans.SQLFetch < 1 Then
         MsgBox "Database ERROR ! ", vbCritical, wis_MESSAGE_TITLE
         Exit Function
      End If
       SchemeName = FormatField(gDbTrans.Rst("SchemeName"))

    'ComputeRegularInterest = Balance * Duration * InterestRate / 36500
     IntRate = ClsInt.InterestRate(wis_Loans, SchemeName, lastDate, TillIndianDate)
     NextDate = TillIndianDate
     'If the LoanHolder has got any concession in the interest, The same concession will be considered
     'for all Interest slabs
     ' So calculate Concession Given ' If concession is not given it will be zero
     IntDiff = IntRate - InterestRate
     'The below loop wil count the interest amount of the peiod wheenver interestrate has been changed
      Do
         NextDate = ClsInt.NextInterestDate
         If NextDate = "" Then
            ActualIntRate = IntRate - IntDiff
            Duration = WisDateDiff(lastDate, TillIndianDate)
            IntAmount = IntAmount + ((Duration / 365) * (ActualIntRate / 100) * Balance)
            Exit Do
         Else
            ActualIntRate = IntRate - IntDiff
            Duration = WisDateDiff(lastDate, TillIndianDate)
            IntAmount = IntAmount + ((Duration / 365) * (ActualIntRate / 100) * Balance)
            lastDate = DateAdd("d", 1, CDate(FormatDate(NextDate)))
            IntRate = ClsInt.NextInterestRate
         End If
     Loop
     ComputeRegularInterest = IntAmount
Else
    ComputeRegularInterest = 0
    
End If

Exit_Line:
    Exit Function

Err_Line:
    If Err Then
        MsgBox "ComputeRegularInterest: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
     End If
End Function
Private Function GetInstalmentNames(LoanID As Long) As String
' Rewrtie the Function  According to new code

Dim InstalmentName As String
Dim Rst As Recordset

Dim i As Integer
Dim frm As frmLoanName
On Error GoTo ErrLine

gDbTrans.SQLStmt = "Select * From Loancomponent where LoanId = " & LoanID
InstalmentName = ""
If gDbTrans.SQLFetch > 0 Then
    Set Rst = gDbTrans.Rst.Clone
    While Not Rst.EOF
        InstalmentName = InstalmentName & FormatField(Rst("CompName")) & ";"
        InstalmentName = InstalmentName & FormatField(Rst("CompAmount")) & ";"
        Rst.MoveNext
    Wend
    InstalmentName = Left(InstalmentName, Len(InstalmentName) - 1)
End If
m_InstalmentDetails = InstalmentName
GetInstalmentNames = InstalmentName
Exit Function
#If Old Then
Pos = InStr(1, InstalmentNames, ";")
InstalmentNames = Mid(InstalmentNames, Pos + 1)
Load frmLoanName
        For i = 1 To NoOfInst
                Pos = InStr(1, InstalmentNames, ";")
                InstName(i) = Left(InstalmentNames, Pos - 1)
                InstalmentNames = Mid(InstalmentNames, Pos + 1)
                
                Pos = InStr(1, InstName(i), ",")
                InstAmt(i) = CInt(Mid(InstName(i), Pos + 1))
                frmLoanName.txtLoanAmt(i - 1).Text = InstAmt(i)
                InstName(i) = Left(InstName(i), Pos - 1)
                frmLoanName.txtInstName(i - 1) = InstName(i)
                If Trim(InstalmentNames) = "" Then Exit For
        Next i
#End If
ErrLine:
        
        'MsgBox ""
End Function

Private Function InstalmentDue(LoanID As Long, TillIndianDate As String) As Boolean

Dim Rst As Recordset
Dim InstalmentMode As Byte
Dim NextInstalmentDate As String
Dim LastInstalmentDate As String
Dim LastInstalmentPaidDate As String
Dim LoanAmount As Currency
Dim AmountRepaid As Currency
Dim TransType As Byte
Dim Lret As Long
Dim BalanceAmt As Currency
Dim LoanDate As String
Dim InstalmentAmt As Currency
    'Get The LoanDetails
    gDbTrans.SQLStmt = "Select * from LoanMaster Where Loanid = " & LoanID
    If gDbTrans.SQLFetch <= 0 Then Exit Function
    
    Set Rst = gDbTrans.Rst.Clone
    InstalmentMode = Val(FormatField(Rst("InstalmentMode")))
    LoanAmount = Val(FormatField(Rst("LoanAmt")))
    'Get The AmountRepaid till today
    TransType = wDeposit
    gDbTrans.SQLStmt = "Select Sum(Amount) as RepaidAmount From LoanTrans Where LoanId = " & _
                            LoanID & " And TransDate <= #" & FormatDate(TillIndianDate) & "# And TransType = " & TransType
    
    Lret = gDbTrans.SQLFetch
    If Lret < 0 Then
        MsgBox "Error retrieving loan details from database.", _
                vbCritical, wis_MESSAGE_TITLE
        GoTo ExitLine
    Else
        AmountRepaid = CCur(FormatField(gDbTrans.Rst("RepaidAmount")))
        BalanceAmt = LoanAmount - AmountRepaid
    End If
    '''************

    'Get The Last Instalment Paid Date & Loan Balance as on today ***********
    Dim AddMonth As Byte
    If InstalmentMode >= 3 And InstalmentMode < 6 Then
        AddMonth = InstalmentMode - 2
    ElseIf InstalmentMode = 6 Then
        AddMonth = 6
    Else
        AddMonth = 12
    End If
    '''****************
    'Calulate the no of instalments till last paid date
    '**** WARNING *****
    ' Im using dateadd function So i'm converting all date s into American format
    ' Later I'm changing it to Indian Format???
    TillIndianDate = FormatDate(TillIndianDate)
    TransType = wDeposit
    gDbTrans.SQLStmt = "Select Top 1 TransDate,Balance From LoanTrans Where LoanId = " & LoanID & _
                "  And TransDate <= #" & FormatDate(TillIndianDate) & "# And TransType = " & TransType & " ORDER BY transid DESC"
    If gDbTrans.SQLFetch > 0 Then
        LastInstalmentPaidDate = FormatDate(FormatField(gDbTrans.Rst("TransDate")))
    Else
        LastInstalmentPaidDate = FormatDate(LoanDate)
    End If
   
    Dim InstalmentNo As Integer
    InstalmentNo = 0
    NextInstalmentDate = FormatDate(LoanDate)
    'Find the NoOf instalmets over till LastInstamentpaid date And
    'laso the LastInstament Date & LasInstalment Due Date w.r.t LstInstamentPaidDate
    Do
        If InstalmentMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstalmentMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, NextInstalmentDate)
        End If
        
        ' Check if the 'NextInstalmentDate' has crossed the last instalment date.  If so, exit.
        If DateDiff("d", CDate(NextInstalmentDate), CDate(LastInstalmentPaidDate)) <= 0 Then Exit Do
        LastInstalmentDate = NextInstalmentDate
        InstalmentNo = InstalmentNo + 1
    Loop
    
    If LastInstalmentDate = "" Then
        LastInstalmentDate = FormatDate(LoanDate)
    End If
    
    Dim TobePaidInstAmount As Currency
    Dim Count As Integer
    Dim InstalmentBalance As Currency
    Dim PenalInterestAmount
    'Dim ActualPaidAmount As Currency
    
     Count = 1
     'NextInstalmentDate = LastInstalmentDate
    'Find panel interest till the next Instlament Date if it is
    TobePaidInstAmount = InstalmentAmt * InstalmentNo
        InstalmentBalance = TobePaidInstAmount - AmountRepaid
        If InstalmentBalance > 0 Then
            InstalmentDue = True
            Exit Function
        End If
        ' We considered the Penal interest till next instamnet date
        'There fore
        LastInstalmentDate = NextInstalmentDate
    'Find the no of skipped instalments
    Do
        ' IF the next instalment date is greater than current date, then exit
        If DateDiff("d", CDate(NextInstalmentDate), CDate(TillIndianDate)) <= 0 Then
            Exit Do
        End If
        If InstalmentMode = 1 Then ' if Weekly
            NextInstalmentDate = DateAdd("d", 7, CDate(NextInstalmentDate))
        ElseIf InstalmentMode = 2 Then 'if Fortnigthly
            NextInstalmentDate = DateAdd("d", 15, CDate(NextInstalmentDate))
        Else
            NextInstalmentDate = DateAdd("m", AddMonth, NextInstalmentDate)
        End If
        
        'After Getting A next InstalmentDate
        'Calulate the amount  to be repaid on the previous instalment date
        TobePaidInstAmount = InstalmentAmt * (InstalmentNo + Count)
        InstalmentBalance = TobePaidInstAmount - AmountRepaid
        If InstalmentBalance > 0 Then
            InstalmentDue = True
            Exit Function
        End If

        LastInstalmentDate = NextInstalmentDate: Count = Count + 1
    Loop
    
ExitLine:
    '''Convert All The Date into Indian Format
    LastInstalmentDate = FormatDate(LastInstalmentDate)
    NextInstalmentDate = FormatDate(NextInstalmentDate)
    TillIndianDate = FormatDate(TillIndianDate)
    ' Add the previous interest balance, if any.
    'ComputePenalInterest =  ComputePenalInterest

End Function

Private Function LoanLoadTypes() As Boolean
On Error GoTo Err_Line

' Declare variables for this routine.
Dim Lret As Long

' Raise an event to indicate that loan types are geing loaded.
RaiseEvent SetStatus("Loading loan types...")

' Query the loan types.
gDbTrans.SQLStmt = "SELECT * FROM loantypes"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then Exit Function

' Load them to the combobox.
With cmbLoanType
    .Clear
    .AddItem ""
    Do While Not gDbTrans.Rst.EOF
        .AddItem FormatField(gDbTrans.Rst("SchemeName"))
        .ItemData(.NewIndex) = Val(FormatField(gDbTrans.Rst("schemeID")))
        gDbTrans.Rst.MoveNext
    Loop
End With

Err_Line:
    RaiseEvent SetStatus("")
    If Err Then
        MsgBox "LoanLoadTypes: " & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 705) & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
End Function
Private Function LoanRepay() As Boolean
On Error GoTo Err_Line

' Variables used in this procedure...
' -----------------------------------
Dim lLoanID As Long
Dim newTransID As Long
Dim inTransaction As Boolean
Dim NewBalance As Currency
Dim IntAmount  As Currency
Dim PenalIntAmount As Currency
Dim PrincAmount  As Currency
Dim IntBalance As Currency
Dim PayAmount As Currency
Dim RegInt As Currency
Dim PenalInt As Currency
Dim IntPaidDate As String
Dim TransType As wisTransactionTypes
' -----------------------------------

' Check if a valid amount is entered.
If Not CurrencyValidate(txtRepayAmt.Text, False) Then
    'MsgBox "Enter valid amount.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 506), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRepayAmt
    GoTo Exit_Line
End If

Dim RepayDate As String
RepayDate = txtRepayDate.Text

'If Package is Considering The Interest Paid date Then
   If Not M_InterestBalance Then  'Check the Validate
      If Not DateValidate(txtIntBalance.Text, "/", True) Then
        MsgBox LoadResString(gLangOffSet + 501), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtIntBalance
        Exit Function
      End If
   
   ' Check if the repayment date is later than today's date.
         If WisDateDiff(FormatDate(gStrDate), txtIntBalance.Text) > 0 Then
             'MsgBox "Repayment date cannot be greater than today's date", _
                     vbExclamation, wis_MESSAGE_TITLE
             MsgBox "Repayment date cannot be greater than today's date", _
                     vbExclamation, wis_MESSAGE_TITLE
             GoTo Exit_Line
         End If

   End If

' Get the loanID.
lLoanID = Val(Mid(tabLOans.SelectedItem.Key, 4))

IntAmount = CCur(Val(txtRegInterest.Text))
PenalIntAmount = CCur(Val(txtPenalInterest.Text))
PayAmount = CCur(txtRepayAmt.Text)
PrincAmount = PayAmount - IntAmount - PenalIntAmount - CCur(Val(txtMIsc.Text))

If PayAmount < CCur(Val(txtIntBalance.Text)) And M_InterestBalance Then
    'If MsgBox("The amount he is paying is less than his Previous Interest Balance " & _
        vbCrLf & " Do you want to continue with transaction", vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    If MsgBox(LoadResString(gLangOffSet + 769) & _
        vbCrLf & LoadResString(gLangOffSet + 541), vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
   
   ' As he is not pay interest till last paid date also so cofirm once again
    'If MsgBox("The amount he is paying is less than his Previous Interest Balance " & _
        vbCrLf & " Do you want to continue with transaction", vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    If MsgBox(LoadResString(gLangOffSet + 769) & _
        vbCrLf & LoadResString(gLangOffSet + 541), vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function
End If

'Calculate the RegInt & Penal Int on the Specified date
    RegInt = ComputeRegularInterest(txtRepayDate.Text) + Val(txtIntBalance.Text)
    RegInt = RegInt \ 1
    PenalInt = ComputePenalInterest(txtRepayDate.Text)
    PenalInt = PenalInt \ 1

If (Val(txtRegInterest.Text) + Val(txtPenalInterest.Text)) - 1 > (RegInt + PenalInt) Then
'If IntAmount > (RegInt + PenalInt) Then
   'If MsgBox("The amount he is paying is more than his Interest till date " & _
        vbCrLf & " Do you want to keep this Extra amount as advance interest", vbYesNo + vbQuestion + _
        vbDefaultButton1, wis_MESSAGE_TITLE) = vbNo Then
   
    If MsgBox(LoadResString(gLangOffSet + 772) & _
        vbCrLf & LoadResString(gLangOffSet + 786), vbYesNo + vbQuestion + _
        vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            IntBalance = 0
    Else
        IntBalance = (RegInt + PenalInt) - IntAmount
    End If
ElseIf IntAmount < 0 Then
    IntBalance = IntAmount
End If

If PrincAmount < 0 Then
    IntBalance = IntAmount - PayAmount
    'Upadate the Varaible as paying interest
    IntAmount = PayAmount
    PrincAmount = 0
End If

' Get a new transactionID.
gDbTrans.SQLStmt = "SELECT MAX(TransID) FROM LoanTrans " _
        & "WHERE loanid = " & lLoanID
If gDbTrans.SQLFetch <= 0 Then GoTo Exit_Line
newTransID = Val(FormatField(gDbTrans.Rst(0)))

' Begin the transaction
If Not gDbTrans.BeginTrans Then GoTo Exit_Line
inTransaction = True
    
''Very First update the Penal Interest amount (If Any)
    TransType = wCharges
    ' Get the balance.
    'He is paying Only interest need not to change the bALANCE
    NewBalance = CCur(txtBalance.Caption)

    ' Update the Penal interest amount.
        newTransID = newTransID + 1
        gDbTrans.SQLStmt = "INSERT INTO loantrans (LoanID, TransID, " _
                & "TransType, Amount, transDate, Balance, Particulars ) " _
                & "VALUES (" & lLoanID & ", " & newTransID & ", " _
                & TransType & ", " & PenalIntAmount & ", #" _
                & FormatDate(txtRepayDate.Text) & "#, " _
                & NewBalance & ",'" & LoadResString(gLangOffSet + 345) & "')"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        m_TotalInterest = m_TotalInterest - PenalIntAmount

    ' Update the interest amount.
    If IntAmount < 0 And NewBalance - PrincAmount > 0 Then
        IntAmount = 0
    ElseIf NewBalance - PrincAmount <= 0 Then
    'Repay the interest if he  has paid extra interst earlier
        If IntAmount < 0 Then
            TransType = wInterest
            IntAmount = Abs(IntAmount)
           'MsgBox "He has paid extra interest in the prevoius Payment " & vbCrLf & "Return the extra interest Rs." & IntAmount
           MsgBox LoadResString(gLangOffSet + 792) & vbCrLf & LoadResString(gLangOffSet + 792) & " " & LoadResString(gLangOffSet + 792) & " " & IntAmount
       End If
    End If
    If PenalIntAmount > 0 Then ' Increment the transaction ID.
    End If
    
        newTransID = newTransID + 1
        gDbTrans.SQLStmt = "INSERT INTO loantrans (LoanID, TransID, " _
                & "TransType, Amount, transDate, Balance, Particulars ) " _
                & "VALUES (" & lLoanID & ", " & newTransID & ", " _
                & TransType & ", " & IntAmount & ", #" _
                & FormatDate(txtRepayDate.Text) & "#, " _
                & NewBalance & ", '" & LoadResString(gLangOffSet + 344) & "')"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        m_TotalInterest = m_TotalInterest - IntAmount
''Next Pay the Principle amount
    TransType = wDeposit
    ' Increment the transaction ID.
    newTransID = newTransID + 1

    ' Update the principal amount.
    NewBalance = NewBalance - PrincAmount

        gDbTrans.SQLStmt = "INSERT INTO loanTrans (LoanID, TransID, " _
            & "TransType, Amount, transDate, Balance, Particulars) " _
            & "VALUES (" & lLoanID & ", " & newTransID & ", " _
            & TransType & ", " & PrincAmount & ", #" _
            & FormatDate(txtRepayDate.Text) & "#, " _
            & NewBalance & ", '" & LoadResString(gLangOffSet + 343) & "')"
        
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
    'End If
    
    
    ' Update the loan master with interest balance.
    ' In this case, the total repaid amt is less than the interest payable.
    ' Therefore, put this difference amt, to loan master table.
    If m_TotalInterest >= 0 Then
        gDbTrans.SQLStmt = "UPDATE loanmaster SET [InterestBalance] = " _
                & AddQuotes(CStr(m_TotalInterest), True) & " WHERE loanid = " & lLoanID
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If

    ' If the balance amount is fully paidup, then set the flag "LoanClosed" to True.
    If NewBalance <= 0 Then
        gDbTrans.SQLStmt = "UPDATE loanmaster SET loanclosed = TRUE Where Loanid =  " & lLoanID
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    End If


    ' Commit the transaction.
    If Not gDbTrans.CommitTrans Then GoTo Exit_Line
    inTransaction = False
    
'Now UpDate the  Miscalleneous Amount
Dim BankClass As New clsBankAcc
If Val(txtMIsc.Text) > 0 Then
    ' while undoing transaction we have to undo this profit also
    ' so for the identification I'm sendig Loan Id & TransID
    Call BankClass.UPDateMiscProfit(Val(txtMIsc.Text), txtRepayDate.Text, "LoanRepay " & lLoanID & "-" & newTransID)
End If
Set BankClass = Nothing

    LoanRepay = True
    'MsgBox "Loan repayment accepted.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 706), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    Exit Function

Err_Line:
    If Err Then
        MsgBox "AcceptPayment: " & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 707) & Err.Description, _
            vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line
End Function
Private Sub ArrangeLoanIssuePropSheet()

Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean
lblLoanIssueDesc.BorderStyle = 0
lblLoanIssueHeading.BorderStyle = 0
fraLoanAccounts.Caption = ""

' Arrange the Slider panel.
With picLoanIssueSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCountLoanIssue
    .Height = txtLoanIssue(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    ' If the height is greater than viewport height,
    ' the scrollbar needs to be displayed.  So,
    ' reduce the width accordingly.
    If .Height > picLoanIssueViewPort.ScaleHeight Then
        NeedsScrollbar = True
        .Width = picLoanIssueViewPort.ScaleWidth - _
                vscLoanIssue.Width
    Else
        .Width = picLoanIssueViewPort.ScaleWidth
    End If

End With

' Set/Reset the properties of scrollbar.
With vscLoanIssue
    .Height = picLoanIssueViewPort.ScaleHeight
    .Min = 0
    .Max = picLoanIssueSlider.Height - picLoanIssueViewPort.ScaleHeight
    If .Max < 0 Then .Max = 0
    .SmallChange = txtLoanIssue(0).Height
    .LargeChange = picLoanIssueViewPort.ScaleHeight / 2
End With

' Adjust the text controls on this panel.
Dim i As Integer
For i = 0 To txtLoanIssue.Count - 1
    txtLoanIssue(i).Width = picLoanIssueSlider.ScaleWidth _
            - lblLoanIssue(i).Width - CTL_MARGIN
Next


If NeedsScrollbar Then
    vscLoanIssue.Visible = True
End If

For i = 0 To txtLoanIssue.Count - 1
    txtLoanIssue(i).Width = picLoanIssueSlider.ScaleWidth - _
        (lblLoanIssue(i).Left + lblLoanIssue(i).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For i = 0 To cmbLoanIssue.Count - 1
    cmbLoanIssue(i).Width = txtLoanIssue(i).Width
Next
For i = 0 To cmdLoanIssue.Count - 1
    cmdLoanIssue(i).Left = txtLoanIssue(i).Left _
        + txtLoanIssue(i).Width - cmdLoanIssue(i).Width
Next

'' Draw lines for the remaining portions of the viewport.
'With picLoanIssueViewPort
'    .CurrentX =


End Sub
Private Sub ArrangeLoanSchemePropSheet()

Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean
lblLoanSchemeHeading.BorderStyle = 0
lblLoanSchemeDesc.BorderStyle = 0
fraSchemes.Caption = ""

' Arrange the Slider panel.
With picLoanSchemeSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCountLoanScheme
    .Height = txtLoanScheme(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    ' If the height is greater than viewport height,
    ' the scrollbar needs to be displayed.  So,
    ' reduce the width accordingly.
    If .Height > picLoanSchemeViewPort.ScaleHeight Then
        NeedsScrollbar = True
        .Width = picLoanSchemeViewPort.ScaleWidth - _
                vscLoanScheme.Width
    Else
        .Width = picLoanSchemeViewPort.ScaleWidth
    End If

End With

' Set/Reset the properties of scrollbar.
With vscLoanScheme
    .Height = picLoanSchemeViewPort.ScaleHeight
    .Min = 0
    .Max = picLoanSchemeSlider.Height - picLoanSchemeViewPort.ScaleHeight
    If .Max < 0 Then .Max = 0
    .SmallChange = txtLoanScheme(0).Height
    .LargeChange = picLoanSchemeViewPort.ScaleHeight / 2
End With

' Adjust the text controls on this panel.
Dim i As Integer
For i = 0 To txtLoanScheme.Count - 1
    txtLoanScheme(i).Width = picLoanSchemeSlider.ScaleWidth _
            - lblLoanScheme(i).Width - CTL_MARGIN
Next


If NeedsScrollbar Then
    vscLoanScheme.Visible = True
End If

For i = 0 To txtLoanScheme.Count - 1
    txtLoanScheme(i).Width = picLoanSchemeSlider.ScaleWidth - _
        (lblLoanScheme(i).Left + lblLoanScheme(i).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For i = 0 To cmbLoanScheme.Count - 1
    cmbLoanScheme(i).Width = txtLoanScheme(i).Width
Next
For i = 0 To cmdLoanScheme.Count - 1
    cmdLoanScheme(i).Left = txtLoanScheme(i).Left _
        + txtLoanScheme(i).Width - cmdLoanScheme(i).Width
Next
End Sub
' Returns the value selected category
Private Function GetLoanCategory() As wisLoanCategories
Dim strCategory As String
strCategory = PropSchemeGetVal("Category")
If StrComp(strCategory, LoadResString(gLangOffSet + 440), vbTextCompare) = 0 Then
    GetLoanCategory = wisAgriculural
Else
    GetLoanCategory = wisNonAgriculural
End If

End Function

Private Function GetLoanTerm() As wisLoanTerm
    Dim strLoanTerm As String
    strLoanTerm = PropSchemeGetVal("TermType")
    If StrComp(strLoanTerm, LoadResString(gLangOffSet + 222), vbTextCompare) = 0 Then
        GetLoanTerm = wisShortTerm
    ElseIf StrComp(strLoanTerm, LoadResString(gLangOffSet + 223), vbTextCompare) = 0 Then
        GetLoanTerm = wisMidTerm
    Else
        GetLoanTerm = wisLongTerm
    End If
    
End Function

Private Function ComputeInstalmentAmount() As Currency
'shashi
Dim LoanAmt As Currency
Dim LoanDueDate As String
Dim InstMode As Integer
Static LoanInterest As Currency
Dim strSchemeName As String
Dim Lret As Long, instCount As Integer
Dim LoanDuration As Integer
Dim Interest As Currency
Dim LoanIssueDate As String
Dim InstalmentAmt As Currency
On Error Resume Next
' If the field name is "LoanAmt", then compute the instalment amount.
' Get the Scheme name.
strSchemeName = PropIssueGetVal("LoanName")
If strSchemeName = "" Then Exit Function
' Query the interest.
If LoanInterest = 0 Then
    gDbTrans.SQLStmt = "SELECT InterestRate FROM loantypes " _
            & "WHERE schemename = '" & strSchemeName & "'"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then Exit Function
    LoanInterest = Val(FormatField(gDbTrans.Rst(0)))
End If

' Get the loan amt.
LoanAmt = Val(PropIssueGetVal("LoanAmount"))
' Get the loanduedate.
LoanDueDate = PropIssueGetVal("LoanDueDate")
If LoanDueDate = "" Then Exit Function
' Get the instalment mode.
InstMode = GetInstalmentMode
' Get the loan issue date.
LoanIssueDate = FormatDate(PropIssueGetVal("IssueDate"))
If LoanIssueDate = "" Then
    LoanIssueDate = FormatDate(gStrDate)
End If
' Compute the duration of loan.
LoanDuration = WisDateDiff(FormatDate(gStrDate), LoanDueDate)
' Compute the interest.
Interest = LoanAmt * (LoanInterest / 100) * (LoanDuration / 365)
' Compute the no. of instalments.
instCount = InstalmentCount(LoanIssueDate, LoanDueDate, InstMode)
' Compute the instalment amount.
If instCount = 0 Then
    InstalmentAmt = FormatCurrency(LoanAmt + Interest)
Else
    InstalmentAmt = FormatCurrency((LoanAmt + Interest) / instCount)
End If
On Error GoTo 0

ComputeInstalmentAmount = IIf(InstalmentAmt > 0, InstalmentAmt, 0)
End Function

' Computes the penalty interest for a delayed repayment.
Public Function ZZZ_ComputePenalInterest_OLD(lLoanID As Long) As Currency
' Setup error handler.
On Error GoTo Err_Line

' Define variables for this procedure.
Dim Lret As Long
Dim nInstalmentMode As Integer
Dim InstalmentAmt As Currency
Dim LoanDueDate As String
Dim LoanIssueDate As String
Dim LastPaidDate As String
Dim DelayPeriod As Integer
Dim nInstalments As Integer
Dim MaxInstalments As Integer
Dim nElapsedInstalments As Integer
Dim AmtTobePaid As Currency
Dim TotalRepaidAmt As Currency
Dim AmountOverdue As Currency
Dim TransType As wisTransactionTypes

' Find out, if any instalments are defined for this loan account.
gDbTrans.SQLStmt = "SELECT instalmentmode, instalmentamt, " _
    & "issuedate, loanduedate FROM loanmaster WHERE loanid = " & lLoanID
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
nInstalmentMode = Val(FormatField(gDbTrans.Rst("instalmentmode")))
InstalmentAmt = Val(FormatField(gDbTrans.Rst("instalmentamt")))
LoanDueDate = FormatField(gDbTrans.Rst("loanduedate"))
LoanIssueDate = FormatField(gDbTrans.Rst("issuedate"))

' Get the last date on which repayment was done.
gDbTrans.SQLStmt = "SELECT TOP 1 transdate FROM loantrans WHERE " _
        & "loanid = " & lLoanID & " ORDER BY transid DESC"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    'MsgBox "Error in database!", vbCritical, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 601), vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
LastPaidDate = FormatField(gDbTrans.Rst(0))

' Compute the delay period.
If nInstalmentMode = 0 Then     ' No instalment mode.
    ' Check if the loanduedate has elapsed.
    DelayPeriod = WisDateDiff(FormatDate(gStrDate), LoanDueDate)
Else
    ' Compute the delay period.
    DelayPeriod = WisDateDiff(FormatDate(gStrDate), LastPaidDate)
    ' Amount to be paid, and no. of instalments.
    Select Case nInstalmentMode
        Case 1  ' Weekly
            '
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 7
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
            
        Case 2  ' Fortnightly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 15
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
        Case 3  ' Monthly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 15
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
        Case 4  ' Bi-monthly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 60
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
        Case 5  ' Quarterly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 90
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
        Case 6  ' Half-yearly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 183
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
        Case 7  ' Yearly
            ' Compute the no. of instalments from the loan issue date, till today.
            nInstalments = WisDateDiff(LoanIssueDate, FormatDate(gStrDate)) \ 365
            ' Compute the amount supposed to be paid until now.
            AmtTobePaid = nInstalments * InstalmentAmt
    End Select
End If

' Get the total amount repaid.
TransType = wDeposit
gDbTrans.SQLStmt = "SELECT SUM (amount) FROM loantrans WHERE " _
        & "loanid = " & lLoanID & " AND transtype = " & TransType
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
TotalRepaidAmt = Val(FormatField(gDbTrans.Rst(0)))

' If the repayableamount is greater than the amount repaid,
' compute the penalty interest on this amount.
AmountOverdue = AmtTobePaid - TotalRepaidAmt

' Get the penal interest rate for this loan type.
gDbTrans.SQLStmt = "SELECT penalinterestrate FROM loantypes " _
        & "WHERE schemeid = (SELECT schemeid FROM loanmaster " _
        & "WHERE loanid = " & lLoanID & ")"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Compute the penal intererst amount.
'ComputePenalInterest = FormatCurrency(AmountOverdue * FormatField(gDBTrans.Rst(0)) / 36500)

Exit_Line:
    Exit Function
Err_Line:
    If Err Then
        MsgBox "ComputePenalInterest: " & vbCrLf _
            '& Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 708) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line
End Function

' Fills in the details of the Account holder to the form.
Private Sub FillAcHolderDetails()
#If COMMENTED Then
On Error GoTo Err_Line
Dim i As Integer
Dim strField As String

' Extract the nominee info from the field.
Dim NomineeInfo() As String
If Not IsNull(gDbTrans.Rst("Nominee")) Then
    GetStringArray FormatField(gDbTrans.Rst("Nominee")), NomineeInfo(), ";"
End If
If UBound(NomineeInfo) = 0 Then
    ReDim NomineeInfo(2)
    NomineeInfo(0) = " "
    NomineeInfo(1) = " "
    NomineeInfo(2) = " "
End If

' fill in the details of ac-holder.
For i = 0 To txtPrompt.Count - 1
    ' Read the bound field of this control.
    strField = ExtractToken(txtPrompt(i).Tag, "DataSource")
    If strField <> "" Then
        With txtData(i)
            Select Case UCase$(strField)
                Case "ACCID"
                    .Text = gDbTrans.Rst("Accid")
                Case "ACNAME"
                    .Text = m_AccHolder.FullName
                Case "NOMINEE_NAME"
                    .Text = NomineeInfo(0)
                Case "NOMINEE_AGE"
                    .Text = NomineeInfo(1)
                Case "NOMINEE_RELATION"
                    .Text = NomineeInfo(2)
                Case "JOINTHOLDER"
                    .Text = gDbTrans.Rst("JointHolder")
                Case "INTRODUCERID"
                    .Text = IIf(gDbTrans.Rst("Introduced") = 0, "", gDbTrans.Rst("Introduced"))
                Case "INTRODUCERNAME"

                Case "LEDGERNO"
                    .Text = gDbTrans.Rst("LedgerNo")
                Case "FOLIONO"
                    .Text = gDbTrans.Rst("FolioNO")
                Case "CREATEDATE"
                    .Text = FormatField(gDbTrans.Rst("CreateDate"))
            End Select
        End With
    End If
Next


Err_Line:
    If Err.Number = 9 Then  'Subscript out of range.
        Resume Next
    ElseIf Err Then
        MsgBox "FillAcHolderDetails: " & vbCrLf _
                & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 709) & vbCrLf _
                & Err.Description, vbCritical
    End If

#End If
End Sub
'
' Fills the details of members to report window.
' Instead of standard fillview function, this is
' written to handle the specific requirements of fill members...
'
Private Function FillMembers(ListViewCtl As ListView, rs As Recordset, Optional AutoWidth As Boolean) As Boolean
On Error GoTo fillmembers_error
Const FIELD_MARGIN = 1.5

' Check if there are any records in the recordset.
rs.MoveLast
rs.MoveFirst
If rs.RecordCount = 0 Then
    FillMembers = True
    GoTo Exit_Line
End If

Dim i As Integer
Dim itmX As ListItem
With ListViewCtl
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    For i = 1 To rs.fields.Count - 1
        .ColumnHeaders.Add , rs.fields(i).Name, rs.fields(i).Name, _
                    ListViewCtl.Parent.TextWidth(rs.fields(i).Name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If i > 0 Then
            If rs.fields(i).Type = dbNumeric Or _
                    rs.fields(i).Type = dbInteger Or _
                    rs.fields(i).Type = dbLong Or _
                    rs.fields(i).Type = dbDouble Or _
                    rs.fields(i).Type = dbCurrency Then
                .ColumnHeaders(i).Alignment = lvwColumnRight
            End If
        End If
    Next

    ' Begin a loop for processing rows.
    Do While Not rs.EOF
        ' Add the details.
        Set itmX = .ListItems.Add(, , rs.fields(1))
        ' If the 'Autowidth' property is enabled,
        ' then check if the width needs to be expanded.
        If AutoWidth Then
            If .ColumnHeaders(1).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.fields(1))) Then
                .ColumnHeaders(1).Width = _
                    .Parent.TextWidth(FormatField(rs.fields(1))) * FIELD_MARGIN
            End If
        End If
        ' Add sub-items.
        For i = 2 To rs.fields.Count - 1
            ' If the field name is Category
            If StrComp(rs.fields(i).Name, "Category", vbTextCompare) = 0 Then
                itmX.SubItems(i - 1) = IIf((FormatField(rs.fields(i)) = 1), _
                        LoadResString(gLangOffSet + 440), LoadResString(gLangOffSet + 441))
            Else
                itmX.SubItems(i - 1) = FormatField(rs.fields(i))
            End If

            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then
                If .ColumnHeaders(i).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(itmX.SubItems(i - 1)) Then
                    .ColumnHeaders(i).Width = _
                        .Parent.TextWidth(itmX.SubItems(i - 1)) * FIELD_MARGIN
                End If
            End If
        Next
        rs.MoveNext
    Loop
End With
FillMembers = True

Exit_Line:
ListViewCtl.Visible = True
ListViewCtl.view = lvwReport
Exit Function

fillmembers_error:
    If Err.Number = 3265 Then
        On Error Resume Next
        Resume
    ElseIf Err Then
        MsgBox "FillMembers: The following error occurred." _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 710) _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function
Public Function HasOverdueLoans(lMemberID As Long) As Boolean
On Error GoTo Err_Line

' Variables...
Dim Lret As Long
Dim loansRS As Recordset
Dim Balance As Currency
Dim dueDate As Date

' Check how many loans this member has taken.
gDbTrans.SQLStmt = "SELECT loanID FROM loanMaster WHERE memberID = " & lMemberID
Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    ' Fatal error. Warn and stop execution.
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    ' No loans taken previously. Return False.
    GoTo Exit_Line
End If
' Store the loans recordset.
Set loansRS = gDbTrans.Rst.Clone

' Get the balances remaining for each of the above loans...
Do While Not loansRS.EOF
    ' Get the balance.
    gDbTrans.SQLStmt = "SELECT TOP 1 * FROM loanTrans WHERE " _
        & "loanID = " & FormatField(loansRS(0)) & " ORDER BY transID Desc"
    Lret = gDbTrans.SQLFetch
    If Lret < 0 Then
        MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
        End
    ElseIf Lret = 0 Then
        GoTo Exit_Line
    End If
    ' If the balance is >0, check if the loan due date has elapsed.
    Balance = Val(FormatField(gDbTrans.Rst("Balance")))
    If Balance > 0 Then
        ' Check the due date.
        gDbTrans.SQLStmt = "SELECT * FROM LoanMaster WHERE " _
                & "loanID = " & FormatField(loansRS(0))
        Lret = gDbTrans.SQLFetch
        If Lret < 0 Then
            MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
            GoTo Exit_Line
        End If
        dueDate = FormatField(gDbTrans.Rst("LoanDueDate"))
        If dueDate < FormatDate(gStrDate) Then
            HasOverdueLoans = True
            Exit Do
        End If
    End If
    ' Move to next record.
    loansRS.MoveNext
Loop

Exit_Line:
    Exit Function
Err_Line:
    If Err Then
        MsgBox "HasOverdues: " & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 711) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Function
' Computes the no. of instalments,. given the start date
Private Function InstalmentCount(StartDateIndian As String, LoanDueDateIndian As String, InstalmentMode As Integer) As Integer
Dim Duration As Integer

' Get the date difference between the two dates.
Duration = WisDateDiff(StartDateIndian, LoanDueDateIndian)

Select Case InstalmentMode
    Case 1  ' Weekly
        InstalmentCount = Duration / 7
        If Duration Mod 7 > 3 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 2  ' Fortnightly
        InstalmentCount = Duration / 15
        If Duration Mod 15 > 7 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 3  ' Monthly
        InstalmentCount = Duration / 30
        If Duration Mod 30 > 15 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 4  ' Bi-monthly
        InstalmentCount = Duration / 60
        If Duration Mod 60 > 30 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 5  ' Quarterly
        InstalmentCount = Duration / 90
        If Duration Mod 90 > 45 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 6  ' Half-yearly
        InstalmentCount = Duration / 183
        If Duration Mod 183 > 90 Then
            InstalmentCount = InstalmentCount + 1
        End If
    Case 7  ' Annually
        InstalmentCount = Duration / 365
        If Duration Mod 365 > 183 Then
            InstalmentCount = InstalmentCount + 1
        End If
End Select

End Function
' Gets the instalment mode specified by the user.
Private Function GetInstalmentMode() As Integer
Dim strInstMode As String
strInstMode = PropIssueGetVal("InstalmentMode")
'
On Error GoTo ErrLine
Select Case strInstMode
    Case ""
        GetInstalmentMode = 0
    Case LoadResString(gLangOffSet + 461) '"WEEKLY"
        GetInstalmentMode = 1
    Case LoadResString(gLangOffSet + 462) '"FORTNIGHTLY"
        GetInstalmentMode = 2
    Case LoadResString(gLangOffSet + 463) '"MONTHLY"
        GetInstalmentMode = 3
    Case LoadResString(gLangOffSet + 464) '"BI-MONTHLY"
        GetInstalmentMode = 4
    Case LoadResString(gLangOffSet + 465) '"QUARTERLY"
        GetInstalmentMode = 5
    Case LoadResString(gLangOffSet + 466) '"HALF-YEARLY"
        GetInstalmentMode = 6
    Case LoadResString(gLangOffSet + 467) '"ANNUALLY"
        GetInstalmentMode = 7
End Select
Exit Function

ErrLine:
MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function GetInstalmentModeText(InstalmentMode As Integer) As String
Dim strInstMode As String
strInstMode = PropIssueGetVal("InstalmentMode")
'
On Error GoTo ErrLine
Select Case InstalmentMode
    Case 1, 2, 3, 4, 5, 6, 7
        GetInstalmentModeText = LoadResString(gLangOffSet + 460 + InstalmentMode)
    Case Else
        GetInstalmentModeText = " "
End Select
Exit Function

ErrLine:
MsgBox Err.Number & vbCrLf & Err.Description
End Function
Private Sub LoadRepaymentTab(lMemberID As Long)

' Setup error handler.
On Error GoTo load_Err

' Varaiables...
Dim Lret As Long, i As Integer

' Validate the member id.
' Query for the member name.
If m_MemberObj Is Nothing Then
    Set m_MemberObj = New clsMMAcc
End If
txtMemberName.Text = m_MemberObj.MemberName(lMemberID)

' Query the loan details for the member.
gDbTrans.SQLStmt = "SELECT * FROM LoanMaster, LoanTypes WHERE " _
        & "MemberID = " & txtAccNo.Text & " AND " _
        & "loanMaster.SchemeID = LoanTypes.SchemeID"
Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No loans issued for this member.", _
            vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 582), _
            vbInformation, wis_MESSAGE_TITLE
    ' Clear the loan issue tab.
    LoanTabClear
    GoTo Exit_Line
End If

' The no. of rows retrieved, is the no. of loans,
' this member has taken.  So create that many tabs
' for displaying the loan details.
tabLOans.Tabs.Clear
With gDbTrans.Rst
    Do While Not .EOF
        tabLOans.Tabs.Add , "wis" & .fields("LoanID"), .fields("SchemeName")
        .MoveNext
    Loop
End With

Exit_Line:
    Me.MousePointer = vbDefault
    Exit Sub

load_Err:
    If Err Then
        'MsgBox "Load: " & Err.Description, vbCritical
        MsgBox LoadResString(gLangOffSet + 3) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Sub

Public Function LoanLoad(lLoanID As Long) As Boolean
On Error GoTo LoanLoad_Error

' Declare variables needed for this procedure...
Dim Lret As Long
Dim SkippedInstalments As Integer
Dim lastDate As String
Dim TransType As wisTransactionTypes
Dim MemClass As New clsMMAcc
' Get the loan details from the LoanMaster table.
gDbTrans.SQLStmt = "SELECT * FROM loanMaster WHERE " _
    & "loanID = " & lLoanID '& " AND NOT loanclosed"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then GoTo Exit_Line
' Save the resultset for future references.
Set m_rstLoanMast = gDbTrans.Rst.Clone

' Check wheather loanclosed or not
If FormatField(m_rstLoanMast("Loanclosed")) Then
    Me.cmdInstalment.Enabled = False
    Me.cmdLoanUpdate.Enabled = False
    Me.cmdRepay.Enabled = False
    Me.cmdUndo.Caption = LoadResString(gLangOffSet + 313) '"Reopen"
Else
    Me.cmdInstalment.Enabled = True
    Me.cmdLoanUpdate.Enabled = True
    Me.cmdRepay.Enabled = True
    Me.cmdUndo.Caption = LoadResString(gLangOffSet + 5) '"uNDO"
End If
' Get the scheme details.
gDbTrans.SQLStmt = "SELECT * FROM loantypes WHERE " _
        & "schemeID = (SELECT schemeID FROM loanmaster " _
        & "WHERE loanid = " & lLoanID & ")"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If
Set m_rstScheme = gDbTrans.Rst
m_SchemeId = Val(FormatField(m_rstScheme("SchemeId")))
' Fill the loan amount.
txtLoanAmt.Caption = FormatField(m_rstLoanMast("LoanAmt"))
' Fill the loan issue date.
txtIssueDate.Caption = FormatField(m_rstLoanMast("IssueDate"))

'Get the total repaid amount on this loan.
'iF HE HAS LAREADY REPAID A LOAN COMPLTELY THEN WE HAVE TO
' FETCH THE RECORDS OF NEW LOAN ONLY

gDbTrans.SQLStmt = "SELECT MAX(TransId) FROM loanTrans WHERE " _
        & " LoanID = " & lLoanID _
        & " And Balance < 1 "

Dim MaxTransId As Long

If gDbTrans.SQLFetch > 0 Then
    MaxTransId = FormatField(gDbTrans.Rst(0))
End If

TransType = wDeposit
If MaxTransId = 0 Then
    gDbTrans.SQLStmt = "SELECT SUM(amount) FROM loanTrans WHERE " _
        & "LoanID = " & lLoanID & " AND transtype = " & TransType
Else
    gDbTrans.SQLStmt = "SELECT SUM(amount) FROM loanTrans WHERE " _
        & "LoanID = " & lLoanID & " AND transtype = " & TransType _
        & " And TransID > " & MaxTransId
End If
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then GoTo Exit_Line
    txtRepaidAmt.Caption = FormatField(gDbTrans.Rst(0))

' Get the balance amount on this loan.
gDbTrans.SQLStmt = "SELECT TOP 1 Balance,Transid  FROM loantrans WHERE " _
        & "loanID = " & lLoanID & " ORDER BY transID Desc"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then GoTo Exit_Line
txtBalance.Caption = FormatField(gDbTrans.Rst(0))

' Get all the transactions on this loan.
gDbTrans.SQLStmt = "SELECT * FROM loantrans WHERE " _
        & "loanid = " & lLoanID & " ORDER BY transID"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then GoTo Exit_Line
Set m_rstLoanTrans = gDbTrans.Rst.Clone
With m_rstLoanTrans
    .MoveLast
    .Move -1 * (.AbsolutePosition Mod 10)
End With

' Load the recordset details to loan grid.
Call LoanLoadGrid
 'Set The User InterFace  for to update Loan Account
' cmdLoanUpdate.Enabled = True

Me.cmdSaveLoan.Enabled = False

' If instalment amount is defined for this loan account,
' display the amount.
If Val(FormatField(m_rstLoanMast("InstalmentAmt"))) > 0 Then
    txtInstAmt.Text = FormatField(m_rstLoanMast("InstalmentAmt"))
Else
    txtInstAmt.Text = ""
End If

' Compute and display the instalment amount due...
txtInstAmt.Text = FormatCurrency(ComputeInstalmentTotal)

'Get the Balance Interest if Any
   txtIntBalance.Text = FormatCurrency(FormatField(m_rstLoanMast("InterestBalance")))

Dim IntAmount As Currency
' Compute and display the regular interest.
IntAmount = ComputeRegularInterest(txtRepayDate.Text)
txtRegInterest.Text = FormatCurrency((IntAmount + Val(txtIntBalance.Text)) \ 1) 'FormatCurrency(ComputeRegularInterest(txtRepayDate.Text))

' Compute and display the penal interest for defaulted payments.
IntAmount = FormatCurrency(ComputePenalInterest(txtRepayDate.Text))
If IntAmount >= 0 Then
    txtPenalInterest.Text = FormatCurrency(IntAmount \ 1)
Else
    txtPenalInterest.Text = ""
End If
' Display the total instalment amount.
'txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) _
         + Val(txtRegInterest.Text) + Val(txtPenalInterest.Text))
txtRepayAmt.Text = "0.00" 'txtTotInst.Text
m_TotalInterest = Val(txtRegInterest.Text) + Val(txtPenalInterest.Text)
'Now Load the Same Loan detials into loan issue loan grid

Dim txtindex As Byte
' Check for member details.
txtindex = PropIssueGetIndex("MemberID")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("MemberId"))
txtLoanIssue(txtindex + 1).Text = MemClass.MemberName(FormatField(m_rstLoanMast("MemberId")))
' Loan Scheme .
txtindex = PropIssueGetIndex("LoanName")
txtLoanIssue(txtindex).Text = FormatField(m_rstScheme("SchemeName"))
txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "SchemeID", m_rstScheme("SchemeID"))

' Pledge item
txtindex = PropIssueGetIndex("PledgeDescription")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("PledgeDescription"))
 
' Pledge Value
txtindex = PropIssueGetIndex("PledgeValue")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("PledgeValue"))
 
 ' evaluator .
txtindex = PropIssueGetIndex("Evaluator")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("Evaluator"))

'No of Installment  'vinay
txtindex = PropIssueGetIndex("LoanInstalment")
'txtLoanIssue(txtindex).Text = IIf(FetchInstalmentNo(lLoanID) > 1, FetchInstalmentNo(lLoanID), 1)
 txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("LoanInstalments"))
 
 m_InstalmentDetails = ""
If Val(txtLoanIssue(txtindex).Text) > 1 Then
    m_InstalmentDetails = GetInstalmentNames(lLoanID)
End If
txtindex = PropIssueGetIndex("LoanAmount")
If m_InstalmentDetails <> "" Then
    txtLoanIssue(txtindex).Text = FormatField(m_rstLoanTrans("Amount"))
Else
    txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("LoanAmt"))
End If
' Loan Amount.
txtindex = PropIssueGetIndex("SanctionAmount")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("LoanAmt"))

' Loan due date.
txtindex = PropIssueGetIndex("LoanDueDate")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("LoanDueDate"))

' Instalment amount.
txtindex = PropIssueGetIndex("InstalmentAmount")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("InstalmentAmt"))
txtindex = PropIssueGetIndex("InstalmentMode")
Dim InstalmentMode As Byte
InstalmentMode = Val(ExtractToken(lblLoanIssue(txtindex).Tag, "TextIndex"))
txtLoanIssue(txtindex).Text = Me.cmbLoanIssue(InstalmentMode).List(FormatField(m_rstLoanMast("InstalmentMode")))

'Interest Rate
txtindex = PropIssueGetIndex("InterestRate")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("InterestRate"))
txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "InterestRate", FormatField(m_rstLoanMast("InterestRate")))

'Penal Interest Rate
txtindex = PropIssueGetIndex("PenalInterestRate")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("PenalInterestRate"))
txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "PenalInterestRate", FormatField(m_rstLoanMast("PenalInterestRate")))

txtindex = PropIssueGetIndex("InstalmentMode")
'txtLoanIssue(txtIndex) = FormatField(m_rstLoanMast("InstalmentMode"))
txtLoanIssue(txtindex) = GetInstalmentModeText(Val(FormatField(m_rstLoanMast("InstalmentMode"))))

' Guarantors ..
Dim l_MMObj As New clsMMAcc
txtindex = PropIssueGetIndex("Guarantor1")
txtLoanIssue(txtindex).Text = l_MMObj.MemberName(FormatField(m_rstLoanMast("GuarantorId1")))
txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "GuarantorID", FormatField(m_rstLoanMast("GuarantorID1")))

txtindex = PropIssueGetIndex("Guarantor2")
txtLoanIssue(txtindex).Text = l_MMObj.MemberName(FormatField(m_rstLoanMast("GuarantorID2")))
'Call putToken(txtLoanIssue(txtindex).Tag, "Guarantor2", txtLoanIssue(txtindex).Text)
 txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "GuarantorID", FormatField(m_rstLoanMast("GuarantorID2")))
Set l_MMObj = Nothing
' Create Date
txtindex = PropIssueGetIndex("IssueDate")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("IssueDate"))
txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "IssueDate", FormatField(m_rstLoanMast("IssueDate")))


 'Remarks
 txtindex = PropIssueGetIndex("Remarks")
txtLoanIssue(txtindex).Text = FormatField(m_rstLoanMast("Remarks"))

LoanLoad = True
m_LoanId = lLoanID
m_SchemeId = Val(FormatField(m_rstLoanMast("SchemeId")))
Exit_Line:


    Exit Function

LoanLoad_Error:
    If Err Then
        MsgBox "LoanLoad: " & Err.Description, vbCritical
'        MsgBox LoadResString(gLangOffSet + 713) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Function

Private Sub LoanLoadGrid()
Err.Clear
On Error GoTo Err_Line

' Variables for this procedure...
Dim i As Integer

' If no recordset, exit.
If m_rstLoanTrans Is Nothing Then Exit Sub
' If no records, exit.
If m_rstLoanTrans.RecordCount = 0 Then Exit Sub

cmdNextTrans.Enabled = False
If m_rstLoanTrans.AbsolutePosition >= 9 Then
    m_rstLoanTrans.MovePrevious
    cmdPrevTrans.Enabled = True
Else
    cmdPrevTrans.Enabled = False
End If

'Show 10 records or till eof of the page being pointed to
With grd
    ' Initialize the grid.
    .Clear: .AllowUserResizing = flexResizeBoth
    ' Set the format string for the loan details grid.
    .Cols = 5 '+ (Val(FormatField(m_rstLoanMast("LoanInstalments"))) - 1)
    If .Cols < 6 Then .Cols = 6
    'grd.FormatString = RPad("Date", " ", 13) & "|" _
        & RPad("Particulars", " ", 26) & "|" _
        & RPad("Instalment Amt", " ", 15) & "|" _
        & RPad("Interest Amt", " ", 15) & "|" & RPad("Penal Interest", " ", 15) & "|" _
        & RPad("Balance", " ", 18)
    
    .Rows = 30
    .FixedRows = 1
    .FixedCols = 0
    .Row = 0
    .Col = 0: .Text = LoadResString(gLangOffSet + 37): .ColWidth(0) = 1050  'Date
    .Col = 1: .Text = LoadResString(gLangOffSet + 39): .ColWidth(1) = 1100  'Particulars
    .Col = 2: .Text = LoadResString(gLangOffSet + 343): .ColWidth(2) = 1100 'Instalment Amt
    If m_InstalmentDetails <> "" Then
'''        Debug.Print " Search for the Code to written here "
    End If
    
    .Col = .Cols - 3: .Text = LoadResString(gLangOffSet + 274): .ColWidth(3) = 850 'Interest
    .Col = .Cols - 2: .Text = LoadResString(gLangOffSet + 345): .ColWidth(4) = 800 'Interest
    .Col = .Cols - 1: .Text = LoadResString(gLangOffSet + 42): .ColWidth(5) = 1050  'Balance
    .Visible = False
    i = 0

'    ' Position the recordset on the 2nd record.
'    ' Because, the first contains the amount of loan issued
'    ' against this account, and we want to show only
'    ' repayments made.
'    If m_rstLoanTrans.RecordCount < 2 Then
'        Exit Sub
'    End If
'    m_rstLoanTrans.AbsolutePosition = 2
    
    ' Begin a loop for filling thd details to grid.
    i = 1
    Dim NextRow As Boolean
    If m_rstLoanTrans.AbsolutePosition Mod 2 = 0 And m_rstLoanTrans.AbsolutePosition <> 0 Then
'''        m_rstLoanTrans.MoveLast
    End If
    .Row = 1
    Do
        .Row = i
        If Val(FormatField(m_rstLoanTrans("Amount"))) = 0 Then GoTo NextRecord
        ' Zeroth column.
            .Col = 0: .Text = FormatField(m_rstLoanTrans("TransDate"))
        ' First column
            .Col = 1: .Text = FormatField(m_rstLoanTrans("Particulars"))
        
        ' Decide the column, depending upon the type of transaction.
        If Val(FormatField(m_rstLoanTrans("TransType"))) = wDeposit Then
            .Col = 2
        ElseIf Val(FormatField(m_rstLoanTrans("TransType"))) = wCharges Then
            .Col = IIf(FormatField(m_rstLoanTrans("Particulars")) = LoadResString(gLangOffSet + 345), 4, 3)
        Else ' If it is loan issue TransaCtions
            .Col = 2: .CellAlignment = 1
        End If
        .Text = FormatField(m_rstLoanTrans("Amount"))
        If Val(.Text) > 0 Then i = i + 1
        ' Fourth column.
        .Col = 5: .Text = FormatField(m_rstLoanTrans("Balance"))
        If i < 30 And Not m_rstLoanTrans.EOF Then
NextRecord:
            m_rstLoanTrans.MoveNext
        Else
            cmdNextTrans.Enabled = True
            Exit Do
        End If
        If m_rstLoanTrans.EOF Then
'''            m_rstLoanTrans.AbsolutePosition = m_rstLoanTrans.RecordCount
            Exit Do
        Else
            'cmdNextTrans.Enabled = True
        End If
    Loop
    .Visible = True
    .Row = 1
End With

'Enable the UNDO Button
cmdUndo.Enabled = True
If m_rstLoanTrans.RecordCount = 1 Then ' IF he has not performed any transction on
   cmdUndo.Caption = LoadResString(gLangOffSet + 14) '' "Delete"
End If
'If m_rstLoanTrans.AbsolutePosition > 10 Then Me.cmdPrevTrans.Enabled = True
If i > 10 Then Me.cmdPrevTrans.Enabled = True
Exit Sub

Err_Line:
    If Err Then
        MsgBox "LoanLoadGrid: " & Err.Description
        'MsgBox LoadResString(gLangOffSet + 714) & Err.Description
    End If
'Resume
End Sub
Private Function LoanSave() As Boolean

' Setup error handler.
On Error GoTo LoanSave_Error

' Declare variables for this procedure...
' ------------------------------------------
Dim txtindex As Integer
Dim inTransaction As Boolean
Dim Lret As Long, nRet As Integer
Dim lMemberID As Long
Dim lGuarantorID As Long
Dim NewLoanID As Long
Dim selSchemeID As Long
Dim newTransID As Long
Dim Balance As Currency
Dim Amount As Currency
' ------------------------------------------

' Check for member details.
txtindex = PropIssueGetIndex("MemberID")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the member id of the person availing the loan.", vbInformation
        MsgBox LoadResString(gLangOffSet + 715), vbInformation
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If

    ' Check if the specified memberid is valid.
    gDbTrans.SQLStmt = "SELECT * FROM mmMaster WHERE AccID = " & .Text
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        'MsgBox "MemberID " & .Text & " does not exist.", vbExclamation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 716), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If

    ' ------------------------------------------------
    ' Check for the eligibility of this member...
    ' A member is elibible for taking loan, only if does not have any
    ' over due loans.
    If HasOverdueLoans(.Text) Then
        'nRet = MsgBox("The member hasoverdue loans.  Loan cannot be issued." _
                & vbCrLf & "Do you want to continue anyway?", vbQuestion + _
                vbYesNo, wis_MESSAGE_TITLE)
        nRet = MsgBox(LoadResString(gLangOffSet + 717) _
                & vbCrLf & LoadResString(gLangOffSet + 541), vbQuestion + _
                vbYesNo, wis_MESSAGE_TITLE)
        If nRet = vbNo Then
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End If
    '
    ' ------------------------------------------------
    lMemberID = Val(.Text)
End With

' Loan Scheme validation.
txtindex = PropIssueGetIndex("LoanName")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Select a loan scheme.", vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 718), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
End With


' Pledge item validation.
txtindex = PropIssueGetIndex("PledgeDescription")
If txtLoanIssue(txtindex).Text <> "" Then
    txtindex = PropIssueGetIndex("Pledgevalue")
    With txtLoanIssue(txtindex)
        ' Make sure that pledge value is mentioned,
        ' if pledge item is mentioned.
        If Trim(.Text) = "" Then
            'MsgBox "Specify the value of pledged items.", _
                        vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 719), _
                        vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        
        ' Ensure the value of pledge items is valid.
        If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for pledge item.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 720), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End With

    ' Make sure evaluator is mentioned.
    txtindex = PropIssueGetIndex("Evaluator")
    With txtLoanIssue(txtindex)
        ' Make sure that pledge value is mentioned,
        ' if pledge item is mentioned.
        If Trim(.Text) = "" Then
            'MsgBox "Specify the name of the evaluator for the pledge item.", _
                        vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 721), _
                        vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End With
End If

'LoanSanction Amount
txtindex = PropIssueGetIndex("SanctionAmount")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for loan amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 506), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
    End If
End With

' Loan Amount.
txtindex = PropIssueGetIndex("LoanAmount")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for loan amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 506), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
    End If
    ' If He Has not Mentioned the Loan Amount in Instamnet then
    ' Assign this loan amount as one Loan AMount
    If UBound(m_LoanAmount) = 0 Then m_LoanAmount(0) = Val(.Text)
    'loanamt = Val(.Text)
End With

' Instalment amount.
txtindex = PropIssueGetIndex("InstalmentAmount")
If txtLoanIssue(txtindex) <> "" Or txtLoanIssue(PropIssueGetIndex("InstalmentMode")).Text <> "" Then
    With txtLoanIssue(txtindex)
        If Trim(.Text) = "" Then
            'MsgBox "Specify the instalment amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 722), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
                'MsgBox "Invalid value for instalment amount.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox LoadResString(gLangOffSet + 506), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtLoanIssue(txtindex)
                GoTo Exit_Line
        End If
    End With
End If

' Loan due date.
txtindex = PropIssueGetIndex("LoanDueDate")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the date of maturity for this loan.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
End With

'Loan Interest Rate
txtindex = PropIssueGetIndex("InterestRate")
If Not IsNumeric(txtLoanIssue(txtindex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 646), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'loan Penal Interest Rate
txtindex = PropIssueGetIndex("PenalInterestRate")
If Not IsNumeric(txtLoanIssue(txtindex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 646), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


' Guarantors...
txtindex = PropIssueGetIndex("Guarantor1")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'nRet = MsgBox("Guarantor not specified. Do you want to continue?", _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        nRet = MsgBox(LoadResString(gLangOffSet + 723) & LoadResString(gLangOffSet + 541), _
                vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE)
        If nRet = vbNo Then GoTo Exit_Line
    End If
        
    ' Check if the guarantor is the same as the loan claimer.
    lGuarantorID = PropGuarantorID(1)
    If lGuarantorID = lMemberID Then
        'MsgBox "A person cannot stand guarantee for his own loan !", _
                vbExclamation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 724), _
                vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
    ' Check if the guarantor is eligible for standing guarantee.
    If HasOverdueLoans(lGuarantorID) Then  '//TODO//
        'MsgBox "Guarantor1 " & PropIssueGetVal("Guarantor1") _
                & " has loan overdues.  Please select another " _
                & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 725), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
End With

' Guarantor 2.
txtindex = PropIssueGetIndex("Guarantor2")
With txtLoanIssue(txtindex)
    If Trim(.Text) <> "" Then
        ' Check if both the guarantors are the same.
        If lGuarantorID <> 0 Then
            If lGuarantorID = PropGuarantorID(2) Then
                'MsgBox "Both guarantors should be different persons.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox LoadResString(gLangOffSet + 726), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtLoanIssue(txtindex)
                GoTo Exit_Line
            End If
        End If
        ' Check if the guarantor is the same as the loan claimer.
        lGuarantorID = PropGuarantorID(2)
        If lGuarantorID = lMemberID Then
            'MsgBox "A person cannot stand guarantee for his own loan!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 724), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        
        ' Check if the guarantor is eligible for standing guarantee.
        If HasOverdueLoans(lGuarantorID) Then
            'MsgBox "Guarantor2 " & PropIssueGetVal("Guarantor2") _
                    & " has loan overdues.  Please select another " _
                    & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 727), vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End If
End With


' Get a new loanid.
gDbTrans.SQLStmt = "SELECT MAX(LoanID) FROM LoanMaster"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then GoTo Exit_Line
NewLoanID = Val(FormatField(gDbTrans.Rst(0))) + 1

' Get the loan schemeid selected.
txtindex = PropIssueGetIndex("LoanName")
selSchemeID = Val(ExtractToken(txtLoanIssue(txtindex).Tag, "SchemeID"))

' Get the TransID and balance for updating LoanTrans Table.
gDbTrans.SQLStmt = "SELECT TOP 1 TransID, BALANCE FROM LoanTrans " _
        & "WHERE LoanID = " & NewLoanID
Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    newTransID = 1
    Amount = Val(PropIssueGetVal("LoanAmont"))
    Balance = Amount
Else
    newTransID = Val(FormatField(gDbTrans.Rst(0))) + 1
    Amount = Val(PropIssueGetVal("LoanAmount"))
    Balance = Val(FormatField(gDbTrans.Rst(1))) + Amount
End If

' Compute the interest to be levied on this loan amount.


' Begin transaction.
gDbTrans.BeginTrans
inTransaction = True

' Put an entry into LoanMaster table.
gDbTrans.SQLStmt = "INSERT INTO LoanMaster (LoanID, SchemeID, MemberID, " _
        & "IssueDate, PledgeValue, PledgeDescription, Evaluator, LoanAmt, " _
        & "InstalmentMode, InstalmentAmt, LoanDueDate, GuarantorID1, " _
        & "GuarantorID2, Remarks,InterestRate,PenalInterestRate, LoanInstalments) VALUES ("
gDbTrans.SQLStmt = gDbTrans.SQLStmt & NewLoanID & ",  " & selSchemeID & ", " _
        & PropIssueGetVal("MemberID") & ", " _
        & " #" & FormatDate(PropIssueGetVal("IssueDate")) & "#, " _
        & Val(PropIssueGetVal("PledgeValue")) & ", " _
        & AddQuotes(PropIssueGetVal("PledgeDescription"), True) & ", " _
        & AddQuotes(PropIssueGetVal("Evaluator"), True) & ", " _
        & PropIssueGetVal("SanctionAmount") & ", " _
        & GetInstalmentMode & ", " _
        & Val(PropIssueGetVal("InstalmentAmount")) & ", " _
        & "#" & FormatDate(PropIssueGetVal("LoanDueDate")) & "#, " _
        & PropGuarantorID(1) & ", " _
        & PropGuarantorID(2) & ", " _
        & AddQuotes(PropIssueGetVal("Remarks"), True) & "," _
        & Val(PropIssueGetVal("InterestRate")) & "," _
        & Val(PropIssueGetVal("PenalInterestRate")) & "," _
        & Val(PropIssueGetVal("LoanInstalMent")) _
        & ")"
'        & "," & "'" & M_Instalmentdetail & "')"
        
        
        
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    Dim Count As Integer
    
' Put an entry into LoanTrans table.
'put the the No Of entries as many No of Instalments
    Balance = 0
    Amount = 0
For Count = 0 To UBound(m_LoanAmount)
    'To Make Easy Undo He Has TO Entere Interest Amount as 0 Before Issue A Loan
    If m_LoanAmount(Count) <> 0 Then
        Balance = Balance + m_LoanAmount(Count)
        gDbTrans.SQLStmt = "INSERT INTO loantrans (loanID, TransID, TransType, " _
                & "Amount, TransDate, Balance, Particulars) VALUES (" _
                & NewLoanID & ", " & newTransID & ", -1, " _
                & m_LoanAmount(Count) & ", " _
                & "#" & FormatDate(PropIssueGetVal("IssueDate")) & "#, " _
                & Balance & ", 'Loan Issued')"
                '& Val(PropIssueGetVal("LoanAmount")) & ", 'Loan Issued')"
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        newTransID = newTransID + 1
    End If
Next Count
'if Loan Has any Componentes Then Make respective Transaction
If m_InstalmentDetails <> "" Then
    Dim InstDet() As String
    Count = GetStringArray(m_InstalmentDetails, InstDet(), ";")
    For Count = 0 To UBound(InstDet)
        gDbTrans.SQLStmt = "Insert INto LoanComponent(LoanId,CompName,CompAmount) Values  " & _
                    " ( " & NewLoanID & "," & _
                    "'" & InstDet(Count) & "'," & _
                    Val(InstDet(Count + 1)) & _
                    ")"
                    Count = Count + 1
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    Next Count
End If

' Commit the transaction
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False

' Credit The Corresponding Loan Amount Into SBTransTab
'Get The SB Accid For This MemberID
Dim MemID As Long
Dim CustId As Long
Dim SbAccID As Long
Dim SbObj As clsSBAcc

If chkSb.value Then
    
    MemID = PropIssueGetVal("MemberID")
    gDbTrans.SQLStmt = "Select CustomerID From MMMaster Where Accid = " & MemID
    
    If gDbTrans.SQLFetch <> 1 Then
        'MsgBox "Cannot Update The SB Account ", vbExclamation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 436) & " " & LoadResString(gLangOffSet + 640), vbExclamation, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
    CustId = Val(gDbTrans.Rst("CustomerID"))
    Set SbObj = New clsSBAcc
    If SbObj.CustomerBalance(CustId, SbAccID) > 0 Then
        If Not SbObj.DepositAmount(SbAccID, Balance, "Loan Sanctioned", PropIssueGetVal("IssueDate")) Then
            'MsgBox "Cannot Update The SB Account ", vbExclamation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 436) & " " & LoadResString(gLangOffSet + 640), vbExclamation, wis_MESSAGE_TITLE
        End If
    End If
    Set SbObj = Nothing
    
End If




'MsgBox "New loan created successfully.", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 728), vbInformation, wis_MESSAGE_TITLE

'Now Load the Same loan ito repayment tab
txtAccNo.Text = PropIssueGetVal("MemberID")
If tabLOans.Tabs.Count > 1 Then
    tabLOans.Tabs(tabLOans.Tabs.Count).Selected = True:   '.SelectedItem=
    'Call Tabloans_Click
End If
Me.cmdSaveLoan.Enabled = False
Me.cmdLoanUpdate.Enabled = True

Exit_Line:
    'If Val(txtLoanIssue(PropIssueGetIndex("InstalmentNames")).Text) > 1 Then Unload frmLoanName
    If inTransaction Then gDbTrans.RollBack
    
Exit Function

LoanSave_Error:
    If Err Then
        MsgBox "LoanSave: " & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 729) & Err.Description, vbCritical
    End If
Resume
    GoTo Exit_Line
End Function

Private Sub LoanTabClear()

 txtLoanAmt.Caption = ""
 txtIssueDate.Caption = ""
 txtRepaidAmt.Caption = ""
 txtBalance.Caption = ""
 txtIntBalance.Text = ""
 txtInstAmt.Text = ""
 txtRegInterest.Text = ""
 txtPenalInterest.Text = ""
 txtTotInst.Text = ""
 txtRepayDate.Text = FormatDate(gStrDate)
 txtRepayAmt.Text = ""
 
 grd.Clear
 tabLOans.Tabs.Clear
 cmdInstalment.Enabled = False
 cmdRepay.Enabled = False
 Me.cmdUndo.Enabled = False
 ' Cler the Loan Issue Grid Also
 Call cmdLoanIssueClear_Click
End Sub
' Reverts the last transaction in loan transactions table.
Private Function LoanUndoLastTransaction() As Boolean

' Variables of the procedure...
Dim lastTransID As Long
Dim NextTransID As Long
Dim inTransaction As Boolean
Dim FirstTransType As wisTransactionTypes
Dim SecondTransType As wisTransactionTypes
Dim ThirdTransType As wisTransactionTypes
Dim Lret As Long
Dim TransRst As Recordset
Dim PenalInterest As Currency
Dim Interest As Currency
Dim Amount As Currency
Dim Transdate As String

' Setup the error handler.
On Error GoTo Err_Line

' Check if a loan account is loaded.
If m_rstLoanMast Is Nothing Then GoTo Exit_Line

' Get the last transaction from the loan transaction table.
' Fetch the last three rows.
gDbTrans.SQLStmt = "SELECT Top 3 transid, transtype,Amount,TransDate FROM loantrans WHERE " _
        & "loanid = " & FormatField(m_rstLoanMast("loanid")) _
        & " ORDER BY transid DESC"
Lret = gDbTrans.SQLFetch

If Lret < 0 Then
    MsgBox "Error querying database!", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 1 Then
    'MsgBox "There are no transaction to undo.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 661), vbInformation, wis_MESSAGE_TITLE
    LoanUndoLastTransaction = True
    GoTo Exit_Line
ElseIf Lret < 1 Then
    'MsgBox "There are no transaction to undo.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 661), vbInformation, wis_MESSAGE_TITLE
    LoanUndoLastTransaction = True
    GoTo Exit_Line
End If

'Assign the recordset to a variable
Set TransRst = gDbTrans.Rst.Clone

' Check the transaction types of the records to undo.
With TransRst
    Transdate = FormatField(.fields("TransDate"))
    
    ' First transtype.
    FirstTransType = FormatField(.fields("Transtype"))
    'Interest = Val(FormatField(.fields("Amount")))
    Amount = Val(FormatField(.fields("Amount")))
    
    ' second type
    .MoveNext
    SecondTransType = FormatField(.fields("Transtype"))
    If SecondTransType = wCharges Then
        Interest = Val(FormatField(.fields("Amount")))
    End If
    
    ' Third Type
    'If records are more than two then only
    PenalInterest = 0
    If .EOF = False Then
        .MoveNext
        ThirdTransType = FormatField(.fields("Transtype"))
        
        If ThirdTransType = wCharges Then
            PenalInterest = Val(FormatField(.fields("Amount")))
        Else
            PenalInterest = 0
            .MovePrevious
        End If
    End If
    
    ' Compare.
    If FirstTransType = wDeposit Or FirstTransType = wWithDraw Then
        If SecondTransType <> wCharges Then
            'MsgBox "Invalid transaction type, or the last transaction entries are not proper."
            'MsgBox LoadResString(gLangOffSet + 588)
            'GoTo Exit_Line
        End If
    ElseIf FirstTransType = wCharges Then
        'Assign Amount=INTTEREST AND Interest=Amount
        Amount = Amount + Interest
        Interest = Amount - Interest
        Amount = Amount - Interest
        ' While Issuing Component Loans THis Possible
'        If SecondTransType <> wDeposit Then
'            'MsgBox "Invalid transaction type, or the last transaction entries are not proper."
'            MsgBox LoadResString(gLangOffSet + 588)
'            GoTo Exit_Line
'        End If
    End If
    .MoveFirst
End With

' Get the transaction id for this last transaction.
TransRst.MoveFirst
lastTransID = Val(FormatField(TransRst("Transid")))

' Begin transaction
gDbTrans.BeginTrans
inTransaction = True

' Delete the entry of principle amount from the transaction table.
gDbTrans.SQLStmt = "DELETE FROM loantrans WHERE transid = " & lastTransID _
        & " AND loanid = " & m_rstLoanMast("loanid")
If Not gDbTrans.SQLExecute Then
    MsgBox "Error updating the loan database.", _
            vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Delete the entry of Interest amount from the transaction table.
If SecondTransType = wCharges Then
     NextTransID = lastTransID - 1
    gDbTrans.SQLStmt = "DELETE FROM loantrans WHERE transid = " & NextTransID _
            & " AND loanid = " & m_rstLoanMast("loanid")
    If Not gDbTrans.SQLExecute Then
        MsgBox "Error updating the loan database.", _
                vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
ElseIf SecondTransType = wInterest Then
    'This is the case if he has paid extra Ainterest Prevoius to this tanasction
    'And he has repaid the extra amount then undo this transaction
    ' In such case the loan must have closed & Balance must be zeo
    'If FormatField(m_rstLoanMast("LoanClosed")) = True Or FormatField(m_rstLoanMast("Balance")) = 0 Then
    TransRst.MoveFirst
    If FormatField(TransRst("Balance")) = 0 Then
         NextTransID = lastTransID - 1
        gDbTrans.SQLStmt = "DELETE FROM loantrans WHERE transid = " & NextTransID _
                & " AND loanid = " & m_rstLoanMast("loanid")
        If Not gDbTrans.SQLExecute Then
            MsgBox "Error updating the loan database.", _
                    vbCritical, wis_MESSAGE_TITLE
            GoTo Exit_Line
        End If
    End If
End If


' Delete the entry of Penal interest amount from the transaction table.
If ThirdTransType = wCharges Then
 NextTransID = lastTransID - 2
    gDbTrans.SQLStmt = "DELETE FROM loantrans WHERE transid = " & NextTransID _
            & " AND loanid = " & m_rstLoanMast("loanid")
    If Not gDbTrans.SQLExecute Then
        MsgBox "Error updating the loan database.", _
                vbCritical, wis_MESSAGE_TITLE
        GoTo Exit_Line
    End If
End If


'UpDate The LoanMAster
gDbTrans.SQLStmt = "UPDATE loanmaster SET LoanClosed = False Where Loanid =  " & m_LoanId
        ' Execute the updation.
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
'gDBTrans.SQLStmt = "UPDATE loanmaster SET intpaiddate = NULL Where Loanid =  " & m_LoanId
        ' Execute the updation.
        'If Not gDBTrans.SQLExecute Then GoTo Exit_Line

' Commit the transaction.
gDbTrans.CommitTrans
inTransaction = False

' UpDate The InterestBalance Field Aproximately
Dim RegInt As Currency
Dim PenalInt As Currency
Dim BalanceInt As Currency
PenalInt = ComputePenalInterest(Transdate)
Amount = FormatCurrency(ComputeRegularInterest(Transdate) + IIf(PenalInt < 0, 0, PenalInt))
BalanceInt = Val(FormatField(m_rstLoanMast("InterestBalance")))
If BalanceInt > 0 Then
  ' Now Find The Difference between Calulated Interest & Actual Paid Interest
  ' If the Difference More than One then only Consider the INterestBalance
    If Amount - Interest > 1 Or FirstTransType = wWithDraw Then
        'Now Update the InterestBalance of LoanMaster
        BalanceInt = IIf(Amount - Interest > BalanceInt, BalanceInt - (Amount - Interest), 0)
        If BalanceInt < 0 Then BalanceInt = 0
        gDbTrans.SQLStmt = "UpDate LoanMaster Set InterestBalance = " & BalanceInt & " Where LoanId = " & m_LoanId
        gDbTrans.BeginTrans
        If Not gDbTrans.SQLExecute Then
            'MsgBox "Unable to the Balance Interest", vbExclamation, wis_MESSAGE_TITLE
            gDbTrans.RollBack
            Exit Function
        End If
        gDbTrans.CommitTrans
    End If
End If

' If any Miscelenous amount collected that has to remove
Dim BankClass  As New clsBankAcc
Call BankClass.UndoUPDatedMiscProfit(Transdate, "LoanRepay " & m_LoanId & "-" & lastTransID)
Set BankClass = Nothing

'MsgBox "The last transaction is deleted.", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 730), vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    If inTransaction Then gDbTrans.RollBack
    Exit Function

Err_Line:
    If Err Then
        MsgBox "LoanUndoLastTransaction: " & vbCrLf _
                '& Err.Description, vbCritical, wis_MESSAGE_TITLE
        'MsgBox LoadResString(gLangOffSet + 731) & vbCrLf _
                & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
    GoTo Exit_Line

End Function
Private Function LoanUpDate() As Boolean

' Setup error handler.
On Error GoTo LoanUpdate_Error:

' Declare variables for this procedure...
' ------------------------------------------
Dim txtindex As Integer
Dim inTransaction As Boolean
Dim Lret As Long, nRet As Integer
Dim lMemberID As Long
Dim lGuarantorID As Long
Dim lGuarantorID2 As Long
Dim NewLoanID As Long
Dim selSchemeID As Long
Dim newTransID As Long
Dim Balance As Currency
Dim Amount As Currency
Dim InstalmentNames As String
Dim lRst As Recordset
' ------------------------------------------

' Check for member details.
txtindex = PropIssueGetIndex("MemberID")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the member id of the person availing the loan.", vbInformation
        MsgBox LoadResString(gLangOffSet + 715), vbInformation
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If

    ' Check if the specified memberid is valid.
    gDbTrans.SQLStmt = "SELECT * FROM mmMaster WHERE AccID = " & Val(.Text)
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        'MsgBox "MemberID " & .Text & " does not exist.", vbExclamation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 716), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    Else
        Set lRst = gDbTrans.Rst.Clone
    End If

    ' ------------------------------------------------
    ' Check for the eligibility of this member...
    ' A member is elibible for taking loan, only if does not have any
    ' over due loans.
    If HasOverdueLoans(.Text) Then
        'nRet = MsgBox("The member hasoverdue loans.  Loan cannot be issued." _
                & vbCrLf & "Do you want to continue anyway?", vbQuestion + _
                vbYesNo, wis_MESSAGE_TITLE)
        nRet = MsgBox(LoadResString(gLangOffSet + 717) _
                & vbCrLf & LoadResString(gLangOffSet + 541), vbQuestion + _
                vbYesNo, wis_MESSAGE_TITLE)
        If nRet = vbNo Then
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End If
    '
    ' ------------------------------------------------
    lMemberID = Val(.Text)
End With

' Loan Scheme validation.
txtindex = PropIssueGetIndex("LoanName")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Select a loan scheme.", vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 718), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
End With


' Pledge item validation.
txtindex = PropIssueGetIndex("PledgeDescription")
If txtLoanIssue(txtindex).Text <> "" Then
    txtindex = PropIssueGetIndex("Pledgevalue")
    With txtLoanIssue(txtindex)
        ' Make sure that pledge value is mentioned,
        ' if pledge item is mentioned.
        If Trim(.Text) = "" Then
            'MsgBox "Specify the value of pledged items.", _
                        vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 719), _
                        vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        
        ' Ensure the value of pledge items is valid.
        If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for pledge item.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 720), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End With

    ' Make sure evaluator is mentioned.
    txtindex = PropIssueGetIndex("Evaluator")
    With txtLoanIssue(txtindex)
        ' Make sure that pledge value is mentioned,
        ' if pledge item is mentioned.
        If Trim(.Text) = "" Then
            'MsgBox "Specify the name of the evaluator for the pledge item.", _
                        vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 721), _
                        vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
    End With
End If

' Loan Amount.
txtindex = PropIssueGetIndex("LoanAmount")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the loan amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for loan amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 506), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
    End If
'loanamt = Val(.Text)
End With

' Instalment amount.
txtindex = PropIssueGetIndex("InstalmentAmount")
If Trim(txtLoanIssue(txtindex).Text) <> "" Or Trim(txtLoanIssue(PropIssueGetIndex("InstalmentMode")).Text) <> "" Then
    'txtindex = PropIssueGetIndex("LoanAmount")
    With txtLoanIssue(txtindex)
        If Trim(.Text) = "" Then
            'MsgBox "Specify the instalment amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 722), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        
        If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
                'MsgBox "Invalid value for instalment amount.", _
                        vbInformation, wis_MESSAGE_TITLE
                MsgBox LoadResString(gLangOffSet + 506), _
                        vbInformation, wis_MESSAGE_TITLE
                ActivateTextBox txtLoanIssue(txtindex)
                GoTo Exit_Line
        End If
    End With
End If

txtindex = PropIssueGetIndex("InterestRate")
If Not IsNumeric(txtLoanIssue(txtindex)) Then
    'MsgBox "Please Specify The Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 646), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

txtindex = PropIssueGetIndex("PenalInterestRate")
If Not IsNumeric(txtLoanIssue(txtindex)) Then
    'MsgBox "Please Specify The Penal Interest Rate ", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 742), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If


' Loan due date.
txtindex = PropIssueGetIndex("LoanDueDate")
With txtLoanIssue(txtindex)
    If Trim(.Text) = "" Then
        'MsgBox "Specify the date of maturity for this loan.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanIssue(txtindex)
        GoTo Exit_Line
    End If
End With

' Guarantors...
txtindex = PropIssueGetIndex("Guarantor1")
With txtLoanIssue(txtindex)
    lGuarantorID = PropGuarantorID(1)
    If Trim(.Text) = "" Then
        'nRet = MsgBox("Guarantor not specified. Do you want to continue?", _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        nRet = MsgBox(LoadResString(gLangOffSet + 723) & LoadResString(gLangOffSet + 541), _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        If nRet = vbNo Then GoTo Exit_Line
    End If
        
    ' Check if the guarantor is the same as the loan claimer.
    If lGuarantorID = lMemberID Then
        'nRet =MsgBox( "A person cannot stand guarantee for his own loan !", _
                vbExclamation+vbyesno, wis_MESSAGE_TITLE)
        nRet = MsgBox(LoadResString(gLangOffSet + 724) & LoadResString(gLangOffSet + 541), _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        If nRet = vbNo Then GoTo Exit_Line
        
    End If
    ' Check if the guarantor is eligible for standing guarantee.
    If HasOverdueLoans(lGuarantorID) Then  '//TODO//
        'MsgBox "Guarantor1 " & PropIssueGetVal("Guarantor1") _
                & " has loan overdues.  Please select another " _
                & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
        nRet = MsgBox(LoadResString(gLangOffSet + 725) & LoadResString(gLangOffSet + 541), _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
        If nRet = vbNo Then GoTo Exit_Line
    End If
End With

' Guarantor 2.
txtindex = PropIssueGetIndex("Guarantor2")
With txtLoanIssue(txtindex)
    If Trim(.Text) <> "" Then
        lGuarantorID2 = PropGuarantorID(2)
        ' Check if both the guarantors are the same.
        If lGuarantorID = lGuarantorID2 And lGuarantorID <> 0 Then
            'MsgBox "Both guarantors should be different persons.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 726), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        ' Check if the guarantor is the same as the loan claimer.
        If lGuarantorID = lMemberID Then
            'MsgBox "A person cannot stand guarantee for his own loan!", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 724), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtLoanIssue(txtindex)
            GoTo Exit_Line
        End If
        
        ' Check if the guarantor is eligible for standing guarantee.
        If HasOverdueLoans(lGuarantorID2) Then
            'MsgBox "Guarantor2 " & PropIssueGetVal("Guarantor2") _
                    & " has loan overdues.  Please select another " _
                    & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
            nRet = MsgBox(LoadResString(gLangOffSet + 725) & LoadResString(gLangOffSet + 541), _
                vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
            If nRet = vbNo Then GoTo Exit_Line
        End If
    End If
End With


' Get the loan schemeid selected.
txtindex = PropIssueGetIndex("LoanName")
    selSchemeID = Val(ExtractToken(txtLoanIssue(txtindex).Tag, "SchemeID"))
    newTransID = Val(FormatField(gDbTrans.Rst(0))) + 1
    Amount = Val(PropIssueGetVal("LoanAmt"))
    'Balance = Val(FormatField(gDBTrans.Rst(1))) + Amount
'Get The  Loan previous LoanAmount & The Differencr of the Amount
'Loan Amount
gDbTrans.SQLStmt = "Select Amount From LoanTrans Where LoanId = " & m_LoanId & " And TransId = 1 "
If gDbTrans.SQLFetch < 1 Then GoTo Exit_Line
Dim PrevBalance As Currency, DiffBalance As Currency
PrevBalance = FormatField(gDbTrans.Rst(0))
DiffBalance = PrevBalance - Amount
' Begin transaction.
gDbTrans.BeginTrans
inTransaction = True

' Put an entry into LoanMaster table.
gDbTrans.SQLStmt = "UpDate LoanMaster Set SchemeID =" & selSchemeID & ", " & _
        "MemberID= " & PropIssueGetVal("MemberID") & ", " & _
        "IssueDate = #" & FormatDate(PropIssueGetVal("IssueDate")) & "#, " & _
        "PledgeValue = " & Val(PropIssueGetVal("PledgeValue")) & ", " & _
        "PledgeDescription = " & AddQuotes(PropIssueGetVal("PledgeDescription"), True) & ", " & _
        "Evaluator = " & AddQuotes(PropIssueGetVal("Evaluator"), True) & ", " & _
        "LoanAmt = " & PropIssueGetVal("SanctionAmount") & ", " & _
        "InstalmentMode = " & GetInstalmentMode & ", " & _
        "InstalmentAmt = " & Val(PropIssueGetVal("InstalmentAmount")) & ", " & _
        "LoanDueDate = #" & FormatDate(PropIssueGetVal("LoanDueDate")) & "#, " & _
        " GuarantorID1 = " & lGuarantorID & ", " & _
        "GuarantorID2 = " & lGuarantorID2 & ", " & _
         " Remarks = " & AddQuotes(PropIssueGetVal("Remarks"), True) & "," & _
         " InterestRate = " & Val(PropIssueGetVal("InterestRate")) & "," & _
         " PenalInterestRate =" & Val(PropIssueGetVal("PenalInterestRate")) & "," & _
         " LoanInstalments ='" & Val(PropIssueGetVal("LoanInstalment")) & "'"
         
gDbTrans.SQLStmt = gDbTrans.SQLStmt & " Where LoanID = " & m_LoanId
If Not gDbTrans.SQLExecute Then GoTo Exit_Line

'Put an entry into LoanTrans table by updating the LOan BAlance
'gDBTrans.SQLStmt = "UpDate LoanTrans set  Balance =  Balance - " & DiffBalance & _
        " Where  Loanid = " & m_LoanId
'If Not gDBTrans.SQLExecute Then GoTo Exit_Line

gDbTrans.SQLStmt = "Delete * From LoanComponent  Where  Loanid = " & m_LoanId
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False
gDbTrans.BeginTrans
inTransaction = True

If m_InstalmentDetails <> "" Then
    Dim Count As Integer
    Dim InstDet() As String
    Count = GetStringArray(m_InstalmentDetails, InstDet(), ";")
    For Count = 0 To UBound(InstDet)
        gDbTrans.SQLStmt = "Insert INto LoanComponent(LoanId,CompName,CompAmount) Values  " & _
                    " ( " & NewLoanID & "," & _
                    "'" & InstDet(Count) & "'," & _
                    Val(InstDet(Count + 1)) & _
                    ")"
                    Count = Count + 1
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
    Next Count
End If


' Commit the transaction
If Not gDbTrans.CommitTrans Then GoTo Exit_Line
inTransaction = False
MsgBox LoadResString(gLangOffSet + 707), vbInformation, wis_MESSAGE_TITLE
'MsgBox "Loan updated  successfully.", vbInformation, wis_MESSAGE_TITLE

Exit_Line:
    'If Val(txtLoanIssue(PropIssueGetIndex("InstalmentNames")).Text) > 1 Then Unload frmLoanName
    If inTransaction Then gDbTrans.RollBack
    Exit Function

LoanUpdate_Error:
    If Err Then
        MsgBox "LoanSave: " & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 728) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line

End Function

' Returns the memberID of Guarantor1, selected by the user.
Private Function PropGuarantorID(GuarantorNum As Integer) As Long
Dim strTmp As String
Dim txtindex As Integer
txtindex = PropIssueGetIndex("Guarantor" & GuarantorNum)
If txtindex >= 0 Then
    strTmp = ExtractToken(txtLoanIssue(txtindex).Tag, "GuarantorID")
    PropGuarantorID = Val(strTmp)
End If

End Function
Private Sub PropInitializeForm()

' Set the proerties for tab strip control.
TabStrip.ZOrder 1
TabStrip.Tabs(1).Selected = True

' Load the properties for loan creation panel.
LoadLoanTypeProp
' Load the properties for loan issue panel.
LoadLoanIssueProp

' Remove all tabs of tabloans.
tabLOans.Tabs.Clear

' Load the loan types to reports tab.
LoanLoadTypes

End Sub
' Returns the text value from a control array
' bound the field "FieldName".
Private Function PropIssueGetVal(FieldName As String) As String
Dim i As Integer
Dim strTxt As String
For i = 0 To txtLoanIssue.Count - 1
    strTxt = ExtractToken(lblLoanIssue(i).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        PropIssueGetVal = txtLoanIssue(i).Text
        Exit For
    End If
Next
End Function


Private Function PropIssueGetIndex(strDataSrc As String) As Integer
PropIssueGetIndex = -1
Dim strTmp As String
Dim i As Integer
For i = 0 To lblLoanIssue.Count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(lblLoanIssue(i).Tag, "DataSource")
    If StrComp(strDataSrc, strTmp, vbTextCompare) = 0 Then
        PropIssueGetIndex = i
        Exit For
    End If
Next

End Function
Private Sub PropSchemeClear()
Dim i As Integer
For i = 0 To txtLoanScheme.Count - 1
    txtLoanScheme(i).Text = ""
Next

i = PropSchemeGetIndex("CreateDate")
txtLoanScheme(i).Text = FormatDate(gStrDate)
m_SchemeLoaded = False
Dim txtindex As Integer

txtindex = PropSchemeGetIndex("LoanName")
txtLoanScheme(txtindex).Locked = False

lblOpMode.Caption = "Operation Mode: <INSERT>"
End Sub
' Returns the index of the control bound to "strDatasrc".
Private Function PropSchemeGetIndex(strDataSrc As String) As Integer
PropSchemeGetIndex = -1
Dim strTmp As String
Dim i As Integer
For i = 0 To lblLoanScheme.Count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(lblLoanScheme(i).Tag, "DataSource")
    If StrComp(strDataSrc, strTmp, vbTextCompare) = 0 Then
        PropSchemeGetIndex = i
        Exit For
    End If
Next
End Function


'****************************************************************************************
'Returns a new account number
'Author: Girish
'Date : 29th Dec, 1999
'Modified by Ravindra on 25th Jan, 2000
'****************************************************************************************
Private Function GetNewAccountNumber() As Long
    Dim NewAccNo As Long
    'gDBTrans.SQLStmt = "Select TOP 1 AccID from SBMaster order by AccID desc"
    gDbTrans.SQLStmt = "SELECT MAX(AccID) FROM SBMaster"
    If gDbTrans.SQLFetch = 0 Then
        NewAccNo = 1
    Else
        NewAccNo = Val(FormatField(gDbTrans.Rst(0))) + 1
    End If
    GetNewAccountNumber = NewAccNo
End Function
Private Sub DisplayResults(MoveDirection As Integer)
#If RAVI_COMMENTED Then
'On Error Resume Next
'Uses the record set to display results
    If m_rstSearchResults Is Nothing Then
        Exit Sub
    End If
    
    If MoveDirection = 1 Then
        m_rstSearchResults.MoveNext
    ElseIf MoveDirection = -1 Then
        m_rstSearchResults.MovePrevious
    End If

    
    Dim CustomerId As Long
    CustomerId = FormatField(m_rstSearchResults("SBMaster.CustomerID"))
    If Not m_AccHolder.LoadCustomerInfo(CustomerId) Then
        Exit Sub
    End If
    txtNewAccNo.Text = m_rstSearchResults("AccID")
    m_AccNo = FormatField(m_rstSearchResults("AccID"))
    txtName.Text = m_AccHolder.FullName
    
    cmdDetails.Enabled = True

    'txtNewDate.Text = m_rstsearchresults("CreateDate")
    txtNewDate.Text = FormatField(m_rstSearchResults("CreateDate"))
    txtLedger.Text = FormatField(m_rstSearchResults("LedgerNo"))
    txtFolio.Text = FormatField(m_rstSearchResults("FolioNo"))
    
    Dim i As Integer
    Dim Arr() As String
    Call GetStringArray(FormatField(m_rstSearchResults("Nominee")), Arr, ";")
    txtNominee.Text = "": txtAge.Text = "": cmbRelation.ListIndex = -1
    For i = 0 To UBound(Arr)
        Select Case i
            Case 0: txtNominee.Text = Arr(i)
            Case 1: txtAge.Text = Arr(i)
            Case 2: cmbRelation.ListIndex = Val(Arr(i))
        End Select
    Next i
    
    Call GetStringArray(FormatField(m_rstSearchResults("JointHolder")), Arr, ";")
    cmbHolders.Clear
    For i = 0 To UBound(Arr)
        cmbHolders.AddItem Arr(i)
    Next i
    
    txtIntroAccNo.Text = FormatField(m_rstSearchResults("Introduced"))
    If Val(txtIntroAccNo.Text) = 0 Then
        txtIntroAccNo.Text = ""
    End If
    
    'Load the name now
    Call m_AccHolder.LoadCustomerInfo(FormatField(m_rstSearchResults("SBMaster.CustomerID")))
    txtIntroName.Text = m_AccHolder.FullName
    
    m_rstSearchResults.MovePrevious
    If m_rstSearchResults.BOF Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If
    If m_rstSearchResults.BOF Then
        m_rstSearchResults.MoveFirst
    Else
        m_rstSearchResults.MoveNext
    End If
    
    m_rstSearchResults.MoveNext
    If m_rstSearchResults.EOF Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If m_rstSearchResults.EOF Then
        m_rstSearchResults.MoveLast
    Else
        m_rstSearchResults.MovePrevious
    End If
    
#End If
End Sub
Private Function GetNewTransID(AccNo As Long) As Long
If AccNo = 0 Then
    Exit Function
End If
gDbTrans.SQLStmt = "Select AccID from SBMaster where AccID = " & AccNo
If gDbTrans.SQLFetch() <> 1 Then
    'MsgBox "Account number not found !", vbExclamation, wis_MESSAGE_TITLE & " - Error"
    MsgBox LoadResString(gLangOffSet + 500), vbExclamation, wis_MESSAGE_TITLE & " - Error"
    Exit Function
End If

gDbTrans.SQLStmt = "Select TOP 1 AccID, TransID from SBTrans where AccID = " & AccNo & " order by TransID desc"
If gDbTrans.SQLFetch <= 0 Then
    GetNewTransID = 1
Else
    GetNewTransID = FormatField(gDbTrans.Rst("TransID")) + 1
End If

End Function


' Returns the text value from a control array
' bound the field "FieldName".
Private Function PropSchemeGetVal(FieldName As String) As String
Dim i As Integer
Dim strTxt As String
For i = 0 To txtLoanScheme.Count - 1
    strTxt = ExtractToken(lblLoanScheme(i).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        PropSchemeGetVal = txtLoanScheme(i).Text
        Exit For
    End If
Next
End Function
Private Function LoadLoanIssueProp() As Boolean

'
' Read the data from Loans.ini and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
If gLangOffSet = wis_KannadaOffset Then
    PropFile = App.Path & "\Loanskan.PRP"
Else
    PropFile = App.Path & "\Loans.PRP"
End If
If Dir(PropFile, vbNormal) = "" Then

    'MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox LoadResString(gLangOffSet + 602) _
            & PropFile & "' !", vbExclamation
    Exit Function
End If

' Declare required variables...
Dim strTmp As String
Dim strPropType As String
Dim FirstImgCtl As Boolean
Dim FirstControl As Boolean
Dim i As Integer, CtlIndex As Integer
Dim strRet As String, imgCtlIndex As Integer
FirstControl = True
FirstImgCtl = True
Dim strTag As String

' Read all the prompts and load accordingly...
Do
    ' Read a line.
    strTag = Trim(ReadFromIniFile("LoanIssue", _
                "Prop" & i + 1, PropFile))
    If strTag = "" Then Exit Do

    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load lblLoanIssue(lblLoanIssue.Count)
        Load txtLoanIssue(txtLoanIssue.Count)
    End If
    CtlIndex = lblLoanIssue.Count - 1
'    Debug.Assert I <> 15
    ' Get the property type.
    strPropType = Trim$(ExtractToken(strTag, "PropType"))
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With lblLoanIssue(CtlIndex)
                .FontBold = True
                '.Text = ""
                .Caption = ""
            End With
            txtLoanIssue(CtlIndex).Enabled = False
        Case "EDITABLE"
            ' Add 4 spaces for indentation purposes.
            With lblLoanIssue(CtlIndex)
                .Caption = Space(2)
                .FontBold = False
            End With
            txtLoanIssue(CtlIndex).Enabled = True
        Case Else
            'MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            MsgBox LoadResString(gLangOffSet + 603) _
                    & "in Property file!", vbCritical
            Exit Function
    End Select
    ' Set the PROPERTIES for controls.
'    Debug.Assert I <> 12
    With lblLoanIssue(CtlIndex)
        strRet = putToken(strTag, "Visible", "True")
        .Tag = strRet
        .Caption = .Caption & ExtractToken(.Tag, "Prompt")
        If CtlIndex = 0 Then
            .Top = 0
        Else
            .Top = lblLoanIssue(CtlIndex - 1).Top _
                + lblLoanIssue(CtlIndex - 1).Height + CTL_MARGIN
        End If
        .Left = 0
        .Visible = True
    End With
    With txtLoanIssue(CtlIndex)
        .Top = lblLoanIssue(CtlIndex).Top
        .Left = lblLoanIssue(CtlIndex).Left + _
            lblLoanIssue(CtlIndex).Width + CTL_MARGIN
        .Visible = True
        ' Check the LockEdit property.
        strRet = ExtractToken(strTag, "LockEdit")
        If StrComp(strRet, "True", vbTextCompare) = 0 Then
            .Locked = True
        Else
            .Locked = False
        End If
    End With

    ' Get the display type. If its a List or Browse,
    ' then load a combo or a cmd button.
    Dim CmdLoaded As Boolean
    Dim ListLoaded As Boolean
    strPropType = ExtractToken(strTag, "DisplayType")
    Select Case UCase$(strPropType)
        Case "LIST"
            'Load a combo.
            If Not ListLoaded Then
                ListLoaded = True
            Else
                Load cmbLoanIssue(cmbLoanIssue.Count)
            End If
            ' Set the alignment.
            With cmbLoanIssue(cmbLoanIssue.Count - 1)
                '.Index = i
                .Left = txtLoanIssue(i).Left
                .Top = txtLoanIssue(i).Top
                .Width = txtLoanIssue(i).Width
                ' Set it's tab order.
                .TabIndex = txtLoanIssue(i).TabIndex + 1
                ' Update the tag with the text index.
                .Tag = putToken(.Tag, "TextIndex", CStr(i))
                ' Write back this button index to text tag.
                lblLoanIssue(i).Tag = putToken(lblLoanIssue(i).Tag, _
                        "TextIndex", CStr(cmbLoanIssue.Count - 1))
                'txtData(i).Visible = False
                ' If the list data is given, load it.
                Dim List() As String, J As Integer
                Dim strListData As String
                strListData = ExtractToken(strTag, "ListData")
                If strListData <> "" Then
                    ' Break up the data into array elements.
                    GetStringArray strListData, List(), ","
                    cmbLoanIssue(cmbLoanIssue.Count - 1).Clear
                    For J = 0 To UBound(List)
                        cmbLoanIssue(cmbLoanIssue.Count - 1).AddItem List(J)
                    Next
                End If
            End With

        Case "BROWSE"
            'Load a command button.
            If Not CmdLoaded Then CmdLoaded = True _
                    Else Load cmdLoanIssue(cmdLoanIssue.Count)
            With cmdLoanIssue(cmdLoanIssue.Count - 1)
                '.Index = i
                .Width = txtLoanIssue(i).Height
                .Height = .Width
                .Left = txtLoanIssue(i).Left + txtLoanIssue(i).Width - .Width
                .Top = txtLoanIssue(i).Top
                .TabIndex = txtLoanIssue(i).TabIndex + 1
                .ZOrder 0
                
                ' If width and caption are mentioned for
                ' the command button, apply them.
                strTmp = ExtractToken(strTag, "ButtonWidth")
                If strTmp <> "" Then
                    If Val(strTmp) > txtLoanIssue(i).Width / 2 Then
                        ' Restrict the width to half the textbox width.
                        .Width = txtLoanIssue(i).Width / 2
                    ElseIf Val(strTmp) < .Width Then
                        .Width = txtLoanIssue(i).Height
                    Else
                        .Width = Val(strTmp)
                    End If
                End If
                strTmp = ExtractToken(strTag, "ButtonCaption")
                If strTmp <> "" Then
                    .Caption = strTmp
                Else
                    .Caption = "..."
                End If
                ' Update the tag with the text index.
                .Tag = putToken(.Tag, "TextIndex", CStr(i))
                ' Write back this button index to text tag.
                lblLoanIssue(i).Tag = putToken(lblLoanIssue(i).Tag, _
                        "TextIndex", CStr(cmdLoanIssue.Count - 1))
            End With
    End Select

    ' Increment the loop count.
    i = i + 1
Loop
ArrangeLoanIssuePropSheet

' Display today's date, for date field.
i = PropIssueGetIndex("IssueDate")
If i >= 0 Then
    txtLoanIssue(i).Text = FormatDate(gStrDate)
End If

Dim cmbIndex As Integer
    ' Find out the textbox bound to InstalamentMode.
    i = PropIssueGetIndex("InstalMentMode")
    ' Get the combobox index for this text.
    cmbIndex = ExtractToken(lblLoanIssue(i).Tag, "TextIndex")
    cmbLoanIssue(cmbIndex).Clear
    ' Now Laod the Ins MOde
    Dim Count As Integer
    cmbLoanIssue(cmbIndex).AddItem " "
    For Count = 1 To 7
        cmbLoanIssue(cmbIndex).AddItem LoadResString(gLangOffSet + 460 + Count)
    Next
End Function
Private Function LoadLoanTypeProp() As Boolean

'
' Read the data from Loans.ini and load the relevant data.
'

' Check for the existence of the file.
Dim PropFile As String
If gLangOffSet = wis_KannadaOffset Then
    PropFile = App.Path & "\LoansKan.PRP"
Else
    PropFile = App.Path & "\Loans.PRP"
End If
If Dir(PropFile, vbNormal) = "" Then
    'MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox LoadResString(gLangOffSet + 602) _
            & PropFile & "' !", vbExclamation
    Exit Function
End If

' Declare required variables...
Dim strTmp As String
Dim strPropType As String
Dim FirstImgCtl As Boolean
Dim FirstControl As Boolean
Dim i As Integer, CtlIndex As Integer
Dim strRet As String, imgCtlIndex As Integer
FirstControl = True
FirstImgCtl = True
Dim strTag As String

' Read all the prompts and load accordingly...
Do
    ' Read a line.
    strTag = ReadFromIniFile("LoanTypes", _
                "Prop" & i + 1, PropFile)
    If strTag = "" Then Exit Do

    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load lblLoanScheme(lblLoanScheme.Count)
        Load txtLoanScheme(txtLoanScheme.Count)
    End If
    CtlIndex = lblLoanScheme.Count - 1

    ' Get the property type.
    strPropType = ExtractToken(strTag, "PropType")
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With lblLoanScheme(CtlIndex)
                .FontBold = True
                '.Text = ""
                .Caption = ""
            End With
            txtLoanScheme(CtlIndex).Enabled = False

        Case "EDITABLE"
            ' Add 4 spaces for indentation purposes.
            With lblLoanScheme(CtlIndex)
                .Caption = Space(2)
                .FontBold = False
            End With
            txtLoanScheme(CtlIndex).Enabled = True
        Case Else
            
            'MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            MsgBox LoadResString(gLangOffSet + 603) _
                    & "in Property file!", vbCritical
            Exit Function

    End Select

    ' Set the PROPERTIES for controls.
    With lblLoanScheme(CtlIndex)
        strRet = putToken(strTag, "Visible", "True")
        .Tag = strRet
        .Caption = .Caption & ExtractToken(.Tag, "Prompt")
        If CtlIndex = 0 Then
            .Top = 0
        Else
            .Top = lblLoanScheme(CtlIndex - 1).Top _
                + lblLoanScheme(CtlIndex - 1).Height + CTL_MARGIN
        End If
        .Left = 0
        .Visible = True
    End With
    With txtLoanScheme(CtlIndex)
        .Top = lblLoanScheme(CtlIndex).Top
        .Left = lblLoanScheme(CtlIndex).Left + _
            lblLoanScheme(CtlIndex).Width + CTL_MARGIN
        .Visible = True
        ' Check the LockEdit property.
        strRet = ExtractToken(strTag, "LockEdit")
        If StrComp(strRet, "True", vbTextCompare) = 0 Then
            .Locked = True
        Else
            .Locked = False
        End If
        '.Enabled = True
    End With

    ' Get the display type. If its a List or Browse,
    ' then load a combo or a cmd button.
    Dim CmdLoaded As Boolean
    Dim ListLoaded As Boolean
    strPropType = ExtractToken(strTag, "DisplayType")
    Select Case UCase$(strPropType)
        Case "LIST"
            'Load a combo.
            If Not ListLoaded Then
                ListLoaded = True
            Else
                Load cmbLoanScheme(cmbLoanScheme.Count)
            End If
            ' Set the alignment.
            With cmbLoanScheme(cmbLoanScheme.Count - 1)
                '.Index = i
                .Left = txtLoanScheme(i).Left
                .Top = txtLoanScheme(i).Top
                .Width = txtLoanScheme(i).Width
                ' Set it's tab order.
                .TabIndex = txtLoanScheme(i).TabIndex + 1
                ' Update the tag with the text index.
                .Tag = putToken(.Tag, "TextIndex", CStr(i))
                ' Write back this button index to text tag.
                lblLoanScheme(i).Tag = putToken(lblLoanScheme(i).Tag, _
                        "TextIndex", CStr(cmbLoanScheme.Count - 1))
                'txtData(i).Visible = False
                ' If the list data is given, load it.
                Dim List() As String, J As Integer
                Dim strListData As String
                strListData = ExtractToken(strTag, "ListData")
                If strListData <> "" Then
                    ' Break up the data into array elements.
                    GetStringArray strListData, List(), ","
                    cmbLoanScheme(cmbLoanScheme.Count - 1).Clear
                    For J = 0 To UBound(List)
                        cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem List(J)
                    Next
                End If
                strListData = ExtractToken(strTag, "DataSource")
                If UCase(Trim$(strListData)) = "CATEGORY" Then
                    cmbLoanScheme(cmbLoanScheme.Count - 1).Clear
                    cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem LoadResString(gLangOffSet + 440)
                    cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem LoadResString(gLangOffSet + 441)
                End If
                If UCase(Trim$(strListData)) = "TERMTYPE" Then
                    cmbLoanScheme(cmbLoanScheme.Count - 1).Clear
                    cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem LoadResString(gLangOffSet + 222)
                    cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem LoadResString(gLangOffSet + 223)
                    cmbLoanScheme(cmbLoanScheme.Count - 1).AddItem LoadResString(gLangOffSet + 224)
                End If
            End With

        Case "BROWSE"
            'Load a command button.
            If Not CmdLoaded Then
                CmdLoaded = True
            Else
                Load cmdLoanScheme(cmdLoanScheme.Count)
            End If
            With cmdLoanScheme(cmdLoanScheme.Count - 1)
                '.Index = i
                .Width = txtLoanScheme(i).Height
                .Height = .Width
                .Left = txtLoanScheme(i).Left + txtLoanScheme(i).Width - .Width
                .Top = txtLoanScheme(i).Top
                .TabIndex = txtLoanScheme(i).TabIndex + 1
                .ZOrder 0
                ' Update the tag with the text index.
                .Tag = putToken(.Tag, "TextIndex", CStr(i))
                ' Write back this button index to text tag.
                lblLoanScheme(i).Tag = putToken(lblLoanScheme(i).Tag, _
                        "TextIndex", CStr(cmdLoanScheme.Count - 1))
            End With

    End Select

    ' Increment the loop count.
    i = i + 1
Loop

ArrangeLoanSchemePropSheet

' display today's date in CreateDate field.
i = PropSchemeGetIndex("CreateDate")
txtLoanScheme(i).Text = FormatDate(gStrDate)

End Function
Private Sub PropScrollLoanIssueWindow(Ctl As Control)

If picLoanIssueSlider.Top + Ctl.Top + Ctl.Height > picLoanIssueViewPort.ScaleHeight Then
    ' The control is below the viewport.
    Do While picLoanIssueSlider.Top + Ctl.Top + Ctl.Height > _
                    picLoanIssueViewPort.ScaleHeight
        ' scroll down by one row.
        With vscLoanIssue
            If .value + .SmallChange <= .Max Then
                    .value = .value + .SmallChange
            Else
                    .value = .Max
            End If
        End With
    Loop

ElseIf picLoanIssueSlider.Top + Ctl.Top < 0 Then
    ' The control is above the viewport.
    ' Keep scrolling until it is in viewport.
    Do While picLoanIssueSlider.Top + Ctl.Top < 0
        With vscLoanIssue
            If .value - .SmallChange >= .Min Then
                .value = .value - .SmallChange
            Else
                .value = .Min
            End If
        End With
    Loop
End If

End Sub
Private Sub ZZZ_ReportDailyCashbook(grd As MSFlexGrid, Optional SchemeType As String, Optional stdate As String, Optional enddate As String, Optional stAmt As String, Optional endAmt As String)

' Declare variables...
Dim Lret As Long
Dim rptRS As Recordset
Dim PrevDate As String
Dim TotAmt As Currency

' Setup error handler.
On Error GoTo Err_Line

' Display status.
RaiseEvent SetStatus("Querying loan details...")
If Trim$(SchemeType) = "" Then
    gDbTrans.SQLStmt = " Select sum(Amount) as TotalAmount , TransDate, TransType " & _
            " From LoanTrans"
Else
    gDbTrans.SQLStmt = " Select sum(Amount) as TotalAmount , TransDate, TransType, SchemeName " & _
            " From LoanTrans a, LoanMaster b, LoanTypes c Where  a.LoanId = b.LoanId " & _
            " And b.SchemeId = c.SchemeId AND c.SchemeName = '" & SchemeType & "' "
End If

#If Junk Then
If Trim$(stdate) <> "" Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate >= #" & FormatDate(stdate) & "#"
End If
If Trim$(enddate) <> "" Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.transdate <= #" & FormatDate(enddate) & "#"
End If
If Trim$(stAmt) <> "" Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.amount <= " & Val(stAmt)
End If
If Trim$(endAmt) <> "" Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND a.amount <= " & Val(endAmt)
End If
#End If

gDbTrans.SQLStmt = gDbTrans.SQLStmt & " Group By TransDate, TransType"
If Trim$(SchemeType) <> "" Then
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & " , SchemeName"
End If

Lret = gDbTrans.SQLFetch
If Lret < 0 Then
    ' Error in database.
    MsgBox "Error retrieving loan details.", vbCritical, wis_MESSAGE_TITLE
    GoTo Exit_Line
ElseIf Lret = 0 Then
    'MsgBox "No records...", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 278), vbInformation, wis_MESSAGE_TITLE
End If
Set rptRS = gDbTrans.Rst.Clone

' Populate the record set.
rptRS.MoveLast
rptRS.MoveFirst

' Update the status message.
RaiseEvent SetStatus("Formatting loan results...")

' Initialize the grid.
With grd
    .Visible = False
    .Clear
    .Rows = rptRS.RecordCount + 1
    .FixedRows = 1
    .FixedCols = 0
    .FormatString = "Date   |Loan scheme name   |Loan issued |Repayment "
End With

' Fill the rows
Dim TransType  As wisTransactionTypes
Dim Transdate As Date
grd.Row = 0
With rptRS
    Do While Not .EOF
        ' Set the row.
        If .fields("Transdate") <> Transdate Then
            grd.Row = grd.Row + 1 '.AbsolutePosition + 1
            Transdate = .fields("TransDate")
        End If

        ' Fill the transaction date.
        grd.Col = 0
        grd.Text = FormatField(.fields("transdate"))

        ' Fill the loan scheme name.
        grd.Col = 1
        If Trim$(SchemeType) <> "" Then
            grd.Text = .fields("schemename")
        End If

        ' Fill the loan amount.  If the transaction type is
        ' -1, show it in the loan amount column, else
        ' show it in the repaid amount column.
        If .fields("TransType") = -1 Then
            grd.Col = 2
        Else
            grd.Col = 3
        End If
        grd.Text = .fields("TotalAmount")

        ' Move to next row.
        .MoveNext
    Loop
End With

' Display the grid.
RaiseEvent SetStatus("")
grd.Visible = True
frmLoanView.Caption = "INDEX-2000  [List of payments made...]"
frmLoanView.Show vbModal, Me

Exit_Line:
    RaiseEvent SetStatus("")
    Unload frmLoanView
    Exit Sub

Err_Line:
    If Err Then
        'MsgBox "ReportLoansIssued: " & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 733) & vbCrLf _
            & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
Resume
    GoTo Exit_Line


End Sub

' Checks if a specified scheme name is already present in the database.
Private Function SchemeExists(strSchemeName As String) As Boolean
' Setup error handler.
On Error GoTo Err_Line

Dim Lret As Long

' Query the loan database.
gDbTrans.SQLStmt = "SELECT * FROM loantypes WHERE " _
        & "schemeName = '" & strSchemeName & "'"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then Exit Function
SchemeExists = True

Err_Line:
    If Err Then
        MsgBox "SchemeExists: " & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 734) & Err.Description, vbCritical
    End If
End Function

'Checks if loans are issued for a given loan type.
Private Function SchemeInUse(strSchemeName As String) As Boolean
' Setup error handler.
On Error GoTo Err_Line

Dim Lret As Long

' Query the loan database.
gDbTrans.SQLStmt = "SELECT * FROM loanmaster WHERE " _
        & "schemeID = (SELECT SchemeID FROM loantypes WHERE " _
        & "schemename = '" & strSchemeName & "')"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then Exit Function
SchemeInUse = True

Err_Line:
    If Err Then
        MsgBox "SchemeExists: " & Err.Description, vbCritical
        'MsgBox LoadResString(gLangOffSet + 734) & Err.Description, vbCritical
    End If

End Function
Private Function SchemeLoad(lSchemeID As Long) As Boolean
On Error GoTo SchemeLoad_error

' Procedure variables...
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*
Dim Lret As Long, i As Integer
Dim Rst As Recordset
Dim strDataSource As String
Dim strTmpIndex As String
' -*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*

' Check if the specified loanID is a valid number.
If lSchemeID < 0 Then GoTo Exit_Line

' Query the database for the loan scheme.
gDbTrans.SQLStmt = "SELECT * FROM loantypes WHERE schemeid = " & lSchemeID
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    'MsgBox "Loan scheme does not exists !", _
            vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 735), _
            vbExclamation, gAppName & " - Error"
    GoTo Exit_Line
End If
Set Rst = gDbTrans.Rst.Clone

' Display the scheme details to the user interface...
For i = 0 To txtLoanScheme.Count - 1
    With txtLoanScheme(i)
        ' Get the bound data source name.
        strDataSource = ExtractToken(lblLoanScheme(i).Tag, "DataSource")
        ' Fill the data.
        Select Case UCase$(strDataSource)
            Case "LOANNAME"
                .Text = FormatField(gDbTrans.Rst("SchemeName"))
                m_SchemeName = .Text
            Case "CATEGORY"
                ' Get the index of the combo control for this field.
                strTmpIndex = ExtractToken(lblLoanScheme(i).Tag, "TextIndex")
                If strTmpIndex <> "" Then
                    With cmbLoanScheme(strTmpIndex)
                        .ListIndex = Val(FormatField(gDbTrans.Rst("Category"))) - 1
                    End With
                    .Text = cmbLoanScheme(strTmpIndex).Text
                    .Locked = True
                End If
            Case "TERMTYPE"
                ' Get the index of the combo control for this field.
                strTmpIndex = ExtractToken(lblLoanScheme(i).Tag, "TextIndex")
                If strTmpIndex <> "" Then
                    With cmbLoanScheme(strTmpIndex)
                        .ListIndex = Val(FormatField(gDbTrans.Rst("TermType"))) - 1
                    End With
                    .Text = cmbLoanScheme(strTmpIndex).Text
                End If
            Case "INTERESTRATE"
                .Text = FormatField(gDbTrans.Rst("InterestRate"))
            Case "PENALINTERESTRATE"
                .Text = FormatField(gDbTrans.Rst("PenalInterestRate"))
            Case "MAXREPAYMENTTIME"
                .Text = FormatField(gDbTrans.Rst("MaxRepaymentTime"))
            Case "INSURANCEFEE"
                .Text = FormatField(gDbTrans.Rst("InsuranceFee"))
            Case "LEGALFEE"
                .Text = FormatField(gDbTrans.Rst("LegalFee"))
            Case "CREATEDATE"
                .Text = FormatField(gDbTrans.Rst("CreateDate"))
            Case "DESCRIPTION"
                .Text = FormatField(gDbTrans.Rst("Description"))
        End Select
    End With
Next

SchemeLoad = True
Exit_Line:
    Exit Function

SchemeLoad_error:
    If Err Then
        'MsgBox "SchemeLoad: " & Err.Description, vbCritical
        MsgBox LoadResString(gLangOffSet + 736) & Err.Description, vbCritical
    End If
Resume
    GoTo Exit_Line

End Function
Private Function SchemeSave() As Boolean

' Setup Error handler
On Error GoTo SchemeSave_Error

' Variables used in this routine...
' ---------------------------------------
Dim txtindex As Byte
Dim nRet As Integer, Lret As Long
Dim ExistingScheme As Boolean
Dim inTransaction As Boolean
Dim NewSchemeId As Long
' ---------------------------------------

' Check for valid scheme name.
txtindex = PropSchemeGetIndex("LoanName")
With txtLoanScheme(txtindex)
    ' Warn if the name not specified.
    If Trim$(.Text) = "" Then
        'MsgBox "Specify a name for the loan scheme.", vbInformation
        MsgBox LoadResString(gLangOffSet + 737), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Check the loan category.
txtindex = PropSchemeGetIndex("Category")
With txtLoanScheme(txtindex)
    If Trim$(.Text) = "" Then
        'MsgBox "Select a loan category.", vbInformation
        MsgBox LoadResString(gLangOffSet + 738), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Check the loan term type.
txtindex = PropSchemeGetIndex("TermType")
With txtLoanScheme(txtindex)
    If Trim$(.Text) = "" Then
        'MsgBox "Select a loan term type.", vbInformation
        MsgBox LoadResString(gLangOffSet + 739), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Repay time limit.
txtindex = PropSchemeGetIndex("MaxRepaymentTime")
With txtLoanScheme(txtindex)
    If (.Text) = "" Then
        'MsgBox "Specify the maximum repayment time limit " _
                & "for this type of loan.", vbInformation
        MsgBox LoadResString(gLangOffSet + 740), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for maximum repayment time limit.", vbInformation
        MsgBox LoadResString(gLangOffSet + 740), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Check for valid interest rate.
txtindex = PropSchemeGetIndex("InterestRate")
With txtLoanScheme(txtindex)
    If (.Text) = "" Then
        'MsgBox "Specify the rate of interest to be applied " _
            & "for this loan.", vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 741), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for rate of interest.", vbInformation
        MsgBox LoadResString(gLangOffSet + 505), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Check for valid interest rate.
txtindex = PropSchemeGetIndex("PenalInterestRate")
With txtLoanScheme(txtindex)
    If (.Text) = "" Then
        'MsgBox "Specify the penal interest to be applied " _
            & "for this loan.", vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 742), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
    If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
        'MsgBox "Invalid value for penal interest.", vbInformation
        MsgBox LoadResString(gLangOffSet + 743), vbInformation
        ActivateTextBox txtLoanScheme(txtindex)
        GoTo Exit_Line
    End If
End With

' Check for Insurance fee
txtindex = PropSchemeGetIndex("InsuranceFee")
With txtLoanScheme(txtindex)
    If (.Text) <> "" Then
        ' Check if it is a vaid numeric data.
        If Not IsNumeric(.Text) Then
            'MsgBox "Invalid value for Insurance fee.", vbInformation
            MsgBox LoadResString(gLangOffSet + 744), vbInformation
            ActivateTextBox txtLoanScheme(txtindex)
            GoTo Exit_Line
        End If
        
        ' Check for permissible range.
        If Val(.Text) < 0 Then
            'MsgBox "Invalid value for Insurance fee.", vbInformation
            MsgBox LoadResString(gLangOffSet + 744), vbInformation
            ActivateTextBox txtLoanScheme(txtindex)
            GoTo Exit_Line
        End If
    End If
End With

' Check for Legal fee
txtindex = PropSchemeGetIndex("Legalfee")
With txtLoanScheme(txtindex)
    If (.Text) <> "" Then
        ' Check if it is a vaid numeric data.
        If Not IsNumeric(.Text) Then
            'MsgBox "Invalid value for legal fee.", vbInformation
            MsgBox LoadResString(gLangOffSet + 745), vbInformation
            ActivateTextBox txtLoanScheme(txtindex)
            GoTo Exit_Line
        End If
        ' Check for permissible range.
        If Val(.Text) < 0 Then
            'MsgBox "Invalid value for legal fee.", vbInformation
            MsgBox LoadResString(gLangOffSet + 745), vbInformation
            ActivateTextBox txtLoanScheme(txtindex)
            GoTo Exit_Line
        End If
    
    End If
End With

txtindex = PropSchemeGetIndex("LoanName")
With txtLoanScheme(txtindex)
    ' Check for duplicate name.
    If SchemeExists(.Text) Then
        ExistingScheme = True
        ' Check if loans are issued based on this loan type.
        ' If it is not used, prompt if the user wants to modify the scheme.
        If SchemeInUse(.Text) Then
            'MsgBox "The scheme name '" & .Text & "' is in use.  " _
                & "Do You want to continue?.", vbInformation
            If MsgBox(LoadResString(gLangOffSet + 734) & vbCrLf & LoadResString(gLangOffSet + 747), _
                    vbInformation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
                ActivateTextBox txtLoanScheme(txtindex)
                GoTo Exit_Line
            End If
        Else
            ' Confirm, if the user wants to update this loan type.
            'nRet = MsgBox("The scheme '" & .Text & "' already exists." & vbCrLf _
                    & "Do you want to modify it?", vbYesNo + vbQuestion)
            nRet = MsgBox(LoadResString(gLangOffSet + 734) & vbCrLf _
                    & LoadResString(gLangOffSet + 747), vbYesNo + vbQuestion)
            If nRet = vbNo Then GoTo Exit_Line
        End If
    End If
End With

' Save the loan type.
If ExistingScheme Or m_SchemeLoaded Then
    'Now get the SchemeId of which scheme updating
    gDbTrans.SQLStmt = "Select SchemeID From LoanTypes " & _
        " where SchemeName = " & AddQuotes(m_SchemeName, True)
    If gDbTrans.SQLFetch < 1 Then GoTo Exit_Line
    NewSchemeId = FormatField(gDbTrans.Rst("SchemeId"))
    
    If Not gDbTrans.BeginTrans Then GoTo Exit_Line
    inTransaction = True
    gDbTrans.SQLStmt = "UPDATE Loantypes SET category = " _
            & GetLoanCategory & ", maxrepaymenttime = " & _
            PropSchemeGetVal("MaxRepaymentTime") _
            & ", InterestRate = " & PropSchemeGetVal("InterestRate") _
            & ", PenalInterestRate = " & PropSchemeGetVal("PenalInterestRate") _
            & " Where SchemeId = " & NewSchemeId
   
Else
    ' Get a new scheme id for this loan.
    gDbTrans.SQLStmt = "SELECT MAX(SchemeID) FROM loantypes"
    Lret = gDbTrans.SQLFetch
    If Lret <= 0 Then
        MsgBox "SchemeSave: Internal error.  Error querying " _
            & "the database.", vbCritical
        GoTo Exit_Line
    End If
    NewSchemeId = Val(FormatField(gDbTrans.Rst(0))) + 1
    ' Begin the transaction.
    If Not gDbTrans.BeginTrans Then GoTo Exit_Line
    inTransaction = True
    gDbTrans.SQLStmt = "INSERT INTO loantypes (SchemeID," _
            & "SchemeName, Category, TermType, MaxRepaymentTime, " _
            & "InterestRate, PenalInterestRate, " _
            & "InsuranceFee, LegalFee, Description, CreateDate) " _
            & "Values (" & NewSchemeId & ", '" & PropSchemeGetVal("LoanName") _
            & "', " & GetLoanCategory & _
            ", " & GetLoanTerm & ", " & _
            PropSchemeGetVal("MaxRepaymentTime") _
            & ", " & PropSchemeGetVal("InterestRate") _
            & ", " & PropSchemeGetVal("PenalInterestRate") _
            & ", " & Val(PropSchemeGetVal("InsuranceFee")) & ", " _
            & Val(PropSchemeGetVal("LegalFee")) & ", '" _
            & PropSchemeGetVal("Description") _
            & "', '" & PropSchemeGetVal("CreateDate") & "')"
    
End If

If Not gDbTrans.SQLExecute Then GoTo Exit_Line

' Commit the transaction.
gDbTrans.CommitTrans
SchemeSave = True
inTransaction = False
'MsgBox "Saved the loan scheme.", vbInformation
MsgBox LoadResString(gLangOffSet + 748), vbInformation

Exit_Line:

     If inTransaction Then gDbTrans.RollBack
    Exit Function

SchemeSave_Error:
    If Err Then
        MsgBox "SchemeSave: " & Err.Description, vbCritical
'        MsgBox LoadResString(gLangOffSet + 749) & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Function

Private Sub PropScrollLoanSchemeWindow(Ctl As Control)

If picLoanSchemeSlider.Top + Ctl.Top + Ctl.Height > picLoanSchemeViewPort.ScaleHeight Then
    ' The control is below the viewport.
    Do While picLoanSchemeSlider.Top + Ctl.Top + Ctl.Height > _
                    picLoanSchemeViewPort.ScaleHeight
        ' scroll down by one row.
        With vscLoanScheme
            If .value + .SmallChange <= .Max Then
                .value = .value + .SmallChange
            Else
                .value = .Max
            End If
        End With
    Loop

ElseIf picLoanSchemeSlider.Top + Ctl.Top < 0 Then
    ' The control is above the viewport.
    ' Keep scrolling until it is in viewport.
    Do While picLoanSchemeSlider.Top + Ctl.Top < 0
        With vscLoanScheme
            If .value - .SmallChange >= .Min Then
                .value = .value - .SmallChange
            Else
                .value = .Min
            End If
        End With
    Loop
End If

End Sub

Private Sub PropSetIssueDescription(Ctl As Control)
' Extract the description title.
lblLoanIssueHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblLoanIssueDesc.Caption = ExtractToken(Ctl.Tag, "Description")

End Sub
Private Sub PropSetSchemeDescription(Ctl As Control)

' Extract the description title.
lblLoanSchemeHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblLoanSchemeDesc.Caption = ExtractToken(Ctl.Tag, "Description")
End Sub

Private Sub ShowPassBookPage()
#If COMMENTED Then
Dim i As Integer
grd.Visible = False
grd.Rows = 1
grd.Rows = 11
For i = 1 To 10
    grd.Row = i
    grd.Col = 0
    grd.Text = FormatField(m_rstPassBook("TransDate"))
    
    grd.Col = 1
    grd.Text = FormatField(m_rstPassBook("Particulars"))
    
    grd.Col = 2
    If Val(FormatField(m_rstPassBook("ChequeNo"))) > 0 Then
        grd.Text = FormatField(m_rstPassBook("ChequeNo"))
    End If
    
        
    
    If FormatField(m_rstPassBook("TransType")) < 0 Then
        grd.Col = 3
    Else
        grd.Col = 4
    End If
    
    grd.Text = FormatField(m_rstPassBook("Amount"))
   
    grd.Col = 5
    grd.Text = FormatField(m_rstPassBook("Balance"))
    
    
    m_rstPassBook.MoveNext
    If m_rstPassBook.EOF Then
        m_rstPassBook.MoveLast
        Exit For
    End If
    
Next i

grd.Visible = True
#End If
End Sub
' Returns the number of items that are visible for a control array.
' Looks in the control's tag for visible property, rather than
' depend upon the control's visible property for some obvious reasons.
Private Function VisibleCountLoanScheme() As Integer
On Error GoTo Err_Line
Dim i As Integer
Dim strVisible As String

For i = 0 To lblLoanScheme.Count - 1
    strVisible = ExtractToken(lblLoanScheme(i).Tag, "Visible")
    If StrComp(strVisible, "True", vbTextCompare) = 0 Then
        VisibleCountLoanScheme = VisibleCountLoanScheme + 1
    End If
Next

Err_Line:
End Function

' Returns the number of items that are visible for a control array.
' Looks in the control's tag for visible property, rather than
' depend upon the control's visible property for some obvious reasons.
Private Function VisibleCountLoanIssue() As Integer
On Error GoTo Err_Line
Dim i As Integer
Dim strVisible As String

For i = 0 To lblLoanIssue.Count - 1
    strVisible = ExtractToken(lblLoanIssue(i).Tag, "Visible")
    If StrComp(strVisible, "True", vbTextCompare) = 0 Then
        VisibleCountLoanIssue = VisibleCountLoanIssue + 1
    End If
Next

Err_Line:
End Function


Private Sub chkCaste_Click()
If chkCaste.value = vbChecked Then
   cmbCaste.Enabled = True
   cmbCaste.BackColor = vbWhite
Else
   cmbCaste.Enabled = False
   cmbCaste.BackColor = wisGray
   cmbCaste.ListIndex = -1
End If
End Sub

Private Sub chkPlace_Click()
If chkPlace.value = vbChecked Then
   cmbPlace.Enabled = True
   cmbPlace.BackColor = vbWhite
Else
   cmbPlace.Enabled = False
   cmbPlace.BackColor = wisGray
   cmbPlace.ListIndex = -1
End If
End Sub

Private Sub cmbLoanIssue_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmbLoanIssue_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtindex As String
txtindex = ExtractToken(cmbLoanIssue(Index).Tag, "TextIndex")
If txtindex <> "" Then
    txtLoanIssue(Val(txtindex)).Text = cmbLoanIssue(Index).Text
End If

End Sub

Private Sub cmbLoanScheme_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmbloanscheme_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtindex As String
txtindex = ExtractToken(cmbLoanScheme(Index).Tag, "TextIndex")
If txtindex <> "" Then
    txtLoanScheme(Val(txtindex)).Text = cmbLoanScheme(Index).Text
End If

End Sub


Private Sub cmdClearScheme_Click()
PropSchemeClear
End Sub


Private Sub cmdEndDate_Click()
With Calendar
    .Left = Me.Left + fraReports.Left + fraDateRange.Left + cmdEndDate.Left - .Width
    .Top = Me.Top + fraReports.Top + fraDateRange.Top + cmdEndDate.Top + 300
    .SelDate = txtEndDate.Text
    .Show vbModal, Me
    If .SelDate <> "" Then txtEndDate.Text = Calendar.SelDate
End With

End Sub

Private Sub cmdInstalment_Click()
If m_LoanId = 0 Then Exit Sub
frmLoanInst.p_AccId = m_AccNo
frmLoanInst.p_LoanId = m_LoanId
Load frmLoanInst
frmLoanInst.Show vbModal
Call LoanLoad(m_LoanId)
End Sub

Private Sub cmdIntdate_Click()
With Calendar
    .Left = Me.Left + Me.fraRepayments.Left + Me.cmdRepayDate.Left
    .Top = Me.Top + Me.fraRepayments.Top + Me.cmdRepayDate.Top
    If DateValidate(txtIntBalance.Text, "/", True) Then
      .SelDate = txtIntBalance.Text
   End If
    .Show vbModal
    If .SelDate <> "" Then txtIntBalance.Text = .SelDate
End With


End Sub

Private Sub cmdLoad_Click()
Me.MousePointer = vbHourglass

' Validate the input data.
If Not IsNumeric(txtAccNo.Text) Then
    'MsgBox "Enter numeric data for Member ID.", _
            vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 510), _
            vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtAccNo
    GoTo Exit_Line
End If
LoadRepaymentTab Val(txtAccNo.Text)

' If loans were found for this member,
' select the first loantab by default and
' fill in the details.
If tabLOans.Tabs.Count > 0 Then
    cmdInstalment.Enabled = True
    tabLOans.Tabs(1).Selected = True
    ' Set the default date for repayment.
'    txtRepayDate.Text = FormatDate(gStrDate)
   End If

txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)


Exit_Line:
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdLoanIssue_Click(Index As Integer)
Me.MousePointer = vbHourglass
' Variables of this routine...
Dim txtindex As String
Dim strField As String
Dim Lret As Long
Dim NameStr As String

' Check to which text index it is mapped.
txtindex = ExtractToken(cmdLoanIssue(Index).Tag, "TextIndex")

' Extract the Bound field name.
strField = ExtractToken(lblLoanIssue(Val(txtindex)).Tag, "DataSource")

Select Case UCase$(strField)
    Case "MEMBERID"
        ' Query for member details...
'        gDBTrans.SQLStmt = "SELECT accid, firstname ," _
                & "middlename , lastname , " _
                & "profession FROM NameTab, mmMaster WHERE " _
                & "nametab.customerID = mmMaster.customerID"
 gDbTrans.SQLStmt = "SELECT accid, title + space(1) + firstname " _
               & "+ space(1) + middlename + space(1) + lastname AS Name, " _
              & "profession FROM NameTab, mmMaster WHERE " _
             & "nametab.customerID = mmMaster.customerID"
        'now Check Whether He Want Search   any particular name
         'NameStr = InputBox("Enter customer name , You want search", "Name Search")
         NameStr = InputBox(LoadResString(gLangOffSet + 785), "Name Search")
         If NameStr <> "" Then
             gDbTrans.SQLStmt = gDbTrans.SQLStmt & " And ( FirstNAme like '" & NameStr & "*' " & _
                                         " Or MiddleName like '" & NameStr & "*' Or LAstName like '" & NameStr & "*' )"
         End If
             gDbTrans.SQLStmt = gDbTrans.SQLStmt & " Order By IsciName"
         
         Lret = gDbTrans.SQLFetch
        If Lret <= 0 Then
            MsgBox LoadResString(gLangOffSet + 569), vbInformation, wis_MESSAGE_TITLE
              'MsgBox "No members present. Create members using the " _
              & "members module.", vbInformation, wis_MESSAGE_TITLE
            GoTo ExitLine
        End If
        'Fill the details to report dialog and display it.
        If m_LookUp Is Nothing Then
            Set m_LookUp = New frmLookUp
        End If
        If Not FillView(m_LookUp.lvwReport, gDbTrans.Rst, True) Then
            MsgBox LoadResString(gLangOffSet + 562), _
                    vbCritical, wis_MESSAGE_TITLE
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical , wis_MESSAGE_TITLE
            GoTo ExitLine
        End If
        With m_LookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
            .lvwReport.ColumnHeaders(2).Width = 3750
            '.lvwReport.ColumnHeaders(3).Width = 3750
            .Title = "Select the member..."
            .m_SelItem = ""
            .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem <> "" Then
                txtLoanIssue(txtindex).Text = .lvwReport.SelectedItem.Text
                txtLoanIssue(txtindex + 1).Text = .lvwReport.SelectedItem.SubItems(1)
            End If
        End With

    Case "LOANNAME"
        ' Query for Loan scheme details...
        gDbTrans.SQLStmt = "SELECT * FROM loantypes Order by Category,SchemeName"
        Lret = gDbTrans.SQLFetch
        If Lret <= 0 Then GoTo ExitLine
        'Fill the details to report dialog and display it.
        If m_LookUp Is Nothing Then
            Set m_LookUp = New frmLookUp
        End If
        If Not FillMembers(m_LookUp.lvwReport, gDbTrans.Rst, True) Then
            MsgBox LoadResString(gLangOffSet + 521), _
                    vbCritical, wis_MESSAGE_TITLE
            'MsgBox "Error loading loan schemes.", _
                    vbCritical, wis_MESSAGE_TITLE
            GoTo ExitLine
        End If
        With m_LookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            .Title = "Select the loan scheme..."
            .m_SelItem = ""
            .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem <> "" Then
                txtLoanIssue(txtindex).Text = .lvwReport.SelectedItem.Text
                ' Move the record pointer to the selected record.
                gDbTrans.Rst.AbsolutePosition = .lvwReport.SelectedItem.Index - 1
                txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, _
                            "SchemeID", FormatField(gDbTrans.Rst("SchemeID")))
                Dim tempIndex As Byte
                      tempIndex = PropIssueGetIndex("InterestRate")
                      txtLoanIssue(tempIndex).Text = FormatField(gDbTrans.Rst("InterestRate"))
                      txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, _
                            "InterestRate", FormatField(gDbTrans.Rst("InterestRate")))
                      
                      tempIndex = PropIssueGetIndex("PenalInterestRate")
                      txtLoanIssue(tempIndex).Text = FormatField(gDbTrans.Rst("PenalInterestRate"))
                      txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, _
                            "PenalInterestRate", FormatField(gDbTrans.Rst("PenalInterestRate")))
            Else
                txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, _
                            "SchemeID", "")
            End If
        End With

    Case "LOANDUEDATE"
        With Calendar
            .Left = txtLoanIssue(txtindex).Left + Me.Left _
                    + picLoanIssueViewPort.Left + fraLoanAccounts.Left + 50
            .Top = Me.Top + txtLoanIssue(txtindex).Top _
                    + picLoanIssueSlider.Top + picLoanIssueViewPort.Top _
                    + fraLoanAccounts.Top + 300
            .Width = txtLoanIssue(txtindex).Width
            .Height = .Width
            .SelDate = txtLoanIssue(txtindex).Text
            .Show vbModal, Me
            If .SelDate <> "" Then txtLoanIssue(txtindex).Text = .SelDate
        End With

    Case "ISSUEDATE"
        With Calendar
            .Left = txtLoanIssue(txtindex).Left + Me.Left _
                    + picLoanIssueViewPort.Left + fraLoanAccounts.Left + 50
            .Top = Me.Top + txtLoanIssue(txtindex).Top _
                    + picLoanIssueSlider.Top + picLoanIssueViewPort.Top _
                    + fraLoanAccounts.Top + 300
            .Width = txtLoanIssue(txtindex).Width
            .Height = .Width
            .SelDate = txtLoanIssue(txtindex).Text
            .Show vbModal, Me
            If .SelDate <> "" Then txtLoanIssue(txtindex).Text = .SelDate
        End With

    Case "GUARANTOR1", "GUARANTOR2"
        ' Query for member details...
        gDbTrans.SQLStmt = "SELECT accid, title + space(1) + " _
                & "firstname + space(1) + middlename" _
                & " + space(1) + lastname AS name, profession FROM NameTab, " _
                & "mmMaster WHERE nametab.customerID = mmMaster.customerID"
        
        'now Check Whether He Want Search   any particular name
         'NameStr = InputBox("Enter customer name , You want search", "Name Search")
         NameStr = InputBox(LoadResString(gLangOffSet + 785), "Name Search")
         If NameStr <> "" Then
             gDbTrans.SQLStmt = gDbTrans.SQLStmt & " And ( FirstNAme like '" & NameStr & "*' " & _
                                         " Or MiddleName like '" & NameStr & "*' Or LAstName like '" & NameStr & "*' )"
         End If
             gDbTrans.SQLStmt = gDbTrans.SQLStmt & " Order By IsciName"

        Lret = gDbTrans.SQLFetch
        If Lret <= 0 Then
            MsgBox LoadResString(gLangOffSet + 569), vbInformation, wis_MESSAGE_TITLE
            'MsgBox "No members present. Create members using the members module.", vbInformation, wis_MESSAGE_TITLE
            GoTo ExitLine
        End If
        'Fill the details to report dialog and display it.
        If m_LookUp Is Nothing Then
            Set m_LookUp = New frmLookUp
        End If
        If Not FillView(m_LookUp.lvwReport, gDbTrans.Rst, True) Then
           MsgBox LoadResString(gLangOffSet + 562), _
                    vbCritical, wis_MESSAGE_TITLE
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
           GoTo ExitLine
        End If
        With m_LookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
            .lvwReport.ColumnHeaders(2).Width = 3750
            '.lvwReport.ColumnHeaders(3).Width = 3750
            .Title = "Select the Guarantor..."
            .m_SelItem = ""
            .Show vbModal, Me
            'If .Status = wis_OK Then
            If .m_SelItem <> "" Then
                txtLoanIssue(txtindex).Text = .lvwReport.SelectedItem.SubItems(1)
                ' Add the guarantorID to the tag property.
                txtLoanIssue(txtindex).Tag = putToken(txtLoanIssue(txtindex).Tag, "GuarantorID", .lvwReport.SelectedItem.Text)
            End If
        End With
        Case "SANCTIONAMOUNT"    'cmdsaveloan   cmdloanupdate
                If Val(txtLoanIssue(PropIssueGetIndex("LoanInstalment")).Text) > 1 Then
                        frmLoanName.P_TxtCount = Val(txtLoanIssue(PropIssueGetIndex("LoanInstalment")).Text)
                        Load frmLoanName
                        If m_InstalmentDetails <> "" Then
                                Dim InstDets() As String
                                Dim J As Integer, Count As Integer
                                Call GetStringArray(m_InstalmentDetails, InstDets, ";")
                                On Error Resume Next
                                For J = 0 To UBound(InstDets) / 2 'frmLoanName.txtInstName.Count - 1
                                    frmLoanName.txtInstName(J).Text = InstDets(Count): Count = Count + 1
                                    frmLoanName.txtLoanAmt(J).Text = InstDets(Count): Count = Count + 1
                                Next
                        End If
                        frmLoanName.Show vbModal, Me
                        m_InstalmentDetails = frmLoanName.P_InstalmentDetails
                        If m_InstalmentDetails <> "" Then
                            Dim SanctionAmount As Currency
                            Call GetStringArray(m_InstalmentDetails, InstDets, ";")
                            For J = 0 To UBound(InstDets)
                                J = J + 1
                                SanctionAmount = SanctionAmount + Val(InstDets(J))
                            Next
                            txtLoanIssue(txtindex).Text = FormatCurrency(SanctionAmount)
                            txtLoanIssue(PropIssueGetIndex("LOanAmount")).Text = FormatCurrency(Val(InstDets(1)))
                        End If
                        Set frmLoanName = Nothing
                End If
                
            Case "LOANAMOUNT"    'cmdsaveloan   cmdloanupdate
                If Val(txtLoanIssue(PropIssueGetIndex("LoanInstalment")).Text) > 1 Then
                        frmLoanName.P_TxtCount = Val(txtLoanIssue(PropIssueGetIndex("LoanInstalment")).Text)
                        Load frmLoanName
                        If m_InstalmentDetails <> "" Then
                                Call GetStringArray(m_InstalmentDetails, InstDets, ";")
                                On Error Resume Next
                                Count = 0
                                For J = 0 To UBound(InstDets) / 2 'frmLoanName.txtInstName.Count - 1
                                    frmLoanName.txtInstName(J).Text = InstDets(Count): Count = Count + 2
                                    frmLoanName.txtLoanAmt(J).Text = "" 'InstDets(Count): Count = Count + 1
                                    frmLoanName.cmdOk.TabIndex = J: frmLoanName.txtLoanAmt(J).TabIndex = J
                                Next
                        End If
                        frmLoanName.Show vbModal, Me
                        Dim RetStr As String
                        RetStr = frmLoanName.P_InstalmentDetails
                        If RetStr <> "" Then
                            Call GetStringArray(RetStr, InstDets, ";")
                            ReDim m_LoanAmount(0)
                            For J = 0 To UBound(InstDets)
                                ReDim Preserve m_LoanAmount(UBound(m_LoanAmount) + 1)
                                J = J + 1
                                m_LoanAmount(UBound(m_LoanAmount) - 1) = Val(InstDets(J))
                                SanctionAmount = SanctionAmount + Val(InstDets(J))
                            Next
                            On Error Resume Next
                            ReDim Preserve m_LoanAmount(UBound(m_LoanAmount) - 1)
                            On Error GoTo ExitLine
                            txtLoanIssue(txtindex).Text = FormatCurrency(SanctionAmount)
                            'txtLoanIssue(PropIssueGetIndex("LOanAmount")).Text = FormatCurrency(Val(InstDets(1)))
                        End If
                        Set frmLoanName = Nothing
                End If
End Select
ExitLine:
Me.MousePointer = vbDefault
End Sub


Private Sub cmdLoanIssueClear_Click()
Dim i As Integer
For i = 0 To txtLoanIssue.Count - 1
    txtLoanIssue(i).Text = ""
Next

' If a date field, display today's date.
m_InstalmentDetails = ""
ReDim m_LoanAmount(0)
i = PropIssueGetIndex("IssueDate")
If i >= 0 Then
    txtLoanIssue(i).Text = FormatDate(gStrDate)
End If
Me.cmdSaveLoan.Enabled = True
cmdLoanUpdate.Enabled = False
End Sub

Private Sub cmdloanscheme_Click(Index As Integer)
Dim txtindex As String

' Check to which text index it is mapped.
txtindex = ExtractToken(cmdLoanScheme(Index).Tag, "TextIndex")
 
' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(lblLoanScheme(Val(txtindex)).Tag, "DataSource")

Select Case UCase$(strField)
    Case "ACCID"
        If m_accUpdatemode = wis_INSERT Then
            txtLoanScheme(txtindex).Text = GetNewAccountNumber
        End If

    #If COMMNETED Then
    Case "ACNAME"
        m_AccHolder.ShowDialog
        txtData(txtindex).Text = m_AccHolder.FullName

    Case "JOINTHOLDER"
        With frmJointHolders
            .Left = Me.Left + picViewport.Left + _
                txtData(txtindex).Left + fraNew.Left + CTL_MARGIN
            .Top = Me.Top + picViewport.Top + txtData(txtindex).Top _
                + fraNew.Top + 300
            .JointHolders = txtData(txtindex).Text
            .Show vbModal
            If .Status = "OK" Then
                txtData(txtindex).Text = .JointHolders
            End If
        End With
        Unload frmJointHolders
    #End If

    Case "INTRODUCERID"
        ' Build a query for getting introducer details.
        ' If an account number specified, exclude it from the list.
        gDbTrans.SQLStmt = "SELECT SBMaster.AccID as [Acc No], " _
                    & "Title + FirstName + Space(1) + Middlename " _
                    & "+ space(1) + LastName as Name, HomeAddress, " _
                    & "OfficeAddress FROM SBMaster, NameTab WHERE " _
                    & "SBMaster.CustomerID = NameTab.CustomerID"
        Dim intIndex As Integer
        intIndex = PropSchemeGetIndex("AccID")
        If txtLoanScheme(intIndex).Text <> "" And _
                IsNumeric(txtLoanScheme(intIndex).Text) Then
            gDbTrans.SQLStmt = gDbTrans.SQLStmt & " AND " _
                & "AccID <> " & txtLoanScheme(intIndex).Text
        End If
        Dim Lret As Long
        Lret = gDbTrans.SQLFetch
        If Lret <= 0 Then
            'MsgBox "No accounts present!", vbExclamation, wis_MESSAGE_TITLE
            MsgBox LoadResString(gLangOffSet + 525), vbExclamation, wis_MESSAGE_TITLE
            Exit Sub
        End If
        'Fill the details to report dialog and display it.
        If m_LookUp Is Nothing Then
            Set m_LookUp = New frmLookUp
        End If
        If Not FillView(m_LookUp.lvwReport, gDbTrans.Rst) Then
            'MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
            MsgBox "Error loading introducer accounts.", _
                    vbCritical, wis_MESSAGE_TITLE
            Exit Sub
        End If
        With m_LookUp
            ' Hide the print and save buttons.
            .cmdPrint.Visible = False
            .cmdSave.Visible = False
            ' Set the column widths.
            .lvwReport.ColumnHeaders(2).Width = 3750
            .lvwReport.ColumnHeaders(3).Width = 3750
            .Title = "Select Introducer..."
            .Show vbModal, Me
            txtLoanScheme(txtindex).Text = .lvwReport.SelectedItem.Text
            txtLoanScheme(txtindex + 1).Text = .lvwReport.SelectedItem.SubItems(1)
        End With
    Case "CREATEDATE"
        With Calendar
            .Left = txtLoanScheme(txtindex).Left + Me.Left _
                    + picLoanSchemeViewPort.Left + fraSchemes.Left + 50
            .Top = Me.Top + txtLoanScheme(txtindex).Top _
                    + picLoanSchemeViewPort.Top + fraSchemes.Top + 300
            .Width = txtLoanScheme(txtindex).Width
            .Height = .Width
            .SelDate = txtLoanScheme(txtindex).Text
            .Show vbModal, Me
            If .SelDate <> "" Then txtLoanScheme(txtindex).Text = .SelDate
        End With
End Select
End Sub
Private Sub cmdNextTrans_Click()
 
If m_rstLoanTrans.AbsolutePosition <= m_rstLoanTrans.RecordCount - 11 Then
    If m_rstLoanTrans.AbsolutePosition Mod 10 <> 0 Then
        m_rstLoanTrans.Move 10 - m_rstLoanTrans.AbsolutePosition Mod 10
        If m_rstLoanTrans.AbsolutePosition >= m_rstLoanTrans.RecordCount - 10 Then
            cmdNextTrans.Enabled = False
        End If
    End If
Else
    cmdNextTrans.Enabled = False
End If
Call LoanLoadGrid

End Sub
Private Sub cmdOK_Click()
Dim Cancel As Boolean
Unload Me
End Sub





Private Sub cmdPrevTrans_Click()
Dim CHECK As Boolean
If m_rstLoanTrans.EOF Then
    m_rstLoanTrans.Move -19
    CHECK = True
End If
If m_rstLoanTrans.AbsolutePosition >= 10 Then
    If m_rstLoanTrans.AbsolutePosition Mod 10 = 0 Then
        'm_rstloantrans.MovePrevious
        If Not CHECK Then m_rstLoanTrans.Move -1 * (m_rstLoanTrans.AbsolutePosition Mod 10 + 20)
    Else
        If Not CHECK Then m_rstLoanTrans.Move -1 * (m_rstLoanTrans.AbsolutePosition Mod 10 + 20)
    End If
End If
If m_rstLoanTrans.AbsolutePosition < 10 Then
    m_rstLoanTrans.MoveFirst
End If
Call LoanLoadGrid
End Sub

Private Sub cmdPrint_Click()
      
      ' Call the print class services...
      Dim my_printClass As New clsPrint
      Set my_printClass.DataSource = grd
      my_printClass.ReportDestination = "PREVIEW"
      my_printClass.ReportTitle = Me.txtMemberName.Text
      my_printClass.HeaderRectangle = True
      my_printClass.CompanyName = gCompanyName

      'See the Code - Ravi
      frmPrint.picPrint.Cls   'vinay
      
      my_printClass.PrintReport

End Sub

Private Sub cmdRepay_Click()
If LoanRepay Then
    Tabloans_Click
Else
    Exit Sub
End If

' Reload the details so that the recordsets or latest.
If Not LoanLoad(FormatField(m_rstLoanMast("loanid"))) Then
    'MsgBox "Error reloading the loan details.", vbCritical, wis_MESSAGE_TITLE
    MsgBox "Error reloading the loan details.", vbCritical, wis_MESSAGE_TITLE
End If
txtRepayDate.SetFocus
txtRepayDate.SelLength = Len(txtRepayDate.Text)

End Sub

Private Sub cmdRepayDate_Click()
With Calendar
    .Left = Me.Left + Me.fraRepayments.Left + Me.cmdRepayDate.Left
    .Top = Me.Top + Me.fraRepayments.Top + Me.cmdRepayDate.Top
    .SelDate = txtRepayDate.Text
    .Show vbModal
    If .SelDate <> "" Then txtRepayDate.Text = .SelDate
End With

End Sub

Private Sub cmdSaveLoan_Click()
Call LoanSave
Call LoanLoad(m_LoanId)
End Sub

Private Sub cmdSchemeLoad_Click()
Me.MousePointer = vbHourglass
Dim Lret As Long

' Get a list of available loan schemes.
gDbTrans.SQLStmt = "SELECT SchemeID ,SchemeName ,InterestRate  FROM loantypes"
Lret = gDbTrans.SQLFetch
If Lret <= 0 Then
    'MsgBox "No loan types defined.", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 768), vbInformation, wis_MESSAGE_TITLE
    GoTo Exit_Line
End If

' Load the details to report window.
If m_LookUp Is Nothing Then
    Set m_LookUp = New frmLookUp
End If
With m_LookUp
    FillView .lvwReport, gDbTrans.Rst, True
    .Show vbModal, Me
    If .m_SelItem <> "" Then
        If SchemeLoad(Val(.m_SelItem)) Then
            m_SchemeLoaded = True
            lblOpMode.Caption = "Operation Mode : <UPDATE>"
        End If
    End If
End With

Exit_Line:
    Me.MousePointer = vbDefault
    Exit Sub

Err_Line:
    If Err Then
        ' _

        MsgBox "LoadScheme: " & Err.Description, _
                vbCritical, wis_MESSAGE_TITLE
    End If
    GoTo Exit_Line
End Sub
Private Sub cmdSchemeSave_Click()
Me.MousePointer = vbHourglass
If Not SchemeSave Then GoTo Exit_Line

Dim InterestRate As Single
Dim txtindex As Integer
Dim StartDate  As String


InterestRate = PropSchemeGetVal("InterestRate")
StartDate = PropSchemeGetVal("CreateDate")

'Instead Of Calling Function SaveInterest in SchemeSave It Has Been Called Here
'get the SchemeName From DataBase
'Get The SchemeName From LoanTypes
Dim SchemeName As String
If m_SchemeLoaded Then
    SchemeName = PropSchemeGetVal("LoanName")
Else
    SchemeName = PropSchemeGetVal("LoanName")
End If
Dim ClsInt As New clsInterest
Call ClsInt.SaveInterest(wis_Loans, SchemeName, CSng(InterestRate), StartDate)

Exit_Line:
Me.MousePointer = vbDefault
End Sub

Private Sub cmdStDate_Click()
With Calendar
    .SelDate = FormatDate(gStrDate)
    .Left = Me.Left + fraReports.Left + fraDateRange.Left + cmdStDate.Left
    .Top = Me.Top + fraReports.Top + fraDateRange.Top + cmdStDate.Top + 300
    .SelDate = txtStartDate.Text
    .Show vbModal, Me
    If .SelDate <> "" Then txtStartDate.Text = Calendar.SelDate
End With

End Sub

Private Sub cmdUndo_Click()

Dim nRet As Integer

If cmdUndo.Caption = LoadResString(gLangOffSet + 5) Then  '"Undo
   'nRet = MsgBox("This will undo the last transaction for this loan account. " _
        & vbCrLf & "Do you want to continue ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(LoadResString(gLangOffSet + 583), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   Me.MousePointer = vbHourglass
   
   Call LoanUndoLastTransaction

ElseIf cmdUndo.Caption = LoadResString(gLangOffSet + 313) Then  '"Reopen
   'nRet = MsgBox("Are you sure to reopen this loan ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(LoadResString(gLangOffSet + 538), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   Me.MousePointer = vbHourglass
   gDbTrans.SQLStmt = "UpDate LoanMaster set LoanClosed = False where LoanId = " & m_LoanId
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to reopen the loan"
      MsgBox LoadResString(gLangOffSet + 536), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.CommitTrans
   'MsgBox "Account reopened succefully"
   MsgBox LoadResString(gLangOffSet + 522), vbInformation, wis_MESSAGE_TITLE
   
ElseIf cmdUndo.Caption = LoadResString(gLangOffSet + 14) Then  '"Delete
   'nRet = MsgBox("Are you sure to delete this accoutnt ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(LoadResString(gLangOffSet + 539), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   Me.MousePointer = vbHourglass
   gDbTrans.SQLStmt = "Delete * From LoanTrans Where LoanId = " & m_LoanId
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete the loan"
      MsgBox LoadResString(gLangOffSet + 532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.SQLStmt = "Delete * From LoanMaster Where LoanId = " & m_LoanId
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete"
      MsgBox LoadResString(gLangOffSet + 532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.CommitTrans
   'MsgBox "Account reopened succefully"
   MsgBox LoadResString(gLangOffSet + 730), vbInformation, wis_MESSAGE_TITLE
   Me.MousePointer = vbDefault
   'Exit Sub
End If

' Reload the loan details.
If Not m_rstLoanMast Is Nothing Then
    LoanLoad FormatField(m_rstLoanMast("loanid"))
End If
Me.MousePointer = vbDefault
End Sub
Private Sub cmdLoanUpdate_Click()
    Call LoanUpDate
End Sub

Private Sub cmdView_Click()
'Call SetKannadaCaption
' Validate the user input.
' Check for starting date.
With txtStartDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        'MsgBox "Enter a valid starting date.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtStartDate
        GoTo Exit_Line
    End If
End With

' Check for ending date.
With txtEndDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        'MsgBox "Enter a valid ending date.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 501), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEndDate
        GoTo Exit_Line
    End If
End With

' Check for starting amount.
With txtStartAmt
    If .Enabled Then
    If .Text <> "" And Not IsNumeric(.Text) Then
        'MsgBox "Enter a valid starting amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtStartAmt
        GoTo Exit_Line
    End If
    End If
End With
    
' Check for ending amount.
With txtEndAmt
    If .Enabled Then
        If .Text <> "" And Not IsNumeric(.Text) Then
        'MsgBox "Enter a valid ending amount.", _
                vbInformation, wis_MESSAGE_TITLE
        MsgBox LoadResString(gLangOffSet + 506), _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEndAmt
        GoTo Exit_Line
        End If
    End If
End With

gCancel = False
frmCancel.Show
frmCancel.Refresh
Set m_frmLoanReport = New frmLoanView
Screen.MousePointer = vbHourglass
Load m_frmLoanReport
If gCancel Then GoTo ExitLine
Unload frmCancel
Screen.MousePointer = vbDefault
m_frmLoanReport.Show vbModal

ExitLine:
   On Error Resume Next
   Set frmCancel = Nothing
   Set m_frmLoanReport = Nothing
Exit_Line:
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'#If COMMENTED Then
' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
Select Case KeyCode
    Case vbKeyUp
        ' Scroll up.
        With vscLoanScheme
            If .value - .SmallChange > .Min Then
                .value = .value - .SmallChange
            Else
                .value = .Min
            End If
        End With
    Case vbKeyDown
        ' Scroll down.
        With vscLoanScheme
            If .value + .SmallChange < .Max Then
                .value = .value + .SmallChange
            Else
                .value = .Max
            End If
        End With
   Case vbKeyTab
      If TabStrip.SelectedItem.Index = TabStrip.Tabs.Count Then
            TabStrip.Tabs(1).Selected = True
      Else
            TabStrip.Tabs(TabStrip.SelectedItem.Index + 1).Selected = True
      End If
End Select
'#End If
End Sub

Private Sub Form_Load()
      Screen.MousePointer = vbHourglass
      'set icon for the form caption
      Me.Icon = LoadResPicture(161, vbResIcon)
      cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
      'call SetKannadaCaption
      Call SetKannadaCaption
      'Centre the form
      
          Me.Move (Screen.Width - Me.Width) \ 2, _
                  (Screen.Height - Me.Height) \ 2
      PropInitializeForm
      
      'Shashi 28/7/00
      'Now get the details from the install table to get the information whether this package
      ' is Considers the Interest till Date or interest balance
      ' Description : if interest on a loan as on today(1/10/99) is Rs200. And customer want to pay only Rs100
      ' In such case The interest can be considered Upto date and remaining Rs 100 will be put as Interest Balance
      ' or find the date till where paid interest will be equal so Rs 100(payingAmount). The date will be 10 days earlier than the today
      ' so then the Interest paid date is 10 days earlier (i.e 20/9/99)
      M_InterestBalance = True
      gDbTrans.SQLStmt = "Select * from Install Where KeyData = " & AddQuotes("InterestDateOrBalance", True)
      If gDbTrans.SQLFetch > 0 Then
         If StrComp(FormatField(gDbTrans.Rst("ValueData")), "Balance", vbTextCompare) = 0 Then
               M_InterestBalance = True
         End If
      End If
      If M_InterestBalance Then
            lblIntBalance.Caption = LoadResString(gLangOffSet + 387)
            cmdIntdate.Visible = False
      Else
            lblIntBalance.Caption = LoadResString(gLangOffSet + 348)
            txtIntBalance.Width = txtRepayDate.Width
            cmdIntdate.Visible = False
      End If
      '** shashi 28/7/99
      Me.txtRepayDate.Text = FormatDate(gStrDate)
      ReDim m_LoanAmount(0)
      
      'In Reports TaB Load The Combos with Caste and Places  Respectively
      chkCaste.Enabled = False
      cmbCaste.Enabled = False
      cmbCaste.BackColor = wisGray
      chkPlace.Enabled = False
      cmbPlace.Enabled = False
      cmbPlace.BackColor = wisGray
      
      Call LoadCastes
      Call LoadPlaces
      
      Screen.MousePointer = vbDefault
      Me.OptLoanBalance.value = True
End Sub
Private Sub LoadPlaces()
    Dim i As Integer
    Me.cmbPlace.Clear
   gDbTrans.SQLStmt = "Select Places from PlaceTab order by Places"
    If gDbTrans.SQLFetch > 0 Then
    gDbTrans.Rst.MoveFirst
        For i = 1 To gDbTrans.Records
            If gDbTrans.Rst("Places") <> "" Then
                cmbPlace.AddItem gDbTrans.Rst("Places")
                gDbTrans.Rst.MoveNext
            End If
        Next i
    Else
        cmbPlace.AddItem "Bangalore"
    End If
 
End Sub




Private Sub LoadCastes()
Me.cmbCaste.Clear
Dim i As Integer
 gDbTrans.SQLStmt = "Select Caste from CasteTab order by Caste"
    If gDbTrans.SQLFetch > 0 Then
    gDbTrans.Rst.MoveFirst
        For i = 1 To gDbTrans.Records
            If gDbTrans.Rst("Caste") <> "" Then
                cmbCaste.AddItem gDbTrans.Rst("Caste")
                gDbTrans.Rst.MoveNext
            End If
        Next i
    Else
        cmbCaste.AddItem "Hindu "
        
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)


' Report form.
If Not m_LookUp Is Nothing Then
    Unload m_LookUp
    Set m_LookUp = Nothing
End If
gWindowHandle = 0
End Sub

Private Sub lblLoanIssue_DblClick(Index As Integer)
On Error Resume Next
txtLoanIssue(Index).SetFocus

End Sub


Private Sub m_frmLoanReport_Initialise(Min As Long, Max As Long)

On Error Resume Next
With frmCancel
    If Max <> 0 Then
        .prg.Visible = True
        .Refresh
        .prg.Min = Min
        If Max > 32500 Then Max = 32500
        .prg.Max = Max
    End If
End With

End Sub

Private Sub m_frmLoanReport_Processing(strMessage As String, Ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS: " & vbCrLf & strMessage
    If Ratio > 0 Then
        If Ratio > 1 Then
                .prg.value = Ratio * .prg.Max
        End If
    End If
End With
End Sub


Private Sub optDailyCashBook_Click()
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optGeneralLedger_Click()
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optInstOverDue_Click()
    txtEndDate.Enabled = False
    txtEndDate.BackColor = wisGray
    cmdEndDate.Enabled = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optInterest_Click()
 txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optLoanBalance_Click()
    txtEndDate.Enabled = False
    txtEndDate.BackColor = wisGray
    cmdEndDate.Enabled = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optLoanHolders_Click()
    
    txtEndDate.Enabled = False
    txtEndDate.BackColor = wisGray
    cmdEndDate.Enabled = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = vbWhite
    chkCaste.Enabled = True
    chkPlace.Enabled = True
    
End Sub


Private Sub OptLoanSanction_Click()
    txtEndDate.Enabled = False
    txtEndDate.BackColor = wisGray
    cmdEndDate.Enabled = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
    
End Sub

Private Sub optLoansIssued_Click()
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optOverDueLoans_Click()
    txtEndDate.Enabled = False
    txtEndDate.BackColor = wisGray
    cmdEndDate.Enabled = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub


Private Sub optRepaymentsMade_Click()
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optTransaction_Click()
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    chkCaste.Enabled = False: chkCaste.value = vbUnchecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = False: chkPlace.value = vbUnchecked
    'chkPlace.BackColor = wisGray
End Sub


Private Sub picLoanIssueSlider_Resize()

'If Not arrangerequired Then Exit Sub
On Error Resume Next

' Arrange the controls on this container.
Dim Ctl As Control
For Each Ctl In Me.Controls
    If Ctl.Container.Name = picLoanIssueSlider.Name Then
        If TypeOf Ctl Is Label Then
            
        ElseIf TypeOf Ctl Is TextBox Then
        
        ElseIf TypeOf Ctl Is ComboBox Then
        
        ElseIf TypeOf Ctl Is CommandButton Then
        
        End If
    End If
Next


End Sub

Private Sub SearchDialog_SelectClick(strAcno As Long)
#If COMMENTED Then
Dim Lret As Long

' Load the name details to CustRegistration dialog.
m_AccHolder.LoadCustomerInfo (strAcno)

' Get the details of the selected account number.
With gDbTrans
    .SQLStmt = "SELECT * FROM SBMaster WHERE AccID = " & strAcno
    Lret = .SQLFetch
    If Lret <= 0 Then
        'Fatal error.
        'MsgBox "Fatal error in procedure '" _
            & "SearchDialog_SelectClick'!", vbExclamation
        MsgBox "Fatal error in procedure '" _
            & "SearchDialog_SelectClick'!", vbExclamation
        Exit Sub
    End If

    ' Fill the details of account holder.
    FillAcHolderDetails
    lblOperation.Caption = LoadResString(gLangOffSet + 56) '"Operation Mode : <UPDATE>"
    m_accUpdatemode = wis_UPDATE
End With

#End If
End Sub
Private Sub TabStrip_Click()
On Error Resume Next
Dim txtindex As Integer
Select Case UCase$(TabStrip.SelectedItem.Key)
    Case "LOANSCHEMES"
        fraSchemes.Visible = True
        fraLoanAccounts.Visible = False
        fraRepayments.Visible = False
        fraReports.Visible = False
        txtindex = PropSchemeGetIndex("LoanName")
        txtLoanScheme(txtindex).SetFocus
    Case "LOANACCOUNTS"
        fraSchemes.Visible = False
        fraLoanAccounts.Visible = True
        fraRepayments.Visible = False
        fraReports.Visible = False
        txtindex = PropIssueGetIndex("MemberID")
        txtLoanIssue(txtindex).SetFocus

    Case "REPAYMENTS"
        fraSchemes.Visible = False
        fraLoanAccounts.Visible = False
        fraRepayments.Visible = True
        fraReports.Visible = False
    
    Case "REPORTS"
        fraSchemes.Visible = False
        fraLoanAccounts.Visible = False
        fraRepayments.Visible = False
        fraReports.Visible = True

End Select

End Sub
Private Sub Tabloans_Click()

' Get the loan scheme name of the selected tab,
' by stripping the first three characters.
Dim lLoanID As Long
lLoanID = Val(Mid(tabLOans.SelectedItem.Key, 4))
If Not LoanLoad(lLoanID) Then
    fraLoanGrid.Visible = False
    Exit Sub
Else
    fraLoanGrid.Visible = True
End If
End Sub

Private Sub txtAccNo_Change()
If Trim$(txtAccNo.Text) <> "" Then
    cmdLoad.Enabled = True
    txtRepayAmt.Enabled = True
    txtRepayDate.Enabled = True
Else
    cmdLoad.Enabled = False
    txtRepayAmt.Enabled = False
    txtRepayDate.Enabled = False
End If
End Sub

Private Sub txtAccNo_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'Me.ActiveControl
End Sub


Private Sub txtEndAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtEndDate_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtInstAmt_Change()
   txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest) + Val(txtMIsc.Text))
End Sub

Private Sub txtInstAmt_GotFocus()
On Error Resume Next
If txtLoanAmt.Caption <> "" Then
    m_LoanId = Val(Mid(tabLOans.SelectedItem.Key, 4))
    LoanLoad (m_LoanId)
End If
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtIntBalance_Click()
   txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest))
End Sub

Private Sub txtIntPaidDate_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtLoanIssue_Change(Index As Integer)

' In case of fields bound to loanduedate and
' instalment mode, compute the instalment amount
' and display it.
Dim strDataSrc As String
Dim Amt As Currency
Dim txtindex As Integer

strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")
If StrComp(strDataSrc, "LoanDueDate", vbTextCompare) = 0 Or _
        StrComp(strDataSrc, "InstalmentMode", vbTextCompare) = 0 Then
'    Amt = ComputeInstalmentAmount
    
    ' Display it in the relevant fieldbox.
    txtindex = PropIssueGetIndex("InstalmentAmount")
    txtLoanIssue(txtindex).Text = Amt
End If

End Sub
Private Sub txtLoanIssue_DblClick(Index As Integer)
Dim strDispType As String
' Get the display type.
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    txtLoanIssue_KeyPress Index, vbKeyReturn
End If


End Sub
Private Sub txtLoanIssue_GotFocus(Index As Integer)
lblLoanIssue(Index).ForeColor = vbBlue
PropSetIssueDescription lblLoanIssue(Index)

' Scroll the window, so that the
' control in focus is visible.
PropScrollLoanIssueWindow txtLoanIssue(Index)

' Select the text, if any.
With txtLoanIssue(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' If the display type is Browse, then
' show the command button for this text.
Dim strDispType As String
Dim TextIndex As String
strDispType = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(lblLoanIssue(Index).Tag, "textindex")
    If TextIndex <> "" Then cmdLoanIssue(Val(TextIndex)).Visible = True
End If


' Hide all other command buttons...
Dim i As Integer
For i = 0 To cmdLoanIssue.Count - 1
    If i <> Val(TextIndex) Or TextIndex = "" Then
        cmdLoanIssue(i).Visible = False
    End If
Next

' Hide all other combo boxes.
For i = 0 To cmbLoanIssue.Count - 1
    If i <> Val(TextIndex) Or TextIndex = "" Then
        cmbLoanIssue(i).Visible = False
    End If
Next

End Sub
Private Sub txtLoanIssue_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(lblLoanIssue(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(lblLoanIssue(Index).Tag, "TextIndex")
        If Trim$(strIndex) <> "" Then
            cmbLoanIssue(Val(strIndex)).Visible = True
            cmbLoanIssue(Val(strIndex)).ZOrder
            cmbLoanIssue(Val(strIndex)).SetFocus
        End If
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub txtLoanIssue_LostFocus(Index As Integer)
lblLoanIssue(Index).ForeColor = vbBlack

' Declare the reqd. member variables here...
Dim strDataSrc As String
Dim Lret As Long
Dim txtindex As Integer
' Get the name of the data source bound to this control.
strDataSrc = ExtractToken(lblLoanIssue(Index).Tag, "DataSource")

' For the textbox bound to MemberID, we need to
' fill in the corresponding member name.
If StrComp(strDataSrc, "MemberID", vbTextCompare) = 0 Then
    
    ' Get the name of the member for the given memberid.
    Dim MMObj As New clsMMAcc
    txtLoanIssue(Index + 1).Text = MMObj.MemberName(Val(txtLoanIssue(Index).Text))
    Set MMObj = Nothing
ElseIf StrComp(strDataSrc, "LoanAmt", vbTextCompare) = 0 Then
    ' Get the index of the textbox bound to instalment amount field.
    
    Dim InstalmentAmt As Currency
    InstalmentAmt = ComputeInstalmentAmount
    txtindex = PropIssueGetIndex("InstalmentAmount")
    txtLoanIssue(txtindex).Text = InstalmentAmt
End If
'for Loan Installment Vinay
 txtindex = PropIssueGetIndex("LoanInstalment")
    If Val(txtLoanIssue(txtindex).Text) = 0 Then txtLoanIssue(txtindex).Text = "1"


End Sub

Private Sub txtloanscheme_DblClick(Index As Integer)
Dim strDispType As String
' Get the display type.
strDispType = ExtractToken(lblLoanScheme(Index).Tag, "DisplayType")
If StrComp(strDispType, "List", vbTextCompare) = 0 Then
    txtLoanScheme_KeyPress Index, vbKeyReturn
End If
End Sub
Private Sub txtloanscheme_GotFocus(Index As Integer)
lblLoanScheme(Index).ForeColor = vbBlue
PropSetSchemeDescription lblLoanScheme(Index)

' Scroll the window, so that the
' control in focus is visible.
PropScrollLoanIssueWindow txtLoanScheme(Index)

' Select the text, if any.
With txtLoanScheme(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' If the display type is Browse, then
' show the command button for this text.
Dim strDispType As String
Dim TextIndex As String
strDispType = ExtractToken(lblLoanScheme(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(lblLoanScheme(Index).Tag, "textindex")
    If TextIndex <> "" Then cmdLoanScheme(Val(TextIndex)).Visible = True
'ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
'    ' Get the cmdbutton index.
'    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
'    If TextIndex <> "" Then
'        cmb(TextIndex).Visible = True
'    End If
End If


' Hide all other command buttons...
Dim i As Integer
For i = 0 To cmdLoanScheme.Count - 1
    If i <> Val(TextIndex) Or TextIndex = "" Then
        cmdLoanScheme(i).Visible = False
    End If
Next

' Hide all other combo boxes.
For i = 0 To cmbLoanScheme.Count - 1
    If i <> Val(TextIndex) Or TextIndex = "" Then
        cmbLoanScheme(i).Visible = False
    End If
Next

End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(lblLoanScheme(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(lblLoanScheme(Index).Tag, "TextIndex")
        If Trim$(strIndex) <> "" Then
            cmbLoanScheme(Val(strIndex)).Visible = True
            cmbLoanScheme(Val(strIndex)).SetFocus
        End If
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub
Private Sub lblloanscheme_DblClick(Index As Integer)
On Error Resume Next
txtLoanScheme(Index).SetFocus
End Sub

Private Sub txtLoanScheme_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(lblLoanScheme(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(lblLoanScheme(Index).Tag, "TextIndex")
        If Trim$(strIndex) <> "" Then
            cmbLoanScheme(Val(strIndex)).Visible = True
            cmbLoanScheme(Val(strIndex)).ZOrder
            cmbLoanScheme(Val(strIndex)).SetFocus
        End If
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub txtLoanScheme_LostFocus(Index As Integer)
lblLoanScheme(Index).ForeColor = vbBlack

End Sub

Private Sub txtMIsc_Change()
   txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest) + Val(txtMIsc.Text))
End Sub

Private Sub txtPenalInterest_Change()
   txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest) + Val(txtMIsc.Text))
End Sub

Private Sub txtPenalInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtRegInterest_Change()
      txtTotInst.Text = FormatCurrency(Val(txtInstAmt.Text) + Val(txtRegInterest.Text) + Val(txtPenalInterest) + Val(txtMIsc.Text))
End Sub

Private Sub txtRegInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtRepayAmt_Change()
If Trim(txtRepayAmt.Text) <> "" Then
    cmdRepay.Enabled = True
Else
    cmdRepay.Enabled = False
End If
End Sub

Private Sub txtRepayAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtStartAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtTotInst_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub vscLoanIssue_Change()
' Move the picLoanissueSlider.
picLoanIssueSlider.Top = -vscLoanIssue.value
End Sub

Private Sub vscLoanScheme_Change()
' Move the picLoanSchemeSlider.
picLoanSchemeSlider.Top = -vscLoanScheme.value
End Sub
Private Sub SetKannadaCaption()
'to trap an error
On Error Resume Next

'declare variables
Dim Ctrl As Control

For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next
On Error GoTo 0
'Set kannada caption for the generally used  controls
Me.cmdOk.Caption = LoadResString(gLangOffSet + 1)

'Set captions for the all the tabs
Me.TabStrip.Tabs(1).Caption = LoadResString(gLangOffSet + 216)
Me.TabStrip.Tabs(2).Caption = LoadResString(gLangOffSet + 215)
Me.TabStrip.Tabs(3).Caption = LoadResString(gLangOffSet + 214)
Me.TabStrip.Tabs(4).Caption = LoadResString(gLangOffSet + 212)

' Set captions for loan schemes
Me.fraSchemes.Caption = LoadResString(gLangOffSet + 214)
Me.lblOpMode.Caption = LoadResString(gLangOffSet + 54)
Me.cmdSchemeLoad.Caption = LoadResString(gLangOffSet + 3)
Me.cmdSchemeSave.Caption = LoadResString(gLangOffSet + 7)
Me.cmdClearScheme.Caption = LoadResString(gLangOffSet + 8)
'Set captions for the Loan Accounts
Me.chkSb.Caption = LoadResString(gLangOffSet + 436) & " " & LoadResString(gLangOffSet + 271)
Me.fraLoanAccounts.Caption = LoadResString(gLangOffSet + 215)
Me.cmdSaveLoan.Caption = LoadResString(gLangOffSet + 15)
Me.cmdLoanUpdate.Caption = LoadResString(gLangOffSet + 7)
Me.cmdLoanIssueClear.Caption = LoadResString(gLangOffSet + 8)

'Set kannada captions for the Repayments

Me.lblMisc.Caption = LoadResString(gLangOffSet + 327)
Me.lblMemberId.Caption = LoadResString(gLangOffSet + 49)
Me.cmdLoad.Caption = LoadResString(gLangOffSet + 3)
Me.lblName.Caption = LoadResString(gLangOffSet + 35)
Me.lblLoanAmount.Caption = LoadResString(gLangOffSet + 58)
Me.lblIssueDate.Caption = LoadResString(gLangOffSet + 340)
Me.lblRepaidAmt.Caption = LoadResString(gLangOffSet + 341)
Me.lblBalance.Caption = LoadResString(gLangOffSet + 342)
Me.lblInstAmt.Caption = LoadResString(gLangOffSet + 343)
Me.lblRegInterest.Caption = LoadResString(gLangOffSet + 344)
Me.lblPenalInterest.Caption = LoadResString(gLangOffSet + 345)
Me.lblIntBalance.Caption = LoadResString(gLangOffSet + 387)
Me.lblTotInst.Caption = LoadResString(gLangOffSet + 346)
Me.lblRepayAmt.Caption = LoadResString(gLangOffSet + 341)
Me.cmdUndo.Caption = LoadResString(gLangOffSet + 5)
Me.cmdRepay.Caption = LoadResString(gLangOffSet + 20)
Me.cmdInstalment.Caption = LoadResString(gLangOffSet + 57)
Me.lblRepayDate.Caption = LoadResString(gLangOffSet + 347)
'Set kannadacaption for Reports tab
optLoansIssued.Caption = LoadResString(gLangOffSet + 81)
optRepaymentsMade.Caption = LoadResString(gLangOffSet + 82)
optLoanHolders.Caption = LoadResString(gLangOffSet + 83)
optOverDueLoans.Caption = LoadResString(gLangOffSet + 84)
optDailyCashBook.Caption = LoadResString(gLangOffSet + 85)
optGeneralLedger.Caption = LoadResString(gLangOffSet + 86)
OptLoanBalance.Caption = LoadResString(gLangOffSet + 67)
optTransaction.Caption = LoadResString(gLangOffSet + 62)
optInstOverDue.Caption = LoadResString(gLangOffSet + 113)
OptLoanSanction.Caption = LoadResString(gLangOffSet + 262)
optInterest.Caption = LoadResString(gLangOffSet + 483)
optGuarantors.Caption = LoadResString(gLangOffSet + 483)

fraReports.Caption = LoadResString(gLangOffSet + 212)
lblLoanType.Caption = LoadResString(gLangOffSet + 89)
fraChooseReport.Caption = LoadResString(gLangOffSet + 288)
fraDateRange.Caption = LoadResString(gLangOffSet + 106)
lblDate1.Caption = LoadResString(gLangOffSet + 109)
lblDate2.Caption = LoadResString(gLangOffSet + 110)
fraAmountRange.Caption = LoadResString(gLangOffSet + 105)
lblAmt1.Caption = LoadResString(gLangOffSet + 107)
lblAmt2.Caption = LoadResString(gLangOffSet + 108)
cmdView.Caption = LoadResString(gLangOffSet + 13)
chkPlace.Caption = LoadResString(gLangOffSet + 112)
chkCaste.Caption = LoadResString(gLangOffSet + 128)

End Sub


