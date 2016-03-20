VERSION 5.00
Begin VB.Form frmLoanReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select report"
   ClientHeight    =   7785
   ClientLeft      =   330
   ClientTop       =   465
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7995
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   400
      Left            =   5280
      TabIndex        =   30
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   6600
      TabIndex        =   29
      Top             =   7380
      Width           =   1215
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   7215
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7785
      Begin VB.Frame fraOrder 
         Caption         =   " List Order"
         Height          =   1845
         Left            =   150
         TabIndex        =   26
         Top             =   5280
         Width           =   7425
         Begin VB.ComboBox cmbPurpose 
            Height          =   315
            Left            =   1860
            TabIndex        =   40
            Top             =   1380
            Width           =   3045
         End
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   400
            Left            =   6150
            TabIndex        =   38
            Top             =   1380
            Width           =   1215
         End
         Begin VB.TextBox txtStartDate 
            Height          =   315
            Left            =   1845
            TabIndex        =   36
            Top             =   930
            Width           =   1290
         End
         Begin VB.TextBox txtEndDate 
            Height          =   315
            Left            =   5745
            TabIndex        =   35
            Top             =   930
            Width           =   1230
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   315
            Left            =   3210
            TabIndex        =   34
            Top             =   930
            Width           =   315
         End
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   315
            Left            =   7035
            TabIndex        =   33
            Top             =   930
            Width           =   315
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   300
            Left            =   420
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   2385
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   300
            Left            =   4140
            TabIndex        =   27
            Top             =   330
            Width           =   2205
         End
         Begin VB.Label lblPurpose 
            Caption         =   "Loan Purpose :"
            Height          =   300
            Left            =   270
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   7350
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "&Starting date :"
            Height          =   300
            Left            =   240
            TabIndex        =   31
            Top             =   975
            Width           =   1470
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "&Ending date :"
            Height          =   300
            Left            =   3855
            TabIndex        =   37
            Top             =   990
            Width           =   1605
         End
      End
      Begin VB.ComboBox cmbLoanType 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3825
      End
      Begin VB.Frame fraSchedule 
         Caption         =   "Schedules"
         Height          =   2565
         Left            =   1980
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1905
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &5"
            Height          =   300
            Index           =   5
            Left            =   60
            TabIndex        =   24
            Top             =   1900
            Width           =   1575
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &1"
            Height          =   300
            Index           =   1
            Left            =   60
            TabIndex        =   20
            Top             =   620
            Width           =   1785
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &2"
            Height          =   300
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   940
            Width           =   1725
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &3"
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   60
            TabIndex        =   22
            Top             =   1260
            Width           =   1665
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &4"
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   60
            TabIndex        =   23
            Top             =   1580
            Width           =   1695
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &6"
            Height          =   300
            Index           =   6
            Left            =   60
            TabIndex        =   25
            Top             =   2220
            Width           =   1695
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "&Monthly report"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report :"
         Height          =   4650
         Left            =   150
         TabIndex        =   3
         Top             =   780
         Width           =   7425
         Begin VB.Frame fraReceivable 
            Caption         =   "Receivables"
            Height          =   1320
            Left            =   1920
            TabIndex        =   41
            Top             =   2880
            Visible         =   0   'False
            Width           =   2625
            Begin VB.OptionButton optLoanReceivable 
               Caption         =   "&Monthly report"
               Height          =   300
               Left            =   60
               TabIndex        =   44
               Top             =   300
               Value           =   -1  'True
               Width           =   2085
            End
            Begin VB.OptionButton optIntReceivableTill 
               Caption         =   "Schedule &2"
               Height          =   300
               Left            =   60
               TabIndex        =   43
               Top             =   990
               Width           =   1725
            End
            Begin VB.OptionButton optIntReceivable 
               Caption         =   "Schedule &1"
               Height          =   300
               Left            =   60
               TabIndex        =   42
               Top             =   645
               Width           =   1785
            End
         End
         Begin VB.OptionButton optReceivable 
            Caption         =   "Other receivable"
            Height          =   300
            Left            =   375
            TabIndex        =   32
            Top             =   4050
            Width           =   3030
         End
         Begin VB.OptionButton optCustRp 
            Caption         =   "Customers Receipt && Payment"
            Height          =   300
            Left            =   4140
            TabIndex        =   17
            Top             =   3522
            Width           =   3075
         End
         Begin VB.OptionButton optGeneralLedger 
            Caption         =   "&General Ledger"
            Height          =   300
            Left            =   4140
            TabIndex        =   7
            Top             =   1434
            Width           =   3060
         End
         Begin VB.OptionButton optOverDueLoans 
            Caption         =   "&Over due loans"
            Height          =   300
            Left            =   4140
            TabIndex        =   11
            Top             =   1956
            Width           =   3030
         End
         Begin VB.OptionButton optGuarantors 
            Caption         =   "&Guarantor's List"
            Height          =   300
            Left            =   4140
            TabIndex        =   15
            Top             =   3000
            Width           =   3060
         End
         Begin VB.OptionButton OptLoanSanction 
            Caption         =   "S&anctioned Loans"
            Height          =   300
            Left            =   4140
            TabIndex        =   13
            Top             =   2478
            Width           =   3030
         End
         Begin VB.OptionButton optSubDayBook 
            Caption         =   "Sub day Book"
            Height          =   300
            Left            =   4140
            TabIndex        =   5
            Top             =   390
            Width           =   3000
         End
         Begin VB.OptionButton optLoansRecovery 
            Caption         =   "Loan &Recovery"
            Height          =   300
            Left            =   390
            TabIndex        =   9
            Top             =   1434
            Width           =   2730
         End
         Begin VB.OptionButton optLoansAdvance 
            Caption         =   "Loans &Advanced"
            Height          =   300
            Left            =   390
            TabIndex        =   8
            Top             =   912
            Width           =   3150
         End
         Begin VB.OptionButton optLoanBalance 
            Caption         =   "&Balances where"
            Height          =   300
            Left            =   375
            TabIndex        =   4
            Top             =   390
            Width           =   3150
         End
         Begin VB.OptionButton optSubCashBook 
            Caption         =   "Sub Cash Book"
            Height          =   300
            Left            =   4140
            TabIndex        =   6
            Top             =   912
            Width           =   3120
         End
         Begin VB.OptionButton optInstOverDue 
            Caption         =   "&Instalment over due loans"
            Height          =   300
            Left            =   375
            TabIndex        =   10
            Top             =   1956
            Width           =   3150
         End
         Begin VB.OptionButton optLoanHolders 
            Caption         =   "&List of loan holders as on "
            Height          =   300
            Left            =   375
            TabIndex        =   14
            Top             =   3000
            Width           =   3180
         End
         Begin VB.OptionButton optInterest 
            Caption         =   "&Interest Collected"
            Height          =   300
            Left            =   375
            TabIndex        =   12
            Top             =   2478
            Width           =   3150
         End
         Begin VB.OptionButton optRegReports 
            Caption         =   "&Regular reports"
            Height          =   300
            Left            =   375
            TabIndex        =   16
            Top             =   3522
            Value           =   -1  'True
            Width           =   3165
         End
      End
      Begin VB.Label lblLoanType 
         AutoSize        =   -1  'True
         Caption         =   "Select a &loan type :"
         Height          =   300
         Left            =   345
         TabIndex        =   1
         Top             =   330
         Width           =   2625
      End
   End
End
Attribute VB_Name = "frmLoanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsRepOption As clsRepOption

Private M_InterestBalance As Boolean
Private m_AccNo As Long
Private m_LoanID As Long
Private m_SchemeId As Long
Private m_SchemeName As String
Public m_rstSearchResults As Recordset
Private m_rstLoanTrans As Recordset
Private m_rstLoanMast As Recordset
Private m_rstScheme As Recordset
Private M_ModuleID  As Integer
Private m_TotalInterest As Currency

Private WithEvents m_LookUp As frmLookUp
Attribute m_LookUp.VB_VarHelpID = -1

Const CTL_MARGIN = 15

Private Recursive As Boolean
Private m_accUpdatemode As Integer

Private m_PrevTab As Integer
Private m_InstalmentDetails As String
Private m_LoanAmount() As Currency

' If schemeloaded is True, the Save option
' will update the record, else it will insert.
Private m_SchemeLoaded As Boolean

' Declare events.
Public Event SetStatus(strMsg As String)
            
Public Event ShowReport(SchemeID As Integer, ReportType As wis_LoanReports, _
             ReportOrder As wis_ReportOrder, _
             fromDate As String, toDate As String, _
             LoanPurpose As String, ByRef clsReportOption As clsRepOption)
 
Public Event WindowClosed()

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
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
    
    fraReports = GetResourceString(283) & GetResourceString(92) 'Reports
    lblLoanType = GetResourceString(89) 'Select Loan Type
    
    fraChooseReport.Caption = GetResourceString(288)  'Chosse report
    optLoanBalance.Caption = GetResourceString(80, 67)  'Loan Balance
    optSubCashBook.Caption = GetResourceString(390, 85) 'Sub Day Bok
    optLoansAdvance.Caption = GetResourceString(290) 'LOan Issue
    optInstOverDue.Caption = GetResourceString(84, 18) '113 'Over Due Loans
    optInterest.Caption = GetResourceString(483) 'Interest received
    optLoanHolders.Caption = GetResourceString(83) 'Loan Holders list
    optOverDueLoans.Caption = GetResourceString(84, 18) 'Over due loans
    
    optSubDayBook.Caption = GetResourceString(390, 63) 'Sub Cash Book
    optGeneralLedger.Caption = GetResourceString(58, 93) 'Loan General Ledger
    optLoansRecovery.Caption = GetResourceString(82) 'repayment Made
    optInstOverDue.Caption = GetResourceString(113) 'Istalment Over due loans
    OptLoanSanction.Caption = GetResourceString(262) 'Sanctioned Loan
    optGuarantors.Caption = GetResourceString(58, 389) 'Guarntor
    
    fraReceivable.Caption = GetResourceString(364)  'Receivables
    optReceivable.Caption = GetResourceString(364)
    optIntReceivable.Caption = GetResourceString(376, 47, 36)
    optIntReceivableTill.Caption = GetResourceString(376, 47)
    optCustRp.Caption = GetResourceString(49, 271, 272)
    optLoanReceivable.Caption = GetResourceString(237, 364) 'Other Receivables
    Debug.Print "Kannada"
    
    optAccId.Caption = GetResourceString(58, 68)
    optName.Caption = GetResourceString(58, 69)

fraReports.Caption = GetResourceString(283) & GetResourceString(92)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
cmdView.Caption = GetResourceString(13)

lblPurpose.Caption = GetResourceString(80, 221)
cmdOk.Caption = GetResourceString(1) 'OK
cmdView.Caption = GetResourceString(13)    'View
cmdAdvance.Caption = GetResourceString(491)    'Options

End Sub

Private Sub cmbLoanType_Click()
'optLoanReceivable.Enabled = CBool(cmbLoanType.ListIndex > 0)
'If cmbLoanType.ListIndex < 1 Then optLoanReceivable.Value = False
End Sub

Private Sub cmdEndDate_Click()
With Calendar
    .Left = Me.Left + fraReports.Left + fraOrder.Left + cmdEndDate.Left - .Width
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdEndDate.Top + 300
    .selDate = txtEndDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtEndDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdAdvance_Click()
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
            
    m_clsRepOption.ShowDialog
End Sub

Private Sub cmdStDate_Click()
With Calendar
    .selDate = gStrDate
    .Left = Me.Left + fraReports.Left + fraOrder.Left + cmdStDate.Left
    .Top = Me.Top + fraReports.Top + fraOrder.Top + cmdStDate.Top + 300
    .selDate = txtStartDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtStartDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdView_Click()

' Validate the user input.
' Check for starting date.
With txtStartDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        MsgBox "Enter a valid starting date.", _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtStartDate
        Exit Sub
    End If
End With

' Check for ending date.
With txtEndDate
    If .Enabled And Not DateValidate(.Text, "/", True) Then
        MsgBox "Enter a valid ending date.", _
                vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEndDate
        Exit Sub
    End If
End With

If txtStartDate.Enabled And txtEndDate.Enabled Then
    If WisDateDiff(txtStartDate, txtEndDate) < 0 Then
        MsgBox "To date is earlier than the from date", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEndDate
        Exit Sub
    End If
End If

If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
        
Call ShowReport

Screen.MousePointer = vbDefault

End Sub
Private Sub ShowReport()

Dim SchemeID As Integer

Dim StartDate As String
Dim EndDate As String
Dim StartAmt As Currency
Dim endAmt As Currency

Dim StrCaste As String
Dim strPLace As String

Dim RetBool As Boolean
Dim ReportType As wis_LoanReports

StartDate = ""
EndDate = ""
StartAmt = 0
endAmt = 0
If txtStartDate.Enabled Then StartDate = txtStartDate
If txtEndDate.Enabled Then EndDate = txtEndDate

If cmbLoanType.ListIndex < 0 Then cmbLoanType.ListIndex = 0
    
SchemeID = cmbLoanType.ItemData(cmbLoanType.ListIndex)
If optRegReports Then

    If optSchedule(0) Then ReportType = repMonthlyRegister
    If optSchedule(1) Then ReportType = repShedule_1
    If optSchedule(2) Then ReportType = repShedule_2
    If optSchedule(3) Then ReportType = repShedule_3
    If optSchedule(4) Then ReportType = repShedule_4A
    If optSchedule(5) Then ReportType = repShedule_5
    If optSchedule(6) Then ReportType = repShedule_6
ElseIf optReceivable Then
    If optIntReceivable Then ReportType = repLoanIntReceivable
    If optLoanReceivable Then ReportType = repLoanReceivable
    If optIntReceivableTill Then ReportType = repLoanIntReceivableTill
Else
    If optLoanBalance Then
          ReportType = repLoanBalance
    ElseIf optSubCashBook Then
          ReportType = repLoanCashBook
    ElseIf optLoansAdvance Then
          ReportType = repLoanIssued
    ElseIf optInstOverDue Then
          ReportType = repLoanInstOD
    ElseIf optInterest Then
          ReportType = repLoanIntCol
    ElseIf optLoanHolders Then
          ReportType = repLoanHolder
    End If
    
    If optSubDayBook Then
          ReportType = repLoanDailyCash
    ElseIf optGeneralLedger Then
          ReportType = repLoanGLedger
    ElseIf optLoansRecovery Then
          ReportType = repLoanRepMade
    ElseIf optOverDueLoans Then
          ReportType = repLoanOD
    ElseIf OptLoanSanction Then
          ReportType = repLoanSanction
    ElseIf optGuarantors Then
          ReportType = repLoanGuarantor
    ElseIf optIntReceivable Then
          ReportType = repLoanIntReceivable
    ElseIf optCustRp Then
          ReportType = repLoanCustRP
    
    End If
    
End If

RaiseEvent ShowReport(SchemeID, ReportType, IIf(optAccId, wisByAccountNo, wisByName), _
            StartDate, EndDate, cmbPurpose.Text, m_clsRepOption)

End Sub

Private Sub Form_Load()

Call CenterMe(Me)
Call SetKannadaCaption
      
      Screen.MousePointer = vbHourglass
      Call LoadLoanSchemes(cmbLoanType)
      Call LoadLoanPurposes(cmbPurpose)
      
      cmbLoanType.AddItem "", 0
      Screen.MousePointer = vbDefault
      optLoanBalance.Value = True
      txtEndDate = gStrDate
      'fraReceivable.Left = optReceivable.Left
End Sub
Private Sub Form_Unload(Cancel As Integer)
' Report form.
If Not m_LookUp Is Nothing Then
    Unload m_LookUp
    Set m_LookUp = Nothing
End If
RaiseEvent WindowClosed
gWindowHandle = 0
End Sub


Private Sub optCustRp_Click()

    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, False)

End Sub

Private Sub optIntReceivable_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, False)


End Sub

Private Sub optIntReceivable_LostFocus()
Call optReceivable_LostFocus
End Sub

Private Sub optIntReceivableTill_Click()
     txtStartDate.Enabled = Not optIntReceivableTill.Value
     txtStartDate.BackColor = IIf(optIntReceivableTill.Value, vbGrayed, vbWhite)
End Sub

Private Sub optIntReceivableTill_LostFocus()
Call optReceivable_LostFocus
End Sub

Private Sub optLoanReceivable_Click()
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(False, False)

End Sub

Private Sub optLoanReceivable_LostFocus()
    Call optReceivable_LostFocus
End Sub

Private Sub optReceivable_Click()
    fraReceivable.Visible = optReceivable.Value
    Call optIntReceivableTill_Click
End Sub

Private Sub optReceivable_GotFocus()
    fraReceivable.Visible = True
End Sub

Private Sub optReceivable_LostFocus()
Dim ctrName As String
ctrName = Me.ActiveControl.name
If ctrName <> fraReceivable.name And ctrName <> optIntReceivable.name _
    And ctrName <> optLoanReceivable.name And ctrName <> optIntReceivableTill.name Then
        
        fraReceivable.Visible = False
End If

End Sub

Private Sub optSUbDayBook_Click()
     
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, False)

End Sub

Private Sub optGeneralLedger_Click()
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True

    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, False)

End Sub

Private Sub optGuarantors_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(False, False)

End Sub

Private Sub optInstOverDue_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub optInterest_Click()
   
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, False)

End Sub

Private Sub optLoanBalance_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True

    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub optLoanHolders_Click()
        
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub OptLoanSanction_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
   
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)


End Sub

Private Sub optLoansAdvance_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True

    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub optOverDueLoans_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False

    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)


End Sub


Private Sub optRegReports_Click()
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    
    fraSchedule.Visible = True
    fraReceivable.Visible = False
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(False, True)

End Sub

Private Sub optRegReports_GotFocus()
optRegReports.Value = True
fraSchedule.Visible = True
fraReceivable.Visible = False
End Sub


Private Sub optRegReports_LostFocus()
If Me.ActiveControl.name <> fraSchedule.name And Me.ActiveControl.name <> optSchedule(1).name Then
    fraSchedule.Visible = False
End If

End Sub


Private Sub optLoansRecovery_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True

    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub optSchedule_LostFocus(Index As Integer)

If Me.ActiveControl.name = optSchedule(1).name Or _
    ActiveControl.name = optRegReports.name Then Exit Sub
fraSchedule.Visible = False
fraReceivable.Visible = False
End Sub


Private Sub optSubCashBook_Click()
    
    fraSchedule.Visible = False
    fraReceivable.Visible = False
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    'set the Place,Caste Gender
    Call SetOptionDialogContorls(True, True)

End Sub

Private Sub txtEndAmt_GotFocus()

ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub

Private Sub txtEndDate_GotFocus()

ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub

Private Sub txtStartAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub

Private Sub SetOptionDialogContorls(EnableCastePlaceContols As Boolean, EnableAmountRange As Boolean)
    If m_clsRepOption Is Nothing Then _
            Set m_clsRepOption = New clsRepOption
    
    With m_clsRepOption
        .EnableCasteControls = EnableCastePlaceContols
        .EnableAmountRange = EnableAmountRange
    End With

End Sub
