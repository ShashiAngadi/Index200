VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmLoanReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select report"
   ClientHeight    =   7335
   ClientLeft      =   1290
   ClientTop       =   1200
   ClientWidth     =   7950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   345
      Left            =   5760
      TabIndex        =   38
      Top             =   6900
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   345
      Left            =   6840
      TabIndex        =   37
      Top             =   6900
      Width           =   975
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   6735
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7785
      Begin VB.Frame fraSchedule 
         Caption         =   "Schedules"
         Height          =   2445
         Left            =   1920
         TabIndex        =   19
         Top             =   1530
         Visible         =   0   'False
         Width           =   1905
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &5"
            Height          =   285
            Index           =   5
            Left            =   30
            TabIndex        =   25
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &1"
            Height          =   285
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   450
            Width           =   1785
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &2"
            Height          =   285
            Index           =   2
            Left            =   30
            TabIndex        =   22
            Top             =   750
            Width           =   1725
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &3"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   30
            TabIndex        =   23
            Top             =   1065
            Width           =   1665
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &4"
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   30
            TabIndex        =   24
            Top             =   1365
            Width           =   1695
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule &6"
            Height          =   285
            Index           =   6
            Left            =   30
            TabIndex        =   26
            Top             =   1980
            Width           =   1695
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "&Monthly report"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   180
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin VB.Frame fraOrder 
         Caption         =   " List Order"
         Height          =   2565
         Left            =   150
         TabIndex        =   27
         Top             =   4080
         Width           =   7425
         Begin VB.TextBox txtStartDate 
            Height          =   315
            Left            =   1845
            TabIndex        =   44
            Top             =   1620
            Width           =   1290
         End
         Begin VB.TextBox txtEndDate 
            Height          =   315
            Left            =   5745
            TabIndex        =   43
            Top             =   1620
            Width           =   1230
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   285
            Left            =   3210
            TabIndex        =   42
            Top             =   1620
            Width           =   285
         End
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   285
            Left            =   7035
            TabIndex        =   41
            Top             =   1620
            Width           =   285
         End
         Begin VB.ComboBox cmbAccGroup 
            Height          =   315
            Left            =   4560
            TabIndex        =   36
            Top             =   210
            Width           =   2775
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   315
            Left            =   150
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   315
            Left            =   2550
            TabIndex        =   31
            Top             =   210
            Width           =   2475
         End
         Begin VB.ComboBox cmbCastes 
            Height          =   315
            Left            =   3045
            TabIndex        =   30
            Top             =   990
            Width           =   1800
         End
         Begin VB.ComboBox cmbPlaces 
            Height          =   315
            Left            =   180
            TabIndex        =   29
            Top             =   990
            Width           =   1860
         End
         Begin VB.ComboBox cmbGender 
            Height          =   315
            Left            =   5730
            TabIndex        =   28
            Top             =   990
            Width           =   1620
         End
         Begin WIS_Currency_Text_Box.CurrText txtStartAmt 
            Height          =   345
            Left            =   1830
            TabIndex        =   48
            Top             =   2040
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin WIS_Currency_Text_Box.CurrText txtEndAmt 
            Height          =   345
            Left            =   5760
            TabIndex        =   46
            Top             =   2010
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            CurrencySymbol  =   ""
            TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
            NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
            FontSize        =   8.25
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   7350
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "&Starting date :"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   1665
            Width           =   990
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "&Ending date :"
            Height          =   195
            Left            =   3855
            TabIndex        =   49
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label lblAmt1 
            AutoSize        =   -1  'True
            Caption         =   "Between :"
            Height          =   255
            Left            =   285
            TabIndex        =   47
            Top             =   2070
            Width           =   1200
         End
         Begin VB.Label lblAmt2 
            AutoSize        =   -1  'True
            Caption         =   "And :"
            Height          =   255
            Left            =   3915
            TabIndex        =   45
            Top             =   2085
            Width           =   1695
         End
         Begin VB.Line Line1 
            X1              =   7290
            X2              =   30
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblPlace 
            Caption         =   "Place"
            Height          =   225
            Left            =   240
            TabIndex        =   35
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label lblCaste 
            Caption         =   "Caste"
            Height          =   255
            Left            =   3390
            TabIndex        =   34
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label lblGender 
            Caption         =   "Label1"
            Height          =   255
            Left            =   5820
            TabIndex        =   33
            Top             =   690
            Width           =   1275
         End
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report :"
         Height          =   3420
         Left            =   150
         TabIndex        =   3
         Top             =   660
         Width           =   7425
         Begin VB.OptionButton optLoanReceivable 
            Caption         =   "Other receivable"
            Height          =   285
            Left            =   390
            TabIndex        =   40
            Top             =   3000
            Width           =   3150
         End
         Begin VB.OptionButton optCustRp 
            Caption         =   "Customers Receipt && Payment"
            Height          =   255
            Left            =   4140
            TabIndex        =   18
            Top             =   3030
            Width           =   3075
         End
         Begin VB.OptionButton optIntReceivable 
            Caption         =   "Interest Receivable Balance"
            Height          =   315
            Left            =   4140
            TabIndex        =   17
            Top             =   2640
            Width           =   3105
         End
         Begin VB.OptionButton optGeneralLedger 
            Caption         =   "&General Ledger"
            Height          =   270
            Left            =   4140
            TabIndex        =   7
            Top             =   1050
            Width           =   3060
         End
         Begin VB.OptionButton optOverDueLoans 
            Caption         =   "&Over due loans"
            Height          =   270
            Left            =   4140
            TabIndex        =   11
            Top             =   1458
            Width           =   3030
         End
         Begin VB.OptionButton optGuarantors 
            Caption         =   "&Guarantor's List"
            Height          =   270
            Left            =   4140
            TabIndex        =   15
            Top             =   2250
            Width           =   3060
         End
         Begin VB.OptionButton OptLoanSanction 
            Caption         =   "S&anctioned Loans"
            Height          =   270
            Left            =   4140
            TabIndex        =   13
            Top             =   1854
            Width           =   3030
         End
         Begin VB.OptionButton optSubDayBook 
            Caption         =   "Sub day Book"
            Height          =   270
            Left            =   4140
            TabIndex        =   5
            Top             =   270
            Width           =   3000
         End
         Begin VB.OptionButton optLoansRecovery 
            Caption         =   "Loan &Recovery"
            Height          =   270
            Left            =   390
            TabIndex        =   9
            Top             =   1050
            Width           =   2730
         End
         Begin VB.OptionButton optLoansAdvance 
            Caption         =   "Loans &Advanced"
            Height          =   285
            Left            =   390
            TabIndex        =   8
            Top             =   660
            Width           =   3150
         End
         Begin VB.OptionButton optLoanBalance 
            Caption         =   "&Balances where"
            Height          =   285
            Left            =   375
            TabIndex        =   4
            Top             =   270
            Width           =   3150
         End
         Begin VB.OptionButton optSubCashBook 
            Caption         =   "Sub Cash Book"
            Height          =   285
            Left            =   4140
            TabIndex        =   6
            Top             =   630
            Width           =   3120
         End
         Begin VB.OptionButton optInstOverDue 
            Caption         =   "&Instalment over due loans"
            Height          =   285
            Left            =   375
            TabIndex        =   10
            Top             =   1455
            Width           =   3150
         End
         Begin VB.OptionButton optLoanHolders 
            Caption         =   "&List of loan holders as on "
            Height          =   285
            Left            =   375
            TabIndex        =   14
            Top             =   2245
            Width           =   3180
         End
         Begin VB.OptionButton optInterest 
            Caption         =   "&Interest Collected"
            Height          =   285
            Left            =   375
            TabIndex        =   12
            Top             =   1850
            Width           =   3150
         End
         Begin VB.OptionButton optRegReports 
            Caption         =   "&Regular reports"
            Height          =   315
            Left            =   375
            TabIndex        =   16
            Top             =   2640
            Value           =   -1  'True
            Width           =   3165
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
      Begin VB.Label lblLoanType 
         AutoSize        =   -1  'True
         Caption         =   "Select a &loan type :"
         Height          =   195
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
'Private WithEvents m_frmLoanReport  As frmLoanView
'Private WithEvents m_frmReportLoan  As frmRegReport

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
             FromDate As String, ToDate As String, _
             FromAmount As Currency, ToAmount As Currency, _
             Gender As wis_Gender, Place As String, _
             Caste As String, AccountGroup As Integer)
             
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
    
    fraReports = LoadResString(gLangOffSet + 283) & LoadResString(gLangOffSet + 92) 'Reports
    lblLoanType = LoadResString(gLangOffSet + 89) 'Select Loan Type
    
    fraChooseReport.Caption = LoadResString(gLangOffSet + 288)  'Chosse report
    optLoanBalance.Caption = LoadResString(gLangOffSet + 80) & " " & LoadResString(gLangOffSet + 67)   'Loan Balance
    optSubCashBook.Caption = LoadResString(gLangOffSet + 390) & " " & LoadResString(gLangOffSet + 85) 'Sub Day Bok
    optLoansAdvance.Caption = LoadResString(gLangOffSet + 290) 'LOan Issue
    optInstOverDue.Caption = LoadResString(gLangOffSet + 84) & " " & LoadResString(gLangOffSet + 18) '113 'Over Due Loans
    optInterest.Caption = LoadResString(gLangOffSet + 483) 'Interest received
    optLoanHolders.Caption = LoadResString(gLangOffSet + 83) 'Loan Holders list
    optOverDueLoans.Caption = LoadResString(gLangOffSet + 84) & " " & LoadResString(gLangOffSet + 18) 'Over due loans
    
    optSubDayBook.Caption = LoadResString(gLangOffSet + 390) & " " & LoadResString(gLangOffSet + 63) 'Sub Cash Book
    optGeneralLedger.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 93)  'Loan General Ledger
    optLoansRecovery.Caption = LoadResString(gLangOffSet + 82) 'repayment Made
    optInstOverDue.Caption = LoadResString(gLangOffSet + 113) 'Istalment Over due loans
    OptLoanSanction.Caption = LoadResString(gLangOffSet + 262) 'Sanctioned Loan
    optGuarantors.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 389)  'Guarntor
    optIntReceivable.Caption = LoadResString(gLangOffSet + 376) & " " & LoadResString(gLangOffSet + 47)  'Guarntor
    optCustRp.Caption = LoadResString(gLangOffSet + 49) & " " & _
        LoadResString(gLangOffSet + 271) & " " & LoadResString(gLangOffSet + 272)
    optLoanReceivable.Caption = LoadResString(gLangOffSet + 237) & " " & LoadResString(gLangOffSet + 364)  'Other Receivables
    Debug.Print "Kannada"
    'fraSchedule.Caption
    'optRegReports.Caption
    'optSchedule(0).Caption
    'optSchedule(1).Caption
    'optSchedule(2).Caption
    'optSchedule(3).Caption
    'optSchedule(4).Caption
    'optSchedule(5).Caption
    'optSchedule(6).Caption
    
    optAccId.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 68)
    optName.Caption = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 69)
    lblPlace.Caption = LoadResString(gLangOffSet + 112)
    lblCaste.Caption = LoadResString(gLangOffSet + 111)
    lblGender.Caption = LoadResString(gLangOffSet + 125)

fraReports.Caption = LoadResString(gLangOffSet + 283) & LoadResString(gLangOffSet + 92)
'fraDaterange.Caption = LoadResString(gLangOffSet + 106)
lblDate1.Caption = LoadResString(gLangOffSet + 109)
lblDate2.Caption = LoadResString(gLangOffSet + 110)
lblAmt1.Caption = LoadResString(gLangOffSet + 147) & " " & LoadResString(gLangOffSet + 42)
lblAmt2.Caption = LoadResString(gLangOffSet + 148) & " " & LoadResString(gLangOffSet + 42)
cmdView.Caption = LoadResString(gLangOffSet + 13)


cmdOK.Caption = LoadResString(gLangOffSet + 1) 'OK
cmdView.Caption = LoadResString(gLangOffSet + 13)    'View
    
End Sub

Private Sub SetSortDetail(Enable As Boolean)

With cmbCastes
    .Enabled = Enable
    If Not Enable Then .ListIndex = 0
    .BackColor = IIf(Enable, wisWhite, wisGray)
End With
lblCaste.Enabled = Enable

With cmbPlaces
    .Enabled = Enable
    If Not Enable Then .ListIndex = 0
    .BackColor = IIf(Enable, wisWhite, wisGray)
End With

lblPlace.Enabled = Enable

With cmbGender
    .Enabled = Enable
    If Not Enable Then .ListIndex = 0
    .BackColor = IIf(Enable, wisWhite, wisGray)
End With
lblGender.Enabled = Enable

End Sub

Private Sub cmbLoanType_Click()
optLoanReceivable.Enabled = CBool(cmbLoanType.ListIndex > 0)
If cmbLoanType.ListIndex < 1 Then optLoanReceivable.Value = False
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
Dim Cancel As Boolean


Unload Me
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

' Check for starting amount.
With txtStartAmt
    If .Enabled Then
        If .Text <> "" And Not IsNumeric(.Text) Then
            MsgBox "Enter a valid starting amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtStartAmt
            Exit Sub
        End If
    End If
End With
    
' Check for ending amount.
With txtEndAmt
    If .Enabled Then
        If .Text <> "" And Not IsNumeric(.Text) Then
            MsgBox "Enter a valid ending amount.", _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtEndAmt
            Exit Sub
        End If
    End If
End With

If cmbGender.ListIndex < 0 Then cmbGender.ListIndex = 0
If cmbLoanType.ListIndex < 0 Then cmbLoanType.ListIndex = 0

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
If txtStartAmt.Enabled Then StartAmt = txtStartAmt
If txtEndAmt.Enabled Then endAmt = txtEndAmt

If cmbLoanType.ListIndex < 0 Then cmbLoanType.ListIndex = 0
    
SchemeID = cmbLoanType.ItemData(cmbLoanType.ListIndex)
If cmbCastes.Enabled Then StrCaste = cmbCastes.Text
If cmbPlaces.Enabled Then strPLace = cmbPlaces.Text
If cmbGender.ListIndex < 0 Then cmbGender.ListIndex = 0
If cmbAccGroup.ListIndex < 0 Then cmbAccGroup.ListIndex = 0

If optRegReports Then

    If optSchedule(0) Then ReportType = repMonthlyRegister
    If optSchedule(1) Then ReportType = repShedule_1
    If optSchedule(2) Then ReportType = repShedule_2
    If optSchedule(3) Then ReportType = repShedule_3
    If optSchedule(4) Then ReportType = repShedule_4A
    If optSchedule(5) Then ReportType = repShedule_5
    If optSchedule(6) Then ReportType = repShedule_6
    
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
    ElseIf optLoanReceivable Then
          ReportType = repLoanReceivable
    End If
    
End If



RaiseEvent ShowReport(SchemeID, ReportType, IIf(optAccId, wisByAccountNo, wisByName), _
            StartDate, EndDate, StartAmt, endAmt, cmbGender.ItemData(cmbGender.ListIndex), _
            strPLace, StrCaste, cmbAccGroup.ItemData(cmbAccGroup.ListIndex))


End Sub



Private Sub Form_Load()

'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption
      
      Screen.MousePointer = vbHourglass
      'set icon for the form caption
      
      'Centre the form
      
      'In Reports TaB Load The Combos with Caste and Places  Respectively
      Call LoadCastes(cmbCastes)
      Call LoadPlaces(cmbPlaces)
      Call LoadGender(cmbGender)
      Call LoadAccountGroups(cmbAccGroup)
      
      With cmbCastes
        .Enabled = False
        .ListIndex = 0
        .BackColor = wisGray
      End With
      
      With cmbPlaces
        .Enabled = False
        .ListIndex = 0
        .BackColor = wisGray
      End With
      
      With cmbGender
        .Enabled = False
        .ListIndex = 0
        .BackColor = wisGray
      End With
      
      Call LoadLoanSchemes(cmbLoanType)
      cmbLoanType.AddItem "", 0
      Screen.MousePointer = vbDefault
      optLoanBalance.Value = True
      txtEndDate = gStrDate
      
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

Private Sub m_frmLoanReport_Initialise(Min As Long, Max As Long)

End Sub

Private Sub m_frmLoanReport_Processing(strMessage As String, Ratio As Single)

End Sub


Private Sub m_frmReportLoan_Initialise(Min As Long, Max As Long)


End Sub

 
Private Sub m_frmReportLoan_Processing(strMessage As String, Ratio As Single)

End Sub
Private Sub optCustRp_Click()

    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub

Private Sub optIntReceivable_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    'Call SetSortDetail(True)

End Sub

Private Sub optLoanReceivable_Click()
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray

    'set the Place,Caste Gender
    Call SetSortDetail(False)

End Sub

Private Sub optSUbDayBook_Click()
     
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub

Private Sub optGeneralLedger_Click()
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    
    'set the Place,Caste Gender
    Call SetSortDetail(False)
End Sub

Private Sub optGuarantors_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    Call SetSortDetail(False)
    
End Sub

Private Sub optInstOverDue_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub

Private Sub optInterest_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    Call SetSortDetail(True)
End Sub

Private Sub optLoanBalance_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    
    'set the Place,Caste Gender
    Call SetSortDetail(True)
End Sub

Private Sub optLoanHolders_Click()
        
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub


Private Sub OptLoanSanction_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    'set the Place,Caste Gender
    Call SetSortDetail(True)
   
End Sub

Private Sub optLoansAdvance_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub

Private Sub optOverDueLoans_Click()
    
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    'set the Place,Caste Gender
    Call SetSortDetail(True)

End Sub


Private Sub optRegReports_Click()
    txtStartDate.Enabled = False
    txtStartDate.BackColor = wisGray
    cmdStDate.Enabled = False
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    
    'set the Place,Caste Gender
    Call SetSortDetail(False)
    
    fraSchedule.Visible = True
    
End Sub

Private Sub optRegReports_GotFocus()
optRegReports.Value = True
fraSchedule.Visible = True

End Sub


Private Sub optRegReports_LostFocus()
If Me.ActiveControl.Name <> fraSchedule.Name And Me.ActiveControl.Name <> optSchedule(1).Name Then
    fraSchedule.Visible = False
End If

End Sub


Private Sub optLoansRecovery_Click()
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = False
    txtStartAmt.BackColor = wisGray
    txtEndAmt.Enabled = False
    txtEndAmt.BackColor = wisGray
    'set the Place,Caste Gender
    Call SetSortDetail(True)
End Sub

Private Sub optSchedule_LostFocus(Index As Integer)

If Me.ActiveControl.Name = optSchedule(1).Name Or _
    ActiveControl.Name = optRegReports.Name Then Exit Sub
fraSchedule.Visible = False

End Sub


Private Sub optSubCashBook_Click()
    
    fraSchedule.Visible = False
    
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite

    'set the Place,Caste Gender
    Call SetSortDetail(True)

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


