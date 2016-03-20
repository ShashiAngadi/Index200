VERSION 5.00
Begin VB.Form frmLoanReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select report"
   ClientHeight    =   5955
   ClientLeft      =   2640
   ClientTop       =   1935
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7020
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   255
      Left            =   6060
      TabIndex        =   0
      Top             =   5670
      Width           =   915
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports..."
      Height          =   5595
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   6885
      Begin VB.Frame fraConReport 
         Caption         =   "Consol reports"
         Height          =   1185
         Left            =   4890
         TabIndex        =   49
         Top             =   2580
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton optConreport 
            Caption         =   "Over due loans"
            Height          =   270
            Index           =   2
            Left            =   60
            TabIndex        =   52
            Top             =   810
            Width           =   1650
         End
         Begin VB.OptionButton optConreport 
            Caption         =   "Instalment over due"
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   51
            Top             =   510
            Width           =   1710
         End
         Begin VB.OptionButton optConreport 
            Caption         =   "Loan Balance"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   50
            Top             =   240
            Width           =   1590
         End
      End
      Begin VB.Frame fraSchedule 
         Caption         =   "Schedules"
         Height          =   2295
         Left            =   2190
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
         Begin VB.OptionButton optSchedule 
            Caption         =   "Monthly report"
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   48
            Top             =   180
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 6"
            Height          =   285
            Index           =   6
            Left            =   30
            TabIndex        =   47
            Top             =   1950
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 4"
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   30
            TabIndex        =   46
            Top             =   1365
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 3"
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   30
            TabIndex        =   44
            Top             =   1065
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 2"
            Height          =   285
            Index           =   2
            Left            =   30
            TabIndex        =   43
            Top             =   750
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 1"
            Height          =   285
            Index           =   1
            Left            =   30
            TabIndex        =   42
            Top             =   450
            Width           =   1395
         End
         Begin VB.OptionButton optSchedule 
            Caption         =   "Schedule 5"
            Height          =   285
            Index           =   5
            Left            =   30
            TabIndex        =   41
            Top             =   1680
            Width           =   1395
         End
      End
      Begin VB.Frame fraChooseReport 
         Caption         =   "Choose a report :"
         Height          =   2460
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   6585
         Begin VB.OptionButton optRegReports 
            Caption         =   "Regular reports"
            Height          =   315
            Left            =   525
            TabIndex        =   45
            Top             =   2100
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optConsoleReport 
            Caption         =   "Consolidated reports"
            Height          =   270
            Left            =   4080
            TabIndex        =   30
            Top             =   2130
            Width           =   2370
         End
         Begin VB.OptionButton optGeneralLedger 
            Caption         =   "General Ledger"
            Height          =   270
            Left            =   4080
            TabIndex        =   17
            Top             =   555
            Width           =   2370
         End
         Begin VB.OptionButton optOverDueLoans 
            Caption         =   "Over due loans"
            Height          =   270
            Left            =   4080
            TabIndex        =   16
            Top             =   1200
            Width           =   2370
         End
         Begin VB.OptionButton optGuarantors 
            Caption         =   "Guarantor's List"
            Height          =   270
            Left            =   4080
            TabIndex        =   15
            Top             =   1830
            Width           =   2370
         End
         Begin VB.OptionButton optInterest 
            Caption         =   "Interest Collected"
            Height          =   285
            Left            =   525
            TabIndex        =   14
            Top             =   1530
            Width           =   2550
         End
         Begin VB.OptionButton optLoanHolders 
            Caption         =   "List of loan holders as on "
            Height          =   285
            Left            =   525
            TabIndex        =   13
            Top             =   1830
            Width           =   2550
         End
         Begin VB.OptionButton OptLoanSanction 
            Caption         =   "Sanctioned Loans"
            Height          =   270
            Left            =   4080
            TabIndex        =   12
            Top             =   1530
            Width           =   2370
         End
         Begin VB.OptionButton optInstOverDue 
            Caption         =   "Instalment over due loans"
            Height          =   285
            Left            =   525
            TabIndex        =   11
            Top             =   1200
            Width           =   2550
         End
         Begin VB.OptionButton optTransaction 
            Caption         =   "Total transactions made"
            Height          =   285
            Left            =   525
            TabIndex        =   10
            Top             =   555
            Width           =   2550
         End
         Begin VB.OptionButton optLoanBalance 
            Caption         =   "Balances where"
            Height          =   285
            Left            =   525
            TabIndex        =   9
            Top             =   240
            Width           =   2550
         End
         Begin VB.OptionButton optDailyCashBook 
            Caption         =   "Daily cash book"
            Height          =   270
            Left            =   4080
            TabIndex        =   8
            Top             =   240
            Width           =   2370
         End
         Begin VB.OptionButton optRepaymentsMade 
            Caption         =   "Repayments made"
            Height          =   270
            Left            =   4080
            TabIndex        =   7
            Top             =   885
            Width           =   2370
         End
         Begin VB.OptionButton optLoansIssued 
            Caption         =   "Loans issued"
            Height          =   285
            Left            =   525
            TabIndex        =   6
            Top             =   885
            Width           =   2550
         End
      End
      Begin VB.Frame fraThru 
         Height          =   465
         Left            =   150
         TabIndex        =   36
         Top             =   3030
         Width           =   6585
         Begin VB.OptionButton optBoth 
            Caption         =   "Both Loans"
            Height          =   225
            Left            =   4800
            TabIndex        =   39
            Top             =   180
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optThruSoc 
            Caption         =   "Loans Thru Societies"
            Height          =   195
            Left            =   2370
            TabIndex        =   38
            Top             =   180
            Width           =   1965
         End
         Begin VB.OptionButton optDirect 
            Caption         =   "Direct Loans"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   210
            Width           =   1965
         End
      End
      Begin VB.Frame fraCustClass 
         Height          =   555
         Left            =   150
         TabIndex        =   31
         Top             =   3450
         Width           =   6585
         Begin VB.CheckBox chkCaste 
            Caption         =   "Caste"
            Height          =   225
            Left            =   3600
            TabIndex        =   35
            Top             =   240
            Width           =   1035
         End
         Begin VB.CheckBox chkPlace 
            Caption         =   " Place"
            Height          =   285
            Left            =   135
            TabIndex        =   34
            Top             =   210
            WhatsThisHelpID =   33
            Width           =   1095
         End
         Begin VB.ComboBox cmbPlace 
            Height          =   315
            Left            =   1470
            TabIndex        =   33
            Top             =   180
            Width           =   1515
         End
         Begin VB.ComboBox cmbCaste 
            Height          =   315
            Left            =   4875
            TabIndex        =   32
            Top             =   180
            Width           =   1515
         End
      End
      Begin VB.Frame fraAmountRange 
         Caption         =   "Amount range :"
         Height          =   570
         Left            =   150
         TabIndex        =   18
         Top             =   4620
         Width           =   6585
         Begin VB.TextBox txtStartAmt 
            Height          =   285
            Left            =   1470
            TabIndex        =   20
            Top             =   195
            Width           =   1485
         End
         Begin VB.TextBox txtEndAmt 
            Height          =   285
            Left            =   4875
            TabIndex        =   19
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label lblAmt1 
            AutoSize        =   -1  'True
            Caption         =   "Between :"
            Height          =   195
            Left            =   135
            TabIndex        =   22
            Top             =   225
            Width           =   1440
         End
         Begin VB.Label lblAmt2 
            AutoSize        =   -1  'True
            Caption         =   "And :"
            Height          =   195
            Left            =   3600
            TabIndex        =   21
            Top             =   210
            Width           =   525
         End
      End
      Begin VB.Frame fraDateRange 
         Caption         =   "Date range :"
         Height          =   570
         Left            =   150
         TabIndex        =   23
         Top             =   4020
         Width           =   6585
         Begin VB.CommandButton cmdEndDate 
            Caption         =   "..."
            Height          =   250
            Left            =   6120
            TabIndex        =   27
            Top             =   210
            Width           =   250
         End
         Begin VB.CommandButton cmdStDate 
            Caption         =   "..."
            Height          =   250
            Left            =   2715
            TabIndex        =   26
            Top             =   210
            Width           =   250
         End
         Begin VB.TextBox txtEndDate 
            Height          =   285
            Left            =   4875
            TabIndex        =   25
            Top             =   180
            Width           =   1200
         End
         Begin VB.TextBox txtStartDate 
            Height          =   285
            Left            =   1470
            TabIndex        =   24
            Top             =   195
            Width           =   1200
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "Ending date :"
            Height          =   195
            Left            =   3600
            TabIndex        =   29
            Top             =   210
            Width           =   1155
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "Starting date :"
            Height          =   195
            Left            =   135
            TabIndex        =   28
            Top             =   225
            Width           =   1170
         End
      End
      Begin VB.ComboBox cmbLoanType 
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   4485
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   285
         Left            =   5730
         TabIndex        =   3
         Top             =   5220
         Width           =   975
      End
      Begin VB.Label lblLoanType 
         AutoSize        =   -1  'True
         Caption         =   "Select a loan type :"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   330
         Width           =   1365
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
Private M_LoanId As Long
Private m_SchemeId As Long
Private m_SchemeName As String
Public m_rstSearchResults As Recordset
Private m_rstLoanTrans As Recordset
Private m_rstLoanMast As Recordset
Private m_rstScheme As Recordset
Private m_ModuleID  As Integer
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

' Declare events.
Public Event SetStatus(strMsg As String)



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
Private Sub ShowReport()
Dim SchemeID As Integer

Dim StartDate As String
Dim EndDate As String
Dim StartAmt As Currency
Dim EndAmt As Currency

Dim StrCaste As String
Dim StrPlace As String
Dim RepSociety As wisRepSocietyLoans


Dim RetBool As Boolean
StartDate = ""
EndDate = ""
StartAmt = 0
EndAmt = 0
If txtStartDate.Enabled Then StartDate = txtStartDate
If txtEndDate.Enabled Then EndDate = txtEndDate
If txtStartAmt.Enabled Then StartAmt = Val(txtStartAmt)
If txtEndAmt.Enabled Then EndAmt = Val(txtEndAmt)
If cmbLoanType.ListIndex < 0 Then
    SchemeID = 0
Else
    SchemeID = cmbLoanType.ItemData(cmbLoanType.ListIndex)
End If

If cmbCaste.Enabled Then
    StrCaste = Trim$(cmbCaste.Text)
End If
If cmbPlace.Enabled Then
    StrPlace = Trim$(cmbPlace.Text)
End If
If Not optDirect Then
    StrCaste = ""
    StrPlace = ""
End If

If optDirect Then RepSociety = DirectLoans
If optThruSoc Then RepSociety = SocietyThruLoans
If optBoth Then RepSociety = BothLoans

gCancel = False
frmCancel.Show
frmCancel.Refresh
Screen.MousePointer = vbHourglass
If optRegReports Then
    If optSchedule(0) Then RetBool = frmRegReport.ShowMeetingRegistar(SchemeID, StartDate, BothLoans)
    If optSchedule(1) Then RetBool = frmRegReport.ShowShed1(StartDate, RepSociety)
    If optSchedule(2) Then RetBool = frmRegReport.ShowShed2(StartDate, RepSociety)
    If optSchedule(5) Then RetBool = frmRegReport.ShowShed5(StartDate)
    If optSchedule(6) Then RetBool = frmRegReport.ShowShed6(StartDate)
    Unload frmCancel
    Set frmCancel = Nothing
    Screen.MousePointer = vbDefault
    If gCancel Then Exit Sub
    If RetBool Then frmRegReport.Show vbModal
    Exit Sub
End If
If optConsoleReport Then
    If optConreport(0) Then RetBool = frmRegReport.ShowConsoleBalance(StartDate)
    If optConreport(1) Then RetBool = frmRegReport.ShowConsoleInstOverDue(StartDate)
    If optConreport(2) Then RetBool = frmRegReport.ShowConsoleODBalance(StartDate)
    
    Unload frmCancel
    Set frmCancel = Nothing
    Screen.MousePointer = vbDefault
    If gCancel Then Exit Sub
    If RetBool Then frmRegReport.Show vbModal
    Exit Sub
End If


Set m_frmLoanReport = New frmLoanView
With m_frmLoanReport
  If optLoanBalance Then _
    RetBool = .ReportLoanBalance(SchemeID, StartDate, EndDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  If optTransaction Then _
    RetBool = .ReportTransactionMade(SchemeID, StartDate, EndDate, StartAmt, EndAmt, RepSociety)
  If optLoansIssued Then _
    RetBool = .ReportLoanIssued(SchemeID, StartDate, EndDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  If optInstOverDue Then _
    RetBool = .ReportOverdueInstalments(SchemeID, StartDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  If optInterest Then _
    RetBool = .ReportInterestRecieved(SchemeID, StartDate, EndDate)
  If optLoanHolders Then _
    RetBool = .ReportLoanHolders(SchemeID, StartDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  
  
  If optDailyCashBook Then _
    RetBool = .ReportDailyCashbook(SchemeID, StartDate, EndDate, StartAmt, EndAmt)
  If optGeneralLedger Then _
    RetBool = .ReportGeneralLedger(SchemeID, StartDate, EndDate, StartAmt, EndAmt)
  If optRepaymentsMade Then _
    RetBool = .ReportPaymentsMade(SchemeID, StartDate, EndDate, StartAmt, EndAmt)
  If optOverDueLoans Then _
    RetBool = .ReportOverdueLoans(SchemeID, StartDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  If OptLoanSanction Then _
    RetBool = .ReportSanctionedLoans(SchemeID, StartDate, EndDate, StartAmt, EndAmt, StrCaste, StrPlace, RepSociety)
  If optGuarantors Then _
    RetBool = .ReportGuarantors(SchemeID, StartDate)
End With

Unload frmCancel
Set frmCancel = Nothing
Screen.MousePointer = vbDefault

If gCancel Then RetBool = False
gCancel = False
If RetBool Then m_frmLoanReport.Show vbModal

ExitLine:

Set m_frmLoanReport = Nothing

End Sub

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





Private Sub cmbCaste_Click()
If cmbCaste.ListIndex >= 0 Then
    cmbCaste.Tag = cmbCaste.ListIndex
End If
End Sub


Private Sub cmbPlace_Click()
If cmbPlace.ListIndex >= 0 Then
    cmbPlace.Tag = cmbPlace.ListIndex
End If
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

Private Sub cmdOK_Click()
Dim Cancel As Boolean
Unload Me
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

Call ShowReport

Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
      
      Screen.MousePointer = vbHourglass
      'set icon for the form caption
      
      'Centre the form
      
      'In Reports TaB Load The Combos with Caste and Places  Respectively
      chkCaste.Enabled = False
      cmbCaste.Enabled = False
      cmbCaste.BackColor = wisGray
      chkPlace.Enabled = False
      cmbPlace.Enabled = False
      cmbPlace.BackColor = wisGray
      
      Call LoadCastes
      Call LoadPlaces
      
      Call LoadLoanSchemes(cmbLoanType)
      cmbLoanType.AddItem "", 0
      cmbLoanType.ItemData(cmbLoanType.NewIndex) = 0
      Screen.MousePointer = vbDefault
      
      fraSchedule.Visible = False
      optLoanBalance.value = True
    

End Sub
Private Sub LoadPlaces()
    Dim i As Integer
    Me.cmbPlace.Clear
   gDbTrans.SqlStmt = "Select Place from PlaceTab order by Place"
    If gDbTrans.SQLFetch > 0 Then
    gDbTrans.Rst.MoveFirst
        For i = 1 To gDbTrans.Records
            If gDbTrans.Rst("Place") <> "" Then
                cmbPlace.AddItem gDbTrans.Rst("Place")
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
 gDbTrans.SqlStmt = "Select Caste from CasteTab order by Caste"
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

Private Sub m_frmLoanReport_Processing(strMessage As String, ratio As Single)
On Error Resume Next
With frmCancel
    .lblMessage = "PROCESS: " & vbCrLf & strMessage
    If ratio > 0 Then
        If ratio > 1 Then
                .prg.value = ratio * .prg.Max
        End If
    End If
End With
End Sub


Private Sub optBoth_Click()
    chkPlace.Enabled = False
    cmbPlace.ListIndex = -1
    cmbPlace.Enabled = False
    chkCaste.Enabled = False
    cmbCaste.ListIndex = -1
    cmbCaste.Enabled = False
End Sub

Private Sub optConreport_LostFocus(Index As Integer)
    If Me.ActiveControl.Name = optConreport(1).Name Or Me.ActiveControl.Name = optConsoleReport.Name Then Exit Sub
    fraConReport.Visible = False

End Sub

Private Sub optConsoleReport_Click()
    
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
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
    'chkPlace.BackColor = wisGray

End Sub

Private Sub optConsoleReport_GotFocus()
fraConReport.Visible = True
End Sub

Private Sub optConsoleReport_LostFocus()
If ActiveControl.Name = optConreport(1).Name Then Exit Sub
fraConReport.Visible = False
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

Private Sub optDirect_Click()
    chkPlace.Enabled = True
    cmbPlace.Enabled = True
    chkCaste.Enabled = True
    cmbCaste.Enabled = True

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
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
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

Private Sub Option1_Click()

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
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
    'chkPlace.BackColor = wisGray
End Sub

Private Sub optLoanHolders_Click()
    
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
    chkCaste.Enabled = True
    chkPlace.Enabled = True
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
    'chkPlace.BackColor = wisGray
End Sub


Private Sub OptLoanSanction_Click()
    
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
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

Private Sub optLoansIssued_Click()
    
    txtEndDate.Enabled = True
    txtEndDate.BackColor = vbWhite
    cmdEndDate.Enabled = True
    txtStartDate.Enabled = True
    txtStartDate.BackColor = vbWhite
    cmdStDate.Enabled = True
    txtStartAmt.Enabled = True
    txtStartAmt.BackColor = vbWhite
    txtEndAmt.Enabled = True
    txtEndAmt.BackColor = vbWhite
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
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
    chkCaste.Enabled = True: chkCaste.value = vbChecked
    'chkCaste.BackColor = wisGray
    chkPlace.Enabled = True: chkPlace.value = vbChecked
    'chkPlace.BackColor = wisGray
End Sub


Private Sub optRegReports_Click()
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
    
    fraSchedule.Visible = True
    
End Sub

Private Sub optRegReports_GotFocus()
optRegReports.value = True
fraSchedule.Visible = True
End Sub

Private Sub optRegReports_LostFocus()
If Me.ActiveControl.Name <> fraSchedule.Name And Me.ActiveControl.Name <> optSchedule(1).Name Then
    fraSchedule.Visible = False
End If
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

Private Sub optSchedule_LostFocus(Index As Integer)
If Me.ActiveControl.Name = optSchedule(1).Name Or ActiveControl.Name = optRegReports.Name Then Exit Sub
fraSchedule.Visible = False
End Sub


Private Sub optThruSoc_Click()
    chkPlace.Enabled = False
    cmbPlace.ListIndex = -1
    cmbPlace.Enabled = False
    chkCaste.Enabled = False
    cmbCaste.ListIndex = -1
    cmbCaste.Enabled = False
    
End Sub

Private Sub optTransaction_Click()
    
    fraSchedule.Visible = False
    
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


Private Sub txtEndAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtEndDate_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txtStartAmt_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


