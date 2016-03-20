VERSION 5.00
Begin VB.Form frmAccReports 
   Caption         =   "Account Reports ..."
   ClientHeight    =   3795
   ClientLeft      =   2025
   ClientTop       =   2190
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDate1 
      Enabled         =   0   'False
      Height          =   395
      Left            =   1920
      TabIndex        =   7
      Top             =   2310
      Width           =   1275
   End
   Begin VB.TextBox txtDate2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   5490
      TabIndex        =   9
      Text            =   "12/12/2222"
      Top             =   2280
      Width           =   1185
   End
   Begin VB.ComboBox cmbReportList 
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   1
      Text            =   "Report List"
      Top             =   180
      Width           =   4005
   End
   Begin VB.ComboBox cmbRepParentHead 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   990
      Width           =   4035
   End
   Begin VB.ComboBox cmbRepHeadID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1500
      Width           =   4035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5580
      TabIndex        =   11
      Top             =   3150
      Width           =   1185
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   4050
      TabIndex        =   10
      Top             =   3150
      Width           =   1185
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   6730
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label lblDate1 
      Caption         =   "From Date"
      Height          =   390
      Left            =   90
      TabIndex        =   6
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label lblDate2 
      Caption         =   "To Date"
      Height          =   390
      Left            =   3840
      TabIndex        =   8
      Top             =   2295
      Width           =   1485
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   6730
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6730
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Select Report Type"
      BeginProperty Font 
         Name            =   "Nudi B-Akshar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label lblRepAccHead 
      Caption         =   " Account Head  :"
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   2535
   End
   Begin VB.Label lblRepAccName 
      Caption         =   "Account Name :"
      Height          =   480
      Left            =   60
      TabIndex        =   4
      Top             =   1530
      Width           =   2535
   End
End
Attribute VB_Name = "frmAccReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event OKClick(StDate As String, EndDate As String, ParentID As Long, headID As Long, ReportSelected As Wis_AccountReportList)
Public Event CancelClick()

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'fraReport.Caption = LoadResString(gLangOffSetNew + 505) & " " & GetResourceString(27)
lblRepAccHead.Caption = GetResourceString(232)
lblRepAccName.Caption = GetResourceString(36, 35)
lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
Me.Label1.Caption = GetResourceString(288)
cmdView.Caption = GetResourceString(13)
cmdCancel.Caption = GetResourceString(2)
End Sub

Private Sub cmbReportList_Click()

Dim ReportList As Wis_AccountReportList

With cmbReportList
    If .ListIndex = -1 Then Exit Sub
    ReportList = .ItemData(.ListIndex)
End With

cmbRepParentHead.Enabled = False
cmbRepHeadID.Enabled = False
txtDate1.Enabled = False
txtDate2.Enabled = False
lblDate1.Enabled = False
lblDate2.Enabled = False
cmdView.Enabled = False

Select Case ReportList

    Case AccountLedger
        
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        txtDate1.Enabled = True
        txtDate2.Enabled = True
        lblDate1.Enabled = True
        lblDate2.Enabled = True
        cmdView.Enabled = True
    
    Case AccountLedgerOnDate
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        txtDate2.Enabled = True
        lblDate2.Enabled = True
        cmdView.Enabled = True
                
    Case DayBook
        txtDate2.Enabled = True
        lblDate2.Enabled = True
        cmdView.Enabled = True
    
    Case AccountsClosed
    
    Case SubDayBook
        
        cmbRepParentHead.Enabled = True
        lblDate2.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True
    
    Case BalancesAsON
    
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        lblDate2.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True
        

    Case GeneralLedger
        lblDate1.Enabled = True
        txtDate1.Enabled = True
        lblDate2.Enabled = True
        txtDate2.Enabled = True
        cmdView.Enabled = True
        
    Case AccountGeneralLedger
        lblDate1.Enabled = True
        txtDate1.Enabled = True
        lblDate2.Enabled = True
        txtDate2.Enabled = True
        cmbRepParentHead.Enabled = True
        cmbRepHeadID.Enabled = True
        cmdView.Enabled = True
        
    Case ProfitandLossTrans
    Case ReportNothing
    Case TotalTransActionsMade
    
End Select


End Sub


Private Sub cmbRepParentHead_Click()

With cmbRepParentHead
    If .ListIndex = -1 Then Exit Sub
    Call LoadLedgersToCombo(cmbRepHeadID, .ItemData(.ListIndex))
End With

End Sub


Private Sub cmdCancel_Click()

RaiseEvent OKClick("", "", 0, 0, 0)
Unload Me
End Sub


Private Sub cmdView_Click()

Dim StartDate As String
Dim EndDate As String
Dim ParentID As Long
Dim headID As Long

Dim ReportList As Wis_AccountReportList

With cmbReportList
    If .ListIndex = -1 Then Exit Sub
    ReportList = .ItemData(.ListIndex)
End With
   

If cmbRepParentHead.Enabled Then
    With cmbRepParentHead
        If Not .ListIndex = -1 Then ParentID = .ItemData(.ListIndex)
    End With
End If

If (ReportList = AccountLedger Or ReportList = AccountGeneralLedger _
                    Or ReportList = SubDayBook) And ParentID = 0 Then Exit Sub

If cmbRepHeadID.Enabled Then
    With cmbRepHeadID
        If Not .ListIndex = -1 Then headID = .ItemData(.ListIndex)
    End With
End If

If ReportList = AccountLedger And headID = 0 Then Exit Sub
If ReportList = AccountLedgerOnDate And headID = 0 Then Exit Sub

If txtDate1.Enabled Then
    StartDate = txtDate1.Text
    'Check For Validate of Dates
    If Not DateValidate(StartDate, "/", True) Then
        txtDate1.SetFocus
        Exit Sub
    End If
Else
    StartDate = txtDate2.Text
End If

'If txtDate2.Enabled Then
    EndDate = txtDate2.Text
    If Not DateValidate(EndDate, "/", True) Then
        txtDate2.SetFocus
        Exit Sub
    End If
'End If

If DateDiff("d", GetSysFormatDate(StartDate), GetSysFormatDate(EndDate)) < 0 Then
    MsgBox "Start date should be earlier than the end date ", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Me.Hide

Screen.MousePointer = vbHourglass
RaiseEvent OKClick(StartDate, EndDate, ParentID, headID, ReportList)

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

CenterMe Me

Me.Icon = LoadResPicture(147, vbResIcon)

If gLangOffSet <> 0 Then SetKannadaCaption

LoadReportList

Call LoadParentHeads(cmbRepParentHead, True)
cmbRepParentHead.AddItem "", 0
cmbRepParentHead.ItemData(cmbRepParentHead.newIndex) = 0
End Sub

Private Sub LoadReportList()

Dim ReportList As Wis_AccountReportList



With cmbReportList
    .Text = ""
    
    ReportList = BalancesAsON
    .AddItem GetResourceString(67) & " " & _
            GetFromDateString(GetResourceString(37)) '"Balances As On"
    .ItemData(.newIndex) = ReportList
    
    ReportList = AccountsClosed
    '"Account Closed"
    '.ItemData(.NewIndex) = ReportList
    
    ReportList = GeneralLedger
    .AddItem GetResourceString(93) '"General Ledger"
    .ItemData(.newIndex) = ReportList
    
    ReportList = AccountGeneralLedger
    .AddItem GetResourceString(36) & " " & _
                    GetResourceString(93) '"Account General Ledger"
    .ItemData(.newIndex) = ReportList
    
    ReportList = ProfitandLossTrans
    '.AddItem GetResourceString(443) & " " & _
            GetResourceString(28) .AddItem "Profit and Loss Transactions"
    '.ItemData(.NewIndex) = ReportList
    
    ReportList = TotalTransActionsMade
    '.AddItem GetResourceString(52) & " " & _
            GetResourceString(28)  '."Total TransActionsMade"
    '.ItemData(.NewIndex) = ReportList
    
    ReportList = AccountLedger
    If gLangOffSet = wis_NoLangOffset Then
        .AddItem "Ledger of Account"
    Else
        .AddItem GetResourceString(36, 295)
    End If
    .ItemData(.newIndex) = ReportList
    
    ReportList = AccountLedgerOnDate
    If gLangOffSet Then
       .AddItem GetResourceString(36, 295) & " " & _
            GetFromDateString(GetResourceString(37))
    Else
        .AddItem "Ledger of Account As On Date"
    End If
    .ItemData(.newIndex) = ReportList
    
    ReportList = DayBook
    .AddItem GetResourceString(63) '"Day Book"
    .ItemData(.newIndex) = ReportList
    
    ReportList = SubDayBook
    .AddItem GetResourceString(390) & " " & _
                GetResourceString(63) '.AddItem "Sub Day Book"
    .ItemData(.newIndex) = ReportList
    
End With

End Sub


Private Sub Form_Resize()
On Error Resume Next

cmbReportList.SetFocus

End Sub


