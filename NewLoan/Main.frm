VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Managment Information Systems"
   ClientHeight    =   6120
   ClientLeft      =   1095
   ClientTop       =   1860
   ClientWidth     =   8820
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8820
   Begin ComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   552
      Top             =   360
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   6150
      Left            =   450
      Picture         =   "Main.frx":000C
      Top             =   480
      Width           =   7860
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   72
      Top             =   336
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483637
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   32896
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9D5F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9D910
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9DC2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9DF44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9E25E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9E578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9E892
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9EBAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9EEC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9F1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9F4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9F814
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9FB2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":9FE48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuGroup 
         Caption         =   "&Groups"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuNewReport 
         Caption         =   "New Report"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transactions"
      Begin VB.Menu mnuBankTrans 
         Caption         =   "Bank Transaction"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuLoanAccount 
         Caption         =   "Loan Account"
      End
      Begin VB.Menu mnuLoanTrans 
         Caption         =   "Loan Transaction"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDataShed4 
         Caption         =   "Data of Shedule 4"
      End
      Begin VB.Menu mnuConsolOB 
         Caption         =   "Consolidated Opening Balance"
      End
      Begin VB.Menu mnuAccDetail 
         Caption         =   "Account Details"
      End
      Begin VB.Menu mnuOb 
         Caption         =   "Opening Balance"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuRP 
         Caption         =   "Reciept && Payments"
      End
      Begin VB.Menu mnuPl 
         Caption         =   "Profit && Loss"
      End
      Begin VB.Menu mnuBalances 
         Caption         =   "Balance Sheet"
      End
      Begin VB.Menu mnuKStatement 
         Caption         =   "K Statement"
      End
      Begin VB.Menu mnuShedules 
         Caption         =   "Schedules"
         Begin VB.Menu mnushedule1 
            Caption         =   "Schedule 1"
         End
         Begin VB.Menu mnushedule2 
            Caption         =   "Schedule 2"
         End
         Begin VB.Menu mnushedule4 
            Caption         =   "Schedule 4"
            Begin VB.Menu mnushedule4A 
               Caption         =   "Part A"
            End
            Begin VB.Menu mnushedule4B 
               Caption         =   "Part B"
            End
            Begin VB.Menu mnushedule4C 
               Caption         =   "Part C"
            End
         End
         Begin VB.Menu mnushedule5 
            Caption         =   "Schedule 5"
         End
         Begin VB.Menu mnuSchedule6 
            Caption         =   "Schedule 6"
         End
      End
      Begin VB.Menu mnuConShed 
         Caption         =   "Consolidate Schedules"
         WindowList      =   -1  'True
         Begin VB.Menu mnuConShedule1 
            Caption         =   "Schedule 1"
         End
         Begin VB.Menu mnuConShedule2 
            Caption         =   "Schedule 2"
         End
         Begin VB.Menu mnuConShedule5 
            Caption         =   "Schedule 5"
         End
      End
      Begin VB.Menu mnuCorrTrans 
         Caption         =   "Correct Transaction"
      End
      Begin VB.Menu mnuRegReports 
         Caption         =   "Regular Reports"
         Begin VB.Menu mnuWeekReport 
            Caption         =   "Weekly Reports"
            Begin VB.Menu mnuW1 
               Caption         =   "W1"
            End
            Begin VB.Menu mnuW2 
               Caption         =   "W2"
            End
         End
         Begin VB.Menu mnuMonthReport 
            Caption         =   "Monthly Reports"
            Begin VB.Menu mnuM1 
               Caption         =   "M1"
            End
            Begin VB.Menu mnuM2 
               Caption         =   "M2"
            End
         End
         Begin VB.Menu mnuQuarterReport 
            Caption         =   "Quarterly Reports"
            Begin VB.Menu mnuQ1 
               Caption         =   "Q1"
            End
            Begin VB.Menu mnuQ2 
               Caption         =   "Q2"
            End
            Begin VB.Menu mnuQ4 
               Caption         =   "Q4"
            End
         End
         Begin VB.Menu mnuHalfYearly 
            Caption         =   "Half Yearly Report"
            Begin VB.Menu mnuH4 
               Caption         =   "H4"
            End
         End
         Begin VB.Menu mnuYearlyReport 
            Caption         =   "Yearly Reports"
            Begin VB.Menu mnuY10D 
               Caption         =   "Y10D"
            End
            Begin VB.Menu mnuY10F 
               Caption         =   "Y10F"
            End
            Begin VB.Menu mnuY10H 
               Caption         =   "Y10H"
            End
            Begin VB.Menu mnuY4 
               Caption         =   "Y4"
            End
            Begin VB.Menu mnuY5 
               Caption         =   "Y5"
            End
         End
      End
   End
   Begin VB.Menu mnuBank 
      Caption         =   "&Banks"
      Begin VB.Menu mnuBankLists 
         Caption         =   "&List Banks.."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuNewBank 
         Caption         =   "&New Bank"
      End
      Begin VB.Menu mnuSociety 
         Caption         =   "New Society"
      End
      Begin VB.Menu mnuBankQuery 
         Caption         =   "&Query..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuPatient 
      Caption         =   "&Heads"
      Visible         =   0   'False
      Begin VB.Menu mnuListHeads 
         Caption         =   "&List Heads"
      End
      Begin VB.Menu mnuNewHead 
         Caption         =   "&NewHead"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuLoanHead 
         Caption         =   "New Loan Head"
      End
      Begin VB.Menu mnuAddCaste 
         Caption         =   "Add Caste"
      End
      Begin VB.Menu mnuAddCrop 
         Caption         =   "Add Crops"
      End
   End
   Begin VB.Menu mnuRemind 
      Caption         =   "&Remind"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Help Topics..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents M_LookUp As frmLookUp
Attribute M_LookUp.VB_VarHelpID = -1
Private WithEvents M_DateForm  As frmRptDt
Attribute M_DateForm.VB_VarHelpID = -1
Private WithEvents M_frmMonth  As frmReptMonth
Attribute M_frmMonth.VB_VarHelpID = -1
Private WithEvents M_BankSelect As frmBankReport
Attribute M_BankSelect.VB_VarHelpID = -1
Private WithEvents M_frmOptions  As frmOption
Attribute M_frmOptions.VB_VarHelpID = -1

Private m_HeadId As Long
'Private m_AccType As wis_AccountType

Private m_varEventReturned As Variant
Private m_BankId As Long
Private M_BankCode As String
Private M_SelId As Long
Private m_FromDate As String
Private m_ToDate As String
Private m_SelectedBank As Integer

Sub AddToolBar()

        Dim btnX As Button
        tbMain.ImageList = ImageList1
        Set btnX = tbMain.Buttons.Add(1, , , , 13)
                btnX.ToolTipText = "Transactions"
        Set btnX = tbMain.Buttons.Add(2, , , , 14)
                btnX.ToolTipText = "New Bank"
        Set btnX = tbMain.Buttons.Add(3, , , , 2)
                btnX.ToolTipText = "List All Bank"
        Set btnX = tbMain.Buttons.Add(4, , , , 7)
                btnX.ToolTipText = "New Head"
        Set btnX = tbMain.Buttons.Add(5, , , , 11)
                btnX.ToolTipText = "List Heads"
        Set btnX = tbMain.Buttons.Add(6, , , , 9)
                btnX.ToolTipText = "Reminder"
        Set btnX = tbMain.Buttons.Add(7, , , , 12)
                btnX.ToolTipText = "Help"
        Set btnX = tbMain.Buttons.Add(8, , , , 10)
                btnX.ToolTipText = "About"
End Sub

Private Function GetBankIDFromBankList(Optional BankType As Integer) As Long

    GetBankIDFromBankList = 0
    
Dim Rst As Recordset
Dim SqlStr As String
Dim BankId As String
If BankType = 0 Then
    SqlStr = "SELECT BankCode, BankName From BankDet " 'Order By BankCode"
Else
    SqlStr = "SELECT BankCode, BankName From BankDet " 'Order By BankCode"
    Dim locBankType As wis_BankType
    locBankType = HeadOffice
    If locBankType And BankType Then
        SqlStr = SqlStr & " WHERE BranchType = " & locBankType
    End If
    locBankType = Divisionaloffice
    If locBankType And BankType Then
        SqlStr = SqlStr & IIf(InStr(1, SqlStr, "WHERE", vbTextCompare), " OR ", " WHERE ")
        SqlStr = SqlStr & " BranchType = " & locBankType
    End If
    locBankType = TalukaBranch
    If locBankType And BankType Then
        SqlStr = SqlStr & IIf(InStr(1, SqlStr, "WHERE", vbTextCompare), " OR ", " WHERE ")
        SqlStr = SqlStr & " BranchType = " & locBankType
    End If
    locBankType = Branch
    If locBankType And BankType Then
        SqlStr = SqlStr & IIf(InStr(1, SqlStr, "WHERE", vbTextCompare), " OR ", " WHERE ")
        SqlStr = SqlStr & " BranchType = " & locBankType
    End If
    locBankType = Society
    If locBankType And BankType Then
        SqlStr = SqlStr & IIf(InStr(1, SqlStr, "WHERE", vbTextCompare), " OR ", " WHERE ")
        SqlStr = SqlStr & " BranchType = " & locBankType
    End If
    
End If


gDbTrans.SQLStmt = SqlStr & " Order By BankCode"
If gDbTrans.SQLFetch Then
    Set Rst = gDbTrans.Rst.Clone
    Set M_LookUp = New frmLookUp
    If FillView(M_LookUp.lvwReport, Rst, True) Then
        M_LookUp.Show vbModal
    End If
Else
    Exit Function
End If

    SqlStr = "SELECT BankId From BankDet " & _
       " Where BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
gDbTrans.SQLStmt = SqlStr

If gDbTrans.SQLFetch Then
    GetBankIDFromBankList = FormatField(gDbTrans.Rst(0))
End If
   m_BankId = GetBankIDFromBankList
End Function

Private Sub Form_Activate()

    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'arrange the Form
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Left = 0: Me.Top = 0
'Arrange the Logo
    imgLogo.Left = (Width - imgLogo.Width) / 2
    imgLogo.Top = (Height - imgLogo.Height) / 2
            Dim Piece As Integer
            Me.Caption = gAppName & "-" & "Main "
          
            Call AddToolBar
            Piece = Me.Width / 6
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gDbTrans.CloseDB
    End
End Sub

Private Sub M_BankSelect_OkClicked(intSelection As Integer)
    m_varEventReturned = intSelection
End Sub

Private Sub M_DateForm_CancelClick()
m_FromDate = ""
 m_ToDate = ""
End Sub

Private Sub M_DateForm_OKClick(stdate As String, enddate As String)
 m_FromDate = stdate
 m_ToDate = enddate
End Sub

Private Sub M_frmMonth_OKClicked(strIndainDate As String)
    m_ToDate = strIndainDate
End Sub


Private Sub M_frmOptions_CancelCicked(intAccType As Integer)
    m_varEventReturned = intAccType
End Sub

Private Sub M_frmOptions_OkCicked(intAccType As Integer)
    m_varEventReturned = intAccType
End Sub

Private Sub M_LookUp_PrintClick(strSelection As String)
    m_varEventReturned = strSelection
End Sub

Private Sub M_LookUp_SaveClick(strSelection As String)
    m_varEventReturned = strSelection
End Sub

Private Sub M_LookUp_SelectClick(strSelection As String)
    m_varEventReturned = strSelection
End Sub




Private Sub M_MonthForm_CancelClicked()
m_varEventReturned = ""
End Sub


Private Sub M_MonthForm_OKClicked(intMonth As Integer, IntYear As Integer)
    m_FromDate = intMonth & "/" & "1/" & IntYear
    m_ToDate = DateAdd("M", 1, m_FromDate)
    
    'Convert the date into Indian format
    m_ToDate = FormatDate(DateAdd("d", -1, m_ToDate))
    m_FromDate = FormatDate(m_FromDate)
    
End Sub
Private Sub mnuAbout_Click()
'    frmAbout.Show vbModal
End Sub

Private Sub mnuAccDetail_Click()
    Dim BankId As Long
    Dim BankType As Integer
    m_varEventReturned = ""
    BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
    BankId = GetBankIDFromBankList(BankType)
    If BankId = 0 Then Exit Sub
'    frmAccTrans.p_BankId = BankId
 '   frmAccTrans.Show 1
    
End Sub

Private Sub mnuAddCaste_Click()
frmCaste.Show vbModal
End Sub

Private Sub mnuAddCrop_Click()
'frmCrops.Show 1
End Sub

Private Sub mnuBalances_Click()
    Dim BankId As Long
    Dim BankType As Integer
    m_varEventReturned = ""
    
    If M_BankSelect Is Nothing Then Set M_BankSelect = New frmBankReport
    M_BankSelect.Show vbModal
    BankId = 0
    If m_varEventReturned = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf m_varEventReturned = 0 Then
        Exit Sub
    End If
    
    Dim intMonth As Integer
    Dim FromDate As String
    
    FromDate = Format(Now, "dd/mm/yyyy")
    intMonth = Month(Now)
    If intMonth > 3 Then
        FromDate = "1/4/" & CStr(Year(Now))
    Else
        FromDate = "1/4/" & CStr(Year(Now) - 1)
    End If
    
    If M_DateForm Is Nothing Then Set M_DateForm = New frmRptDt
    M_DateForm.txtStDate = FromDate 'Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_ToDate = "" Then Exit Sub
    If m_FromDate = "" Then Exit Sub
'    frmReport.p_BankId = BankId
'
'    If frmReport.ShowBalanceSheet(frmReport.grd, m_FromDate, m_ToDate) Then
'      frmReport.Show vbModal
'    Else
'        MsgBox "No records", vbInformation, wis_MESSAGE_TITLE
'    End If
    
End Sub

Private Sub mnuBankLists_Click()
     If M_frmOptions Is Nothing Then Set M_frmOptions = New frmOption
     Load M_frmOptions
     
     Dim BankType As wis_BankType
     
     BankType = Divisionaloffice
     M_frmOptions.chkOption(0).Caption = "Divisional Branches"
     M_frmOptions.chkOption(0).Tag = BankType
     BankType = TalukaBranch
     M_frmOptions.chkOption(1).Caption = "Taluka Branches"
     M_frmOptions.chkOption(1).Tag = BankType
     BankType = Branch
     M_frmOptions.chkOption(2).Caption = "Other Branches"
     M_frmOptions.chkOption(2).Tag = BankType
     BankType = Society
     
     M_frmOptions.chkOption(3).Caption = "Societies"
     M_frmOptions.chkOption(3).Tag = BankType
     m_varEventReturned = ""
'     M_frmOptions.Show vbModal
'     If Val(m_varEventReturned) = 0 Then Exit Sub
'     If frmReport.ShowBanks(frmReport.grd, Val(m_varEventReturned)) Then
'        frmReport.Show vbModal
'     End If
     
End Sub

Private Sub mnuBankTrans_Click()
    Dim BankId As Long
    Dim BankType As Integer
    BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
    BankId = GetBankIDFromBankList(BankType)
    If BankId = 0 Then Exit Sub
'    frmTreeView.p_BankId = BankId 'FormatField(gDbTrans.Rst(0))
'    Load frmTreeView
'
'    frmTreeView.Show vbModal
'    frmTrans.p_BankId = FormatField(gDbTrans.Rst(0))
'    Load frmTrans
'    frmTrans.Show vbModal
   
End Sub

Private Sub mnuConShedule1_Click()
    If M_DateForm Is Nothing Then
        Set M_DateForm = New frmRptDt
    End If
    
    M_DateForm.txtStDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_ToDate = "" Then Exit Sub
    If m_FromDate = "" Then Exit Sub
    
'    If frmGrids.ShowConsoleShed1(m_FromDate, m_ToDate) Then
'        frmGrids.Show vbModal
'    End If

End Sub

Private Sub mnuConShedule2_Click()
    If M_DateForm Is Nothing Then
        Set M_DateForm = New frmRptDt
    End If
    
    M_DateForm.txtStDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_ToDate = "" Then Exit Sub
    If m_FromDate = "" Then Exit Sub
    
'    If frmGrids.ShowConsoleShed2(m_FromDate, m_ToDate) Then
'        frmGrids.Show vbModal
'    End If


End Sub


Private Sub mnuCorrTrans_Click()
Dim Rst As Recordset
Dim SqlStr As String
Dim BankId As Long
Dim BankType As Integer
    
    m_varEventReturned = ""
    
    If M_BankSelect Is Nothing Then Set M_BankSelect = New frmBankReport
    M_BankSelect.Show vbModal
    BankId = 0
    If m_varEventReturned = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf m_varEventReturned = 0 Then
        Exit Sub
    End If
        
'   frmReportCorrTrans.p_BankId = BankId 'FormatField(gDbTrans.Rst(0))
'   Load frmReportCorrTrans
'   frmReportCorrTrans.Show vbModal

End Sub


Private Sub mnuDataShed4_Click()


    Dim BankId As Long
    Dim BankType As Integer
    Set M_BankSelect = New frmBankReport
    M_BankSelect.optConsol.Caption = "All Bank"
    M_BankSelect.optIndividual.value = True
    m_varEventReturned = ""
    M_BankSelect.Show vbModal
    Set M_BankSelect = Nothing
    If Val(m_varEventReturned) = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf Val(m_varEventReturned) = 0 Then
      Exit Sub
    End If
    
'    frmShed4.p_BankId = BankId
'    frmShed4.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGroup_Click()
'    frmGroups.Show vbModal
End Sub


Private Sub mnuH4_Click()
'If frmGrids.H4("", "") Then
'    frmGrids.Show vbModal
'End If

End Sub

Private Sub mnuHelpTopics_Click()
'                cdb.HelpCommand = cdlHelpContents
'                cdb.HelpCommand = cdlHelpIndex
'                cdb.ShowHelp
End Sub

Private Sub mnuKStatement_Click()
  Dim BankId As Long
    Dim BankType As Integer
    m_varEventReturned = ""
    
    BankId = 0
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
      
    Dim intMonth As Integer
    Dim FromDate As String
    If Val(m_varEventReturned) = 0 Then Exit Sub
    FromDate = Format(Now, "dd/mm/yyyy")
    intMonth = Month(Now)
    If intMonth > 3 Then
        FromDate = "1/4/" & CStr(Year(Now))
    Else
        FromDate = "1/4/" & CStr(Year(Now) - 1)
    End If
    
    If M_DateForm Is Nothing Then Set M_DateForm = New frmRptDt
    M_DateForm.txtStDate = FromDate 'Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_FromDate = "" Or m_ToDate = "" Then Exit Sub
    
'    frmReport.p_BankId = m_BankId
'    If frmReport.ShowK_Statement(frmReport.grd, m_FromDate, m_ToDate) Then
'        frmReport.Show vbModal
'    Else
'        MsgBox "No records", vbInformation, wis_MESSAGE_TITLE
'    End If
End Sub


Private Sub mnuListHeads_Click()
'
'm_varEventReturned = ""
'
'    If M_frmOptions Is Nothing Then Set M_frmOptions = New frmOption
'
'     Load M_frmOptions
'
'    ' Dim BankType As wis_AccountType
'
'     BankType = Asset
'     M_frmOptions.chkOption(0).Caption = "Asset"
'     M_frmOptions.chkOption(0).Tag = BankType
'
'     BankType = Liability
'     M_frmOptions.chkOption(1).Caption = "Laibility"
'     M_frmOptions.chkOption(1).Tag = BankType
'
'     BankType = Loss
'     M_frmOptions.chkOption(2).Caption = "Loss"
'     M_frmOptions.chkOption(2).Tag = BankType
'
'     BankType = Profit
'     M_frmOptions.chkOption(3).Caption = "Profit"
'     M_frmOptions.chkOption(3).Tag = BankType
'
'      m_varEventReturned = ""
'     M_frmOptions.Show vbModal
'
'    If Val(m_varEventReturned) = 0 Then Exit Sub
'
'    If frmReport.ShowHeads(frmReport.grd, Val(m_varEventReturned)) Then
'        frmReport.Show
'    End If

End Sub

Private Sub mnuLoanAccount_Click()
    Dim BankId As Long
    Dim BankType As Integer
    BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
    BankId = GetBankIDFromBankList(BankType)
    If BankId = 0 Then Exit Sub
'    frmLoanInit.p_BankId = BankId
'    Load frmLoanInit
'    frmLoanInit.Show vbModal
    
End Sub

Private Sub mnuLoanHead_Click()
'    frmLoanHeads.Show 1
End Sub

Private Sub mnuLoanTrans_Click()
    Dim BankId As Long
    Dim BankType As Integer
    Set M_BankSelect = New frmBankReport
    M_BankSelect.optConsol.Caption = "All Bank"
    M_BankSelect.optIndividual.value = True
    m_varEventReturned = ""
    M_BankSelect.Show vbModal
    Set M_BankSelect = Nothing
    If Val(m_varEventReturned) = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf Val(m_varEventReturned) = 0 Then
      Exit Sub
    End If
'    frmLoanTrans.p_BankId = BankId
'    Load frmLoanTrans
'    frmLoanTrans.Show 1
    
End Sub


Private Sub mnuNewBank_Click()
     Load frmBank
     frmBank.optSociety.Enabled = False
     frmBank.Show vbModal
End Sub

Private Sub mnuNewHead_Click()
'    frmSchemes.Show
End Sub

Private Sub mnuNewReport_Click()
'    frmReportDesign.Show vbModal
    
End Sub

Private Sub mnuPl_Click()
    Dim BankId As Long
    Dim BankType As Integer
    m_varEventReturned = ""
    
    If M_BankSelect Is Nothing Then Set M_BankSelect = New frmBankReport
    M_BankSelect.Show vbModal
    BankId = 0
    If m_varEventReturned = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf m_varEventReturned = 0 Then
        Exit Sub
    End If
    
    Dim intMonth As Integer
    Dim FromDate As String
    
    FromDate = Format(Now, "dd/mm/yyyy")
    intMonth = Month(Now)
    If intMonth > 3 Then
        FromDate = "1/4/" & CStr(Year(Now))
    Else
        FromDate = "1/4/" & CStr(Year(Now) - 1)
    End If
    
    If M_DateForm Is Nothing Then Set M_DateForm = New frmRptDt
    M_DateForm.txtStDate = FromDate 'Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_FromDate = "" Or m_ToDate = "" Then Exit Sub
'    frmReport.p_BankId = BankId
'
'    If frmReport.ShowProfitLoss(frmReport.grd, m_FromDate, m_ToDate) Then
'        frmReport.Show vbModal
'    Else
'        MsgBox "No records", vbInformation, wis_MESSAGE_TITLE
'    End If
End Sub

Private Sub ShowAllOffices()

'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim SqlStr As String
Dim Rst As Recordset

SqlStr = "Select BankCode,BankName,Manager, Address,BankID from BankDet "
gDbTrans.SQLStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set Rst = gDbTrans.Rst.Clone
Else
    GoTo ExitLine
End If
If M_LookUp Is Nothing Then
    Set M_LookUp = New frmLookUp
End If

Call FillView(M_LookUp.lvwReport, Rst, True)

M_BankCode = ""
M_LookUp.Show vbModal

 If Val(m_varEventReturned) <> 0 Then  '"get the BnakID
            Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
            If Not Rst.NoMatch Then
                m_BankId = FormatField(Rst("BankID"))
            Else
                m_BankId = 0
            End If
 End If

ExitLine:
    Exit Sub

ErrLine:
    
    If Err Then
        MsgBox " ShowAllOffices: " & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Err.Clear
        GoTo ExitLine
    End If

End Sub



Private Sub mnuQ4_Click()

MsgBox "This is under construction"

'If frmGrids.Q4("", "") Then
 '   frmGrids.Show vbModal
'End If

End Sub


Private Sub mnuRemind_Click()

'frmRemind.Show vbModal

End Sub



Private Sub mnuRP_Click()
    
    Dim BankId As Long
    Dim BankType As Integer
    m_varEventReturned = ""
    
    If M_BankSelect Is Nothing Then Set M_BankSelect = New frmBankReport
    M_BankSelect.Show vbModal
    BankId = 0
    If m_varEventReturned = 1 Then
      BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
      BankId = GetBankIDFromBankList(BankType)
      If BankId = 0 Then Exit Sub
    ElseIf m_varEventReturned = 0 Then
        Exit Sub
    End If
    
    Dim intMonth As Integer
    Dim FromDate As String
    
    FromDate = Format(Now, "dd/mm/yyyy")
    intMonth = Month(Now)
    If intMonth > 3 Then
        FromDate = "1/4/" & CStr(Year(Now))
    Else
        FromDate = "1/4/" & CStr(Year(Now) - 1)
    End If
    
    If M_DateForm Is Nothing Then Set M_DateForm = New frmRptDt
    M_DateForm.txtStDate = FromDate 'Format(Now, "dd/mm/yyyy")
    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
    M_DateForm.Show vbModal
    Set M_DateForm = Nothing
    If m_FromDate = "" Or m_ToDate = "" Then Exit Sub
    
'    frmReport.p_BankId = BankId
'    frmReport.ReporTtype = ReceiptsAndPayments
'    If frmReport.ShowReciptPayment(frmReport.grd, m_FromDate, m_ToDate) Then
'        frmReport.Show vbModal
'    Else
'        MsgBox "No records", vbInformation, wis_MESSAGE_TITLE
'    End If

End Sub

Private Sub mnuTreeView_Click()
'frmTreeView.Show vbModal
End Sub

Private Sub mnuSchedule6_Click()
Dim SqlStr As String
Dim Rst As Recordset
    
    SqlStr = "SELECT BankCode,BankName,BankID From BankDet " & _
            " WHERE BankID Mod " & BO_Offset & " = 0"
    
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch < 1 Then Exit Sub
    Set Rst = gDbTrans.Rst.Clone
    
    Set M_LookUp = New frmLookUp
    
    Call FillView(M_LookUp.lvwReport, Rst, True)
    m_varEventReturned = ""
    M_LookUp.Show vbModal
    Set M_LookUp = Nothing
       
    If Trim(m_varEventReturned) = "" Then Exit Sub
    
    Rst.MoveFirst
    Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
    
    If Rst.NoMatch Then
        MsgBox "Unable to Deteck Bank", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
'    frmGrids.p_BankId = FormatField(Rst("BankId"))
'    Set Rst = Nothing
'
'    If M_DateForm Is Nothing Then
'        Set M_DateForm = New frmRptDt
'    End If
'
'    M_DateForm.txtStDate = Format(Now, "dd/mm/yyyy")
'    M_DateForm.txtEndDate = Format(Now, "dd/mm/yyyy")
'    M_DateForm.Show vbModal
'    Set M_DateForm = Nothing
'    If m_ToDate = "" Then Exit Sub
'    If m_FromDate = "" Then Exit Sub
'
'    If frmGrids.ShowShed6(m_FromDate, m_ToDate) Then
'        frmGrids.Show vbModal
'    End If
End Sub

Private Sub mnushedule1_Click()

Dim SqlStr As String
Dim Rst As Recordset
    
    SqlStr = "SELECT BankCode,BankName,BankID From BankDet " & _
            " WHERE BankID Mod " & BO_Offset & " = 0"
    
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch < 1 Then Exit Sub
    Set Rst = gDbTrans.Rst.Clone
    
    Set M_LookUp = New frmLookUp
    
    Call FillView(M_LookUp.lvwReport, Rst, True)
    m_varEventReturned = ""
    M_LookUp.Show vbModal
    Set M_LookUp = Nothing
       
    If Trim(m_varEventReturned) = "" Then Exit Sub
    
    Rst.MoveFirst
    Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
    
    If Rst.NoMatch Then
        MsgBox "Unable to Deteck Bank", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
'    frmGrids.p_BankId = FormatField(Rst("BankId"))
'    Set Rst = Nothing
'
'    If M_frmMonth Is Nothing Then
'        Set M_frmMonth = New frmReptMonth
'    End If
'
'    M_frmMonth.Show vbModal
'    Set M_frmMonth = Nothing
'    If m_ToDate = "" Then Exit Sub
'
'    If frmGrids.ShowShed1(m_ToDate) Then
'        frmGrids.Show vbModal
'    End If
End Sub

Private Sub mnushedule2_Click()

Dim SqlStr As String
Dim Rst As Recordset
    
    SqlStr = "SELECT BankCode,BankName,BankID From BankDet " & _
            " WHERE BankID Mod " & BO_Offset & " = 0"
    
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch < 1 Then Exit Sub
    Set Rst = gDbTrans.Rst.Clone
    
    Set M_LookUp = New frmLookUp
    
    Call FillView(M_LookUp.lvwReport, Rst, True)
    m_varEventReturned = ""
    M_LookUp.Show vbModal
    Set M_LookUp = Nothing
       
    If Trim(m_varEventReturned) = "" Then Exit Sub
    
    Rst.MoveFirst
    Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
    
    If Rst.NoMatch Then
        MsgBox "Unable to Deteck Bank", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
'    frmGrids.p_BankId = FormatField(Rst("BankId"))
'    Set Rst = Nothing
'
'    If M_frmMonth Is Nothing Then
'        Set M_frmMonth = New frmReptMonth
'    End If
'
'    M_frmMonth.Show vbModal
'    Set M_frmMonth = Nothing
'    If m_ToDate = "" Then Exit Sub
'
'    If frmGrids.ShowShed2(m_ToDate) Then
'        frmGrids.Show vbModal
'    End If
End Sub


Private Sub mnushedule4B_Click()

Dim SqlStr As String
Dim Rst As Recordset
    
    SqlStr = "SELECT BankCode,BankName,BankID From BankDet " & _
            " WHERE BankID Mod " & BO_Offset & " = 0"
    
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch < 1 Then Exit Sub
    Set Rst = gDbTrans.Rst.Clone
    
    Set M_LookUp = New frmLookUp
    
    Call FillView(M_LookUp.lvwReport, Rst, True)
    m_varEventReturned = ""
    M_LookUp.Show vbModal
    Set M_LookUp = Nothing
       
    If Trim(m_varEventReturned) = "" Then Exit Sub
    
    Rst.MoveFirst
    Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
    
    If Rst.NoMatch Then
        MsgBox "Unable to Deteck Bank", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
'    frmGrids.p_BankId = FormatField(Rst("BankId"))
'    Set Rst = Nothing
'
'    If M_frmMonth Is Nothing Then
'        Set M_frmMonth = New frmReptMonth
'    End If
'
'    M_frmMonth.Show vbModal
'    Set M_frmMonth = Nothing
'    If m_ToDate = "" Then Exit Sub
'
'    If frmGrids.ShowShed4B("1/4/2001", m_ToDate) Then
'        frmGrids.Show vbModal
'    End If
End Sub

Private Sub mnushedule4A_Click()

Dim SqlStr As String
Dim Rst As Recordset
    
    SqlStr = "SELECT BankCode,BankName,BankID From BankDet " & _
            " WHERE BankID Mod " & BO_Offset & " = 0"
    
    gDbTrans.SQLStmt = SqlStr
    If gDbTrans.SQLFetch < 1 Then Exit Sub
    Set Rst = gDbTrans.Rst.Clone
    
    Set M_LookUp = New frmLookUp
    
    Call FillView(M_LookUp.lvwReport, Rst, True)
    m_varEventReturned = ""
    M_LookUp.Show vbModal
    Set M_LookUp = Nothing
       
    If Trim(m_varEventReturned) = "" Then Exit Sub
    
    Rst.MoveFirst
    Rst.FindFirst "BankCode = " & AddQuotes(CStr(m_varEventReturned), True)
    
    If Rst.NoMatch Then
        MsgBox "Unable to Deteck Bank", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
'    frmGrids.p_BankId = FormatField(Rst("BankId"))
'    Set Rst = Nothing
'
'    If M_frmMonth Is Nothing Then
'        Set M_frmMonth = New frmReptMonth
'    End If
'
'    M_frmMonth.Show vbModal
'    Set M_frmMonth = Nothing
'    If m_ToDate = "" Then Exit Sub
'
'    If frmGrids.ShowShed4A("1/4/2001", m_ToDate) Then
'        frmGrids.Show vbModal
'    End If
'

End Sub


Private Sub mnuSociety_Click()
    Dim BankId As Long
    Dim BankType As Integer
    BankType = Branch + Divisionaloffice + HeadOffice + TalukaBranch
    BankId = GetBankIDFromBankList(BankType)
    If BankId = 0 Then Exit Sub
    Dim Count As Integer
    Load frmBank
    With frmBank.cmbParent
        For Count = 0 To .ListCount
            If .ItemData(Count) = BankId Then
                .ListIndex = Count
                .Locked = True
                Exit For
            End If
        Next
    End With
    frmBank.optBranch.Enabled = False
    frmBank.optDivisional.Enabled = False
    frmBank.optTaluka.Enabled = False
    frmBank.optSociety.value = True
    frmBank.chkLiq.Enabled = True
    frmBank.Show vbModal
    
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As ComctlLib.Button)

    If Button.Index = 1 Then
        Call mnuBankTrans_Click
    ElseIf Button.Index = 2 Then
        Call mnuNewBank_Click
    ElseIf Button.Index = 3 Then
        Call mnuBankLists_Click
    ElseIf Button.Index = 4 Then
        Call mnuNewHead_Click
    ElseIf Button.Index = 5 Then
            Call mnuListHeads_Click
    ElseIf Button.Index = 6 Then
        Call mnuRemind_Click
    ElseIf Button.Index = 7 Then
    
    ElseIf Button.Index = 8 Then
        Call mnuAbout_Click
    End If
End Sub

