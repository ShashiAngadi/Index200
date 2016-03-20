VERSION 5.00
Begin VB.Form frmLoanScheme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New loan scheme"
   ClientHeight    =   6855
   ClientLeft      =   2145
   ClientTop       =   1350
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6795
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   90
      TabIndex        =   21
      Top             =   1680
      Width           =   6525
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   2160
         TabIndex        =   26
         Top             =   1140
         Width           =   4185
      End
      Begin VB.ComboBox cmbTerm 
         Height          =   315
         Left            =   2160
         TabIndex        =   25
         Top             =   1590
         Width           =   4185
      End
      Begin VB.ComboBox cmbLoanType 
         Height          =   315
         Left            =   2160
         TabIndex        =   24
         Top             =   2040
         Width           =   4185
      End
      Begin VB.ComboBox cmbPurpose 
         Height          =   315
         Left            =   2160
         TabIndex        =   23
         Top             =   690
         Width           =   4155
      End
      Begin VB.CheckBox chkMemOnly 
         Caption         =   "       Loan only for Members"
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   210
         Value           =   1  'Checked
         Width           =   5685
      End
      Begin VB.Label lblCategary 
         Caption         =   "Loan &Categary"
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label lblTerm 
         Caption         =   "&Term :"
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   1620
         Width           =   1905
      End
      Begin VB.Label lblClassification 
         Caption         =   "Loan T&ype classifiaction"
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Top             =   2070
         Width           =   2085
      End
      Begin VB.Label lblPurpose 
         Caption         =   "&Purpose"
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1995
      End
   End
   Begin VB.Frame fraName 
      Height          =   1575
      Left            =   90
      TabIndex        =   12
      Top             =   120
      Width           =   6525
      Begin VB.TextBox txtLoanNameEnglish 
         Height          =   345
         Left            =   2280
         TabIndex        =   18
         Top             =   1080
         Width           =   4185
      End
      Begin VB.TextBox txtLoanName 
         Height          =   345
         Left            =   2040
         TabIndex        =   16
         Top             =   630
         Width           =   3945
      End
      Begin VB.CommandButton cmdLoanName 
         Caption         =   "..."
         Height          =   315
         Left            =   6090
         TabIndex        =   15
         Top             =   630
         Width           =   315
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3810
         TabIndex        =   13
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblLoanNameEnglish 
         Caption         =   "Loan &Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label lblLoanName 
         Caption         =   "Loan &Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label lblDate 
         Caption         =   "&Date :"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   1800
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2025
      Left            =   90
      TabIndex        =   3
      Top             =   4170
      Width           =   6525
      Begin VB.TextBox txtEmpPenalInt 
         Height          =   315
         Left            =   5490
         TabIndex        =   32
         Top             =   1410
         Width           =   825
      End
      Begin VB.TextBox txtEmpIntrate 
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Width           =   825
      End
      Begin VB.TextBox txtPenalInt 
         Height          =   315
         Left            =   5490
         TabIndex        =   7
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtMonthDuration 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox txtIntrate 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtDayDuration 
         Height          =   315
         Left            =   5490
         TabIndex        =   4
         Top             =   690
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   6360
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblEmployee 
         Caption         =   "Interest rate for Emplyees"
         Height          =   300
         Left            =   60
         TabIndex        =   35
         Top             =   1110
         Width           =   2205
      End
      Begin VB.Label lblEmpPenalInt 
         Caption         =   "P&enal interest :"
         Height          =   300
         Left            =   3390
         TabIndex        =   34
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label lblEmpIntrate 
         Caption         =   "Rate of &interest :"
         Height          =   300
         Left            =   120
         TabIndex        =   33
         Top             =   1470
         Width           =   1425
      End
      Begin VB.Label lblPenalInt 
         Caption         =   "&Penal interest :"
         Height          =   300
         Left            =   3390
         TabIndex        =   11
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblMonthDuaration 
         Caption         =   "Duration (&Month)"
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   1365
      End
      Begin VB.Label lblIntrate 
         Caption         =   "&Rate of interest :"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblDayDuration 
         Caption         =   "Duration (d&ays)"
         Height          =   300
         Left            =   3390
         TabIndex        =   8
         Top             =   720
         Width           =   1965
      End
   End
   Begin VB.ListBox lstHidden 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   1380
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Canc&el"
      Height          =   400
      Left            =   5370
      TabIndex        =   0
      Top             =   6300
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   400
      Left            =   3990
      TabIndex        =   1
      Top             =   6300
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoanScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_dbOperation As wis_DBOperation
Private m_SchemeId As String
Private m_Ctrl As Control

Public Event SchemeCreated(ByVal SchemeID As Integer)
Public Event SchemeModified(ByVal SchemeID As Integer)
Public Event WindowClosed()

Private Function CheckValidation() As Boolean

CheckValidation = False

If Not DateValidate(txtDate, "/", True) Then
    MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If
If txtLoanName.Text = "" Then
    MsgBox "Please Enter LoanName", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtLoanName
    Exit Function
End If
If cmbPurpose.ListIndex = -1 Then
    MsgBox "Please Select Loan Purpose", vbInformation, wis_MESSAGE_TITLE
    cmbPurpose.SetFocus
    Exit Function
End If
If cmbCategory.ListIndex = -1 Then
    MsgBox "Please Select Loan category", vbInformation, wis_MESSAGE_TITLE
    cmbCategory.SetFocus
    Exit Function
End If
If cmbLoanType.ListIndex = -1 Then
    MsgBox "Please Select Loan Type", vbInformation, wis_MESSAGE_TITLE
    cmbLoanType.SetFocus
    Exit Function
End If

If txtIntrate = "" Then
    MsgBox "Please Enter Rate of Interest", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntrate
    Exit Function
End If
If Not IsNumeric(txtIntrate) Then
    MsgBox "Invalid Interest rate Specified", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtIntrate
    Exit Function
End If
If txtPenalInt <> "" Then
    If Not IsNumeric(txtPenalInt) Then
        MsgBox "Invalid penal Interest rate Specified", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtPenalInt
        Exit Function
    End If
End If
If txtMonthDuration <> "" Then
    If Not IsNumeric(txtMonthDuration) Then
        MsgBox "Invalid loan duration Specified", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtMonthDuration
        Exit Function
    End If
End If
If txtDayDuration <> "" Then
    If Not IsNumeric(txtDayDuration) Then
        MsgBox "Invalid loan duration Specified", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtDayDuration
        Exit Function
    End If
End If

'If txtEmpIntrate = "" Then
'    MsgBox "Please Enter Rate of Interest", vbInformation, wis_MESSAGE_TITLE
'    ActivateTextBox txtEmpIntrate
'    Exit Function
'End If
'If Not IsNumeric(txtEmpIntrate) Then
'    MsgBox "Invalid Interest rate Specified", vbInformation, wis_MESSAGE_TITLE
'    ActivateTextBox txtEmpIntrate
'    Exit Function
'End If
If txtEmpPenalInt <> "" Then
    If Not IsNumeric(txtEmpPenalInt) Then
        MsgBox "Invalid penal Interest rate Specified", vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEmpPenalInt
        Exit Function
    End If
End If

CheckValidation = True

End Function
Private Sub ClearControls()
    
    txtLoanName = ""
    txtLoanNameEnglish = ""
    cmbPurpose.Text = ""
    cmbCategory.ListIndex = -1
    cmbTerm.ListIndex = -1
    txtIntrate = ""
    txtPenalInt = ""
    txtMonthDuration = ""
    txtDayDuration = ""
    txtEmpIntrate = ""
    txtEmpPenalInt = ""
    
    cmbLoanType.ListIndex = -1
    chkMemOnly.Value = vbChecked

    cmdCreate.Caption = GetResourceString(15)
    m_dbOperation = Insert
    
End Sub

Private Function LoadLoanSchemeDetails(SchemeID As Long) As Boolean
LoadLoanSchemeDetails = False
Dim SqlStr As String
Dim rst As Recordset
Dim CCLoan As Boolean
Dim CropLoan As Boolean
Dim ItemCount As Integer
Dim Category As Byte
Dim Term As Byte
Dim ln As Integer
SqlStr = "SELECT * from LoanScheme where SchemeID = " & SchemeID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

Dim LoanType As wis_LoanType

'Now set the values on the form controls
txtLoanName = FormatField(rst("SchemeName"))
txtLoanNameEnglish = FormatField(rst("SchemeNameEnglish"))
cmbPurpose.Text = FormatField(rst("LoanPurpose"))
txtIntrate = FormatField(rst("IntRate"))
txtPenalInt = FormatField(rst("PenalIntRate"))
txtMonthDuration = FormatField(rst("MonthDuration"))
txtDayDuration = FormatField(rst("DayDuration"))
txtEmpIntrate = FormatField(rst("EmpIntRate"))
txtEmpPenalInt = FormatField(rst("EmpPenalIntRate"))
chkMemOnly.Value = IIf(FormatField(rst("OnlyMember")), vbChecked, vbUnchecked)
Category = FormatField(rst("Category"))
LoanType = FormatField(rst("LoanType"))
Term = FormatField(rst("TermType"))

ItemCount = 0
ln = cmbPurpose.ListCount - 1
For ItemCount = 0 To ln
    If cmbPurpose.List(ItemCount) = cmbPurpose.Text Then
        cmbPurpose.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount

ItemCount = 0
ln = cmbCategory.ListCount - 1
For ItemCount = 0 To ln
    If cmbCategory.ItemData(ItemCount) = Category Then
        cmbCategory.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount
ItemCount = 0
ln = 0
ln = cmbTerm.ListCount - 1
For ItemCount = 0 To ln
    If cmbTerm.ItemData(ItemCount) = Category Then
        cmbTerm.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount

ln = cmbLoanType.ListCount - 1
For ItemCount = 0 To ln
    If cmbLoanType.ItemData(ItemCount) = LoanType Then
        cmbLoanType.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount

LoadLoanSchemeDetails = True
End Function

Private Sub LoadLoanSchemes()
Dim rst As Recordset
Dim SqlStr As String

'Get the LoanSchemeName & SchemeID from the database to
'show it into lsthidden Listbox
SqlStr = "SELECT SchemeID,SchemeName from LoanScheme"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub

lstHidden.Clear
While Not rst.EOF
    lstHidden.AddItem FormatField(rst("SchemeName"))
    lstHidden.ItemData(lstHidden.newIndex) = FormatField(rst("SchemeID"))
    rst.MoveNext
Wend
End Sub


'This function Saves the LoanSchemes into the LoanScheme table
'
Private Function SaveLoanSchemeDetails() As Boolean

SaveLoanSchemeDetails = False
Dim LoanSchemeName As String
Dim LoanSchemeNameEnglish As String
Dim LoanSchemeID As Long
Dim LoanPurpose As String
Dim LoanCategory As Byte
Dim LoanTerm As Byte
Dim LoanType As wis_LoanType
Dim SqlStr As String
Dim rst As Recordset

Dim AsOnDate As Date
AsOnDate = GetSysFormatDate(txtDate)

LoanSchemeName = Trim$(txtLoanName)
LoanSchemeNameEnglish = Trim$(txtLoanNameEnglish)
LoanCategory = cmbCategory.ItemData(cmbCategory.ListIndex)

LoanTerm = cmbTerm.ItemData(cmbTerm.ListIndex)
LoanPurpose = Trim$(cmbPurpose.Text)

'Get the LoanSchemeId
SqlStr = "Select Max(SchemeID) from LoanScheme"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) <> 1 Then Exit Function
LoanSchemeID = FormatField(rst(0)) + 1
LoanType = cmbLoanType.ItemData(cmbLoanType.ListIndex)
If gLangOffSet = wis_NoLangOffset Then LoanSchemeNameEnglish = LoanSchemeName
'Now Insert the values
SqlStr = ""
gDbTrans.BeginTrans

If m_dbOperation = Insert Then
    m_SchemeId = LoanSchemeID
    SqlStr = "INSERT INTO LoanScheme (SchemeID,SchemeName,SchemeNameEnglish,Category," & _
        " TermType,LoanType,MonthDuration,DayDuration," & _
        " IntRate,PenalIntRate,LoanPurpose,OnlyMember,UserID )" & _
        " Values ( " & _
        LoanSchemeID & "," & _
        AddQuotes(LoanSchemeName, True) & "," & _
        AddQuotes(LoanSchemeNameEnglish, True) & "," & _
        LoanCategory & "," & _
        LoanTerm & "," & _
        LoanType & "," & _
        Val(txtMonthDuration.Text) & "," & _
        Val(txtDayDuration.Text) & "," & _
        Val(Trim$(txtIntrate.Text)) & "," & _
        Val(Trim$(txtPenalInt.Text)) & "," & _
        AddQuotes(LoanPurpose, True) & ", " & _
        IIf(chkMemOnly = vbChecked, True, False) & "," & gUserID & " )"
    
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    RaiseEvent SchemeCreated(LoanSchemeID)
    
ElseIf m_dbOperation = Update Then
    SqlStr = "UPDATE LoanScheme set " & _
        " SchemeName = " & AddQuotes(LoanSchemeName, True) & "," & _
        " SchemeNameEnglish = " & AddQuotes(LoanSchemeNameEnglish, True) & "," & _
        " Category = " & LoanCategory & "," & _
        " TermType = " & LoanTerm & "," & _
        " LoanType = " & LoanType & "," & _
        " MonthDuration = " & Val(txtMonthDuration.Text) & "," & _
        " DayDuration = " & Val(txtDayDuration.Text) & "," & _
        " IntRate = " & Val(Trim$(txtIntrate.Text)) & "," & _
        " PenalIntRate = " & Val(Trim$(txtPenalInt.Text)) & "," & _
        " EmpIntRate = " & Val(Trim$(txtEmpIntrate.Text)) & "," & _
        " EmpPenalIntRate = " & Val(Trim$(txtEmpPenalInt.Text)) & "," & _
        " LoanPurpose = " & AddQuotes(LoanPurpose, True) & "," & _
        " Onlymember = " & IIf(chkMemOnly = vbChecked, True, False) & _
        " Where SchemeID = " & m_SchemeId

    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    RaiseEvent SchemeModified(m_SchemeId)
End If
gDbTrans.CommitTrans

'Now store the Loan interest

Dim ClsInt As New clsInterest
Call ClsInt.SaveInterest(wis_Loans, "MemRegularInterest" & m_SchemeId, Val(txtIntrate), AsOnDate)
Call ClsInt.SaveInterest(wis_Loans, "MemPenalInterest" & m_SchemeId, Val(txtPenalInt), AsOnDate)
Call ClsInt.SaveInterest(wis_Loans, "EmpRegularInterest" & m_SchemeId, Val(txtEmpIntrate), AsOnDate)
Call ClsInt.SaveInterest(wis_Loans, "EmpPenalInterest" & m_SchemeId, Val(txtEmpPenalInt), AsOnDate)

Set ClsInt = Nothing

SaveLoanSchemeDetails = True

End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblDate = GetResourceString(37) 'Date
lblLoanName = GetResourceString(214) 'Loan Scheme
lblLoanNameEnglish = GetResourceString(214, 468) 'Loan Scheme
lblCategary = GetResourceString(317) 'Categaey
lblTerm = GetResourceString(319)      'Term
lblClassification = GetResourceString(58, 318) 'Loan Classification
lblIntrate = GetResourceString(186) 'Rate Of Interest
lblPenalInt = GetResourceString(345, 305) 'Penal Int rate
lblEmpIntrate = GetResourceString(186) 'Rate Of Int
lblEmpPenalInt = GetResourceString(345, 305) 'Penal Int rate
chkMemOnly.Caption = GetResourceString(434)
Me.lblDayDuration.Caption = GetResourceString(433) & " (" & GetResourceString(44) & ")"
Me.lblEmployee.Caption = GetResourceString(155)
Me.lblMonthDuaration = GetResourceString(433) & " (" & GetResourceString(192) & ")"
Me.lblPurpose = GetResourceString(80, 221)

cmdCreate.Caption = GetResourceString(15)  'Create
cmdCancel.Caption = GetResourceString(11) 'Close


Call SkipFontToControls(lblLoanName, lblLoanNameEnglish, txtLoanNameEnglish)
If gLangOffSet = wis_NoLangOffset Then
    lblLoanNameEnglish.Visible = False
    txtLoanNameEnglish.Visible = False
    Dim Ht As Integer
    Ht = txtLoanNameEnglish.Top - txtLoanName.Top
    Me.Height = Me.Height - Ht
    Call ReduceControlsTopPosition(Ht, Frame2, Frame4, cmdCancel, cmdCreate)
End If
End Sub

Private Sub cmdCancel_Click()

Unload Me
End Sub


Private Sub cmdCreate_Click()
'Check the validation
If Not CheckValidation Then Exit Sub

If Not SaveLoanSchemeDetails Then Exit Sub

'MsgBox "Saved the Loan Scheme Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(748), vbInformation, wis_MESSAGE_TITLE

Call ClearControls
End Sub


Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + cmdDate.Left
    .Top = Me.Top + cmdDate.Top - .Height / 2
    .selDate = txtDate
    .Show vbModal
    txtDate = .selDate
End With
End Sub


Private Sub cmdLoanName_Click()
Dim rst As Recordset
Dim SqlStr As String

SqlStr = "SELECT SchemeID,SchemeName,SchemeNameEnglish,LoanPurpose from LoanScheme Order by SchemeName"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Call FillView(m_frmLookUp.lvwReport, rst, True)
m_SchemeId = ""
m_frmLookUp.Show vbModal
If m_SchemeId = "" Then Exit Sub
If LoadLoanSchemeDetails(CLng(m_SchemeId)) Then
    'cmdCreate.Caption = "&Update"
    cmdCreate.Caption = GetResourceString(171)
    m_dbOperation = Update
    
End If
End Sub

Private Sub Form_Load()
'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
txtDate = gStrDate
Call SetKannadaCaption

m_dbOperation = Insert

txtEmpIntrate.Enabled = False
txtEmpPenalInt.Enabled = False

'Load loan Purposes
cmbPurpose.AddItem GetResourceString(88) '"Agriculural"
cmbPurpose.AddItem GetResourceString(252) '"Individual"
cmbPurpose.AddItem "Industrial"
cmbPurpose.AddItem "Business"
cmbPurpose.AddItem "Horticulure"

'Load Loan Category
cmbCategory.AddItem GetResourceString(440) '"Agricutural"
cmbCategory.ItemData(cmbCategory.newIndex) = 1
cmbCategory.AddItem GetResourceString(441) '"Non Agricutural"
cmbCategory.ItemData(cmbCategory.newIndex) = 2

'Load Loan TYpes
Dim LoanType As wis_LoanType
LoanType = wisCashCreditLoan
cmbLoanType.AddItem "Cash Credit Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType
LoanType = wisCropLoan
cmbLoanType.AddItem "Crop Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType
LoanType = wisVehicleloan
cmbLoanType.AddItem "Vehicle Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType
LoanType = wisIndividualLoan
cmbLoanType.AddItem "Individual Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType
LoanType = wisBKCC
cmbLoanType.AddItem "BKCC Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType
LoanType = wisSHGLoan
cmbLoanType.AddItem "SHG Loan"
cmbLoanType.ItemData(cmbLoanType.newIndex) = LoanType

'Loan Terms
cmbTerm.AddItem GetResourceString(222) '"Short Term"
cmbTerm.ItemData(cmbTerm.newIndex) = 1
cmbTerm.AddItem GetResourceString(223) '"Mid Term"
cmbTerm.ItemData(cmbTerm.newIndex) = 2
cmbTerm.AddItem GetResourceString(224) '"Long Term"
cmbTerm.ItemData(cmbTerm.newIndex) = 3
cmbTerm.AddItem "Conversion" 'GetResourceString(222) '
cmbTerm.ItemData(cmbTerm.newIndex) = 4
m_dbOperation = Insert


End Sub

Private Sub Form_Unload(Cancel As Integer)
gWindowHandle = 0
RaiseEvent WindowClosed
End Sub


Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_SchemeId = strSelection
End Sub
Private Sub txtLoanName_Change()
    If Len(Trim$(txtLoanName.Text)) < 1 Then
        lstHidden.Visible = False
    ElseIf lstHidden.ListCount Then
        lstHidden.Visible = True
        lstHidden.Top = 1100
    End If
End Sub

Private Sub txtLoanName_GotFocus()
    cmdCreate.Default = False
    lstHidden.Top = txtLoanName.Top + txtLoanName.Height
    lstHidden.Left = txtLoanName.Left
    lstHidden.Width = txtLoanName.Width
    lstHidden.Clear
    Set m_Ctrl = Me.ActiveControl
    lstHidden.Tag = ActiveControl.name
    
    Call LoadLoanSchemes
End Sub

Private Sub txtLoanName_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim txt As String
    Static SelNo As Integer
    Dim count As Integer
    Dim I As Integer
    txt = txtLoanName.Text
    count = lstHidden.ListCount - 1
    
    For I = 0 To count
        If InStr(1, lstHidden.List(I), txtLoanName.Text, vbTextCompare) = 1 Then
            lstHidden.Selected(I) = True
            lstHidden.ListIndex = I
            SelNo = lstHidden.ListIndex
            Exit For
        End If
        If I = count And Len(txtLoanName.Text) > 1 Then _
                        lstHidden.Selected(SelNo) = False
    Next I
    
End Sub

Private Sub txtLoanName_LostFocus()
lstHidden.Clear
lstHidden.Visible = False
cmdCreate.Default = True
End Sub



Private Sub txtLoanNameEnglish_GotFocus()
    Call ToggleWindowsKey(winScrlLock, False)
End Sub

Private Sub txtLoanNameEnglish_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub

Private Sub txtMonthDuration_LostFocus()

With txtMonthDuration
    If .Text = "" Then .Text = 0
       If .Text = 1 Then
          txtDayDuration = 30
       ElseIf .Text = 2 Then
           txtDayDuration = 60
       ElseIf .Text = 3 Then
           txtDayDuration = 90
       ElseIf .Text = 12 Then
          txtDayDuration.Text = 365
    End If
End With

End Sub

