VERSION 5.00
Begin VB.Form FrmDepScheme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New loan scheme"
   ClientHeight    =   5595
   ClientLeft      =   2715
   ClientTop       =   2235
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4995
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Canc&el"
      Height          =   315
      Left            =   3990
      TabIndex        =   31
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   315
      Left            =   3000
      TabIndex        =   32
      Top             =   5160
      Width           =   885
   End
   Begin VB.Frame fraLoanScheme 
      Caption         =   "LoanScheme"
      Height          =   5010
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4785
      Begin VB.ListBox lstHidden 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   1590
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1605
         Left            =   90
         TabIndex        =   33
         Top             =   3300
         Width           =   4605
         Begin VB.TextBox txtDayDuration 
            Height          =   315
            Left            =   3960
            TabIndex        =   25
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtIntrate 
            Height          =   315
            Left            =   1620
            TabIndex        =   19
            Top             =   210
            Width           =   555
         End
         Begin VB.TextBox txtMonthDuration 
            Height          =   315
            Left            =   1620
            TabIndex        =   23
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtPenalInt 
            Height          =   315
            Left            =   3960
            TabIndex        =   21
            Top             =   210
            Width           =   555
         End
         Begin VB.Frame Frame1 
            Caption         =   "Employee's"
            Height          =   645
            Left            =   0
            TabIndex        =   26
            Top             =   960
            Width           =   4605
            Begin VB.TextBox txtEmpPenalInt 
               Height          =   315
               Left            =   3840
               TabIndex        =   30
               Top             =   210
               Width           =   675
            End
            Begin VB.TextBox txtEmpIntrate 
               Height          =   315
               Left            =   1560
               TabIndex        =   28
               Top             =   210
               Width           =   675
            End
            Begin VB.Label lblEmpPenalInt 
               Caption         =   "P&enal interest :"
               Height          =   255
               Left            =   2370
               TabIndex        =   29
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label lblEmpIntrate 
               Caption         =   "Rate of &interest :"
               Height          =   255
               Left            =   60
               TabIndex        =   27
               Top             =   270
               Width           =   1335
            End
         End
         Begin VB.Label lblDayDuration 
            Caption         =   "Duration (d&ays)"
            Height          =   225
            Left            =   2280
            TabIndex        =   24
            Top             =   630
            Width           =   1545
         End
         Begin VB.Label lblIntrate 
            Caption         =   "&Rate of interest :"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblMonthDuaration 
            Caption         =   "Duration (&Month)"
            Height          =   225
            Left            =   60
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblPenalInt 
            Caption         =   "&Penal interest :"
            Height          =   255
            Left            =   2310
            TabIndex        =   20
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame fraName 
         Caption         =   "Frame3"
         Height          =   1065
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   4605
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   285
            Left            =   2760
            TabIndex        =   4
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtDate 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton cmdLoanName 
            Caption         =   "..."
            Height          =   285
            Left            =   4290
            TabIndex        =   6
            Top             =   570
            Width           =   285
         End
         Begin VB.TextBox txtLoanName 
            Height          =   315
            Left            =   1320
            TabIndex        =   7
            Top             =   570
            Width           =   2835
         End
         Begin VB.Label lblDate 
            Caption         =   "&Date :"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label lblLoanName 
            Caption         =   "Loan &Name :"
            Height          =   225
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Classification"
         Height          =   2055
         Left            =   90
         TabIndex        =   8
         Top             =   1260
         Width           =   4605
         Begin VB.CheckBox chkMemOnly 
            Alignment       =   1  'Right Justify
            Caption         =   "Loan only for Members"
            Height          =   345
            Left            =   1920
            TabIndex        =   9
            Top             =   150
            Value           =   1  'Checked
            Width           =   2565
         End
         Begin VB.ComboBox cmbPurpose 
            Height          =   315
            Left            =   1980
            TabIndex        =   11
            Top             =   540
            Width           =   2535
         End
         Begin VB.ComboBox cmbLoanType 
            Height          =   315
            Left            =   1980
            TabIndex        =   17
            Top             =   1620
            Width           =   2565
         End
         Begin VB.ComboBox cmbTerm 
            Height          =   315
            Left            =   1980
            TabIndex        =   15
            Top             =   1260
            Width           =   2565
         End
         Begin VB.ComboBox cmbCategory 
            Height          =   315
            Left            =   1980
            TabIndex        =   13
            Top             =   900
            Width           =   2565
         End
         Begin VB.Label lblPurpose 
            Caption         =   "&Purpose"
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label lblClassification 
            Caption         =   "Loan T&ype classifiaction"
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   1680
            Width           =   1905
         End
         Begin VB.Label lblTerm 
            Caption         =   "&Term :"
            Height          =   225
            Left            =   120
            TabIndex        =   14
            Top             =   1290
            Width           =   1575
         End
         Begin VB.Label lblCategary 
            Caption         =   "Loan &Categary"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   930
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmDepScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
'Private m_DbOperation As wis_DbOperation
Private m_SchemeId As String
Private m_Ctrl As Control

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

If txtEmpIntrate = "" Then
    MsgBox "Please Enter Rate of Interest", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtEmpIntrate
    Exit Function
End If
If Not IsNumeric(txtEmpIntrate) Then
    MsgBox "Invalid Interest rate Specified", vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtEmpIntrate
    Exit Function
End If
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
    chkMemOnly.value = vbChecked

    cmdCreate.Caption = LoadResString(gLangOffSet + 15)
    
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
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(rst, adOpenForwardOnly) < 0 Then Exit Function

Dim LoanType As wis_LoanType

'Now set the values on the form controls
txtLoanName = FormatField(rst("SchemeName"))
cmbPurpose.Text = FormatField(rst("LoanPurpose"))
txtIntrate = FormatField(rst("IntRate"))
txtPenalInt = FormatField(rst("PenalIntRate"))
txtMonthDuration = FormatField(rst("MonthDuration"))
txtDayDuration = FormatField(rst("DayDuration"))
txtEmpIntrate = FormatField(rst("EmpIntRate"))
txtEmpPenalInt = FormatField(rst("EmpPenalIntRate"))
chkMemOnly.value = IIf(FormatField(rst("OnlyMember")), vbChecked, vbUnchecked)
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

Public Function LoadSchemeDetail(SchmeID As Integer) As Boolean
LoadSchemeDetail = False



LoadSchemeDetail = True

End Function

Private Sub LoadLoanSchemes()
Dim rst As Recordset
Dim SqlStr As String

'Get the LoanSchemeName & SchemeID from the database to
'show it into lsthidden Listbox
SqlStr = "SELECT SchemeID,SchemeName from LoanScheme"
gDBTrans.SQLStmt = SqlStr
If gDBTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub

lstHidden.Clear
While Not rst.EOF
    lstHidden.AddItem FormatField(rst("SchemeName"))
    lstHidden.ItemData(lstHidden.NewIndex) = FormatField(rst("SchemeID"))
    rst.MoveNext
Wend
End Sub

'This function Saves the LoanSchemes into the LoanScheme table
'
Private Function SaveLoanSchemeDetails() As Boolean

SaveLoanSchemeDetails = False
Dim LoanSchemeName As String
Dim LoanSchemeID As Long
Dim LoanPurpose As String
Dim LoanCategory As Byte
Dim LoanTerm As Byte
Dim LoanType As wis_LoanType
Dim SqlStr As String

Dim AsonDate As Date
AsonDate = FormatDate(txtDate)

LoanSchemeName = Trim$(txtLoanName)
LoanCategory = cmbCategory.ItemData(cmbCategory.ListIndex)

LoanTerm = cmbTerm.ItemData(cmbTerm.ListIndex)
LoanPurpose = Trim$(cmbPurpose.Text)

'Get the LoanSchemeId
SqlStr = "Select Max(SchemeID) from LoanScheme"
gDBTrans.SQLStmt = SqlStr
If gDBTrans.SQLFetch <> 1 Then Exit Function
LoanSchemeID = FormatField(gDBTrans.rst(0)) + 1
LoanType = cmbLoanType.ItemData(cmbLoanType.ListIndex)

'Now Insert the values
SqlStr = ""
gDBTrans.BeginTrans

If m_DbOperation = Insert Then
    m_SchemeId = LoanSchemeID
    SqlStr = "INSERT INTO LoanScheme (SchemeID,SchemeName,Category," & _
        " TermType,LoanType,MonthDuration,DayDuration," & _
        " IntRate,PenalIntRate,LoanPurpose,OnlyMember )" & _
        " Values ( " & _
        LoanSchemeID & "," & _
        AddQuotes(LoanSchemeName, True) & "," & _
        LoanCategory & "," & _
        LoanTerm & "," & _
        LoanType & "," & _
        Val(txtMonthDuration.Text) & "," & _
        Val(txtDayDuration.Text) & "," & _
        Val(Trim$(txtIntrate.Text)) & "," & _
        Val(Trim$(txtPenalInt.Text)) & "," & _
        AddQuotes(LoanPurpose, True) & ", " & _
        IIf(chkMemOnly = vbChecked, True, False) & " )"
    
    gDBTrans.SQLStmt = SqlStr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If
        
ElseIf m_DbOperation = Update Then
    SqlStr = "UPDATE LoanScheme set " & _
        " SchemeName = " & AddQuotes(LoanSchemeName, True) & "," & _
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

    gDBTrans.SQLStmt = SqlStr
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Function
    End If

End If
gDBTrans.CommitTrans

'Now store the Loan interest

'Dim clsInt As New clsInterest
'Call clsInt.SaveInterest(wisLoanAccount, "MemRegularInterest" & m_SchemeId, Val(txtIntrate), AsOnDate)
'Call clsInt.SaveInterest(wisLoanAccount, "MemPenalInterest" & m_SchemeId, Val(txtPenalInt), AsOnDate)
'Call clsInt.SaveInterest(wisLoanAccount, "EmpRegularInterest" & m_SchemeId, Val(txtEmpIntrate), AsOnDate)
'Call clsInt.SaveInterest(wisLoanAccount, "EmpPenalInterest" & m_SchemeId, Val(txtEmpPenalInt), AsOnDate)

'Set clsInt = Nothing

SaveLoanSchemeDetails = True

End Function

Private Sub SetKannadaCaption()
Dim ctrl As Control
On Error Resume Next
    For Each ctrl In Me
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then ctrl.Font.Size = gFontSize
    Next
Err.Clear

lblDate = LoadResString(gLangOffSet + 37) 'Date
lblLoanName = LoadResString(gLangOffSet + 214) 'Loan Scheme
lblCategary = LoadResString(gLangOffSet + 317) 'Categaey
lblTerm = LoadResString(gLangOffSet + 319)      'Term
lblClassification = LoadResString(gLangOffSet + 58) & " " & LoadResString(gLangOffSet + 318) 'Loan Classification
lblIntrate = LoadResString(gLangOffSet + 186) 'Rate Of Interest
lblPenalInt = LoadResString(gLangOffSet + 345) & " " & LoadResString(gLangOffSet + 305) 'Penal Int rate
lblEmpIntrate = LoadResString(gLangOffSet + 186) 'Rate Of Int
lblEmpPenalInt = LoadResString(gLangOffSet + 345) & " " & LoadResString(gLangOffSet + 305) 'Penal Int rate

cmdCreate.Caption = LoadResString(gLangOffSet + 15)  'Create
cmdCancel.Caption = LoadResString(gLangOffSet + 2) 'Cancel


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdCreate_Click()
'Check the validation
If Not CheckValidation Then
    Exit Sub
End If
If Not SaveLoanSchemeDetails Then
    Exit Sub
End If
MsgBox "Saved the Loan Scheme Details ", vbInformation, wis_MESSAGE_TITLE
Call ClearControls
End Sub


Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + fraLoanScheme.Left + cmdDate.Left
    .Top = Me.Top + fraLoanScheme.Top + cmdDate.Top - .Height / 2
    .SelDate = txtDate
    .Show vbModal
    txtDate = .SelDate
End With
End Sub

Private Sub cmdLoanName_Click()
Dim rst As Recordset
Dim SqlStr As String

SqlStr = "SELECT SchemeID,SchemeName,LoanPurpose from LoanScheme Order by SchemeName"

gDBTrans.SQLStmt = SqlStr
If gDBTrans.SQLFetch < 1 Then Exit Sub
Set rst = gDBTrans.rst.Clone
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Call FillView(m_frmLookUp.lvwReport, rst, True)
m_SchemeId = ""
m_frmLookUp.Show vbModal
If m_SchemeId = "" Then Exit Sub
If LoadLoanSchemeDetails(CLng(m_SchemeId)) Then
    'cmdCreate.Caption = "&Update"
    cmdCreate.Caption = LoadResString(gLangOffSet + 171)
    m_DbOperation = Update
End If
End Sub


Private Sub Form_Load()
Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
txtDate = FormatDate(gStrDate)
Call SetKannadaCaption

m_DbOperation = Insert

'Load loan Purposes
cmbPurpose.AddItem "Agriculural"
cmbPurpose.AddItem "Individual"
cmbPurpose.AddItem "Industrial"
cmbPurpose.AddItem "Business"
cmbPurpose.AddItem "Horticulure"

'Load Loan Category
cmbCategory.AddItem "Agricutural"
cmbCategory.ItemData(cmbCategory.NewIndex) = 1
cmbCategory.AddItem "Non Agricutural"
cmbCategory.ItemData(cmbCategory.NewIndex) = 2

'Load Loan TYpes
Dim LoanType As wis_LoanType
LoanType = wisCashCreditLoan
cmbLoanType.AddItem "Cash Credit Loan"
cmbLoanType.ItemData(cmbLoanType.NewIndex) = LoanType
LoanType = wisCropLoan
cmbLoanType.AddItem "Crop Loan"
cmbLoanType.ItemData(cmbLoanType.NewIndex) = LoanType
LoanType = wisVehicleloan
cmbLoanType.AddItem "Vehicle Loan"
cmbLoanType.ItemData(cmbLoanType.NewIndex) = LoanType
LoanType = wisIndividualLoan
cmbLoanType.AddItem "Individual Loan"
cmbLoanType.ItemData(cmbLoanType.NewIndex) = LoanType
LoanType = wisBKCC
cmbLoanType.AddItem "BKCC Loan"
cmbLoanType.ItemData(cmbLoanType.NewIndex) = LoanType

'Loan Terms
cmbTerm.AddItem LoadResString(gLangOffSet + 222) '"Short Term"
cmbTerm.ItemData(cmbTerm.NewIndex) = 1
cmbTerm.AddItem LoadResString(gLangOffSet + 223) '"Mid Term"
cmbTerm.ItemData(cmbTerm.NewIndex) = 2
cmbTerm.AddItem LoadResString(gLangOffSet + 224) '"Long Term"
cmbTerm.ItemData(cmbTerm.NewIndex) = 3
cmbTerm.AddItem "Conversion" 'LoadResString(gLangOffSet + 222) '
cmbTerm.ItemData(cmbTerm.NewIndex) = 4
m_DbOperation = Insert

End Sub





Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_SchemeId = strSelection

End Sub


Private Sub txtLoanName_Change()
    If Len(Trim$(txtLoanName.Text)) < 1 Then
        lstHidden.Visible = False
    ElseIf lstHidden.ListCount Then
        lstHidden.Visible = True
    End If
End Sub

Private Sub txtLoanName_GotFocus()
    cmdCreate.Default = False
    lstHidden.Top = fraLoanScheme.Top + txtLoanName.Top + txtLoanName.Height
    lstHidden.Left = fraLoanScheme.Left + txtLoanName.Left
    lstHidden.Width = txtLoanName.Width
    lstHidden.Clear
    Set m_Ctrl = Me.ActiveControl
    lstHidden.Tag = ActiveControl.Name
    
    Call LoadLoanSchemes
End Sub

Private Sub txtLoanName_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Txt As String
    Static SelNo As Integer
    Dim Count As Integer
    Dim I As Integer
    Txt = txtLoanName.Text
    Count = lstHidden.ListCount - 1
    For I = 0 To Count
        If InStr(1, lstHidden.List(I), txtLoanName.Text, vbTextCompare) = 1 Then
            lstHidden.Selected(I) = True
            lstHidden.ListIndex = I
            SelNo = lstHidden.ListIndex
            Exit For
        End If
        If I = Count And _
                        Len(txtLoanName.Text) > 1 Then _
            lstHidden.Selected(SelNo) = False
    Next I
    

End Sub


Private Sub txtLoanName_LostFocus()
lstHidden.Clear
lstHidden.Visible = False
cmdCreate.Default = True
End Sub



