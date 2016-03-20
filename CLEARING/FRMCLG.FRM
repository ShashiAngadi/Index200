VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmClearing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques clearing scheduler"
   ClientHeight    =   6705
   ClientLeft      =   855
   ClientTop       =   1365
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5310
      TabIndex        =   14
      Top             =   6250
      Width           =   1215
   End
   Begin VB.Frame fraClearing 
      Caption         =   "Cleared Cheques"
      Height          =   5500
      Index           =   2
      Left            =   210
      TabIndex        =   37
      Top             =   540
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   3945
         Left            =   90
         TabIndex        =   19
         Top             =   960
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6959
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   2
      End
      Begin VB.TextBox txtgrd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4170
         TabIndex        =   40
         Top             =   390
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdDiscount 
         Caption         =   "Discount"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5010
         TabIndex        =   23
         Top             =   4950
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5040
         TabIndex        =   39
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeSelectAll 
         Caption         =   "&Deselect All"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5040
         TabIndex        =   38
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   400
         Left            =   2460
         TabIndex        =   21
         Top             =   4950
         Width           =   1215
      End
      Begin VB.TextBox txtDate2 
         Height          =   315
         Left            =   2130
         TabIndex        =   17
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   400
         Left            =   1050
         TabIndex        =   20
         Top             =   4950
         Width           =   1215
      End
      Begin VB.CommandButton cmdBounce 
         Caption         =   "Bounce"
         Enabled         =   0   'False
         Height          =   400
         Left            =   3750
         TabIndex        =   22
         Top             =   4950
         Width           =   1215
      End
      Begin VB.Label lblCheque 
         Caption         =   "Cheque Details"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   690
         Width           =   3285
      End
      Begin VB.Label lblDate2 
         Caption         =   "Date :"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   330
         Width           =   1815
      End
   End
   Begin VB.Frame fraClearing 
      Caption         =   "Reposrts"
      Height          =   5500
      Index           =   3
      Left            =   210
      TabIndex        =   36
      Top             =   540
      Width           =   6375
      Begin VB.OptionButton optReports 
         Caption         =   "Cheques Discount"
         Height          =   315
         Index           =   4
         Left            =   510
         TabIndex        =   28
         Top             =   2730
         Width           =   3045
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Cheques pending"
         Height          =   315
         Index           =   3
         Left            =   510
         TabIndex        =   27
         Top             =   2136
         Width           =   3045
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Cheque Submitted "
         Height          =   315
         Index           =   0
         Left            =   510
         TabIndex        =   24
         Top             =   360
         Width           =   3045
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Cheques Cleared "
         Height          =   315
         Index           =   1
         Left            =   510
         TabIndex        =   25
         Top             =   952
         Width           =   3045
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Cheques not cleared "
         Height          =   315
         Index           =   2
         Left            =   510
         TabIndex        =   26
         Top             =   1544
         Width           =   3045
      End
      Begin VB.TextBox txtFromDate 
         Height          =   345
         Left            =   1470
         TabIndex        =   31
         Top             =   3960
         Width           =   1275
      End
      Begin VB.CommandButton cmdFromDate 
         Caption         =   "..."
         Height          =   315
         Left            =   2820
         TabIndex        =   30
         Top             =   3960
         Width           =   315
      End
      Begin VB.TextBox txtToDate 
         Height          =   345
         Left            =   4380
         TabIndex        =   34
         Top             =   3960
         Width           =   1245
      End
      Begin VB.CommandButton cmdToDate 
         Caption         =   "..."
         Height          =   315
         Left            =   5730
         TabIndex        =   33
         Top             =   3960
         Width           =   315
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   400
         Left            =   4920
         TabIndex        =   35
         Top             =   4650
         Width           =   1215
      End
      Begin VB.Label lblFrom 
         Caption         =   "From :"
         Height          =   315
         Left            =   270
         TabIndex        =   29
         Top             =   3990
         Width           =   1095
      End
      Begin VB.Label lblTo 
         Caption         =   "To :"
         Height          =   315
         Left            =   3510
         TabIndex        =   32
         Top             =   3990
         Width           =   705
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6000
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6070
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   10716
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Schedule for Clearing"
            Key             =   "clearing"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Update accounts"
            Key             =   "update"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraClearing 
      Caption         =   "Enter clearing instrument details here :"
      Height          =   5500
      Index           =   1
      Left            =   210
      TabIndex        =   15
      Top             =   540
      Width           =   6375
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   375
         Left            =   2130
         TabIndex        =   46
         Top             =   3831
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.CheckBox chkDiscount 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount"
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Top             =   3900
         Width           =   1815
      End
      Begin VB.ComboBox cmbBank 
         Height          =   315
         Left            =   2130
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2322
         Width           =   3915
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   2130
         TabIndex        =   5
         Top             =   1296
         Width           =   1725
      End
      Begin VB.ComboBox cmbAcType 
         Height          =   315
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   813
         Width           =   3885
      End
      Begin VB.TextBox txtChqNo 
         Height          =   345
         Left            =   2130
         TabIndex        =   9
         Top             =   3318
         Width           =   3885
      End
      Begin VB.TextBox txtBank 
         Height          =   345
         Left            =   2130
         TabIndex        =   8
         Top             =   2805
         Width           =   3885
      End
      Begin VB.TextBox txtDate1 
         Height          =   345
         Left            =   2130
         TabIndex        =   2
         Top             =   300
         Width           =   1725
      End
      Begin VB.CommandButton cmdLookUp 
         Caption         =   "..."
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   1320
         Width           =   315
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Accept"
         Height          =   400
         Left            =   4830
         TabIndex        =   13
         Top             =   5010
         Width           =   1200
      End
      Begin VB.TextBox txtRemarks 
         Height          =   345
         Left            =   2130
         TabIndex        =   12
         Top             =   4380
         Width           =   3885
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3930
         TabIndex        =   1
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblRemarks 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   4530
         Width           =   1800
      End
      Begin VB.Label lblDate1 
         Caption         =   "Date :"
         Height          =   315
         Left            =   150
         TabIndex        =   49
         Top             =   345
         Width           =   1800
      End
      Begin VB.Label lblBank 
         AutoSize        =   -1  'True
         Caption         =   "Bank Details :"
         Height          =   315
         Left            =   150
         TabIndex        =   48
         Top             =   2960
         Width           =   1800
      End
      Begin VB.Label lblInstrNo 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No. :"
         Height          =   315
         Left            =   150
         TabIndex        =   47
         Top             =   3483
         Width           =   1800
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         Height          =   315
         Left            =   150
         TabIndex        =   45
         Top             =   4006
         Width           =   1800
      End
      Begin VB.Label lblAccountType 
         Caption         =   "Account Type :"
         Height          =   315
         Left            =   150
         TabIndex        =   44
         Top             =   868
         Width           =   1800
      End
      Begin VB.Label lblAccountNo 
         AutoSize        =   -1  'True
         Caption         =   "Account No. :"
         Height          =   315
         Left            =   150
         TabIndex        =   43
         Top             =   1391
         Width           =   1800
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Custmer Name :"
         Height          =   315
         Left            =   150
         TabIndex        =   42
         Top             =   1914
         Width           =   1800
      End
      Begin VB.Label lblClearBank 
         Caption         =   "Clearing Bank"
         Height          =   315
         Left            =   150
         TabIndex        =   41
         Top             =   2437
         Width           =   1800
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2130
         TabIndex        =   6
         Top             =   1809
         Width           =   3885
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmClearing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents LookUpDialog As frmLookUp
Attribute LookUpDialog.VB_VarHelpID = -1
Private m_SelectedAccount As String
Private m_Status() As Integer
'Private m_CustID As Long
Private m_Loading As Boolean

Public Event ShowReport(ReportType As wis_ClearingReports, _
        fromDate As String, toDate As String)
Public Event WindowClosed()

Private Function DeleteNotes(TransID As Long) As Boolean
Dim I As Integer
Dim MaxI As Integer
Dim Selected As Boolean
Dim rst As Recordset

I = 0

gDbTrans.BeginTrans
With grd
    MaxI = .Rows - 1
    For I = 1 To MaxI
       If m_Status(I) < 0 Then
          'First Get The Moule Id & NotEId Which  You Have to Delete
          gDbTrans.SqlStmt = "Select * From ClearingTab " & _
                    "WHERE ChequeID = " & Val(.RowData(I))
          If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Selected = True
       End If
    Next I
End With
gDbTrans.CommitTrans

If Selected Then Call LoadNotes

End Function

Private Function LoadNotes() As Boolean

On Error GoTo ErrLine

Dim Notetype As Byte


Dim rst As Recordset

Notetype = 1  ' bankers Note

Dim I As Integer
Dim Status As wis_ChequeStatus
'first query the Table To Load Only Bankers Note
Status = wisPending
gDbTrans.SqlStmt = "SELECT A.ChequeID,ChequeNo,BankName,Status," & _
    " SubmitDate,AccHeadID,AccID,Amount FROM ClearingTab A,ClearingTrans B" & _
    " WHERE A.ChequeId = b.ChequeID And TransID = (Select Max(TransID)" & _
            " From ClearingTrans C Where C.ChequeId = a.ChequeID )" & _
    " And (Status = " & wisPending & " OR Status = " & wisDiscount & ")"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function

I = 0
rst.MoveFirst
TabStrip1.Tabs(2).Selected = True
m_Loading = True
Dim strField As String
    
With grd
    .Visible = False
    .Clear
    .Rows = 2
'    .RowHeight(0) = .RowHeight(1) '* 1.5
'    .Row = 0
'    .Col = 0: txtgrd = "Collection Charges"
'    .Text = txtgrd
'    .Col = 1: txtgrd = "Discount Charges"
'    .Text = txtgrd
    txtgrd = ""
    .Cols = 7
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = "Select"
    .Col = 1: .Text = "Account Name"
    .Col = 2: .Text = "Account No"
    .Col = 3: .Text = "Cheque No"
    .Col = 4: .Text = "Amount"
    .Col = 5: .Text = "Collection Charges"
    .Col = 6: .Text = "Discount Charges"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 1800
    .ColWidth(2) = 800
    .ColWidth(3) = 1200
    .ColWidth(4) = 1500
    .ColWidth(5) = 800
    .ColWidth(6) = 800
End With

I = 0
ReDim m_Status(0)
grd.Row = 0
Do While rst.EOF = False
    I = I + 1
    
    With grd
        ReDim Preserve m_Status(I)
        m_Status(I) = FormatField(rst("Status"))
        m_Status(I) = Not m_Status(I)
        .Row = .Row + 1
        .Col = 0: Set .CellPicture = LoadResPicture(143, vbResIcon)
        .Col = 1: .Text = GetHeadName(FormatField(rst("AccHeadID")))
        .Col = 2: .Text = GetAccountNumber(rst("AccHeadID"), rst("AccID"))
        .Col = 3: .Text = FormatField(rst("ChequeNo"))
        .Col = 4: .Text = FormatField(rst("Amount"))
        .RowData(.Row) = rst("chequeID")
        Status = FormatField(rst("Status"))
        If Status = wisPending Then
            .Col = 6
            .CellBackColor = &H8000000F 'vbBlack
        End If
        .Rows = .Rows + 1
    End With
    rst.MoveNext
Loop
grd.Rows = grd.Rows - 1

With grd
    .Rows = rst.RecordCount + 1
    'If .ScrollBars And flexScrollBarVertical Then
        '.ScrollBars = flexScrollBarVertical
    '    .ColWidth(4) = 830
    '    .ColWidth(5) = 830
    'End If

    .Visible = True
End With

If rst.RecordCount > 0 Then
    cmdDelete.Enabled = True
    cmdClear.Enabled = True
    cmdBounce.Enabled = True
    cmdDiscount.Enabled = True
    cmdSelectAll.Enabled = True
    cmdDeSelectAll.Enabled = True
End If

LoadNotes = True
m_Loading = False

Exit Function

ErrLine:

LoadNotes = False
grd.Visible = True
m_Loading = False
End Function

Private Sub ResetUserInterface()
Dim Ctrl As Control

For Each Ctrl In Me
    If TypeOf Ctrl Is TextBox Then
        Ctrl.Text = ""
    ElseIf TypeOf Ctrl Is ComboBox Or TypeOf Ctrl Is ListBox Then
        Ctrl.ListIndex = IIf(Ctrl.ListCount, 0, -1)
    End If
Next Ctrl

txtName = ""
txtDate1.Text = gStrDate
txtDate2.Text = gStrDate

End Sub

'function saves the details into notetab and Transaction Tab
Private Function SaveDetails() As Boolean

On Error GoTo ErrLine

Dim AccId As Long
Dim AccNum As String

Dim ChqNo As String
Dim Amount As Currency
Dim AccHeadID  As Long
Dim ModuleId As wisModules
Dim NoteId As Long
Dim Notes As String
Dim Notetype As Byte
Dim TransID As Long
Dim BankID As Long
Dim TransDate As Date
Dim UserID As Integer

UserID = gCurrUser.UserID
TransDate = GetSysFormatDate(txtDate1)
AccHeadID = cmbAcType.ItemData(cmbAcType.ListIndex)
ModuleId = GetModuleIDFromHeadID(AccHeadID)
AccNum = txtAccNo.Text
AccId = CLng(txtAccNo.Tag)
ChqNo = Trim(txtChqNo.Text)
Amount = txtAmount
NoteId = 0
Notetype = 1   'for Bankers Note

'Store The Varaible In Necessary Tables first Query Table To Get The New Noteid
gDbTrans.SqlStmt = "SELECT MAX(NoteID) FROM NoteTab " & _
        " WHERE ModuleId = " & ModuleId

Dim rst As Recordset

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                NoteId = Val(FormatField(rst(0))) + 1
NoteId = NoteId + 1

'GetThe Max TransID for This ModuleId
Dim ChequeId As Long
ChequeId = 0

gDbTrans.SqlStmt = "Select MAX(ChequeID) From ClearingTab "
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then ChequeId = Val(FormatField(rst(0)))
ChequeId = ChequeId + 1

' insert the Corresponding Val into NoteTeb
Notes = "ChequeNo : " & ChqNo & " of Amount :" & CStr(Amount) & _
            " of bank = " & txtBank.Text & _
            " sent For Clearing on " & txtDate1.Text
gDbTrans.SqlStmt = "INSERT INTO NoteTab (NoteId , ModuleID ,NoteType ," _
                    & "Accid,Notes )VALUES (" _
                    & NoteId & "," & ModuleId & "," & Notetype & "," _
                    & AccId & "," & AddQuotes(Notes, True) & ")"
gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Go For Insertion Into Clearing Tab
Notes = ""
If cmbAcType.ListIndex = 0 Then Notes = "SB "
If cmbAcType.ListIndex = 1 Then Notes = "CA "

Dim Status As wis_ChequeStatus
Dim Loan As Boolean
Loan = IIf(cmbAcType.ListIndex > 3, True, False)
Status = wisPending
Notes = Notes & " Account No = " & txtAccNo.Text & " Cheque " & ChqNo & _
        " sent for clearing on  " & txtDate1.Text

BankID = cmbBank.ItemData(cmbBank.ListIndex)

gDbTrans.SqlStmt = "INSERT INTO ClearingTab " & _
        " (ChequeId, AccHeadID,AccId,SubmitDate, " & _
        " ChequeNo, Amount, BankName," & _
        " ClearingBankID,Remarks) VALUES (" & _
        ChequeId & "," & AccHeadID & "," & _
        Val(txtAccNo.Tag) & ",#" & TransDate & "#," & _
        AddQuotes(ChqNo, True) & "," & Amount & "," & _
        AddQuotes(txtBank, True) & "," & BankID & "," & _
        AddQuotes(txtRemarks, True) & ")"
                    
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

TransID = 1
Status = wisPending
gDbTrans.SqlStmt = "INSERT INTO ClearingTrans " & _
        " (ChequeId, TransID,TransDate,Status) " & _
        " VALUES (" & ChequeId & "," & TransID & "," & _
        " #" & TransDate & "#," & Status & ")"
                    
If Not gDbTrans.SQLExecute Then GoTo ErrLine


If chkDiscount.Value = vbChecked Then
    TransID = TransID + 1
    Status = wisDiscount
    
    gDbTrans.SqlStmt = "INSERT INTO ClearingTrans " & _
            " (ChequeId, TransID,TransDate,Status) " & _
            " VALUES (" & ChequeId & "," & TransID & "," & _
            " #" & TransDate & "#," & Status & ")"
                        
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    Dim rstTemp As Recordset
    Dim bankClass As clsBankAcc
    Dim ClgHeadId As Long
    Dim ContraID As Long
    
    ContraID = GetMaxContraTransID + 1
    
    Set bankClass = New clsBankAcc
    
    ClgHeadId = bankClass.GetHeadIDCreated(GetResourceString(168, 36), _
                LoadResourceStringS(168, 36), parReceivable, 0, wis_BankAccounts)
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,TransID,AccHeadID,AccId,TransType,Amount,UserID) " & _
        " VALUES (" & ContraID & "," & TransID & "," & ClgHeadId & "," & ClgHeadId & "," & _
        wContraWithdraw & "," & Amount & "," & UserID & ")"
        
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
    'Now Make the Transaction To the Ledger heads
    If Not bankClass.UpdateContraTrans(ClgHeadId, AccHeadID, _
                            txtAmount, TransDate) Then GoTo ErrLine
    Dim ClassObj As Object
    If ModuleId = wis_SBAcc Then Set ClassObj = New clsSBAcc
    If ModuleId = wis_CAAcc Then Set ClassObj = New ClsCAAcc
    
    If ClassObj.DepositAmount(Val(txtAccNo.Tag), _
            txtAmount, "Discount " & txtChqNo.Text, GetSysFormatDate(txtDate1)) = 0 Then GoTo ErrLine
    Set ClassObj = Nothing
End If

gDbTrans.CommitTrans

SaveDetails = True

ErrLine:
    gDbTrans.RollBack
End Function

Private Function SaveIntoTransTab(Status As wis_ChequeStatus) As Boolean

'Check the Date Validation
If Not DateValidate(txtDate2.Text, "/", True) Then
   MsgBox GetResourceString(501), vbExclamation, wis_MESSAGE_TITLE
   ActivateTextBox txtDate2
   Exit Function
End If

Dim rst As Recordset
Dim ModuleId As wisModules
Dim ClgHeadId As Long
Dim AccHeadID As Long
Dim BankHeadID As Long
Dim ClassObj As Object
Dim TransDate As Date
Dim ChequeId As Long
Dim PrevStatus As wis_ChequeStatus
Dim bankClass As clsBankAcc


Dim TotaClrAmount As Currency
Dim TotalDiscAmount As Currency
Dim TransID As Integer
Dim SbAmount As Currency
Dim SbHeadID As Long

Dim CaAmount As Currency
Dim CAHeadID As Long

Dim DiscAmount As Currency
Dim ClrAmount As Currency
Dim Amount As Currency

Set bankClass = New clsBankAcc

'Get the Details of the Cheque
gDbTrans.SqlStmt = "Select A.ChequeID,AccHeadID,AccID,Amount," & _
    " ClearingBankID,Status,TransDate From ClearingTab A,ClearingTrans B" & _
    " Where A.ChequeId = B.ChequeId AND TransId = " & _
    "(SELECT Max(TransID) From ClearingTrans C Where C.ChequeId = A.ChequeId And " & _
        "(Status = " & wisPending & " OR Status = " & wisDiscount & " ));"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    MsgBox "Database ERROR - Clearing Tab", vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If


Dim I As Integer
Dim MaxI As Integer
I = 0

SbHeadID = GetIndexHeadID(GetResourceString(421))
CAHeadID = GetIndexHeadID(GetResourceString(422))


MaxI = grd.Rows - 1 'lstCheques.ListCount - 1
For I = 1 To MaxI
    If m_Status(I) < 0 Then
        ChequeId = Val(grd.RowData(I))
        rst.MoveFirst
        rst.Find "ChequeID = " & ChequeId
        PrevStatus = FormatField(rst("Status"))
        If PrevStatus = Status Then
            'msgbox "Unable to perform this operation"
            MsgBox GetResourceString(535), , wis_MESSAGE_TITLE
            GoTo ErrLine
        End If
        Amount = FormatField(rst("Amount"))
        If rst("AccHeadID") = SbHeadID Then
            SbAmount = SbAmount + Amount
        Else
            CaAmount = CaAmount + Amount
        End If
        
        If PrevStatus = wisDiscount Then _
            TotalDiscAmount = TotalDiscAmount + Amount
        If PrevStatus = wisPending Then _
            TotaClrAmount = TotaClrAmount + Amount
        If Status = wisCleared Then
            With grd
                .Row = I + 1
                .Col = 0: DiscAmount = DiscAmount + Val(.Text)
                .Col = 1: ClrAmount = ClrAmount + Val(.Text)
            End With
        End If
    End If
Next


TransID = 100
Dim rstTemp As Recordset
TransID = GetMaxContraTransID + 1

gDbTrans.BeginTrans

ClgHeadId = bankClass.GetHeadIDCreated(GetResourceString(168, 36), _
            LoadResourceStringS(168, 36), parReceivable, 0, wis_BankAccounts)
rst.MoveFirst
BankHeadID = FormatField(rst("ClearingBankID"))

    'Now Transfer the Amount From Clering account to indvidual account
    'Make transaction to contra trans table
    Dim UserID As Long
    UserID = gCurrUser.UserID
    
gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
        " VALUES (" & TransID & "," & BankHeadID & "," & BankHeadID & "," & _
        wContraWithdraw & "," & TotaClrAmount + TotalDiscAmount & "," & UserID & ")"
        
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

If TotalDiscAmount Then
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
                " VALUES (" & TransID & "," & BankHeadID & "," & BankHeadID & "," & _
                wContraWithdraw & "," & TotalDiscAmount & "," & UserID & ")"
        
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo ErrLine
    End If
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
                " VALUES (" & TransID & "," & ClgHeadId & "," & ClgHeadId & "," & _
                wContraDeposit & "," & TotalDiscAmount & "," & UserID & ")"
        
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        GoTo ErrLine
    End If
End If


For I = 1 To MaxI
    If m_Status(I) >= 0 Then GoTo NextLine
    rst.MoveFirst
    
    ChequeId = Val(grd.RowData(I))
    rst.Find "ChequeID = " & ChequeId
    
    PrevStatus = FormatField(rst("Status"))
    BankHeadID = FormatField(rst("ClearingBankId"))
    AccHeadID = FormatField(rst("AccHeadId"))

    ModuleId = GetModuleIDFromHeadID(AccHeadID)
    TransDate = GetSysFormatDate(txtDate2.Text)
    
    If ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then Set ClassObj = New clsSBAcc
    If ModuleId = wis_CAAcc Then Set ClassObj = New ClsCAAcc
    If ModuleId = wis_BKCCLoan Or ModuleId = wis_BKCC Then Set ClassObj = New clsBkcc
    
    If PrevStatus = wisPending Then
        If ClassObj.DepositAmount(Val(FormatField(rst("AccId"))), _
           FormatField(rst("Amount")), _
           "Clg " & FormatField(rst("ChequeId")), TransDate) = 0 Then GoTo ErrLine
    End If
    If Status = wisCleared Then
        With grd
            .Row = I + 1
            .Col = 0: DiscAmount = Val(.Text)
            .Col = 1: DiscAmount = DiscAmount + Val(.Text)
        End With
        If DiscAmount Then
            If ClassObj.WithdrawAmount(Val(FormatField(rst("AccId"))), _
               DiscAmount, "Clearing Charges", TransDate) = 0 Then GoTo ErrLine
        End If
    End If
    'Upadate this Cheque in the Clearing Tab
    
    'Get MAx Transaction ID
    gDbTrans.SqlStmt = "Select Max(TransID) From ClearingTrans " & _
                        " WHERE ChequeID = " & ChequeId
    TransID = 0
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then TransID = FormatField(rstTemp(0))
    TransID = TransID + 1
    
    gDbTrans.SqlStmt = "Insert Into ClearingTrans " & _
                "(ChequeID,Status,TransDate,Transid) " & _
                " VALUES ( " & ChequeId & "," & Status & "," & _
                " #" & TransDate & "#," & TransID & ")"
            
    If Not gDbTrans.SQLExecute Then
       gDbTrans.RollBack
       GoTo ErrLine
    End If
    
NextLine:

Next

gDbTrans.CommitTrans

'Now Go For Deletion Of Corresponding In Records
Dim Notes As clsNotes
Set Notes = New clsNotes
'Call Notes.DeleteNotes(FormatField(Rst("Moduleid")), FormatField(Rst("Accid")), FormatField(Rst("Noteid")))
SaveIntoTransTab = True

Exit Function

ErrLine:
   If Err Then _
      MsgBox "Error In SaveTransTab - Clearing Module", vbCritical, wis_MESSAGE_TITLE
   
End Function

Private Function SaveClearing() As Boolean

Dim Status As wis_ChequeStatus
Status = wisCleared

'Check the Date Validation
If Not DateValidate(txtDate2.Text, "/", True) Then
   MsgBox GetResourceString(501), vbExclamation, wis_MESSAGE_TITLE
   ActivateTextBox txtDate2
   Exit Function
End If

Dim rst As Recordset
Dim ModuleId As wisModules
Dim ClgHeadId As Long
Dim ClgIncomeID As Long
Dim AccHeadID As Long
Dim BankHeadID As Long
Dim SbHeadID As Long
Dim CAHeadID As Long

Dim ClassObj As Object

Dim TransDate As Date
Dim ChequeId As Long
Dim PrevStatus As wis_ChequeStatus
Dim bankClass As clsBankAcc

Dim TotalDiscAmount As Currency
Dim SBCharges As Currency
Dim CACharges As Currency
Dim SbAmount As Currency
Dim CaAmount As Currency

Dim TransID As Integer

Dim DiscAmount As Currency
Dim ClrAmount As Currency
Dim Amount As Currency

Set bankClass = New clsBankAcc

TransDate = GetSysFormatDate(txtDate2)

'Get the Details of the Cheque
gDbTrans.SqlStmt = "Select A.ChequeID,AccHeadID,AccID,Amount,ChequeNO," & _
    " ClearingBankID,Status,TransDate From ClearingTab A,ClearingTrans B" & _
    " Where A.ChequeId = B.ChequeId AND TransId = " & _
    "(SELECT Max(TransID) From ClearingTrans C Where C.ChequeId = A.ChequeId And " & _
        "(Status = " & wisPending & " OR Status = " & wisDiscount & " ));"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    MsgBox "Database ERROR - Clearing Tab", vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If


Dim I As Integer
Dim MaxI As Integer
I = 0

SbHeadID = GetIndexHeadID(GetResourceString(421))
CAHeadID = GetIndexHeadID(GetResourceString(422))

m_Loading = True

MaxI = grd.Rows - 1 'lstCheques.ListCount - 1
For I = 1 To MaxI
    If m_Status(I) < 0 Then
        PrevStatus = Not m_Status(I)
        ChequeId = grd.RowData(I)
        
        rst.MoveFirst
        rst.Find "ChequeID = " & ChequeId
        If PrevStatus = Status Then
            'msgbox "Unable to perform this operation"
            MsgBox GetResourceString(535), , wis_MESSAGE_TITLE
            GoTo ErrLine
        End If
        Amount = FormatField(rst("Amount"))
        AccHeadID = FormatField(rst("AccHeadID"))
        If PrevStatus = wisPending Then
            If rst("AccHeadID") = SbHeadID Then
                SbAmount = SbAmount + Amount
            Else
                CaAmount = CaAmount + Amount
            End If
        Else
            TotalDiscAmount = TotalDiscAmount + Amount
        End If
        
        With grd
            .Row = I '+ 1
            .Col = 0: ClrAmount = ClrAmount + Val(.Text)
            'If PrevStatus = wisDiscount Then
                .Col = 0
                If AccHeadID = SbHeadID Then
                    SBCharges = SBCharges + Val(.Text)
                    .Col = 1
                    SBCharges = SBCharges + Val(.Text)
                Else
                    CACharges = CACharges + Val(.Text)
                    .Col = 1
                    CACharges = CACharges + Val(.Text)
                End If
            'End If
        End With
    End If
Next


Dim rstTemp As Recordset

gDbTrans.BeginTrans

ClgHeadId = bankClass.GetHeadIDCreated(GetResourceString(168, 36), _
            LoadResourceStringS(168, 36), parReceivable, 0, wis_BankAccounts)
ClgIncomeID = bankClass.GetHeadIDCreated(GetResourceString(168, 366), _
            LoadResourceStringS(168, 366), parBankIncome, 0, wis_BankAccounts)

rst.MoveFirst
BankHeadID = FormatField(rst("ClearingBankID"))

'Now Transfer the Amount From Clering account to indvidual account
'Make transaction to contra trans table
Dim UserID As Long
UserID = gCurrUser.UserID
TransID = GetMaxContraTransID + 1
If TotalDiscAmount Then
    
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
                " VALUES (" & TransID & "," & ClgHeadId & "," & ClgHeadId & "," & _
                wContraDeposit & "," & TotalDiscAmount & "," & UserID & ")"
        
    'If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

'Now Make the Ledger head transaction
Dim AccTransClass As clsAccTrans
Set AccTransClass = New clsAccTrans

With AccTransClass
    If (CaAmount + SbAmount) > 0 Then
        If .AllTransHeadsAdd(BankHeadID, SbAmount + CaAmount, 0) <> Success Then GoTo ErrLine
        If SbAmount Then _
            If .AllTransHeadsAdd(SbHeadID, 0, SbAmount) <> Success Then GoTo ErrLine
        If CaAmount Then _
            If .AllTransHeadsAdd(CAHeadID, 0, CaAmount) <> Success Then GoTo ErrLine

        'If .SaveVouchers(Receipt, TransDate) <> Success Then GoTo ErrLine
        If .SaveVouchers(Journal, TransDate) <> Success Then GoTo ErrLine
        
    End If
    If TotalDiscAmount Then
        If .AllTransHeadsAdd(BankHeadID, TotalDiscAmount, 0) <> Success Then GoTo ErrLine
        If .AllTransHeadsAdd(ClgHeadId, 0, TotalDiscAmount) <> Success Then GoTo ErrLine
        'If .SaveVouchers(Receipt, TransDate) <> Success Then GoTo ErrLine
        If .SaveVouchers(Journal, TransDate) <> Success Then GoTo ErrLine
    End If
    If SBCharges Then
        If .AllTransHeadsAdd(ClgIncomeID, 0, SBCharges) <> Success Then GoTo ErrLine
        If .AllTransHeadsAdd(SbHeadID, SBCharges, 0) <> Success Then GoTo ErrLine
        If .SaveVouchers(Journal, TransDate) <> Success Then GoTo ErrLine
    End If
    If CACharges Then
        If .AllTransHeadsAdd(ClgIncomeID, 0, CACharges) <> Success Then GoTo ErrLine
        If .AllTransHeadsAdd(CAHeadID, CACharges, 0) <> Success Then GoTo ErrLine
        If .SaveVouchers(Journal, TransDate) <> Success Then GoTo ErrLine
    End If

End With
TransID = GetMaxContraTransID + 1
For I = 1 To MaxI
    If m_Status(I) >= 0 Then GoTo NextLine
    
    rst.MoveFirst
    
    ChequeId = Val(grd.RowData(I))
    rst.Find "ChequeID = " & ChequeId
    Debug.Assert Not rst.EOF
    PrevStatus = FormatField(rst("Status"))
    AccHeadID = FormatField(rst("AccHeadId"))
    Amount = FormatField(rst("Amount"))
    
    ModuleId = GetModuleIDFromHeadID(AccHeadID)
    TransDate = GetSysFormatDate(txtDate2.Text)
    
    If ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then Set ClassObj = New clsSBAcc
    If ModuleId = wis_CAAcc Then Set ClassObj = New ClsCAAcc
    If ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then Set ClassObj = New clsBkcc
    
    If PrevStatus = wisPending Then
        'TransID = GetMaxContraTransID + 1

        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
            " VALUES (" & TransID & "," & BankHeadID & "," & BankHeadID & "," & _
            wContraWithdraw & "," & Amount & "," & UserID & ")"
            
        If Not gDbTrans.SQLExecute Then GoTo ErrLine

        If ClassObj.DepositAmount(Val(FormatField(rst("AccId"))), Amount, _
                "Clg " & FormatField(rst("ChequeNo")), TransDate, " ", False, 10) = 0 Then GoTo ErrLine
    End If
    With grd
        .Row = I '+ 1
        .Col = 0: DiscAmount = Val(.Text)
        .Col = 1: DiscAmount = DiscAmount + Val(.Text)
    End With
    If DiscAmount Then
        'Dim Sbclass As clsSBAcc
        'Call Sbclass.WithdrawAmount(1, 100, "dd", "12/10/03")
        'TransID = GetMaxContraTransID + 1
        If ClassObj.WithdrawAmount(rst("AccId"), _
                DiscAmount, "Clearing Charges", TransDate, " ", False, TransID) = 0 Then GoTo ErrLine
        
        
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
            " VALUES (" & TransID & "," & ClgIncomeID & "," & ClgIncomeID & "," & _
            wContraDeposit & "," & DiscAmount & "," & UserID & ")"
            
        If Not gDbTrans.SQLExecute Then GoTo ErrLine

    End If
    
    'Get MAx Transaction ID
    'gDbTrans.SQLStmt = "Select Max(TransID) From ClearingTrans " & _
                        " WHERE ChequeID = " & ChequeId
    'If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                        TransID = FormatField(rstTemp(0)) + 1
    
    gDbTrans.SqlStmt = "Insert Into ClearingTrans " & _
                "(ChequeID,Status,TransDate,Transid) " & _
                " VALUES ( " & ChequeId & "," & Status & "," & _
                " #" & TransDate & "#," & TransID & ")"
            
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
NextLine:

Next

gDbTrans.CommitTrans

MsgBox GetResourceString(829), vbOKOnly, wis_MESSAGE_TITLE '"Transaction done successfully"

'Now Go For Deletion Of Corresponding In Records
'Dim Notes As clsNotes
'Set Notes = New clsNotes
'Call Notes.DeleteNotes(FormatField(Rst("Moduleid")), _
        FormatField(Rst("Accid")), FormatField(Rst("Noteid")))
SaveClearing = True
m_Loading = False


Exit Function

ErrLine:
    gDbTrans.RollBack
   If Err Then _
      MsgBox "Error In SaveTransTab - Clearing Module", vbCritical, wis_MESSAGE_TITLE
    m_Loading = False

End Function

Private Function SaveDiscount() As Boolean

'Check the Date Validation
If Not DateValidate(txtDate2.Text, "/", True) Then
   MsgBox GetResourceString(501), vbExclamation, wis_MESSAGE_TITLE
   ActivateTextBox txtDate2
   Exit Function
End If

Dim rst As Recordset
Dim ModuleId As wisModules
Dim ClgHeadId As Long
Dim AccHeadID As Long
Dim BankHeadID As Long
Dim ClassObj As Object
Dim TransDate As Date
Dim ChequeId As Long
Dim bankClass As clsBankAcc


Dim TotaClrAmount As Currency
Dim TotalDiscAmount As Currency
Dim TransID As Integer

Dim SbAmount As Currency
Dim SbHeadID As Long
Dim CaAmount As Currency
Dim CAHeadID As Long

Dim DiscAmount As Currency
Dim ClrAmount As Currency
Dim Amount As Currency

Set bankClass = New clsBankAcc

'Get the Details of the Cheque
gDbTrans.SqlStmt = "Select A.ChequeID,AccHeadID,AccID,Amount,ChequeNo," & _
    " ClearingBankID,Status,TransDate From ClearingTab A,ClearingTrans B" & _
    " Where A.ChequeId = B.ChequeId AND TransId = " & _
    "(SELECT Max(TransID) From ClearingTrans C Where C.ChequeId = A.ChequeId And " & _
        "(Status = " & wisPending & " OR Status = " & wisDiscount & " ));"

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    MsgBox "Database ERROR - Clearing Tab", vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If


Dim I As Integer
Dim MaxI As Integer
I = 0

SbHeadID = GetIndexHeadID(GetResourceString(421))
CAHeadID = GetIndexHeadID(GetResourceString(422))


MaxI = grd.Rows - 1 'lstCheques.ListCount - 1
For I = 1 To MaxI
    
    If m_Status(I) > 0 Then
        'ChequeId = Val(lstCheques.ItemData(I))
        ChequeId = grd.RowData(I)
        rst.MoveFirst
        rst.Find "ChequeID = " & ChequeId
        Debug.Assert Not rst.EOF
        If (Not m_Status(I)) = wisDiscount Then
            'msgbox "Unable to perform this operation"
            MsgBox GetResourceString(535), , wis_MESSAGE_TITLE
            GoTo ErrLine
        End If
        Amount = FormatField(rst("Amount"))
        If rst("AccHeadID") = SbHeadID Then
            SbAmount = SbAmount + Amount
        Else
            CaAmount = CaAmount + Amount
        End If
    End If
Next


TransID = 100
Dim rstTemp As Recordset
TransID = GetMaxContraTransID + 1

gDbTrans.BeginTrans

ClgHeadId = bankClass.GetHeadIDCreated(GetResourceString(168, 36), _
              LoadResourceStringS(168, 36), parReceivable, 0, wis_BankAccounts)
rst.MoveFirst
BankHeadID = FormatField(rst("ClearingBankID"))

'Now Transfer the Amount From Clering account to indvidual account
'Make transaction to contra trans table
Dim UserID As Long
UserID = gCurrUser.UserID
    
'Now Make the Transaction To the Ledger heads
If Not bankClass.UpdateContraTrans(ClgHeadId, SbHeadID, SbAmount, TransDate) Then GoTo ErrLine
If Not bankClass.UpdateContraTrans(ClgHeadId, CAHeadID, CaAmount, TransDate) Then GoTo ErrLine
    
gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
        "(ContraID,AccHeadID,AccId,TransType,Amount,UserID) " & _
        " VALUES (" & TransID & "," & ClgHeadId & "," & ClgHeadId & "," & _
        wContraWithdraw & "," & SbAmount + CaAmount & "," & UserID & ")"
        
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo ErrLine
End If

For I = 1 To MaxI
    If Not m_Status(I) >= 0 Then GoTo NextLine
    
    rst.MoveFirst
    'ChequeId = Val(lstCheques.ItemData(I))
    ChequeId = grd.RowData(I)
    rst.Find "ChequeID = " & ChequeId
    
    'PrevStatus = Not m_Status(I)
    BankHeadID = FormatField(rst("ClearingBankId"))
    AccHeadID = FormatField(rst("AccHeadId"))

    ModuleId = GetModuleIDFromHeadID(AccHeadID)
    TransDate = GetSysFormatDate(txtDate2.Text)
    
    If ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then Set ClassObj = New clsSBAcc
    If ModuleId = wis_CAAcc Then Set ClassObj = New ClsCAAcc
    
    If ClassObj.DepositAmount(Val(FormatField(rst("AccId"))), _
       FormatField(rst("Amount")), _
       "Discount " & FormatField(rst("ChequeNo")), TransDate) = 0 Then GoTo ErrLine
        
    'Upadate this Cheque in the Clearing Tab
    
    'Get MAx Transaction ID
    gDbTrans.SqlStmt = "Select Max(TransID) From ClearingTrans " & _
                            " WHERE ChequeID = " & ChequeId
    TransID = 0
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then TransID = FormatField(rstTemp(0))
    TransID = TransID + 1
    
    gDbTrans.SqlStmt = "Insert Into ClearingTrans " & _
                "(ChequeID,Status,TransDate,Transid) " & _
                " VALUES (" & ChequeId & "," & wisDiscount & "," & _
                " #" & TransDate & "#," & TransID & ")"
            
    If Not gDbTrans.SQLExecute Then
       gDbTrans.RollBack
       GoTo ErrLine
    End If
    
NextLine:

Next

gDbTrans.CommitTrans

'Now Go For Deletion Of Corresponding In Records
Dim Notes As clsNotes
Set Notes = New clsNotes
'Call Notes.DeleteNotes(FormatField(Rst("Moduleid")), FormatField(Rst("Accid")), FormatField(Rst("Noteid")))
SaveDiscount = True

Exit Function

ErrLine:
    gDbTrans.RollBack
   If Err Then _
      MsgBox "Error In SaveTransTab - Clearing Module", vbCritical, wis_MESSAGE_TITLE
   
End Function

Private Sub SetKannadaCaption()

    Call SetFontToControls(Me)
    'Set kannada captions for all the  controls
    'Now for  the tabs
    TabStrip1.Tabs(1).Caption = GetResourceString(167)
    TabStrip1.Tabs(2).Caption = GetResourceString(168)
    TabStrip1.Tabs(3).Caption = GetResourceString(283) & GetResourceString(92)
    cmdCancel.Caption = GetResourceString(1)
    
    'Set the captions for the controls in the schedule for the clearing
    fraClearing(1).Caption = GetResourceString(169)
    lblDate1 = GetResourceString(37)
    lblAccountType = GetResourceString(36) & " " & _
                GetResourceString(35)
    lblAccountNo = GetResourceString(36) & " " & _
                GetResourceString(60)
'    lblDeposit.Caption = GetResourceString(241)
    lblBank.Caption = GetResourceString(418) & " " & _
            GetResourceString(295)
    lblInstrNo.Caption = GetResourceString(41)
    lblName = GetResourceString(205, 36)
    lblAmount.Caption = GetResourceString(40)
    cmdAccept.Caption = GetResourceString(4)
    lblClearBank = GetResourceString(168, 418)
    lblCheque = GetResourceString(275)
    lblRemarks = GetResourceString(261)
    chkDiscount.Caption = GetResourceString(293)
    
    'set the captions for the controls in the update tab
    fraClearing(2).Caption = GetResourceString(469)
    lblDate2.Caption = GetResourceString(37)
    cmdSelectAll.Caption = GetResourceString(25)
    cmdDeSelectAll.Caption = GetResourceString(26)
    cmdClear.Caption = GetResourceString(8)
    cmdDelete.Caption = GetResourceString(14)
    cmdBounce.Caption = GetResourceString(470)
    cmdDiscount.Caption = GetResourceString(293)
    
    'Set the captions of the reports tab
    fraClearing(3).Caption = GetResourceString(283) & GetResourceString(92)
    optReports(0).Caption = GetResourceString(471)
    optReports(1).Caption = GetResourceString(469)
    optReports(2).Caption = GetResourceString(472)
    optReports(3).Caption = GetResourceString(473)
    optReports(4).Caption = GetResourceString(293) & " " & _
            GetResourceString(275)
    lblFrom.Caption = GetResourceString(109)
    lblTo.Caption = GetResourceString(110)
    cmdView.Caption = GetResourceString(13)

End Sub


Private Sub cmdAccept_Click()
Dim AccNo As Long
Dim ModuleId As wisModules
'First Validate The Enteries
   
    'Check the Date Validation
    If Not DateValidate(txtDate1.Text, "/", True) Then
        MsgBox GetResourceString(501), vbExclamation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate1
        Exit Sub
    End If
    
    If cmbAcType.ListIndex = -1 Then
        'MsgBox "Invalid Account Type Specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(673), vbInformation, wis_MESSAGE_TITLE
        cmbAcType.SetFocus
        Exit Sub
    End If
    
    ModuleId = Me.cmbAcType.ItemData(Me.cmbAcType.ListIndex)
    If cmbBank.ListIndex = -1 Then
        MsgBox "Select the Clearing Bank", vbInformation, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(673), vbInformation, wis_MESSAGE_TITLE
        cmbBank.SetFocus
        Exit Sub
    End If
    
    If Trim(txtAccNo.Text) = "" Or Val(txtAccNo.Tag) = 0 Then
        'MsgBox "Invalid Account Number Specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtAccNo
        Exit Sub
    End If
    
    AccNo = Val(txtAccNo.Text)
    
    If Trim(txtChqNo.Text) = "" Then
        'MsgBox "Invalid Cheque Number Specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(511), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtChqNo
        Exit Sub
    End If
    
    If txtAmount = 0 Then
        'MsgBox "Invalid Amount Specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtAmount
        Exit Sub
    End If
    
    'if all the Above Entries Are Correct then Call Function To Save Val INto Tab and asign the Values To It
     If Not SaveDetails Then  'Saves The Details IntoNoteTab and Clearing Tab
       
       'MsgBox "Unable To Save Details", vbCritical, wis_MESSAGE_TITLE
       MsgBox GetResourceString(533), vbCritical, wis_MESSAGE_TITLE
    Else
       MsgBox "Details Saved Successfully ", vbInformation, wis_MESSAGE_TITLE
       'MsgBox GetResourceString(530), vbInformation, wis_MESSAGE_TITLE
       Call ResetUserInterface
    End If
    Call LoadNotes
    
ErrLine:

End Sub

Private Sub cmdBounce_Click()

If SaveIntoTransTab(wisBounced) Then Call LoadNotes

End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Public Function LoadModuleId() As Boolean
Dim rstDeposit As Recordset
Dim rstLoan As Recordset

gDbTrans.SqlStmt = "Select * FROM DepositName"
Call gDbTrans.Fetch(rstDeposit, adOpenStatic)
gDbTrans.SqlStmt = "Select * FROM LoanScheme"
Call gDbTrans.Fetch(rstLoan, adOpenStatic)

'Dim ModuleId As wisModules
Dim AccHeadID As Long
Dim ClsBank As clsBankAcc
Dim AccName As String

'Set ClsBank = New clsBankAcc

With cmbAcType
    'Withdrawing of the amount willbe only from the
    'account type where cheque facility is given
    'Such type of account are only two(three) types
    'They are
    '1) Saving Bank Account
    '2) Current Account
    '3) Od Account 'Presetly no such accounts are in our s/w
    
    .Clear
    'ModuleId = wis_SBAcc
    AccName = GetResourceString(421) '"SB AccOUnt
    
    AccHeadID = GetIndexHeadID(AccName)
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID 'ModuleId
    End If
    
    'ModuleId = wis_CAAcc
    AccName = GetResourceString(422) '"CA AccOUnt
    AccHeadID = GetIndexHeadID(GetResourceString(422))
    If AccHeadID Then
        .AddItem AccName
        .ItemData(.newIndex) = AccHeadID 'ModuleId
    End If
End With

LoadModuleId = True
Exit Function


ErrLine:
LoadModuleId = False
End Function


Private Sub cmdClear_Click()

'If SaveIntoTransTab(wisCleared) Then Call LoadNotes
If SaveClearing Then Call LoadNotes

End Sub

Private Sub cmdDate_Click()

With Calendar
    .Left = Left + fraClearing(1).Left + cmdDate.Left + cmdDate.Width
    .Top = Top + fraClearing(1).Top + cmdDate.Top - .Height / 2
    If DateValidate(txtDate1, "/", True) Then .selDate = txtDate1 _
            Else .selDate = gStrDate
    .Show vbModal
    txtDate1 = .selDate
End With

End Sub

Private Sub cmdDelete_Click()

Dim Note As New clsNotes
Dim rst As Recordset
Dim Found As Boolean
'Check   Whether Any Item Has Been Selected or Not
Dim I As Integer
Dim MaxI As Integer
MaxI = grd.Rows - 1 ' lstCheques.ListCount - 1
I = 0

For I = 1 To MaxI
    
    If m_Status(I) < 0 And Not m_Status(I) = wisPending Then
         If Not Found Then
             If MsgBox(GetResourceString(789), _
                         vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
             'If MsgBox("Are you sure you want to delete the transaction", _
              '           vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
             Found = True
         End If
         
         Dim ModuleId As wisModules
         Dim AccId As Long
         Dim NoteId As Long
         
         'Get The Neccessary Values
         gDbTrans.SqlStmt = "Select * From ClearingTrans Where " & _
               " TransId = " & Val(grd.RowData(I))
         If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub
         ModuleId = FormatField(rst("Moduleid"))
         AccId = CLng(FormatField(rst("Accid")))
         NoteId = CLng(FormatField(rst("Noteid")))
         
         'Go for Deletion
          gDbTrans.SqlStmt = "Delete * From ClearingTab " & _
                " Where TransId = " & grd.RowData(I) 'Val(lstCheques.ItemData(I))
          gDbTrans.BeginTrans
          If Not gDbTrans.SQLExecute Then
             gDbTrans.RollBack
             Exit Sub
          End If
          gDbTrans.CommitTrans
          Call Note.DeleteNotes(ModuleId, AccId, NoteId)
    End If
Next I

   Call LoadNotes
End Sub

Private Sub cmdDeSelectAll_Click()
Dim I As Integer
Dim MaxI As Integer
With grd
    MaxI = .Rows - 1
    .Col = 0
    For I = 1 To MaxI
        If m_Status(I) >= 0 Then
        m_Status(I) = Not m_Status(I)
        .Row = I
        Set .CellPicture = LoadResPicture(143 + IIf(m_Status(I) > 0, 0, 1), vbResIcon)
        End If
    Next I
End With

End Sub

Private Sub cmdDiscount_Click()

If SaveDiscount Then Call LoadNotes

End Sub

Private Sub cmdFromDate_Click()
With Calendar
    .Left = Me.Left + fraClearing(1).Left + cmdFromDate.Left - .Width / 2
    .Top = Me.Top + fraClearing(1).Top + cmdFromDate.Top
    If DateValidate(txtFromDate.Text, "/", True) Then
        .selDate = txtFromDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    If .selDate <> "" Then txtFromDate.Text = .selDate
End With

End Sub

Private Sub cmdLookup_Click()

If cmbAcType.ListIndex < 0 Then GoTo ExitLine

Dim ModuleId As wisModules
Dim AccHeadID As Long
Dim NameStr As String

AccHeadID = cmbAcType.ItemData(cmbAcType.ListIndex)

Dim RstAcc As Recordset

'Now Check Whether He Want Search   any particular name
'NameStr = InputBox("Enter customer name , You want search", "Name Search")
NameStr = InputBox(GetResourceString(785), "Name Search")

Set RstAcc = GetAccRecordSet(AccHeadID, , NameStr)
If RstAcc Is Nothing Then Exit Sub
 
 If LookUpDialog Is Nothing Then Set LookUpDialog = New frmLookUp
 m_SelectedAccount = ""
 With LookUpDialog
      .m_SelItem = ""
      .Title = "Select Deposit "
      ' Fill the data to report dialog.
      If Not FillView(.lvwReport, RstAcc, True) Then
          'MsgBox "Error filling the customer details.", vbCritical
          MsgBox "Error filling the customer details.", vbCritical
          GoTo ExitLine
      End If
      Screen.MousePointer = vbDefault
      'Display the dialog.
      .Show vbModal
End With


txtAccNo.Text = m_SelectedAccount
Set RstAcc = GetAccRecordSet(AccHeadID, txtAccNo)

If RstAcc Is Nothing Then Exit Sub
txtAccNo.Tag = FormatField(RstAcc("ID"))
txtName = FormatField(RstAcc("CustName"))


ExitLine:
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSelectAll_Click()
Dim I As Integer
Dim MaxI As Integer

With grd
    MaxI = .Rows - 1
    .Col = 0
    For I = 1 To MaxI
        If m_Status(I) < 0 Then
        m_Status(I) = Not m_Status(I)
        .Row = I
        Set .CellPicture = LoadResPicture(143 + IIf(m_Status(I) > 0, 0, 1), vbResIcon)
        End If
    Next I
End With

End Sub


Private Sub cmdToDate_Click()
With Calendar
    .Left = Me.Left + fraClearing(1).Left + cmdToDate.Left - .Width / 2
    .Top = Me.Top + fraClearing(1).Top + cmdToDate.Top
    If DateValidate(txtToDate.Text, "/", True) Then
        .selDate = txtToDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    If .selDate <> "" Then txtToDate.Text = .selDate
End With

End Sub

Private Sub cmdView_Click()
'Make one round of date checks
If txtFromDate.Enabled And Trim$(txtFromDate.Text) <> "" And Not DateValidate(txtFromDate.Text, "/", True) Then
    'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
    ActivateTextBox txtFromDate
    Exit Sub
End If
If txtToDate.Enabled And Trim$(txtToDate.Text) <> "" And Not DateValidate(txtToDate.Text, "/", True) Then
     'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
    ActivateTextBox txtToDate
    Exit Sub
End If

Dim ReportType As wis_ClearingReports

If optReports(0).Value Then
    ReportType = repChqSubmitted
ElseIf optReports(1).Value Then
    ReportType = repChqCleared
ElseIf optReports(2).Value Then
    ReportType = repChqNotCleared
ElseIf optReports(3).Value Then
    ReportType = repChqPending
ElseIf optReports(4).Value Then
    ReportType = repChqDiscount
End If
RaiseEvent ShowReport(ReportType, _
        txtFromDate, txtToDate)

'frmCLReport.Show ' vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
Dim I As Byte
If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
    With TabStrip1
        I = .SelectedItem.Index
        If Shift = 2 Then
            I = I + 1
            If I > .Tabs.count Then I = 1
        Else
            I = I - 1
            If I = 0 Then I = .Tabs.count
        End If
        
        .Tabs(I).Selected = True
    End With
End If

End Sub

Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2

'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
txtDate1.Text = gStrDate
txtDate2.Text = gStrDate

'Set kannada caption
Call SetKannadaCaption
'lstCheques.Font.Size = 12
' Set the border style for pics.
Dim I As Integer

TabStrip1.Tabs(1).Selected = True
'TODO...
    ' 1. Load the account types into the combobox.
If Not LoadModuleId Then
    'MsgBox "Unable To Load  Account types", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(757), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Call LoadNotes

'Me.lstCheques.MultiSelect = 2 'Extended MultiSelect
cmbBank.Clear
Call LoadLedgersToCombo(cmbBank, parBankAccount)
gDbTrans.SqlStmt = "Select * From Install Where KeyData = 'ClearingBankId'"

Dim rst As Recordset
Dim RstBank As Recordset

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    gDbTrans.SqlStmt = "Select HeadName,HeadID " & _
                   " FROM Heads WHERE ParentID =  " & parBankAccount & _
                   " ORDER BY HeadName"
    Call gDbTrans.Fetch(RstBank, adOpenDynamic)
    With cmbBank
        .Clear
        While Not rst.EOF
            RstBank.MoveFirst
            RstBank.Find "HeadID = " & Val(rst("ValueData"))
            If Not RstBank.EOF Then
                .AddItem RstBank("HeadName")
                .ItemData(.newIndex) = RstBank("HeadID")
            End If
            rst.MoveNext
        Wend
    End With
End If

If cmbBank.ListCount = 1 Then cmbBank.ListIndex = 0

If gOnLine Then
    txtDate1.Locked = True
    cmdDate.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
RaiseEvent WindowClosed
gWindowHandle = 0

End Sub

Private Sub grd_Click()
With grd
    If .Row < .FixedRows Then Exit Sub
    If .Col = 5 Or .Col = 6 Then Call grd_EnterCell
    
    If .Col <> 0 Then Exit Sub
    
    If m_Status(.Row) <> wisCleared Then
        Set .CellPicture = LoadResPicture(143, vbResIcon)
        m_Status(.Row) = wisCleared
    Else
        Set .CellPicture = LoadResPicture(144, vbResIcon)
        m_Status(.Row) = wisPending
    End If
    
End With

End Sub

Private Sub grd_EnterCell()

If m_Loading Then Exit Sub

With grd
    If .Col < 4 Then Exit Sub
    If .Row < .FixedRows Then Exit Sub
    
    If .CellBackColor = &H8000000F Then Exit Sub
    'If lstCheques.ListIndex <> .Row - 1 Then Exit Sub
    Dim I As Integer
    I = grd.Row
    txtgrd.Text = .Text
    txtgrd.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
    txtgrd.Visible = .RowIsVisible(.Row)
    txtgrd.ZOrder 0
End With
   
   On Error Resume Next
   ActivateTextBox txtgrd
   Err.Clear
End Sub

Private Sub grd_LeaveCell()
    
If m_Loading Then Exit Sub
If grd.Col < 4 Then Exit Sub
If grd.CellBackColor = &H8000000F Then Exit Sub

With txtgrd
   grd.Text = .Text
   .Visible = False
   .ZOrder 1
   .Text = ""
End With

End Sub

Private Sub grd_Scroll()

With grd
    txtgrd.Visible = .RowIsVisible(.Row)
    If txtgrd.Visible Then txtgrd.Visible = .RowIsVisible(.Row)
    If Not txtgrd.Visible Then Exit Sub
    txtgrd.Move .Left + .ColPos(.Col), .Top + .RowPos(.Row) ', grd.CellWidth, grd.CellHeight
End With

End Sub


Private Sub txtgrd_GotFocus()
txtgrd.SelStart = 0
txtgrd.SelLength = Len(txtgrd)
End Sub

Private Sub txtgrd_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

With grd
    
    If KeyCode = 37 Then 'Press Left Arrow
        If txtgrd.SelStart = 0 And .Col > .FixedCols Then .Col = .Col - 1
    
    ElseIf KeyCode = 39 Then 'Press Right Arrow
        If txtgrd.SelStart = Len(txtgrd.Text) And .Col < .Cols - 1 Then .Col = .Col + 1
    
    ElseIf KeyCode = 38 Then  'Press UpArrow
        If .Row > .FixedRows Then .Row = .Row - 1
    
    ElseIf KeyCode = 40 Then ' Press Doun Arroow
        If .Row < .Rows - 1 Then .Row = .Row + 1
    
    ElseIf KeyCode = 33 Then  'Press PageUp
        If .Row > 0 Then .Row = .Row - 1
        
    ElseIf KeyCode = 34 Then ' Press PageDown
        If .Row < .Rows - 1 Then .Row = .Row + 1
    
    End If
    
End With

End Sub

Private Sub txtgrd_LostFocus()

    grd.Text = txtgrd.Text
    txtgrd.Visible = False
    'grd.SetFocus
End Sub

Private Sub LookUpDialog_SaveClick(strSelection As String)
m_SelectedAccount = strSelection

End Sub

Private Sub LookUpDialog_SelectClick(strSelection As String)
m_SelectedAccount = strSelection
End Sub

Private Sub LookUpDialog_SubItems(strSubItem() As String)
Dim count As Integer
ReDim m_SubItems(UBound(strSubItem))
For count = LBound(strSubItem) To UBound(strSubItem)
   m_SubItems(count) = strSubItem(count)
Next

End Sub

Private Sub TabStrip1_Click()

fraClearing(TabStrip1.SelectedItem.Index).ZOrder 0
If TabStrip1.SelectedItem.Index = 3 Then optReports(0).Value = True

End Sub

Private Sub txtAccNo_LostFocus()

If cmbAcType.ListIndex < 0 Then Exit Sub
If Len(Trim$(txtAccNo)) = 0 Then Exit Sub

Dim AccHeadID As Long

AccHeadID = cmbAcType.ItemData(cmbAcType.ListIndex)

Dim RstAcc As Recordset
Set RstAcc = GetAccRecordSet(AccHeadID, txtAccNo)

If RstAcc Is Nothing Then Exit Sub
txtAccNo.Tag = FormatField(RstAcc("ID"))
txtName = FormatField(RstAcc("CustName"))


End Sub


