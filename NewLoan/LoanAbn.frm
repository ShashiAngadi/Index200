VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmLoanAbn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABN & EP Details"
   ClientHeight    =   6975
   ClientLeft      =   1530
   ClientTop       =   1560
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   400
      Left            =   150
      TabIndex        =   17
      Top             =   6510
      Width           =   1215
   End
   Begin VB.Frame fraLoan 
      Height          =   3225
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6945
      Begin WIS_Currency_Text_Box.CurrText txtInterest 
         Height          =   345
         Left            =   1890
         TabIndex        =   16
         Top             =   2730
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPenal 
         Height          =   345
         Left            =   5160
         TabIndex        =   19
         Top             =   2730
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtDueDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5160
         TabIndex        =   12
         Top             =   1794
         Width           =   1395
      End
      Begin VB.Label lblDueDate 
         Caption         =   "&Date :"
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         Top             =   1794
         Width           =   1515
      End
      Begin VB.Label lblInterest 
         Caption         =   "&Interest till date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2850
         Width           =   1695
      End
      Begin VB.Label lblPenal 
         Caption         =   "&Penal Interest "
         Height          =   225
         Left            =   3720
         TabIndex        =   18
         Top             =   2730
         Width           =   1485
      End
      Begin VB.Label lblCustID 
         Caption         =   "Member Number"
         Height          =   270
         Left            =   3480
         TabIndex        =   3
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label lblName 
         Caption         =   "Loan Holder Name :"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   858
         Width           =   1545
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "Loan Account No :"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label lblLoanName 
         Caption         =   "Loan  Name :"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   1326
         Width           =   1545
      End
      Begin VB.Label lblLoanBalance 
         Caption         =   "Loan Balance :"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lblLoanDate 
         Caption         =   "Loan Date :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1794
         Width           =   1425
      End
      Begin VB.Label txtLoanName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1890
         TabIndex        =   8
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1890
         TabIndex        =   6
         Top             =   855
         Width           =   4695
      End
      Begin VB.Label txtLoanAccNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   2
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label txtCustID 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5160
         TabIndex        =   4
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label txtLoanDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   10
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label txtLoanBalance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   14
         Top             =   2280
         Width           =   1395
      End
   End
   Begin VB.Frame fraInstall 
      Height          =   3255
      Left            =   120
      TabIndex        =   42
      Top             =   3180
      Width           =   6945
      Begin VB.TextBox txtEpNo 
         Height          =   345
         Left            =   5280
         TabIndex        =   35
         Top             =   1840
         Width           =   1395
      End
      Begin VB.TextBox txtAbnNo 
         Height          =   345
         Left            =   5160
         TabIndex        =   25
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox txtEPDate 
         Height          =   345
         Left            =   1890
         TabIndex        =   32
         Top             =   1840
         Width           =   1395
      End
      Begin VB.CommandButton cmdEpDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3270
         TabIndex        =   33
         Top             =   1840
         Width           =   315
      End
      Begin VB.TextBox txtEPRemark 
         Height          =   345
         Left            =   1890
         TabIndex        =   41
         Top             =   2760
         Width           =   4785
      End
      Begin VB.CheckBox chkEP 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit to Loan Account"
         Height          =   315
         Left            =   3720
         TabIndex        =   39
         Top             =   2300
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.TextBox txtAbnDate 
         Height          =   345
         Left            =   1890
         TabIndex        =   22
         Top             =   300
         Width           =   1395
      End
      Begin VB.CommandButton cmdAbndate 
         Caption         =   "..."
         Height          =   315
         Left            =   3240
         TabIndex        =   23
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtAbnRemark 
         Height          =   345
         Left            =   1890
         TabIndex        =   30
         Top             =   1220
         Width           =   4785
      End
      Begin VB.CheckBox chkAbn 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit to Loan Account"
         Height          =   315
         Left            =   3720
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   2805
      End
      Begin WIS_Currency_Text_Box.CurrText txtAbnFee 
         Height          =   345
         Left            =   1890
         TabIndex        =   27
         Top             =   765
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtEPFee 
         Height          =   345
         Left            =   1890
         TabIndex        =   37
         Top             =   2295
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label Label1 
         Caption         =   "Ep &No"
         Height          =   315
         Left            =   3720
         TabIndex        =   34
         Top             =   1840
         Width           =   1335
      End
      Begin VB.Label lblAbnNo 
         Caption         =   "&ABN No"
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblEPDate 
         Caption         =   "&Date :"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1840
         Width           =   1335
      End
      Begin VB.Label lblEpRemark 
         Caption         =   "&Remarks"
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   2760
         Width           =   1395
      End
      Begin VB.Label lblEPFee 
         Caption         =   "&EP Fee"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   2300
         Width           =   1515
      End
      Begin VB.Label lblABNDate 
         Caption         =   "&Date :"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblABNRemark 
         Caption         =   "&Remarks"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1220
         Width           =   1395
      End
      Begin VB.Label lblABNFee 
         Caption         =   "&ABN Fee"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   760
         Width           =   1515
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   6660
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4650
      TabIndex        =   28
      Top             =   6510
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5940
      TabIndex        =   38
      Top             =   6510
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoanAbn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_LoanID As Long
Private m_SchemeName As String
Private m_SchemeId As Integer
Private m_IntTable As String
Private m_TransTable As String
Private m_MasterTable As String
Private m_Kcc As Boolean

'Private m_AbnDate As Date
'Private m_EpDate As Date
'Private m_AbnAmount As Currency
'Private m_EpAmount As Currency
Private m_FormLoaded As Boolean

Public Event WindowClosed()

Public Property Let LoanAccountID(NewValue As Long)
   
    m_LoanID = NewValue
    If m_LoanID = 0 Then Exit Property
    
    
    Dim IsBkcc As Boolean
    m_Kcc = IsBkcc
    If IsBkcc Then
        m_IntTable = "BKCCIntTrans"
        m_TransTable = "BKCCTrans"
        m_MasterTable = "BKCCMaster"
    Else
        m_IntTable = "LoanIntTrans"
        m_TransTable = "LoanTrans"
        m_MasterTable = "LoanMaster"
    End If

If m_FormLoaded Then Call UpdateDetails

End Property


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblLoanAccNo = GetResourceString(80, 36, 60)
lblCustID = GetResourceString(49, 60) 'Customer ID
lblName = GetResourceString(35) 'Name
lblLoanName = GetResourceString(214) 'loanScheme
lblLoanBalance = GetResourceString(80, 42) 'Loan Balance
lblLoanDate = GetResourceString(80, 37) 'Loan Date
'lblDueDate = GetResourceString(209) 'Loan due Date
lblDueDate = GetResourceString(84, 37) 'Loan due Date
lblInterest = GetResourceString(344)
lblPenal = GetResourceString(345) '"PenalInterest

lblABNDate = GetResourceString(372, 37) 'Date
lblABNFee = GetResourceString(372, 191) 'Date
lblABNRemark = GetResourceString(372, 261) 'Remarks
chkAbn.Caption = GetResourceString(80, 36, 272)
lblEPDate = GetResourceString(373, 37) 'Date
lblEPFee = GetResourceString(373, 191) 'Date
lblEpRemark = GetResourceString(373, 261) 'Remarks
chkEP.Caption = GetResourceString(80, 36, 272)

cmdCancel.Caption = GetResourceString(2)  'Cancel
cmdOk.Caption = GetResourceString(4)      'OK
'errorr nnn
'on clicking the overdue loans it give run time error

cmdUndo.Caption = GetResourceString(19)     'Undo (Delete)
End Sub
Private Function UpdateAbnEp1() As Boolean

'Now Check whether to or Insert the Abn ep
Dim rst As Recordset
Dim DBOperation As wis_DBOperation

Dim AbnDate As String
Dim EPDate As String

gDbTrans.SqlStmt = "SELECT * From LoanAbnEp " & _
            " Where LoanID = " & m_LoanID & " And BKCC = 0"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then DBOperation = Update
Set rst = Nothing

AbnDate = GetSysFormatDate(txtAbnDate)
If txtEPDate.Text <> "" Then EPDate = GetSysFormatDate(txtEPDate)

If DBOperation = Update Then
    gDbTrans.SqlStmt = "UPDate LOanAbnEp Set " & _
        " abndate = #" & AbnDate & "#," & _
        " AbnAmount = " & txtAbnFee & " ," & _
        " AbnDesc = " & AddQuotes(txtAbnRemark, True) & " ," & _
        " AbnNo = " & AddQuotes(txtAbnNo, True) & " ," & _
        " EPDate = " & IIf(txtEPDate = "", " Null ", "#" & EPDate & "#") & "," & _
        " EpAmount = " & txtEPFee & _
        " EPDesc = " & AddQuotes(txtEPRemark, True) & " ," & _
        " EPDesc = " & AddQuotes(txtEpNo, True) & " ," & _
        " Where Loanid = " & m_LoanID & _
        " ANd Bkcc = 0 "
        
Else
    gDbTrans.SqlStmt = "Insert Into LoanAbnEp " & _
                " (LoanID,BKCC,AbnDate,AbnAmount,AbnDesc,abnNo, " & _
                " EpDate,EpAmount,EpDesc, EpNo ) VALUES " & _
        "(" & m_LoanID & ", 0 , #" & AbnDate & "#, " & txtAbnFee & " ," & _
        AddQuotes(txtAbnRemark, True) & "," & AddQuotes(txtAbnNo, True) & "," & _
        IIf(txtEPDate = "", " Null ", "#" & EPDate & "#") & "," & txtEPFee & "," & _
        AddQuotes(txtEPRemark, True) & "," & AddQuotes(txtEpNo, True) & ")"
End If


gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If
gDbTrans.CommitTrans

UpdateAbnEp1 = True

End Function

Private Function UpdateAbnEp() As Boolean
On Error GoTo Err_line

'Now Check whether to or Insert the Abn ep
Dim rst As Recordset
Dim DBOperation As wis_DBOperation

Dim AbnDate As Date
Dim EPDate As Date

Dim PrevAbnAmount As Currency
Dim PrevEPAmount As Currency

Dim Balance As Currency
Dim tmpAmount As Currency

Dim TransID As Long
Dim ContraID As Long
Dim UserID As Integer

Dim AccHeadID As Long
Dim DueHeadID As Long

Dim bankClass As clsBankAcc

gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
                " Where SchemeID = (Select SchemeID From LoanMaster" & _
                        " Where LoanID = " & m_LoanID & " )"
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
AccHeadID = GetHeadID(FormatField(rst("SchemeName")), parMemberLoan)

Dim AddToAccount As Boolean
Dim InTrans As Boolean

'PrevAbnDate = "1/1/1000"
'PrevEPDate = "1/1/1000"

gDbTrans.SqlStmt = "SELECT * From LoanAbnEp " & _
            "Where LoanID = " & m_LoanID & " And Bkcc = 0"
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    DBOperation = Update
    'If Not IsNull(Rst("abnDate")) Then PrevAbnDate = Rst("abnDate")
    'If Not IsNull(Rst("EpDate")) Then PrevEPDate = Rst("EpDate")
    'PrevAbnAmount = FormatField(Rst("AbnAmount"))
    'PrevEPAmount = FormatField(Rst("EpAmount"))
    Set rst = Nothing
    'Here Check For The Abn or Ep  Entry in the AmountReceivable

End If

'if me.chkAbn
AbnDate = GetSysFormatDate(txtAbnDate)
If txtEPDate.Text <> "" Then EPDate = GetSysFormatDate(txtEPDate)

    'Here insert the same amount to Receivable table
    'Now Get the pending Balance From the Previous account
Balance = 0
gDbTrans.SqlStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & AccHeadID & _
        " ANd AccId = " & m_LoanID
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    rst.MoveLast
    Balance = FormatField(rst("Balance"))
    Set rst = Nothing
End If

TransID = GetLoanMaxTransID(m_LoanID)
UserID = gCurrUser.UserID
'Begin the transaction
InTrans = gDbTrans.BeginTrans
Set bankClass = New clsBankAcc
            
If DBOperation = Update Then
    gDbTrans.SqlStmt = "UPDate LoanAbnEp Set " & _
                " Abndate = #" & AbnDate & "#," & _
                " AbnAmount = " & txtAbnFee & " ," & _
                " AbnDesc = " & AddQuotes(txtAbnRemark, True) & " ," & _
                " AbnNo = " & AddQuotes(txtAbnNo, True) & " ," & _
                " EPDate = " & IIf(txtEPDate = "", "Null", "#" & EPDate & "#") & "," & _
                " EpAmount = " & txtEPFee & _
                " EPDesc = " & AddQuotes(txtEPRemark, True) & " ," & _
                " EPDesc = " & AddQuotes(txtEpNo, True) & " ," & _
                " Where Loanid = " & m_LoanID & _
                " And Bkcc = 0"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    If chkEP.Value = vbChecked Then
        'Check for the Last transaction Date
        If EPDate > GetLoanLastTransDate(m_LoanID) Then GoTo Err_line
        'Now Get the loan Balance
        gDbTrans.SqlStmt = "Select Balance From LoanTrans" & _
                " Where LoanId = " & m_LoanID & _
                " And TrasnID = " & TransID
        If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Err_line
        Balance = FormatField(rst("Balance"))
        
        Balance = Balance + txtEPFee.Value
        TransID = TransID + 1
        'Now Insert the amoun to the Loan trasnaction table
        gDbTrans.SqlStmt = "Insert Into LoanTrans " & _
                    " (LoanId,TransID,TransType," & _
                    "Amount,TransDate, Balance,UserID,Particulars) Values (" & _
                    m_LoanID & ", " & _
                    TransID & "," & _
                    wContraWithdraw & "," & _
                    txtEPFee.Value & "," & _
                    "#" & EPDate & "#," & _
                    Balance & "," & _
                    UserID & ", " & AddQuotes(GetResourceString(373)) & ")"
    
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        Balance = 0
        
        If TransID < 1 Then GoTo Err_line
            'Now get the Contra ID
        'Dim ContraId As Long
        'Miscelenoeus Head
        DueHeadID = bankClass.GetHeadIDCreated(GetResourceString(373, 366), _
                            LoadResourceStringS(373, 366), parBankIncome, 0, wis_None)
        
        ContraID = GetMaxContraTransID
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(DueHeadID) & "," & DueHeadID & _
                        wContraDeposit & "," & TransID & "," & _
                        txtEPFee.Value & "," & UserID & "," & _
                        "'')"
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        If Not bankClass.UpdateContraTrans(AccHeadID, DueHeadID, txtEPFee.Value, EPDate) Then GoTo Err_line
    Else
        
        DueHeadID = bankClass.GetHeadIDCreated(GetResourceString(373), LoadResString(373), parReceivable, 0, wis_None)
        'Now insert this details Into amount receivable table
        gDbTrans.SqlStmt = "Insert Into AmountReceivAble" & _
                        "( AccHeadId,AccID," & _
                        " TransType,TransDate,TransID,Amount," & _
                        " Balance,UserID,DueHeadID) VALUES (" & _
                        AccHeadID & "," & m_LoanID & "," & _
                        wWithdraw & ",#" & EPDate & "#," & TransID & "," & _
                        txtEPFee.Value & "," & Balance & "," & UserID & "," & _
                        DueHeadID & ")"
        'If Not gDbTrans.SQLExecute Then GoTo Err_line
        Debug.Assert txtEPFee.Value = 0
        If Not AddToAmountReceivable(AccHeadID, m_LoanID, _
                    TransID, EPDate, txtEPFee.Value, DueHeadID) Then GoTo Err_line

    End If
Else
    gDbTrans.SqlStmt = "Insert Into LoanAbnEp " & _
            " (LoanID,BKCC,AbnDate,AbnAmount,AbnDesc,abnNo, " & _
            " EpDate,EpAmount,EpDesc, EpNo ) VALUES " & _
        "(" & m_LoanID & ", 0 , #" & AbnDate & "#, " & txtAbnFee & " ," & _
        AddQuotes(txtAbnRemark, True) & "," & AddQuotes(txtAbnNo, True) & "," & _
        IIf(txtEPDate = "", " Null ", "#" & EPDate & "#") & "," & txtEPFee & "," & _
        AddQuotes(txtEPRemark, True) & "," & AddQuotes(txtEpNo, True) & ")"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    
    If chkAbn.Value = vbChecked Then
        'Check for the Last transaction Date
        If AbnDate > GetLoanLastTransDate(m_LoanID) Then GoTo Err_line
        'Now Get the loan Balance
        gDbTrans.SqlStmt = "Select Balance From LoanTrans" & _
                " Where LoanId = " & m_LoanID & _
                " And TrasnID = " & TransID
        If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Err_line
        Balance = FormatField(rst("Balance"))
        
        Balance = Balance + txtAbnFee.Value
        TransID = TransID + 1
        'Now Insert the amoun to the Loan trasnaction table
        gDbTrans.SqlStmt = "Insert Into LoanTrans " & _
                    " (LoanId,TransID,TransType," & _
                    "Amount,TransDate, Balance,UserID,Particulars) Values (" & _
                    m_LoanID & ", " & _
                    TransID & "," & _
                    wContraWithdraw & "," & _
                    txtAbnFee.Value & "," & _
                    "#" & AbnDate & "#," & _
                    Balance & "," & _
                    UserID & ", " & AddQuotes(GetResourceString(372)) & ")"
    
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        Balance = 0
        If TransID < 1 Then GoTo Err_line
            'Now get the Contra ID
        'Miscelenoeus Head
        DueHeadID = bankClass.GetHeadIDCreated(GetResourceString(372, 366), _
                           LoadResourceStringS(372, 366), parBankIncome, 0, wis_None)
        
        ContraID = GetMaxContraTransID
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(DueHeadID) & "," & DueHeadID & _
                        wContraDeposit & "," & TransID & "," & _
                        txtAbnFee.Value & "," & UserID & "," & _
                        "'')"
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        If Not bankClass.UpdateContraTrans(AccHeadID, DueHeadID, txtAbnFee.Value, EPDate) Then GoTo Err_line
        
    Else
        DueHeadID = bankClass.GetHeadIDCreated(GetResourceString(372), LoadResString(372), _
                                                        parReceivable, 0, wis_None)
        'Now insert this details Into amount receivable table
        'gDbTrans.SQLStmt = "Insert Into AmountReceivAble" & _
                        "( AccHeadId,AccID," & _
                        " TransType,TransDate,TransID,Amount," & _
                        " Balance,UserID,DueHeadID) VALUES (" & _
                        AccHeadId & "," & m_LoanID & "," & _
                        wWithdraw & ",#" & AbnDate & "#," & TransID & "," & _
                        txtABNFee.Value & "," & Balance & "," & UserID & "," & _
                        DueHeadId & ")"
        'If Not gDbTrans.SQLExecute Then GoTo Err_line
        If Not AddToAmountReceivable(AccHeadID, m_LoanID, _
                    TransID, AbnDate, txtAbnFee.Value, DueHeadID) Then GoTo Err_line

    End If
    
End If


gDbTrans.CommitTrans
InTrans = False

UpdateAbnEp = True

Exit Function

Err_line:
    If InTrans Then gDbTrans.RollBack
    
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE

End Function

Private Sub UpdateDetails()

If m_LoanID = 0 Then
    Err.Raise 2002, "Loan Payment", "You have not set the LoanId or Account Id"
    Exit Sub
End If


Dim Retval As Long
Dim rstMaster As Recordset
Dim rstTrans As Recordset
Dim rstTemp As Recordset
'Set LoanMaster Recpords
gDbTrans.SqlStmt = "Select * From LoanMaster Where LoanId = " & m_LoanID
Call gDbTrans.Fetch(rstMaster, adOpenStatic)
txtLoanAccNo = FormatField(rstMaster("AccNum"))
txtLoanDate = FormatField(rstMaster("IssueDate"))
txtDueDate = FormatField(rstMaster("LOanDueDate"))

'Set the Member Details
txtCustID = GetMemberNumber(FormatField(rstMaster("CustomerId")))
m_SchemeId = FormatField(rstMaster("SchemeId"))

'Set Interest amount to 0
txtInterest = 0
txtPenal = 0

'Set LoanTrans Records
gDbTrans.SqlStmt = "SELECT * From LoanTrans Where LoanId = " & m_LoanID & _
                " Order By TransId "
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Exit Sub

rstTrans.MoveLast
txtLoanBalance = FormatField(rstTrans("Balance"))

gDbTrans.SqlStmt = "SELECT * From LoanIntTrans Where " & _
            " LoanId = " & m_LoanID & " Order By TransId"
If gDbTrans.Fetch(rstTrans, adOpenStatic) > 0 Then
    rstTrans.MoveLast
    txtInterest = FormatField(rstTrans("IntBalance"))
    txtPenal = FormatField(rstTrans("PenalIntBalance"))
End If

'Set LoanScheme  Records
gDbTrans.SqlStmt = "SELECT * From LoanScheme Where Schemeid = " & m_SchemeId
Call gDbTrans.Fetch(rstTemp, adOpenStatic)

m_SchemeName = FormatField(rstTemp("SchemeName"))
txtLoanName = m_SchemeName

Dim l_CustClass As New clsCustReg
txtName = l_CustClass.CustomerName(rstMaster("customerID"))
Set l_CustClass = Nothing

Dim L_class As New clsLoan
Dim AsOnDate As Date
AsOnDate = gStrDate 'FormatDate(txtAbnDate.Text)
txtInterest = txtInterest + L_class.RegularInterest(m_LoanID, , AsOnDate) \ 1
txtPenal = txtPenal + L_class.PenalInterest(m_LoanID, , AsOnDate) \ 1

Set L_class = Nothing

'm_AbnAmount = 0
'm_AbnDate = vbNull
'm_EpAmount = 0
'm_EpDate = vbNull

txtAbnFee = 0
txtAbnRemark = ""
txtAbnNo = ""
txtEPDate = ""
txtEPFee = 0
txtEPRemark = ""
txtEpNo = ""

With txtAbnDate
    .Text = ""
    .Locked = False
End With
With txtAbnFee
    .Value = 0
    .Locked = False
End With
txtAbnRemark = ""
txtAbnNo = ""
chkAbn.Enabled = True


With txtEPDate
    .Text = ""
    .Locked = False
End With
With txtEPFee
    .Value = 0
    .Locked = False
End With
txtEPRemark = ""
txtEpNo = ""
chkEP.Enabled = True


'Now check for the abn or Ep
'gDbTrans.SQLStmt = "Select * From LOanAbnEp Where LoanID = " & m_LoanID
gDbTrans.SqlStmt = "Select * From LoanAbnEp " & _
        " Where LoanID = " & m_LoanID & " And BKCC = 0"

cmdUndo.Enabled = False
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    cmdUndo.Enabled = True
    txtAbnDate = FormatField(rstTemp("abnDate"))
    txtAbnFee = FormatField(rstTemp("abnAmount"))
    txtAbnRemark = FormatField(rstTemp("abnDesc"))
    txtAbnNo = FormatField(rstTemp("abnNo"))
    If Not IsNull(rstTemp("abnDate")) Then
        txtEPDate = gStrDate
        txtAbnDate.Locked = True
        txtAbnFee.Locked = True
        chkAbn.Enabled = False
        gDbTrans.SqlStmt = "SELECT * From LoanTrans" & _
                " Where LoanId = " & m_LoanID & _
                " AND TransDate = #" & rstTemp("abnDate") & "#" & _
                " AND Amount = " & txtAbnFee.Value
        If gDbTrans.Fetch(rstTrans, adOpenStatic) > 0 Then chkAbn.Value = vbChecked
    End If
    'Details Of Ep
    
    txtEPDate = FormatField(rstTemp("EpDate"))
    txtEPFee = FormatField(rstTemp("EpAmount"))
    txtEPRemark = FormatField(rstTemp("EpDesc"))
    txtEpNo = FormatField(rstTemp("EpNo"))
    
    If Not IsNull(rstTemp("EPDate")) Then
        txtEPDate.Locked = True
        txtEPFee.Locked = True
        chkEP.Enabled = False
        gDbTrans.SqlStmt = "SELECT * From LoanTrans" & _
                " Where LoanId = " & m_LoanID & _
                " AND TransDate = #" & rstTemp("EPDate") & "#" & _
                " AND Amount = " & txtEPFee.Value
        If gDbTrans.Fetch(rstTrans, adOpenStatic) > 0 Then chkEP.Value = vbChecked
    End If
End If

If txtAbnDate = "" Then txtAbnDate = gStrDate


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdabnDate_Click()
With Calendar
    
    .Left = Me.Left + fraInstall.Left + cmdAbndate.Left - (.Width / 2)
    .Top = Me.Top + fraInstall.Top + cmdAbndate.Top
    If DateValidate(txtAbnDate.Text, "/", True) Then
        .selDate = txtAbnDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtAbnDate.Text = .selDate
End With

End Sub

Private Sub cmdEpDate_Click()
With Calendar
    
    .Left = Me.Left + fraInstall.Left + cmdEpDate.Left - (.Width / 2)
    .Top = Me.Top + fraInstall.Top + cmdEpDate.Top
    If DateValidate(txtEPDate.Text, "/", True) Then
        .selDate = txtEPDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtEPDate.Text = .selDate
End With

End Sub


Private Sub cmdOk_Click()
If Not DateValidate(txtAbnDate, "/", True) Then
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtAbnDate
    Exit Sub
End If
If txtAbnFee = 0 Then
    If MsgBox(GetResourceString(506) & vbCrLf & _
        GetResourceString(541), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
    ActivateTextBox txtAbnFee
    Exit Sub
End If

If txtEPDate <> "" Then
    If Not DateValidate(txtEPDate, "/", True) Then
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtEPDate
        Exit Sub
    End If
    If txtEPFee = 0 Then
        If MsgBox(GetResourceString(506) & vbCrLf & _
            GetResourceString(541), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
        ActivateTextBox txtEPFee
        Exit Sub
    End If
End If


If UpdateAbnEp Then Unload Me

End Sub

Private Sub cmdUndo_Click()

If MsgBox(GetResourceString(583), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub

Dim rstAbnEp As Recordset

'Now check for the abn or Ep
gDbTrans.SqlStmt = "Select * From LoanAbnEp " & _
            " Where LoanID = " & m_LoanID & " And BKCC = 0"
If gDbTrans.Fetch(rstAbnEp, adOpenDynamic) < 1 Then Exit Sub

Dim Amount As Currency
Dim DueHeadID As Long
Dim AccHeadID As Long
Dim TransID As Long
Dim TransDate As Date

Dim rstTemp  As Recordset
Dim InTrans As Boolean
Dim bankClass As New clsBankAcc

gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
                " Where SchemeID = (Select SchemeID From LoanMaster" & _
                        " Where LoanID = " & m_LoanID & " )"
Call gDbTrans.Fetch(rstTemp, adOpenForwardOnly)
AccHeadID = GetHeadID(FormatField(rstTemp("SchemeName")), parMemberLoan)
    
'First Decide whther to Delete the Abn Or Ep
If Not IsNull(rstAbnEp("EpDate")) Then
    
    TransDate = rstAbnEp("EpDate")
    'Delete the only Ep Details
    'First Check for the Contra Details in  Trans table
    If chkEP.Value = vbChecked Then
        'Then Check the Last ID
        gDbTrans.SqlStmt = "SELECT * From LoanTrans" & _
                " Where LoanId = " & m_LoanID & _
                " AND TransDate = #" & TransDate & "#" & _
                " AND Amount = " & txtEPFee.Value
        If gDbTrans.Fetch(rstTemp, adOpenStatic) > 0 Then _
                        TransID = FormatField(rstTemp("TransID"))
        'Now chec the Transaction id with the Last thransction Id
        'if both are not same then exit the sub
        If TransID = GetLoanMaxTransID(m_LoanID) <> TransID Then GoTo Err_line
    End If
    
    InTrans = gDbTrans.BeginTrans
    
    gDbTrans.SqlStmt = "UPDate LoanAbnEp Set " & _
                " EPDate = Null ," & _
                " EpAmount = 0 " & _
                " EPDesc = ''," & _
                " EPDesc = '', " & _
                " Where Loanid = " & m_LoanID & _
                " And Bkcc = 0"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    If TransID Then
        DueHeadID = GetHeadID(GetResourceString(373) & " " & _
                            GetResourceString(366), parBankIncome)

        'Now Delete the Transaction In Bkcc trans
        gDbTrans.SqlStmt = "Delete * From LoanTrans " & _
                " WHERE LoanID = " & m_LoanID & _
                " And TransID = " & TransID
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        
        'Now Delete the Transaction Details in Contra Table
        gDbTrans.SqlStmt = "Delete * From ContraTrans " & _
                    " Where AccHeadID = " & AccHeadID & _
                    " And Amount = " & txtEPFee.Value & _
                    " And TransID = " & TransID & _
                    " And TransType = & wContraWithdraw "
        If Not gDbTrans.SQLExecute Then GoTo Err_line
                    
        If Not bankClass.UndoContraTrans(AccHeadID, DueHeadID, txtEPFee.Value, TransDate) Then GoTo Err_line

    Else
        'Delete the Ep Details In AmountReceivable
        DueHeadID = GetHeadID(GetResourceString(373) & " " & _
                            GetResourceString(366), parBankIncome)
        gDbTrans.SqlStmt = "Delete From AmountReceivAble" & _
                " WHERE AccHeadID = " & AccHeadID & _
                " And DueHeadID = " & DueHeadID & _
                " And Amount = " & txtEPFee.Value
               
        If Not gDbTrans.SQLExecute Then GoTo Err_line
    End If
    gDbTrans.CommitTrans
    InTrans = False
    Call UpdateDetails
    Exit Sub
    
End If


'Delete the Tranaction
TransID = 0
TransDate = rstAbnEp("AbnDate")
'Delete the only Ep Details
'First Check for the Contra Details in  Trans table
If chkAbn.Value = vbChecked Then
    'Then Check the Last ID
    gDbTrans.SqlStmt = "SELECT * From LoanTrans" & _
            " Where LoanId = " & m_LoanID & _
            " AND TransDate = #" & TransDate & "#" & _
            " AND Amount = " & txtAbnFee.Value
    If gDbTrans.Fetch(rstTemp, adOpenStatic) > 0 Then _
                        TransID = FormatField(rstTemp("TransID"))
    'Now chec the Transaction id with the Last thransction Id
    'if both are not same then exit the sub
    If TransID = GetLoanMaxTransID(m_LoanID) <> TransID Then GoTo Err_line
End If

InTrans = gDbTrans.BeginTrans

gDbTrans.SqlStmt = "Delete * From LoanAbnEp " & _
            " Where Loanid = " & m_LoanID & _
            " And Bkcc = 0"
If Not gDbTrans.SQLExecute Then GoTo Err_line
If TransID Then
    DueHeadID = GetHeadID(GetResourceString(372) & " " & _
                        GetResourceString(366), parBankIncome)

    'Now Delete the Transaction In Bkcc trans
    gDbTrans.SqlStmt = "Delete * From LoanTrans " & _
            " WHERE LoanID = " & m_LoanID & _
            " And TransID = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    
    'Now Delete the Transaction Details in Contra Table
    gDbTrans.SqlStmt = "Delete * From ContraTrans " & _
                " Where AccHeadID = " & AccHeadID & _
                " And Amount = " & txtAbnFee.Value & _
                " And TransID = " & TransID & _
                " And TransType = & wContraWithdraw "
    If Not gDbTrans.SQLExecute Then GoTo Err_line
                
    If Not bankClass.UndoContraTrans(AccHeadID, DueHeadID, txtEPFee.Value, TransDate) Then GoTo Err_line

Else
    'Delete the Ep Details In AmountReceivable
    DueHeadID = GetHeadID(GetResourceString(372) & " " & _
                    GetResourceString(366), parBankIncome)
    gDbTrans.SqlStmt = "Delete From AmountReceivAble" & _
            " WHERE AccHeadID = " & AccHeadID & _
            " And DueHeadID = " & DueHeadID & _
            " And Amount = " & txtEPFee.Value
           
    If Not gDbTrans.SQLExecute Then GoTo Err_line
End If

Set bankClass = Nothing
gDbTrans.CommitTrans
InTrans = False
Call UpdateDetails

Unload Me

Exit Sub

Err_line:

    Set bankClass = Nothing
    If InTrans Then gDbTrans.RollBack
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE

End Sub


Private Sub Form_Load()
'set icon for the form caption
'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption

Me.txtName.FONTSIZE = Me.txtName.FONTSIZE + 1
'Call SetKannadaCaption
txtAbnDate.Text = gStrDate

m_FormLoaded = True
If m_LoanID Then Call UpdateDetails


End Sub



Private Sub Form_Unload(Cancel As Integer)
m_FormLoaded = False
End Sub




Private Sub txtInterest_GotFocus()
Me.ActiveControl.SelStart = 0
Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


