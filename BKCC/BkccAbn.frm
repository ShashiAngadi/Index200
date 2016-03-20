VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CURRTEXT.OCX"
Begin VB.Form frmBKCCAbn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ABN & EP Details"
   ClientHeight    =   6540
   ClientLeft      =   1065
   ClientTop       =   1905
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo"
      Height          =   400
      Left            =   150
      TabIndex        =   39
      Top             =   5970
      Width           =   1215
   End
   Begin VB.Frame fraLoan 
      Height          =   2685
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6555
      Begin WIS_Currency_Text_Box.CurrText txtPenal 
         Height          =   360
         Left            =   4980
         TabIndex        =   17
         Top             =   2070
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtInterest 
         Height          =   360
         Left            =   1770
         TabIndex        =   14
         Top             =   2070
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label txtDueDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   4920
         TabIndex        =   10
         Top             =   1260
         Width           =   1395
      End
      Begin VB.Label lblDueDate 
         Caption         =   "&Date :"
         Height          =   465
         Left            =   3450
         TabIndex        =   9
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label lblInterest 
         Caption         =   "&Interest till date"
         Height          =   300
         Left            =   180
         TabIndex        =   13
         Top             =   2100
         Width           =   1695
      End
      Begin VB.Label lblPenal 
         Caption         =   "&Penal Interest "
         Height          =   300
         Left            =   3450
         TabIndex        =   16
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label lblCustID 
         Caption         =   "Member Number"
         Height          =   300
         Left            =   3390
         TabIndex        =   3
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label lblName 
         Caption         =   "Loan Holder Name :"
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "Loan Account No :"
         Height          =   300
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label lblLoanBalance 
         Caption         =   "Loan Balance :"
         Height          =   300
         Left            =   180
         TabIndex        =   11
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblLoanDate 
         Caption         =   "Loan Date :"
         Height          =   300
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1425
      End
      Begin VB.Label txtName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1770
         TabIndex        =   6
         Top             =   690
         Width           =   4545
      End
      Begin VB.Label txtLoanAccNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1770
         TabIndex        =   2
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label txtCustID 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   4890
         TabIndex        =   4
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label txtLoanDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1770
         TabIndex        =   8
         Top             =   1230
         Width           =   1395
      End
      Begin VB.Label txtLoanBalance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1770
         TabIndex        =   12
         Top             =   1650
         Width           =   1365
      End
   End
   Begin VB.Frame fraInstall 
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   6555
      Begin VB.TextBox txtEpNo 
         Height          =   360
         Left            =   5040
         TabIndex        =   32
         Top             =   1770
         Width           =   1305
      End
      Begin VB.TextBox txtAbnNo 
         Height          =   360
         Left            =   5010
         TabIndex        =   22
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtEPDate 
         Height          =   360
         Left            =   1770
         TabIndex        =   29
         Top             =   1770
         Width           =   1395
      End
      Begin VB.CommandButton cmdEpDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3270
         TabIndex        =   40
         Top             =   1770
         Width           =   315
      End
      Begin VB.TextBox txtEPRemark 
         Height          =   360
         Left            =   1770
         TabIndex        =   37
         Top             =   2760
         Width           =   4575
      End
      Begin VB.CheckBox chkEP 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit to Loan Account"
         Height          =   315
         Left            =   3720
         TabIndex        =   35
         Top             =   2310
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txtAbnDate 
         Height          =   360
         Left            =   1770
         TabIndex        =   19
         Top             =   270
         Width           =   1395
      End
      Begin VB.CommandButton cmdAbndate 
         Caption         =   "..."
         Height          =   315
         Left            =   3240
         TabIndex        =   20
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtAbnRemark 
         Height          =   360
         Left            =   1770
         TabIndex        =   27
         Top             =   1110
         Width           =   4575
      End
      Begin VB.CheckBox chkAbn 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit to Loan Account"
         Height          =   315
         Left            =   3750
         TabIndex        =   25
         Top             =   690
         Visible         =   0   'False
         Width           =   2565
      End
      Begin WIS_Currency_Text_Box.CurrText txtABNFee 
         Height          =   360
         Left            =   1770
         TabIndex        =   24
         Top             =   690
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtEPFee 
         Height          =   360
         Left            =   1770
         TabIndex        =   34
         Top             =   2310
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label Label1 
         Caption         =   "Ep &No"
         Height          =   315
         Left            =   3750
         TabIndex        =   31
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label lblAbnNo 
         Caption         =   "&ABN No"
         Height          =   315
         Left            =   3780
         TabIndex        =   21
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblEPDate 
         Caption         =   "&Date :"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label lblEpRemark 
         Caption         =   "&Remarks"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   1395
      End
      Begin VB.Label lblEPFee 
         Caption         =   "&EP Fee"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   2370
         Width           =   1515
      End
      Begin VB.Label lblABNDate 
         Caption         =   "&Date :"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblABNRemark 
         Caption         =   "&Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label lblABNFee 
         Caption         =   "&ABN Fee"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1515
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   60
         X2              =   6450
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4170
      TabIndex        =   38
      Top             =   5970
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5460
      TabIndex        =   30
      Top             =   5970
      Width           =   1215
   End
End
Attribute VB_Name = "frmBKCCAbn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_LoanID As Long
'Private m_SchemeName As String
'Private m_SchemeId As Integer

'Private m_AbnDate As Date
'Private m_EpDate As Date
'Private m_AbnAmount As Currency
'Private m_EpAmount As Currency
Private m_FormLoaded As Boolean

Public Event WindowClosed()

Public Property Let LoanAccountID(NewValue As Long)
   
    m_LoanID = NewValue
    If m_LoanID = 0 Then Exit Property
    If m_FormLoaded Then Call UpdateDetails
    
End Property


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblLoanAccNo = GetResourceString(80, 36, 60)
lblCustID = GetResourceString(49, 60) 'Customer ID
lblName = GetResourceString(35) 'Name
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
Dim AbnHeadID As Long

Dim bankClass As clsBankAcc

AccHeadID = GetHeadID(GetResourceString(229) & " " & _
                        GetResourceString(58), parMemberLoan)

Dim AddToAccount As Boolean
Dim InTrans As Boolean

'PrevAbnDate = "1/1/1000"
'PrevEPDate = "1/1/1000"

gDbTrans.SqlStmt = "SELECT * From LoanAbnEp " & _
            "Where LoanID = " & m_LoanID & " And Bkcc = 1"
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

TransID = GetKCCMaxTransID(m_LoanID)
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
                " And Bkcc = 1"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    If chkEP.Value = vbChecked Then
        Dim KccClass  As New clsBkcc
        TransID = KccClass.WithdrawAmount(CInt(m_LoanID), txtEPFee.Value, _
                                    GetResourceString(373), EPDate)
        Set KccClass = Nothing
        If TransID < 1 Then GoTo Err_line
            'Now get the Contra ID
        'Dim ContraId As Long
        'Miscelenoeus Head
        AbnHeadID = bankClass.GetHeadIDCreated(GetResourceString(373, 366), _
                    LoadResourceStringS(373, 366), parBankIncome, 0, wis_None)
        
        ContraID = GetMaxContraTransID
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(AbnHeadID) & "," & AbnHeadID & _
                        wContraDeposit & "," & TransID & "," & _
                        txtEPFee.Value & "," & UserID & "," & _
                        "'')"
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        If Not bankClass.UpdateContraTrans(AccHeadID, AbnHeadID, txtEPFee.Value, EPDate) Then GoTo Err_line
    Else
        
        AbnHeadID = bankClass.GetHeadIDCreated(GetResourceString(373), LoadResString(373), parReceivable, 0, wis_None)
        'Now insert this details Into amount receivable table
        gDbTrans.SqlStmt = "Insert Into AmountReceivAble" & _
                        "( AccHeadId,AccID," & _
                        " TransType,TransDate,TransID,Amount," & _
                        " Balance,UserID,DueHeadID) VALUES (" & _
                        AccHeadID & "," & m_LoanID & "," & _
                        wWithdraw & ",#" & EPDate & "#," & TransID & "," & _
                        txtEPFee.Value & "," & Balance & "," & UserID & "," & _
                        AbnHeadID & ")"
        'If Not gDbTrans.SQLExecute Then GoTo Err_line
        Debug.Assert txtEPFee = 0
        If Not AddToAmountReceivable(AccHeadID, m_LoanID, TransID, _
                     EPDate, txtEPFee.Value, AbnHeadID) Then GoTo Err_line

    End If
Else
    gDbTrans.SqlStmt = "Insert Into LoanAbnEp " & _
            " (LoanID,BKCC,AbnDate,AbnAmount,AbnDesc,abnNo, " & _
            " EpDate,EpAmount,EpDesc, EpNo ) VALUES " & _
        "(" & m_LoanID & ", 1 , #" & AbnDate & "#, " & txtAbnFee & " ," & _
        AddQuotes(txtAbnRemark, True) & "," & AddQuotes(txtAbnNo, True) & "," & _
        IIf(txtEPDate = "", " Null ", "#" & EPDate & "#") & "," & txtEPFee & "," & _
        AddQuotes(txtEPRemark, True) & "," & AddQuotes(txtEpNo, True) & ")"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    
    If chkAbn.Value = vbChecked Then
        'Dim KccClass  As New clsBkcc
        TransID = KccClass.WithdrawAmount(CInt(m_LoanID), txtAbnFee.Value, _
                                    GetResourceString(372), AbnDate)
        Set KccClass = Nothing
        If TransID < 1 Then GoTo Err_line
            'Now get the Contra ID
        'Miscelenoeus Head
        AbnHeadID = bankClass.GetHeadIDCreated(GetResourceString(372, 366), _
                LoadResourceStringS(372, 366), parBankIncome, 0, wis_None)
        
        ContraID = GetMaxContraTransID
        gDbTrans.SqlStmt = "Insert Into ContraTrans " & _
                        "( ContraID,AccHeadId,AccID," & _
                        " TransType,TransID,Amount," & _
                        " UserID,VoucherNo) VALUES (" & _
                        ContraID & "," & _
                        GetParentID(AbnHeadID) & "," & AbnHeadID & _
                        wContraDeposit & "," & TransID & "," & _
                        txtAbnFee.Value & "," & UserID & "," & _
                        "'')"
        If Not gDbTrans.SQLExecute Then GoTo Err_line
        If Not bankClass.UpdateContraTrans(AccHeadID, AbnHeadID, txtAbnFee.Value, EPDate) Then GoTo Err_line
        
    Else
        AbnHeadID = bankClass.GetHeadIDCreated(GetResourceString(372), LoadResString(372), parReceivable, 0, wis_None)
        'Now insert this details Into amount receivable table
        gDbTrans.SqlStmt = "Insert Into AmountReceivAble" & _
                        "( AccHeadId,AccID," & _
                        " TransType,TransDate,TransID,Amount," & _
                        " Balance,UserID,DueHeadID) VALUES (" & _
                        AccHeadID & "," & m_LoanID & "," & _
                        wWithdraw & ",#" & AbnDate & "#," & TransID & "," & _
                        txtAbnFee.Value & "," & Balance & "," & UserID & "," & _
                        AbnHeadID & ")"
        'If Not gDbTrans.SQLExecute Then GoTo Err_line
        Debug.Assert txtAbnFee.Value = 0
        If Not AddToAmountReceivable(AccHeadID, m_LoanID, TransID, _
                            AbnDate, txtAbnFee.Value, AbnHeadID) Then GoTo Err_line

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
'Set BKCC Master Records
gDbTrans.SqlStmt = "Select * From BKCCMaster Where LoanId = " & m_LoanID
Call gDbTrans.Fetch(rstMaster, adOpenStatic)
txtLoanAccNo = FormatField(rstMaster("AccNum"))
txtLoanDate = FormatField(rstMaster("IssueDate"))

Dim KccClass As New clsBkcc
txtDueDate = GetIndianDate(KccClass.LoanDueDate(m_LoanID))

'Set the Member Details
txtCustID = GetMemberNumber(FormatField(rstMaster("CustomerId")))

'Set Interest amount to 0
txtInterest = 0
txtPenal = 0

'Set Trans Records
gDbTrans.SqlStmt = "SELECT * From BKCCTrans " & _
            " Where LoanId = " & m_LoanID & _
            " Order By TransId "
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Exit Sub

rstTrans.MoveLast
txtLoanBalance = FormatField(rstTrans("Balance"))

gDbTrans.SqlStmt = "SELECT * From BKCCIntTrans" & _
                " Where LoanId = " & m_LoanID & _
                " Order By TransId"
If gDbTrans.Fetch(rstTrans, adOpenStatic) > 0 Then
    rstTrans.MoveLast
    txtInterest = FormatField(rstTrans("IntBalance"))
    txtPenal = FormatField(rstTrans("PenalIntBalance"))
End If

Dim l_CustClass As New clsCustReg
txtName = l_CustClass.CustomerName(rstMaster("CustomerID"))
Set l_CustClass = Nothing

Dim L_class As New clsBkcc
Dim AsOnDate As Date
AsOnDate = gStrDate 'FormatDate(txtAbnDate.Text)
txtInterest = txtInterest + L_class.RegularInterest(m_LoanID, AsOnDate) \ 1
txtPenal = txtPenal + L_class.PenalInterest(m_LoanID, AsOnDate) \ 1

Set L_class = Nothing

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
gDbTrans.SqlStmt = "Select * From LoanAbnEp " & _
            " Where LoanID = " & m_LoanID & " And BKCC = 1"

cmdUndo.Enabled = False
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
    cmdUndo.Enabled = True
    'Details Of Abn
    txtAbnDate = FormatField(rstTemp("abnDate"))
    txtAbnFee = FormatField(rstTemp("abnAmount"))
    txtAbnRemark = FormatField(rstTemp("abnDesc"))
    txtAbnNo = FormatField(rstTemp("abnNo"))
    If Not IsNull(rstTemp("abnDate")) Then
        txtEPDate = gStrDate
        txtAbnDate.Locked = True
        txtAbnFee.Locked = True
        chkAbn.Enabled = False
        gDbTrans.SqlStmt = "SELECT * From BKCCTrans" & _
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
        gDbTrans.SqlStmt = "SELECT * From BKCCTrans" & _
                " Where LoanId = " & m_LoanID & _
                " AND TransDate = # " & rstTemp("EPDate") & "#" & _
                " AND Amount = " & txtEPFee.Value
        If gDbTrans.Fetch(rstTrans, adOpenStatic) > 0 Then chkEP.Value = vbChecked
    End If
End If
    
txtAbnFee.Tag = txtAbnFee.Value
txtEPFee.Tag = txtEPFee.Value

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
Else
    txtEPFee.Value = 0
End If


If UpdateAbnEp Then Unload Me

End Sub

Private Sub cmdUndo_Click()

If MsgBox(GetResourceString(583), vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Sub

Dim rstAbnEp As Recordset

'Now check for the abn or Ep
gDbTrans.SqlStmt = "Select * From LoanAbnEp " & _
            " Where LoanID = " & m_LoanID & " And BKCC = 1"
If gDbTrans.Fetch(rstAbnEp, adOpenDynamic) < 1 Then Exit Sub

Dim Amount As Currency
Dim DueHeadID As Long
Dim AccHeadID As Long
Dim TransID As Long
Dim TransDate As Date

Dim rstTemp  As Recordset
Dim InTrans As Boolean
Dim bankClass As New clsBankAcc

AccHeadID = GetHeadID(GetResourceString(229) & " " & _
                        GetResourceString(58), parMemberLoan)

'First Decide whther to Delete the Abn Or Ep
If Not IsNull(rstAbnEp("EpDate")) Then
    
    TransDate = rstAbnEp("EpDate")
    'Delete the only Ep Details
    'First Check for the Contra Details in  Trans table
    If chkEP.Value = vbChecked Then
        'Then Check the Last ID
        gDbTrans.SqlStmt = "SELECT * From BKCCTrans" & _
                " Where LoanId = " & m_LoanID & _
                " AND TransDate = #" & TransDate & "#" & _
                " AND Amount = " & txtEPFee.Value
        If gDbTrans.Fetch(rstTemp, adOpenStatic) > 0 Then TransID = FormatField(rstTemp("TransID"))
        'Now chec the Transaction id with the Last thransction Id
        'if both are not same then exit the sub
        If TransID = GetKCCMaxTransID(m_LoanID) <> TransID Then GoTo Err_line
    End If
    
    InTrans = gDbTrans.BeginTrans
    
    gDbTrans.SqlStmt = "UPDate LoanAbnEp Set " & _
                " EPDate = Null ," & _
                " EpAmount = 0 " & _
                " EPDesc = ''," & _
                " EPDesc = '', " & _
                " Where Loanid = " & m_LoanID & _
                " And Bkcc = 1"
    If Not gDbTrans.SQLExecute Then GoTo Err_line
    If TransID Then
        DueHeadID = GetHeadID(GetResourceString(373) & " " & _
                            GetResourceString(366), parBankIncome)

        'Now Delete the Transaction In Bkcc trans
        gDbTrans.SqlStmt = "Delete * From BkccTrans " & _
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

TransDate = rstAbnEp("AbnDate")
'Delete the only Ep Details
'First Check for the Contra Details in  Trans table
If chkAbn.Value = vbChecked Then
    'Then Check the Last ID
    gDbTrans.SqlStmt = "SELECT * From BKCCTrans" & _
            " Where LoanId = " & m_LoanID & _
            " AND TransDate = #" & TransDate & "#" & _
            " AND Amount = " & txtAbnFee.Value
    If gDbTrans.Fetch(rstTemp, adOpenStatic) > 0 Then _
        TransID = FormatField(rstTemp("TransID"))
    'Now chec the Transaction id with the Last thransction Id
    'if both are not same then exit the sub
    If TransID = GetKCCMaxTransID(m_LoanID) <> TransID Then GoTo Err_line
End If

InTrans = gDbTrans.BeginTrans

gDbTrans.SqlStmt = "Delete * From LoanAbnEp " & _
            " Where Loanid = " & m_LoanID & _
            " And Bkcc = 1"
If Not gDbTrans.SQLExecute Then GoTo Err_line
If TransID Then
    DueHeadID = GetHeadID(GetResourceString(372) & " " & _
                        GetResourceString(366), parBankIncome)

    'Now Delete the Transaction In Bkcc trans
    gDbTrans.SqlStmt = "Delete * From BkccTrans " & _
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

'Call SetKannadaCaption
txtAbnDate.Text = gStrDate

m_FormLoaded = True
If m_LoanID Then Call UpdateDetails


End Sub



Private Sub Form_Unload(Cancel As Integer)
m_FormLoaded = False
End Sub




Private Sub txtInterest_GotFocus()

ActiveControl.SelStart = 0
ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


