VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmLoanTrans 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Detials"
   ClientHeight    =   9210
   ClientLeft      =   1410
   ClientTop       =   1230
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleMode       =   0  'User
   ScaleWidth      =   13380
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   8280
      TabIndex        =   36
      Top             =   8750
      Width           =   1515
   End
   Begin VB.Frame fraCust 
      Height          =   1635
      Left            =   68
      TabIndex        =   62
      Top             =   -100
      Width           =   9705
      Begin VB.ComboBox cmbMemberType 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   3645
      End
      Begin VB.CommandButton cmdCustName 
         BackColor       =   &H8000000A&
         Caption         =   "..."
         Height          =   315
         Left            =   8640
         TabIndex        =   63
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   7320
         MaxLength       =   9
         TabIndex        =   10
         Top             =   750
         Width           =   990
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "..."
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtLoanAccNo 
         Height          =   315
         Left            =   7290
         TabIndex        =   3
         Top             =   300
         Width           =   1005
      End
      Begin VB.CommandButton cmdLoan 
         Caption         =   "..."
         Height          =   375
         Left            =   8640
         TabIndex        =   4
         Top             =   250
         Width           =   915
      End
      Begin VB.ComboBox cmbLoanScheme 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   300
         Width           =   3645
      End
      Begin VB.Label lblMemberType 
         AutoSize        =   -1  'True
         Caption         =   "Member Type :"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblCustID 
         AutoSize        =   -1  'True
         Caption         =   "Member No :"
         Height          =   300
         Left            =   5640
         TabIndex        =   9
         Top             =   780
         Width           =   915
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "&Loan Account No"
         Height          =   300
         Left            =   5640
         TabIndex        =   2
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   1155
         Width           =   6555
      End
      Begin VB.Label lblCustName 
         Caption         =   "&Name :"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLoanScheme 
         Caption         =   "Loan Scheme"
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.Frame fra2 
      Height          =   1155
      Left            =   68
      TabIndex        =   51
      Top             =   1410
      Width           =   9705
      Begin VB.CommandButton cmdAbn 
         Caption         =   "..."
         Height          =   315
         Left            =   9120
         TabIndex        =   32
         Top             =   600
         Width           =   315
      End
      Begin VB.Label txtLoanAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label txtRepaid 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2520
         TabIndex        =   59
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label txtOverDue 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7560
         TabIndex        =   55
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label txtSanction 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4920
         TabIndex        =   53
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblSanction 
         Caption         =   "Sanction Loan Amount :"
         Height          =   300
         Left            =   4800
         TabIndex        =   52
         Top             =   250
         Width           =   1905
      End
      Begin VB.Label lblOverDue 
         Caption         =   "Over Due Amount :"
         Height          =   300
         Left            =   7440
         TabIndex        =   54
         Top             =   250
         Width           =   1605
      End
      Begin VB.Label lblRepaid 
         Caption         =   "Repaid Amount :"
         Height          =   300
         Left            =   2640
         TabIndex        =   58
         Top             =   250
         Width           =   1845
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "Loan Amount :"
         Height          =   300
         Left            =   240
         TabIndex        =   56
         Top             =   250
         Width           =   1545
      End
   End
   Begin VB.Frame fraRepay 
      Height          =   2955
      Left            =   68
      TabIndex        =   61
      Top             =   2400
      Width           =   9705
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   2445
         Width           =   2325
      End
      Begin VB.CommandButton cmdMisc 
         Caption         =   "..."
         Height          =   315
         Left            =   9120
         TabIndex        =   18
         Top             =   1500
         Width           =   315
      End
      Begin VB.CommandButton cmdPenalBalance 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   23
         Top             =   2040
         Width           =   315
      End
      Begin VB.CommandButton cmdIntBalance 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Top             =   1560
         Width           =   315
      End
      Begin VB.TextBox txtTransDate 
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton cmdTransDate 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   13
         Top             =   210
         Width           =   315
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotalIntBalance 
         Height          =   345
         Left            =   7920
         TabIndex        =   20
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtRegInt 
         Height          =   345
         Left            =   7920
         TabIndex        =   22
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPenalINT 
         Height          =   345
         Left            =   7920
         TabIndex        =   25
         Top             =   1065
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMisc 
         Height          =   345
         Left            =   7920
         TabIndex        =   28
         Top             =   1485
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtPrincAmount 
         Height          =   345
         Left            =   7920
         TabIndex        =   31
         Top             =   1905
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtTotAmt 
         Height          =   345
         Left            =   7800
         TabIndex        =   16
         Top             =   2445
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblParticulars 
         Caption         =   "Particulars"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   2445
         Width           =   2295
      End
      Begin VB.Label lblTotalIntBalance 
         Caption         =   "Interest balance"
         Height          =   300
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   5400
         X2              =   10170
         Y1              =   2315
         Y2              =   2315
      End
      Begin VB.Label lblPenalBalance 
         Caption         =   "Peanl Int Balance"
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   240
         TabIndex        =   43
         Top             =   2010
         Width           =   2295
      End
      Begin VB.Label txtPenalBalance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2640
         TabIndex        =   44
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label txtLastTransDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2640
         TabIndex        =   38
         Top             =   675
         Width           =   2205
      End
      Begin VB.Label txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2640
         TabIndex        =   40
         Top             =   1125
         Width           =   2205
      End
      Begin VB.Label LblTotAmt 
         Caption         =   "Principal Amount :"
         ForeColor       =   &H00404000&
         Height          =   300
         Left            =   5520
         TabIndex        =   30
         Top             =   1905
         Width           =   2175
      End
      Begin VB.Label lblMiscAmount 
         Caption         =   "&Misc Amount"
         Height          =   300
         Left            =   5520
         TabIndex        =   27
         Top             =   1485
         Width           =   1905
      End
      Begin VB.Label lblPenalInt 
         Caption         =   "&Penal interest :"
         Height          =   300
         Left            =   5520
         TabIndex        =   24
         Top             =   1065
         Width           =   1965
      End
      Begin VB.Label lblTransDate 
         Caption         =   "Transaction &Date"
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblRegIntBalance 
         Caption         =   "&Interest Balance"
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   240
         TabIndex        =   41
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblRegInt 
         Caption         =   "&Regular Interest"
         Height          =   300
         Left            =   5520
         TabIndex        =   21
         Top             =   660
         Width           =   2295
      End
      Begin VB.Label lblRepayAmount 
         Caption         =   "Repay &Amount :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5520
         TabIndex        =   15
         Top             =   2445
         Width           =   1875
      End
      Begin VB.Label lblBalance 
         Caption         =   "Loan Balance"
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   240
         TabIndex        =   39
         Top             =   1125
         Width           =   2295
      End
      Begin VB.Label lblLastTransDate 
         Caption         =   "Last Transaction "
         Height          =   300
         Left            =   240
         TabIndex        =   37
         Top             =   675
         Width           =   2385
      End
      Begin VB.Label txtRegIntBalance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2640
         TabIndex        =   42
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Height          =   3555
      Left            =   75
      TabIndex        =   60
      Top             =   5100
      Width           =   9705
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "P&hoto"
         Height          =   400
         Left            =   1560
         TabIndex        =   65
         Top             =   3050
         Width           =   1215
      End
      Begin VB.CommandButton cmdRepay 
         Caption         =   "&Repay"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   400
         Left            =   8400
         TabIndex        =   33
         Top             =   3050
         Width           =   1215
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo Last"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6480
         TabIndex        =   34
         Top             =   3050
         Width           =   1545
      End
      Begin VB.CommandButton cmdPay 
         Caption         =   "&Payment"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   195
         MaskColor       =   &H00800080&
         TabIndex        =   35
         Top             =   3050
         Width           =   1215
      End
      Begin VB.Frame fraInstructions 
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   2025
         Left            =   360
         TabIndex        =   64
         Top             =   840
         Width           =   9075
         Begin VB.CommandButton cmdAddNote 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   0
            Width           =   405
         End
         Begin RichTextLib.RichTextBox rtfNote 
            Height          =   1965
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   3466
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"LnTrans.frx":0000
         End
      End
      Begin ComctlLib.TabStrip tabLoans 
         Height          =   2655
         Left            =   150
         TabIndex        =   45
         Top             =   300
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   4683
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   ""
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraLoanGrid 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   2085
         Left            =   315
         TabIndex        =   50
         Top             =   750
         Width           =   9165
         Begin VB.CommandButton cmdNextTrans 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   420
            Width           =   375
         End
         Begin VB.CommandButton cmdPrevTrans 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8820
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   15
            Width           =   375
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   315
            Left            =   8790
            Picture         =   "LnTrans.frx":0082
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1470
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   1830
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Visible         =   0   'False
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   3228
            _Version        =   393216
            Rows            =   10
            Cols            =   5
            AllowBigSelection=   0   'False
            ScrollBars      =   2
         End
      End
   End
End
Attribute VB_Name = "frmLoanTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_LoanID As Long
Dim m_CustomerID As Long
Dim m_ClsCust As clsCustReg
Dim m_rstLoanMast As Recordset
Dim m_rstLoanTrans As Recordset

Private m_retVar As Variant
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmLoanInst As frmLoanInst
Attribute m_frmLoanInst.VB_VarHelpID = -1
Private WithEvents m_frmPrintTrans As frmPrintTrans
Attribute m_frmPrintTrans.VB_VarHelpID = -1
Private m_Notes As New clsNotes

'Private WithEvents m_clsReceivable As  clsReceive
Private m_clsReceivable As clsReceive

Public Event AccountChanged(ByVal LoanID As Long)
Public Event AccountTransaction(TransacType As wisTransactionTypes)
Public Event PaymentClicked(ByVal LoanID As Long)
Public Event RePaymentClicked(ByVal LoanID As Long)
Public Event WindowClosed()
Private Sub ClearControls()

If m_LoanID = 0 Then Exit Sub

Set m_rstLoanMast = Nothing
Set m_rstLoanTrans = Nothing
Set m_frmLoanInst = Nothing
Set m_frmPrintTrans = Nothing
Call m_ClsCust.NewCustomer

txtTransDate.Tag = ""

txtLoanAccNo = ""
txtLoanAccNo.Locked = False
txtSanction = ""
txtOverDue = ""
txtBalance = ""

txtLoanAmount = ""
txtMisc = 0
Set m_clsReceivable = Nothing
txtLastTransDate = ""

txtRegIntBalance = ""
txtTotalIntBalance = 0
If Not DateValidate(txtTransDate, "/", True) Then
    txtTransDate.Text = gStrDate
    txtTransDate.Tag = txtTransDate.Text
End If
txtTotAmt = 0
txtTotAmt.Tag = 0
txtPenalInt = 0
txtRegInt = 0
txtTotalIntBalance = 0
txtRepaid = ""
txtPrincAmount = 0

txtRegInt.BackColor = wisWhite
txtTotalIntBalance.BackColor = wisWhite
txtPenalInt.BackColor = wisWhite
txtRegIntBalance.BackColor = wisWhite
txtMisc.BackColor = wisWhite
txtCustNAme.BackColor = wisWhite
txtLoanAccNo.BackColor = wisWhite
txtPrincAmount.BackColor = wisWhite
txtTotAmt.BackColor = wisWhite
txtRegInt.Enabled = True
txtTotalIntBalance.Enabled = True
txtPenalInt.Enabled = True
txtRegIntBalance.Enabled = True
txtMisc.Enabled = True
txtCustNAme.Enabled = True
txtLoanAccNo.Enabled = True
txtPrincAmount.Enabled = True
txtTotAmt.Enabled = True

cmdRepay.Enabled = False
cmdIntBalance.Enabled = False
cmdPenalBalance.Enabled = False
cmdPay.Enabled = True

grd.Clear
Set m_frmLoanInst = Nothing

m_LoanID = 0

Err_line:

    If Err Then
        MsgBox "Error in Clear Controls", vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    End If

End Sub


Private Sub InitGrid()
grd.Clear
With grd
    .Clear
    .Cols = 8
    .Rows = 11
    .AllowUserResizing = flexResizeBoth
    .FixedCols = 1
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = "SL No": .ColWidth(0) = 400
    .Col = 1: .Text = GetResourceString(37): .ColWidth(1) = 1200
    .Col = 2: .Text = GetResourceString(39): .ColWidth(2) = 1300 'Particulars
    .Col = 3: .Text = GetResourceString(277): .ColWidth(2) = 800
    .Col = 4: .Text = GetResourceString(276): .ColWidth(3) = 1200
    .Col = 5: .Text = GetResourceString(344): .ColWidth(4) = 900
    .Col = 6: .Text = GetResourceString(345): .ColWidth(5) = 650
    .Col = 7: .Text = GetResourceString(42): .ColWidth(5) = 1150
End With


End Sub

Private Sub LoadCustomerLoans()

If m_CustomerID = 0 Then Exit Sub
Call ClearControls

Dim SqlStr As String
Dim rst As Recordset
cmdRepay.Enabled = False
cmdIntBalance.Enabled = False
cmdPenalBalance.Enabled = False
cmdPrevTrans.Enabled = False
cmdNextTrans.Enabled = False
cmdUndo.Enabled = False
cmdPay.Enabled = False

tabLoans.Tabs.Clear
SqlStr = "SELECT LoanID,A.SchemeID,SchemeName FROM LoanMaster A, " & _
            " LoanScheme B WHERE A.SchemeID = b.SchemeID " & _
            " AND CustomerID = " & m_CustomerID & _
            " ANd (LoanClosed <> 2 or LoanClosed is Null)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub
SqlStr = ""
Dim count As Integer

tabLoans.Tabs.Add count + 1, "KEY_Note", GetResourceString(219)
count = 1
cmdPay.Enabled = True
While Not rst.EOF
    tabLoans.Tabs.Add count + 1, "KEY" & FormatField(rst("LoanId")), _
        FormatField(rst("SchemeName"))
    tabLoans.Tabs(tabLoans.Tabs.count).Tag = FormatField(rst("SchemeID"))
    rst.MoveNext
    count = count + 1
Wend
tabLoans.Tabs(count).Selected = True

txtLoanAccNo.Locked = True

End Sub

Private Function LoadClosedCustomerLoan(LoanID As Long)

On Error GoTo ErrLine

If m_CustomerID = 0 Then Exit Function

Dim SqlStr As String
Dim rst As Recordset

SqlStr = "SELECT LoanID,A.SchemeID,SchemeName FROM LoanMaster A, " & _
            " LoanScheme B WHERE A.SchemeID = b.SchemeID " & _
            " AND CustomerID = " & m_CustomerID & _
            " AND LoanId = " & LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Function
SqlStr = ""
'Dim count As Integer
'count = 1
'tabLoans.Tabs.Add count, "KEY_Note", FormatField(Rst("SchemeName"))
While Not rst.EOF
    tabLoans.Tabs.Add 1, "KEY" & FormatField(rst("LoanId")), _
        FormatField(rst("SchemeName"))
    tabLoans.Tabs(tabLoans.Tabs.count).Tag = FormatField(rst("SchemeID"))
    rst.MoveNext
    
Wend
tabLoans.Tabs(1).Selected = True

'txtLoanAccNo.Locked = True
LoadClosedCustomerLoan = True

ErrLine:

End Function


Public Sub LoadLoanDetail(ByVal LoanID As Long)
On Error Resume Next

If LoanID = 0 Then Exit Sub
Call ClearControls

Dim LoanHeadID As Long

Dim SqlStr As String
Dim rst As Recordset
Dim SchemeID As Long
Dim hasEMI As Boolean
'Get the loan master detial
SqlStr = "SELECT * FROM LoanMaster WHERE LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(m_rstLoanMast, adOpenDynamic) < 1 Then Exit Sub

m_LoanID = LoanID
SchemeID = FormatField(m_rstLoanMast("SchemeID"))
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
                " Where SchemeID = " & m_rstLoanMast("SchemeID")
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
LoanHeadID = GetHeadID(FormatField(rst("SchemeName")), parMemberLoan)

fraLoanGrid.Visible = True
fraLoanGrid.ZOrder 0
grd.Visible = True
txtCustID = GetMemberNumber(FormatField(m_rstLoanMast("CustomerId")))
txtLoanAccNo = FormatField(m_rstLoanMast("AccNum"))
txtSanction = FormatField(m_rstLoanMast("LoanAmount"))
txtLoanAmount = txtSanction

'cmdPay.Enabled = True
Dim InstType As wisInstallmentTypes
InstType = FormatField(m_rstLoanMast("InstMode"))
Dim count As Integer
For count = 0 To cmbLoanScheme.ListCount - 1
    If SchemeID = cmbLoanScheme.ItemData(count) Then
        cmbLoanScheme.ListIndex = count
        Exit For
    End If
Next count

'Load the installment details
SqlStr = "SELECT * FROM LoanInst WHERE LoanId = " & LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    hasEMI = True
    If m_frmLoanInst Is Nothing Then Set m_frmLoanInst = New frmLoanInst
    m_frmLoanInst.LoanID = LoanID
    m_frmLoanInst.Operation = InstRepay
    txtRegIntBalance.Tag = txtRegIntBalance
    txtPenalBalance.Tag = txtPenalBalance
    txtTotalIntBalance.Tag = Val(Val(txtRegIntBalance) + Val(txtPenalBalance))
    Load m_frmLoanInst
    If Err Then Debug.Print Err.Number & "  " & Err.Description
    
    m_frmLoanInst.txtCustNAme = txtCustNAme
    m_frmLoanInst.txtLoanAmount = txtLoanAmount
End If

'Get the amount Details
Dim transType As wisTransactionTypes
transType = wWithdraw
SqlStr = "SELECT SUM(Amount) FROM LoanTrans " & _
        " WHERE LoanID = " & LoanID & _
        " AND (TransType = " & wWithdraw & _
        " OR TransType = " & wContraWithdraw & ")"
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then cmdPay.SetFocus: Exit Sub 'No Payment has made

txtLoanAmount = FormatField(rst(0))
txtBalance = txtLoanAmount

If Val(txtLoanAmount) > 0 Then
    cmdRepay.Enabled = True
    cmdIntBalance.Enabled = True
    cmdPenalBalance.Enabled = True
    cmdUndo.Enabled = gCurrUser.IsAdmin
    'txtTransDate = GetAppFormatDate(gstrdate)
    txtTransDate.Tag = txtTransDate.Text
    
    cmdUndo.Enabled = gCurrUser.IsAdmin
    cmdPay.Enabled = True
Else
    cmdUndo.Enabled = False
End If

transType = wDeposit
SqlStr = "SELECT SUM(Amount) FROM LoanTrans " & _
        " WHERE LoanID = " & LoanID & _
        " AND (TransType = " & wDeposit & _
        " OR TransType = " & wContraDeposit & ")"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub 'No Payment has made
txtRepaid = FormatField(rst(0))

'Get the Loan Balance as on Today
txtBalance = FormatCurrency(Val(txtLoanAmount) - Val(txtRepaid))

Dim TransDate As Date

If transType = wDeposit And DateValidate(txtTransDate, "/", True) Then
    TransDate = GetSysFormatDate(txtTransDate)
    Dim L_clsLoan As New clsLoan
    txtRegIntBalance = FormatCurrency(L_clsLoan.RegInterestBalance(LoanID, TransDate))
    txtPenalBalance = FormatCurrency(L_clsLoan.PenalInterestBalance(LoanID, TransDate))
'    txtTotalIntBalance = L_clsLoan.InterestBalance(LoanID, TransDate)
    txtTotalIntBalance = Val(txtRegIntBalance) + Val(txtPenalBalance)
    txtRegInt = L_clsLoan.RegularInterest(LoanID, , TransDate)
    txtPenalInt = L_clsLoan.PenalInterest(LoanID, , TransDate)
    If hasEMI Then txtPrincAmount = L_clsLoan.PrincipalAmountAsOn(LoanID, TransDate)
    Call CheckForDueAmount
    txtOverDue = FormatCurrency(L_clsLoan.OverDueAmount(LoanID, , TransDate))
    cmdAbn.Enabled = IIf(txtOverDue > 0, True, False)
    txtTransDate.Tag = txtTransDate.Text
End If
'Now Check for the any miscaleneous amount
txtMisc = GetReceivAbleAmount(LoanHeadID, m_LoanID)
If txtMisc Then
    gDbTrans.SqlStmt = "Select * From AmountReceivable" & _
            " WHERE AccHeadID = " & LoanHeadID & _
            " And AccID = " & m_LoanID & _
            " And TransID = " & GetLoanMaxTransID(m_LoanID)
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Set m_clsReceivable = New clsReceive
        While Not rst.EOF
            Call m_clsReceivable.AddHeadAndAmount(rst("DueHeadID"), rst("Amount"))
            rst.MoveNext
        Wend
        Set rst = Nothing
    End If
End If

Set L_clsLoan = Nothing
'Get the last Transaction Detail
SqlStr = "SELECT Top 1 TransDate FROM LoanTrans " & _
        " WHERE LoanID = " & LoanID & " ORDER BY TransID desc"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                txtLastTransDate = FormatField(rst(0))
                
SqlStr = "SELECT Top 1 TransDate FROM LoanIntTrans " & _
    " WHERE LoanID = " & LoanID & " ORDER BY TransID desc"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    If DateDiff("d", GetSysFormatDate(txtLastTransDate), rst(0)) > 0 Then _
        txtLastTransDate = FormatField(rst(0))
End If

If txtLastTransDate = "" Then
    cmdUndo.Enabled = False
    Exit Sub
End If

If DateValidate(txtTransDate, "/", True) Then
    cmdRepay.Enabled = IIf(WisDateDiff(txtLastTransDate, txtTransDate) >= 0, -1, 0)
    cmdIntBalance.Enabled = cmdRepay.Enabled
    cmdPenalBalance.Enabled = cmdRepay.Enabled
    cmdUndo.Enabled = cmdRepay.Enabled And gCurrUser.IsAdmin
End If

'Get the Loan transction detail
SqlStr = "SELECT 'Principal',LoanId,TransDate,TransId," & _
    " TransType,Amount,Balance,Particulars," & _
    " 0 as MiscAmount FROM LoanTrans WHERE LoanId = " & LoanID & _
    " UNION " & _
    " SELECT 'Interest',LoanId,TransDate,TransId," & _
    " TransType,IntAmount as Amount,PenalIntAmount as Balance, Particulars," & _
    " MiscAmount FROM LoanIntTrans WHERE LoanId = " & LoanID & _
    " ORDER BY TransDate,TransID"

gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(m_rstLoanTrans, adOpenDynamic)

With rtfNote
    '.BackColor = IIf(ClosedDate = "", vbWhite, wisGray)
    '.Enabled = IIf(ClosedDate = "", True, False)
    Call m_Notes.LoadNotes(wis_Loans, m_LoanID)
End With
Call m_Notes.DisplayNote(rtfNote)

fraLoanGrid.Visible = True
grd.Visible = True
cmdNextTrans.Tag = 0
cmdPrevTrans.Tag = 0
cmdPrevTrans.Enabled = False
Dim TransID As Integer

Do
    If m_rstLoanTrans.recordCount > 19 Then
        cmdPrevTrans.Enabled = True
        m_rstLoanTrans.MoveLast
        TransID = m_rstLoanTrans("transID")
        m_rstLoanTrans.MoveFirst
        m_rstLoanTrans.Find "TransID >= " & TransID, , adSearchForward
        'm_rstLoanTrans.Move -(m_rstLoanTrans.RecordCount Mod 20)
    End If
    Call ShowTransaction
    If cmdNextTrans.Enabled = False Then Exit Do
'    m_rstLoanTrans.MoveNext  'Cusror has moved one step Behind
                            'in ShowTransaction Function To Equal that we are moving
                            'On step ahead
    cmdPrevTrans.Enabled = True
Loop

'Check for the loan closure
If FormatField(m_rstLoanMast("LoanClosed")) Then
    cmdAbn.Enabled = False
    lblLastTransDate = GetResourceString(58, 282)
    cmdRepay.Enabled = False
    cmdIntBalance.Enabled = False
    cmdPenalBalance.Enabled = False
    txtPrincAmount.Enabled = False
    txtPrincAmount.BackColor = wisGray
    txtTotalIntBalance = 0
    txtRegInt = 0
    txtPenalInt = 0
    txtRegIntBalance = "0.00"
    txtPenalBalance = "0.00"
    txtTotalIntBalance.BackColor = wisGray
    txtRegInt.BackColor = wisGray
    txtPenalInt.BackColor = wisGray
    txtRegIntBalance.BackColor = wisGray
    txtPenalBalance.BackColor = wisGray
    txtMisc.BackColor = wisGray
    txtCustNAme.BackColor = wisGray
    'txtLoanAccNo.BackColor = wisGray
    txtTotAmt.BackColor = wisGray
    txtTotAmt.Enabled = False
    txtTotalIntBalance.Enabled = False
    txtRegInt.Enabled = False
    txtPenalInt.Enabled = False
    'txtIntBalance.Enabled = False
    txtMisc.Enabled = False
    txtCustNAme.Enabled = False
  If FormatCurrency(Val(txtBalance)) = 0 Then
    cmdUndo.Caption = GetResourceString(313) '"&Reopen"
    'cmdUndo.Tag = "Reopen"
    'txtLoanAccNo.Enabled = False
  Else
    'lblLastTransDate.Caption = "Last transaction"
    lblLastTransDate.Caption = GetResourceString(391) & " " & _
    GetResourceString(38, 37)
    txtPrincAmount.BackColor = wisWhite
    cmdRepay.Enabled = True
    cmdIntBalance.Enabled = True
    cmdPenalBalance.Enabled = True
  End If
End If

If FormatField(m_rstLoanMast("LoanClosed")) = 0 Then
    txtRegIntBalance.Enabled = True
    txtRegIntBalance.BackColor = &H80000005
    txtPenalBalance.Enabled = True
    txtPenalBalance.BackColor = &H80000005
    txtTotalIntBalance.Enabled = True
    txtTotalIntBalance.BackColor = &H80000005
    txtRegInt.Enabled = True
    txtRegInt.BackColor = &H80000005
    txtPenalInt.Enabled = True
    txtPenalInt.BackColor = &H80000005
    txtPrincAmount.Enabled = True
    txtPrincAmount.BackColor = &H80000005
    txtMisc.Enabled = True
    txtMisc.BackColor = &H80000005
    txtTotAmt.Enabled = True
    txtTotAmt.BackColor = &H80000005
    grd.Visible = True
    cmdUndo.Caption = GetResourceString(19) & " (" & GetResourceString(391) & ")" '"&Undo(Last)"
End If

m_LoanID = LoanID

End Sub

Private Sub SetKannadaCaption()
Call SetFontToControlsSkipGrd(Me)

lblLoanScheme = GetResourceString(214)  'LOan Scheme
lblLoanAccNo = GetResourceString(58, 60) 'LOan Acc NO
lblCustID = GetResourceString(49, 60) 'Member No
lblCustName = GetResourceString(35) 'Name
lblMemberType.Caption = GetResourceString(101)  'Member type
cmdPhoto.Caption = GetResourceString(415)

cmdLoad.Caption = GetResourceString(3) 'Detail
cmdLoan.Caption = GetResourceString(3) 'Detail

lblSanction = GetResourceString(247)
lblOverDue = GetResourceString(84, 18)
lblLoanAmount = GetResourceString(80, 91)
lblRepaid = GetResourceString(341) 'Repaid Amount

lblTransDate = GetResourceString(38, 37)
lblBalance = GetResourceString(67, 58) '"balance Amount
lblRegInt = GetResourceString(344) 'Reg Interest
lblMiscAmount = GetResourceString(327)
lblLastTransDate = GetResourceString(391, 38, 37)
lblRegIntBalance = GetResourceString(67, 47) 'Interest Balance
lblPenalBalance = GetResourceString(67, 345) 'Interest Balance
lblTotalIntBalance = GetResourceString(52, 67, 47) 'Interest Balance
lblPenalInt = GetResourceString(345) 'Penal Interest
LblTotAmt = GetResourceString(310) 'Penal Interest
lblRepayAmount = GetResourceString(341) 'Repaid Amount

lblParticulars = GetResourceString(39) 'Particulars

cmdPay.Caption = GetResourceString(289) 'Payments
cmdUndo.Caption = GetResourceString(19) & GetResourceString(391) 'Undo Last
cmdRepay.Caption = GetResourceString(20) 'Repay

cmdOk.Caption = GetResourceString(1) 'OK

End Sub

Private Sub ShowTransaction()

Call InitGrid

Dim count As Integer
Dim TransID As Long
Dim transType As wisTransactionTypes

grd.Visible = False
'cmdPrevTrans.Tag = cmdNextTrans.Tag
cmdNextTrans.Tag = m_rstLoanTrans.AbsolutePosition
Do
  With grd
    If m_rstLoanTrans.EOF Then Exit Do
    If TransID < m_rstLoanTrans("TransID") Then
        TransID = m_rstLoanTrans("TransID")
        count = count + 1
        If count > 10 Then Exit Do           'm_rstLoanTrans.MovePrevious
        .Row = count
    End If
    transType = m_rstLoanTrans("TransType")
    If FormatField(m_rstLoanTrans(0)) = "Principal" Then
        .Col = 0: .Text = Format(count, "00")
        .Col = 1: .Text = FormatField(m_rstLoanTrans("TransDate"))
        .Col = 2: .Text = FormatField(m_rstLoanTrans("Particulars"))
        .Col = IIf(transType = wWithdraw Or transType = wContraWithdraw, 3, 4)
        .Text = FormatField(m_rstLoanTrans(5))
        '.Col = 7: .Text = FormatField(m_rstLoanTrans(6))
        .Col = 7: .Text = FormatField(m_rstLoanTrans("Balance"))
    Else
        .Col = 0: .Text = Format(count, "00")
        .Col = 1: .Text = FormatField(m_rstLoanTrans("transDate"))
        .Col = 2: .Text = FormatField(m_rstLoanTrans("Particulars"))
        .Col = 5: .Text = FormatField(m_rstLoanTrans(5))
        .Col = 6: .Text = FormatField(m_rstLoanTrans(6))
    End If
  End With
  
  m_rstLoanTrans.MoveNext
Loop

grd.ScrollBars = flexScrollBarBoth
grd.Visible = True
'If Count < 11 Then grd.Rows = IIf(Count < 5, 6, Count)
cmdNextTrans.Enabled = Not m_rstLoanTrans.EOF

End Sub

Private Function UndoLastTransaction() As Boolean

If m_CustomerID = 0 Or m_LoanID = 0 Then Exit Function

Dim SqlStr As String
Dim rst As Recordset
Dim lastTransID  As Long
Dim Amount As Currency
Dim IntAmount As Currency
Dim PenalIntAmount As Currency
Dim MiscAmount As Currency

Dim InstAmount As Currency
Dim PaidInst As Currency
Dim transType As wisTransactionTypes
Dim IntTransType As wisTransactionTypes
Dim LastIntDate As Date

Dim bankClass As clsBankAcc
Dim TransDate As Date

Dim LoanHeadID As Long

Dim InstBalance() As Currency
Dim InstNo() As Integer
Dim count As Integer

If FormatField(m_rstLoanMast("LastIntDate")) = "" Then
    LastIntDate = vbNull
Else
    LastIntDate = m_rstLoanMast("LastIntDate")
End If

InstAmount = FormatField(m_rstLoanMast("InstAmount"))

'Now Get the Loan Tranacation Id to Delete i.e. the Max Transaction ID
lastTransID = GetLoanMaxTransID(m_LoanID)
If lastTransID < 1 Then Exit Function


Dim SchemeName As String
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
            " Where SchemeID = " & m_rstLoanMast("schemeID")
Call gDbTrans.Fetch(rst, adOpenDynamic)
SchemeName = FormatField(rst("SchemeName"))

LoanHeadID = GetIndexHeadID(SchemeName)


SqlStr = "SELECT Top 1 * FROM LoanTrans " & _
        " WHERE LoanID = " & m_LoanID & _
        " AND TransID = " & lastTransID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Amount = FormatField(rst("Amount"))
    transType = rst("TransType")
    TransDate = rst("TransDate")
End If

SqlStr = "SELECT Top 1 * FROM LoanIntTrans " & _
        " WHERE LoanID = " & m_LoanID & _
        " AND TransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    IntTransType = rst("TransType")
    TransDate = rst("TransDate")
    IntAmount = FormatField(rst("IntAmount"))
    PenalIntAmount = FormatField(rst("PenalIntAmount"))
    MiscAmount = FormatField(rst("MiscAmount"))
End If

'Check With The Amount in Payable/Receivable Field
SqlStr = "SELECT * FROM AmountReceivAble " & _
        " WHERE AccHeadid = " & LoanHeadID & _
        " AND AccID = " & m_LoanID & _
        " AND AccTransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
Dim MiscHeadCount As Integer
Dim DueAmount() As Currency
Dim DueHeadID() As Long
Dim I As Integer
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    MiscAmount = 0
    'Set m_clsReceivable = New clsReceive
    While Not rst.EOF
        ReDim Preserve DueHeadID(MiscHeadCount)
        ReDim Preserve DueAmount(MiscHeadCount)
        MiscAmount = MiscAmount + DueAmount(MiscHeadCount)
        DueAmount(MiscHeadCount) = FormatField(rst("Amount"))
        DueHeadID(MiscHeadCount) = FormatField(rst("DueHeadID"))
        'Call m_clsReceivable.AddHeadAndAmount(Rst("DueHeadID"), Rst("Amount"))
        rst.MoveNext
        MiscHeadCount = MiscHeadCount + 1
    Wend
End If


Dim TotAmount As Currency
If m_rstLoanMast("EMI") Then
    TotAmount = PenalIntAmount + IntAmount + Amount
Else
    TotAmount = IntAmount + Amount + PenalIntAmount + MiscAmount
End If

'Confirm the deletion of transaction
'If MsgBox("Are you sure you wanto delete then transaction of amount " & totAmount _
    , vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
If MsgBox(GetResourceString(583) & " Rs." & TotAmount _
            , vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Function

If transType = wContraDeposit Or transType = wContraWithdraw Or _
            IntTransType = wContraDeposit Or IntTransType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & LoanHeadID & _
            " And Accid = " & m_LoanID & " And  TransID = " & lastTransID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        If ContraClass.UndoTransaction(rst("ContraID"), TransDate) = Success Then _
                UndoLastTransaction = True
        Set ContraClass = Nothing
        Exit Function
    End If
End If

'Fetch the details of Installment table
If FormatField(m_rstLoanMast("InstMode")) <> 0 And Amount > 0 Then
    SqlStr = "SELECT * FROM LoanInst" & _
            " WHERE LoanID = " & m_LoanID & _
            " AND InstBalance < " & InstAmount & _
            " ORDER BY InstDate Desc"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
        count = 0
        Do
            If rst.EOF Then Exit Do
            
            ReDim Preserve InstBalance(count)
            ReDim Preserve InstNo(count)
            InstBalance(count) = FormatField(rst("InstBalance"))
            InstNo(count) = FormatField(rst("InstNo"))
            count = count + 1
            rst.MoveNext
        Loop
        Set rst = Nothing
    End If
End If


Dim InTrans As Boolean
InTrans = gDbTrans.BeginTrans

'Now Delete the Transcation in the loan trans
SqlStr = "DELETE * FROM LoanTrans WHERE " & _
        " LoanID = " & m_LoanID & " AND TransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

'Now Delete the Transcation of interest
SqlStr = "DELETE * FROM LoanIntTrans WHERE " & _
        " LoanID = " & m_LoanID & " AND TransID = " & lastTransID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

'Check With The Amount in Payable/Receivable Field
If Not UndoAmountReceivable(LoanHeadID, m_LoanID, lastTransID) Then GoTo ExitLine

If DateDiff("d", LastIntDate, TransDate) = 0 And _
                (IntTransType = wDeposit Or IntTransType = wContraDeposit) Then
    SqlStr = "UPDATE LoanMaster Set LoanClosed = 0, " & _
                "LastINtDate = NULL WHERE LoanID = " & m_LoanID
Else
    SqlStr = "UPDATE LoanMaster Set LoanClosed = 0" & _
                " WHERE LoanID = " & m_LoanID
End If

gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then GoTo ExitLine

'NOW UPDate the Installment detailas
If (transType = wWithdraw Or transType = wContraWithdraw) And lastTransID <> 1 Then

Else
  If count > 0 Then
    count = 0
    If FormatField(m_rstLoanMast("EMI")) Then Amount = TotAmount
    TotAmount = Amount
    Do
        If count > UBound(InstBalance) Then Exit Do
        If TotAmount <= 0 Then Exit Do
        PaidInst = InstAmount - InstBalance(count)
        If TotAmount <= PaidInst Then
            InstBalance(count) = InstBalance(count) + TotAmount
            TotAmount = 0
        Else
          'If his last payment has affected more than one
          'installment amount then 'Consider it
            InstBalance(count) = InstAmount
            TotAmount = TotAmount - PaidInst
        End If
        If lastTransID = 1 Then InstBalance(count) = InstAmount
        SqlStr = "UPDATE LoanInst Set InstBalance = " & InstBalance(count) & _
                    " WHERE LoanID = " & m_LoanID & _
                    " AND InstNo = " & InstNo(count)
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        count = count + 1
    Loop
  End If
End If

'Varible for the ledger heads
Dim RegIntHeadID As Long
Dim PenalHeadID As Long
Dim MiscHeadId As Long

RegIntHeadID = GetIndexHeadID(SchemeName & " " & GetResourceString(344))
PenalHeadID = GetIndexHeadID(SchemeName & " " & GetResourceString(345))
MiscHeadId = GetIndexHeadID(GetResourceString(327))

Set bankClass = New clsBankAcc
If IntTransType = wDeposit Then
    If Not bankClass.UndoCashDeposits(RegIntHeadID, IntAmount, TransDate) Then GoTo ExitLine
    If Not bankClass.UndoCashDeposits(PenalHeadID, PenalIntAmount, TransDate) Then GoTo ExitLine
    If MiscAmount > 0 Then
        If MiscHeadCount = 0 Then
            If Not bankClass.UndoCashDeposits(MiscHeadId, MiscAmount, TransDate) Then GoTo ExitLine
        Else
         For I = 0 To MiscHeadCount - 1
            If Not bankClass.UndoCashDeposits(DueHeadID(I), _
                             DueAmount(I), TransDate) Then GoTo ExitLine
         Next
        End If
    End If
ElseIf IntTransType = wWithdraw Then
    If Not bankClass.UndoCashWithdrawls(RegIntHeadID, IntAmount, TransDate) Then GoTo ExitLine
    If Not bankClass.UndoCashWithdrawls(PenalHeadID, PenalIntAmount, TransDate) Then GoTo ExitLine
End If
 
If transType = wDeposit Then
    If Not bankClass.UndoCashDeposits(LoanHeadID, Amount, TransDate) Then GoTo ExitLine
ElseIf transType = wWithdraw Then
    If Not bankClass.UndoCashWithdrawls(LoanHeadID, Amount, TransDate) Then GoTo ExitLine
ElseIf Amount Then
    If Amount = (IntAmount + PenalIntAmount + MiscAmount) Then
        If transType <> IntTransType Then
            If Not bankClass.UndoContraTrans(LoanHeadID, RegIntHeadID, _
                IntAmount, TransDate) Then GoTo ExitLine
            If Not bankClass.UndoContraTrans(LoanHeadID, PenalHeadID, _
                PenalIntAmount, TransDate) Then GoTo ExitLine
            If Not bankClass.UndoContraTrans(LoanHeadID, MiscHeadId, _
                MiscAmount, TransDate) Then GoTo ExitLine
        End If
    Else
        MsgBox "No Undo For Contra Transction"
    End If
End If


gDbTrans.CommitTrans
InTrans = False

'MsgBox "Last transaction deleted", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE

UndoLastTransaction = True

ExitLine:
If InTrans Then gDbTrans.RollBack

End Function

Private Sub cmbLoanScheme_Change()
txtLoanAccNo.Locked = False
End Sub

Private Sub cmbLoanScheme_Click()
txtLoanAccNo.Locked = False
End Sub


Private Sub cmdAbn_Click()
If m_LoanID = 0 Then Exit Sub
frmLoanAbn.LoanAccountID = m_LoanID
frmLoanAbn.Show 1
End Sub

Private Sub cmdAddNote_Click()

If m_Notes.ModuleID = 0 Or m_Notes.AccId = 0 Then Exit Sub

Call m_Notes.Show
Call m_Notes.DisplayNote(rtfNote)

End Sub

Private Sub cmdCustName_Click()
'Get the CustomerId
If m_rstLoanMast Is Nothing Then Exit Sub

m_CustomerID = FormatField(m_rstLoanMast("CustomerId"))

Dim FrmLnMast As New frmLoanMaster

With FrmLnMast
     Load FrmLnMast
    .LoadLoan (m_LoanID)
    '.cmdCreate.Visible = False
    .cmdDelete.Visible = False
    .Show vbModal
End With

Unload FrmLnMast
Set FrmLnMast = Nothing
Call tabLoans_Click

End Sub


Private Sub cmdIntBalance_Click()

If Not DateValidate(txtTransDate, "/", True) Then Exit Sub
Dim IntBalance As String
IntBalance = InputBox("Enter the new regular interest balance", "Interest Balance", "0")
If Len(IntBalance) = 0 Then Exit Sub
If Not IsNumeric(IntBalance) Then Exit Sub


gDbTrans.BeginTrans
gDbTrans.SqlStmt = "Update LoanIntTrans Set IntBalance = " & Val(IntBalance) & _
         " WHERE LoanID = " & m_LoanID & " AND TransID = " & _
            "(SElect Top 1 TransID From LoanIntTrans Where LoanID = " & _
            m_LoanID & " And TransDate <= #" & GetSysFormatDate(txtTransDate) & "#" & _
            " Order By TransID Desc)"
 

If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
    txtRegIntBalance = IntBalance
Else
    gDbTrans.RollBack
End If

End Sub

Private Sub cmdLoad_Click()

If Len(Trim(txtCustID)) = 0 Then Exit Sub

Dim CustName As String
Dim memberType As Integer

txtCustNAme.Caption = GetMemberNameCustIDByMemberNum(Trim(txtCustID), m_CustomerID, memberType)

If m_CustomerID < 1 Then Exit Sub

m_ClsCust.LoadCustomerInfo (m_CustomerID)
txtCustNAme = m_ClsCust.CustomerName(m_CustomerID)
Call SetComboIndex(cmbMemberType, ,memberType)
Call LoadCustomerLoans

Exit Sub



'Check For The Customer existance
Dim SqlStr As String
Dim rst As Recordset

SqlStr = "SELECT * FROM MemMaster Where AccNum = " & AddQuotes(txtCustID, True)
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    m_CustomerID = 0
    Exit Sub
Else
    m_CustomerID = FormatField(rst("Customerid"))
End If

RetryLine:

SqlStr = "SELECT * FROM NameTab Where CustomerID = " & m_CustomerID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    m_CustomerID = 0
    SqlStr = "SELECT CustomerID, " & _
        "Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName as Name " & _
        " FROM NameTab ORDER By firstName"

    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
    If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
    Call FillView(m_frmLookUp.lvwReport, rst, True)
    m_CustomerID = 0
    m_frmLookUp.Show vbModal
    If m_CustomerID = 0 Then Exit Sub
    txtCustID = GetMemberNumber(m_CustomerID)
    GoTo RetryLine
    Exit Sub
End If

m_ClsCust.LoadCustomerInfo (m_CustomerID)
txtCustNAme = m_ClsCust.CustomerName(m_CustomerID)
Call LoadCustomerLoans
End Sub

'
Private Sub cmdLoan_Click()

'If txtLoanAccNo.Locked Then Exit Sub
If cmbLoanScheme.ListIndex = -1 Then
    'MsgBox "Select the Loantype"
    MsgBox GetResourceString(89), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
ElseIf Trim(txtLoanAccNo.Text) = "" Then
    'MsgBox "Please specify the loan account NO", vbInformation, wis_MESSAGE_TITLE '
    MsgBox GetResourceString(606), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Dim SqlStr As String
Dim rst As Recordset
Dim AccNum As String
Dim RecFetch As Boolean

AccNum = Trim$(txtLoanAccNo)
Call ClearControls


SqlStr = "SELECT * FROM LoanMaster WHERE AccNum = " & AddQuotes(AccNum, True)
If cmbLoanScheme.ListIndex >= 0 Then _
    SqlStr = SqlStr & " AND SchemeID = " & cmbLoanScheme.ItemData(cmbLoanScheme.ListIndex)

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then RecFetch = True
    
If RecFetch Then
    'MsgBox "There is no Loan in this account", vbInformation
    MsgBox GetResourceString(582), vbInformation
    txtCustNAme.Caption = ""
    
    txtTransDate.Text = ""
    txtTransDate.Tag = txtTransDate.Text
    txtLoanAmount.Caption = ""
    txtSanction.Caption = ""
    txtBalance.Caption = ""
    txtLastTransDate.Caption = ""
    grd.Clear
    Exit Sub
End If

Dim custId As Long
Dim LoanID As Long
If rst.recordCount > 1 Then
    Dim multiLoan As Boolean
    multiLoan = False
    While rst.EOF = False
        If Not FormatField(rst("LoanClosed")) Then
            If (custId > 0) Then multiLoan = True
            custId = FormatField(rst("CustomerID"))
            LoanID = FormatField(rst("LoanID"))
        End If
        rst.MoveNext
    Wend
    If multiLoan Then
        MsgBox "You have more than one loan account with this number." & vbCrLf & "Please enter the member number to get the details"
        Exit Sub
    End If
Else
    custId = FormatField(rst("CustomerID"))
    LoanID = FormatField(rst("LoanID"))
End If


Set rst = Nothing
Dim memberType As Integer
txtCustID = GetMemberNumberAndType(custId, memberType)
m_CustomerID = custId
m_ClsCust.LoadCustomerInfo (m_CustomerID)
txtCustNAme = m_ClsCust.CustomerName(m_CustomerID)
Call SetComboIndex(cmbMemberType, , memberType)
Call LoadCustomerLoans

Dim count As Integer
count = 1
Dim loanFound As Boolean
loanFound = False
Do
    count = count + 1
    If count > tabLoans.Tabs.count Then Exit Do
    If Val(Mid(tabLoans.Tabs(count).Key, 4)) = LoanID Then loanFound = True: Exit Do
Loop
On Error Resume Next
If Not loanFound Then
    'Search for the existance of the loan
    'in a not show list
    loanFound = LoadClosedCustomerLoan(LoanID)
End If

If count < tabLoans.Tabs.count Then tabLoans.Tabs(count).Selected = True
If Not loanFound Then tabLoans.Tabs(count).Selected = True

Err.Clear
End Sub


Private Sub cmdMisc_Click()

If m_LoanID = 0 Then Exit Sub
If m_clsReceivable Is Nothing Then _
            Set m_clsReceivable = New clsReceive
    
m_clsReceivable.Show
    
If m_clsReceivable.TotalAmount Then
    txtMisc.Locked = True
    txtMisc.Value = m_clsReceivable.TotalAmount
Else
    txtMisc.Locked = False
    Set m_clsReceivable = Nothing
End If

End Sub

Private Sub cmdNextTrans_Click()
cmdPrevTrans.Enabled = True
If m_rstLoanTrans.AbsolutePosition <= 1 Then cmdPrevTrans.Enabled = False
Call ShowTransaction
End Sub

'this SubRoutinue procedure will not be used ..
'this is just for checking purpose

Private Sub cmdOk_Click()

    Unload Me
End Sub

Private Sub cmdPay_Click()
Dim FrmPay As New frmLoanPay
With FrmPay
    .LoanAccountID = m_LoanID
     Load FrmPay
    .Show vbModal
End With
RaiseEvent PaymentClicked(m_LoanID)
Call tabLoans_Click
End Sub

Private Sub cmdPenalBalance_Click()

If Not DateValidate(txtTransDate, "/", True) Then Exit Sub
Dim IntBalance As String
IntBalance = InputBox("Enter the new Penal interest balance", "Interest Balance", "0")

If Len(IntBalance) = 0 Then Exit Sub
If Not IsNumeric(IntBalance) Then Exit Sub

gDbTrans.BeginTrans
gDbTrans.SqlStmt = "Update LoanIntTrans Set PenalIntBalance = " & Val(IntBalance) & _
         " WHERE LoanID = " & m_LoanID & " AND TransID = " & _
            "(SElect Top 1 TransID From LoanIntTrans Where LoanID = " & _
            m_LoanID & " And TransDate <= #" & GetSysFormatDate(txtTransDate) & "#" & _
            " Order By TransID Desc)"

If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
    txtPenalBalance = IntBalance
Else
    gDbTrans.RollBack
End If

End Sub

Private Sub cmdPhoto_Click()
    frmPhoto.setAccNo (m_CustomerID)
    If (m_CustomerID > 0) Then frmPhoto.Show vbModal

End Sub

Private Sub cmdPrevTrans_Click()
On Error Resume Next

m_rstLoanTrans.MoveFirst
m_rstLoanTrans.Move Val(cmdNextTrans.Tag)
m_rstLoanTrans.Move -20 'Val(cmdNextTrans.Tag)

If m_rstLoanTrans.AbsolutePosition < 2 Then
    m_rstLoanTrans.MoveFirst
    cmdPrevTrans.Enabled = False
Else
    cmdPrevTrans.Enabled = True
End If
Call ShowTransaction

End Sub

Private Sub cmdPrint_Click()
    If m_frmPrintTrans Is Nothing Then _
      Set m_frmPrintTrans = New frmPrintTrans
    
    m_frmPrintTrans.Show vbModal

End Sub

Private Sub cmdRepay_Click()

If Val(txtTotAmt.Tag) Then
    If txtTotAmt.Value <> (txtTotalIntBalance + txtRegInt + txtPenalInt + txtMisc + txtPrincAmount) Then
        MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtTotAmt
        Exit Sub
    End If
End If

RaiseEvent RePaymentClicked(m_LoanID)

If LoanRepayment Then
    'MsgBox "repayment made succefully"
     MsgBox GetResourceString(586), vbInformation, wis_MESSAGE_TITLE
     'Check wether LoanBalance=0 or Not if bal is null then reopen the form
     Call tabLoans_Click
End If


End Sub

Private Function LoanRepayment() As Boolean

'VALIDATE THE CONTROLS
'Validate the date of transacdtion
If Not DateValidate(txtTransDate, "/", True) Then
    'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Function
End If

If txtTotAmt = 0 Then
    'MsgBox "Invalid amount specified ", vbInformation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtRegInt
    Exit Function
End If

'Get the Particulars
If Trim$(cmbParticulars.Text) = "" Then
    'MsgBox "Transaction particulars not specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(621), vbExclamation, gAppName & " - Error"
    cmbParticulars.SetFocus
    Exit Function
End If

Dim SqlStr As String
Dim rst As Recordset
Dim TransDate As Date
Dim InstType As wisInstallmentTypes

TransDate = GetSysFormatDate(txtTransDate)
InstType = FormatField(m_rstLoanMast("InstMode"))
Dim RstInst As Recordset

'Compare the date w.r.t. Last TransDate
If DateDiff("d", TransDate, GetLoanLastTransDate(m_LoanID)) > 0 Then
    'Earlier Date Specified
    MsgBox GetResourceString(572), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTransDate
    Exit Function
End If

Dim TransID As Long
'Get the MAx transaction Id
TransID = GetLoanMaxTransID(m_LoanID)

'VALIDATION COMPLETES HERE INSERTION CODE STARTS HERE
Dim transType As wisTransactionTypes
Dim Amount As Currency
Dim PrincAmount As Currency
Dim IntAmount As Currency
Dim C_IntAmount As Currency
Dim c_PenalInt As Currency
Dim PenalIntAmount As Currency
Dim Balance As Currency
Dim IntBalance As Currency
Dim PrevIntBalance As Currency

Dim PenalBalance As Currency
Dim MiscAmount As Currency
Dim RetMsg As VbMsgBoxResult

'txtTotalIntBalance

Amount = Val(txtTotAmt)
Balance = Val(txtBalance)
IntAmount = txtRegInt
PenalIntAmount = Val(txtPenalInt)
PrevIntBalance = Val(txtRegIntBalance) + Val(txtPenalBalance)
IntBalance = txtTotalIntBalance
PenalBalance = Val(txtPenalBalance)
MiscAmount = txtMisc
'Amount = Amount - MiscAmount

'Get the Loan Balance
gDbTrans.SqlStmt = "SELECT SUM(Amount) FROM LoanTrans " & _
                " WHERE LoanID = " & m_LoanID & _
                " AND (TransType = " & wWithdraw & _
                    " OR TransType = " & wContraWithdraw & ")"

Balance = 0
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Balance = FormatField(rst(0))

gDbTrans.SqlStmt = "SELECT SUM(Amount) FROM LoanTrans " & _
                " WHERE LoanID = " & m_LoanID & _
                " AND (TransType = " & wDeposit & _
                " OR TransType = " & wContraDeposit & ")"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then Balance = Balance - Val(FormatField(rst(0)))

'Remove the IntPaid
'Now Calulate the Interest & Principal From Repaying amount
PrincAmount = Amount - IntAmount - MiscAmount - PenalIntAmount - IntBalance
If PrincAmount < 0 Then
    'MsgBox "invalid amount specified"
    MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Now Get the Penla iNterest & Regular interest as on Transaction date
Dim ClassLoan As New clsLoan
C_IntAmount = ClassLoan.RegularInterest(m_LoanID, , TransDate)
c_PenalInt = ClassLoan.PenalInterest(m_LoanID, , TransDate)
Set ClassLoan = Nothing


Dim SecMsg As String

'First check whether he is pay any interest amount
'First Case If He is paying the interest amount then
If IntAmount + PenalIntAmount + IntBalance > 0 Then
    'if he has the interest balance
    'without paying it , He is paying the interet amount  then
    'do not allow him to make the transaction
    If IntAmount + PenalIntAmount > 0 And IntBalance < PrevIntBalance Then
        'MsgBox "The amount he is paying is lesser then his interest balance " & _
            vbYesNo, wis_MESSAGE_TITLE)
        MsgBox GetResourceString(670) & _
                    vbYesNo + vbQuestion, wis_MESSAGE_TITLE
        Exit Function
    End If
    'The Amount He is paying less than the Balance interest
    'If IntAmount < IntBalance Then
    If IntBalance < PrevIntBalance Then
        'RetMsg = MsgBox("The amount he is paying is lesser then his interest balance " & _
            "Do you want to continue", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(670) & _
            GetResourceString(541), vbYesNo + vbQuestion, wis_MESSAGE_TITLE)
        If RetMsg = vbNo Then Exit Function
    'ElseIf IntAmount + PenalIntAmount < C_IntAmount + c_PenalInt + IntBalance + PenalBalance Then
    ElseIf IntAmount + PenalIntAmount + IntBalance < C_IntAmount + c_PenalInt + PrevIntBalance Then
        
        SecMsg = GetResourceString(IIf(PrincAmount, 541, 669))
        'RetMsg = MsgBox("The amount he is paying is lesser than his interest till date" & _
            "Do you want keep the difference as balance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(668) & vbCrLf & SecMsg, _
            IIf(PrincAmount, vbYesNo, vbYesNoCancel) + vbQuestion, wis_MESSAGE_TITLE)
        If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then
            If PrincAmount Then Exit Function
            IntBalance = PrevIntBalance
            C_IntAmount = IntAmount '- IntBalance
            c_PenalInt = PenalIntAmount '- PenalBalance
        End If
    
    'ElseIf IntAmount + PenalIntAmount > C_IntAmount + c_PenalInt + IntBalance + PenalBalance Then
    'The interest he is paying more than the actual calcualated
    'Than Ask for the to take the extra interest amount as Advance Interest
    ElseIf IntAmount + PenalIntAmount > C_IntAmount + c_PenalInt Then
    'In some banks/Soceities they collect the interest till 31st March
    'This case for such banks
        SecMsg = GetResourceString(IIf(PrincAmount, 541, 667))
        'RetMsg = MsgBox("The amount he is paying is more than his interest till date" & _
            "Do you want keep the difference as blance interest ", vbYesNo, wis_MESSAGE_TITLE)
        RetMsg = MsgBox(GetResourceString(666) & vbCrLf & SecMsg, _
             IIf(PrincAmount, vbYesNo, vbYesNoCancel) + vbQuestion, wis_MESSAGE_TITLE)
        If RetMsg = vbCancel Then Exit Function
        If RetMsg = vbNo Then
            If PrincAmount Then Exit Function
            C_IntAmount = IntAmount ' - IntBalance
            c_PenalInt = PenalIntAmount ' - PenalBalance
        End If
    End If
    
End If

Balance = Balance - PrincAmount
Set RstInst = Nothing
'IntBalAmount = IntBalance + c_IntAmount + c_PenalInt - IntAmount - PenalIntAmount
'IntBalance = IntBalance + C_IntAmount - IntAmount
'PenalBalance = PenalBalance + c_PenalInt - PenalIntAmount

'Calculate the interest balance whic will cary forward
IntBalance = Val(txtRegIntBalance) + C_IntAmount - IntAmount
PenalBalance = Val(txtPenalBalance) + c_PenalInt - PenalIntAmount

If PrincAmount Then IntBalance = 0: PenalBalance = 0

If InstType <> Inst_No Then
    SqlStr = "SELECT * FROM LoanInst Where LoanID = " & m_LoanID & _
                " AND InstBalance > 0 ORDER BY InstDate"
    gDbTrans.SqlStmt = SqlStr
    Call gDbTrans.Fetch(RstInst, adOpenStatic)
End If

'Varible for the ledger heads
Dim LoanHeadID As Long
Dim RegIntHeadID As Long
Dim PenalHeadID As Long
Dim MiscHeadId As Long

Dim UserID As Integer
Dim bankClass As clsBankAcc
Dim AccType As Long
Dim SchemeName As String
Dim SchemeNameEnglish As String

gDbTrans.SqlStmt = "SELECT SchemeName,SchemeNameEnglish From LoanScheme " & _
                " Where SchemeID = " & m_rstLoanMast("SchemeID")
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
SchemeName = FormatField(rst("SchemeName"))
SchemeNameEnglish = FormatField(rst("SchemeNameEnglish"))
AccType = wis_Loans + m_rstLoanMast("SchemeID")
If bankClass Is Nothing Then Set bankClass = New clsBankAcc

gDbTrans.BeginTrans

LoanHeadID = bankClass.GetHeadIDCreated(SchemeName, SchemeNameEnglish, parMemberLoan, 0, AccType)

UserID = gCurrUser.UserID

'Now Chek the Any receivable Amount

If GetReceivAbleAmount(LoanHeadID, m_LoanID) Then
    If m_clsReceivable Is Nothing Then
        If MsgBox("This customer has due in Receivable" & _
            vbCrLf & GetResourceString(541), _
            vbYesNo, wis_MESSAGE_TITLE) = vbNo Then Exit Function
    End If
End If

Set bankClass = New clsBankAcc
Dim InTrans As Boolean


InTrans = True
TransID = TransID + 1

If PrincAmount > 0 Then
    LoanHeadID = bankClass.GetHeadIDCreated(SchemeName, SchemeNameEnglish, parMemberLoan, 0, AccType)
    
    transType = wDeposit
    SqlStr = "INSERT INTO LoanTrans(LoanId,TransDate,TransId," _
            & " TransType,Amount,Balance,Particulars,UserId)" _
            & " VALUES ( " _
            & m_LoanID & "," _
            & " #" & TransDate & "#," _
            & TransID & "," _
            & transType & "," _
            & PrincAmount & "," _
            & Balance & "," _
            & AddQuotes(cmbParticulars.Text) & "," _
            & UserID & ")"
    
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    'Update it to ledger head
    If Not bankClass.UpdateCashDeposits(LoanHeadID, PrincAmount, TransDate) Then GoTo ExitLine
    
End If

'Then Insert the records into LoanInt Trans
If IntAmount + PenalIntAmount + MiscAmount > 0 Then
    Dim engHeadName As String
    transType = wDeposit
    If Len(SchemeNameEnglish) > 0 Then engHeadName = SchemeNameEnglish & " " & LoadResString(344)
    RegIntHeadID = bankClass.GetHeadIDCreated(SchemeName & " " & GetResourceString(344), engHeadName, _
                 parMemLoanIntReceived, 0, AccType)
    If Len(SchemeNameEnglish) > 0 Then engHeadName = SchemeNameEnglish & " " & LoadResString(345)
    PenalHeadID = bankClass.GetHeadIDCreated(SchemeName & " " & GetResourceString(345), engHeadName, _
                parMemLoanPenalInt, 0, AccType)
    
    MiscHeadId = bankClass.GetHeadIDCreated(GetResourceString(327), LoadResString(327), _
                parBankIncome, 0)
                
    SqlStr = " INSERT INTO LoanIntTrans(LoanId,TransDate," _
            & " TransId,TransType,IntAmount,PenalIntAmount," _
            & " MiscAmount,IntBalance,PenalIntBalance,UserID) " _
            & " VALUES ( " _
            & m_LoanID & "," _
            & " #" & TransDate & "#," _
            & TransID & "," _
            & transType & "," _
            & IntAmount & "," _
            & PenalIntAmount & "," _
            & MiscAmount & "," _
            & IntBalance & "," _
            & PenalBalance & "," _
            & UserID & ")"

    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    If IntAmount > 0 Or PenalIntAmount > 0 Then
        SqlStr = "UPDATE LoanMaster Set LastIntDate = #" & TransDate & "#" & _
                    " WHERE LoanId = " & m_LoanID
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
    End If
    'Update the amounts to the ledgr head
    'Update Regular interest
    If IntAmount > 0 Then
        If Not bankClass.UpdateCashDeposits(RegIntHeadID, IntAmount, TransDate) Then GoTo ExitLine
    End If
    'Update penal interest
    If PenalIntAmount > 0 Then
        If Not bankClass.UpdateCashDeposits(PenalHeadID, PenalIntAmount, TransDate) Then GoTo ExitLine
    End If
    'Update misceleneous amount
    If MiscAmount > 0 Then
        If m_clsReceivable Is Nothing Then
            If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                        MiscAmount, TransDate) Then GoTo ExitLine
        Else
            Call m_clsReceivable.GetHeadAndAmount(MiscHeadId, 0, MiscAmount)
            Do While MiscHeadId > 0
                If MiscHeadId = 0 Then Exit Do
                
                If Not bankClass.UpdateCashDeposits(MiscHeadId, _
                                 MiscAmount, TransDate) Then GoTo ExitLine
                'Do the same transaction in Amount Receivble table
                If Not RemoveFromAmountReceivable(LoanHeadID, m_LoanID, TransID, _
                            TransDate, MiscAmount, MiscHeadId) Then GoTo ExitLine
                
                Call m_clsReceivable.NextHeadAndAmount(MiscHeadId, 0, MiscAmount)
            Loop
        End If
    End If
End If

'If Loan HAs Instalments then
'Update the Installment details To The LoanInst table

If Not RstInst Is Nothing And PrincAmount > 0 Then
    'If FormatField(m_rstLoanMast("EMI")) = True Then
        'MsgBox "Doubts has to be Cleared about bifuracation of amount"
        Amount = PrincAmount 'Val(txtRepayAmount)
    'End If
    
    Dim InstAmount As Currency
    Dim InstBalance As Currency
    Dim InstPayment As Currency
    Dim InstNo As Integer
    Do
        If RstInst.EOF Then Exit Do
        If Amount <= 0 Then Exit Do
        InstAmount = FormatField(RstInst("InstBalance"))
        InstNo = FormatField(RstInst("InstNo"))
        If InstAmount > Amount Then
            InstPayment = InstAmount
            InstBalance = InstAmount - Amount
            Amount = 0
            PrincAmount = PrincAmount - InstAmount
        Else
            Amount = Amount - InstAmount
            'InstAmount = PrincAmount
            InstBalance = 0
            'InstBalance = InstAmount - PrincAmount
            'PrincAmount = 0
            
        End If
        SqlStr = "UPDATE LoanInst  Set InstBalance = " & InstBalance & _
                ", PaidDate = #" & TransDate & "#" & _
                " WHERE LoanID = " & m_LoanID & _
                " AND InstNo = " & InstNo
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        
        RstInst.MoveNext
    Loop
End If

'If Balance - PrincAmount = 0 Then
If Balance = 0 Then
    SqlStr = "UPDATE LoanMaster SET LoanClosed = 1 " & _
        " WHERE LoanID = " & m_LoanID
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo ExitLine

    'Reduce the amount of installment
    SqlStr = "UPDATE LoanMaster SET InstAmount = 0 " & _
        " WHERE LoanID = " & m_LoanID
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
End If

gDbTrans.CommitTrans
InTrans = False

'Now writ the pariculars to the File
Call WriteParticularstoFile(cmbParticulars.Text, App.Path & "\Loan.ini")
'Now Load the pariculars from the File
Call LoadParticularsFromFile(cmbParticulars, App.Path & "\Loan.ini")


LoanRepayment = True
Exit Function

ExitLine:

    If InTrans Then gDbTrans.RollBack
    'MsgBox "Unable to repay the loan", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
            
End Function

Private Function CheckForDueAmount() As Currency

'Now check whether he has any amount due
'like Insurence amount, Vehicle insusrence,Abn ,EP, etc
'First Get the Max TransID From this acount
Dim TransID As Integer
Dim AccType As Long
Dim SchemeName As String
Dim LoanHeadID As Long
Dim rst As Recordset

gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
                " Where SchemeID = " & m_rstLoanMast("SchemeID")
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
SchemeName = FormatField(rst("SchemeName"))
AccType = wis_Loans + m_rstLoanMast("SchemeID")

Dim bankClass As clsBankAcc
If bankClass Is Nothing Then Set bankClass = New clsBankAcc
LoanHeadID = GetHeadID(SchemeName, parMemberLoan)
If LoanHeadID = 0 Then Exit Function

'Now Check whether there is any transction
'in Amount receivable & whs transid is moter than TransID
TransID = GetReceivAbleAmountID(LoanHeadID, m_LoanID)
If TransID Then 'There is AMount due int
    Dim rstTemp As Recordset
    gDbTrans.SqlStmt = "Select * From AmountReceivAble" & _
            " Where AccHeadID = " & LoanHeadID & _
            " And AccID = " & m_LoanID '& _
            " And TransID > (Select Max(TransID) From" & _
                " AmountReceivAble Where AccHeadID = " & LoanHeadID & _
                " And AccID = " & m_LoanID & _
                " AND Balance = 0" & ")"
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        Set m_clsReceivable = New clsReceive
        While Not rstTemp.EOF
            Call m_clsReceivable.AddHeadAndAmount(rstTemp("DueHeadID"), rstTemp("Amount"))
            rstTemp.MoveNext
        Wend
        With txtMisc
            .Value = m_clsReceivable.TotalAmount * -1
            .Locked = True
            CheckForDueAmount = m_clsReceivable.TotalAmount
        End With
    End If
End If

End Function

Private Sub cmdTransDate_Click()
txtTransDate.Tag = txtTransDate.Text
With Calendar
    If DateValidate(txtTransDate, "/", True) Then
    .selDate = txtTransDate
    Else
    .selDate = gStrDate
    End If
    .Left = Left + fra2.Left + fraRepay.Left + Me.cmdTransDate.Left
    .Top = Top + fra2.Top + fraRepay.Top + cmdTransDate.Top - .Height / 2
    .Show vbModal
    txtTransDate = .selDate
End With

If DateValidate(txtTransDate, "/", True) Then
    Call tabLoans_Click
End If
txtTransDate.Tag = txtTransDate.Text

End Sub

Private Sub cmdUndo_Click()

Dim nRet As Integer
If cmdUndo.Caption = GetResourceString(19) & " (" & GetResourceString(391) & ")" Then 'Undo Last
'If cmdUndo.Caption = GetResourceString(5) Then   '"Undo
   Me.MousePointer = vbHourglass
   Call UndoLastTransaction

ElseIf cmdUndo.Caption = GetResourceString(313) Then  '"&Reopen" Then
   
   'nRet = MsgBox("Are you sure to reopen this loan ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(538), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   Me.MousePointer = vbHourglass
   
   gDbTrans.SqlStmt = "UpDate LoanMaster set LoanClosed = 0 where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to reopen the loan"
      MsgBox GetResourceString(536), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.CommitTrans
   
   'MsgBox "Account reopened succefully"
   MsgBox GetResourceString(522), vbInformation, wis_MESSAGE_TITLE
   Call tabLoans_Click
   
   'set the Undo caption.
   cmdUndo.Caption = GetResourceString(19) & " (" & GetResourceString(391) & ")"
   'cmdUndo.Caption = "Undo (Last)"
   lblLastTransDate.Caption = GetResourceString(391) & " (" & GetResourceString(38) & " (" & GetResourceString(37) & ")"
   Me.MousePointer = vbDefault
   Exit Sub
ElseIf cmdUndo.Caption = GetResourceString(14) Then  '"Delete
   'nRet = MsgBox("Are you sure to delete this accoutnt ?", vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   nRet = MsgBox(GetResourceString(539), vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
   If nRet = vbNo Then Exit Sub
   gDbTrans.BeginTrans
   Me.MousePointer = vbHourglass
   gDbTrans.SqlStmt = "Delete * From Loanmaster Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete the loan"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
   gDbTrans.SqlStmt = "Delete * From LoanMaster Where LoanId = " & m_LoanID
   If Not gDbTrans.SQLExecute Then
      'MsgBox "Unable to delete"
      MsgBox GetResourceString(532), vbExclamation + vbCritical, wis_MESSAGE_TITLE
      gDbTrans.RollBack
      Exit Sub
   End If
    
   gDbTrans.CommitTrans
   'MsgBox "Account reopened succefully"
   MsgBox GetResourceString(730), vbInformation, wis_MESSAGE_TITLE
   Me.MousePointer = vbDefault
   'Exit Sub
End If

    Call tabLoans_Click
    Me.MousePointer = vbNormal

End Sub

Private Sub Form_Load()

cmdPrevTrans.Picture = LoadResPicture(101, vbResIcon)
cmdNextTrans.Picture = LoadResPicture(102, vbResIcon)
cmdPrint.Picture = LoadResPicture(120, vbResBitmap)
cmdPhoto.Enabled = Len(gImagePath)

Me.Caption = Me.Caption
Call CenterMe(Me)
Call SetKannadaCaption
txtCustNAme.FONTSIZE = txtCustNAme.FONTSIZE + 1

'Load The instalment types
txtTransDate = gStrDate
txtTransDate.Tag = txtTransDate.Text
tabLoans.Tabs.Clear

'Now Load The Loan schemes
cmbLoanScheme.Clear
Call LoadLoanSchemes(cmbLoanScheme)

Call LoadMemberTypes(cmbMemberType)

Dim InstType As wisInstallmentTypes

Set m_ClsCust = New clsCustReg

If gOnLine Then
    txtTransDate.Locked = True
    cmdTransDate.Enabled = False
End If

'Now Load the pariculars from the File
Call LoadParticularsFromFile(cmbParticulars, App.Path & "\Loan.ini")

End Sub


Private Sub Form_Unload(cancel As Integer)
Call ClearControls
Set m_ClsCust = Nothing
gWindowHandle = 0
RaiseEvent WindowClosed

End Sub


Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub

Private Sub m_frmPrintTrans_DateClick(StartIndiandate As String, EndIndianDate As String)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim rst As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim TransID As Long
Dim lastPrintRow As Integer
Const HEADER_ROWS = 2
Dim curPrintRow As Integer

'1. Fetch last print row from sb master table.
'First get the last printed txnID From the SbMaster
SqlStr = "SELECT LastPrintID, LastPrintRow,Name,B.AccNum,SchemeName,A.MemberNum  From  QryMemName A,LoanMaster B,LoanScheme C " & _
  " Where a.CustomerID = B.CustomerID And B.SchemeID = c.SchemeID And LoanID = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintRow = IIf(IsNull(metaRst("LastPrintrow")), 0, metaRst("LastPrintrow"))

'2. count how many records are present in the table between the two given dates
    SqlStr = "SELECT count(*) From LoanTrans WHERE LoanId = " & m_LoanID
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
   
' If there are no records to print, since the last printed txn,
' display a message and exit.
    If (rst(0) = 0) Then
       MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    'Print [or don't print] header part
    'lastPrintRow = IIf(IsNull(Rst("LastPrintRow")), 0, Rst("LastPrintRow"))
    If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE - 1) Then
    'clsPrint.newPage
    clsPrint.isNewPage = True
    
End If
    
'3. Getting matching records for passbook printing

SqlStr = "SELECT LoanId,TransDate,TransId,TransType, 0 as IntAmount, 0 as PenalIntAmount,Amount,Balance, 'PRINCIPAL' as TableType " & _
    " FROM LoanTrans WHERE LoanId = " & m_LoanID & _
    " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
    " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
    " UNION " & _
    " SELECT LoanId,TransDate,TransId,TransType,IntAmount as Amount,PenalIntAmount, MiscAmount, IntBalance, 'INTEREST' as TableType " & _
    " FROM LoanIntTrans WHERE LoanId = " & m_LoanID & _
    " AND TransDate >= #" & GetSysFormatDate(StartIndiandate) & "#" & _
    " AND TransDate <= #" & GetSysFormatDate(EndIndianDate) & "#" & _
    " ORDER BY TransID, TransType"
    

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Set clsPrint = New clsTransPrint

'Printer.PaperSize = 9
Printer.Font = "Courier New"
Printer.FONTSIZE = 9
With clsPrint
   ' .Header = gCompanyName & vbCrLf & vbCrLf & m_CustReg.FullName
    .Cols = 6
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .ColHeader(1) = GetResourceString(37) 'Date
    '.ColWidth(1) = 5: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(2) = 12: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(3) = 13: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(4) = 10: .COlHeader(4) = GetResourceString(344) 'Interest
    '.ColWidth(5) = 8: .COlHeader(5) = GetResourceString(345) 'Interest
    '.ColWidth(6) = 15: .COlHeader(6) = GetResourceString(42) 'Balance
    
        If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
    End If
        
    
     ' column widths for printing txn rows.
        .ColWidth(0) = 8
        .ColWidth(1) = 5
        .ColWidth(2) = 8
        .ColWidth(3) = 9
        .ColWidth(4) = 8
        .ColWidth(5) = 10
        .ColWidth(6) = 6
         
         
  Dim bHeaderPrinted As Boolean
  bHeaderPrinted = False
  
  '''
  Dim Receipt As Currency
  Dim payment As Currency
  Dim IntAmount As Currency
  Dim MicAmount As Currency
  Dim TransDate As String
  Dim PenalInt As Currency
  Dim Balance As Currency
  Dim Total As Currency
        
        
    
    While Not rst.EOF
      If (bHeaderPrinted = False) Then
            Printer.CurrentX = 1000
            
            Printer.Font.name = gFontName
            Printer.Font.Size = Printer.Font.Size + 2
            Printer.CurrentY = 5000 - (Printer.TextWidth(gCompanyName) / 2)
            Printer.Font.Bold = True
            Printer.Print gCompanyName
            Printer.Font.Bold = False
            Printer.Font.Size = Printer.Font.Size - 2
            
            Printer.CurrentX = 1000
            Printer.CurrentY = Printer.CurrentY + 50
            Printer.Print (FormatField(metaRst("Name")))
            Printer.CurrentX = 9000
            Printer.Print FormatField(metaRst("SchemeName")) & " " & GetResourceString(58, 60); ":" & FormatField(metaRst("AccNum"))
            'Printer.CurrentX = 9000
            Printer.CurrentY = 8000
            Printer.Print GetResourceString(79, 60); ":" & FormatField(metaRst("AccNum"))
            
                '.printHead
                .isNewPage = False
                bHeaderPrinted = True
        End If
        
    If TransID <> 0 And TransID <> FormatField(rst("TransID")) Then
         .ColText(0) = TransDate
         .ColText(1) = Receipt
         .ColText(2) = payment
         .ColText(3) = IntAmount
         .ColText(4) = PenalInt
         .ColText(5) = Balance
         .ColText(6) = Total
         
        ' Debug.Print Receipt & " | " & payment & " | " & IntAmount
         .PrintRows
         
         ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE1) Then
        
            ' since we have to print now in a new page,
            ' we need to print the header.
            ' So, set columns widths for header.
            
            .newPage
           ' MsgBox "plz insert new page"
            curPrintRow = 1
        End If
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    
    TransID = FormatField(rst("TransID"))
        
        If rst("TableType") = "INTEREST" Then
        
            TransDate = FormatField(rst("TransDate"))
                          
            'MicAmount = FormatField(Rst("Amount"))
            IntAmount = FormatField(rst("IntAmount"))
            PenalInt = FormatField(rst("PenalIntAmount"))
            'Total = IntAmount + PenalInt + Receipt
            
        ElseIf rst("TableType") = "PRINCIPAL" Then
            TransDate = FormatField(rst("TransDate"))
            
            If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
                   payment = FormatField(rst("Amount"))
            Else
                   Receipt = FormatField(rst("Amount"))
            End If
            
               Balance = FormatField(rst("Balance"))
               Total = IntAmount + PenalInt + payment
        
        ' Debug.Print TransDate & " | " & Receipt & "|" & payment & " | " & IntAmount & "|" & PenalInt & "|" & Balance & "|" & Total
        
        End If
           
       ' .PrintRows
        
        
        rst.MoveNext
       
    Wend
    If TransID > 0 Then
         .ColText(0) = TransDate
         .ColText(1) = Receipt
         .ColText(2) = payment
         .ColText(3) = IntAmount
         .ColText(4) = PenalInt
         .ColText(5) = Balance
         .ColText(6) = Total
         
        ' Debug.Print Receipt & " | " & payment & " | " & IntAmount
         
         .PrintRows
         
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    .newPage
End With
Printer.EndDoc


Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing
            
'Now Update the Last Print Id to the master
SqlStr = "UPDATE LoanMaster set LastPrintRow = " & curPrintRow - 1 & _
        " Where LoanId = " & m_LoanID
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If
End Sub
    

Private Sub m_frmPrintTrans_TransClick(bNewPassbook As Boolean)
Dim clsPrint As clsTransPrint
Dim SqlStr As String
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim metaRst As ADODB.Recordset
Dim lastPrintId, lastPrintRow As Integer
Const HEADER_ROWS = 2
Dim curPrintRow As Integer

'First get the last printed txnId and last printed row From the loanMaster
SqlStr = "SELECT  LastPrintId, LastPrintRow From LoanMaster WHERE LoanId = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(metaRst, adOpenDynamic) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
Set clsPrint = New clsTransPrint
lastPrintId = IIf(IsNull(metaRst("LastPrintId")), 0, metaRst("LastPrintId"))

' count how many records are present in the table, after the last printed txn id
SqlStr = "SELECT count(*) From LoanTrans WHERE LoanId = " & m_LoanID & " AND TransID > " & lastPrintId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(676), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

' Print the first page of passbook, if newPassbook option is chosen.
If bNewPassbook Then
    clsPrint.printPassbookPage1 wis_Loans, m_LoanID
    
    SqlStr = "UPDATE LoanMaster set LastPrintId = LastPrintId - " & m_frmPrintTrans.cmbRecords.Text & _
        ", LastPrintRow = 0 Where LoanId = " & m_LoanID
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
    Else
        gDbTrans.CommitTrans
    End If
    Exit Sub
End If

' If there are no records to print, since the last printed txn,
' display a message and exit.
If (rst(0) = 0) Then
    MsgBox "There are no transactions available for printing."
    Exit Sub
End If

' Fetch records for txns that have been created after lasttxnId.

SqlStr = "SELECT Amount, 0 as IntAmount, 0 as PenalIntAmount, Balance, TransType, TransID,TransDate, 'Principal' as TableType From LoanTrans " & _
    " WHERE LoanId = " & m_LoanID & _
    " ORDER BY TransID"

SqlStr = SqlStr & " UNION " & "SELECT MiscAmount as Amount, IntAmount, PenalIntAmount, IntBalance as Balance, TransType,TransID, TransDate, 'Interest' as TableType From LoanIntTrans " & _
    " WHERE LoanId = " & m_LoanID & _
        " ORDER BY TransID"

gDbTrans.SqlStmt = SqlStr '"SELECT * From LoanTrans WHERE LoanId = " & m_LoanID & " AND TransID > " & lastPrintId
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    MsgBox GetResourceString(278), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


 clsPrint.isNewPage = True
'Print [or don't print] header part
lastPrintRow = IIf(IsNull(metaRst("LastPrintRow")), 0, metaRst("LastPrintRow"))
If (lastPrintRow < 1 Or lastPrintRow > wis_ROWS_PER_PAGE1 - 1) Then
   ' clsPrint.newPage
    clsPrint.isNewPage = True
    
End If

Set clsPrint = New clsTransPrint
'Printer.PaperSize = 9
Printer.Font.name = gFontName
Printer.Font.Size = 12 'gFontSize
With clsPrint
    .Header = gCompanyName & vbCrLf & vbCrLf & m_ClsCust.FullName
    .Cols = 6
    '.ColWidth(0) = 10: .COlHeader(0) = GetResourceString(37) 'Date
    '.ColWidth(1) = 8: .ColHeader(1) = GetResourceString(37) 'Date
    '.ColWidth(1) = 20: .COlHeader(2) = GetResourceString(39) 'Particulars
    '.ColWidth(2) = 10: .COlHeader(3) = GetResourceString(276) 'Debit
    '.ColWidth(3) = 10: .COlHeader(4) = GetResourceString(277) 'Credit
    '.ColWidth(4) = 8: .COlHeader(4) = GetResourceString(344) 'Interest
    '.ColWidth(5) = 8: .COlHeader(5) = GetResourceString(345) 'Interest
    '.ColWidth(6) = 15: .COlHeader(6) = GetResourceString(42) 'Balance
        'TransID = Rst("TransID")
     If (lastPrintRow >= 1 And lastPrintRow <= wis_ROWS_PER_PAGE1) Then
        ' Print as many blank lines as required to match the correct printable row
        Dim count As Integer
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
        curPrintRow = lastPrintRow + 1
    Else
        curPrintRow = 1
        For count = 0 To (HEADER_ROWS + lastPrintRow)
            Printer.Print ""
        Next count
    End If
    
    ' column widths for printing txn rows.
        .ColWidth(0) = 10
        .ColWidth(1) = 6
        .ColWidth(2) = 8
        .ColWidth(3) = 8
        .ColWidth(4) = 8
        .ColWidth(5) = 12
        .ColWidth(6) = 7
      
  Dim bHeaderPrinted As Boolean
  bHeaderPrinted = False
  
  '''
  Dim Receipt As Currency
  Dim payment As Currency
  Dim IntAmount As Currency
  Dim MicAmount As Currency
  Dim TransDate As String
  Dim PenalInt As Currency
  Dim Balance As Currency
  Dim Total As Currency
  Dim PenalInt1 As String
  '''
  While Not rst.EOF
        If .isNewPage Then
            If (bHeaderPrinted = False) Then
               ' .printHead
                .isNewPage = False
                bHeaderPrinted = True
           End If
        Else
            '.printHead
            '.isNewPage = False
        End If

    If TransID <> 0 And TransID <> FormatField(rst("TransID")) Then

         .ColText(0) = TransDate
         .ColText(1) = Receipt
         .ColText(2) = payment
         .ColText(3) = IntAmount
         .ColText(4) = PenalInt
         .ColText(5) = Balance
         .ColText(6) = Total
         
            
         Debug.Print Receipt & " | " & payment & " | " & IntAmount
              
        If Receipt <> 0 Then
         .ColText(1) = Receipt
        Else
         .ColText(1) = " "
        End If
           
        If payment <> 0 Then
         .ColText(2) = payment
        Else
         .ColText(2) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(3) = IntAmount
        Else
         .ColText(3) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(4) = PenalInt
        Else
         .ColText(4) = " "
        End If
      
        If Total <> 0 Then
         .ColText(6) = Total
        Else
         .ColText(6) = " "
        End If
         
         Debug.Print Receipt & " | " & PenalInt & " | " & IntAmount
         .PrintRows
         
         ' Increment the current printed row.
        curPrintRow = curPrintRow + 1
        If (curPrintRow > wis_ROWS_PER_PAGE1) Then
        
            ' since we have to print now in a new page,
            ' we need to print the header.
            ' So, set columns widths for header.
            
            .newPage
           ' MsgBox "plz insert new page"
            curPrintRow = 1
        End If
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    
    TransID = FormatField(rst("TransID"))
        
        If rst("TableType") = "Interest" Then
        
            TransDate = FormatField(rst("TransDate"))
                          
            'MicAmount = FormatField(Rst("Amount"))
            IntAmount = FormatField(rst("IntAmount"))
            
            'Total = IntAmount + PenalInt + Receipt
            
            
            
        ElseIf rst("TableType") = "Principal" Then
            TransDate = FormatField(rst("TransDate"))
            
            If rst("TransType") = wDeposit Or rst("TransType") = wContraDeposit Then
                   payment = FormatField(rst("Amount"))
            Else
                   Receipt = FormatField(rst("Amount"))
            End If
             
            PenalInt = FormatField(rst("PenalIntAmount"))
            
            
               Balance = FormatField(rst("Balance"))
               Total = IntAmount + PenalInt + payment
        End If
           
       ' .PrintRows
        
        
        rst.MoveNext
       
    Wend
    If TransID > 0 Then
         .ColText(0) = TransDate
         .ColText(1) = Receipt
         .ColText(2) = payment
         .ColText(3) = IntAmount
         .ColText(4) = PenalInt
         .ColText(5) = Balance
         .ColText(6) = Total
         
         Debug.Print Receipt & " | " & payment & " | " & IntAmount
         
           If Receipt <> 0 Then
         .ColText(1) = Receipt
        Else
         .ColText(1) = " "
        End If
           
        If payment <> 0 Then
         .ColText(2) = payment
        Else
         .ColText(2) = " "
        End If
           
        If IntAmount <> 0 Then
         .ColText(3) = IntAmount
        Else
         .ColText(3) = " "
        End If
           
        If PenalInt <> 0 Then
         .ColText(4) = PenalInt
        Else
         .ColText(4) = " "
        End If
      
        If Total <> 0 Then
         .ColText(6) = Total
        Else
         .ColText(6) = " "
        End If
         
         
         .PrintRows
         
        Receipt = 0: payment = 0:
        IntAmount = 0: PenalInt = 0: Balance = 0: Total = 0
    End If
    .newPage
End With
Printer.EndDoc


Set rst = Nothing
Set metaRst = Nothing
Set clsPrint = Nothing
            
'Now Update the Last Print Id to the master
SqlStr = "UPDATE LoanMaster set LastPrintId = " & TransID & _
        ", LastPrintRow = " & curPrintRow - 1 & _
        " Where LoanId = " & m_LoanID
gDbTrans.BeginTrans
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
Else
    gDbTrans.CommitTrans
End If
End Sub



Private Sub tabLoans_Click()
On Error Resume Next
    If tabLoans.SelectedItem.Key = "KEY_Note" Then
        fraLoanGrid.Visible = False
        fraInstructions.Visible = True
        fraInstructions.ZOrder 0
        Exit Sub
    Else
        fraInstructions.Visible = False
        fraLoanGrid.Visible = True
        fraLoanGrid.ZOrder 0
    End If
    Dim LoanID As Long
    LoanID = CLng(Mid(tabLoans.SelectedItem.Key, 4))
    
    If m_LoanID = LoanID And Me.ActiveControl.name = tabLoans.name Then
        Err.Clear
        Exit Sub
    End If
    Err.Clear
    If LoanID Then Call LoadLoanDetail(LoanID)
    
End Sub


Private Sub txtCustID_Change()
If m_CustomerID = Val(txtCustID) Then Exit Sub
Call ClearControls
txtCustNAme = ""
fraLoanGrid.Visible = False
tabLoans.Tabs.Clear
m_CustomerID = 0

End Sub


Private Sub txtMisc_Change()
Call txtRegInt_Change
End Sub

Private Sub txtMisc_GotFocus()
With txtMisc
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtPenalBalance_Change()
txtTotalIntBalance = Val(txtRegIntBalance) = Val(txtPenalBalance)
End Sub

Private Sub txtPenalInt_Change()
Call txtRegInt_Change
End Sub

Private Sub txtPenalInt_GotFocus()
With txtPenalInt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtPrincAmount_Change()
'Call txtRegInt_Change
End Sub

Private Sub txtPrincAmount_GotFocus()
With txtPrincAmount
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub txtRegInt_Change()
Dim Amount As Currency
txtRegIntBalance.Tag = Val(txtRegInt)
Amount = txtTotAmt
If Val(txtTotAmt.Tag) Then
    txtPrincAmount = txtTotAmt - (txtTotalIntBalance + txtRegInt + txtPenalInt + txtMisc)
Else
    Amount = txtTotalIntBalance + txtRegInt + txtPenalInt + txtMisc + txtPrincAmount
    If Amount < 0 Then Amount = 0
    txtTotAmt = Amount
End If
End Sub

Private Sub txtRegInt_GotFocus()
With txtRegInt
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


Private Sub txtRegIntBalance_Change()
txtTotalIntBalance = Val(txtRegIntBalance) + Val(txtPenalBalance)
End Sub

Private Sub txtTotalIntBalance_Change()
Call txtRegInt_Change
End Sub

Private Sub txtTotalIntBalance_GotFocus()
With txtTotalIntBalance
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTotAmt_Change()
If Me.ActiveControl.name <> txtTotAmt.name Then Exit Sub
txtPrincAmount = txtTotAmt - txtRegInt - txtPenalInt - txtMisc
txtTotAmt.Tag = 1
End Sub

Private Sub txtTotAmt_GotFocus()
With txtTotAmt
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub


Private Sub txtTransDate_GotFocus()
    txtTransDate.Tag = txtTransDate.Text
End Sub

'
Private Sub txtTransDate_LostFocus()
'cmdRepay.Default = True
Dim TmpStr  As String
'Load the InterestDetails
    If txtTransDate.Tag = txtTransDate.Text Then Exit Sub
    txtTransDate.Tag = txtTransDate.Text
    If Not DateValidate(txtTransDate, "/", True) Then Exit Sub

Call tabLoans_Click
'txtTransDate.Text = TmpStr
txtTransDate.Tag = txtTransDate.Text
End Sub


'This Functionm Returns the Last Transaction Date of the
'Memeber Transaction of the particular account
Private Sub GetLastTransDate(ByVal AccountId As Long, _
                Optional TransID As Long, Optional TransDate As Date)

Dim rst As Recordset
TransID = 0
TransDate = vbNull
'
On Error GoTo ErrLine

'NOw get the Transcation Id from The table
Dim tmpTransID As Integer
'Now Assume deposit date as the last int paid amount
gDbTrans.SqlStmt = "Select Top 1 TransID,TransDate FROM MemTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        TransID = FormatField(rst("TransID")): TransDate = rst("TransDate")

'Get Max Trans From Interest table
gDbTrans.SqlStmt = "Select TransID,TransDate FROM MemIntTrans " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = rst("TransDate")
End If

'Get Max TransID From Payabale Trans
gDbTrans.SqlStmt = "Select TransID,TransDate FROM MemIntPayable " & _
                    " where AccID = " & AccountId & _
                    " ORder By TransId Desc"

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    tmpTransID = FormatField(rst("TransID"))
    If tmpTransID > TransID Then _
        TransID = tmpTransID: TransDate = FormatField(rst("TransDate"))
End If

ErrLine:

End Sub

'This Function Returns the Last Transction Date of The Fd
' of the given account Id
' In case there is no transaction it reurns "1/1/100"
Public Function GetMemberLastTransDate(ByVal AccountId As Long) As Date
Dim TransDate As Date
Call GetLastTransDate(AccountId, , TransDate)
GetMemberLastTransDate = TransDate

End Function


'This Function Returns the Max Transction ID of
'the given Member share account Id
'In case there is no transaction it reurns 0
Public Function GetMemberMaxTransID(ByVal AccountId As Long) As Long
Dim TransID As Long
Call GetLastTransDate(AccountId, TransID)
GetMemberMaxTransID = TransID

End Function


Public Function ComputeTotalMMLiability(AsOnDate As Date) As Currency

Dim ret As Long
Dim rst As Recordset

ComputeTotalMMLiability = 0

Dim SqlStr As String
SqlStr = "SELECT Max(TransID) As MaxTransID,AccID " & _
    " FROM MemTrans WHERE TransDate <= #" & AsOnDate & "# GROUP BY AccID"
gDbTrans.SqlStmt = SqlStr
gDbTrans.CreateView ("qryTemp")

gDbTrans.SqlStmt = "SELECT SUM(Balance) FROM MemTrans A, qryTemp B " & _
    " WHERE A.AccID = B.AccID And A.TransID = B.MaxTransID"

'Dim Rst As Recordset

If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then ComputeTotalMMLiability = FormatField(rst(0))

Exit Function

End Function


Private Function GetCustomerID(SearchString As String) As Integer

Dim SqlStr As String
Dim rst As Recordset

If Trim(SearchString) = "" Then _
    SearchString = InputBox("Eneter Name to search", "SearchString")

SqlStr = "SELECT CustomerID,Title + ' ' + FirstName+' '" & _
        " + MiddleName +' ' + LastName as Name FROM NameTab "

      
If Trim$(SearchString) <> "" Then
    SqlStr = SqlStr & IIf(InStr(1, SqlStr, "where", vbTextCompare), " AND ", " WHERE ")
    SqlStr = SqlStr & " (firstName like '" & Trim$(SearchString) & "%' " & _
        " OR LastName like '" & Trim$(SearchString) & "%' )"
End If

gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(rst, adOpenStatic) <= 0 Then Exit Function

MousePointer = vbHourglass

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Call FillView(m_frmLookUp.lvwReport, rst, True)
m_retVar = ""

MousePointer = vbDefault
m_frmLookUp.Show vbModal

'If Val(m_retVar) <= 0 Then Exit Function
GetCustomerID = Val(m_retVar)
End Function




