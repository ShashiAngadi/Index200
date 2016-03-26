VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmLoanMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create loan account"
   ClientHeight    =   8880
   ClientLeft      =   870
   ClientTop       =   1245
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8460
   Begin VB.Frame fraGuarantor 
      Caption         =   "Guarantor"
      Height          =   1900
      Left            =   90
      TabIndex        =   33
      Top             =   3900
      Width           =   8115
      Begin VB.TextBox txtGName 
         Height          =   375
         Index           =   3
         Left            =   2730
         TabIndex        =   72
         Top             =   1900
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox txtGName 
         Height          =   375
         Index           =   2
         Left            =   2730
         TabIndex        =   71
         Top             =   2400
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.CommandButton cmdGuar 
         Caption         =   "..."
         Height          =   315
         Index           =   3
         Left            =   7620
         TabIndex        =   70
         Top             =   1900
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton cmdGuar 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   7620
         TabIndex        =   69
         Top             =   2400
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtGMem 
         Height          =   375
         Index           =   3
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   68
         Top             =   1900
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtGMem 
         Height          =   375
         Index           =   2
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   67
         Top             =   2400
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtGMem 
         Height          =   375
         Index           =   1
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   41
         Top             =   1440
         Width           =   840
      End
      Begin VB.TextBox txtGMem 
         Height          =   375
         Index           =   0
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   37
         Top             =   930
         Width           =   840
      End
      Begin VB.ComboBox cmbAccGroup 
         Height          =   315
         Left            =   4710
         TabIndex        =   49
         Top             =   400
         Width           =   2505
      End
      Begin VB.CommandButton cmdGuar 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   7620
         TabIndex        =   43
         Top             =   1410
         Width           =   345
      End
      Begin VB.CommandButton cmdGuar 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   7620
         TabIndex        =   39
         Top             =   960
         Width           =   345
      End
      Begin VB.TextBox txtGName 
         Height          =   375
         Index           =   1
         Left            =   2730
         TabIndex        =   42
         Top             =   1410
         Width           =   4845
      End
      Begin VB.TextBox txtGName 
         Height          =   375
         Index           =   0
         Left            =   2730
         TabIndex        =   38
         Top             =   930
         Width           =   4845
      End
      Begin VB.ComboBox cmbPurpose 
         Height          =   315
         Left            =   1770
         TabIndex        =   35
         Top             =   400
         Width           =   2655
      End
      Begin VB.Label lblGuaranteer 
         Caption         =   "I &Guarantor Name :"
         Height          =   300
         Index           =   3
         Left            =   150
         TabIndex        =   74
         Top             =   2460
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblGuaranteer 
         Caption         =   "I &Guarantor Name :"
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   73
         Top             =   1900
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblGuaranteer 
         Caption         =   "I &Guarantor Name :"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lblPurpose 
         Caption         =   "Loan &Purpose"
         Height          =   300
         Left            =   210
         TabIndex        =   34
         Top             =   400
         Width           =   1425
      End
      Begin VB.Label lblGuaranteer 
         Caption         =   "I &Guarantor Name :"
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   40
         Top             =   1440
         Width           =   1485
      End
   End
   Begin VB.CheckBox chkShowLoan 
      Caption         =   "Show loan transaction"
      Height          =   315
      Left            =   180
      TabIndex        =   15
      Top             =   8220
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   90
      TabIndex        =   44
      Top             =   5670
      Width           =   8115
      Begin VB.TextBox txtPledgItem 
         Height          =   315
         Left            =   1770
         TabIndex        =   46
         Top             =   240
         Width           =   5835
      End
      Begin WIS_Currency_Text_Box.CurrText txtPledgevalue 
         Height          =   345
         Left            =   1770
         TabIndex        =   48
         Top             =   690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblPledgevalue 
         Caption         =   "Pledge &value :"
         Height          =   300
         Left            =   90
         TabIndex        =   47
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label lblPledgeItem 
         Caption         =   "&Pledge  Item :"
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   270
         Width           =   1395
      End
   End
   Begin VB.Frame fra2 
      Height          =   2115
      Left            =   90
      TabIndex        =   12
      Top             =   1860
      Width           =   8115
      Begin VB.CommandButton cmdDueDate 
         Caption         =   "..."
         Height          =   315
         Left            =   7620
         TabIndex        =   22
         Top             =   730
         Width           =   315
      End
      Begin VB.TextBox txtDueDate 
         Height          =   315
         Left            =   5850
         TabIndex        =   23
         Top             =   730
         Width           =   1605
      End
      Begin VB.TextBox txtIssueDate 
         Height          =   315
         Left            =   1890
         TabIndex        =   20
         Top             =   730
         Width           =   1215
      End
      Begin VB.TextBox txtIntrate 
         Height          =   315
         Left            =   1890
         TabIndex        =   25
         Top             =   1190
         Width           =   1635
      End
      Begin VB.TextBox txtPenalInt 
         Height          =   315
         Left            =   5850
         TabIndex        =   27
         Top             =   1190
         Width           =   1605
      End
      Begin VB.ComboBox cmbInstType 
         Height          =   315
         Left            =   1890
         TabIndex        =   29
         Top             =   1650
         Width           =   1635
      End
      Begin VB.TextBox txtNoOfINst 
         Height          =   315
         Left            =   5850
         TabIndex        =   31
         Top             =   1650
         Width           =   1605
      End
      Begin VB.CommandButton cmdIssueDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3210
         TabIndex        =   19
         Top             =   730
         Width           =   315
      End
      Begin VB.CheckBox chkEMI 
         Caption         =   "&EMI"
         Height          =   300
         Left            =   7530
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdInst 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   7620
         TabIndex        =   32
         Top             =   1650
         Width           =   315
      End
      Begin VB.TextBox txtLoanAccNo 
         Height          =   315
         Left            =   5850
         TabIndex        =   17
         Top             =   270
         Width           =   1605
      End
      Begin WIS_Currency_Text_Box.CurrText txtLoanAmount 
         Height          =   345
         Left            =   1890
         TabIndex        =   14
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblLoanAmount 
         Caption         =   "&Sanction Amount"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label lblDueDate 
         Caption         =   "&Due date :"
         Height          =   300
         Left            =   3810
         TabIndex        =   21
         Top             =   735
         Width           =   1995
      End
      Begin VB.Label lblIssueDate 
         Caption         =   "&Issue date :"
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   735
         Width           =   1725
      End
      Begin VB.Label lblIntrate 
         Caption         =   "Rate of &interest :"
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   1185
         Width           =   1695
      End
      Begin VB.Label lblPenalInt 
         Caption         =   "&Penal interest :"
         Height          =   300
         Left            =   3810
         TabIndex        =   26
         Top             =   1185
         Width           =   1905
      End
      Begin VB.Label lblNoOfInst 
         Caption         =   "&No of installments"
         Height          =   300
         Left            =   3810
         TabIndex        =   30
         Top             =   1650
         Width           =   1785
      End
      Begin VB.Label lblInstType 
         Caption         =   "Installment &Mode"
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Top             =   1650
         Width           =   1665
      End
      Begin VB.Label lblLoanAccNo 
         Caption         =   "Loan Account &No :"
         Height          =   300
         Left            =   3810
         TabIndex        =   16
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.Frame fraCust 
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8085
      Begin VB.ComboBox cmbMemberType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   2805
      End
      Begin VB.ListBox lstHidden 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   5340
         TabIndex        =   55
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbLoanScheme 
         Height          =   315
         Left            =   1890
         TabIndex        =   6
         Top             =   750
         Width           =   5325
      End
      Begin VB.CommandButton cmdCustID 
         Caption         =   "..."
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   6090
         MaxLength       =   9
         TabIndex        =   3
         Top             =   270
         Width           =   960
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7650
         TabIndex        =   7
         Top             =   720
         Width           =   315
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   8
         Top             =   720
         Width           =   315
      End
      Begin VB.CommandButton cmdCustName 
         Caption         =   "..."
         Height          =   315
         Left            =   7620
         TabIndex        =   11
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1890
         TabIndex        =   10
         Top             =   1200
         Width           =   5445
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLoanScheme 
         Caption         =   "&Loan Scheme :"
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label lblCustID 
         AutoSize        =   -1  'True
         Caption         =   "&Member No :"
         Height          =   300
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label lblCustName 
         Caption         =   "&Name :"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   1260
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   400
      Left            =   4290
      TabIndex        =   65
      Top             =   8190
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6900
      TabIndex        =   66
      Top             =   8190
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Default         =   -1  'True
      Height          =   400
      Left            =   5565
      TabIndex        =   64
      Top             =   8190
      Width           =   1215
   End
   Begin VB.Frame fraAgriDet 
      Height          =   1095
      Left            =   90
      TabIndex        =   50
      Top             =   6750
      Width           =   8115
      Begin VB.ComboBox cmbSeason 
         Height          =   315
         Left            =   1770
         TabIndex        =   52
         Top             =   390
         Width           =   2235
      End
      Begin VB.ComboBox cmbCrop 
         Height          =   315
         Left            =   5940
         TabIndex        =   54
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label lblSeason 
         Caption         =   "Season"
         Height          =   300
         Left            =   240
         TabIndex        =   51
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label LblCrop 
         Caption         =   "Crop :"
         Height          =   300
         Left            =   4560
         TabIndex        =   53
         Top             =   420
         Width           =   1125
      End
   End
   Begin VB.Frame fraVehicle 
      Caption         =   "Vehicle Details"
      Height          =   1125
      Left            =   90
      TabIndex        =   57
      Top             =   6750
      Width           =   8115
      Begin VB.TextBox txtRegNo 
         Height          =   315
         Left            =   1770
         TabIndex        =   59
         Top             =   350
         Width           =   2115
      End
      Begin VB.TextBox txtInsurFee 
         Height          =   315
         Left            =   6060
         TabIndex        =   61
         Top             =   350
         Width           =   1635
      End
      Begin VB.TextBox txtInsureDate 
         Height          =   315
         Left            =   1770
         TabIndex        =   63
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label lblRegNo 
         Caption         =   "Registration No"
         Height          =   300
         Left            =   150
         TabIndex        =   58
         Top             =   350
         Width           =   1395
      End
      Begin VB.Label lblInsurFee 
         Caption         =   "Insurence Fee :"
         Height          =   300
         Left            =   4350
         TabIndex        =   60
         Top             =   345
         Width           =   1155
      End
      Begin VB.Label lblInsureDate 
         Caption         =   "Insurence Up to :"
         Height          =   300
         Left            =   150
         TabIndex        =   62
         Top             =   750
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmLoanMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_FrmInst As frmLoanInst
Attribute m_FrmInst.VB_VarHelpID = -1
Private m_dbOperation As wis_DBOperation
Private m_CustDet As clsCustReg
Attribute m_CustDet.VB_VarHelpID = -1
Private m_NoOfGuaranteers  As Integer

Private m_CustomerID As Long
Private m_LoanID As Long

Private m_retVar As Variant
Private m_SchemeId As Integer
Private m_InstDet As Boolean
Private m_InstIndianDates() As String
Private m_InstAmounts() As Currency
Private m_InstBalance() As Currency

Private m_rstLoanScheme As Recordset
Private m_rstCustLoans As Recordset

Public Event CustomerSelected(custId As Long)
Public Event LoanCreated(ByVal LoanID As Long)
Public Event LoanModified(ByVal LoanID As Long)
Public Event AccountChanged(ByVal LoanID As Long)
Public Event WindowClosed()

Private Function CheckValidation() As Boolean
   Dim AccNum As String
   Dim SchemeID As Integer
   Dim SqlStr As String
   Dim rst As Recordset
   Dim loopCount As Integer
    
   'Check all the validations
   CheckValidation = False
   If txtCustName = "" Then
      MsgBox "Please Select the Customer Name", vbInformation, wis_MESSAGE_TITLE
      Exit Function
   End If
   
   If cmbLoanScheme.ListIndex = -1 Then
      MsgBox "Please Select the Loan scheme", vbInformation, wis_MESSAGE_TITLE
      Exit Function
   End If
   
   If m_rstLoanScheme Is Nothing Then GoTo Exit_Line
If HasOverdueLoans(m_CustomerID) Then
     'nRet = MsgBox("The member has overdue loans.Loan cannot be issued." _
             & vbCrLf & "Do you want to continue anyway?", vbQuestion + _
             vbYesNo, wis_MESSAGE_TITLE)
     If MsgBox(GetResourceString(717) _
             & vbCrLf & GetResourceString(541), vbQuestion + _
             vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
         ActivateTextBox txtCustID
         GoTo Exit_Line
     End If
 End If
    
' Guarantors...
 
 For loopCount = 0 To 3
    
    If Val(txtGName(loopCount).Tag) > 0 Then
        ' Check if the guarantor is the same as the loan claimer.
        If Val(txtGName(loopCount).Tag) = m_CustomerID Then
            'MsgBox "A person cannot stand guarantee for his own loan !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(724), vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtGName(loopCount)
            GoTo Exit_Line
        End If
        
 
        ' Check if the guarantor is the same as the loan claimer.
        If Val(txtGName(loopCount).Tag) = m_CustomerID Then
            'MsgBox "A person cannot stand guarantee for his own loan !", _
                    vbExclamation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(724), _
                    vbExclamation, wis_MESSAGE_TITLE
            ActivateTextBox txtGName(loopCount)
            GoTo Exit_Line
        End If
        
        ' Check if the guarantor is eligible for standing guarantee.
        If HasOverdueLoans(Val(txtGName(loopCount).Tag)) Then  '//TODO//
            'MsgBox "Guarantor1 " & PropIssueGetVal("Guarantor1") _
                    & " has loan overdues.  Please select another " _
                    & "guarantor.", vbExclamation, wis_MESSAGE_TITLE
            If MsgBox(GetResourceString(725), vbYesNo + _
                vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
                ActivateTextBox txtGName(loopCount)
                GoTo Exit_Line
            End If
        End If
        
        If Not IsMemberExistsForCustomerID(Val(txtGName(loopCount).Tag)) Then  '//TODO//
            MsgBox "Guarantor " & txtGName(loopCount).Text _
                & " is not a active member. Please select another guarantor." _
                & "", vbExclamation, wis_MESSAGE_TITLE
        
            ActivateTextBox txtGName(loopCount)
            GoTo Exit_Line
        End If
        
        
    Else
        'Changed On 18/09/2003
        'Need not to confirm whether user want to enter the
        'Guranteer
        'If MsgBox(GetResourceString(723) & GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, _
            wis_MESSAGE_TITLE) = vbNo Then GoTo Exit_line
    End If
 Next
     
 
'Check for the Account Group
If cmbAccGroup.ListIndex = -1 Then cmbAccGroup = 0
'    MsgBox GetResourceString(749), vbInformation, wis_MESSAGE_TITLE
'    cmbAccGroup.SetFocus
'    GoTo Exit_line
'End If
   
If txtPledgItem.Text <> "" Then
    With txtPledgevalue
        ' Make sure that pledge value is mentioned,
        ' if pledge item is mentioned.
        If Trim(.Text) = "" Then
            'MsgBox "Specify the value of pledged items.", _
                        vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(719), _
                        vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtPledgevalue
            GoTo Exit_Line
        End If
        
        ' Ensure the value of pledge items is valid.
        If Not IsNumeric(.Text) Or Val(.Text) <= 0 Then
            'MsgBox "Invalid value for pledge item.", _
                    vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(720), _
                    vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox txtPledgevalue
            GoTo Exit_Line
        End If
    End With
End If
   
   AccNum = Trim(txtLoanAccNo)
   If AccNum = "" Then
      MsgBox "Please Specify the Loan AccountNo", vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtLoanAccNo
      Exit Function
   End If
   'Check For the Same Account No
   SqlStr = "SELECT * FROM LoanMaster WHERE SchemeId = " & SchemeID & _
                " AND AccNum = " & AddQuotes(AccNum, True)
   
   If m_dbOperation = Update Then SqlStr = SqlStr & " AND LoanId <> " & m_LoanID
    
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
        'THis account no has given to othere account
        'So warn abount this to the user
        'MsgBox "This loan account number alread exists ", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(545), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtLoanAccNo
        Exit Function
    End If
    
   If Not CurrencyValidate(txtIntRate, False) Then
      'MsgBox "Please enter the rate of interest", vbInformation, wis_MESSAGE_TITLE
      MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtIntRate
      Exit Function
   End If
   If Not CurrencyValidate(txtPenalINT, True) Then
      'MsgBox "Please enter the Penal interest rate", vbInformation, wis_MESSAGE_TITLE
      MsgBox GetResourceString(505), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtPenalINT
      Exit Function
   End If
   If Not DateValidate(txtIssueDate, "/", True) Then
      'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
      MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtIssueDate
      Exit Function
   End If
   If Not DateValidate(txtDueDate, "/", True) Then
      'MsgBox "Invalid date specified", vbInformation, wis_MESSAGE_TITLE
      MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
      ActivateTextBox txtDueDate
      Exit Function
   End If
'   If Not (optSf.value Or optBF Or optMf Or optOth) Then
'      MsgBox "Please Select type of farmer ", vbInformation, wis_MESSAGE_TITLE
'      Exit Function
'   End If
'''validation of the Agricultural loans
   Dim LoanType  As wis_LoanType
   LoanType = FormatField(m_rstLoanScheme("LoanType"))
   If LoanType = wisCropLoan Then
        If cmbSeason.ListIndex = -1 Then
           MsgBox "Please Select type of Season", vbInformation, wis_MESSAGE_TITLE
           Exit Function
        End If
        If cmbCrop.ListIndex = -1 Then
           MsgBox "Please Select type of Crop", vbInformation, wis_MESSAGE_TITLE
           Exit Function
        End If
    End If
   
'Now Check The Installment Details
'If Installments are allowed then
'Check For The installment Deatils
If cmbInstType.ListIndex > 0 Then
   'Check For The No Of Installment
      If Not IsNumeric(Val(txtNoOfINst.Text)) Then
         MsgBox "Invalid No of Installments specified ", vbInformation, wis_MESSAGE_TITLE
         ActivateTextBox txtNoOfINst
         Exit Function
      End If
      If Not m_InstDet Then 'If installment deatils
         'Not Provided then ask for them
         Call cmdInst_Click
      End If
      If Not m_InstDet Then 'If installet deatils
        'Not do not save it
         Exit Function
      End If
End If

CheckValidation = True
Exit_Line:

End Function
'
Private Sub ClearControls()
Dim count As Integer
    'cmbLoanScheme.ListIndex = -1
    m_SchemeId = 0
    m_LoanID = 0
'    txtCustName = ""
'    If Not m_CustDet Is Nothing And m_CustomerID = 0 Then m_CustDet.NewCustomer
    txtLoanAccNo = ""
    txtLoanAmount.Text = ""
    txtIntRate = ""
    txtPenalINT = ""
    txtIssueDate = ""
    txtDueDate = ""
    
    For count = 0 To 3
        txtGMem(count).Text = ""
        txtGName(count).Text = ""
        txtGName(count).Tag = ""
    Next
    txtPledgItem = ""
    txtPledgevalue.Text = ""
    txtNoOfINst = ""
    txtRegNo = ""
    txtInsurFee = ""
    
    cmbPurpose.ListIndex = -1
    cmbInstType.ListIndex = -1
    cmbSeason.ListIndex = -1
    cmbCrop.ListIndex = -1
    
    cmdDelete.Enabled = False
    'cmdNext.Enabled = False
    cmdPrev.Enabled = False
    
    m_dbOperation = Insert
    'cmdCreate.Caption = "Create"
    cmdCreate.Caption = GetResourceString(15)

    chkShowLoan.Visible = False
    chkShowLoan.Value = vbChecked
    ReDim m_InstIndianDates(0)
    ReDim m_InstAmounts(0)
    ReDim m_InstBalance(0)
    m_InstDet = False
End Sub


Public Property Let CustomerID(NewValue As Long)

Dim rst As Recordset
gDbTrans.SqlStmt = "select * From LoanMaster " & _
    "Where CustomerId = " & NewValue
If m_SchemeId Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " ANd SchemeID = " & m_SchemeId

gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Order by LoanId desc"

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    If m_LoanID Then
        rst.Find "LoanID = " & m_LoanID
        If rst.EOF Then  'no match found then
            rst.MoveFirst
            m_LoanID = rst("LoanId")
        End If
    Else
        m_LoanID = FormatField(rst("LoanID"))
    End If
End If

Set rst = Nothing
    
    m_CustomerID = NewValue
End Property

'If Selected Customer has already have this loan
'Then Load The Interest,Issue date, Due Date and
'other details of this loan
Private Sub LoadCustomerLoan()

m_dbOperation = Insert
m_InstDet = False
If m_CustomerID = 0 Then Exit Sub

Dim SchemeID As Long
Dim SqlStr As String
Dim LoanID As Long

Dim RstInst As Recordset
Dim Season As wisSeason
Dim Crop As Byte
Dim iCount As Integer
Dim InstMode As Byte
Dim rst As Recordset
'Now Get Is any loan related to this customer
SchemeID = cmbLoanScheme.ItemData(cmbLoanScheme.ListIndex)

If m_rstCustLoans Is Nothing Then
    gDbTrans.SqlStmt = "SELECT * from LoanMaster " & _
            " WHERE SchemeID = " & SchemeID & _
            " And CustomerID = " & m_CustomerID
    iCount = gDbTrans.Fetch(m_rstCustLoans, adOpenStatic)
    If iCount < 1 Then Exit Sub
    
    If iCount > 1 Then cmdPrev.Enabled = True
    cmdNext.Enabled = iCount
    'Set m_rstCustLoans = gDbTrans.Rst.Clone
End If

If m_rstCustLoans.EOF Then cmdPrev.Enabled = False
'If m_rstCustLoans.BOF Then cmdNext.Enabled = False

If m_rstCustLoans.AbsolutePosition > 1 And _
            m_rstCustLoans.recordCount > 1 Then cmdPrev.Enabled = True
'cmdNext.Enabled = False
'If m_rstCustLoans.AbsolutePosition = m_rstCustLoans.RecordCount Then cmdNext.Enabled = False

cmdDelete.Enabled = gCurrUser.IsAdmin
m_dbOperation = Update
cmdCreate.Caption = GetResourceString(171) 'Update

'Now Load The Details Of The installment
Dim InstNo As Integer
m_LoanID = FormatField(m_rstCustLoans("LoanID"))
m_SchemeId = FormatField(m_rstCustLoans("SchemeID"))

txtNoOfINst = FormatField(m_rstCustLoans("NoOfInstall"))
If Val(FormatField(m_rstCustLoans("EMI"))) Then chkEMI.Value = vbChecked
txtLoanAccNo = FormatField(m_rstCustLoans("AccNum"))
txtLoanAmount = FormatField(m_rstCustLoans("LoanAmount"))
'Get The Ist TransaCtion Amount

txtIntRate = FormatField(m_rstCustLoans("IntRate"))
txtPenalINT = FormatField(m_rstCustLoans("PenalIntRate"))
txtIssueDate = FormatField(m_rstCustLoans("IssueDate"))
txtDueDate = FormatField(m_rstCustLoans("LoanDueDate"))

InstMode = FormatField(m_rstCustLoans("InstMode"))
For iCount = 0 To cmbInstType.ListCount - 1
    If cmbInstType.ItemData(iCount) = InstMode Then
        cmbInstType.ListIndex = iCount
        Exit For
    End If
Next iCount

'Now set the Account group
InstMode = FormatField(m_rstCustLoans("AccGroupID"))
For iCount = 0 To cmbAccGroup.ListCount - 1
    If cmbAccGroup.ItemData(iCount) = InstMode Then
        cmbAccGroup.ListIndex = iCount
        Exit For
    End If
Next iCount

'Get The Guarantor Details
cmbPurpose.Text = FormatField(m_rstCustLoans("LoanPurpose"))

Dim Guaranteer As Long
Dim GMemNum As String
Dim memberType As Integer
For iCount = 0 To 3
    txtGName(iCount).Tag = 0
    Guaranteer = FormatField(m_rstCustLoans("Guarantor" & CStr(iCount + 1)))
    txtGName(iCount).Text = GetMemberNameNumberByCustID(Guaranteer, GMemNum, memberType)
    If Len(txtGName(iCount).Text) > 0 Then txtGName(iCount).Tag = Guaranteer: txtGMem(iCount).Text = GMemNum
Next

txtPledgItem = FormatField(m_rstCustLoans("PledgeItem"))
txtPledgevalue = FormatField(m_rstCustLoans("PledgeValue"))

'Set the Season combo box
Season = FormatField(m_rstCustLoans("SeasonType"))
For iCount = 0 To cmbSeason.ListCount - 1
    If cmbSeason.ItemData(iCount) = Season Then
        cmbSeason.ListIndex = iCount
        Exit For
    End If
Next iCount

'Set the Crop combo box
Crop = FormatField(m_rstCustLoans("CropType"))
For iCount = 0 To cmbCrop.ListCount - 1
    If cmbCrop.ItemData(iCount) = Crop Then
        cmbCrop.ListIndex = iCount
        Exit For
    End If
Next iCount

Dim retstr As String
Dim strArr() As String

retstr = FormatField(m_rstCustLoans("OtherDets"))
On Error Resume Next
Call GetStringArray(retstr, strArr, gDelim)
ReDim Preserve strArr(2)
txtRegNo = strArr(0)
txtInsurFee = strArr(1)
txtInsureDate = strArr(2)
If Not m_FrmInst Is Nothing Then Unload m_FrmInst
Set m_FrmInst = Nothing
On Error GoTo 0

'Now Check Whether to Show the transaction or not
chkShowLoan.Visible = False
If FormatField(m_rstCustLoans("LoanClosed")) <> 0 Then
    chkShowLoan.Visible = True
    chkShowLoan.Value = IIf(m_rstCustLoans("LoanClosed") = 2, 0, 1)
End If
'Now set the grid with the no fo installments
SqlStr = ""

'Now load The Details Installments
SqlStr = "SELECT * from LoanInst WHERE LoanID = " & m_LoanID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(RstInst, adOpenDynamic) < 1 Then Exit Sub
   'Set RstInst = gDbTrans.Rst.Clone

InstNo = RstInst.recordCount
txtNoOfINst = InstNo

Set m_FrmInst = New frmLoanInst
m_FrmInst.Operation = InstInsert
m_InstDet = True
m_FrmInst.LoanID = m_LoanID
Load m_FrmInst
m_FrmInst.Operation = InstInsert
'Now Load The Details of installment

iCount = 1
While Not RstInst.EOF
   ReDim Preserve m_InstAmounts(iCount - 1)
   ReDim Preserve m_InstIndianDates(iCount - 1)
   ReDim Preserve m_InstBalance(iCount - 1)
   With m_FrmInst.grdInst
      .Rows = iCount + 1
      .Row = iCount
      .Col = 0: .Text = iCount
      .Col = 1: .Text = FormatField(RstInst("InstDate"))
      m_InstIndianDates(iCount - 1) = .Text
      m_InstBalance(iCount - 1) = FormatField(RstInst("InstBalance"))
      .Col = 2: .Text = FormatField(RstInst("InstAmount"))
      m_InstAmounts(iCount - 1) = .Text
   End With
   RstInst.MoveNext
   iCount = iCount + 1
Wend

End Sub
Public Sub LoadLoan(ByVal LoanID As Long)
Dim rst As Recordset

gDbTrans.SqlStmt = "Select LoanId,A.CustomerID,SchemeID, " & _
    " Title +' '+FirstName +' '+MiddleName + " & _
    " ' '+LastName As name From LoanMaster A,NameTab B " & _
    " Where LoanID = " & LoanID & " ANd A.CustomerID = B.CustomerID"
    
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub
Dim SchemeID As Integer
SchemeID = FormatField(rst("SchemeId"))
Call SetComboIndex(cmbLoanScheme, , SchemeID)

'txtCustName = FormatField(rst("Name"))
'txtCustId = GetMemberNumber(rst("CustomerID"))
Dim memNum As String

Dim memberType As Integer
txtCustName.Caption = GetMemberNameNumberByCustID(rst("CustomerID"), memNum, memberType)
txtCustID.Text = memNum
If cmbMemberType.ListCount > 1 Then Call SetComboIndex(cmbMemberType, , memberType)

m_CustomerID = rst("CustomerId")
m_LoanID = rst("LoanID")
gDbTrans.SqlStmt = "SELECT * from LoanMaster WHERE SchemeID = " & SchemeID & _
          " And CustomerID = " & m_CustomerID
Call gDbTrans.Fetch(m_rstCustLoans, adOpenDynamic)

m_rstCustLoans.MoveFirst
m_rstCustLoans.Find "LoanID = " & LoanID

Call LoadCustomerLoan

End Sub

Private Sub LoadLoanSchemeDetail()
Dim SchemeID As Integer
Dim SqlStr As String
Dim Days As Long

SchemeID = cmbLoanScheme.ItemData(cmbLoanScheme.ListIndex)
'Now load the Details of loanscekme
SqlStr = "SELECT * FROM LoanScheme where SchemeID = " & SchemeID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(m_rstLoanScheme, adOpenStatic) < 1 Then Exit Sub

'If Loanscehme Is Non Agri type then
Dim LoanCategary As wisLoanCategories
LoanCategary = FormatField(m_rstLoanScheme("Category"))
If LoanCategary = wisNonAgriculural Then
    'Disable the FraAgriDet
    fraAgriDet.Visible = False
    fraAgriDet.Enabled = False
Else
    fraAgriDet.Visible = True
    fraAgriDet.Enabled = True
    cmbCrop.Enabled = True
    cmbSeason.Enabled = cmbCrop.Enabled
End If
Dim RetBool As Boolean
Dim LoanType As wis_LoanType
LoanType = FormatField(m_rstLoanScheme("LoanType"))
cmbCrop.Enabled = IIf(LoanType = wisCropLoan, -1, 0)
cmbSeason.Enabled = cmbCrop.Enabled
If LoanType = wisVehicleloan Then
    fraVehicle.Visible = True
    fraAgriDet.Visible = False
    fraVehicle.ZOrder 0
Else
    fraVehicle.Visible = False
End If

'Get The Ist Instllment

txtIntRate = FormatField(m_rstLoanScheme("IntRate"))
txtPenalINT = FormatField(m_rstLoanScheme("PenalIntRate"))
txtIssueDate = Format(Now, "dd/mm/yyyy")
On Error Resume Next
Days = FormatField(m_rstLoanScheme("MonthDuration"))
txtDueDate = GetIndianDate(DateAdd("M", Days, GetSysFormatDate(txtIssueDate)))
Days = FormatField(m_rstLoanScheme("DayDuration"))
txtDueDate = GetIndianDate(DateAdd("d", Days, GetSysFormatDate(txtDueDate)))

On Error GoTo 0
'Set Rst = Nothing

'Load the loan purpose from databse
Call LoadLoanPurposes(cmbPurpose, SchemeID)


End Sub

Public Property Let LoanID(NewValue As Long)
    
Dim rst As Recordset

gDbTrans.SqlStmt = "Select * From LoanMaster " & _
        " Where LoanId = " & NewValue

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
    m_CustomerID = FormatField(rst("CustomerID"))

Set rst = Nothing
m_LoanID = NewValue
    
End Property

'
Private Function SaveInstallmentDetails(LoanID As Long, rst As Recordset) As Boolean

Dim InstNo As Integer
Dim NoOfInst As Integer
Dim lpCount As Integer
Dim SqlStr As String
'Dim Rst As Recordset

NoOfInst = Val(txtNoOfINst)

Dim InstBalance As Currency
''
Dim sqlLoop As Integer
sqlLoop = 0
If m_dbOperation = Insert Then
InsertLIne:
    '*Begin and Commit of the Transaction should in the calling function
    'so below code is commented
   'gDbTrans.BeginTrans
   
    SqlStr = "Delete * FROM LoanInst WHERE LoanID = " & LoanID
   gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
       '*gDbTrans.RollBack
       Exit Function
    End If
   For sqlLoop = LBound(m_InstAmounts) To UBound(m_InstAmounts)
      SqlStr = "INSERT INTO LoanInst (LoanID,InstNo," & _
               " InstDate,InstAmount,InstBalance )" & _
            " Values ( " & _
            LoanID & "," & _
            sqlLoop - LBound(m_InstAmounts) + 1 & "," & _
            "#" & GetSysFormatDate(m_InstIndianDates(sqlLoop)) & "#," & _
            m_InstAmounts(sqlLoop) & "," & _
            m_InstAmounts(sqlLoop) & " ) "
      gDbTrans.SqlStmt = SqlStr
      
      If Not gDbTrans.SQLExecute Then
         '*gDbTrans.RollBack
         Exit Function
      End If
   Next sqlLoop
   
   '*gDbTrans.CommitTrans

ElseIf m_dbOperation = Update Then

   SqlStr = ""
   sqlLoop = 0
    '*Begin and Commit of the Transaction should in the calling function
    'so below code is commented
   'gDbTrans.BeginTrans
   If rst Is Nothing Then GoTo InsertLIne
   
    For sqlLoop = LBound(m_InstAmounts) To UBound(m_InstAmounts)
      'If Any Installment paid consider that amount
      InstBalance = m_InstAmounts(sqlLoop) - _
            (FormatField(rst("InstAmount")) - FormatField(rst("InstBalance")))
      InstNo = FormatField(rst("Instno"))
      SqlStr = "UPDATE LoanInst  SET " & _
            " InstDate = #" & GetSysFormatDate(m_InstIndianDates(sqlLoop)) & "#, " & _
            " InstAmount = " & m_InstAmounts(sqlLoop) & ", " & _
            " InstBalance = " & InstBalance & _
            " WHERE LoanID = " & LoanID & _
            " AND InstNo = " & InstNo
      gDbTrans.SqlStmt = SqlStr
      If Not gDbTrans.SQLExecute Then Exit Function
      
      rst.MoveNext
   Next sqlLoop
   '*gDbTrans.CommitTrans
End If

SaveInstallmentDetails = True
End Function

'This function will Save the loan account details
'LoanMaster & LoanInst Tables are used
Private Function SaveLoanAccount() As Boolean

On Error GoTo ErrLine
    Dim AccNum As String
    Dim SchemeID As Long
    Dim IssueDate As String
    Dim DueDate As String
    Dim SqlStr As String
    Dim LoanID As Long
    Dim SeasonType As wisSeason
    Dim CropType As Byte
    Dim InstType As wisInstallmentTypes
    Dim EMI As Boolean
    Dim rstInstDetail As Recordset
    Dim rst As Recordset
    
    
    
SaveLoanAccount = False
AccNum = txtLoanAccNo

'Get the Max LoanID
gDbTrans.SqlStmt = "SELECT Max(LoanId) from LoanMaster "
LoanID = 1
If gDbTrans.Fetch(rst, adOpenStatic) = 1 Then LoanID = FormatField(rst(0)) + 1

'Get SchemeId
SchemeID = cmbLoanScheme.ItemData(cmbLoanScheme.ListIndex)

'CusdtomerID is ModuleLevel m_CustomerID
IssueDate = GetSysFormatDate(Trim$(txtIssueDate))
DueDate = GetSysFormatDate(Trim$(txtDueDate))

If cmbSeason.ListIndex >= 0 Then _
    SeasonType = cmbSeason.ItemData(cmbSeason.ListIndex)
If cmbCrop.ListIndex >= 0 Then _
    CropType = cmbCrop.ItemData(cmbCrop.ListIndex)

'Check for the Guareanteer Details
If m_rstLoanScheme("OnlyMember") Then
    'Check whether Customer Is Member Or not
    
End If

'Check for the existance of accoun number
gDbTrans.SqlStmt = "SELECT LoanID from LoanMaster" & _
    " Where SchemeId = " & SchemeID & " AND AccNum = " & AddQuotes(AccNum)
If m_dbOperation = Update Then
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & " AND LoanID <> " & m_LoanID
End If
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    'MsgBox "Account number " & .Text & "already exists." & vbCrLf & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(545) & vbCrLf & "Please specify another account number !", vbExclamation, gAppName & " - Error"
    Exit Function
End If

Dim InstAmount As Currency
Dim MemID As Long
InstAmount = 0
InstType = Inst_No

If cmbInstType.ListIndex > 0 And Val(txtNoOfINst) > 0 Then
    If Not m_InstDet Then Call cmdInst_Click
 '   If Trim$(m_InstIndianDates(UBound(m_InstIndianDates))) <> "" Then
'        DueDate = GetSysFormatDate(m_InstIndianDates(UBound(m_InstIndianDates)))
  '  End If
    'CusdtomerID is ModuleLevel m_CustomerID
    IssueDate = GetSysFormatDate(Trim$(txtIssueDate))
    DueDate = GetSysFormatDate(Trim$(txtDueDate))
    
    'If he has cancelled to enter the Installment detials
    If Not m_InstDet Then Exit Function
    InstType = cmbInstType.ItemData(cmbInstType.ListIndex)
    InstAmount = m_InstAmounts(0)
    
End If

EMI = False
If chkEMI.Value = vbChecked Then EMI = True

'If Vehicle Loan then Keep these information
Dim OtherDet As String
OtherDet = txtRegNo & gDelim & txtInsurFee & gDelim & Trim(txtInsureDate)


SqlStr = ""
'Get the Memeber ID
   gDbTrans.SqlStmt = "select AccId from MemMaster where customerID= " & m_CustomerID
   If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then MemID = FormatField(rst(0))
   
'm_CustomerID = txtCustID
If m_dbOperation = Insert Then
   '
   SqlStr = "INSERT INTO LoanMaster (LoanID,SchemeID,CustomerID," & _
         " AccNum,MemId,IssueDate,PledgeItem,PledgeValue," & _
         " LoanAmount,InstMode,InstAmount,NoOfInstall,EMI," & _
         " LoanDueDate,IntRate,PenalIntRate," & _
         " Guarantor1, Guarantor2,Guarantor3, Guarantor4,OtherDetS, " & _
         " LoanPurpose,SeasonType,CropType,AccGroupID,UserID) "
   SqlStr = SqlStr & " VALUES (" & _
         LoanID & "," & _
         SchemeID & "," & _
         m_CustomerID & ", " & _
         AddQuotes(Trim(txtLoanAccNo), True) & ", " & _
         MemID & ", " & _
         "#" & IssueDate & "#," & _
         "'" & Trim$(txtPledgItem) & "'," & _
         Val(Trim$(txtPledgevalue)) & "," & _
         Val(Trim$(txtLoanAmount)) & "," & _
         InstType & "," & InstAmount & "," & Val(txtNoOfINst) & "," & EMI & "," & _
         "#" & DueDate & "#," & _
         CSng(Val(Trim$(txtIntRate))) & "," & CSng(Val(Trim$(txtPenalINT))) & "," & _
         Val(txtGName(0).Tag) & ", " & Val(txtGName(1).Tag) & ", " & _
         Val(txtGName(2).Tag) & ", " & Val(txtGName(3).Tag) & ", " & _
         AddQuotes(OtherDet, True) & ", " & _
         AddQuotes(cmbPurpose.Text, True) & ", " & _
         SeasonType & "," & _
         CropType & "," & _
         cmbAccGroup.ItemData(cmbAccGroup.ListIndex) & "," & gUserID & " )"
         
    gDbTrans.BeginTrans
    'iF hE HAS CREATED nEW CUSTOMER
    'FirST save the customer
    m_CustDet.ModuleID = wis_Loans
    If Not m_CustDet.SaveCustomer Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    'Now save the loan detal in LOan Master
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
       Exit Function
    End If
    
    'iF THIS LOAN HAS iNSTALLMENTS
    'then sace the installment details
    Set rstInstDetail = Nothing
    If m_InstDet Then
        If Not SaveInstallmentDetails(LoanID, rstInstDetail) Then
            gDbTrans.RollBack
            Exit Function
        End If
    End If
    
    gDbTrans.CommitTrans
         
ElseIf m_dbOperation = Update Then
    Dim LoanClosed As Byte
    LoanClosed = 0
    With chkShowLoan
        If .Visible Then
            LoanClosed = IIf(.Value = vbChecked, 1, 2)
        End If
    End With
    
   SqlStr = "UPDATE LoanMaster SET  " & _
         " IssueDate = #" & IssueDate & "#," & _
         " PledgeItem = " & AddQuotes(Trim$(txtPledgItem), True) & "," & _
         " PledgeValue = " & Val(Trim$(txtPledgevalue)) & ", " & _
         " LoanAmount = " & Val(Trim$(txtLoanAmount)) & ", " & _
         " MemId = " & MemID & ", " & _
         " AccNum = " & AddQuotes(Trim(txtLoanAccNo), True) & ", " & _
         " InstMode = " & InstType & ", " & _
         " InstAmount = " & InstAmount & ", " & _
         " NoOfInstall = " & Val(txtNoOfINst) & ", " & _
         " EMI = " & EMI & ", " & _
         " LoanDueDate = #" & DueDate & "#," & _
         " IntRate = " & CSng(Val(Trim$(txtIntRate))) & ", " & _
         " PenalIntRate = " & CSng(Val(Trim$(txtPenalINT))) & ", " & _
         " AccGroupID = " & cmbAccGroup.ItemData(cmbAccGroup.ListIndex) & ", " & _
         " Guarantor1 = " & Val(txtGName(0).Tag) & ", " & " Guarantor2 = " & Val(txtGName(1).Tag) & ", " & _
         " Guarantor3 = " & Val(txtGName(2).Tag) & ", " & " Guarantor4 = " & Val(txtGName(3).Tag) & ", " & _
         " LoanClosed  = " & LoanClosed & ", " & _
         " OtherDets = " & AddQuotes(OtherDet, True) & ", " & _
         " LoanPurpose = " & AddQuotes(cmbPurpose.Text, True) & ", " & _
         " SeasonType = " & SeasonType & ", " & _
         " CropType = " & CropType & _
         " WHERE CustomerID = " & m_CustomerID & _
         " AND SchemeID = " & m_SchemeId & _
         " AND LoanID = " & m_LoanID
         
    gDbTrans.SqlStmt = SqlStr
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
       gDbTrans.RollBack
       Exit Function
    End If
         
    'iF THIS LOAN HAS iNSTALLMENTS
    'then sace the installment details
    If m_InstDet Then
          'Check Whether these installment records are
          'already exist int the database or not
          gDbTrans.SqlStmt = "SELECT * FROM LoanInst WHERE LOanId = " & m_LoanID & _
               " ORDER BY InstNo"
          
          Set rstInstDetail = Nothing
          Call gDbTrans.Fetch(rstInstDetail, adOpenStatic)
    
        gDbTrans.SqlStmt = SqlStr
        If Not gDbTrans.SQLExecute Then
           gDbTrans.RollBack
           Exit Function
        End If
        If Not SaveInstallmentDetails(LoanID, rstInstDetail) Then
            gDbTrans.RollBack
            Exit Function
        End If
    End If
    
    gDbTrans.CommitTrans
End If

If m_dbOperation = Insert Then
    RaiseEvent LoanCreated(LoanID)
Else
    RaiseEvent LoanModified(m_LoanID)
End If

SqlStr = ""
'Insert the data into purpose Table als0
   Call SaveLoanPurpose(cmbPurpose, m_SchemeId)

MsgBox "Saved the Details ", vbInformation, wis_MESSAGE_TITLE

SaveLoanAccount = True
ErrLine:
    gDbTrans.RollBack
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    End If
End Function



Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'fraNewLoan.Caption = GetResourceString(754)  'Dummy Value
lblCustID = GetResourceString(49, 60) 'Member No
lblCustName = GetResourceString(35)
lblLoanScheme = GetResourceString(214) 'Loan Scheme
lblLoanAccNo = GetResourceString(58, 36, 60) 'Account No

lblLoanAmount = GetResourceString(80, 91) 'Loan Amount
lblIssueDate = GetResourceString(340)
lblIntrate = GetResourceString(186) 'Rate oF interest
lblInstType = GetResourceString(57) 'Installment
lblDueDate = GetResourceString(209) 'Due date
lblPenalInt = GetResourceString(345, 305) 'Penal Interest
lblNoOfInst = GetResourceString(55) 'Instllment NO


Dim count As Integer
fraGuarantor.Caption = GetResourceString(389) 'Guaranters
For count = 0 To 3
    lblGuaranteer(count).Caption = GetResourceString(389, 35)
Next
'lblGuaranteer1 = GetResourceString(389,35)
'lblGuaranteer2 = GetResourceString(389,35)



lblPurpose = GetResourceString(80, 221) 'LOan Purposie
'lblRegNo =loadresstring(glangoffset+1)

lblPledgeItem = GetResourceString(392)
lblPledgevalue = GetResourceString(392, 393)
lblSeason = GetResourceString(394) 'Season
LblCrop = GetResourceString(395)    'Crop

lblInsureDate = GetResourceString(396, 37) 'Insure Date
lblInsurFee = GetResourceString(396, 40) 'Insurence fee

cmdCancel.Caption = GetResourceString(2) 'Cancel
cmdCreate.Caption = GetResourceString(15) 'Create
cmdDelete.Caption = GetResourceString(14) 'Delete

chkShowLoan.Caption = GetResourceString(80) & " " & _
                        GetResourceString(38, 13) 'Show Loan transation

End Sub

'
Private Sub cmbInstType_Change()
   cmbInstType_Click
   'cmdInst.Enabled = IIf(Val(txtNoOfINst) > 0, True, False)
   cmdInst.Enabled = Val(txtNoOfINst) > 0
   
End Sub

Private Sub cmbInstType_Click()

If cmbInstType.ListIndex < 0 Then Exit Sub


cmdInst.Enabled = True

Dim NoOfInst As Integer
Dim SlNo As Integer
Dim InstAmount As Currency
Dim LoanAmount As Currency
Dim InstDate As String
Dim InstType As Integer
Dim IssueDate As String
'Dim InstTyp
If Not DateValidate(txtIssueDate, "/", True) Then
   Exit Sub
End If
LoanAmount = Val(txtLoanAmount)
IssueDate = GetSysFormatDate(txtIssueDate)

InstType = cmbInstType.ItemData(cmbInstType.ListIndex)
If InstType = Inst_Monthly Then
   InstDate = DateAdd("m", 1, IssueDate)
End If
NoOfInst = Val(txtNoOfINst)
SlNo = 1
If NoOfInst = 0 Then Exit Sub
InstAmount = LoanAmount / NoOfInst

Dim lpCount As Integer
lpCount = 0
'Call GridInit(NoOfInst + 1)
'With grdInst
'   .Row = 0
'   For lpCount = 0 To NoOfInst - 1
'      If .Rows <= .Row + 2 Then .Rows = .Rows + 1
'      .Row = .Row + 1
'      .Col = 0: .Text = SlNo
'      InstDate = DateAdd("m", 1, IssueDate)
'      .Col = 1: .Text = FormatDate(IssueDate)
'      .Col = 2: .Text = InstAmount
'      IssueDate = InstDate
'      SlNo = SlNo + 1
'   Next
'End With

End Sub

Private Sub cmbInstType_KeyUp(KeyCode As Integer, Shift As Integer)

'if typed key is not of charector then
' exit sub
If KeyCode < 64 Or KeyCode > 91 Then  'Check alphbets
    If KeyCode < 47 Or KeyCode > 57 Then 'Check for numeral
        Exit Sub
    End If
End If


Dim CursPos As Integer
Dim count As Integer
Dim loopCount As Integer
Dim strTyped As String
Dim Found As Boolean
strTyped = cmbInstType.Text

CursPos = cmbInstType.SelStart
loopCount = cmbInstType.ListCount - 1
'now compare this text with list
Found = False
For count = 0 To loopCount
    If InStr(1, cmbInstType.List(count), strTyped, vbTextCompare) = 1 Then
        Found = True
        Exit For
    End If
Next count

If Not Found Then Exit Sub
cmbInstType.Text = strTyped & Mid(cmbInstType.List(count), CursPos + 1)
cmbInstType.SelStart = CursPos
cmbInstType.SelStart = CursPos
cmbInstType.SelLength = Len(cmbInstType.Text) - CursPos

End Sub


Private Sub cmbInstType_LostFocus()
If Trim(cmbInstType.Text) = "" Then Exit Sub
Dim count As Integer
Dim Found As Boolean

If cmbInstType.ListIndex >= 0 Then Exit Sub
For count = 0 To cmbInstType.ListCount - 1
    If StrComp(cmbInstType.Text, cmbInstType.List(count), vbTextCompare) = 0 Then
        Found = True
        Exit For
    End If
Next

'IF he has not specified then Installmnet  type is 0
If Not Found Then count = 0
cmbInstType.ListIndex = count

End Sub


Private Sub cmbLoanScheme_Click()

Call ClearControls

m_dbOperation = Insert
If cmbLoanScheme.ListIndex < 0 Then Exit Sub

Call LoadLoanSchemeDetail

If m_CustomerID = 0 Then Exit Sub
If txtCustName = "" Then Exit Sub

Dim SchemeID As Integer
Dim SqlStr As String
Dim count As Integer
Dim rst As Recordset

SchemeID = cmbLoanScheme.ItemData(cmbLoanScheme.ListIndex)

'Clear the Controls
SqlStr = "SELECT * from LoanMaster WHERE SchemeID = " & SchemeID & _
      " And CustomerID = " & m_CustomerID

gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(m_rstCustLoans, adOpenStatic) < 1 Then Exit Sub

count = m_rstCustLoans.recordCount

If count > 1 Then cmdPrev.Enabled = True
cmdNext.Enabled = count
m_rstCustLoans.MoveLast

Call LoadCustomerLoan

End Sub



Private Sub cmbPurpose_KeyUp(KeyCode As Integer, Shift As Integer)

'if typed key is not of charector then
' exit sub
If KeyCode < 64 Or KeyCode > 91 Then  'Check alphbets
    If KeyCode < 47 Or KeyCode > 57 Then 'Check for numeral
        Exit Sub
    End If
End If


Dim CursPos As Integer
Dim count As Integer
Dim loopCount As Integer
Dim strTyped As String
Dim Found As Boolean
strTyped = cmbPurpose.Text

CursPos = cmbPurpose.SelStart
loopCount = cmbPurpose.ListCount - 1
'now compare this text with list
Found = False
For count = 0 To loopCount
    If InStr(1, cmbPurpose.List(count), strTyped, vbTextCompare) = 1 Then
        Found = True
        Exit For
    End If
Next count

If Not Found Then Exit Sub
cmbPurpose.Text = strTyped & Mid(cmbPurpose.List(count), CursPos + 1)
cmbPurpose.SelStart = CursPos
cmbPurpose.SelStart = CursPos
cmbPurpose.SelLength = Len(cmbPurpose.Text) - CursPos

End Sub


Private Sub cmbPurpose_LostFocus()

If Trim(cmbPurpose.Text) = "" Then Exit Sub
Dim count As Integer
Dim Found As Boolean

For count = 0 To cmbPurpose.ListCount - 1
    If StrComp(cmbPurpose.Text, cmbPurpose.List(count), vbTextCompare) = 0 Then
        Found = True
        Exit For
    End If
Next
   
   If Not Found Then Exit Sub
   cmbPurpose.Text = cmbPurpose.List(count)

End Sub


Private Sub cmdCancel_Click()


Unload Me
End Sub

Private Sub cmdCreate_Click()
If Not CheckValidation Then Exit Sub
    
    If Not SaveLoanAccount Then Exit Sub
    
    Dim lstIndex As Integer
    With cmbLoanScheme
        lstIndex = .ListIndex
        .ListIndex = -1
        .ListIndex = lstIndex
    End With
    
End Sub

Private Sub cmdCustID_Click()

If Len(txtCustID) = 0 Then
    Call ClearControls
    If Not m_CustDet Is Nothing Then m_CustDet.NewCustomer
    m_CustomerID = 0
    txtCustName = ""
    ''Search for the user to load
    m_CustomerID = SearchAndGetCustomerID("")
    If m_CustomerID = 0 Then Exit Sub
End If

Dim SqlStr As String
Dim rst As Recordset
Dim memberType As Integer

' This will get the Member name
SqlStr = "SELECT * FROM MemMaster Where AccNum = " & AddQuotes(txtCustID, True)

If m_CustomerID > 0 Then
    SqlStr = "SELECT * FROM MemMaster Where CustomerID = " & m_CustomerID
Else
    If cmbMemberType.ListCount > 1 And cmbMemberType.ListIndex >= 0 Then memberType = cmbMemberType.ItemData(cmbMemberType.ListIndex)
    txtCustName.Caption = GetMemberNameCustIDByMemberNum(Trim(txtCustID), m_CustomerID, memberType)
    If m_CustomerID = 0 Then Exit Sub
    SqlStr = "SELECT * FROM MemMaster Where CustomerID = " & m_CustomerID
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    m_CustomerID = 0
    Exit Sub
Else
    txtCustID = FormatField(rst("AccNum"))
    m_CustomerID = FormatField(rst("customerid"))
    memberType = FormatField(rst("MemberType"))
End If

If m_CustomerID = 0 Then Exit Sub

If cmbMemberType.ListCount > 1 Then Call SetComboIndex(cmbMemberType, , memberType)

m_CustDet.LoadCustomerInfo (m_CustomerID)

'Get The Customer name
txtCustName = m_CustDet.CustomerName(m_CustomerID)

RaiseEvent CustomerSelected(m_CustomerID)

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


Public Function SearchAndGetCustomerID(SearchString As String) As Integer

Dim SqlStr As String
Dim rst As Recordset

If Trim(SearchString) = "" Then _
    SearchString = InputBox("Eneter Name to search", "SearchString")

SqlStr = "SELECT CustomerID,Title + ' ' + FirstName+' '" & _
        " + MiddleName +' ' + LastName as Name FROM NameTab "

If cmbLoanScheme.ListIndex >= 0 Then
    If FormatField(m_rstLoanScheme("OnlyMember")) = True Then
        SqlStr = "SELECT A.CustomerID, Title + ' ' + FirstName+' '" & _
                " + MiddleName +' ' + LastName as Name FROM NameTab A,MemMaster B" & _
                " WHERE A.CustomerID = B.CustomerID"
    End If
End If
      
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
SearchAndGetCustomerID = Val(m_retVar)
End Function



'this sub routine will reload the customer into the form
'With loan details
Private Function LoadCustomerDetails(CustomerID As Long) As Boolean

Dim SqlStr As String
Dim LoanRst As Recordset
Dim CustomerRst As Recordset
Dim SchemeID As Long
Dim retstr As String
Dim strArr() As String

SqlStr = "SELECT CustomerID,Title + ' ' + FirstName + ' ' + " & _
    " Middlename + ' ' + LastName as Name " & _
    " From NameTab WHERE CustomerID = " & CustomerID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(CustomerRst, adOpenStatic) < 1 Then
    MsgBox "Database Error ", vbCritical, wis_MESSAGE_TITLE
    Exit Function
End If

txtCustName = FormatField(CustomerRst("Name"))
Set CustomerRst = Nothing

SqlStr = "SELECT * from LoanMaster where CustomerID  = " & CustomerID
gDbTrans.SqlStmt = SqlStr

If gDbTrans.Fetch(LoanRst, adOpenStatic) < 1 Then Exit Function

'Set the loan schemname on loanschename combo box
Dim ItemCount As Integer
m_SchemeId = FormatField(LoanRst("SchemeID"))
ItemCount = 0
For ItemCount = 0 To cmbLoanScheme.ListCount - 1
    If cmbLoanScheme.ItemData(ItemCount) = SchemeID Then
        cmbLoanScheme.ListIndex = ItemCount
    End If
Next ItemCount

txtIntRate = FormatField(LoanRst("IntRate"))
txtPenalINT = FormatField(LoanRst("PenalIntRate"))
txtIssueDate = FormatField(LoanRst("IssueDate"))
txtDueDate = FormatField(LoanRst("LoanDueDate"))
'Set the appropriate Option button
        
''
txtPledgItem = FormatField(LoanRst("PledgeItem"))
txtPledgevalue = FormatField(LoanRst("PledgeValue"))

'Set the season Combo box
''Oad the Guarantor details
On Error Resume Next
For ItemCount = 0 To 3
    retstr = FormatField(LoanRst("Guarantor" & CStr(ItemCount)))
    Call GetStringArray(retstr, strArr, gDelim)
    txtGName(ItemCount) = strArr(0)
Next

retstr = FormatField(LoanRst("OtherDets"))
Call GetStringArray(retstr, strArr, gDelim)
txtRegNo = strArr(0)
txtInsurFee = strArr(1)
On Error GoTo 0


End Function

Private Sub cmdCustName_Click()
    m_CustDet.ShowDialog
    m_CustomerID = m_CustDet.CustomerID
    txtCustID = GetMemberNumber(m_CustomerID)
    txtCustName = m_CustDet.FullName
    If cmbLoanScheme.ListIndex >= 0 Then Call cmbLoanScheme_Click
    
End Sub

'
Private Sub cmdDelete_Click()
If m_LoanID = 0 Or m_CustomerID = 0 Then Exit Sub
Dim SqlStr As String
Dim rst As Recordset

SqlStr = "SELECT * FROM LoanMaster WHERE Loanid = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub
If MsgBox(GetResourceString(539), _
        vbQuestion + vbYesNo + vbDefaultButton2 _
            , wis_MESSAGE_TITLE) = vbNo Then Exit Sub
    
SqlStr = "SELECT * FROM LoanTrans WHERE Loanid = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then
    If MsgBox(GetResourceString(539) & vbCrLf & _
        GetResourceString(539), vbQuestion + _
    vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
End If

gDbTrans.BeginTrans

'Delete the Transaction details
SqlStr = "DELETE * FROM LoanTrans WHERE " & _
    " LoanID = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo LastLine
End If

'Delete the Installment detials
SqlStr = "DELETE * FROM LoanInst WHERE " & _
     " LoanID = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo LastLine
End If

'Delete the Interest details
SqlStr = "DELETE * FROM LoanIntTrans WHERE " & _
    " LoanID = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo LastLine
End If

'Delete the Master detaisl
SqlStr = "DELETE * FROM LoanMaster WHERE " & _
        " LoanID = " & m_LoanID
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    GoTo LastLine
End If

gDbTrans.CommitTrans

    MsgBox "Loan account deleted", vbInformation, wis_MESSAGE_TITLE
    
    Call cmbLoanScheme_Click
Exit Sub

LastLine:
    
    MsgBox "Unable to delete the loan account", vbInformation, wis_MESSAGE_TITLE
    

End Sub

Private Sub cmdDueDate_Click()
With Calendar
    .Left = Me.Left + fra2.Left + cmdDueDate.Left - (.Width / 2)
    .Top = Me.Top + fra2.Top + cmdDueDate.Top
    If DateValidate(txtDueDate.Text, "/", True) Then
        .selDate = txtDueDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtDueDate.Text = .selDate
End With

End Sub

Private Sub cmdDueDate_LostFocus()
txtDueDate.Enabled = True
End Sub


Private Sub GetGuranteerName(txtGuaranteer As TextBox, txtMem As TextBox)
Dim SqlStr As String
Dim rst As Recordset

Dim SearchString As String
Dim Lret As Long

SearchString = Trim$(txtGuaranteer.Text)
    
If Trim(SearchString) = "" Then _
    SearchString = Trim$(InputBox("Enter Name to search", "SearchString"))
SqlStr = GetResourceString(49, 60)
With gDbTrans
    .SqlStmt = "SELECT A.AccNum as '" & SqlStr & "' ," & _
        " Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName AS '" & GetResourceString(35) & "'" & _
        " FROM MemMaster A, NameTab B WHERE A.CustomerID = B.CustomerID "
    If Trim(SearchString) <> "" Then
        .SqlStmt = .SqlStmt & " AND (FirstName like '" & SearchString & "%' " & _
            " Or LastName like '" & SearchString & "%')"
        .SqlStmt = .SqlStmt & " Order by IsciName"
    Else
        .SqlStmt = .SqlStmt & " Order by val(AccNum)"
    End If
    Lret = .Fetch(rst, adOpenStatic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        Exit Sub
    End If
End With

' Create a report dialog.
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

Me.MousePointer = vbHourglass
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp
Call FillView(m_frmLookUp.lvwReport, rst, True)

m_retVar = ""
m_frmLookUp.Show vbModal
Me.MousePointer = vbNormal
If Val(m_retVar) <= 0 Then Exit Sub

Dim custId As Long
Dim retValue As String
retValue = m_retVar
txtGuaranteer = GetMemberNameCustIDByMemberNum(retValue, custId)
txtGuaranteer.Tag = custId
txtMem.Text = m_retVar

End Sub


Private Sub cmdGuar_Click(Index As Integer)
    Call GetGuranteerName(txtGName(Index), txtGMem(Index))
End Sub

Private Sub cmdInst_Click()

If cmbInstType.ListIndex < 0 Or Val(txtNoOfINst) < 1 Then Exit Sub
If Val(txtIntRate) = 0 Then
    If MsgBox(GetResourceString(505) & vbCrLf & _
        GetResourceString(541), vbQuestion + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
        ActivateTextBox txtIntRate
        Exit Sub
    End If
End If
If Val(txtLoanAmount) = 0 Then
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    txtLoanAmount.SetFocus
    Exit Sub
End If
   
If m_FrmInst Is Nothing Then Set m_FrmInst = New frmLoanInst
m_FrmInst.Operation = InstInsert
'Load the Form
m_FrmInst.LoanID = m_LoanID
m_FrmInst.InterestRate = Val(txtIntRate)
Load m_FrmInst
   
'Load The Details
With m_FrmInst
    'Set the Values
    .cmbInstType.ListIndex = cmbInstType.ListIndex
    .cmbLoanScheme.ListIndex = cmbLoanScheme.ListIndex
    .txtCustName = txtCustName
    .txtIssueDate = txtIssueDate
    .txtLoanAmount = Me.txtLoanAmount
    .txtNoOfINst = txtNoOfINst
    .grdInst.Enabled = True
    .chkEMI = chkEMI
End With

m_FrmInst.Show vbModal, Me
   
End Sub


Private Sub cmdIssueDate_Click()
With Calendar
    .Left = Me.Left + fra2.Left + cmdIssueDate.Left - (.Width / 2)
    .Top = Me.Top + fra2.Top + cmdIssueDate.Top
    If DateValidate(txtIssueDate.Text, "/", True) Then
        .selDate = txtIssueDate.Text
    Else
        .selDate = gStrDate
    End If
    .Show vbModal
    txtIssueDate.Text = .selDate
End With
End Sub


Private Sub cmdNext_Click()
'If m_rstCustLoans Is Nothing Then Exit Sub
If Not m_rstCustLoans.EOF Then m_rstCustLoans.MoveNext

If m_rstCustLoans.EOF Then
    Call ClearControls
    Call LoadLoanSchemeDetail
    cmdPrev.Enabled = True
Else
    LoadCustomerLoan
End If

End Sub

Private Sub cmdPrev_Click()

    If m_rstCustLoans Is Nothing Then Exit Sub
    If m_rstCustLoans.BOF Then Exit Sub
    
    m_rstCustLoans.MovePrevious
    cmdPrev.Enabled = Not m_rstCustLoans.BOF 'm_rstCustLoans.AbsolutePosition
    If m_rstCustLoans.AbsolutePosition = 1 Then cmdPrev.Enabled = False
    
    cmdNext.Enabled = True
    LoadCustomerLoan
End Sub
Private Sub Form_Load()

'Me.Caption = Me.Caption & " - " & gBankName
Call CenterMe(Me)
Call SetKannadaCaption
txtCustName.FONTSIZE = txtCustName.FONTSIZE + 1

cmdPrev.Picture = LoadResPicture(101, vbResIcon)
cmdNext.Picture = LoadResPicture(102, vbResIcon)

'disabling the control
Set m_CustDet = New clsCustReg
'txtFstInst.Enabled = False
cmdCustName.Enabled = False
txtDueDate.Enabled = False
txtCustName = "" 'True

cmbMemberType.Clear
Call LoadMemberTypes(cmbMemberType)
Call LoadLoanSchemes(cmbLoanScheme)
Call LoadAccountGroups(cmbAccGroup)

m_InstDet = False
'Load The instalment types
Dim InstType As wisInstallmentTypes

With cmbInstType
    InstType = Inst_No
    .AddItem ""
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Daily
    .AddItem GetResourceString(410) '"Daily"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Weekly
    .AddItem GetResourceString(411) '"Weekly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_FortNightly
    .AddItem GetResourceString(412) '"Fortnightly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Monthly
    .AddItem GetResourceString(463) '"Monthly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_BiMonthly
    .AddItem GetResourceString(413) '"Bi-Monthly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Quartery
    .AddItem GetResourceString(414) '"Quarterly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_HalfYearly
    .AddItem "6 " & GetResourceString(463) '"Half Yearly"
    .ItemData(cmbInstType.newIndex) = InstType
    
    InstType = Inst_Yearly
    .AddItem "1 " & GetResourceString(208) '"Yearly"
    .ItemData(cmbInstType.newIndex) = InstType
 End With

'Load Season Combobox
 With cmbSeason
    .AddItem ""
    .ItemData(.newIndex) = 0
    .AddItem "Khariff"
    .ItemData(.newIndex) = 1
    .AddItem "Rabi"
    .ItemData(.newIndex) = 2
    .AddItem "T_Belt"
    .ItemData(.newIndex) = 3
    .AddItem "Annual"
    .ItemData(.newIndex) = 4
    .AddItem "Other"
    .ItemData(.newIndex) = 5
 End With
''
'Load the crop combo box.
  With cmbCrop
    .AddItem ""
    .ItemData(.newIndex) = 0
    .AddItem "Oil Seeds"
    .ItemData(.newIndex) = 1
    .AddItem "Sugar Cane"
    .ItemData(.newIndex) = 2
    .AddItem "Paddy"
    .ItemData(.newIndex) = 3
    .AddItem "Food crops"
    .ItemData(.newIndex) = 4
    .AddItem "Horticulture"
    .ItemData(.newIndex) = 5
  End With

Dim SetUp As New clsSetup
m_NoOfGuaranteers = 2
m_NoOfGuaranteers = CInt(SetUp.ReadSetupValue("Loan", "NoOfGuaranteers", "2"))
If m_NoOfGuaranteers < 2 Then m_NoOfGuaranteers = 2
If m_NoOfGuaranteers > 4 Then m_NoOfGuaranteers = 4
Set SetUp = Nothing

Dim extraHeight As Integer
Dim extraGuaranteers As Integer
Dim count As Integer
extraHeight = txtGMem(1).Top - txtGMem(0).Top
For count = 1 To 3
    lblGuaranteer(count).Top = lblGuaranteer(count - 1).Top + extraHeight
    txtGMem(count).Top = txtGMem(count - 1).Top + extraHeight
    txtGMem(count).Height = txtGMem(count - 1).Height
    txtGName(count).Top = txtGName(count - 1).Top + extraHeight
    txtGName(count).Height = txtGName(count - 1).Height
    cmdGuar(count).Top = cmdGuar(count - 1).Top + extraHeight
    cmdGuar(count).Height = cmdGuar(count - 1).Height
Next
    
'Now Set Up the Gurantteers
If m_NoOfGuaranteers > 2 Then
    lblGuaranteer(2).Visible = True
    txtGMem(2).Visible = True
    txtGName(2).Visible = True
    cmdGuar(2).Visible = True
    
    If m_NoOfGuaranteers > 3 Then
        extraHeight = extraHeight * 2
        lblGuaranteer(3).Visible = True
        txtGMem(3).Visible = True
        txtGName(3).Visible = True
        cmdGuar(3).Visible = True
    End If
    fraGuarantor.Height = fraGuarantor.Height + extraHeight + 100
    Frame2.Top = Frame2.Top + extraHeight
    Me.Height = Me.Height + extraHeight + 100
    Call CenterMe(Me)
    
    fraAgriDet.Top = fraAgriDet.Top + extraHeight
    chkShowLoan.Top = chkShowLoan.Top + extraHeight
    cmdDelete.Top = cmdDelete.Top + extraHeight
    cmdCreate.Top = cmdCreate.Top + extraHeight
    cmdCancel.Top = cmdCancel.Top + extraHeight
Else
    extraHeight = 0
    fraGuarantor.Height = fraGuarantor.Height + extraHeight + 100
End If

m_dbOperation = Insert
End Sub


Private Sub Form_Resize()
'Call GridResize
End Sub



Private Sub Form_Terminate()
Set m_frmLookUp = Nothing
Set m_FrmInst = Nothing
Set m_CustDet = Nothing

End Sub

Private Sub Form_Unload(cancel As Integer)
gWindowHandle = 0
RaiseEvent WindowClosed

End Sub

Private Sub lblGuaranteer1_Click()

End Sub

Private Sub m_FrmInst_CancelClicked()
   m_InstDet = False
End Sub

Private Sub m_FrmInst_OkClicked(InstIndianDates() As String, InstAmounts() As Currency)

' The INstallmentde dates & Amounts to varible
Dim count As Integer
On Error Resume Next
ReDim m_InstIndianDates(LBound(InstIndianDates) To UBound(InstIndianDates))
ReDim m_InstAmounts(LBound(InstIndianDates) To UBound(InstIndianDates))
For count = LBound(InstIndianDates) To UBound(InstIndianDates)
   m_InstAmounts(count) = InstAmounts(count)
   m_InstIndianDates(count) = InstIndianDates(count)
Next
m_InstDet = True
End Sub

Private Sub m_frmLookUp_SaveClick(strSelection As String)

m_retVar = strSelection
End Sub

Private Sub m_frmLookUp_CancelClik()
m_retVar = ""
End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_retVar = strSelection
End Sub




Private Sub txtCustName_Change()
    If Len(Trim$(txtCustName)) < 1 Then
        lstHidden.Visible = False
    ElseIf lstHidden.ListCount Then
        lstHidden.Visible = True
    End If
End Sub
Private Sub txtCustName_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Exit Sub
    
    Dim txt As String
    Static SelNo As Integer
    Dim count As Integer
    Dim I As Integer
    txt = txtCustName
    count = lstHidden.ListCount - 1
    For I = 0 To count
        If InStr(1, lstHidden.List(I), txtCustName, vbTextCompare) = 1 Then
            lstHidden.Selected(I) = True
            lstHidden.ListIndex = I
            SelNo = lstHidden.ListIndex
            Exit For
        End If
        If I = count And _
                        Len(txtCustName) > 1 Then _
            lstHidden.Selected(SelNo) = False
    Next I
    

End Sub


Private Sub txtCustName_LostFocus()
lstHidden.Clear
lstHidden.Visible = False
cmdCreate.Default = True
'modified on 30/9/01
End Sub



Private Sub txtDueDate_LostFocus()
'checking for the validate date
If Not DateValidate(txtIssueDate, "/", True) Then Exit Sub
If Not DateValidate(txtDueDate, "/", True) Then Exit Sub

If WisDateDiff(txtIssueDate, txtDueDate) < 0 Then
    MsgBox "Duedate cannot be less than IssueDate"
    cmdDueDate.SetFocus
End If
End Sub















Private Sub txtGMem_Change(Index As Integer)
txtGMem(Index).Tag = ""
End Sub

Private Sub txtGMem_LostFocus(Index As Integer)
    Dim custId As Long
    If Len(txtGName(Index).Tag) > 0 Then Exit Sub
    txtGName(Index).Text = GetMemberNameCustIDByMemberNum(txtGMem(Index).Text, custId)
    txtGName(Index).Tag = custId
End Sub


Private Sub txtGName_LostFocus(Index As Integer)
    If Trim$(txtGName(Index).Text) = "" Then txtGName(Index).Tag = ""
    If Me.ActiveControl.name = cmdGuar(Index).name Then Exit Sub
    If Val(txtGName(Index).Tag) = 0 Then txtGName(Index).Text = ""
End Sub


Private Sub txtIntrate_LostFocus()
If Val(txtIntRate.Text) = 0 And Len(txtIntRate.Text) > 0 Then
 MsgBox "Invalid amount specified...", vbOKOnly
  txtIntRate.SetFocus
 End If
 End Sub


Private Sub txtLoanAmount_Change()
On Error Resume Next
    If ActiveControl.name = txtLoanAmount.name Then
        'txtFstInst = txtLoanAmount
    End If
Err.Clear
End Sub

Private Sub txtNoOfINst_Change()
If cmbInstType.ListIndex > 0 And Val(txtNoOfINst) > 0 Then
   cmdInst.Enabled = True
Else
   cmdInst.Enabled = False
End If
End Sub

Private Sub txtNoOfINst_LostFocus()
Dim NoOfInst As Integer

End Sub


Private Sub txtPenalInt_LostFocus()
If Val(txtPenalINT.Text) = 0 And Len(txtPenalINT.Text) > 0 Then
 MsgBox "Invalid amount specified", vbOKOnly
  txtPenalINT.SetFocus
  End If
   End Sub


Private Sub txtPledgItem_LostFocus()
If Not Val(txtPledgItem.Text) = 0 And Len(txtPledgItem.Text) > 0 Then
 MsgBox "Invalid item specified...", vbOKOnly
 txtPledgItem.SetFocus
End If
 End Sub


