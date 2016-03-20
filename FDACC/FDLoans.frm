VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFDLoans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loans"
   ClientHeight    =   5700
   ClientLeft      =   2280
   ClientTop       =   1740
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   6660
      TabIndex        =   12
      Top             =   5280
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   5085
      Left            =   90
      TabIndex        =   13
      Top             =   60
      Width           =   7665
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   285
         Left            =   3390
         TabIndex        =   0
         Top             =   780
         Width           =   255
      End
      Begin VB.CommandButton cmdRepay 
         Caption         =   "Repay"
         Height          =   315
         Left            =   4860
         TabIndex        =   10
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   150
         TabIndex        =   28
         Top             =   2490
         Width           =   7365
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Undo"
         Height          =   315
         Left            =   3510
         TabIndex        =   11
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   7365
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Height          =   315
         Left            =   6330
         TabIndex        =   9
         Top             =   2640
         Width           =   1125
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   780
         Width           =   945
      End
      Begin VB.TextBox txtAvailable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1770
         Width           =   1155
      End
      Begin VB.TextBox txtLoan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1110
         Width           =   1125
      End
      Begin VB.TextBox txtSanctioned 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   780
         Width           =   1065
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   150
         TabIndex        =   14
         Top             =   3030
         Width           =   7365
      End
      Begin VB.TextBox txtInterest 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtInterestAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox txtIssuedAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1440
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   1815
         Left            =   150
         TabIndex        =   24
         Top             =   3150
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   3201
         _Version        =   393216
         ScrollBars      =   2
         AllowUserResizing=   1
      End
      Begin VB.Label lblCaption1 
         Caption         =   "New loans will be issued only after deducting interest prevailing on previous loans first."
         Height          =   495
         Left            =   3960
         TabIndex        =   27
         Top             =   1860
         Width           =   3435
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCaption 
         Caption         =   "Total loan amount drawn on this deposit:"
         Height          =   285
         Left            =   2550
         TabIndex        =   26
         Top             =   270
         Width           =   4455
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label lblLoanAmtAvail 
         Caption         =   "Available loan amount : "
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   1830
         Width           =   2145
      End
      Begin VB.Label lblPrevLoanAmt 
         Caption         =   "Previous loan amount : "
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   2130
         Width           =   2145
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposit amount : "
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label lblLoanSanctioned 
         Caption         =   "Sanctioned amount : "
         Height          =   255
         Left            =   3990
         TabIndex        =   19
         Top             =   810
         Width           =   2265
      End
      Begin VB.Label lblRateofIntForLoans 
         Caption         =   "Rate of interest for loans:"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   1470
         Width           =   2145
      End
      Begin VB.Label lblDepositNo 
         Caption         =   "Deposit No: "
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label lblLessIntOnPrevLoan 
         Caption         =   "Less interest on previous loans: "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3990
         TabIndex        =   16
         Top             =   1110
         Width           =   2295
      End
      Begin VB.Label lblTotAmtIssued 
         Caption         =   "Total amount to be issued:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   1440
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmFDLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_AccID As Long
Public m_DepositID As Long
Private M_setUp As New clsSetup
Private Sub SetKannadaCaption()
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
      Ctrl.Font.Size = gFontSize
    End If
Next Ctrl

Me.lblDepositNo.Caption = LoadResString(gLangOffSet + 241)   ' "«œ˙Û∆æÚ ÕÆ≤˙Â"
Me.lblCaption.Caption = LoadResString(gLangOffSet + 242)   '"§ «œ˙Û∆æÚÆıÙ ∆˙ÙÛë¬ ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblDate.Caption = LoadResString(gLangOffSet + 37)    '"ä¬ÒÆ∞"
Me.lblDepAmount.Caption = LoadResString(gLangOffSet + 243)    '"«œ˙Û∆æÚ ∆˙˜¿ﬁ"
Me.lblRateofIntForLoans.Caption = LoadResString(gLangOffSet + 244)    '"ÕÒ»¡ ∆˙ÙÛë¬ ƒà‹ÆıÙ ¡«"
Me.lblLoanAmtAvail.Caption = LoadResString(gLangOffSet + 245)    '"ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblPrevLoanAmt.Caption = LoadResString(gLangOffSet + 246)    '"ñÆä¬ ÕÒ»¡ ∆˙˜¿ﬁ"
Me.lblLoanSanctioned.Caption = LoadResString(gLangOffSet + 247)    '"∆ÙÆ∏˜«Ò¡ ÕÒ»"
Me.lblLessIntOnPrevLoan.Caption = LoadResString(gLangOffSet + 248)    '"ñÆä¬ ÕÒ»¡ ƒà‹ "
Me.lblTotAmtIssued.Caption = LoadResString(gLangOffSet + 249)    '"´ªÙ⁄ ∆˙˜¿ﬁ ∞˙˜Ω≈˙Û∞ÒÇ¡Ù‡"
Me.lblCaption1.Caption = LoadResString(gLangOffSet + 250)    '"ñÆä¬ ÕÒ»¡ ƒà‹ÆıÙ¬Ù· ∞ ˙¡Ù Œ˙˜Õ ÕÒ»∆¬Ù· ∞˙˜Ω≈˙Û∞Ù"
Me.cmdUndo.Caption = LoadResString(gLangOffSet + 19)    '"°íÕÙ(∞˙˜¬˙ÆıÙ)"
Me.cmdRepay.Caption = LoadResString(gLangOffSet + 20)    '"∆Ù«Ù√Ò∆â"
Me.cmdAccept.Caption = LoadResString(gLangOffSet + 4)    '"°ÆÇÛ∞êÕÙ"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 11)    '"∆ÙÙ∂Ù’"
End Sub


Private Sub cmdAccept_Click()

Dim TransType As wisTransactionTypes

'Check out the date
    If Not DateValidate(txtDate.Text, "/", True) Then
        'MsgBox "Date of transaction not in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Sub
    End If

'See if account has already been matured or closed
    gDBTrans.SQLStmt = "Select * from FDmaster where AccID = " & m_AccID & _
                        " and DepositID = " & m_DepositID '& " order by TransID"
    
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - Error"
        Exit Sub
    End If
    'Check for closing
'''    If FormatField(gDBTrans.Rst("ClosedDate")) <> "" Then
'''        'MsgBox "This deposit has already been closed !", vbExclamation, gAppName & " - Error"
'''        MsgBox LoadResString(gLangOffSet + 524), vbExclamation, gAppName & " - Error"
'''        Exit Sub
'''    End If

    If WisDateDiff(FormatField(gDBTrans.Rst("MaturityDate")), txtDate.Text) >= 0 Then
        'MsgBox "You have specified a date that is later than the maturity date i.e " & FormatField(gDBTrans.Rst("MaturityDate")) & vbCrLf & "This means that you are trying to issue loans on a matured deposit !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 578) & FormatField(gDBTrans.Rst("MaturityDate")) & vbCrLf & "This means that you are trying to issue loans on a matured deposit !", vbExclamation, gAppName & " - Error"
'''        ActivateTextBox txtDate
'''        Exit Sub
    End If
    
   
'Check date range
    gDBTrans.SQLStmt = "Select TOP 1 TransDate from FDTrans where AccID = " & m_AccID & _
                        " and DepositID = " & m_DepositID & " order by TransID desc"
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbCritical, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbCritical, gAppName & " - Error"
        Exit Sub
    End If
'''    If WisDateDiff(FormatField(gDBTrans.Rst("TransDate")), txtDate.Text) < 0 Then
'''        'MsgBox "Date of transaction is lesser than the previous transaction date", vbExclamation, gAppName & " - Error"
'''        MsgBox LoadResString(gLangOffSet + 572), vbExclamation, gAppName & " - Error"
'''        ActivateTextBox txtDate
'''        Exit Sub
'''    End If

'Check out if the interest rate is valid
    If Val(txtInterest.Text) < 0 Then
        'MsgBox "Interest rate has not been specified for this period." & vbCrLf & vbCrLf & "Please set the value of interest for this period in the properties of this account !", vbInformation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 505) & vbCrLf & vbCrLf & "Please set the value of interest for this period in the properties of this account !", vbInformation, gAppName & " - Error"
        Exit Sub
    End If

'Check out the sanctioned amount
    If Not CurrencyValidate(txtSanctioned.Text, False) Then
        'MsgBox "Invalid amount sanctioned !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 506), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtSanctioned
        Exit Sub
    End If
    
    If Val(txtSanctioned.Text) > Val(txtAvailable.Text) Then
        'MsgBox "Loan amount sanctioned is greater than the available loan amount !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 581), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtSanctioned
        Exit Sub
    End If
    
'Get new transaction ID
    Dim TransID As Long
    Dim Balance As Currency
    Dim Loan As Boolean
    
    Loan = True
    gDBTrans.SQLStmt = "Select TOP 1 TransID,Balance from FDTrans where AccID = " & m_AccID & _
                    " and DepositID = " & m_DepositID & " ANd Loan = " & Loan & _
                    " Order by TransID desc"
    
    If gDBTrans.SQLFetch < 0 Then
        'MsgBox "Error in performing transaction !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 640), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    If gDBTrans.Rst.RecordCount > 0 Then
        TransID = Val(FormatField(gDBTrans.Rst("TransID"))) + 1
        Balance = CCur(FormatField(gDBTrans.Rst("Balance")))
    Else
        TransID = 1
    End If
    
'Start data base transactions
gDBTrans.BeginTrans
    'First insert any interest of previous loans

'''    TransType = wCharges
'''    Loan = True
'''        gDBTrans.SQLStmt = "Insert into FDTrans (AccID, DepositID,Loan, TransID, TransType, " & _
'''                            " TransDate, Amount,Balance, " & _
'''                            " Particulars) values ( " & _
'''                            m_AccID & "," & _
'''                            m_DepositID & "," & _
'''                            Loan & ", " & _
'''                            TransID & "," & _
'''                            TransType & "," & _
'''                            "#" & FormatDate(txtDate.Text) & "#," & _
'''                            CCur(Val(Me.txtInterestAmount.Text)) & "," & _
'''                            Balance & ", " & _
'''                            "'" & "By interest" & "'" & _
'''                            ")"
'''        If Not gDBTrans.SQLExecute Then
'''            gDBTrans.RollBack
'''            Exit Sub
'''        End If
'''        TransID = TransID + 1

    
    'Now insert record of the new loan
    
    TransType = wWithDraw
    Loan = True
    Balance = CCur(txtSanctioned.Text) + Balance
    gDBTrans.SQLStmt = "Insert into FDTrans (AccID, DepositID, Loan,TransID, TransType, " & _
                        " TransDate, Amount, Balance, " & _
                        " Particulars) values ( " & _
                        m_AccID & "," & _
                        m_DepositID & "," & _
                        Loan & ", " & _
                        TransID & "," & _
                        TransType & "," & _
                        "#" & FormatDate(txtDate.Text) & "#," & _
                        CCur(txtSanctioned.Text) & "," & _
                        Balance & ", " & _
                        "'" & "To Loans" & "'" & _
                        ")"
                    
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        Exit Sub
    End If

'COmmit transactions
gDBTrans.CommitTrans

'Udate date with todays date (By default)
    txtDate.Text = FormatDate(gStrDate)
    txtSanctioned.Text = ""
'Update the details on the UI
    Call UpdateUserInterface
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDate_Click()
With Calendar
    .Left = Me.Left + cmdDate.Left - .Width / 2
    .Top = Me.Top + cmdDate.Top
    .SelDate = txtDate.Text
    .Show vbModal
    Me.txtDate.Text = .SelDate
End With
End Sub

Private Sub cmdRepay_Click()
frmFDRepay.Show vbModal
Call UpdateUserInterface
End Sub

Private Sub cmdUndo_Click()
Dim Loan As Boolean
Dim TransID As Long
'Get the last transaction ID
Loan = True
    gDBTrans.SQLStmt = "Select Top 1 TransID, TransType from FDTrans where AccID = " & _
                        m_AccID & " and DepositID = " & m_DepositID & _
                        " And Loan = " & Loan & " order by TransID desc"
    Call gDBTrans.SQLFetch
    If FormatField(gDBTrans.Rst("TransID")) < 1 Then
        'MsgBox "You do not have any loans on this deposit !", vbInformation, gAppName & " - Message"
        MsgBox LoadResString(gLangOffSet + 582), vbInformation, gAppName & " - Message"
        Exit Sub
    End If
    TransID = FormatField(gDBTrans.Rst("TransID"))

'Check out the transaction before the last transaction, because it may the interest
'added. Since we are performing interest charges automatically, we've got to remove
'this also automatically
gDBTrans.SQLStmt = "Select transType from FDTrans where " & _
                    " AccID = " & m_AccID & _
                    " and DepositID = " & m_DepositID & _
                    " and Loan = " & Loan & " And TransID = " & TransID - 1
Call gDBTrans.SQLFetch
Dim TransType As wisTransactionTypes
TransType = FormatField(gDBTrans.Rst("TransType"))

'Confirm about transaction
'If MsgBox("Are you sure you want to undo a previous loan transaction ?", vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
If MsgBox(LoadResString(gLangOffSet + 583), vbQuestion + vbYesNo, gAppName & " - Error") = vbNo Then
    Exit Sub
End If

gDBTrans.BeginTrans
    'Remove the last transaction
    gDBTrans.SQLStmt = "Delete from FDTrans where AccID = " & m_AccID & _
                        " and DepositID = " & m_DepositID & _
                        " And LOan = " & Loan & " and TransID = " & TransID
    If Not gDBTrans.SQLExecute Then
        'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
        MsgBox LoadResString(gLangOffSet + 609), vbCritical, gAppName & " - Critical Error"
        gDBTrans.RollBack
        Exit Sub
    End If


    'Remove the transaction previous to the one removed only if it is of type
    'CHARGES levied. b'cause this record would have been added automatically
    'If TransType = wCharges Then
        gDBTrans.SQLStmt = "Delete from FDTrans where AccID = " & m_AccID & _
                            " and DepositID = " & m_DepositID & _
                            " And Loan = " & Loan & " And TransID = " & TransID - 1
        If Not gDBTrans.SQLExecute Then
            'MsgBox "Unable to undo transactions !", vbCritical, gAppName & " - Critical Error"
            MsgBox LoadResString(gLangOffSet + 609), vbCritical, gAppName & " - Critical Error"
            gDBTrans.RollBack
            Exit Sub
        End If
    'End If
gDBTrans.CommitTrans

'Udate date with todays date (By default)
    txtDate.Text = FormatDate(gStrDate)

Call UpdateUserInterface
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

'Centre the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    txtDate.Text = FormatDate(gStrDate)
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

'set kannada fonts
Call SetKannadaCaption
'Initialize the grid
    grd.Rows = 6
    grd.Cols = 5
    grd.FixedRows = 1
    grd.FixedCols = 0
    grd.Row = 0
    grd.Col = 0: grd.Text = LoadResString(gLangOffSet + 37): grd.ColWidth(0) = (grd.Width / 5) '- (grd.Width / 12) 'Some shit adjustment '"Date"
    grd.Col = 1: grd.Text = LoadResString(gLangOffSet + 235): grd.ColWidth(1) = grd.Width / 5 '"Loan Amount"
    grd.Col = 2: grd.Text = LoadResString(gLangOffSet + 216): grd.ColWidth(2) = grd.Width / 5 ' "Repayment"
    grd.Col = 3: grd.Text = LoadResString(gLangOffSet + 274): grd.ColWidth(3) = grd.Width / 6 ' "Interest"
    grd.Col = 4: grd.Text = LoadResString(gLangOffSet + 42): grd.ColWidth(3) = grd.Width / 5 ' "Balance"
'Fill up the two module level variables
    m_AccID = frmFDAcc.m_AccID
    
'Obtain the rate of interest as applicable to this deposit
    Dim Days As Integer
    Dim Loan As Boolean
    Dim TransType As wisTransactionTypes
    Loan = False
    TransType = wDeposit
    gDBTrans.SQLStmt = "Select  * from FDMaster where " & _
                        " AccID = " & m_AccID & _
                        " and DepositID = " & m_DepositID
    If gDBTrans.SQLFetch <= 0 Then
        'MsgBox "Error accessing data base !", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 601), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
    'Check if deposit is closed
    If FormatField(gDBTrans.Rst("ClosedDate")) <> "" Then
        cmdUndo.Enabled = False: cmdAccept.Enabled = False: cmdRepay.Enabled = False
    Else
        cmdUndo.Enabled = True: cmdAccept.Enabled = True: cmdRepay.Enabled = True
    End If
    
    Days = WisDateDiff(FormatField(gDBTrans.Rst("CreateDate")), FormatField(gDBTrans.Rst("MaturityDate")))
   
   Dim SchemeName As String
   
   
    SchemeName = "0.5_1_Loan"
   
    If Days > 0 And Days <= 30 Then
        'txtInterest.Text = M_setUp.ReadSetupValue("FDAcc", "1/2_1_Loan", 0)
        SchemeName = "0.5_1_Loan"
    ElseIf Days > 30 And Days <= 45 Then
        SchemeName = "1_1.5_Loan"
    ElseIf Days > 45 And Days <= 90 Then
        SchemeName = "1.5_3_Loan"
    ElseIf Days > 90 And Days <= 180 Then
        SchemeName = "3_6_Loan"
    ElseIf Days > 180 And Days <= 365 Then
        SchemeName = "6_12_Loan"
    ElseIf Days > 365 And Days <= 730 Then
        SchemeName = "12_24_Loan"
    ElseIf Days > 730 And Days < 1096 Then
        SchemeName = "24_36_Loan"
   ElseIf Days >= 1096 Then
         SchemeName = "36_Above_Loan"
    End If
    
   Dim ClsInt As New clsInterest
    txtInterest.Text = ClsInt.InterestRate(wis_FDAcc, SchemeName, FormatDate(gStrDate))
    
    txtInterest.Text = Format(txtInterest.Text, "#0.00")
    
Call UpdateUserInterface
cmdUndo.Enabled = True: cmdAccept.Enabled = True: cmdRepay.Enabled = True

End Sub

Private Sub UpdateUserInterface()
Dim MasterRst As Recordset
Dim TransRst As Recordset
Dim TransType As wisTransactionTypes
Dim Balance As Currency

'Get details of this account and deposit
    gDBTrans.SQLStmt = "Select * from FDMaster where AccID = " & _
                    m_AccID & " and DepositID = " & m_DepositID
    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        Set MasterRst = gDBTrans.Rst.Clone
    End If

'Get The TransCtion Details of this deposit
    gDBTrans.SQLStmt = "Select * from FDTrans where AccID = " & _
                    m_AccID & " and DepositID = " & m_DepositID
    
    If gDBTrans.SQLFetch <= 0 Then
        Exit Sub
    Else
        Set TransRst = gDBTrans.Rst.Clone
    End If

'Get the Maturity date
    Dim MatDate As String
    MatDate = FormatField(MasterRst("MaturityDate"))
    
'Update command buttons only if deposit is not closed
If FormatField(MasterRst("ClosedDate")) = "" Then
    If TransRst.RecordCount = 1 Then          'No loans on this account
        cmdRepay.Enabled = False
        cmdUndo.Enabled = False
    Else
        cmdRepay.Enabled = True
        cmdUndo.Enabled = True
    End If
Else
    cmdRepay.Enabled = False: cmdUndo.Enabled = False: cmdAccept.Enabled = False
End If
    
    Dim LoanAmount As Currency
    Dim TransDate As String
    Dim i As Integer
    LoanAmount = 0
    grd.Rows = 1: grd.Rows = 7
    grd.Row = 0
    txtDeposit.Text = ""
    
    While Not TransRst.EOF
        'Register the TransType first
        TransType = FormatField(TransRst("TransTYpe"))
        
        'Check out if field is displayable
        If TransRst("Amount") = 0 Then
            'If TransType = wCharges Then GoTo NextRecord
            
        End If
        
        If TransType = wDeposit And TransRst("Loan") = False Then  'Deposit
            txtDeposit.Text = Val(txtDeposit.Text) + Val(FormatField(TransRst("Amount")))
            'grd.Row = grd.Row - 1
            GoTo NextRecord
        End If
        
        'if record is not related to loan then read next record
        If TransRst("Loan") = False Then
            'grd.Row = grd.Row - 1
            GoTo NextRecord
        End If
        
        If Balance <> TransRst("Balance") Then
            'Set new row number for displaying the record
            If grd.Rows = grd.Row + 2 Then grd.Rows = grd.Rows + 3
            grd.Row = grd.Row + 1
            Balance = TransRst("Balance")
            grd.Col = 4: grd.Text = FormatCurrency(Balance)
        End If
        
        If TransType = wWithDraw Then  'Loans Drawn
            grd.Col = 0: grd.Text = FormatField(TransRst("TransDate"))
            grd.Col = 1: grd.Text = FormatField(TransRst("Amount"))
            LoanAmount = LoanAmount + CCur(FormatField(TransRst("Amount")))
        ElseIf TransType = wDeposit Then  ' Loan Repaument
            grd.Col = 0: grd.Text = FormatField(TransRst("TransDate"))
            grd.Col = 2: grd.Text = FormatField(TransRst("Amount"))
            LoanAmount = LoanAmount - CCur(FormatField(TransRst("Amount")))
        'ElseIf TransType = wCharges Then  'Interest charged
            grd.Col = 0: grd.Text = FormatField(TransRst("TransDate"))
            grd.Col = 3: grd.Text = FormatField(TransRst("Amount"))
        End If
        
        TransDate = FormatField(TransRst("TransDate"))
NextRecord:
        TransRst.MoveNext
    Wend
    txtLoan.Text = FormatCurrency(LoanAmount)
    
    
'Calculate the available amount form loan ( 80 % of deposit - Total loans drawn till date)
                                            'is what is available. Take this from setup
    Dim LoanPercent As Single
    Dim SetupClass As New clsSetup
    LoanPercent = SetupClass.ReadSetupValue("FDAcc", "MaxLoanPercent", "80")
    If LoanPercent > 1 Then
        LoanPercent = LoanPercent / 100
    End If
    txtAvailable.Text = FormatCurrency((Val(txtDeposit.Text) * LoanPercent) - Val(txtLoan.Text))
    Set SetupClass = Nothing

'Calculate the interest for loan if a loan has been drawn previously
    Dim Days As Integer

    On Error Resume Next
    
    'See if deposit has matured
    If TransDate <> "" Then
        If WisDateDiff(txtDate.Text, MatDate) <= 0 Then
'''           txtInterestAmount.Text = ComputeFDInterest(LoanAmount, Transdate, MatDate, CSng(txtInterest.Text), True)
        Else
            txtInterestAmount.Text = ComputeFDInterest(LoanAmount, TransDate, txtDate.Text, CSng(txtInterest.Text), True)
        End If
'''        txtInterestAmount.Text = FormatCurrency(Val(txtInterestAmount.Text) \ 1)
    Else
        txtInterestAmount.Text = "0.00"
    End If
    
    lblCaption.Caption = LoadResString(gLangOffSet + 242) & " " & LoadResString(gLangOffSet + 312) & LoanAmount  '"Total loans on this deposit : Rs."
cmdUndo.Enabled = True: cmdAccept.Enabled = True: cmdRepay.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFDLoans = Nothing
   
End Sub



Private Sub txtDate_LostFocus()
If Me.ActiveControl.Name = cmdAccept.Name Then txtDeposit.SetFocus

If Not DateValidate(txtDate.Text, "/", True) Then
    Exit Sub
End If
Call UpdateUserInterface

End Sub


Private Sub txtInterestAmount_Change()
'COmpute the total amount to be actully issued
txtIssuedAmount.Text = FormatCurrency(Val(txtSanctioned.Text) - Val(txtInterestAmount.Text))

End Sub

Private Sub txtSanctioned_Change()
'COmpute the total amount to be actully issued
'''txtIssuedAmount.Text = FormatCurrency(Val(txtSanctioned.Text) - Val(txtInterestAmount.Text))
End Sub


