VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmFDRenew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renew Matured FD"
   ClientHeight    =   6240
   ClientLeft      =   1350
   ClientTop       =   1995
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frarenew 
      Caption         =   "Renew details"
      Height          =   2415
      Left            =   30
      TabIndex        =   7
      Top             =   3120
      Width           =   7485
      Begin VB.TextBox txtInterest 
         Height          =   315
         Left            =   5760
         TabIndex        =   17
         Top             =   920
         Width           =   585
      End
      Begin VB.TextBox txtDays 
         Height          =   345
         Left            =   2190
         TabIndex        =   15
         Top             =   920
         Width           =   585
      End
      Begin VB.TextBox txtMatureDate 
         Height          =   315
         Left            =   5760
         TabIndex        =   13
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox txtRenewDate 
         Height          =   315
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton cmdDepositDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3570
         TabIndex        =   9
         Top             =   420
         Width           =   315
      End
      Begin VB.CommandButton cmdMatureDate 
         Caption         =   ".."
         Height          =   315
         Left            =   7020
         TabIndex        =   12
         Top             =   420
         Width           =   315
      End
      Begin VB.TextBox txtCertificate 
         Height          =   345
         Left            =   2190
         TabIndex        =   24
         Top             =   1920
         Width           =   5085
      End
      Begin WIS_Currency_Text_Box.CurrText txtDepositAmount 
         Height          =   345
         Left            =   2190
         TabIndex        =   19
         Top             =   1425
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin WIS_Currency_Text_Box.CurrText txtMatureAmount 
         Height          =   375
         Left            =   5760
         TabIndex        =   22
         Top             =   1425
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblDepAmount 
         Caption         =   "Deposit amount (Rs) : "
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1425
         Width           =   2025
      End
      Begin VB.Label lblInterest 
         Caption         =   "Interest (%) : "
         Height          =   315
         Left            =   3900
         TabIndex        =   16
         Top             =   915
         Width           =   1125
      End
      Begin VB.Label lblDays 
         Caption         =   "Days : "
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   915
         Width           =   1185
      End
      Begin VB.Label lblMatureDate 
         Caption         =   "Matures on : "
         Height          =   315
         Left            =   3930
         TabIndex        =   11
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label lblrenewDate 
         Caption         =   "Renew date :"
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label lblMatureAmount 
         Caption         =   "Maturity amount (Rs) : "
         Height          =   315
         Left            =   3870
         TabIndex        =   21
         Top             =   1425
         Width           =   1725
      End
      Begin VB.Label lblCertificate 
         Caption         =   "Certificate No :"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   1995
      End
   End
   Begin VB.Frame fraOldFD 
      Caption         =   "Mature details"
      Height          =   1785
      Left            =   30
      TabIndex        =   28
      Top             =   1440
      Width           =   7485
      Begin VB.Label txtOldMatAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5790
         TabIndex        =   20
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label txtOldDepAmount 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2190
         TabIndex        =   23
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label txtOldMatDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5790
         TabIndex        =   34
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label txtOldDepdate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2190
         TabIndex        =   32
         Top             =   780
         Width           =   1545
      End
      Begin VB.Label txtOldCertificate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2190
         TabIndex        =   30
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label lblOldMatAmount 
         Caption         =   "Maturity amount (Rs) : "
         Height          =   315
         Left            =   3810
         TabIndex        =   36
         Top             =   1380
         Width           =   1785
      End
      Begin VB.Label lblOldDepAmount 
         Caption         =   "Deposit amount (Rs) : "
         Height          =   315
         Left            =   90
         TabIndex        =   35
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label lblOldMatDate 
         Caption         =   "Matured Date"
         Height          =   315
         Left            =   3870
         TabIndex        =   33
         Top             =   840
         Width           =   1725
      End
      Begin VB.Label lblOldDepDate 
         Caption         =   "Deposit Date"
         Height          =   315
         Left            =   90
         TabIndex        =   31
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label lblOldCertificate 
         Caption         =   "Certificate No :"
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   400
      Left            =   3450
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdRenew 
      Caption         =   "Renew"
      Height          =   400
      Left            =   4845
      TabIndex        =   26
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   6210
      TabIndex        =   27
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame fraAccount 
      Caption         =   "account details"
      Height          =   1335
      Left            =   30
      TabIndex        =   0
      Top             =   180
      Width           =   7485
      Begin VB.CommandButton cmdTransDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3540
         TabIndex        =   3
         Top             =   780
         Width           =   315
      End
      Begin VB.TextBox txtTransDate 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label txtAccNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6300
         TabIndex        =   6
         Top             =   750
         Width           =   915
      End
      Begin VB.Label lblTransdate 
         Caption         =   "Transaction  Date"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No"
         Height          =   315
         Left            =   4230
         TabIndex        =   5
         Top             =   810
         Width           =   1725
      End
      Begin VB.Label LblAccHolder 
         Caption         =   "Name of the account Holder"
         Height          =   330
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmFDRenew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_AccID As Long



Public Property Let AccountId(NewValue As Long)
m_AccID = NewValue
End Property


Private Function CloseMFD() As Boolean
    Dim TransDate As Date
    Dim rst As Recordset
    If Not DateValidate(Trim$(txtTransDate.Text), "/", True) Then
        'MsgBox "Please specify deposit date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtTransDate
        Exit Function
    Else
        TransDate = GetSysFormatDate(txtTransDate.Text)
    End If
    
    
    'Check the Amount to be withdrawn
    Dim Amount As Currency
    If Not CurrencyValidate(txtOldMatAmount, False) Then Exit Function
    Amount = Val(txtOldMatAmount.Caption)
    
    'Now withdraw this amount from The MFD Account
    Dim transType As wisTransactionTypes
    Dim ContraTransType As wisTransactionTypes
    Dim TransID As Long
    
    gDbTrans.SqlStmt = "SELECT MAX(Transid) From MatFDTrans WHERE " & _
        " ACCID = " & m_AccID
    Call gDbTrans.Fetch(rst, adOpenStatic)
    TransID = FormatField(rst(0)) + 1
    transType = wWithdraw
    gDbTrans.SqlStmt = "INSERT INTO MatFDTrans (AccID, TransID, " & _
        " TransDate, Amount,TransType,Balance,Particulars,UserID) VALUES " & _
        "(" & m_AccID & "," & _
        TransID & "," & _
        "#" & TransDate & "#, " & _
        Amount & ", " & _
        transType & "," & _
        " 0,'WithDrwan', " & _
        gUserID & " )"
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        'MsgBox "Unable to do this operation"
        MsgBox GetResourceString(1), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    gDbTrans.CommitTrans
    
CloseMFD = True
End Function

Private Function RenewMFD() As Boolean
    Dim TransDate As Date
    Dim rst As Recordset
    If Not DateValidate(Trim$(txtTransDate.Text), "/", True) Then
        'MsgBox "Please specify deposit date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtTransDate
        Exit Function
    Else
        TransDate = GetSysFormatDate(txtTransDate.Text)
    End If
    
    Dim EffectiveDate As Date
    If Not DateValidate(Trim$(txtRenewDate.Text), "/", True) Then
        'MsgBox "Please specify deposit date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtRenewDate
        Exit Function
    Else
        EffectiveDate = GetSysFormatDate(txtRenewDate.Text)
    End If
    

'Validate the Maturity date
    Dim MatureDate As Date
    If Not DateValidate(Trim$(txtMatureDate.Text), "/", True) Then
        'MsgBox "Please specify maturity date in DD/MM/YYYY format !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(573), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtMatureDate
        Exit Function
    Else
        MatureDate = GetSysFormatDate(txtMatureDate.Text)
    End If
    
    'Check this date with transaction date and Renew date
    If DateDiff("d", TransDate, MatureDate) <= 0 Then
        'MsgBox "Invalid date specified!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtMatureDate
        Exit Function
    End If
    If WisDateDiff(txtRenewDate, txtMatureDate) <= 0 Then
        'MsgBox "Invalid date specified!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtMatureDate
        Exit Function
    End If
    
    'Get the Last Transaction date
    gDbTrans.SqlStmt = "SELECT Max(TransDate) as MaxDate From MatFDTrans " & _
        " Where accID = " & m_AccID
    Call gDbTrans.Fetch(rst, adOpenStatic)
    If DateDiff("d", rst("MaxDate"), TransDate) <= 0 Then
        'MsgBox "You have specified a transaction date that is earlier than the last date of transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(572), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtTransDate
        Exit Function
    End If
    

'Validate the rate of interest
    Dim Interest As Double
    If Val(Trim$(txtInterest.Text)) < 1 Or Val(Trim$(txtInterest.Text)) >= 100 Then
        'MsgBox "Invalid rate of interest specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(505), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtInterest
        Exit Function
    Else
        Interest = Val(txtInterest.Text)
    End If
        
'Validate the deposit Amount
    Dim DepositAmount As Currency
'    If Not CurrencyValidate(txtDepositAmount.Text, False) Then
'        'MsgBox "Invalid deposit amount specified !", vbExclamation, gAppName & " - Error"
'        MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
'        ActivateTextBox txtDepositAmount
'        Exit Function
'    End If
    DepositAmount = txtDepositAmount
    

'Validate the maturity Amount
    Dim MatureAmount As Currency
'    If Not CurrencyValidate(txtMatureAmount.Text, False) Then
'        'MsgBox "Invalid maturity amount specified !", vbExclamation, gAppName & " - Error"
'        MsgBox GetResourceString(506), vbExclamation, gAppName & " - Error"
'        ActivateTextBox txtMatureAmount
'        Exit Function
'    End If
    MatureAmount = txtMatureAmount


'CERTIFICATE NO
'Get the Particulars(Certificate no)
    Dim Particulars As String
    Particulars = "By Deposit"
    If Trim(txtCertificate.Text) = "" Then
        'If MsgBox("Certificate no not speicfied " & vbCrLf & _
            " Do yoy want to continue", vbInformation, "wis_message_title") = vbNo Then
        If MsgBox(GetResourceString(337, 60, 296) & _
                GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            ActivateTextBox txtCertificate
            Exit Function
        End If
'        Exit Function
    ElseIf Val(txtCertificate.Text) = 0 Then
        'If MsgBox("Invalid Certificate no speicfied " & vbCrLf & _
            " Do yoy want to continue", vbInformation, "wis_message_title") = vbNo Then
        If MsgBox(GetResourceString(337, 60, 296) & _
                GetResourceString(541), vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then
            ActivateTextBox txtCertificate
            Exit Function
        End If
    Else
        Particulars = Trim(txtCertificate.Text)
        If Len(Particulars) > 49 Then Particulars = Left(Particulars, 49)
    End If
    
    Dim CustomerID As Long
    Dim AccNum As String
    gDbTrans.SqlStmt = "Select CustomerId,AccNum From FDMaster where " & _
                    " AccId = " & m_AccID
    If gDbTrans.Fetch(rst, adOpenStatic) <> 1 Then Exit Function
    CustomerID = Val(FormatField(rst(0)))
    AccNum = FormatField(rst(1))
    
    'Get the Deposit Number
    Dim AccId As Long
    'if Account deposit is there then DEposit ID is 2, else it is 1
    gDbTrans.SqlStmt = "Select Max(AccID) from FDMaster"
    If gDbTrans.Fetch(rst, adOpenStatic) <= 0 Then
        AccId = 1
    Else
        AccId = FormatField(rst(0)) + 1
    End If

    'Check the Amount to be withdrawn From MAture Fd account
    Dim Amount As Currency
    If Not CurrencyValidate(txtOldMatAmount, False) Then Exit Function
    Amount = Val(txtOldMatAmount.Caption)
    
    'Now with draw this amount from The MFD Account
    Dim transType As wisTransactionTypes
    Dim TransID As Long
    
    gDbTrans.SqlStmt = "SELECT MAX(Transid) From MatFDTrans WHERE " & _
        " ACCID = " & m_AccID
    
    Call gDbTrans.Fetch(rst, adOpenStatic)
    TransID = FormatField(rst(0)) + 1
    
    transType = wContraWithdraw
    gDbTrans.SqlStmt = "INSERT INTO MatFDTrans (AccID, TransID, " & _
            " TransDate, Amount,TransType,Balance,Particulars,UserID) VALUES " & _
            "(" & m_AccID & "," & _
            TransID & "," & _
            "#" & TransDate & "#, " & _
            Amount & "," & _
            transType & "," & _
            " 0, 'Transferred to FD', " & _
            gUserID & " )"
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        'MsgBox "Unable to do this operation"
        MsgBox GetResourceString(1), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If
    
'Now TransFer the Details to the FD account
    gDbTrans.SqlStmt = "Insert into FDMaster (AccID, AccNum, " & _
            " CustomerID, CreateDate, EffectiveDate,DepositAmount," & _
            " MaturityDate,MaturityAmount,RateOfInterest,CertificateNo,UserID) values ( " & _
            AccId & "," & _
            AddQuotes(AccNum, True) & "," & _
            CustomerID & ", " & _
            "#" & TransDate & "#," & _
            "#" & EffectiveDate & "#," & _
            DepositAmount & "," & _
            "#" & MatureDate & "#," & _
            txtMatureAmount & ", " & _
            Interest & "," & _
            AddQuotes(txtCertificate.Text, True) & _
            "," & gUserID & " )"
    
    'Fire the query
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    'Now insert the Depoist Amount & other details in to the FDTrans Table
    
   'Get the new Transaction ID
   'This is a new deposit, and will have a transaction ID of 1
    TransID = 1
    transType = wContraDeposit
    gDbTrans.SqlStmt = "Insert into FDTrans (AccID,TransID, TransType, " & _
            " TransDate, Amount, Balance, " & _
            " Particulars,UserId) values ( " & _
            AccId & "," & _
            TransID & "," & _
            transType & "," & _
            "#" & TransDate & "#," & _
            DepositAmount & "," & DepositAmount & "," & _
            "'From MatFD of " & m_AccID & "'" & _
            "," & gUserID & " )"
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    
    gDbTrans.CommitTrans

    RenewMFD = True
    
End Function

Private Sub UpdateDetails()
Dim transType As wisTransactionTypes
Dim rst As Recordset

If Not DateValidate(txtTransDate, "/", True) Then txtTransDate = gStrDate

If m_AccID = 0 Then
    cmdRenew.Enabled = False
    cmdClose.Enabled = False
    Exit Sub
End If

'First Acc holder details
gDbTrans.SqlStmt = "SELECT * FROM FDMaster WHERE " & _
    "AccID = " & m_AccID
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub
'Now Load the acc name
Dim L_Custreg As clsCustReg
Set L_Custreg = New clsCustReg
LblAccHolder.Caption = L_Custreg.CustomerName(FormatField(rst("customerID")))
txtAccNo.Caption = FormatField(rst("AccNum"))

'Get the Certificate no & othere details
transType = wDeposit
gDbTrans.SqlStmt = "Select * from FDMaster where AccID = " & m_AccID
If gDbTrans.Fetch(rst, adOpenStatic) <= 0 Then Exit Sub

txtOldCertificate = FormatField(rst("CertificateNo"))
txtOldDepAmount = FormatField(rst("DepositAmount"))
txtOldDepdate = FormatField(rst("EffectiveDate"))
    
Set L_Custreg = Nothing
'Get the Details From the MFD trans
gDbTrans.SqlStmt = "SELECT * FROM MatFDTrans WHERE " & _
    "AccID = " & m_AccID & " ORDER BY TransID Desc "
     
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub

If rst("Balance") = 0 Then
    'Deposited Renewed
    cmdRenew.Enabled = False
    cmdClose.Enabled = False
    rst.MoveNext
End If

txtOldMatDate = FormatField(rst("TransDate"))
txtRenewDate = FormatField(rst("TransDate"))
txtOldMatAmount = FormatField(rst("Balance"))
txtDepositAmount = FormatField(rst("Balance"))


End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
If Not CloseMFD Then Exit Sub
Unload Me
End Sub

Private Sub cmdDepositDate_Click()
With Calendar
    .Left = Me.Left + frarenew.Left + cmdDepositDate.Left
    .Top = Me.Top + frarenew.Top + cmdDepositDate.Top - .Height / 2
    .selDate = txtRenewDate
    .Show 1
    txtRenewDate = .selDate
End With

End Sub

Private Sub cmdMatureDate_Click()
With Calendar
    .Left = Me.Left + frarenew.Left + cmdMatureDate.Left
    .Top = Me.Top + frarenew.Top + cmdMatureDate.Top - .Height / 2
    .selDate = txtMatureDate
    .Show 1
    txtMatureDate.Text = .selDate
End With

End Sub


Private Sub cmdRenew_Click()
If Not RenewMFD Then Exit Sub
Unload Me
End Sub

Private Sub cmdTransDate_Click()
With Calendar
    .Left = Me.Left + fraAccount.Left + cmdTransDate.Left
    .Top = Me.Top + fraAccount.Top + cmdTransDate.Top - .Height / 2
    .selDate = txtTransDate
    .Show 1
    txtTransDate = .selDate
End With
End Sub

Private Sub Form_Load()
'Center the form
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

' Set Kannada Fonts
Call SetKannadaCaption

Call UpdateDetails
If gOnLine Then
    txtTransDate.Locked = True
    cmdTransDate.Enabled = False
End If

End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

fraAccount.Caption = GetResourceString(70) & " " & _
        GetResourceString(295) 'Account Details
fraOldFD.Caption = GetResourceString(43, 295) 'Mature detaisl
frarenew.Caption = GetResourceString(46, 295) 'Mature details

lblTransDate = GetResourceString(37)
lblAccNO.Caption = GetResourceString(36, 60)

lblOldCertificate = GetResourceString(337, 60) 'certificate
lblOldDepAmount = GetResourceString(43, 42) 'Deposited Amount
lblOldMatAmount = GetResourceString(46, 40) 'MAtured amount
lblOldDepDate = GetResourceString(43, 37) 'Deposit date
lblOldMatDate = GetResourceString(48, 37)  'MAturity date

lblCertificate = GetResourceString(337, 60) 'Certificate
lblrenewDate.Caption = GetResourceString(265, 37) 'Renew date
lblMatureDate = GetResourceString(48, 67) 'Mature date
lblDays.Caption = GetResourceString(44) & GetResourceString(92)  'days
lblInterest = GetResourceString(186)
lblDepAmount = GetResourceString(43, 42) 'Deposit amount
lblMatureAmount = GetResourceString(48, 40) 'Mature Amount

'Set the Kannada caption to the Command buttons
cmdClose.Caption = GetResourceString(11) 'Close
cmdCancel.Caption = GetResourceString(2) 'Cancel
cmdRenew.Caption = GetResourceString(265) '"Renew"




End Sub






Private Sub txtDays_Change()
On Error GoTo ExitLine

If Me.ActiveControl.name <> txtDays.name Then Exit Sub
Dim Days As Long

If Val(txtDays.Text) > 99999 Then Exit Sub
If Val(txtDays.Text) <= 0 Or Not IsNumeric(txtDays.Text) Then
    txtInterest.Text = "0.00"
    txtMatureDate.Text = txtRenewDate.Text
    Exit Sub
Else
    Days = Val(txtDays.Text)
End If


If Not DateValidate(txtRenewDate, "/", True) Then Exit Sub

If Trim(txtDays.Text) = "" Then
    txtMatureDate.Text = txtRenewDate.Text
    Exit Sub
End If

Dim DateDep As String
Dim DateMature As Date
Dim SchemeName As String

DateDep = GetSysFormatDate(txtRenewDate.Text)  'get the american date
DateMature = DateAdd("d", Days, DateDep)
txtMatureDate.Text = GetIndianDate(DateMature)

If Days > 0 And Days <= 30 Then
    'txtInterest.Text = M_setUp.ReadSetupValue("FDAcc", "1/2_1_Deposit", Val(txtInterestRates(0).Text))
    SchemeName = "0.5_1_Deposit"
    
ElseIf Days > 30 And Days <= 45 Then
    SchemeName = "1_1.5_Deposit"
ElseIf Days > 45 And Days <= 90 Then
    SchemeName = "1.5_3_Deposit"
ElseIf Days > 90 And Days <= 180 Then
    SchemeName = "3_6_Deposit"
ElseIf Days > 180 And Days <= 365 Then
    SchemeName = "6_12_Deposit"
ElseIf Days > 365 And Days <= 730 Then
    SchemeName = "12_24_Deposit"
ElseIf Days > 730 And Days <= 1095 Then
    SchemeName = "24_36_Deposit"
ElseIf Days > 1095 Then
    SchemeName = "36_Above_Deposit"
End If

Dim ClsInt As New clsInterest
Me.txtInterest.Text = ClsInt.InterestRate(wis_Deposits, SchemeName, txtRenewDate.Text, txtMatureDate.Text)

txtMatureAmount = txtDepositAmount + _
            CCur(ComputeFDInterest(txtDepositAmount, txtRenewDate.Text, txtMatureDate.Text, _
            CSng(txtInterest.Text), False))

ExitLine:
End Sub


Private Sub txtMatureDate_Change()

If Me.ActiveControl.name <> txtMatureDate.name Or _
        ActiveControl.name <> cmdMatureDate.name Then Exit Sub


On Error GoTo ExitLine

Dim Days As Long

If Val(txtDays.Text) > 99999 Then Exit Sub
If Val(txtDays.Text) <= 0 Or Not IsNumeric(txtDays.Text) Then
    txtInterest.Text = "0.00"
    txtMatureDate.Text = txtRenewDate.Text
    Exit Sub
Else
    Days = Val(txtDays.Text)
End If

If Not DateValidate(txtRenewDate, "/", True) Then Exit Sub

If Trim(txtDays.Text) = "" Then
    txtMatureDate.Text = txtRenewDate.Text
    Exit Sub
End If

Dim DateDep As String
Dim DateMature As String
Dim SchemeName As String

DateDep = GetSysFormatDate(txtRenewDate.Text)  'get the american date
DateMature = CStr(DateAdd("d", Days, CDate(DateDep)))
txtMatureDate.Text = GetIndianDate(CDate(DateMature))

If Days > 0 And Days <= 30 Then
    'txtInterest.Text = M_setUp.ReadSetupValue("FDAcc", "1/2_1_Deposit", Val(txtInterestRates(0).Text))
    SchemeName = "0.5_1_Deposit"
ElseIf Days > 30 And Days <= 45 Then
    SchemeName = "1_1.5_Deposit"
ElseIf Days > 45 And Days <= 90 Then
    SchemeName = "1.5_3_Deposit"
ElseIf Days > 90 And Days <= 180 Then
    SchemeName = "3_6_Deposit"
ElseIf Days > 180 And Days <= 365 Then
    SchemeName = "6_12_Deposit"
ElseIf Days > 365 And Days <= 730 Then
    SchemeName = "12_24_Deposit"
ElseIf Days > 730 And Days <= 1095 Then
    SchemeName = "24_36_Deposit"
ElseIf Days > 1095 Then
    SchemeName = "36_Above_Deposit"
End If

Dim ClsInt As New clsInterest

txtInterest.Text = ClsInt.InterestRate(wis_Deposits, SchemeName, txtRenewDate.Text, txtMatureDate.Text)
txtMatureAmount = txtDepositAmount + _
        CCur(ComputeFDInterest(txtDepositAmount, txtRenewDate.Text, _
        txtMatureDate.Text, CSng(txtInterest.Text), False))

ExitLine:

End Sub


