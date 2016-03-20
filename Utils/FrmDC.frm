VERSION 5.00
Begin VB.Form frmDC 
   Caption         =   "Material/Pigmy Debit/Credit Entry"
   ClientHeight    =   2895
   ClientLeft      =   3645
   ClientTop       =   2955
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6960
   Begin VB.Frame fraDC 
      Caption         =   "Debit/Credit"
      Height          =   2145
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   6705
      Begin VB.TextBox txtCLStock 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Width           =   1665
      End
      Begin VB.ComboBox cmbAccNames 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1665
      End
      Begin VB.TextBox txtDebit 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   930
         Width           =   1665
      End
      Begin VB.TextBox txtCredit 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1590
         Width           =   1665
      End
      Begin VB.Label lblCLStock 
         Caption         =   "Closing Stock :"
         Height          =   615
         Left            =   3690
         TabIndex        =   14
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblSelection 
         Caption         =   "Select an Account :"
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lblDate 
         Caption         =   "Date :"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lblDebit 
         Caption         =   "Debit :"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lblCredit 
         Caption         =   "Credit :"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   1290
         Width           =   1485
      End
      Begin VB.Label lblStock 
         Caption         =   "Stock :"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   1620
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdCalcel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   345
      Left            =   5790
      TabIndex        =   7
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4710
      TabIndex        =   6
      Top             =   2370
      Width           =   1005
   End
End
Attribute VB_Name = "frmDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccType As wisAccHeads



Private Sub cmbAccNames_Click()
If cmbAccNames.ListIndex = -1 Then Exit Sub
If cmbAccNames.ItemData(cmbAccNames.ListIndex) <> wis_Stock Then
    lblStock.Caption = LoadResString(gLangOffSet + 47)
'''    txtProfit.Enabled = False
    txtCLStock.Enabled = False
Else
    lblStock.Caption = LoadResString(gLangOffSet + 373)
'''    txtProfit.Enabled = True
    txtCLStock.Enabled = True
End If
cmbAccNames_LostFocus
txtDate_Change
End Sub

Private Sub cmbAccNames_LostFocus()
If cmbAccNames.ListIndex = -1 Then Exit Sub
AccType = cmbAccNames.ItemData(cmbAccNames.ListIndex)
End Sub


Private Sub cmdCalcel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Transdate As String
Dim TransType As wisTransactionTypes
Dim DBOperation As Integer
Dim AccountType As Integer

If cmbAccNames.ListIndex = -1 Then Exit Sub
If Not IsNumeric(txtDebit.Text) Or Not IsNumeric(txtCredit.Text) _
   Or Not IsNumeric(txtStock.Text) Then
   MsgBox LoadResString(gLangOffSet + 499), vbInformation, "Not A Numeric"
   Exit Sub
End If

AccountType = cmbAccNames.ItemData(cmbAccNames.ListIndex)

Transdate = FormatDate(txtDate.Text)
gDBTrans.SQLStmt = "SELECT * FROM Material WHERE Module = " & AccountType & _
    " AND TransDate = #" & Transdate & "#"
If gDBTrans.SQLFetch <= 1 Then
   DBOperation = wis_INSERT
Else
   DBOperation = wis_UPDATE
End If

gDBTrans.BeginTrans
If DBOperation = wis_INSERT Then
   If Val(txtDebit.Text) >= 0 Then
        gDBTrans.SQLStmt = "INSERT INTO Material (Module, TransDate, Amount, " & _
           "TransType) VALUES (" & AccountType & ", #" & Transdate & "#, " & _
           Val(txtDebit.Text) & ", " & wDeposit & ")"
            If Not gDBTrans.SQLExecute Then
               gDBTrans.RollBack
               Exit Sub
            End If
   End If
   If Val(txtCredit.Text) >= 0 Then
        gDBTrans.SQLStmt = "INSERT INTO Material (Module, TransDate, Amount, " & _
           "TransType) VALUES (" & AccountType & ", #" & Transdate & "#, " & _
           Val(txtCredit.Text) & ", " & wWithDraw & ")"
        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           Exit Sub
        End If
   End If
   If Val(txtStock.Text) >= 0 Then
        gDBTrans.SQLStmt = "INSERT INTO Material (Module, TransDate, Amount, " & _
           "TransType) VALUES (" & AccountType & ", #" & Transdate & "#, " & _
           Val(txtStock.Text) & ", " & IIf((AccountType = wis_Stock), wStock, _
           IIf((AccountType = wis_PD), wInterest, wCharges)) & ")"
        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           Exit Sub
        End If
   End If
   If AccountType = wis_Stock Then
        If Val(txtCLStock.Text) >= 0 Then
        gDBTrans.SQLStmt = "INSERT INTO Material (Module, TransDate, Amount, " & _
           "TransType) VALUES (" & AccountType & ", #" & Transdate & "#, " & _
           Val(txtCLStock.Text) & ", " & wCharges & ")"
        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           Exit Sub
        End If
        End If
   End If
Else
   If Val(txtDebit.Text) >= 0 Then
    gDBTrans.SQLStmt = "UPDATE Material SET Amount = " & Val(txtDebit.Text) & _
       " WHERE Module = " & AccountType & " AND TransDate = #" & Transdate & "#" & _
       " AND TransType = " & wDeposit
    If Not gDBTrans.SQLExecute Then
       gDBTrans.RollBack
       Exit Sub
    End If
   End If
   If Val(txtCredit.Text) >= 0 Then
    gDBTrans.SQLStmt = "UPDATE Material SET Amount = " & Val(txtCredit.Text) & _
       " WHERE Module = " & AccountType & " AND TransDate = #" & Transdate & "#" & _
       " AND TransType = " & wWithDraw
    If Not gDBTrans.SQLExecute Then
       gDBTrans.RollBack
       Exit Sub
    End If
   End If
   If Val(txtStock.Text) >= 0 Then
    gDBTrans.SQLStmt = "UPDATE Material SET Amount = " & Val(txtStock.Text) & _
       " WHERE Module = " & AccountType & " AND TransDate = #" & Transdate & "#" & _
       " AND TransType = " & IIf((AccountType = wis_Stock), wStock, _
       IIf((AccountType = wis_PD), wInterest, wCharges))
    If Not gDBTrans.SQLExecute Then
       gDBTrans.RollBack
       Exit Sub
    End If
   End If
   If AccountType = wis_Stock Then
        If Val(txtCLStock.Text) >= 0 Then
            gDBTrans.SQLStmt = "UPDATE Material SET Amount = " & Val(txtCLStock.Text) & _
               " WHERE Module = " & AccountType & " AND TransDate = #" & Transdate & "#" & _
               " AND TransType = " & wCharges
            If Not gDBTrans.SQLExecute Then
               gDBTrans.RollBack
               Exit Sub
            End If
        End If
    End If
End If
gDBTrans.CommitTrans
MsgBox "Operation is Successful", vbInformation, "Debit/Credit into Material"
End Sub

Private Sub Command1_Click()
Dim Transdate As String
Dim Amount As Currency
Transdate = "4/1/01"
While CDate(Transdate) < Now
    gDBTrans.SQLStmt = "SELECT SUM(Amount) FROM PDTrans WHERE TransDate = #" & _
        Transdate & "# AND TransType = " & wWithDraw & " And Loan = False"
    gDBTrans.SQLFetch
    Amount = FormatField(gDBTrans.Rst(0))
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "INSERT INTO Material (Module, TransDate, Amount, " & _
       "TransType) VALUES (" & wis_PD & ", #" & Transdate & "#, " & _
       Amount & ", " & wWithDraw & ")"
        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           Exit Sub
        End If
    gDBTrans.CommitTrans
    Transdate = DateAdd("d", 1, Transdate)
Wend
End Sub


Private Sub Form_Load()
SetKannadaCaption
cmbAccNames.AddItem LoadResString(gLangOffSet + 401) & " / " & LoadResString(gLangOffSet + 402)
cmbAccNames.ItemData(cmbAccNames.NewIndex) = wis_Stock
cmbAccNames.AddItem LoadResString(gLangOffSet + 425)
cmbAccNames.ItemData(cmbAccNames.NewIndex) = wis_PD
cmbAccNames.AddItem LoadResString(gLangOffSet + 417)
cmbAccNames.ItemData(cmbAccNames.NewIndex) = wis_PDLoan
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub SetKannadaCaption()
Dim Ctrl As Control
    For Each Ctrl In Me
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then
            Ctrl.Font.Size = gFontSize
        End If
    Next
    lblSelection.Caption = LoadResString(gLangOffSet + 34)
    lblDate.Caption = LoadResString(gLangOffSet + 37)
    lblDebit.Caption = LoadResString(gLangOffSet + 276)
    lblCredit.Caption = LoadResString(gLangOffSet + 277)
    lblStock.Caption = LoadResString(gLangOffSet + 373)
    fraDC.Caption = LoadResString(gLangOffSet + 276) & " / " & _
        LoadResString(gLangOffSet + 277)
    lblCLStock.Caption = LoadResString(gLangOffSet + 285) & _
        " " & LoadResString(gLangOffSet + 373)
End Sub

Private Sub txtCredit_LostFocus()
If Val(txtCredit.Text) > 0 And AccType = wis_Stock Then
    txtStock.Text = FormatCurrency(txtCredit.Text)
End If
End Sub

Private Sub txtDate_Change()
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
gDBTrans.SQLStmt = "SELECT Amount FROM Material WHERE TransDate = #" & _
   FormatDate(txtDate.Text) & "# AND Module = " & AccType & _
   " AND TransType = " & wDeposit
If gDBTrans.SQLFetch > 0 Then
    txtDebit.Text = FormatField(gDBTrans.Rst(0))
Else
    txtDebit.Text = FormatCurrency(0)
End If

gDBTrans.SQLStmt = "SELECT Amount FROM Material WHERE TransDate = #" & _
   FormatDate(txtDate.Text) & "# AND Module = " & AccType & _
   " AND TransType = " & wWithDraw
If gDBTrans.SQLFetch > 0 Then
    txtCredit.Text = FormatField(gDBTrans.Rst(0))
Else
    txtCredit.Text = FormatCurrency(0)
End If

gDBTrans.SQLStmt = "SELECT Amount FROM Material WHERE TransDate = #" & _
   FormatDate(txtDate.Text) & "# AND Module = " & AccType & _
   " AND TransType = " & IIf((AccType = wis_Stock), wStock, _
           IIf((AccType = wis_PD), wInterest, wCharges))
If gDBTrans.SQLFetch > 0 Then
    txtStock.Text = FormatField(gDBTrans.Rst(0))
Else
    txtStock.Text = FormatCurrency(0)
End If

End Sub

Private Sub txtDate_LostFocus()
If Not DateValidate(txtDate.Text, "/", True) Then
   MsgBox LoadResString(gLangOffSet + 501), vbCritical, "Date is not valid"
End If
End Sub


Private Sub txtStock_LostFocus()
Dim OPStock As Currency
If Not DateValidate(txtDate.Text, "/", True) Then Exit Sub
If AccType = wis_Stock Then
    gDBTrans.SQLStmt = "Select TOP 1 Amount From Material WHERE Module = " & AccType & _
        " And TransType = " & wCharges & " And TransDate < #" & FormatDate(txtDate.Text) & "#" & _
        " ORDER BY TransDate DESC"
    
    If gDBTrans.SQLFetch >= 1 Then
        OPStock = FormatField(gDBTrans.Rst(0))
    End If

    If OPStock <= 0 Then
        OPStock = OBOfAccount(wis_Stock, txtDate.Text)
    End If
    
    If IsNumeric(Val(txtStock.Text)) Then
        txtCLStock.Text = FormatCurrency(OPStock + Val(txtStock.Text))
    End If
End If
End Sub

