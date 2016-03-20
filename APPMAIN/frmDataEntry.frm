VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDataEntry 
   Caption         =   "DataEntry"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLabel 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9315
      TabIndex        =   15
      Top             =   0
      Width           =   9375
      Begin VB.TextBox txtShareValue 
         Height          =   300
         Left            =   7440
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMemFee 
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Text            =   "31/3/2014"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblShareValue 
         Caption         =   "Share Face Value"
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "member Fee"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picOut 
      Height          =   5625
      Left            =   600
      ScaleHeight     =   5565
      ScaleWidth      =   6225
      TabIndex        =   4
      Top             =   600
      Width           =   6285
      Begin VB.ComboBox cmbGender 
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.ComboBox cmbPlace 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cmbCaste 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtBox 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   0
         Width           =   1035
      End
      Begin VB.TextBox txtBefore 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtAfter 
         Height          =   285
         Left            =   1350
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         Top             =   420
         Width           =   435
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   5145
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   9075
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.ComboBox cmbTitle 
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   345
      Left            =   5700
      TabIndex        =   6
      Top             =   6450
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   4350
      TabIndex        =   5
      Top             =   6450
      Width           =   945
   End
   Begin VB.TextBox txtTemp 
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5760
      Width           =   150
   End
   Begin VB.Label lblTitle 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblHead 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Shown As Boolean
Private m_Module As wisModules
Private cmb As ComboBox
Private m_CustReg As New clsCustReg
Private custReg(100) As New clsCustReg

Private m_colNo, m_rowNo As Integer
Private m_TotalCols As Integer
Private m_Fixedcols As Integer

Public Event Initialised(Min As Long, Max As Long)
Public Event Processed(Ratio As Single)

Public Property Let AccNum(Index As Integer, NewValue As String)

With grd
    '.Col = 0
    '.Row = Index
    '.Text = Newvalue
    .TextMatrix(Index, 0) = NewValue
    RaiseEvent Processed(.Row / .Rows)
End With


End Property

Private Sub InitGrid()

Dim Wid As Single
With grd
    .ZOrder 0
    .Top = 0: .Left = 0
    .Width = picOut.Width - 50
    .Height = picOut.Height - 50
    .AllowUserResizing = flexResizeNone
    
    .FixedCols = 1: .FixedRows = 1
    .ScrollBars = flexScrollBarVertical
    
    'If m_Module = wis_Members Then
     Wid = (.Width - cmbTitle.Width - cmbCaste.Width - cmbPlace.Width - cmd.Width - cmbGender.Width) / (.Cols - 5)
     .ColWidth(0) = Wid * 0.3
     .ColWidth(1) = Wid * 0.5
     .ColWidth(2) = cmd.Width
     .ColWidth(3) = cmbTitle.Width 'Wid * 1.5
     .ColWidth(4) = Wid * 1.5
     .ColWidth(5) = Wid
     .ColWidth(6) = Wid * 1.5
     .ColWidth(7) = cmbGender.Width 'Wid * 0.5
     .ColWidth(8) = cmbCaste.Width 'Wid * 0.5
     .ColWidth(9) = cmbPlace.Width 'Wid * 1.5
     If m_Module <> wis_Members Then .ColWidth(10) = Wid * 0.9
     'If m_Module = wis_Members Then .ColWidth(11) = Wid
    'End If
End With
End Sub








Public Property Let Balance(Index As Integer, NewValue As Currency)

With grd
    '.Col = 2
    '.Row = Index
    '.Text = FormatCurrency(Newvalue)
    .TextMatrix(Index, 2) = NewValue
End With

End Property
Private Function SaveMembers() As Boolean
    SaveMembers = False
    m_Shown = False
    Dim count As Integer
    Dim CustomerID As Integer
    Dim AccId  As Long
    Dim AccNum  As String
    Dim headID As Long
    Dim TransDate As Date
    Dim MemFee As Currency
    Dim Balance As Currency
    Dim ShareFaceValue As Currency
    
    Dim LeavesCount As Integer
    Dim LeavesCount_2 As Integer
    Dim transLoop As Integer
    Dim transType As wisTransactionTypes
    Dim bankClass As New clsBankAcc
    Dim CreateDate As Date
    
Dim rst As Recordset

    If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
        'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    Else
        TransDate = GetSysFormatDate(txtDate.Text)
    End If
    
    If Len(Trim$(txtMemFee.Text)) > 0 And Not IsNumeric(Trim$(txtMemFee.Text)) Then
        If MsgBox("Member Fee not specified! Do you want to continue", vbYesNo + vbQuestion _
                + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then GoTo Exit_Line
        ActivateTextBox txtMemFee
        Exit Function
    End If
    MemFee = Val(txtMemFee.Text)
    If Len(Trim$(txtShareValue.Text)) > 0 And Not IsNumeric(Trim$(txtShareValue.Text)) Then
        If MsgBox("Share face value not specified! Do you want to continue", vbYesNo + vbQuestion _
                + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then GoTo Exit_Line
        ActivateTextBox txtMemFee
        Exit Function
    End If
    ShareFaceValue = Val(txtShareValue.Text)
    
    gDbTrans.SqlStmt = "Select Max(AccID) from MemMaster"
    If gDbTrans.Fetch(rst, adOpenDynamic) Then AccId = FormatField(rst(0))
        
    gDbTrans.BeginTrans
    headID = bankClass.GetHeadIDCreated(GetResourceString(79, 191), _
                LoadResString(79) & " " & LoadResString(191), parBankIncome, 0, wis_Members)
    gDbTrans.CommitTrans
    
    For count = 1 To grd.Rows - 1
        LeavesCount = 0
        LeavesCount_2 = 0
        AccNum = Trim$(grd.TextMatrix(count, 1))
        If Len(AccNum) < 1 Then GoTo NextCount
        
        
        If grd.RowData(count) < 1 Then
            ''Save New Customer
            If custReg(count) Is Nothing Then Set custReg(count) = New clsCustReg: custReg(count).NewCustomer
            If custReg(count).FormValue Is Nothing Then custReg(count).NewCustomer: custReg(count).Modified = True
            If Not custReg(count) Is Nothing Then
                custReg(count).FormValue.cmbTitle.ListIndex = GetComboIndex(cmbTitle, grd.TextMatrix(count, 3))
                custReg(count).FormValue.txtFirstName = grd.TextMatrix(count, 4)
                custReg(count).FormValue.txtMiddleName = grd.TextMatrix(count, 5)
                custReg(count).FormValue.txtLastName = grd.TextMatrix(count, 6)
                custReg(count).FormValue.cmbGender.ListIndex = GetComboIndex(cmbGender, grd.TextMatrix(count, 7))
                custReg(count).FormValue.txtCaste = grd.TextMatrix(count, 8)
                custReg(count).FormValue.txtHomeCity = grd.TextMatrix(count, 9)
                
                custReg(count).SaveCustomer
                CustomerID = custReg(count).CustomerID
            End If
        Else
            CustomerID = grd.RowData(count)
        End If
        
        ''Now Save the Member Details
        gDbTrans.SqlStmt = "Select * FROM MemMAster Where AccNum = " & AddQuotes(AccNum)
        
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            '"This Account number already exists"
            MsgBox "Account number " & AccNum & " already exists", vbInformation, wis_MESSAGE_TITLE
            GoTo NextCount
        End If
        
        'Check Whethe This Customer already becomethe member
        gDbTrans.SqlStmt = "Select * from MemMaster Where CustomerID = " & CustomerID
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            If MsgBox("This Customer already has Member Number" & _
                FormatField(rst("AccNum")) & vbCrLf & _
                "Do you want to continue!", vbYesNo + vbQuestion _
                + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then GoTo NextCount
        End If
        
        Balance = Val(grd.TextMatrix(count, 10))
        CreateDate = TransDate
        If DateValidate(grd.TextMatrix(count, 11), "/", True) Then CreateDate = GetSysFormatDate(grd.TextMatrix(count, 11))
        gDbTrans.BeginTrans
        
        AccId = AccId + 1
        
        gDbTrans.SqlStmt = "Insert into MemMaster (AccNum,AccID, CustomerID, " & _
                " CreateDate,MemberType,AccGroupID, UserID" & _
                ") " & _
                "values (" & AddQuotes(AccNum, True) & "," & _
                AccId & "," & _
                CustomerID & "," & _
                "#" & CreateDate & "#,1, 1," & _
                gUserID & " )"
        
        If Not gDbTrans.SQLExecute Then
            AccId = AccId - 1
            gDbTrans.RollBack
            GoTo NextCount
        End If
        
        'Build Sql To Insert values into MemTrans
        transType = wDeposit
        gDbTrans.SqlStmt = "Insert Into MemIntTrans(AccId,TransId,TransDate," & _
                " Amount,TransType,Balance,Particulars,UserId)" & _
                " Values(" & _
                AccId & ", 1, " & _
                "#" & TransDate & "#" & _
                ", " & MemFee & ", " & _
                transType & ", 0,'Memberdhip Fee'," & gUserID & " ) "
        If Not gDbTrans.SQLExecute Then
            AccId = AccId - 1
            gDbTrans.RollBack
            GoTo NextCount
        End If
                ''update the Member fee
        If Not bankClass.UpdateCashDeposits(headID, MemFee, TransDate) Then
            AccId = AccId - 1
            gDbTrans.RollBack
            'Exit Function
        End If

        
        'Now Insert the Shares Transcation
        If ShareFaceValue > 0 And Balance > 0 Then
            LeavesCount = Balance / ShareFaceValue
            
            If Balance > (LeavesCount * ShareFaceValue) Then LeavesCount_2 = 1
            
            For transLoop = 1 To LeavesCount
                gDbTrans.SqlStmt = "Insert into ShareTrans (AccID, SaleTransID, " & _
                    " CertNo, FaceValue) values (" & _
                    AccId & ", " & _
                    1 & ", " & _
                    AddQuotes(AccNum + CStr(transLoop)) & ", " & _
                    ShareFaceValue & ")"
            
                If Not gDbTrans.SQLExecute Then
                    'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
                    MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
                    gDbTrans.RollBack
                    GoTo NextCount
                End If
            Next
            
            If LeavesCount_2 > 0 Then
                ShareFaceValue = Balance - (LeavesCount * ShareFaceValue)
                gDbTrans.SqlStmt = "Insert into ShareTrans (AccID, SaleTransID, " & _
                    " CertNo, FaceValue) values (" & _
                    AccId & ", " & _
                    1 & ", " & _
                    AddQuotes(AccNum & "_" & CStr(transLoop)) & ", " & _
                    ShareFaceValue & ")"
            
                If Not gDbTrans.SQLExecute Then
                    'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
                    MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
                    gDbTrans.RollBack
                    GoTo NextCount
                End If

            End If
            
            'Insert into MemTrans Tab
            gDbTrans.SqlStmt = "INSERT INTO MemTrans (AccID, TransID, TransDate, " & _
                        " Leaves, Amount, TransType, Balance,UserID) values ( " & _
                        AccId & ",1," & _
                        "#" & TransDate & "#, " & _
                        LeavesCount + LeavesCount_2 & ", " & _
                        Balance & ", " & _
                        transType & ", " & _
                        Balance & "," & gUserID & " )"
            If Not gDbTrans.SQLExecute Then
                'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
                gDbTrans.RollBack
                Exit Function
            End If
        
        End If
        
        gDbTrans.CommitTrans
    
NextCount:
    Next
    

SaveMembers = True

Exit_Line:

m_Shown = True
End Function

Private Function SaveSavings() As Boolean
    SaveSavings = False
    m_Shown = False
    Dim count As Integer
    Dim CustomerID As Integer
    Dim AccId  As Long
    Dim AccNum  As String
    Dim headID As Long
    Dim TransDate As Date
    Dim Balance As Currency
    Dim transType As wisTransactionTypes
    Dim bankClass As New clsBankAcc
    Dim rst As Recordset
    
    
    If Not DateValidate(Trim$(txtDate.Text), "/", True) Then
        'MsgBox "Invalid transaction date specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtDate
        Exit Function
    Else
        TransDate = GetSysFormatDate(txtDate.Text)
    End If
    
    gDbTrans.SqlStmt = "Select Max(AccID) from SBMaster"
    If gDbTrans.Fetch(rst, adOpenDynamic) Then AccId = FormatField(rst(0))
        
    gDbTrans.BeginTrans
        headID = bankClass.GetHeadIDCreated(GetResourceString(79, 191), _
                LoadResString(79) & " " & LoadResString(191), parBankIncome, 0, wis_Members)
    gDbTrans.CommitTrans
    transType = wDeposit

    For count = 1 To grd.Rows - 1
        AccNum = Trim$(grd.TextMatrix(count, 1))
        If Len(AccNum) < 1 Then GoTo NextCount
        
        
        If grd.RowData(count) < 1 Then
            ''Save New Customer
            If custReg(count) Is Nothing Then Set custReg(count) = New clsCustReg: custReg(count).NewCustomer
            If custReg(count).FormValue Is Nothing Then custReg(count).NewCustomer: custReg(count).Modified = True
            If Not custReg(count) Is Nothing Then
                custReg(count).FormValue.cmbTitle.ListIndex = GetComboIndex(cmbTitle, grd.TextMatrix(count, 3))
                custReg(count).FormValue.txtFirstName = grd.TextMatrix(count, 4)
                custReg(count).FormValue.txtMiddleName = grd.TextMatrix(count, 5)
                custReg(count).FormValue.txtLastName = grd.TextMatrix(count, 6)
                custReg(count).FormValue.cmbGender.ListIndex = GetComboIndex(cmbGender, grd.TextMatrix(count, 7))
                custReg(count).FormValue.txtCaste = grd.TextMatrix(count, 8)
                custReg(count).FormValue.txtHomeCity = grd.TextMatrix(count, 9)
                
                custReg(count).SaveCustomer
                CustomerID = custReg(count).CustomerID
            End If
        Else
            CustomerID = grd.RowData(count)
        End If
        
        ''Now Save the Member Details
        gDbTrans.SqlStmt = "Select * FROM SBMaster Where AccNum = " & AddQuotes(AccNum)
        
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            '"This Account number already exists"
            MsgBox "Account number " & AccNum & " already exists", vbInformation, wis_MESSAGE_TITLE
            GoTo NextCount
        End If
        
        'Check Whethe This Customer already becomethe member
        gDbTrans.SqlStmt = "Select * from SBMaster Where CustomerID = " & CustomerID
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            If MsgBox("This Customer already has Saving account" & _
                FormatField(rst("AccNum")) & vbCrLf & _
                "Do you want to continue!", vbYesNo + vbQuestion _
                + vbDefaultButton2, gAppName & " - Confirmation") = vbNo Then GoTo NextCount
        End If
        
        AccId = AccId + 1
        
        gDbTrans.SqlStmt = "Insert into SBMaster (AccNum,AccID, CustomerID, " & _
                " CreateDate,AccGroupID, UserID" & _
                ") " & _
                "values (" & AddQuotes(AccNum, True) & "," & _
                AccId & "," & _
                CustomerID & "," & _
                "#" & TransDate & "#, 1," & _
                gUserID & " )"
        
        gDbTrans.BeginTrans
        If Not gDbTrans.SQLExecute Then
            AccId = AccId - 1
            gDbTrans.RollBack
            GoTo NextCount
        End If
        
        
        Balance = Val(grd.TextMatrix(count, 10))
        If Balance > 0 Then
            'Build Sql To Insert values into SBTrans
            
            gDbTrans.SqlStmt = "Insert Into SBTrans(AccId,TransId,TransDate," & _
                    " Amount,TransType,Balance,Particulars,UserId)" & _
                    " Values(" & _
                    AccId & ", 1, " & _
                    "#" & TransDate & "#" & _
                    ", " & Balance & ", " & _
                    transType & "," & Balance & ",'Opening Balance'," & gUserID & " ) "
            If Not gDbTrans.SQLExecute Then
                AccId = AccId - 1
                gDbTrans.RollBack
                GoTo NextCount
            End If
        End If
        
        ''update the Member fee
        If Not bankClass.UpdateCashDeposits(headID, Balance, TransDate) Then
            AccId = AccId - 1
            gDbTrans.RollBack
            Exit Function
        End If
        
        gDbTrans.CommitTrans
    
NextCount:
    Next
    

SaveSavings = True

Exit_Line:

m_Shown = True
End Function

Public Sub ShowForm(moduleId As wisModules)

m_Module = moduleId

m_TotalCols = 12
If moduleId = wis_SBAcc Then m_TotalCols = 11

'Load the Column

    m_Fixedcols = 1
    'Call LoadContorls(100, 8)
    With grd
        .Row = 0
        
        .TextMatrix(0, 0) = GetResourceString(33)
        .TextMatrix(0, 1) = GetResourceString(49)
        .TextMatrix(0, 2) = ""  'Load Customer
        .TextMatrix(0, 3) = GetResourceString(119) 'Title
        .TextMatrix(0, 4) = GetResourceString(120) 'Fname
        .TextMatrix(0, 5) = GetResourceString(121)  'Mname
        .TextMatrix(0, 6) = GetResourceString(122) 'Lname
        .TextMatrix(0, 7) = GetResourceString(125) 'Gender
        .TextMatrix(0, 8) = GetResourceString(111) 'Caste
        .TextMatrix(0, 9) = GetResourceString(112) 'Place
        .TextMatrix(0, 10) = GetResourceString(42) 'Balance
        If .Cols > 11 Then .TextMatrix(0, 11) = GetResourceString(281) 'Balance
    End With
    
'End If

m_Shown = True
Me.Show 1

End Sub

Public Function LoadContorls(ControlLoopCount As Integer, ColCount As Integer)
'On Error Resume Next
If ControlLoopCount < 1 Then Exit Function
With grd
    .ZOrder 0
    .Top = 0: .Left = 0
    .Width = picOut.Width - 50
    .Height = picOut.Height - 50
    .AllowUserResizing = flexResizeNone
    .Cols = ColCount
    .Rows = ControlLoopCount + 1
    .FixedCols = 1: .FixedRows = 1
    Dim count As Integer
    .Col = 0
    For count = .FixedRows To .Rows - 1
        .Row = count
        .TextMatrix(count, 0) = count
    Next
End With

Call InitGrid

LoadContorls = True

RaiseEvent Initialised(0, ControlLoopCount + 1)

End Function




Private Sub cmbCaste_GotFocus()
    cmbCaste.TabIndex = 2
    txtAfter.TabIndex = 3
End Sub


Private Sub cmbGender_GotFocus()
    cmbGender.TabIndex = 2
    txtAfter.TabIndex = 3

End Sub


Private Sub cmbPlace_GotFocus()
    cmbPlace.TabIndex = 2
    txtAfter.TabIndex = 3
End Sub


Private Sub cmbTitle_GotFocus()
    cmbTitle.TabIndex = 2
    txtAfter.TabIndex = 3

End Sub




Private Sub cmd_Click()
    If custReg(m_rowNo) Is Nothing Then custReg(m_rowNo) = New clsCustReg
    
    custReg(m_rowNo).moduleId = wis_Members
    If grd.RowData(m_rowNo) > 0 Then
        'm_CustReg.CustomerID = grd.RowData(m_rowNo)
        custReg(m_rowNo).LoadCustomerInfo (grd.RowData(m_rowNo))
    Else
        If Not custReg(m_rowNo).CustomerLoaded Then custReg(m_rowNo).NewCustomer
    End If
    custReg(m_rowNo).ShowDialog
    'txtData(txtIndex).Text = m_CustReg.FullName
    With grd
        
        
        If Not custReg(m_rowNo).FormValue Is Nothing Then
            .TextMatrix(m_rowNo, 3) = custReg(m_rowNo).FormValue.cmbTitle.Text
            .TextMatrix(m_rowNo, 4) = custReg(m_rowNo).FormValue.txtFirstName
            .TextMatrix(m_rowNo, 5) = custReg(m_rowNo).FormValue.txtMiddleName
            .TextMatrix(m_rowNo, 6) = custReg(m_rowNo).FormValue.txtLastName
            .TextMatrix(m_rowNo, 7) = custReg(m_rowNo).FormValue.cmbGender
            .TextMatrix(m_rowNo, 8) = custReg(m_rowNo).FormValue.cmbCaste
            .TextMatrix(m_rowNo, 9) = custReg(m_rowNo).FormValue.cmbHomeCity
        End If
        If Not custReg(m_rowNo).IsNewCustomerLoaded Then
            grd.RowData(m_rowNo) = custReg(m_rowNo).CustomerID
            Set custReg(m_rowNo) = Nothing
        Else
            'If Not custReg(m_rowNo).CustomerLoaded Then Set custReg(m_rowNo) = Nothing
        End If
        
    End With
End Sub

Private Sub cmd_GotFocus()
    cmd.TabIndex = 2
    txtAfter.TabIndex = 3
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If m_Module = wis_Members Then
        If Not SaveMembers() Then Exit Sub
    ElseIf m_Module = wis_SBAcc Then
        If Not SaveSavings() Then Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()

Dim SetUp As New clsSetup
If SetUp Is Nothing Then Set SetUp = New clsSetup

Call SetKannadaCaption

'Load places for Places tab
    Call LoadPlaces(Me.cmbPlace)
''Load Caste From Table
    Call LoadCastes(cmbCaste)
With Me.cmbTitle
    .Clear
    .AddItem GetResourceString(385)
    .ItemData(.newIndex) = wisMale
    .AddItem GetResourceString(386)
    .ItemData(.newIndex) = wisFemale
    .AddItem GetResourceString(237)
    .ItemData(.newIndex) = wisNoGender
End With
With cmbGender
    .Clear
    .AddItem GetResourceString(385)
    .ItemData(.newIndex) = wisMale
    .AddItem GetResourceString(386)
    .ItemData(.newIndex) = wisFemale
    .AddItem GetResourceString(237)
    .ItemData(.newIndex) = wisNoGender
End With
With cmbTitle
    .Clear
    .AddItem GetResourceString(321)
    .AddItem GetResourceString(322)
    .AddItem GetResourceString(323)
    .AddItem GetResourceString(324)
    .AddItem GetResourceString(325)
    .AddItem GetResourceString(326)
    .AddItem "M/S"
End With

Dim Ctrl As Control
On Error Resume Next
    Dim rst As Recordset
    
    If m_Module = wis_Members Then
        txtMemFee.Text = SetUp.ReadSetupValue("MMAcc", "MemberShipFee", "0.00")
        gDbTrans.SqlStmt = "Select top 1 AccNum from MemMaster order by AccID Desc"
        txtMemFee.Visible = True
        Label1.Visible = True
        txtShareValue.Visible = True
        lblShareValue.Visible = True
    
    ElseIf m_Module = wis_SBAcc Then
        gDbTrans.SqlStmt = "Select top 1 AccNum from SBMaster order by AccID Desc"
        
    End If
    
    'Get the Member/Account No
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
         txtBox.Text = Val(FormatField(rst(0))) + 1
         grd.TextMatrix(1, 1) = txtBox.Text
    End If
    
    Call LoadContorls(100, m_TotalCols)
    
Me.Width = Screen.Width * 3 / 4
Me.Height = Screen.Height * 3 / 4
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Err.Clear

End Sub

Private Sub Form_Resize()

With picOut
    .Width = Me.Width - 3 * .Left
    .Height = Me.Height - (cmdSave.Height * 2 + lblTitle.Height + 400 + picLabel.Height)
End With
With cmdClose
    .Left = Me.Width - .Width * 1.5
    .Top = picOut.Top + picOut.Height + 100
End With
With cmdSave
    .Left = cmdClose.Left - .Width - 100
    .Top = cmdClose.Top
End With

lblHead.Top = cmdSave.Top
'cmbHead.Top = cmdSave.Top


With grd
    .ZOrder 0
    .Top = 0: .Left = 0
    .Width = picOut.Width - 50
    .Height = picOut.Height - 50
    '.AllowUserResizing = flexResizeNone
    Call InitGrid
  
End With

End Sub

Private Sub grd_EnterCell()
If Not m_Shown Then Exit Sub

Dim curText As String
Dim count As Integer
With grd
    m_colNo = .Col
    m_rowNo = .Row
    curText = .Text
    If .Row = 0 Then Exit Sub
    'If .Col <> 3 Then Exit Sub
    If .Col < m_Fixedcols Then Exit Sub
    If .Col = 2 Then
        cmd.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
        'cmb.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        cmd.ZOrder 0
        cmd.Visible = True
        cmd.SetFocus
        
    ElseIf .Col = 3 Or .Col = 7 Or .Col = 8 Or .Col = 9 Then
    'ElseIf .Col = 2 Or .Col = 6 Or .Col = 7 Then
        Set cmb = Nothing
        If .Col = 3 Then Set cmb = cmbTitle
        If .Col = 7 Then Set cmb = cmbGender
        If .Col = 8 Then Set cmb = cmbCaste
        If .Col = 9 Then Set cmb = cmbPlace
        
        'Check for the Value
        For count = 0 To cmb.ListCount - 1
            If cmb.List(count) = curText Then
                cmb.ListIndex = count
                Exit For
            End If
        Next
        cmb.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
        'cmb.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        cmb.Visible = True
        
        On Error Resume Next
        ActivateTextBox cmb
        cmb.ZOrder 0
        
        Err.Clear
    Else
        txtBox.Text = .Text
        
        'txtBox.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight
        txtBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        txtBox.Visible = True
        
        On Error Resume Next
        ActivateTextBox txtBox
        txtBox.ZOrder 0
        
        Err.Clear
    End If
End With

End Sub

Private Sub grd_GotFocus()
    Call grd_EnterCell
End Sub

Private Sub grd_LeaveCell()

If Not m_Shown Then Exit Sub

Static InLoop As Boolean
If InLoop Then Exit Sub
Dim Bal As Currency
Dim Amount As Double
Dim PrevAmount As Double

With grd
    'If .Col <> 3 Then Exit Sub
    'If .Col < m_Fixedcols Then Exit Sub
    
    'mARK AS CONTROL IS IN LOOP
    InLoop = True
    If .Col = 2 Then
        cmd.Visible = False
    ElseIf .Col = 3 Or .Col = 7 Or .Col = 8 Or .Col = 9 Then
        .Text = cmb.Text
        cmb.ListIndex = -1
        cmb.Visible = False
        Set cmb = Nothing
        
        Err.Clear
    Else
        'Now update the New Amount to the grid
        .Text = txtBox
        txtBox.Visible = False
        txtBox = ""
    End If
    
End With

'MARK AS CONTROL IS OUT OF LOOP
InLoop = False

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdClose.Caption = GetResourceString(11)
cmdSave.Caption = GetResourceString(7)

End Sub




Private Sub grd_Scroll()
txtBox.Visible = True
txtBox.Move grd.Left + grd.ColPos(grd.Col), grd.Top + grd.RowPos(grd.Row) ', grd.CellWidth, grd.CellHeight
'txtbox.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight
If txtBox.Left < grd.Left + grd.ColPos(0) = txtBox.Width Then txtBox.Visible = False
If txtBox.Top < grd.Top + grd.RowPos(0) + txtBox.Height Then txtBox.Visible = False

End Sub

Private Sub txtAfter_GotFocus()

Call txtBefore_GotFocus
Exit Sub

'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the next amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)
'If txtNo < txtAmount.Count - 1 Then
If txtNo < grd.Rows - 1 Then
    txtNo = txtNo + 1
    'Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub

Private Sub txtBefore_GotFocus()
Dim rowno As Integer
Dim colno As Integer
colno = m_colNo
rowno = m_rowNo

If colno = grd.Cols - 1 Then
    If rowno = grd.Rows - 1 Then
        Exit Sub
    Else
        colno = grd.FixedCols
        rowno = rowno + 1
        grd.Col = colno
        grd.Row = rowno
    End If
Else
    colno = colno + 1
    grd.Col = colno
End If

Exit Sub

'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the previous amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)

If txtNo > 0 Then
    txtNo = txtNo - 1
'    Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub

Private Sub txtBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode & " DOWN " & Shift
If KeyCode <> vbKeyTab Then Exit Sub

With grd
    If Shift = 1 Then
        If .Row = 1 Then Exit Sub
        .Row = .Row - 1
    Else
        If .Rows = .Rows - 1 Then Exit Sub
        .Row = .Row + 1
    End If
End With

End Sub

Private Sub txtBox_GotFocus()
    txtBox.TabIndex = 2 'txtAfter.TabIndex
    txtAfter.TabIndex = 3
End Sub

Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 37 Then 'Press Left Arrow
    If txtBox.SelStart = 0 And grd.Col > grd.FixedCols Then grd.Col = grd.Col - 1
    
ElseIf KeyCode = 39 Then 'Press Right Arrow
    If txtBox.SelStart = Len(txtBox.Text) And grd.Col < grd.Cols - 1 Then grd.Col = grd.Col + 1
    
ElseIf KeyCode = 38 Then  'Press UpArrow
    If grd.Row > grd.FixedRows Then grd.Row = grd.Row - 1
    
ElseIf KeyCode = 40 Then ' Press Down Arroow
    If grd.Row < grd.Rows - 1 Then grd.Row = grd.Row + 1
    
ElseIf KeyCode = 33 Then  'Press PageUp
    If grd.Row > 0 Then grd.Row = grd.Row - 1
    
ElseIf KeyCode = 34 Then ' Press PageDown
    If grd.Row < grd.Rows - 1 Then grd.Row = grd.Row + 1
    
End If

End Sub



