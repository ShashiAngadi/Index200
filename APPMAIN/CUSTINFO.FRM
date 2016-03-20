VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmCustInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX2000  -  Customer information"
   ClientHeight    =   4860
   ClientLeft      =   2040
   ClientTop       =   1920
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   315
      Left            =   -615
      TabIndex        =   2
      Top             =   735
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   5055
      TabIndex        =   4
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtCustName 
      Height          =   300
      Left            =   1650
      TabIndex        =   7
      Top             =   165
      Width           =   3885
   End
   Begin VB.CommandButton cmdLookUP 
      Caption         =   "..."
      Height          =   315
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   360
   End
   Begin VB.TextBox txtCustID 
      Height          =   330
      Left            =   1635
      TabIndex        =   1
      Top             =   -180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin ComctlLib.ListView lvw 
      Height          =   3075
      Index           =   2
      Left            =   420
      TabIndex        =   9
      Top             =   1065
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "CustInfo.frx":0000
      NumItems        =   0
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3720
      Left            =   330
      TabIndex        =   3
      Top             =   600
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   6562
      _Version        =   327680
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Accounts"
            Key             =   "Bal"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Loans"
            Key             =   "Loan"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      MouseIcon       =   "CustInfo.frx":001C
   End
   Begin ComctlLib.ListView lvw 
      Height          =   3075
      Index           =   1
      Left            =   435
      TabIndex        =   8
      Top             =   1065
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5424
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327680
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "CustInfo.frx":0038
      NumItems        =   0
   End
   Begin VB.Label lblCustomerName 
      Caption         =   "Customer Name :"
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   195
      Width           =   1290
   End
   Begin VB.Label lblCustomerId 
      Caption         =   "Customer ID :"
      Height          =   300
      Left            =   285
      TabIndex        =   5
      Top             =   -135
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "frmCustInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents LookUpDialog As frmLookUp
Attribute LookUpDialog.VB_VarHelpID = -1
Private m_CustID As Long
Private Sub LoadCustomerInfo()

Dim CustomerHasBalance As Boolean

'First Fill the Deposit Balnace & AccOUnt Balance
CustomerHasBalance = False
With lvw(1)
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear
    
    .ColumnHeaders.Add , "AccName", LoadResString(gLangOffSet + 34)  ', 1000, 0   '" Account Name"
    .ColumnHeaders.Add , "AccID", LoadResString(gLangOffSet + 36)     '" Account No"
    .ColumnHeaders.Add , "Balance", LoadResString(gLangOffSet + 42)     ' "BAlance"
    .ColumnHeaders(1).Alignment = lvwColumnLeft
    .ColumnHeaders(2).Alignment = lvwColumnLeft
    .ColumnHeaders(3).Alignment = lvwColumnRight
    .Visible = True

Dim CustObj As Object
Dim Balance As Currency
Dim AccId As Long
Dim ItemX As ListItem

    Set CustObj = New clsMMAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If AccId > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 49) & " " & LoadResString(gLangOffSet + 60))   '" Member ID"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsSBAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 436))
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsCAAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 422))  '" Curreen Account"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsFDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 423))  ' "Fixed Deposit"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsRDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 424))  ' " REcurring Deposit"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsPDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 425))  ' Pigmy Deposit"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    Set CustObj = New clsDLAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , CustObj.DepositName)  ' "DhanLAxmi"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = Nothing
    
    .Visible = True
End With
 
 ' Now Load the Loan Details of Customer
 CustomerHasBalance = False
With lvw(2)
    .ZOrder 0
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear
    
    .ColumnHeaders.Add , "Loans", LoadResString(gLangOffSet + 18)  ', 1000, 0   '" LOanName"
    .ColumnHeaders.Add , "AccID", LoadResString(gLangOffSet + 36)     '" Account No"
    .ColumnHeaders.Add , "Balance", LoadResString(gLangOffSet + 42)    ' "BAlance"
    .ColumnHeaders(1).Alignment = lvwColumnLeft
    .ColumnHeaders(2).Alignment = lvwColumnLeft
    .ColumnHeaders(3).Alignment = lvwColumnRight
    .view = lvwReport
    Set CustObj = New clsFDAcc
    Balance = CustObj.CustomerLoanBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 423))
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = New clsRDAcc
    Balance = CustObj.CustomerLoanBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 424))
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    Set CustObj = New clsPDAcc
    Balance = CustObj.CustomerLoanBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , LoadResString(gLangOffSet + 425))
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    
    Set CustObj = New clsDLAcc
    Balance = CustObj.CustomerLoanBalance(m_CustID, AccId)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , CustObj.DepositName)  ' "DhanLAxmi"
        ItemX.SubItems(1) = AccId
        ItemX.SubItems(2) = FormatCurrency(Balance)
    End If
    
' Now Load the General Loans
    Dim SchemeId() As Long
    Dim LoanName() As String
    Dim Retval As Integer
    Set CustObj = New clsLoan
    Retval = CustObj.LoanList(LoanName, SchemeId)
    If Retval = 0 Then GoTo Exit_Line
    For Retval = LBound(SchemeId) To UBound(SchemeId)
        Balance = CustObj.CustomerBalance(m_CustID, SchemeId(Retval), AccId)
        If Balance > 0 Then
            CustomerHasBalance = True
            '.ColumnHeaders(0).SubItemIndex
            Set ItemX = .ListItems.Add(, , LoanName(Retval))
            ItemX.SubItems(1) = AccId
            ItemX.SubItems(2) = FormatCurrency(Balance)
        End If
    Next Retval
    .Visible = True
End With

Exit_Line:
lvw(1).Visible = True
lvw(2).Visible = True
lvw(1).ZOrder 0
End Sub

Private Sub ResetUserInterface()
Me.lvw(1).ListItems.Clear
Me.lvw(2).ListItems.Clear
txtCustID.Text = ""
txtCustName.Text = ""
End Sub
Private Sub SetKannadaCaption()
Dim Ctrl As Control
For Each Ctrl In Me
    Ctrl.Font.Name = gFontName
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Size = gFontSize
    End If
Next
Me.lblCustomerId.Caption = LoadResString(gLangOffSet + 445) '& " " & LoadResString(gLangOffSet + 140)
Me.lblCustomerName.Caption = LoadResString(gLangOffSet + 35) ' & " " & LoadResString(gLangOffSet + 69)
Me.cmdLoad.Caption = LoadResString(gLangOffSet + 3)
Me.cmdClose.Caption = LoadResString(gLangOffSet + 11)
Me.TabStrip1.Tabs(1).Caption = LoadResString(gLangOffSet + 90)
Me.TabStrip1.Tabs(2).Caption = LoadResString(gLangOffSet + 18)


End Sub


Private Sub cmdClose_Click()
Unload Me
Set frmCustInfo = Nothing
End Sub

Private Sub cmdLoad_Click()
    Dim CustObj As New clsCustReg
    CustObj.LoadCustomerInfo (m_CustID)
    Me.txtCustName.Text = CustObj.FullName
    Call LoadCustomerInfo

End Sub

Private Sub cmdLookup_Click()
Screen.MousePointer = vbHourglass
On Error GoTo Err_Line
Dim Lret As Long

' Query the database to get all the customer names...
With gDBTrans
    .SQLStmt = "SELECT CustomerId, title + ' ' + FirstName " & _
        " + ' ' + MiddleName + ' ' + LastName as Name   FROM NameTab "
        
        'Now Check Whether He Want Search   any particular name
        'NameStr = InputBox("Enter customer name , You want search", "Name Search")
        Dim NameStr As String
        Screen.MousePointer = vbDefault
        NameStr = InputBox(LoadResString(gLangOffSet + 785), "Name Search")
        If NameStr <> "" Then
            .SQLStmt = gDBTrans.SQLStmt & " Where ( FirstNAme like '" & NameStr & "*' " & _
                                        " Or MiddleName like '" & NameStr & "*' Or LAstName like '" & NameStr & "*' )"
        End If
            .SQLStmt = .SQLStmt & " Order By IsciName"
    Lret = .SQLFetch
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox "No data available!", vbExclamation
        GoTo Exit_Line
    End If
End With

' Create a report dialog.
If LookUpDialog Is Nothing Then
    Set LookUpDialog = New frmLookUp
End If
m_CustID = 0
With LookUpDialog
    .m_SelItem = ""
    ' Fill the data to report dialog.
    If Not FillView(.lvwReport, gDBTrans.Rst, True) Then
        'MsgBox "Error filling the customer details.", vbCritical
        MsgBox "Error filling the customer details.", vbCritical
        GoTo Exit_Line
    End If
    Screen.MousePointer = vbDefault
    ' Display the dialog.
    .Show vbModal
    txtCustID.Text = m_CustID
End With
    Dim CustClass As New clsCustReg
    If m_CustID = 0 Then
        ResetUserInterface
    Else
        Me.txtCustName = CustClass.CustomerName(m_CustID)
        LoadCustomerInfo
    End If

Exit_Line:

Exit Sub

Err_Line:
MsgBox Err.Number & " " & Err.Description, vbCritical, wis_MESSAGE_TITLE

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
      If Me.TabStrip1.SelectedItem.Index = TabStrip1.Tabs.Count Then
            TabStrip1.Tabs(1).Selected = True
      Else
            TabStrip1.Tabs(TabStrip1.SelectedItem.Index + 1).Selected = True
      End If
End If

End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   ""(Me.hwnd, False)
End Sub

Private Sub LookUpDialog_SelectClick(strSelection As String)
    m_CustID = Val(strSelection)
End Sub


Private Sub TabStrip1_Click()
    lvw(TabStrip1.SelectedItem.Index).ZOrder 0
    lvw(TabStrip1.SelectedItem.Index).Visible = True
End Sub

Private Sub txtCustID_Change()
    
    m_CustID = Val(txtCustID.Text)
    If m_CustID = 0 Then Exit Sub
    
End Sub


