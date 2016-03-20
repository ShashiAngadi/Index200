VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCustInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INDEX2000  -  Customer information"
   ClientHeight    =   4935
   ClientLeft      =   135
   ClientTop       =   2940
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6135
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
      Height          =   400
      Left            =   4785
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtCustName 
      Height          =   345
      Left            =   1650
      TabIndex        =   7
      Top             =   165
      Width           =   3885
   End
   Begin VB.CommandButton cmdLookUP 
      Caption         =   "..."
      Height          =   285
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   315
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
      Top             =   1080
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3750
      Left            =   330
      TabIndex        =   3
      Top             =   600
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   6615
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
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
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Surity"
            Key             =   "Surity"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Surity"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lvw 
      Height          =   3075
      Index           =   1
      Left            =   435
      TabIndex        =   8
      Top             =   1080
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5424
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView lvw 
      Height          =   3075
      Index           =   3
      Left            =   390
      TabIndex        =   10
      Top             =   1080
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblCustomerName 
      Caption         =   "Customer Name :"
      Height          =   315
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
Public Event WindowClosed()
Private Sub LoadCustomerInfo()

Dim CustomerHasBalance As Boolean
Dim AsOnDate As Date
Dim rst As Recordset

AsOnDate = GetSysFormatDate(gStrDate)

'First Fill the Deposit Balnace & AccOUnt Balance
CustomerHasBalance = False
Me.Refresh
With lvw(1)
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear
    
    .ColumnHeaders.Add , "AccName", GetResourceString(36, 35) ', 1000, 0   '" Account Name"
    .ColumnHeaders.Add , "AccID", GetResourceString(36, 60)    '" Account No"
    .ColumnHeaders.Add , "Balance", GetResourceString(42)     ' "BAlance"
    .ColumnHeaders(1).Alignment = lvwColumnLeft
    .ColumnHeaders(2).Alignment = lvwColumnLeft
    .ColumnHeaders(3).Alignment = lvwColumnRight
    .Visible = True

    Dim CustObj As Object
    Dim Balance As Currency
    Dim AccId As Long
    Dim AccNum As String
    Dim ItemX As ListItem
    Dim SchemeID() As Long
    Dim LoanName() As String
    Dim Retval As Integer
    Dim AccountTypeName As String

    Set CustObj = New clsMMAcc
    Balance = CustObj.CustomerBalanceNew(m_CustID, AccNum, AccountTypeName, AsOnDate)
    If Len(AccNum) > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , IIf(Len(AccountTypeName), AccountTypeName, GetResourceString(49, 60))) '" Member No"
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        'Balance = 0
        While Balance
            Balance = CustObj.NextBalance(AccNum)
            If Balance > 0 Then
                CustomerHasBalance = True
                '.ColumnHeaders(0).SubItemIndex
                Set ItemX = .ListItems.Add(, , GetResourceString(49, 60))  '" Member ID"
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    
    Set CustObj = Nothing
    
    Set CustObj = New clsSBAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , GetResourceString(421))
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        While Balance
            Balance = CustObj.NextBalance(AccNum)
            If Balance > 0 Then
                CustomerHasBalance = True
                '.ColumnHeaders(0).SubItemIndex
                Set ItemX = .ListItems.Add(, , GetResourceString(421))
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    Set CustObj = Nothing
    
    Set CustObj = New ClsCAAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , GetResourceString(422))  '" Current Account"
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        While Balance
            Balance = CustObj.NextBalance(AccNum)
            If Balance > 0 Then
                CustomerHasBalance = True
                '.ColumnHeaders(0).SubItemIndex
                Set ItemX = .ListItems.Add(, , GetResourceString(422))  '" Current Account"
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    Set CustObj = Nothing
    Dim StrName As String
    Me.Refresh
    Set CustObj = New clsFDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AsOnDate, AccNum, StrName)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , StrName)  ' "Fixed Deposit"
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        While Balance
            Balance = CustObj.NextBalance(AccNum, StrName)
            If Balance > 0 Then
                CustomerHasBalance = True
                Set ItemX = .ListItems.Add(, , StrName)  ' "Fixed Deposit"
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    Set CustObj = Nothing
    Me.Refresh
    Set CustObj = New clsRDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    If Balance > 0 Then
        CustomerHasBalance = True
        Set ItemX = .ListItems.Add(, , GetResourceString(424))  ' " REcurring Deposit"
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        While Balance
            Balance = CustObj.NextBalance(AccNum)
            If Balance > 0 Then
                CustomerHasBalance = True
                Set ItemX = .ListItems.Add(, , GetResourceString(424))  ' " REcurring Deposit"
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    Set CustObj = Nothing
    Me.Refresh
    Set CustObj = New clsPDAcc
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    If Balance > 0 Then
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , GetResourceString(425))  ' Pigmy Deposit"
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        While Balance
            Balance = CustObj.NextBalance(AccNum)
            If Balance > 0 Then
                CustomerHasBalance = True
                '.ColumnHeaders(0).SubItemIndex
                Set ItemX = .ListItems.Add(, , GetResourceString(425))  ' Pigmy Deposit"
                ItemX.SubItems(1) = AccNum
                ItemX.SubItems(2) = FormatCurrency(Balance)
            End If
        Wend
    End If
    
    Set CustObj = Nothing
    Me.Refresh
    Set CustObj = New clsBkcc
    StrName = GetResourceString(229, 43)
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    While Balance
        If Balance < 0 Then
            CustomerHasBalance = True
            Set ItemX = .ListItems.Add(, , StrName)
            ItemX.SubItems(1) = AccNum
            ItemX.SubItems(2) = FormatCurrency(Balance)
        End If
        Balance = CustObj.NextBalance(AccNum)
    Wend
    
    Me.Refresh
    .Visible = True
End With
 
'Now Load the Loan Details of Customer
CustomerHasBalance = False
With lvw(2)
    .ZOrder 0
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear
    
    .ColumnHeaders.Add , "Loans", GetResourceString(18)  ', 1000, 0   '" LOanName"
    .ColumnHeaders.Add , "AccID", GetResourceString(36)      '" Account No"
    .ColumnHeaders.Add , "Balance", GetResourceString(42)    ' "BAlance"
    .ColumnHeaders(1).Alignment = lvwColumnLeft
    .ColumnHeaders(2).Alignment = lvwColumnLeft
    .ColumnHeaders(3).Alignment = lvwColumnRight
    .view = lvwReport
    
    Set CustObj = New clsDepLoan
    Dim Deptype As Long
    Balance = CustObj.CustomerBalance(m_CustID, AsOnDate, Deptype, AccNum)
    While Balance
        CustomerHasBalance = True
        '.ColumnHeaders(0).SubItemIndex
        Set ItemX = .ListItems.Add(, , GetDepositTypeText(CInt(Deptype)))
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        Balance = CustObj.NextBalance(Deptype, AccNum)
    Wend
    
    Set CustObj = New clsBkcc
    StrName = GetResourceString(229, 58)
    Balance = CustObj.CustomerBalance(m_CustID, AccNum, AsOnDate)
    While Balance
        CustomerHasBalance = True
        If Balance > 0 Then
            CustomerHasBalance = True
            Set ItemX = .ListItems.Add(, , StrName)
            ItemX.SubItems(1) = AccNum
            ItemX.SubItems(2) = FormatCurrency(Balance)
        End If
        Balance = CustObj.NextBalance(AccNum)
    Wend
    
    Set CustObj = New clsLoan
    Balance = CustObj.CustomerBalance(m_CustID, AsOnDate, AccNum, StrName)
    While Balance > 0
        CustomerHasBalance = True
        Set ItemX = .ListItems.Add(, , StrName)
        ItemX.SubItems(1) = AccNum
        ItemX.SubItems(2) = FormatCurrency(Balance)
        
        Balance = CustObj.NextBalance(AccNum, StrName)
    Wend
    Set CustObj = Nothing
    .Visible = True
End With
Me.Refresh

With lvw(3)
    Dim SqlStr As String

    .ZOrder 0
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    .ColumnHeaders.Add , "AccID", GetResourceString(36)  ', 1000, 0   '" LOanName"
    .ColumnHeaders.Add , "LOanID", GetResourceString(18)      '" Account No"
    .ColumnHeaders.Add , "CustName", GetResourceString(35)       '" Account No"
    .ColumnHeaders.Add , "Balance", GetResourceString(42)    ' "BAlance"
    .ColumnHeaders(1).Alignment = lvwColumnLeft
    .ColumnHeaders(2).Alignment = lvwColumnLeft
    .ColumnHeaders(3).Alignment = lvwColumnLeft
    .ColumnHeaders(4).Alignment = lvwColumnRight
    .view = lvwReport

    'If this customer is not a membere then
    'he can notbe guaranteer
    
    SqlStr = "SELECT H.LoanID FROM LoanMaster H " & _
        " Where Guarantor1 = " & m_CustID & " OR Guarantor2 = " & m_CustID & _
        " OR Guarantor3 = " & m_CustID & " OR Guarantor4 = " & m_CustID
        
    gDbTrans.SqlStmt = "SELECT A.AccNum,A.LoanID, SchemeID,Balance, " _
        & " Title + ' ' + FirstName + ' '+ MiddleName +' '+ LastName AS CustName" _
        & " From LoanMaster A, LoanTrans B, NameTab D " _
        & " WHERE TransID = (SELECT Max(TransID) From LoanTrans E " & _
            " WHERE E.LoanID = B.LoanID ) " _
        & " AND Balance > 0 And A.LoanID IN (" & SqlStr & ") AND B.LOanID = A.LOanID " _
        & " AND D.CustomerID = A.CustomerID "
        
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Set CustObj = New clsLoan
        While Not rst.EOF
            CustomerHasBalance = True
            Set ItemX = .ListItems.Add(, , rst("AccNum"))
            ItemX.SubItems(1) = CustObj.LoanName(rst("SchemeID"))
            ItemX.SubItems(2) = rst("CustName")
            ItemX.SubItems(3) = rst("Balance")
            
            rst.MoveNext
        Wend
        Set CustObj = Nothing
    End If

    SqlStr = "SELECT H.LoanID FROM BKCCMaster H " & _
        " Where Guarantor1 = " & m_CustID & " OR Guarantor2 = " & m_CustID
        
    gDbTrans.SqlStmt = "SELECT A.AccNUm,A.LoanID, Balance, " _
        & " Title + ' ' + FirstName + ' '+ MiddleName + ' ' + LastName AS CustName" _
        & " From BKCCMaster A, BKCCTrans B, NameTab D " _
        & " WHERE A.LoanID IN (" & SqlStr & ") " _
        & " AND TransID = (SELECT Max(TransID)" & _
            " From BKCCTrans E WHERE E.LoanID = B.LoanID ) " _
        & " AND Balance > 0 And B.LOanID = A.LOanID " _
        & " AND D.CustomerID = A.CustomerID "
        
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        While rst.EOF
            CustomerHasBalance = True
            '.ColumnHeaders(0).SubItemIndex
            Set ItemX = .ListItems.Add(, , rst("AccNum"))
            ItemX.SubItems(1) = GetResourceString(229)
            ItemX.SubItems(2) = rst("CustName")
            ItemX.SubItems(3) = rst("Balance")
            
            rst.MoveNext
        Wend
    End If
    .Visible = True
End With

Exit_Line:
lvw(1).Visible = True
lvw(2).Visible = True
lvw(3).Visible = True
Set CustObj = Nothing

lvw(1).ZOrder 0
End Sub

Private Sub ResetUserInterface()
Me.lvw(1).ListItems.Clear
Me.lvw(2).ListItems.Clear
txtCustID.Text = ""
txtCustName.Text = ""
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
lblCustomerId.Caption = GetResourceString(205, 60)
lblCustomerName.Caption = GetResourceString(35)
cmdLoad.Caption = GetResourceString(3)
cmdClose.Caption = GetResourceString(11)

TabStrip1.Tabs(1).Caption = GetResourceString(36) & GetResourceString(92)
TabStrip1.Tabs(2).Caption = GetResourceString(18)
TabStrip1.Tabs(3).Caption = GetResourceString(389)

End Sub


Private Sub cmdClose_Click()
Unload Me
Set frmCustInfo = Nothing
End Sub

Private Sub cmdLoad_Click()
    Dim CustObj As clsCustReg
    Set CustObj = New clsCustReg
    CustObj.LoadCustomerInfo (m_CustID)
    Me.txtCustName.Text = CustObj.FullName
    'Call LoadCustomerInfo
    Set CustObj = Nothing
End Sub

Private Sub cmdLookup_Click()
Screen.MousePointer = vbHourglass
On Error GoTo Err_line
Dim Lret As Long
Dim rst As Recordset

' Query the database to get all the customer names...
With gDbTrans
    .SqlStmt = "SELECT CustomerId, title + ' ' + FirstName " & _
        " + ' ' + MiddleName + ' ' + LastName as Name, Place, HomeAddress  FROM NameTab "
        
        'Now Check Whether He Want Search   any particular name
        'NameStr = InputBox("Enter customer name , You want search", "Name Search")
        Dim NameStr As String
        Screen.MousePointer = vbDefault
        NameStr = InputBox(GetResourceString(785), "Name Search")
        If NameStr <> "" Then
            .SqlStmt = gDbTrans.SqlStmt & " Where ( FirstNAme like " & _
                    "'" & NameStr & "%' Or MiddleName like '" & NameStr & "%'" & _
                    "Or LAstName like '" & NameStr & "%' )"
        End If
            .SqlStmt = .SqlStmt & " Order By IsciName"
    Lret = .Fetch(rst, adOpenDynamic)
    If Lret <= 0 Then
        'MsgBox "No data available!", vbExclamation
        MsgBox "No data available!", vbExclamation
        GoTo Exit_Line
    End If
End With

' Create a report dialog.
If LookUpDialog Is Nothing Then Set LookUpDialog = New frmLookUp

m_CustID = 0
With LookUpDialog
    .m_SelItem = ""
    ' Fill the data to report dialog.
    If Not FillView(.lvwReport, rst, True) Then
        'MsgBox "Error filling the customer details.", vbCritical
        MsgBox "Error filling the customer details.", vbCritical
        GoTo Exit_Line
    End If
    Screen.MousePointer = vbDefault
    ' Display the dialog.
    .Show vbModal
    txtCustID.Text = m_CustID
End With

    Dim CustClass As clsCustReg
    Set CustClass = New clsCustReg
    If m_CustID = 0 Then
        ResetUserInterface
    Else
        txtCustName = CustClass.CustomerName(m_CustID)
        LoadCustomerInfo
        TabStrip1_Click
    End If
    Set CustClass = Nothing
    
Exit_Line:
Set CustClass = Nothing
Exit Sub

Err_line:
    MsgBox Err.Number & " " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    'Resume
    Err.Clear

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyTab Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
If Not CtrlDown Then Exit Sub


Dim IndxNo As Integer
IndxNo = TabStrip1.SelectedItem.Index
IndxNo = IndxNo + IIf(Shift = 2, 1, -1)
If IndxNo < 1 Then IndxNo = TabStrip1.Tabs.count
If IndxNo > TabStrip1.Tabs.count Then IndxNo = 1
TabStrip1.Tabs(IndxNo).Selected = True


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
    Set LookUpDialog = Nothing
    RaiseEvent WindowClosed
    
End Sub

Private Sub LookUpDialog_SelectClick(strSelection As String)
    m_CustID = Val(strSelection)
End Sub


Private Sub TabStrip1_Click()
    lvw(3).Visible = False
    lvw(1).Visible = False
    lvw(2).Visible = False
    lvw(TabStrip1.SelectedItem.Index).ZOrder 0
    lvw(TabStrip1.SelectedItem.Index).Visible = True
End Sub

Private Sub txtCustID_Change()
    m_CustID = Val(txtCustID.Text)
    If m_CustID = 0 Then Exit Sub
End Sub
