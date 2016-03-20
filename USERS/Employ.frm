VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmEmpoyee 
   Caption         =   "Employee Details"
   ClientHeight    =   3675
   ClientLeft      =   1905
   ClientTop       =   2355
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6795
   Begin VB.TextBox txtEpfNo 
      Height          =   315
      Left            =   4830
      TabIndex        =   18
      Top             =   2094
      Width           =   1605
   End
   Begin VB.ComboBox cmbDesignation 
      Height          =   315
      ItemData        =   "Employ.frx":0000
      Left            =   1470
      List            =   "Employ.frx":001F
      TabIndex        =   3
      Text            =   "Designation"
      Top             =   606
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   4050
      TabIndex        =   6
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox txtDA 
      Height          =   315
      Left            =   1470
      TabIndex        =   11
      Top             =   1608
      Width           =   855
   End
   Begin VB.TextBox txtHra 
      Height          =   315
      Left            =   4830
      TabIndex        =   13
      Top             =   1608
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   5400
      TabIndex        =   9
      Top             =   3150
      Width           =   1215
   End
   Begin WIS_Currency_Text_Box.CurrText txtBasic 
      Height          =   345
      Left            =   1470
      TabIndex        =   5
      Top             =   1092
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin WIS_Currency_Text_Box.CurrText txtTotalSalary 
      Height          =   345
      Left            =   1470
      TabIndex        =   15
      Top             =   2094
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin WIS_Currency_Text_Box.CurrText txtOther 
      Height          =   345
      Left            =   4860
      TabIndex        =   8
      Top             =   1092
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.Label txtPassword 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4830
      TabIndex        =   16
      Top             =   2610
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label txtLoginName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1470
      TabIndex        =   20
      Top             =   2610
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEpfNo 
      Caption         =   "EPF Number :"
      Height          =   315
      Left            =   3420
      TabIndex        =   17
      Top             =   2094
      Width           =   1335
   End
   Begin VB.Label lblDeignation 
      Caption         =   "Designation"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label lblLoginName 
      Caption         =   "Login name:"
      Height          =   315
      Left            =   60
      TabIndex        =   19
      Top             =   2610
      Width           =   1335
   End
   Begin VB.Label lblLoginPassword 
      Caption         =   "Login password:"
      Height          =   315
      Left            =   3450
      TabIndex        =   21
      Top             =   2610
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   6450
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Label lblTotalsalary 
      Caption         =   "Total salry"
      Height          =   285
      Left            =   30
      TabIndex        =   14
      Top             =   2094
      Width           =   1365
   End
   Begin VB.Label lblOther 
      Caption         =   "Other renumeration"
      Height          =   285
      Left            =   3420
      TabIndex        =   7
      Top             =   1092
      Width           =   1365
   End
   Begin VB.Label lblHRA 
      Caption         =   "HRA in %"
      Height          =   315
      Left            =   3420
      TabIndex        =   12
      Top             =   1608
      Width           =   1305
   End
   Begin VB.Label lblDA 
      Caption         =   "DA in %"
      Height          =   285
      Left            =   60
      TabIndex        =   10
      Top             =   1608
      Width           =   1365
   End
   Begin VB.Label lblBasic 
      Caption         =   "Basic Salary"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1092
      Width           =   1365
   End
   Begin VB.Label txtEmpName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "employee Name"
      Height          =   345
      Left            =   1470
      TabIndex        =   1
      Top             =   90
      Width           =   4965
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEmpName 
      Caption         =   "Emp Name :"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmpoyee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CustomerID As Long
Private m_UserID As Integer

Private m_LoadedForm As Boolean
Private m_dbOperation As wis_DBOperation

Public Property Let UserID(NewValue As Long)

m_UserID = NewValue
cmdOk.Enabled = False
If m_LoadedForm And m_UserID Then Call GetEmployeeDetail

End Property

Private Sub GetEmployeeDetail()

If m_UserID <= 0 Then Exit Sub
cmdOk.Enabled = True
m_dbOperation = Insert

'Now Get the Details of Emloyee
'Get Empoyee Name
Dim rst As Recordset
gDbTrans.SqlStmt = "Select A.UserID,A.CustomerID,LoginName,LoginPassword," & _
    " Title +' ' +FirstName+' '+MiddleName+' '+LastName as CustName,FullName" & _
    " From UserTab A,NameTab B" & _
    " Where A.CustomerID = B.CustomerID ANd A.UserID = " & m_UserID

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Unload Me

txtEmpName = Trim$(FormatField(rst("CustName")))
txtEmpName.Tag = Trim$(FormatField(rst("FullName")))
txtLoginName = FormatField(rst("LoginName"))
txtPassword = FormatField(rst("LoginPassword"))
m_CustomerID = FormatField(rst("CustomerID"))

gDbTrans.SqlStmt = "Select * From EmpDetails " & _
    " Where CustomerID = " & m_CustomerID
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

m_dbOperation = Update

cmbDesignation.ListIndex = FormatField(rst("Designation")) - 1
txtBasic = FormatField(rst("BasicSalary"))
txtTotalSalary = FormatField(rst("Netsalary"))
txtDA = FormatField(rst("DA"))
txtHra = FormatField(rst("HRA"))
txtOther = FormatField(rst("OtherRenum"))
txtEpfNo = FormatField(rst("EpfNum"))


End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'lblCaption.Caption = GetResourceString(149)
'lblUserReg.Caption = GetResourceString(150)
'cmdDetail.Caption = GetResourceString(295)  'Details

lblLoginName.Caption = GetResourceString(151, 35)
lblLoginPassword.Caption = GetResourceString(151, 153)
'lblConfirmPassword.Caption = GetResourceString(154,153)

'lblPermissions.Caption = GetResourceString(156)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)

End Sub

Private Sub cmdCancel_Click()
    m_LoadedForm = False
    m_dbOperation = Insert
    Unload Me
End Sub

Private Sub cmdOk_Click()

'Now Check the Desigantion
If cmbDesignation.ListIndex < 0 Then
    MsgBox "Please select the empoyees designation"
    Exit Sub
End If

'Basic salary
If txtBasic = 0 Then
    'Invalid currency
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If


'DA
If Not CurrencyValidate(txtDA, True) Then
    'Invalid currency
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDA
    Exit Sub
End If

If Not CurrencyValidate(txtHra, True) Then
    'Invalid currency
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtHra
    Exit Sub
End If

If txtTotalSalary = 0 Then
    'Invalid currency
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTotalSalary
    Exit Sub
End If

Dim DA As Single
Dim HRA As Single
Dim EpfNum As String
Dim AdvanceID As Long
Dim SalaryID As Long
Dim headName As String
Dim headNameEnglish As String
Dim Designation As Integer

Designation = cmbDesignation.ListIndex + 1

DA = Val(txtDA)
HRA = Val(txtHra)
EpfNum = Trim$(txtEpfNo)

Dim bankClass As clsBankAcc

If m_dbOperation = Update Then
    Dim rst As Recordset
    gDbTrans.SqlStmt = "Select * From EmpDetails " & _
        " WHERE CustomerID = " & m_CustomerID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        AdvanceID = FormatField(rst("AdvanceID"))
        SalaryID = FormatField(rst("SalaryID"))
    End If
End If

'Now INsert the details
Dim InTrans As Boolean
gDbTrans.BeginTrans
InTrans = True

headName = Left(Trim(txtEmpName), 200)
headNameEnglish = Trim$(txtEmpName.Tag)
If AdvanceID = 0 And Designation > 2 Then
    
    Set bankClass = New clsBankAcc
    'salary expense
    SalaryID = bankClass.GetHeadIDCreated(headName, headNameEnglish, parSalaryExpense, 0, wis_Users)
    'Salary advance
    AdvanceID = bankClass.GetHeadIDCreated(headName & "1", headNameEnglish & "1", parSalaryAdvance, 0, wis_Users)
    Set bankClass = Nothing
    'BankClass.up

ElseIf AdvanceID And Designation > 2 Then
    
    gDbTrans.SqlStmt = "UPdate Heads Set HeadName = " & AddQuotes(headName, True) & _
        " Where HeadID = " & SalaryID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
    gDbTrans.SqlStmt = "UPdate Heads Set HeadName = " & AddQuotes(headName, True) & _
        " Where HeadID = " & AdvanceID
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
End If


If m_dbOperation = Insert Then
    
    gDbTrans.SqlStmt = "Insert Into EmpDetails " & _
        "(CustomerID,UserID,Designation,BasicSalary," & _
        "EPFNum,DA,HRA,OtherRenum,SalaryId,AdvanceID)" & _
        " VALUES (" & m_CustomerID & "," & m_UserID & "," & Designation & "," & _
        txtBasic & "," & AddQuotes(txtEpfNo, True) & "," & _
        DA & "," & HRA & "," & txtOther & "," & _
        SalaryID & "," & AdvanceID & ")"
    
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
Else
    
    gDbTrans.SqlStmt = "UPDATE EmpDetails SET " & _
        "Designation = " & Designation & "," & _
        "NetSalary = " & txtTotalSalary & ",BasicSalary = " & txtBasic & "," & _
        "EPFNum = " & AddQuotes(txtEpfNo, True) & "," & _
        "DA = " & DA & ", HRA = " & HRA & "," & _
        "OtherRenum = " & txtOther & "," & _
        "SalaryId = " & SalaryID & ",AdvanceID = " & AdvanceID & _
        " WHERE CustomerID = " & m_CustomerID
    
    If Not gDbTrans.SQLExecute Then GoTo ExitLine
    
End If


gDbTrans.CommitTrans
InTrans = False
Unload Me
    
ExitLine:
    If InTrans Then
        gDbTrans.RollBack
        MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

cmdOk.Enabled = False

Call CenterMe(Me)

'Now Set the Keannada captin
Call SetKannadaCaption
m_LoadedForm = True
If m_UserID Then Call GetEmployeeDetail

End Sub


