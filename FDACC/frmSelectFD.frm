VERSION 5.00
Begin VB.Form frmSelectFD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Deposit"
   ClientHeight    =   1965
   ClientLeft      =   4365
   ClientTop       =   2865
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2985
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   495
      TabIndex        =   3
      Top             =   360
      Width           =   2385
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1965
      TabIndex        =   2
      Top             =   1530
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1005
      TabIndex        =   1
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label lblSelect 
      Caption         =   "Select Deposit :"
      Height          =   285
      Left            =   495
      TabIndex        =   0
      Top             =   30
      Width           =   1995
   End
End
Attribute VB_Name = "frmSelectFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_grpType As wis_GroupType
Private M_ModuleID As wisModules
Private m_MultiDeposit As Boolean
Public Event OKClick(ByVal DepositType As Integer)
Public Event CancelClick()
Public Event DepositCount(ByVal NumOfDeposits As Integer)


Public Property Let GroupType(NewValue As wis_GroupType)
    m_grpType = NewValue
End Property
Public Property Get multiDeposit() As Boolean
    multiDeposit = m_MultiDeposit
End Property

Public Property Let ModuleID(NewValue As wisModules)
    M_ModuleID = NewValue
End Property

Private Sub LoadMemberTypes()

Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "SELECT * FROM MemberTypeTab"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        List1.Clear
        While Not rst.EOF
            List1.AddItem rst("MemberTypeName")
            List1.ItemData(List1.newIndex) = rst("MemberTYpe")
            rst.MoveNext
        Wend
        Set rst = Nothing
        m_MultiDeposit = True
        If List1.ListCount = 1 Then
            m_MultiDeposit = False
            RaiseEvent OKClick(List1.ItemData(0))
        End If
'    Else
'        Unload Me
    End If
End Sub

Private Sub LoadFixedDepositTypes()
    Dim rst As ADODB.Recordset
    Dim recordCount As Integer
    Dim DepositType As Integer
    gDbTrans.SqlStmt = "SELECT * FROM DepositTypeTab where moduleid =" & M_ModuleID
    recordCount = gDbTrans.Fetch(rst, adOpenForwardOnly)
    List1.Clear
    If recordCount > 0 Then
        m_MultiDeposit = True
        List1.Clear
        While Not rst.EOF
            List1.AddItem rst("DepositTypeName")
            List1.ItemData(List1.newIndex) = rst("DepositType")
            rst.MoveNext
        Wend
        Set rst = Nothing
        If List1.ListCount = 1 Then
            m_MultiDeposit = False
            'DepositType = IIf(recordCount = 0, 0, rst("DepositType"))
            RaiseEvent OKClick(List1.ItemData(0))
        End If
'    Else
'        Unload Me
    Else
        m_MultiDeposit = False
        DepositType = IIf(recordCount = 0, 0, rst("DepositType"))
        RaiseEvent OKClick(DepositType)
        'Me.Hide
        'Unload Me
        
    End If
End Sub
Private Sub LoadDeposits()
Dim rst As ADODB.Recordset
    List1.Clear
    gDbTrans.SqlStmt = "SELECT * FROM DepositName"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        List1.Clear
        While Not rst.EOF
            List1.AddItem rst("DepositName")
            List1.ItemData(List1.newIndex) = rst("DepositID")
            rst.MoveNext
        Wend
        Set rst = Nothing
        m_MultiDeposit = True
        If List1.ListCount = 1 Then
            m_MultiDeposit = False
            'DepositType = IIf(recordCount = 0, 0, rst("DepositType"))
            RaiseEvent OKClick(List1.ItemData(0))
        End If
'    Else
'        Unload Me
    End If
End Sub

'
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
Me.cmdOk.Caption = GetResourceString(1)
Me.cmdCancel.Caption = GetResourceString(2)

End Sub

'
Private Sub cmdCancel_Click()
    Me.Hide
    RaiseEvent DepositCount(List1.ListCount)
    RaiseEvent CancelClick
End Sub


'
Private Sub cmdOk_Click()
Dim DepositType As Integer

If List1.ListIndex = -1 Then Exit Sub

DepositType = List1.ItemData(List1.ListIndex)

Me.Hide
RaiseEvent DepositCount(List1.ListCount)
If Trim$(List1.Text) <> "" Then RaiseEvent OKClick(DepositType)

End Sub

Private Sub Form_Initialize()
    m_MultiDeposit = False
End Sub

Private Sub Form_Load()
    SetKannadaCaption
    If m_grpType = grpMember Then
        LoadMemberTypes
    ElseIf m_grpType = grpAllDeposit Then
        LoadFixedDepositTypes
    Else
         LoadDeposits
    End If
    'If List1.ListCount Then List1.ListIndex = 0
    If List1.ListCount < 2 Then
        'Unload Me
    Else
        List1.ListIndex = 0
    End If
    
End Sub

Private Sub List1_DblClick()
    'cmdOk_Click
End Sub


