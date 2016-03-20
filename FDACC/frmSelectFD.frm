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
Public Event OKClick(ByVal DepositType As Integer)
Public Event CancelClick()

Public Property Let GroupType(NewValue As wis_GroupType)
    m_grpType = NewValue
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
'    Else
'        Unload Me
    End If
End Sub
Private Sub LoadDeposits()
Dim rst As ADODB.Recordset
    gDbTrans.SqlStmt = "SELECT * FROM DepositName"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        List1.Clear
        While Not rst.EOF
            List1.AddItem rst("DepositName")
            List1.ItemData(List1.newIndex) = rst("DepositID")
            rst.MoveNext
        Wend
        Set rst = Nothing
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
    RaiseEvent CancelClick
End Sub


'
Private Sub cmdOk_Click()
Dim DepositType As Integer

If List1.ListIndex = -1 Then Exit Sub

DepositType = List1.ItemData(List1.ListIndex)

Me.Hide

If Trim$(List1.Text) <> "" Then _
    RaiseEvent OKClick(DepositType)

End Sub

Private Sub Form_Load()
    SetKannadaCaption
    If m_grpType = grpMember Then
        LoadMemberTypes
    Else
         LoadDeposits
    End If
    If List1.ListCount Then List1.ListIndex = 0
End Sub


'
Private Sub List1_DblClick()
    'cmdOk_Click
End Sub


