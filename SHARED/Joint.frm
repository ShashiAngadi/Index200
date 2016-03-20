VERSION 5.00
Begin VB.Form frmJoint 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Joint holder names"
   ClientHeight    =   2025
   ClientLeft      =   3840
   ClientTop       =   2190
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Top             =   1275
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   860
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   445
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   1305
      TabIndex        =   0
      Top             =   1580
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2235
      TabIndex        =   1
      Top             =   1580
      Width           =   885
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   400
      Left            =   90
      TabIndex        =   10
      Top             =   1580
      Width           =   885
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   1230
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   830
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   430
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   2625
   End
End
Attribute VB_Name = "frmJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Status As String
Private m_JointCust(3) As clsCustReg
Private M_ModuleID As wisModules


Public Property Let ModuleID(NewValue As wisModules)
    M_ModuleID = NewValue
End Property


Public Function SaveJointCustomers() As Boolean
'if any customer is there in the joint list
'then save their details
Dim count As Integer
For count = 0 To 3
    If m_JointCust(count) Is Nothing Then Exit For
    If m_JointCust(count).customerID = 0 Then Exit For
    
    If Not m_JointCust(count).SaveCustomer Then Exit Function
Next

SaveJointCustomers = True
End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
Me.cmdOk.Caption = GetResourceString(1)  '"«Œâ³ú"
Me.cmdCancel.Caption = GetResourceString(2)  '"ÇðÁðôà"
End Sub

Private Sub cmdCancel_Click()
Me.Status = "CANCEL"
Me.Hide
End Sub


Private Sub cmdClear_Click()
Dim I As Integer
For I = 1 To 3
    Set m_JointCust(I) = Nothing
    txtName(I) = ""
Next
End Sub


Private Sub cmdJoint_Click(Index As Integer)


If Index > 0 Then
    If m_JointCust(Index - 1) Is Nothing Then Exit Sub
    If m_JointCust(Index - 1).customerID <= 0 Then Exit Sub
End If

If m_JointCust(Index) Is Nothing Then Set m_JointCust(Index) = New clsCustReg

m_JointCust(Index).ModuleID = wis_SBAcc
m_JointCust(Index).ShowDialog

'Check whether the selected customer is already  customer of this account
'If he is customer of this account then do not show him
Dim count As Integer
Dim From As Integer


For count = 0 To 3
    If m_JointCust(count) Is Nothing Then Exit For
    If count <> Index Then
        If m_JointCust(Index).customerID = m_JointCust(count).customerID Then
            m_JointCust(Index).NewCustomer
            Exit Sub
        End If
    End If
Next

If m_JointCust(Index).customerID > 0 Then _
    txtName(Index) = m_JointCust(Index).FullName
    
End Sub

Private Sub cmdOk_Click()
Me.Status = "OK"
'if any customer is there in the joint list
'then save their details
Dim count As Integer
For count = 1 To 3
    If m_JointCust(count) Is Nothing Then Exit For
    If m_JointCust(count).customerID = 0 Then Exit For
'    Call m_JointCust(Count).SaveCustomer
Next

Me.Hide
End Sub

Private Sub Form_Load()

'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
Dim I As Integer
For I = 0 To txtName.count - 1
    txtName(I).Width = Me.ScaleWidth - cmdJoint(I).Width * 2
    cmdJoint(I).Left = txtName(I).Left + txtName(I).Width + 100
Next
cmdCancel.Left = cmdJoint(0).Left + cmdJoint(0).Width - cmdCancel.Width
cmdOk.Left = cmdCancel.Left - 50 - cmdOk.Width

Height = cmdOk.Top + cmdOk.Height + 500
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Public Property Get JointCustId(Index As Integer) As Long

If Not m_JointCust(Index) Is Nothing Then
    JointCustId = m_JointCust(Index).customerID
End If

Exit Property


Dim I As Integer
Dim Str As String
For I = 0 To txtName.count - 1
    If Trim$(txtName(I)) <> "" Then
        Str = Str & txtName(I).Tag & ";"
    End If
Next
'JointCustId = Str
End Property
Public Property Let JointCustId(Index As Integer, ByVal NewValue As Long)
Dim rst As ADODB.Recordset

'While Creating the new accounts we load the account holder detail hust into the form
'which not loades in such case It may give error
'So in such time just do not shoe the details

On Error GoTo Err_line
'Check the existance of customer
gDbTrans.SqlStmt = "SELECT * FROM NameTab WHERE CustomerId = " & Val(NewValue)
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    'MsgBox "No customer"
    If Index > 0 Then MsgBox GetResourceString(662) & NewValue, vbExclamation, wis_MESSAGE_TITLE
    Set m_JointCust(Index) = Nothing
    Exit Property
End If
If NewValue = 0 Then Exit Property

If m_JointCust(Index) Is Nothing Then _
    Set m_JointCust(Index) = New clsCustReg

m_JointCust(Index).ModuleID = M_ModuleID
m_JointCust(Index).LoadCustomerInfo (NewValue)
txtName(Index) = m_JointCust(Index).FullName
Exit Property

' Breakup the string into array components.
Dim strTmp() As String
Dim I As Integer
'
Err_line:
End Property
Public Property Get JointHolders() As String
Dim I As Integer
Dim Str As String

For I = 0 To txtName.count - 1
    If Trim$(txtName(I)) <> "" Then Str = Str & txtName(I) & ";"
Next
JointHolders = Str
End Property

Private Sub Form_Unload(Cancel As Integer)
Dim I As Integer
For I = 0 To txtName.count - 1
    Set m_JointCust(I) = Nothing
Next

End Sub


