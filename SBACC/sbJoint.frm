VERSION 5.00
Begin VB.Form frmSBJoint 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Joint holder names"
   ClientHeight    =   1710
   ClientLeft      =   3840
   ClientTop       =   2190
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Top             =   975
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   660
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   285
   End
   Begin VB.CommandButton cmdJoint 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1245
      TabIndex        =   0
      Top             =   1350
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2175
      TabIndex        =   1
      Top             =   1350
      Width           =   915
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   930
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   330
      Width           =   2625
   End
   Begin VB.Label txtName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   2625
   End
End
Attribute VB_Name = "frmSBJoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Status As String
Private m_JointCust(3) As clsCustReg

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)
Me.cmdOk.Caption = LoadResString(gLangOffSet + 1)  '"«Œâ³ú"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)  '"ÇðÁðôà"
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

If Index = 0 Then Exit Sub

If m_JointCust(Index - 1) Is Nothing Then Exit Sub
If m_JointCust(Index - 1).CustomerID = 0 Then Exit Sub

If m_JointCust(Index) Is Nothing Then Set m_JointCust(Index) = New clsCustReg

m_JointCust(Index).ModuleID = wis_SBAcc
m_JointCust(Index).ShowDialog

'Check whether the selected customer is already  customer of this account
'If he is customer of this account then do not show him
Dim Count As Integer
For Count = 0 To 3
    If m_JointCust(Count) Is Nothing Then Exit For
    If Count <> Index Then
        If m_JointCust(Index).CustomerID = m_JointCust(Count).CustomerID Then
            m_JointCust(Index) = 0
            Exit Sub
        End If
    End If
Next

If m_JointCust(Index).CustomerID > 0 Then _
    txtName(Index) = m_JointCust(Index).FullName
    
End Sub

Private Sub cmdOk_Click()
Me.Status = "OK"
'if any customer is there in the joint list
'then save their details
Dim Count As Integer
For Count = 1 To 3
    If m_JointCust(0).CustomerID = 0 Then Exit For
    Call m_JointCust(0).SaveCustomer
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
For I = 0 To txtName.Count - 1
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
    JointCustId = m_JointCust(Index).CustomerID
End If

Exit Property


Dim I As Integer
Dim Str As String
For I = 0 To txtName.Count - 1
    If Trim$(txtName(I)) <> "" Then
        Str = Str & txtName(I).Tag & ";"
    End If
Next
'JointCustId = Str
End Property
Public Property Let JointCustId(Index As Integer, ByVal NewValue As Long)
On Error GoTo Err_line
If NewValue = 0 Then Exit Property

If m_JointCust(Index) Is Nothing Then _
    Set m_JointCust(Index) = New clsCustReg
    
m_JointCust(Index).ModuleID = wis_SBAcc
m_JointCust(Index).LoadCustomerInfo (NewValue)
txtName(Index) = m_JointCust(Index).FullName
Exit Property

' Breakup the string into array components.
Dim strTmp() As String
Dim I As Integer
'GetStringArray NewValue, strTmp(), ";"
'For i = 0 To UBound(strTmp)
'    If i > Text1.Count - 1 Then Exit For
'    Text1(i).Tag = strTmp(i)
'Next
'
Err_line:
End Property
Public Property Get JointHolders() As String
Dim I As Integer
Dim Str As String

For I = 0 To txtName.Count - 1
    If Trim$(txtName(I)) <> "" Then
        Str = Str & txtName(I) & ";"
    End If
Next
JointHolders = Str
End Property
Public Property Let JointHolders(ByVal vNewValue As String)
''On Error GoTo Err_Line
''If vNewValue = "" Then Exit Property
''
''' Breakup the string into array components.
''Dim strTmp() As String
''Dim i As Integer
''GetStringArray vNewValue, strTmp(), ";"
''For i = 0 To UBound(strTmp)
''    If i > txtName.Count - 1 Then Exit For
''    txtName(i).Text = strTmp(i)
''Next

Err_line:
End Property

