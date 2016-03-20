VERSION 5.00
Begin VB.Form frmCreateItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating Product or Item"
   ClientHeight    =   2310
   ClientLeft      =   1245
   ClientTop       =   2025
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CreateItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbGroup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   9
      Top             =   210
      Width           =   2805
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4200
      TabIndex        =   8
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2130
      TabIndex        =   3
      Top             =   1170
      Width           =   315
   End
   Begin VB.ComboBox cmbAliasName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   5
      Top             =   1170
      Width           =   2805
   End
   Begin VB.CheckBox chkHasAliasName 
      Caption         =   "Has any alias name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2490
      TabIndex        =   4
      Top             =   660
      Width           =   2775
   End
   Begin VB.TextBox txtProductName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   2
      Top             =   1170
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2820
      TabIndex        =   7
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1440
      TabIndex        =   6
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label lblGroup 
      Caption         =   "Select group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   2025
   End
   Begin VB.Label lblProductName 
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   2025
   End
End
Attribute VB_Name = "frmCreateItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_dbOperation  As wis_DBOperation
Private m_ProductId As Long

Private WithEvents m_frmLookUp  As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls
lblGroup.Caption = GetResourceString(157)
lblProductName.Caption = GetResourceString(158, 35)
chkHasAliasName.Caption = GetResourceString(159, 35)
cmdOk.Caption = GetResourceString(1)
cmdDelete.Caption = GetResourceString(14)
cmdCancel.Caption = GetResourceString(2)


End Sub

Private Function DeleteProducts() As Boolean
Dim intGroupID As Long
Dim rst As ADODB.Recordset

DeleteProducts = False

On Error GoTo ErrLine

intGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
'Check for the properties set for the product
gDbTrans.SqlStmt = " SELECT RelationID FROM RelationMaster " & _
                " WHERE GroupID = " & intGroupID & _
                " AND ProductID = " & m_ProductId

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    MsgBox GetResourceString(532) '"You can not delete while properties exist", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'Ask for the confirmation
'If MsgBox(GetResourceString(789) & vbCrLf & GetResourceString(541) & _
'             vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE = vbNo) Then Exit Function

gDbTrans.SqlStmt = " DELETE * FROM Products " & _
                 " WHERE GroupID = " & intGroupID & _
                 " AND ProductID = " & m_ProductId
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

DeleteProducts = True
m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"
cmbGroup.Enabled = True
MsgBox GetResourceString(909) '"Product Deleted ", vbInformation

Exit Function

ErrLine:
    MsgBox Err.Description, vbCritical
    
End Function

Private Sub SetFont()
On Error Resume Next

Dim Ctrl As Control

For Each Ctrl In Me
      Ctrl.FontName = "Arial"
      Ctrl.FONTSIZE = 10
Next Ctrl


End Sub


Private Sub ClearControls()

txtProductName.Text = ""

chkHasAliasName.Value = vbUnchecked
cmbAliasName.ListIndex = -1
chkHasAliasName.Enabled = True

On Error Resume Next
cmbGroup.SetFocus
End Sub

Private Sub LoadAliasNames()
Dim rstAlias As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT ProductID,ProductName " & _
                  " FROM Products " & _
                  " WHERE AliasID = " & 0


If gDbTrans.Fetch(rstAlias, adOpenForwardOnly) < 1 Then Exit Sub


cmbAliasName.Clear
cmbAliasName.AddItem ""
cmbAliasName.ItemData(cmbAliasName.newIndex) = 0

Do
   If rstAlias.EOF Then Exit Sub
   
   cmbAliasName.AddItem FormatField(rstAlias("ProductName"))
   cmbAliasName.ItemData(cmbAliasName.newIndex) = FormatField(rstAlias("ProductID"))
   
   rstAlias.MoveNext
Loop

End Sub

Private Function LoadPrdouctDetails() As Boolean
Dim rstProducts As ADODB.Recordset
Dim AliasID As Long
Dim GroupID As Integer


LoadPrdouctDetails = False

gDbTrans.SqlStmt = " SELECT * FROM Products " & _
                  " WHERE ProductID = " & m_ProductId


If gDbTrans.Fetch(rstProducts, adOpenStatic) < 1 Then Exit Function

txtProductName = FormatField(rstProducts("ProductName"))
AliasID = FormatField(rstProducts("AliasID"))
GroupID = FormatField(rstProducts("GroupID"))

cmbGroup.ItemData(cmbGroup.ListIndex) = GroupID


If AliasID > 0 Then
   cmbAliasName.ListIndex = AliasID
   cmbAliasName.Enabled = True
   chkHasAliasName.Enabled = True
   chkHasAliasName.Value = vbChecked
End If
  

LoadPrdouctDetails = True

m_dbOperation = Update
cmdOk.Caption = GetResourceString(171) '"&Update"
cmdDelete.Enabled = True
cmbGroup.Enabled = False
End Function

Private Function SaveProducts() As Boolean
Dim AliasID As Long
Dim ProductID As Long
Dim rstProducts As ADODB.Recordset
Dim GroupID As Integer

SaveProducts = False

On Error GoTo ErrLine


AliasID = 0
If cmbGroup.ListIndex = -1 Then Exit Function

If chkHasAliasName.Value = vbChecked Then
   If cmbAliasName.ListIndex = -1 Then Exit Function
   AliasID = cmbAliasName.ItemData(cmbAliasName.ListIndex)
End If

GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)



gDbTrans.SqlStmt = " SELECT MAX(ProductID) FROM Products"

ProductID = 0
   
If gDbTrans.Fetch(rstProducts, adOpenStatic) > 0 Then ProductID = FormatField(rstProducts(0)) + 1
      
gDbTrans.SqlStmt = " INSERT INTO Products (ProductID,GroupID,ProductName,AliasID) " & _
                " VALUES ( " & _
                ProductID & "," & _
                GroupID & ",'" & _
                Trim$(txtProductName.Text) & "'," & _
                AliasID & " ) "
                        
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

gDbTrans.CommitTrans

SaveProducts = True
Call ClearControls
cmdDelete.Enabled = False


Exit Function

ErrLine:
    MsgBox "SaveProducts" & vbCrLf & Err.Description, vbCritical
End Function
Private Function UpdateProducts() As Boolean
Dim AliasID As Long
Dim GroupID As Integer

UpdateProducts = False
AliasID = 0
If cmbGroup.ListIndex = -1 Then Exit Function

If chkHasAliasName.Value = vbChecked Then
   If cmbAliasName.ListIndex = -1 Then Exit Function
   AliasID = cmbAliasName.ItemData(cmbAliasName.ListIndex)
End If

GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)

gDbTrans.SqlStmt = " UPDATE Products SET " & _
                " ProductName = '" & Trim$(txtProductName.Text) & "'," & _
                " AliasID = " & AliasID & "," & _
                " GroupID = " & GroupID & _
                " WHERE ProductID = " & m_ProductId



gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

UpdateProducts = True
cmdDelete.Enabled = False
cmbGroup.Enabled = True
Call ClearControls

Exit Function

ErrLine:
    MsgBox "UpdateProducts" & vbCrLf & Err.Description, vbCritical

End Function
Private Function Validated() As Boolean
Dim strProductName As String
Dim rst As ADODB.Recordset
Dim GroupID As Integer

Validated = False

If cmbGroup.ListIndex = -1 Then Exit Function

GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)

strProductName = Trim$(txtProductName.Text)

If strProductName = "" Then Exit Function
If m_dbOperation = Insert Then

    gDbTrans.SqlStmt = " SELECT * FROM Products " & _
                       " WHERE GroupID = " & GroupID & _
                       " AND ProductName = '" & strProductName & "'"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
       'MsgBox "Duplicate Entry", vbInformation, wis_MESSAGE_TITLE
       MsgBox GetResourceString(607), vbInformation, wis_MESSAGE_TITLE
       Exit Function
    End If
End If
Validated = True

End Function

Private Sub chkHasAliasName_Click()

cmbAliasName.Enabled = False
If chkHasAliasName.Value = vbChecked Then
   cmbAliasName.Enabled = True
   LoadAliasNames
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
If DeleteProducts Then ClearControls
End Sub

Private Sub cmdOk_Click()

If Not Validated Then Exit Sub
If m_dbOperation = Insert Then
    If Not SaveProducts Then Exit Sub
ElseIf m_dbOperation = Update Then
 If Not UpdateProducts Then Exit Sub
End If
'MsgBox "Saved the details", vbInformation, wis_MESSAGE_TITLE
MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE
End Sub

Private Sub cmdSearch_Click()
Dim rstProducts As ADODB.Recordset
Dim intGroupID As Integer

If cmbGroup.ListIndex = -1 Then Exit Sub

intGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)


gDbTrans.SqlStmt = " SELECT ProductID,ProductName " & _
                   " FROM  Products " & _
                   " WHERE GroupID = " & intGroupID & _
                   " ORDER BY ProductName"
                      

If gDbTrans.Fetch(rstProducts, adOpenForwardOnly) < 1 Then Exit Sub

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

'If Not FillView(m_frmLookUp.lvwReport, rstProducts, "ProductID", True) Then Exit Sub
If Not FillView(m_frmLookUp.lvwReport, rstProducts, True) Then Exit Sub

m_ProductId = 0
m_frmLookUp.Show vbModal

If m_ProductId > 0 Then Call LoadPrdouctDetails


End Sub


Private Sub Form_Load()
'Declare the variables
Dim rstProducts As ADODB.Recordset

'Cnter the form
CenterMe Me

'set the icon
Me.Icon = LoadResPicture(147, vbResIcon)

'set kannada captions
SetKannadaCaption

m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"

chkHasAliasName.Enabled = False
cmbAliasName.Enabled = False
cmdDelete.Enabled = False
cmbGroup.Enabled = True
'Load the product groups
LoadProductGroups


gDbTrans.SqlStmt = " SELECT * FROM Products"

If gDbTrans.Fetch(rstProducts, adOpenForwardOnly) < 1 Then Exit Sub

Set rstProducts = Nothing

'Load alias names to the alis combobox
LoadAliasNames

chkHasAliasName.Enabled = True


End Sub


Private Sub LoadProductGroups()
Dim rstGroups As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT GroupID,GroupName FROM ProductGroup " & _
                   " ORDER BY GroupID "
        
Call gDbTrans.Fetch(rstGroups, adOpenForwardOnly)

cmbGroup.Clear

Do While Not rstGroups.EOF
   
   cmbGroup.AddItem FormatField(rstGroups.Fields("GroupName"))
   cmbGroup.ItemData(cmbGroup.newIndex) = FormatField(rstGroups.Fields("GroupID"))
   
   'Move to next record
   rstGroups.MoveNext
Loop


End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCreateItem = Nothing
End Sub


Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_ProductId = CLng(strSelection)
End Sub

Private Sub txtProductName_LostFocus()
'txtProductName.Text = ConvertToProperCase(txtProductName.Text)
End Sub


