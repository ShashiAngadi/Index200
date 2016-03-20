VERSION 5.00
Begin VB.Form frmProductGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Group Creation"
   ClientHeight    =   1605
   ClientLeft      =   3675
   ClientTop       =   2925
   ClientWidth     =   4410
   Icon            =   "ProductGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4410
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   435
      Left            =   3390
      TabIndex        =   5
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   345
      Left            =   3900
      TabIndex        =   2
      Top             =   210
      Width           =   375
   End
   Begin VB.TextBox txtGroupName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   210
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2400
      TabIndex        =   4
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   1410
      TabIndex        =   3
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblGroupName 
      Caption         =   "Group Name"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmProductGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_GroupID As Integer

Private m_DBOperation As wis_DBOperation

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls

lblGroupName.Caption = LoadResString(gLangOffSet + 157) & " " & LoadResString(gLangOffSet + 35)
cmdOk.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
cmdDelete.Caption = LoadResString(gLangOffSet + 14)
End Sub


Private Sub ClearControls()
txtGroupName.Text = ""

cmdOk.Caption = LoadResString(gLangOffSet + 1) '"&Ok"
m_DBOperation = Insert
End Sub

Private Function DeleteGroup() As Boolean
Dim Rst As ADODB.Recordset

On Error GoTo ErrLine

DeleteGroup = False
gDbTrans.SQLStmt = " SELECT ProductName FROM Products " & _
                   " WHERE GroupID = " & m_GroupID
                   
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    'MsgBox "Group can not be deleted while product from this group exists", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 772)
    Exit Function
End If

If MsgBox("Do you want to Delete this group ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function

gDbTrans.SQLStmt = " DELETE FROM ProductGroup " & _
                   " WHERE GroupID = " & m_GroupID


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

'MsgBox "Group Deleted", vbInformation
MsgBox LoadResString(gLangOffSet + 677), vbInformation
DeleteGroup = True
cmdDelete.Enabled = False

Exit Function

ErrLine:
    MsgBox "DeleteGroup" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Sub LoadProductGroup()

Dim rstGroup As ADODB.Recordset

gDbTrans.SQLStmt = " SELECT * FROM ProductGroup " & _
                  " WHERE GroupID = " & m_GroupID
Call gDbTrans.Fetch(rstGroup, adOpenForwardOnly)

txtGroupName.Text = FormatField(rstGroup.Fields("GroupName"))


End Sub

Private Function SaveGroupName() As Boolean
'Declare the variables
Dim strGroupName As String
Dim lngGroupID As Long
Dim Rst As ADODB.Recordset

On Error GoTo ErrLine

SaveGroupName = False

strGroupName = Trim$(txtGroupName.Text)

gDbTrans.SQLStmt = " SELECT MAX(GroupID) FROM ProductGroup"

Call gDbTrans.Fetch(Rst, adOpenForwardOnly)
lngGroupID = FormatField(Rst.Fields(0)) + 1

'insert into the database
gDbTrans.SQLStmt = " INSERT INTO ProductGroup (GroupID,GroupName) " & _
                  " VALUES ( " & _
                  lngGroupID & ",'" & _
                  strGroupName & "' ) "


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 528), vbInformation, wis_MESSAGE_TITLE
SaveGroupName = True

Exit Function

ErrLine:
    MsgBox "SaveGroupName" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function UpDateGroupName() As Boolean

On Error GoTo ErrLine

UpDateGroupName = False

gDbTrans.SQLStmt = " UPDATE ProductGroup " & _
                   " SET GroupName = " & AddQuotes(txtGroupName.Text) & _
                   " WHERE GroupID = " & m_GroupID


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 707), vbInformation, wis_MESSAGE_TITLE

UpDateGroupName = True

Exit Function

ErrLine:
    MsgBox "UpdateGroupName" & vbCrLf & Err.Description, vbCritical
    
End Function


Private Function Validated() As Boolean
Dim strGroupName As String
Dim Rst As ADODB.Recordset

Validated = False


strGroupName = Trim$(txtGroupName.Text)


If strGroupName = "" Then Exit Function

If m_DBOperation = Insert Then

    gDbTrans.SQLStmt = " SELECT * FROM ProductGroup " & _
                       " WHERE GroupName = '" & strGroupName & "'"
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
       'MsgBox "Duplicate Entry", vbInformation, wis_MESSAGE_TITLE
       MsgBox LoadResString(gLangOffSet + 607)
       Exit Function
    End If
End If
Validated = True
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
If DeleteGroup Then ClearControls
End Sub

Private Sub cmdOk_Click()
If Not Validated Then Exit Sub

If m_DBOperation = Insert Then
    If SaveGroupName Then ClearControls
ElseIf m_DBOperation = Update Then
    If UpDateGroupName Then ClearControls
End If

End Sub


Private Sub cmdSearch_Click()
Dim Rst As ADODB.Recordset

gDbTrans.SQLStmt = " SELECT GroupID,GroupName " & _
                  " FROM ProductGroup " & _
                  " ORDER BY GroupID"

If gDbTrans.Fetch(Rst, adOpenStatic) < 1 Then Exit Sub

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

'If Not FillView(m_frmLookUp.lvwReport, Rst,  "GroupID", True) Then Exit Sub
If Not FillView(m_frmLookUp.lvwReport, Rst, True) Then Exit Sub

m_GroupID = 0

m_frmLookUp.Show vbModal

If m_GroupID > 0 Then
   LoadProductGroup
   m_DBOperation = Update
   cmdOk.Caption = LoadResString(gLangOffSet + 171) '"&Update"
   cmdDelete.Enabled = True
End If

End Sub

Private Sub Form_Load()
CenterMe Me


m_DBOperation = Insert
cmdOk.Caption = LoadResString(gLangOffSet + 1) '"&Ok"
cmdDelete.Enabled = False

'set icon for the form caption
Me.Icon = LoadResPicture(147, vbResIcon)
  
Call SetKannadaCaption

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmProductGroup = Nothing
End Sub


Private Sub Image1_Click()

End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_GroupID = CLng(strSelection)
End Sub


Private Sub txtGroupName_LostFocus()
'txtGroupName.Text = ConvertToProperCase(txtGroupName.Text)
End Sub


