VERSION 5.00
Begin VB.Form frmUnits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units Creation"
   ClientHeight    =   1185
   ClientLeft      =   1770
   ClientTop       =   2910
   ClientWidth     =   4275
   Icon            =   "Units.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4275
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3390
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   345
      Left            =   3900
      TabIndex        =   2
      Top             =   30
      Width           =   375
   End
   Begin VB.TextBox txtUnitName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   30
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1650
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblUnitName 
      Caption         =   "Unit Name"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_UnitID As Integer

Private m_DBOperation As wis_DBOperation

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)


'set the Kannada for all controls
lblUnitName.Caption = LoadResString(gLangOffSet + 161) & " " & LoadResString(gLangOffSet + 35)
cmdOk.Caption = LoadResString(gLangOffSet + 1)
cmdCancel.Caption = LoadResString(gLangOffSet + 2)
cmdDelete.Caption = LoadResString(gLangOffSet + 14)
End Sub


Private Sub ClearControls()
txtUnitName.Text = ""
'txtSymbol.Text = ""

ActivateTextBox txtUnitName

End Sub

Private Function DeleteUnit() As Boolean

DeleteUnit = False

On Error GoTo ErrLine

If MsgBox("Do you want to Delete this Unit ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function

gDbTrans.SQLStmt = " DELETE FROM Units " & _
                   " WHERE UnitID = " & m_UnitID


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

'MsgBox "Deleted the Unit", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 678), vbInformation, wis_MESSAGE_TITLE
DeleteUnit = True
cmdDelete.Enabled = False
cmdOk.Caption = "&Ok"
m_DBOperation = Insert

Exit Function

ErrLine:
    MsgBox "DeleteUnit" & vbCrLf & Err.Description, vbCritical


End Function

Private Sub LoadUnits()

Dim rstUnits As ADODB.Recordset

gDbTrans.SQLStmt = " SELECT * FROM Units " & _
                  " WHERE UnitID = " & m_UnitID
Call gDbTrans.Fetch(rstUnits, adOpenForwardOnly)

txtUnitName.Text = FormatField(rstUnits.Fields("UnitName"))
'txtSymbol.Text = FormatField(rstUnits.Fields("UnitName"))

End Sub

Private Function SaveUnitName() As Boolean
'Declare the variables
Dim strUnitName As String
Dim strSymbol As String
Dim lngUnitID As Integer
Dim Rst As ADODB.Recordset

SaveUnitName = False

On Error GoTo ErrLine

strUnitName = Trim$(txtUnitName.Text)

gDbTrans.SQLStmt = " SELECT MAX(UnitID) FROM Units"

Call gDbTrans.Fetch(Rst, adOpenForwardOnly)
lngUnitID = FormatField(Rst.Fields(0)) + 1

'insert into the database
gDbTrans.SQLStmt = " INSERT INTO Units (UnitID,UnitName) " & _
                  " VALUES ( " & _
                  lngUnitID & ",'" & _
                  strUnitName & "' ) "


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 528), vbInformation, wis_MESSAGE_TITLE
SaveUnitName = True

Exit Function

ErrLine:
    MsgBox "SaveUnitName" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function UpDateUnitName() As Boolean

UpDateUnitName = False

On Error GoTo ErrLine

If txtUnitName = "" Then Exit Function


gDbTrans.SQLStmt = " UPDATE Units " & _
                   " SET UnitName = " & AddQuotes(txtUnitName.Text) & _
                   " WHERE UnitID = " & m_UnitID


gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
MsgBox LoadResString(gLangOffSet + 707), vbInformation, wis_MESSAGE_TITLE
cmdOk.Caption = "&ok"
m_DBOperation = Insert

UpDateUnitName = True

Exit Function

ErrLine:
    MsgBox "UpdateUnitName" & vbCrLf & Err.Description, vbCritical
    
End Function



Private Function Validated() As Boolean
Dim strUnitName As String
Dim Rst As ADODB.Recordset

Validated = False

strUnitName = Trim$(txtUnitName.Text)


If m_DBOperation = Update Then
    Validated = True
    Exit Function
End If

If strUnitName = "" Then Exit Function

gDbTrans.SQLStmt = " SELECT * FROM Units " & _
                   " WHERE UnitName = '" & strUnitName & "'"
                   
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
   'MsgBox "Duplicate Entry", vbInformation, wis_MESSAGE_TITLE
   MsgBox LoadResString(gLangOffSet + 607), vbInformation, wis_MESSAGE_TITLE
   Exit Function
End If

Validated = True
End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
If DeleteUnit Then ClearControls
End Sub

Private Sub cmdOk_Click()
If Validated Then
   If m_DBOperation = Insert Then
      If SaveUnitName Then ClearControls
   ElseIf m_DBOperation = Update Then
      If UpDateUnitName Then ClearControls
   End If
End If
End Sub


Private Sub cmdSearch_Click()
Dim Rst As ADODB.Recordset

gDbTrans.SQLStmt = " SELECT UnitID,UnitName " & _
                  " FROM Units " & _
                  " ORDER BY UnitID"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

'If Not FillView(m_frmLookUp.lvwReport, Rst, "UnitID", True) Then Exit Sub
If Not FillView(m_frmLookUp.lvwReport, Rst, True) Then Exit Sub

m_UnitID = 0

m_frmLookUp.Show vbModal

If m_UnitID > 0 Then
   
   LoadUnits
   
   m_DBOperation = Update
   cmdOk.Caption = "&Update"
   cmdDelete.Enabled = True
End If

End Sub

Private Sub Form_Load()
CenterMe Me

m_DBOperation = Insert
cmdOk.Caption = "&Ok"
cmdDelete.Enabled = False

'set icon for the form caption
Me.Icon = LoadResPicture(147, vbResIcon)
  
Call SetKannadaCaption




End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmProductGroup = Nothing
End Sub


Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_UnitID = CLng(strSelection)
End Sub




Private Sub txtUnitName_LostFocus()
'txtUnitName.Text = ConvertToProperCase(txtUnitName.Text)
End Sub


