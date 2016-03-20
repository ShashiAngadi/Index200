VERSION 5.00
Begin VB.Form frmTransPort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Transportation Mode"
   ClientHeight    =   1245
   ClientLeft      =   2520
   ClientTop       =   4740
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4320
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1650
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2550
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtTransPort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1530
      TabIndex        =   1
      Top             =   90
      Width           =   2265
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3870
      TabIndex        =   2
      Top             =   90
      Width           =   405
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3450
      TabIndex        =   5
      Top             =   720
      Width           =   825
   End
   Begin VB.Label lblTransport 
      Caption         =   "Transport mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "frmTransPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_TransportID As Integer
Private m_TransportModeName As String

Private m_DBOperation As wis_DBOperation

Private Sub ClearControls()
txtTransPort.Text = ""

m_DBOperation = Insert
cmdOk.Caption = "&Ok"

cmdDelete.Enabled = False

End Sub

Private Function DeleteGroup() As Boolean

Dim Rst As ADODB.Recordset

On Error GoTo ErrLine:

DeleteGroup = False

gDbTrans.SqlStmt = " SELECT ProductID" & _
                   " FROM Products" & _
                   " WHERE GroupID = " & m_TransportID
                   
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    MsgBox "Group can not be deleted while product from this group exists", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If MsgBox("Do you want to Delete this group ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function

gDbTrans.SqlStmt = " DELETE FROM ProductGroup" & _
                   " WHERE GroupID = " & m_TransportID

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

MsgBox "Group Deleted", vbInformation
DeleteGroup = True
cmdDelete.Enabled = False

Exit Function

ErrLine:
    MsgBox "DeleteGroup" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function SaveDetails() As Boolean
'Declare the variables
Dim TransportModeName As String
Dim TransportID As Long
Dim Rst As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

TransportModeName = Trim$(txtTransPort.Text)

gDbTrans.SqlStmt = " SELECT MAX(TransModeID)" & _
                   " FROM TransportMode"

Call gDbTrans.Fetch(Rst, adOpenForwardOnly)

TransportID = FormatField(Rst.Fields(0)) + 1

'insert into the database
gDbTrans.SqlStmt = " INSERT INTO TransportMode (TransModeID,TransModeName)" & _
                  " VALUES ( " & _
                  TransportID & "," & _
                  AddQuotes(TransportModeName) & " )"

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE

SaveDetails = True

Exit Function

ErrLine:
    MsgBox "SaveDetails" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function UpdateDetails() As Boolean


'Setup an error handler...
On Error GoTo ErrLine

gDbTrans.SqlStmt = " UPDATE TransportMode" & _
                   " SET TransModeName = " & AddQuotes(txtTransPort.Text) & _
                   " WHERE TransModeID = " & m_TransportID

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
gDbTrans.CommitTrans

MsgBox "Details Updated", vbInformation, wis_MESSAGE_TITLE

UpdateDetails = True

Exit Function

ErrLine:
    MsgBox "UpDateDetails" & vbCrLf & Err.Description, vbCritical
    
End Function
Private Function Validated() As Boolean
Dim TransportModeName As String
Dim Rst As ADODB.Recordset

Validated = False

TransportModeName = Trim$(txtTransPort.Text)

If TransportModeName = "" Then Exit Function

Validated = True
If m_DBOperation = Update Then Exit Function

Validated = False

gDbTrans.SqlStmt = " SELECT *" & _
                   " FROM TransportMode" & _
                   " WHERE TransModeName = " & AddQuotes(TransportModeName)
                   
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
   MsgBox "Duplicate Entry", vbInformation, wis_MESSAGE_TITLE
   Exit Function
End If


Validated = True

End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
Call DeleteGroup
Call ClearControls
End Sub

Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

If m_DBOperation = Insert Then SaveDetails
If m_DBOperation = Update Then UpdateDetails
 
Call ClearControls

End Sub


Private Sub cmdSearch_Click()
Dim Rst As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT TransModeID,TransModeName" & _
                  " FROM TransPortMode" & _
                  " ORDER BY TransModeName"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

If Not FillView(m_frmLookUp.lvwReport, Rst, "TransModeID", True) Then Exit Sub

m_TransportID = 0
m_TransportModeName = ""
m_frmLookUp.Show vbModal

If m_TransportID = 0 Then Exit Sub

txtTransPort.Text = m_TransportModeName

m_DBOperation = Update
cmdOk.Caption = "&Update"

cmdDelete.Enabled = True
   


End Sub

Private Sub Form_Load()
'Center the form
CenterMe Me

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

m_DBOperation = Insert
cmdOk.Caption = "&Ok"

cmdDelete.Enabled = False


End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmTransPort = Nothing
End Sub


Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_TransportID = CLng(strSelection)
End Sub






Private Sub m_frmLookUp_SubItems(strSubItem() As String)
m_TransportModeName = strSubItem(0)
End Sub


