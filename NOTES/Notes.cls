VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NotesStruct
    Note As String
    Notetype As Byte
    NoteId As Byte
    Deleted As Boolean
End Type

Dim m_AccID As Long
Private M_ModuleID As Long
Private m_Notes() As NotesStruct
Private m_NoteCount As Byte

Private WithEvents m_frmNotes As frmNotes
Attribute m_frmNotes.VB_VarHelpID = -1

Public Property Get ObjectName()
    ObjectName = "NoteClass"
End Property


Public Function DeleteNotes(moduleID As wisModules, AccId As Long, NoteId As Long) As Boolean
On Error GoTo ErrLine

gDbTrans.SqlStmt = "DELETE * FROM NoteTab " & _
        " WHERE Accid = " & AccId & _
        " AND ModuleId =" & moduleID & _
        " AND NoteID = " & NoteId
        
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
   gDbTrans.RollBack
   GoTo ErrLine
End If
gDbTrans.CommitTrans
DeleteNotes = True

ErrLine:
End Function


Public Sub DisplayNote(txtRichTextBox As RichTextBox)
    
If txtRichTextBox Is Nothing Then Exit Sub
    
If M_ModuleID = 0 Then Exit Sub

Dim I As Integer

txtRichTextBox.Text = ""
Dim St As Long
txtRichTextBox.Visible = False
For I = 0 To UBound(m_Notes) - 1
    txtRichTextBox.Text = txtRichTextBox.Text & m_Notes(I).Note & vbCrLf
Next I
Dim TotalLen As Long
If gLangOffSet = wis_NoLangOffset Then
    txtRichTextBox.Font.name = "arial"
    txtRichTextBox.Font.Size = 10
Else
    txtRichTextBox.Font.name = gFontName
    txtRichTextBox.Font.Size = gFontSize

End If
If UBound(m_Notes) <= 0 Then
    txtRichTextBox.Text = GetResourceString(259) '"< No notes defined >"
    txtRichTextBox.Visible = True
    Exit Sub
End If


TotalLen = 0
For I = 0 To m_NoteCount - 1
    txtRichTextBox.SelStart = TotalLen
    txtRichTextBox.SelLength = Len(m_Notes(I).Note) + 2
    TotalLen = TotalLen + txtRichTextBox.SelLength
    If m_Notes(I).Notetype = 1 Then
        txtRichTextBox.SelColor = vbRed
    Else
        txtRichTextBox.SelColor = vbBlue
    End If
    txtRichTextBox.SelBullet = True
    txtRichTextBox.SelIndent = 100
    txtRichTextBox.SelHangingIndent = 200
Next I
txtRichTextBox.SelStart = 0
txtRichTextBox.Enabled = False

txtRichTextBox.Visible = True

End Sub
Public Property Get moduleID() As Long
    moduleID = M_ModuleID
End Property

Public Property Get AccId() As Long
AccId = m_AccID
End Property
Public Sub LoadNotes(ByVal moduleID As Long, ByVal AccId As Long)
Dim rst As ADODB.Recordset
    m_NoteCount = 0
    ReDim m_Notes(0)
    gDbTrans.SqlStmt = "Select * from NoteTab where ModuleID  = " & moduleID & _
                                " and Accid = " & AccId & " order by ModuleID"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        M_ModuleID = moduleID
        m_AccID = AccId
        Exit Sub
    End If
    M_ModuleID = moduleID
    m_AccID = AccId
    
    While Not rst.EOF
        
        m_Notes(m_NoteCount).Deleted = False
        m_Notes(m_NoteCount).Note = FormatField(rst("Notes"))
        m_Notes(m_NoteCount).NoteId = FormatField(rst("NoteID"))
        m_Notes(m_NoteCount).Notetype = FormatField(rst("NoteType"))
        
        m_NoteCount = m_NoteCount + 1
        ReDim Preserve m_Notes(m_NoteCount)
        
        rst.MoveNext
    Wend
    
End Sub

Public Property Get NoteCount() As Byte
    NoteCount = m_NoteCount
End Property

Public Sub Show()
    Dim I As Integer
    
    If M_ModuleID = 0 Or m_AccID = 0 Then
         Err.Raise 233, , "NameID not set or Note not loaded"
    End If
    

    
    Set m_frmNotes = New frmNotes
    Load m_frmNotes
    
    For I = 0 To UBound(m_Notes) - 1
        m_frmNotes.lstNote.AddItem m_Notes(I).Note
        m_frmNotes.lstNote.ItemData(m_frmNotes.lstNote.newIndex) = m_Notes(I).Notetype
    Next I
        
    m_frmNotes.Show vbModal
    Set m_frmNotes = Nothing
End Sub

Private Sub Class_Initialize()
ReDim m_Notes(0)
End Sub


Private Sub m_frmNotes_OKClicked()

'Update the Data base here from the form

'Start DB Transactions
gDbTrans.BeginTrans

'Delete all previous entries
gDbTrans.SqlStmt = "Delete from NoteTab " & _
        " where ModuleID  = " & M_ModuleID & _
        " And Accid = " & m_AccID
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If

'Insert new records into the data base
    Dim I As Integer
    Dim Notes As String
    Dim Notetype As Byte
    ReDim m_Notes(0)
    For I = 0 To m_frmNotes.lstNote.ListCount - 1
        Notes = m_frmNotes.lstNote.List(I)
        Notetype = m_frmNotes.lstNote.ItemData(I)
        gDbTrans.SqlStmt = "Insert Into NoteTab (ModuleID , " & _
                    " NoteID, Notes, NoteType, Accid) " & _
                    " values ( " & _
                    M_ModuleID & "," & _
                    I + 1 & "," & _
                    AddQuotes(Notes, True) & "," & _
                    Notetype & "," & m_AccID & ")"
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Sub
        End If
        'Update the notes struct
        m_Notes(I).Note = Notes
        m_Notes(I).NoteId = I
        m_Notes(I).Notetype = Notetype
        ReDim Preserve m_Notes(UBound(m_Notes) + 1)
        
    Next I
    
'COmmit transactions
    gDbTrans.CommitTrans

End Sub


