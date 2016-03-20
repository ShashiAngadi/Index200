VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmReceive As frmReceivable
Attribute m_frmReceive.VB_VarHelpID = -1
Private m_AllTrans As AllTrans
Private m_RetCount As Integer
Private m_TotalAmount As Currency
Private m_strKeys


Public Sub AddHeadAndAmount(ByVal AccHeadID As Long, ByVal Amount As Currency)
If m_frmReceive Is Nothing Then _
    Set m_frmReceive = New frmReceivable
If m_AllTrans Is Nothing Then Set m_AllTrans = New AllTrans

'FIrst Load the AMountothte heads
With m_frmReceive
    Dim I As Integer
    
    'First Select the Head
    With .cmbHead
        Do
            If I = .ListCount Then Exit Do
            If AccHeadID = .ItemData(I) Then .ListIndex = I: Exit Do
            I = I + 1
        Loop
    End With
    .txtAmount = Amount
End With

'Now Add this amount to the Grids
Call AddToGrid(AccHeadID, Amount)

End Sub

Private Sub AddToGrid(HeadID As Long, Amount As Currency)
'Now Check for the existence of the same head then
If m_AllTrans.Exists("Head" & HeadID) Then
    'MsgBox "This head Already exists"
    MsgBox GetResourceString(545), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If m_AllTrans.Add(HeadID, 0, Amount, GetParentID(HeadID), "Head" & HeadID) Is Nothing Then
    'msgbox "Unable to perform this operatin"
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Now Get the row, where these details are to be entered
With m_frmReceive.grd
    .Col = 0: .Row = 0
    Do
        .Row = .Row + 1
        If Val(.Text) = 0 Then Exit Do
        If .Row = .Rows - 1 Then .Rows = .Rows + 1
    Loop
    .RowData(.Row) = HeadID
    .Col = 0: .Text = .Row
    .Col = 1: .Text = m_frmReceive.cmbHead.Text
    .Col = 2: .Text = m_frmReceive.txtAmount
End With

m_TotalAmount = m_TotalAmount + Amount
'Now clear the controls of the form
Call ClearForm


End Sub


Public Property Let IsAddToAccount(NewValue As Boolean)
'If m_AllTrans Is Nothing Then Exit Property

If m_frmReceive Is Nothing Then Exit Property
m_frmReceive.chkAdd.Value = NewValue

End Property


Public Property Get IsAddToAccount() As Boolean
If m_AllTrans Is Nothing Then Exit Property
If m_frmReceive Is Nothing Then Exit Property

IsAddToAccount = m_frmReceive.chkAdd.Value

End Property



Private Sub ClearForm()
If m_frmReceive Is Nothing Then Exit Sub

With m_frmReceive
    .cmdAccept.Caption = GetResourceString(4)
    
    .cmbHead.ListIndex = -1
    .cmbHead.Locked = False
    .txtAmount.Value = 0
End With

End Sub

Public Sub GetHeadAndAmount(HeadID As Long, DrAmount As Currency, CrAmount As Currency)
HeadID = 0: DrAmount = 0: CrAmount = 0
If m_AllTrans Is Nothing Then Exit Sub
If m_AllTrans.count = 0 Then Exit Sub
m_RetCount = 0
m_strKeys = m_AllTrans.GetKeys

HeadID = m_AllTrans(m_strKeys(m_RetCount)).HeadID

CrAmount = m_AllTrans(m_strKeys(m_RetCount)).CrAmount
DrAmount = m_AllTrans(m_strKeys(m_RetCount)).DrAmount

m_RetCount = m_RetCount + 1

End Sub
Public Sub NextHeadAndAmount(HeadID As Long, DrAmount As Currency, CrAmount As Currency)
HeadID = 0: DrAmount = 0: CrAmount = 0
'if m_RetCount =0
If m_AllTrans Is Nothing Then Exit Sub
If m_RetCount = m_AllTrans.count Then Exit Sub

If m_RetCount = 0 Then
    Call GetHeadAndAmount(HeadID, DrAmount, CrAmount)
    Exit Sub
End If


HeadID = m_AllTrans(m_strKeys(m_RetCount)).HeadID
CrAmount = m_AllTrans(m_strKeys(m_RetCount)).CrAmount
DrAmount = m_AllTrans(m_strKeys(m_RetCount)).DrAmount
'Now increase the count
m_RetCount = m_RetCount + 1

End Sub


Public Sub Show()

If m_frmReceive Is Nothing Then
    Set m_frmReceive = New frmReceivable
    Load m_frmReceive
End If

'First check the whether heads alreay loaded or not
If Not m_AllTrans Is Nothing Then
'    Unload m_frmReceive
'    Load m_frmReceive
    With m_frmReceive.grd
        .Rows = m_AllTrans.count + 1
        If .Rows < 7 Then .Rows = 7
        .Row = 0
    End With
End If

m_frmReceive.Show vbModal


End Sub

Public Property Get TotalAmount() As Currency
TotalAmount = m_TotalAmount
End Property

Private Sub Class_Terminate()

If Not m_frmReceive Is Nothing Then _
            Set m_frmReceive = Nothing
            
End Sub


Private Sub m_frmReceive_AcceptClicked()
If m_AllTrans Is Nothing Then _
                Set m_AllTrans = New AllTrans
    
Dim HeadID As Long
Dim Amount As Currency

With m_frmReceive
    HeadID = .cmbHead.ItemData(.cmbHead.ListIndex)
    Amount = .txtAmount.Value
End With

Call AddToGrid(HeadID, Amount)
Exit Sub

'Now Check for the existence of the same head then
If m_AllTrans.Exists("Head" & HeadID) Then
    'MsgBox "This head Already exists"
    MsgBox GetResourceString(545), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
If m_AllTrans.Add(HeadID, 0, Amount, GetParentID(HeadID), "Head" & HeadID) Is Nothing Then
    'msgbox "Unable to perform this operatin"
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Now Get the row, where these details are to be entered
With m_frmReceive.grd
    .Col = 0: .Row = 0
    Do
        .Row = .Row + 1
        If Val(.Text) = 0 Then Exit Do
        If .Row = .Rows - 1 Then .Rows = .Rows + 1
    Loop
    .RowData(.Row) = HeadID
    .Col = 0: .Text = .Row
    .Col = 1: .Text = m_frmReceive.cmbHead.Text
    .Col = 2: .Text = m_frmReceive.txtAmount
End With

m_TotalAmount = m_TotalAmount + Amount
'Now clear the controls of the form
Call ClearForm

End Sub
Private Sub m_frmReceive_CancelClicked()
If Not m_AllTrans Is Nothing Then _
                Set m_AllTrans = Nothing
                
m_TotalAmount = 0
Unload m_frmReceive

End Sub


Private Sub m_frmReceive_GridClicked()
Dim HeadID As Long
Dim strKey As String

With m_frmReceive.grd
    HeadID = .RowData(.Row)
    If HeadID = 0 Then GoTo Exit_line
    strKey = "Head" & HeadID
End With

'Now Search the same ead in the combo box
With m_frmReceive.cmbHead
    .Locked = False
    Dim I As Integer, MaxI As Integer
    MaxI = .ListCount - 1
    For I = 0 To MaxI
        If .ItemData(I) = HeadID Then
            .ListIndex = I
            .Locked = True
            Exit For
        End If
    Next
End With

With m_frmReceive
    If Not .cmbHead.Locked Then GoTo Exit_line
    .cmdAccept = GetResourceString(171)
    .grd.Col = .grd.Cols - 1
    .txtAmount = Val(.grd.Text)
End With


Exit_line:

End Sub

Private Sub m_frmReceive_OkClicked()

If m_AllTrans Is Nothing Then Unload m_frmReceive: Exit Sub
If m_AllTrans.count = 0 Then Set m_AllTrans = Nothing: Unload m_frmReceive: Exit Sub
'Unload m_frmReceive

End Sub

Private Sub m_frmReceive_UpdateClicked()

On Error GoTo Exit_line

If m_AllTrans Is Nothing Then GoTo Exit_line
Dim strKey As String
Dim HeadID As Long
Dim PrevAmount As Currency
Dim Amount As Currency

With m_frmReceive
  With .cmbHead
    HeadID = .ItemData(.ListIndex)
    If HeadID = 0 Then GoTo Exit_line
    strKey = "Head" & HeadID
  End With
  Amount = .txtAmount
End With

If Not m_AllTrans.Exists(strKey) Then GoTo Exit_line

PrevAmount = m_AllTrans(strKey).CrAmount

If m_frmReceive.txtAmount = 0 Then
    m_AllTrans.Remove (strKey)
Else
    If m_AllTrans.Update(HeadID, 0, Amount, _
                GetParentID(HeadID), HeadID, strKey) Is Nothing Then GoTo Exit_line
End If

'Now Get the row, where these details are to be entered
With m_frmReceive.grd
    .Col = 0: .Row = 0
    Do
        .Row = .Row + 1
        If .RowData(.Row) = HeadID Then Exit Do
        If .Row = .Rows - 1 Then Exit Do
    Loop
    If m_frmReceive.txtAmount = 0 Then
        .RemoveItem (.Row)
        GoTo LastLine
    End If
    .RowData(.Row) = HeadID
    .Col = 0: .Text = .Row
    .Col = 1: .Text = m_frmReceive.cmbHead.Text
    .Col = 2: .Text = m_frmReceive.txtAmount
End With

LastLine:

m_TotalAmount = m_TotalAmount - PrevAmount + Amount

Call ClearForm

Exit_line:
    
    MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE

End Sub


Private Sub m_frmReceive_WindowClosed()
Set m_frmReceive = Nothing

End Sub


