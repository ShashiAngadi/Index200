VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDepSelect 
   Caption         =   "Deposit Lst"
   ClientHeight    =   3300
   ClientLeft      =   2865
   ClientTop       =   1860
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   1035
   End
   Begin ComctlLib.ListView lstDeposits 
      Height          =   2625
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4630
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDepSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkClicked(AccList() As String, DepList() As Integer, TotalBalance As Currency)

Public Event CancelClicked()
Public Function LoadDeposits(CustomerID As Long)
Dim rst As Recordset
Dim SqlStr As String
SqlStr = "Select DepositType,A.AccID,AccNum,Balance " & _
    " FROM FDMAster A,FDTrans B Where B.AccID = A.AccID" & _
    " And B.TransID = (Select Max(TransID) From FDTrans C " & _
        " WHERE C.AccID = A.AccID ) " & _
    "AND Balance > 0  ANd A.CustomerID = " & CustomerID
SqlStr = SqlStr & " UNION " & _
    "Select " & wisDeposit_RD & " as DepositType,A.AccID,A.AccNum,Balance " & _
    " FROM RDMAster A,RDTrans B Where B.AccID = A.AccID" & _
    " And B.TransID = (Select Max(TransID) From RDTrans C " & _
        " WHERE C.AccID = A.AccID ) " & _
    "AND Balance > 0  ANd A.CustomerID = " & CustomerID

SqlStr = SqlStr & " UNION " & _
    "Select " & wisDeposit_PD & " as DepositType,A.AccID,A.AccNUm, Balance " & _
    " FROM PDMAster A,PDTrans B Where B.AccID = A.AccID" & _
    " And B.TransID = (Select Max(TransID) From PDTrans C " & _
        " WHERE C.AccID = A.AccID ) " & _
    "AND Balance > 0  ANd A.CustomerID = " & CustomerID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then GoTo Exit_Line
Dim lstItem As ListItem
With lstDeposits
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , "AccNum", GetResourceString(36, 60), 300
    .ColumnHeaders.Add , "AccID", , 0
    .ColumnHeaders.Add , "DepType", GetResourceString(43), 1200
    .ColumnHeaders.Add , "Balance", GetResourceString(42), 1500
    
End With
Dim DepCount As Integer
'Call FillView(Me.lstDeposits, Rst, False)

While Not rst.EOF
    With lstDeposits
       DepCount = DepCount + 1
       Set lstItem = .ListItems.Add(, "Dep" & DepCount, FormatField(rst("AccNum")))
       'Set lstItem = .ListItems.Add(, "Dep" & DepCount, FormatField(rst("AccID")))
       lstItem.SubItems(1) = FormatField(rst("AccNum"))
       lstItem.SubItems(2) = GetDepositTypeText(FormatField(rst("DepositType")))
       lstItem.SubItems(3) = FormatField(rst("Balance"))
       'If DepCount Mod 3 = 0 Then .ListItems(DepCount).Selected = True
    End With
    rst.MoveNext
Wend
lstDeposits.MultiSelect = True
lstDeposits.view = lvwReport

Exit_Line:



End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
cmdOk.Caption = GetResourceString(1) 'ok
cmdCancel.Caption = GetResourceString(2) 'Cancel
End Sub



Private Sub cmdCancel_Click()

RaiseEvent CancelClicked
Me.Hide

End Sub

Private Sub cmdOk_Click()

Dim Balance As Currency
Dim ArrCount As Integer

Dim count As Integer
Dim MaxCount As Integer

Dim AccList() As String
'Dim AccNum As String
Dim DepList() As Integer

ReDim AccList(0)
ReDim DepList(0)

Dim strAccNum As String

MaxCount = lstDeposits.ListItems.count

strAccNum = ""
For count = 1 To MaxCount
    If lstDeposits.ListItems(count).Selected Then
        ReDim Preserve AccList(ArrCount)
        ReDim Preserve DepList(ArrCount)
       ' AccList(ArrCount) = Val(lstDeposits.ListItems(Count))
        AccList(ArrCount) = lstDeposits.ListItems(count).SubItems(1)
        DepList(ArrCount) = GetDepositType(lstDeposits.ListItems(count).SubItems(2))
        Balance = Balance + Val(lstDeposits.ListItems(count).SubItems(3))
        ArrCount = ArrCount + 1
    End If
Next

If UBound(AccList) = 0 And AccList(0) = "" Then
    MsgBox "You have not selected any deposit", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Now remove the first delimeter
strAccNum = Mid(strAccNum, 2)
'strDep = Mid(strDep, 2)

RaiseEvent OkClicked(AccList(), DepList(), Balance)
Me.Hide

End Sub

Private Sub cmdOLDOK_Click()

Dim Balance As Currency
Dim ArrCount As Integer

Dim strDep As String
Dim strAccNum As String
Dim count As Integer
Dim MaxCount As Integer

Dim AccList() As Long
Dim DepList() As Integer

MaxCount = lstDeposits.ListItems.count

strAccNum = ""
For count = 1 To MaxCount
    If lstDeposits.ListItems(count).Selected Then
        ReDim Preserve AccList(ArrCount)
        ReDim Preserve DepList(ArrCount)
        AccList(ArrCount) = lstDeposits.ListItems(count)
        DepList(ArrCount) = GetDepositType(lstDeposits.ListItems(count).SubItems(2))
        'DepName(ArrCount) = GetDepositInteger(lstDeposits.ListItems(Count).SubItems(1))
        Balance = Balance + Val(lstDeposits.ListItems(count).SubItems(3))
        strDep = strDep & gDelim & lstDeposits.ListItems(count).SubItems(1)
        strAccNum = strAccNum & gDelim & lstDeposits.ListItems(count)
        ArrCount = ArrCount + 1
    End If
Next

If strAccNum = "" Then
    MsgBox "You have not selected any deposit", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

'Now remove the first delimeter
strAccNum = Mid(strAccNum, 2)
strDep = Mid(strDep, 2)

'RaiseEvent OKClicked(strAccNum(), strDep(), Balance)

Me.Hide

End Sub


Private Sub Form_Load()
Call SetKannadaCaption
'Call LoadDeposits(3970)

End Sub

