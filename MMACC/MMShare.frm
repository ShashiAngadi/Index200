VERSION 5.00
Begin VB.Form frmMMShare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques"
   ClientHeight    =   3225
   ClientLeft      =   2700
   ClientTop       =   2790
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert Selection"
      Height          =   495
      Left            =   210
      TabIndex        =   3
      Top             =   2490
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3210
      TabIndex        =   4
      Top             =   2490
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Frame fraPurchase 
      Caption         =   "Issue new Shares"
      Height          =   2325
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   4245
      Begin VB.TextBox txtPrefix 
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtLeaves 
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtStartNo 
         Height          =   375
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblPrefix 
         Caption         =   "Cert Prefix:"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lblLeaves 
         Caption         =   "No of leaves:"
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   1650
         Width           =   1755
      End
      Begin VB.Label lblStartNo 
         Caption         =   "Start no:"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   990
         Width           =   1785
      End
   End
   Begin VB.Frame fraSale 
      Caption         =   "Current leaf set"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   4245
      Begin VB.ListBox lstCheque 
         Height          =   1860
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   270
         Width           =   3865
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   465
      Left            =   1680
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMMShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Event LeavesPurchased(strConstant As String, StartNo As Long, Leaves As Long, Cancel As Boolean)
'Public Event LeavesReturned(Leaves() As String)
Public Event ShareIssued(ShareNos() As String, Cancel As Boolean)
Public Event ShareReturned(Leaves() As String)


'
Private Sub cmdClose_Click()
Unload Me
End Sub


'
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

fraPurchase.Caption = GetResourceString(143)
lblStartNo = GetResourceString(144)
lblLeaves = GetResourceString(145)
cmdInvert.Caption = GetResourceString(21)
cmdAdd.Caption = GetResourceString(10)
cmdCancel.Caption = GetResourceString(11)

End Sub

'
Private Sub cmdAdd_Click()

Dim LeafCount As Long
Dim strLeft As String
Dim strRight As String
Dim strTemp As String
Dim strNo  As String
Dim count As Integer
Dim FstShareNo As Long
Dim StrShare() As String

'Validate the Cheque Start No
strLeft = Trim(txtPrefix)
If Len(strLeft) > 0 Then
    If IsNumeric(strNo) Then
        MsgBox "Prefix has to be alphabet", vbOKOnly, gAppName & " - Info"
        ActivateTextBox txtPrefix
        Exit Sub
    End If
End If
strNo = Trim(txtStartNo)
If Not IsNumeric(strNo) Then
    MsgBox GetResourceString(504), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtStartNo
    Exit Sub
    'There may be a case that share certificate may contain
    'String's 'In such case consider the string
    'so saperate the String From The Loan
    'Get the left part
'    Count = 1
    'Do
'        If Count = Len(strNo) + 1 Then Exit Do
 '       If IsNumeric(Mid(strNo, Count)) Then
  '          strLeft = Left(strNo, Count - 1)
   '         strNo = Mid(strNo, Count)
    '        Exit Do
     '   End If
      '  Count = Count + 1
    'Loop
    
    'Get the right part
    'First reverse the string
'    Count = 1
'    strTemp = strReverse(strNo)
 '   Do
  '      If Count = Len(strTemp) + 1 Then Exit Do
   '     If IsNumeric(Mid(strTemp, Count)) Then
    '        strRight = Left(strTemp, Count - 1)
     '       strTemp = Mid(strTemp, Count)
      '      Exit Do
       ' End If
        'Count = Count + 1
'    Loop
'    strNo = strReverse(strTemp)
End If
    
If Not IsNumeric(strNo) <= 0 Or Val(strNo) <= 0 Then
    'MsgBox "Invalid cheque book start number specified !", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(503), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtStartNo
    Exit Sub
End If


'Validate the no of leaves
If Val(txtLeaves.Text) <= 0 Then
    'MsgBox "Invalid number of leaves specified !" & vbCrLf & "Leaves should be 10, 25, 50 , 100", vbExclamation, gAppName & " - Error"
    MsgBox GetResourceString(504), vbExclamation, gAppName & " - Error"
    ActivateTextBox txtLeaves
    Exit Sub
Else
    LeafCount = Val(txtLeaves.Text)
End If

'RaiseEvent LeavesPurchased
ReDim StrShare(LeafCount - 1)
'Get the first share No

FstShareNo = Val(strNo)
For count = 0 To LeafCount - 1
    StrShare(count) = strLeft & CStr(FstShareNo + count) & strRight
Next

Dim Cancel As Boolean

'RaiseEvent LeavesPurchased(strLeft, Val(strNo), Val(txtLeaves.Text), Cancel)
RaiseEvent ShareIssued(StrShare, Cancel)

If Cancel = True Then Exit Sub

Unload Me

End Sub

'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'
Private Sub cmdInvert_Click()
Dim I As Integer
For I = 0 To lstCheque.ListCount - 1
    lstCheque.Selected(I) = Not lstCheque.Selected(I)
Next I


End Sub



'
Private Sub cmdRemove_Click()
#If junk Then
Dim LeafArr() As Long
Dim I As Integer
'Prepare the array
ReDim Preserve LeafArr(0)
For I = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(I) = True Then
        LeafArr(UBound(LeafArr)) = lstCheque.List(I)
        ReDim Preserve LeafArr(UBound(LeafArr) + 1)
    End If
Next I

If UBound(LeafArr) = 0 Then
    Exit Sub
End If
RaiseEvent RemoveLeaves(LeafArr)

'MsgBox "Specified leaves remove successfully.", vbInformation, gAppName & " - Error"
MsgBox GetResourceString(557), vbInformation, gAppName & " - Error"

Call FillChequeList
#End If
End Sub

'
Private Sub cmdSelect_Click()
Dim LeafArr() As String
Dim I As Integer
'Prepare the array
ReDim Preserve LeafArr(0)
For I = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(I) = True Then
        LeafArr(UBound(LeafArr)) = lstCheque.List(I)
        ReDim Preserve LeafArr(UBound(LeafArr) + 1)
    End If
Next I
If UBound(LeafArr) = 0 Then LeafArr(0) = "": Exit Sub

'If He has Selected Any leaf then
ReDim Preserve LeafArr(UBound(LeafArr) - 1)
RaiseEvent ShareReturned(LeafArr)

Unload Me

End Sub

Private Sub Form_Load()
Dim Rst As Recordset

Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'Find the Max imum Share no
gDbTrans.SqlStmt = "Select CertNo from ShareTrans where CertID = " & _
    "(Select Max(Certid) from ShareTrans)"
If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
         txtStartNo.Text = FormatField(Rst(0))

If IsNumeric(txtStartNo.Text) Then txtStartNo.Text = Val(txtStartNo.Text) + 1
Call SetKannadaCaption

End Sub
Private Sub Form_Unload(Cancel As Integer)

Set frmMMShare = Nothing
End Sub


Private Sub txtLeaves_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
End If

End Sub

Private Sub txtStartNo_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdAdd.Enabled = True
Else
    'cmdOK.Enabled = False
    cmdAdd.Enabled = False
End If
End Sub


