VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   2310
   ClientTop       =   3210
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit..."
      Height          =   375
      Left            =   3690
      TabIndex        =   6
      Top             =   3810
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4860
      TabIndex        =   7
      Top             =   3810
      Width           =   1185
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options..."
      Height          =   375
      Left            =   2430
      TabIndex        =   5
      Top             =   3810
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3585
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   5955
      Begin VB.CheckBox chk 
         Caption         =   "View remaining accounts only"
         Height          =   255
         Left            =   690
         TabIndex        =   4
         Top             =   3210
         Value           =   1  'Checked
         Width           =   3765
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Accept..."
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   345
         Left            =   4710
         TabIndex        =   3
         Top             =   930
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   4515
      End
      Begin VB.ListBox lst 
         ForeColor       =   &H00000000&
         Height          =   2400
         Index           =   0
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   690
         Width           =   4515
      End
      Begin VB.ListBox lst 
         Height          =   2400
         Index           =   1
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   4920
         Top             =   300
         Width           =   525
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Auto As Boolean

Dim FilledArr() As Currency
Private WithEvents m_Amount As frmAmount
Attribute m_Amount.VB_VarHelpID = -1


Private Function GetListIndex(TextToSearch As String, lst As ListBox) As Long
Dim Low As Integer, High As Integer, Middle As Integer
Dim PrevMiddle As Integer
Dim Retval As Integer
Dim TargetPosition As Integer
Dim SrchLen As Integer
Dim BinarySearch As Integer
Dim SrchText As String

'Initialize variables
    SrchLen = Len(TextToSearch)
    SrchText = UCase(TextToSearch)
    TargetPosition = 0
    Low = 0
    High = lst.ListCount - 1

'Perform binary search on the data in the list box - lst
    BinarySearch = 1
    Do
        Middle = (Low + High) \ 2
        If Middle = PrevMiddle Then
            BinarySearch = -1
            Exit Do
        End If
        Select Case UCase(Left(lst.List(Middle), SrchLen))
            Case Is = SrchText
                TargetPosition = Middle
                Exit Do
            Case Is > SrchText
                High = Middle - 1
            Case Is < SrchText
                Low = Middle + 1
        End Select
        PrevMiddle = Middle
    Loop Until TargetPosition <> 0

'Trace backwards to first match
    If BinarySearch <> -1 Then
        While Middle > 0 And UCase(Left(lst.List(Middle - 1), SrchLen)) = SrchText
            Middle = Middle - 1
        Wend
    End If

'Return the right index to caller
    If BinarySearch <> -1 Then
        GetListIndex = Middle
    Else
        GetListIndex = -1
    End If

End Function

Private Sub chk_Click()
    If chk.value = vbChecked Then
        lst(0).Visible = True
        lst(1).Visible = False
        lst(0).TabIndex = 1
    Else
        lst(0).Visible = False
        lst(1).Visible = True
        lst(1).TabIndex = 1
    End If
    
End Sub


Private Sub cmdAccept_Click()
    Set m_Amount = New frmAmount
    m_Amount.Show vbModal
    Text1.SetFocus
    Text1.SelStart = 0: Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
Dim i As Integer
Dim FIleNo As Integer
Dim str As String
Dim Pos As Integer
FIleNo = FreeFile
Open App.Path & "\" & "win.txt" For Input As #FIleNo
While Not EOF(FIleNo)
NextLine:
    'On Error GoTo NextLine
    Line Input #FIleNo, str
    If Trim$(str) = "" Then GoTo NextLine
    'Debug.Print Str
    Pos = InStr(1, str, " ", vbTextCompare)
    str = Mid(str, Pos + 1, Len(str) - Pos)
    Pos = InStr(1, str, " ", vbTextCompare)
    str = Mid(str, Pos + 1, Len(str) - Pos)
    Pos = InStr(1, str, " ", vbTextCompare)
    str = Left(str, Pos - 1)
    lst(0).AddItem str
    lst(1).AddItem str
Wend
Close FIleNo
'List1.ListIndex = 0
ReDim FilledArr(lst(0).ListCount)
End Sub





Private Sub lst2_Click()

End Sub

Private Sub lst_DblClick(Index As Integer)
Text1.Text = lst(Index).List(lst(Index).ListIndex)
End Sub


Private Sub lst_GotFocus(Index As Integer)
If lst(Index).ListIndex >= 0 Then
    cmdAccept.Enabled = True
Else
    cmdAccept.Enabled = False
End If

End Sub


Private Sub m_Amount_OKClicked(Amount As Currency)
Dim strName As String
Dim idx As Long
Dim Count As Long
'Update lst 0
    If lst(0).Visible Then
        'Search for str in lst(1)
        idx = GetListIndex(lst(0).List(lst(0).ListIndex), lst(1))
        If idx >= 0 Then
            lst(1).RemoveItem idx
        End If
        'You have to modify the string in lst(0)
        idx = lst(0).ListIndex
        strName = lst(0).List(idx)
        If FilledArr(idx) = 0 Then
            strName = strName & "....." & CStr(Amount)
        Else
            'strName = strName & "     "
            strName = Left(strName, Len(strName) - Len(CStr(Amount)) - Len("....."))
            strName = strName & "....." & CStr(Amount)
        End If
        lst(0).RemoveItem idx
        lst(0).AddItem strName
        
        FilledArr(idx) = Amount
        'Count = lst(0).ListIndex
    ElseIf lst(1).Visible Then
        'Search for str in lst(0)
        idx = GetListIndex(lst(1).List(lst(1).ListIndex), lst(0))
        lst(1).RemoveItem lst(1).ListIndex
        strName = lst(0).List(idx)
        strName = strName & "....." & Amount
        FilledArr(idx) = Amount
        'Count = Idx
        lst(0).RemoveItem idx
        lst(0).AddItem strName
    End If
    
'Now modify the string in lst(0)
    Dim MyStr As String
    
    'Text1.SetFocus
End Sub

Private Sub Text1_Change()
Dim Lines As Integer
On Error Resume Next
'Prelim check
    If Text1.Text = "" Then
        lst(0).ListIndex = 0
        lst(0).ListIndex = 0
        Exit Sub
    End If
    Lines = (lst(0).Height / 200) - 2

Dim idx As Long
If lst(0).Visible Then
    idx = GetListIndex(Text1.Text, lst(0))
    If idx >= 0 Then
        lst(0).ListIndex = idx + Lines
        lst(0).ListIndex = idx
    End If
Else
    idx = GetListIndex(Text1.Text, lst(1))
    
    lst(1).ListIndex = idx + Lines
    lst(1).ListIndex = idx
End If

End Sub


Private Sub Text1_LostFocus()
'If lst(0).Visible Then
'    lst(0).SetFocus
'Else
'    lst(1).SetFocus
'End If
End Sub

