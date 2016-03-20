VERSION 5.00
Begin VB.Form frmPDAccEx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   2205
   ClientTop       =   2265
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit..."
      Height          =   375
      Left            =   3570
      TabIndex        =   6
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4875
      TabIndex        =   7
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options..."
      Height          =   375
      Left            =   2265
      TabIndex        =   5
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   5955
      Begin VB.CheckBox chk 
         Caption         =   "View all accounts "
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   3660
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Accept..."
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   345
         Left            =   4710
         TabIndex        =   3
         Top             =   1245
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1140
         TabIndex        =   1
         Top             =   600
         Width           =   3495
      End
      Begin VB.ListBox lst 
         ForeColor       =   &H00000000&
         Height          =   2595
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   1020
         Width           =   4515
      End
      Begin VB.ListBox lst 
         Height          =   2595
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.Label txtAccID 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   915
      End
      Begin VB.Label txtdate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2820
         TabIndex        =   11
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label lblDate 
         Caption         =   "Label1"
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   2265
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   4920
         Top             =   630
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmPDAccEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Auto As Boolean
Dim m_Accid() As Long
Dim FilledArr() As Currency
Private WithEvents m_Amount As frmAmount
Attribute m_Amount.VB_VarHelpID = -1
Public M_pigmyAmount As Currency
Dim m_rstMaster As Recordset
Public ArrayIdx As Long
Private m_AccHolders() As String
Public Function GetListIndex(TextToSearch As String, lst As ListBox) As Long
    
    Dim Count As Long
    Dim RetIndex As Long
    RetIndex = -1
    For Count = 0 To lst.ListCount
       If StrComp(TextToSearch, lst.List(Count)) = 0 Then
          RetIndex = Count
          Exit For
       End If
    Next
    
    GetListIndex = RetIndex
    
End Function

Private Sub SetKannadaCaption()
Dim ctrl As Control
For Each ctrl In Me
   On Error GoTo NextCount
   ctrl.FontName = gFontName
   If Not TypeOf ctrl Is ComboBox Then
      ctrl.FontSize = gFontSize
   End If
NextCount:
Next
lblDate.Caption = LoadResString(gLangOffSet + 38) & " " & _
    LoadResString(gLangOffSet + 37)  'Transaction Date
cmdAccept.Caption = LoadResString(gLangOffSet + 4)
cmdClose.Caption = LoadResString(gLangOffSet + 11)
cmdSave.Caption = LoadResString(gLangOffSet + 7)
cmdSubmit.Caption = LoadResString(gLangOffSet + 490)
cmdOptions.Caption = LoadResString(gLangOffSet + 491)
chk.Caption = LoadResString(gLangOffSet + 492)

End Sub

'Public Expdate As Date
Public Function InsertIntoPDTransTab()
'Coded By Vinay

'Declare The Variables
Dim TransDep As wisTransactionTypes
Dim i As Integer
Dim Loan As Boolean
Dim Amount() As Currency
Dim Balance() As Currency
Dim TransId() As Long
Dim TransDate As String
Dim Lret As Long
Dim Particulars As String
'Initialize The Variables
TransDep = wDeposit
Loan = False
TransDate = gStrDate  'This has To Be Discussed
Particulars = "Pigmy Payment"
Debug.Print "DATE Variable "

ReDim Amount(UBound(m_Accid))
ReDim Balance(UBound(m_Accid))
ReDim TransId(UBound(m_Accid))

gDBTrans.SQLStmt = "Select * from Pdtrans where TransDate= #" & TransDate & _
                            "# and userId= " & gUserID
Dim Rst As Recordset

If gDBTrans.Fetch(Rst, adOpenForwardOnly) >= 1 Then
'If MsgBox("The Pigmy Collector Has Already Paid His Collections for This Date ," & vbCrLf & "Do You Wish OverWrite The previous Entries ", vbYesNo + vbDefaultButton2, gAppName) = vbNo Then Exit Function
If MsgBox(LoadResString(gLangOffSet + 800) & vbCrLf & LoadResString(gLangOffSet + 801), vbYesNo + vbDefaultButton2, gAppName) = vbNo Then Exit Function
Debug.Print "To Be Discussed"
For i = 1 To UBound(m_Accid)
      If FilledArr(i - 1) > 0 Then
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Delete * from Pdtrans where TransDate= #" & TransDate & _
                            "# and userId= " & gUserID & "and Accid = " & m_Accid(i) & " and loan = " & Loan
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            ''MsgBox "Unable To Perform Transactions", vbCritical, gAppName & "ERROR !"
            MsgBox LoadResString(gLangOffSet + 535), vbCritical, gAppName & "ERROR !"
            Exit Function
        Else
            gDBTrans.CommitTrans
        End If
    End If
Next i
End If
Me.MousePointer = vbHourglass

For i = 1 To UBound(m_Accid)
gDBTrans.SQLStmt = "Select TransID,Balance,TransDate from Pdtrans where UserId = " & gUserID & _
                                                " and Accid = " & m_Accid(i) & " and loan = " & Loan & " order By transId desc "
Lret = gDBTrans.Fetch(Rst, adOpenStatic)
    If Lret <= 0 Then
        GoTo NextCount
    End If
    'Check For Transaction date
    If WisDateDiff(FormatField(Rst("TransDate")), CStr(TransDate)) < 0 Then
         ' MsgBox "Transaction Date  Of Accid  " & CStr(m_Accid(i)) & " is Earlier " & vbCrLf & "Unable To Perform Transactions" _
            , vbCritical, "Date ERROR !"
    
         MsgBox LoadResString(gLangOffSet + 572), vbExclamation, wis_MESSAGE_TITLE
           Me.MousePointer = vbDefault
           Exit Function
    End If
    Amount(i) = FilledArr(i - 1)
    
    If Lret = 0 Then
        TransId(i) = 1
        Balance(i) = Amount(i)
    Else
        TransId(i) = Val(FormatField(Rst("transid"))) + 1
        Balance(i) = CCur(Val(FormatField(Rst("Balance")))) + Amount(i)
    End If
If FilledArr(i - 1) > 0 Then
             
gDBTrans.BeginTrans

gDBTrans.SQLStmt = "Insert into Pdtrans (Accid, UserId , Amount, TransId , transDate , Loan ,Balance ,Transtype, Particulars) " & _
                        " values (" & m_Accid(i) & " ," & gUserID & " ," & Amount(i) & _
                        " ," & TransId(i) & " , #" & TransDate & "#  ," & Loan & " ," & _
                        Balance(i) & " ," & TransDep & ",'" & Particulars & "')"
                
    If Not gDBTrans.SQLExecute Then
        ''MsgBox "Unable To Perform Transactions", vbCritical, gAppName & "ERROR !"
        MsgBox LoadResString(gLangOffSet + 535), vbCritical, gAppName & "ERROR !"
        gDBTrans.RollBack
        Me.MousePointer = vbDefault
        Exit Function
    Else
        gDBTrans.CommitTrans
    End If
End If
NextCount:
DoEvents
 Next i
''MsgBox "Values Stored In The table", vbInformation, gAppName & "Save!"
MsgBox LoadResString(gLangOffSet + 802), vbInformation, wis_MESSAGE_TITLE
Me.MousePointer = vbDefault
End Function




Public Function FillCustomerNames(UserID As Long) As Boolean

'Show the express dialog relevant to that user only
'UserID = 2

'Get an rst of depositors for the user
    gDBTrans.SQLStmt = "Select Title +' ' + " & _
            "FirstName + ' ' + " & _
            " MiddleName + ' ' + " & _
            " LastName  as name ," & _
            " PdMaster.AccID" & _
            " from NameTab, PDMaster where " & _
            " NameTab.CustomerID = PDMaster.CustomerID " & _
            " And PDMaster.UserID = " & UserID & _
            " Order By PdMaster.AccId"
    
'Dim Rst As Recordset
If gDBTrans.Fetch(m_rstMaster, adOpenStatic) <= 0 Then Exit Function


'Fill names into list boxes
ReDim m_Accid(m_rstMaster.RecordCount - 1)
    
    Dim i As Long
    ReDim m_AccHolders(m_rstMaster.RecordCount - 1)
    
    While Not m_rstMaster.EOF
        lst(0).AddItem FormatField(m_rstMaster("Name"))
        lst(1).AddItem FormatField(m_rstMaster("Name"))
        m_Accid(i) = Val(FormatField(m_rstMaster("AccId")))
        m_AccHolders(i) = FormatField(m_rstMaster("Name"))
        m_rstMaster.MoveNext
        i = i + 1
    Wend
    
    ReDim FilledArr(lst(0).ListCount)

End Function

'
Public Function ZZZGetListIndex(TextToSearch As String, lst As ListBox) As Long
Dim Low As Integer, High As Integer, Middle As Integer
Dim PrevMiddle As Integer
Dim Retval As Integer
Dim TargetPosition As Integer
Dim SrchLen As Integer
Dim BinarySearch As Integer
Dim SrchText As String
PrevMiddle = -1





'Initialize variables
    SrchLen = Len(TextToSearch)
    'SrchText = UCase(TextToSearch)
    SrchText = TextToSearch
    TargetPosition = 0
    Low = 0
    High = lst.ListCount - 1
    PrevMiddle = -1
'Perform binary search on the data in the list box - lst
    
    'binary Search wil work for sorted array
    ' as here we can not have the array sorted because
    ' we have to load account holders by theid accont no
    
   Exit Function
   
    BinarySearch = 1
    Do
        Middle = (Low + High) / 2
        If Middle = PrevMiddle Then
            BinarySearch = -1
            Exit Do
        End If
        Select Case Left(lst.List(Middle), SrchLen) '
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
        While Middle > 0 And Left(lst.List(Middle - 1), SrchLen) = SrchText
            Middle = Middle - 1
        Wend
    End If

'Return the right index to caller
    If BinarySearch <> -1 Then
        ZZZGetListIndex = Middle
    Else
        ZZZGetListIndex = -1
    End If

End Function

Public Function PigmyAmountCollected(FilledArr() As Currency, UboundOfArr As Integer) As Currency
'Coded By Vinay
PigmyAmountCollected = -1
On Error GoTo ErrorLine

Dim Amount As Currency
Dim Count As Integer
If LBound(FilledArr) = UBound(FilledArr) Then
Count = LBound(FilledArr)
Amount = LBound(FilledArr)
Else
For Count = LBound(FilledArr) To UBound(FilledArr)
Amount = Amount + FilledArr(Count)
Next Count
End If

PigmyAmountCollected = Amount
Exit Function
ErrorLine:
        MsgBox Err.Description
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
Dim i As Integer
Dim LstVisible As Byte
If Me.lst.Item(1).Visible = True Then
    LstVisible = 1
    i = lst(1).ListIndex
    If Me.lst.Item(1).ListIndex < 0 Then
        ''MsgBox "Please Select The Persons Name From The List ", vbOKOnly, gAppName & "Error!"
        MsgBox LoadResString(gLangOffSet + 803), vbOKOnly, gAppName & "Error!"
        Exit Sub
    End If
End If

If Me.lst.Item(0).Visible = True Then
    LstVisible = 0
    i = lst(0).ListIndex
    If Me.lst.Item(0).ListIndex < 0 Then
        ''MsgBox "Please Select The Persons Name From The List ", vbOKOnly, gAppName & "Error!"
        MsgBox LoadResString(gLangOffSet + 803), vbOKOnly, gAppName & "Error!"
        Exit Sub
    End If
End If

    If m_Amount Is Nothing Then Set m_Amount = New frmAmount
    
    m_Amount.txtAmount = FilledArr(lst(LstVisible).ListIndex)
    m_Amount.Show vbModal
    Text1.Text = lst(LstVisible).List(i)

'    Text1.SetFocus
 '   Text1.SelStart = 0
  '  Text1.SelLength = Len(Text1.Text)
    If lst(0).ListIndex < lst(0).ListCount - 1 Then lst(0).ListIndex = lst(0).ListIndex + 1

End Sub

Private Sub cmdClose_Click()

        ' Temp Code Starts
'        Dim Sum As Currency
'        Dim i As Integer
'        Sum = 0
'        If UBound(FilledArr) = LBound(FilledArr) Then
'        Sum = 0
'        Else
'        For i = 0 To UBound(FilledArr)
'        Sum = Sum + FilledArr(i)
'        Next i
'        End If
'        MsgBox "Total Pigmy Amount Collected  is " & Sum, vbOKOnly, gAppName & " Error !"
        ' Temp Code Ends

gDBTrans.CloseDB
Unload Me

End Sub




Private Sub cmdOptions_Click()
frmOption.Show vbModal, Me
End Sub


Private Sub cmdSave_Click()
Call InsertIntoPDTransTab
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo ErrLine
M_pigmyAmount = PigmyAmountCollected(FilledArr, UBound(FilledArr))
frmDenomination.TxtExpectedAmount.Text = M_pigmyAmount
frmDenomination.Show vbModal
ErrLine:
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
#If JUNK Then
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
    
#End If

  Call SetKannadaCaption
  txtdate = FormatDate(gStrDate)
  Image1.Picture = LoadResPicture(115, vbResIcon)
  cmdSubmit.Enabled = True
  cmdSave.Enabled = False
  Call FillCustomerNames(CLng(gUserID))
  Text1.Locked = True
  Text1.TabStop = False

End Sub






Private Sub lst_Click(Index As Integer)

If lst(Index).ListIndex >= 0 Then
    cmdAccept.Enabled = True
Else
    cmdAccept.Enabled = False
End If

End Sub

Private Sub lst_DblClick(Index As Integer)
Text1.Text = lst(Index).List(lst(Index).ListIndex)
txtAccID = m_Accid(lst(Index).ListIndex)
End Sub





Private Sub m_Amount_OKClicked(Amount As Currency)
Dim strName As String
Dim idx As Long
Dim IdxofLst0 As Long
Dim IdxofLst1 As Long
Dim Count As Long
Dim Vis1 As Boolean
Dim Vis2  As Boolean
'Update lst 0
Vis1 = lst(0).Visible
Vis2 = lst(1).Visible

lst(0).Visible = False
lst(1).Visible = True

If Vis1 Then
    'Search for str in lst(1)
    IdxofLst0 = lst(0).ListIndex
    idx = GetListIndex(lst(0).Text, lst(1))
    If idx >= 0 Then lst(1).RemoveItem idx
    
    'You have to modify the string in lst(0)
    idx = lst(0).ListIndex
    strName = lst(0).List(idx)
    
    strName = Left(strName, Len(strName) - Len(CStr(Amount)) - Len("....."))
    strName = strName & "....." & CStr(Amount)
    
    FilledArr(idx) = Amount
    'Count = lst(0).ListIndex
ElseIf Vis2 Then
    'Search for str in lst(0)
    idx = GetListIndex(lst(1).Text, lst(0))
    'Remove the name form the listbox
    If idx >= 0 Then lst(1).RemoveItem lst(1).ListIndex
    
    strName = lst(0).List(idx)
    strName = strName & "....." & Amount
    FilledArr(idx) = Amount
    'Count = Idx
   ' lst(0).RemoveItem idx
    'lst(0).AddItem strName
End If

     m_AccHolders(idx) = strName
        lst(0).Clear
        For Count = LBound(m_AccHolders) To UBound(m_AccHolders)
            lst(0).AddItem m_AccHolders(Count)
            lst(0).ItemData(lst(0).NewIndex) = m_Accid(Count)
        Next
    
lst(0).Visible = Vis1
lst(1).Visible = Vis2

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
   '   idx = lst(0).ListIndex
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

