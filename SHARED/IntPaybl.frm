VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIntPayble 
   Caption         =   "Interest Payable"
   ClientHeight    =   6540
   ClientLeft      =   2040
   ClientTop       =   1230
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web"
      Height          =   345
      Left            =   2610
      TabIndex        =   12
      Top             =   6120
      Width           =   915
   End
   Begin VB.CheckBox ChkAdd 
      Caption         =   "Add interest receivable amount to Loan head"
      Height          =   345
      Left            =   150
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox cmbHead 
      Height          =   315
      Left            =   1290
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   345
      Left            =   3540
      TabIndex        =   7
      Top             =   6120
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4500
      TabIndex        =   1
      Top             =   6120
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5490
      TabIndex        =   2
      Top             =   6120
      Width           =   945
   End
   Begin VB.PictureBox picOut 
      Height          =   5625
      Left            =   150
      ScaleHeight     =   5565
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   390
      Width           =   6285
      Begin VB.TextBox txtAfter 
         Height          =   285
         Left            =   1350
         ScrollBars      =   1  'Horizontal
         TabIndex        =   5
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox txtBefore 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtBox 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   0
         Width           =   1035
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   5145
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   9075
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin VB.Label lblHead 
      Caption         =   "Label1"
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   6150
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Height          =   315
      Left            =   390
      TabIndex        =   8
      Top             =   30
      Width           =   5895
   End
End
Attribute VB_Name = "frmIntPayble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PutTotal As Boolean
Private m_ShowBalance  As Boolean
Private m_ShowTotal As Boolean

Private m_PrevBal As Double
Private m_CurAmount As Double
Private m_PrevAmount As Double
Private m_Shown As Boolean

Public Event Initialised(Min As Long, Max As Long)
Public Event Processed(Ratio As Single)
Public Property Let AccNum(Index As Integer, NewValue As String)

With grd
    '.Col = 0
    '.Row = Index
    '.Text = Newvalue
    .TextMatrix(Index, 0) = NewValue
    RaiseEvent Processed(.Row / .Rows)
End With


End Property

Public Property Let KeyData(Index As Integer, NewValue As Long)

With grd
    If Index >= .Rows Then Exit Property
    .RowData(Index) = NewValue
End With

End Property


Public Property Let Amount(Index As Integer, NewValue As Currency)
Dim BalAmount As Currency

With grd
    If Index Then .Row = Index
    .Col = 3
    .Text = NewValue
    txtBox = NewValue
End With

End Property

Public Property Get KeyData(Index As Integer) As Long

With grd
    If Index > .Rows - 2 And m_PutTotal Then Exit Property
    KeyData = .RowData(Index)
End With

End Property
Public Property Get Amount(Index As Integer) As Currency

With grd
    If Index > grd.Rows - 2 And m_PutTotal Then Exit Property
    .Col = 3
    .Row = Index
    Amount = Val(.Text) '
End With

End Property


Public Property Let Balance(Index As Integer, NewValue As Currency)

With grd
    '.Col = 2
    '.Row = Index
    '.Text = FormatCurrency(Newvalue)
    .TextMatrix(Index, 2) = NewValue
End With

End Property
Public Sub ShowForm()
m_Shown = True
Me.Show 1

End Sub

Public Property Let TotalColoumn(NewValue As Boolean)
    m_ShowTotal = NewValue
End Property

Public Property Let BalanceColoumn(NewValue As Boolean)
    m_ShowBalance = NewValue
End Property

Public Property Get Total(Index As Integer) As Currency
With grd
    .Col = 4
    .Row = Index
    Total = Val(.Text)
End With

End Property

Public Property Let Total(Index As Integer, NewValue As Currency)
With grd
    .Col = 4
    .Row = Index
    .Text = FormatCurrency(NewValue)
End With
End Property

Public Property Let CustName(Index As Integer, NewValue As String)
With grd
    '.Row = Index
    '.Col = 1
    '.Text = Newvalue
    .TextMatrix(Index, 1) = NewValue
End With
End Property

Public Property Let PutTotal(NewValue As Boolean)
    m_PutTotal = NewValue
End Property

Public Property Let Title(ColIndex As Integer, NewValue As String)

With grd
    If ColIndex > grd.Cols - 1 Then Exit Property
    .Col = ColIndex
    .Row = 0
    .Text = NewValue
    .CellAlignment = 4
    .CellFontBold = True
    If .Col = 3 Then txtBox = NewValue: .Col = 0
End With

End Property


'
Public Function LoadContorls(ControlLoopCount As Integer, CtrlGap As Integer)
'On Error Resume Next
If ControlLoopCount < 1 Then Exit Function

Dim Wid As Single
With grd
    .ZOrder 0
    .Top = 0: .Left = 0
    .Width = picOut.Width - 50
    .Height = picOut.Height - 50
    .AllowUserResizing = flexResizeNone
    
    .Cols = 5
    .Rows = ControlLoopCount + 1
    .FixedCols = 1: .FixedRows = 1
    .ScrollBars = flexScrollBarVertical
    Wid = .Width / .Cols
    .ColWidth(0) = Wid * 0.35
    .ColWidth(1) = Wid * 2
    .ColWidth(2) = Wid * 0.95
    .ColWidth(3) = Wid * 0.65
    .ColWidth(4) = Wid * 0.75
End With

LoadContorls = True

RaiseEvent Initialised(0, ControlLoopCount + 1)


End Function

Private Sub cmdCancel_Click()

grd.Clear
grd.Rows = 1
grd.Cols = 1
Me.Hide

End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
m_Shown = False

With wisMain.grdPrint
    .ReportTitle = lblTitle
    .GridObject = grd
    .CompanyName = gCompanyName
    .PrintGrid
End With

m_Shown = True

End Sub

Private Sub cmdWeb_Click()
Dim clswebGrid As New clsgrdWeb
With clswebGrid
    Set .GridObject = grd
    .CompanyAddress = ""
    .CompanyName = gCompanyName
    .ReportTitle = "Interest Payable"
    Call clswebGrid.ShowWebView '(grd)

End With
End Sub

Private Sub Form_Load()
m_ShowBalance = True
m_ShowTotal = True

Dim Ctrl As Control
On Error Resume Next
'Now Assign the Kannada fonts to the All controls
For Each Ctrl In Me
    Debug.Print Ctrl.Name & " " & Ctrl.Height
    If Not TypeOf Ctrl Is ComboBox Then
        Ctrl.Font.Name = gFontName
        Ctrl.Font.Size = gFontSize
    End If
Next Ctrl

cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
Err.Clear

End Sub

Private Sub Form_Resize()

With picOut
    .Width = Me.Width - 3 * .Left
    .Height = Me.Height - (cmdOk.Height * 2 + lblTitle.Height + 400)
End With
With cmdCancel
    .Left = Me.Width - .Width * 1.5
    .Top = picOut.Top + picOut.Height + 100
End With
With cmdOk
    .Left = cmdCancel.Left - .Width - 100
    .Top = cmdCancel.Top
End With
With cmdPrint
    .Top = cmdCancel.Top
    '.Left = picOut.Left
    .Left = cmdOk.Left - .Width - 100
    cmdWeb.Left = .Left - 100 - cmdWeb.Width
    cmdWeb.Top = .Top
End With

lblHead.Top = cmdOk.Top
cmbHead.Top = cmdOk.Top
With chkAdd
    .Top = cmdOk.Top + (cmdOk.Height - chkAdd.Height) / 2
    .Width = cmdPrint.Left - .Left - 100
End With

With grd
    .ZOrder 0
    .Top = 0: .Left = 0
    .Width = picOut.Width - 50
    .Height = picOut.Height - 50
    .AllowUserResizing = flexResizeNone
    
    .Cols = 5
    .FixedCols = 1: .FixedRows = 1
    .ScrollBars = flexScrollBarVertical
    
    Dim Wid As Single
    
    Wid = .Width / .Cols
    If Not m_ShowBalance Then Wid = Wid * 1.25
    If Not m_ShowTotal Then Wid = Wid * 1.15
    
    .ColWidth(0) = Wid * 0.35
    .ColWidth(1) = Wid * 2
    .ColWidth(2) = Wid * 0.95
    .ColWidth(3) = Wid * 0.65
    .ColWidth(4) = Wid * 0.85
    
    If Not m_ShowBalance Then .ColWidth(2) = 0.5
    If Not m_ShowTotal Then .ColWidth(4) = 0.5
    
End With

End Sub

Private Sub grd_EnterCell()
If Not m_Shown Then Exit Sub

With grd
    
    If .Row = 0 Then Exit Sub
    If .Col <> 3 Then Exit Sub
    If m_PutTotal And .Rows = .Row + 1 Then Exit Sub
    
    txtBox.Text = .Text
    m_PrevAmount = Val(.Text)
    'txtBox.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight
    txtBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
    txtBox.Visible = True
    
    On Error Resume Next
    ActivateTextBox txtBox
    txtBox.ZOrder 0
    Err.Clear
End With

End Sub

Private Sub grd_GotFocus()
    Call grd_EnterCell
End Sub

Private Sub grd_LeaveCell()

If Not m_Shown Then Exit Sub

Static InLoop As Boolean
If InLoop Then Exit Sub
Dim Bal As Currency
Dim Amount As Double
Dim PrevAmount As Double

With grd
    If .Col <> 3 Then Exit Sub
    If m_PutTotal And .Rows = .Row + 1 Then Exit Sub
    'mARK AS CONTROL IS IN LOOP
    InLoop = True
    
    'PrRow = .Row
    'Get the Previous amount
    PrevAmount = m_PrevAmount 'Val(.Text)
    
    'Now update the New Amount to the grid
    .Text = txtBox
    
    txtBox.Visible = False
    Amount = Val(txtBox)
    txtBox = ""
    
    'Get the present balance (i.e. from last coloumn)
    Bal = Val(.TextMatrix(.Row, .Cols - 1))
    
    '.Text = FormatCurrency(PrevBal + PrAmount - m_PrevAmount)
    'ANd Update the balance
    .TextMatrix(.Row, .Cols - 1) = FormatCurrency(Bal + Amount - PrevAmount)
    
    ''Now Update the Grans total(i.e. at the LAst row,ans same col
    '.Row = .Rows - 1
    '.Col = 3
    'txtBox = FormatCurrency(Val(.Text) + PrAmount - CurrAmount)
    '.Text = txtBox
    Bal = FormatCurrency(Val(.TextMatrix(.Rows - 1, .Col)) + Amount - PrevAmount)
    .TextMatrix(.Rows - 1, .Col) = Bal
    
    If m_PutTotal Then
        'And now put the Grand total of balance
        '.Col = .Cols - 1
        'PrevBal = Val(.Text)
        Bal = Val(.TextMatrix(.Rows - 1, .Cols - 1))
        .TextMatrix(.Rows - 1, .Cols - 1) = FormatCurrency(Bal + Amount - PrevAmount)
    End If
    '.Col = 3
    '.Row = PrRow

End With

'MARK AS CONTROL IS OUT OF LOOP
InLoop = False

End Sub

Private Sub grd_Scroll()
txtBox.Visible = True
txtBox.Move grd.Left + grd.ColPos(grd.Col), grd.Top + grd.RowPos(grd.Row) ', grd.CellWidth, grd.CellHeight
'txtbox.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight
If txtBox.Left < grd.Left + grd.ColPos(0) = txtBox.Width Then txtBox.Visible = False
If txtBox.Top < grd.Top + grd.RowPos(0) + txtBox.Height Then txtBox.Visible = False

End Sub

Private Sub txtAfter_GotFocus()
'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the next amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)
'If txtNo < txtAmount.Count - 1 Then
If txtNo < grd.Rows - 1 Then
    txtNo = txtNo + 1
    'Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub


Private Sub txtBefore_GotFocus()
'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the previous amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)

If txtNo > 0 Then
    txtNo = txtNo - 1
'    Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub

Private Sub txtBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode & " DOWN " & Shift
If KeyCode <> vbKeyTab Then Exit Sub

With grd
    If Shift = 1 Then
        If .Row = 1 Then Exit Sub
        .Row = .Row - 1
    Else
        If .Rows = .Rows - 1 Then Exit Sub
        .Row = .Row + 1
    End If
End With

End Sub

Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 37 Then 'Press Left Arrow
    If txtBox.SelStart = 0 And grd.Col > grd.FixedCols Then grd.Col = grd.Col - 1
    
ElseIf KeyCode = 39 Then 'Press Right Arrow
    If txtBox.SelStart = Len(txtBox.Text) And grd.Col < grd.Cols - 1 Then grd.Col = grd.Col + 1
    
ElseIf KeyCode = 38 Then  'Press UpArrow
    If grd.Row > grd.FixedRows Then grd.Row = grd.Row - 1
    
ElseIf KeyCode = 40 Then ' Press Doun Arroow
    If grd.Row < grd.Rows - 1 Then grd.Row = grd.Row + 1
    
ElseIf KeyCode = 33 Then  'Press PageUp
    If grd.Row > 0 Then grd.Row = grd.Row - 1
    
ElseIf KeyCode = 34 Then ' Press PageDown
    If grd.Row < grd.Rows - 1 Then grd.Row = grd.Row + 1
    
End If

End Sub

