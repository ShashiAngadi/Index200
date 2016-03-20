VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Ledger"
   ClientHeight    =   6360
   ClientLeft      =   2940
   ClientTop       =   2025
   ClientWidth     =   6030
   Icon            =   "LedgerCreation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "Clea&r"
      Enabled         =   0   'False
      Height          =   400
      Left            =   240
      TabIndex        =   13
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CheckBox chkContra 
      Caption         =   "Is Contra Head"
      Height          =   345
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   3075
   End
   Begin VB.TextBox txtLedgerEnglish 
      Height          =   395
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   3645
   End
   Begin VB.CheckBox chkNegBal 
      Caption         =   "Negative Balance"
      Height          =   345
      Left            =   3660
      TabIndex        =   9
      Top             =   1830
      Width           =   2235
   End
   Begin VB.TextBox txtOpBalance 
      Height          =   395
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtLedgerName 
      Height          =   395
      Left            =   2160
      TabIndex        =   4
      Top             =   810
      Width           =   3645
   End
   Begin ComctlLib.ListView lvwLedger 
      Height          =   2745
      Left            =   180
      TabIndex        =   12
      Top             =   2670
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4842
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox cmbParent 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   3765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4590
      TabIndex        =   1
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   3150
      TabIndex        =   0
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblLedgerEnglish 
      Caption         =   "Ledger Name"
      Height          =   390
      Left            =   60
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5880
      Y1              =   5595
      Y2              =   5595
   End
   Begin VB.Label lblOpBalance 
      Caption         =   "Opening Balance"
      Height          =   390
      Left            =   180
      TabIndex        =   7
      Top             =   1830
      Width           =   1935
   End
   Begin VB.Label lblLedger 
      Caption         =   "Ledger Name"
      Height          =   390
      Left            =   60
      TabIndex        =   11
      Top             =   810
      Width           =   2055
   End
   Begin VB.Label lblParentName 
      Caption         =   "Select Parent Ledger"
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   2055
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1

Private m_HeadID As Long

' are used in the class

Public Event OKClick()
Public Event CancelClick()
Public Event ClearClick()
Public Event LookUpClick(ParentID As Long)
Public Event LvwLedgerClick(headID As Long)

Private m_dbOperation As wis_DBOperation
'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
If gLangOffSet = 0 Then txtLedgerEnglish.Enabled = False
txtLedgerEnglish.Font.name = "MS Sans Serif"
txtLedgerEnglish.Font.Size = 8

'set the Kannada for all controls
lblParentName.Caption = GetResourceString(160, 36)
lblLedger.Caption = GetResourceString(36, 35)
lblLedgerEnglish.Caption = GetResourceString(36, 35, 468)

lblOpBalance.Caption = GetResourceString(284)
chkContra.Caption = GetResourceString(270, 36)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(11)
cmdCancel.Caption = GetResourceString(8)

End Sub



Private Sub ClearControls()

cmbParent.ListIndex = -1
txtLedgerName.Text = ""
txtLedgerEnglish.Text = ""
txtOpBalance.Text = ""
lvwLedger.ColumnHeaders.Clear
cmdOk.Caption = GetResourceString(10)

End Sub


Private Function Validated() As Boolean

Validated = False

If cmbParent.ListIndex = -1 Then
    MsgBox "Select Parent Name ", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

If Not CurrencyValidate(txtOpBalance.Text, True) Then
    MsgBox "Invalid opening balance specified", vbInformation, wis_MESSAGE_TITLE
    txtOpBalance.SetFocus
    Exit Function
End If

If Len(txtLedgerName.Text) = 0 Then
    MsgBox "No LedgerName specified", vbInformation, wis_MESSAGE_TITLE
    txtLedgerName.SetFocus
    Exit Function
End If

If Val(txtLedgerName.Text) > 0 Then
    MsgBox "Invalid LedgerName specified", vbInformation, wis_MESSAGE_TITLE
    txtLedgerName.SetFocus
    Exit Function
End If

Validated = True

End Function
Private Sub cmbParent_Click()

Dim ParentID As Long
Dim rstHead As Recordset

If cmbParent.ListIndex = -1 Then Exit Sub

ParentID = cmbParent.ItemData(cmbParent.ListIndex)

Me.lblLedger.Refresh

RaiseEvent LookUpClick(ParentID)

gDbTrans.SqlStmt = "Select * FROM ParentHeads WHERE ParentID = " & ParentID
Dim rstTemp As Recordset
Call gDbTrans.Fetch(rstTemp, adOpenDynamic)

cmdOk.Enabled = IIf(FormatField(rstTemp("UserCreated")) <= 2, True, False)

End Sub

Private Sub cmdCancel_Click()
Unload Me
RaiseEvent CancelClick

End Sub

Private Sub cmdClear_Click()
    RaiseEvent ClearClick
End Sub

Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

RaiseEvent OKClick

m_dbOperation = Insert

End Sub

Private Sub Form_Load()

'Center the form
CenterMe Me

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)


If gLangOffSet <> 0 Then SetKannadaCaption

Call LoadParentHeads(cmbParent)

m_dbOperation = Insert


End Sub

Private Sub LoadParentHeads(ctrlComboBox As ComboBox)

Dim rstParent  As ADODB.Recordset

ctrlComboBox.Clear

gDbTrans.SqlStmt = " SELECT ParentName,ParentID,ParentNameEnglish " & _
                   " FROM ParentHeads " & _
                   " ORDER BY AccountType,ParentID"
'WHERE UserCreated <= 2
Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)

Do While Not rstParent.EOF

    ctrlComboBox.AddItem FormatField(rstParent.Fields("ParentName"))
    ctrlComboBox.ItemData(ctrlComboBox.newIndex) = FormatField(rstParent.Fields("ParentID"))
    
    'Move to the next record
    rstParent.MoveNext
    
Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLedger = Nothing
End Sub

Private Sub lvwLedger_DblClick()
On Error Resume Next

Dim opBalance As Currency

' selected item key will be like this - "A2001"
' so we have to fetch only value

With lvwLedger.SelectedItem
    m_HeadID = Val(Mid(.Key, 4))
    txtLedgerName.Text = .Text
    opBalance = Val(.SubItems(1))
    txtOpBalance.Text = FormatCurrency(Abs(opBalance))
    chkNegBal = IIf(opBalance < 0, vbChecked, vbUnchecked)
    cmbParent.Locked = True
    txtLedgerEnglish.Text = .SubItems(2)
End With

m_dbOperation = Update
cmdOk.Caption = GetResourceString(171)
cmdOk.Enabled = True

RaiseEvent LvwLedgerClick(m_HeadID)


End Sub

Private Sub txtLedgerEnglish_GotFocus()
    Call ToggleWindowsKey(winScrlLock, False)
    Call Translate(txtLedgerName, txtLedgerEnglish)
End Sub

Private Sub txtLedgerEnglish_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub

Private Sub txtLedgerName_LostFocus()
'txtLedgerName = ConvertToProperCase(txtLedgerName.Text)
End Sub
