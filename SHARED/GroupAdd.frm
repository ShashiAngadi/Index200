VERSION 5.00
Begin VB.Form frmAddGroup 
   Caption         =   "Add Group of Place,Caste,Account,Customer"
   ClientHeight    =   2580
   ClientLeft      =   3045
   ClientTop       =   2550
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtItemEnglish 
      Height          =   350
      Left            =   2415
      TabIndex        =   3
      Top             =   600
      Width           =   2955
   End
   Begin VB.ComboBox cmbCumulative 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1980
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CheckBox chkCumulative 
      Caption         =   "Cumulative"
      Height          =   300
      Left            =   150
      TabIndex        =   6
      Top             =   1530
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4320
      TabIndex        =   9
      Top             =   2130
      Width           =   1080
   End
   Begin VB.ComboBox cmbList 
      Height          =   315
      Left            =   2235
      TabIndex        =   5
      Top             =   1110
      Width           =   3165
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   400
      Left            =   4335
      TabIndex        =   7
      Top             =   1575
      Width           =   1050
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   400
      Left            =   3060
      TabIndex        =   8
      Top             =   1590
      Width           =   1050
   End
   Begin VB.TextBox txtItem 
      Height          =   350
      Left            =   2340
      TabIndex        =   1
      Top             =   105
      Width           =   3075
   End
   Begin VB.Label lblItemEnglish 
      Caption         =   "Label1"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1800
   End
   Begin VB.Label lblList 
      Caption         =   "Label1"
      Height          =   300
      Left            =   165
      TabIndex        =   4
      Top             =   1005
      Width           =   1815
   End
   Begin VB.Label lblItem 
      Caption         =   "Label1"
      Height          =   300
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   1800
   End
End
Attribute VB_Name = "frmAddGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public Event AddClick(strName As String, Cancel As Integer)
Public Event AddClick(StrName As String, Cancel As Integer, strNameEnglish As String)
Public Event UpDateClick(StrName As String, Id As Long, Cancel As Integer, strNameEnglish As String)
Public Event RemoveClick(StrName As String, Id As Long, Cancel As Integer)
Public Event CancelClick(Cancel As Boolean)
Public Event ItemSelected(Id As Long)
Private Sub SetKannadaCaption()
    
    Call SetFontToControls(Me)
    
   cmdRemove.Caption = GetResourceString(12)     'Remove
   cmdAdd.Caption = GetResourceString(10)        'Add
   cmdCancel.Caption = GetResourceString(2)      'Close

    ' Labels has to load from the Calling Function
End Sub


Private Sub cmblist_Click()

txtItem.Text = cmbList.Text
If cmbList.ListIndex > 0 Then
    cmdAdd.Caption = GetResourceString(171)
    cmdRemove.Enabled = True
    'txtItem.Text = cmbList.Text
    RaiseEvent ItemSelected(cmbList.ItemData(cmbList.ListIndex))
Else
    cmdAdd.Caption = GetResourceString(10)
    cmdRemove.Enabled = False
End If
    
If cmbList.Text <> "" Then
    cmdRemove.Enabled = True
    Me.txtItem.Text = cmbList.Text
Else
    cmdRemove.Enabled = False
    cmdAdd.Enabled = True
End If


End Sub

Private Sub cmdAdd_Click()

If Trim$(txtItem.Text) = "" Then
    MsgBox GetResourceString(688), vbOKOnly, wis_MESSAGE_TITLE
    Call ActivateTextBox(txtItem)
    Exit Sub
End If
If txtItemEnglish.Visible = True And Trim$(txtItemEnglish.Text) = "" Then
    MsgBox GetResourceString(688), vbOKOnly, wis_MESSAGE_TITLE
    ActivateTextBox txtItemEnglish
    Exit Sub
End If

If Me.cmbCumulative.Visible = True And Me.chkCumulative.Value = vbChecked And cmbCumulative.ListIndex < 1 Then
    MsgBox GetResourceString(230), vbOKOnly, wis_MESSAGE_TITLE
    ActivateTextBox cmbCumulative
    Exit Sub
End If

Dim Cancel As Integer

If gLangOffSet = 0 Then txtItemEnglish.Text = txtItem.Text

With cmbList
    If .ListIndex > 0 Then
        RaiseEvent UpDateClick(txtItem, .ItemData(.ListIndex), Cancel, txtItemEnglish.Text)
    Else
        RaiseEvent AddClick(txtItem, Cancel, txtItemEnglish.Text)
    End If
End With

If Cancel = 0 Then Me.Hide

End Sub


Private Sub cmdCancel_Click()
    RaiseEvent CancelClick(True)
    Hide
End Sub

Private Sub cmdRemove_Click()

    If Trim$(txtItem.Text) = "" Then Exit Sub
    Dim Cancel As Integer
    Dim Id As Long
    With cmbList
        If .ListIndex = -1 Then
          MsgBox "Invalid Group Specified", vbInformation, wis_MESSAGE_TITLE
          Exit Sub
        End If
        Id = IIf(.ListIndex >= 0, .ItemData(.ListIndex), 0)
    End With
    RaiseEvent RemoveClick(txtItem.Text, Id, Cancel)
    If Cancel = 0 Then Me.Hide
    
End Sub

'
Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption

If gLangOffSet = 0 Then
    Dim Ht As Integer
    Ht = txtItemEnglish.Top - txtItem.Top
    txtItemEnglish.Visible = False
    lblItemEnglish.Visible = False
    Call ReduceControlsTopPosition(Ht, lblList, cmbList, chkCumulative, cmbCumulative, cmdAdd, cmdCancel, cmdRemove)
    Me.Height = Me.Height - Ht
Else
    Call SkipFontToControls(lblItem, lblItemEnglish, txtItemEnglish)
End If

With cmbCumulative
    .Clear
    .AddItem ""
    .ItemData(.newIndex) = 0
    .AddItem GetResourceString(463) 'Monthly
    .ItemData(.newIndex) = Inst_Monthly
    .AddItem GetResourceString(413) 'Bi Monthly
    .ItemData(.newIndex) = Inst_BiMonthly
    .AddItem GetResourceString(414) 'Quarterly
    .ItemData(.newIndex) = Inst_Quartery
    .AddItem "Half Yearly" 'GetResourceString(463) 'halfe
    .ItemData(.newIndex) = Inst_HalfYearly
    .AddItem GetResourceString(208) 'Yearly
    .ItemData(.newIndex) = Inst_Yearly
    
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
 RaiseEvent CancelClick(True)

End Sub

Private Sub txtItemEnglish_GotFocus()
    Call ToggleWindowsKey(winScrlLock, False)
    Call Translate(txtItem, txtItemEnglish)
    
End Sub

Private Sub txtItemEnglish_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub
