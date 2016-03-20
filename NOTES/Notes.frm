VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Note"
   ClientHeight    =   3270
   ClientLeft      =   1950
   ClientTop       =   2115
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   400
      Left            =   5640
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraNotes 
      Caption         =   "Notes"
      Height          =   2385
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   6945
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   400
         Left            =   5610
         TabIndex        =   3
         Top             =   1830
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5610
         TabIndex        =   2
         Top             =   690
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   400
         Left            =   5610
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNote 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   5295
      End
      Begin VB.ListBox lstNote 
         Height          =   1620
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Frame fraProp 
      Caption         =   "Properties"
      Height          =   675
      Left            =   90
      TabIndex        =   8
      Top             =   2490
      Width           =   5445
      Begin VB.OptionButton optCustNote 
         Caption         =   "This is a customer's note"
         Height          =   315
         Left            =   2700
         TabIndex        =   5
         Top             =   240
         Width           =   2595
      End
      Begin VB.OptionButton optBankNote 
         Caption         =   "This is a bankers's note"
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkClicked()
Private Sub SetKannadaCaption()
Call SetFontToControls(Me)
'with frame Notes
fraNotes.Caption = GetResourceString(256)
cmdClear.Caption = GetResourceString(8)
cmdRemove.Caption = GetResourceString(12)
cmdAdd.Caption = GetResourceString(10)
cmdClose.Caption = GetResourceString(11)

'with fra Properties
fraProp.Caption = GetResourceString(213)
optBankNote.Caption = GetResourceString(257)
optCustNote.Caption = GetResourceString(258)

ErrLine:

End Sub


Private Sub cmdAdd_Click()
'Just add to the Note List
    If Trim$(txtNote.Text) = "" Then
        'MsgBox "Note not specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(566), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtNote
        Exit Sub
    End If
    lstNote.AddItem Trim$(txtNote.Text)
    lstNote.ItemData(lstNote.newIndex) = IIf(optBankNote.Value, 1, 2)
    txtNote.Text = ""
End Sub

Private Sub cmdClear_Click()
If Me.lstNote.ListCount <= 0 Then Exit Sub
'If MsgBox("Are you sure you want to clear all the notes associated with this account ?", vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
If MsgBox(GetResourceString(567), vbQuestion + vbYesNo, gAppName & " - Confirmation") = vbNo Then
    Exit Sub
End If
Me.lstNote.Clear
End Sub

Private Sub cmdClose_Click()
RaiseEvent OkClicked
Unload Me
End Sub

Private Sub cmdRemove_Click()
If lstNote.ListCount = 0 Then
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
    Exit Sub
End If
If lstNote.ListIndex < 0 Then
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
    Exit Sub
End If

lstNote.RemoveItem (lstNote.ListIndex)
txtNote.Text = ""
End Sub

Private Sub Form_Activate()
If Me.lstNote.ListCount > 0 Then
   cmdRemove.Enabled = True
   cmdClear.Enabled = True
Else
   cmdRemove.Enabled = False
   cmdClear.Enabled = False
End If
End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption

End Sub


Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
Set frmNotes = Nothing
End Sub

Private Sub lstNote_Click()
If lstNote.ListIndex >= 0 Then
    cmdRemove.Enabled = True
Else
    cmdRemove.Enabled = False
End If
If lstNote.ItemData(lstNote.ListIndex) = 1 Then
    optBankNote.Value = True

Else
    optCustNote.Value = True
End If
    txtNote.Text = lstNote.List(lstNote.ListIndex)
End Sub

Private Sub txtNote_Change()
If Trim$(txtNote.Text) = "" Then
    cmdAdd.Enabled = False
Else
    cmdAdd.Enabled = True
End If
End Sub


