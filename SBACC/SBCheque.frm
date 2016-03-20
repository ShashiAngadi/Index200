VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSBCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques"
   ClientHeight    =   2775
   ClientLeft      =   4290
   ClientTop       =   1860
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4380
      TabIndex        =   7
      Top             =   2310
      Width           =   1290
   End
   Begin VB.Frame fraRemove 
      Caption         =   "Current leaf set"
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   540
      Width           =   5295
      Begin VB.CommandButton cmdInvert 
         Caption         =   "Invert Selection"
         Height          =   345
         Left            =   3750
         TabIndex        =   6
         Top             =   720
         Width           =   1395
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   3750
         TabIndex        =   5
         Top             =   270
         Width           =   1395
      End
      Begin VB.ListBox lstCheque 
         Height          =   1185
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   270
         Width           =   3465
      End
   End
   Begin VB.Frame fraAdd 
      Caption         =   "Issue new cheque book"
      Height          =   1590
      Left            =   270
      TabIndex        =   4
      Top             =   540
      Width           =   5295
      Begin VB.TextBox txtLeaves 
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   2
         Top             =   780
         Width           =   705
      End
      Begin VB.TextBox txtStartNo 
         Height          =   315
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   1
         Top             =   330
         Width           =   1995
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Add"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   3705
         TabIndex        =   3
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblLeaves 
         Caption         =   "No of leaves:"
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label lblStartNo 
         Caption         =   "Start no:"
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   1245
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   3836
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add cheque book"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Add new cheque book"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Remove cheque leaves"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSBCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OKClicked(Series As String, StartNo As Long, Leaves As Long, Cancel As Boolean)
Public Event RemoveLeaves(LeafArr() As Long, Opeation As wis_ChequeTrans)
Private Function FillChequeList()
Dim i As Integer
    lstCheque.Clear
    With frmSBAcc
        For i = 1 To .cmbCheque.ListCount - 1
            lstCheque.AddItem .cmbCheque.List(i)
        Next i
    End With

End Function

Private Sub SetKannadaCaption()
Dim Ctrl As Control
    For Each Ctrl In Me
        Ctrl.Font.Name = gFontName
        If Not TypeOf Ctrl Is ComboBox Then
            Ctrl.Font.Size = gFontSize
        End If
    Next
'with general form.
TabStrip1.Tabs(1).Caption = LoadResString(gLangOffSet + 220)
TabStrip1.Tabs(2).Caption = LoadResString(gLangOffSet + 221)
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 11)   '"∆ÙÙ∂Ù’"
'with tabstrip1 or frame add
Me.fraAdd.Caption = LoadResString(gLangOffSet + 141)    '"Œ˙˜Õ ∂˙∞Ù– √ˆÕﬁ∞ "
Me.lblStartNo.Caption = LoadResString(gLangOffSet + 144)  '"¢«Æ≈˛¡ ÕÆ≤˙Â"
Me.lblLeaves.Caption = LoadResString(gLangOffSet + 145)     '"∂˙∞˝ ®…˙≥  ÕÆ≤˙Â"
Me.cmdOK.Caption = LoadResString(gLangOffSet + 10)     '"Õ˙ÛêÕÙ"
'with tabstrip2 or Frame Remove
fraRemove.Caption = LoadResString(gLangOffSet + 142)   '"∂ÒëﬁèÙ«Ù∆ ∂˙∞Ù–≥ Ù"
Me.cmdRemove.Caption = LoadResString(gLangOffSet + 12)    '"¿˙≥˙"
Me.cmdInvert.Caption = LoadResString(gLangOffSet + 21)    '"•…Ò⁄ Õ˙…˙∞Î¬˝ "
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInvert_Click()
Dim i As Integer
For i = 0 To lstCheque.ListCount - 1
    lstCheque.Selected(i) = Not lstCheque.Selected(i)
Next i


End Sub

Private Sub cmdOK_Click()

Dim LeafCount As Long
Dim Cancel As Boolean

'Validate the Cheque Start No
    If Val(txtStartNo.Text) <= 0 Then
        MsgBox LoadResString(gLangOffSet + 503), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtStartNo
        Exit Sub
    End If

'Validate the no of leaves
    If Val(txtLeaves.Text) <= 0 Then
        MsgBox LoadResString(gLangOffSet + 504), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtLeaves
        Exit Sub
    Else
        LeafCount = Val(txtLeaves.Text)
    End If

'Validate the number of leaves
    If LeafCount <> 1 And LeafCount <> 10 And LeafCount <> 25 And LeafCount <> 50 And LeafCount <> 100 Then
         'MsgBox "Invalid number of leaves specified !" & vbCrLf & "Number of leaves should be 10, 25, 50 , 100", vbExclamation, gAppName & " - Error"
        MsgBox LoadResString(gLangOffSet + 504), vbExclamation, gAppName & " - Error"
        ActivateTextBox txtLeaves
        Exit Sub
    End If

RaiseEvent OKClicked(Val(txtStartNo.Text), Val(txtLeaves.Text), Cancel)
Call FillChequeList

'MsgBox "Cheque Book added successfully !", vbExclamation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 637), vbExclamation, gAppName & " - Error"
txtStartNo.Text = ""
txtLeaves.Text = ""

End Sub

Private Sub cmdRemove_Click()
Dim LeafArr() As Long
Dim i As Integer
'Prepare the array
ReDim Preserve LeafArr(0)
For i = 0 To lstCheque.ListCount - 1
    If lstCheque.Selected(i) = True Then
        LeafArr(UBound(LeafArr)) = lstCheque.List(i)
        ReDim Preserve LeafArr(UBound(LeafArr) + 1)
    End If
Next i

If UBound(LeafArr) = 0 Then
    Exit Sub
End If
RaiseEvent RemoveLeaves(LeafArr)

'MsgBox "Specified leaves remove successfully.", vbInformation, gAppName & " - Error"
MsgBox LoadResString(gLangOffSet + 557), vbInformation, gAppName & " - Error"

Call FillChequeList

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode = vbKeyTab Then
      If Me.TabStrip1.SelectedItem.Index = TabStrip1.Tabs.Count Then
            TabStrip1.Tabs(1).Selected = True
      Else
            TabStrip1.Tabs(TabStrip1.SelectedItem.Index + 1).Selected = True
      End If
End If

End Sub

Private Sub Form_Load()
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
fraAdd.Visible = True
fraRemove.Visible = False
fraAdd.ZOrder 0

'Fill Remove list box
Call FillChequeList
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSBCheque = Nothing
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem.Index = 1 Then
        fraAdd.Visible = True
        fraRemove.Visible = False
    End If
    If TabStrip1.SelectedItem.Index = 2 Then
        fraAdd.Visible = False
        fraRemove.Visible = True
    End If

End Sub

Private Sub txtLeaves_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If

End Sub

Private Sub txtStartNo_Change()
If Trim$(txtStartNo.Text) <> "" And Trim$(txtLeaves.Text) <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
End Sub


