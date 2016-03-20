VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   Caption         =   "SB Account - Search wizard"
   ClientHeight    =   4470
   ClientLeft      =   1320
   ClientTop       =   1740
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7020
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   210
      TabIndex        =   9
      Top             =   1650
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5490
      TabIndex        =   6
      Top             =   270
      Width           =   1200
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "&Transactions..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5490
      TabIndex        =   7
      Top             =   720
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   375
      Left            =   5490
      TabIndex        =   8
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Frame fraFind 
      Caption         =   "Find Account "
      Height          =   1365
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   5055
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   3735
         TabIndex        =   3
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find Now"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3735
         TabIndex        =   2
         Top             =   225
         Width           =   1200
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   315
         TabIndex        =   1
         Top             =   390
         Width           =   3345
      End
      Begin VB.OptionButton optAcno 
         Caption         =   "By account number"
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   810
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton optAccName 
         Caption         =   "By name"
         Height          =   255
         Left            =   2670
         TabIndex        =   5
         Top             =   825
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Define the Events that will be raised by this object.
Public Event FindClick(FindType As Integer, strSearch As String)
Public Event TransClick(strAcno As String)
Public Event SelectClick(strAcno As Long)
Private Sub SetKannadaCaption()
Dim ctrl As Control
#If Kannada Then
#Else
Exit Sub
#End If
For Each ctrl In Me
    ctrl.Font.Name = gFontName
    If TypeOf ctrl Is ComboBox Then
        ctrl.Font.Size = gFontSize
    End If
Next
'General form
Me.fraFind.Caption = "≤Ò¿˙ÆıÙ¬Ù· ŒÙΩÙ∞Ù"
Me.cmdFind.Caption = "ŒÙΩÙ∞Ù"
Me.cmdClear.Caption = "ÓÁÆıÙ«"
Me.cmdSelect.Caption = "¢êÕÙ"
Me.cmdTrans.Caption = "∆Â∆ŒÒ«"
Me.cmdClose.Caption = "∆ÙÙ∂Ù’"
Me.optAcno.Caption = "≤Ò¿˙ ÕÆ≤˙ÂèÙÆ¡"
Me.optAccName.Caption = "Œ˙ÕêãÆ¡"
End Sub
Private Sub cmdClear_Click()
txtFind.Text = ""
ListView1.ListItems.Clear
cmdSelect.Enabled = False
cmdTrans.Enabled = False
End Sub

Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdFind_Click()
' Check if valid data is entered.
If optAcno.value Then
    RaiseEvent FindClick(wis_QUERY_BY_CUSTOMERID, txtFind.Text)
Else
    RaiseEvent FindClick(wis_QUERY_BY_CUSTOMERNAME, txtFind.Text)
End If
End Sub
Private Sub cmdSelect_Click()
Me.Hide
RaiseEvent SelectClick(ListView1.SelectedItem.Text)
End Sub
Private Sub cmdTrans_Click()
RaiseEvent TransClick(ListView1.SelectedItem.Text)
End Sub

Private Sub Form_Load()
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption

End Sub
Private Sub Form_Resize()
On Error Resume Next

' Align all the controls of this form
' according to the new dimensions.
Const MARGIN = 200
Const CTL_MARGIN = 50

' Specify min_width and min_height.
Dim min_height As Long
Dim min_width As Long
min_width = cmdSelect.Width + optAcno.Width + optAccName.Width _
        + cmdFind.Width + MARGIN * 5 + 4 * CTL_MARGIN
If Me.Width < min_width Then Me.Width = min_width
min_height = fraFind.Height + 2 * MARGIN _
            + 2 * CTL_MARGIN + 1800
If Me.Height < min_height Then
    Me.Height = min_height
End If

' Set the position of command buttons.
With cmdSelect
    .Top = MARGIN + CTL_MARGIN * 2
    .Left = Me.ScaleWidth - MARGIN - .Width
End With
With cmdTrans
    .Top = cmdSelect.Top + cmdSelect.Height + CTL_MARGIN
    .Left = cmdSelect.Left
End With
With cmdClose
    .Top = cmdTrans.Top + cmdTrans.Height + CTL_MARGIN
    .Left = cmdSelect.Left
End With

' Align the Frame.
With fraFind
    .Left = MARGIN
    .Top = MARGIN
    .Width = Me.ScaleWidth - cmdSelect.Width _
            - CTL_MARGIN * 2 - MARGIN * 2
End With

' Align the controls within the frame...
With cmdFind
    .Left = fraFind.Width - MARGIN - .Width
End With
With cmdClear
    .Left = cmdFind.Left
End With
With txtFind
    .Left = MARGIN
    .Width = fraFind.Width - cmdFind.Width - 3 * MARGIN
End With
optAcno.Left = MARGIN
optAccName.Left = txtFind.Left + txtFind.Width - optAccName.Width
'Label1.Left = MARGIN

' Align the list view control.
With ListView1
    .Left = MARGIN
    .Top = fraFind.Top + fraFind.Height + CTL_MARGIN * 2
    .Width = Me.ScaleWidth - 2 * MARGIN
    .Height = Me.ScaleHeight - fraFind.Height _
                - 2 * MARGIN - 2 * CTL_MARGIN
End With

End Sub

Private Sub ListView1_DblClick()
cmdSelect_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
' Enable the Transactions button.
cmdTrans.Enabled = True
cmdSelect.Enabled = True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdSelect_Click
End If
End Sub

Private Sub txtFind_Change()
If Trim$(txtFind.Text) = "" Then
    cmdFind.Enabled = False
Else
    cmdFind.Enabled = True
End If
End Sub
