VERSION 5.00
Begin VB.Form frmAmount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   1545
   ClientLeft      =   5475
   ClientTop       =   3270
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1980
      TabIndex        =   4
      Top             =   1170
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   990
      TabIndex        =   3
      Top             =   1170
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2745
      Begin VB.TextBox txtAmount 
         Height          =   345
         Left            =   900
         TabIndex        =   1
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount:"
         Height          =   225
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Okclicked(Amount As Currency)

Private Sub SetKannadaCaption()
Dim ctrl As Control
For Each ctrl In Me
   ctrl.FontName = gFontName
   If Not TypeOf ctrl Is ComboBox Then
      ctrl.FontSize = gFontSize
   End If
Next
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)
Me.lblAmount.Caption = LoadResString(gLangOffSet + 40)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'If Trim(txtAmount.Text) = "" Or txtAmount.Text = "0" Then
If Not CurrencyValidate(txtAmount.Text, True) Then
   'MsgBox "Enter The Pigmy Amount", vbOKOnly, "ERROR"
   MsgBox LoadResString(gLangOffSet + 796), vbOKOnly, "ERROR"
   ActivateTextBox txtAmount
   Exit Sub
End If
frmPDAccEx.cmdSubmit.Enabled = True
RaiseEvent Okclicked(Val(txtAmount.Text))
Me.Hide
End Sub


Private Sub Form_Activate()
Me.txtAmount.SetFocus
Me.txtAmount.SelStart = 0
Me.txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub Form_Load()
Call SetKannadaCaption
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Dim i As Byte
Dim idx As Long
        If frmPDAccEx.lst(0).Visible = True Then
        i = 0
        Else
        i = 1
        End If
Dim Name As String
Name = frmPDAccEx.lst(i).List(frmPDAccEx.lst(i).ListIndex)
idx = frmPDAccEx.GetListIndex(Name, frmPDAccEx.lst(i))
If idx = -1 Then
Me.txtAmount.Text = ""
End If
'txtAmount.SetFocus
txtAmount.SelStart = 0
txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAmount = Nothing
End Sub


