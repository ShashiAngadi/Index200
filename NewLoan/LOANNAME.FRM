VERSION 5.00
Begin VB.Form frmLoanName 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Define Installment Name"
   ClientHeight    =   1845
   ClientLeft      =   2595
   ClientTop       =   3105
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLoanAmt 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   300
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   850
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      Left            =   1980
      TabIndex        =   1
      Top             =   1080
      Width           =   850
   End
   Begin VB.TextBox txtInstName 
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   1785
   End
   Begin VB.Label lblAmt 
      Caption         =   "Installment Amount"
      Height          =   225
      Left            =   2040
      TabIndex        =   5
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label lblInstName 
      Caption         =   "Installment Name"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   1545
   End
End
Attribute VB_Name = "frmLoanName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Status As String
Public P_InstalmentDetails As String
Public P_TxtCount As Integer
Private Sub SetKannadaCaption()

'Declare variables
Dim Ctrl As Control
On Error Resume Next
For Each Ctrl In Me
Ctrl.Font.Name = gFontName
If Not TypeOf Ctrl Is ComboBox Then
    Ctrl.Font.Size = gFontSize
End If
Next
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)        '"«Œâ³ú"
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)  '"ÇðÁðôà"
Me.lblInstName.Caption = LoadResString(gLangOffSet + 57) & " " & LoadResString(gLangOffSet + 35)
Me.lblAmt.Caption = LoadResString(gLangOffSet + 57) & " " & LoadResString(gLangOffSet + 40)



End Sub
Private Sub cmdCancel_Click()
P_InstalmentDetails = ""
Unload Me
End Sub


Private Sub cmdOK_Click()
Dim Count As Integer
P_InstalmentDetails = ""
For Count = 0 To txtInstName.Count - 1
    If txtInstName(Count).Text = "" Then
        
       'MsgBox "You not save specified the Loan Component name ", vbExclamation, wis_MESSAGE_TITLE
       MsgBox LoadResString(gLangOffSet + 770), vbExclamation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    If Val(txtLoanAmt(Count).Text) <= 0 Then
        txtLoanAmt(Count).Text = "0.00"
        If MsgBox(LoadResString(gLangOffSet + 506) & vbCrLf & LoadResString(gLangOffSet + 541), vbExclamation + vbYesNo, wis_MESSAGE_TITLE) = vbNo Then
            Exit Sub
        End If
    End If
    P_InstalmentDetails = P_InstalmentDetails & txtInstName(Count).Text & ";" & txtLoanAmt(Count).Text & ";"
Next

P_InstalmentDetails = Left(P_InstalmentDetails, Len(P_InstalmentDetails) - 1)
Unload Me
End Sub

Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'Declare variables
Dim Count As Integer
Dim MARGIN As Integer
Dim InstDet() As String
Dim J As Integer

    MARGIN = 25
    For J = 1 To P_TxtCount - 1
        On Error Resume Next
        Load txtInstName(J)
        Load txtLoanAmt(J)
        txtInstName(J).Width = txtInstName(J - 1).Width
        txtInstName(J).Height = txtInstName(J - 1).Height
        txtInstName(J).Top = txtInstName(J - 1).Top + MARGIN + txtInstName(J).Height
        txtInstName(J).Left = txtInstName(J - 1).Left
        txtInstName(J).Visible = True
        txtInstName(J).TabIndex = txtLoanAmt(J - 1).TabIndex + 1
        
        txtLoanAmt(J).Width = txtLoanAmt(J - 1).Width
        txtLoanAmt(J).Height = txtLoanAmt(J - 1).Height
        txtLoanAmt(J).Top = txtLoanAmt(J - 1).Top + MARGIN + txtLoanAmt(J).Height
        txtLoanAmt(J).Left = txtLoanAmt(J - 1).Left
        txtLoanAmt(J).Visible = True
        txtLoanAmt(J).TabIndex = txtInstName(J).TabIndex + 1
    Next
        cmdOK.Left = Me.Width - cmdOK.Width - 150
        cmdOK.Top = txtInstName(J - 1).Top + txtInstName(J - 1).Height + 100
        cmdOK.Visible = True
        cmdCancel.Left = Me.Width - (cmdOK.Width + cmdCancel.Width + 350)
        cmdCancel.Top = cmdOK.Top 'txtInstName(J - 1).Top + txtInstName(J).Height + 100
        cmdCancel.Visible = True
        Me.Height = cmdOK.Top + cmdOK.Height * 1.5 + 250
'Now set the Values to the text box
On Error Resume Next

'Set kannada Caption for the controls
Call SetKannadaCaption


ExitLine:
' centre the form
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

