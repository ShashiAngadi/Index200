VERSION 5.00
Begin VB.Form frmDenomination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Enter The Denomination....."
   ClientHeight    =   5730
   ClientLeft      =   2250
   ClientTop       =   1695
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5115
   Begin VB.CommandButton Cmdok 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4020
      TabIndex        =   49
      Top             =   5265
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   48
      Top             =   5265
      Width           =   960
   End
   Begin VB.Frame fraDenomination 
      Caption         =   "Denomination Details"
      Height          =   5115
      Left            =   60
      TabIndex        =   45
      Top             =   30
      Width           =   4965
      Begin VB.TextBox txt1000 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   25
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txt1000 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   120
         TabIndex        =   52
         Top             =   4350
         Width           =   4725
      End
      Begin VB.TextBox TxtExpectedAmount 
         Height          =   315
         Left            =   3210
         TabIndex        =   51
         Top             =   4590
         Width           =   1395
      End
      Begin VB.TextBox txt500 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   4
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox txt100 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   6
         Text            =   " "
         Top             =   1155
         Width           =   975
      End
      Begin VB.TextBox txt50 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   8
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox txt20 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   1830
         Width           =   975
      End
      Begin VB.TextBox txt10 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txt5 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   14
         Text            =   " "
         Top             =   2505
         Width           =   975
      End
      Begin VB.TextBox txt2 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   16
         Text            =   " "
         Top             =   2835
         Width           =   975
      End
      Begin VB.TextBox txt1 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   18
         Top             =   3180
         Width           =   975
      End
      Begin VB.TextBox txtcoin 
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   20
         Top             =   3510
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   165
         TabIndex        =   44
         Top             =   3810
         Width           =   4665
      End
      Begin VB.TextBox txttotal 
         Height          =   315
         Index           =   1
         Left            =   915
         TabIndex        =   22
         Text            =   " "
         Top             =   4020
         Width           =   1125
      End
      Begin VB.TextBox txttotal 
         Height          =   315
         Index           =   0
         Left            =   3405
         TabIndex        =   46
         Text            =   " "
         Top             =   4005
         Width           =   1095
      End
      Begin VB.TextBox txtcoin 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   43
         Top             =   3510
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   41
         Top             =   3180
         Width           =   915
      End
      Begin VB.TextBox txt2 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   39
         Text            =   " "
         Top             =   2835
         Width           =   915
      End
      Begin VB.TextBox txt5 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   37
         Text            =   " "
         Top             =   2505
         Width           =   915
      End
      Begin VB.TextBox txt10 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   35
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox txt20 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   33
         Top             =   1830
         Width           =   915
      End
      Begin VB.TextBox txt50 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   31
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txt100 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   29
         Text            =   " "
         Top             =   1155
         Width           =   915
      End
      Begin VB.TextBox txt500 
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   27
         Top             =   810
         Width           =   915
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100&0 *"
         Height          =   195
         Index           =   17
         Left            =   2910
         TabIndex        =   24
         Top             =   495
         Width           =   465
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100&0 *"
         Height          =   195
         Index           =   16
         Left            =   330
         TabIndex        =   1
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lblNetAmouint 
         Caption         =   "Net Amount :"
         Height          =   225
         Left            =   1560
         TabIndex        =   50
         Top             =   4650
         Width           =   1155
      End
      Begin VB.Label lblTitle 
         Caption         =   "Returned Amount"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   210
         Width           =   1605
      End
      Begin VB.Label lblTittle 
         Caption         =   "Submitted Amount"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&500 *"
         Height          =   195
         Index           =   15
         Left            =   435
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&100 *"
         Height          =   195
         Index           =   14
         Left            =   435
         TabIndex        =   5
         Top             =   1215
         Width           =   375
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&50 *"
         Height          =   195
         Index           =   13
         Left            =   525
         TabIndex        =   7
         Top             =   1545
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&20 *"
         Height          =   195
         Index           =   12
         Left            =   525
         TabIndex        =   9
         Top             =   1890
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&10 *"
         Height          =   195
         Index           =   11
         Left            =   525
         TabIndex        =   11
         Top             =   2220
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&5 *"
         Height          =   195
         Index           =   10
         Left            =   615
         TabIndex        =   13
         Top             =   2565
         Width           =   195
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&2 *"
         Height          =   195
         Index           =   9
         Left            =   615
         TabIndex        =   15
         Top             =   2895
         Width           =   195
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&1 *"
         Height          =   195
         Index           =   8
         Left            =   615
         TabIndex        =   17
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label lblSumOfCoin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Coins"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   19
         Top             =   3570
         Width           =   390
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total :"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   4035
         Width           =   450
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total   :"
         Height          =   195
         Index           =   0
         Left            =   2490
         TabIndex        =   47
         Top             =   4050
         Width           =   720
      End
      Begin VB.Label lblSumOfCoin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Coins"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   42
         Top             =   3555
         Width           =   390
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&1 *"
         Height          =   195
         Index           =   7
         Left            =   3195
         TabIndex        =   40
         Top             =   3225
         Width           =   195
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&2 *"
         Height          =   195
         Index           =   6
         Left            =   3195
         TabIndex        =   38
         Top             =   2880
         Width           =   195
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&5 *"
         Height          =   195
         Index           =   5
         Left            =   3195
         TabIndex        =   36
         Top             =   2550
         Width           =   195
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&10 *"
         Height          =   195
         Index           =   4
         Left            =   3105
         TabIndex        =   34
         Top             =   2205
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&20 *"
         Height          =   195
         Index           =   3
         Left            =   3105
         TabIndex        =   32
         Top             =   1875
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&50 *"
         Height          =   195
         Index           =   2
         Left            =   3105
         TabIndex        =   30
         Top             =   1530
         Width           =   285
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&100 *"
         Height          =   195
         Index           =   1
         Left            =   3015
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbldenomination 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&500 *"
         Height          =   195
         Index           =   0
         Left            =   3015
         TabIndex        =   26
         Top             =   825
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmDenomination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Function DisplayAmount(Index As Integer) As String
Dim NoOf1000 As Currency
Dim NoOf500 As Currency
Dim NoOf100 As Currency
Dim NoOf50 As Currency
Dim NoOf20 As Currency
Dim NoOf10 As Currency
Dim NoOf5 As Currency
Dim NoOf2 As Currency
Dim NoOf1 As Currency
Dim TotalCoins As Currency

NoOf1000 = CCur(1000 * Val(txt1000(Index).Text))
NoOf500 = CCur(500 * Val(txt500(Index).Text))
NoOf100 = CCur(100 * Val(txt100(Index).Text))
NoOf50 = CCur(50 * Val(txt50(Index).Text))
NoOf20 = CCur(20 * Val(txt20(Index).Text))
NoOf10 = CCur(10 * Val(txt10(Index).Text))
NoOf5 = CCur(5 * Val(txt5(Index).Text))
NoOf2 = CCur(2 * Val(txt2(Index).Text))
NoOf1 = CCur((Val(txt1(Index).Text)))
TotalCoins = CCur(Val(txtcoin(Index).Text))

DisplayAmount = CStr(NoOf1000 + NoOf500 + NoOf100 + NoOf50 + NoOf20 + NoOf10 + NoOf5 + NoOf2 _
                            + NoOf1 + TotalCoins)

End Function

Private Sub SetKannadaCaption()
Dim ctrl As Control
For Each ctrl In Me
   ctrl.FontName = gFontName
   If Not TypeOf ctrl Is ComboBox Then
      ctrl.FontSize = gFontSize
   End If
Next
Me.fraDenomination.Caption = LoadResString(gLangOffSet + 478) & " " & LoadResString(gLangOffSet + 295)
Me.cmdOK.Caption = LoadResString(gLangOffSet + 1)
Me.cmdCancel.Caption = LoadResString(gLangOffSet + 2)
Me.lblSumOfCoin(0).Caption = LoadResString(gLangOffSet + 479)
Me.lblSumOfCoin(1).Caption = LoadResString(gLangOffSet + 479)
Me.lbltotal(0).Caption = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 42)
Me.lbltotal(1).Caption = LoadResString(gLangOffSet + 52) & " " & LoadResString(gLangOffSet + 42)
End Sub



Private Function TxtLostfocus(txtBox As TextBox) As Boolean
         On Error GoTo ErrLine
         TxtLostfocus = False
        
             If txtBox.Text = " " Or txtBox.Text = "0" Or Trim(txtBox.Text) = "" Then
                 TxtLostfocus = True
                 txtBox.Text = "0"
                 Exit Function
             End If
          If Not CurrencyValidate(Trim$(txtBox.Text), True) Then
              GoTo ErrLine
         End If
         TxtLostfocus = True
         Exit Function
ErrLine:
         If Me.ActiveControl.Name <> Me.cmdCancel.Name Then
            ActivateTextBox txtBox
         End If
End Function


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
              
'        If Val(Me.txttotal.Text) < Val(Me.TxtExpectedAmount.Text) Then
'                MsgBox " Amount Returned Is less then Collected Amount  !", vbCritical, "ERROR!"
'                Exit Sub
'        End If
'        Me.txtRepayableamount = Val(Me.TxtTotalRecieved.Text) - Val(Me.TxtExpectedAmount.Text)
        
        If (CLng(txttotal(1).Text) - CLng(txttotal(0).Text)) = CLng(TxtExpectedAmount.Text) Then
               frmPDAccEx.cmdSave.Enabled = True
         Else
               'MsgBox "Amount Do Not Tally ", vbInformation, wis_MESSAGE_TITLE
               MsgBox LoadResString(gLangOffSet + 797), vbInformation, wis_MESSAGE_TITLE
               Exit Sub
         End If
        Unload Me
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Call SetKannadaCaption
frmDenomination.Caption = "Denomination "
txttotal(0).Locked = True
txttotal(1).Locked = True
txttotal(0).Text = "0"
txttotal(1).Text = "0"
txt500(0).Text = "0"
txt100(0).Text = "0"
txt50(0).Text = "0"
Me.txt20(0).Text = "0"
Me.txt10(0).Text = "0"
Me.txt5(0).Text = "0"
Me.txt2(0).Text = "0"
Me.txt1(0).Text = "0"
Me.txtcoin(0).Text = "0"

txt500(1).Text = "0"
txt100(1).Text = "0"
txt50(1).Text = "0"
Me.txt20(1).Text = "0"
Me.txt10(1).Text = "0"
Me.txt5(1).Text = "0"
Me.txt2(1).Text = "0"
Me.txt1(1).Text = "0"
Me.txtcoin(1).Text = "0"

Me.txttotal(0).TabStop = False
Me.txttotal(1).TabStop = False
Me.TxtExpectedAmount.Locked = True
Me.cmdOK.Enabled = False
End Sub







Private Sub txt1_Change(Index As Integer)

txttotal(Index).Text = DisplayAmount(Index)

End Sub

Private Sub txt1_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub



Private Sub txt1_LostFocus(Index As Integer)
  If Not TxtLostfocus(txt1(Index)) Then Exit Sub
End Sub


Private Sub txt10_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)

End Sub

Private Sub txt10_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub

Private Sub txt10_LostFocus(Index As Integer)
  If Not TxtLostfocus(txt10(Index)) Then Exit Sub
End Sub


Private Sub txt100_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt100_LostFocus(Index As Integer)
If Not TxtLostfocus(txt100(Index)) Then Exit Sub
End Sub
Private Sub txt100_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txt1000_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt1000_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)

End Sub


Private Sub txt2_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt2_LostFocus(Index As Integer)
    If Not TxtLostfocus(txt2(Index)) Then Exit Sub
End Sub
Private Sub txt2_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txt20_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt20_LostFocus(Index As Integer)
  If Not TxtLostfocus(txt20(Index)) Then Exit Sub
End Sub

Private Sub txt20_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txt5_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt5_LostFocus(Index As Integer)
    If Not TxtLostfocus(txt5(Index)) Then Exit Sub
End Sub
Private Sub txt5_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txt50_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt50_LostFocus(Index As Integer)
    If Not TxtLostfocus(txt50(Index)) Then Exit Sub
End Sub
Private Sub txt50_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txt500_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txt500_LostFocus(Index As Integer)
   If Not TxtLostfocus(txt500(Index)) Then Exit Sub
End Sub

Private Sub txt500_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtcoin_Change(Index As Integer)
txttotal(Index).Text = DisplayAmount(Index)
End Sub

Private Sub txtcoin_GotFocus(Index As Integer)
   Me.ActiveControl.SelStart = 0
   Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtcoin_LostFocus(Index As Integer)
  If Not TxtLostfocus(txtcoin(Index)) Then Exit Sub
txttotal(Index).Text = DisplayAmount(Index)
End Sub




Private Sub txttotal_Change(Index As Integer)
If (Val(txttotal(1).Text) - Val(txttotal(0).Text)) = Val(TxtExpectedAmount.Text) Then
   cmdOK.Enabled = True
Else
   cmdOK.Enabled = False
End If
End Sub
