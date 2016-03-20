VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmReportOption 
   Caption         =   "Selection Creteria"
   ClientHeight    =   3630
   ClientLeft      =   1005
   ClientTop       =   660
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5310
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   2550
      TabIndex        =   12
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   3900
      TabIndex        =   13
      Top             =   2940
      Width           =   1215
   End
   Begin VB.ComboBox cmbGender 
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      Top             =   1740
      Width           =   3255
   End
   Begin VB.ComboBox cmbPlaces 
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.ComboBox cmbCastes 
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   1230
      Width           =   3255
   End
   Begin VB.ComboBox cmbAccGroup 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   210
      Width           =   3255
   End
   Begin WIS_Currency_Text_Box.CurrText txtStartAmt 
      Height          =   345
      Left            =   1860
      TabIndex        =   9
      Top             =   2250
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin WIS_Currency_Text_Box.CurrText txtEndAmt 
      Height          =   345
      Left            =   4380
      TabIndex        =   11
      Top             =   2250
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Tag             =   "v"
      X1              =   30
      X2              =   5500
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label lblAmt2 
      AutoSize        =   -1  'True
      Caption         =   "And :"
      Height          =   300
      Left            =   2820
      TabIndex        =   10
      Top             =   2250
      Width           =   1485
   End
   Begin VB.Label lblAmt1 
      AutoSize        =   -1  'True
      Caption         =   "Between :"
      Height          =   300
      Left            =   60
      TabIndex        =   8
      Top             =   2250
      Width           =   1320
   End
   Begin VB.Label lblGroup 
      Caption         =   "Group Name"
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   1365
   End
   Begin VB.Label lblGender 
      Caption         =   "Gender :"
      Height          =   300
      Left            =   60
      TabIndex        =   6
      Top             =   1740
      Width           =   1365
   End
   Begin VB.Label lblCaste 
      Caption         =   "Caste"
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Top             =   1230
      Width           =   1365
   End
   Begin VB.Label lblPlace 
      Caption         =   "Place"
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "frmReportOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Range As Boolean
Private m_Enable As Boolean
'Variables to Store Selectd values
Private m_GroupIndex As Byte
Private m_PlaceIndex As Byte
Private m_CasteIndex As Byte
Private m_GenderIndex As Byte
Private m_FromAmount As Double
Private m_ToAmount As Double

Public Event WindowClosed()

Public Property Let EnableAmountRange(NewValue As Boolean)
    
    If m_Range = NewValue Then Exit Property
    m_Range = NewValue
    
    txtStartAmt.Enabled = NewValue
    txtEndAmt.Enabled = NewValue
    lblAmt1.Enabled = NewValue
    lblAmt2.Enabled = NewValue
    
    txtStartAmt.BackColor = IIf(NewValue, wisWhite, wisGray)
    txtEndAmt.BackColor = IIf(NewValue, wisWhite, wisGray)
    txtStartAmt.Text = IIf(NewValue, m_FromAmount, 0)
    txtEndAmt.Text = IIf(NewValue, m_ToAmount, 0)
End Property

Public Property Let EnableControls(Enable As Boolean)
    
    If m_Enable = Enable Then Exit Property
    m_Enable = Enable
    
    With cmbCastes
        .Enabled = Enable
        .ListIndex = IIf(Enable, m_CasteIndex, 0)
        .BackColor = IIf(Enable, wisWhite, wisGray)
    End With
    lblCaste.Enabled = Enable
    
    With cmbPlaces
        .Enabled = Enable
        .ListIndex = IIf(Enable, m_PlaceIndex, 0)
        .BackColor = IIf(Enable, wisWhite, wisGray)
    End With
    
    lblPlace.Enabled = Enable
    
    With cmbGender
        .Enabled = Enable
        .ListIndex = IIf(Enable, m_GenderIndex, 0)
        .BackColor = IIf(Enable, wisWhite, wisGray)
    End With
    lblGender.Enabled = Enable
    
End Property

Public Property Get AccountGroupID() As Integer
    With cmbAccGroup
        If .ListCount = 1 Then
            AccountGroupID = 1
            Exit Property
        End If
        If .ListIndex < 0 Then
            AccountGroupID = 1
            Exit Property
        End If
        AccountGroupID = .ItemData(.ListIndex)
    End With
End Property

Public Property Get Caste() As String
    If Not cmbCastes.Enabled Then Exit Property
    Caste = cmbCastes.Text
End Property

Public Property Get FromAmount() As Double
    If txtStartAmt.Enabled Then
        FromAmount = txtStartAmt.Value
    End If
    
End Property

Public Property Get ToAmount() As Double
    If txtEndAmt.Enabled Then
        ToAmount = txtEndAmt.Value
    End If
    
End Property

Public Property Get Gender() As wis_Gender
    Gender = wisNoGender
    
    If Not cmbGender.Enabled Then Exit Property
    
    With cmbGender
        If .ListIndex >= 0 Then
            Gender = .ItemData(.ListIndex)
        End If
    End With
End Property

Public Property Get Place() As String
    If Not cmbPlaces.Enabled Then Exit Property
    Place = cmbPlaces.Text
End Property

Private Sub SetKannadaCaption()

    Call SetFontToControls(Me)
        
    lblGroup.Caption = GetResourceString(36, 157)
    lblPlace.Caption = GetResourceString(112)
    lblCaste.Caption = GetResourceString(111)
    lblGender.Caption = GetResourceString(125)
    
    lblAmt1.Caption = GetResourceString(147, 42)
    lblAmt2.Caption = GetResourceString(148, 42)
    
    cmdOk.Caption = GetResourceString(1) 'OK
    cmdCancel.Caption = GetResourceString(2)    'Cancel
    
End Sub

Private Sub cmbAccGroup_Click()
    If cmbAccGroup.ListIndex > 0 Then _
        m_GenderIndex = cmbAccGroup.ListIndex
End Sub

Private Sub cmbCastes_Click()
    If cmbCastes.ListIndex > 0 Then _
        m_CasteIndex = cmbCastes.ListIndex
End Sub

Private Sub cmbGender_Click()
    If cmbGender.ListIndex > 0 Then _
        m_GenderIndex = cmbGender.ListIndex
End Sub

Private Sub cmbPlaces_Click()
    If cmbPlaces.ListIndex > 0 Then _
        m_PlaceIndex = cmbPlaces.ListIndex

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    RaiseEvent WindowClosed
    Me.Hide
End Sub

Private Sub Form_Load()
    CenterMe Me
    Call SetKannadaCaption
    
    'In Reports TaB Load The Combos with Caste and Places  Respectively
    Call LoadCastes(cmbCastes)
    Call LoadPlaces(cmbPlaces)
    Call LoadGender(cmbGender)
    Call LoadAccountGroups(cmbAccGroup)
        
    m_Range = True
    m_Enable = True
    
    If cmbAccGroup.ListCount < 2 Then
        lblGroup.Visible = False
        Dim Gap As Single
        Gap = cmbAccGroup.Height
        lblPlace.Top = lblPlace.Top - Gap
        cmbPlaces.Top = cmbPlaces.Top - Gap
        
        lblCaste.Top = lblCaste.Top - Gap
        cmbCastes.Top = cmbCastes.Top - Gap
        
        lblGender.Top = lblGender.Top - Gap
        cmbGender.Top = cmbGender.Top - Gap
        
        lblAmt1.Top = lblAmt1.Top - Gap
        txtStartAmt.Top = txtStartAmt.Top - Gap
        lblAmt2.Top = lblAmt2.Top - Gap
        txtEndAmt.Top = txtEndAmt.Top - Gap
        
        cmdCancel.Top = cmdCancel.Top - Gap
        cmdOk.Top = cmdCancel.Top
        Line1.Y1 = Line1.Y1 - Gap
        Line1.Y2 = Line1.Y1
        Height = Height - Gap
    End If

End Sub

Private Sub txtEndAmt_Change()
    m_ToAmount = txtEndAmt
End Sub

Private Sub txtStartAmt_Click()
    m_FromAmount = txtStartAmt.Text
End Sub

