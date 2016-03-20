VERSION 5.00
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmReportOption 
   Caption         =   "Selection Creteria"
   ClientHeight    =   3075
   ClientLeft      =   2715
   ClientTop       =   2265
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4500
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   345
      Left            =   2310
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   345
      Left            =   3390
      TabIndex        =   8
      Top             =   2610
      Width           =   975
   End
   Begin VB.ComboBox cmbGender 
      Height          =   315
      Left            =   1650
      TabIndex        =   3
      Top             =   1470
      Width           =   2715
   End
   Begin VB.ComboBox cmbPlaces 
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   510
      Width           =   2715
   End
   Begin VB.ComboBox cmbCastes 
      Height          =   315
      Left            =   1650
      TabIndex        =   1
      Top             =   990
      Width           =   2715
   End
   Begin VB.ComboBox cmbAccGroup 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   30
      Width           =   2715
   End
   Begin WIS_Currency_Text_Box.CurrText txtStartAmt 
      Height          =   345
      Left            =   1650
      TabIndex        =   10
      Top             =   1950
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin WIS_Currency_Text_Box.CurrText txtEndAmt 
      Height          =   345
      Left            =   3480
      TabIndex        =   11
      Top             =   1950
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   609
      CurrencySymbol  =   ""
      TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
      NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
      FontSize        =   8.25
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Tag             =   "v"
      X1              =   4485
      X2              =   90
      Y1              =   2460
      Y2              =   2475
   End
   Begin VB.Label lblAmt2 
      AutoSize        =   -1  'True
      Caption         =   "And :"
      Height          =   255
      Left            =   2730
      TabIndex        =   13
      Top             =   2010
      Width           =   585
   End
   Begin VB.Label lblAmt1 
      AutoSize        =   -1  'True
      Caption         =   "Between :"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   2010
      Width           =   1350
   End
   Begin VB.Label lblGroup 
      Caption         =   "Group Name"
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   90
      Width           =   1395
   End
   Begin VB.Label lblGender 
      Caption         =   "Gender :"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   1530
      Width           =   1395
   End
   Begin VB.Label lblCaste 
      Caption         =   "Caste"
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1050
      Width           =   1395
   End
   Begin VB.Label lblPlace 
      Caption         =   "Place"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   1395
   End
End
Attribute VB_Name = "frmReportOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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


Public Property Let DisableAmountRange(NewValue As Boolean)
    txtStartAmt.Enabled = NewValue
    txtEndAmt.Enabled = NewValue
End Property

Public Property Get Caste() As String
    If Not cmbCastes.Enabled Then Exit Property
    Caste = cmbCastes.Text
End Property


Public Property Get ToAmount() As Double
    If txtEndAmt.Enabled Then
        ToAmount = txtEndAmt.Value
    End If
    
End Property

Public Property Get FromAmount() As Double
    If txtStartAmt.Enabled Then
        FromAmount = txtStartAmt.Value
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
        
        
    lblGroup.Caption = LoadResString(gLangOffSet + 36) & " " & LoadResString(gLangOffSet + 157)
    lblPlace.Caption = LoadResString(gLangOffSet + 112)
    lblCaste.Caption = LoadResString(gLangOffSet + 111)
    lblGender.Caption = LoadResString(gLangOffSet + 125)
    
    lblAmt1.Caption = LoadResString(gLangOffSet + 147) & " " & LoadResString(gLangOffSet + 42)
    lblAmt2.Caption = LoadResString(gLangOffSet + 148) & " " & LoadResString(gLangOffSet + 42)
    
    
    cmdOk.Caption = LoadResString(gLangOffSet + 1) 'OK
    cmdCancel.Caption = LoadResString(gLangOffSet + 2)    'Cancle
    
End Sub

Private Sub Form_Load()
    CenterMe Me
    Call SetKannadaCaption
    
    'In Reports TaB Load The Combos with Caste and Places  Respectively
    Call LoadCastes(cmbCastes)
    Call LoadPlaces(cmbPlaces)
    Call LoadGender(cmbGender)
    Call LoadAccountGroups(cmbAccGroup)
    
    If Not cmbAccGroup.Visible Then
        
        lblGroup.Visible = False
        lblPlace.Top = lblPlace.Top - 100
        cmbPlaces.Top = cmbPlaces.Top - 100
        
        lblCaste.Top = lblCaste.Top - 100
        cmbCastes.Top = cmbCastes.Top - 100
        
        lblGender.Top = lblGender.Top - 100
        cmbGender.Top = cmbGender.Top - 100
        
    End If

End Sub


