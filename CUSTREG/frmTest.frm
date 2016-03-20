VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test form"
   ClientHeight    =   2265
   ClientLeft      =   -150
   ClientTop       =   3015
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8865
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save customer"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   1140
      Width           =   2715
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5190
      TabIndex        =   8
      Top             =   1620
      Width           =   2715
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Create new customer"
      Height          =   405
      Left            =   5190
      TabIndex        =   7
      Top             =   540
      Width           =   2715
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   720
      MouseIcon       =   "frmTest.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   1650
      Width           =   315
      Begin VB.Shape cirIn 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   2
         Left            =   150
         Shape           =   3  'Circle
         Top             =   60
         Width           =   75
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   -30
         Shape           =   3  'Circle
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmTest.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   1170
      Width           =   315
      Begin VB.Shape cirIn 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   1
         Left            =   150
         Shape           =   3  'Circle
         Top             =   60
         Width           =   75
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   -30
         Shape           =   3  'Circle
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   720
      MouseIcon       =   "frmTest.frx":0614
      MousePointer    =   99  'Custom
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   660
      Width           =   315
      Begin VB.Shape cirIn 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   0
         Left            =   150
         Shape           =   3  'Circle
         Top             =   60
         Width           =   75
      End
      Begin VB.Shape cirOut 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   -30
         Shape           =   3  'Circle
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to exit this window."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   1200
      MouseIcon       =   "frmTest.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   2085
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to test invoke the name register dialog."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1200
      MouseIcon       =   "frmTest.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1170
      Width           =   3480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to create CustReg database."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1200
      MouseIcon       =   "frmTest.frx":0F32
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   690
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Name register module testing."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5130
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_CustReg As clsCustReg




Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNew_Click()

Set m_CustReg = New clsCustReg
'MsgBox m_CustReg.CustomerID
m_CustReg.ModuleID = wis_CustReg
m_CustReg.ShowDialog
MsgBox m_CustReg.CustomerId

End Sub




Private Sub cmdSave_Click()
m_CustReg.ModuleID = wis_PDAcc

m_CustReg.LoadCustomerInfo 1

m_CustReg.SaveCustomer
End Sub


Private Sub Form_Load()

Me.Left = 0
'Me.Top = Screen.Height - Me.Height
Dim I As Byte
For I = pic.LBound To pic.UBound
    pic(I).BackColor = Me.BackColor
    cirIn(I).ZOrder 1
Next

Call Initialize
frmTest.Show

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
#If Junk Then
Dim I As Byte
For I = cirIn.LBound To cirIn.UBound
    cirIn(I).ZOrder 1
Next
#End If
End Sub

Private Sub lbl_Click(Index As Integer)
On Error GoTo Err_Line
'Dim dbObj As New clsTransact
Select Case Index
    Case 0
        If Not gDBTrans.CreateDB(App.Path & "\CustReg.ini", "wis") Then
            MsgBox "Error in creating database.", vbCritical
        End If
    Case 1
        Dim NameRegister As New clsCustReg
        With NameRegister
            ' Open the database.
            If Not gDBTrans.OpenDB(App.Path & "\CustReg.MDB", "wis") Then Exit Sub
            If Not .LoadCustomerInfo(1) Then Exit Sub
            '.Clear
            '.NewCustomer
            .ShowDialog
            
            ' Save the details.
            If .Modified Then
                If Not .SaveCustomer Then
                    MsgBox "Error saving the details.", vbCritical, wis_MESSAGE_TITLE
                    Exit Sub
                End If
            End If
        End With

    Case 2
        Unload Me
End Select

Err_Line:
    'If Err Then
    '    ' User defined error.
    '    MsgBox Err.Description, vbCritical
    'End If
'Resume
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim I As Byte
For I = pic.LBound To pic.UBound
    If I = Index Then
        cirIn(I).ZOrder
    Else
        cirIn(I).ZOrder 1
    End If
Next
End Sub


Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim I As Byte
For I = cirIn.LBound To cirIn.UBound
    If I <> Index Then
        cirIn(I).ZOrder 1
    Else
        cirIn(I).ZOrder
    End If
Next
End Sub
