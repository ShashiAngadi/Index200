VERSION 5.00
Begin VB.Form frmPlaceCaste 
   Caption         =   "Add place or caste"
   ClientHeight    =   1740
   ClientLeft      =   1260
   ClientTop       =   4440
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5025
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   1380
      Width           =   1080
   End
   Begin VB.ComboBox cmbPlace 
      Height          =   315
      Left            =   2115
      TabIndex        =   1
      Top             =   510
      Width           =   2685
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   330
      Left            =   3735
      TabIndex        =   2
      Top             =   915
      Width           =   1050
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   930
      Width           =   1050
   End
   Begin VB.TextBox txtPlace 
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Top             =   105
      Width           =   2715
   End
   Begin VB.Label lblPlaceList 
      Caption         =   "Label1"
      Height          =   285
      Left            =   165
      TabIndex        =   6
      Top             =   525
      Width           =   1815
   End
   Begin VB.Label lblPlace 
      Caption         =   "Label1"
      Height          =   285
      Left            =   165
      TabIndex        =   5
      Top             =   105
      Width           =   1800
   End
End
Attribute VB_Name = "frmPlaceCaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AddClick(strName As String)
Public Event RemoveClick(strName As String)
Public Event CancelClick(Cancel As Boolean)
Private PlaceBool As Boolean 'true=place false=caste





Public Sub LoadCombobox(TableName As String)
Dim Rst As ADODB.Recordset

gDbTrans.SQLStmt = "Select * from " & TableName
cmbPlace.Clear
cmbPlace.AddItem ""
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    While Not Rst.EOF
        cmbPlace.AddItem FormatField(Rst(0))
        Rst.MoveNext
    Wend
End If
End Sub

Private Sub SetKannadaCaption()
On Error Resume Next
Dim ctrl As Control
    For Each ctrl In Me
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then
            ctrl.Font.Size = gFontSize
        End If
    Next
    cmdRemove.Caption = LoadResString(gLangOffSet + 12)
    cmdAdd.Caption = LoadResString(gLangOffSet + 10)
    cmdCancel.Caption = LoadResString(gLangOffSet + 2)
    
    ' Labels has to load from the Calling Function
End Sub


Private Sub cmbPlace_Click()
If cmbPlace.Text <> "" Then
cmdRemove.Enabled = True
cmdAdd.Enabled = False
Me.txtPlace.Text = cmbPlace.Text
Else
cmdRemove.Enabled = False
cmdAdd.Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()

If Trim$(txtPlace.Text) = "" Then Exit Sub
RaiseEvent AddClick(txtPlace.Text)
Unload Me

End Sub


Private Sub cmdCancel_Click()
RaiseEvent CancelClick(True)
Unload Me
End Sub

Private Sub cmdRemove_Click()
If Trim$(txtPlace.Text) = "" Then Exit Sub
    RaiseEvent RemoveClick(txtPlace.Text)
    Unload Me
End Sub


Private Sub Form_Activate()
Dim Rst As ADODB.Recordset
Dim StrPlace As String
Dim StrCaste As String
Dim strDepositname As String


Call SetKannadaCaption
strDepositname = "DepositName"
StrPlace = LoadResString(gLangOffSet + 99)
StrCaste = LoadResString(gLangOffSet + 100)
If Me.lblPlace.Caption = StrPlace Then
PlaceBool = True
StrPlace = "PlaceTab"
End If
If Me.lblPlace.Caption = StrCaste Then
PlaceBool = False
StrPlace = "CasteTab"
End If
If Me.lblPlace.Caption = strDepositname Then
PlaceBool = False
StrPlace = strDepositname
End If

gDbTrans.SQLStmt = "Select * From " & StrPlace
Me.cmbPlace.Clear
cmbPlace.AddItem ""
'toggle the command Buttons
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
Me.cmdAdd.Enabled = True
Me.cmdRemove.Enabled = False
GoTo ErrLine
End If

Do While Rst.EOF = False
Me.cmbPlace.AddItem CStr(FormatField(Rst(0)))
Rst.MoveNext
Loop
ErrLine:
End Sub


'
Private Sub Form_Load()
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

Call SetKannadaCaption

'Dim Rst As ADODB.Recordset
'Dim StrPlace As String
'Dim StrCaste As String
'
'StrPlace = LoadResString(gLangOffSet + 99)
'StrCaste = LoadResString(gLangOffSet + 100)
'If Me.lblPlace.Caption = StrPlace Then
'PlaceBool = True
'StrPlace = "PlaceTab"
'End If
'If Me.lblPlace.Caption = StrCaste Then
'PlaceBool = False
'StrPlace = "CasteTab"
'End If
'gDBTrans.SQLStmt = "Select * From " & AddQuotes(StrPlace, True)
'Me.cmbPlace.Clear
''toggle the command Buttons
'If gDBTrans.SQLFetch < 1 Then
'Me.cmdAdd.Enabled = True
'Me.cmdRemove.Enabled = False
'GoTo ErrLine
'End If
'Set Rst = gDBTrans.Rst.Clone
'Rst.MoveFirst
'
'Do While Rst.EOF = False
'Me.cmbPlace.AddItem FormatField(Rst)
'Rst.MoveNext
'Loop
'ErrLine:
End Sub


