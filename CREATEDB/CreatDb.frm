VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCreatDb 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   2010
   ClientTop       =   2880
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5880
   Begin ComctlLib.ProgressBar pbrdb 
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   1740
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2190
      TabIndex        =   0
      Top             =   780
      Width           =   1305
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   315
      Left            =   3675
      TabIndex        =   5
      Top             =   810
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   5130
      TabIndex        =   4
      Top             =   420
      Width           =   420
   End
   Begin VB.TextBox txtDBPath 
      Height          =   300
      Left            =   1770
      TabIndex        =   3
      Top             =   450
      Width           =   3240
   End
   Begin VB.Label Label2 
      Caption         =   "Creating Date Base "
      Height          =   300
      Left            =   90
      TabIndex        =   6
      Top             =   1335
      Width           =   5670
   End
   Begin VB.Label Label1 
      Caption         =   "Data Base Path"
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   405
      Width           =   1515
   End
   Begin VB.Label lblDatabase 
      Caption         =   "Create Server Component DataBase"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3270
   End
End
Attribute VB_Name = "frmCreatDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_CreaDBClass As clsTransact
Attribute m_CreaDBClass.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    gCancel = True
End Sub

Private Sub cmdCreate_Click()
If UCase(cmdCreate.Caption) = "&CREATE" Then
    'Check for the DataBase Path
    If Trim(Me.txtDBPath.Text) = "" Then
        MsgBox "you have not specified the path for dataBase Creation ", vbExclamation, "PATH ERROR"
        Exit Sub
    ElseIf Dir(txtDBPath.Text, vbDirectory) = "" Then
        If MsgBox("Specified Does not exists " & vbCrLf & "Do you want to Create the path ", vbInformation + vbYesNo, _
                    "DB Path Error") = vbNo Then Exit Sub
        
        If Not MakeDirectories(txtDBPath.Text) Then
            MsgBox "Error in creating the path " & txtDBPath _
                & " for " & "DBName", vbCritical
            Exit Sub
            'GoTo dbCreate_err
        End If
    End If
  Dim DbPath As String
  DbPath = Trim$(txtDBPath.Text)
    cmdCreate.Caption = "&OK"
    If m_CreaDBClass Is Nothing Then
        Set m_CreaDBClass = New clsTransact
    End If
    ' Now Set the PathOf DataBase In tab File
    Dim strRet As String
    Dim NewStr As String
    Dim IniFile As String
    Dim I As Integer
    IniFile = DbPath & "\Indx2000.tab"
    IniFile = InputBox("Enter the Name file to create Database", "Create Access DataBase", App.Path & "\Index 2000.tab")
    If Dir(IniFile, vbNormal) = "" Then
        MsgBox "Invalid file name", vbCritical, "Create DataBase"
        End
    End If
    Do
        I = I + 1
        strRet = ReadFromIniFile("DataBases", "database" & I, IniFile)
        If strRet = "" Then Exit Do
        'Now set the DataBase Path To then Ini file
        NewStr = putToken(strRet, "DBPath", Trim(DbPath))
        Call WriteToIniFile("DataBases", "DataBase" & I, NewStr, IniFile)
    Loop
    If m_CreaDBClass.CreateDB(IniFile, "WIS!@#") Then
        Label2.Caption = "Created DataBase"
        NewStr = "dbname=Index 2000"
        If Not WriteToIniFile("DataBases", "Database" & 1, NewStr, IniFile) Then
            MsgBox ""
        End If
    Else
        Label2.Caption = "Error in Creating DataBase"
    End If

ElseIf UCase(cmdCreate.Caption) = "&CANCEL" Then
    cmdCreate.Caption = "&CREATE"
    gCancel = True
End If


End Sub

Private Sub Command1_Click()
frmpath.Show vbModal
txtDBPath = frmpath.txtPath
Unload frmpath
End Sub

Private Sub Form_Load()
pbrdb.Max = 100
pbrdb.Min = 1

End Sub

Private Sub m_CreaDBClass_CreateDBStatus(strMsg As String, CreatedDBRatio As Single)
If CreatedDBRatio > 0 And CreatedDBRatio <= 1 Then
    pbrdb.Value = pbrdb.Max * CreatedDBRatio
End If
Label2.Caption = strMsg
Me.Refresh

End Sub

