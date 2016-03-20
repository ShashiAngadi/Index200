VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Testing ..."
   ClientHeight    =   2580
   ClientLeft      =   2535
   ClientTop       =   2445
   ClientWidth     =   5595
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   165
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   345
      Picture         =   "frmTest.frx":0442
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim prnObj As New clsPrint
Dim db As Database
Dim rs As Recordset

' Open database.
Set db = OpenDatabase("c:\indx2000\appmain\indx2000.mdb")
Set rs = db.OpenRecordset("nametab")
With prnObj
    .ReportTemplate = "c:\mydocu~1\print.txt"
    Set .DataSource = rs
    .ReportDestination = "PREVIEW"
    .PrintReport
End With

End Sub
Private Sub Form_Load()
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

