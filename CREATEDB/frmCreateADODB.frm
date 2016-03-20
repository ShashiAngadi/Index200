VERSION 5.00
Begin VB.Form frmCreateADODB 
   Caption         =   "Creates a Access 2000 Database"
   ClientHeight    =   5325
   ClientLeft      =   2010
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCreateDB 
      Caption         =   "Create Database"
      Height          =   525
      Left            =   1770
      TabIndex        =   0
      Top             =   1170
      Width           =   1665
   End
End
Attribute VB_Name = "frmCreateADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreateDB_Click()
Dim clsADOCreate As clsNewTransact
Dim INIFilePath As String

INIFilePath = "C:\NewDB\Indx2000.tab"
Set clsADOCreate = New clsNewTransact
    If Not clsADOCreate.CreateDB(INIFilePath, "PRAGMANS") Then _
        MsgBox "Cannot create the specified Database", vbInformation, "Database creation"
Set clsADOCreate = Nothing
Unload Me
End Sub


