VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2280
   ClientTop       =   1845
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1830
      TabIndex        =   2
      Top             =   1530
      Width           =   3435
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2325
      Left            =   240
      TabIndex        =   1
      Top             =   2310
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4101
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   765
      Left            =   1440
      TabIndex        =   0
      Top             =   510
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Notes As clsNotes
Set Notes = New clsNotes

Call Notes.LoadNotes(1)

Call Notes.DisplayNote(RichTextBox1)




End Sub


Private Sub Command2_Click()
Dim Notes As clsNotes
Set Notes = New clsNotes

Call Notes.LoadNotes(1)
Call Notes.Show

End Sub


