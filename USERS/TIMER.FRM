VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   10755
   ClientTop       =   3360
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   1560
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   30
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event TimedOut(Status As Boolean)

