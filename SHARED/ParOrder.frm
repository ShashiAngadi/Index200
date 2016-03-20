VERSION 5.00
Begin VB.Form frmParentOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set the Printing Order"
   ClientHeight    =   3810
   ClientLeft      =   2250
   ClientTop       =   3015
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLedger 
      Height          =   2310
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   780
      Width           =   4635
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   4200
      TabIndex        =   6
      Top             =   3390
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   400
      Left            =   2790
      TabIndex        =   5
      Top             =   3390
      Width           =   1215
   End
   Begin VB.ComboBox cmbAccountType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Text            =   "Account Types"
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton cmdDown 
      Height          =   525
      Left            =   4800
      Picture         =   "ParOrder.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1590
      Width           =   585
   End
   Begin VB.CommandButton cmdUP 
      Height          =   525
      Left            =   4800
      Picture         =   "ParOrder.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Select Account"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   5340
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmParentOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub LoadAccountType()

Dim AccountType As wis_AccountType

With cmbAccountType
    .Clear
    
    AccountType = Asset
    .AddItem ("Asset")
    .ItemData(.newIndex) = AccountType
    
    AccountType = Liability
    .AddItem ("Liability")
    .ItemData(.newIndex) = AccountType
    
    AccountType = Loss
    .AddItem ("Expenses")
    .ItemData(.newIndex) = AccountType
    
    AccountType = Profit
    .AddItem ("Income")
    .ItemData(.newIndex) = AccountType
    
    AccountType = ItemPurchase
    .AddItem ("Purchase")
    .ItemData(.newIndex) = AccountType
    
    AccountType = ItemSales
    .AddItem ("Sales")
    .ItemData(.newIndex) = AccountType
End With

End Sub

' This Will Save the Data

Private Sub SaveData()

On Error GoTo Hell:

Dim ListCount As Long
Dim Item As Long
Dim ParentID As Long
Dim PrintOrder As Long
Dim PrintStatus As wis_PrintStatus

gDbTrans.BeginTrans

With lstLedger

    ListCount = .ListCount
        
    For Item = 0 To ListCount - 1
        
        PrintStatus = NoPrintDetailed
        
        If .Selected(Item) Then PrintStatus = PrintDetailed
               
        ParentID = .ItemData(Item)
        
        PrintOrder = Item + 1
        
        gDbTrans.SqlStmt = " UPDATE ParentHeads " & _
                           " SET PrintOrder=" & PrintOrder & "," & _
                           " PrintDetailed=" & PrintStatus & _
                           " WHERE ParentID=" & ParentID
                         
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    Next Item
    
End With

gDbTrans.CommitTrans

MsgBox "Updated the Records!"

Exit Sub

Hell:

    MsgBox "SaveData : " & vbCrLf & Err.Description
    
End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

End Sub

Private Sub cmbAccountType_Click()

LoadParentHeadsToListView
End Sub

Private Sub LoadParentHeadsToListView()

Dim rstParentHeads As ADODB.Recordset
Dim AccountType As wis_AccountType
Dim PrintStatus As wis_PrintStatus

Dim fldParentID As ADODB.Field
'ParentName,PrintOrder,PrintDetailed
Dim fldParentName As ADODB.Field
Dim fldPrintDetailed As ADODB.Field

If cmbAccountType.ListIndex = -1 Then Exit Sub

AccountType = cmbAccountType.ItemData(cmbAccountType.ListIndex)

gDbTrans.SqlStmt = " SELECT ParentID,ParentName,PrintOrder,PrintDetailed" & _
                   " FROM ParentHeads" & _
                   " WHERE AccountType=" & AccountType & _
                   " ORDER BY PrintOrder"
                                   
If gDbTrans.Fetch(rstParentHeads, adOpenForwardOnly) < 0 Then Exit Sub

Set fldParentID = rstParentHeads.Fields("ParentID")
Set fldParentName = rstParentHeads.Fields("ParentName")
Set fldPrintDetailed = rstParentHeads.Fields("PrintDetailed")

With lstLedger
    
    .Clear
    
    Do While Not rstParentHeads.EOF
        
        PrintStatus = fldPrintDetailed.Value
        
        .AddItem rstParentHeads("ParentName")
        .ItemData(.newIndex) = rstParentHeads("ParentID")
        
        If PrintStatus = PrintDetailed Then .Selected(.newIndex) = True
        
        rstParentHeads.MoveNext
    Loop
    
End With
 

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmdDown_Click()

  On Error Resume Next
  Dim nItem As Integer
  Dim ItemData As Long
  
  With lstLedger
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    
    ItemData = .ItemData(nItem)
    
    'move item down
    .AddItem .Text, nItem + 2
    .ItemData(nItem + 2) = ItemData
    'remove old item
    .RemoveItem nItem
    'select the item that was just moved
    .Selected(nItem + 1) = True
  End With

End Sub

Private Sub cmdSave_Click()


SaveData



End Sub

Private Sub cmdUP_Click()
  On Error Resume Next
  Dim nItem As Integer
  Dim ItemData As Long
    
  With lstLedger
  
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    
    ItemData = .ItemData(nItem)
  
    
    'move item up
    .AddItem .Text, nItem - 1
    .ItemData(nItem - 1) = ItemData
  
    
    'remove old item
    .RemoveItem nItem + 1
    'select the item that was just moved
    .Selected(nItem - 1) = True
    
  End With

End Sub

Private Sub Form_Load()

'Center the form
CenterMe Me
'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

' This will load the Account Types
If gLangOffSet <> 0 Then SetKannadaCaption
LoadAccountType
End Sub


Private Sub Form_Unload(Cancel As Integer)
If gWindowHandle = Me.hwnd Then gWindowHandle = 0
End Sub


