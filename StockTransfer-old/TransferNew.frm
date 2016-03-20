VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransferNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Transfer"
   ClientHeight    =   6090
   ClientLeft      =   300
   ClientTop       =   1875
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   8820
      TabIndex        =   24
      Top             =   5490
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   10170
      TabIndex        =   25
      Top             =   5490
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4815
      Index           =   1
      Left            =   4680
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   480
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   8493
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4815
      Index           =   0
      Left            =   4680
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   8493
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Frame fra 
      Height          =   5265
      Index           =   0
      Left            =   90
      TabIndex        =   29
      Top             =   540
      Width           =   4515
      Begin VB.CommandButton cmdProduct 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         TabIndex        =   30
         Top             =   2970
         Width           =   315
      End
      Begin VB.OptionButton optFrom 
         Caption         =   "From"
         Enabled         =   0   'False
         Height          =   405
         Left            =   1650
         TabIndex        =   2
         Top             =   690
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optTo 
         Caption         =   "To"
         Enabled         =   0   'False
         Height          =   405
         Left            =   3210
         TabIndex        =   3
         Top             =   690
         Width           =   1125
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   3210
         TabIndex        =   17
         Top             =   4710
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   690
         TabIndex        =   15
         Top             =   4710
         Width           =   1215
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1950
         TabIndex        =   16
         Top             =   4710
         Width           =   1215
      End
      Begin VB.TextBox txtTransferDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   395
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1425
      End
      Begin VB.ComboBox cmbBranch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         Top             =   1350
         Width           =   2715
      End
      Begin VB.ComboBox cmbUnit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         Top             =   2460
         Width           =   2715
      End
      Begin VB.ComboBox cmbGroup 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   7
         Top             =   1875
         Width           =   2715
      End
      Begin VB.ComboBox cmbProductName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   11
         Top             =   2970
         Width           =   2385
      End
      Begin VB.TextBox txtQuantity 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         TabIndex        =   14
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblPerUnit 
         Height          =   390
         Left            =   3300
         TabIndex        =   31
         Top             =   3510
         Width           =   1155
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   12
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Label lblTransferDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lblBranch 
         Caption         =   "Branch Name"
         Height          =   420
         Left            =   90
         TabIndex        =   4
         Top             =   1395
         Width           =   1575
      End
      Begin VB.Label lblUnit 
         Caption         =   "Unit Name"
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   2475
         Width           =   1635
      End
      Begin VB.Label lblGroup 
         Caption         =   "Group Name"
         Height          =   360
         Left            =   90
         TabIndex        =   6
         Top             =   1965
         Width           =   1635
      End
      Begin VB.Label lblProductName 
         Caption         =   "Product Name"
         Height          =   270
         Left            =   90
         TabIndex        =   10
         Top             =   3000
         Width           =   1785
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   4050
         Width           =   1605
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   4440
         Y1              =   4620
         Y2              =   4620
      End
   End
   Begin VB.PictureBox tabTransfer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   60
      ScaleHeight     =   5775
      ScaleWidth      =   4545
      TabIndex        =   27
      Top             =   60
      Width           =   4605
   End
   Begin VB.Frame fra 
      Height          =   5205
      Index           =   1
      Left            =   90
      TabIndex        =   28
      Top             =   600
      Width           =   4515
      Begin VB.TextBox txtFromDate 
         Height          =   395
         Left            =   1830
         TabIndex        =   20
         Top             =   780
         Width           =   1485
      End
      Begin VB.CommandButton cmdTab1Show 
         Caption         =   "&Show"
         Height          =   435
         Left            =   3420
         TabIndex        =   23
         Top             =   4650
         Width           =   1035
      End
      Begin VB.TextBox txtToDate 
         Height          =   395
         Left            =   1830
         TabIndex        =   22
         Top             =   1290
         Width           =   1485
      End
      Begin VB.Label lblFromDate 
         Caption         =   "From Date"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label lblToDate 
         Caption         =   "To Date"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   1350
         Width           =   1485
      End
      Begin VB.Line Line3 
         X1              =   90
         X2              =   4440
         Y1              =   4560
         Y2              =   4560
      End
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   11280
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmTransferNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For the CONSTANTS
Const TRANSFER_TRANSACTION As Integer = 0
Const TRANSFER_SHOW As Integer = 1

'To Handle Grid Functions
Private m_GrdFunctions As clsGrdFunctions

Private m_dbOperation As wis_DBOperation

Private IsNewItemCreated As Boolean

Public Event AddClicked()
Public Event UnDoClicked()
Public Event SaveClicked()
Public Event OptFromClicked()
Public Event OptToClicked()
Public Event ClearClicked()
Public Event GridShowTransferClicked()
Public Event GridTransactionClicked()
Public Event DeleteClicked()
Public Event WindowClosed()

Public Event ShowClicked(ByVal FromDateUS As String, ByVal ToDateUS As String)

Private Function CheckTransDate() As Boolean

On Error GoTo ErrLine:

Dim LastTransDate As String
Dim CurrentDate As String

CheckTransDate = False

' Get the Last TransDate for the HeadID
LastTransDate = LoadLastTransDate

If LastTransDate = "" Then Exit Function

CurrentDate = txtTransferDate.Text

If Not TextBoxDateValidate(txtTransferDate, "/", True, True) Then Exit Function

If GetSysFormatDate(CurrentDate) < GetSysFormatDate(LastTransDate) Then
    
    If MsgBox("Current Date is Smaller than Last Entered Date!" & _
        vbCrLf & "Do You Want To Continue ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Function
    Else
        CheckTransDate = True
        Exit Function
    End If
Else
    CheckTransDate = True
End If

Exit Function

ErrLine:
    MsgBox "Check TransDate :" & vbCrLf & Err.Description
    
End Function

'This sub will get the last trans date from acctrans and load it to
' m_LastTransDate
Private Function LoadLastTransDate() As String

Dim rstTransDate As ADODB.Recordset
Dim HeadID As Long

LoadLastTransDate = ""

gDbTrans.SQLStmt = " SELECT MAX(TransDate) as MaxTransDate " & _
                   " FROM Stock"
                 
Call gDbTrans.Fetch(rstTransDate, adOpenForwardOnly)

LoadLastTransDate = FormatField(rstTransDate.Fields("MaxTransDate"))

' If no transaction is made then last trans date will be first day
If LoadLastTransDate = "" Then LoadLastTransDate = FinIndianFromDate

Set rstTransDate = Nothing


End Function



Private Sub UnLoadME()
    Unload Me
End Sub






Private Function ValidateShowClicked() As Boolean

'Setup an error handler...
On Error GoTo ErrLine

If Not TextBoxDateValidate(txtFromDate, "/", True, True, True) Then Exit Function
If Not TextBoxDateValidate(txtToDate, "/", True, True, True) Then Exit Function

ValidateShowClicked = True

Exit Function

ErrLine:
    MsgBox Err.Description, vbCritical
    
End Function

Private Sub cmbBranch_Click()
cmbProductName.Clear
cmbUnit.ListIndex = -1
Dim MatClass As New clsMaterial
Dim I As Integer
I = cmbGroup.ListIndex
Call MatClass.LoadProductGroups(cmbGroup)
Set MatClass = Nothing
On Error Resume Next
If cmbGroup.ListCount = 1 Then cmbGroup.ListIndex = 0: cmbUnit.SetFocus
If cmbGroup.Locked Then cmbGroup.ListIndex = I
End Sub


Private Sub cmbGroup_Click()
If cmbUnit.ListCount = 1 Then cmbUnit.ListIndex = 0: cmbProductName.SetFocus

End Sub


Private Sub cmbManufacturer_Click()
cmbProductName.Clear
End Sub


Private Sub cmbProductName_Click()
If cmbProductName.ListIndex = -1 Then Exit Sub
LoadProductDetails
End Sub


Private Sub LoadProductDetails()
Dim lngRelationID As Long
Dim GodownID As Byte
Dim TransDate As String

Dim StockAvailable As Double

Dim Rst As ADODB.Recordset


Dim MaterialClass As clsMaterial

If cmbProductName.ListIndex = -1 Then Exit Sub
If cmbBranch.ListIndex = -1 Then Exit Sub
If Not TextBoxDateValidate(txtTransferDate, "/", True, True, True) Then Exit Sub

TransDate = txtTransferDate.Text
lngRelationID = cmbProductName.ItemData(cmbProductName.ListIndex)
GodownID = cmbBranch.ItemData(cmbBranch.ListIndex)
 
Set MaterialClass = New clsMaterial
'.GetItemClosingStock(lngRelationID, m_FromGodownID)
StockAvailable = MaterialClass.GetItemOnDateClosingStock(lngRelationID, TransDate, GodownID)

lblBalance.Caption = StockAvailable
If optTo.Value Then Exit Sub

If m_dbOperation = Insert Then
    If StockAvailable > 0 Then
        lblBalance.Caption = StockAvailable
    Else
        MsgBox "There is No Stock. Please purchse the product.", vbInformation, wis_MESSAGE_TITLE
        
        With cmbProductName
            If .ListIndex = .ListCount - 1 Then
                .ListIndex = 0
            Else
                .ListIndex = .ListIndex + 1
            End If
        End With
        
        Exit Sub
        
    End If
End If

lblPerUnit.Caption = " " & cmbUnit.Text

End Sub

Private Function LoadBranchRelationID(ByVal TheCombo As ComboBox, ByVal ToGodownID As Byte, ByVal RelationID As Long) As Long
'Declare the variables
Dim Rst As ADODB.Recordset
Dim ProductID As Long
Dim HeadID As Long
Dim GroupID As Integer
Dim UnitID As Long

Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldRelationID As ADODB.Field
   

gDbTrans.SQLStmt = " SELECT HeadID,GroupID,ProductID FROM RelationMaster " & _
                " WHERE RelationID= " & RelationID

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 0 Then Exit Function

ProductID = FormatField(Rst.Fields("ProductID"))
HeadID = FormatField(Rst.Fields("HeadID"))
GroupID = FormatField(Rst.Fields("GroupID"))



gDbTrans.SQLStmt = " SELECT ProductName,UnitName,RelationID FROM RelationMaster A," & _
                    " Products B,ProductGroup C ,Units D " & _
                    " WHERE GodownID = " & ToGodownID & _
                    " AND HeadID = " & HeadID & _
                    " AND A.GroupID = " & GroupID & _
                    " AND A.ProductID = " & ProductID & _
                    " AND A.ProductID = b.ProductID  " & _
                    " AND  A.UnitID = D.UnitID " & _
                    " AND B.GroupID= C.GroupID " & _
                    " AND A.GroupID = b.GroupID  "

TheCombo.Clear
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 0 Then Exit Function


Set fldProductName = Rst.Fields("ProductName")
Set fldUnitName = Rst.Fields("UnitName")
Set fldRelationID = Rst.Fields("RelationID")
 
 


Do While Not Rst.EOF
    TheCombo.AddItem fldProductName.Value & "  " & fldUnitName.Value
    TheCombo.ItemData(TheCombo.newIndex) = fldRelationID.Value
    
    'Move the record set
    Rst.MoveNext
Loop


End Function


Private Sub cmbProductName_GotFocus()
On Error Resume Next
cmdProduct.Enabled = False
If cmbProductName.ListCount = 0 Then
    cmdProduct.Enabled = True
    cmdProduct.SetFocus
End If
End Sub

Private Sub cmbUnit_Click()
Dim MaterialClass As clsMaterial
Dim HeadID As Long
Dim GroupID As Integer
Dim UnitID As Long
Dim GodownID As Byte

'if To Option is enabld the do not load the products because they are loaded accordingly
'If Not IsNewItemCreated And cmbProductName.ListCount Then If m_DBOperation = Insert Then If optTo.value Then Exit Sub

cmbProductName.Clear

If cmbUnit.ListIndex = -1 Then Exit Sub
If cmbBranch.ListIndex = -1 Then Exit Sub
'If cmbManufacturer.ListIndex = -1 Then Exit Sub
If cmbGroup.ListIndex = -1 Then Exit Sub

'HeadId = cmbManufacturer.ItemData(cmbManufacturer.ListIndex)
GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
UnitID = cmbUnit.ItemData(cmbUnit.ListIndex)
GodownID = cmbBranch.ItemData(cmbBranch.ListIndex)

Set MaterialClass = New clsMaterial

'Load the products to the combo
'Load the HO Products to transfer to the  selected Branches
Call MaterialClass.LoadProductsToCombo(cmbProductName, GroupID, UnitID, GodownID)

If cmbProductName.ListCount < 0 Then Call cmbProductName_GotFocus
If cmbProductName.ListCount = 1 Then cmbProductName.ListIndex = 0: txtQuantity.SetFocus
lblPerUnit.Caption = cmbUnit.Text

Set MaterialClass = Nothing

End Sub


Private Sub cmdAdd_Click()

RaiseEvent AddClicked


On Error Resume Next


End Sub


Private Sub cmdClose_Click()
Unload Me
RaiseEvent WindowClosed
End Sub

Private Sub cmdProduct_Click()
Dim GodownID As Byte
Dim HeadID As Long
Dim GroupID As Integer
Dim UnitID As Long

Dim MaxCount As Integer
Dim ItemCount As Integer

'If cmbManufacturer.ListIndex = -1 Then Exit Sub
If cmbBranch.ListIndex = -1 Then Exit Sub

'HeadId = cmbManufacturer.ItemData(cmbManufacturer.ListIndex)
GodownID = cmbBranch.ItemData(cmbBranch.ListIndex)

If cmbGroup.ListIndex >= 0 Then GroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
If cmbUnit.ListIndex >= 0 Then UnitID = cmbUnit.ItemData(cmbUnit.ListIndex)

With frmProductPropertyNew
    .GodownID = GodownID
    '.CompanyID = HeadID
    .GroupID = GroupID
    .UnitID = UnitID
'    .lblCompanyName = cmbManufacturer.Text
    
    .Show vbModal
    
End With

IsNewItemCreated = True

If optFrom.Value Then Call optFrom_Click
If optTo.Value Then Call optTo_Click

ItemCount = 0
MaxCount = 0

MaxCount = cmbUnit.ListCount - 1

For ItemCount = 0 To MaxCount
    If UnitID = cmbUnit.ItemData(ItemCount) Then
        cmbUnit.ListIndex = ItemCount
        Call cmbUnit_Click
        Exit For
    End If
Next ItemCount

On Error Resume Next
cmbProductName.SetFocus

End Sub

Private Sub cmdRefresh_Click()
RaiseEvent ClearClicked

optFrom.Enabled = True
optFrom.Value = True
txtTransferDate.Locked = False
tabTransfer.Tabs(1).Selected = True
txtTransferDate.SetFocus

IsNewItemCreated = False

End Sub

Private Sub cmdSave_Click()

'Save the data
RaiseEvent SaveClicked

'Refresh the form
RaiseEvent ClearClicked

txtTransferDate.Locked = False

Call cmdRefresh_Click

'DayBeginDate = txtTransferDate.Text

IsNewItemCreated = False


End Sub


Private Sub cmdTab1Show_Click()
Dim FromDateUS As String
Dim ToDateUS As String

If Not ValidateShowClicked Then Exit Sub

Call InitGridForTab1

FromDateUS = GetSysFormatDate(txtFromDate.Text)
ToDateUS = GetSysFormatDate(txtToDate.Text)
If DateDiff("d", ToDateUS, FromDateUS) > 0 Then
    MsgBox "Start Date Should be Earlier than the end Date", vbInformation
    ActivateTextBox txtFromDate
    Exit Sub
End If

RaiseEvent ShowClicked(FromDateUS, ToDateUS)


End Sub

Private Sub cmdUndo_Click()
RaiseEvent UnDoClicked

On Error Resume Next
cmbBranch.SetFocus

txtTransferDate.Locked = False

End Sub

Private Sub Form_Initialize()

If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions


End Sub
Private Sub LoadProducts(ByVal HeadID As Long)
'Declare the variables
Dim lngGroupID As Long
Dim lngUnitID As Long
Dim rstRelation As ADODB.Recordset

Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldRelationID As ADODB.Field

cmbProductName.Clear

If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbUnit.ListIndex = -1 Then Exit Sub

'Get the respective Ids from Item data
lngGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)
lngUnitID = cmbUnit.ItemData(cmbUnit.ListIndex)

cmbProductName.Clear

gDbTrans.SQLStmt = " SELECT B.ProductName,RelationID,UnitName" & _
                " FROM RelationMaster A, Products B,Units C" & _
                " WHERE HeadID = " & HeadID & _
                " AND A.GroupID = " & lngGroupID & _
                " AND A.UnitID = " & lngUnitID & _
                " AND A.ProductID = B.ProductID" & _
                " AND A.UnitId=C.UnitID" & _
                " ORDER BY ProductName"

If gDbTrans.Fetch(rstRelation, adOpenForwardOnly) < 1 Then Exit Sub

Set fldProductName = rstRelation.Fields("ProductName")
Set fldRelationID = rstRelation.Fields("RelationID")
Set fldUnitName = rstRelation.Fields("UnitName")

'Load the data to the combo box
Do While Not rstRelation.EOF
   
   cmbProductName.AddItem fldProductName.Value & " " & fldUnitName.Value
   cmbProductName.ItemData(cmbProductName.newIndex) = fldRelationID.Value
   
   'Move the recordset
   rstRelation.MoveNext
Loop

End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
    
tabTransfer.Font = gFontName
tabTransfer.Font.Size = gFontSize
tabTransfer.Tabs(1).Caption = LoadResString(gLangOffSet + 38)
tabTransfer.Tabs(2).Caption = LoadResString(gLangOffSet + 38) & " " & LoadResString(gLangOffSet + 295)

lblTransferDate.Caption = LoadResString(gLangOffSet + 37)

optFrom.Caption = LoadResString(gLangOffSet + 107)
optTo.Caption = LoadResString(gLangOffSet + 108)

lblBranch.Caption = LoadResString(gLangOffSet + 210)
'lblManufacturer.Caption = LoadResString(gLangOffSet + 174)
lblGroup.Caption = LoadResString(gLangOffSet + 157) & " " & LoadResString(gLangOffSet + 27)
lblUnit.Caption = LoadResString(gLangOffSet + 161) & " " & LoadResString(gLangOffSet + 27)
lblProductName.Caption = LoadResString(gLangOffSet + 158) & " " & LoadResString(gLangOffSet + 35) _
                         & " " & LoadResString(gLangOffSet + 27)
lblQuantity.Caption = LoadResString(gLangOffSet + 306)
lblProductName.Caption = LoadResString(gLangOffSet + 215)
lblQuantity.Caption = LoadResString(gLangOffSet + 306)
  
cmdUndo.Caption = LoadResString(gLangOffSet + 19)
cmdSave.Caption = LoadResString(gLangOffSet + 7)
cmdAdd.Caption = LoadResString(gLangOffSet + 10)
cmdClose.Caption = LoadResString(gLangOffSet + 11) 'Close
cmdRefresh.Caption = LoadResString(gLangOffSet + 32)

cmdTab1Show.Caption = LoadResString(gLangOffSet + 13)

lblFromDate.Caption = LoadResString(gLangOffSet + 109)
lblToDate.Caption = LoadResString(gLangOffSet + 110)
 
End Sub



Private Sub GridResize(ByVal TabIndex As Integer)
Dim Ratio  As Single

Select Case TabIndex
    Case 1
        With grd(TRANSFER_TRANSACTION)
           Ratio = .Width / .Cols
           
           .ColWidth(0) = Ratio * 1.2
           .ColWidth(1) = Ratio * 1.25
           .ColWidth(2) = Ratio * 0.7
           .ColWidth(3) = Ratio * 0.8
        End With
    Case 2
        With grd(TRANSFER_SHOW)
           Ratio = .Width / .Cols
           
           .ColWidth(0) = Ratio * 0.3
           .ColWidth(1) = Ratio * 1.05
           .ColWidth(2) = Ratio * 1.5
           .ColWidth(3) = Ratio * 1.5
           .ColWidth(4) = Ratio * 0.6
           .ColWidth(5) = Ratio * 0.95
        End With
End Select
End Sub



Public Sub InitGridForTab1()

With grd(TRANSFER_SHOW)
    .Clear
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 6
    .FixedRows = 1
    .FixedCols = 0
    .Row = 0
    
    .Col = 0: .Text = LoadResString(gLangOffSet + 33): .CellFontBold = True 'Sl
    .Col = 1: .Text = LoadResString(gLangOffSet + 37): .CellFontBold = True
    .Col = 2: .Text = LoadResString(gLangOffSet + 227): .CellFontBold = True
    .Col = 3: .Text = LoadResString(gLangOffSet + 295): .CellFontBold = True
    .Col = 4: .Text = LoadResString(gLangOffSet + 306): .CellFontBold = True
    .Col = 5: .Text = LoadResString(gLangOffSet + 42): .CellFontBold = True
End With

GridResize (TRANSFER_SHOW)

End Sub


Public Sub InitGridForTab0()

With grd(TRANSFER_TRANSACTION)
    .Clear
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 4
    .FixedRows = 1
    .Row = 0
    .Col = 0: .Text = LoadResString(gLangOffSet + 227): .CellFontBold = True  '"Branch Name"
    .Col = 1: .Text = LoadResString(gLangOffSet + 295): .CellFontBold = True '"Description"
    .Col = 2: .Text = LoadResString(gLangOffSet + 306): .CellFontBold = True 'Quantity
    .Col = 3: .Text = LoadResString(gLangOffSet + 42): .CellFontBold = True 'Balance
End With

GridResize (TRANSFER_TRANSACTION)
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyF3 Then frmDateChange.Show    'vbModal

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
Dim IndxNo As Integer

CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub

If KeyCode = vbKeyTab Then
    IndxNo = Me.tabTransfer.SelectedItem.Index
    IndxNo = IndxNo + IIf(Shift = 2, 1, -1)
    If IndxNo < 1 Then IndxNo = tabTransfer.Tabs.Count
    If IndxNo > tabTransfer.Tabs.Count Then IndxNo = 1
    tabTransfer.Tabs(IndxNo).Selected = True
End If
End Sub

Private Sub Form_Load()
'
CenterMe Me

Me.Icon = LoadResPicture(147, vbResIcon)

If gLangOffSet <> 0 Then SetKannadaCaption
tabTransfer.Tabs(1).Selected = True

optFrom.Enabled = True
Dim strDate As String
strDate = gStrDate
txtTransferDate.Text = strDate
txtFromDate.Text = strDate
txtToDate.Text = strDate

InitGridForTab0
InitGridForTab1
optFrom_Click

m_dbOperation = Insert


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set m_GrdFunctions = Nothing

End Sub

Private Sub Form_Resize()
GridResize (1)
GridResize (2)

cmdProduct.Height = cmbProductName.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmAccTrans = Nothing

End Sub




Public Property Get DBOperation() As wis_DBOperation

    DBOperation = m_dbOperation
    
End Property

Public Property Let DBOperation(ByVal NewValue As wis_DBOperation)
    m_dbOperation = NewValue
End Property




Private Sub grd_DblClick(Index As Integer)

If Index = TRANSFER_SHOW Then
    RaiseEvent GridShowTransferClicked
    'LoadDetailsToTransactionGrid
    'Disable the Add button
    cmdAdd.Enabled = False
    grd(TRANSFER_TRANSACTION).Enabled = True
ElseIf Index = TRANSFER_TRANSACTION Then
    RaiseEvent GridTransactionClicked
    'LoadGridSelectedDataToCombo
    
End If


End Sub


Private Sub grd_KeyPress(Index As Integer, KeyAscii As Integer)

If tabTransfer.Tabs(1).Selected And KeyAscii = 4 Then
    If MsgBox("Are you sure Do you want to Delete the Record", vbYesNo + vbDefaultButton2 + vbInformation, wis_MESSAGE_TITLE) = vbNo Then Exit Sub
    
    RaiseEvent DeleteClicked
    
End If

End Sub


Private Sub optFrom_Click()
optTo.Enabled = True
optFrom.Enabled = True

If IsNewItemCreated Then Exit Sub

'If m_DBOperation = Update Then Exit Sub


If optFrom.Value Then
    optTo.Enabled = False
    RaiseEvent OptFromClicked
ElseIf optTo.Value Then
    optFrom.Enabled = False
    RaiseEvent OptToClicked
End If

End Sub


Private Sub optTo_Click()
'optTo.Enabled = True
'optFrom.Enabled = True

If IsNewItemCreated Then Exit Sub

'If m_DBOperation = Update Then Exit Sub


If optFrom.Value Then
    optTo.Enabled = False
    RaiseEvent OptFromClicked
ElseIf optTo.Value Then
    optFrom.Enabled = False
    RaiseEvent OptToClicked
End If
End Sub


Private Sub tabTransfer_Click()
Dim lpCount As Byte
For lpCount = 1 To fra.Count
    fra(lpCount - 1).Visible = False
Next

With tabTransfer
    fra(.SelectedItem.Index - 1).Visible = True
    fra(.SelectedItem.Index - 1).ZOrder 0
    
    If .SelectedItem.Index = 1 Then
        'InitGridForTab0
        'GridResize (1)
        grd(TRANSFER_SHOW).Visible = False
        grd(TRANSFER_TRANSACTION).Visible = True
        grd(TRANSFER_TRANSACTION).ZOrder 0
        
    ElseIf .SelectedItem.Index = 2 Then
        'txtFromDate.Text = FinFromDate
        'txtToDate.Text = DayBeginDate
        'InitGridForTab1
        'GridResize (2)
        grd(TRANSFER_TRANSACTION).Visible = False
        grd(TRANSFER_SHOW).Visible = True
        grd(TRANSFER_SHOW).ZOrder 0
        cmdTab1Show.Default = True
        
        ActivateTextBox txtFromDate
    End If
    
End With



End Sub




Private Sub tabTransfer_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyF3 Then frmDateChange.Show 'vbModal


End Sub


Private Sub tabTransfer_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim CtrlDown
'Dim IndxNo As Integer
'
'CtrlDown = (Shift And vbCtrlMask) > 0
'
'If Not CtrlDown Then Exit Sub
'
'If KeyCode = vbKeyTab Then
'    IndxNo = Me.tabTransfer.SelectedItem.Index
'    IndxNo = IndxNo + IIf(Shift = 2, 1, -1)
'    If IndxNo < 1 Then IndxNo = tabTransfer.Tabs.Count
'    If IndxNo > tabTransfer.Tabs.Count Then IndxNo = 1
'    tabTransfer.Tabs(IndxNo).Selected = True
'End If
End Sub


Private Sub txtFromDate_GotFocus()
ActivateTextBox txtFromDate
End Sub


Private Sub txtQuantity_GotFocus()
ActivateTextBox txtQuantity

End Sub


Private Sub txtToDate_GotFocus()
ActivateTextBox txtToDate
End Sub


Private Sub txtTransferDate_GotFocus()
ActivateDateTextBox txtTransferDate
End Sub


