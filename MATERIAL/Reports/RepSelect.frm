VERSION 5.00
Begin VB.Form frmRepSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Selector"
   ClientHeight    =   5595
   ClientLeft      =   1965
   ClientTop       =   2055
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5940
   Begin VB.TextBox txtFromDate 
      Height          =   345
      Left            =   1560
      TabIndex        =   13
      Text            =   "01/04/2002"
      Top             =   4200
      Width           =   1245
   End
   Begin VB.TextBox txtToDate 
      Height          =   345
      Left            =   4530
      TabIndex        =   15
      Text            =   "31/03/2002"
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Frame fraSelectReport 
      Caption         =   "Select Report"
      Height          =   2385
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      Begin VB.OptionButton optReports 
         Caption         =   "Product Sales"
         Height          =   300
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   1950
         Width           =   2235
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Stock As on"
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Product Purchase"
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Product Sales"
         Height          =   300
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   1470
         Width           =   2415
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Sales Invoices"
         Height          =   300
         Index           =   3
         Left            =   3030
         TabIndex        =   4
         Top             =   480
         Width           =   2355
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Purchase Invoices"
         Height          =   300
         Index           =   4
         Left            =   3030
         TabIndex        =   5
         Top             =   960
         Width           =   2355
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Stock Summary"
         Height          =   300
         Index           =   5
         Left            =   3030
         TabIndex        =   18
         Top             =   1470
         Width           =   2355
      End
   End
   Begin VB.ComboBox cmbGodown 
      Height          =   315
      ItemData        =   "RepSelect.frx":0000
      Left            =   2400
      List            =   "RepSelect.frx":0002
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   4620
      TabIndex        =   17
      Top             =   4890
      Width           =   1200
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   450
      Left            =   3240
      TabIndex        =   16
      Top             =   4890
      Width           =   1200
   End
   Begin VB.ComboBox cmbGroup 
      Height          =   315
      ItemData        =   "RepSelect.frx":0004
      Left            =   2400
      List            =   "RepSelect.frx":0006
      TabIndex        =   9
      Top             =   3060
      Width           =   3375
   End
   Begin VB.ComboBox cmbProductName 
      Height          =   315
      ItemData        =   "RepSelect.frx":0008
      Left            =   2400
      List            =   "RepSelect.frx":000A
      TabIndex        =   11
      Top             =   3510
      Width           =   3375
   End
   Begin VB.Line Line2 
      X1              =   90
      X2              =   5880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5760
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblDate1 
      Caption         =   "From Date"
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   4200
      Width           =   1485
   End
   Begin VB.Label lblDate2 
      Caption         =   "To Date"
      Height          =   315
      Left            =   3150
      TabIndex        =   14
      Top             =   4230
      Width           =   1305
   End
   Begin VB.Label lblGodown 
      Caption         =   "Select Branch"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   2550
      Width           =   2145
   End
   Begin VB.Label lblProductName 
      Caption         =   "Select Product Name"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   3525
      Width           =   2145
   End
   Begin VB.Label lblGroup 
      Caption         =   "Select Group"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   3090
      Width           =   2145
   End
End
Attribute VB_Name = "frmRepSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event WindowClosed()

Private m_ReportClass As Object
Private m_TheReportType As Wis_ReportType

Const STOCK_AS_ON = 0
Const STOCK_PURCHASE = 1
Const STOCK_SALES = 2
Const SALES_INVOICES = 3
Const PURCHASE_INVOICES = 4
Const STOCK_SUMMARY = 5
Const SOOT_REPORT = 6
'Private m_GodownID As Byte 'local copy
Public Sub LoadBranches()
Dim rstBranches As ADODB.Recordset
Dim GodownID As ADODB.Field
Dim GodownName As ADODB.Field

gDbTrans.SqlStmt = " SELECT GodownID,GodownName FROM GodownDet " & _
                   " ORDER BY GodownID "

        
Call gDbTrans.Fetch(rstBranches, adOpenStatic)

Set GodownID = rstBranches.Fields("GodownID")
Set GodownName = rstBranches.Fields("GodownName")

cmbGodown.Clear

cmbGodown.AddItem " "
cmbGodown.ItemData(cmbGodown.newIndex) = 0
   
   
Do While Not rstBranches.EOF
   cmbGodown.AddItem GodownName.Value
   cmbGodown.ItemData(cmbGodown.newIndex) = GodownID.Value
   
   'Move to next record
   rstBranches.MoveNext
Loop

If cmbGodown.ListCount = 1 Then
    cmbGodown.ListIndex = 0
    cmbGodown.Locked = True
    cmbGodown.Enabled = False
ElseIf cmbGodown.ListCount > 1 Then
    cmbGodown.ListIndex = 1
End If

End Sub


Private Function SalesValidated() As Boolean
SalesValidated = False
On Error Resume Next

Screen.MousePointer = vbDefault
If cmbGodown.ListIndex = -1 Then
    MsgBox "Please Select Branch", vbInformation, wis_MESSAGE_TITLE
    cmbGodown.SetFocus
    Exit Function
End If
If cmbGodown.ItemData(cmbGodown.ListIndex) = 0 Then
    MsgBox "Please Select Branch", vbInformation, wis_MESSAGE_TITLE
    cmbGodown.SetFocus
    Exit Function
End If

SalesValidated = True
Screen.MousePointer = vbHourglass
End Function

Private Function PurchaseValidated() As Boolean
PurchaseValidated = False

On Error Resume Next
Screen.MousePointer = vbDefault

PurchaseValidated = True
Screen.MousePointer = vbHourglass
End Function

'set the Kannada option here.
Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

'set the Kannada for all controls
lblGroup.Caption = GetResourceString(157, 27)
lblProductName.Caption = GetResourceString(158, 35, 27)
'for Frame caption
fraSelectReport.Caption = GetResourceString(283) & GetResourceString(92)
'for Options
optReports(0).Caption = GetResourceString(175)
optReports(1).Caption = GetResourceString(158, 176)
optReports(2).Caption = GetResourceString(158, 180)
optReports(3).Caption = GetResourceString(176, 417)
optReports(4).Caption = GetResourceString(180, 417)
optReports(5).Caption = GetResourceString(175, 295)
optReports(6).Caption = "soot" & " " & GetResourceString(283)

cmdView.Caption = GetResourceString(13)
cmdCancel.Caption = GetResourceString(2)
End Sub


Private Sub LoadGroupsToCombo()
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rst As ADODB.Recordset

'Get the data from the database
gDbTrans.SqlStmt = " SELECT GroupID,GroupName " & _
                   " FROM ProductGroup " & _
                   " ORDER BY GroupName "
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Clear the combo
cmbGroup.Clear

'Asign Blank Data for the ListiNdex 0
cmbGroup.AddItem ""
cmbGroup.ItemData(cmbGroup.newIndex) = 0

'Load the Groups to the combo box
Do
    If rst.EOF Then Exit Sub
        
    cmbGroup.AddItem FormatField(rst.Fields("GroupName"))
    cmbGroup.ItemData(cmbGroup.newIndex) = FormatField(rst.Fields("GroupID"))
    
    'move the recordset to next position
    rst.MoveNext
    
Loop

Exit Sub

ErrLine:
    If Err Then
        MsgBox "LoadGroupsToCombo: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Sub
    End If
End Sub

Private Sub LoadCompaniesToCombo(ByVal CompanyType As wis_CompanyType)
   
    MsgBox " NO IMplementetion LoadCompaniesToCombo: "
    
End Sub

Private Sub LoadProductsToCombo()
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rst As ADODB.Recordset
Dim lngGroupID As Long

'Validate the inputs
If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbGroup.ListIndex = 0 Then Exit Sub

lngGroupID = cmbGroup.ItemData(cmbGroup.ListIndex)


'Get the data from the database
gDbTrans.SqlStmt = " SELECT ProductID,ProductName " & _
                   " FROM Products " & _
                   " WHERE GroupID = " & lngGroupID & _
                   " ORDER BY ProductName "
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Clear the combo
cmbProductName.Clear

'Asign Blank Data for the ListiNdex 0
cmbProductName.AddItem ""
cmbProductName.ItemData(cmbProductName.newIndex) = 0

'Load the Groups to the combo box
Do
    If rst.EOF Then Exit Sub
        
    cmbProductName.AddItem FormatField(rst.Fields("ProductName"))
    cmbProductName.ItemData(cmbProductName.newIndex) = FormatField(rst.Fields("ProductID"))
    
    'move the recordset to next position
    rst.MoveNext
    
Loop


Exit Sub

ErrLine:
    If Err Then
        MsgBox "LoadProductsToCombo: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Sub
    End If
End Sub
Private Sub LoadALLProductsToCombo()
'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rst As ADODB.Recordset

'Get the data from the database
gDbTrans.SqlStmt = " SELECT ProductID,ProductName " & _
                   " FROM Products " & _
                   " ORDER BY ProductName "
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

'Clear the combo
cmbProductName.Clear

'Asign Blank Data for the ListiNdex 0
cmbProductName.AddItem ""
cmbProductName.ItemData(cmbProductName.newIndex) = 0

'Load the Groups to the combo box
Do
    If rst.EOF Then Exit Sub
        
    cmbProductName.AddItem FormatField(rst.Fields("ProductName"))
    cmbProductName.ItemData(cmbProductName.newIndex) = FormatField(rst.Fields("ProductID"))
    
    'move the recordset to next position
    rst.MoveNext
    
Loop


Exit Sub

ErrLine:
    If Err Then
        MsgBox "LoadProductsToCombo: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        Exit Sub
    End If
End Sub

Private Function Validated() As Boolean
Dim StartDateUS As String
Dim EndDateUS As String

If Not TextBoxDateValidate(txtFromDate, "/", True, True) Then Exit Function
If Not TextBoxDateValidate(txtToDate, "/", True, True) Then Exit Function

StartDateUS = GetSysFormatDate(txtFromDate.Text)
EndDateUS = GetSysFormatDate(txtToDate.Text)

If DateDiff("d", CDate(StartDateUS), CDate(EndDateUS)) < 0 Then
    MsgBox "Start date should be earlier than the end date ", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

Validated = True

End Function

Private Sub cmbGodown_Click()
Dim GodownID As Byte

If cmbGodown.ListIndex = -1 Then Exit Sub
GodownID = cmbGodown.ItemData(cmbGodown.ListIndex)

If GodownID > 1 Then
    optReports(0).Enabled = True
    optReports(2).Enabled = True
    optReports(1).Enabled = False
    optReports(3).Enabled = True 'Sales Invoice
    optReports(4).Enabled = False
     optReports(5).Enabled = False
    optReports(0).Value = True
Else
    optReports(0).Enabled = True
    optReports(1).Enabled = True
    optReports(2).Enabled = True
    optReports(3).Enabled = True
    optReports(4).Enabled = True
    optReports(5).Enabled = True
End If

End Sub

Private Sub cmbGroup_Click()

'Load the products
cmbProductName.Clear

If cmbGroup.ListIndex = -1 Then Exit Sub
If cmbGroup.ListIndex = 0 Then Exit Sub

Call LoadProductsToCombo

End Sub

Private Sub cmdCancel_Click()

Unload Me
RaiseEvent WindowClosed
End Sub

Private Sub cmdView_Click()

'Declare the variables
Dim ItemIndex As Integer
Dim ReportSuccess As Boolean
Dim headID As Long
Dim GodownID As Byte

If Not Validated Then Exit Sub

For ItemIndex = optReports.LBound To optReports.UBound
    If optReports(ItemIndex).Value Then Exit For
Next ItemIndex

Screen.MousePointer = vbHourglass

m_ReportClass.FromRpDate = txtFromDate.Text
m_ReportClass.ToRpDate = txtToDate.Text

Select Case ItemIndex

    Case 0 'Stock Report
        If m_ReportClass.ShowStockReports Then ReportSuccess = True
        
    Case 1 'Purchase Report
        If m_ReportClass.ShowStockPurchaseReports Then ReportSuccess = True
        
    Case 2 'Sales Report
        If m_ReportClass.ShowStockSalesReports Then ReportSuccess = True
    
    Case 3 'Purchase Invoices
        If Not PurchaseValidated Then Exit Sub
        'HeadID = cmbManufacturer.ItemData(cmbManufacturer.ListIndex)
        If m_ReportClass.ShowPurchaseInvoices() Then ReportSuccess = True
    
    Case 4 'Sales Invoices
        If Not SalesValidated Then Exit Sub
        GodownID = cmbGodown.ItemData(cmbGodown.ListIndex)
        If m_ReportClass.ShowSalesInvoices(GodownID) Then ReportSuccess = True
    
    Case 5
        Dim RepClass As New clsStockRep
        RepClass.ShowStockSummaryReports
        If m_ReportClass Is Nothing Then Set m_ReportClass = New clsStockRep
        If m_ReportClass.ShowStockSummaryReports Then ReportSuccess = True

    Case 6
        'If m_ReportClass.ShowSootReports Then ReportSuccess = True
        m_ReportClass.SootReport = True
        If m_ReportClass.ShowStockSalesReports Then ReportSuccess = True
        m_ReportClass.SootReport = False
        
End Select

'Set ReportClass = Nothing
If Not ReportSuccess Then
    Screen.MousePointer = vbDefault
    MsgBox "No Records", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If

Set m_ReportClass = Nothing
Call optReports_Click(ItemIndex)


Screen.MousePointer = vbDefault

End Sub

Private Function GetReportTypeToLoad(Optional headID As Long, Optional GroupID As Integer, Optional ProductID As Long, Optional GodownID As Byte) As Wis_ReportType
'Declare the variables
Dim ItemIndex As Integer

GetReportTypeToLoad = 0

With frmRepSelect
    For ItemIndex = .optReports.LBound To .optReports.UBound
        If .optReports(ItemIndex).Value Then Exit For
    Next ItemIndex
    
    If GodownID = 0 Then
        Select Case .optReports(ItemIndex).Index
            Case 0
                If GroupID = 0 Then
                    GetReportTypeToLoad = StockIncludingBranches
                    Exit Function
                ElseIf GroupID > 0 Then
                    If ProductID = 0 Then
                        GetReportTypeToLoad = StockOfGroupIncBranches
                        Exit Function
                    Else
                        'GetReportTypeToLoad = StockOfGroupAndProduct
                        'Exit Function
                    End If
                End If
            Case 2
                If GroupID = 0 Then
                    GetReportTypeToLoad = SalesIncludingBranches
                    Exit Function
                ElseIf GroupID > 0 Then
                    If ProductID = 0 Then
                        'GetReportTypeToLoad = SalesFromGroup
                        'Exit Function
                    Else
                        'GetReportTypeToLoad = SalesFromGroupAndProduct
                        'Exit Function
                    End If
                End If
        End Select
    End If

    ''''
    If headID = 0 Then
        Select Case .optReports(ItemIndex).Index
            Case 0
                If GroupID = 0 Then
                    GetReportTypeToLoad = StockAsOn
                ElseIf GroupID > 0 Then
                    If ProductID = 0 Then
                        GetReportTypeToLoad = StockOfGroup
                    Else
                        GetReportTypeToLoad = StockOfGroupAndProduct
                    End If
                End If
            Case 1
                If GroupID = 0 Then
                    GetReportTypeToLoad = PurchaseOfBranches
                ElseIf GroupID > 0 Then
                    If ProductID = 0 Then
                        GetReportTypeToLoad = PurchaseOfGroup
                    Else
                        GetReportTypeToLoad = PurchaseOfGroupAndProduct
                    End If
                End If
            Case 2
                If GroupID = 0 Then
                    GetReportTypeToLoad = SalesOfBranches
                ElseIf GroupID > 0 Then
                    If ProductID = 0 Then
                        GetReportTypeToLoad = SalesOfGroup
                    Else
                        GetReportTypeToLoad = SalesOfGroupAndProduct
                    End If
                End If
            Case 3
        End Select
    Else
        Select Case .optReports(ItemIndex).Index
            Case 0
            Case 1
            Case 2
            Case 3
        End Select
    End If
End With

End Function
Private Sub Form_Load()
'Center the form
CenterMe Me

'set icon for the form caption
'Me.Icon = LoadResPicture(147, vbResIcon)
  
Call SetKannadaCaption

txtFromDate.Text = FinIndianFromDate
txtToDate.Text = gStrDate 'FinIndianEndDate
txtFromDate.Locked = False

If optReports(STOCK_AS_ON).Value Then txtFromDate.Locked = True
'Load baranches
LoadBranches

'Load the companies to ComboBox
Debug.Print "MFR"

'LoadGroups to Combobox
LoadGroupsToCombo

'Show the Blank data in the Combo Box

Call optReports_Click(STOCK_AS_ON)

'If cmbManufacturer.ListCount > 0 Then cmbManufacturer.ListIndex = 0
If cmbGroup.ListCount > 0 Then cmbGroup.ListIndex = 0


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmRepSelect = Nothing
End Sub

Private Sub optReports_Click(Index As Integer)
Dim lpIndex As Integer

cmbGodown.Enabled = True
cmbGroup.Enabled = True
cmbProductName.Enabled = True

txtFromDate.Text = FinIndianFromDate
If txtToDate.Text = "" Then txtToDate.Text = gStrDate 'FinIndianEndDate
txtFromDate.Locked = False

Select Case Index
    Case STOCK_AS_ON, SALES_INVOICES, PURCHASE_INVOICES, STOCK_SUMMARY
        Set m_ReportClass = New clsStockRep
    Case STOCK_PURCHASE
        Set m_ReportClass = New clsPurchaseRep
    Case STOCK_SALES, SOOT_REPORT
        Set m_ReportClass = New clsSalesRep
End Select

Call m_ReportClass.SetStockRepSelectForm(frmRepSelect)

If Index = STOCK_AS_ON Then txtFromDate.Locked = True

If cmbGodown.ListIndex = -1 Then Exit Sub

If Index = STOCK_PURCHASE Then
    For lpIndex = 0 To cmbGodown.ListCount - 1
        If cmbGodown.ItemData(lpIndex) = 1 Then
            cmbGodown.ListIndex = lpIndex
            cmbGodown.Enabled = False
            Exit For
        End If
    Next lpIndex
End If

If Index = SALES_INVOICES Then
    'LoadCompaniesToCombo (Enum_Customers)
    cmbGroup.Enabled = False
    cmbProductName.Enabled = False
Else
    'LoadCompaniesToCombo (Enum_Stockist)
End If

End Sub

Private Sub txtFromDate_GotFocus()
    ActivateTextBox txtFromDate
End Sub


Private Sub txtToDate_GotFocus()
ActivateTextBox txtToDate
End Sub


