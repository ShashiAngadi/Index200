VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrdFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Public grd As MSHFlexGrid
Public fGrd As MSFlexGrid


Public Property Get ObjectName()
    ObjectName = "GridFunctionClass"
End Property


Public Sub fGridRows(HeadArray() As String, Optional LExcel As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer
Dim Items As Integer
With fGrd
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Row = .FixedRows
    .Visible = True
    
    RowNum = .Row
    .Col = 0
    For Items = LBound(HeadArray) To UBound(HeadArray)
        fMoreRows (2)
        fMoreRows (4) ' double check
        .Row = RowNum: .Text = HeadArray(Items, 0): .CellFontBold = True: .CellAlignment = 0
        .RowHeight(RowNum) = Val(HeadArray(Items, 1))
        If LExcel Then
            'With xlWorkSheet
             '   .Cells(RowNum + 1, ColNum + 1) = HeadArray(RowNum)
              '  .Cells(RowNum + 1, ColNum + 1).Font.Bold = True
            'End With
        End If
        RowNum = RowNum + 1
    Next Items
    
End With
          
            


End Sub



Public Sub fGridCols(HeadArray() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer
Dim Cols As Integer
Dim Items As Integer

With fGrd
    RowNum = .Row
    .Col = 0
                
    ' put the main header column wise
    
    For Items = LBound(HeadArray) To UBound(HeadArray)
        .Col = Items: .Text = HeadArray(Items): .CellFontBold = True: .CellAlignment = 4
        '.ColWidth(Items) = Val(HeadArray(Items, 1))
        If LExcel Then
            'With xlWorkSheet
             '   .Cells(RowNum + 1, ColNum + 1) = HeadArray(ColNum)
            '    .Cells(RowNum + 1, ColNum + 1).Font.Bold = True
            'End With
        End If
    Next
    
    ' if lslno is true
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        For ColNum = LBound(HeadArray) To UBound(HeadArray)
            .Col = ColNum: .Text = ColNum + 1: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                'With xlWorkSheet
                 '   .Cells(RowNum + 1, ColNum + 1) = ColNum + 1
                  '  .Cells(RowNum + 1, ColNum + 1).Font.Bold = True
                'End With
            End If
        Next ColNum
    End If
    
End With
          

End Sub

Public Sub fMoreRows(RowNum As Integer)

With fGrd
    If .Rows <= .Row + RowNum Then .Rows = .Rows + RowNum
End With

End Sub


Private Sub Class_Terminate()

On Error Resume Next
Set fGrd = Nothing


End Sub


