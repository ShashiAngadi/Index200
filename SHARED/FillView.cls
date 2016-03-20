VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFillView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ObjectName()
    ObjectName = "FillViewClass"
End Property


' Fills the listview control with the record set data.
Public Function FillView(view As ListView, rs As ADODB.Recordset, KeyField As String, Optional AutoWidth As Boolean) As Boolean

Dim strKey As String
Dim I As Integer
Dim itmX As ListItem
Dim StartPos As Integer
Dim KeyCount As Long
Dim SlNo As Long

Const FIELD_MARGIN = 1.5

' Check if there are any records in the recordset.
If Not rs.EOF Then rs.MoveLast
If Not rs.BOF Then rs.MoveFirst

If rs.EOF And rs.BOF Then
    FillView = True
    Exit Function
End If

With view
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    Dim X As Integer
    X = 4
    If rs.Fields.Count <= X Then X = rs.Fields.Count - 1
    
    For I = 0 To X  'display only selected fields instead of all the fields
        If rs.Fields(I).Name = KeyField Then I = I + 1
        
        .ColumnHeaders.Add , rs.Fields(I).Name, rs.Fields(I).Name, _
                     view.Parent.TextWidth(rs.Fields(I).Name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                        rs.Fields(I).Type = adTinyInt Or _
                        rs.Fields(I).Type = adSmallInt Or _
                        rs.Fields(I).Type = adBigInt Or _
                        rs.Fields(I).Type = adSingle Or _
                        rs.Fields(I).Type = adDouble Or _
                        rs.Fields(I).Type = adCurrency Then
                        
               .ColumnHeaders(I).Alignment = lvwColumnRight
            End If
        End If
        ' If the autowidth property is set,
        ' check if the width of the column is to be adjusted.
    Next

    ' Begin a loop for processing rows.
    KeyCount = 0
    Do While Not rs.EOF
      KeyCount = KeyCount + 1
      DoEvents
      ' Add the details.
      strKey = "A" & FormatField(rs.Fields(KeyField))
      If rs.Fields(0).Name = KeyField Then
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(1)))
        StartPos = 2
      Else
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(0)))
        StartPos = 1
      End If
      ' If the 'Autowidth' property is enabled,
      ' then check if the width needs to be expanded.
      If AutoWidth Then
         If .ColumnHeaders(1).Width \ FIELD_MARGIN < _
                     .Parent.TextWidth(FormatField(rs.Fields(0))) Then
             .ColumnHeaders(1).Width = _
                 .Parent.TextWidth(FormatField(rs.Fields(0))) * FIELD_MARGIN
         End If
      End If
      ' Add sub-items.
      ' For I = 1 To rs.fields.Count - 1
        X = 4
      If rs.Fields.Count <= 4 Then X = rs.Fields.Count - 1

       For I = StartPos To X 'display only necessary fields to user
            If rs.Fields(I).Name = KeyField Then I = I + 1
            
            itmX.SubItems(I - 1) = FormatField(rs.Fields(I))
            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then
                If .ColumnHeaders(I).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) Then
                    .ColumnHeaders(I).Width = _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) * FIELD_MARGIN
                End If
            End If
        Next
        rs.MoveNext
    Loop
End With

FillView = True
view.Visible = True
view.view = lvwReport

End Function
' Fills the listview control with the record set data.
Public Function FillViewForPurchase(view As ListView, rs As ADODB.Recordset, KeyField As String, Optional AutoWidth As Boolean) As Boolean

Dim strKey As String
Dim I As Integer
Dim itmX As ListItem
Dim StartPos As Integer
Dim KeyCount As Long
Dim SlNo As Long
Dim X As Integer
Const FIELD_MARGIN = 1.5

' Check if there are any records in the recordset.
If Not rs.EOF Then rs.MoveLast
If Not rs.BOF Then rs.MoveFirst

If rs.EOF And rs.BOF Then
    FillViewForPurchase = True
    Exit Function
End If

With view
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    X = rs.Fields.Count - 1
    
    For I = 0 To X
        If rs.Fields(I).Name = KeyField Then I = I + 1
        
        .ColumnHeaders.Add , rs.Fields(I).Name, rs.Fields(I).Name, _
                     view.Parent.TextWidth(rs.Fields(I).Name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                        rs.Fields(I).Type = adTinyInt Or _
                        rs.Fields(I).Type = adSmallInt Or _
                        rs.Fields(I).Type = adBigInt Or _
                        rs.Fields(I).Type = adSingle Or _
                        rs.Fields(I).Type = adDouble Or _
                        rs.Fields(I).Type = adCurrency Then
                        
               .ColumnHeaders(I).Alignment = lvwColumnRight
            End If
        End If
        ' If the autowidth property is set,
        ' check if the width of the column is to be adjusted.
    Next

    ' Begin a loop for processing rows.
    KeyCount = 0
    Do While Not rs.EOF
      KeyCount = KeyCount + 1
      DoEvents
      ' Add the details.
      strKey = "KEY" & FormatField(rs.Fields(KeyField))
      If rs.Fields(0).Name = KeyField Then
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(1)))
        StartPos = 2
      Else
        Set itmX = .ListItems.Add(, strKey, FormatField(rs.Fields(0)))
        StartPos = 1
      End If
      ' If the 'Autowidth' property is enabled,
      ' then check if the width needs to be expanded.
      If AutoWidth Then
         If .ColumnHeaders(1).Width \ FIELD_MARGIN < _
                     .Parent.TextWidth(FormatField(rs.Fields(0))) Then
             .ColumnHeaders(1).Width = _
                 .Parent.TextWidth(FormatField(rs.Fields(0))) * FIELD_MARGIN
         End If
      End If
      ' Add sub-items.
      ' For I = 1 To rs.fields.Count - 1
       For I = StartPos To X
            If rs.Fields(I).Name = KeyField Then I = I + 1
            
            itmX.SubItems(I - 1) = FormatField(rs.Fields(I))
            ' If the 'Autowidth' property is enabled,
            ' then check if the width needs to be expanded.
            If AutoWidth Then
                If .ColumnHeaders(I).Width \ FIELD_MARGIN < _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) Then
                    .ColumnHeaders(I).Width = _
                        .Parent.TextWidth(FormatField(rs.Fields(I))) * FIELD_MARGIN
                End If
            End If
        Next
        rs.MoveNext
    Loop
End With

FillViewForPurchase = True
view.Visible = True
view.view = lvwReport

End Function

' Fills the listview control with the record set data.
Public Function FillViewWithSlno(view As ListView, rs As ADODB.Recordset, KeyField As String, Optional AutoWidth As Boolean) As Boolean

Dim strKey As String
Dim I As Integer
Dim itmX As ListItem
Dim StartPos As Integer
Dim KeyCount As Long
Dim SlNo As Long
Dim MaxWidth As Long

Const FIELD_MARGIN = 1.5
Const FILED1_WIDTH = 100


' Check if there are any records in the recordset.
If Not rs.EOF Then rs.MoveLast
If Not rs.BOF Then rs.MoveFirst

If rs.EOF And rs.BOF Then
    FillViewWithSlno = True
    Exit Function
End If

With view
    ' Hide the view control before processing.
    .Visible = False
    .ListItems.Clear
    .ColumnHeaders.Clear

    ' Add column headers.
    Dim X As Integer
    X = 4
    
    ' This is for the SL No
    .ColumnHeaders.Add , "slno", "Sl No"
    If rs.Fields.Count <= X Then X = rs.Fields.Count - 1
    
    For I = 0 To X  'display only selected fields instead of all the fields
        
        If rs.Fields(I).Name = KeyField Then I = I + 1
        
        .ColumnHeaders.Add , rs.Fields(I).Name, rs.Fields(I).Name, _
                     view.Parent.TextWidth(rs.Fields(I).Name) * FIELD_MARGIN
        ' Set the alignment characterstic for the column.
        If I > 0 Then
            If rs.Fields(I).Type = adNumeric Or _
                        rs.Fields(I).Type = adTinyInt Or _
                        rs.Fields(I).Type = adSmallInt Or _
                        rs.Fields(I).Type = adBigInt Or _
                        rs.Fields(I).Type = adSingle Or _
                        rs.Fields(I).Type = adDouble Or _
                        rs.Fields(I).Type = adCurrency Then
                        
               .ColumnHeaders(I).Alignment = lvwColumnRight
            End If
        End If
        ' If the autowidth property is set,
        ' check if the width of the column is to be adjusted.
    Next

    ' Begin a loop for processing rows.
    KeyCount = 0
    SlNo = 1
    Do While Not rs.EOF
      KeyCount = KeyCount + 1
      DoEvents
      ' Add the details.
      strKey = "A" & FormatField(rs.Fields(KeyField))
      
      Set itmX = .ListItems.Add(, strKey, SlNo)
      
      ' Add sub-items.
      ' For I = 1 To rs.fields.Count - 1
      
      X = 4
      If rs.Fields.Count <= 4 Then X = rs.Fields.Count - 1
        
        StartPos = 1
        For I = StartPos To X 'display only necessary fields to user
            If rs.Fields(I).Name = KeyField Then I = I + 1
            
            itmX.SubItems(I) = FormatField(rs.Fields(I))
            
            If MaxWidth < .Parent.TextWidth(Trim$(rs.Fields(1))) * FIELD_MARGIN Then _
            MaxWidth = .Parent.TextWidth(Trim$(rs.Fields(1))) * FIELD_MARGIN
            
           .ColumnHeaders(2).Width = MaxWidth
            
        Next
        rs.MoveNext
        SlNo = SlNo + 1
    Loop
    
    .ColumnHeaders(1).Width = FIELD_MARGIN
    
End With

FillViewWithSlno = True
view.Visible = True
view.view = lvwReport

End Function


