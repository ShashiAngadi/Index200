VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Define member variables of this print class...
'

' Constants...
Private Const TITLE_TOP = 50
Private Const FIELD_MARGIN = 200
Private Const PAGE_MARGIN = 500
Private Const RECT_MARGIN = 15
Private Const WIS_FOOTER = "Waves Information Systems"

' Title Object.
Private m_Title As clsField

Private my_printer As Object                         ' Can be a printer or a picturebox object.
Private m_DataSource As MSFlexGrid           ' Grid to read rows from.

Public PageNums As Boolean                          ' Prints the pagenos, if set to true
Public ReportDate As Boolean                        ' Prints the date on report, if true.
Public HeaderRectangle As Boolean               ' Prints a rectangle around the header, if true
Private print_EOF As Boolean                        ' Flag to identify the end of rows.

Private m_FontSize As Single                        ' To Set The Fotn Size And Font NAme for Printing
Private m_FontName As String


Private WithEvents PrintDlg As frmPrint     ' Preview display object.
Attribute PrintDlg.VB_VarHelpID = -1

Private title_printed As Boolean                     ' Suppresses the title for subsequent pages.
Private rows_per_page As Integer
Private rows_in_lastpage As Integer
Private num_pages As Integer
Private page_start_row As Long
Public PageNumber As Integer
Private pause_between_pages As Boolean      ' Applies while printing to printer.
Private LastPage_Reference As Integer      ' Applies While Moving to next & Previrues pages
Private m_heading_top As Single
Public CompanyName As String             ' This name will appear in the title of reports...
Private curPage As Integer
Private saved_row As Long
Private saved_page_start_row As Long
Private TypeOfReport As Boolean
Private Function AlignData(strSource As String, wid As Single, Optional dataFormat As Integer) As String

If dataFormat = adCurrency Then
    strSource = FormatCurrency(CCur(strSource))
End If

If my_printer.TextWidth(strSource) < wid Then
    While my_printer.TextWidth(strSource) < wid
        strSource = " " & strSource
    Wend
End If
AlignData = strSource
End Function
Private Sub PrintFooter()
#If DEFERRED Then

Dim I As Integer
Dim curRow As Single
Dim curX As Single
Dim strData As String

With my_printer
    ' Draw a line.
'''    Debug.Print "TODO: Check for margin overflow..."
    .CurrentY = my_printer.CurrentY + 200
    my_printer.Line (FIELD_MARGIN, .CurrentY)-(.ScaleWidth - FIELD_MARGIN, .CurrentY), , B

    curRow = my_printer.CurrentY ' Store the coordinates
    curX = FIELD_MARGIN

    ' Loop in the fields collection and print the sum for fields specified.
'''    For i = 1 To m_fields.Count

'''        If m_fields(i).SumField Then

            ' Set the font attributes.
'''            m_page_header.SetAttrib my_printer

            ' Set the co-ordinates for printing
            my_printer.CurrentX = curX
            my_printer.CurrentY = curRow

            ' Set the justification, if any.
'''            If m_fields(i).Align = "RIGHT" Then
'''                strData = AlignData(m_fields(i).SumVal, m_fields(i).Width, dbCurrency)
'''            Else
'''                strData = FormatCurrency(CCur(m_fields(i).SumVal))
'''            End If

            ' Print the data.
            my_printer.Print strData

'''        End If
        ' Update the print co-ordinates.
'''        curX = curX + m_fields(i).Width + FIELD_MARGIN
'''    Next
End With
#End If


End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                        Called by 'PrintReport'
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PrintHeading()
Static PageNumPosition As Single
Dim printX As Single, printY As Single
Dim I As Integer
Dim rectLeft As Single, rectTop As Single
Dim rectRight As Single, rectBottom As Single
Dim strData As String, pagestr As String
Dim saveRow As Long, saveCol As Long

On Error GoTo PrintHeading_Error

If TypeOf my_printer Is PictureBox Then
    PrintDlg.Init
End If

' Set the top position (Y-coordinate) for printing.)
printY = m_Title.Height

' If the 'pagenumbers' option specified, print the page number.
If Me.PageNums Then
    With my_printer
        .FontSize = 10
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
        pagestr = "Page No. : " & PageNumber
        .CurrentX = .ScaleWidth - .TextWidth(pagestr) - FIELD_MARGIN
        If PageNumPosition = 0 Then
            PageNumPosition = .CurrentY
        Else
            .CurrentY = PageNumPosition
        End If
        
    End With
    my_printer.Print pagestr

End If

printX = FIELD_MARGIN
'printY = my_printer.CurrentY + 200
printY = m_heading_top

' Print the header rectangle......
SetFont

If Me.HeaderRectangle Then
    rectLeft = printX - RECT_MARGIN
    rectTop = printY - RECT_MARGIN
    rectRight = rectLeft + (my_printer.ScaleWidth - 2 * FIELD_MARGIN) + RECT_MARGIN
    rectBottom = rectTop + my_printer.TextHeight("A") + RECT_MARGIN
    my_printer.Line (rectLeft, rectTop)-(rectRight, rectBottom), , B
End If

' Save the current grid row and move to top row.
saveRow = m_DataSource.Row
saveCol = m_DataSource.Col
m_DataSource.Row = 0
m_DataSource.Col = 0

' Loop through collection of fields...
For I = 0 To m_DataSource.Cols - 1

    m_DataSource.Col = I
    ' set the attributes of print object for printing.
    SetFont
    With my_printer
        .CurrentX = printX
        .CurrentY = printY
        If TypeOfReport Then
            If m_DataSource.Col = 3 Then
                printX = printX + m_DataSource.CellWidth + FIELD_MARGIN + 200
            ElseIf m_DataSource.Col = 0 Then
                printX = printX + m_DataSource.CellWidth + FIELD_MARGIN + 50
            Else
                printX = printX + m_DataSource.CellWidth + FIELD_MARGIN - 300
            End If
        Else
            printX = printX + m_DataSource.CellWidth + FIELD_MARGIN
        End If
        
        ' Check if the printx exceeds the width of the printing region.
        If printX > my_printer.ScaleWidth - FIELD_MARGIN Then
            MsgBox "Some fields are outside print margins." & vbCrLf _
                    & "Realign the field widths in the grid.", vbInformation
        End If
    End With
    If m_DataSource.ColAlignment(I) = 1 Then
        strData = AlignData(m_DataSource.Text, m_DataSource.CellWidth)
    Else
        strData = m_DataSource.Text
    End If
'''    Debug.Print strData
    If m_DataSource.CellWidth > my_printer.TextWidth("I") Then
        my_printer.Print strData
    Else
        my_printer.Print ""
    End If
Next

' Set the y-coordinate for printing record details...
my_printer.CurrentY = my_printer.CurrentY + 200

PrintHeading_Error:
    If Err Then
        MsgBox "PrintHeading: " & vbCrLf & Err.Description, vbCritical
    End If
'Resume

' Restore the grid row.
m_DataSource.Row = saveRow
m_DataSource.Col = saveCol
DoEvents
End Sub
Public Function PrintReport(Optional BalSheet As Boolean) As Boolean
Dim saveRow As Long
Dim nRet As Integer

' Setup error handler...
On Error GoTo Err_Line

If BalSheet Then TypeOfReport = True

' If no recordset available, exit.
If m_DataSource Is Nothing Then
    MsgBox "No records!  Assign a recordset for printing.", vbOKOnly + vbExclamation
    GoTo Exit_Line
End If

' Hide the grid, if visible.
m_DataSource.Visible = False

' Initialize the print form, if type of print is "PREVIEW"
If TypeOf my_printer Is PictureBox Then PrintDlg.Init

' Set the initial row of the grid to 1,
' because we have to skip the 0th row.
If m_DataSource.Rows > 0 Then
    If PageNumber = 1 Then m_DataSource.Row = 1
Else
    MsgBox "Internal error: No data in grid!!!", vbCritical
    GoTo Exit_Line
End If

' Loop through available pages.
Do While Not print_EOF

    ' Store the current row, in case of printing the same page.
    saveRow = m_DataSource.Row
    
    ' Print a page
    If Not PrintPage Then Exit Do

    ' if print type is "PREVIEW", display the view window.
    If ReportDestination = "PREVIEW" Then
        If PrintDlg Is Nothing Then
            Set PrintDlg = New frmPrint
        End If
        PrintDlg.Show vbModal
        If PrintDlg.Status = wis_CANCEL Then GoTo Exit_Line
    Else        ' Printing to printer.
        If pause_between_pages Then
            nRet = MsgBox("Print next page ?", vbQuestion + vbYesNo)
            If nRet = vbNo Then Exit Do
        End If
    End If

    ' Increment the page counter
    PageNumber = PageNumber + 1

Loop

Exit_Line:
    Screen.MousePointer = vbDefault
    Exit Function
    
Err_Line:
    If Err.Number = 482 Then
        nRet = MsgBox("Printer Error: " & vbCrLf & "Check your printer settings.", vbAbortRetryIgnore)
        If nRet = vbRetry Then
            Resume
        ElseIf nRet = vbIgnore Then
            Resume Next
        End If
    
    ElseIf Err Then
        MsgBox "PrintReport: " & vbCrLf & Err.Description, vbCritical
    End If
    GoTo Exit_Line

End Function
' This sub routine will be called, while processing the
' print request to the printer.  Basically, it saves the curpagenumber,
' and the current grid row, to restore it back, later.
Private Sub SaveSettings()
    
    ' Save the current view page.
    curPage = PageNumber

    ' Set the title printed flag to false,
    ' to force the printing of title on the first page.
    title_printed = False

    ' Set the page_start_row
    saved_row = m_DataSource.Row
    m_DataSource.Row = page_start_row
    
    ' Save the page_start_row
    saved_page_start_row = page_start_row
    
    ' Restore the current printer object.
    Set my_printer = Printer

End Sub
' This sub routine will be called, while processing the
' print request to the printer.  Basically, it saves the curpagenumber,
' and the current grid row, to restore it back, later.
' NOTE: This routine should be followed by a call to
'             RestoreSettings, after print processing.
Private Sub RestoreSettings()
    
    ' Restore the current view page.
    PageNumber = curPage
    
    ' Restore the data source current row
     m_DataSource.Row = saved_row

    ' Reset the title_printed flag
    title_printed = True

  ' Restore the page_start_row
     page_start_row = saved_page_start_row

    ' Save the current printer
    Set my_printer = PrintDlg.picPrint
End Sub
Private Sub SetFont(Optional srcObj As clsField)
On Error GoTo Err_Line
Dim obj As Object

If Not srcObj Is Nothing Then
    Set obj = srcObj
Else
    Set obj = m_DataSource
End If

With my_printer
    .FontName = obj.CellFontName
    .FontSize = obj.CellFontSize
    .FontBold = obj.CellFontBold
    .FontItalic = obj.CellFontItalic
    .FontUnderline = obj.CellFontUnderline
    .FontStrikethru = obj.CellFontStrikeThrough

#If DEFERRED Then
    ' if the height of the font exceeds, the height
    ' of the title section, increase the section height.
'''    If obj.TextHeight("A") > Height Then
'''        Height = obj.TextHeight("A") * 2
'''    End If
#End If

End With

Exit_Line:
    Exit Sub

Err_Line:
    If Err Then
        MsgBox "SetFont: " & vbCrLf _
            & Err.Description, vbCritical
    End If
'Resume
    GoTo Exit_Line
End Sub
Private Sub PrintRectangle(fldObj As clsField, curX As Single, curY As Single)
Dim rectLeft As Single, rectTop As Single
Dim rectRight As Single, rectBottom As Single

With fldObj
    rectLeft = FIELD_MARGIN - .RectMargin
    rectTop = curY - .RectMargin
    rectRight = rectLeft + .RectWidth(my_printer) + .RectMargin
    rectBottom = rectTop + .RectHeight(my_printer) + .RectMargin
End With
my_printer.Line (rectLeft, rectTop)-(rectRight, rectBottom), , B

' Restore the curx and cury.
my_printer.CurrentX = curX
my_printer.CurrentY = curY

End Sub
Public Function PrintPage() As Boolean

' Setup error handler.
On Error GoTo Err_Line
Screen.MousePointer = vbHourglass

' Declare variables...
Dim I As Integer
Dim lcount As Long

' In case of "PREVIEW", clear the screen.
If TypeOf my_printer Is PictureBox Then my_printer.Cls

With my_printer

    ' Print the title.
    If Not title_printed Or PageNumber = 1 Then
        PrintTitle
        title_printed = True
    End If

    ' Print the heading.
    PrintHeading

    ' IF the row is zero, force it to 1,
    ' because we do not want to print the 0th row
    If m_DataSource.Row = 0 Then m_DataSource.Row = 1

    ' Starting row for this page.
    page_start_row = m_DataSource.Row + 1

    ' Begin a loop for printing the records...
    For lcount = m_DataSource.Row To m_DataSource.Rows - 1
    
        ' Print the record details.
        PrintRow

        ' Check if page end has reached.
        If .CurrentY >= .ScaleHeight - PAGE_MARGIN Then
            'the Current row Alreay printed Print from Next Line
            If m_DataSource.Row <= m_DataSource.Rows - 2 Then
                m_DataSource.Row = m_DataSource.Row + 1
            End If
            Exit For
        End If
        If m_DataSource.Row <= m_DataSource.Rows - 2 Then
            m_DataSource.Row = m_DataSource.Row + 1
        End If
    Next
    'm_DataSource.Rows = m_DataSource.Rows - 1
    On Error Resume Next
        'm_DataSource.Row = m_DataSource.Row + 1
    On Error GoTo Err_Line
    #If DEFERRED Then
    ' Print the footer.
        ' Check if any of the fields has a sum property set.
'''        For i = 1 To m_fields.Count
'''            If m_fields(i).SumField Then
                PrintFooter
'''                Exit For
'''            End If
'''        Next
    #End If

End With

' Update the rows
If rows_per_page = 0 Then rows_per_page = m_DataSource.Row - page_start_row

' Time to do some page computations...
If num_pages = 0 Then
    num_pages = m_DataSource.Rows / rows_per_page
    rows_in_lastpage = m_DataSource.Rows Mod rows_per_page
End If

' Update pagecount info...
If TypeOf my_printer Is PictureBox Then
    PrintDlg.txtPageCount.Text = PageNumber & "/" & num_pages
End If

'
' Check if the last row has been reached.
' IF so, set the eof property to TRUE.
If m_DataSource.Row >= m_DataSource.Rows + 1 Then print_EOF = True
If PageNumber = num_pages Then print_EOF = True
Screen.MousePointer = vbDefault
If TypeOf my_printer Is Printer Then my_printer.NewPage
PrintPage = True

Exit_Line:
    Screen.MousePointer = vbDefault
    m_DataSource.Visible = True
    Exit Function

Err_Line:
    If Err Then
        MsgBox "PreviewReport: " & Err.Description
    End If
'Resume
    GoTo Exit_Line

End Function
Private Sub PrintRow()
Dim I As Integer
Dim curRow As Single
Dim curX As Single
Dim strData As String
On Error GoTo printrow_error

' Save the current row.
curRow = my_printer.CurrentY
curX = FIELD_MARGIN

With m_DataSource
    ' Loop through the collection of fields...
    For I = 0 To .Cols - 1

        .Col = I        ' Set the current cell.
        SetFont     ' Set the font for this field.

        ' If the width of the data to be printed is greater than the width of the field, truncate.
        If my_printer.TextWidth(.Text) > .CellWidth Then    'And .CellAlignment = flexAlignLeftCenter Then
            strData = TruncateData(.Text, .CellWidth)
        ElseIf .ColAlignment(I) = 1 Then
            strData = AlignData(.Text, .CellWidth)
        Else
            strData = .Text
        End If
        my_printer.CurrentY = curRow
        my_printer.CurrentX = curX - 10
        If m_DataSource.CellWidth > my_printer.TextWidth("I") Then
            'Debug.Print my_printer.CurrentY
            my_printer.Print strData
        Else
            my_printer.Print ""
        End If
next_field:
        If TypeOfReport Then
            If .Col = 3 Then
                curX = curX + (.CellWidth + 50) + FIELD_MARGIN
            Else
                curX = curX + (.CellWidth - 200) + FIELD_MARGIN
            End If
        Else
                curX = curX + .CellWidth + FIELD_MARGIN
        End If
    
    #If DEFERRED Then
        ' If this field is selected as a sum-field, then
        ' accumulate the total for this field.
'''        If m_fields(i).SumField And m_DataSource(i - 1).Type = dbNumeric Or _
'''                    m_DataSource(i - 1).Type = dbInteger Or m_DataSource(i - 1).Type = dbLong Or _
'''                    m_DataSource(i - 1).Type = dbCurrency Or m_DataSource(i - 1).Type = dbSingle Then
'''            m_fields(i).SumVal = m_fields(i).SumVal + m_DataSource(i - 1)
'''        End If
    #End If

    Next
End With
Exit Sub

printrow_error:
    If Err Then MsgBox "Printrow: " & vbCrLf & Err.Description, vbCritical '
'Resume
End Sub
Private Sub PrintTitle()

If m_Title Is Nothing Then
    Set m_Title = New clsField
End If
'm_Title.FontName = gFontName
m_Title.SetAttrib my_printer       ' Set the font, color

With my_printer
        .CurrentY = TITLE_TOP
        .CurrentX = (.ScaleWidth - .TextWidth(CompanyName)) / 2
        ' If a rectangle specified, print it.
        If m_Title.Rectangle Then
            PrintRectangle m_Title, my_printer.CurrentX, my_printer.CurrentY
        End If
End With

' Print the title
my_printer.Print CompanyName
my_printer.CurrentX = (my_printer.ScaleWidth - my_printer.TextWidth(m_Title.Name)) / 2
If m_Title.Rectangle Then
    PrintRectangle m_Title, my_printer.CurrentX, my_printer.CurrentY
End If
my_printer.Print m_Title.Name

' Set the attributes for printing the company name.
my_printer.FontSize = 12
my_printer.CurrentX = (my_printer.ScaleWidth - my_printer.TextWidth(CompanyName)) / 2
my_printer.FontUnderline = False
'my_printer.Print m_CompanyName

' If date attribute is specified, print the date.
Dim DateStr As String
If Me.ReportDate Then
    With my_printer
        .FontSize = 10
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontStrikethru = False
    End With
    DateStr = "Date : " & FormatDate(Now)
    my_printer.CurrentX = my_printer.ScaleWidth - my_printer.TextWidth(DateStr) - FIELD_MARGIN
    my_printer.Print DateStr
End If

m_heading_top = my_printer.CurrentY + 500
End Sub
Private Function ReadTemplate() As Boolean
#If DEPRECATED Then

On Error GoTo Err_Line
' Variables.
Dim nFile As Integer
Dim Txt As String

' Open the template file, and read the contents.
nFile = FreeFile
Open m_report_template For Input As nFile
Txt = Input(LOF(nFile), nFile)
Close nFile

' Serialize the template details.
'If Serialize(Me, Txt) Then ReadTemplate = True

Err_Line:
    If Err Then
        MsgBox "ReadTemplate: " & vbCrLf & Err.Description, vbCritical
    End If
    Close       ' Clean up.

#End If
End Function
Public Property Get ReportTitle() As String
    ReportTitle = m_Title.Name
End Property

Public Property Let ReportTitle(ByVal vNewValue As String)

' Initialize the title object, if not already done.
If m_Title Is Nothing Then Set m_Title = New clsField
m_Title.Name = vNewValue
End Property

Public Property Get ReportTemplate() As Variant
'ReportTemplate = m_report_template
End Property

Public Property Let ReportTemplate(ByVal strRepFile As Variant)
#If DEPRECATED Then
' Validate, if the template file exists.
If Dir(strRepFile, vbNormal) <> "" Then
    m_report_template = strRepFile
    ReadTemplate
Else
    MsgBox "Invalid file - " & strRepFile, vbInformation
End If
Exit Property

Err_Line:
    MsgBox "Invalid file - " & strRepFile, vbInformation
    

#End If

End Property

Public Property Get ReportDestination() As String

If TypeOf my_printer Is Printer Then
    ReportDestination = "PRINTER"
ElseIf TypeOf my_printer Is PictureBox Then
    ReportDestination = "PREVIEW"
End If

End Property

Public Property Let FontSize(ByVal NewValue As Single)
m_FontSize = NewValue
End Property

Public Property Get FontSize() As Single
FontSize = m_FontSize
End Property

Public Property Get FontName() As String
FontName = m_FontName
End Property
Public Property Let FontName(ByVal NewValue As String)
'Befor Assignin the value Check For the Valid FontName
Dim TmpFontName As StdFont
Dim Retval As Integer, Count As Integer
Retval = Screen.FontCount
For Count = 1 To Retval
    If NewValue = Screen.Fonts(Count) Then GoTo ExitLine
Next
Err.Raise 50003, "Print Class", "Invalid FontName"

Exit Property

ExitLine:
m_FontName = NewValue
End Property

Public Property Let ReportDestination(ByVal vNewValue As String)
If vNewValue = "PRINTER" Then
    Set my_printer = Printer
ElseIf vNewValue = "PREVIEW" Then
    If PrintDlg Is Nothing Then Set PrintDlg = New frmPrint
    Set my_printer = PrintDlg.picPrint
Else
    MsgBox "Invalid Report destination - " & vNewValue, vbExclamation
End If
End Property


' ---------------------------------------------------------------------------------------------
' This routine parses the contents of report template file, and
' initializes required field headers and their attributes.
' Since this function may be called by itself recursively,
' care has to be exercised while coding.  -  Ravindranath.
'
' Called from ReadTemplate function
' ---------------------------------------------------------------------------------------------
Private Function zzzSerialize(obj As Object, strSerial As String) As Boolean
#If COMMENTED Then
' Set up error handler
On Error GoTo serialize_error

' Variables...
Dim Txt As String
Dim m_fieldObj As clsField
Dim token_name As String, token_value As String
Txt = strSerial

While Txt <> ""         '// While there are tokens to get.

    ' Get a pair of token and values from the string.
    GetToken Txt, token_name, token_value

    Select Case LCase$(token_name)
        Case "report_header"
            ' Create a new report_header object.
            Set m_Title = New clsField
            If Not Serialize(m_Title, token_value) Then GoTo serialize_error
    
        Case "page_header"
            ' Create a page header object.
            Set m_page_header = New clsField
            If Not Serialize(m_page_header, token_value) Then GoTo serialize_error
    
        Case "details_section"
            Set m_fields = New Collection
            If Not Serialize(m_fields, token_value) Then GoTo serialize_error

        Case "field"
            Set m_fieldObj = New clsField
            If Not Serialize(m_fieldObj, token_value) Then GoTo serialize_error
            obj.Add m_fieldObj
            Set m_fieldObj = Nothing

        Case "title"
            obj.Name = token_value

        Case "fontname"
            obj.FontName = token_value
        Case "fontsize"
            obj.FontSize = token_value
            
        Case "fontbold"
            obj.FontBold = IIf((StrComp(token_value, True, vbTextCompare) = 0), True, False)
            
        Case "forecolor"
            obj.Color = token_value

        Case "name"
            obj.Name = token_value

        Case "align"
            obj.Align = token_value
    
        Case "rectangle"
            obj.Rectangle = IIf((StrComp(token_value, "true", vbTextCompare) = 0), True, False)
        
        Case "sum"
            obj.SumField = IIf((StrComp(token_value, "true", vbTextCompare) = 0), True, False)
            
        Case "date"
            obj.DatePrint = IIf((StrComp(token_value, "true", vbTextCompare) = 0), True, False)
            
        Case "pagenumbers"
            obj.PageNums = IIf((StrComp(token_value, "true", vbTextCompare) = 0), True, False)

    End Select

Wend
Serialize = True
Exit Function

serialize_error:
    If Err Then
        MsgBox "Serialize : " & vbCrLf & Err.Description, vbCritical
    End If
'Resume

#End If
End Function
Private Function TruncateData(srcString As String, fldwidth As Single) As String
On Error Resume Next
If my_printer.TextWidth(srcString) > fldwidth Then
    While my_printer.TextWidth(srcString) > fldwidth And srcString <> ""
        srcString = Left$(srcString, Len(srcString) - 1)
    Wend
    srcString = Left$(srcString, Len(srcString) - 3) & "..."
End If
TruncateData = srcString
End Function
Private Sub Class_Initialize()

' Set the destination print object.
Set PrintDlg = New frmPrint
Set my_printer = PrintDlg.picPrint

' Create a title object.
Set m_Title = New clsField
m_Title.FontName = gFontName
Me.PageNumber = 1

End Sub
Public Property Set DataSource(ByVal grd As MSFlexGrid)
    
 ' This old method assigning a recordset has been
 ' deprecated.
' Set m_DataSource = rs
Set m_DataSource = grd

End Property

Private Sub Class_Terminate()
If Not PrintDlg Is Nothing Then
    Unload PrintDlg
End If
Set PrintDlg = Nothing
End Sub
Private Sub PrintDlg_ProcessEvent(eventNo As Integer)

' Declare variables...
Dim curPage As Integer

Select Case eventNo
    Case wis_SHOW_PREVIOUS
        If PageNumber = 1 Then Exit Sub
        m_DataSource.Row = IIf((page_start_row - rows_per_page) >= 0, _
                (page_start_row - rows_per_page), 0)
            title_printed = False
        PageNumber = PageNumber - 1
        If PageNumber = 1 Then title_printed = False
        PrintPage

    Case wis_SHOW_NEXT
        If print_EOF Then Exit Sub
        PageNumber = PageNumber + 1
        PrintPage
        
    Case wis_SHOW_FIRST
        ' If already at page 1, exit.
        If PageNumber = 1 Then Exit Sub
        m_DataSource.Row = 1
        PageNumber = 1
        PrintPage

    Case wis_SHOW_PREVIOUS
        If page_start_row - rows_per_page > 1 Then
            m_DataSource.Row = page_start_row - rows_per_page
        Else
            m_DataSource.Row = 1
            title_printed = False
        End If
        PrintPage

    Case wis_SHOW_LAST
        m_DataSource.Row = m_DataSource.Rows - rows_in_lastpage + 2
        PageNumber = num_pages
        PrintPage

    Case wis_PRINT_CURRENT
        ' Save the current settings...
        Call SaveSettings
        'PageNumber = PageNumber - 1        ' What is the necessity ???
        print_EOF = False
        PrintTitle
        PrintPage
        ' Restore the saved settings...
        Call RestoreSettings

    Case wis_PRINT_CURRENT_PAUSE
        pause_between_pages = True
        Call SaveSettings
        print_EOF = False
        PrintReport
        Call RestoreSettings
    
    Case wis_PRINT_ALL, wis_PRINT_ALL_PAUSE
        If eventNo = wis_PRINT_ALL_PAUSE Then
            pause_between_pages = True
        End If
        'm_DataSource.Row = page_start_row
        Call SaveSettings
        PageNumber = 1
        print_EOF = False
        PrintReport
        Call RestoreSettings
        
    Case wis_Print_Excel
'        Call PrintToExcel(m_DataSource, ReportTitle)
End Select

End Sub
