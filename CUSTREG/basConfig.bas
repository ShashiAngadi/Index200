Attribute VB_Name = "basConfig"

Function GetMatchingImgFileCount(acno As String, imgType As Integer) As Integer
Dim strImgType As String
If imgType = 1 Then
   strImgType = "photo"
Else
    strImgType = "sign"
End If
Dim strFileName As String
Dim file As String
strFileName = gImagePath + "\index2000_" + strImgType + "_" + acno + "_*"
file = Dir$(strFileName, vbNormal)
Dim fileCount As Integer
fileCount = 0
Do While Len(file)
  fileCount = fileCount + 1
  file = Dir
Loop

GetMatchingImgFileCount = fileCount
End Function
Function GetImageFilesCount(strPath As String)

Dim nCount As Integer
Dim strFileArr() As String
nCount = GetImageFilesArray(strPath, strFileArr)
GetImageFilesCount = nCount

End Function
Public Function GetImageFilesArray(strPath As String, strFileArray() As String)

Dim strExtn As String, strExtns() As String, file As String
Dim nCount As Integer, nPosExt As Integer, nPos As Integer
nCount = GetStringArray(INDEX2000_IMG_FILE_TYPES, strExtns, "|")
ReDim strFileArray(0)
file = Dir$(strPath)

Do While Len(file)
    
    'Extract the extension of this file
    nPosExt = InstrRev(file, ".")
    If (nPosExt <= 0) Then
        GoTo GetNextFile 'Get next file if we dont have an extension
    End If
    strExtn = Mid(file, nPosExt + 1, Len(file) - nPosExt)
    
    'File number should be compliant to spec
    Dim nFileNo As Integer
    nFileNo = GetImageFileNumber(file)
    If (nFileNo < 0) Then
        GoTo GetNextFile
    End If
    
    For I = 0 To UBound(strExtns)
        'Compare the ext with the standard list
        If (StrComp(strExtn, strExtns(I), vbTextCompare) = 0) Then
            ReDim Preserve strFileArray(UBound(strFileArray) + 1)
            strFileArray(UBound(strFileArray)) = file
            Exit For
        End If
    Next I
    
GetNextFile:
    file = Dir$
Loop
nRet = BubbleSort(strFileArray)

GetImageFilesArray = UBound(strFileArray)
End Function
Function BubbleSort(strFileArray() As String)

Dim bSwapped As Boolean
Dim strTmp As String
bSwapped = True
Do While bSwapped = True
    bSwapped = False
    For I = 1 To UBound(strFileArray) - 1
        nVal_1 = GetImageFileNumber(strFileArray(I))
        nVal_2 = GetImageFileNumber(strFileArray(I + 1))
        If (nVal_1 > nVal_2) Then
            strTmp = strFileArray(I)
            strFileArray(I) = strFileArray(I + 1)
            strFileArray(I + 1) = strTmp
            bSwapped = True
        End If
    Next I
Loop

End Function

Function GetImageFileNumber(strImgFile) As Integer

GetImageFileNumber = -1
Dim nExtn As Integer, nPos As Integer
nExtn = InstrRev(strImgFile, ".")
If (nExtn <= 0) Then
    Exit Function
End If

nPos = InstrRev(strImgFile, "_")
If (nPos <= 0) Then
    Exit Function
End If

If (nExtn < nPos) Then
    Exit Function
End If

Dim strFileNo As String
strFileNo = Mid(strImgFile, nPos + 1, nExtn - nPos - 1)
If (IsNumeric(strFileNo) = False) Then
    Exit Function
End If

GetImageFileNumber = Val(strFileNo)

End Function

Function GetImageFileNext(strCurrImageFile As String, nDirection As Integer, ByRef nCount As Integer, ByRef strDate As String) As String

Dim strFile As String
Dim nPos As Integer

nPos = InstrRev(strCurrImageFile, "_")
strFile = Mid(strCurrImageFile, 1, nPos)
strFile = strFile + "*"

Dim strFileArr() As String
Dim nCnt As Integer
nCnt = GetImageFilesArray(strFile, strFileArr)

Dim bFound As Boolean
bFound = False

For I = 1 To UBound(strFileArr)
    If (StrComp(gImagePath + "\" + strFileArr(I), strCurrImageFile, vbTextCompare) = 0) Then
        If (nDirection = 1 And I < UBound(strFileArr)) Then
            strFileName = strFileArr(I + 1)
            nCount = I + 1
        ElseIf nDirection = -1 And I > 1 Then
            strFileName = strFileArr(I - 1)
            nCount = I - 1
        End If
        Exit For
    End If
Next I

GetImageFileNext = ""
If (strFileName <> "") Then
    strFileName = gImagePath + "\" + strFileName
    dateVar = FileDateTime(strFileName)
    'nTotal = UBound(strFileArr)
    strDate = Mid(dateVar, 1, InStr(1, dateVar, " ", vbTextCompare) - 1)
    GetImageFileNext = strFileName
End If

End Function
Function AddImageFile(strAccNo As String, strImgType As String, strSrcPath As String) As String
AddImageFile = ""

'Get Ext of source file
Dim strExt As String
Dim nPos As Integer

'Extract the extension
nPos = InstrRev(strSrcPath, ".")
If (nPos <= 0) Then Exit Function


strExt = Mid(strSrcPath, nPos + 1, Len(strSrcPath) - nPos)

'Check if extn complies to allowed extension types
Dim strExts() As String
nRet = GetStringArray(INDEX2000_IMG_FILE_TYPES, strExts, "|")
For I = 0 To UBound(strExts)
    If (StrComp(strExt, strExts(I), vbTextCompare) = 0) Then
        Exit For
    End If
Next I
If (I > UBound(strExts)) Then
    Exit Function
End If

'Move to the last file in our db
Dim strFile As String, strFileLast As String
Dim nCount As Integer, strDate As String
strFile = GetImageFileFirst(strAccNo, strImgType, nCount, strDate)
Do While (strFile <> "")
    strFileLast = strFile
    strFile = GetImageFileNext(strFile, 1, nCount, strDate)
Loop


If (strFileLast = "") Then
    strFileNo = "1"
Else
    'Extract the image count and increment it
    nPos = InstrRev(strFileLast, "_")
    nPosExt = InstrRev(strFileLast, ".")
    strFileNo = Mid(strFileLast, nPos + 1, nPosExt - nPos - 1)
    strFileNo = Val(strFileNo) + 1
End If
'Form the dest path
strdestpath = gImagePath + "index2000_" + strImgType + "_" + strAccNo + "_" + CStr(strFileNo) + "." + strExt

'Check if file exists
strFile = Dir$(strSrcPath)
If (strFile = "") Then
    Exit Function
End If

FileCopy strSrcPath, strdestpath

AddImageFile = strdestpath
End Function
Function GetImageFileFirst(strAccNo As String, strImgType As String, ByRef nCount As Integer, ByRef strDate As String) As String


GetImageFileFirst = ""

'stick to default photo
If (strImgType <> INDEX2000_PHOTO And strImgType <> INDEX2000_SIGN) Then
    strImgType = INDEX2000_PHOTO
End If

'form the path to search
Dim strFileName As String
Dim strFileArray() As String
Dim nCnt As Integer
Dim nTotal As Integer
strFileName = gImagePath + "\index2000_" + strImgType + "_" + strAccNo + "_*"
nCnt = GetImageFilesArray(strFileName, strFileArray)
If (UBound(strFileArray) > 0) Then
    strFileName = gImagePath + "\" + strFileArray(1)
    dateVar = FileDateTime(strFileName)
    nCount = 1
    nTotal = UBound(strFileArray)
    strDate = Mid(dateVar, 1, InStr(1, dateVar, " ", vbTextCompare) - 1)
    GetImageFileFirst = strFileName
Else
    nCount = 1
    'GetImageFileFirst = Mid(strFileName, 1, Len(strFileName) - 1) & nCount
End If
End Function


Public Function loadImageFiles(acno As String, imgType As String, imagesArray() As String) As Boolean

On Error GoTo Hell

    Dim strFileName As String
    strFileName = gImagePath + "\index2000_" + imgType + "_" + acno + "_*"
    nCnt = GetImageFilesArray(strFileName, imagesArray)
    
    loadImageFiles = True
    Exit Function
Hell:
    loadImageFiles = False
End Function

Public Function loadsignatureFiles(acno As String, imgType As String, imagesArray() As String) As Boolean

On Error GoTo Hell

    Dim strFileName As String
    strFileName = gImagePath + "\index2000_" + imgType + "_" + acno + "_*"
    nCnt = GetImageFilesArray(strFileName, imagesArray)
    
    loadsignatureFiles = True
    Exit Function
Hell:
    loadsignatureFiles = False
End Function
Public Sub DeleteUnregisteredImageFiles()
On Error Resume Next
'Public Const INDEX2000_PHOTO = "photo"
'Public Const INDEX2000_SIGN = "sign"
Dim acno As String
Dim count As Integer
Dim totalCount As Integer
Dim fileNames() As String
Dim strFileName As String
    acno = ""
    'Delete the Photos
    strFileName = gImagePath + "\index2000_" + INDEX2000_PHOTO + "_" + acno + "_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        Kill gImagePath & fileNames(count)
    Next
    'Delete the Images
    strFileName = gImagePath + "\index2000_" + INDEX2000_SIGN + "_" + acno + "_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        Kill gImagePath & fileNames(count)
    Next
    
    'Now Delete assuming the custid was 0
    acno = "0"
    'Delete the Photos
    strFileName = gImagePath + "\index2000_" + INDEX2000_PHOTO + "_" + acno + "_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        Kill gImagePath & fileNames(count)
    Next
    'Delete the Images
    strFileName = gImagePath + "\index2000_" + INDEX2000_SIGN + "_" + acno + "_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        Kill gImagePath & fileNames(count)
    Next
    
End Sub

Public Sub SaveUnregisteredImageFiles(custId As Long)

On Error Resume Next
Dim count As Integer
Dim totalCount As Integer
Dim fileNames() As String
Dim strFileName As String
Dim strNewFileName As String
    'Delete the Photos
    strFileName = gImagePath + "\index2000_" + INDEX2000_PHOTO + "__*"
    strFileName = gImagePath + "\index2000_" + INDEX2000_PHOTO + "_" + CStr(custId) + "_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        strNewFileName = Replace(fileNames(count), "__", "_" + custId + "_")
        Name fileNames(count) As strNewFileName
    Next
    'Delete the Images
    strFileName = gImagePath + "\index2000_" + INDEX2000_SIGN + "__*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        strNewFileName = Replace(fileNames(count), "__", "_" + custId + "_")
        Name fileNames(count) As strNewFileName
    Next
    
    'Now Delete assuming the custid was 0
    'Delete the Photos
    strFileName = gImagePath + "\index2000_" + INDEX2000_PHOTO + "_0_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        strNewFileName = Replace(fileNames(count), "_0_", "_0_")
        Name fileNames(count) As strNewFileName
    Next
    'Delete the Images
    strFileName = gImagePath + "\index2000_" + INDEX2000_SIGN + "_0_*"
    totalCount = GetImageFilesArray(strFileName, fileNames)
    For count = 1 To totalCount
        strNewFileName = Replace(fileNames(count), "_0_", "_" + custId + "_")
        Name fileNames(count) As strNewFileName
    Next
    
End Sub
