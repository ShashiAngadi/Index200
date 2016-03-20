Attribute VB_Name = "basUninst"
Option Explicit

Public Const gAppName = "iRiS"
Public Const gRegPath = "Software\integra\iris"
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Type Directories
    FileArray() As String
End Type

'*************************************************************************************************************
'                                                       GET ALL FILES
'*************************************************************************************************************
'
'   This function gets a list of files (including the subdirectories ) into an array given the file path.
''   Note : This function uses the structure
'               Type Directories
'                       FileArray() As String
'               End Type
'  Last Modified By : Girish. S. Desai
'  Modified On         : 18 May, 1998
'  Dependencies     :  1.     Structure  : Directories
'                               2.    Function  : CheckFileType
'                               3.    DoesPathExist(FileName)
'
Function GetAllFiles(FilePath As String, FileArr() As String, Optional IncludeSubDirs As Boolean) As Integer

Dim FileName As String
Dim PathNum As Integer
Dim Flag As Boolean
Dim FileType As String
Dim Count As Integer
Dim AttrVal As Integer
Dim MyPath() As Directories
Dim IncSubDir As Boolean
Dim i As Integer
Dim CountInThisLoop As Integer
Const MAXSIZE = 32765

GetAllFiles = -1
ReDim MyPath(0)
ReDim MyPath(0).FileArray(0)
    
        IncSubDir = False
        If Not IsMissing(IncludeSubDirs) Then
            If IncludeSubDirs Then
                IncSubDir = True
            End If
        End If

'Check if the drive letter is valid
    If Len(FilePath) < 3 Then
        Exit Function
    End If
    
    If Mid(FilePath, 2, 2) <> ":\" Then
        Exit Function
    End If
    

FileName = FilePath

If Right(FileName, 1) = "\" Then
    FileName = Left(FileName, Len(FileName) - 1)
End If

    For Count = 1 To Len(FileName)
        'If Mid(FileName, Count, 1) <> "\" Then
        If Left(Right(FileName, Count), 1) <> "\" Then
            FileType = Right(FileName, Count)
        Else
            Exit For
        End If
    Next Count
    
    If InStr(1, FileType, "*") = 0 And InStr(1, FileType, "?") = 0 Then
        FileName = FileName
        FileType = "*.*"
    Else
        FileName = Left(FileName, Len(FileName) - Len(FileType) - 1)
    End If
    
'Check if the specified path exists..
    If DoesPathExist(FileName) = 0 Then
        Exit Function
    End If
    
CountInThisLoop = 0
MyPath(0).FileArray(0) = FileName
        
        Do
                Flag = False
                    
                    If Not IncSubDir And CountInThisLoop = 1 Then
                        Exit Do
                    Else
                        CountInThisLoop = 1
                    End If
                
                    For PathNum = 0 To UBound(MyPath)
                            
                            
                            If MyPath(PathNum).FileArray(0) <> "" Then
                                Flag = True
                                Exit For
                            End If
                    Next PathNum
            
                    If Not Flag Then
                        Exit Do
                    End If
                    
                FileName = Dir(MyPath(PathNum).FileArray(0) & "\*.*", vbDirectory)

                Do
                        If FileName <> "." And FileName <> ".." And FileName <> "" Then
                            AttrVal = GetAttr(MyPath(PathNum).FileArray(0) & "\" & FileName)
                            If AttrVal >= 16 And AttrVal <= 23 Then
                                If UBound(MyPath) + 1 > MAXSIZE Then
                                    GetAllFiles = -2
                                    Exit Function
                                End If
                                ReDim Preserve MyPath(UBound(MyPath) + 1)
                                ReDim Preserve MyPath(UBound(MyPath)).FileArray(0)
                                MyPath(UBound(MyPath)).FileArray(0) = MyPath(PathNum).FileArray(0) & "\" & FileName
                            Else
                                If CheckFileType(MyPath(PathNum).FileArray(0) & "\" & FileName, FileType) Then
                                            If UBound(MyPath(PathNum).FileArray) + 1 > MAXSIZE Then
                                                GetAllFiles = -2
                                                Exit Function
                                            End If
                                    ReDim Preserve MyPath(PathNum).FileArray(UBound(MyPath(PathNum).FileArray) + 1)
                                    MyPath(PathNum).FileArray(UBound(MyPath(PathNum).FileArray)) = MyPath(PathNum).FileArray(0) & "\" & FileName
                                End If
                                
                            End If
                        End If
                    If FileName <> "" Then
                    FileName = Dir()
                    End If
                Loop Until FileName = ""
                MyPath(PathNum).FileArray(0) = ""
        Loop

'Merge all the resulting arrays into one single array
Dim Nooffiles
Nooffiles = 0
ReDim FileArr(Nooffiles)
    For Count = 0 To UBound(MyPath)
        For i = 1 To UBound(MyPath(Count).FileArray)
            If MyPath(Count).FileArray(i) <> "" Then
                        If UBound(FileArr) + 1 > MAXSIZE Then
                            GetAllFiles = -2
                            Exit Function
                        End If
                ReDim Preserve FileArr(Nooffiles)
                FileArr(UBound(FileArr)) = MyPath(Count).FileArray(i)
                Nooffiles = Nooffiles + 1
            End If
        Next i
    Next Count
    
    If FileArr(0) = "" Then
        GetAllFiles = 0
    Else
        GetAllFiles = UBound(FileArr) + 1
    End If

End Function


Function GetWinDir() As String
Dim Retstr As String
Dim RetVal As Integer

    Retstr = String$(255, 0)
    RetVal = GetWindowsDirectory(Retstr, Len(Retstr))
    Retstr = Left(Retstr, RetVal)
    GetWinDir = Retstr
    
End Function

'**************************************************************************************************************
'                                                   CHECK FILE TYPE
'**************************************************************************************************************
'
'   This functions checks a file name if it is in accordance to the the specified type
'    For Example : This function returns true if
'           FileName = c:\integra\Fido2001\anteres.txt
'           FileType = "ant*.txt"
'
' Last Modified On :   18 May, 1998.
' Modified By         : Girish. S. Desai
' Dependencies      : <NONE>
'
Function CheckFileType(FileName As String, FileType As String) As Boolean

    Dim Type1 As String
    Dim Type2 As String
    Dim File1 As String
    Dim File2 As String
    Dim Pos As Integer
    Dim PureName As String
    Dim Count As Integer
    
    CheckFileType = False
    
    If FileName = "" Or FileType = "" Then
        Exit Function
    End If

        Pos = InStr(1, FileType, ".")
        If Pos <> 0 Then
            Type1 = Left(FileType, Pos - 1)
            Type2 = Right(FileType, Len(FileType) - Pos)
        Else
            Type1 = FileType
            Type2 = "*"
        End If
        
        For Count = 1 To Len(FileName)
                If Left(Right(FileName, Count), 1) <> "\" Then
                    PureName = Right(FileName, Count)
                Else
                    Exit For
                End If
        Next Count
    
        Pos = InStr(1, PureName, ".")
        If Pos <> 0 Then
            File1 = Left(PureName, Pos - 1)
            File2 = Right(PureName, Len(PureName) - Pos)
        Else
            File1 = PureName
            File2 = ""
        End If
            
'Do first check
        If Type2 = "" And File2 <> "" Then
            Exit Function
        End If
       ' If Type2 <> "" And File2 = "" Then
       '     Exit Function
       ' End If
        
'Check for part1
        Pos = InStr(1, Type1, "*")
            If Pos = 0 Then
                If Len(Type1) <> Len(File1) Then
                    Exit Function
                End If
                For Count = 1 To Len(Type1)
                    If Right(Left(Type1, Count), 1) <> "?" Then
                        If UCase(Right(Left(Type1, Count), 1)) <> UCase(Right(Left(File1, Count), 1)) Then
                            Exit Function
                        End If
                    End If
                Next Count
            Else
                If UCase(Left(Type1, Pos - 1)) <> UCase(Left(File1, Pos - 1)) Then
                    Exit Function
                End If
            End If
            
'Check for part2
'If Type2 = "" Or File2 = "" Then
'    Exit Function
'End If
        Pos = InStr(1, Type2, "*")
            If Pos = 0 Then
                If Len(Type2) <> Len(File2) Then
                    Exit Function
                End If
                For Count = 1 To Len(Type2)
                    If Right(Left(Type2, Count), 1) <> "?" Then
                        If UCase(Right(Left(Type2, Count), 1)) <> UCase(Right(Left(File2, Count), 1)) Then
                            Exit Function
                        End If
                    End If
                Next Count
            Else
                If UCase(Left(Type2, Pos - 1)) <> UCase(Left(File2, Pos - 1)) Then
                    Exit Function
                End If
            End If


CheckFileType = True
End Function

'***********************************************************************
'                           DOES PATH EXIST
'
''***********************************************************************
'Function to check if the path.
'Returns 0 if path does not exist
'Returns 1 if it is a file
'Returns -1 if it is read only file
'Returns 2 if it is a directory
'Returns -2 if it is a read only directory


Function DoesPathExist(Path As String) As Integer

On Error GoTo Errline
Dim RetVal As Integer
 
  RetVal = GetAttr(Path)
    If RetVal >= 32 Then
        RetVal = RetVal - 32
    End If
    
    If RetVal >= 17 Then
        DoesPathExist = -2 'Read Only Directory
        Exit Function
    End If
        
    If RetVal >= 16 Then
        DoesPathExist = 2 'Normal Only Directory
        Exit Function
    End If
    
    If RetVal = 1 Then
        DoesPathExist = -1  'Read Only File
    Else
        DoesPathExist = 1   'Normal File
    End If
    
Exit Function
Errline:
    DoesPathExist = 0
End Function

