VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   2655
   ClientTop       =   3975
   ClientWidth     =   4530
   Icon            =   "Progress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   1776
      Width           =   4284
      _ExtentX        =   7567
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   45
      Top             =   720
   End
   Begin VB.Shape f 
      BorderWidth     =   2
      Height          =   2148
      Left            =   48
      Top             =   48
      Width           =   4452
   End
   Begin VB.Label lblDDE 
      Caption         =   "For DDE connection"
      Height          =   324
      Left            =   864
      TabIndex        =   2
      Top             =   1392
      Visible         =   0   'False
      Width           =   2964
   End
   Begin VB.Image imgProgress 
      Height          =   345
      Index           =   0
      Left            =   420
      Stretch         =   -1  'True
      Top             =   120
      Width           =   405
   End
   Begin VB.Label lblProgress 
      Caption         =   "sample data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   150
      Width           =   3255
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fCount As Integer

Sub DeleteiRiSDLLS()
Dim WinPath As String

'winpath = getwindowsdirectory(

End Sub

Sub RemoveProgramGroup(strGroupName As String)
Dim intIdx As Integer

    '
    'Initialize for DDE Conversation with Windows Program Manager in
    'manual mode (.LinkMode = 2) where destination control is not auto-
    'matically updated.  Set DDE timeout for 10 seconds.  The loop around
    'DoEvents() is to allow time for the DDE Execute to be processsed.
    '

    Dim intRetry As Integer
    For intRetry = 1 To 20
        On Error Resume Next
        lblDDE.LinkTopic = "PROGMAN|PROGMAN"
        If Err = 0 Then
            Exit For
        End If
        DoEvents
    Next intRetry

    lblDDE.LinkMode = 2     '(Manual mode)
    For intIdx = 1 To 10
      DoEvents
    Next
    lblDDE.LinkTimeout = 100

    On Error Resume Next
    lblDDE.LinkExecute "[DELETEGROUP(" & strGroupName & ")]"


End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'This function deletes   files from a given directory according to
' specification and directory if no more files are there
'or all files and directory if second argument is not existing
'Return Value : returns 1 when success
'             : returns -1 on error
'             : returns 0 when directory does not exist
' Contributed by gyan on 6-4-98
' Recontributed by Ravindra on 22/05/1998.
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function DelTree(DirName As String) As Integer
Dim fileNames As String
Dim strCurDrive As String
Dim strCurDir As String
Dim strTmpString As String
Dim strParentDir As String

    On Error Resume Next
    DelTree = -1            'setting function value to Error
    
    'Check if the full path is specified.
    If InStr(DirName, ":") = 0 Then
        MsgBox "Mention full path including the drive letter", vbInformation
        DelTree = -2
        Exit Function
    End If
    
    If Dir(DirName, vbDirectory) = "" Then 'If given directory doesn't exist
        DelTree = 0
        Exit Function
    End If
    
    'Save the current drive and directory
    strCurDrive = Left$(CurDir, 1)
    strCurDir = CurDir
    
    'Change to the given drive.
    ChDrive (Left$(DirName, 1))
    
    'Save the parent directory name of the given directory.
    ChDir DirName & "\.."
    strParentDir = CurDir
    
    Err.Number = 0                      'Initialize the error handler.
    ChDir DirName                       'Change to the given directory.
    If Err.Number > 0 Then         'If could not change to the given directory
        DelTree = -1                        'return error
        Exit Function
    End If
    
DeleteSubDirectories:
    'If there are files then delete them.
    Kill CurDir & "\*.*"
    Err.Number = 0: Err.Description = ""        'Ignore errors
    'Check for any sub-directories in the given directory.
    strTmpString = Dir(CurDir & "\*.*", vbDirectory)
    
    'Ignore current directory and encompassing directories.
    If strTmpString = "." Then
        strTmpString = Dir
    End If
    If strTmpString = ".." Then
        strTmpString = Dir
    End If
    
    'If there are sub-directories then loop back to 'DeleteSubDirectories'
    If strTmpString <> "" Then
        If (GetAttr(strTmpString) And vbDirectory) = vbDirectory Then
                'Change to the sub-directory.
                ChDir CurDir & "\" & strTmpString
                GoTo DeleteSubDirectories
        End If
    End If
    
    'If there are no more directories, change to the parent directory,
    'and remove it.  If the parent directory is the input directory
    'exit the function.
    strTmpString = CurDir                           'Get the name of the current directory.
    ChDir (CurDir & "\..")                              'Change to the parent.
    RmDir strTmpString                                'Remove the sub-directory.
    If UCase$(CurDir) <> UCase(strParentDir) Then         'If we are not at the parent level of 'DirName'
            GoTo DeleteSubDirectories           'Continue deleting the subdirectories.
   End If
   
    'We are done with deleting the sub-directories, so return success.
    DelTree = 1
    
    'Restore the original drive.
    If UCase$(strCurDrive) <> UCase$(Left$(DirName, 1)) Then
        ChDrive strCurDrive
    End If
    
    'Restore the original directory.
    If UCase$(strCurDir) <> UCase$(DirName) Then
        ChDir strCurDir
    End If

End Function

'Routine to remove the application folder

Sub RemoveApplicationFolder()
Dim Retstr As String
Dim RetVal As Integer

Retstr = GetRegistryValue(HKEY_LOCAL_MACHINE, gRegPath & "\Settings", "InstDir")
    If Trim$(Retstr) = "" Then    'Can be null if it is a demo version
        Exit Sub
    Else
        RetVal = DelTree(Retstr)
    End If


End Sub


Sub RemoveRegistryEntries()
Dim RetBool As Boolean

RetBool = DeleteRegistryKey(HKEY_LOCAL_MACHINE, gRegPath & "\OPTIONS")
RetBool = DeleteRegistryKey(HKEY_LOCAL_MACHINE, gRegPath & "\Preprinted")
RetBool = DeleteRegistryKey(HKEY_LOCAL_MACHINE, gRegPath & "\WinPos")
RetBool = DeleteRegistryKey(HKEY_LOCAL_MACHINE, gRegPath & "\Settings")
RetBool = DeleteRegistryKey(HKEY_LOCAL_MACHINE, gRegPath)
End Sub


Function StartUninstall() As Boolean
Dim Count As Integer
Dim i As Integer

ProgressBar1.Value = 0

End Function

Sub Uninstall(Condition As Boolean, Count As Integer)
    If Not Condition Then
        Exit Sub
    End If

    Select Case Count
        Case 1  ' Remove the IRIS folder
            Call RemoveApplicationFolder
            
        Case 2:
            Call DeleteiRiSDLLS
            
        Case 3:
            Call RemoveRegistryEntries
            
        Case 4:
            Call RemoveProgramGroup("iRiS 3.0")
            
    End Select
End Sub

Private Sub Form_Load()
Dim RetBool As Boolean
Dim Count As Integer

'Initialize...
    
    Me.Left = (frmUninst.Width - Me.Width) \ 2
    Me.Top = (frmUninst.Height - Me.Height) \ 2
    'Me.Left = (Screen.Width - Me.Width) \ 2
    'Me.Top = (Screen.Height - Me.Height) \ 2
    
    Me.Caption = gAppName & " - Uninstall"
    Timer1.Interval = 200
    lblProgress(0).Caption = "Deleting program folders."
    lblProgress(0).Top = imgProgress(0).Top + (imgProgress(0).Height \ 2)
    For Count = 1 To 3
    Load lblProgress(Count): lblProgress(Count).Visible = True
    Load imgProgress(Count): imgProgress(Count).Visible = True
    'lblProgress(Count).Top = lblProgress(Count - 1).Top + lblProgress(Count - 1).Height * 2
    imgProgress(Count).Top = imgProgress(Count - 1).Top + imgProgress(Count - 1).Height
    lblProgress(Count).Top = imgProgress(Count).Top + (imgProgress(Count).Height \ 2)
    
    Next Count
    
    lblProgress(1).Caption = "Deleting dependency files."
    lblProgress(2).Caption = "Deleting registry information."
    lblProgress(3).Caption = "Removing program icons"
    lblProgress(0).ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()

ProgressBar1.Value = ProgressBar1.Value + 10
        
    If ProgressBar1.Value = 100 Then
        imgProgress(fCount).Picture = LoadResPicture(101, vbResIcon)
        ProgressBar1.Value = 0
        fCount = fCount + 1
        Call Uninstall(True, fCount)
        If fCount = 4 Then
            ProgressBar1.Value = 100
            lblProgress(fCount - 1).ForeColor = vbBlack
            Timer1.Enabled = False
            MsgBox gAppName & " has been uninstalled." & vbCrLf & _
            vbCrLf & "It is important that you restart your computer now.", vbInformation, gAppName & " - Uninstall Message"
            On Error Resume Next
            Unload Me
            End
        End If
                lblProgress(fCount - 1).ForeColor = vbBlack
                lblProgress(fCount).ForeColor = vbRed

    End If

End Sub

