VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPhoto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Photo"
   ClientHeight    =   4440
   ClientLeft      =   1515
   ClientTop       =   2040
   ClientWidth     =   6750
   Icon            =   "frmPhoto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSgnDel 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5820
      TabIndex        =   7
      Top             =   3390
      Width           =   540
   End
   Begin VB.CommandButton cmdSgnNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4500
      TabIndex        =   5
      Top             =   3390
      Width           =   540
   End
   Begin VB.CommandButton cmdSgnAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5160
      TabIndex        =   6
      Top             =   3390
      Width           =   540
   End
   Begin VB.CommandButton cmdSgnPrev 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3825
      TabIndex        =   4
      Top             =   3390
      Width           =   540
   End
   Begin VB.CommandButton cmdImgDel 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2370
      TabIndex        =   3
      Top             =   3375
      Width           =   540
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImgPrev 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   375
      TabIndex        =   0
      Top             =   3375
      Width           =   540
   End
   Begin VB.CommandButton cmdImgAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1695
      TabIndex        =   2
      Top             =   3375
      Width           =   540
   End
   Begin VB.CommandButton cmdImgNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1050
      TabIndex        =   1
      Top             =   3375
      Width           =   540
   End
   Begin VB.Line Line9 
      X1              =   3540
      X2              =   6570
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line Line8 
      X1              =   3540
      X2              =   6570
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line7 
      X1              =   6570
      X2              =   6570
      Y1              =   3210
      Y2              =   270
   End
   Begin VB.Line Line6 
      X1              =   3540
      X2              =   3540
      Y1              =   270
      Y2              =   3210
   End
   Begin VB.Line Line5 
      X1              =   180
      X2              =   3210
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line Line4 
      X1              =   180
      X2              =   3210
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line3 
      X1              =   3210
      X2              =   3210
      Y1              =   270
      Y2              =   3210
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   180
      Y1              =   270
      Y2              =   3210
   End
   Begin VB.Image picphoto 
      Height          =   2955
      Left            =   180
      Stretch         =   -1  'True
      Top             =   270
      Width           =   3045
   End
   Begin VB.Image picSign 
      Height          =   2955
      Left            =   3540
      Stretch         =   -1  'True
      Top             =   270
      Width           =   3045
   End
   Begin VB.Label lblSgnDate 
      Caption         =   "Date: 13/06/2013"
      Height          =   345
      Left            =   3600
      TabIndex        =   11
      Top             =   3930
      Width           =   1680
   End
   Begin VB.Label lblImgDate 
      Caption         =   "Date: 2/4/2013"
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   3960
      Width           =   1560
   End
   Begin VB.Label lblSgnCount 
      Alignment       =   2  'Center
      Caption         =   "Signature: 0/0"
      Height          =   345
      Left            =   5190
      TabIndex        =   9
      Top             =   3930
      Width           =   1530
   End
   Begin VB.Label lblImgCount 
      Alignment       =   2  'Center
      Caption         =   "Photo: 0/0"
      Height          =   315
      Left            =   1710
      TabIndex        =   8
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   3375
      X2              =   3375
      Y1              =   270
      Y2              =   4170
   End
End
Attribute VB_Name = "frmPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form level variables...
Dim fAccNo As String

Dim fCurImgPhoto As String
Dim fCurPhotoNum As Integer
Dim fCurPhotoDate As String

Dim fCurImgSignature As String
Dim fCurSignatureNum As Integer
Dim fCurSignatureDate As String

' Form level variables
Private signatures() As String
Private photos() As String

Public Sub PhotoInitialize()

If Len(fAccNo) = 0 Then Exit Sub

On Error GoTo Hell
    
    ' Read the configuration file for finding out the path of image files.
    Dim ret As Long
    Dim strRet As String
    cmdImgNext.Enabled = False
    cmdSgnNext.Enabled = False
    cmdImgPrev.Enabled = False
    cmdSgnPrev.Enabled = False
    cmdImgDel.Enabled = False
    cmdSgnDel.Enabled = False
    
    'Load the image names (photos and signatures) into array
    If (loadImageFiles(fAccNo, INDEX2000_PHOTO, photos) = False) Then
        MsgBox "Error while loading photos.", vbOKOnly
    End If
    
    If (loadImageFiles(fAccNo, INDEX2000_SIGN, signatures) = False) Then
        MsgBox "Error while loading signatures.", vbOKOnly
    End If
    fCurPhotoNum = 0
    ' Display the first picture if there exists one.
    If (UBound(photos) > 0) Then
        fCurPhotoNum = 1
        picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
        cmdImgDel.Enabled = True
    ElseIf Len(Dir(gImagePath & "image_default.jpg")) > 0 Then
        picphoto.Picture = VB.LoadPicture(gImagePath & "image_default.jpg")
        fCurPhotoNum = 0
    End If
    
    fCurSignatureNum = 0
    If (UBound(signatures) > 0) Then
        fCurSignatureNum = 1
        picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
        cmdSgnDel.Enabled = True
    ElseIf Len(Dir(gImagePath & "sign_default.jpg")) Then
        picSign.Picture = VB.LoadPicture(gImagePath & "sign_default.jpg")
        fCurSignatureNum = 0
        cmdSgnDel.Enabled = False
    End If
    
    ' Since the first img will be loaded, there is no prev image to display,
    ' disable the prev button.
    
    'aswell as signature count label
    setPhotosLabel

Hell:

End Sub


Private Sub cmdImgAdd_Click()

    CommonDialog1.Filter = "Pictures (*.bmp;*.jpg;*.tif)|*.jpg;|*.tif;"
    CommonDialog1.ShowOpen
    Dim imgfile As String
    
    If (CommonDialog1.Filename = "") Then Exit Sub
    imgfile = AddImageFile(fAccNo, INDEX2000_PHOTO, CommonDialog1.Filename)
    
    Dim pos As Integer
    pos = InstrRev(imgfile, "\")
    
    If pos Then imgfile = Mid$(imgfile, pos + 1)
    
    Dim newIndex As Integer
    newIndex = UBound(photos) + 1
    ReDim Preserve photos(newIndex)
        photos(newIndex) = imgfile
        fCurPhotoNum = fCurPhotoNum + 1
        picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
    
        
    PhotoInitialize
    
End Sub
Private Sub cmdImgDel_Click()

Dim result As VbMsgBoxResult
    result = MsgBox(GetResourceString(830), vbYesNo, wis_MESSAGE_TITLE)
If (result = vbNo) Then
    Exit Sub
End If
 

On Error GoTo Hell
 
'Now delete the image from file system.
Kill gImagePath & photos(fCurPhotoNum)
For I = fCurPhotoNum To UBound(photos) - 1
       photos(I) = photos(I - 1)
       picphoto.Picture = VB.LoadPicture(gImagePath & photos(I))
Next I

Set picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
If fCurPhotoNum < 1 Then
    ReDim Preserve photos(UBound(photos) - 1)
    Set picphoto.Picture = VB.LoadPicture(gImagePath, "\image_default.jpg")
End If

' Need to handle some border cases.
' if the number of images left after deleting the current one is equal to 1,
' we just keep the prev/next button on, without disabling it.

Hell:

PhotoInitialize
End Sub

Private Sub cmdImgNext_Click()
Dim imgfile As String

'photo display
If (fCurPhotoNum < UBound(photos)) Then
    fCurPhotoNum = fCurPhotoNum + 1
    picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
End If

setPhotosLabel

End Sub
Private Sub cmdImgPrev_Click()
 
    fCurPhotoNum = fCurPhotoNum - 1
    picphoto.Picture = VB.LoadPicture(gImagePath & photos(fCurPhotoNum))
    

setPhotosLabel

End Sub

Private Sub cmdSgnAdd_Click()

    CommonDialog1.Filter = "Pictures (*.bmp;*.jpg;*.tif)|*.jpg;|*.tif;"
    
    CommonDialog1.ShowOpen
    Dim imgfile As String

    If (CommonDialog1.Filename = "") Then Exit Sub

    imgfile = AddImageFile(fAccNo, INDEX2000_SIGN, CommonDialog1.Filename)
    Dim pos As Integer
    pos = InstrRev(imgfile, "\")
    If pos Then imgfile = Mid$(imgfile, pos + 1)
    
    Dim newIndex As Integer
    newIndex = UBound(signatures) + 1
    ReDim Preserve signatures(newIndex)
    signatures(newIndex) = imgfile
    fCurSignatureNum = fCurSignatureNum + 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
    
    
PhotoInitialize

End Sub
Private Sub cmdSgnDel_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox(GetResourceString(831), vbYesNo, wis_MESSAGE_TITLE)
    If (result = vbNo) Then Exit Sub
 
    On Error GoTo Hell

    ' Now delete the image from file system.
    Kill gImagePath & signatures(fCurSignatureNum)
    
    
Hell:
    
PhotoInitialize

End Sub


Private Sub cmdSgnNext_Click()

Dim imgfile As String
'signature display

If (fCurSignatureNum < UBound(signatures)) Then
    fCurSignatureNum = fCurSignatureNum + 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
End If

setPhotosLabel

End Sub
Private Sub cmdSgnPrev_Click()
    
    fCurSignatureNum = fCurSignatureNum - 1
    picSign.Picture = VB.LoadPicture(gImagePath & signatures(fCurSignatureNum))
    
setPhotosLabel
End Sub
Private Sub SetKannadaCaption()
    Call SetFontToControls(Me)
    lblImgDate = GetResourceString(415, 37)
    lblSgnDate = GetResourceString(416, 37)
    lblImgCount = GetResourceString(415, 50)
    lblSgnCount = GetResourceString(416, 50)
    
End Sub

Private Sub Form_Load()
Call CenterMe(Me)

Call SetKannadaCaption

Call PhotoInitialize

' Update the image count label
setPhotosLabel


End Sub

Public Function setAccNo(acno As String)
fAccNo = acno
If Len(fAccNo) = 0 Then
    cmdImgAdd.Enabled = False
    cmdSgnAdd.Enabled = False
End If

End Function

Private Sub setPhotosLabel()
    If (UBound(photos) > 0) Then
        lblImgCount.Caption = GetResourceString(415) & ":" & fCurPhotoNum & "/" & UBound(photos)
    End If
    If (UBound(photos) > 0) Then
        lblImgDate.Caption = GetResourceString(37) & ":" & FileDateTime(gImagePath & photos(fCurPhotoNum))
    End If
    If (UBound(signatures) > 0) Then
        lblSgnCount.Caption = GetResourceString(416) & ":" & fCurSignatureNum & "/" & UBound(signatures)
    End If
    If (UBound(signatures) > 0) Then
        lblSgnDate.Caption = GetResourceString(37) & ":" & FileDateTime(gImagePath & signatures(fCurSignatureNum))
    End If
    
cmdImgNext.Enabled = (UBound(photos) > fCurPhotoNum) And UBound(photos) > 0
cmdSgnNext.Enabled = (UBound(signatures) > fCurSignatureNum) And UBound(signatures) > 0

cmdImgPrev.Enabled = fCurPhotoNum > 1
cmdSgnPrev.Enabled = fCurSignatureNum > 1
   
End Sub

