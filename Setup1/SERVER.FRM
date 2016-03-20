VERSION 5.00
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Index2000 - Server Name"
   ClientHeight    =   3240
   ClientLeft      =   2430
   ClientTop       =   3630
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3900
      TabIndex        =   9
      Top             =   2850
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   2820
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please specify the name of the server"
      Height          =   2610
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   4890
      Begin VB.TextBox txtShareName 
         Height          =   345
         Left            =   2670
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox cboNetworkProtocol 
         Height          =   315
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   4470
      End
      Begin VB.TextBox txtServer 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   1020
         Width           =   3105
      End
      Begin VB.Label lblshareName 
         Caption         =   "Share Name"
         Height          =   225
         Left            =   270
         TabIndex        =   7
         Top             =   1560
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   810
         TabIndex        =   6
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblNetworkProtocol 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   270
         Width           =   2550
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgWelcome 
         Height          =   480
         Left            =   300
         Picture         =   "SERVER.frx":0000
         Top             =   930
         Width           =   480
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   2220
         Width           =   3885
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RpcNetworkIsProtseqValid Lib "rpcrt4.dll" Alias "RpcNetworkIsProtseqValidA" (ByVal strProtseq As String) As Long

Private Sub cmdCancel_Click()
gstrServerName = ""
gstrShareName = ""
gboolServer = True
MsgBox "The software will be installed to the local machine"
Unload Me
End Sub

Private Sub cmdOK_Click()

gstrServerName = Trim$(txtServer.Text)
gstrShareName = Trim$(txtShareName.Text)

Unload Me
End Sub

Private Sub Form_Load()

SetFormFont Me

CenterForm Me

Caption = gstrAppName & " Client Setup"
lblNetworkProtocol.Caption = ResolveResString(resNETWORKPROTOCOL)
cmdOK.Caption = ResolveResString(resOK)
lblHelp.Caption = "Server name is the name of the computer where the database for " & gstrAppName & " resides." & _
                         "If you are not sure about the server name please consult your system administrator."
                        
'Call FillInProtocols
End Sub



Private Sub FillInProtocols()
    Dim i As Integer
    Dim fSuccessReading As Boolean
    
    cboNetworkProtocol.Clear
    fSuccessReading = ReadProtocols(gstrSetupInfoFile, gstrINI_SETUP)
    If Not fSuccessReading Or gcProtocols <= 0 Then
        MsgError ResolveResString(resNOPROTOCOLSINSETUPLST), vbExclamation Or vbOKOnly, gstrTitle
        ExitSetup Me, gintRET_FATAL
    End If
    
    For i = 1 To gcProtocols
        If fIsProtocolSeqSupported(gProtocol(i).strName, gProtocol(i).strFriendlyName) Then
            cboNetworkProtocol.AddItem gProtocol(i).strFriendlyName
        End If
    Next i

    If cboNetworkProtocol.ListCount > 0 Then
        'We were successful in finding at least one protocol available on this machine
        Exit Sub
    End If
    
    'None of the protocols specified in SETUP.LST are available on this machine.  We need
    'to let the user know what's wrong, including which protocol(s) were expected.
    MsgError ResolveResString(resNOPROTOCOLSSUPPORTED1), vbExclamation Or vbOKOnly, gstrTitle
    '
    ' Don't log the rest if this is SMS.  Ok for silent mode since
    ' silent can take more than 255 characters.
    '
    If Not gfSMS Then
        Dim strMsg As String
        strMsg = ResolveResString(resNOPROTOCOLSSUPPORTED2) & vbLf
    
        For i = 1 To gcProtocols
            strMsg = strMsg & vbLf & Chr$(9) & gProtocol(i).strFriendlyName
        Next i
        
        MsgError strMsg, vbExclamation Or vbOKOnly, gstrTitle
    End If
    ExitSetup Me, gintRET_FATAL
End Sub


' Determines whether a given protocol sequence is supported and available on this machine
Function fIsProtocolSeqSupported(ByVal strProto As String, ByVal strProtoFriendlyName) As Boolean
    Const RPC_S_OK = 0&
    Const RPC_S_PROTSEQ_NOT_SUPPORTED = 1703&
    Const RPC_S_INVALID_RPC_PROTSEQ = 1704&

    Dim rcps As Long
    Static fUnexpectedErr As Boolean

    On Error Resume Next

    fIsProtocolSeqSupported = False
    rcps = RpcNetworkIsProtseqValid(strProto)

    Select Case rcps
        Case RPC_S_OK
            fIsProtocolSeqSupported = True
        Case RPC_S_PROTSEQ_NOT_SUPPORTED
            LogNote ResolveResString(resNOTEPROTOSEQNOTSUPPORTED, "|1", strProto, "|2", strProtoFriendlyName)
        Case RPC_S_INVALID_RPC_PROTSEQ
            LogWarning ResolveResString(resNOTEPROTOSEQINVALID, "|1", strProto, "|2", strProtoFriendlyName)
        Case Else
            If Not fUnexpectedErr Then
                MsgWarning ResolveResString(resPROTOSEQUNEXPECTEDERR), vbOKOnly Or vbInformation, gstrTitle
                If gfNoUserInput Then
                    '
                    ' This is probably redundant since this form should never
                    ' be shown if we are running in silent or SMS mode.
                    '
                    ExitSetup Me, gintRET_FATAL
                End If
                fUnexpectedErr = True
            End If
        'End Case
    End Select
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not cmdOK.Enabled Then
    Cancel = False
End If
End Sub


Private Sub optLocal_Click()
txtServer.Enabled = False
txtServer.BackColor = &H80000000
End Sub


Private Sub optRemote_Click()
txtServer.Enabled = True
txtServer.ForeColor = vbWhite
End Sub


Private Sub txtServer_Change()
If Trim$(txtServer.Text) <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
    
End Sub


