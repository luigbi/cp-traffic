VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form UserLogEMail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox edcFromEmail 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CheckBox ckcMessageType 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   405
      Width           =   1485
   End
   Begin VB.CheckBox ckcMessageType 
      Caption         =   "Screen Alert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3015
      TabIndex        =   2
      Top             =   405
      Width           =   1485
   End
   Begin VB.CheckBox ckcBlockSystem 
      Caption         =   "Affiliate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2940
      TabIndex        =   6
      Top             =   795
      Width           =   1050
   End
   Begin VB.CheckBox ckcBlockSystem 
      Caption         =   "Traffic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1845
      TabIndex        =   5
      Top             =   795
      Width           =   1050
   End
   Begin VB.ListBox lbcResults 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      ItemData        =   "UserLogEMail.frx":0000
      Left            =   885
      List            =   "UserLogEMail.frx":0002
      TabIndex        =   19
      Top             =   5325
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.CommandButton cmcExit 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5175
      TabIndex        =   17
      Top             =   4770
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7665
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox edcAttachment 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6510
   End
   Begin VB.TextBox edcSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   10
      Top             =   1995
      Width           =   8130
   End
   Begin VB.TextBox edcMessage 
      Height          =   1800
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2415
      Width           =   8130
   End
   Begin VB.CommandButton cmcAttachment 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8130
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmcSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2835
      TabIndex        =   16
      Top             =   4770
      Width           =   2010
   End
   Begin VB.TextBox edcCountdown 
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   1185
      Width           =   645
   End
   Begin VB.Label lblFromEmail 
      Caption         =   "From Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   21
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label lacMessageType 
      Caption         =   "Message Type to be Sent as:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   1
      Top             =   405
      Width           =   2625
   End
   Begin VB.Label lacBlockSystem 
      Caption         =   "Block Users on:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   4
      Top             =   780
      Width           =   1470
   End
   Begin VB.Label lacMessage 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   11
      Top             =   2385
      Width           =   855
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8520
      OleObjectBlob   =   "UserLogEMail.frx":0004
      Top             =   4680
   End
   Begin VB.Label lacResults 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   18
      Top             =   5355
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lacAttachment 
      Caption         =   "Attachment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   13
      Top             =   4335
      Width           =   1185
   End
   Begin VB.Label lacCountdown 
      Caption         =   "Countdown in Minutes to Shutdown:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   7
      Top             =   1230
      Width           =   3300
   End
   Begin VB.Label lacScreen 
      Caption         =   "User Log Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label lacSubject 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   9
      Top             =   1995
      Width           =   855
   End
End
Attribute VB_Name = "UserLogEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of UserLogEMail.Frm on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmUlfSrchKey1                                                                         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mTranslateDynaErrors                                                                  *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  minitZIPCmdStruct             mAddFileToZip                                           *
'******************************************************************************************

Option Explicit
Option Compare Text

'Email sent
Private smEMail As String
Private smUserName As String
Private imTerminate As Integer
'Private Mail As MailSender
'E-mail
Dim tmCef As CEF            'CEF record image
Dim hmCef As Integer        'CEF Handle
Dim imCefRecLen As Integer      'CEF record length
Dim tmCefSrchKey0 As LONGKEY0    'CEF key record image

Dim hmUlf As Integer
Dim tmUlf As ULF        'GSF record image
Dim tmUlfSrchKey0 As LONGKEY0    'GSF key record image
Dim imUlfRecLen As Integer        'GSF record length

Dim hmUrf As Integer
Dim tmUrf As URF
Dim tmUrfSrchKey0 As INTKEY0    'URF key record image
Dim imUrfRecLen As Integer        'URF record length

Dim hmUst As Integer
Dim tmUst As UST
Dim tmUstSrchKey0 As INTKEY0    'UST key record image
Dim imUstRecLen As Integer        'UST record length
'Private imEMailAllowed As Integer   'Dan M 9/11/09 moved where email object is set..this is no longer doing anything
'Private smZipPathName As String     'commented out because not zipping
'Private bmZipFileCreated As Boolean 'Dan M set but never used..also not zipping
'Private smToAddress As String         'Dan M set but never used
'Dan M 11/04/09 update Cef email address if currently empty
Dim bmAddFromAddress As Boolean
'Dan M 11/06/09 set by userlog form to determine whether to enable email
Public bmEnableEmail As Boolean
Private Sub ckcMessageType_Click(Index As Integer)
    mSetCommands
End Sub

Private Sub edcCountdown_Change()
    mSignOffMsg
End Sub

Private Sub edcCountdown_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bmAddFromAddress And LenB(Trim(edcFromEmail)) > 0 Then
        'gInsertNewFromAddress hmCef, edcFromEmail
        If Not ogEmailer Is Nothing Then
            ogEmailer.SetEmailAddressThisUser (Trim$(edcFromEmail.Text))
        End If
    End If
    btrDestroy hmUst
    btrDestroy hmUrf
    btrDestroy hmUlf
    btrDestroy hmCef
    Set ogEmailer = Nothing
    Unload UserLogEMail

End Sub

Sub mShowResults(flag As Boolean)

    lbcResults.Visible = flag
    lacResults.Visible = flag

End Sub
Private Sub cmcAttachment_Click()

    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
    CommonDialog1.ShowOpen 'Show the Open Dialog

    If CommonDialog1.fileName <> "" Then
        edcAttachment.Text = Trim$(CommonDialog1.fileName)
    Else
        edcAttachment.Text = ""
    End If

    gChDrDir
End Sub


Private Sub cmcExit_Click()

    Unload UserLogEMail

End Sub

'Dan M 9/11/09 made into gSendEmail
'Private Function mSendEmail() As Integer
'Dim slErrorFact As String
'' Dan M 9/09/09 flag never used
''    Dim iOkToSend As Integer
''
''    iOkToSend = True
'    mSendEmail = True
'' Dan M 9/09/09 commented out stopping of email 1)don't have to have username/password 2)let error be shown in lbcresults
'    If Trim$(tgSite.sEmailHost) = "" Then
'        Mail.Host = ""
''        iOkToSend = False
'        slErrorFact = "SMTP Host is undefined in Site Options"
''        gMsgBox "SMTP Host is undefined in Site Options", vbCritical, "Email Error"
''        mSendEmail = False
''        Exit Function
'    Else
'        Mail.Host = Trim$(tgSite.sEmailHost)
'    End If
'
'    If tgSite.iEmailPort = 0 Then
'        Mail.Port = 0
'        slErrorFact = "Port Number is undefined in Site Options"
''        iOkToSend = False
''        gMsgBox "Port Number is undefined in Site Options", vbCritical, "Email Error"
''        mSendEmail = False
''        Exit Function
'    Else
'        Mail.Port = Trim$(tgSite.iEmailPort)
'    End If
''dan test must fix
'    If Trim$(tgSite.sEmailAcctName) = "" Then
'        Mail.Username = ""
''        iOkToSend = False
''        gMsgBox "Account Name is undefined in Site Options", vbCritical, "Email Error"
''        mSendEmail = False
''        Exit Function
'    Else
'        Mail.Username = Trim$(tgSite.sEmailAcctName)
'    End If
'
'    If Trim$(tgSite.sEmailPassword) = "" Then
'        Mail.Password = ""
''        iOkToSend = False
''        gMsgBox "Password is undefined in Site Options", vbCritical, "Email Error"
''        mSendEmail = False
''        Exit Function
'    Else
'        Mail.Password = Trim$(tgSite.sEmailPassword)
'    End If
'
'    'Use the site option if it exists otherwise use the current users name
'    If Trim$(tgSite.sEmailFromName) = "" Then
'        Mail.FromName = Trim$(smUserName)
'    Else
'        Mail.FromName = Trim$(tgSite.sEmailFromName)
'    End If
'
'    'Use the site option if it exists otherwise use the current users email address
'    If Trim$(tgSite.sEmailFromAddress) = "" Then
'        Mail.From = Trim$(smEMail)
'    Else
'        Mail.From = Trim$(tgSite.sEmailFromAddress)
'    End If
'
'
'     'Subject
'     If Trim(edcSubject.Text) <> "" Then
'         Mail.Subject = Trim(edcSubject.Text)
'     Else
'        MsgBox "Please enter a subject."
'        edcSubject.SetFocus
'        Exit Function
'     End If
'
'    'body
'     If Trim(edcMessage.Text) <> "" Then
'         Mail.Body = Trim(edcMessage.Text)
'     Else
'         MsgBox "Please enter the email body."
'         edcMessage.SetFocus
'         Exit Function
'     End If
'
'    'uncomment this if you want attachments zipped.
'    'Also search for the comment: "uncomment this line to zip the attachment" and uncomment as well
'    ' If smZipPathName <> "" Then
'    '     Mail.AddAttachment smZipPathName
'    ' End If
'
'    'attachments will not be zipped using this code.  Comment out if your going to zip them
'    If Trim$(edcAttachment.Text) <> "" Then
'        Mail.AddAttachment edcAttachment.Text
'    End If
'    'Dan M 9/09/09 so error will be read
'   On Error Resume Next
'    Mail.Send ' send message
'    If Err <> 0 Then ' error occurred
'        lbcResults.ForeColor = vbRed
'        lbcResults.AddItem Err.Description & "  " & slErrorFact
'    Else
'        lbcResults.Clear
'        lbcResults.ForeColor = vbGreen
'        lbcResults.AddItem "Mail sent successfully"
'    End If
'
''    Mail.Send ' send message
''    If Err <> 0 Then ' error occurred
''        lbcResults.ForeColor = vbRed
''         lbcResults.Text = Err.Description
''    Else
''        lbcResults.Clear
''        lbcResults.ForeColor = vbGreen
''       lbcResults.Text = "Mail sent successfully"
''    End If
'    lbcResults.Refresh
'End Function

Private Sub cmcSend_Click()

    Dim ilUlf As Integer
    Dim slStr As String
    Dim slSignOffDate As String
    Dim ilRet As Integer
    Dim llEMailCefCode As Long
    Dim slEMail As String
    Dim slUserName As String
    Dim ilSendEmail As Integer
   ' Dim tlEmailInfo As EmailInformation 'Dan M. to send to global procedure
    Dim slTempString As String
    'Dan M 09/06/11 use new email object
    'Dan M made local
   ' Dim Mail As MailSender
    
    'dan m 090909
    lbcResults.Clear
    lbcResults.ForeColor = vbBlack
    DoEvents
    'dan M 9/07/11
    'Set Mail = New ASPEMAILLib.MailSender
    ogEmailer.Clear False, True
    Screen.MousePointer = vbHourglass
    If ckcMessageType(0).Value = vbChecked Then 'alert
        If igUserLogButton = 2 Then
            'Allow zero to be used to turn it off
            If Len(Trim$(edcCountdown.Text)) = 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Countdown must be defined! ", vbCritical
                Exit Sub
            End If
            If Val(edcCountdown.Text) < 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Countdown must a positive number or zero! ", vbCritical
                Exit Sub
            End If
            If (ckcBlockSystem(0).Value = vbUnchecked) And (ckcBlockSystem(1).Value = vbUnchecked) Then
                Screen.MousePointer = vbDefault
                MsgBox "Block User on Traffic and/or Affiliate must be defined! ", vbCritical
                Exit Sub
            End If
        End If
    End If
    If ckcMessageType(1).Value = vbChecked Then
        If Trim$(edcSubject.Text) = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "Subject must be defined! ", vbCritical
            Exit Sub
        End If
    End If
    'Dan M 9/18/09 save cef messages only for alerts..move to if below
'        slStr = Trim$(edcMessage.Text)
'        'tmCef.iStrLen = Len(slStr)
'        tmCef.sComment = Trim$(slStr) & Chr$(0) '& Chr$(0) 'sgTB
'        imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
'        'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
'        If Trim$(slStr) <> "" Then
'            tmCef.lCode = 0 'Autoincrement
'            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
'        Else
'            tmCef.lCode = 0
'            ilRet = BTRV_ERR_NONE
'        End If
'        If ilRet <> BTRV_ERR_NONE Then
'            tmCef.lCode = 0
'        End If
    If ckcMessageType(0).Value = vbChecked Then
        slStr = Trim$(edcMessage.Text)
        'tmCef.iStrLen = Len(slStr)
        tmCef.sComment = Trim$(slStr) & Chr$(0) '& Chr$(0) 'sgTB
        imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
        'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
        If Trim$(slStr) <> "" Then
            tmCef.lCode = 0 'Autoincrement
            ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
        Else
            tmCef.lCode = 0
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tmCef.lCode = 0
        End If
        If igUserLogButton = 2 Then
            'Send block
            slStr = Trim$(str$(tgUrf(0).iCode)) & "|" & Trim$(str$(tmCef.lCode)) & "|" & Trim$(edcCountdown.Text)
            If (ckcBlockSystem(0).Value = vbChecked) And (ckcBlockSystem(1).Value = vbChecked) Then
                ilRet = gAlertAdd("B", "B", 0, 0, slStr)
            ElseIf (ckcBlockSystem(0).Value = vbChecked) Then
                ilRet = gAlertAdd("B", "T", 0, 0, slStr)
            Else
                ilRet = gAlertAdd("B", "A", 0, 0, slStr)
            End If
        Else
            'Send alert notice
            For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
                tmUlfSrchKey0.lCode = lgUserLogUlfCode(ilUlf)
                ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDate tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1), slSignOffDate
                    If gDateValue(slSignOffDate) = gDateValue("12/31/2069") Then
                        slStr = Trim$(str$(tmUlf.lCode)) & "|" & Trim$(str$(tmCef.lCode))
                        ilRet = gAlertAdd("N", "", 0, 0, slStr)
                    End If
                End If
            Next ilUlf
        End If
    End If
    If ckcMessageType(1).Value = vbChecked Then
        ' Dan M 9/11/09 removed zipping ability (not implemented anyway) and moved to global
        mShowResults True
'        If Len(edcAttachment.Text) > 0 Then
'           ' mShowResults True
'            mZipAllFiles ' - Zip up the files
'            'Dan M 9/11/09 variable set but never used.  On another form, was used to delete file when done.
'            bmZipFileCreated = True
'        Else
'           ' mShowResults True
'            bmZipFileCreated = False
'        End If

        cmcExit.Caption = "Done"
        'Dan M 9/14/09 this currently not doing anything, commented out procedure
       ' mWriteEmailFile
        ilSendEmail = False
        For ilUlf = 0 To UBound(lgUserLogUlfCode) - 1 Step 1
            llEMailCefCode = 0
            tmUlfSrchKey0.lCode = lgUserLogUlfCode(ilUlf)
            ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If tmUlf.sSystemType = "T" Then
                    tmUrfSrchKey0.iCode = tmUlf.iUrfCode
                    ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmUrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        llEMailCefCode = tmUrf.lEMailCefCode
                        If Trim$(gDecryptField(Trim$(tmUrf.sRept))) <> "" Then
                            slUserName = Trim$(gDecryptField(Trim$(tmUrf.sRept)))
                        Else
                            slUserName = Trim$(gDecryptField(Trim$(tmUrf.sName)))
                        End If
                        ' Dan M 9/11/09 set but never used
                       ' smToAddress = Trim$(slUserName)
                    End If
                Else
                    tmUstSrchKey0.iCode = tmUlf.iUstCode
                    ilRet = btrGetEqual(hmUst, tmUst, imUstRecLen, tmUstSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        llEMailCefCode = tmUst.lEMailCefCode
                        If Trim$(tmUst.sReportName) <> "" Then
                            slUserName = Trim$(tmUst.sReportName)
                        Else
                            slUserName = Trim$(tmUst.sName)
                        End If
                    End If
                     ' Dan M 9/11/09 set but never used
                   ' smToAddress = Trim$(slUserName)
                End If
            End If
            If llEMailCefCode > 0 Then
                'slEMail = mGetEMail(llEMailCefCode)
                slEMail = ogEmailer.GetEmailAddress(llEMailCefCode)
                If Trim$(slEMail) <> "" Then
                    ogEmailer.AddTOAddress slEMail, slUserName
                   ' Mail.AddAddress slEMail, slUserName
                    'Dan M 11/06/09 count number of emails sending; display to user.
                   ' ilSendEmail = True
                    ilSendEmail = ilSendEmail + 1
                End If
            End If
        Next ilUlf
        'Dan made msendemail global, added else statements 9/09/09
        lbcResults.ForeColor = vbRed
        'Dan M 11/06/09 display # of emails sent
        'If ilSendEmail Then
        If ilSendEmail > 0 Then
            With ogEmailer
                .FromAddress = Trim(edcFromEmail.Text)
                .FromName = smUserName
                .Subject = edcSubject.Text
                .Message = edcMessage.Text
                .Attachment = edcAttachment.Text
            End With
'            tlEmailInfo.sFromAddress = Trim(edcFromEmail.Text)
'            'tlEmailInfo.sFromAddress = smEMail
'            tlEmailInfo.sFromName = smUserName
'            tlEmailInfo.sSubject = edcSubject.Text
'            tlEmailInfo.sMessage = edcMessage.Text
'            tlEmailInfo.sAttachment = edcAttachment.Text
            'gSendEmail tlEmailInfo, lbcResults, , Mail
            'to zip, send zip control
             'gSendEmail tlEmailInfo, lbcResults, zpcDZip, Mail
            If ogEmailer.Send(lbcResults, zpcDZip) Then
            'dan 8/6/11 added '='
                If UBound(lgUserLogUlfCode) > ilSendEmail Then
            'If gSendEmail(tlEmailInfo, lbcResults, , Mail) And UBound(lgUserLogUlfCode) > ilSendEmail Then
                    If ilSendEmail = 1 Then
                        slTempString = " 1 email was sent: "
                    Else
                        slTempString = ilSendEmail & " emails were sent: "
                    End If
                    If UBound(lgUserLogUlfCode) - ilSendEmail = 1 Then
                        lbcResults.AddItem slTempString & " a selected user does not have the capacity to receive email."
                    Else
                        lbcResults.AddItem slTempString & UBound(lgUserLogUlfCode) - ilSendEmail & " selected users do not have the capacity to receive email."
                    End If
                'handled in ogEmailer
'                Else
'                    lbcResults.ForeColor = vbBlack
'                    lbcResults.AddItem "mail was sent."
                End If
            Else
                'handled inside 'send'
                'lbcResults.AddItem "Email was not sent. " & ogEmailer.ErrorMessage
            End If
        'Dan M 11/06/09 now that 'logUser' makes sure selected, don't need this
'        ElseIf UBound(lgUserLogUlfCode) = 0 Then
'            lbcResults.AddItem "Please select users to receive email"
        Else
            lbcResults.AddItem "Email was not sent."
        End If

'        If ilSendEmail Then
'            ilRet = mSendEmail()
'        End If
    End If

    If igUserLogButton = 3 Then
        ilRet = gAlertClear("A", "B", "", 0, 0, "")
    End If
    cmcExit.Caption = "Done"
    'Set Mail = Nothing
    Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Load()

    gCenterStdAlone UserLogEMail
    mInit

End Sub


Private Sub mInit()
    Dim ilRet As Integer
    'Dan M 90909 moved object setting to cmcSend-- pressing 'send' clears object and starts fresh (in case user stays on form and sends more than one email)
    ' This made a lot of the code unnecessary.
'    On Error GoTo CreateObjectErr:
'    imEMailAllowed = True
    'Set Mail = CreateObject("Persits.MailSender")
'    If Not imEMailAllowed Then
'        MsgBox "Unable to Send E-Mail, Error # " & Err.Number & " Description " & Err.Description, vbExclamation + vbOKOnly, "Send Message"
'    End If
'    On Error GoTo 0
    'dan m 9/7/11
    Set ogEmailer = New CEmail
    imTerminate = False
    hmUlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UserLogEMail
    On Error GoTo 0
    imUlfRecLen = Len(tmUlf)  'Get and save ARF record length

    hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UserLogEMail
    On Error GoTo 0
    imUrfRecLen = Len(tmUrf)  'Get and save ARF record length
    hmUst = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUst, "", sgDBPath & "Ust.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", UserLogEMail
    On Error GoTo 0
    imUstRecLen = Len(tmUst)  'Get and save ARF record length
    If Trim$(tgUrf(0).sRept) <> "" Then
        smUserName = Trim$(tgUrf(0).sRept)
    Else
        smUserName = Trim$(tgUrf(0).sName)
    End If
    smEMail = ""
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
       If tgUrf(0).lEMailCefCode > 0 Then
           'tmCefSrchKey0.lCode = tgUrf(0).lEMailCefCode
           'tmCef.sComment = ""
           'imCefRecLen = 1009
           'ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
           'If ilRet = BTRV_ERR_NONE Then
           '    If tmCef.iStrLen > 0 Then
           '        smEMail = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
           '    End If
           'End If
          ' smEMail = mGetEMail(tgUrf(0).lEMailCefCode)
          smEMail = ogEmailer.GetEmailAddress(tgUrf(0).lEMailCefCode)
       Else    'no email, insert at unload if possible
           bmAddFromAddress = True
       End If
    End If
   'Dan M 11/06/09
    If bmEnableEmail Then
        ckcMessageType(1).Enabled = True
        edcFromEmail.Text = smEMail
        mSetSubject
    Else
        ckcMessageType(1).Enabled = False
    End If
    'Dan M 11/06/09 code already here: if email enabled, default value is checked. moved msetCommands after and mSetSubject now set above

    'mSetCommands
    If igUserLogButton = 1 Then 'Message
        If ckcMessageType(1).Enabled Then
            ckcMessageType(1).Value = vbChecked
           ' mSetSubject
        Else
            ckcMessageType(0).Value = vbChecked
        End If
        mMessageMsg
    ElseIf igUserLogButton = 2 Then 'Sign Off
        ckcMessageType(0).Value = vbChecked
        If ckcMessageType(1).Enabled Then
            ckcMessageType(1).Value = vbChecked
            'mSetSubject
        End If
        edcCountdown.Text = "10"
        ckcBlockSystem(0).Value = vbChecked
        ckcBlockSystem(1).Value = vbChecked
        mSignOffMsg
    ElseIf igUserLogButton = 3 Then 'Sign on
        ckcMessageType(0).Enabled = False
        'mSetSubject
        mSignOnMsg
    End If
    mSetCommands
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    'Exit Sub
    'Dan M 9/11/09 object not created here any longer
'CreateObjectErr:
'    imEMailAllowed = False
'    Resume Next
End Sub
'Dan M 9/14/09 commented out.  currently only erasing info in file.
'Private Sub mWriteEmailFile()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilRet                                                                                 *
''******************************************************************************************
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slLocation As String
'
'    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(tgUrf(0).iCode) & ".txt"
'    If fs.FileExists(slLocation) Then
'        fs.DeleteFile slLocation
'        DoEvents
'    End If
'
'    fs.CreateTextFile slLocation, False
'    DoEvents
'    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
'    DoEvents
'    'Fix
'    tlTxtStream.WriteLine ("")
'    DoEvents
'    tlTxtStream.Close
'    DoEvents
'
'End Sub

Private Sub mSetCommands()
    If (tgUrf(0).sAllowedToBlock <> "Y") Or (igUserLogButton <> 2) Then
        edcCountdown.Enabled = False
        lacCountdown.Enabled = False
        ckcBlockSystem(0).Enabled = False
        ckcBlockSystem(1).Enabled = False
        lacBlockSystem.Enabled = False
    End If
    'Dan M 9/11/09 moved setting of mail object...imEMailAllowed no longer used
   ' If (smEMail = "" And Trim$(tgSite.sEmailFromAddress) = "") Or (Not imEMailAllowed) Then
'    If (smEMail = "" And Trim$(tgSite.sEmailFromAddress) = "") Then
    ' tgsite.sEmailFromAddress no longer used Dan M 11/04/09-- always allow email option--but still see if 'checked'
    mEmailOption ckcMessageType(1).Value
'    If (smEMail = "") Then
'
'        ckcMessageType(1).Enabled = False
'        ' Dan M 11/04/09 moved boxes into own method
'        mEmailOption False
''        edcSubject.Enabled = False
''        lacSubject.Enabled = False
''        edcAttachment.Enabled = False
''        lacAttachment.Enabled = False
''        cmcAttachment.Enabled = False
''        edcFromEmail.Enabled = False
'    Else
'        If ckcMessageType(1).Value = vbChecked Then
'            mEmailOption True
''            edcSubject.Enabled = True
''            edcAttachment.Enabled = True
''            cmcAttachment.Enabled = True
''            edcFromEmail.Enabled = True
''            edcFromEmail.Text = smEMail
'        Else
'            mEmailOption False
''            edcSubject.Enabled = False
''            edcAttachment.Enabled = False
''            cmcAttachment.Enabled = False
''            edcFromEmail.Enabled = False
''            edcFromEmail.Text = ""
'        End If
'    End If
    If igUserLogButton = 3 Then 'Sign on
        ckcMessageType(0).Enabled = False
    End If

End Sub
Private Sub mEmailOption(blAllow As Boolean)
    edcSubject.Enabled = blAllow
    lacSubject.Enabled = blAllow
    edcAttachment.Enabled = blAllow
    lacAttachment.Enabled = blAllow
    cmcAttachment.Enabled = blAllow
    'dan M 11/04/09 added from name and from email because site options doesn't contain this info anymore
    edcFromEmail.Enabled = blAllow
    lblFromEmail.Enabled = blAllow
End Sub
Private Sub mSignOffMsg()
    Dim slNamePlusPhone As String
    Dim slZone As String
    Dim ilPos As Integer
    Dim slLen As String
    Dim ilLen As Integer

    If Trim$(tgUrf(0).sRept) <> "" Then
        slNamePlusPhone = Trim$(tgUrf(0).sRept)
    Else
        slNamePlusPhone = Trim$(tgUrf(0).sName)
    End If
    If Trim$(tgUrf(0).sPhoneNo) <> "" Then
        slNamePlusPhone = slNamePlusPhone & " at " & Trim$(tgUrf(0).sPhoneNo)
    End If
    slZone = gGetLocalTZName()
    ilPos = InStr(1, slZone, " ", vbTextCompare)
    If ilPos > 0 Then
        slZone = Left$(slZone, ilPos - 1) & " Time"
    End If
    slLen = Trim$(edcCountdown.Text)
    ilLen = Val(slLen)
    edcMessage.Text = "System will be shutdown on " & Format$(Now, "m/d/yy") & " at " & gFormatTimeLong(gTimeToLong(Now, False) + 60 * ilLen, "A", "2") & " " & slZone & " " & "(" & slLen & " minutes from now)" & sgCR & sgLF & "If you have any problems with this, call " & slNamePlusPhone
End Sub

Private Sub mSignOnMsg()
    Dim slNamePlusPhone As String

    If Trim$(tgUrf(0).sRept) <> "" Then
        slNamePlusPhone = Trim$(tgUrf(0).sRept)
    Else
        slNamePlusPhone = Trim$(tgUrf(0).sName)
    End If
    If Trim$(tgUrf(0).sPhoneNo) <> "" Then
        slNamePlusPhone = slNamePlusPhone & " at " & Trim$(tgUrf(0).sPhoneNo)
    End If
    edcMessage.Text = "System is back up!!" & sgCR & sgLF & "Thank you for your patience" & sgCR & sgLF & "If you have any questions call " & slNamePlusPhone
End Sub

Private Sub mSetSubject()
    If igUserLogButton = 1 Then 'Message
        If Trim$(tgUrf(0).sRept) <> "" Then
            edcSubject.Text = "Message from " & Trim$(tgUrf(0).sRept)
        Else
            edcSubject.Text = "Message from " & Trim$(tgUrf(0).sName)
        End If
        'Dan M 9/10/09 line below wipes out if statment above
        'edcSubject.Text = "Message from "
    ElseIf igUserLogButton = 2 Then 'Sign Off
        edcSubject.Text = "Sign Off Message"
    ElseIf igUserLogButton = 3 Then 'Sign on
        edcSubject.Text = "Sign On Message"
    End If
End Sub

Private Sub mMessageMsg()
    Dim slName As String

    If Trim$(tgUrf(0).sRept) <> "" Then
        slName = Trim$(tgUrf(0).sRept)
    Else
        slName = Trim$(tgUrf(0).sName)
    End If
    edcMessage.Text = "Message from " & slName & " on " & Now
End Sub

'Private Function mGetEMail(llEMailCefCode As Long) As String
'    Dim ilRet As Integer
'
'    mGetEMail = ""
'    If llEMailCefCode > 0 Then
'        tmCefSrchKey0.lCode = llEMailCefCode
'        tmCef.sComment = ""
'        imCefRecLen = Len(tmCef)    '1009
'        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            'If tmCef.iStrLen > 0 Then
'            '    mGetEMail = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
'            'End If
'            mGetEMail = gStripChr0(tmCef.sComment)
'        End If
'    End If
'End Function
''unused zip procedures commented out Dan M 9/11/09
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Sub minitZIPCmdStruct() 'VBC NR
'  zpcDZip.ActionDZ = 0 'NO_ACTION 'VBC NR
'  zpcDZip.AddCommentFlag = False 'VBC NR
'  zpcDZip.AfterDateFlag = False 'VBC NR
'  zpcDZip.BackgroundProcessFlag = False 'VBC NR
'  zpcDZip.COMMENT = "" 'VBC NR
'  zpcDZip.CompressionFactor = 5 'VBC NR
'  zpcDZip.ConvertLFtoCRLFFlag = False 'VBC NR
'  zpcDZip.Date = "" 'VBC NR
'  zpcDZip.DeleteOriginalFlag = False 'VBC NR
'  zpcDZip.DiagnosticFlag = False 'VBC NR
'  zpcDZip.DontCompressTheseSuffixesFlag = False 'VBC NR
'  zpcDZip.DosifyFlag = False 'VBC NR
'  zpcDZip.EncryptFlag = False 'VBC NR
'  zpcDZip.FixFlag = False 'VBC NR
'  zpcDZip.FixHarderFlag = False 'VBC NR
'  zpcDZip.GrowExistingFlag = False 'VBC NR
'  zpcDZip.IncludeFollowing = "" 'VBC NR
'  zpcDZip.IncludeOnlyFollowingFlag = False 'VBC NR
'  zpcDZip.IncludeSysandHiddenFlag = False 'VBC NR
'  zpcDZip.IncludeVolumeFlag = False 'VBC NR
'  zpcDZip.ItemList = "" 'VBC NR
'  zpcDZip.MajorStatusFlag = True 'VBC NR
'  zpcDZip.MessageCallbackFlag = True 'VBC NR
'  zpcDZip.MinorStatusFlag = True 'VBC NR
'  zpcDZip.MultiVolumeControl = 0 'VBC NR
'  zpcDZip.NoDirectoryEntriesFlag = True 'VBC NR
'  zpcDZip.NoDirectoryNamesFlag = True 'VBC NR
'
'  zpcDZip.OldAsLatestFlag = False 'VBC NR
'  zpcDZip.PathForTempFlag = False 'VBC NR
'  zpcDZip.QuietFlag = False 'VBC NR
'  zpcDZip.RecurseFlag = False 'VBC NR
'  zpcDZip.StoreSuffixes = "" 'VBC NR
'  zpcDZip.TempPath = "" 'VBC NR
'  zpcDZip.ZIPFile = "" 'VBC NR
'
'  'Write out a log file in the windows sub directory
'  zpcDZip.ZipSubOptions = 256 'VBC NR
'
'  ' added for rev 3.00
'  zpcDZip.RenameCallbackFlag = False 'VBC NR
'  zpcDZip.ExtProgTitle = "" 'VBC NR
'  zpcDZip.ZIPString = "" 'VBC NR
'
'End Sub 'VBC NR
'Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)
'
'    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)
'
'    lbcResults.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'
'Function mAddFileToZip(szZip As String, szFile As String) As Integer 'VBC NR
'
'    'Init the Zip control structure
'    Call minitZIPCmdStruct 'VBC NR
'
'    zpcDZip.ZIPFile = szZip    'The ZIP file name 'VBC NR
'    zpcDZip.ItemList = szFile  'The file list to be added 'VBC NR
'    zpcDZip.BackgroundProcessFlag = True 'VBC NR
'    zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file 'VBC NR
'    'Returns the error code.  This code can be translated by the sub mTranslateErrors.
'    'It is not currently being used to log to a file.
'    mAddFileToZip = zpcDZip.ErrorCode 'VBC NR
'
'End Function 'VBC NR
'
'Sub mZipAllFiles()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilRet                         ilPos                         slData                    *
''*  slStr                         ilLoop                                                  *
''******************************************************************************************
'
'
'    Dim slDateTime As String
'    Dim slCurDir As String
'
'    slCurDir = CurDir
'
'    smZipPathName = ""
'    DoEvents
'    slDateTime = " " & Format$(gNow(), "mmddyy")
'    smZipPathName = sgDBPath & gFileNameFilter(Trim$(tgSpf.sGClient)) & slDateTime & ".zip"  'BuildZipName
'
'    'uncomment this line to zip the attachment
'    'ilRet = mAddFileToZip(smZipPathName, Trim$(edcAttachment.text))
'    gChDrDir
'    DoEvents
'
'End Sub
' end zip functions

