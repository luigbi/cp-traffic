VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PBEmail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
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
   ScaleHeight     =   5715
   ScaleWidth      =   9510
   Begin VB.ListBox lstSendings 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   885
      TabIndex        =   17
      Top             =   4785
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.ComboBox cbcToName 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   570
      Width           =   3390
   End
   Begin VB.CommandButton cmdClrAttach 
      Caption         =   "Clear Attachments"
      Height          =   330
      Left            =   150
      TabIndex        =   8
      Top             =   4350
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
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
      TabIndex        =   10
      Top             =   4230
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7455
      Top             =   4155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAttachments 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3570
      Width           =   7605
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1185
      TabIndex        =   5
      Top             =   1140
      Width           =   8130
   End
   Begin VB.TextBox txtMsgBox 
      Height          =   1800
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1650
      Width           =   9195
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "Add Attachments"
      Height          =   300
      Left            =   150
      TabIndex        =   7
      Top             =   3930
      Width           =   1500
   End
   Begin VB.CommandButton cmdSend 
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
      TabIndex        =   9
      Top             =   4230
      Width           =   2010
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   315
      Left            =   5910
      TabIndex        =   2
      Top             =   150
      Width           =   3375
   End
   Begin VB.TextBox txtToEmail 
      Height          =   315
      Left            =   5925
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtFromName 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   135
      Width           =   3390
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8205
      OleObjectBlob   =   "PBEmail.frx":0000
      Top             =   4095
   End
   Begin VB.Label lblResults 
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
      Top             =   4815
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblAttachments 
      Caption         =   "Attachments:"
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
      Left            =   405
      TabIndex        =   16
      Top             =   3570
      Width           =   1185
   End
   Begin VB.Label lblToName 
      Caption         =   "To Name:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   600
      Width           =   900
   End
   Begin VB.Label lblToEmail 
      Caption         =   "To Email:"
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
      Left            =   5025
      TabIndex        =   14
      Top             =   615
      Width           =   870
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
      Left            =   4845
      TabIndex        =   13
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label lblFrom 
      Caption         =   "From Name:"
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
      Left            =   90
      TabIndex        =   12
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label lblSubject 
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
      Left            =   315
      TabIndex        =   11
      Top             =   1140
      Width           =   855
   End
End
Attribute VB_Name = "PBEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
''Private Mail As MailSender
'' PBEMailer removed, zipping global Dan M 9/11/09
''Email sent
''Private objMyEmail As Email
''''Email sender used to send out the emails
''Private WithEvents objEmailSender As EmailSender
''Dim smZipPathName As String
''Dim bmZipFileCreated As Boolean
'Dim tmEmpEmails(7) As EMPLOYEEEMAILS
'Dim bmAddFromAddress As Boolean
'Dim hmCef As Integer
'' dan M 11/04/09 this is a safety to make sure table opened ok: probably don't need?
'Dim bmCefOpened As Boolean
'Private Sub Form_Unload(Cancel As Integer)
'    'no email address for user?
'    If bmAddFromAddress And LenB(Trim(txtFromEmail)) > 0 And bmCefOpened Then
'        gInsertNewFromAddress hmCef, txtFromEmail
'    End If
'    Erase tmEmpEmails
'    If bmCefOpened Then
'        btrClose hmCef
'        btrDestroy hmCef
'    End If
'    Unload PBEmail
'
'End Sub
''Private Sub mInsertNewFromAddress()
''Dan M 11/04/09 made global
'''Dan M no cef for email? update
''Dim olMail As ASPEMAILLib.MailSender
''Dim slAddress As String
''Dim ilRet As Integer
''Dim ilRecLen As Integer
''Dim tlCef As CEF
''Dim hlUrf As Integer
''Dim tlUrf As URF
''Dim tlUrfSearchKey As INTKEY0    'URF key record image
''    slAddress = txtFromEmail
''    Set olMail = New ASPEMAILLib.MailSender
''    If olMail.ValidateAddress(slAddress) = 0 Then
''        tlCef.lCode = 0
''        tlCef.sComment = slAddress
''        ilRecLen = Len(tlCef)
''        ilRet = btrInsert(hmCef, tlCef, ilRecLen, 0)
''        If ilRet = BTRV_ERR_NONE Then   'update urf with cef key
''            hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
''            ilRet = btrOpen(hlUrf, "", sgDBPath & "urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
''            If ilRet = BTRV_ERR_NONE Then
''                ilRecLen = Len(tlUrf)
''                tlUrfSearchKey.iCode = tgUrf(0).iCode
''                ilRet = btrGetEqual(hlUrf, tlUrf, ilRecLen, tlUrfSearchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
''                If ilRet = BTRV_ERR_NONE Then
''                    gUrfDecrypt tlUrf
''                    tlUrf.lEMailCefCode = tlCef.lCode
''                    gUrfEncrypt tlUrf
''                    ilRet = btrUpdate(hlUrf, tlUrf, ilRecLen)
''                    If ilRet = BTRV_ERR_NONE Then
''                       tgUrf(0).lEMailCefCode = tlCef.lCode
''                    End If  'updated urf
''                End If ' found urf
''                ilRet = btrClose(hlUrf)
''                btrDestroy hlUrf
''            End If ' opened urf
''        End If  'updated cef
''    End If  'valid address
''    Set olMail = Nothing
''End Sub
'Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)
'    lstSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)
'
'    lstSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
'
'End Sub
'
'Sub mShowResults(flag As Boolean)
'
'    lstSendings.Visible = flag
'    lblResults.Visible = flag
'
'End Sub
'Sub mLoadCbcToName()
''Dan updated no longer dynamic
'   ' ReDim tmEmpEmails(0 To 7) As EMPLOYEEEMAILS
'    Dim llCount As Long
'
'    tmEmpEmails(0).Name = "Nora Browne"
'    tmEmpEmails(1).Name = "Anna Dow"
'    tmEmpEmails(2).Name = "Eric Garst"
'    tmEmpEmails(3).Name = "Mary Nelson"
'    tmEmpEmails(4).Name = "Martina Patterson"
'    tmEmpEmails(5).Name = "Melinda Lachance Pomeroy"
'   ' tmEmpEmails(5).Name = "Ruth Pender"
'    tmEmpEmails(6).Name = "Sara Vidar"
'    tmEmpEmails(7).Name = "Service Department"
'    tmEmpEmails(0).Email = "norabrowne@counterpoint.net"
'    tmEmpEmails(1).Email = "annadow@counterpoint.net"
'    tmEmpEmails(2).Email = "ericgarst@counterpoint.net"
'    tmEmpEmails(3).Email = "marynelson@counterpoint.net"
'    tmEmpEmails(4).Email = "martinapatterson@counterpoint.net"
'    tmEmpEmails(5).Email = "melindalachance@counterpoint.net"
'    tmEmpEmails(6).Email = "saravidar@counterpoint.net"
'    tmEmpEmails(7).Email = "service@counterpoint.net"
'
'    For llCount = 0 To 7
'      cbcToName.AddItem tmEmpEmails(llCount).Name
'    Next
'
'End Sub
'Private Sub cbcToName_Click()
'    txtToEmail.Text = tmEmpEmails(cbcToName.ListIndex).Email
'End Sub
'
'Private Sub cmdAttach_Click()
'' Dan M can't use object's attachment to add to text box
'    Dim ilLoop As Integer
'   ' Dim slCurDir As String
'
'   ' slCurDir = CurDir
'    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
'    CommonDialog1.ShowOpen 'Show the Open Dialog
'    If CommonDialog1.fileName <> "" Then
'        If LenB(Trim(txtAttachments)) > 0 Then
'            txtAttachments.Text = txtAttachments.Text & ";" & CommonDialog1.fileName
'        Else
'            txtAttachments.Text = CommonDialog1.fileName
'        End If
'    End If
''    If CommonDialog1.fileName <> "" Then
''        If objMyEmail.Attachments.Add(CommonDialog1.fileName) = False Then
''                MsgBox "Unable to Add Attachment! ", vbCritical
''        End If
''    End If
'
''    txtAttachments.Text = ""    'Update Attachment Text
'
''    For ilLoop = 1 To objMyEmail.Attachments.Count
''        txtAttachments.Text = txtAttachments.Text & objMyEmail.Attachments.ItemName(ilLoop) & "; "
''    Next ilLoop
'    'ChDir slCurDir 'Resign the Current Dir
'    gChDrDir
'End Sub
'
'Private Sub cmdClrAttach_Click()
''Dan change method
''    Dim ilLoop As Integer
''    'use reset instead
''   ' objMyEmail.Attachments.Clear    'Clear the object
'    txtAttachments.Text = ""    'Clear the Attachment Text
''    For ilLoop = 1 To objMyEmail.Attachments.Count
''     '   txtAttachments.Text = txtAttachments.Text & objMyEmail.Attachments.ItemName(ilLoop) & "; "
''    Next ilLoop
'
'End Sub
'
'Private Sub cmdExit_Click()
'
'    Unload PBEmail
'
'End Sub
''' Dan use global gSendEmail
''Sub mSendEmail()
''    Set objMyEmail = New Email ' Reset Object Var
''    objMyEmail.Attachments.Add (smZipPathName)
''    objEmailSender.UseSsl = True 'Use SSL/TLS, but allow sending if not supported.
''    objEmailSender.AllowNonSsl = True
''    objMyEmail.FromEmail = txtFromEmail.Text    'Asing the From Information
''    objMyEmail.FromName = txtFromName.Text
''
''    If txtMsgBox.Text = "" Then    'Recipient Information
''        txtMsgBox.Text = " ** No Message **"
''    End If
''    objMyEmail.ToEmail = txtToEmail.Text
''    objMyEmail.ToName = cbcToName.Text
''    objMyEmail.Priority = Priorities.Normal    'Priority
''    objMyEmail.Subject = txtSubject.Text    'Message and Subject
''    objMyEmail.BodyText = txtMsgBox.Text
''
''    Dim EmailTag As String    'Find an Empty space in the List
''    EmailTag = 0
''    Do While Len(lstSendings.List(Int(EmailTag))) > 0 And (Not Mid(lstSendings.List(Int(EmailTag)), 1, 4) = "Done")
''        EmailTag = Int(EmailTag) + 1
''    Loop
''    objMyEmail.Tag = EmailTag
''
''    Call objEmailSender.SendEmail(objMyEmail)    'Send the Email thru the EmailSender with the ListIndex as Tag, to show information on that place of the List.
''
''End Sub
'Private Sub cmdSend_Click()
'Dim tlEmailInfo As EmailInformation 'Dan M. to send to global procedure
'    Screen.MousePointer = vbHourglass
'    lstSendings.Clear
'    lstSendings.ForeColor = vbBlack
'    'Dan M 9/16/09 gSendEMail has stronger validation
'    If txtFromEmail.Text = "" Then ' Or InStr(1, txtFromEmail.Text, "@") = 0 Then
'        MsgBox "You must enter Your Email Address! ", vbCritical
'        txtFromEmail.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If txtToEmail.Text = "" Then ' Or InStr(1, txtToEmail.Text, "@") = 0 Then
'        MsgBox "You must enter the Recipient's Email Address! ", vbCritical
'        txtToEmail.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    mShowResults True
'    With tlEmailInfo
'        .sFromAddress = txtFromEmail
'        .sFromName = txtFromName
'        .sToAddress = txtToEmail
'        .sToName = cbcToName.Text
'        .sSubject = txtSubject
'        .sMessage = txtMsgBox
'        .sAttachment = txtAttachments
'        '.bUserFromHasPriority = True
'    End With
'    gSendEmail tlEmailInfo, lstSendings, zpcDZip
'   ' mSendEmail
''    If Len(txtAttachments.Text) > 0 Then
''        mShowResults True
''      '  mZipAllFiles ' - Zip up the files
''        ' Dan  for deletion later
''      '  bmZipFileCreated = True
''
''        mSendEmail
''    Else
''        mShowResults True
''        bmZipFileCreated = False
''        mSendEmail
''    End If
'
'    cmdExit.Caption = "Done"
'    ' Dan M 9/15/09 don't clear attachment
''    cmdClrAttach_Click
'    ' dan 9/14/09 call to cef instead--don't need file anymore
'    'mWriteEmailFile
'    Screen.MousePointer = vbDefault
'
'End Sub
'
'
'Private Sub Form_Load()
'
''    Set objEmailSender = New EmailSender    'Initialize Sender and Email
''    Set objMyEmail = New Email
''
''    Dim objReg As KeyCodeRegistration
''    Set objReg = New KeyCodeRegistration
''    Call objReg.EnterKeyCode("QwBvAHUAbgB0AGUAcgBwAG8AaQBuAHQAIABTAG8AZgB0AHcAYQByAGUAAABEAEIARQBBADAANQAxADUAOQA3AEQAQQA3ADUAQQAzADMANwA5ADQAQwBFADgARgBFADYANwA4ADAAMAAwAEIANgBEAEYARgBFAEEAMQAxADgAOQAzAEUAMAAyADEARABCAEEAQwA0ADkARgBCADMAQgBEADEAMAA3ADUANABCAEEANwBBADAANQBEAEIARQAzADQAMQA2ADUAMgA5AEIAQQA5ADkARAA0ADYAMgBEADUARABBADIAMgBCADMARQAwADAAQwA0ADAARAAxAEUAOABCAEMAOQBBAEIAQwBBAA==")
'
'    gCenterStdAlone PBEmail
'    mLoadCbcToName
'    mInit
'
'End Sub
'
'
'Private Sub mInit()
'Dim ilRet As Integer
'    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet = BTRV_ERR_NONE Then
'        bmCefOpened = True
'    End If
'    'read from file.... show attachments
'    If Trim$(tgUrf(0).sRept) <> "" Then
'        txtFromName.Text = Trim$(tgUrf(0).sRept) ' & " - " & Trim$(tgSpf.sGClient)
'    Else
'        txtFromName.Text = Trim$(tgUrf(0).sName) ' & " - " & Trim$(tgSpf.sGClient)
'    End If
'    'Dan M 9/14/09 replaced read/write to file with using user's email.
'    txtFromEmail.Text = mGetEMail(tgUrf(0).lEMailCefCode)
'    'mReadEmailFile
'    If sgFileAttachment <> "" Then
'    ' Dan 9/14/09 need path and name
'        'txtAttachments.Text = sgFileAttachmentName
'        txtAttachments.Text = sgFileAttachment
'        txtSubject.Text = Trim$(tgSpf.sGClient) & " - " & Trim$(sgFileAttachmentName)
'    Else
'        txtSubject.Text = Trim$(tgSpf.sGClient)
'    End If
'End Sub
'
''Dan M replace read and write email file with user's defined email address 9/14/09
'
''Private Sub mWriteEmailFile()
'''save from email for this user as a text file
''    Dim tlTxtStream As TextStream
''    Dim fs As New FileSystemObject
''    Dim slLocation As String
''    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(tgUrf(0).iCode) & ".txt"
''    If fs.FileExists(slLocation) Then
''        fs.DeleteFile slLocation
''        DoEvents
''    End If
''
''    fs.CreateTextFile slLocation, False
''    DoEvents
''    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
''    DoEvents
''    tlTxtStream.WriteLine (txtFromEmail.Text)
''    DoEvents
''    tlTxtStream.Close
''    DoEvents
''
''End Sub
''
''Private Sub mReadEmailFile()
''    Dim tlTxtStream As TextStream
''    Dim fs As New FileSystemObject
''    Dim slLocation As String
''    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(tgUrf(0).iCode) & ".txt"
''    If fs.FileExists(slLocation) Then
''        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
''        'right now there is only one line, but we are ready for more
''        Do While tlTxtStream.AtEndOfStream <> True
''            txtFromEmail.Text = tlTxtStream.ReadLine
''        Loop
''        tlTxtStream.Close
''    Else
''        txtFromEmail.Text = ""
''    End If
''End Sub
'Private Function mGetEMail(llEMailCefCode As Long) As String
'    Dim ilRet As Integer
'    Dim tlCefSrchKey0 As LONGKEY0
'    Dim tlCef As CEF
'    Dim ilRecLen As Integer
'   ' Dim hlCef As Integer
'    mGetEMail = ""
'    If llEMailCefCode > 0 And bmCefOpened Then
'       ' hlCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'       ' ilRet = btrOpen(hlCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'       ' If ilRet = BTRV_ERR_NONE Then
'        tlCefSrchKey0.lCode = llEMailCefCode
'        tlCef.sComment = ""
'        ilRecLen = Len(tlCef)    '1009
'        ilRet = btrGetEqual(hmCef, tlCef, ilRecLen, tlCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            mGetEMail = gStripChr0(tlCef.sComment)
'        End If
'    Else    'no email, insert at unload if possible
'        bmAddFromAddress = True
'    End If
'    'ilRet = btrClose(hlCef)
'    'btrDestroy hlCef
'End Function
'
'' Dan M the following events don't have a corresponding aspemail event.
'
''Private Sub objEmailSender_EmailStatusChanged(ByVal EventArgs As PBEmail7.EmailStatusChangedArgs)
''    'dan lose
''    lstSendings.List(Int(EventArgs.Email.Tag)) = EventArgs.Status.Description    'Show the Status in the Sendings List in the Email Tag position
''
''End Sub
''
''
''
''Private Sub objEmailSender_ProgressChanged(ByVal EventArgs As PBEmail7.ProgressChangedArgs)
''    'Dan lose
''    lstSendings.List(Int(EventArgs.Email.Tag)) = "Progress: " & EventArgs.Progress.Progress & "% - " & EventArgs.Progress.BytesSent & " bytes sent."    'Show the Progress in the Sendings List in the Email Tag position
''
''End Sub
''
''
''Private Sub objEmailSender_SendCompleted(ByVal EventArgs As PBEmail7.SendCompletedArgs)
'''Dan clean zip path and send messages: whole thing must now be based on "ok" message
''Dim slMsg As String
''
''    If bmZipFileCreated Then
''        Kill smZipPathName 'Delete zip file
''    End If
''
''    If EventArgs.Result.Result = Results.AllOk Then 'Show the Result in the Sendings List in the Email Tag position
''        lstSendings.List(Int(EventArgs.Email.Tag)) = "Done: Message sent ok."
''    Else
''        slMsg = "A general error has occured in the Email form: " & EventArgs.Result.Description
''        gLogMsg slMsg, "TrafficErrors.Txt", False
''        gMsgBox slMsg, vbCritical, "Email Error"
''    End If
''
''End Sub
