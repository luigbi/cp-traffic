VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
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
      Top             =   4815
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.ComboBox cbcToName 
      Height          =   315
      ItemData        =   "AffEMail.frx":0000
      Left            =   1170
      List            =   "AffEMail.frx":0002
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
      Left            =   7650
      Top             =   4005
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
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8415
      OleObjectBlob   =   "AffEMail.frx":0004
      Top             =   4020
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
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dan M 9/16/09 lost pbemail object and moved to aspemail
'Email sent
'Private objMyEmail As Email
'Private WithEvents objEmailSender As EmailSender
'Dim smZipPathName As String
'Dim bmZipFileCreated As Boolean
Dim tmEmpEmails(0 To 7) As EMPLOYEEEMAILS
Dim bmAddFromAddress As Boolean

Private Sub Form_Unload(Cancel As Integer)
    If bmAddFromAddress And LenB(Trim(txtFromEmail)) > 0 Then
        'mInsertNewFromAddress
        gInsertNewFromAddress (Trim(txtFromEmail))
    End If
    Erase tmEmpEmails
    Set frmEmail = Nothing
    
End Sub
'Private Sub mInsertNewFromAddress()
''Dan M no cef for email? update
'Dim olMail As ASPEMAILLib.MailSender
'Dim slAddress As String
'Dim llNewCefCode As Long
'Dim slQuery As String
'    slAddress = txtFromEmail
'    Set olMail = New ASPEMAILLib.MailSender
'    If olMail.ValidateAddress(slAddress) = 0 Then
'        slQuery = "INSERT into cef_comments_events (cefCode, cefComment) values( Replace, '" & slAddress & "')"
'        llNewCefCode = gInsertAndReturnCode(slQuery, "cef_comments_events", "cefCode", "Replace")
'        If llNewCefCode > 0 Then
'            SQLQuery = "UPDATE ust SET ustEMailCefCode = " & llNewCefCode & " WHERE ustname = '" & sgUserName & "'"
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then  'fail to update, remove from cef. no message box
'                gMsg = "A general error has occured in frmEmail-mInsertNewFromAddress: "
'                gLogMsg "Error: " & gMsg & Err.Description & "; Error # " & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
'                SQLQuery = "delete FROM cef_comments_events WHERE cefCode = " & llNewCefCode
'                gSQLWaitNoMsgBox SQLQuery, False
'            End If  'failed to update Ust
'        End If  'inserted to CEF
'    End If  'valid address
'    Set olMail = Nothing
'
'End Sub

Private Sub zpcDZip_ZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)

    lstSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
    
End Sub
Private Sub zpcDZip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)

    lstSendings.List(0) = "Zipping... " & ItemName & "  " & Percent & " % "
    
End Sub



Sub mLoadCbcToName()

'Dan updated
    Dim llCount As Long

    tmEmpEmails(0).Name = "Nora Browne"
    tmEmpEmails(1).Name = "Anna Dow"
    tmEmpEmails(2).Name = "Eric Garst"
    tmEmpEmails(3).Name = "Mary Nelson"
    tmEmpEmails(4).Name = "Martina Patterson"
    tmEmpEmails(5).Name = "Melinda Lachance Pomeroy"
   ' tmEmpEmails(5).Name = "Ruth Pender"
    tmEmpEmails(6).Name = "Sara Vidar"
    tmEmpEmails(7).Name = "Service Department"
    tmEmpEmails(0).Email = "norabrowne@counterpoint.net"
    tmEmpEmails(1).Email = "annadow@counterpoint.net"
    tmEmpEmails(2).Email = "ericgarst@counterpoint.net"
    tmEmpEmails(3).Email = "marynelson@counterpoint.net"
    tmEmpEmails(4).Email = "martinapatterson@counterpoint.net"
    tmEmpEmails(5).Email = "melindalachance@counterpoint.net"
    tmEmpEmails(6).Email = "saravidar@counterpoint.net"
    tmEmpEmails(7).Email = "service@counterpoint.net"


    For llCount = 0 To 7
      cbcToName.AddItem tmEmpEmails(llCount).Name
    Next


'    Dim llCount As Long
'
'    tmEmpEmails(0).Name = "Anna Dow"
'    tmEmpEmails(1).Name = "Eric Garst"
'    tmEmpEmails(2).Name = "Mary Nelson"
'    tmEmpEmails(3).Name = "Melinda Lechance"
'    tmEmpEmails(4).Name = "Nora Browne"
'    tmEmpEmails(5).Name = "Sara Vidar"
'    tmEmpEmails(6).Name = "Service Department"
'    tmEmpEmails(0).Email = "annadow@counterpoint.net"
'    tmEmpEmails(1).Email = "ericgarst@counterpoint.net"
'    tmEmpEmails(2).Email = "marynelson@counterpoint.net"
'    tmEmpEmails(3).Email = "melindalechance@counterpoint.net"
'    tmEmpEmails(4).Email = "norabrowne@counterpoint.net"
'    tmEmpEmails(5).Email = "saravidar@counterpoint.net"
'    tmEmpEmails(6).Email = "service@counterpoint.net"
'
'
'    For llCount = 0 To 6
'      cbcToName.AddItem tmEmpEmails(llCount).Name
'    Next

End Sub
Private Sub cbcToName_Click()

    txtToEmail.text = tmEmpEmails(cbcToName.ListIndex).Email

End Sub

Private Sub cmdAttach_Click()
' Dan M can't use object's attachment to add to text box
    Dim ilLoop As Integer
    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
    CommonDialog1.ShowOpen 'Show the Open Dialog
    If CommonDialog1.fileName <> "" Then
        If LenB(Trim(txtAttachments)) > 0 Then
            txtAttachments.text = txtAttachments.text & ";" & CommonDialog1.fileName
        Else
            txtAttachments.text = CommonDialog1.fileName
        End If
    End If
    gChDrDir

'    Dim ilLoop As Integer
'  '  Dim slCurDir As String
'
'   ' slCurDir = CurDir
'    CommonDialog1.Filter = "All Files|*.*"    'Setup the CommonDialog
'    CommonDialog1.ShowOpen 'Show the Open Dialog
'
'    If CommonDialog1.fileName <> "" Then
'        If objMyEmail.Attachments.Add(CommonDialog1.fileName) = False Then
'                gMsgBox "Unable to Add Attachment! ", vbCritical
'        End If
'    End If
'
'    txtAttachments.text = ""    'Update Attachment Text
'
'    For ilLoop = 1 To objMyEmail.Attachments.Count
'        txtAttachments.text = txtAttachments.text & objMyEmail.Attachments.ItemName(ilLoop) & "; "
'    Next ilLoop
' '   ChDir slCurDir 'Resign the Current Dir
'    gChDrDir
End Sub

Private Sub cmdClrAttach_Click()
    txtAttachments.text = ""    'Clear the Attachment Text

'    Dim ilLoop As Integer
'
'    objMyEmail.Attachments.Clear    'Clear the object
'    txtAttachments.text = ""    'Clear the Attachment Text
'    For ilLoop = 1 To objMyEmail.Attachments.Count
'        txtAttachments.text = txtAttachments.text & objMyEmail.Attachments.ItemName(ilLoop) & "; "
'    Next ilLoop

End Sub

Private Sub cmdExit_Click()

    Unload frmEmail

End Sub
'' Dan use global gSendEmail

'Sub mSendEmail()
'
'    Set objMyEmail = New Email ' Reset Object Var
'    objMyEmail.Attachments.Add (smZipPathName)
'    'lstSendings.List(0) = ""
'    'Validate Data
'
'
'    objEmailSender.UseSsl = True 'Use SSL/TLS, but allow sending if not supported.
'    objEmailSender.AllowNonSsl = True
'    objMyEmail.FromEmail = txtFromEmail.text    'Asing the From Information
'    objMyEmail.FromName = txtFromName.text
'
'    If txtMsgBox.text = "" Then    'Recipient Information
'        txtMsgBox.text = " ** No Message **"
'    End If
'    objMyEmail.ToEmail = txtToEmail.text
'    objMyEmail.ToName = cbcToName.text
'    objMyEmail.Priority = Priorities.Normal    'Priority
'    objMyEmail.Subject = txtSubject.text    'Message and Subject
'    objMyEmail.BodyText = txtMsgBox.text
'
'    Dim EmailTag As String    'Find an Empty space in the List
'    EmailTag = 0
'    Do While Len(lstSendings.List(Int(EmailTag))) > 0 And (Not Mid(lstSendings.List(Int(EmailTag)), 1, 4) = "Done")
'        EmailTag = Int(EmailTag) + 1
'    Loop
'    objMyEmail.Tag = EmailTag
'
'    Call objEmailSender.SendEmail(objMyEmail)    'Send the Email thru the EmailSender with the ListIndex as Tag, to show information on that place of the List.
'
'End Sub
Private Sub cmdSend_Click()
Dim tlEmailInfo As EmailInformation 'Dan M. to send to global procedure
    Screen.MousePointer = vbHourglass
    lstSendings.Clear
    lstSendings.ForeColor = vbBlack
        'Dan M 9/16/09 gSendEMail has stronger validation
    If txtFromEmail.text = "" Then ' Or InStr(1, txtFromEmail.Text, "@") = 0 Then
        MsgBox "You must enter Your Email Address! ", vbCritical
        txtFromEmail.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If txtToEmail.text = "" Then ' Or InStr(1, txtToEmail.Text, "@") = 0 Then
        MsgBox "You must enter the Recipient's Email Address! ", vbCritical
        txtToEmail.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    mShowResults True
    With tlEmailInfo
        .sFromAddress = txtFromEmail
        .sFromName = txtFromName
        .sToAddress = txtToEmail
        .sToName = cbcToName.text
        .sSubject = txtSubject
        .sMessage = txtMsgBox
        .sAttachment = txtAttachments
      '  .bUserFromHasPriority = True
    End With
    gSendEmail tlEmailInfo, lstSendings, zpcDZip
    cmdExit.Caption = "Done"
    Screen.MousePointer = vbDefault


'    Screen.MousePointer = vbHourglass
'    If txtFromEmail.text = "" Or InStr(1, txtFromEmail.text, "@") = 0 Then
'        gMsgBox "You must enter Your Email Address! ", vbCritical
'        txtFromEmail.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If txtToEmail.text = "" Or InStr(1, txtToEmail.text, "@") = 0 Then
'        gMsgBox "You must enter the Recipient's Email Address! ", vbCritical
'        txtToEmail.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If Len(txtAttachments.text) > 0 Then
'        mShowResults True
'        mZipAllFiles ' - Zip up the files
'        bmZipFileCreated = True
'        mSendEmail
'    Else
'        mShowResults True
'        bmZipFileCreated = False
'        mSendEmail
'    End If
'
'    cmdExit.Caption = "Done"
'    cmdClrAttach_Click
'    mWriteEmailFile
'    Screen.MousePointer = vbDefault

End Sub
Sub mShowResults(flag As Boolean)

    lstSendings.Visible = flag
    lblResults.Visible = flag
    
End Sub

Private Sub Form_Load()
' Dan M 9/16/09 object gone.
'    Set objEmailSender = New EmailSender    'Initialize Sender and Email
'    Set objMyEmail = New Email
'
'    Dim objReg As KeyCodeRegistration
'    Set objReg = New KeyCodeRegistration
'    Call objReg.EnterKeyCode("QwBvAHUAbgB0AGUAcgBwAG8AaQBuAHQAIABTAG8AZgB0AHcAYQByAGUAAABEAEIARQBBADAANQAxADUAOQA3AEQAQQA3ADUAQQAzADMANwA5ADQAQwBFADgARgBFADYANwA4ADAAMAAwAEIANgBEAEYARgBFAEEAMQAxADgAOQAzAEUAMAAyADEARABCAEEAQwA0ADkARgBCADMAQgBEADEAMAA3ADUANABCAEEANwBBADAANQBEAEIARQAzADQAMQA2ADUAMgA5AEIAQQA5ADkARAA0ADYAMgBEADUARABBADIAMgBCADMARQAwADAAQwA0ADAARAAxAEUAOABCAEMAOQBBAEIAQwBBAA==")

    gCenterForm frmEmail
    mLoadCbcToName
    mInit
    
End Sub


Private Sub mInit()
    txtFromName.text = Trim$(sgUserName) & " - " & Trim$(sgClientName)
    'Dan M 9/16/09 find if have email; no longer read/write to file.
    txtFromEmail.text = gGetEMailAddress(bmAddFromAddress)
    
    If sgFileAttachment <> "" Then
        'Dan M not as nice looking, but easier
        'txtAttachments.text = sgFileAttachmentName
        txtAttachments.text = sgFileAttachment
        txtSubject.text = Trim$(sgClientName) & " - " & Trim$(sgFileAttachmentName)
'Dan M no longer exists
'        If objMyEmail.Attachments.Add(sgFileAttachment) = False Then
'            gMsgBox "Unable to Add Attachment! ", vbCritical
'        End If
    Else
        txtSubject.text = Trim$(sgClientName)
    End If

'    txtFromName.text = Trim$(sgUserName) & " - " & Trim$(sgClientName)
'    mReadEmailFile
'    If sgFileAttachment <> "" Then
'        txtAttachments.text = sgFileAttachmentName
'        txtSubject.text = Trim$(sgClientName) & " - " & Trim$(sgFileAttachmentName)
'
'        If objMyEmail.Attachments.Add(sgFileAttachment) = False Then
'            gMsgBox "Unable to Add Attachment! ", vbCritical
'        End If
'    Else
'        txtSubject.text = Trim$(sgClientName)
'    End If

End Sub
'Dan M 11/06/09 made global
'Private Function mGetEMailAddress() As String
'' search ust for sgClientName; get ustmailcefcode.  Look at Cef and see if address;
'Dim rstUst As ADODB.Recordset
'Dim rstCef As ADODB.Recordset
'    SQLQuery = "SELECT ustemailcefcode FROM ust Where ustname = '" & sgUserName & "'"
'    Set rstUst = cnn.Execute(SQLQuery)
'    bmAddFromAddress = True
'    If Not rstUst.EOF Then  'just in case no ust record
'        If rstUst!ustEMailCefCode > 0 Then
'            SQLQuery = "SELECT cefComment FROM cef_comments_events where cefCode = " & rstUst!ustEMailCefCode
'            Set rstCef = cnn.Execute(SQLQuery)
'            If Not rstCef.EOF Then
'                mGetEMailAddress = Trim(rstCef!cefComment)
'                If LenB(mGetEMailAddress) > 0 Then
'                    bmAddFromAddress = False
'                End If
'            End If
'        End If
'    End If
'    Set rstUst = Nothing
'    Set rstCef = Nothing
'End Function

'Dan 9/16/09 no longer write/read to file: use function above instead
'Private Sub mWriteEmailFile()
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slLocation As String
'    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(igUstCode) & ".txt"
'    If fs.FileExists(slLocation) Then
'        fs.DeleteFile slLocation
'        DoEvents
'    End If
'
'    fs.CreateTextFile slLocation, False
'    DoEvents
'    Set tlTxtStream = fs.OpenTextFile(slLocation, ForWriting, False)
'    DoEvents
'    tlTxtStream.WriteLine (txtFromEmail.text)
'    DoEvents
'    tlTxtStream.Close
'    DoEvents
'
'End Sub
'
'Private Sub mReadEmailFile()
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slLocation As String
'    slLocation = sgDBPath & "Messages\CSIEmail" & CStr(igUstCode) & ".txt"
'    If fs.FileExists(slLocation) Then
'        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
'        'right now there is only one line, but we are ready for more
'        Do While tlTxtStream.AtEndOfStream <> True
'            txtFromEmail.text = tlTxtStream.ReadLine
'        Loop
'        tlTxtStream.Close
'    Else
'        txtFromEmail.text = ""
'    End If
'End Sub
'
'Dan M 9/16/09 objects no longer exist
'Private Sub objEmailSender_EmailStatusChanged(ByVal EventArgs As PBEmail7.EmailStatusChangedArgs)
'
'    lstSendings.List(Int(EventArgs.Email.Tag)) = EventArgs.Status.Description    'Show the Status in the Sendings List in the Email Tag position
'
'End Sub
'
'
'
'Private Sub objEmailSender_ProgressChanged(ByVal EventArgs As PBEmail7.ProgressChangedArgs)
'
'    lstSendings.List(Int(EventArgs.Email.Tag)) = "Progress: " & EventArgs.Progress.Progress & "% - " & EventArgs.Progress.BytesSent & " bytes sent."    'Show the Progress in the Sendings List in the Email Tag position
'
'End Sub
'
'
'Private Sub objEmailSender_SendCompleted(ByVal EventArgs As PBEmail7.SendCompletedArgs)
'
'    If bmZipFileCreated Then
'        Kill smZipPathName 'Delete zip file
'    End If
'
'    If EventArgs.Result.Result = Results.AllOk Then 'Show the Result in the Sendings List in the Email Tag position
'        lstSendings.List(Int(EventArgs.Email.Tag)) = "Done: Message sent ok."
'    Else
'        gMsg = "A general error has occured in the Email form: " & EventArgs.Result.Description
'        gLogMsg gMsg, "AffErrorLog.txt", False
'        gMsgBox gMsg, vbCritical
'    End If
'
'End Sub
'Dan M 9/16/09 unused zip procedures : now global
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Sub minitZIPCmdStruct()
'  zpcDZip.ActionDZ = 0 'NO_ACTION
'  zpcDZip.AddCommentFlag = False
'  zpcDZip.AfterDateFlag = False
'  zpcDZip.BackgroundProcessFlag = False
'  zpcDZip.Comment = ""
'  zpcDZip.CompressionFactor = 5
'  zpcDZip.ConvertLFtoCRLFFlag = False
'  zpcDZip.Date = ""
'  zpcDZip.DeleteOriginalFlag = False
'  zpcDZip.DiagnosticFlag = False
'  zpcDZip.DontCompressTheseSuffixesFlag = False
'  zpcDZip.DosifyFlag = False
'  zpcDZip.EncryptFlag = False
'  zpcDZip.FixFlag = False
'  zpcDZip.FixHarderFlag = False
'  zpcDZip.GrowExistingFlag = False
'  zpcDZip.IncludeFollowing = ""
'  zpcDZip.IncludeOnlyFollowingFlag = False
'  zpcDZip.IncludeSysandHiddenFlag = False
'  zpcDZip.IncludeVolumeFlag = False
'  zpcDZip.ItemList = ""
'  zpcDZip.MajorStatusFlag = True
'  zpcDZip.MessageCallbackFlag = True
'  zpcDZip.MinorStatusFlag = True
'  zpcDZip.MultiVolumeControl = 0
'  zpcDZip.NoDirectoryEntriesFlag = True
'  zpcDZip.NoDirectoryNamesFlag = True
'
'  zpcDZip.OldAsLatestFlag = False
'  zpcDZip.PathForTempFlag = False
'  zpcDZip.QuietFlag = False
'  zpcDZip.RecurseFlag = False
'  zpcDZip.StoreSuffixes = ""
'  zpcDZip.TempPath = ""
'  zpcDZip.ZIPFile = ""
'
'  'Write out a log file in the windows sub directory
'  zpcDZip.ZipSubOptions = 256
'
'  ' added for rev 3.00
'  zpcDZip.RenameCallbackFlag = False
'  zpcDZip.ExtProgTitle = ""
'  zpcDZip.ZIPString = ""
'
'End Sub
'Function mAddFileToZip(szZip As String, szFile As String) As Integer
'
'    'Init the Zip control structure
'    Call minitZIPCmdStruct
'
'    zpcDZip.ZIPFile = szZip    'The ZIP file name
'    zpcDZip.ItemList = szFile  'The file list to be added
'    zpcDZip.BackgroundProcessFlag = True
'    zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file
'    'Returns the error code.  This code can be translated by the sub mTranslateErrors.
'    'It is not currently being used to log to a file.
'    mAddFileToZip = zpcDZip.ErrorCode
'
'End Function

'Private Function mTranslateDynaErrors(iError As Integer) As String
'
'    Dim slErrMsg As String
'
'    Select Case iError
'        Case 0
'            slErrMsg = "Backup was Successful."
'        Case 1
'            slErrMsg = "Busy, can't re-enter now."
'        Case 2
'            slErrMsg = "Unexpected end of Zip file."
'        Case 3
'            slErrMsg = "Zip file structure invalid."
'        Case 4
'            slErrMsg = "Out of memory."
'        Case 5
'            slErrMsg = "Internal logic error."
'        Case 6
'            slErrMsg = "Entry too big to split."
'        Case 7
'            slErrMsg = "Invalid comment format."
'        Case 8
'            slErrMsg = "Zip file invalid or insufficient memory."
'        Case 9
'            slErrMsg = "Operation interrupted by application."
'        Case 10
'            slErrMsg = "Temporary file failure."
'        Case 11
'            slErrMsg = "Input file read failure."
'        Case 12
'            slErrMsg = "Nothing to do!"
'        Case 13
'            slErrMsg = "Missing or empty Zip file."
'        Case 14
'            slErrMsg = "Output file write failure, possible disk full."
'        Case 15
'            slErrMsg = "Could not create output file."
'        Case 16
'            slErrMsg = "Invalid control parameters."
'        Case 17
'            slErrMsg = "Could not complete operation."
'        Case 18
'            slErrMsg = "File not found or no read permission."
'        Case 19
'            slErrMsg = "Media Error Encountered."
'        Case 20
'            slErrMsg = "Invalid Multi-Volume control parameters."
'        Case 21
'            slErrMsg = "Improper use of Multi-Volume Zip file."
'    End Select
'
'    mTranslateDynaErrors = slErrMsg
'
'End Function


'Sub mZipAllFiles()
'
'    Dim ilRet As Integer
'    Dim ilPos As Integer
'    Dim slData As String
'    Dim slDateTime As String
'    Dim slStr As String
'    Dim ilLoop As Integer
'    Dim slCurDir As String
'
'    slCurDir = CurDir
'
'    smZipPathName = ""
'    DoEvents
'    slDateTime = " " & Format$(gNow(), "ddmmyy")
'    smZipPathName = sgDBPath & Trim$(sgClientName) & slDateTime & ".zip"  'BuildZipName
'
'    'slData = sgFileAttachment & " " & sgDBPath & "Messages\TrafficErrors.Txt"
'    For ilLoop = 1 To objMyEmail.Attachments.Count
'        slData = slData & objMyEmail.Attachments.Item(ilLoop) & " "
'    Next ilLoop
'
'    ilRet = mAddFileToZip(smZipPathName, slData)
'    objMyEmail.Attachments.Clear
'    ChDir slCurDir
'    DoEvents
'
'
'End Sub
