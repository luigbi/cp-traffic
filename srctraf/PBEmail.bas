Attribute VB_Name = "PBEmailMod"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PBEmail.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
Option Explicit 'added with 10016
Public ogEmailer As CEmail
'10564 (removed)
'Public Const CSISITE As String = "smtpauth.hosting.earthlink.net"  '"smtp.office365.com" '
'Public Const CSIPORT As Integer = 587
'Public Const CSIUSERNAME As String = "csitest@counterpoint.net" '"noreply@counterpoint.net" '
'Public Const CSIPASSWORD As String = "TestMyEmail12!" '"Csi44Sic"  '"csi#8x21" '
'Public Const CSITLS As Boolean = True

'TTP 10837,10564,10564    2023-10-30 JJB
Public Const CSISITE As String = "smtp.office365.com"
Public Const CSIPORT As Integer = 587
Public Const CSIUSERNAME As String = "noreply@counterpoint.net" '
Public Const CSIPASSWORD As String = "csi#8x21"
Public Const CSITLS As Boolean = True


'Public Const CSISITE As String = "10.70.0.15"
'Public Const CSIPORT As Integer = 25
'Public Const CSIUSERNAME As String = "csitest@counterpoint.net"
'Public Const CSIPASSWORD As String = "TestMyEmail12!"
'Public Const CSITLS As Boolean = False


Type EMPLOYEEEMAILS
    Name As String
    Email As String
End Type
'10016 moved to invoice
''Dan M 11/7/16 email for darlene's invoicing 8245 10016 added optional test
'Public Function gPDFEmailStart() As Boolean
'    Dim blRet As Boolean
'    Dim slLogName As String
'    blRet = True
'    Set MyFile = New CLogger
'    slLogName = MyFile.CreateLogName(sgDBPath & "Messages\" & "PDFEmail.Txt")
'    MyFile.LogPath = slLogName
'    MyFile.WriteFacts "Starting sending emails", True
'    '10016
'    If bgPDFEmailTestMode Then
'        MyFile.WriteWarning "In test mode!  Not sending Emails", True
'    End If
'    Set ogEmailer = New CEmail
'    If Len(ogEmailer.ErrorMessage) > 0 Then
'        blRet = False
'        gLogMsg "Email could not be set up in mod PBEmail-gPDFEmailStart: " & ogEmailer.ErrorMessage, "TrafficErrors.Txt", False
'    End If
'    gPDFEmailStart = blRet
'End Function
'Public Sub gPDFEmailEnd()
'    Set ogEmailer = Nothing
'    MyFile.WriteFacts "Ending sending emails", True
'
'End Sub
'Public Function gSendPDFEmail(slRecipient As String, slFromName As String, slFromAddress As String, slSubject As String, slMessage As String, slAttachment As String, slPayee As String, Optional slErrorMessage As String = "") As Boolean
'    'note: ogemailer must already be created.  I: slRecipient may be single email or multiple separated by ;  slFromName,slMessage,slSubject,slAttachment may be blank  slFromAddress and slRecipient wil be tested for validity
'    ' slAttachment should be complete path to file and will be tested that it actually exists.  O: true if sent ok.
'    'errors written to trafficErrors.txt, but not shown as msg box
'    '10016, added the payee for error messages
'
'    Dim blRet As Boolean
'    Dim slTo() As String
'    Dim c As Integer
'    Dim blAtLeastOne As Boolean
'On Error GoTo ERRBOX
'    blRet = True
'    slErrorMessage = ""
'    blAtLeastOne = False
'    If InStr(slRecipient, ";") > 0 Then
'        slTo = Split(slRecipient, ";")
'    Else
'        ReDim slTo(0)
'        slTo(0) = slRecipient
'    End If
'    'now build and send
'    If blRet Then
'        If Not ogEmailer Is Nothing Then
'            With ogEmailer
'                'because previous error message wasn't erased
'                .Clear False, True
'                For c = 0 To UBound(slTo)
'                    If Len(slTo) > 0 Then
'                        If .TestAddress(slTo(c)) Then
'                            blAtLeastOne = True
'                            .AddTOAddress slTo(c)
'                            '10016
'                            MyFile.WriteFacts " Payee: " & slPayee & " To: " & slTo(c)
'                            If Len(.ErrorMessage) > 0 Then
'                                blRet = False
'                                slErrorMessage = .ErrorMessage
'                                MyFile.WriteError .ErrorMessage, False, False
'                                GoTo CONTINUE
'                            End If
'                        Else
'                           ' slErrorMessage = slBadAddress & slTo(c) & ","
'                            slErrorMessage = slErrorMessage & " Payee: " & slPayee & ": " & slTo(c) & " " & .ErrorMessage & ","
'                        End If
'                    Else
'                        MyFile.WriteFacts " Payee " & slPayee & " missing email address."
'                        slErrorMessage = slErrorMessage & " Payee " & slPayee & " missing email address,"
'                    End If
'                Next c
'                If Len(slErrorMessage) > 0 Then
'                    slErrorMessage = "invalid emails: " & mLoseLastLetterIfComma(slErrorMessage)
'                    MyFile.WriteWarning slErrorMessage
'                End If
'                If blAtLeastOne Then
'                    .FromAddress = slFromAddress
'                    .FromName = slFromName
'                    .Subject = slSubject
'                    .Message = slMessage
'                    .Attachment = slAttachment
'                    '10016
'                    If bgPDFEmailTestMode = False Then
'                        If Not .Send() Then
'                            blRet = False
'                            slErrorMessage = .ErrorMessage & slErrorMessage
'                            MyFile.WriteError "Send failed.  " & slErrorMessage, True, False
'                        End If
'                    Else
'                        MyFile.WriteWarning "Send skipped", True
'                    End If
'                Else
'                    blRet = False
'                    slErrorMessage = "no valid 'to' address. Did not send. " & slErrorMessage
'                    MyFile.WriteError "no valid 'to' address. Did not send.", True, False
'                End If
'                '10016
'                If blRet Then
'                    MyFile.WriteFacts "Sent from " & slFromAddress & " Subject: " & slSubject & " Message: " & slMessage
'                End If
'            End With
'        Else
'            blRet = False
'            slErrorMessage = "ogEmailer does not exist"
'        End If
'    End If
'CONTINUE:
'    gSendPDFEmail = blRet
'    Exit Function
'ERRBOX:
'    gSendPDFEmail = False
'    slErrorMessage = err.Description
'End Function
Public Sub gSendServiceEmail(slSubject As String, slBody As String)
    'Dan M 12/9/14 for monitor program
    
    If slSubject = "" Then
        slSubject = "Automated message from a client"
    End If
    
    Set ogEmailer = New CEmail
    With ogEmailer
        .FromAddress = CSIUSERNAME '"AClient@Counterpoint.net"  'TTP 10837 JJB
        .FromName = Trim$(tgSpf.sGClient)
        .AddTOAddress "Service@counterpoint.net", "Service"
        .ToName = "Service"
        .Subject = slSubject
        .Message = slBody
        '10564 set tls to constant
        '.SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, False
        .SetHost CSISITE, CSIPORT, CSIUSERNAME, CSIPASSWORD, CSITLS
        If Not .Send() Then
            gLogMsg "Email could not be sent from mod PBEmail-gSendServiceEmail: ", "TrafficErrors.Txt", False
        End If
    End With
    Set ogEmailer = Nothing
End Sub
'Private Function mLoseLastLetterIfComma(slInput As String) As String
'    Dim llLength As Long
'    Dim slNewString As String
'    Dim llLastLetter As Long
'
'    llLength = Len(slInput)
'    llLastLetter = InStrRev(slInput, ",")
'    If llLength > 0 And llLastLetter = llLength Then
'        slNewString = Mid(slInput, 1, llLength - 1)
'    Else
'        slNewString = slInput
'    End If
'    mLoseLastLetterIfComma = slNewString
'End Function
'Type EmailInformation
'    sFromName As String
'    sToName As String
'    sFromAddress As String
'    sToAddress As String
'    sSubject As String
'    sMessage As String
'    sAttachment As String
'    'Dan M 9/7/10
'    sToMultiple As String
'    sCCMultiple As String   '"XX@counterpoint.net,YY@counterpoint.net,etc..."
'    sBCCMulitple As String
'    bTLSSet As Boolean
'    'Dan M 11/04/09 after losing site options, this no longer needed
'   ' bUserFromHasPriority As Boolean  'does what the user types in for from name, from address override what is in site options?
'End Type
'Public Function gEmailWithCounterpointHost(tlEmailInfo As EmailInformation) As Boolean
'    Dim myEmail As MailSender
'    Const CSISITE = "smtpauth.hosting.earthlink.net"
'    Const CSIPORT = 587
'    Const CSIUSERNAME = "emailSender@counterpoint.net"
'    Const CSIPASSWORD = "Csi44Sic"
'
'    Set myEmail = New MailSender
'    With myEmail
'        .Host = CSISITE
'        .Port = CSIPORT
'        .Username = CSIUSERNAME
'        .Password = CSIPASSWORD
'       ' .TLS = False
'    End With
'    tlEmailInfo.bTLSSet = True
'    gEmailWithCounterpointHost = gSendEmail(tlEmailInfo, , , myEmail)
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:gSendEmail                      *
''*                                                     *
''*             Created:9/10/09       By:D. Michaelson  *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: validate email, send, and      *
''*                     return result.  Was an 'm'      *
''*                     form function used on 2         *
''*                     different forms; now general    *
''*                                                     *
''*******************************************************
'
'Public Function gSendEmail(tlMyEmailInfo As EmailInformation, Optional ByRef ctrResultBox As control, Optional ZipBox As dzactxctrl, Optional Mail As MailSender) As Boolean
'' In: tlMyEmailInfo-- to/from/subject/message/attachment.  bUserFromHasPriority is confusing: use the values here over default(site options), or rather use default first and these values if default blank?
''                  -- blank to/from and blank default will mean message not sent.
'' In: ctrResultBox control to display result  text or listbox
'' In: Mail object.  frmLogEmail sets the 'to' of the mail object on the form, but could pass CC:, reply, etc.
'Dim slErrorFact As String
'Dim slNonZipFact As String
'Dim blthrowError As Boolean
'Dim slAttachments() As String
'Dim slFailedToZip() As String
'Dim slZipFile As String
'Dim blDeleteZipFile As Boolean
'Dim slWord As String    'variable word in string for zip failure
''Dan M 9/7/10 added to keep in sync with affiliate
'Dim slNames() As String
'Dim c As Integer
'Const MAILREGKEY = "16374-29451-54460"
'
'    If Mail Is Nothing Then
'        Set Mail = New ASPEMAILLib.MailSender
'    End If
'    Mail.RegKey = MAILREGKEY
'    'to allow verifying, passed mailsender will have host and other variables set.
'    If LenB(Mail.Host) = 0 Then
'    'Dan M 9/14/09 not sure if needed. If it doesn't find it here, it looks in registry
'    ' Dan M 10/08/09 replaced client's info with counterpoint's info
'    ' Dan M 02/07/11 use client's info.
'        'smtp
'        If Trim$(tgSite.sEmailHost) = "" Then
'            Mail.Host = ""
'            slErrorFact = "SMTP Host is undefined in Site Options."
'            blthrowError = True
'        Else
'            Mail.Host = Trim$(tgSite.sEmailHost)
'        End If
'        'port
'        If tgSite.iEmailPort = 0 Then
'            Mail.Port = 0
'            slErrorFact = slErrorFact & " Port Number is undefined in Site Options"
'            blthrowError = True
'        Else
'            Mail.Port = Trim$(tgSite.iEmailPort)
'        End If
'        'username
'        If LenB(Trim(tgSite.sEmailAcctName)) = 0 Then
'            Mail.Username = ""
'        Else
'            Mail.Username = Trim$(tgSite.sEmailAcctName)
'        End If
'        'password
'        If LenB(Trim(tgSite.sEmailPassword)) = 0 Then
'            Mail.Password = ""
'        Else
'            Mail.Password = Trim$(tgSite.sEmailPassword)
'        End If
'    End If
'    '     From Name
'    '    11/04/09 site options email options gone.
'     '    is what the user typed in first choice, or the one in site options?  Messages is first choice, User Status is the latter, so need to flag-bUserFromHasPriority
''        If (tlMyEmailInfo.bUserFromHasPriority Or LenB(Trim(tgSite.sEmailFromName)) = 0) Then
''            Mail.FromName = tlMyEmailInfo.sFromName
''        Else
''            Mail.FromName = Trim$(tgSite.sEmailFromName)
''        End If
'    ' verify passes in from name and from address
'    If Len(Mail.FromName) = 0 Then
'        Mail.FromName = tlMyEmailInfo.sFromName
'    End If
'     ' from address must be filled or will fail
'      'email site options no longer exist.
'    If Len(Mail.From) = 0 Then
'        If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0) Then
'    '        If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0 And LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
'            slErrorFact = slErrorFact & " No available 'from' address"
'            blthrowError = True
'        Else
'            Mail.From = tlMyEmailInfo.sFromAddress
'        End If
'    End If
''        If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0 And LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
''            slErrorFact = slErrorFact & " No available 'from' address"
''            blThrowError = True
''        ElseIf (tlMyEmailInfo.bUserFromHasPriority Or LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
''            Mail.From = tlMyEmailInfo.sFromAddress
''        Else
''            Mail.From = Trim$(tgSite.sEmailFromAddress)
''        End If
'     If Mail.ValidateAddress(Mail.From) <> 0 Then
'         slErrorFact = slErrorFact & " 'from' address is not a valid address."
'         blthrowError = True
'     End If
'     ' receipt? add this code Dan M 9/17/10
'    ' Mail.AddCustomHeader "Return-Receipt-To: " & Mail.From
'    ' Mail.AddCustomHeader "Disposition-Notification-To: " & Mail.From
'     ' To address, To Name
'     If LenB(tlMyEmailInfo.sToAddress) > 0 Then  'some mail objects have already set to address before coming to this procedure
'         If Mail.ValidateAddress(tlMyEmailInfo.sToAddress) <> 0 Then
'             slErrorFact = slErrorFact & " 'To' address is not a valid address."
'             blthrowError = True
'         Else
'             Mail.AddAddress tlMyEmailInfo.sToAddress, tlMyEmailInfo.sToName
'         End If
'    End If
'        'Dan added 9/7/10 for....changed type to include more To addresses, CC, Bcc
'     'To addresses
'     If Len(tlMyEmailInfo.sToMultiple) > 0 Then
'         slNames = Split(tlMyEmailInfo.sToMultiple, ",")
'         If Not mTestAddresses(slNames, Mail) Then
'             slErrorFact = slErrorFact & " One of the multiple 'To' addresses is not a valid address."
'             blthrowError = True
'         Else
'             For c = 0 To UBound(slNames)
'                 Mail.AddAddress slNames(c)
'             Next c
'         End If
'     End If
'     'CC
'     If Len(tlMyEmailInfo.sCCMultiple) > 0 Then
'         slNames = Split(tlMyEmailInfo.sCCMultiple, ",")
'         If Not mTestAddresses(slNames, Mail) Then
'             slErrorFact = slErrorFact & " One of the multiple 'CC' addresses is not a valid address."
'             blthrowError = True
'         Else
'             For c = 0 To UBound(slNames)
'                 Mail.AddCC slNames(c)
'             Next c
'         End If
'     End If
'     'bcc
'     If Len(tlMyEmailInfo.sBCCMulitple) > 0 Then
'         slNames = Split(tlMyEmailInfo.sBCCMulitple, ",")
'         If Not mTestAddresses(slNames, Mail) Then
'             slErrorFact = slErrorFact & " One of the multiple 'BCC' addresses is not a valid address."
'             blthrowError = True
'         Else
'             For c = 0 To UBound(slNames)
'                 Mail.AddBcc slNames(c)
'             Next c
'         End If
'     End If
'     'subject
'     Mail.Subject = tlMyEmailInfo.sSubject
'    'body
'      If LenB(Trim(tlMyEmailInfo.sMessage)) > 0 Then
'          Mail.Body = tlMyEmailInfo.sMessage
'      Else
'         Mail.Body = " ** No Message **"
'      End If
'      'attachments
'      If LenB(Trim(tlMyEmailInfo.sAttachment)) > 0 Then
'         slAttachments = Split(tlMyEmailInfo.sAttachment, ";")
'         If mAllFilesExist(slAttachments) Then
'             If ZipBox Is Nothing Then  'don 't zip
'             'Dan M 9/7/10 now defined at top
'                ' Dim c As Integer
'                 For c = 0 To UBound(slAttachments)  'test multiple not zipped
'                     Mail.AddAttachment slAttachments(c)
'                 Next c
'             Else
'                 ReDim slFailedToZip(0)
'                 slZipFile = mZipAllFiles(slAttachments, slFailedToZip, ZipBox)
'                 If (StrComp(slZipFile, "NoXne", vbBinaryCompare) <> 0) Then  ' not error zipping
'                     If (UBound(slAttachments) + 1 <> UBound(slFailedToZip)) Then  '  if all files can't be zipped, don't add attachment
'                         Mail.AddAttachment slZipFile
'                         blDeleteZipFile = True
'                     End If
'                     If UBound(slFailedToZip) > 0 Then   'nothing to do error from zipping: send with unzipped
'                         'code to write out message:
'                         If UBound(slFailedToZip) = UBound(slAttachments) + 1 Then   'all attachments failed
'                             slWord = " attached file"
'                             If UBound(slFailedToZip) > 1 Then 'more than one
'                                 slWord = slWord & "s"
'                             End If
'                         Else
'                             If UBound(slFailedToZip) > 1 Then
'                                 slWord = " some attached files"     'only some failed
'                             Else
'                                 slWord = " an attached file"
'                             End If
'                         End If
'                         slNonZipFact = " But" & slWord & " could not be zipped."
'                         For c = 0 To UBound(slFailedToZip) - 1 'test multiple not zipped
'                             Mail.AddAttachment slFailedToZip(c)
'                         Next c
'                     End If
'                 Else 'couldn't zip stop email
'                     slErrorFact = slErrorFact & " Email not sent. Attached files could not be zipped."
'                     blthrowError = True
'                 End If 'error zipping
'             End If  ' zip?
'         Else
'             slErrorFact = slErrorFact & " Email not sent.  Some attached files do not exist."
'             blthrowError = True
'         End If 'files exist?
'         Erase slAttachments
'     End If  'attachment?
'     'TLS
'     'failed and commented out 6/30/11
'     'dan added tls 6/28/11.  Verify passes value, don't grab from site.
'     If Not tlMyEmailInfo.bTLSSet Then
'        If tlMyEmailInfo.sFromName = "1" Then
'            Mail.TLS = True
'        Else
'            Mail.TLS = False
'        End If
'     End If
'  '  End If 'fill mail info if needed.
'    On Error Resume Next
'    If Not blthrowError Then
'        gSendEmail = Mail.Send ' send message
'    Else    'no address, attachment file issue, zipping issue
'        Err.Raise 5555, , " "
'    End If
'    If Not ctrResultBox Is Nothing Then
'        With ctrResultBox
'            If TypeOf ctrResultBox Is ListBox Then
'                .Clear
'                If Err <> 0 Then ' error occurred
'                    .ForeColor = vbRed
'                    .AddItem Err.Description & "  " & slErrorFact
'                Else
'                    .ForeColor = vbGreen
'                    .AddItem "Mail sent." & slNonZipFact     'attachments sent but could not be zipped.
'                End If
'            ElseIf TypeOf ctrResultBox Is TextBox Then
'                .Text = ""
'                If Err <> 0 Then ' error occurred
'                    .ForeColor = vbRed
'                    .Text = Err.Description & "  " & slErrorFact
'                Else
'                    .ForeColor = vbGreen
'                    .Text = "Mail sent." & slNonZipFact
'                End If
'            End If  'list box/text box
'        End With
'    End If      'send control?
'    If blDeleteZipFile Then
'        Kill slZipFile
'    End If
'    Erase slNames
'    Set Mail = Nothing
'End Function
''Public Function gSendEmail(tlMyEmailInfo As EmailInformation, Optional ByRef ctrResultBox As control, Optional ZipBox As dzactxctrl, Optional Mail As MailSender) As Boolean
''' In: tlMyEmailInfo-- to/from/subject/message/attachment.  bUserFromHasPriority is confusing: use the values here over default(site options), or rather use default first and these values if default blank?
'''                  -- blank to/from and blank default will mean message not sent.
''' In: ctrResultBox control to display result  text or listbox
''' In: Mail object.  frmLogEmail sets the 'to' of the mail object on the form, but could pass CC:, reply, etc.
''Dim slErrorFact As String
''Dim slNonZipFact As String
''Dim blThrowError As Boolean
''Dim slAttachments() As String
''Dim slFailedToZip() As String
''Dim slZipFile As String
''Dim blDeleteZipFile As Boolean
''Dim slWord As String    'variable word in string for zip failure
'''Dan M 9/7/10 added to keep in sync with affiliate
''Dim slNames() As String
''Dim c As Integer
''
'''Const CSISITE = "smtpauth.hosting.earthlink.net"
'''Const CSIPORT = 587
'''Const CSIUSERNAME = "emailSender@counterpoint.net"
'''Const CSIPASSWORD = "Csi44Sic"
''Const MAILREGKEY = "16374-29451-54460"
''
''    If Mail Is Nothing Then
''        Set Mail = New ASPEMAILLib.MailSender
''    End If
''    'to allow verifying, passed mailsender will have host and other variables set.
''    If LenB(Mail.Host) = 0 Then
''    'Dan M 9/14/09 not sure if needed. If it doesn't find it here, it looks in registry
'''to do this should be put in all objects!  move the other stuff to where might want to use our earthlink object.
'''        Mail.RegKey = MAILREGKEY '"16374-29451-54460"
''
'''        Mail.Host = CSISITE
'''        Mail.Port = CSIPORT
'''        Mail.Username = CSIUSERNAME
'''        Mail.Password = CSIPASSWORD
''    ' Dan M 10/08/09 replaced client's info with counterpoint's info
''    ' Dan M 02/07/11 use client's info.
''        'smtp
''        If Trim$(tgSite.sEmailHost) = "" Then
''            Mail.Host = ""
''            slErrorFact = "SMTP Host is undefined in Site Options."
''            blThrowError = True
''        Else
''            Mail.Host = Trim$(tgSite.sEmailHost)
''        End If
''        'port
''        If tgSite.iEmailPort = 0 Then
''            Mail.Port = 0
''            slErrorFact = slErrorFact & " Port Number is undefined in Site Options"
''            blThrowError = True
''        Else
''            Mail.Port = Trim$(tgSite.iEmailPort)
''        End If
''        'username
''        If LenB(Trim(tgSite.sEmailAcctName)) = 0 Then
''            Mail.Username = ""
''        Else
''            Mail.Username = Trim$(tgSite.sEmailAcctName)
''        End If
''        'password
''        If LenB(Trim(tgSite.sEmailPassword)) = 0 Then
''            Mail.Password = ""
''        Else
''            Mail.Password = Trim$(tgSite.sEmailPassword)
''        End If
''    '     From Name
''    '    11/04/09 site options email options gone.
''     '    is what the user typed in first choice, or the one in site options?  Messages is first choice, User Status is the latter, so need to flag-bUserFromHasPriority
''    '    If (tlMyEmailInfo.bUserFromHasPriority Or LenB(Trim(tgSite.sEmailFromName)) = 0) Then
''    '        Mail.FromName = tlMyEmailInfo.sFromName
''    '    Else
''    '        Mail.FromName = Trim$(tgSite.sEmailFromName)
''    '    End If
''        Mail.FromName = tlMyEmailInfo.sFromName
''        ' from address must be filled or will fail
''         'email site options no longer exist.
''        If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0) Then
''       ' If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0 And LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
''            slErrorFact = slErrorFact & " No available 'from' address"
''            blThrowError = True
''        Else
''            Mail.From = tlMyEmailInfo.sFromAddress
''        End If
''    '    If (LenB(Trim(tlMyEmailInfo.sFromAddress)) = 0 And LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
''    '        slErrorFact = slErrorFact & " No available 'from' address"
''    '        blThrowError = True
''    '    ElseIf (tlMyEmailInfo.bUserFromHasPriority Or LenB(Trim(tgSite.sEmailFromAddress)) = 0) Then
''    '        Mail.From = tlMyEmailInfo.sFromAddress
''    '    Else
''    '        Mail.From = Trim$(tgSite.sEmailFromAddress)
''    '    End If
''
''        If Mail.ValidateAddress(Mail.From) <> 0 Then
''            slErrorFact = slErrorFact & " 'from' address is not a valid address."
''            blThrowError = True
''        End If
''        ' receipt? add this code Dan M 9/17/10
''       ' Mail.AddCustomHeader "Return-Receipt-To: " & Mail.From
''       ' Mail.AddCustomHeader "Disposition-Notification-To: " & Mail.From
''        ' To address, To Name
''        If LenB(tlMyEmailInfo.sToAddress) > 0 Then  'some mail objects have already set to address before coming to this procedure
''            If Mail.ValidateAddress(tlMyEmailInfo.sToAddress) <> 0 Then
''                slErrorFact = slErrorFact & " 'To' address is not a valid address."
''                blThrowError = True
''            Else
''                Mail.AddAddress tlMyEmailInfo.sToAddress, tlMyEmailInfo.sToName
''            End If
''       End If
''           'Dan added 9/7/10 for....changed type to include more To addresses, CC, Bcc
''        'To addresses
''        If Len(tlMyEmailInfo.sToMultiple) > 0 Then
''            slNames = Split(tlMyEmailInfo.sToMultiple, ",")
''            If Not mTestAddresses(slNames, Mail) Then
''                slErrorFact = slErrorFact & " One of the multiple 'To' addresses is not a valid address."
''                blThrowError = True
''            Else
''                For c = 0 To UBound(slNames)
''                    Mail.AddAddress slNames(c)
''                Next c
''            End If
''        End If
''        'CC
''        If Len(tlMyEmailInfo.sCCMultiple) > 0 Then
''            slNames = Split(tlMyEmailInfo.sCCMultiple, ",")
''            If Not mTestAddresses(slNames, Mail) Then
''                slErrorFact = slErrorFact & " One of the multiple 'CC' addresses is not a valid address."
''                blThrowError = True
''            Else
''                For c = 0 To UBound(slNames)
''                    Mail.AddCC slNames(c)
''                Next c
''            End If
''        End If
''        'bcc
''        If Len(tlMyEmailInfo.sBCCMulitple) > 0 Then
''            slNames = Split(tlMyEmailInfo.sBCCMulitple, ",")
''            If Not mTestAddresses(slNames, Mail) Then
''                slErrorFact = slErrorFact & " One of the multiple 'BCC' addresses is not a valid address."
''                blThrowError = True
''            Else
''                For c = 0 To UBound(slNames)
''                    Mail.AddBcc slNames(c)
''                Next c
''            End If
''        End If
''
''        'subject
''        Mail.Subject = tlMyEmailInfo.sSubject
''       'body
''         If LenB(Trim(tlMyEmailInfo.sMessage)) > 0 Then
''             Mail.Body = tlMyEmailInfo.sMessage
''         Else
''            Mail.Body = " ** No Message **"
''         End If
''         'attachments
''         If LenB(Trim(tlMyEmailInfo.sAttachment)) > 0 Then
''            slAttachments = Split(tlMyEmailInfo.sAttachment, ";")
''            If mAllFilesExist(slAttachments) Then
''                If ZipBox Is Nothing Then  'don 't zip
''                'Dan M 9/7/10 now defined at top
''                   ' Dim c As Integer
''                    For c = 0 To UBound(slAttachments)  'test multiple not zipped
''                        Mail.AddAttachment slAttachments(c)
''                    Next c
''                Else
''                    ReDim slFailedToZip(0)
''                    slZipFile = mZipAllFiles(slAttachments, slFailedToZip, ZipBox)
''                    If (StrComp(slZipFile, "NoXne", vbBinaryCompare) <> 0) Then  ' not error zipping
''                        If (UBound(slAttachments) + 1 <> UBound(slFailedToZip)) Then  '  if all files can't be zipped, don't add attachment
''                            Mail.AddAttachment slZipFile
''                            blDeleteZipFile = True
''                        End If
''                        If UBound(slFailedToZip) > 0 Then   'nothing to do error from zipping: send with unzipped
''                            'code to write out message:
''                            If UBound(slFailedToZip) = UBound(slAttachments) + 1 Then   'all attachments failed
''                                slWord = " attached file"
''                                If UBound(slFailedToZip) > 1 Then 'more than one
''                                    slWord = slWord & "s"
''                                End If
''                            Else
''                                If UBound(slFailedToZip) > 1 Then
''                                    slWord = " some attached files"     'only some failed
''                                Else
''                                    slWord = " an attached file"
''                                End If
''                            End If
''                            slNonZipFact = " But" & slWord & " could not be zipped."
''                            For c = 0 To UBound(slFailedToZip) - 1 'test multiple not zipped
''                                Mail.AddAttachment slFailedToZip(c)
''                            Next c
''                        End If
''                    Else 'couldn't zip stop email
''                        slErrorFact = slErrorFact & " Email not sent. Attached files could not be zipped."
''                        blThrowError = True
''                    End If 'error zipping
''                End If  ' zip?
''            Else
''                slErrorFact = slErrorFact & " Email not sent.  Some attached files do not exist."
''                blThrowError = True
''            End If 'files exist?
''            Erase slAttachments
''        End If  'attachment?
''    End If 'fill mail info if needed.
''    Mail.RegKey = MAILREGKEY
''    On Error Resume Next
''    If Not blThrowError Then
''        gSendEmail = Mail.Send ' send message
''    Else    'no address, attachment file issue, zipping issue
''        Err.Raise 5555, , " "
''    End If
''    If Not ctrResultBox Is Nothing Then
''        With ctrResultBox
''            If TypeOf ctrResultBox Is ListBox Then
''                .Clear
''                If Err <> 0 Then ' error occurred
''                    .ForeColor = vbRed
''                    .AddItem Err.Description & "  " & slErrorFact
''                Else
''                    .ForeColor = vbGreen
''                    .AddItem "Mail sent." & slNonZipFact     'attachments sent but could not be zipped.
''                End If
''            ElseIf TypeOf ctrResultBox Is TextBox Then
''                .Text = ""
''                If Err <> 0 Then ' error occurred
''                    .ForeColor = vbRed
''                    .Text = Err.Description & "  " & slErrorFact
''                Else
''                    .ForeColor = vbGreen
''                    .Text = "Mail sent." & slNonZipFact
''                End If
''            End If  'list box/text box
''        End With
''    End If      'send control?
''    If blDeleteZipFile Then
''        Kill slZipFile
''    End If
''    Erase slNames
''    Set Mail = Nothing
''End Function
'
'Private Function mTestAddresses(slNames() As String, olMail As MailSender) As Boolean
'    Dim c As Integer
'
'    For c = 0 To UBound(slNames)
'        If olMail.ValidateAddress(slNames(c)) <> 0 Then
'            mTestAddresses = False
'            Exit Function
'        End If
'    Next c
'    mTestAddresses = True
'End Function
'
'Private Function mAllFilesExist(slFiles() As String) As Boolean
'Dim olFile As FileSystemObject
'Dim c As Integer
'    Set olFile = New FileSystemObject
'    For c = 0 To UBound(slFiles)
'        If Not olFile.FileExists(slFiles(c)) Then
'            mAllFilesExist = False
'            Set olFile = Nothing
'            Exit Function
'        End If
'    Next c
'    mAllFilesExist = True
'    Set olFile = Nothing
'End Function
''zipping procedures
'Private Function mZipAllFiles(ByRef slAttachments() As String, ByRef slFailure() As String, zpcDZip As dzactxctrl) As String
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilPos                         slStr                                                   *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim slData As String
'    Dim slDateTime As String
'    Dim ilLoop As Integer
'    Dim slZipPathName As String
'    Dim ilIndex As Integer
'    DoEvents
'    slDateTime = " " & Format$(gNow(), "ddmmyy")
'    slZipPathName = sgDBPath & gFileNameFilter(Trim$(tgSpf.sGClient)) & slDateTime & ".zip"  'BuildZipName
'    On Error Resume Next
'    Kill slZipPathName  'if errors zipping, file might exist from before
'    On Error GoTo 0
'    For ilLoop = 0 To UBound(slAttachments)
'        ilRet = mAddFileToZip(slZipPathName, slAttachments(ilLoop), zpcDZip)
'        If ilRet > 0 Then
'            If ilRet = 12 Then  'nothing to zip
'                ilIndex = UBound(slFailure)
'                ReDim Preserve slFailure(0 To ilIndex + 1)
'                slFailure(ilIndex) = slAttachments(ilLoop)
'            Else    'error
'               mZipAllFiles = "NoXne"
'               gChDrDir
'               Exit Function
'            End If
'        End If
'    Next ilLoop
'    mZipAllFiles = slZipPathName
'    gChDrDir
'    DoEvents
'
'End Function
'Private Function mAddFileToZip(szZip As String, szFile As String, zpcDZip As dzactxctrl) As Integer
'
'    'Init the Zip control structure
'    Call minitZIPCmdStruct(zpcDZip)
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
'
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Private Sub minitZIPCmdStruct(zpcDZip As dzactxctrl)
'    zpcDZip.ActionDZ = 0 'NO_ACTION
'    zpcDZip.AddCommentFlag = False
'    zpcDZip.AfterDateFlag = False
'    zpcDZip.BackgroundProcessFlag = False
'    zpcDZip.COMMENT = ""
'    zpcDZip.CompressionFactor = 5
'    zpcDZip.ConvertLFtoCRLFFlag = False
'    zpcDZip.Date = ""
'    zpcDZip.DeleteOriginalFlag = False
'    zpcDZip.DiagnosticFlag = False
'    zpcDZip.DontCompressTheseSuffixesFlag = False
'    zpcDZip.DosifyFlag = False
'    zpcDZip.EncryptFlag = False
'    zpcDZip.FixFlag = False
'    zpcDZip.FixHarderFlag = False
'    zpcDZip.GrowExistingFlag = False
'    zpcDZip.IncludeFollowing = ""
'    zpcDZip.IncludeOnlyFollowingFlag = False
'    zpcDZip.IncludeSysandHiddenFlag = False
'    zpcDZip.IncludeVolumeFlag = False
'    zpcDZip.ItemList = ""
'    zpcDZip.MajorStatusFlag = True
'    zpcDZip.MessageCallbackFlag = True
'    zpcDZip.MinorStatusFlag = True
'    zpcDZip.MultiVolumeControl = 0
'    zpcDZip.NoDirectoryEntriesFlag = True
'    zpcDZip.NoDirectoryNamesFlag = True
'
'    zpcDZip.OldAsLatestFlag = False
'    zpcDZip.PathForTempFlag = False
'    zpcDZip.QuietFlag = False
'    zpcDZip.RecurseFlag = False
'    zpcDZip.StoreSuffixes = ""
'    zpcDZip.TempPath = ""
'    zpcDZip.ZIPFile = ""
'
'    'Write out a log file in the windows sub directory
'    zpcDZip.ZipSubOptions = 256
'
'    ' added for rev 3.00
'    zpcDZip.RenameCallbackFlag = False
'    zpcDZip.ExtProgTitle = ""
'    zpcDZip.ZIPString = ""
'    'Dan m 9/14/09 don't show error message
'    zpcDZip.AllQuiet = True
'End Sub
'Public Sub gInsertNewFromAddress(hmCef As Integer, slAddress As String)
''Dan M no cef for email? update
'Dim olMail As ASPEMAILLib.MailSender
'Dim ilRet As Integer
'Dim ilRecLen As Integer
'Dim tlCef As CEF
'Dim hlUrf As Integer
'Dim tlUrf As URF
'Dim tlUrfSearchKey As INTKEY0    'URF key record image
'    Set olMail = New ASPEMAILLib.MailSender
'    If olMail.ValidateAddress(slAddress) = 0 Then
'        tlCef.lCode = 0
'        tlCef.sComment = slAddress
'        ilRecLen = Len(tlCef)
'        ilRet = btrInsert(hmCef, tlCef, ilRecLen, 0)
'        If ilRet = BTRV_ERR_NONE Then   'update urf with cef key
'            hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'            ilRet = btrOpen(hlUrf, "", sgDBPath & "urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'            If ilRet = BTRV_ERR_NONE Then
'                ilRecLen = Len(tlUrf)
'                tlUrfSearchKey.iCode = tgUrf(0).iCode
'                ilRet = btrGetEqual(hlUrf, tlUrf, ilRecLen, tlUrfSearchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                If ilRet = BTRV_ERR_NONE Then
'                    gUrfDecrypt tlUrf
'                    tlUrf.lEMailCefCode = tlCef.lCode
'                    gUrfEncrypt tlUrf
'                    ilRet = btrUpdate(hlUrf, tlUrf, ilRecLen)
'                    If ilRet = BTRV_ERR_NONE Then
'                       tgUrf(0).lEMailCefCode = tlCef.lCode
'                    End If  'updated urf
'                End If ' found urf
'                ilRet = btrClose(hlUrf)
'                btrDestroy hlUrf
'            End If ' opened urf
'        End If  'updated cef
'    End If  'valid address
'    Set olMail = Nothing
'End Sub
''Public Function gGetEmailAddress(hmCef As Integer, llEMailCefCode As Long) As String
''    Dim ilRet As Integer
''    Dim ilRecLen As Integer
''    Dim tlCef As CEF
''
''    gGetEmailAddress = ""
''    If llEMailCefCode > 0 Then
''        tlCef.sComment = ""
''        ilRecLen = Len(tlCef)    '1009
''        ilRet = btrGetEqual(hmCef, tlCef, ilRecLen, llEMailCefCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
''        If ilRet = BTRV_ERR_NONE Then
''            gGetEmailAddress = gStripChr0(tlCef.sComment)
''        End If
''    End If
''
''End Function
''Public Sub GCreateWaitPicture(myImage As PictureBox, ilIndex As Integer, myForm As Form)
''        Load myImage(ilIndex)
''        With myImage(ilIndex)
''            .Left = (myForm.ScaleWidth - .Width) / 2
''            .Top = (myForm.ScaleHeight - .Height) / 2
''            '.Top = Me.Top + (0.5 * Me.Height)
''            .ZOrder 0
''            .Picture = LoadPicture("")
''            .Width = 800
''            .Height = 800
''        End With
''End Sub
