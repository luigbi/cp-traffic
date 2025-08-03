VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'traffic and affiliate match.  CEmailUser and CEmailInfo change for each.
'public:  Send   Clear   SetHost(override getting from site)   testAddress  GetEmailAddress
'public:  AddToAddress AddCCAddress AddBCCAddress--these are the only ways to add a to address.  'toName' and 'toaddress' are only to send values back and forth to forms.
'public:  additional message is to pass info between forms.  Used in proposal xml.
'There is a global ogEmailer that uses this class.  After .send, should reset all values..from address, etc. because a failure to 'send' will set values to "".
Public FromName As String
Public FromAddress As String
'To is used just to pass info back and forth to email form
Public ToName As String
Public ToAddress As String
Public Subject As String
Public Message As String
Public Attachment As String
Public Success As Boolean
'proposal xml-write out that they should manually send attachment
Public AdditionalMessageIfSuccess As String
Public AdditionalMessageIfFail As String
'"XX@counterpoint.net,YY@counterpoint.net,etc..."
Private smBCCMultiple As String
Private AddressInfo As CEmailUser
Private bmIsHostOverride As Boolean
Private bmIsTLS As Boolean
Private smHost As String
Private imPort As Integer
Private smUserName As String
Private smPassword As String
Private smErrorMessage As String
' in order to have multiple 'to addresses' with 'to names', I write directly to mail object.  When I check that 'to' is set, need to tell that
'it is already in mailer.
Private bmToDirect As Boolean
'Private myEMailer As ASPEMAILLib.MailSender
Private myEmailer As CsiNetUtilities.CsiEmailer
Private Const MAILREGKEY = "16374-29451-54460"

Public Property Get UserName() As String
    UserName = AddressInfo.UserName
End Property
Public Property Get UserEmailAddress() As String
    UserEmailAddress = AddressInfo.UserEmail
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = Trim$(smErrorMessage)
End Property
Public Property Let ErrorMessage(ByVal slvalue As String)
    smErrorMessage = slvalue
End Property

Private Sub Class_Initialize()
Dim blTls As Boolean

On Error GoTo errbox
    'Set myEMailer = New ASPEMAILLib.MailSender
    'myEMailer.RegKey = MAILREGKEY
    Set myEmailer = New CsiNetUtilities.CsiEmailer
    Set AddressInfo = New CEmailUser
    mSetHostToSite
Exit Sub
errbox:
    smErrorMessage = "Could not initialize mail object"
    MsgBox smErrorMessage, vbCritical, "Email class error"
End Sub
Private Sub Class_Terminate()
    Set myEmailer = Nothing
    Set AddressInfo = Nothing
End Sub
Public Sub AddTOAddress(slAddress As String, Optional slName As String)
    '8245 blank address passes test but will cause error if passed to 'addaddress', so I added length test
    If Len(slAddress) > 0 Then
        If TestAddress(slAddress) Then
            myEmailer.AddAddress slAddress, slName
            bmToDirect = True
        Else
            smErrorMessage = " a 'To' is address not valid."
        End If
    Else
        smErrorMessage = " a 'To' address is blank."
    End If
End Sub
Public Sub AddCCAddress(slAddress As String, Optional slName As String)
    If TestAddress(slAddress) Then
        myEmailer.AddCC slAddress, slName
    Else
        smErrorMessage = " a 'CC' address is not valid."
    End If
End Sub
'must validate before adding!
Public Sub AddBCCAddress(slAddress As String)
    If LenB(smBCCMultiple) = 0 Then
        smBCCMultiple = slAddress
    Else
        smBCCMultiple = smBCCMultiple & "," & slAddress
    End If
End Sub
Public Sub Clear(blClearHost As Boolean, blClearReturn As Boolean)
    'addressInfo not cleared. Generally don't want to clear host, because must then reget from site.
    bmToDirect = False
    Success = False
    FromName = ""
    ToName = ""
    FromAddress = ""
    ToAddress = ""
    Subject = ""
    Message = ""
    Attachment = ""
    smBCCMultiple = ""
    '8245 move to 'blClearReturn
    'smErrorMessage = ""
    If blClearReturn Then
        smErrorMessage = ""
        AdditionalMessageIfSuccess = ""
        AdditionalMessageIfFail = ""
    End If
    ' can't just clear host, must get values from site.
    If blClearHost Then
        mSetHostToSite
    End If
    bmIsHostOverride = False
On Error GoTo errbox
'    Set myEMailer = New ASPEMAILLib.MailSender
'    myEMailer.RegKey = MAILREGKEY
    Set myEmailer = New CsiNetUtilities.CsiEmailer
Exit Sub
errbox:
    smErrorMessage = "Could not clear mail object; could not initialize."
End Sub

Public Function SetHost(slHost As String, ilPort As Integer, slUserName As String, slPassword As String, blIsTLS As Boolean) As Boolean
    Dim blOk As Boolean
      
    blOk = True
    smErrorMessage = ""
    If Len(slHost) > 0 Then
        smHost = slHost
    Else
        blOk = False
        smErrorMessage = smErrorMessage & "SMTP Host is undefined."
    End If
    If Len(slUserName) > 0 Then
        smUserName = slUserName
    Else
        blOk = False
        smErrorMessage = smErrorMessage & " Username is undefined."
    End If
    If Len(slPassword) > 0 Then
        smPassword = slPassword
    Else
        blOk = False
        smErrorMessage = smErrorMessage & " Password is undefined."
    End If
    bmIsTLS = blIsTLS
    If ilPort > 0 Then
        imPort = ilPort
    Else
        blOk = False
        smErrorMessage = smErrorMessage & " Port Number is undefined."
    End If
    If blOk Then
        bmIsHostOverride = True
    End If
    SetHost = blOk
End Function
Public Function TestAddress(slAddress As String) As Boolean
    Dim slNames(0) As String
    
    slNames(0) = slAddress
    TestAddress = mTestAddresses(slNames)
End Function
Public Function GetEmailAddress(llCefCode As Long) As String
    AddressInfo.Index = llCefCode
    GetEmailAddress = AddressInfo.GetEmail()
End Function
Public Function SetEmailAddress(slAddress As String) As Long
'returns cefCode
    AddressInfo.SetEmail slAddress
    SetEmailAddress = AddressInfo.Index
End Function
Public Function SetEmailAddressThisUser(slAddress As String) As Boolean
    AddressInfo.SetEmailThisUser slAddress
    If Len(AddressInfo.ErrorMessage) = 0 Then
        SetEmailAddressThisUser = True
    Else
        SetEmailAddressThisUser = False
    End If
End Function
Public Function Send(Optional ByRef ctrResultBox As control, Optional ZipBox As dzactxctrl) As Boolean
' In: ctrResultBox control to display result  text or listbox
'if not all fields not valid, all reset to ""
    Dim blOk As Boolean
    Dim slZipFile As String
    Dim blDeleteZipFile As Boolean
    
    If Len(Trim$(smErrorMessage)) > 0 Then
        'force mDisplayResult to show 'not sent' and write out smErrorMessage
        Err.Number = 77
        '10035
        If Not ctrResultBox Is Nothing Then
            mDisplayResult ctrResultBox, False
        End If
        Send = False
        Exit Function
    End If
    smErrorMessage = ""
    blOk = True
    If Not mValuesSet() Then
        Send = False
        If Not ctrResultBox Is Nothing Then
            mDisplayResult ctrResultBox, False
        End If
        If bmIsHostOverride Then
            Me.Clear True, False
        Else
            Me.Clear False, False
        End If
        Exit Function
    End If
    mSetToEmailer
    ' if true, error message is that some files couldn't be zipped; if false, error message says why
    If ZipBox Is Nothing Then
        blOk = mHandleAttachments()
    Else
        blOk = mHandleAttachmentsZip(ZipBox, blDeleteZipFile, slZipFile)
    End If
    If blOk Then
    On Error Resume Next
        'set to false, so if error in .send, this method will return false
        blOk = False
        blOk = myEmailer.Send()
        Send = blOk
        If Not blOk Then
            smErrorMessage = myEmailer.FailureReason
        End If
    End If
    ' read error if occurred
    If Not ctrResultBox Is Nothing Then
        mDisplayResult ctrResultBox, blOk
    End If      'send control?
    If blDeleteZipFile Then
        Kill slZipFile
    End If
End Function

'****************************************************************************
'
'****************************************************************************
Public Sub SpellCheckUsingMSWord(edcCtrl As TextBox)
    On Error GoTo Err_SpellCheckUsingMSWord
    Dim slText As String
    Dim SpellCheck As Object

    slText = edcCtrl.Text
    'Dan M 10/18/11 a blank textbox generates an error
    If Len(slText) = 0 Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    App.OleRequestPendingTimeout = 999999   ' Prevent the "Switch To" dialog from appearing.
    Set SpellCheck = CreateObject("Word.Application")
    SpellCheck.Visible = False
    Call mMinimizeWordIfOpen
    SpellCheck.Documents.Add                              'Open New Document (Hidden)
    Clipboard.Clear
    Clipboard.SetText slText, vbCFText                    'Copy Text To Clipboard
    SpellCheck.Selection.Paste                            'Paste Text Into WORD
    Call mBringWindowToTopMost
    SpellCheck.Visible = False
    SpellCheck.ActiveDocument.CheckSpelling               'Activate The Spell Checker
    SpellCheck.Visible = False                            'Hide WORD From User
    SpellCheck.ActiveDocument.Select                      'Select The Corrected Text
    SpellCheck.Selection.Cut                              'Cut The Text To Clipboard
    edcCtrl.Text = Clipboard.GetText(vbCFText)  'Assign Text To SPELLCHECKER Function
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
    edcCtrl.SetFocus
    edcCtrl.SelStart = Len(edcCtrl.Text)
    Screen.MousePointer = vbDefault
    Exit Sub

Err_SpellCheckUsingMSWord:
    Screen.MousePointer = vbDefault
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
    MsgBox "Error: " & Err.Number & ", " & Err.Description & vbCrLf & vbCrLf & "Please note you must have Microsoft Word installed to utilize the spell check feature.", vbExclamation, "Spell Check Problem"
End Sub
Private Sub mBringWindowToTopMost()
    Dim hwnd As Long
    Dim ilResult As Long

    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        ilResult = SetWindowPos(hwnd, WNDNOTOPMOST, 0, 0, 0, 0, FRMNOMOVE Or FRMNOSIZE)
    End If
End Sub

'****************************************************************************
' This function will look for a word doc that is currently open with the
' title of "Document1 - Microsoft Word", indicating a new blank word doc.
' If this is found, we need to minimize it to avoid having it become the
' top most visible window.
'
'****************************************************************************
Private Sub mMinimizeWordIfOpen()
    Dim hwnd As Long
    Dim wp As WINDOWPLACEMENT

    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        If GetWindowPlacement(hwnd, wp) > 0 Then
            wp.LENGTH = Len(wp)
            wp.Flags = 0&
            wp.showCmd = SW_SHOWMINIMIZED
            SetWindowPlacement hwnd, wp
        End If
    End If
End Sub
Private Sub mSetHostToSite()
    Dim myInfo As CEmailInfo
    
    Set myInfo = New CEmailInfo
    With myInfo
        .GetHostInfo
        Me.SetHost .Host, .Port, .User, .Password, .TLS
    End With
    bmIsHostOverride = False
    Set myInfo = Nothing
End Sub
Private Sub mDisplayResult(ByRef ctrResultBox As control, blOk As Boolean)
    Dim slNotSent As String
    
    slNotSent = "Mail not sent. "
    With ctrResultBox
        If TypeOf ctrResultBox Is ListBox Then
            .Clear
            If Err <> 0 Then ' error occurred
                .ForeColor = vbRed
                .AddItem slNotSent & Err.Description & "  " & smErrorMessage
            ElseIf Not blOk Then
                .ForeColor = vbRed
                .AddItem slNotSent & smErrorMessage
            Else
                .ForeColor = vbGreen
                .AddItem "Mail sent." & smErrorMessage     'attachments sent but could not be zipped.
            End If
        ElseIf TypeOf ctrResultBox Is TextBox Then
            .Text = ""
            If Err <> 0 Then ' error occurred
                .ForeColor = vbRed
                .Text = slNotSent & Err.Description & "  " & smErrorMessage
            ElseIf Not blOk Then
                .ForeColor = vbRed
                .Text = slNotSent & smErrorMessage
            Else
                .ForeColor = vbGreen
                .Text = "Mail sent." & smErrorMessage
            End If
        End If  'list box/text box
    End With
End Sub
Private Function mHandleAttachments() As Boolean
    Dim slAttachments() As String
    Dim c As Integer
    Dim blThrowError As Boolean
    
    If LenB(Trim(Attachment)) > 0 Then
       slAttachments = Split(Attachment, ";")
       If mAllFilesExist(slAttachments) Then
          For c = 0 To UBound(slAttachments)  'test multiple not zipped
              myEmailer.AddAttachment slAttachments(c)
          Next c
       Else
           smErrorMessage = "Some attached files do not exist."
           blThrowError = True
       End If 'files exist?
       Erase slAttachments
    End If  'attachment?
    mHandleAttachments = Not blThrowError
End Function
Private Function mHandleAttachmentsZip(ZipBox As dzactxctrl, blDeleteZipFile As Boolean, slZipFile As String) As Boolean
' O- blDeleteZipFile and zlZipFile

    Dim slAttachments() As String
    Dim c As Integer
    Dim blThrowError As Boolean
    Dim slFailedToZip() As String
    Dim slWord As String
    Dim slNonZipFact As String
    
    If LenB(Trim(Attachment)) > 0 Then
        slAttachments = Split(Attachment, ";")
        If mAllFilesExist(slAttachments) Then
        ReDim slFailedToZip(0)
        slZipFile = mZipAllFiles(slAttachments, slFailedToZip, ZipBox)
        If (StrComp(slZipFile, "NoXne", vbBinaryCompare) <> 0) Then  ' not error zipping
            If (UBound(slAttachments) + 1 <> UBound(slFailedToZip)) Then  '  if all files can't be zipped, don't add attachment
                myEmailer.AddAttachment slZipFile
                blDeleteZipFile = True
            End If
            If UBound(slFailedToZip) > 0 Then   'nothing to do error from zipping: send with unzipped
                'code to write out message:
                If UBound(slFailedToZip) = UBound(slAttachments) + 1 Then   'all attachments failed
                    slWord = " attached file"
                    If UBound(slFailedToZip) > 1 Then 'more than one
                        slWord = slWord & "s"
                    End If
                Else
                    If UBound(slFailedToZip) > 1 Then
                        slWord = " some attached files"     'only some failed
                    Else
                        slWord = " an attached file"
                    End If
                End If
                'slNonZipFact = " But" & slWord & " could not be zipped."
                smErrorMessage = " But" & slWord & " could not be zipped."
                For c = 0 To UBound(slFailedToZip) - 1 'test multiple not zipped
                    myEmailer.AddAttachment slFailedToZip(c)
                Next c
            End If
        Else 'couldn't zip stop email
            smErrorMessage = "Attached files could not be zipped."
            blThrowError = True
        End If 'error zipping
        Else
            smErrorMessage = "Some attached files do not exist."
            blThrowError = True
        End If 'files exist?
        Erase slAttachments
    End If  'attachment?
    mHandleAttachmentsZip = Not blThrowError
End Function
Private Sub mSetToEmailer()
    Dim slNames() As String
    Dim c As Integer

    With myEmailer
        .Host = smHost
        .Port = imPort
        .UserName = smUserName
        .Password = smPassword
        .From = FromAddress
        If Len(FromName) > 0 Then
            .FromName = FromName
        End If
        .TLS = bmIsTLS
        slNames = Split(smBCCMultiple, ",")
        If Not bmToDirect Then
            AddTOAddress ToAddress, ToName
        End If
        For c = 0 To UBound(slNames)
            .AddBcc slNames(c)
        Next c
        If LenB(Trim(Subject)) > 0 Then
            .Subject = Subject
        Else
           .Subject = " ** No Subject **"
        End If
        If LenB(Trim(Message)) > 0 Then
            .Body = Message
        Else
           .Body = " ** No Message **"
        End If
    End With

End Sub
Private Function mValuesSet() As Boolean
    Dim blThrowError As Boolean
    Dim blToSet As Boolean
    Dim slErrorFact As String
    Dim slNames() As String
'from address
    If LenB(FromAddress) = 0 Then
            slErrorFact = slErrorFact & " No available 'from' address"
            blThrowError = True
    End If
     If myEmailer.ValidateAddress(FromAddress) <> 0 Then
         slErrorFact = slErrorFact & " 'from' address is not a valid address."
         blThrowError = True
     End If
     'Dan 11/16/11
'to address
    'If Not bmToDirect Then
    If Not bmToDirect And Len(ToAddress) = 0 Then
        blThrowError = True
        slErrorFact = slErrorFact & " No 'To' address is set."
    End If
' to address and cc address validated when set.
     If Len(smBCCMultiple) > 0 Then
         slNames = Split(smBCCMultiple, ",")
         If Not mTestAddresses(slNames) Then
             slErrorFact = slErrorFact & " One of the 'BCC' addresses is not a valid address."
             blThrowError = True
         End If
     End If
'host info
    If Len(smHost) = 0 Or Len(smPassword) = 0 Or Len(smUserName) = 0 Or imPort = 0 Then
        blThrowError = True
        slErrorFact = slErrorFact & " Outgoing email options must be set in master site options-email tab."
    End If
    Erase slNames
    mValuesSet = Not blThrowError
    smErrorMessage = smErrorMessage & " " & slErrorFact
End Function
Private Function mTestAddresses(slNames() As String) As Boolean
    Dim c As Integer
      '9938
'    For c = 0 To UBound(slNames)
'        If myEMailer.ValidateAddress(slNames(c)) <> 0 Then
'            mTestAddresses = False
'            Exit Function
'        End If
'        If mTestForScrewyChars(slNames(c)) Then
'            mTestAddresses = False
'            Exit Function
'        End If
'    Next c
    smErrorMessage = ""
    For c = 0 To UBound(slNames)
        If mAdditionalValidation(slNames(c)) = False Then
            mTestAddresses = False
            Exit Function
        End If
        If myEmailer.ValidateAddress(slNames(c)) <> 0 Then
            mTestAddresses = False
            Exit Function
        End If
    Next c
    mTestAddresses = True
End Function
'8345 '9938 part of 'mAdditionalValidation' 9888 moved there
'Private Function mTestForScrewyChars(slString As String) As Boolean
'    Dim blRet As Boolean
'    Dim c As Integer
'    Dim ilAsc As Integer
'    Dim ilLen As Integer
'
'    ilLen = Len(slString)
'    blRet = False
'    For c = ilLen To 1 Step -1
'        ilAsc = Asc(Mid(slString, c, 1))
'        If ilAsc > 126 Then
'            blRet = True
'            Exit For
'        End If
'    Next c
''    '9888 can't end with period
''    If blRet = False Then
''        If InStrRev(slString, ".") = ilLen Then
''            blRet = True
''        End If
''    End If
'    mTestForScrewyChars = blRet
'End Function
Private Function mAllFilesExist(slFiles() As String) As Boolean
Dim olFile As FileSystemObject
Dim c As Integer
    Set olFile = New FileSystemObject
    For c = 0 To UBound(slFiles)
        If Not olFile.FILEEXISTS(slFiles(c)) Then
            mAllFilesExist = False
            Set olFile = Nothing
            Exit Function
        End If
    Next c
    mAllFilesExist = True
    Set olFile = Nothing
End Function
'zipping procedures
Private Function mZipAllFiles(ByRef slAttachments() As String, ByRef slFailure() As String, zpcDZip As dzactxctrl) As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                         slStr                                                   *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slData As String
    Dim slDateTime As String
    Dim ilLoop As Integer
    Dim slZipPathName As String
    Dim ilIndex As Integer
    DoEvents
    slDateTime = " " & Format$(gNow(), "ddmmyy")
    slZipPathName = sgDBPath & AddressInfo.Client & slDateTime & ".zip"
    On Error Resume Next
    Kill slZipPathName  'if errors zipping, file might exist from before
    On Error GoTo 0
    For ilLoop = 0 To UBound(slAttachments)
        ilRet = mAddFileToZip(slZipPathName, slAttachments(ilLoop), zpcDZip)
        If ilRet > 0 Then
            If ilRet = 12 Then  'nothing to zip
                ilIndex = UBound(slFailure)
                ReDim Preserve slFailure(0 To ilIndex + 1)
                slFailure(ilIndex) = slAttachments(ilLoop)
            Else    'error
               mZipAllFiles = "NoXne"
               ChDrDir
               Exit Function
            End If
        End If
    Next ilLoop
    mZipAllFiles = slZipPathName
    ChDrDir
    DoEvents

End Function
Private Function mAddFileToZip(szZip As String, szFile As String, zpcDZip As dzactxctrl) As Integer

    'Init the Zip control structure
    Call minitZIPCmdStruct(zpcDZip)

    zpcDZip.ZIPFile = szZip    'The ZIP file name
    zpcDZip.ItemList = szFile  'The file list to be added
    zpcDZip.BackgroundProcessFlag = True
    zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file
    'Returns the error code.  This code can be translated by the sub mTranslateErrors.
    'It is not currently being used to log to a file.
    mAddFileToZip = zpcDZip.ERRORCODE

End Function
'9938 add some vehicle email validation
Private Function mAdditionalValidation(slEmailAddress As String) As Boolean
    Dim blRet As Boolean
    Dim ilPos As Integer
    Dim ilLen As Integer
    Dim slTestForAnother As String
    
    Dim ilResult As Integer
    Dim slTemp As String
    Dim ilIdx As Integer
    Dim slEMsg As String
    Dim slExt As String
    Dim ilMax As Integer
    Dim ilYesNo As Integer
    
    blRet = True
    ilLen = Len(slEmailAddress)
    If ilLen < 6 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address is shorter than 6 characters. "
      '  sgErrorMsg = "ERROR: " & """" & slEmailAddress(ilLoop) & """" & " is not a valid email address." & Chr(13) & Chr(10) & "Your email address is shorter than 6 characters which is impossible."
    End If
    ilPos = InStrRev(slEmailAddress, ".")
    If ilPos = 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address does not contain a period. "
    ElseIf ilPos = ilLen Then
        blRet = False
        smErrorMessage = smErrorMessage & "A period cannot be the last character. "
    End If
    ilPos = InStr(1, slEmailAddress, "@", vbTextCompare)
    If ilPos = 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address does not contain an ampersand. "
    Else
        If ilPos = 1 Then
            blRet = False
            smErrorMessage = smErrorMessage & "@ Cannot be the first character. "
        ElseIf ilPos = ilLen Then
            blRet = False
            smErrorMessage = smErrorMessage & "@ Cannot be the last character. "
        'doesn't test if already first or last
        Else
            slTestForAnother = Mid(slEmailAddress, ilPos + 1)
            ilPos = InStr(1, slTestForAnother, "@", vbTextCompare)
            If ilPos > 0 Then
                blRet = False
                smErrorMessage = smErrorMessage & "Address contains more than one ampersand. "
            End If
        End If
    End If
    If InStr(1, slEmailAddress, ";", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains a semicolon. "
    End If
     If InStr(1, slEmailAddress, "..", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains two or more periods together. "
    End If
     If InStr(1, slEmailAddress, " ", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains a space. "
    End If
     If InStr(1, slEmailAddress, "[", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains a [ character. "
    End If
    If InStr(1, slEmailAddress, "]", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains a ] character. "
    End If
        If InStr(1, slEmailAddress, ",", vbTextCompare) <> 0 Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains a comma. "
    End If
    If mTestForScrewyChars(slEmailAddress) Then
        blRet = False
        smErrorMessage = smErrorMessage & "Address contains an ASCII character greater than 126. "
    End If
    mAdditionalValidation = blRet
End Function
Private Function mTestForScrewyChars(slString As String) As Boolean
    Dim blRet As Boolean
    Dim c As Integer
    Dim ilAsc As Integer
    
    blRet = False
    For c = Len(slString) To 1 Step -1
        ilAsc = Asc(Mid(slString, c, 1))
        If ilAsc > 126 Then
            blRet = True
            Exit For
        End If
    Next c
    mTestForScrewyChars = blRet
End Function

' **************************************************************************************
'
'  Procedure:  initZIPCmdStruct()
'
'  Purpose:  Set the ZIP control values
'
' **************************************************************************************
Private Sub minitZIPCmdStruct(zpcDZip As dzactxctrl)
    zpcDZip.ActionDZ = 0 'NO_ACTION
    zpcDZip.AddCommentFlag = False
    zpcDZip.AfterDateFlag = False
    zpcDZip.BackgroundProcessFlag = False
    zpcDZip.Comment = ""
    zpcDZip.CompressionFactor = 5
    zpcDZip.ConvertLFtoCRLFFlag = False
    zpcDZip.Date = ""
    zpcDZip.DeleteOriginalFlag = False
    zpcDZip.DiagnosticFlag = False
    zpcDZip.DontCompressTheseSuffixesFlag = False
    zpcDZip.DosifyFlag = False
    zpcDZip.EncryptFlag = False
    zpcDZip.FixFlag = False
    zpcDZip.FixHarderFlag = False
    zpcDZip.GrowExistingFlag = False
    zpcDZip.IncludeFollowing = ""
    zpcDZip.IncludeOnlyFollowingFlag = False
    zpcDZip.IncludeSysandHiddenFlag = False
    zpcDZip.IncludeVolumeFlag = False
    zpcDZip.ItemList = ""
    zpcDZip.MajorStatusFlag = True
    zpcDZip.MessageCallbackFlag = True
    zpcDZip.MinorStatusFlag = True
    zpcDZip.MultiVolumeControl = 0
    zpcDZip.NoDirectoryEntriesFlag = True
    zpcDZip.NoDirectoryNamesFlag = True
    
    zpcDZip.OldAsLatestFlag = False
    zpcDZip.PathForTempFlag = False
    zpcDZip.QuietFlag = False
    zpcDZip.RecurseFlag = False
    zpcDZip.StoreSuffixes = ""
    zpcDZip.TempPath = ""
    zpcDZip.ZIPFile = ""
    
    'Write out a log file in the windows sub directory
    zpcDZip.ZipSubOptions = 256
    
    ' added for rev 3.00
    zpcDZip.RenameCallbackFlag = False
    zpcDZip.ExtProgTitle = ""
    zpcDZip.ZIPString = ""
    'Dan m 9/14/09 don't show error message
    zpcDZip.AllQuiet = True
End Sub

Public Sub ChDrDir()
    If InStr(1, CurDir(), ":") > 0 Then 'colon exists
        ChDrive Left$(CurDir(), 2)  'Set the default drive
        ChDir CurDir()
    End If
End Sub
