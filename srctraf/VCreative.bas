Attribute VB_Name = "VCreative"
' Copyright 2013 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: vCreative.Bas
'
' June 20, 2013
' Written by: Doug Smith

Option Explicit
Option Compare Text

Type STATIONINFO
    iStationCode As Integer
    sMarket As String * 60
    sNetwork As String * 6
    sCallLetters As String * 40
    sBand As String * 2
End Type
Private tmStationInfo() As STATIONINFO

Type SALESPEOPLEINFO
    iSalesUserIDs As Integer
    iUrfCode As Integer
    lCefEmailCode As String
    sSalesFirstName As String * 40
    sSalesLastName As String * 40
    sSalesEmail As String * 40
End Type
Private tmSalsePeopleInfo() As SALESPEOPLEINFO

'Copy Rotations
Private tmCsf As CSF            'CSF record image
Private tmCsfSrchKey As LONGKEY0  'CSF key record image
Private hmCsf As Integer        'CSF Handle
Private imCsfRecLen As Integer      'CSF record length

'Vehicle
Private tmVef As VEF            'VEF record image
Private tmVefSrchKey As INTKEY0  'VEF key record image
Private hmVef As Integer        'VEF Handle
Private imVefRecLen As Integer      'VEF record length

'Multi-Name
Private hmMnf As Integer
Private tmMnf As MNF
Private tmMnfGp3 As MNF
Private imMnfRecLen As Integer
Private tmMnfSrchKey As INTKEY0
Private tmMnfGp3SrchKey As INTKEY0

Private smMMnfStamp As String
Private tmMMnf() As MNF

Private hmMef As Integer 'Media file handle
Private tmMef As MEF        'MCF record image
Private tmMefSrchKey1 As MEFKEY1    'MCF key record image
Private imMefRecLen As Integer        'MCF record length

Private hmCef As Integer
Private tmCef As CEF
Private tmCefSrchKey0 As LONGKEY0    'CEF key record image
Private imCefRecLen As Integer        'CEF record length

Private hmSaf As Integer
Private tmSaf As SAF            'Schedule Attributes record image
Private tmSafSrchKey1 As SAFKEY1    'Vef key record image
Private imSafRecLen As Integer

'Comments
Private hmCxf As Integer            'Comments file handle
Private tmCxf As CXF               'CXF record image
Private tmCxfSrchKey As LONGKEY0     'CXF key record image
Private imCxfRecLen As Integer         'CXF record length

'Contract
Private tmChf As CHF            'CHF record image
Private tmChfSrchKey As LONGKEY0  'CHF key record image
Private tmChfSrchKey1 As CHFKEY1  'CHF key record image
Private hmCHF As Integer        'CHF Handle
Private imCHFRecLen As Integer      'CHF record length

'Rotation
Private tmCrf As CRF            'CRF record image
Private tmCrfSrchKey As LONGKEY0  'CRF key record image
Private tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Private tmCrfSrchKey2 As CRFKEY2  'CRF key record image
Private tmCrfSrchKey3 As CRFKEY3  'CRF key record image
Private hmCrf As Integer        'CRF Handle
Private imCrfRecLen As Integer      'CRF record length
Private lmCrfOverlap() As Long  'Record position of records totally overlapped
'5/20/15
Private imCrfVefCode() As Integer

'5/20/15: Copy Vehicles
Dim hmCvf As Integer        'Contract header file handle
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey1 As LONGKEY0  'CVF key record image
Dim imCvfRecLen As Integer      'CVF record length

'Copy Usage
Private tmCuf As CUF            'CUF record image
Private tmCufSrchKey As CUFKEY0  'CUF key record image
Private tmCufSrchKey1 As CUFKEY1  'CUF key record image
Private hmCuf As Integer        'CUF Handle
Private imCufRecLen As Integer      'CUF record length

'Media Code
Private tmMcf As MCF            'MCF record image
Private tmMcfSrchKey As INTKEY0  'MCF key record image
Private imMcfRecLen As Integer      'MCF record length
Private imMcfCode As Integer

'Product
Private tmCpf As CPF            'CPF record image
Private tmCpfSrchKey As LONGKEY0  'CPF key record image
Private imCpfRecLen As Integer      'CPF record length

Private hmSlf As Integer            'Salesperson file handle
Private imSlfRecLen As Integer      'SLF record length
Private tmSlfSrchKey As INTKEY0     'SLF key image
Private tmSlf As SLF

Private tmCif As CIF            'CIF record image
Private tmCifSrchKey4 As CIFKEY4  'CIF key record image - used for vCreative

Private slPsErrors As String
Private smCallType As String

Private lmChfCode As Long
Private smMsg As String

'*** Start variables for the vCreative send ***

'User info
Private smCurUserFirstName As String
Private smCurUserLastName As String
Private smUserPswd As String
Private imUserID As Integer
Private smUserEmail As String
Private bmErrorsFound As Boolean

'INI info
Private smvCUserAgent As String
Private smvCFirmId As String
Private smvCStatus As String
Private smvCScheduledFlag As String
Private smvCTypeSalesperson As String
Private smvCUseCSICopyCodes As String
Private smvCUserName As String
Private smvCUserPswd As String

'INI URL Calls
Private smReqAuthURL As String
Private smAddNewCopyURL As String
Private smGetCompCopyURL As String
         
'Access Key obtained from vCreative
Private smvCCurUserAPIKey As String

'General Info
Private imAdf As Integer
Private smAdfname As String
Private lmCartID As Long
Private smCart As String
Private smScript As String
Private lmCIF_CopyID As Long
Private smCreativeTitle As String
Private smISCI As String
Private smProductName As String
Private smRotStartDate As String
Private smRotEndDate As String
Private smRotDueDate As String
Private imSpotLen As Integer
Private smContractNo As String
Private imBusCatID As Integer
Private smBusCatName As String
Private smNetwork As String
Private smLastRetrievalDate As String
Private lmContractNo As Long
Private smCallLetters As String
Private smCopySuccCount As Long
Private smCopyFailCount As Long
Private smCopyCompletedCount As Long
Private lmSlfCode() As Long
Private lmErrorCnt As Long
Private bmDemoMode As Boolean

Private Function mInit() As Boolean

    Dim ilRet As Integer
    Dim slStr As String
    
    mInit = False
    On Error GoTo ErrHand:
    
    lmErrorCnt = 0
    ilRet = gObtainUrf()
    ilRet = gObtainSalesperson()
    
    hmSaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSafRecLen = Len(tmSaf)
    
    hmCsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCsfRecLen = Len(tmCsf)
    
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVefRecLen = Len(tmVef)
    
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMnfRecLen = Len(tmMnf)
    
    hmMef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMef, "", sgDBPath & "Mef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMefRecLen = Len(tmMef)
    
    hmCef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCefRecLen = Len(tmCef)
    
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "CHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCHFRecLen = Len(tmChf)
    
    hmCxf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "CXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCxfRecLen = Len(tmCxf)
    
    hmCuf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCuf, "", sgDBPath & "Cuf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCufRecLen = Len(tmCuf)
    
    '5/20/15
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCvfRecLen = Len(tmCvf)
    
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCrfRecLen = Len(tmCrf)
    
    hmSlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSlfRecLen = Len(tmSlf)
    
    'Get the information in the vCreative INI file
    Call mLoadOption("GenInfo", "FirmID", smvCFirmId)
    Call mLoadOption("GenInfo", "Status", smvCStatus)
    Call mLoadOption("GenInfo", "ScheduledFlag", smvCScheduledFlag)
    Call mLoadOption("GenInfo", "TypeSalesperson", smvCTypeSalesperson)
    Call mLoadOption("GenInfo", "TrafficID", smvCUseCSICopyCodes)
    Call mLoadOption("Authorization", "Password", smvCUserPswd)
    Call mLoadOption("Authorization", "User", smvCUserName)
    Call mLoadOption("Authorization", "UserAgent", smvCUserAgent)
    Call mLoadOption("URL", "ReqAuthURL", smReqAuthURL)
    Call mLoadOption("URL", "AddNewCopyURL", smAddNewCopyURL)
    Call mLoadOption("URL", "GetCompCopyURL", smGetCompCopyURL)
    
    mInit = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mInit: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gGetvCreative(hCIF As Integer, iCifRecLen As Integer, hMcf As Integer, hCPF As Integer) As Integer

    Dim ilRet As Integer
    Dim ilUpdated As Integer
    Dim slDate As String
    Dim slNowDate As String
    Dim slTemp As String
    Dim ilOk As Boolean
    Dim slErrMsg As String
    Dim ilSlf As Long
    
    On Error GoTo ErrHand
    gGetvCreative = False
    bmErrorsFound = False
    ProgressMsg.Show
    DoEvents
    
    'D.S. 07/01/15
    'Demo mode for testing without a connection to vCreative's sand box.  This creates a text file in the Messages folder with all of the data that
    'would have been sent to vCreative.  In demo mode there is no testing of getting completed copy back from vCreative.
    bmDemoMode = False
    If igTestSystem Then
        bmDemoMode = True
    Else
        If sgUserName = "Guide" Then
            ilRet = MsgBox("Would you like to use Demo mode that does" & vbCrLf & "not require a connection to vCreative?", vbYesNo)
            If ilRet = vbYes Then
                bmDemoMode = True
            End If
        End If
    End If
    
    ProgressMsg.SetMessage 1, "Initializing Files"
    DoEvents
    ilOk = mInit()
    If Not ilOk Then
        gMsgBox "**** Initialize Program Failed. *****" & vbCrLf & "Please See Error Log For Details"
        Exit Function
    End If
    ReDim lmSlfCode(0 To 0) As Long
    If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
        'get authorization - if not get out
        mGetUserInfo (0) 'Current user info
        smvCCurUserAPIKey = ""
        
        ilRet = gAddNewUser()
        If Not ilRet Or (smvCCurUserAPIKey = "") Then
            gLogMsg "Get Authorization Failed.", "vCreativeErrors.Txt", False
            gMsgBox "Get Authorization Failed. See vCreativeErrors.Txt."
            lmErrorCnt = lmErrorCnt + 1
            Unload ProgressMsg
            DoEvents
            mClose
            bmErrorsFound = True
            Exit Function
        End If
        
        tmCifSrchKey4.sCleared = "N"
        ilRet = btrGetEqual(hCIF, tmCif, iCifRecLen, tmCifSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORWRITE)
        smCopySuccCount = 0
        smCopyFailCount = 0
        ProgressMsg.SetMessage 1, "Initializing Files" & vbCrLf & vbCrLf & "Sending Copy to vCreative"
        DoEvents
        Do While (ilRet = BTRV_ERR_NONE) And (tmCif.sCleared = "N")
            DoEvents
            ilUpdated = False
            gUnpackDate tmCif.iInvSentDate(0), tmCif.iInvSentDate(1), slDate
            If (slDate <> "1/1/1970") And (slDate <> "") And (tmCif.sPurged = "A") Then
                gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slDate
                If (slDate <> "1/1/1970") And (slDate <> "") Then
                    '*** Build fields to send to vCreative ***'
                    lmCIF_CopyID = tmCif.lCode
                    mGetLastSentdate
                    'rotation Dates
                    smRotStartDate = slDate
                    smRotDueDate = smRotStartDate
                    gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slDate
                    smRotEndDate = slDate
                    'Advertiser Code and Name
                    imAdf = gBinarySearchAdf(tmCif.iAdfCode)
                    If imAdf <> -1 Then
                        smAdfname = Trim$(tgCommAdf(imAdf).sName)
                    Else
                        smAdfname = "Missing: " & tmCif.iAdfCode
                    End If
                    'Spot Length
                    imSpotLen = tmCif.iLen
                    'ISCI, Copy, Creative
                    
                    'D.S. D.L. 06/27/19 Start Change
                    ilRet = mGetCopyData(hCPF, hMcf, tmCif, smCart, smISCI, smProductName, smCreativeTitle)
                    lmCartID = tmCif.lCode
                    'Contract Number
                    If Trim$(tmCif.sReel) <> "" Then
                        smContractNo = Trim$(tmCif.sReel)
                        lmContractNo = Val(smContractNo)
                        'this is acually the script in vCreative
                        mGetBusCatID lmContractNo
                        mGetScript tmCif.lCsfCode
                        mGetBusCatName
                        mBuildStaInfo tmCif.lCode
                        'sales people
                        mGetSalesUserInfo
                    Else
                        smContractNo = "Undefined"
                        imBusCatID = 9999
                        mGetScript tmCif.lCsfCode
                        smBusCatName = "Undefined"
                        mBuildStaInfo tmCif.lCode
                        For ilSlf = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
                              tmChf.iSlfCode(ilSlf) = 0
                        Next ilSlf
                        tmChf.iSlfCode(LBound(tmChf.iSlfCode)) = tgUrf(0).iSlfCode
                        mGetSalesUserInfo
                    End If
                    '***************************************
                    ilUpdated = False
                    ProgressMsg.SetMessage 0, "Initializing Files" & vbCrLf & vbCrLf & "Sending Copy to vCreative: " & smCopySuccCount   '& vbCrLf & vbCrLf & "Checking for Completed Copy"
                    DoEvents
                    ilRet = gAddNewCopy
                    If ilRet Then
                        ilUpdated = True
                    End If
                    '*************************************
                    'D.S. D.L. 06/27/19 End Change

'                    ilRet = mGetCopyData(hCPF, hMcf, tmCif, smCart, smISCI, smProductName, smCreativeTitle)
'                    lmCartID = tmCif.lCode
'                    'Contract Number
'                    'D.S. D.L. 6/25/19
'                    If Trim$(tmCif.sReel) <> "" Then
'                        smContractNo = Trim$(tmCif.sReel)
'                        lmContractNo = Val(smContractNo)
'                        'this is acually the script in vCreative
'                        mGetBusCatID lmContractNo
'                        If tmChf.lCode <> -1 Then
'                            mGetScript tmCif.lCsfCode
'                            mGetBusCatName
'                            mBuildStaInfo tmCif.lCode
'                            'sales people
'                            mGetSalesUserInfo
'                            '***************************************
'                            ilUpdated = False
'                            ProgressMsg.SetMessage 0, "Initializing Files" & vbCrLf & vbCrLf & "Sending Copy to vCreative: " & smCopySuccCount   '& vbCrLf & vbCrLf & "Checking for Completed Copy"
'                            DoEvents
'                            ilRet = gAddNewCopy
'                            If ilRet Then
'                                ilUpdated = True
'                            End If
'                            '*************************************
                End If
            End If
            If ilUpdated Then
                slNowDate = Format$(gNow(), "m/d/yy")
                tmCif.sCleared = "S"

                gPackDate slNowDate, tmCif.iInvSentDate(0), tmCif.iInvSentDate(1)
                ilRet = btrUpdate(hCIF, tmCif, iCifRecLen)
                tmCifSrchKey4.sCleared = "N"
                ilRet = btrGetEqual(hCIF, tmCif, iCifRecLen, tmCifSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORWRITE)
            Else
                ilRet = btrGetNext(hCIF, tmCif, iCifRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            End If
            
            DoEvents
        Loop
        If Not bmDemoMode Then
        'All done sending new copy now call for all of the completed copy.
        ilRet = gGetVCreativeCompCopy(False)
        If Not ilRet Then
            gLogMsg "Get Completed Copy Failed.", "vCreativeErrors.Txt", False
            gMsgBox "Get Completed Failed. See vCreativeErrors.Txt."
            lmErrorCnt = lmErrorCnt + 1
            bmErrorsFound = True
        End If
    End If
    End If
    If Not bmErrorsFound Then
        ProgressMsg.SetMessage 1, "Copy Sent: " & smCopySuccCount & vbCrLf & vbCrLf & "Copy Failed: " & smCopyFailCount & vbCrLf & vbCrLf & "Completed Copy: " & smCopyCompletedCount & vbCrLf & vbCrLf & "Program Completed Successfully."
        DoEvents
        gLogMsg "Copy Sent: " & smCopySuccCount & " Copy Failed: " & smCopyFailCount & " Completed Copy: " & smCopyCompletedCount, "vCreativeLog.Txt", False
        gLogMsg "**** Program Completed Successfully *****", "vCreativeLog.Txt", False
        DoEvents
    Else
        ProgressMsg.SetMessage 2, "Copy Sent: " & smCopySuccCount & vbCrLf & vbCrLf & "Copy Failed: " & smCopyFailCount & vbCrLf & vbCrLf & "Completed Copy: " & smCopyCompletedCount & vbCrLf & vbCrLf & "Program Completed With Errors."
        DoEvents
        gLogMsg "Copy Sent: " & smCopySuccCount & " Copy Failed: " & smCopyFailCount & " Completed Copy: " & smCopyCompletedCount, "vCreativeErrors.Txt", False
        gLogMsg "**** Program Completed With " & lmErrorCnt & " Errors *****" & vbCrLf & "Please See Error Log For Details", "vCreativeErrors.Txt", False
    End If
    
    DoEvents
    mClose
    'Sleep 2500
    'Unload ProgressMsg
    gGetvCreative = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - gGetvCreative: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl & vbCrLf & "Please See Error Log For Details", vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gAddNewUser() As Boolean
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    Dim slBody As String
    Dim slRet As String
    Dim slResponse As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilRetries As Integer
    
    On Error GoTo ErrHand
    gAddNewUser = False
    
    slBody = "username=" & smvCUserName
    slBody = slBody & "&password=" & smvCUserPswd
    slBody = slBody & "&firmid=" & smvCFirmId
    slBody = slBody & "&personid=" & imUserID
    slBody = slBody & "&firstname=" & smCurUserFirstName
    slBody = slBody & "&lastname=" & smCurUserLastName
    slBody = slBody & "&emailaddress=" & smUserEmail
    
    For ilRetries = 1 To 5 Step 1
        Set objHttp = New MSXML2.XMLHTTP60
        objHttp.Open "post", smReqAuthURL, False
        objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objHttp.setRequestHeader "User-Agent", smvCUserAgent
        If bmDemoMode Then
            llReturn = 200
            slResponse = "Test Mode, Test Mode"
        Else
        objHttp.Send slBody
        llReturn = objHttp.Status
        slResponse = objHttp.responseText
        End If
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn = 200 Then
            gLogMsg "AddNewUser - " & llReturn & " " & slResponse & " " & slBody, "vCreativeLog.Txt", False
            gAddNewUser = True
            Exit For
        Else
            gLogMsg "AddNewUser - " & llReturn & " " & slResponse & " " & slBody, "vCreativeErrors.Txt", False
            MsgBox "AddNewUser - " & llReturn & " " & slResponse & " " & slBody
            lmErrorCnt = lmErrorCnt + 1
            Call mSleep(2)
        End If
    Next ilRetries
    
    If gAddNewUser Then
        Call parse(slResponse)
        If bmDemoMode Then
            smCallType = "token"
            smvCCurUserAPIKey = "Test API Key"
        End If
        'basic sanity check
        If smCallType <> "token" Then
            gMsgBox "token Failed a Sanity Check " & vbCrLf & "Please See Error Log For Details", vbCritical
            Exit Function
        End If
        gAddNewUser = True
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - gAddNewUser "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gAddNewCopy() As Boolean
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    Dim slBody As String
    Dim slRet As String
    Dim slReturn As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim blFound As Boolean
    Dim llSale As Long
    
    On Error GoTo ErrHand
    gAddNewCopy = False

    'debug
    'If smCart = "AID1089" Then
    '    ilLoop = ilLoop
    'End If

    Set objHttp = New MSXML2.XMLHTTP60
    objHttp.Open "post", smAddNewCopyURL & lmCIF_CopyID, False
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.setRequestHeader "token", smvCCurUserAPIKey
    
    'Copy
    slBody = "{"
    slBody = slBody & """" & "clientname" & """" & " : " & """" & mFixDoubleQuoteAddEsc(smAdfname) & """" & ","
    slBody = slBody & """" & "title" & """" & " : " & """" & mFixDoubleQuoteAddEsc(smCreativeTitle) & """" & ","
    slBody = slBody & """" & "iscicode" & """" & " : " & """" & Trim$(smISCI) & """" & ","
    slBody = slBody & """" & "duedate" & """" & " : " & """" & Format(Trim$(smRotDueDate), "yyyy-mm-dd") & """" & ","
    slBody = slBody & """" & "startdate" & """" & " : " & """" & Format(Trim$(smRotStartDate), "yyyy-mm-dd") & """" & ","
    slBody = slBody & """" & "enddate" & """" & " : " & """" & Format(Trim$(smRotEndDate), "yyyy-mm-dd") & """" & ","
    slBody = slBody & """" & "length" & """" & " : " & """" & imSpotLen & """" & ","
    slBody = slBody & """" & "rotation" & """" & " : " & """" & "100" & """" & ","
    slBody = slBody & """" & "status" & """" & " : " & """" & smvCStatus & """" & ","
    slBody = slBody & """" & "adtypeid" & """" & " : " & """" & imBusCatID & """" & ","
    slBody = slBody & """" & "adtypedesc" & """" & " : " & """" & mFixDoubleQuoteAddEsc(smBusCatName) & """" & ","
    
    ' TTP 10646 JD 02-03-23
    If Len(Trim(smScript)) > 1 Then
        slBody = slBody & """" & "Production" & """" & " : " & "{"
    
        slTemp = smScript
        smScript = mRemoveCRLFAddEsc(slTemp)
        slTemp = smScript
        smScript = mFixDoubleQuoteAddEsc(slTemp)
        slBody = slBody & """" & "script" & """" & " : " & """" & smScript & """" & "},"
    End If
    
    'Station
    slBody = slBody & """" & "Station" & """" & " : [{"
    For ilLoop = LBound(tmStationInfo()) To UBound(tmStationInfo()) - 1 Step 1
        slBody = slBody & """" & "stationid" & """" & " : " & """" & Trim$(tmStationInfo(ilLoop).iStationCode) & """" & ","
        slBody = slBody & """" & "callletters" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(tmStationInfo(ilLoop).sCallLetters)) & """" & ","
        smCallLetters = Trim$(tmStationInfo(ilLoop).sCallLetters)
        slBody = slBody & """" & "stationband" & """" & " : " & """" & Trim$(tmStationInfo(ilLoop).sBand) & """" & ","
        slBody = slBody & """" & "market" & """" & " : " & """" & Trim$(tmStationInfo(ilLoop).sMarket) & """" & ","
        slBody = slBody & """" & "network" & """" & " : " & """" & Trim$(tmStationInfo(ilLoop).sNetwork) & """" & ","
        slBody = slBody & """" & "contractid" & """" & " : " & """" & Trim$(smContractNo) & """" & ","
        slBody = slBody & """" & "cartid" & """" & " : " & """" & Trim$(smCart) & """" & ","
        slBody = slBody & """" & "trafficid" & """" & " : " & """" & lmCIF_CopyID & """" & ","
        If ilLoop < UBound(tmStationInfo()) - 1 Then
            slBody = slBody & """" & "scheduledflag" & """" & " : " & """" & Trim$(smvCScheduledFlag) & """" & "},{"
        Else
            slBody = slBody & """" & "scheduledflag" & """" & " : " & """" & Trim$(smvCScheduledFlag) & """" & "}],"
        End If
    Next ilLoop
    
    'Sales
    slBody = slBody & """" & "POC" & """" & ": [{"
    For ilLoop = LBound(tmSalsePeopleInfo()) To UBound(tmSalsePeopleInfo()) Step 1
        If tmSalsePeopleInfo(ilLoop).iUrfCode > 0 Then

            slBody = slBody & """" & "personid" & """" & " : " & """" & tmSalsePeopleInfo(ilLoop).iUrfCode & """" & ","
            slBody = slBody & """" & "firstname" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(tmSalsePeopleInfo(ilLoop).sSalesFirstName)) & """" & ","
            slBody = slBody & """" & "lastname" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(tmSalsePeopleInfo(ilLoop).sSalesLastName)) & """" & ","
            slBody = slBody & """" & "emailaddress" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(tmSalsePeopleInfo(ilLoop).sSalesEmail)) & """" & ","
            slBody = slBody & """" & "firmid" & """" & " : " & """" & Trim$(smvCFirmId) & """" & ","
            slBody = slBody & """" & "howid" & """" & " : " & """" & Trim$(smvCTypeSalesperson) & """" & "},{"
            
            slBody = slBody & """" & "personid" & """" & " : " & """" & tgUrf(0).iCode & """" & ","
            slBody = slBody & """" & "firstname" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(smCurUserFirstName)) & """" & ","
            slBody = slBody & """" & "lastname" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(smCurUserLastName)) & """" & ","
            slBody = slBody & """" & "emailaddress" & """" & " : " & """" & mFixDoubleQuoteAddEsc(Trim$(smUserEmail)) & """" & ","
            slBody = slBody & """" & "firmid" & """" & " : " & """" & Trim$(smvCFirmId) & """" & ","
            slBody = slBody & """" & "howid" & """" & " : " & """" & Trim$("8") & """"
            
            
            If ilLoop < UBound(tmSalsePeopleInfo) Then
                slBody = slBody & "}"
                slBody = slBody & ","
                slBody = slBody & "{"
            Else
                slBody = slBody & "}"
                slBody = slBody & "]}"
            End If
        Else
            blFound = False
            For llSale = 0 To UBound(lmSlfCode) - 1 Step 1
                If lmSlfCode(llSale) = tmSalsePeopleInfo(ilLoop).iSalesUserIDs Then
                    blFound = True
                    Exit For
                End If
            Next llSale
            If Not blFound Then
                lmSlfCode(UBound(lmSlfCode)) = tmSalsePeopleInfo(ilLoop).iSalesUserIDs
                ReDim Preserve lmSlfCode(0 To UBound(lmSlfCode) + 1) As Long
                gLogMsg "gAddNewCopy - " & "Salesperson Not Defined in the User Options.  Please Correct. Salesperson " & Trim$(tmSalsePeopleInfo(ilLoop).sSalesFirstName) & " " & Trim$(tmSalsePeopleInfo(ilLoop).sSalesLastName), "vCreativeErrors.Txt", False
                lmErrorCnt = lmErrorCnt + 1
            End If

            Exit Function
        End If
    Next ilLoop
    
    slTemp = slBody
    slBody = gStripCntrlChars(slTemp)
    slTemp = Replace(slBody, "\\", "\")
    slBody = slTemp
    DoEvents
    If InStr(smCallLetters, "MISS") Then
        gLogMsg "gAddNewCopy - " & "No Call Letters Defined.  Please Correct. " & slBody, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        smCopyFailCount = smCopyFailCount + 1
        bmErrorsFound = True
        Exit Function
    Else
    If bmDemoMode Then
        llReturn = 200
        slReturn = "Test Mode, Test Mode"
    Else
        objHttp.Send slBody
        llReturn = objHttp.Status
        slReturn = objHttp.responseText
    End If
        gLogMsg "gAddNewCopy - " & llReturn & " " & slReturn & " " & slBody, "vCreativeLog.Txt", False
    
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn <> 200 Then
            gLogMsg "gAddNewCopy - " & llReturn & " " & slReturn & " " & slBody, "vCreativeErrors.Txt", False
            smCopyFailCount = smCopyFailCount + 1
            lmErrorCnt = lmErrorCnt + 1
            bmErrorsFound = True
            Exit Function
        End If
    
        If InStr(slReturn, "error") And llReturn = 200 Then
            gLogMsg "gAddNewCopy - " & llReturn & " " & slReturn & " " & slBody, "vCreativeErrors.Txt", False
            smCopyFailCount = smCopyFailCount + 1
            lmErrorCnt = lmErrorCnt + 1
            bmErrorsFound = True
            Exit Function
        End If
        
        If (Not InStr(slReturn, "error")) And llReturn = 200 Then
            smCopySuccCount = smCopySuccCount + 1
        End If
        
    End If
    
    gAddNewCopy = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - gAddNewCopy "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gGetVCreativeCompCopy(bFromReports As Boolean) As Integer

    Dim objHttp As MSXML2.XMLHTTP60
    Dim hlCif As Integer
    Dim ilCifRecLen As Integer
    Dim tlCif As CIF
    Dim tlCifSrchKey0 As LONGKEY0
    Dim llReturn As Long
    Dim slBody As String
    Dim slRet As String
    Dim slReturn As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilStartPos As Integer
    Dim ilEndPos As Integer
    Dim smRecordsArray() As String
    Dim ilLoop As Integer
    Dim llCifCode As Long
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim ilCompCopyfound As Integer
    Dim slStr As String
    
    On Error GoTo ErrHand
    
    gGetVCreativeCompCopy = False
    smCopyCompletedCount = 0
    
    If bFromReports Then
        DoEvents
        mInit
        mGetUserInfo 0
        If Not gAddNewUser() Then
            gGetVCreativeCompCopy = 1
            Exit Function
        End If
    End If
    
    mGetLastSentdate
    Set objHttp = New MSXML2.XMLHTTP60
    
    'D.S. 08/31/15 Change post to a get and pass all the variables on the URL
    'objHTTP.Open "post", smGetCompCopyURL, False
    objHttp.Open "get", smGetCompCopyURL & Format(Trim$(smLastRetrievalDate), "yyyy-mm-dd") & "/" & "trafficid", False
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.setRequestHeader "token", smvCCurUserAPIKey
    
    'D.S. 08/31/15
    'slbody = "{" & """" & "sincedate" & """"
    'slbody = slbody & " : " & """" & Format(Trim$(smLastRetrievalDate), "yyyy-mm-dd") & """"
    'slbody = slbody & "," & """" & "trafficidflag" & """" & " : "
    'slbody = slbody & """" & smvCUseCSICopyCodes & """" & "}"
    '
    objHttp.Send
    llReturn = objHttp.Status
    slReturn = objHttp.responseText
    'gLogMsg "gGetVCreativeCompCopy: " & slbody, "vCreativeLog.Txt", False
    gLogMsg "gGetVCreativeCompCopy: " & llReturn & ": " & slReturn, "vCreativeLog.Txt", False
    
    Set objHttp = Nothing
    'Anything but 200 is an error.
    If llReturn <> 200 Then
        gLogMsg slReturn, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
        If bFromReports Then
            gGetVCreativeCompCopy = 2
        End If
        Exit Function
    End If
    
    'Example
    'slReturn = "{" & """" & "trafficid" & """" & ":" & """" & "80806,80801,80876,80082,80056" & """" & "," & """" & "total" & """" & ":" & """" & "3" & """}"
    'slReturn = "{" & """" & "trafficid" & """" & ":" & """" & "80056" & """" & "," & """" & "total" & """" & ":" & """" & "3" & """}"
    
    'slReturn = "{trafficid : 20123, 20122, 20159, 20158, 20157, 100006,total : 6}"

    ilCompCopyfound = True
    If InStr(slReturn, "null") Then
        ilCompCopyfound = False
    End If
    
    If ilCompCopyfound Then
        hlCif = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        ilCifRecLen = Len(tmCif)
        
        ilStartPos = InStr(slReturn, ":")
        slRetStr = Mid(slReturn, ilStartPos + 2, Len(slReturn) - ilStartPos)
        ilEndPos = InStr(slRetStr, """")
        slRetStr = Left(slRetStr, ilEndPos - 1)
        smRecordsArray = Split(slRetStr, ",")
        If Not IsArray(smRecordsArray) Then
            If bFromReports Then
                gGetVCreativeCompCopy = 3
                Exit Function
            Else
                gLogMsg "gGetVCreativeCompCopy: " & slReturn & " Could not Interpret Return.", "vCreativeLog.Txt", False
            End If
            Exit Function
        End If
        
        For ilLoop = 0 To UBound(smRecordsArray)
            llCifCode = smRecordsArray(ilLoop)
            tlCifSrchKey0.lCode = llCifCode
            ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilCifRecLen = Len(tlCif)
                tlCif.sCleared = "Y"
                ilRet = btrUpdate(hlCif, tlCif, ilCifRecLen)
                If ilRet <> BTRV_ERR_NONE Then
                    'error
                    ilRet = ilRet
                Else
                    smCopyCompletedCount = smCopyCompletedCount + 1
                End If
            End If
        Next ilLoop
    End If
    
    If ilRet = BTRV_ERR_NONE Then
        tmSafSrchKey1.iVefCode = 0
        ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        slNowDate = Format$(gNow(), "m/d/yy")
        gPackDate slNowDate, tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
        ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
    End If
    
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    
    If bFromReports Then
        mClose
    End If
    
    gGetVCreativeCompCopy = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - gGetVCreativeCompCopy "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function
Private Function mFixDoubleQuoteAddEsc(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim llLoop As Long
    
    On Error GoTo ErrHand
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For llLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, llLoop, 1)
            If sChar = """" Then
                sOutStr = sOutStr & "\" & sChar
            Else
                sOutStr = sOutStr & sChar
            End If
        Next llLoop
    End If
    mFixDoubleQuoteAddEsc = sOutStr
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mFixDoubleQuoteAddEsc"
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Function mRemoveCRLFAddEsc(sInStr As String) As String
    
    Dim sOutStr As String
    Dim sChar As String
    Dim llPos As Long
    
    On Error GoTo ErrHand
    sOutStr = ""
    'Replace 0d 0a with \n
    llPos = InStr(sInStr, sgCR & sgLF)
    Do While llPos > 0
        Mid$(sInStr, llPos, 2) = "\n"
        llPos = InStr(sInStr, sgCR & sgLF)
    Loop

    mRemoveCRLFAddEsc = sInStr
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mRemoveCRLFAddEsc"
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function


Private Function mGetSalesUserInfo() As Boolean

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilSlf As Integer
    Dim ilIdx As Integer
    Dim ilFirstSlf As Integer
    Dim blFound As Boolean
    Dim ilSlfIndex As Integer
    Dim slFirstName As String
    Dim slLastName As String
    'Dim ilTemp As Integer
    'Dim tlSlfSrchKey As Integer
    'Dim tlSlf As SLF
    'Dim ilSlfRecLen As Integer
    
    On Error GoTo ErrHand
    mGetSalesUserInfo = False

    'debug
    'If smCart = "ADI1089" Then
    '    ilRet = ilRet
    'End If
    
    ilIdx = 0
    ilFirstSlf = -1
    ReDim tmSalsePeopleInfo(0 To 0) As SALESPEOPLEINFO
    For ilSlf = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
        If tmChf.iSlfCode(ilSlf) > 0 Then
            If ilFirstSlf = -1 Then
                ilFirstSlf = ilSlf
            End If
            blFound = False
            ilSlfIndex = gBinarySearchSlf(tmChf.iSlfCode(ilSlf))
            If ilSlfIndex >= 0 Then
                slFirstName = Trim$(tgMSlf(ilSlfIndex).sFirstName)
                slLastName = Trim$(tgMSlf(ilSlfIndex).sLastName)
                For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                    If tgPopUrf(ilLoop).iSlfCode = tmChf.iSlfCode(ilSlf) Then
                        blFound = True
                        ReDim Preserve tmSalsePeopleInfo(0 To ilIdx) As SALESPEOPLEINFO
                        tmSalsePeopleInfo(ilIdx).lCefEmailCode = tgPopUrf(ilLoop).lEMailCefCode
                        tmSalsePeopleInfo(ilIdx).sSalesEmail = Trim$(mGetEMail(tgPopUrf(ilLoop).lEMailCefCode))
                        tmSalsePeopleInfo(ilIdx).iSalesUserIDs = tmChf.iSlfCode(ilSlf)
                        tmSalsePeopleInfo(ilIdx).iUrfCode = tgPopUrf(ilLoop).iCode
                        'tlSlfSrchKey = tmChf.iSlfCode(ilSlf)
                        'ilRet = btrGetEqual(hmSlf, tlSlf, imSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        'If ilRet = BTRV_ERR_NONE Then
                        '    tmSalsePeopleInfo(ilIdx).sSalesFirstName = Trim$(tlSlf.sFirstName)
                        '    tmSalsePeopleInfo(ilIdx).sSalesLastName = Trim$(tlSlf.sLastname)
                        'End If
                        tmSalsePeopleInfo(ilIdx).sSalesFirstName = slFirstName
                        tmSalsePeopleInfo(ilIdx).sSalesLastName = slLastName
                        ilIdx = ilIdx + 1
                        Exit For
                    End If
                Next ilLoop
                If Not blFound Then
                    'Use the current user info
                    blFound = True
                    ReDim Preserve tmSalsePeopleInfo(0 To ilIdx) As SALESPEOPLEINFO
                    tmSalsePeopleInfo(ilIdx).lCefEmailCode = tgUrf(0).lEMailCefCode
                    tmSalsePeopleInfo(ilIdx).sSalesEmail = Trim$(mGetEMail(tgUrf(0).lEMailCefCode))
                    tmSalsePeopleInfo(ilIdx).iSalesUserIDs = 0
                    tmSalsePeopleInfo(ilIdx).iUrfCode = tgUrf(0).iCode
                    'tlSlfSrchKey = tmChf.iSlfCode(ilSlf)
                    'ilRet = btrGetEqual(hmSlf, tlSlf, imSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If ilRet = BTRV_ERR_NONE Then
                    '    tmSalsePeopleInfo(ilIdx).sSalesFirstName = Trim$(tlSlf.sFirstName)
                    '    tmSalsePeopleInfo(ilIdx).sSalesLastName = Trim$(tlSlf.sLastname)
                    'End If
                    tmSalsePeopleInfo(ilIdx).sSalesFirstName = smCurUserFirstName
                    tmSalsePeopleInfo(ilIdx).sSalesLastName = smCurUserLastName
                    ilIdx = ilIdx + 1
                End If
                If Not blFound Then
                    ReDim Preserve tmSalsePeopleInfo(0 To ilIdx) As SALESPEOPLEINFO
                    tmSalsePeopleInfo(ilIdx).lCefEmailCode = 0
                    tmSalsePeopleInfo(ilIdx).sSalesEmail = ""
                    tmSalsePeopleInfo(ilIdx).iSalesUserIDs = tmChf.iSlfCode(ilSlf)
                    tmSalsePeopleInfo(ilIdx).iUrfCode = -1
                    'tlSlfSrchKey = tmChf.iSlfCode(ilSlf)
                    'ilRet = btrGetEqual(hmSlf, tlSlf, imSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If ilRet = BTRV_ERR_NONE Then
                    '    tmSalsePeopleInfo(ilIdx).sSalesFirstName = Trim$(tlSlf.sFirstName)
                    '    tmSalsePeopleInfo(ilIdx).sSalesLastName = Trim$(tlSlf.sLastname)
                    'End If
                    tmSalsePeopleInfo(ilIdx).sSalesFirstName = slFirstName
                    tmSalsePeopleInfo(ilIdx).sSalesLastName = slLastName
                    ilIdx = ilIdx + 1
                    'gLogMsg "Can't find sales person in URF - User File.  Please Add", "vCreativeErrors.Txt", False
                End If
            Else
                gLogMsg "Error Salesperson not in global sales array.  Please Add", "vCreativeErrors.Txt", False
                lmErrorCnt = lmErrorCnt + 1
            End If
        End If
    Next ilSlf
    If ilFirstSlf = -1 Then
        'Use the current user info
        ReDim Preserve tmSalsePeopleInfo(0 To ilIdx) As SALESPEOPLEINFO
        tmSalsePeopleInfo(ilIdx).lCefEmailCode = tgUrf(0).lEMailCefCode
        tmSalsePeopleInfo(ilIdx).sSalesEmail = Trim$(mGetEMail(tgUrf(0).lEMailCefCode))
        tmSalsePeopleInfo(ilIdx).iSalesUserIDs = 0
        tmSalsePeopleInfo(ilIdx).iUrfCode = tgUrf(0).iCode
        tmSalsePeopleInfo(ilIdx).sSalesFirstName = smCurUserFirstName
        tmSalsePeopleInfo(ilIdx).sSalesLastName = smCurUserLastName
    End If
    'There was no salesperson so we grab the first salesperson in the contract
'    If tmSalsePeopleInfo(ilIdx).iUrfCode = -1 And ilFirstSlf <> -1 Then
'        ilIdx = 0
'        ilSlf = ilFirstSlf
'        tmSalsePeopleInfo(ilIdx).lCefEmailCode = 0
'        tmSalsePeopleInfo(ilIdx).sSalesEmail = ""
'        tmSalsePeopleInfo(ilIdx).iSalesUserIDs = tmChf.iSlfCode(ilSlf)
'        tmSalsePeopleInfo(ilIdx).iUrfCode = -1
'        'tlSlfSrchKey = tmChf.iSlfCode(ilSlf)
'        'ilRet = btrGetEqual(hmSlf, tlSlf, imSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        'If ilRet = BTRV_ERR_NONE Then
'        '    tmSalsePeopleInfo(ilIdx).sSalesFirstName = Trim$(tlSlf.sFirstName)
'        '    tmSalsePeopleInfo(ilIdx).sSalesLastName = Trim$(tlSlf.sLastName)
'        'End If
'    End If
    mGetSalesUserInfo = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetSalesUserInfo: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Sub mGetBusCatID(llCntrNo As Long)
    
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilTemp As Integer
    
    On Error GoTo ErrHand
    'this is acually the script in vCreative
    slStr = ""
    imCHFRecLen = Len(tmChf)
    tmChfSrchKey1.lCntrNo = llCntrNo
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) And (tmChf.sSchStatus <> "F")
        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) And (tmChf.sSchStatus = "F") Then
    
        'debug
        'If llCntrNo = 16167 Then
        '    ilTemp = ilTemp
        'End If
    
        ilTemp = tmChf.iMnfBus
        If ilTemp = 0 Then
            imBusCatID = 9999
        'D.S. D.L. 6/25/19
        Else
            imBusCatID = tmChf.iMnfBus
        End If
    Else
        tmChf.lCode = -1
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetBusCatID: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub

Private Function mGetContractNo() As String

    Dim ilRet As Integer
    Dim slResult As String
    Dim llResult As Long
    
    'use the CifCode to look in the Cuf for the Crf code for the Chf code and Vef code
    On Error GoTo ErrHand
    llResult = tmCuf.lCrfCode(0)
    mGetContractNo = slResult
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetContractNo: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function


Private Function mGetCopyData(hCPF As Integer, hMcf As Integer, tlCif As CIF, slCart As String, slISCI As String, slProductName As String, slCreativeTitle As String) As Integer

    Dim ilRet As Integer
    Dim slCopy As String
    
    On Error GoTo ErrHand
    
    mGetCopyData = False
    imCpfRecLen = Len(tmCpf)
    imMcfRecLen = Len(tmMcf)
    'Read CPF using lCpfCode from CIF to get COPY data
     slCopy = Trim$(tlCif.sName)
     slISCI = ""
     slProductName = ""
     slCreativeTitle = ""
    
     If tlCif.lcpfCode > 0 Then
         tmCpfSrchKey.lCode = tlCif.lcpfCode
         ilRet = btrGetEqual(hCPF, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
         If ilRet = BTRV_ERR_NONE Then
             slISCI = Trim$(tmCpf.sISCI)  ' ISCI Code
             smISCI = Trim$(slISCI)
             slProductName = Trim$(tmCpf.sName)
             If Trim$(tmCpf.sCreative) <> "" Then
                slCreativeTitle = Trim$(tmCpf.sCreative)
            Else
                slCreativeTitle = "   "
            End If
         End If
     End If

     ' Concatinate Copy from Media Code, Inv. Name
     smNetwork = ""
     If (tgSpf.sUseCartNo <> "N") And (tlCif.iMcfCode <> 0) Then
         tmMcfSrchKey.iCode = tlCif.iMcfCode
         ilRet = btrGetEqual(hMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
         If ilRet = BTRV_ERR_NONE Then
             ' Media Code is tmMcf.sName
             slCopy = Trim$(tmMcf.sName) & slCopy
             tmMefSrchKey1.iMcfCode = tmMcf.iCode
             ilRet = btrGetEqual(hmMef, tmMef, imMefRecLen, tmMefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
             If ilRet = BTRV_ERR_NONE Then
                 smNetwork = tmMef.sNetworkID
                Else
                    smNetwork = "Undefined"
             End If
         End If
     End If
     smCart = slCopy
     mGetCopyData = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetCopyData: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Sub mClose()

    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    ilRet = btrClose(hmCxf)
    btrDestroy hmCHF
    ilRet = btrClose(hmCuf)
    btrDestroy hmCuf
    '5/20/15
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmMef)
    btrDestroy hmMef
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    
    Erase tmStationInfo
    Erase tmSalsePeopleInfo
    Erase lmSlfCode
    Erase tmMMnf
    Erase lmCrfOverlap
    Erase lmSlfCode
    '5/20/15
    Erase imCrfVefCode
    
    Exit Sub

ErrHand:
    Resume Next
End Sub

Private Sub mGetUserInfo(iUrfCode As Integer)
    
    Dim ilPos As Integer
    Dim slLen As Integer
    Dim slUserName As String
    
    On Error GoTo ErrHand
        
    imUserID = tgUrf(iUrfCode).iCode
    If Len(Trim(tgUrf(iUrfCode).sRept)) > 0 Then
        slUserName = Trim(tgUrf(iUrfCode).sRept)
    Else
        slUserName = Trim(tgUrf(iUrfCode).sName)
    End If
    
    smUserEmail = mGetEMail(tgUrf(iUrfCode).lEMailCefCode)
    If smUserEmail = "" Then
        smUserEmail = "Undefined"
    End If
    smUserPswd = Trim$(tgUrf(iUrfCode).sPassword)
    
    ilPos = InStr(1, slUserName, " ", vbTextCompare)
    If ilPos > 0 Then
        smCurUserFirstName = Left$(slUserName, ilPos - 1)
        smCurUserLastName = Mid$(slUserName, ilPos + 1, Len(slUserName))
        If smCurUserLastName = "" Then
            smCurUserLastName = "Undefined"
        End If
    Else
        smCurUserFirstName = Trim(slUserName)
        smCurUserLastName = ""
    End If
    
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetUserInfo: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub
    
Private Function mGetEMail(llEMailCefCode As Long) As String
    
    Dim ilRet As Integer

    On Error GoTo ErrHand
    mGetEMail = ""
    If llEMailCefCode > 0 Then
        tmCefSrchKey0.lCode = llEMailCefCode
        tmCef.sComment = ""
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            mGetEMail = gStripChr0(tmCef.sComment)
        End If
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetEmail: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function
Public Function gStripCntrlChars(slInStr As String) As String
    
    Dim slStr As String
    Dim llChr As Long
    Dim llLoop As Long
    Dim llLen As Long

    slStr = ""
    If Len(slInStr) > 0 Then
        llLen = Len(slInStr)
        For llLoop = 1 To Len(slInStr)
            llChr = Asc(Mid(slInStr, llLoop, 1))
            If llChr < 32 Or llChr > 126 Then
                llLoop = llLoop
            Else
                slStr = slStr & Chr(llChr)
            End If
        Next
    End If
    gStripCntrlChars = Trim$(slStr)
End Function

Private Sub mGetBusCatName()

    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    If imBusCatID = 9999 Then
        smBusCatName = "Undefined"
        Exit Sub
    End If
    
    tmMnfSrchKey.iCode = 0
    tmMnfSrchKey.iCode = imBusCatID
    If tmMnf.iCode <> tmMnfSrchKey.iCode Then
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If Len(Trim$(tmMnf.sName)) = 0 Then
                tmMnf.sName = "Undefined"
            End If
        Else
            tmMnf.sName = "BTR Error " & ilRet
        End If
     End If
    smBusCatName = Trim$(tmMnf.sName)
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetBusCatName: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub

Private Sub mBuildStaInfo(lCifCode As Long)

    Dim ilVpf As Integer
    Dim ilRet As Integer
    Dim ilCrf As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim ilTempCode As Long
    Dim ilIdx As Integer
    Dim ilDuplicate As Long
    '5/20/15
    Dim ilCrfVef As Integer
    Dim ilfirstTime As Boolean
    Dim blFound As Boolean
    
    On Error GoTo ErrHand
    ilfirstTime = True
    ilIdx = 0
    ReDim tmStationInfo(0 To 0) As STATIONINFO
    tmCufSrchKey1.lCifCode = tmCif.lCode
    ilRet = btrGetEqual(hmCuf, tmCuf, imCufRecLen, tmCufSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCuf.lCifCode = tmCif.lCode)
        For ilCrf = 0 To UBound(tmCuf.lCrfCode) Step 1
            If tmCuf.lCrfCode(ilCrf) > 0 Then
                tmStationInfo(ilIdx).sMarket = ""
                tmCrfSrchKey.lCode = tmCuf.lCrfCode(ilCrf)
                ilRet = btrGetEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                '5/20/15
                mObtainCrfVehicle tmCrf.lCode, tmCrf.iVefCode
                For ilCrfVef = 0 To UBound(imCrfVefCode) - 1 Step 1
                    tmCrf.iVefCode = imCrfVefCode(ilCrfVef)
                    
                    tmStationInfo(ilIdx).iStationCode = tmCrf.iVefCode
                    ilVef = gBinarySearchVef(tmCrf.iVefCode)
                    If ilVef <> -1 Then
                        tmMnfGp3SrchKey.iCode = tgMVef(ilVef).iMnfVehGp3Mkt
                        If tmMnfGp3.iCode <> tmMnfGp3SrchKey.iCode Then
                            ilRet = btrGetEqual(hmMnf, tmMnfGp3, imMnfRecLen, tmMnfGp3SrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                If Len(Trim$(tmMnfGp3.sName)) = 0 Then
                                    tmMnfGp3.sName = "Undefined"
                                End If
                            Else
                                tmMnfGp3.sName = "BTR Error " & ilRet
                            End If
                         End If
                        tmStationInfo(ilIdx).sMarket = Trim$(tmMnfGp3.sName)
                    Else
                        'error condition
                    End If
                    ilVff = gBinarySearchVff(tmCrf.iVefCode)
                    If ilVff <> -1 Then
                        tmStationInfo(ilIdx).sCallLetters = Trim$(tgVff(ilVff).sASICallLetters)
                        tmStationInfo(ilIdx).sBand = Trim$(tgVff(ilVff).sASIBand)
                    Else
                        tmStationInfo(ilIdx).sCallLetters = "MISS"
                        'tmStationInfo(ilIdx).sBand = Trim$(tgVpf(ilVpf).sEDIBand) & "M"
                        tmStationInfo(ilIdx).sBand = "XX"
                    End If
                    If Trim$(tmStationInfo(ilIdx).sCallLetters) = "" Then
                        tmStationInfo(ilIdx).sCallLetters = "MISS"
                        'tmStationInfo(ilIdx).sBand = Trim$(tgVpf(ilVpf).sEDIBand) & "M"
                        tmStationInfo(ilIdx).sBand = "XX"
                    End If
                    tmStationInfo(ilIdx).sNetwork = smNetwork
                    
                    If ilfirstTime Then
                    ilIdx = ilIdx + 1
                    ReDim Preserve tmStationInfo(0 To ilIdx) As STATIONINFO
                    End If
                    If Not ilfirstTime Then
                        blFound = False
                        For ilDuplicate = 0 To UBound(tmStationInfo) - 1
                            If tmStationInfo(ilDuplicate).sBand = tmStationInfo(ilIdx).sBand And tmStationInfo(ilDuplicate).sCallLetters = tmStationInfo(ilIdx).sCallLetters And tmStationInfo(ilDuplicate).sMarket = tmStationInfo(ilIdx).sMarket And tmStationInfo(ilDuplicate).sNetwork = tmStationInfo(ilIdx).sNetwork Then
                                blFound = True
                                Exit For
                            End If
                        Next ilDuplicate
                        If Not blFound Then
                                ilIdx = ilIdx + 1
                                ReDim Preserve tmStationInfo(0 To ilIdx) As STATIONINFO
                            End If
                    End If
                    ilfirstTime = False
                '5/20/15
                Next ilCrfVef
            End If
        Next ilCrf
        ilRet = btrGetNext(hmCuf, tmCuf, imCufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mBuildStaInfo: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub

Private Sub mGetScript(lCsfCode As Long)

    Dim ilRet As Integer
    Dim slTemp As String
    Dim slRTFText As String
    
    On Error GoTo ErrHand
    smScript = ""
    tmCsfSrchKey.lCode = lCsfCode
    If lCsfCode <> 0 Then
        tmCsf.sComment = ""
        ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmCsf.lCode = 0
            tmCsf.sComment = ""
            'tmCsf.iStrLen = 0
        Else
            'comments are stored in RTF which is not HTML compatable.
            'Strip out all of the RTF control characters
            slRTFText = gStripChr0(tmCsf.sComment)
            Traffic.RichTextBox1.TextRTF = ""
            Traffic.RichTextBox1.TextRTF = Trim$(slRTFText)
            smScript = Trim$(Traffic.RichTextBox1.Text)
        End If
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mGetScript: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub

Public Function mLoadOption(Section As String, Key As String, sValue As String) As Boolean
    
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    Dim slFileName As String

    On Error GoTo ErrHand
    mLoadOption = False
    If igDirectCall = -1 Then
        slFileName = sgIniPath & "vCreative.Ini"
    Else
        slFileName = CurDir$ & "\vCreative.Ini"
    End If

    mLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            mLoadOption = True
        End If
    End If
    mLoadOption = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mLoadOption: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function

End Function

Private Sub mGetLastSentdate()

    Dim slStr As String

    On Error GoTo ErrHand
    gUnpackDate tgSaf(0).iVCreativeDate(0), tgSaf(0).iVCreativeDate(1), slStr
    If gValidDate(slStr) Then
        If gDateValue(slStr) <> gDateValue("1/1/1990") Then
            smLastRetrievalDate = Format(Trim$(slStr), "yyyy-mm-dd")
        Else
            smLastRetrievalDate = "1990-01-01"
        End If
    Else
        smLastRetrievalDate = "0000-00-00"
    End If
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - mLoadOption: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub
Sub mSleep(ilTotalSeconds As Long)

    Dim ilLoop As Integer
    For ilLoop = 0 To ilTotalSeconds
        DoEvents
        Sleep (1000)   ' Wait 1 second
    Next
End Sub


Public Function parse(ByRef str As String) As Object

   Dim Index As Long
   
    On Error GoTo ErrHand
    Index = 1
    slPsErrors = ""
    On Error Resume Next
    Call skipChar(str, Index)
    Select Case Mid(str, Index, 1)
       Case "{"
          Set parse = parseObject(str, Index)
       Case "["
          'Set parse = parseArray(str, Index)
       Case Else
          slPsErrors = "Invalid JSON"
    End Select
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parse: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Sub skipChar(ByRef str As String, ByRef Index As Long)
   
    Dim bComment As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment As Boolean
   
    On Error GoTo ErrHand
    Do While Index > 0 And Index <= Len(str)
        Select Case Mid(str, Index, 1)
        Case vbCr, vbLf
            If Not bLongComment Then
                bStartComment = False
                bComment = False
            End If
        Case vbTab, " ", "(", ")"
         
        Case "/"
            If Not bLongComment Then
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                Else
                    bStartComment = True
                    bComment = False
                    bLongComment = False
                End If
            Else
                If bStartComment Then
                    bLongComment = False
                    bStartComment = False
                    bComment = False
                    End If
                End If
        Case "*"
            If bStartComment Then
                bStartComment = False
                bComment = True
                bLongComment = True
            Else
                bStartComment = True
            End If
         
        Case Else
            If Not bComment Then
                Exit Do
            End If
        End Select
      
        Index = Index + 1
    Loop
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - skipChar: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Sub
End Sub

'Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection
'
'    On Error GoTo ErrHand
'    Set parseArray = New Collection
'
'   Call skipChar(str, Index)
'   If Mid(str, Index, 1) <> "[" Then
'      slPsErrors = slPsErrors & "Invalid Array at position " & Index & " : " + Mid(str, Index, 20) & vbCrLf
'      Exit Function
'   End If
'   Index = Index + 1
'
'   Do
'      Call skipChar(str, Index)
'      If "]" = Mid(str, Index, 1) Then
'         Index = Index + 1
'         Exit Do
'      ElseIf "," = Mid(str, Index, 1) Then
'         Index = Index + 1
'         Call skipChar(str, Index)
'      ElseIf Index > Len(str) Then
'         slPsErrors = slPsErrors & "Missing ']': " & right(str, 20) & vbCrLf
'         Exit Do
'      End If
'
'      ' add value
'      On Error Resume Next
'      parseArray.Add parseValue(str, Index)
'      If Err.Number <> 0 Then
'         slPsErrors = slPsErrors & Err.Description & ": " & Mid(str, Index, 20) & vbCrLf
'         Exit Do
'      End If
'   Loop
'    Exit Function
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'    If (Err.Number <> 0) Then
'        smMsg = "A general error has occured in vCreative.bas - parseArray: "
'        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
'        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'        bmErrorsFound = True
'    End If
'    Exit Function
'End Function

Private Function parseValue(ByRef str As String, ByRef Index As Long) As Object

    On Error GoTo ErrHand
    Call skipChar(str, Index)

    Select Case Mid(str, Index, 1)
        Case "{"
            Set parseValue = parseObject(str, Index)
        Case "["
            'Set parseValue = parseArray(str, Index)
        Case """", "'"
            parseValue = parseString(str, Index)
        Case "t", "f"
            parseValue = parseBoolean(str, Index)
        Case "n"
            'parseValue = parseNull(str, Index)
        Case Else
            parseValue = parseNumber(str, Index)
    End Select
    Exit Function
    
ErrHand:
'    Screen.MousePointer = vbDefault
'    If (Err.Number <> 0) Then
'        smMsg = "A general error has occured in vCreative.bas - parseValue: "
'        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
'        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'        bmErrorsFound = True
'    End If
    Exit Function
End Function

Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary

    Set parseObject = New Dictionary
    Dim sKey As String
   
    On Error GoTo ErrHand
    ' "{"
    Call skipChar(str, Index)
    If Mid(str, Index, 1) <> "{" Then
        slPsErrors = slPsErrors & "Invalid Object at position " & Index & " : " & Mid(str, Index) & vbCrLf
        Exit Function
    End If
    
    Index = Index + 1

    Do
        Call skipChar(str, Index)
        If "}" = Mid(str, Index, 1) Then
            Index = Index + 1
            Exit Do
        ElseIf "," = Mid(str, Index, 1) Then
            Index = Index + 1
            Call skipChar(str, Index)
        ElseIf Index > Len(str) Then
            slPsErrors = slPsErrors & "Missing '}': " & right(str, 20) & vbCrLf
            Exit Do
        End If

        'add key/value pair
        sKey = parseKey(str, Index)
        smCallType = sKey
        On Error Resume Next
      
        parseObject.Add sKey, parseValue(str, Index)
        If err.Number <> 0 Then
            slPsErrors = slPsErrors & err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parseValue: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Function parseString(ByRef str As String, ByRef Index As Long) As String

    Dim quote As String
    Dim Char  As String
    Dim Code  As String
    Dim SB As New cStringBuilder

    On Error GoTo ErrHand
    parseString = ""
    Call skipChar(str, Index)
    quote = Mid(str, Index, 1)
    Index = Index + 1

    Do While Index > 0 And Index <= Len(str)
        Char = Mid(str, Index, 1)
        Select Case (Char)
            Case "\"
                Index = Index + 1
                Char = Mid(str, Index, 1)
                Select Case (Char)
                    Case """", "\", "/", "'"
                        SB.Append Char
                        Index = Index + 1
                    Case "b"
                        SB.Append vbBack
                        Index = Index + 1
                    Case "f"
                        SB.Append vbFormFeed
                        Index = Index + 1
                    Case "n"
                        SB.Append vbLf
                        Index = Index + 1
                    Case "r"
                        SB.Append vbCr
                        Index = Index + 1
                    Case "t"
                        SB.Append vbTab
                        Index = Index + 1
                    Case "u"
                        Index = Index + 1
                        Code = Mid(str, Index, 4)
                        SB.Append ChrW(Val("&h" + Code))
                        Index = Index + 4
                    End Select
            Case quote
                Index = Index + 1
                parseString = SB.toString
                smvCCurUserAPIKey = parseString
                Set SB = Nothing
                Exit Function
            Case Else
                SB.Append Char
                Index = Index + 1
        End Select
    Loop

    parseString = SB.toString
    Set SB = Nothing
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parseString: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean

    On Error GoTo ErrHand
    
    parseBoolean = False
    Call skipChar(str, Index)
    If Mid(str, Index, 4) = "true" Then
        parseBoolean = True
        Index = Index + 4
    ElseIf Mid(str, Index, 5) = "false" Then
        parseBoolean = False
        Index = Index + 5
    Else
        slPsErrors = slPsErrors & "Invalid Boolean at position " & Index & " : " & Mid(str, Index) & vbCrLf
    End If
    parseBoolean = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parseBoolean: "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

'Private Function parseNull(ByRef str As String, ByRef Index As Long) As Boolean
'
'    On Error GoTo ErrHand
'
'    parseNull = False
'    Call skipChar(str, Index)
'    If Mid(str, Index, 4) = "null" Then
'        parseNull = Null
'        Index = Index + 4
'    Else
'        slPsErrors = slPsErrors & "Invalid null value at position " & Index & " : " & Mid(str, Index) & vbCrLf
'    End If
'    parseNull = True
'    Exit Function
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'    If (Err.Number <> 0) Then
'        smMsg = "A general error has occured in vCreative.bas - parseNull: "
'        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
'        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'        bmErrorsFound = True
'    End If
'    Exit Function
'End Function

Private Function parseNumber(ByRef str As String, ByRef Index As Long) As Boolean

    Dim Value As String
    Dim Char As String
    
    On Error GoTo ErrHand
    parseNumber = False
    Call skipChar(str, Index)
    Do While Index > 0 And Index <= Len(str)
        Char = Mid(str, Index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            Index = Index + 1
        Else
            parseNumber = CDec(Value)
            Exit Function
        End If
    Loop
    parseNumber = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parseNumber "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function


Private Function parseKey(ByRef str As String, ByRef Index As Long) As String

    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim Char    As String

    On Error GoTo ErrHand
    parseKey = ""
    Call skipChar(str, Index)
    Do While Index > 0 And Index <= Len(str)
        Char = Mid(str, Index, 1)
        Select Case (Char)
            Case """"
                dquote = Not dquote
                Index = Index + 1
                If Not dquote Then
                    Call skipChar(str, Index)
                    If Mid(str, Index, 1) <> ":" Then
                        slPsErrors = slPsErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case "'"
                squote = Not squote
                Index = Index + 1
                If Not squote Then
                    Call skipChar(str, Index)
                    If Mid(str, Index, 1) <> ":" Then
                        slPsErrors = slPsErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case ":"
                Index = Index + 1
                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & Char
                End If
            Case Else
                If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
                Else
                    parseKey = parseKey & Char
                End If
                Index = Index + 1
        End Select
    Loop
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (err.Number <> 0) Then
        smMsg = "A general error has occured in vCreative.bas - parseKey "
        gLogMsg "Error: " & smMsg & err.Description & " Error #" & err.Number & "; Line #" & Erl, "vCreativeErrors.Txt", False
        lmErrorCnt = lmErrorCnt + 1
        gMsgBox smMsg & err.Description & "; Error #" & err.Number & "; Line #" & Erl, vbCritical
        bmErrorsFound = True
    End If
    Exit Function
End Function

Private Sub mObtainCrfVehicle(llCrfCode As Long, ilCrfVefCode As Integer)
    Dim ilRet As Integer
    Dim ilCvf As Integer
    Dim ilVef As Integer
        
    If ilCrfVefCode > 0 Then
        ilVef = gBinarySearchVef(ilCrfVefCode)
        If ilVef = -1 Then
            ReDim imCrfVefCode(0 To 0) As Integer
            imCrfVefCode(0) = 0
        ElseIf tgMVef(ilVef).sType <> "P" Then
            ReDim imCrfVefCode(0 To 1) As Integer
            imCrfVefCode(0) = ilCrfVefCode
        Else
            ReDim imCrfVefCode(0 To 0) As Integer
            imCrfVefCode(0) = 0
        End If
        Exit Sub
    End If
    ReDim imCrfVefCode(0 To 0) As Integer
    imCvfRecLen = Len(tmCvf)
    tmCvfSrchKey1.lCode = llCrfCode
    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCvf.lCrfCode = llCrfCode)
        For ilCvf = 0 To 99 Step 1
            If tmCvf.iVefCode(ilCvf) > 0 Then
                ilVef = gBinarySearchVef(tmCvf.iVefCode(ilCvf))
                If ilVef <> -1 Then
                    If tgMVef(ilVef).sType <> "P" Then
                        imCrfVefCode(UBound(imCrfVefCode)) = tmCvf.iVefCode(ilCvf)
                        ReDim Preserve imCrfVefCode(0 To UBound(imCrfVefCode) + 1) As Integer
                    End If
                End If
            End If
        Next ilCvf
        ilRet = btrGetNext(hmCvf, tmCvf, imCvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
