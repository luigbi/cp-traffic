Attribute VB_Name = "EDSSubs"
'EDS Support Sub Routines
'Doug Smith 7/16/15
'Copyright 2015 Counterpoint Software, Inc. All rights reserved.
'Proprietary Software, Do not copy
Option Explicit
Option Compare Text

Public sgToken As String
Public sgGuiPDFDocumentID As String
Public sgNetworkName As String
Public bgEDSIsActive As Boolean
Public smUserAgent As String
Public smFirmId As String
Public smUserPswd As String

'INI URL Calls
Public smReqAuthURL As String
Public smRootURL As String
         
'Access Key obtained from EDS
Public smCurUserAPIKey As String

'User info
Public smCurUserFirstName As String
Public smCurUserLastName As String
Public imUserID As Integer
'Public smUserEmail As String
Public bmErrorsFound As Boolean
Private smUserName As String
Private smUserRights As String
Private smUserEmail As String

'Misc
Public lmLen As Long
Public lmErrorCnt As Long
Public bmDemoMode As Boolean
Public smMsg As String
Public hmPDF As Integer
Private smStationUserEmails() As String
Type VEHICLE_INFO
    sVehicleName As String * 40
    sVehicleNetIntCode As Integer
End Type

Public Function gSend_Post_URLs(sBody As String, sAPI As String) As Boolean
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    Dim slBody As String
    Dim slRet As String
    Dim slResponse As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilRetries As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    gSend_Post_URLs = False
    ilRet = gGetEDSAutorization
    If Not ilRet Then
        Exit Function
    End If
    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
        smRootURL = smRootURL & "DataSync/" & sAPI & sBody
    End If
    For ilRetries = 1 To 1 Step 1
        Set objHttp = New MSXML2.XMLHTTP60
        objHttp.Open "POST", smRootURL
        objHttp.setRequestHeader "Content-Type", "application/jsonrequest"
        objHttp.setRequestHeader "X-CCS-Gateway-Token", sgToken
        If bmDemoMode Then
            llReturn = 200
            slResponse = "Test Mode, Test Mode"
        Else
            objHttp.Send
            llReturn = objHttp.Status
            slResponse = objHttp.responseText
            sgToken = Trim$(objHttp.getResponseHeader("X-CCS-Gateway-Token"))
        End If
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn = 200 Then
            gLogMsg "gSend_Post_URLs - " & sAPI & " " & llReturn & " " & slResponse & " " & smRootURL, "EDS_Log.txt", False
            gSend_Post_URLs = True
            Exit For
        Else
            gLogMsg "gSend_Post_URLs - " & sAPI & " " & llReturn & " " & slResponse & " " & smRootURL, "EDS_ErrorLog.Txt", False
            MsgBox "gSend_Post_URLs - " & sAPI & " " & llReturn & " " & slResponse & " " & smRootURL
            lmErrorCnt = lmErrorCnt + 1
            Call gSleep(2)
        End If
    Next ilRetries
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gSend_Post_URLs "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDS_ErrorLog.Txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gSend_Post_APIs(sBody As String, sAPI As String) As Boolean
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    'Dim slBody As String
    Dim slRet As String
    Dim slResponse As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilRetries As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    gSend_Post_APIs = False
    
    ilRet = gGetEDSAutorization()
    If Not ilRet Then
        Exit Function
    End If
    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
        If sAPI = "SubmitApprovalRequest" Then
            smReqAuthURL = smRootURL & "EDS/" & sAPI
        Else
            smReqAuthURL = smRootURL & "DataSync/" & sAPI
        End If
    End If

    For ilRetries = 1 To 1 Step 1
        Set objHttp = New MSXML2.XMLHTTP60
        objHttp.Open "POST", smReqAuthURL
        objHttp.setRequestHeader "Content-Type", "application/json"
        objHttp.setRequestHeader "X-CCS-Gateway-Token", sgToken
        If bmDemoMode Then
            llReturn = 200
            slResponse = "Test Mode, Test Mode"
        Else
            objHttp.Send (sBody)
            llReturn = objHttp.Status
            slResponse = objHttp.responseText
            sgToken = Trim$(objHttp.getResponseHeader("X-CCS-Gateway-Token"))
            DoEvents
        End If
        
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn = 200 Then
            gLogMsg "gSend_Post_APIs - " & sAPI & " " & llReturn & " " & slResponse & " " & sBody, "EDS_Log.txt", False
            gSend_Post_APIs = True
            Exit For
        Else
            gLogMsg "gSend_Post_APIs - " & sAPI & " " & llReturn & " " & slResponse & " " & sBody, "EDS_ErrorLog.Txt", False
            MsgBox "gSend_Post_APIs - " & sAPI & " " & llReturn & " " & slResponse & " " & sBody
            lmErrorCnt = lmErrorCnt + 1
            Call gSleep(2)
        End If
    Next ilRetries
    Exit Function
    MsgBox "Submit Approval Complete", vbOK
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gSend_Post_APIs "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDS_ErrorLog.Txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gGetEDSAutorization() As Boolean
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    Dim sBody As String
    Dim slRet As String
    Dim slResponse As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilRetries As Integer
    Dim ilRet As Integer
    Dim slIntegratorID As String
    Dim slIntegratorKey As String
    
    On Error GoTo ErrHand
    gGetEDSAutorization = False
    If igTestSystem Then
        Exit Function
    End If
    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
        smReqAuthURL = smRootURL & "Integrator/Authenticate"
    Else
        MsgBox "[EDS-API] - EDS_RootURL definition not found in the Traffic.ini file"
        Exit Function
    End If
    
    If Not gLoadOption("EDS-API", "EDS_IntegratorID", slIntegratorID) Then
        MsgBox "[EDS-API] - EDS_IntegratorID definition not found in the Traffic.ini file"
        Exit Function
    End If
    
    If Not gLoadOption("EDS-API", "EDS_IntegratorKey", slIntegratorKey) Then
        MsgBox "[EDS-API] - EDS_IntegratorKey definition not found in the Traffic.ini file"
        Exit Function
    End If
    
    sBody = "{" & """" & "IntegratorID" & """" & " : " & """" & slIntegratorID & """" & "," & """" & "IntegratorKey" & """" & " : " & """" & slIntegratorKey & """" & "}"
    For ilRetries = 1 To 3 Step 1
        Set objHttp = New MSXML2.XMLHTTP60
        objHttp.Open "POST", smReqAuthURL
        objHttp.setRequestHeader "Content-Type", "application/json"
        If bmDemoMode Then
            llReturn = 200
            slResponse = "Test Mode, Test Mode"
        Else
            objHttp.Send (sBody)
            llReturn = objHttp.Status
            slResponse = objHttp.responseText
            sgToken = Trim$(objHttp.getResponseHeader("X-CCS-Gateway-Token"))
        End If
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn = 200 Then
            'gLogMsg "gGetEDSAutorization - " & llReturn & " " & sBody, "EDS_Log.txt", False
            gGetEDSAutorization = True
            Exit For
        Else
            If ilRetries = 3 Then
                gLogMsg "gGetEDSAutorization - " & llReturn & " " & slResponse & " " & sBody, "EDS_ErrorLog.Txt", False
                MsgBox "gGetEDSAutorization - " & llReturn & " " & sBody
            End If
            lmErrorCnt = lmErrorCnt + 1
            'Call Sleep(500)
            Call Sleep(3000)
        End If
    Next ilRetries
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gGetEDSAutorization "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDS_ErrorLog.Txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function ChangeNetworkName(sOldNetwork As String, sNewNetworkName As String) As Boolean
    
    Dim slStr As String
    
    On Error GoTo ErrHand
    ChangeNetworkName = False
    
    ChangeNetworkName = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function


Public Function gLinkStationToNetwork(sCallLetters As String, sClientName As String) As Boolean

    Dim blRet As Boolean
    Dim slBody As String
    Dim slTemp As String
    
    On Error GoTo ErrHand
    gLinkStationToNetwork = False
    slTemp = Replace(Trim$(sClientName), " ", "%20")
    slBody = "?" & "stationName" & "=" & Trim$(sCallLetters) & "&" & "networkName" & "=" & slTemp
    blRet = gSend_Post_URLs(slBody, "LinkStationToNetwork")
    If blRet Then
        gLinkStationToNetwork = True
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gAddOrUpdateSingleNetworkUser(mURF_Code As Integer) As Boolean

    Dim ilTemp As Integer
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slSQLQuery As String
    Dim slBody As String
    Dim rst_Temp As ADODB.Recordset
    'Items passed to EDS web
    Dim slUserNameEmail  As String 'username (same as their email address in this project)
    Dim slUserRights As String  'not implemented yet, pass an empty list
    Dim blIsActive As Boolean
    Dim slFullName As String
    Dim slPassword As String
    Dim slNetworkName As String
    Dim ilMinOffset As Integer
    Dim slID As String

    On Error GoTo ErrHand
    gAddOrUpdateSingleNetworkUser = False
    ilRet = gObtainUrf()
    On Error GoTo ErrHand
    ilRet = gObtainUrf()
    ilTemp = 0
    blRet = gGetEDSAutorization()
    If blRet Then
        'If tgPopUrf(mURF_Code).sDelete <> "Y" And tgPopUrf(mURF_Code).lEMailCefCode > 0 Then
        If tgPopUrf(mURF_Code).sDelete <> "Y" And tgPopUrf(mURF_Code).lEMailCefCode > 0 Then
            slSQLQuery = "Select cefComment from CEF_Comments_Events where cefCode = " & tgPopUrf(mURF_Code).lEMailCefCode
            'Set rst_Temp = cnn.Execute(slSQLQuery)
            Set rst_Temp = gSQLSelectCall(slSQLQuery)
            If Not rst_Temp.EOF Then
                slUserNameEmail = Trim$(rst_Temp!cefComment)
            Else
                slUserNameEmail = "UnDefined"
            End If
            slUserRights = "V"
            blIsActive = 1
            slFullName = Trim$(tgPopUrf(mURF_Code).sName)
            slPassword = Trim$(tgPopUrf(mURF_Code).sPassword)
            slNetworkName = Trim$(tgSpf.sGClient)
            ilMinOffset = mGetUTCMinutesOffset("")
            slID = "URF" & Format(mURF_Code, String(5, "0"))
            slBody = "{" & """" & "UsernameEmail" & """" & ":" & """" & slUserNameEmail & """" & ","
            slBody = slBody & """" & "UserRights" & """" & ":[" & "4" & "],"
            slBody = slBody & """" & "IsActive" & """" & ":" & "true" & ","
            slBody = slBody & """" & "FullName" & """" & ":" & """" & slFullName & """" & ","
            slBody = slBody & """" & "Password" & """" & ":" & """" & slPassword & """" & ","
            slBody = slBody & """" & "UTCMinutesOffset" & """" & ":" & ilMinOffset & ","
            'slBody = slBody & """" & "ID" & """" & ":" & slID & ","
            slBody = slBody & """" & "NetworkName" & """" & ":" & """" & slNetworkName & """" & "}"
            blRet = gSend_Post_APIs(slBody, "AddOrUpdateNetworkUser")
            rst_Temp.Close
        End If
    Else
        gAddOrUpdateSingleNetworkUser = False
        Exit Function
    End If
    gAddOrUpdateSingleNetworkUser = True
    Exit Function
ErrHand:
    Resume Next
    rst_Temp.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gReadInPDF(sFromFileName As String) As Boolean

    Dim hmPDF As Integer
    Dim llLen As Long
    Dim i As Long
    
    Dim objHttp As MSXML2.XMLHTTP60
    Dim llReturn As Long
    Dim slBody As String
    Dim slRet As String
    Dim slResponse As String
    Dim slRetStr As String
    Dim slTemp As String
    Dim ilRetries As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    'Dim slFileName As String
    Dim slRetFilename As String
    Dim slRetID As String
    Dim slNetworkName As String
    Dim slStr As String
    Dim llIdx As Long
    Dim slBoundary As String
    Dim llCnt As Long
       
    On Error GoTo ErrHand
    'ilPos = InStr(1, sFromFileName, "\")
    'slFileName = right$(sFromFileName, Len(sFromFileName) - ilPos)
    gReadInPDF = True
    slNetworkName = Trim$(tgSpf.sGClient)
    slBoundary = vbCrLf & "--AaB03x" & vbCrLf
    slBody = slBoundary
    slBody = slBody & sFromFileName
    slBody = slBody & slBoundary
    slBody = slBody & "Content-Disposition: form-data; name=""file"";" & " filename" & "=" & """" & sFromFileName & """" & vbCrLf
    slBody = slBody & "Content-Type: application/pdf" & vbCrLf & vbCrLf
    
    Dim blFileBytes() As Byte
    ReDim blByteheader(0 To Len(slBody) - 1) As Byte
    blFileBytes = mReadFile(sFromFileName)
    blByteheader = StrConv(slBody, vbFromUnicode)
    ReDim blPostByte(0 To UBound(blFileBytes) + UBound(blByteheader) + 1) As Byte
    For i = 0 To UBound(blByteheader)
        blPostByte(i) = blByteheader(i)
    Next i
    llCnt = UBound(blByteheader) + 1
    For i = 0 To UBound(blFileBytes)
        blPostByte(i + llCnt) = blFileBytes(i)
    Next i
    slBody = vbCrLf & "--AaB03x--" & vbCrLf
    ReDim blByteFooter(0 To Len(slBody) - 1) As Byte
    blByteFooter = StrConv(slBody, vbFromUnicode)
    llCnt = llCnt + UBound(blFileBytes) + 1
    ReDim Preserve blPostByte(0 To UBound(blFileBytes) + UBound(blByteheader) + UBound(blByteFooter) + 2) As Byte
    For i = 0 To UBound(blByteFooter)
        blPostByte(i + llCnt) = blByteFooter(i)
    Next i
    llLen = UBound(blPostByte) + 1
    'ilRet = gGetEDSAutorization()
    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
        smReqAuthURL = smRootURL & "File/UploadPDF"
    End If
    For ilRetries = 1 To 1 Step 1
        Set objHttp = New MSXML2.XMLHTTP60
        objHttp.Open "POST", smReqAuthURL
        objHttp.setRequestHeader "Content-Type", "multipart/form-data;boundary= " & "AaB03x"
        objHttp.setRequestHeader "Content-Length", llLen
        objHttp.setRequestHeader "X-CCS-Gateway-Token", sgToken
        objHttp.setRequestHeader "X-CCS-Gateway-Network", Trim$(slNetworkName)
        objHttp.setRequestHeader "X-CCS-Gateway-Filename", Trim$(sFromFileName)
        If bmDemoMode Then
            llReturn = 200
            slResponse = "Test Mode, Test Mode"
        Else
            objHttp.Send (blPostByte)
            llReturn = objHttp.Status
            slResponse = objHttp.responseText
            sgToken = Trim$(objHttp.getResponseHeader("X-CCS-Gateway-Token"))
        End If
       
        slTemp = Mid(slResponse, 9, Len(slResponse))
        ilPos = InStr(slTemp, ",")
        'GUI ID for submit approval requests
        sgGuiPDFDocumentID = Mid(slTemp, 1, ilPos - 2)
        Set objHttp = Nothing
        'Anything but 200 is an error.
        If llReturn = 200 Then
            gLogMsg "gReadInPDF - " & llReturn & " " & slResponse & " ", "EDS_Log.txt", False
            gReadInPDF = True
            Exit For
        Else
            gLogMsg "gReadInPDF - " & llReturn & " " & slResponse & " ", "EDS_ErrorLog.Txt", False
            MsgBox "gReadInPDF -  " & llReturn & " " & slResponse & " "
            lmErrorCnt = lmErrorCnt + 1
        End If
    Next ilRetries
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function mReadFile(sFileName As String) As Byte()

    Dim llLen As Long
    Dim llIdx As Long
    Dim slStr As String
    Dim ilRet As Integer

    'On Error GoTo ErrHand

    'hmPDF = FreeFile
    'Open sgExportPath & sFileName For Binary As #hmPDF
    ilRet = gFileOpen(sgExportPath & sFileName, "Binary", hmPDF)
    If ilRet <> 0 Then
        ReDim blbyte(0 To 0) As Byte
        mReadFile = blbyte
        Exit Function
    End If
    llLen = LOF(hmPDF)
    ReDim blbyte(0 To llLen) As Byte
    blbyte = InputB(llLen, #hmPDF)
    Close hmPDF
    mReadFile = blbyte

    Exit Function
'ErrHand:
'    Screen.MousePointer = vbDefault
'    On Error GoTo 0
End Function
    
    

Public Function gSubmitApprovalRequest(tmEmail_Info() As EMAILINFO) As Boolean

    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilCnt As Integer
    Dim slTemp As String
    Dim ilTemp As Integer
    Dim slBody As String

    Dim blRet As Boolean
    Dim rst_Temp As ADODB.Recordset
    Dim rst_artt As ADODB.Recordset
    'API vars
    Dim slFromNetworkName As String
    Dim slNetworkUserEmail As String
    Dim slToStationName As String
    Dim slTransactionType As String
    Dim slContractNumber As String
    Dim slEstimateNumber As String
    Dim slAdvertiserName  As String
    Dim slProductName As String
    Dim slAgencyName As String
    Dim slTimeZone As String
    Dim slActualZone As String
    Dim slRespondByDate As String
    Dim slPDFDocumentID As String
    Dim slEmailContent As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slMasterEmail As String
    Dim blIsMaster As Boolean
    Dim slMissingStations As String
    Dim slMissingEmail As String
    Dim ilVef As Integer
    Dim ilPos As Integer
    Dim blError As Boolean
    Dim ilExtRevNo As Integer
    
    On Error GoTo ErrHand
    slActualZone = "EST"
    slTimeZone = "TimeZone"
    Screen.MousePointer = vbHourglass
    slMissingStations = ""
    slMissingEmail = ""
    blError = False
    ilExtRevNo = -1
    'ilRet = gObtainUrf()
    blRet = gGetEDSAutorization()
    If Not blRet Then
        gSubmitApprovalRequest = False
        Exit Function
    End If
    'First we have to read in the PDF and send it to get back the response containing the GuiPDFDocumentID required for the
    'Submit Approval Request
    For ilCnt = LBound(tmEmail_Info) To UBound(tmEmail_Info) - 1 Step 1
        If ilCnt <> LBound(tmEmail_Info) Then
            If tmEmail_Info(ilCnt).lChfCode <> tmEmail_Info(ilCnt - 1).lChfCode Then
                'If (Not blError) Then
                    slTemp = "UPDATE chf_Contract_Header SET "
                    slTemp = slTemp & "chfEDSSentDate = " & "'" & Format(gNow(), sgSQLDateForm) & "'" & ", "
                    slTemp = slTemp & "chfEDSSentTime = " & "'" & Format(gNow(), sgSQLTimeForm) & "'" & ", "
                    slTemp = slTemp & "chfEDSSentUrfCode = " & tgUrf(0).iCode & ", "
                    slTemp = slTemp & "chfEDSSentExtRevNo = " & ilExtRevNo
                    slTemp = slTemp & " WHERE chfCode = " & tmEmail_Info(ilCnt - 1).lChfCode
                    'Set rst_Temp = cnn.Execute(slTemp)
                    If gSQLWaitNoMsgBox(slTemp, False) <> 0 Then
                        gHandleError "TrafficErrors.txt", "EDSSubs: gSubmitApprovalRequest"
                    End If
                'End If
                blError = False
                ilExtRevNo = -1
            End If
        End If
        If ilExtRevNo = -1 Then
            slTemp = "Select chfExtRevNo from chf_Contract_Header where chfCode = " & tmEmail_Info(ilCnt).lChfCode
            'Set rst_Temp = cnn.Execute(slTemp)
            Set rst_Temp = gSQLSelectCall(slTemp)
            If Not rst_Temp.EOF Then
                ilExtRevNo = Trim$(rst_Temp!chfExtRevNo)
            Else
                ilExtRevNo = 0
            End If
        End If
        slTemp = Trim$(tmEmail_Info(ilCnt).sPDFFileName) & ".pdf"
        blRet = gReadInPDF(slTemp)
        gSubmitApprovalRequest = False
        'user must have an email address
        If tgUrf(0).lEMailCefCode = 0 Then
            Exit Function
        Else
            slTemp = "Select cefComment from CEF_Comments_Events where cefCode = " & tgUrf(0).lEMailCefCode
            'Set rst_Temp = cnn.Execute(slTemp)
            Set rst_Temp = gSQLSelectCall(slTemp)
            If Not rst_Temp.EOF Then
                slNetworkUserEmail = Trim$(rst_Temp!cefComment)
            Else
                slNetworkUserEmail = ""
            End If
        End If
        ReDim smStationUserEmails(0 To 0) As String
        slFromNetworkName = Trim$(tgSpf.sGClient)
        ilTemp = gGetStationFromVehicle(tmEmail_Info(ilCnt).iVefCode)
        
        If ilTemp <> 0 Then
'            blIsMaster = mIsMasterStation(ilTemp)
            slToStationName = gGetCallLettersByShttCode(ilTemp)
'            If blIsMaster Then
'                SQLQuery = "Select * from artt where arttShttCode = " & ilTemp & " And ArttType = " & " '" & "P" & "'" & " And ArttState = 0 "
'                SQLQuery = SQLQuery & " and arttEmailRights In ('M', 'A', 'V')"
'                Set rst_artt = cnn.Execute(SQLQuery)
'                ilIdx = 0
'                While Not rst_artt.EOF
'                    If Len(Trim(rst_artt!arttEmail)) > 0 Then
'                        smStationUserEmails(ilIdx) = Trim$(rst_artt!arttEmail)
'                        ilIdx = ilIdx + 1
'                        ReDim Preserve smStationUserEmails(0 To ilIdx)
'                    End If
'                    rst_artt.MoveNext
'                Wend
'            Else
'                slMasterEmail = mGetMasterEmail(ilTemp)
'            End If
            'If UBound(smStationUserEmails) > 0 Then
            blRet = mGetPersonnel(ilTemp)
            If blRet Then
                slTransactionType = "I"
                slContractNumber = tmEmail_Info(ilCnt).lCntrNo
                slEstimateNumber = tmEmail_Info(ilCnt).sAgyEstNo
                slAdvertiserName = tmEmail_Info(ilCnt).sAdvtName
                slProductName = tmEmail_Info(ilCnt).sProduct
                slAgencyName = " "
                slRespondByDate = Format(tmEmail_Info(ilCnt).sResponseDate, "mm/dd/yyyy")
                slPDFDocumentID = sgGuiPDFDocumentID
                SQLQuery = "Select * from emf_Email_Content where emfCode = " & tmEmail_Info(ilCnt).lEmfCode
                'Set rst_Temp = cnn.Execute(SQLQuery)
                Set rst_Temp = gSQLSelectCall(SQLQuery)
                If Not rst_Temp.EOF Then
                    slEmailContent = Trim$(rst_Temp!emfContent)
                End If
                slStartDate = Format(tmEmail_Info(ilCnt).sStartDate, "mm/dd/yyyy")
                slEndDate = Format(tmEmail_Info(ilCnt).sEndDate, "mm/dd/yyyy")
                slBody = ""
                slBody = "{""FromNetworkName""" & ":" & """" & slFromNetworkName & ""","
                slBody = slBody & """NetworkUserEmail""" & ":" & """" & slNetworkUserEmail & ""","
                slBody = slBody & """ToStationName""" & ":" & """" & slToStationName & ""","
                slBody = slBody & smUserEmail
                slBody = slBody & smUserName
                slBody = slBody & smUserRights

                'slBody = slBody & """StationUserEmails""" & ":["
                'slTemp = ""
                'For ilIdx = 0 To UBound(smStationUserEmails) - 1 Step 1
                '    slTemp = slTemp & """" & Trim$(smStationUserEmails(ilIdx)) & """"
                '    If ilIdx < UBound(smStationUserEmails) - 1 Then
                '        slTemp = slTemp & ","
                '    End If
                'Next ilIdx
                'slBody = slBody & slTemp & "],"
                
                slBody = slBody & """TransactionType""" & ":" & """" & Trim$(slTransactionType) & ""","
                slBody = slBody & """ContractNumber""" & ":" & """" & Trim$(slContractNumber) & ""","
                slBody = slBody & """EstimateNumber""" & ":" & """" & Trim$(slEstimateNumber) & ""","
                slBody = slBody & """AdvertiserName""" & ":" & """" & Trim$(slAdvertiserName) & ""","
                slBody = slBody & """ProductName""" & ":" & """" & Trim$(slProductName) & ""","
                slBody = slBody & """AgencyName""" & ":" & """" & Trim$(slAgencyName & "") & ""","
                slBody = slBody & """RespondByDate""" & ":" & """" & Trim$(slRespondByDate) & ""","
                slBody = slBody & """PDFDocumentID""" & ":" & """" & Trim$(slPDFDocumentID) & ""","
                
                slEmailContent = slEmailContent & "<BR/>" & "<BR/>" & "Please click on the following link to view the Insertion Order:" & "<BR/>"
                slEmailContent = slEmailContent & "[EDSLINK]" & "<BR/>" & "<BR/>"
                slEmailContent = slEmailContent & "[NEWPASSWORD]" & "<BR/>" & "<BR/>"
                slEmailContent = slEmailContent & "Contract Number: " & Trim$(slContractNumber) & "<BR/>"
                slEmailContent = slEmailContent & "Estimate Number: " & Trim$(slEstimateNumber) & "<BR/>"
                slEmailContent = slEmailContent & "Advertiser Name: " & Trim$(slAdvertiserName) & "<BR/>"
                slEmailContent = slEmailContent & "Product Name: " & Trim$(slProductName) & "<BR/>" & "<BR/>"
                slEmailContent = slEmailContent & "Respond by Date: " & Trim$(slRespondByDate) & "<BR/>"
                'slEmailContent = slEmailContent & "Flight Date: " & Trim$("") & vbCrLf
                slEmailContent = slEmailContent & "<BR/>" & "<BR/>"
                
                slBody = slBody & """EmailContent""" & ":" & """" & Trim$(slEmailContent) & ""","
                slBody = slBody & """StartDate""" & ":" & """" & Trim$(slStartDate) & ""","
                slBody = slBody & """EndDate""" & ":" & """" & Trim$(slEndDate) & ""","
                'slBody = slBody & """TimeZone""" & ":" & """EST""}"
                slBody = slBody & """" & slTimeZone & """" & ":" & """" & slActualZone & """" & "}"
                blRet = gSend_Post_APIs(slBody, "SubmitApprovalRequest")
                If blRet Then
                    gSubmitApprovalRequest = True
                    slTemp = "UPDATE clf_Contract_Line SET "
                    slTemp = slTemp & "clfEDSSentExtRevNo = " & ilExtRevNo
                    slTemp = slTemp & " WHERE clfchfCode = " & tmEmail_Info(ilCnt).lChfCode
                    slTemp = slTemp & " AND clfVefCode = " & tmEmail_Info(ilCnt).iVefCode
                    'Set rst_Temp = cnn.Execute(slTemp)
                    If gSQLWaitNoMsgBox(slTemp, False) <> 0 Then
                        gHandleError "TrafficErrors.txt", "EDSSubs: gSubmitApprovalRequest"
                    End If
                Else
                    blError = False
                End If
            Else
                blError = False
                ilVef = gBinarySearchVef(tmEmail_Info(ilCnt).iVefCode)
                If ilVef <> -1 Then
                    If slMissingEmail = "" Then
                        slMissingEmail = Trim(tgMVef(ilVef).sName)
                    Else
                        If InStr(1, slMissingEmail, Trim(tgMVef(ilVef).sName), vbTextCompare) <= 0 Then
                            slMissingEmail = slMissingEmail & "," & Trim(tgMVef(ilVef).sName)
                        End If
                    End If
                Else
                    If slMissingEmail = "" Then
                        slMissingEmail = "vef code:" & tmEmail_Info(ilCnt).iVefCode
                    Else
                        If InStr(1, slMissingEmail, "vef code:" & tmEmail_Info(ilCnt).iVefCode, vbTextCompare) <= 0 Then
                            slMissingEmail = slMissingEmail & ", " & "vef code:" & tmEmail_Info(ilCnt).iVefCode
                        End If
                    End If
                End If
            End If
        Else
            blError = False
            ilVef = gBinarySearchVef(tmEmail_Info(ilCnt).iVefCode)
            If ilVef <> -1 Then
                If slMissingStations = "" Then
                    slMissingStations = Trim(tgMVef(ilVef).sName)
                Else
                    If InStr(1, slMissingStations, Trim(tgMVef(ilVef).sName), vbTextCompare) <= 0 Then
                        slMissingStations = slMissingStations & "," & Trim(tgMVef(ilVef).sName)
                    End If
                End If
            Else
                If slMissingStations = "" Then
                    slMissingStations = "vef code:" & tmEmail_Info(ilCnt).iVefCode
                Else
                    If InStr(1, slMissingStations, "vef code:" & tmEmail_Info(ilCnt).iVefCode, vbTextCompare) <= 0 Then
                        slMissingStations = slMissingStations & ", " & "vef code:" & tmEmail_Info(ilCnt).iVefCode
                    End If
                End If
            End If
        End If
    Next ilCnt
    'If (Not blError) Then
        slTemp = "UPDATE chf_Contract_Header SET "
        slTemp = slTemp & "chfEDSSentDate = " & "'" & Format(gNow(), sgSQLDateForm) & "'" & ", "
        slTemp = slTemp & "chfEDSSentTime = " & "'" & Format(gNow(), sgSQLTimeForm) & "'" & ", "
        slTemp = slTemp & "chfEDSSentUrfCode = " & tgUrf(0).iCode & ", "
        slTemp = slTemp & "chfEDSSentExtRevNo = " & ilExtRevNo
        slTemp = slTemp & " WHERE chfCode = " & tmEmail_Info(UBound(tmEmail_Info) - 1).lChfCode
        'Set rst_Temp = cnn.Execute(slTemp)
        If gSQLWaitNoMsgBox(slTemp, False) <> 0 Then
            gHandleError "TrafficErrors.txt", "EDSSubs: gSubmitApprovalRequest"
        End If
    'End If
    If slMissingEmail <> "" Then
        MsgBox "The following stations are missings emails or email rights are not defined: " & slMissingEmail & " Emails will not be sent to the stations listed."
    End If
    If slMissingStations <> "" Then
        MsgBox "The following vehicles are missings stations: " & slMissingStations & " Emails will not be sent to the stations listed."
    End If
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function SetAllUserStatusesAtStation(sStationName As String, bIsActive As Boolean) As Boolean
    
    Dim slBody As String
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    SetAllUserStatusesAtStation = False
    slBody = "{" & """" & Trim$(sStationName) & """" & ":" & """" & Trim$(bIsActive) & """" & "}"
    blRet = gSend_Post_APIs(slBody, "SetAllUserStatusesAtStation")
    If blRet Then
        SetAllUserStatusesAtStation = True
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gChangeNetworkName(sOldName As String, sNewName As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    gChangeNetworkName = False
    slBody = "{" & """" & "oldName" & """" & ":" & """" & Trim$(sOldName) & """" & "," & """" & "newName" & """" & ":" & """" & Trim$(sNewName) & """" & "}"
    blRet = gSend_Post_APIs(slBody, "ChangeNetworkName")
    If blRet Then
        gChangeNetworkName = True
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gChangeStationName(sOldName As String, sNewName As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    gChangeStationName = False
    slBody = "{" & """" & Trim$(sOldName) & """" & ":" & """" & Trim$(sNewName) & """" & "}"
    blRet = gSend_Post_APIs(slBody, "ChangeStationName")
    If blRet Then
        gChangeStationName = True
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gChangeUsernameEmail(sOldName As String, sNewName As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    
'    On Error GoTo ErrHand
'    gChangeUsernameEmail = False
'    slBody = "?" & "oldUsername" & "=" & Trim$(sOldName) & "&" & "newUsername" & "=" & Trim$(sNewName)
'    blRet = gSend_Post_URLs(slBody, "ChangeUsernameEmail")
'    If blRet Then
'        gChangeUsernameEmail = True
'    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function ChangeUserPassword(sUserName As String, sNewPswd As String) As Boolean
    
    Dim slBody As String
    Dim blRet As Boolean
    Dim slUserNameEmail As String
    Dim rst_Temp As ADODB.Recordset
    Dim ilUrf As Integer
    Dim ilLoop As Integer
    Dim ilURF_Code  As Integer
    Dim slSQLQuery As String
    
    ChangeUserPassword = False
    'For ilUrf = 1 To UBound(tgPopUrf) - 1 Step 1
    For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
        If Trim(tgPopUrf(ilUrf).sName) = Trim(sUserName) Then
            ilURF_Code = tgPopUrf(ilUrf).iAutoCode
            Exit For
        End If
    Next ilUrf
    slSQLQuery = "Select cefComment from CEF_Comments_Events where cefCode = " & tgPopUrf(ilURF_Code).lEMailCefCode
    'Set rst_Temp = cnn.Execute(slSQLQuery)
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        slUserNameEmail = Trim$(rst_Temp!cefComment)
    Else
        slUserNameEmail = "UnDefined"
    End If
    slBody = "?" & "username" & "=" & Trim$(slUserNameEmail) & "&" & "newPassword" & "=" & Trim$(sNewPswd)
    blRet = gSend_Post_URLs(slBody, "ChangeUserPassword")
    ChangeUserPassword = True
    rst_Temp.Close
End Function

Public Function gRemoveStationFromNetwork(sStationName As String, sNetworkName As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    gRemoveStationFromNetwork = False
    slBody = "{" & """" & Trim$(sStationName) & """" & ":" & """" & Trim$(sNetworkName) & """" & "}"
    blRet = gSend_Post_APIs(slBody, "gRemoveStationFromNetwork")
    If blRet Then
        gRemoveStationFromNetwork = True
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gUpdateVehicle(iVefCode As Integer) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    Dim rst_Shtt As ADODB.Recordset
    Dim rst_vef As ADODB.Recordset
    Dim slTemp As String
    
    On Error GoTo ErrHand
    SQLQuery = "select * from VEF_Vehicles where vefcode = " & iVefCode
    'Set rst_vef = cnn.Execute(SQLQuery)
    Set rst_vef = gSQLSelectCall(SQLQuery)
    'Loop through all of the vehicles
    If Not rst_vef.EOF Then
        'Do we have a station with a name that matches the vehicle name?
        slTemp = "Select * from shtt where shttCallLetters = " & "'" & Trim$(rst_vef!VEFNAME) & "'" & "And shttType = 0"
        'Set rst_Shtt = cnn.Execute(slTemp)
        Set rst_Shtt = gSQLSelectCall(slTemp)
        If Not rst_Shtt.EOF Then
            slBody = "{" & """" & "Name" & """" & ":" & """" & Trim$(rst_Shtt!shttCallLetters) & """" & "}"
            blRet = gSend_Post_APIs(slBody, "AddOrUpdateStation")
            sgNetworkName = Trim$(tgSpf.sGClient)
            blRet = gLinkStationToNetwork(rst_Shtt!shttCallLetters, sgNetworkName)
        End If
    End If
    On Error Resume Next
    rst_Shtt.Close
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

'Public Function gLinkstations() As Boolean
'
'    Dim slTemp As String
'    Dim slTemp2 As String
'    Dim slBody As String
'    Dim blRet As Boolean
'    Dim rst_Master As ADODB.Recordset
'    Dim rst_Sister As ADODB.Recordset
'    Dim rst_EmailRights As ADODB.Recordset
'    Dim blFound As Boolean
'
'    gLinkstations = False
'    'Find all of the master stations
'    SQLQuery = "select shttCode, shttCallLetters, shttClusterGroupID, shttMAsterCluster from shtt where shttMasterCluster = " & "'" & "Y" & "'"
'    Set rst_Master = cnn.Execute(SQLQuery)
'    While Not rst_Master.EOF
'        'debug
'        'If rst_Master!shttClusterGroupID = 20 Then
'        '    blFound = blFound
'        'End If
'        'Now find the master station's sister stations
'        slTemp = "select shttCallLetters, shttCode from shtt where shttClusterGroupID = " & rst_Master!shttClusterGroupID & " and shttMasterCluster <> " & "'" & "Y" & "'"
'        Set rst_Sister = cnn.Execute(slTemp)
'        slBody = "{" & """" & "MasterStationName" & """" & ":" & """" & Trim(rst_Master!shttCallLetters) & """," & """" & "StationNamesInGroup" & """" & ":" & "["
'        While Not rst_Sister.EOF
'            blFound = False
'            slTemp2 = "select arttEmailRights from artt where arttShttCode = " & rst_Master!shttCode
'            slTemp2 = slTemp2 & " and arttEmailrights in ('M','A','V')"
'            Set rst_EmailRights = cnn.Execute(slTemp2)
'            If Not rst_EmailRights.EOF Then
'                If Trim(rst_EmailRights!arttEmailRights) <> "" And rst_EmailRights!arttEmailRights <> "N" Then
'                    slBody = slBody & """" & Trim(rst_Sister!shttCallLetters) & """"
'                    blFound = True
'                End If
'            End If
'            rst_Sister.MoveNext
'            If Not rst_Sister.EOF And blFound Then
'                slBody = slBody & ","
'            End If
'        Wend
'        While right(slBody, 1) = ","
'            slBody = Left(slBody, Len(slBody) - 1)
'        Wend
'        slBody = slBody & "]}"
'        blRet = gSend_Post_APIs(slBody, "AssociateStationsToMasterGroup")
'        rst_Master.MoveNext
'    Wend
'    gLinkstations = True
'    On Error Resume Next
'    rst_Master.Close
'    rst_Sister.Close
'    rst_EmailRights.Close
'    Exit Function
'ErrHand:
'    On Error Resume Next
'    rst_Master.Close
'    rst_Sister.Close
'    rst_EmailRights.Close
'    Screen.MousePointer = vbDefault
'    On Error GoTo 0
'End Function

Private Function mEDSWebErrors(mRoutine As String, mErrorCode As Integer) As String

    Dim slErrorText As String

    mEDSWebErrors = ""
    Select Case mErrorCode
        Case 499
            slErrorText = "Bad Integrator ID"
        Case 498
            slErrorText = "Bad Integrator Key"
        Case 497
            slErrorText = "Bad Integrator Token"
        Case 496
            slErrorText = "Integrator Not Active"
        Case 495
            slErrorText = "Integrator Not Authorized"
        Case 494
            slErrorText = "Bad UserName"
        Case 493
            slErrorText = "Bad Password"
        Case 492
            slErrorText = "Bad User Token"
        Case 491
            slErrorText = "User Not Active"
        Case 490
            slErrorText = "User Not Authorized"
        Case 489
            slErrorText = "Organization Not Found"
        Case 488
            slErrorText = "Network Name Already Exists"
        Case 487
            slErrorText = "Add/Update User Error"
        Case 486
            slErrorText = "UserName Not Found"
        Case 485
            slErrorText = "Update Username Error"
        Case 484
            slErrorText = "Bad Entity Relationship"
        Case 483
            slErrorText = "User type was found invalid (operation is valid for a network user, not a station user, etc.)"
        Case 482
            slErrorText = "Network Not Found"
        Case 481
            slErrorText = "User account entity can't be found in the database"
        Case 480
            slErrorText = "Organization type was found invalid (network was used when it should've been a station)"
        Case 479
            slErrorText = "Approval Request Recipient is Invalid"
        Case 478
            slErrorText = "An error occurred in regards to a PDF document"
        Case Else
            slErrorText = "Unknown Error code was returned from the EDS Web Site"
    End Select
    gLogMsg mRoutine & ": " & mErrorCode & " - " & slErrorText, "EDS_ErrorLog.Txt", False
End Function

'Public Function mIsMasterStation(ilShttCode) As Boolean
'
'    Dim rst_Shtt As ADODB.Recordset
'    Dim slTemp As String
'
'    mIsMasterStation = False
'    slTemp = "Select shttClusterGroupID, shttMasterCluster from shtt where shttCode = " & ilShttCode
'    Set rst_Shtt = cnn.Execute(slTemp)
'    If Not rst_Shtt.EOF Then
'        If Trim(rst_Shtt!shttMasterCluster) = "Y" Or rst_Shtt!shttClusterGroupID = 0 Then
'            mIsMasterStation = True
'        End If
'    End If
'End Function
'
'
'Private Function mGetMasterEmail(iShttCode As Integer) As String
'
'    Dim rst_Shtt As ADODB.Recordset
'    Dim rst_artt As ADODB.Recordset
'    Dim ilShttGroupID As Integer
'    Dim slTemp As String
'    Dim ilShttCode As String
'    Dim ilIdx As Integer
'
'    mGetMasterEmail = ""
'    slTemp = "Select shttCode, shttMasterCluster, shttClusterGroupID from shtt where shttCode = " & iShttCode
'    Set rst_Shtt = cnn.Execute(slTemp)
'    If Not rst_Shtt.EOF Then
'        ilShttGroupID = rst_Shtt!shttClusterGroupID
'    End If
'
'    ilShttCode = iShttCode
'    If ilShttGroupID > 0 Then
'        slTemp = "Select * from shtt where shttClusterGroupID = " & ilShttGroupID & " and shttMasterCluster = " & "'" & "Y" & "'"
'        Set rst_Shtt = cnn.Execute(slTemp)
'    End If
'
'    SQLQuery = "Select * from artt where arttShttCode = " & rst_Shtt!shttCode
'    SQLQuery = SQLQuery & " and arttEmailRights In ('M', 'A', 'V')"
'    Set rst_artt = cnn.Execute(SQLQuery)
'    slTemp = ""
'    ilIdx = 0
'    While Not rst_artt.EOF
'        If Len(Trim(rst_artt!arttEmail)) > 0 Then
'            smStationUserEmails(ilIdx) = Trim(rst_artt!arttEmail)
'            ilIdx = ilIdx + 1
'            ReDim Preserve smStationUserEmails(0 To ilIdx)
'        End If
'        rst_artt.MoveNext
'    Wend
'    mGetMasterEmail = slTemp
'    rst_Shtt.Close
'    rst_artt.Close
'    Exit Function
'
'End Function



'Open the default browser
'ShellExecute 0, vbNullString, "http://www.sony.com/", vbNullString, vbNullString, vbNormalFocus

'*** Open the default emailer ***
'ShellExecute 0, vbNullString, "mailto:dicklevine@counterpoint.net", vbNullString, vbNullString, vbNormalFocus

'*** Open a document with the default document viewer ***
'ShellExecute 0, vbNullString, """C:\house.docx""", vbNullString, vbNullString, vbNormalFocu


Public Function mGetUTCMinutesOffset(sTimeZone) As Integer
    mGetUTCMinutesOffset = -360  'EST with no daylight saving time accounted for
End Function

'End Function


'Public Function gClearAllMasterLinks()
'
'    Dim objHTTP As MSXML2.XMLHTTP60
'    Dim llReturn As Long
'    Dim slRet As String
'    Dim slResponse As String
'    Dim slRetStr As String
'    Dim ilRet As Integer
'
'
'    Screen.MousePointer = vbHourglass
'    ilRet = gGetEDSAutorization
'    gClearAllMasterLinks = False
'    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
'        smRootURL = smRootURL & "DataSync/" & "RemoveAllMasterStationGroupings"
'    End If
'    Set objHTTP = New MSXML2.XMLHTTP60
'    objHTTP.Open "POST", smRootURL
'    objHTTP.setRequestHeader "Content-Type", "application/jsonrequest"
'    objHTTP.setRequestHeader "X-CCS-Gateway-Token", sgToken
'    If bmDemoMode Then
'        llReturn = 200
'        slResponse = "Test Mode, Test Mode"
'    Else
'        objHTTP.Send
'        llReturn = objHTTP.Status
'        slResponse = objHTTP.responseText
'        sgToken = Trim$(objHTTP.getResponseHeader("X-CCS-Gateway-Token"))
'    End If
'    Set objHTTP = Nothing
'    'Anything but 200 is an error.
'        If llReturn = 200 Then
'            gLogMsg "gSend_Post_APIs - " & smRootURL & " " & llReturn & " " & slResponse, "EDS_Log.txt", False
'            gClearAllMasterLinks = True
'        Else
'            gLogMsg "gSend_Post_APIs - " & smRootURL & " " & llReturn & " " & slResponse, "EDS_ErrorLog.Txt", False
'            MsgBox "gSend_Post_APIs - " & smRootURL & " " & llReturn & " " & slResponse
'            lmErrorCnt = lmErrorCnt + 1
'        End If
'End Function

Private Function mGetPersonnel(iShttCode As Integer) As Boolean
    
    Dim rst_artt As ADODB.Recordset
    Dim blRet As Boolean
    Dim slUserNameEmail  As String 'username (same as their email address in this project)
    Dim slUserRights As String 'not implemented yet, pass an empty list
    Dim slFullName As String

    mGetPersonnel = False
    smUserName = ""
    smUserRights = ""
    smUserEmail = ""
    SQLQuery = "Select * from artt where arttShttCode = " & iShttCode & " And ArttType = " & "'" & "P" & "'" & " And ArttState = 0 "
    SQLQuery = SQLQuery & "and arttEmailRights In ('M', 'A', 'V')"
    'Set rst_artt = cnn.Execute(SQLQuery)
    Set rst_artt = gSQLSelectCall(SQLQuery)
    Do While Not rst_artt.EOF
        If Len(rst_artt!arttEmail) > 0 Then
            mGetPersonnel = True
            slUserNameEmail = rst_artt!arttEmail
            'M = Master Accept/Reject; A = Alternate Accept/Reject; V = View; N or Blank = No
            'M = 2, A = 5, V = 4
            If rst_artt!arttEmailRights = "M" Then
                slUserRights = 2
            ElseIf rst_artt!arttEmailRights = "A" Then
                slUserRights = 5
            ElseIf rst_artt!arttEmailRights = "V" Or rst_artt!arttEmailRights = " " Then
                slUserRights = 4
            End If
            slFullName = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
            If smUserName = "" Then
                smUserName = """" & "StationUserFullNames" & """" & ":[" & """" & Trim$(slFullName) & """"
                smUserEmail = """" & "StationUserEmails" & """" & ":[" & """" & Trim$(slUserNameEmail) & """"
                smUserRights = """" & "StationUserRights" & """" & ":[" & """" & Trim$(slUserRights) & """"
            Else
                smUserName = smUserName & "," & """" & Trim$(slFullName) & """"
                smUserEmail = smUserEmail & "," & """" & Trim$(slUserNameEmail) & """"
                smUserRights = smUserRights & "," & """" & Trim$(slUserRights) & """"
            End If
        End If
        rst_artt.MoveNext
    Loop
    If smUserName <> "" Then
        smUserName = smUserName & "],"
        smUserEmail = smUserEmail & "],"
        smUserRights = smUserRights & "],"
    End If
    
End Function
