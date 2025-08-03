Attribute VB_Name = "modEDSSubs"
'EDS Support Sub Routines
'Doug Smith 7/16/15
'Copyright 2015 Counterpoint Software, Inc. All rights reserved.
'Proprietary Software, Do not copy
Option Explicit
Option Compare Text

Private sgToken As String
Private sgGuiPDFDocumentID As String
Public sgNetworkName As String
Public bgEDSIsActive As Boolean
Private smUserAgent As String
Private smFirmId As String
Private smUserName As String
Private smUserPswd As String
Private smEMailDistribution As String
Private smEDSIniPathFileName As String
Private bmDemoMode As Boolean
Private lmErrorCnt As Long
Private smMsg As String
Private bmErrorsFound As Boolean

'INI URL Calls
Private smReqAuthURL As String
Private smRootURL As String
         
'Access Key obtained from EDS
Private smCurUserAPIKey As String

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
    Screen.MousePointer = vbHourglass
    ilRet = mGetEDSAutorization()
    If Not ilRet Then
        Exit Function
    End If
    If mLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
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
            llReturn = 2000000
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
            gLogMsg "gSend_Post_APIs - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody, "EDS_Log.txt", False
            gSend_Post_APIs = True
            Exit For
        Else
            gLogMsg "gSend_Post_APIs - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody, "EDS_ErrorLog.txt", False
            MsgBox "gSend_Post_APIs - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody
            lmErrorCnt = lmErrorCnt + 1
        End If
    Next ilRetries
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gSend_Post_APIs "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDS_Errors.txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function mGetEDSAutorization() As Boolean
    
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
    Screen.MousePointer = vbHourglass
    mGetEDSAutorization = False
    
    If mLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
        smReqAuthURL = smRootURL & "Integrator/Authenticate"
    Else
        MsgBox "[EDS-API] - EDS_RootURL definition not found in the Traffic.ini file"
        Exit Function
    End If
    
    If Not mLoadOption("EDS-API", "EDS_IntegratorID", slIntegratorID) Then
        MsgBox "[EDS-API] - EDS_IntegratorID definition not found in the Traffic.ini file"
        Exit Function
    End If
    
    If Not mLoadOption("EDS-API", "EDS_IntegratorKey", slIntegratorKey) Then
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
            'gLogMsg "mGetEDSAutorization - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody, "EDSLinkLog.txt", False
            mGetEDSAutorization = True
            Exit For
        Else
            If ilRetries = 3 Then
                gLogMsg "mGetEDSAutorization - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody, "EDSLinkErrorsLogLog.txt", False
                MsgBox "gSend_Post_APIs - " & smReqAuthURL & " " & llReturn & " " & slResponse & " " & sBody
            End If
            lmErrorCnt = lmErrorCnt + 1
            'Call Sleep(500)
            Call Sleep(3000)
        End If
    Next ilRetries
    mGetEDSAutorization = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gGetEDSAutorization "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDSLinkLog.txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gRemoveStationFromNetwork(sCallLetters As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    Dim slNetworkName As String
    Dim slTemp As String
        
    On Error GoTo ErrHand
    gRemoveStationFromNetwork = False
    If sCallLetters <> "" Then
        slNetworkName = Trim$(sgClientName)
        If mIsStationAVehicle(sCallLetters) Then
            slTemp = Replace(Trim$(slNetworkName), " ", "%20")
            slBody = "?" & "stationName" & "=" & Trim$(sCallLetters) & "&" & "networkName" & "=" & Trim$(slTemp)
            blRet = gSend_Post_URLs(slBody, "RemoveStationFromNetwork")
            If blRet Then
                gRemoveStationFromNetwork = True
            End If
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function
Public Function mAddOrUpdateSingleStationUser(lArttCode As Long, iShttCode As Integer) As Boolean

    Dim rst_vef As ADODB.Recordset
    Dim rst_artt As ADODB.Recordset
    Dim rst_Shtt As ADODB.Recordset
    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim blRet As Boolean
    Dim slUsernameEmail  As String 'username (same as their email address in this project)
    Dim slUserRights() As String 'not implemented yet, pass an empty list
    Dim blIsActive As Boolean
    Dim slFullName As String
    Dim slBody As String
    Dim slPassword As String
    Dim slStationName As String
    Dim slNetworkName As String
    Dim slCallLetters As String
    Dim ilMinOffset As Integer
    Dim slId As String
    Dim slTemp As String
        
    On Error GoTo ErrHand:
    Screen.MousePointer = vbHourglass
    ilRet = gPopStations()
    ReDim slUserRights(0 To 0)
    slStationName = Trim$(gGetCallLettersByShttCode(iShttCode))
    SQLQuery = "Select vefCode, vefName From VEF_Vehicles where vefType in ('C', 'A', 'G') and vefName = " & "'" & slStationName & "'"
    Set rst_vef = gSQLSelectCall(SQLQuery)
    If Not rst_vef.EOF Then
        ilIdx = 0
        'Do we have a station with a name that matches the vehicle name?
        'slTemp = "Select * from shtt where shttCallLetters = " & "'" & Trim$(rst_vef!vefName) & "'" & "And shttType = 0"
        'Set rst_Shtt = gSQLSelectCall(slTemp)
        'If rst_Shtt!shttClusterGroupID = 0 Or (rst_Shtt!shttClusterGroupID <> 0 And rst_Shtt!shttMasterCluster = "Y") Then
        'If Not rst_Shtt.EOF Then
            SQLQuery = "Select * from artt where arttCode = " & lArttCode & " And ArttType = " & "'" & "P" & "'" & " And ArttState = 0 "
            SQLQuery = SQLQuery & "and arttEmailRights In ('M', 'A', 'V')"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            If Not rst_artt.EOF Then
                If Len(rst_artt!arttEmail) > 0 Then
                    slUsernameEmail = rst_artt!arttEmail
                    'M = Master Accept/Reject; A = Alternate Accept/Reject; V = View; N or Blank = No
                    'M = 2, A = 5, V = 4
                    If rst_artt!arttEmailRights = "M" Then
                        slUserRights(ilIdx) = 2
                    ElseIf rst_artt!arttEmailRights = "A" Then
                        slUserRights(ilIdx) = 5
                    ElseIf rst_artt!arttEmailRights = "V" Or rst_artt!arttEmailRights = " " Then
                        slUserRights(ilIdx) = 4
                    End If
                    'A=Active; D=Dormant
                    If rst_artt!ArttState = 0 Then
                        blIsActive = True
                    Else
                        blIsActive = False
                    End If
                    slFullName = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
                    'Don't supply a password at this time
                    slPassword = ""
                    'Future use, we don't have multiple user rights at this time
                    'ilIdx = ilIdx + 1
                    ilMinOffset = mGetUTCMinuteOffset("")
                    slId = "ARTT" & Format(rst_artt!arttCode, String(5, "0"))
                    slNetworkName = Trim$(sgClientName)
                    'ilMinOffset = mGetUTCMinuteOffset("")
                    slBody = "{" & """" & "UsernameEmail" & """" & ":" & """" & Trim$(slUsernameEmail) & """" & ","
                    slBody = slBody & """" & "UserRights" & """" & ":[" & slUserRights(ilIdx) & "],"
                    slBody = slBody & """" & "IsActive" & """" & ":" & "true" & ","
                    slBody = slBody & """" & "FullName" & """" & ":" & """" & slFullName & """" & ","
                    slBody = slBody & """" & "Password" & """" & ":" & """" & slPassword & """" & ","
                    slBody = slBody & """" & "StationName" & """" & ":" & """" & slStationName & """" & ","
                    slBody = slBody & """" & "UTCMinuteOffset" & """" & ":" & ilMinOffset & ","
                    'slBody = slBody & """" & "ID" & """" & ":" & slId & ","
                    slBody = slBody & """" & "NetworkName" & """" & ":" & """" & slNetworkName & """" & "}"
                    blRet = gSend_Post_APIs(slBody, "AddOrUpdateStationUser")
                End If
            End If
        'End If
    End If
    'blRet = gSend_Post_APIs(slBody, "ChangeStationName")
    mAddOrUpdateSingleStationUser = True
    On Error Resume Next
    rst_vef.Close
    rst_artt.Close
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    rst_vef.Close
    
    Resume Next
    rst_artt.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function

Public Function gAddSingleStation(sStaCallLetters As String, sOldMaster As String, sNewMaster As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    Dim rst_vef As ADODB.Recordset
    Dim rst_Vff As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If gGetEMailDistribution Then
        gAddSingleStation = False
        'Do we have a vehicle with a name that matches the station name?
        SQLQuery = "Select * From VEF_Vehicles where vefName = " & "'" & sStaCallLetters & "'"
        Set rst_vef = gSQLSelectCall(SQLQuery)
        If Not rst_vef.EOF Then
            SQLQuery = "select * from VFF_Vehicle_Features where vffVefCode = " & rst_vef!vefCode
            Set rst_Vff = gSQLSelectCall(SQLQuery)
            If Not rst_Vff.EOF Then
                If rst_Vff!vffOnInsertions = "Y" Then
                    slBody = "{" & """" & "Name" & """" & ":" & """" & Trim$(sStaCallLetters) & """" & "}"
                    blRet = gSend_Post_APIs(slBody, "AddOrUpdateStation")
                    sgNetworkName = Trim$(sgClientName)
                    blRet = gLinkStationToNetwork(Trim$(sStaCallLetters), sgNetworkName)
'                    If blRet Then
'                        blRet = gLinkstations(sStaCallLetters, sOldMaster, UCase(sNewMaster))
'                        If blRet Then
'                            gAddSingleStation = True
'                        End If
'                    End If
                End If
            Else
                gAddSingleStation = True
                Exit Function
            End If
        Else
            gAddSingleStation = True
            Exit Function
        End If
    Else
        gAddSingleStation = True
        Exit Function
    End If
        
    On Error Resume Next
    Screen.MousePointer = vbDefault
    gAddSingleStation = True
    Exit Function
ErrHand:
    On Error Resume Next
    
Resume Next
    
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
Private Function gSend_Post_URLs(sBody As String, sAPI As String) As Boolean
    
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
    Screen.MousePointer = vbHourglass
    ilRet = mGetEDSAutorization()
    gSend_Post_URLs = False
    If mLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
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
            gLogMsg "gSend_Post_APIs - " & smReqAuthURL & " " & llReturn & " " & slResponse, "EDS_Log.txt", False
            gSend_Post_URLs = True
            Exit For
        Else
            gLogMsg "gSend_Post_URLs - " & smRootURL & " " & llReturn & " " & slResponse, "EDS_Errors.txt", False
            MsgBox "gSend_Post_URLs - " & smRootURL & " " & llReturn & " " & slResponse
            lmErrorCnt = lmErrorCnt + 1
            'Call gSleep(2)
        End If
    Next ilRetries
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) Then
        smMsg = "A general error has occured in EDSSubs.bas - gSend_Post_URLs "
        gLogMsg "Error: " & smMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "EDS_Errors.txt", False
        gMsgBox smMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        lmErrorCnt = lmErrorCnt + 1
        bmErrorsFound = True
    End If
    Exit Function
End Function

Public Function gGetEMailDistribution() As Boolean

    Dim slSQLQuery As String
    Dim saf_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gGetEMailDistribution = False
    slSQLQuery = "Select safFeatures2 From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set saf_rst = gSQLSelectCall(slSQLQuery)
    If Not saf_rst.EOF Then
        If (Asc(saf_rst!safFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then
            gGetEMailDistribution = True
        End If
    End If
    saf_rst.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gGetEMailDistribution"
    saf_rst.Close
    Exit Function
    
End Function

Public Function gChangeStationName(sOldName As String, sNewName As String) As Boolean

    Dim slBody As String
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    gChangeStationName = False
    slBody = "?" & "oldName" & "=" & Trim$(sOldName) & "&" & "newName" & "=" & Trim$(sNewName)
    blRet = gSend_Post_URLs(slBody, "ChangeStationName")
    If blRet Then
        gChangeStationName = True
    End If
    Exit Function
ErrHand:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    On Error GoTo 0
End Function


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
    gLogMsg mRoutine & ": " & mErrorCode & " - " & slErrorText, "EDS_Errors.txt", False
End Function

Public Function mLoadOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_mLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    mLoadOption = False
    
    smEDSIniPathFileName = sgStartupDirectory & "\Traffic.Ini"
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, smEDSIniPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            mLoadOption = True
        End If
    End If
    Exit Function
ERR_mLoadOption:
    ' return now if an error occurs
End Function

Private Function mGetUTCMinuteOffset(sTimeZone) As Integer
    mGetUTCMinuteOffset = -360  'EST with no daylight saving time accounted for
End Function

Public Function mIsStationAVehicle(sStationName As String) As Boolean

    Dim rst_tmp As ADODB.Recordset
    
    mIsStationAVehicle = False
    SQLQuery = "Select vefCode, vefName From VEF_Vehicles where vefType in ('C', 'A', 'G') and vefName = " & "'" & Trim$(sStationName) & "'"
    Set rst_tmp = gSQLSelectCall(SQLQuery)
    If Not rst_tmp.EOF Then
        mIsStationAVehicle = True
    End If
End Function

'Public Function gLinkstations(sCallLetters As String, sOldMaster As String, sNewMaster As String) As Boolean
'
'    Dim slTemp As String
'    Dim slTemp2 As String
'    Dim slBody As String
'    Dim blRet As Boolean
'    Dim ilShttCode As Integer
'    Dim rst_Master As ADODB.Recordset
'    Dim rst_Sister As ADODB.Recordset
'    Dim rst_EmailRights As ADODB.Recordset
'    Dim blFound As Boolean
'    Dim ilMasterGrpNum As Integer
'
'    On Error GoTo ErrHand
'    gLinkstations = False
'    ilShttCode = gGetShttCodeFromCallLetters(sCallLetters)
'    SQLQuery = "select * from shtt where shttCode = " & ilShttCode
'    Set rst_Master = gSQLSelectCall(SQLQuery)
'
'    blRet = gRemoveAllStationsFromMasterGroup(sOldMaster)
'
'    While Not rst_Master.EOF
'        'Now find the master station's sister stations
'        slTemp = "select shttCallLetters, shttCode from shtt where shttClusterGroupID = " & rst_Master!shttClusterGroupID & " and shttMasterCluster <> " & "'" & "Y" & "'"
'        Set rst_Sister = gSQLSelectCall(slTemp)
'        slBody = "{" & """" & "MasterStationName" & """" & ":" & """" & Trim(sNewMaster) & """," & """" & "StationNamesInGroup" & """" & ":" & "["
'        While Not rst_Sister.EOF
'            blFound = False
'            slTemp2 = "select arttEmailRights from artt where arttShttCode = " & rst_Master!shttCode
'            slTemp2 = slTemp2 & " and arttEmailrights in ('M','A','V')"
'            Set rst_EmailRights = gSQLSelectCall(slTemp2)
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
'    Set rst_Master = gSQLSelectCall(SQLQuery)
'    While Not rst_Master.EOF
'        'Now find the master station's sister stations
'        slTemp = "select shttCallLetters, shttCode from shtt where shttClusterGroupID = " & rst_Master!shttClusterGroupID & " and shttMasterCluster <> " & "'" & "Y" & "'"
'        Set rst_Sister = gSQLSelectCall(slTemp)
'        slBody = "{" & """" & "MasterStationName" & """" & ":" & """" & Trim(rst_Master!shttCallLetters) & """," & """" & "StationNamesInGroup" & """" & ":" & "["
'        While Not rst_Sister.EOF
'            blFound = False
'            slTemp2 = "select arttEmailRights from artt where arttShttCode = " & rst_Sister!shttCode
'            slTemp2 = slTemp2 & " arttEmailRights In ('M', 'A', 'V')"
'            Set rst_EmailRights = gSQLSelectCall(slTemp2)
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


'Public Function gRemoveAllStationsFromMasterGroup(sMasterStation As String)
'
'    Dim objHttp As MSXML2.XMLHTTP60
'    Dim llReturn As Long
'    Dim slRet As String
'    Dim slResponse As String
'    Dim slRetStr As String
'    Dim ilRet As Integer
'
'    Screen.MousePointer = vbHourglass
'    ilRet = mGetEDSAutorization
'    gRemoveAllStationsFromMasterGroup = False
'    If gLoadOption("EDS-API", "EDS_RootURL", smRootURL) Then
'        smRootURL = smRootURL & "DataSync/" & "RemoveAllStationsFromMasterGroup" & "?" & "masterStationName=" & sMasterStation
'    End If
'    Set objHttp = New MSXML2.XMLHTTP60
'    objHttp.Open "POST", smRootURL
'    objHttp.setRequestHeader "Content-Type", "application/jsonrequest"
'    objHttp.setRequestHeader "X-CCS-Gateway-Token", sgToken
'    If bmDemoMode Then
'        llReturn = 200
'        slResponse = "Test Mode, Test Mode"
'    Else
'        objHttp.Send
'        llReturn = objHttp.Status
'        slResponse = objHttp.responseText
'        sgToken = Trim$(objHttp.getResponseHeader("X-CCS-Gateway-Token"))
'    End If
'    Set objHttp = Nothing
'    'Anything but 200 is an error.
'    If llReturn = 200 Then
'        gLogMsg smReqAuthURL & " " & llReturn & " " & slResponse, "EDS_Log.txt", False
'        gRemoveAllStationsFromMasterGroup = True
'    Else
'        gLogMsg smReqAuthURL & " " & llReturn & " " & slResponse, "EDS_ErrorLog.txt", False
'        MsgBox smReqAuthURL & " " & llReturn & " " & slResponse
'        lmErrorCnt = lmErrorCnt + 1
'    End If
'End Function
'
'Public Function mIsMasterStation(ilShttCode) As Boolean
'
'    Dim rst_Shtt As ADODB.Recordset
'    Dim slTemp As String
'
'    mIsMasterStation = False
'    slTemp = "Select shttClusterGroupID, shttMasterCluster from shtt where shttCode = " & ilShttCode
'    Set rst_Shtt = gSQLSelectCall(slTemp)
'    If Not rst_Shtt.EOF Then
'        If Trim(rst_Shtt!shttMasterCluster) = "Y" Or rst_Shtt!shttClusterGroupID = 0 Then
'            mIsMasterStation = True
'        End If
'    End If
'End Function

'Open the default browser
'ShellExecute 0, vbNullString, "http://www.sony.com/", vbNullString, vbNullString, vbNormalFocus

'*** Open the default emailer ***
'ShellExecute 0, vbNullString, "mailto:dicklevine@counterpoint.net", vbNullString, vbNullString, vbNormalFocus

'*** Open a document with the default document viewer ***
'ShellExecute 0, vbNullString, """C:\house.docx""", vbNullString, vbNullString, vbNormalFocus



